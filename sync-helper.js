// sync-helper.js
// 共享同步模块：支持在 options.js、popup.js 以及 background.js 中使用。

// 立即清理历史所有 DNR 规则，恢复正常的浏览器 SharePoint 网页访问！
clearDNRRules().catch(err => console.error('Failed to clear rules on load:', err));

async function clearDNRRules() {
  const rules = await new Promise((resolve) => {
    chrome.declarativeNetRequest.getDynamicRules(resolve);
  });
  const ids = (rules || []).map(r => r.id);
  if (ids.length > 0) {
    await new Promise((resolve) => {
      chrome.declarativeNetRequest.updateDynamicRules({
        removeRuleIds: ids
      }, resolve);
    });
    console.log('[SharePoint Map] Cleared all active DNR rules to restore normal browsing:', ids);
  }
}

// 辅助函数：安全转义 SharePoint 相对路径 URL (保留斜杠并处理空格与单引号)
function cleanRelativePathUrl(siteUrl, relativePath, apiType) {
  // 1. 转义单引号以防止 OData 语句解析截断 (两单引号表示转义)
  const escapedPath = relativePath.replace(/'/g, "''");
  // 2. 使用 encodeURIComponent 对路径进行完整编码，确保所有的特殊字符 (例如 #、%、& 等) 被正确编码，避免 HTTP 解析问题。
  // 在 GetFolderByServerRelativePath 中，路径是作为字符串参数传递的，因此斜杠 / 编码为 %2F 也是完全被支持的。
  const encodedPath = encodeURIComponent(escapedPath);
  // 3. 拼接 API 地址 (使用 GetFolderByServerRelativePath API)
  return `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${encodedPath}')/${apiType}`;
}

// 核心函数：使用 url 参数读取 Cookie（自动获取父域名如 .sharepoint.com 的授权 Cookie）
async function setupCookieDNRRule(siteUrl) {
  const parsedUrl = new URL(siteUrl);
  const domain = parsedUrl.hostname;
  const targetOrigin = parsedUrl.origin;

  // 1. 利用 chrome.cookies API 读取该 URL 有效的所有 Cookie
  const cookies = await new Promise((resolve) => {
    chrome.cookies.getAll({ url: siteUrl }, (result) => {
      resolve(result || []);
    });
  });

  if (cookies.length === 0) {
    throw new Error('未检测到您在浏览器中登录过此站点。请先在常规标签页中登录该 SharePoint 网站。');
  }

  // 格式化拼接为 Cookie 头字符串
  const cookieStr = cookies.map(c => `${c.name}=${c.value}`).join('; ');
  const extensionId = chrome.runtime.id;

  // 2. 配置 DNR 规则 (动态规则 ID 定为 1)
  const ruleId = 1;
  const rule = {
    id: ruleId,
    priority: 1,
    action: {
      type: "modifyHeaders",
      requestHeaders: [
        {
          header: "Cookie",
          operation: "set",
          value: cookieStr
        },
        {
          header: "Origin",
          operation: "remove"
        },
        {
          header: "Referer",
          operation: "set",
          value: targetOrigin + "/"
        }
      ]
    },
    condition: {
      urlFilter: "||" + domain,
      initiatorDomains: [extensionId], 
      resourceTypes: ["xmlhttprequest"]
    }
  };

  // 3. 应用动态规则
  await new Promise((resolve, reject) => {
    chrome.declarativeNetRequest.updateDynamicRules({
      removeRuleIds: [ruleId],
      addRules: [rule]
    }, () => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        console.log(`[SharePoint Map] CORS-bypass DNR rule registered for domain: ${domain}`);
        resolve();
      }
    });
  });
}

// 1. 同步并缓存 1 级目录
async function syncLevel1() {
  const { sp_config } = await chrome.storage.local.get('sp_config');
  if (!sp_config || !sp_config.siteUrl || !sp_config.libraryName) {
    throw new Error('请先在配置页面设置站点 URL 和文档库名称');
  }

  const { siteUrl, libraryName } = sp_config;
  
  try {
    await setupCookieDNRRule(siteUrl);

    const foldersApi = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryName)}')/RootFolder/Folders`;
    const filesApi = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryName)}')/RootFolder/Files`;

    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json'
    };

    const [foldersRes, filesRes] = await Promise.all([
      fetch(foldersApi, { method: 'GET', headers }),
      fetch(filesApi, { method: 'GET', headers })
    ]);

    if (!foldersRes.ok) {
      throw new Error(`获取文件夹失败 (HTTP ${foldersRes.status}): 请确保文档库名称正确。`);
    }
    if (!filesRes.ok) {
      throw new Error(`获取文件失败 (HTTP ${filesRes.status}): 请确保文档库名称正确。`);
    }

    const foldersData = await foldersRes.json();
    const filesData = await filesRes.json();

    const folders = foldersData.value || [];
    const files = filesData.value || [];

    // 按名称字母升序自然排序（支持中文拼音与数字自然排序）
    folders.sort((a, b) => a.Name.localeCompare(b.Name, 'zh-CN', { numeric: true }));
    files.sort((a, b) => a.Name.localeCompare(b.Name, 'zh-CN', { numeric: true }));

    const items = [];

    // 处理文件夹
    folders.forEach(item => {
      if (item.Name === 'Forms') return;
      items.push({
        id: item.UniqueId,
        name: item.Name,
        type: 'folder',
        relativeUrl: item.ServerRelativeUrl,
        webUrl: `${siteUrl.split('/sites/')[0]}${item.ServerRelativeUrl}`,
        level: 1
      });
    });

    // 处理文件
    files.forEach(item => {
      items.push({
        id: item.UniqueId,
        name: item.Name,
        type: 'file',
        relativeUrl: item.ServerRelativeUrl,
        webUrl: `${siteUrl.split('/sites/')[0]}${item.ServerRelativeUrl}`,
        level: 1
      });
    });

    // 自动更新收藏夹中可能发生重命名或路径变更的 1 级项目
    const { favorites } = await chrome.storage.local.get('favorites');
    if (favorites && Array.isArray(favorites)) {
      let updatedFavs = false;
      favorites.forEach(fav => {
        const matchingItem = items.find(item => item.id === fav.id);
        if (matchingItem) {
          if (fav.name !== matchingItem.name || fav.relativeUrl !== matchingItem.relativeUrl || fav.webUrl !== matchingItem.webUrl) {
            fav.name = matchingItem.name;
            fav.relativeUrl = matchingItem.relativeUrl;
            fav.webUrl = matchingItem.webUrl;
            updatedFavs = true;
          }
        }
      });
      if (updatedFavs) {
        await chrome.storage.local.set({ favorites: favorites });
        console.log('[SharePoint Map] Self-healed Level 1 items in favorites.');
      }
    }

    // 写入本地存储
    await chrome.storage.local.set({
      l1_cache: {
        last_updated: Date.now(),
        items: items
      }
    });

    return items;

  } finally {
    await clearDNRRules();
  }
}

// 2. 并行且高效地抓取已收藏的 1 级文件夹子树 (支持增量和全量更新)
async function syncSubtree(l1FolderId, l1FolderRelativeUrl) {
  const { sp_config } = await chrome.storage.local.get('sp_config');
  if (!sp_config || !sp_config.siteUrl || !sp_config.libraryName) {
    throw new Error('站点配置或文档库名称丢失，无法抓取子树');
  }

  const { siteUrl, libraryName } = sp_config;
  const baseUrl = siteUrl.split('/sites/')[0];

  let folderCount = 0;
  let nodeCount = 0;
  let lastReportedFolder = l1FolderRelativeUrl;
  let isSyncActive = true;

  // 定期将当前进度写入 storage，彻底避免多任务并发时的写冲突，节省性能
  const progressTimer = setInterval(async () => {
    if (!isSyncActive) return;
    try {
      const statusData = await chrome.storage.local.get('sync_status');
      const syncStatus = statusData.sync_status || {};
      const pathParts = lastReportedFolder.split('/');
      const currentFolderName = pathParts[pathParts.length - 1] || lastReportedFolder;

      syncStatus[l1FolderId] = {
        status: 'syncing',
        folderCount: folderCount,
        nodeCount: nodeCount,
        currentFolder: currentFolderName
      };
      await chrome.storage.local.set({ sync_status: syncStatus });
    } catch (err) {
      console.warn('Failed to write progress:', err);
    }
  }, 300);

  try {
    await setupCookieDNRRule(siteUrl);

    // 获取历史缓存
    const storageData = await chrome.storage.local.get('subtree_cache');
    const subtreeCache = storageData.subtree_cache || {};
    const cachedData = subtreeCache[l1FolderId];
    const oldTree = cachedData?.tree;
    const lastSyncTime = cachedData?.last_updated || 0;

    // 如果没有历史缓存，或者缓存的根目录路径与当前路径不符，强制进行首次全量同步
    const forceFullSync = !oldTree || !cachedData || !oldTree[l1FolderRelativeUrl];
    let isFullSync = forceFullSync || lastSyncTime === 0;

    let folderTreeCache = {};
    const allDiscoveredItems = {};
    const MAX_FOLDERS = 500;
    const CONCURRENCY = 6;

    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json'
    };

    let changedFolders = new Set();

    if (!isFullSync) {
      console.log(`[SharePoint Map] Performing incremental sync for subtree: ${l1FolderRelativeUrl} since ${new Date(lastSyncTime).toISOString()}`);
      // 查询在此之后修改过的所有文件和文件夹，以定位发生过变动的目录
      const isoString = new Date(lastSyncTime).toISOString();
      const modifiedItemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryName)}')/items?$filter=Modified gt datetime'${isoString}'&$select=FileRef,FileSystemObjectType,UniqueId&$top=5000`;

      try {
        const res = await fetch(modifiedItemsUrl, { method: 'GET', headers });
        if (res.ok) {
          const data = await res.json();
          const items = data.value || [];
          
          items.forEach(item => {
            const path = item.FileRef;
            if (!path) return;
            
            // 过滤：只保留当前 1 级目录下的变动
            const isUnderL1 = path === l1FolderRelativeUrl || path.startsWith(l1FolderRelativeUrl + '/');
            if (!isUnderL1) return;

            if (item.FileSystemObjectType === 1) {
              // 变动项本身是文件夹，直接加入重拉取列表
              changedFolders.add(path);
            } else {
              // 变动项是文件，将其父级文件夹加入重拉取列表
              const lastSlashIndex = path.lastIndexOf('/');
              if (lastSlashIndex > 0) {
                const parentPath = path.substring(0, lastSlashIndex);
                if (parentPath === l1FolderRelativeUrl || parentPath.startsWith(l1FolderRelativeUrl + '/')) {
                  changedFolders.add(parentPath);
                }
              }
            }
          });
        } else {
          console.warn(`Incremental query failed with HTTP ${res.status}. Falling back to full sync.`);
          isFullSync = true;
        }
      } catch (err) {
        console.warn('Incremental query failed. Falling back to full sync:', err);
        isFullSync = true;
      }
    }

    if (isFullSync) {
      console.log(`[SharePoint Map] Performing full sync for subtree: ${l1FolderRelativeUrl}`);
      const queue = [l1FolderRelativeUrl];

      while (queue.length > 0 && folderCount < MAX_FOLDERS) {
        // 一次性取出 CONCURRENCY 个要处理的路径
        const batch = queue.splice(0, CONCURRENCY);

        await Promise.all(batch.map(async (currentRelativeUrl) => {
          folderCount++;

          // 拼接展开子目录与文件的 URL，一次请求获取该目录下全部子项
          let folderUrl = cleanRelativePathUrl(siteUrl, currentRelativeUrl, '');
          if (folderUrl.endsWith('/')) {
            folderUrl = folderUrl.slice(0, -1);
          }
          folderUrl += '?$expand=Folders,Files';

          try {
            const res = await fetch(folderUrl, { method: 'GET', headers });
            if (!res.ok) {
              console.error(`Failed to fetch subfolder: ${currentRelativeUrl}, HTTP ${res.status}`);
              if (currentRelativeUrl === l1FolderRelativeUrl) {
                throw new Error(`获取该文件夹的子目录失败 (HTTP ${res.status})。请确保您已登录网页版，且对该文件夹有访问权限。`);
              }
              return;
            }

            const data = await res.json();
            const subFolders = data.Folders || [];
            const subFiles = data.Files || [];

            // 自然排序
            subFolders.sort((a, b) => a.Name.localeCompare(b.Name, 'zh-CN', { numeric: true }));
            subFiles.sort((a, b) => a.Name.localeCompare(b.Name, 'zh-CN', { numeric: true }));

            const parsedFolders = subFolders
              .filter(item => item.Name !== 'Forms')
              .map(item => {
                queue.push(item.ServerRelativeUrl);
                const folderObj = {
                  id: item.UniqueId,
                  name: item.Name,
                  type: 'folder',
                  relativeUrl: item.ServerRelativeUrl,
                  webUrl: `${baseUrl}${item.ServerRelativeUrl}`
                };
                allDiscoveredItems[item.UniqueId] = folderObj;
                return folderObj;
              });

            const parsedFiles = subFiles.map(item => {
              const fileObj = {
                id: item.UniqueId,
                name: item.Name,
                type: 'file',
                relativeUrl: item.ServerRelativeUrl,
                webUrl: `${baseUrl}${item.ServerRelativeUrl}`
              };
              allDiscoveredItems[item.UniqueId] = fileObj;
              return fileObj;
            });

            folderTreeCache[currentRelativeUrl] = {
              folders: parsedFolders,
              files: parsedFiles
            };

            nodeCount += parsedFolders.length + parsedFiles.length;
            lastReportedFolder = currentRelativeUrl;

          } catch (err) {
            console.error(`Error requesting folder data for ${currentRelativeUrl}:`, err);
            if (currentRelativeUrl === l1FolderRelativeUrl) {
              throw err;
            }
          }
        }));

        // 每次并发批处理后留微弱间隔，配合 SharePoint 的防爬虫策略
        await new Promise(resolve => setTimeout(resolve, 50));
      }
    } else {
      // 增量同步路径：拷贝原缓存，按需仅拉取有变动的文件夹
      folderTreeCache = { ...oldTree };
      const changedFoldersArray = Array.from(changedFolders);

      if (changedFoldersArray.length > 0) {
        console.log(`[SharePoint Map] Incremental sync: ${changedFoldersArray.length} folders modified. Updating them...`);
        const queue = [...changedFoldersArray];

        while (queue.length > 0) {
          const batch = queue.splice(0, CONCURRENCY);

          await Promise.all(batch.map(async (folderPath) => {
            folderCount++;

            let folderUrl = cleanRelativePathUrl(siteUrl, folderPath, '');
            if (folderUrl.endsWith('/')) {
              folderUrl = folderUrl.slice(0, -1);
            }
            folderUrl += '?$expand=Folders,Files';

            try {
              const res = await fetch(folderUrl, { method: 'GET', headers });
              if (!res.ok) {
                console.warn(`Failed to fetch modified subfolder: ${folderPath}, HTTP ${res.status}`);
                // 若该目录已从 SharePoint 移除，忽略即可，垃圾回收会处理它
                return;
              }

              const data = await res.json();
              const subFolders = data.Folders || [];
              const subFiles = data.Files || [];

              subFolders.sort((a, b) => a.Name.localeCompare(b.Name, 'zh-CN', { numeric: true }));
              subFiles.sort((a, b) => a.Name.localeCompare(b.Name, 'zh-CN', { numeric: true }));

              const parsedFolders = subFolders
                .filter(item => item.Name !== 'Forms')
                .map(item => {
                  const folderObj = {
                    id: item.UniqueId,
                    name: item.Name,
                    type: 'folder',
                    relativeUrl: item.ServerRelativeUrl,
                    webUrl: `${baseUrl}${item.ServerRelativeUrl}`
                  };
                  allDiscoveredItems[item.UniqueId] = folderObj;
                  return folderObj;
                });

              const parsedFiles = subFiles.map(item => {
                const fileObj = {
                  id: item.UniqueId,
                  name: item.Name,
                  type: 'file',
                  relativeUrl: item.ServerRelativeUrl,
                  webUrl: `${baseUrl}${item.ServerRelativeUrl}`
                };
                allDiscoveredItems[item.UniqueId] = fileObj;
                return fileObj;
              });

              folderTreeCache[folderPath] = {
                folders: parsedFolders,
                files: parsedFiles
              };

              nodeCount += parsedFolders.length + parsedFiles.length;
              lastReportedFolder = folderPath;

            } catch (err) {
              console.error(`Error requesting modified folder data for ${folderPath}:`, err);
            }
          }));

          await new Promise(resolve => setTimeout(resolve, 50));
        }

        // 垃圾回收：清理已在 SharePoint 中被删除的文件夹缓存
        const reachable = new Set([l1FolderRelativeUrl]);
        const scanQueue = [l1FolderRelativeUrl];
        while (scanQueue.length > 0) {
          const current = scanQueue.shift();
          const node = folderTreeCache[current];
          if (node && node.folders) {
            node.folders.forEach(f => {
              if (!reachable.has(f.relativeUrl)) {
                reachable.add(f.relativeUrl);
                scanQueue.push(f.relativeUrl);
              }
            });
          }
        }

        // 剔除不可达路径
        Object.keys(folderTreeCache).forEach(path => {
          if (!reachable.has(path)) {
            delete folderTreeCache[path];
          }
        });
      } else {
        console.log('[SharePoint Map] Incremental sync: No files/folders modified since last sync.');
      }
    }

    // 自动更新收藏夹中可能发生重命名或路径变更的深层项目
    const favData = await chrome.storage.local.get('favorites');
    const favoritesList = favData.favorites;
    if (favoritesList && Array.isArray(favoritesList)) {
      let updatedFavs = false;
      favoritesList.forEach(fav => {
        const matchingItem = allDiscoveredItems[fav.id];
        if (matchingItem) {
          if (fav.name !== matchingItem.name || fav.relativeUrl !== matchingItem.relativeUrl || fav.webUrl !== matchingItem.webUrl) {
            fav.name = matchingItem.name;
            fav.relativeUrl = matchingItem.relativeUrl;
            fav.webUrl = matchingItem.webUrl;
            updatedFavs = true;
          }
        }
      });
      if (updatedFavs) {
        await chrome.storage.local.set({ favorites: favoritesList });
        console.log('[SharePoint Map] Self-healed deep items in favorites.');
      }
    }

    // 更新总 subtree 缓存
    subtreeCache[l1FolderId] = {
      last_updated: Date.now(),
      tree: folderTreeCache
    };

    await chrome.storage.local.set({ subtree_cache: subtreeCache });
    return nodeCount;

  } finally {
    isSyncActive = false;
    clearInterval(progressTimer);
    
    await clearDNRRules();
    // 清除正在同步的状态
    const statusData = await chrome.storage.local.get('sync_status');
    const syncStatus = statusData.sync_status || {};
    delete syncStatus[l1FolderId];
    await chrome.storage.local.set({ sync_status: syncStatus });
  }
}

// 3. 后台定时器全量重新同步主函数
async function performAllSync() {
  try {
    const { sp_config } = await chrome.storage.local.get('sp_config');
    if (!sp_config || !sp_config.siteUrl || !sp_config.libraryName) {
      console.warn('SharePoint config is not completed. Sync aborted.');
      return;
    }

    // A. 同步 1 级目录
    console.log('Syncing Level 1...');
    const l1Items = await syncLevel1();
    
    // B. 同步所有已收藏 of 1 级目录子树
    const { favorites } = await chrome.storage.local.get('favorites');
    if (favorites && Array.isArray(favorites)) {
      const l1Folders = favorites.filter(fav => fav.level === 1 && fav.type === 'folder');
      console.log(`Syncing ${l1Folders.length} favorited Level 1 subtrees...`);
      for (const favFolder of l1Folders) {
        const matchingL1 = l1Items.find(item => item.id === favFolder.id);
        if (matchingL1) {
          try {
            await syncSubtree(favFolder.id, matchingL1.relativeUrl);
          } catch (err) {
            console.error(`Failed to sync subtree for folder ${favFolder.name}:`, err);
          }
        }
      }
    }
    console.log('All scheduled sync completed successfully.');
  } catch (error) {
    console.error('Scheduled sync failed:', error);
  }
}
