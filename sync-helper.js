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
  // 1. 转义单引号以防止 OData 语句解析截断
  const escapedPath = relativePath.replace(/'/g, "''");
  // 2. 拼接原始 API 地址
  const rawUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapedPath}')/${apiType}`;
  // 3. 使用 encodeURI 对整条 URL 进行编码。
  // encodeURI 的精妙之处在于它会转义空格 (%20) 和中文字符，但会严格保留斜杠 (/) 和单引号 (') 不予转义。
  // 这能完美避免 encodeURIComponent 将斜杠转义为 %2F 从而导致 SharePoint 路由报 404 的问题。
  return encodeURI(rawUrl);
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

// 2. 递归抓取已收藏的 1 级文件夹子树 (BFS)
async function syncSubtree(l1FolderId, l1FolderRelativeUrl) {
  const { sp_config } = await chrome.storage.local.get('sp_config');
  if (!sp_config || !sp_config.siteUrl) {
    throw new Error('站点配置丢失，无法抓取子树');
  }

  const { siteUrl } = sp_config;
  const baseUrl = siteUrl.split('/sites/')[0];

  try {
    await setupCookieDNRRule(siteUrl);

    const folderTreeCache = {};
    const queue = [l1FolderRelativeUrl];
    let nodeCount = 0;
    
    const MAX_FOLDERS = 500; 
    let folderCount = 0;

    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json'
    };

    while (queue.length > 0) {
      const currentRelativeUrl = queue.shift();
      folderCount++;

      if (folderCount > MAX_FOLDERS) {
        console.warn(`Subtree sync exceeded limit (${MAX_FOLDERS}). Stopping.`);
        break;
      }

      // 【修复】使用 cleanRelativePathUrl 对相对路径进行安全转义，保留关键的斜杠 /，防止报 404
      const foldersUrl = cleanRelativePathUrl(siteUrl, currentRelativeUrl, 'Folders');
      const filesUrl = cleanRelativePathUrl(siteUrl, currentRelativeUrl, 'Files');
      
      try {
        const [foldersRes, filesRes] = await Promise.all([
          fetch(foldersUrl, { method: 'GET', headers }),
          fetch(filesUrl, { method: 'GET', headers })
        ]);

        if (!foldersRes.ok || !filesRes.ok) {
          console.error(`Failed to fetch subfolder data for: ${currentRelativeUrl}, HTTP ${foldersRes.status} / ${filesRes.status}`);
          continue;
        }

        const foldersData = await foldersRes.json();
        const filesData = await filesRes.json();

        const subFolders = foldersData.value || [];
        const subFiles = filesData.value || [];

        // 按名称字母升序自然排序
        subFolders.sort((a, b) => a.Name.localeCompare(b.Name, 'zh-CN', { numeric: true }));
        subFiles.sort((a, b) => a.Name.localeCompare(b.Name, 'zh-CN', { numeric: true }));

        const parsedFolders = subFolders
          .filter(item => item.Name !== 'Forms')
          .map(item => {
            queue.push(item.ServerRelativeUrl);
            nodeCount++;
            return {
              id: item.UniqueId,
              name: item.Name,
              type: 'folder',
              relativeUrl: item.ServerRelativeUrl,
              webUrl: `${baseUrl}${item.ServerRelativeUrl}`
            };
          });

        const parsedFiles = subFiles.map(item => {
          nodeCount++;
          return {
            id: item.UniqueId,
            name: item.Name,
            type: 'file',
            relativeUrl: item.ServerRelativeUrl,
            webUrl: `${baseUrl}${item.ServerRelativeUrl}`
          };
        });

        folderTreeCache[currentRelativeUrl] = {
          folders: parsedFolders,
          files: parsedFiles
        };

      } catch (err) {
        console.error(`Error requesting folder data for ${currentRelativeUrl}:`, err);
      }

      // 延迟 50ms 避开限流
      await new Promise(resolve => setTimeout(resolve, 50));
    }

    // 读取并更新缓存
    const storageData = await chrome.storage.local.get('subtree_cache');
    const subtreeCache = storageData.subtree_cache || {};

    subtreeCache[l1FolderId] = {
      last_updated: Date.now(),
      tree: folderTreeCache
    };

    await chrome.storage.local.set({ subtree_cache: subtreeCache });
    return nodeCount;

  } finally {
    await clearDNRRules();
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
