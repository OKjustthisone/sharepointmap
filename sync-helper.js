// sync-helper.js
// 共享同步模块：支持在 options.js、popup.js 以及 background.js 中使用。

const NOTIFICATION_FILE_TYPE_OPTIONS = Object.freeze([
  Object.freeze({
    id: 'ppt',
    extensions: Object.freeze(['ppt', 'pptx', 'pptm', 'pps', 'ppsx', 'ppsm', 'pot', 'potx', 'potm'])
  }),
  Object.freeze({
    id: 'word',
    extensions: Object.freeze(['doc', 'docx', 'docm', 'dot', 'dotx', 'dotm'])
  }),
  Object.freeze({
    id: 'excel',
    extensions: Object.freeze(['xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'xltm'])
  }),
  Object.freeze({
    id: 'pdf',
    extensions: Object.freeze(['pdf'])
  })
]);
const DEFAULT_NOTIFICATION_FILE_TYPES = Object.freeze(['ppt']);
const PRESENTATION_FILE_EXTENSIONS = new Set(
  NOTIFICATION_FILE_TYPE_OPTIONS.find(option => option.id === 'ppt').extensions
);
const FILE_UPDATE_NOTIFICATIONS_KEY = 'file_update_notifications';
const MAX_FILE_UPDATE_NOTIFICATIONS = 100;

function normalizeNotificationFileTypes(value) {
  const values = Array.isArray(value) ? value : DEFAULT_NOTIFICATION_FILE_TYPES;
  const selected = new Set(
    values
      .map(item => String(item).trim().toLowerCase().replace(/^\./, ''))
      .filter(Boolean)
  );

  return NOTIFICATION_FILE_TYPE_OPTIONS
    .filter(option => selected.has(option.id) || option.extensions.some(ext => selected.has(ext)))
    .map(option => option.id);
}

function getNotificationFileExtensions(config) {
  const selectedTypes = normalizeNotificationFileTypes(config?.notificationFileTypes);
  const extensions = new Set();

  selectedTypes.forEach(typeId => {
    const option = NOTIFICATION_FILE_TYPE_OPTIONS.find(item => item.id === typeId);
    option?.extensions.forEach(extension => extensions.add(extension));
  });

  return extensions;
}

function isNotificationFileName(name, config) {
  if (!name || typeof name !== 'string') return false;
  const match = name.toLowerCase().match(/\.([a-z0-9]+)$/);
  return Boolean(match && getNotificationFileExtensions(config).has(match[1]));
}

function isPresentationFileName(name) {
  if (!name || typeof name !== 'string') return false;
  const match = name.toLowerCase().match(/\.([a-z0-9]+)$/);
  return Boolean(match && PRESENTATION_FILE_EXTENSIONS.has(match[1]));
}

function normalizeSharePointDate(value) {
  if (!value) return '';
  const timestamp = Date.parse(value);
  return Number.isNaN(timestamp) ? '' : new Date(timestamp).toISOString();
}

function getItemModifiedAt(item) {
  return normalizeSharePointDate(
    item?.Modified || item?.TimeLastModified || item?.Last_x0020_Modified || ''
  );
}

function getItemCreatedAt(item) {
  return normalizeSharePointDate(
    item?.Created || item?.TimeCreated || item?.Created_x0020_Date || ''
  );
}

function getFileNameFromPath(relativeUrl) {
  if (!relativeUrl) return '';
  const lastPart = relativeUrl.split('/').pop() || '';
  try {
    return decodeURIComponent(lastPart);
  } catch (err) {
    return lastPart;
  }
}

function getCachedItemsById(tree) {
  const itemsById = new Map();
  if (!tree || typeof tree !== 'object') return itemsById;

  Object.values(tree).forEach(node => {
    if (!node) return;
    [...(node.folders || []), ...(node.files || [])].forEach(item => {
      if (item && item.id) itemsById.set(item.id, item);
    });
  });
  return itemsById;
}

function buildFileUpdateEvent(item, eventType, config) {
  if (!item || item.type !== 'file' || !isNotificationFileName(item.name, config)) return null;

  const eventTime = eventType === 'uploaded'
    ? (item.createdAt || item.modifiedAt || 'new')
    : item.modifiedAt;

  // 修改提醒必须有可比较的 Modified 时间，避免每次同步都重复生成提醒。
  if (eventType === 'modified' && !eventTime) return null;

  return {
    fileId: item.id,
    name: item.name,
    relativeUrl: item.relativeUrl,
    webUrl: item.webUrl,
    modifiedAt: item.modifiedAt || '',
    createdAt: item.createdAt || '',
    eventType,
    eventTime
  };
}

async function updateFileUpdateBadge() {
  if (!chrome.action || typeof chrome.action.setBadgeText !== 'function') return;

  try {
    const data = await chrome.storage.local.get(FILE_UPDATE_NOTIFICATIONS_KEY);
    const notifications = Array.isArray(data[FILE_UPDATE_NOTIFICATIONS_KEY])
      ? data[FILE_UPDATE_NOTIFICATIONS_KEY]
      : [];
    const unreadCount = notifications.filter(item => !item.read).length;
    const badgeText = unreadCount === 0 ? '' : (unreadCount > 99 ? '99+' : String(unreadCount));

    await chrome.action.setBadgeText({ text: badgeText });
    if (typeof chrome.action.setBadgeBackgroundColor === 'function') {
      await chrome.action.setBadgeBackgroundColor({ color: '#ff4d6d' });
    }
  } catch (err) {
    console.warn('[SharePoint Map] Failed to update notification badge:', err);
  }
}

async function recordFileUpdateNotifications(configId, config, updateEvents) {
  const events = Array.isArray(updateEvents) ? updateEvents.filter(Boolean) : [];
  if (events.length === 0) {
    await updateFileUpdateBadge();
    return [];
  }

  const data = await chrome.storage.local.get(FILE_UPDATE_NOTIFICATIONS_KEY);
  const existing = Array.isArray(data[FILE_UPDATE_NOTIFICATIONS_KEY])
    ? data[FILE_UPDATE_NOTIFICATIONS_KEY]
    : [];
  const existingKeys = new Set(existing.map(item => item.dedupeKey).filter(Boolean));
  const effectiveConfigId = configId || 'legacy';
  const detectedAt = Date.now();
  const newNotifications = [];

  events.forEach(event => {
    const dedupeKey = [effectiveConfigId, event.fileId, event.eventType, event.eventTime].join('|');
    if (existingKeys.has(dedupeKey)) return;

    const notification = {
      id: `file_update_${effectiveConfigId}_${event.fileId}_${event.eventType}_${event.eventTime}`,
      dedupeKey,
      configId: effectiveConfigId,
      configName: config?.name || config?.libraryName || 'SharePoint 文档库',
      fileId: event.fileId,
      name: event.name,
      type: 'file',
      relativeUrl: event.relativeUrl,
      webUrl: event.webUrl,
      eventType: event.eventType,
      modifiedAt: event.modifiedAt || '',
      createdAt: event.createdAt || '',
      detectedAt,
      read: false
    };

    existingKeys.add(dedupeKey);
    newNotifications.push(notification);
  });

  if (newNotifications.length === 0) {
    await updateFileUpdateBadge();
    return [];
  }

  const allNotifications = [...existing, ...newNotifications]
    .sort((a, b) => (b.detectedAt || 0) - (a.detectedAt || 0))
    .slice(0, MAX_FILE_UPDATE_NOTIFICATIONS);

  await chrome.storage.local.set({
    [FILE_UPDATE_NOTIFICATIONS_KEY]: allNotifications
  });
  await updateFileUpdateBadge();
  return newNotifications;
}

async function markFileUpdateNotificationsRead(ids) {
  const idSet = new Set(Array.isArray(ids) ? ids : []);
  if (idSet.size === 0) return 0;

  const data = await chrome.storage.local.get(FILE_UPDATE_NOTIFICATIONS_KEY);
  const existing = Array.isArray(data[FILE_UPDATE_NOTIFICATIONS_KEY])
    ? data[FILE_UPDATE_NOTIFICATIONS_KEY]
    : [];
  let changedCount = 0;
  const updated = existing.map(item => {
    if (idSet.has(item.id) && !item.read) {
      changedCount += 1;
      return { ...item, read: true };
    }
    return item;
  });

  if (changedCount > 0) {
    await chrome.storage.local.set({ [FILE_UPDATE_NOTIFICATIONS_KEY]: updated });
  }
  await updateFileUpdateBadge();
  return changedCount;
}

async function markAllFileUpdateNotificationsRead(configId) {
  const data = await chrome.storage.local.get(FILE_UPDATE_NOTIFICATIONS_KEY);
  const existing = Array.isArray(data[FILE_UPDATE_NOTIFICATIONS_KEY])
    ? data[FILE_UPDATE_NOTIFICATIONS_KEY]
    : [];
  let changedCount = 0;
  const updated = existing.map(item => {
    if ((!configId || item.configId === configId) && !item.read) {
      changedCount += 1;
      return { ...item, read: true };
    }
    return item;
  });

  if (changedCount > 0) {
    await chrome.storage.local.set({ [FILE_UPDATE_NOTIFICATIONS_KEY]: updated });
  }
  await updateFileUpdateBadge();
  return changedCount;
}

async function removeFileUpdateNotificationsForConfig(configId) {
  if (!configId) return;
  const data = await chrome.storage.local.get(FILE_UPDATE_NOTIFICATIONS_KEY);
  const existing = Array.isArray(data[FILE_UPDATE_NOTIFICATIONS_KEY])
    ? data[FILE_UPDATE_NOTIFICATIONS_KEY]
    : [];
  const updated = existing.filter(item => item.configId !== configId);
  if (updated.length !== existing.length) {
    await chrome.storage.local.set({ [FILE_UPDATE_NOTIFICATIONS_KEY]: updated });
  }
  await updateFileUpdateBadge();
}

// 核心配置迁移函数
async function migrateConfigsIfNeeded() {
  const data = await chrome.storage.local.get(['sp_config', 'sp_configs', 'current_config_id', 'l1_cache', 'favorites']);
  const updates = {};

  if (Array.isArray(data.sp_configs)) {
    let hasChanges = false;
    const normalizedConfigs = data.sp_configs.map(config => {
      const notificationFileTypes = normalizeNotificationFileTypes(config?.notificationFileTypes);
      const currentTypes = Array.isArray(config?.notificationFileTypes)
        ? config.notificationFileTypes
        : null;
      const isSame = currentTypes && currentTypes.length === notificationFileTypes.length
        && currentTypes.every((type, index) => type === notificationFileTypes[index]);

      if (isSame) return config;
      hasChanges = true;
      return { ...config, notificationFileTypes };
    });

    if (hasChanges) updates.sp_configs = normalizedConfigs;
  } else if (data.sp_config && data.sp_config.siteUrl) {
    const defaultId = 'config_default';
    const siteUrl = data.sp_config.siteUrl;
    const libraryName = data.sp_config.libraryName || 'Shared Documents';
    const siteOrigin = data.sp_config.siteOrigin || new URL(siteUrl).origin;
    
    const defaultConfigs = [{
      id: defaultId,
      name: libraryName === 'Shared Documents' ? '默认文档库' : libraryName,
      siteUrl: siteUrl,
      libraryName: libraryName,
      siteOrigin: siteOrigin,
      notificationFileTypes: normalizeNotificationFileTypes(data.sp_config.notificationFileTypes)
    }];
    
    Object.assign(updates, {
      sp_configs: defaultConfigs,
      current_config_id: defaultId
    });
    
    if (data.l1_cache) {
      updates[`l1_cache_${defaultId}`] = data.l1_cache;
    }
    
    if (data.favorites) {
      updates[`favorites_${defaultId}`] = data.favorites;
    }
    
  }

  if (Object.keys(updates).length > 0) {
    await chrome.storage.local.set(updates);
    console.log('[Migration] Normalized SharePoint configuration settings successfully.');
  }
}

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
async function syncLevel1(configId) {
  await migrateConfigsIfNeeded();

  let targetConfigId = configId;
  if (!targetConfigId) {
    const data = await chrome.storage.local.get('current_config_id');
    targetConfigId = data.current_config_id;
  }

  let siteUrl, libraryName;
  if (targetConfigId) {
    const data = await chrome.storage.local.get('sp_configs');
    const configs = data.sp_configs || [];
    const config = configs.find(c => c.id === targetConfigId);
    if (!config) {
      throw new Error('未找到对应的站点配置');
    }
    siteUrl = config.siteUrl;
    libraryName = config.libraryName;
  } else {
    const { sp_config } = await chrome.storage.local.get('sp_config');
    if (!sp_config || !sp_config.siteUrl || !sp_config.libraryName) {
      throw new Error('请先在配置页面设置站点 URL 和文档库名称');
    }
    siteUrl = sp_config.siteUrl;
    libraryName = sp_config.libraryName;
  }
  
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
        modifiedAt: getItemModifiedAt(item),
        createdAt: getItemCreatedAt(item),
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
        modifiedAt: getItemModifiedAt(item),
        createdAt: getItemCreatedAt(item),
        level: 1
      });
    });

    // 自动更新收藏夹中可能发生重命名或路径变更的 1 级项目
    const favKey = targetConfigId ? `favorites_${targetConfigId}` : 'favorites';
    const favData = await chrome.storage.local.get(favKey);
    const favorites = favData[favKey];
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
        await chrome.storage.local.set({ [favKey]: favorites });
        console.log('[SharePoint Map] Self-healed Level 1 items in favorites.');
      }
    }

    // 写入本地存储
    const cacheKey = targetConfigId ? `l1_cache_${targetConfigId}` : 'l1_cache';
    await chrome.storage.local.set({
      [cacheKey]: {
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
async function syncSubtree(l1FolderId, l1FolderRelativeUrl, syncOptions = {}) {
  await migrateConfigsIfNeeded();

  let siteUrl = '';
  let libraryName = '';
  let targetConfigId = '';
  let targetConfig = null;

  const data = await chrome.storage.local.get(['sp_configs', 'sp_config', 'current_config_id']);
  const configs = data.sp_configs || [];
  
  for (const config of configs) {
    const cacheData = await chrome.storage.local.get(`l1_cache_${config.id}`);
    const l1Cache = cacheData[`l1_cache_${config.id}`];
    if (l1Cache && l1Cache.items && l1Cache.items.some(item => item.id === l1FolderId)) {
      siteUrl = config.siteUrl;
      libraryName = config.libraryName;
      targetConfigId = config.id;
      targetConfig = config;
      break;
    }
  }

  if (!siteUrl && data.sp_config) {
    const l1CacheData = await chrome.storage.local.get('l1_cache');
    const l1Cache = l1CacheData.l1_cache;
    if (l1Cache && l1Cache.items && l1Cache.items.some(item => item.id === l1FolderId)) {
      siteUrl = data.sp_config.siteUrl;
      libraryName = data.sp_config.libraryName;
      targetConfig = data.sp_config;
    }
  }

  if (!siteUrl) {
    const currentId = data.current_config_id || (configs[0] ? configs[0].id : '');
    const config = configs.find(c => c.id === currentId);
    if (config) {
      siteUrl = config.siteUrl;
      libraryName = config.libraryName;
      targetConfigId = config.id;
      targetConfig = config;
    } else if (data.sp_config) {
      siteUrl = data.sp_config.siteUrl;
      libraryName = data.sp_config.libraryName;
      targetConfig = data.sp_config;
    }
  }

  if (!siteUrl || !libraryName) {
    throw new Error('站点配置或文档库名称丢失，无法抓取子树');
  }

  const baseUrl = siteUrl.split('/sites/')[0];

  let folderCount = 0;
  let nodeCount = 0;
  let lastReportedFolder = l1FolderRelativeUrl;
  let isSyncActive = true;
  const syncStatusKey = 'sync_status_' + l1FolderId;

  // 定期将当前进度写入 storage，彻底避免多任务并发时的写冲突，节省性能
  const progressTimer = setInterval(async () => {
    if (!isSyncActive) return;
    try {
      const pathParts = lastReportedFolder.split('/');
      const currentFolderName = pathParts[pathParts.length - 1] || lastReportedFolder;

      if (!isSyncActive) return;
      await chrome.storage.local.set({
        [syncStatusKey]: {
          status: 'syncing',
          folderCount: folderCount,
          nodeCount: nodeCount,
          currentFolder: currentFolderName
        }
      });
    } catch (err) {
      console.warn('Failed to write progress:', err);
    }
  }, 300);

  try {
    // 立即写入初始同步状态，避免 300ms 延迟导致看不到同步状态
    try {
      const pathParts = l1FolderRelativeUrl.split('/');
      const currentFolderName = pathParts[pathParts.length - 1] || l1FolderRelativeUrl;
      await chrome.storage.local.set({
        [syncStatusKey]: {
          status: 'syncing',
          folderCount: 0,
          nodeCount: 0,
          currentFolder: currentFolderName
        }
      });
    } catch (err) {
      console.warn('Failed to write initial sync status:', err);
    }

    await setupCookieDNRRule(siteUrl);

    // 获取历史缓存
    const storageKey = 'subtree_cache_' + l1FolderId;
    const storageData = await chrome.storage.local.get(storageKey);
    const cachedData = storageData[storageKey];
    const oldTree = cachedData?.tree;
    const lastSyncTime = cachedData?.last_updated || 0;
    const hasPreviousSnapshot = Boolean(cachedData && cachedData.tree);
    const oldItemsById = getCachedItemsById(oldTree);
    const shouldDetectUpdates = syncOptions.notifyUpdates === true;
    let updateEvents = [];

    // 如果没有历史缓存，或者缓存的根目录路径与当前路径不符，强制进行首次全量同步
    const forceFullSync = !oldTree || !cachedData || !oldTree[l1FolderRelativeUrl];
    let isFullSync = forceFullSync || lastSyncTime === 0;

    let folderTreeCache = {};
    const MAX_FOLDERS = 500;
    const CONCURRENCY = 6;

    const headers = {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json'
    };

    let changedFolders = new Set();

    if (!isFullSync) {
      console.log(`[SharePoint Map] Performing incremental sync for subtree: ${l1FolderRelativeUrl} since ${new Date(lastSyncTime).toISOString()}`);
      
      // 先把旧的缓存树完全复制过来，后续对其进行局部修改和垃圾回收
      folderTreeCache = { ...oldTree };

      // 从原有缓存中构建 id -> { path, type, parent } 映射，用于识别重命名和移动
      const cacheIdToPath = {};
      Object.keys(oldTree).forEach(parentPath => {
        const node = oldTree[parentPath];
        if (node.folders) {
          node.folders.forEach(f => {
            cacheIdToPath[f.id] = { ...f, path: f.relativeUrl, type: 'folder', parent: parentPath };
          });
        }
        if (node.files) {
          node.files.forEach(f => {
            cacheIdToPath[f.id] = { ...f, path: f.relativeUrl, type: 'file', parent: parentPath };
          });
        }
      });

      // 查询在此时间之后修改过的所有文件和文件夹
      const isoString = new Date(lastSyncTime).toISOString();
      const modifiedItemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryName)}')/items?$filter=Modified gt datetime'${isoString}'&$select=FileRef,FileSystemObjectType,UniqueId,FileLeafRef,Created,Modified&$top=5000`;

      try {
        const res = await fetch(modifiedItemsUrl, { method: 'GET', headers });
        if (res.ok) {
          const data = await res.json();
          const items = data.value || [];
          
          items.forEach(item => {
            const newPath = item.FileRef;
            if (!newPath) return;
            
            // 过滤：只保留当前 1 级目录下的变动
            const isUnderL1 = newPath === l1FolderRelativeUrl || newPath.startsWith(l1FolderRelativeUrl + '/');
            if (!isUnderL1) return;

            const isFolderItem = Number(item.FileSystemObjectType) === 1;
            const oldInfo = cacheIdToPath[item.UniqueId];
            const hasPathChanged = oldInfo && oldInfo.path !== newPath;

            if (shouldDetectUpdates && !isFolderItem) {
              const currentItem = {
                id: item.UniqueId,
                name: item.FileLeafRef || getFileNameFromPath(newPath),
                type: 'file',
                relativeUrl: newPath,
                webUrl: `${baseUrl}${newPath}`,
                modifiedAt: getItemModifiedAt(item),
                createdAt: getItemCreatedAt(item)
              };
              const previousItem = oldItemsById.get(item.UniqueId);
              const currentModifiedMs = Date.parse(currentItem.modifiedAt);
              const previousModifiedMs = Date.parse(previousItem?.modifiedAt || '');

              if (!previousItem && hasPreviousSnapshot) {
                const uploadEvent = buildFileUpdateEvent(currentItem, 'uploaded', targetConfig);
                if (uploadEvent) updateEvents.push(uploadEvent);
              } else if (
                previousItem &&
                Number.isFinite(currentModifiedMs) &&
                Number.isFinite(previousModifiedMs) &&
                currentModifiedMs > previousModifiedMs
              ) {
                const modifiedEvent = buildFileUpdateEvent(currentItem, 'modified', targetConfig);
                if (modifiedEvent) updateEvents.push(modifiedEvent);
              }
            }

            if (isFolderItem) {
              // 1. 文件夹变动
              if (hasPathChanged) {
                console.log(`[SharePoint Map] Incremental sync - Folder renamed/moved: ${oldInfo.path} -> ${newPath}`);
                
                // 将新、旧 parent 目录都加入重拉取列表，以刷新子项列表
                const newParentPath = newPath.substring(0, newPath.lastIndexOf('/'));
                changedFolders.add(newParentPath);
                changedFolders.add(oldInfo.parent);

                // 从缓存中删除该文件夹及其下所有子目录旧路径的缓存键
                Object.keys(folderTreeCache).forEach(pathKey => {
                  if (pathKey === oldInfo.path || pathKey.startsWith(oldInfo.path + '/')) {
                    delete folderTreeCache[pathKey];
                  }
                });

                // 将新文件夹路径本身加入重拉取，启动递归抓取其子树
                changedFolders.add(newPath);
              } else {
                // 没有路径变动，只需更新文件夹本身
                changedFolders.add(newPath);
                // 同时也应拉取父级以确保最新
                const parentPath = newPath.substring(0, newPath.lastIndexOf('/'));
                if (parentPath === l1FolderRelativeUrl || parentPath.startsWith(l1FolderRelativeUrl + '/')) {
                  changedFolders.add(parentPath);
                }
              }
            } else {
              // 2. 文件变动
              const parentPath = newPath.substring(0, newPath.lastIndexOf('/'));
              if (parentPath === l1FolderRelativeUrl || parentPath.startsWith(l1FolderRelativeUrl + '/')) {
                changedFolders.add(parentPath);
              }

              if (hasPathChanged) {
                console.log(`[SharePoint Map] Incremental sync - File renamed/moved: ${oldInfo.path} -> ${newPath}`);
                // 新旧父目录均重拉，以便删除旧文件项、添加新文件项
                changedFolders.add(parentPath);
                changedFolders.add(oldInfo.parent);
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
                  webUrl: `${baseUrl}${item.ServerRelativeUrl}`,
                  modifiedAt: getItemModifiedAt(item),
                  createdAt: getItemCreatedAt(item)
                };
                return folderObj;
              });

            const parsedFiles = subFiles.map(item => {
              const fileObj = {
                id: item.UniqueId,
                name: item.Name,
                type: 'file',
                relativeUrl: item.ServerRelativeUrl,
                webUrl: `${baseUrl}${item.ServerRelativeUrl}`,
                modifiedAt: getItemModifiedAt(item),
                createdAt: getItemCreatedAt(item)
              };
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
        const fetchedFolders = new Set();

        while (queue.length > 0) {
          // 过滤掉已经抓取过的路径，防止环路或重复请求
          const nextBatch = [];
          while (queue.length > 0 && nextBatch.length < CONCURRENCY) {
            const path = queue.shift();
            if (!fetchedFolders.has(path)) {
              nextBatch.push(path);
            }
          }
          if (nextBatch.length === 0) continue;

          await Promise.all(nextBatch.map(async (folderPath) => {
            fetchedFolders.add(folderPath);
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
                  // 如果该子文件夹在缓存中不存在且不在已抓取列表中，说明是新增或重命名得到的，加入队列进行深度同步
                  if (!folderTreeCache[item.ServerRelativeUrl] && !fetchedFolders.has(item.ServerRelativeUrl)) {
                    queue.push(item.ServerRelativeUrl);
                  }
                  const folderObj = {
                    id: item.UniqueId,
                    name: item.Name,
                    type: 'folder',
                    relativeUrl: item.ServerRelativeUrl,
                    webUrl: `${baseUrl}${item.ServerRelativeUrl}`,
                    modifiedAt: getItemModifiedAt(item),
                    createdAt: getItemCreatedAt(item)
                  };
                  return folderObj;
                });

              const parsedFiles = subFiles.map(item => {
                const fileObj = {
                  id: item.UniqueId,
                  name: item.Name,
                  type: 'file',
                  relativeUrl: item.ServerRelativeUrl,
                  webUrl: `${baseUrl}${item.ServerRelativeUrl}`,
                  modifiedAt: getItemModifiedAt(item),
                  createdAt: getItemCreatedAt(item)
                };
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

    // 全量同步时用同步前快照比较新增/修改的 PPT。首次建立缓存只做基线，不生成提醒。
    if (shouldDetectUpdates && isFullSync && hasPreviousSnapshot) {
      const newItemsById = getCachedItemsById(folderTreeCache);
      newItemsById.forEach(currentItem => {
        if (currentItem.type !== 'file' || !isNotificationFileName(currentItem.name, targetConfig)) return;

        const previousItem = oldItemsById.get(currentItem.id);
        if (!previousItem) {
          const uploadEvent = buildFileUpdateEvent(currentItem, 'uploaded', targetConfig);
          if (uploadEvent) updateEvents.push(uploadEvent);
          return;
        }

        const currentModifiedMs = Date.parse(currentItem.modifiedAt || '');
        const previousModifiedMs = Date.parse(previousItem.modifiedAt || '');
        if (
          Number.isFinite(currentModifiedMs) &&
          Number.isFinite(previousModifiedMs) &&
          currentModifiedMs > previousModifiedMs
        ) {
          const modifiedEvent = buildFileUpdateEvent(currentItem, 'modified', targetConfig);
          if (modifiedEvent) updateEvents.push(modifiedEvent);
        }
      });
    }

    // 从最终的 folderTreeCache 中重新构建完整的 allDiscoveredItems 映射，确保自愈机制能覆盖到未变动的项目
    const allDiscoveredItems = {};
    Object.keys(folderTreeCache).forEach(parentPath => {
      const node = folderTreeCache[parentPath];
      if (node.folders) {
        node.folders.forEach(f => {
          allDiscoveredItems[f.id] = f;
        });
      }
      if (node.files) {
        node.files.forEach(f => {
          allDiscoveredItems[f.id] = f;
        });
      }
    });

    // 自动更新收藏夹中可能发生重命名或路径变更的深层项目
    const favoritesKey = targetConfigId ? `favorites_${targetConfigId}` : 'favorites';
    const favData = await chrome.storage.local.get(favoritesKey);
    const favoritesList = favData[favoritesKey];
    if (favoritesList && Array.isArray(favoritesList)) {
      let updatedFavs = false;
      favoritesList.forEach(fav => {
        const matchingItem = allDiscoveredItems[fav.id];
        if (matchingItem) {
          if (
            fav.name !== matchingItem.name ||
            fav.relativeUrl !== matchingItem.relativeUrl ||
            fav.webUrl !== matchingItem.webUrl ||
            fav.modifiedAt !== matchingItem.modifiedAt ||
            fav.createdAt !== matchingItem.createdAt
          ) {
            fav.name = matchingItem.name;
            fav.relativeUrl = matchingItem.relativeUrl;
            fav.webUrl = matchingItem.webUrl;
            fav.modifiedAt = matchingItem.modifiedAt;
            fav.createdAt = matchingItem.createdAt;
            updatedFavs = true;
          }
        }
      });
      if (updatedFavs) {
        await chrome.storage.local.set({ [favoritesKey]: favoritesList });
        console.log('[SharePoint Map] Self-healed deep items in favorites.');
      }
    }

    // 更新总 subtree 缓存
    await chrome.storage.local.set({
      [storageKey]: {
        last_updated: Date.now(),
        tree: folderTreeCache
      }
    });

    if (shouldDetectUpdates && updateEvents.length > 0) {
      await recordFileUpdateNotifications(targetConfigId, targetConfig, updateEvents);
    } else if (shouldDetectUpdates) {
      await updateFileUpdateBadge();
    }
    return nodeCount;

  } finally {
    isSyncActive = false;
    clearInterval(progressTimer);
    
    await clearDNRRules();
    // 清除正在同步的状态 (直接删除对应的 Key)
    try {
      await chrome.storage.local.remove(syncStatusKey);
    } catch (err) {
      console.warn('Failed to clear sync status in finally:', err);
    }
  }
}

// 3. 后台定时器全量重新同步主函数
async function performAllSync(options = {}) {
  const notifyUpdates = options.notifyUpdates === true;
  try {
    await migrateConfigsIfNeeded();
    const { sp_configs } = await chrome.storage.local.get('sp_configs');
    
    if (sp_configs && Array.isArray(sp_configs) && sp_configs.length > 0) {
      console.log(`[SharePoint Map] performAllSync: starting sync for ${sp_configs.length} configurations...`);
      for (const config of sp_configs) {
        try {
          console.log(`[SharePoint Map] Syncing config "${config.name}" (${config.id})...`);
          
          // A. 同步 1 级目录
          const l1Items = await syncLevel1(config.id);
          
          // B. 同步所有已收藏
          const favKey = `favorites_${config.id}`;
          const favData = await chrome.storage.local.get(favKey);
          const favorites = favData[favKey];
          if (favorites && Array.isArray(favorites)) {
            const l1Folders = favorites.filter(fav => fav.level === 1 && fav.type === 'folder');
            console.log(`[SharePoint Map] Syncing ${l1Folders.length} favorited Level 1 subtrees for "${config.name}"...`);
            for (const favFolder of l1Folders) {
              const matchingL1 = l1Items.find(item => item.id === favFolder.id);
              if (matchingL1) {
                try {
                  await syncSubtree(favFolder.id, matchingL1.relativeUrl, { notifyUpdates });
                } catch (err) {
                  console.error(`Failed to sync subtree for folder ${favFolder.name}:`, err);
                }
              }
            }
          }
        } catch (configErr) {
          console.error(`Failed to sync config ${config.name} (${config.id}):`, configErr);
        }
      }
    } else {
      // 兼容旧版
      const { sp_config } = await chrome.storage.local.get('sp_config');
      if (!sp_config || !sp_config.siteUrl || !sp_config.libraryName) {
        console.warn('SharePoint config is not completed. Sync aborted.');
        return;
      }
      console.log('Syncing Level 1...');
      const l1Items = await syncLevel1();
      const { favorites } = await chrome.storage.local.get('favorites');
      if (favorites && Array.isArray(favorites)) {
        const l1Folders = favorites.filter(fav => fav.level === 1 && fav.type === 'folder');
        for (const favFolder of l1Folders) {
          const matchingL1 = l1Items.find(item => item.id === favFolder.id);
          if (matchingL1) {
            try {
              await syncSubtree(favFolder.id, matchingL1.relativeUrl, { notifyUpdates });
            } catch (err) {
              console.error(`Failed to sync subtree for folder ${favFolder.name}:`, err);
            }
          }
        }
      }
    }
    console.log('All scheduled sync completed successfully.');
  } catch (error) {
    console.error('Scheduled sync failed:', error);
  }
}

// 共享模糊搜索逻辑
function performFuzzySearchInCache(query, l1Cache, subtreeCache) {
  if (!query) return [];
  const q = query.trim().toLowerCase();
  if (!q) return [];

  // 模糊/宽容字符匹配辅助函数
  function fuzzyMatch(target, searchStr) {
    if (!target) return false;
    const t = target.toLowerCase();
    
    // 1. 直截了当的直接包含匹配
    if (t.includes(searchStr)) return true;

    // 2. 忽略/替换特殊分隔符 (如 -, _, /, \, ., 空格) 为单个空格进行规范化匹配
    const normalize = (str) => str.replace(/[-_/\s.]+/g, ' ').trim();
    const normT = normalize(t);
    const normQ = normalize(searchStr);
    if (normT.includes(normQ)) return true;

    // 3. 无序多词匹配
    const qWords = normQ.split(' ').filter(w => w.length > 0);
    if (qWords.length > 1) {
      const allWordsMatched = qWords.every(word => normT.includes(word));
      if (allWordsMatched) return true;
    }

    // 4. 紧凑无缝匹配
    const strip = (str) => str.replace(/[-_/\s.]+/g, '');
    const strippedT = strip(t);
    const strippedQ = strip(searchStr);
    if (strippedT.includes(strippedQ)) return true;

    return false;
  }

  const results = [];

  // 1. 收集 1 级目录
  if (l1Cache && l1Cache.items) {
    l1Cache.items.forEach(item => {
      results.push(item);
    });
  }

  // 2. 收集所有已缓存的子树节点
  if (subtreeCache) {
    for (const l1Id in subtreeCache) {
      const cacheRoot = subtreeCache[l1Id];
      if (cacheRoot && cacheRoot.tree) {
        Object.keys(cacheRoot.tree).forEach(parentPath => {
          const children = cacheRoot.tree[parentPath];
          if (children) {
            (children.folders || []).forEach(f => results.push(f));
            (children.files || []).forEach(f => results.push(f));
          }
        });
      }
    }
  }

  // 去重
  const uniqueMap = new Map();
  results.forEach(item => {
    uniqueMap.set(item.id || item.relativeUrl, item);
  });

  const searchPool = Array.from(uniqueMap.values());

  // 模糊匹配
  return searchPool.filter(item => {
    return fuzzyMatch(item.name, q) || fuzzyMatch(item.relativeUrl, q);
  });
}
