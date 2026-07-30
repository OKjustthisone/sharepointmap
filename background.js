// background.js
// 引入共享同步逻辑
importScripts('sync-helper.js');

// 启动时清除任何遗留的正在同步状态，防止 Service Worker 重启或崩溃后状态卡在“同步中”
chrome.storage.local.get(null, (allData) => {
  const keysToRemove = Object.keys(allData).filter(key => key.startsWith('sync_status_') || key === 'sync_status');
  if (keysToRemove.length > 0) {
    chrome.storage.local.remove(keysToRemove, () => {
      console.log('[SharePoint Map] Cleared stale sync statuses on startup:', keysToRemove);
    });
  }
});

// 计算下一个工作日 10 点的时间戳
function getNextWeekday10AM(now) {
  const date = new Date(now);
  date.setHours(10, 0, 0, 0);

  // 如果当前时间已经过了今天的 10 点，则移到明天
  if (date.getTime() <= now) {
    date.setDate(date.getDate() + 1);
  }

  // 过滤掉周六 (6) 和周日 (0)，如果是周末则持续往后移直到周一
  while (date.getDay() === 0 || date.getDay() === 6) {
    date.setDate(date.getDate() + 1);
  }

  return date.getTime();
}

function scheduleNextWeekdayAlarm() {
  const nextTime = getNextWeekday10AM(Date.now());
  chrome.alarms.create('weekday_sync_alarm', { when: nextTime });
  console.log('Scheduled next weekday sync for:', new Date(nextTime).toString());
}

// 检查并确保工作日 10 点的 Alarm 已设置
chrome.alarms.get('weekday_sync_alarm', (alarm) => {
  if (!alarm) {
    scheduleNextWeekdayAlarm();
  } else {
    console.log('Weekday sync alarm already scheduled for:', new Date(alarm.scheduledTime).toString());
  }
});

// 监听安装事件，设置定时任务与右键菜单
chrome.runtime.onInstalled.addListener(() => {
  console.log('SharePoint Quick Access extension installed.');
  // 设置 7 天定时任务 (7 * 24 * 60 分钟)
  chrome.alarms.create('sync_all_data', { periodInMinutes: 7 * 24 * 60 });
  // 设置工作日 10 点定时同步
  scheduleNextWeekdayAlarm();

  // 创建右键菜单以支持选中文本在网页中搜索
  chrome.contextMenus.create({
    id: "search_sp_map",
    title: "在 SharePoint Map 中搜索 '%s'",
    contexts: ["selection"]
  });
});

// 监听 Alarm 触发
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === 'weekday_sync_alarm') {
    console.log('Weekday 10 AM sync alarm triggered. Syncing all favorited directories...');
    performAllSync()
      .then(() => {
        console.log('Weekday 10 AM sync completed successfully.');
      })
      .catch((err) => {
        console.error('Weekday 10 AM sync failed:', err);
      })
      .finally(() => {
        // 无论成功还是失败，都安排下一次的工作日同步
        scheduleNextWeekdayAlarm();
      });
  } else if (alarm.name === 'sync_all_data') {
    console.log('Scheduled alarm triggered. Syncing all SharePoint data...');
    performAllSync();
  }
});

// 消息监听保留，以备将来需要
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'sync_all') {
    performAllSync()
      .then(() => sendResponse({ success: true }))
      .catch((err) => sendResponse({ success: false, error: err.message || err }));
    return true; // 异步通道
  }
  if (request.action === 'sync_level1') {
    syncLevel1(request.configId)
      .then((items) => sendResponse({ success: true, count: items.length }))
      .catch((err) => sendResponse({ success: false, error: err.message || err }));
    return true; // 异步通道
  }
  if (request.action === 'sync_subtree') {
    const { folderId, relativeUrl } = request;
    syncSubtree(folderId, relativeUrl)
      .then((nodeCount) => sendResponse({ success: true, count: nodeCount }))
      .catch((err) => sendResponse({ success: false, error: err.message || err }));
    return true;
  }
});

// 监听右键菜单点击
chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId === "search_sp_map" && info.selectionText) {
    const query = info.selectionText;
    
    // 执行配置迁移（防万一）
    await migrateConfigsIfNeeded();

    const configData = await chrome.storage.local.get(['sp_configs', 'current_config_id', 'sp_config']);
    const spConfigs = configData.sp_configs || [];
    const currentConfigId = configData.current_config_id || '';
    
    let activeConfig = null;
    if (currentConfigId) {
      activeConfig = spConfigs.find(c => c.id === currentConfigId);
    } else if (configData.sp_config) {
      activeConfig = configData.sp_config;
    }

    if (!activeConfig || !activeConfig.siteUrl) {
      // 提示未配置
      try {
        chrome.tabs.sendMessage(tab.id, {
          action: 'show_search_results',
          query: query,
          results: []
        });
      } catch (err) {
        console.warn('Failed to send message to tab:', err);
      }
      return;
    }

    // 获取该配置下的 L1 cache
    const l1CacheKey = currentConfigId ? `l1_cache_${currentConfigId}` : 'l1_cache';
    const l1Data = await chrome.storage.local.get(l1CacheKey);
    const l1Cache = l1Data[l1CacheKey];

    // 获取所有 subtree_cache
    const allStorage = await chrome.storage.local.get(null);
    const subtreeCache = {};
    
    if (l1Cache && l1Cache.items) {
      const l1Ids = l1Cache.items.map(item => item.id);
      Object.keys(allStorage).forEach(key => {
        if (key.startsWith('subtree_cache_')) {
          const l1Id = key.substring('subtree_cache_'.length);
          if (l1Ids.includes(l1Id)) {
            subtreeCache[l1Id] = allStorage[key];
          }
        }
      });
    }

    // 收集一级子目录列表
    const l1Directories = [];
    if (l1Cache && l1Cache.items) {
      l1Cache.items.forEach(item => {
        if (item.type === 'folder') {
          l1Directories.push({
            name: item.name,
            relativeUrl: item.relativeUrl
          });
        }
      });
    }

    // 执行共享搜索
    const results = performFuzzySearchInCache(query, l1Cache, subtreeCache);

    // 发送消息给当前的 Content Script 展示浮层
    try {
      chrome.tabs.sendMessage(tab.id, {
        action: 'show_search_results',
        query: query,
        results: results,
        l1Directories: l1Directories
      });
    } catch (err) {
      console.warn('Failed to send message to content script:', err);
    }
  }
});
