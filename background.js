// background.js
// 引入共享同步逻辑
importScripts('sync-helper.js');

const DAILY_UPDATE_ALARM_NAME = 'daily_update_alarm';
const LEGACY_WEEKDAY_ALARM_NAME = 'weekday_sync_alarm';
const LEGACY_WEEKLY_ALARM_NAME = 'sync_all_data';
const DAILY_SYNC_HOURS = [9, 13];

// 启动时清除任何遗留的正在同步状态，防止 Service Worker 重启或崩溃后状态卡在“同步中”
chrome.storage.local.get(null, (allData) => {
  const keysToRemove = Object.keys(allData).filter(key => key.startsWith('sync_status_') || key === 'sync_status');
  if (keysToRemove.length > 0) {
    chrome.storage.local.remove(keysToRemove, () => {
      console.log('[SharePoint Map] Cleared stale sync statuses on startup:', keysToRemove);
    });
  }
});

// Service Worker 启动时恢复扩展图标上的未读提醒数量。
updateFileUpdateBadge().catch(err => {
  console.warn('[SharePoint Map] Failed to initialize notification badge:', err);
});

// 计算下一个本地时间 09:00 或 13:00 的时间戳。
function getNextDailySyncTime(now) {
  const date = new Date(now);
  const todayCandidates = DAILY_SYNC_HOURS.map(hour => {
    const candidate = new Date(date);
    candidate.setHours(hour, 0, 0, 0);
    return candidate.getTime();
  }).filter(timestamp => timestamp > now);

  if (todayCandidates.length > 0) {
    return Math.min(...todayCandidates);
  }

  date.setDate(date.getDate() + 1);
  date.setHours(DAILY_SYNC_HOURS[0], 0, 0, 0);
  return date.getTime();
}

function scheduleNextDailyAlarm() {
  const nextTime = getNextDailySyncTime(Date.now());
  chrome.alarms.create(DAILY_UPDATE_ALARM_NAME, { when: nextTime });
  console.log('Scheduled next daily sync for:', new Date(nextTime).toString());
}

function getAutomaticSyncWindow(scheduledTime) {
  const scheduledDate = new Date(scheduledTime);
  const syncHour = scheduledDate.getHours() === 9 ? 9 : 13;
  const end = new Date(scheduledDate);
  end.setHours(syncHour, 0, 0, 0);

  const start = new Date(end);
  if (syncHour === 9) {
    start.setDate(start.getDate() - 1);
    start.setHours(13, 0, 0, 0);
  } else {
    start.setHours(9, 0, 0, 0);
  }

  return {
    modifiedAfter: start.getTime(),
    modifiedBefore: end.getTime(),
    checkpointTime: end.getTime(),
    label: `${syncHour}:00`
  };
}

function isSupportedDailySyncAlarm(alarm) {
  if (!alarm || !Number.isFinite(alarm.scheduledTime)) return false;
  return DAILY_SYNC_HOURS.includes(new Date(alarm.scheduledTime).getHours());
}

// 检查并确保每天 09:00/13:00 的 Alarm 已设置。
chrome.alarms.get(DAILY_UPDATE_ALARM_NAME, (alarm) => {
  if (!isSupportedDailySyncAlarm(alarm)) {
    if (alarm) {
      chrome.alarms.clear(DAILY_UPDATE_ALARM_NAME, () => {
        scheduleNextDailyAlarm();
      });
      return;
    }
    scheduleNextDailyAlarm();
  } else {
    console.log('Daily sync alarm already scheduled for:', new Date(alarm.scheduledTime).toString());
  }
});

// 清理旧版本的工作日 Alarm，避免同一扩展产生两次同步。
chrome.alarms.get(LEGACY_WEEKDAY_ALARM_NAME, (alarm) => {
  if (alarm) {
    chrome.alarms.clear(LEGACY_WEEKDAY_ALARM_NAME);
  }
});

// 清理旧版本每周全量同步 Alarm，避免绕过新的时间窗口逻辑。
chrome.alarms.get(LEGACY_WEEKLY_ALARM_NAME, (alarm) => {
  if (alarm) {
    chrome.alarms.clear(LEGACY_WEEKLY_ALARM_NAME);
  }
});

// 监听安装事件，设置定时任务与右键菜单
chrome.runtime.onInstalled.addListener(() => {
  console.log('SharePoint Quick Access extension installed.');
  // 设置每天 09:00/13:00 定时同步并生成文件更新提醒
  scheduleNextDailyAlarm();

  // 创建右键菜单以支持选中文本在网页中搜索
  chrome.contextMenus.create({
    id: "search_sp_map",
    title: "在 SharePoint Map 中搜索 '%s'",
    contexts: ["selection"]
  });
});

// 监听 Alarm 触发
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === DAILY_UPDATE_ALARM_NAME) {
    const syncWindow = getAutomaticSyncWindow(alarm.scheduledTime || Date.now());
    console.log(`Daily ${syncWindow.label} sync alarm triggered. Syncing favorited directories and checking configured file updates from ${new Date(syncWindow.modifiedAfter).toISOString()} to ${new Date(syncWindow.modifiedBefore).toISOString()}...`);
    performAllSync({
      mode: 'automatic',
      notifyUpdates: true,
      modifiedAfter: syncWindow.modifiedAfter,
      modifiedBefore: syncWindow.modifiedBefore,
      overlapMs: 0,
      automaticCheckpointTime: syncWindow.checkpointTime
    })
      .then(() => {
        console.log(`Daily ${syncWindow.label} sync and update check completed successfully.`);
      })
      .catch((err) => {
        console.error(`Daily ${syncWindow.label} sync failed:`, err);
      })
      .finally(() => {
        // 无论成功还是失败，都安排下一次的每日同步。
        scheduleNextDailyAlarm();
      });
  } else if (alarm.name === LEGACY_WEEKLY_ALARM_NAME) {
    // 兼容已存在的旧 Alarm：启动时会清理，这里不再执行旧的同步逻辑。
    chrome.alarms.clear(LEGACY_WEEKLY_ALARM_NAME);
  }
});

// 消息监听保留，以备将来需要
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'sync_all') {
    performAllSync({ mode: 'manual', configId: request.configId || '' })
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
    syncSubtree(folderId, relativeUrl, { mode: 'manual', notifyUpdates: false })
      .then((nodeCount) => sendResponse({ success: true, count: nodeCount }))
      .catch((err) => sendResponse({ success: false, error: err.message || err }));
    return true;
  }
  if (request.action === 'get_sync_logs') {
    chrome.storage.local.get('sync_logs')
      .then((data) => sendResponse({ success: true, logs: data.sync_logs || [] }))
      .catch((err) => sendResponse({ success: false, error: err.message || String(err) }));
    return true;
  }
  if (request.action === 'mark_update_notifications_read') {
    markFileUpdateNotificationsRead(request.ids)
      .then((count) => sendResponse({ success: true, count }))
      .catch((err) => sendResponse({ success: false, error: err.message || err }));
    return true;
  }
  if (request.action === 'mark_all_update_notifications_read') {
    markAllFileUpdateNotificationsRead(request.configId || '')
      .then((count) => sendResponse({ success: true, count }))
      .catch((err) => sendResponse({ success: false, error: err.message || err }));
    return true;
  }
});

// 其他扩展页面直接修改提醒状态时，也及时同步徽标。
chrome.storage.onChanged.addListener((changes, namespace) => {
  if (namespace === 'local' && changes[FILE_UPDATE_NOTIFICATIONS_KEY]) {
    updateFileUpdateBadge();
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
