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

// 监听安装事件，设置定时任务
chrome.runtime.onInstalled.addListener(() => {
  console.log('SharePoint Quick Access extension installed.');
  // 设置 7 天定时任务 (7 * 24 * 60 分钟)
  chrome.alarms.create('sync_all_data', { periodInMinutes: 7 * 24 * 60 });
  // 设置工作日 10 点定时同步
  scheduleNextWeekdayAlarm();
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
