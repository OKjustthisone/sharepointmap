// background.js
// 引入共享同步逻辑
importScripts('sync-helper.js');

// 监听安装事件，设置 7 天定时任务
chrome.runtime.onInstalled.addListener(() => {
  console.log('SharePoint Quick Access extension installed.');
  // 设置 7 天定时任务 (7 * 24 * 60 分钟)
  chrome.alarms.create('sync_all_data', { periodInMinutes: 7 * 24 * 60 });
});

// 监听 Alarm 触发
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === 'sync_all_data') {
    console.log('Scheduled alarm triggered. Syncing all SharePoint data...');
    // 在后台静默运行定时更新（如果 Cookie 丢失则会自动被捕获并忽略，等用户打开 Popup 时会自动基于 UI 线程 Cookie 重新更新）
    performAllSync();
  }
});

// 消息监听保留，以备将来需要
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'sync_level1') {
    syncLevel1()
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
