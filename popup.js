// popup.js

document.addEventListener('DOMContentLoaded', async () => {
  // DOM 元素获取
  const settingsBtn = document.getElementById('settingsBtn');
  const syncL1Btn = document.getElementById('syncL1Btn');
  const alertBanner = document.getElementById('alertBanner');
  const alertText = document.getElementById('alertText');
  const alertActionBtn = document.getElementById('alertActionBtn');
  const searchInput = document.getElementById('searchInput');
  const clearSearchBtn = document.getElementById('clearSearchBtn');
  const searchResultsSection = document.getElementById('searchResultsSection');
  const searchResultsList = document.getElementById('searchResultsList');
  const defaultViews = document.getElementById('defaultViews');
  const favoritesList = document.getElementById('favoritesList');
  const directoryTree = document.getElementById('directoryTree');
  const directoryToggleBtn = document.getElementById('directoryToggleBtn');
  const directoryToggleArrow = document.getElementById('directoryToggleArrow');
  const syncTimeSpan = document.getElementById('syncTime');
  const toastContainer = document.getElementById('toastContainer');
  const l1FilterSelect = document.getElementById('l1FilterSelect');
  const filterChips = document.querySelectorAll('.filter-chip');
  const tabsContainer = document.getElementById('tabsContainer');
  const notificationsBtn = document.getElementById('notificationsBtn');
  const notificationBellDot = document.getElementById('notificationBellDot');
  const updateNotificationsSection = document.getElementById('updateNotificationsSection');
  const updateNotificationsList = document.getElementById('updateNotificationsList');
  const updateNotificationCount = document.getElementById('updateNotificationCount');
  const markNotificationsReadBtn = document.getElementById('markNotificationsReadBtn');

  // 全局数据状态缓存与过滤器状态
  let activeFilter = 'all';
  let activeL1Path = 'all';
  let spConfig = null;
  let spConfigs = [];
  let currentConfigId = '';
  let favorites = [];
  let l1Cache = null;
  let subtreeCache = {};
  let syncStatus = {};
  let updateNotifications = [];
  let isNotificationsPanelOpen = false;
  let activeNotificationL1Path = 'all';

  function svgIcon(name, className = '') {
    const extraClass = className ? ` ${className}` : '';
    return `<svg class="svg-icon${extraClass}" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><use href="#icon-${name}"></use></svg>`;
  }

  function updateDirectoryToggleIndicator(isExpanded) {
    const arrowIcon = directoryToggleArrow.querySelector('.toggle-arrow-svg');
    const arrowLabel = directoryToggleArrow.querySelector('.toggle-arrow-label');
    if (arrowIcon) arrowIcon.classList.toggle('expanded', isExpanded);
    if (arrowLabel) arrowLabel.innerText = isExpanded ? '收起目录' : '展开浏览';
    directoryToggleArrow.style.color = isExpanded ? 'var(--primary-cyan)' : 'var(--text-muted)';
  }
  
  // 树状图折叠展开状态映射 (folderId -> boolean)
  let expandedState = {};
  let isTreeExpanded = false; // 全部目录折叠展开状态

  // 1. 初始化检查配置
  await initApp();

  // 2. 绑定页面通用交互事件
  settingsBtn.addEventListener('click', () => chrome.runtime.openOptionsPage());
  alertActionBtn.addEventListener('click', () => chrome.runtime.openOptionsPage());

  notificationsBtn.addEventListener('click', () => {
    const visibleNotifications = getVisibleUpdateNotifications();
    if (visibleNotifications.length === 0) {
      showToast('暂无文件更新提醒', 'bell');
      return;
    }

    isNotificationsPanelOpen = !isNotificationsPanelOpen;
    renderUpdateNotifications();
  });

  document.getElementById('notificationL1Filter')?.addEventListener('change', (event) => {
    activeNotificationL1Path = event.target.value;
    renderUpdateNotifications();
  });

  markNotificationsReadBtn.addEventListener('click', async () => {
    const visibleNotifications = getVisibleUpdateNotifications();
    const unreadIds = visibleNotifications.filter(item => !item.read).map(item => item.id);
    if (unreadIds.length === 0) return;

    try {
      await markFileUpdateNotificationsRead(unreadIds);
      await loadUpdateNotificationsFromStorage();
      renderUpdateNotifications();
    } catch (err) {
      console.error('Failed to mark update notifications as read:', err);
      showToast('更新已读状态失败，请重试', 'alert');
    }
  });
  
  // 3. 绑定“全部目录”展开/折叠事件
  directoryToggleBtn.addEventListener('click', () => {
    isTreeExpanded = !isTreeExpanded;
    saveUIState(); // 保存状态
    if (isTreeExpanded) {
      directoryTree.classList.remove('collapsed');
    } else {
      directoryTree.classList.add('collapsed');
    }
    updateDirectoryToggleIndicator(isTreeExpanded);
  });
  
  // 手动同步 1 级目录事件
  syncL1Btn.addEventListener('click', (e) => {
    e.stopPropagation();
    syncL1Btn.classList.add('loading');
    showToast('已在后台启动 1 级目录同步...', 'sync');
    chrome.runtime.sendMessage({ action: 'sync_level1', configId: currentConfigId }, (response) => {
      syncL1Btn.classList.remove('loading');
      if (chrome.runtime.lastError) {
        console.error('Background sync level 1 failed:', chrome.runtime.lastError);
        showToast('同步 1 级目录异常: ' + chrome.runtime.lastError.message, 'alert', true);
      } else if (response && !response.success) {
        console.error('Background sync level 1 returned error:', response.error);
        showToast(`同步 1 级目录失败: ${response.error}`, 'alert', true);
      } else {
        showToast('1 级目录同步完成！', 'check');
      }
    });
  });

  // 搜索过滤芯片事件绑定

  filterChips.forEach(chip => {
    chip.addEventListener('click', () => {
      filterChips.forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      activeFilter = chip.getAttribute('data-filter');
      
      saveUIState(); // 保存状态

      const query = searchInput.value.trim().toLowerCase();
      if (query) {
        performSearch(query);
      }
    });
  });

  // 1级目录下拉筛选事件
  l1FilterSelect.addEventListener('change', () => {
    activeL1Path = l1FilterSelect.value;
    if (activeL1Path !== 'all') {
      l1FilterSelect.classList.add('active');
    } else {
      l1FilterSelect.classList.remove('active');
    }
    
    saveUIState(); // 保存状态

    const query = searchInput.value.trim().toLowerCase();
    if (query) {
      performSearch(query);
    }
  });

  // 搜索输入过滤
  searchInput.addEventListener('input', () => {
    const query = searchInput.value.trim().toLowerCase();
    saveUIState(); // 保存状态
    if (query) {
      clearSearchBtn.classList.remove('hide');
      performSearch(query);
    } else {
      clearSearchBtn.classList.add('hide');
      searchResultsSection.classList.add('hide');
      defaultViews.classList.remove('hide');
    }
  });

  // 清除搜索
  clearSearchBtn.addEventListener('click', () => {
    searchInput.value = '';
    clearSearchBtn.classList.add('hide');
    searchResultsSection.classList.add('hide');
    defaultViews.classList.remove('hide');
    
    // 重置过滤器
    activeFilter = 'all';
    activeL1Path = 'all';
    l1FilterSelect.value = 'all';
    l1FilterSelect.classList.remove('active');
    filterChips.forEach(c => c.classList.remove('active'));
    const allChip = document.querySelector('[data-filter="all"]');
    if (allChip) allChip.classList.add('active');
    
    saveUIState(); // 保存状态

    searchInput.focus();
  });

  // ==================== 初始化与数据加载 ====================

  async function initApp() {
    // A. 运行配置迁移（如果需要）
    await migrateConfigsIfNeeded();

    // B. 获取配置和当前激活的配置 ID
    const configData = await chrome.storage.local.get(['sp_configs', 'current_config_id', 'sp_config', 'ui_state']);
    spConfigs = configData.sp_configs || [];
    currentConfigId = configData.current_config_id || '';
    
    if (!currentConfigId && configData.sp_config) {
      spConfig = configData.sp_config;
    } else {
      spConfig = spConfigs.find(c => c.id === currentConfigId) || null;
    }

    // C. 根据激活的配置加载对应的 favorites 和 l1_cache
    const cacheKeys = [];
    if (currentConfigId) {
      cacheKeys.push(`favorites_${currentConfigId}`);
      cacheKeys.push(`l1_cache_${currentConfigId}`);
    } else {
      cacheKeys.push('favorites');
      cacheKeys.push('l1_cache');
    }

    const data = await chrome.storage.local.get(cacheKeys);
    
    if (currentConfigId) {
      favorites = data[`favorites_${currentConfigId}`] || [];
      l1Cache = data[`l1_cache_${currentConfigId}`];
    } else {
      favorites = data.favorites || [];
      l1Cache = data.l1_cache;
    }

    syncStatus = await loadSyncStatusFromStorage(favorites);
    subtreeCache = await loadSubtreeCacheFromStorage(favorites);
    await loadUpdateNotificationsFromStorage();
    populateNotificationL1Filter();
    
    // 渲染切换 Tab 栏
    renderTabs();
    renderUpdateNotifications();

    // 恢复 UI 状态变量
    const uiState = configData.ui_state || {};
    expandedState = uiState.expandedState || {};
    isTreeExpanded = uiState.isTreeExpanded || false;
    activeFilter = uiState.activeFilter || 'all';
    activeL1Path = uiState.activeL1Path || 'all';

    if (!spConfig || !spConfig.siteUrl || !spConfig.libraryName) {
      showAlert('请先配置您的 SharePoint 站点与文档库。', 'alert');
      directoryTree.innerHTML = `
        <div class="empty-list-placeholder">
          配置未设置。请点击右上方设置按钮进入设置页面配置站点。
        </div>
      `;
      tabsContainer.classList.add('hide');
      return;
    }

    hideAlert();

    // 恢复全部目录的 UI 展开/收起状态
    if (isTreeExpanded) {
      directoryTree.classList.remove('collapsed');
    } else {
      directoryTree.classList.add('collapsed');
    }
    updateDirectoryToggleIndicator(isTreeExpanded);
    
    // 恢复过滤器 Chip 的 active 状态
    filterChips.forEach(chip => {
      if (chip.getAttribute('data-filter') === activeFilter) {
        chip.classList.add('active');
      } else {
        chip.classList.remove('active');
      }
    });

    // 如果本地没有缓存，提示并自动尝试首次同步
    if (!l1Cache) {
      directoryTree.innerHTML = `
        <div class="loading-spinner-wrapper">
          <div class="spinner"></div>
          <span>首次使用，正在拉取 1 级目录...</span>
        </div>
      `;
      try {
        await syncLevel1(currentConfigId);
        showToast('1 级目录同步成功！', 'check');
        await loadDataFromStorage();
        populateL1FilterDropdown();
        renderFavorites();
        renderDirectoryTree();
        updateFavoritesSyncProgressDisplay();
      } catch (err) {
        console.error(err);
        directoryTree.innerHTML = `
          <div class="empty-list-placeholder" style="color: var(--danger-red);">
            同步失败，请点击右上角同步按钮手动重试。<br>
            错误原因: ${err.message || err}
          </div>
        `;
      }
    } else {
      // 检查是否已过期 (7 天)
      const isExpired = Date.now() - l1Cache.last_updated > 7 * 24 * 60 * 60 * 1000;
      if (isExpired) {
        showToast('正在自动更新已过期的缓存...', 'sync');
        syncLevel1(currentConfigId)
          .then(async () => {
            console.log('Auto refresh of Level 1 completed.');
            await loadDataFromStorage();
            populateL1FilterDropdown();
            renderFavorites();
            renderDirectoryTree();
            updateFavoritesSyncProgressDisplay();
            const query = uiState.searchQuery || '';
            if (query) {
              performSearch(query);
            }
          })
          .catch(err => console.warn('Auto refresh failed:', err));
      }
      
      updateSyncTimeDisplay();
      populateL1FilterDropdown();
      renderFavorites();
      renderDirectoryTree();
      updateFavoritesSyncProgressDisplay();

      // 恢复搜索输入与触发搜索
      const query = uiState.searchQuery || '';
      if (query) {
        searchInput.value = query;
        clearSearchBtn.classList.remove('hide');
        performSearch(query);
      }
    }
  }

  function renderTabs() {
    if (spConfigs.length <= 1) {
      tabsContainer.classList.add('hide');
      return;
    }

    tabsContainer.classList.remove('hide');
    tabsContainer.innerHTML = '';

    spConfigs.forEach(config => {
      const btn = document.createElement('button');
      btn.className = `tab-btn ${config.id === currentConfigId ? 'active' : ''}`;
      btn.innerText = config.name || '未命名站点';
      btn.title = `${config.siteUrl} (${config.libraryName})`;
      btn.addEventListener('click', async () => {
        if (config.id === currentConfigId) return;
        currentConfigId = config.id;
        await chrome.storage.local.set({ current_config_id: config.id });
        
        // 切换配置时清除界面部分状态
        expandedState = {};
        isTreeExpanded = false;
        isNotificationsPanelOpen = false;
        activeNotificationL1Path = 'all';
        
        // 重新初始化并加载新站点的数据
        await initApp();
      });
      tabsContainer.appendChild(btn);
    });
  }

  // 辅助函数：从本地存储中聚合所有一级目录的子树缓存
  async function loadSubtreeCacheFromStorage(favoriteItems = []) {
    const l1FolderIds = new Set(
      favoriteItems
        .filter(item => item.type === 'folder' && item.level === 1)
        .map(item => item.id)
    );
    if (l1FolderIds.size === 0) return {};

    const cacheKeys = Array.from(l1FolderIds, id => `subtree_cache_${id}`);
    const allData = await chrome.storage.local.get(cacheKeys);
    const subtreeCache = {};
    Object.keys(allData).forEach(key => {
      if (key.startsWith('subtree_cache_')) {
        const l1Id = key.substring('subtree_cache_'.length);
        subtreeCache[l1Id] = allData[key];
      }
    });
    return subtreeCache;
  }

  // 辅助函数：从本地存储中聚合所有一级目录的同步状态
  async function loadSyncStatusFromStorage(favoriteItems = []) {
    const folderIds = favoriteItems
      .filter(item => item.type === 'folder' && item.level === 1)
      .map(item => item.id);
    if (folderIds.length === 0) return {};

    const statusKeys = folderIds.map(id => `sync_status_${id}`);
    const allData = await chrome.storage.local.get(statusKeys);
    const syncStatus = {};
    Object.keys(allData).forEach(key => {
      if (key.startsWith('sync_status_')) {
        const folderId = key.substring('sync_status_'.length);
        syncStatus[folderId] = allData[key];
      }
    });
    return syncStatus;
  }

  async function loadUpdateNotificationsFromStorage() {
    const data = await chrome.storage.local.get(FILE_UPDATE_NOTIFICATIONS_KEY);
    const notifications = Array.isArray(data[FILE_UPDATE_NOTIFICATIONS_KEY])
      ? data[FILE_UPDATE_NOTIFICATIONS_KEY]
      : [];
    const cutoff = Date.now() - 30 * 24 * 60 * 60 * 1000;
    updateNotifications = notifications
      .filter(item => Number(item?.detectedAt) >= cutoff)
      .sort((a, b) => (b.detectedAt || 0) - (a.detectedAt || 0));

    if (updateNotifications.length !== notifications.length) {
      await chrome.storage.local.set({ file_update_notifications: updateNotifications });
    }
  }

  function getVisibleUpdateNotifications() {
    const configId = currentConfigId || 'legacy';
    return updateNotifications.filter(item => {
      if ((currentConfigId || spConfig) && item.configId !== configId) return false;
      if (activeNotificationL1Path === 'all') return true;
      return getNotificationFirstLevelDirectory(item, getNotificationConfig(item)) === activeNotificationL1Path;
    });
  }

  function populateNotificationL1Filter() {
    const select = document.getElementById('notificationL1Filter');
    if (!select) return;

    const notifications = updateNotifications.filter(item => {
      const configId = currentConfigId || 'legacy';
      return (!(currentConfigId || spConfig) || item.configId === configId);
    });
    const names = new Map();
    notifications.forEach(item => {
      const config = getNotificationConfig(item);
      const name = getNotificationFirstLevelDirectory(item, config);
      names.set(name, name);
    });

    select.innerHTML = '<option value="all">所有一级目录</option>';
    Array.from(names.values())
      .sort((a, b) => a.localeCompare(b, 'zh-CN', { numeric: true }))
      .forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        select.appendChild(option);
      });

    if (!names.has(activeNotificationL1Path)) activeNotificationL1Path = 'all';
    select.value = activeNotificationL1Path;
  }

  function formatUpdateNotificationTime(notification) {
    const rawTime = notification.eventType === 'uploaded'
      ? (notification.createdAt || notification.modifiedAt || notification.detectedAt)
      : (notification.modifiedAt || notification.detectedAt);
    const date = new Date(rawTime);
    if (Number.isNaN(date.getTime())) return '时间未知';
    return date.toLocaleString('zh-CN', {
      month: 'numeric',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  }

  function decodeNotificationPathPart(value) {
    try {
      return decodeURIComponent(value);
    } catch (err) {
      return value;
    }
  }

  function getNotificationPathSegments(value) {
    return String(value || '')
      .split('/')
      .filter(Boolean)
      .map(decodeNotificationPathPart);
  }

  function getNotificationConfig(notification) {
    return spConfigs.find(config => config.id === notification?.configId) || spConfig || null;
  }

  function getNotificationSiteName(notification, config) {
    if (config?.siteUrl) {
      try {
        const siteUrl = new URL(config.siteUrl);
        const sitePath = getNotificationPathSegments(siteUrl.pathname);
        return sitePath[sitePath.length - 1] || siteUrl.hostname;
      } catch (err) {
        // 继续从文件的 server-relative 路径中解析站点名称。
      }
    }

    const relativePath = getNotificationPathSegments(notification?.relativeUrl);
    const siteMarkerIndex = relativePath.findIndex(part => ['sites', 'teams'].includes(part.toLowerCase()));
    return relativePath[siteMarkerIndex + 1] || config?.name || 'SharePoint';
  }

  function findNotificationPathSequence(parts, sequence) {
    if (sequence.length === 0) return -1;

    for (let index = 0; index <= parts.length - sequence.length; index += 1) {
      const isMatch = sequence.every((part, offset) => (
        parts[index + offset].toLowerCase() === part.toLowerCase()
      ));
      if (isMatch) return index;
    }

    return -1;
  }

  function getNotificationFirstLevelDirectory(notification, config) {
    const pathParts = getNotificationPathSegments(notification?.relativeUrl);
    if (pathParts.length < 2) return '文档库根目录';

    const parentParts = pathParts.slice(0, -1);
    const libraryName = decodeNotificationPathPart(String(config?.libraryName || '').trim()).toLowerCase();
    let libraryIndex = libraryName
      ? parentParts.findIndex(part => part.toLowerCase() === libraryName)
      : -1;

    if (libraryIndex < 0 && config?.siteUrl) {
      try {
        const sitePath = getNotificationPathSegments(new URL(config.siteUrl).pathname);
        const siteIndex = findNotificationPathSequence(parentParts, sitePath);
        if (siteIndex >= 0) libraryIndex = siteIndex + sitePath.length;
      } catch (err) {
        // 使用下面的 SharePoint 默认路径规则兜底。
      }
    }

    if (libraryIndex < 0) {
      const knownLibraries = ['shared documents', 'documents', '共享文档', '文档'];
      libraryIndex = parentParts.findIndex(part => knownLibraries.includes(part.toLowerCase()));
    }

    if (libraryIndex < 0) {
      const siteMarkerIndex = parentParts.findIndex(part => ['sites', 'teams'].includes(part.toLowerCase()));
      libraryIndex = siteMarkerIndex >= 0 ? siteMarkerIndex + 2 : 0;
    }

    return parentParts[libraryIndex + 1] || '文档库根目录';
  }

  function getNotificationLocation(notification) {
    const config = getNotificationConfig(notification);
    const siteName = getNotificationSiteName(notification, config);
    const firstLevelDirectory = getNotificationFirstLevelDirectory(notification, config);
    return `${siteName}/${firstLevelDirectory}`;
  }

  function updateNotificationBell(visibleNotifications) {
    const unreadCount = visibleNotifications.filter(item => !item.read).length;
    const hasUnread = unreadCount > 0;

    notificationBellDot.classList.toggle('hide', !hasUnread);
    notificationBellDot.innerText = hasUnread
      ? (unreadCount > 99 ? '99+' : String(unreadCount))
      : '';
    notificationsBtn.classList.toggle('has-unread', hasUnread);
    notificationsBtn.title = hasUnread ? `${unreadCount} 条未读文件更新` : '查看文件更新提醒';
    notificationsBtn.setAttribute('aria-label', notificationsBtn.title);
    notificationsBtn.setAttribute('aria-expanded', String(isNotificationsPanelOpen));
  }

  function renderUpdateNotifications() {
    if (!updateNotificationsSection || !updateNotificationsList) return;

    const visibleNotifications = getVisibleUpdateNotifications();
    updateNotificationBell(visibleNotifications);
    updateNotificationsList.innerHTML = '';

    if (visibleNotifications.length === 0) {
      updateNotificationCount.innerText = '';
      markNotificationsReadBtn.disabled = true;
      updateNotificationsSection.classList.add('hide');
      return;
    }

    if (!isNotificationsPanelOpen) {
      updateNotificationsSection.classList.add('hide');
      return;
    }

    updateNotificationsSection.classList.remove('hide');
    const unreadCount = visibleNotifications.filter(item => !item.read).length;
    updateNotificationCount.innerText = unreadCount > 0
      ? `${unreadCount} 条待查看`
      : `${visibleNotifications.length} 条记录`;
    markNotificationsReadBtn.disabled = unreadCount === 0;

    visibleNotifications.slice(0, 30).forEach(notification => {
      const card = document.createElement('article');
      card.className = `update-notification-card ${notification.read ? 'is-read' : 'is-unread'} ${notification.eventType}`;

      const icon = document.createElement('div');
      icon.className = 'update-notification-icon';
      icon.innerHTML = svgIcon('file-presentation', 'update-notification-svg');

      const body = document.createElement('div');
      body.className = 'update-notification-body';

      const titleRow = document.createElement('div');
      titleRow.className = 'update-notification-title-row';

      const kind = document.createElement('span');
      kind.className = `update-notification-kind ${notification.eventType}`;
      kind.innerText = notification.eventType === 'uploaded' ? '新上传' : '已修改';
      titleRow.appendChild(kind);

      const fileLink = document.createElement('a');
      fileLink.className = 'update-notification-file-link';
      fileLink.innerText = notification.name || '未命名文件';
      fileLink.title = '点击直接打开文件';
      fileLink.href = getOnlineViewUrl(notification.webUrl);
      fileLink.target = '_blank';
      fileLink.rel = 'noopener noreferrer';
      fileLink.addEventListener('click', () => {
        if (!notification.read) {
          markFileUpdateNotificationsRead([notification.id]).catch(err => {
            console.warn('Failed to mark clicked notification as read:', err);
          });
        }
      });
      titleRow.appendChild(fileLink);
      body.appendChild(titleRow);

      const description = document.createElement('div');
      description.className = 'update-notification-description';
      description.innerText = `${notification.eventType === 'uploaded' ? '已上传到' : '最后修改于'} ${formatUpdateNotificationTime(notification)}`;
      body.appendChild(description);

      const path = document.createElement('div');
      path.className = 'update-notification-path';
      path.title = notification.relativeUrl || '';
      path.innerHTML = svgIcon('folder', 'path-svg');
      const pathLabel = document.createElement('span');
      pathLabel.innerText = getNotificationLocation(notification);
      path.appendChild(pathLabel);
      body.appendChild(path);

      const actions = document.createElement('div');
      actions.className = 'update-notification-actions';

      const openLink = document.createElement('a');
      openLink.className = 'update-notification-open';
      openLink.innerHTML = `${svgIcon('external-link', 'inline-action-svg')}<span>打开文件</span>`;
      openLink.href = getOnlineViewUrl(notification.webUrl);
      openLink.target = '_blank';
      openLink.rel = 'noopener noreferrer';
      openLink.addEventListener('click', () => {
        if (!notification.read) {
          markFileUpdateNotificationsRead([notification.id]).catch(err => {
            console.warn('Failed to mark clicked notification as read:', err);
          });
        }
      });
      actions.appendChild(openLink);

      if (!notification.read) {
        const readBtn = document.createElement('button');
        readBtn.className = 'update-notification-read-btn';
        readBtn.type = 'button';
        readBtn.title = '标记为已读';
        readBtn.setAttribute('aria-label', '标记为已读');
        readBtn.innerHTML = svgIcon('check', 'read-svg');
        readBtn.addEventListener('click', async () => {
          await markFileUpdateNotificationsRead([notification.id]);
          await loadUpdateNotificationsFromStorage();
          renderUpdateNotifications();
        });
        actions.appendChild(readBtn);
      }

      card.appendChild(icon);
      card.appendChild(body);
      card.appendChild(actions);
      updateNotificationsList.appendChild(card);
    });
  }

    // 重新从 storage 读取最新数据
  async function loadDataFromStorage() {
    const configData = await chrome.storage.local.get(['sp_configs', 'current_config_id', 'sp_config']);
    spConfigs = configData.sp_configs || [];
    currentConfigId = configData.current_config_id || '';
    if (currentConfigId) {
      spConfig = spConfigs.find(c => c.id === currentConfigId) || null;
    } else {
      spConfig = configData.sp_config;
    }

    const cacheKeys = [];
    if (currentConfigId) {
      cacheKeys.push(`favorites_${currentConfigId}`);
      cacheKeys.push(`l1_cache_${currentConfigId}`);
    } else {
      cacheKeys.push('favorites');
      cacheKeys.push('l1_cache');
    }

    const data = await chrome.storage.local.get(cacheKeys);
    if (currentConfigId) {
      favorites = data[`favorites_${currentConfigId}`] || [];
      l1Cache = data[`l1_cache_${currentConfigId}`];
    } else {
      favorites = data.favorites || [];
      l1Cache = data.l1_cache;
    }

    syncStatus = await loadSyncStatusFromStorage(favorites);
    subtreeCache = await loadSubtreeCacheFromStorage(favorites);
    await loadUpdateNotificationsFromStorage();
    populateNotificationL1Filter();
    
    renderTabs();
    updateSyncTimeDisplay();
    renderUpdateNotifications();
  }

  function updateSyncTimeDisplay() {
    if (l1Cache && l1Cache.last_updated) {
      const date = new Date(l1Cache.last_updated);
      syncTimeSpan.innerText = `上次同步: ${date.toLocaleDateString()} ${date.toTimeString().substring(0, 5)}`;
    } else {
      syncTimeSpan.innerText = '上次同步: --';
    }
  }

  // ==================== 渲染收藏夹 ====================

  function renderFavorites() {
    const favCountSpan = document.getElementById('favCount');
    if (favCountSpan) {
      favCountSpan.innerText = favorites.length > 0 ? `(${favorites.length})` : '';
    }

    favoritesList.innerHTML = '';
    if (favorites.length === 0) {
      favoritesList.className = 'empty-list-placeholder';
      favoritesList.innerText = '暂无收藏。点击目录树中文件夹或文件旁的收藏按钮即可加入收藏。';
      return;
    }

    favoritesList.className = '';
    
    // 按名称排序：文件夹优先，然后按字母/数字自然排序
    const sortedFavorites = [...favorites].sort((a, b) => {
      if (a.type !== b.type) {
        return a.type === 'folder' ? -1 : 1;
      }
      return a.name.localeCompare(b.name, 'zh-CN', { numeric: true });
    });

    sortedFavorites.forEach(fav => {
      const favNode = createTreeNodeElement(fav, 1, true);
      favoritesList.appendChild(favNode);
    });
  }

  // ==================== 渲染目录树 ====================

  function renderDirectoryTree() {
    directoryTree.innerHTML = '';
    if (!l1Cache || !l1Cache.items || l1Cache.items.length === 0) {
      directoryTree.innerHTML = '<div class="empty-list-placeholder">没有找到 1 级目录，请刷新。</div>';
      return;
    }

    // 渲染 1 级目录
    l1Cache.items.forEach(item => {
      const nodeEl = createTreeNodeElement(item, 1);
      directoryTree.appendChild(nodeEl);
    });
  }

  // 创建树节点 DOM 元素
  function createTreeNodeElement(item, depth, isFavList = false) {
    const container = document.createElement('div');
    container.className = 'tree-node-wrapper';

    const isFolder = item.type === 'folder';
    const isExpanded = expandedState[item.id] || false;
    const isFav = favorites.some(fav => fav.id === item.id);

    // 判断该文件夹是否已在缓存中 (如果是 1 级收藏文件夹，或者其父辈已被收藏)
    const hasCache = getCachedSubtree(item.id, item.relativeUrl);

    // 折叠展开箭头状态
    let toggleClass = 'node-toggle';
    if (!isFolder) {
      toggleClass += ' empty';
    } else if (isExpanded) {
      toggleClass += ' expanded';
    }

    const nodeEl = document.createElement('div');
    nodeEl.className = 'tree-node';
    nodeEl.style.paddingLeft = `${(depth - 1) * 16 + 10}px`;

    const iconName = isFolder ? (isExpanded ? 'folder-open' : 'folder') : 'file';
    const icon = svgIcon(iconName, 'node-svg-icon');

    let syncBtnHtml = '';
    const isNodeSyncing = syncStatus[item.id] && syncStatus[item.id].status === 'syncing';
    if (isFolder && item.level === 1 && isFav) {
      const loadingClass = isNodeSyncing ? 'loading' : '';
      const syncTitle = isNodeSyncing ? '正在同步子树...' : '同步该目录下的子树';
      syncBtnHtml = `<button class="action-btn sync-btn ${loadingClass}" title="${syncTitle}" aria-label="${syncTitle}">${svgIcon('sync', 'action-svg')}</button>`;
    }

    let favBtnHtml = '';
    if (!isFavList) {
      const favTitle = isFav ? '取消快捷收藏' : '加入快捷收藏';
      favBtnHtml = `<button class="fav-btn ${isFav ? 'active' : ''}" title="${favTitle}" aria-label="${favTitle}">${svgIcon(isFav ? 'star-filled' : 'star', 'fav-svg')}</button>`;
    }

    nodeEl.innerHTML = `
      <div class="node-left">
        <span class="${toggleClass}">${svgIcon('chevron-right', 'node-toggle-svg')}</span>
        <span class="node-icon ${isFolder ? 'folder-icon' : 'file-icon'}">${icon}</span>
        <span class="node-name ${isFolder ? 'folder-node' : ''}" title="${item.relativeUrl}">${item.name}</span>
      </div>
      <div class="node-right ${isFav ? 'is-fav' : ''}">
        ${syncBtnHtml}
        <button class="action-btn open-btn" data-url="${item.webUrl}" title="在浏览器中打开网页" aria-label="在浏览器中打开网页">${svgIcon('external-link', 'action-svg')}</button>
        <button class="action-btn copy-btn" data-url="${item.webUrl}" title="复制 SharePoint 链接" aria-label="复制 SharePoint 链接">${svgIcon('clipboard', 'action-svg')}</button>
        ${favBtnHtml}
      </div>
    `;

    container.appendChild(nodeEl);

    // 子树容器 (用于放置折叠子目录)
    const childrenContainer = document.createElement('div');
    childrenContainer.className = 'children-container';
    if (!isExpanded) {
      childrenContainer.style.display = 'none';
    }
    container.appendChild(childrenContainer);

    // 1. 打开与复制事件
    nodeEl.querySelector('.open-btn').addEventListener('click', (e) => {
      e.stopPropagation();
      handleOpen(item.webUrl);
    });
    nodeEl.querySelector('.copy-btn').addEventListener('click', (e) => {
      e.stopPropagation();
      handleCopy(item.webUrl);
    });

    // 2. 收藏按钮事件
    if (!isFavList) {
      nodeEl.querySelector('.fav-btn').addEventListener('click', (e) => {
        e.stopPropagation();
        toggleFavorite(item);
      });
    }

    // 2.5 同步子树按钮事件
    if (isFolder && item.level === 1 && isFav) {
      const syncBtn = nodeEl.querySelector('.sync-btn');
      if (syncBtn) {
        syncBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          syncBtn.classList.add('loading');
          
          showToast(`已在后台启动目录 [${item.name}] 的同步，可以关闭此窗口...`, 'sync');
          
          chrome.runtime.sendMessage({
            action: 'sync_subtree',
            folderId: item.id,
            relativeUrl: item.relativeUrl
          }, (response) => {
            syncBtn.classList.remove('loading');
            if (chrome.runtime.lastError) {
              console.error('Background sync subtree failed:', chrome.runtime.lastError);
              showToast('同步子目录异常: ' + chrome.runtime.lastError.message, 'alert', true);
            } else if (response && !response.success) {
              console.error('Background sync subtree returned error:', response.error);
              showToast(`同步子目录失败: ${response.error}`, 'alert', true);
            } else {
              showToast(`目录 [${item.name}] 同步完成！`, 'check');
            }
          });
        });
      }
    }

    // 3. 展开折叠事件 (如果是文件夹)
    if (isFolder) {
      const toggleNode = () => {
        const nextExpanded = !expandedState[item.id];
        expandedState[item.id] = nextExpanded;
        
        saveUIState(); // 保存状态
        
        // 刷新节点图标
        const arrow = nodeEl.querySelector('.node-toggle');
        const folderIcon = nodeEl.querySelector('.node-icon');
        if (nextExpanded) {
          arrow.classList.add('expanded');
          folderIcon.innerHTML = svgIcon('folder-open', 'node-svg-icon');
          childrenContainer.style.display = 'block';
          
          // 加载子项
          renderSubtreeItems(item, childrenContainer, depth + 1);
        } else {
          arrow.classList.remove('expanded');
          folderIcon.innerHTML = svgIcon('folder', 'node-svg-icon');
          childrenContainer.style.display = 'none';
        }
      };

      nodeEl.querySelector('.node-toggle').addEventListener('click', (e) => {
        e.stopPropagation();
        toggleNode();
      });
      nodeEl.querySelector('.node-left').addEventListener('click', toggleNode);

      // 如果初始化状态为展开，则自动渲染子目录
      if (isExpanded) {
        renderSubtreeItems(item, childrenContainer, depth + 1);
      }
    }

    return container;
  }

  // 渲染子目录项
  function renderSubtreeItems(parentItem, container, depth) {
    container.innerHTML = '';

    // 判断该父级文件夹的子树是否存在于缓存中
    const subItems = getCachedSubtree(parentItem.id, parentItem.relativeUrl);

    if (subItems) {
      // 存在缓存，进行子树渲染
      const folders = subItems.folders || [];
      const files = subItems.files || [];

      if (folders.length === 0 && files.length === 0) {
        const emptyEl = document.createElement('div');
        emptyEl.className = 'tree-node';
        emptyEl.style.paddingLeft = `${depth * 16 + 10}px`;
        emptyEl.innerHTML = `<span class="node-name uncached-alert">文件夹为空</span>`;
        container.appendChild(emptyEl);
        return;
      }

      // 优先显示子文件夹，再显示文件
      folders.forEach(sub => {
        container.appendChild(createTreeNodeElement({ ...sub, level: depth }, depth));
      });
      files.forEach(sub => {
        container.appendChild(createTreeNodeElement({ ...sub, level: depth }, depth));
      });

      // 如果是 1 级收藏的文件夹，在其展开列表的最下方展示全量统计汇总
      if (parentItem.level === 1) {
        let totalFolders = 0;
        let totalFiles = 0;
        const cacheRoot = subtreeCache[parentItem.id];
        if (cacheRoot && cacheRoot.tree) {
          Object.values(cacheRoot.tree).forEach(node => {
            totalFolders += (node.folders || []).length;
            totalFiles += (node.files || []).length;
          });
        }
        
        const summaryEl = document.createElement('div');
        summaryEl.className = 'tree-node-summary';
        summaryEl.style.paddingLeft = `${depth * 16 + 10}px`;
        summaryEl.innerHTML = `${svgIcon('grid', 'summary-svg')}<span>该目录下共含有 <strong>${totalFolders}</strong> 个文件夹，<strong>${totalFiles}</strong> 个文件</span>`;
        container.appendChild(summaryEl);
      }

    } else {
      // 不在缓存中：这只能是 1 级目录
      if (parentItem.level === 1) {
        const isSyncing = syncStatus[parentItem.id] && syncStatus[parentItem.id].status === 'syncing';
        const isFav = favorites.some(fav => fav.id === parentItem.id);
        const tipEl = document.createElement('div');
        tipEl.className = 'uncached-tip';
        tipEl.style.marginLeft = `${(depth - 1) * 16 + 10}px`;
        
        if (isSyncing) {
          tipEl.innerHTML = `
            <div class="sync-progress-wrapper" style="display: flex; align-items: center; gap: 8px;">
              <div class="spinner mini"></div>
              <span>正在后台同步中...</span>
            </div>
          `;
        } else if (isFav) {
          tipEl.innerHTML = `
            <span>该目录缓存未同步。请点击 <button class="inline-sync-btn">${svgIcon('sync', 'inline-action-svg')}<span>重新同步子目录</span></button> 尝试拉取。</span>
          `;
          tipEl.querySelector('.inline-sync-btn').addEventListener('click', () => {
            const btn = tipEl.querySelector('.inline-sync-btn');
            btn.disabled = true;
            btn.innerHTML = `${svgIcon('sync', 'inline-action-svg')}<span>正在同步...</span>`;
            triggerSubtreeSyncInBackground(parentItem.id, parentItem.relativeUrl);
          });
        } else {
          tipEl.innerHTML = `
            <span>该目录未在本地缓存中。请先 <button class="inline-fav-btn">${svgIcon('star-filled', 'inline-action-svg')}<span>收藏该文件夹</span></button> 以在后台自动同步其子目录。</span>
          `;
          tipEl.querySelector('.inline-fav-btn').addEventListener('click', () => {
            toggleFavorite(parentItem);
          });
        }
        container.appendChild(tipEl);
      } else {
        // 如果是深层目录没有缓存（不应发生，因为 1 级收藏会拉取整树，这里做容错处理）
        const alertEl = document.createElement('div');
        alertEl.className = 'tree-node';
        alertEl.style.paddingLeft = `${depth * 16 + 10}px`;
        alertEl.innerHTML = `<span class="node-name uncached-alert">缓存未加载</span>`;
        container.appendChild(alertEl);
      }
    }
  }

  // 判断指定路径是否有缓存子项，并返回子项列表 {folders, files}
  function getCachedSubtree(itemId, relativeUrl) {
    // 1. 如果自己就是 1 级收藏文件夹，直接查 subtreeCache
    if (subtreeCache[itemId] && subtreeCache[itemId].tree) {
      return subtreeCache[itemId].tree[relativeUrl] || null;
    }

    // 2. 如果自己是深层目录，我们需要找到其所属的 1 级祖先文件夹
    // 在 subtree_cache 中遍历所有缓存树，看看哪棵树中包含了这个 relativeUrl 的路径
    for (const l1Id in subtreeCache) {
      const cacheRoot = subtreeCache[l1Id];
      if (cacheRoot && cacheRoot.tree && cacheRoot.tree[relativeUrl]) {
        return cacheRoot.tree[relativeUrl];
      }
    }

    return null;
  }

  // 保存 UI 状态到本地存储
  async function saveUIState() {
    try {
      await chrome.storage.local.set({
        ui_state: {
          expandedState,
          isTreeExpanded,
          searchQuery: searchInput.value,
          activeFilter,
          activeL1Path
        }
      });
    } catch (e) {
      console.warn('Failed to save UI state:', e);
    }
  }

  // 动态填充 1 级目录筛选下拉框（仅包含已收藏的 1 级文件夹）
  function populateL1FilterDropdown() {
    const selectEl = document.getElementById('l1FilterSelect');
    if (!selectEl) return;

    // 保留第一个“所有 1 级目录”选项，清除其他选项
    selectEl.innerHTML = '<option value="all">所有 1 级目录</option>';

    if (favorites && Array.isArray(favorites)) {
      // 筛选出已收藏的 1 级目录文件夹，并按名称自然排序
      const favoritedL1Folders = favorites.filter(item => item.type === 'folder' && item.level === 1);
      favoritedL1Folders.sort((a, b) => a.name.localeCompare(b.name, 'zh-CN', { numeric: true }));
      
      favoritedL1Folders.forEach(item => {
        const option = document.createElement('option');
        option.value = item.relativeUrl;
        option.textContent = item.name;
        selectEl.appendChild(option);
      });
    }

    // 恢复之前的选中状态，如果当前选中的 relativeUrl 仍存在
    if (activeL1Path && activeL1Path !== 'all') {
      const exists = Array.from(selectEl.options).some(opt => opt.value === activeL1Path);
      if (exists) {
        selectEl.value = activeL1Path;
        selectEl.classList.add('active');
      } else {
        activeL1Path = 'all';
        selectEl.value = 'all';
        selectEl.classList.remove('active');
        
        // 如果在搜索视图下且选择的一级目录被取消收藏，刷新搜索面板
        const query = searchInput.value.trim().toLowerCase();
        if (query) {
          performSearch(query);
        }
      }
    } else {
      selectEl.value = 'all';
      selectEl.classList.remove('active');
    }
  }

  // 渲染并更新快捷收藏部分的同步进度显示
  function updateFavoritesSyncProgressDisplay() {
    const progressEl = document.getElementById('favSyncProgress');
    if (!progressEl) return;

    if (!syncStatus) {
      progressEl.classList.add('hide');
      progressEl.innerHTML = '';
      return;
    }

    const activeSyncs = [];
    Object.keys(syncStatus).forEach(folderId => {
      const isFavorited = favorites.some(fav => fav.id === folderId);
      if (isFavorited && syncStatus[folderId] && syncStatus[folderId].status === 'syncing') {
        activeSyncs.push(syncStatus[folderId]);
      }
    });
    
    if (activeSyncs.length === 0) {
      progressEl.classList.add('hide');
      progressEl.innerHTML = '';
      return;
    }

    // 汇总进度
    let totalFolders = 0;
    let totalNodes = 0;
    let currentPath = '';

    activeSyncs.forEach(s => {
      totalFolders += s.folderCount || 0;
      totalNodes += s.nodeCount || 0;
      if (s.currentFolder) {
        // 取最后一个路径节点做简短展示，避免溢出
        const parts = s.currentFolder.split('/');
        currentPath = parts[parts.length - 1] || s.currentFolder;
      }
    });

    progressEl.classList.remove('hide');

    const progressContent = document.createElement('span');
    progressContent.className = 'fav-sync-progress-content';
    progressContent.innerHTML = `${svgIcon('sync', 'progress-svg')}<span>更新中:</span>${svgIcon('folder', 'progress-inline-svg')}<span>${totalFolders}</span>${svgIcon('file', 'progress-inline-svg')}<span>${totalNodes}</span>`;
    if (currentPath) {
      progressContent.insertAdjacentHTML('beforeend', `<span>(${currentPath})</span>`);
    }
    
    progressEl.innerHTML = '';
    progressEl.appendChild(progressContent);
    progressEl.title = `正在更新：已扫描 ${totalFolders} 个文件夹，发现 ${totalNodes} 个文件${currentPath ? `。当前: ${currentPath}` : ''}`;
  }

  // ==================== 搜索逻辑 ====================

  function performSearch(query) {
    searchResultsList.innerHTML = '';
    
    // 调用共享模糊搜索逻辑
    const matchedItems = performFuzzySearchInCache(query, l1Cache, subtreeCache);

    // B. 进行过滤器筛选
    const filtered = matchedItems.filter(item => {
      // 1. 应用 Level 1 目录过滤器
      if (activeL1Path !== 'all') {
        const isMatchL1 = item.relativeUrl === activeL1Path || item.relativeUrl.startsWith(activeL1Path + '/');
        if (!isMatchL1) return false;
      }

      // 2. 应用类型过滤器
      if (activeFilter === 'folder') {
        if (item.type !== 'folder') return false;
      } else if (activeFilter !== 'all') {
        // 文件类型过滤
        if (item.type !== 'file') return false;
        const ext = item.name.split('.').pop().toLowerCase();
        
        if (activeFilter === 'word' && ext !== 'doc' && ext !== 'docx') return false;
        if (activeFilter === 'excel' && ext !== 'xls' && ext !== 'xlsx') return false;
        if (activeFilter === 'ppt' && ext !== 'ppt' && ext !== 'pptx') return false;
        if (activeFilter === 'pdf' && ext !== 'pdf') return false;
        if (activeFilter === 'prism' && ext !== 'prism' && ext !== 'pzfx') return false;
      }

      return true;
    });

    // 更新搜索结果总条数
    const countEl = document.getElementById('searchResultCount');
    if (countEl) {
      countEl.innerText = `共 ${filtered.length} 条`;
    }

    // C. 渲染搜索结果
    if (filtered.length === 0) {
      searchResultsList.innerHTML = '<div class="empty-list-placeholder">未找到匹配的结果。</div>';
    } else {
      // 限制最大渲染结果，防止卡顿
      const limit = Math.min(filtered.length, 50);
      for (let i = 0; i < limit; i++) {
        const item = filtered[i];
        const itemEl = document.createElement('div');
        itemEl.className = 'result-item';

        const isFolder = item.type === 'folder';
        const iconName = isFolder ? 'folder' : 'file';
        const icon = svgIcon(iconName, 'node-svg-icon');

        // 提取父级目录路径展示
        let parentPath = '';
        const pathParts = item.relativeUrl.split('/');
        if (pathParts.length > 2) {
          parentPath = pathParts.slice(0, -1).join('/');
        } else {
          parentPath = '/';
        }

        const isFav = favorites.some(fav => fav.id === item.id);

        itemEl.innerHTML = `
          <div class="result-info">
            <div class="result-title-row">
              <span class="node-icon ${isFolder ? 'folder-icon' : 'file-icon'}">${icon}</span>
              <span class="node-name" style="font-weight: 500;">${item.name}</span>
            </div>
            <div class="result-path" title="${item.relativeUrl}">路径: ${parentPath}</div>
          </div>
          <div class="node-right ${isFav ? 'is-fav' : ''}">
            <button class="action-btn open-btn" data-url="${item.webUrl}" title="在浏览器中打开网页" aria-label="在浏览器中打开网页">${svgIcon('external-link', 'action-svg')}</button>
            <button class="action-btn copy-btn" data-url="${item.webUrl}" title="复制 SharePoint 链接" aria-label="复制 SharePoint 链接">${svgIcon('clipboard', 'action-svg')}</button>
            <button class="fav-btn ${isFav ? 'active' : ''}" title="${isFav ? '取消快捷收藏' : '加入快捷收藏'}" aria-label="${isFav ? '取消快捷收藏' : '加入快捷收藏'}">${svgIcon(isFav ? 'star-filled' : 'star', 'fav-svg')}</button>
          </div>
        `;

        itemEl.querySelector('.open-btn').addEventListener('click', () => handleOpen(item.webUrl));
        itemEl.querySelector('.copy-btn').addEventListener('click', () => handleCopy(item.webUrl));
        itemEl.querySelector('.fav-btn').addEventListener('click', () => toggleFavorite(item));

        searchResultsList.appendChild(itemEl);
      }
    }

    defaultViews.classList.add('hide');
    searchResultsSection.classList.remove('hide');
  }

  // 触发后台同步子目录（通过 Background 避免弹窗关闭终止任务）
  function triggerSubtreeSyncInBackground(folderId, relativeUrl) {
    showToast('已在后台启动子目录同步，可以关闭此窗口...', 'sync');
    chrome.runtime.sendMessage({
      action: 'sync_subtree',
      folderId: folderId,
      relativeUrl: relativeUrl
    }, (response) => {
      if (chrome.runtime.lastError) {
        console.error('Background sync trigger error:', chrome.runtime.lastError);
        showToast('触发后台同步失败', 'alert', true);
      } else if (response && !response.success) {
        console.error('Background sync failed:', response.error);
        showToast(`同步失败: ${response.error}`, 'alert', true);
      } else {
        showToast('子树目录同步完成！已完全缓存。', 'check');
      }
    });
  }

  // ==================== 打开与复制行为 ====================

  // 获取在线预览 URL (对 Office 文档及 PDF 附加 ?web=1 以便在浏览器中在线打开)
  function getOnlineViewUrl(url) {
    if (!url) return url;
    
    let targetUrl = url;
    if (!targetUrl.includes('?')) {
      const officeExtensions = [
        'doc', 'docx', 'docm', 'dot', 'dotx',
        'xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx',
        'ppt', 'pptx', 'pps', 'ppsx', 'pptm',
        'pdf'
      ];
      
      const ext = targetUrl.split('.').pop().toLowerCase();
      if (officeExtensions.includes(ext)) {
        targetUrl = `${targetUrl}?web=1`;
      }
    }
    
    try {
      // 先对 URL 进行解码，再通过 encodeURI 进行标准化转码，确保空格、中文等字符在复制和打开时是一致的转码链接
      let decoded = targetUrl;
      try {
        decoded = decodeURIComponent(targetUrl);
      } catch (de) {
        try {
          decoded = decodeURI(targetUrl);
        } catch (de2) {}
      }

      let mainUrl = decoded;
      let suffix = '';
      const qIndex = decoded.indexOf('?web=1');
      if (qIndex !== -1) {
        mainUrl = decoded.substring(0, qIndex);
        suffix = decoded.substring(qIndex);
      }

      // 特殊字符转义：把文件名中可能出现的 # 替换为 %23，避免浏览器将其作为 Hash 片段截断
      const encodedMain = encodeURI(mainUrl).replace(/#/g, '%23');
      return encodedMain + suffix;
    } catch (e) {
      console.warn('URL decoding/encoding failed:', e);
      return encodeURI(targetUrl).replace(/#/g, '%23');
    }
  }

  function handleOpen(url) {
    const finalUrl = getOnlineViewUrl(url);
    chrome.tabs.create({ url: finalUrl });
  }

  function handleCopy(url) {
    const finalUrl = getOnlineViewUrl(url);
    navigator.clipboard.writeText(finalUrl)
      .then(() => {
        showToast('链接已成功复制到剪贴板！', 'clipboard');
      })
      .catch(err => {
        console.error('Failed to copy text: ', err);
        showToast('复制失败，请重试', 'alert');
      });
  }

  // ==================== 收藏管理逻辑 ====================

  async function toggleFavorite(item) {
    const index = favorites.findIndex(fav => fav.id === item.id);
    const isAdding = index === -1;

    if (isAdding) {
      // 添加收藏
      favorites.push({
        id: item.id,
        name: item.name,
        type: item.type,
        level: item.level || 2, // 默认为 2 级，以便识别是否为 1 级收藏
        relativeUrl: item.relativeUrl,
        webUrl: item.webUrl
      });
      showToast('已添加至快捷收藏', 'star-filled');

      // 核心要求：如果是 1 级文件夹，收藏后立即触发对子树进行全量递归同步
      if (item.type === 'folder' && item.level === 1) {
        triggerSubtreeSyncInBackground(item.id, item.relativeUrl);
      }
    } else {
      // 取消收藏
      favorites.splice(index, 1);
      showToast('已取消收藏', 'star');
      
      // 如果被删除的是 1 级目录，同时从 subtreeCache 中移除以释放存储空间
      if (item.type === 'folder' && item.level === 1) {
        delete subtreeCache[item.id];
        await chrome.storage.local.remove('subtree_cache_' + item.id);
      }
    }

    // 保存至 storage 并重绘界面
    const favKey = currentConfigId ? `favorites_${currentConfigId}` : 'favorites';
    await chrome.storage.local.set({ [favKey]: favorites });
    renderFavorites();
    renderDirectoryTree();
    
    // 如果在搜索视图下，刷新搜索面板
    const query = searchInput.value.trim().toLowerCase();
    if (query) {
      performSearch(query);
    }
  }

  // ==================== 提示条与 Toast 通用组件 ====================

  function showToast(message, iconName = '', isError = false) {
    const toast = document.createElement('div');
    toast.className = 'toast';
    if (isError) toast.classList.add('toast-error');

    if (iconName) {
      const icon = document.createElement('span');
      icon.className = `toast-icon toast-icon-${iconName}`;
      icon.innerHTML = svgIcon(iconName, 'toast-svg');
      toast.appendChild(icon);
    }

    const text = document.createElement('span');
    text.innerText = message;
    toast.appendChild(text);

    if (isError) {
      const copyButton = document.createElement('button');
      copyButton.className = 'toast-copy-btn';
      copyButton.type = 'button';
      copyButton.textContent = '复制错误';
      copyButton.addEventListener('click', async () => {
        try {
          await navigator.clipboard.writeText(message);
          copyButton.textContent = '已复制';
        } catch (err) {
          copyButton.textContent = '复制失败';
        }
      });
      toast.appendChild(copyButton);
    }
    
    toastContainer.appendChild(toast);
    
    // 动画结束后自动移除 DOM
    setTimeout(() => {
      toast.remove();
    }, isError ? 30000 : 2500);
  }

  function showAlert(msg, iconName = '') {
    alertText.innerHTML = '';
    if (iconName) {
      const icon = document.createElement('span');
      icon.className = `alert-icon alert-icon-${iconName}`;
      icon.innerHTML = svgIcon(iconName, 'alert-svg');
      alertText.appendChild(icon);
    }
    const text = document.createElement('span');
    text.innerText = msg;
    alertText.appendChild(text);
    alertBanner.classList.remove('hide');
  }

  function hideAlert() {
    alertBanner.classList.add('hide');
  }

  // 监听本地存储变化，实现实时响应刷新
  chrome.storage.onChanged.addListener(async (changes, namespace) => {
    if (namespace === 'local') {
      const hasSubtreeChange = Object.keys(changes).some(key => key.startsWith('subtree_cache_'));
      const hasSyncStatusChange = Object.keys(changes).some(key => key.startsWith('sync_status_'));
      const hasFavoritesChange = changes.favorites || Object.keys(changes).some(key => key.startsWith('favorites_'));
      const hasL1CacheChange = changes.l1_cache || Object.keys(changes).some(key => key.startsWith('l1_cache_'));
      const hasUpdateNotificationsChange = Boolean(changes[FILE_UPDATE_NOTIFICATIONS_KEY]);
      
      if (hasSubtreeChange || hasSyncStatusChange || hasFavoritesChange || hasL1CacheChange || hasUpdateNotificationsChange || changes.current_config_id || changes.sp_configs) {
        await loadDataFromStorage();
        populateL1FilterDropdown();
        renderFavorites();
        renderDirectoryTree();
        updateFavoritesSyncProgressDisplay();
      }
    }
  });
});
