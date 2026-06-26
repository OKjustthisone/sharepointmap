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

  // 全局数据状态缓存与过滤器状态
  let activeFilter = 'all';
  let activeL1Path = 'all';
  let spConfig = null;
  let favorites = [];
  let l1Cache = null;
  let subtreeCache = {};
  let syncStatus = {};
  
  // 树状图折叠展开状态映射 (folderId -> boolean)
  let expandedState = {};
  let isTreeExpanded = false; // 全部目录折叠展开状态

  // 1. 初始化检查配置
  await initApp();

  // 2. 绑定页面通用交互事件
  settingsBtn.addEventListener('click', () => chrome.runtime.openOptionsPage());
  alertActionBtn.addEventListener('click', () => chrome.runtime.openOptionsPage());
  
  // 3. 绑定“全部目录”展开/折叠事件
  directoryToggleBtn.addEventListener('click', () => {
    isTreeExpanded = !isTreeExpanded;
    saveUIState(); // 保存状态
    if (isTreeExpanded) {
      directoryTree.classList.remove('collapsed');
      directoryToggleArrow.innerText = '▼ 收起目录';
      directoryToggleArrow.style.color = 'var(--primary-cyan)';
    } else {
      directoryTree.classList.add('collapsed');
      directoryToggleArrow.innerText = '▶ 展开浏览';
      directoryToggleArrow.style.color = 'var(--text-muted)';
    }
  });
  
  // 手动同步 1 级目录事件
  syncL1Btn.addEventListener('click', (e) => {
    e.stopPropagation();
    syncL1Btn.classList.add('loading');
    showToast('🔄 已在后台启动 1 级目录同步...');
    
    chrome.runtime.sendMessage({ action: 'sync_level1' }, (response) => {
      syncL1Btn.classList.remove('loading');
      if (chrome.runtime.lastError) {
        console.error('Background sync level 1 failed:', chrome.runtime.lastError);
        showToast('❌ 同步 1 级目录异常: ' + chrome.runtime.lastError.message);
      } else if (response && !response.success) {
        console.error('Background sync level 1 returned error:', response.error);
        showToast(`❌ 同步 1 级目录失败: ${response.error}`);
      } else {
        showToast('✨ 1 级目录同步完成！');
      }
    });
  });

  // 搜索过滤芯片事件绑定

  const filterChips = document.querySelectorAll('.filter-chip');
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
    const data = await chrome.storage.local.get(['sp_config', 'favorites', 'l1_cache', 'subtree_cache', 'sync_status', 'ui_state']);
    spConfig = data.sp_config;
    favorites = data.favorites || [];
    l1Cache = data.l1_cache;
    subtreeCache = data.subtree_cache || {};
    syncStatus = data.sync_status || {};

    // 恢复 UI 状态变量
    const uiState = data.ui_state || {};
    expandedState = uiState.expandedState || {};
    isTreeExpanded = uiState.isTreeExpanded || false;
    activeFilter = uiState.activeFilter || 'all';
    activeL1Path = uiState.activeL1Path || 'all';

    if (!spConfig || !spConfig.siteUrl || !spConfig.libraryName) {
      showAlert('⚠️ 请先配置您的 SharePoint 站点与文档库。');
      directoryTree.innerHTML = `
        <div class="empty-list-placeholder">
          配置未设置。请点击右上方 ⚙️ 按钮进入设置页面配置站点。
        </div>
      `;
      return;
    }

    hideAlert();

    // 恢复全部目录的 UI 展开/收起状态
    if (isTreeExpanded) {
      directoryTree.classList.remove('collapsed');
      directoryToggleArrow.innerText = '▼ 收起目录';
      directoryToggleArrow.style.color = 'var(--primary-cyan)';
    } else {
      directoryTree.classList.add('collapsed');
      directoryToggleArrow.innerText = '▶ 展开浏览';
      directoryToggleArrow.style.color = 'var(--text-muted)';
    }
    
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
        await syncLevel1();
        showToast('✨ 1 级目录同步成功！');
        await loadDataFromStorage();
        populateL1FilterDropdown();
        renderFavorites();
        renderDirectoryTree();
      } catch (err) {
        console.error(err);
        directoryTree.innerHTML = `
          <div class="empty-list-placeholder" style="color: var(--danger-red);">
            ⚠️ 同步失败，请点击右上角 🔄 手动重试。<br>
            错误原因: ${err.message || err}
          </div>
        `;
      }
    } else {
      // 检查是否已过期 (7 天)
      const isExpired = Date.now() - l1Cache.last_updated > 7 * 24 * 60 * 60 * 1000;
      if (isExpired) {
        showToast('🔄 正在自动更新已过期的缓存...');
        syncLevel1()
          .then(async () => {
            console.log('Auto refresh of Level 1 completed.');
            await loadDataFromStorage();
            populateL1FilterDropdown();
            renderFavorites();
            renderDirectoryTree();
            // 在数据后台同步完成后，如果刚才有恢复搜索，重新执行一下搜索
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

      // 恢复搜索输入与触发搜索
      const query = uiState.searchQuery || '';
      if (query) {
        searchInput.value = query;
        clearSearchBtn.classList.remove('hide');
        performSearch(query);
      }
    }
  }

  // 重新从 storage 读取最新数据
  async function loadDataFromStorage() {
    const data = await chrome.storage.local.get(['favorites', 'l1_cache', 'subtree_cache', 'sync_status']);
    favorites = data.favorites || [];
    l1Cache = data.l1_cache;
    subtreeCache = data.subtree_cache || {};
    syncStatus = data.sync_status || {};
    updateSyncTimeDisplay();
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
      favoritesList.innerText = '暂无收藏。点击目录树中文件夹或文件旁的 ⭐ 即可加入收藏。';
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
      const favNode = createTreeNodeElement(fav, 1);
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
  function createTreeNodeElement(item, depth) {
    const container = document.createElement('div');
    container.className = 'tree-node-wrapper';

    const isFolder = item.type === 'folder';
    const isExpanded = expandedState[item.id] || false;
    const isFav = favorites.some(fav => fav.id === item.id);

    // 判断该文件夹是否已在缓存中 (如果是 1 级收藏文件夹，或者其父辈已被收藏)
    const hasCache = getCachedSubtree(item.id, item.relativeUrl);

    // 折叠展开箭头状态
    let toggleIcon = '▶';
    let toggleClass = 'node-toggle';
    if (!isFolder) {
      toggleClass += ' empty';
    } else if (isExpanded) {
      toggleClass += ' expanded';
    }

    const nodeEl = document.createElement('div');
    nodeEl.className = 'tree-node';
    nodeEl.style.paddingLeft = `${(depth - 1) * 16 + 10}px`;

    const icon = isFolder ? (isExpanded ? '📂' : '📁') : '📄';

    let syncBtnHtml = '';
    if (isFolder && item.level === 1 && isFav) {
      syncBtnHtml = `<button class="action-btn sync-btn" title="同步该目录下的子树">同步 🔄</button>`;
    }

    nodeEl.innerHTML = `
      <div class="node-left">
        <span class="${toggleClass}">${toggleIcon}</span>
        <span class="node-icon">${icon}</span>
        <span class="node-name ${isFolder ? 'folder-node' : ''}" title="${item.relativeUrl}">${item.name}</span>
      </div>
      <div class="node-right ${isFav ? 'is-fav' : ''}">
        ${syncBtnHtml}
        <button class="action-btn open-btn" data-url="${item.webUrl}">打开 🌐</button>
        <button class="action-btn copy-btn" data-url="${item.webUrl}">复制 📋</button>
        <button class="fav-btn ${isFav ? 'active' : ''}">${isFav ? '⭐' : '☆'}</button>
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
    nodeEl.querySelector('.fav-btn').addEventListener('click', (e) => {
      e.stopPropagation();
      toggleFavorite(item);
    });

    // 2.5 同步子树按钮事件
    if (isFolder && item.level === 1 && isFav) {
      const syncBtn = nodeEl.querySelector('.sync-btn');
      if (syncBtn) {
        syncBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          syncBtn.classList.add('loading');
          
          showToast(`🚀 已在后台启动目录 [${item.name}] 的同步，可以关闭此窗口...`);
          
          chrome.runtime.sendMessage({
            action: 'sync_subtree',
            folderId: item.id,
            relativeUrl: item.relativeUrl
          }, (response) => {
            syncBtn.classList.remove('loading');
            if (chrome.runtime.lastError) {
              console.error('Background sync subtree failed:', chrome.runtime.lastError);
              showToast('❌ 同步子目录异常: ' + chrome.runtime.lastError.message);
            } else if (response && !response.success) {
              console.error('Background sync subtree returned error:', response.error);
              showToast(`❌ 同步子目录失败: ${response.error}`);
            } else {
              showToast(`✨ 目录 [${item.name}] 同步完成！`);
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
          folderIcon.innerText = '📂';
          childrenContainer.style.display = 'block';
          
          // 加载子项
          renderSubtreeItems(item, childrenContainer, depth + 1);
        } else {
          arrow.classList.remove('expanded');
          folderIcon.innerText = '📁';
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
        summaryEl.innerHTML = `📊 该目录下共含有 <strong>${totalFolders}</strong> 个文件夹，<strong>${totalFiles}</strong> 个文件`;
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
          const progress = syncStatus[parentItem.id];
          const folderCount = progress ? progress.folderCount : 0;
          const nodeCount = progress ? progress.nodeCount : 0;
          const currentFolder = progress ? progress.currentFolder : '';
          
          tipEl.innerHTML = `
            <div class="sync-progress-wrapper" style="display: flex; align-items: center; gap: 8px;">
              <div class="spinner mini"></div>
              <span>⏳ 正在后台同步: 已扫描 <strong>${folderCount}</strong> 个目录，发现 <strong>${nodeCount}</strong> 个项目...</span>
            </div>
            ${currentFolder ? `<div style="font-size: 11px; color: var(--text-muted); margin-top: 4px; padding-left: 22px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 320px;" title="${currentFolder}">当前: ${currentFolder}</div>` : ''}
          `;
        } else if (isFav) {
          tipEl.innerHTML = `
            <span>该目录缓存未同步。请点击 <button class="inline-sync-btn">🔄 重新同步子目录</button> 尝试拉取。</span>
          `;
          tipEl.querySelector('.inline-sync-btn').addEventListener('click', () => {
            const btn = tipEl.querySelector('.inline-sync-btn');
            btn.disabled = true;
            btn.innerText = '⏳ 正在同步...';
            triggerSubtreeSyncInBackground(parentItem.id, parentItem.relativeUrl);
          });
        } else {
          tipEl.innerHTML = `
            <span>该目录未在本地缓存中。请先 <button class="inline-fav-btn">★ 收藏该文件夹</button> 以在后台自动同步其子目录。</span>
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
    selectEl.innerHTML = '<option value="all">所有 1 级目录 📂</option>';

    if (favorites && Array.isArray(favorites)) {
      // 筛选出已收藏的 1 级目录文件夹，并按名称自然排序
      const favoritedL1Folders = favorites.filter(item => item.type === 'folder' && item.level === 1);
      favoritedL1Folders.sort((a, b) => a.name.localeCompare(b.name, 'zh-CN', { numeric: true }));
      
      favoritedL1Folders.forEach(item => {
        const option = document.createElement('option');
        option.value = item.relativeUrl;
        option.textContent = `📁 ${item.name}`;
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

  // ==================== 搜索逻辑 ====================

  function performSearch(query) {
    searchResultsList.innerHTML = '';
    const results = [];

    // A. 收集所有可供搜索的节点
    // 1. 收集 1 级目录
    if (l1Cache && l1Cache.items) {
      l1Cache.items.forEach(item => {
        results.push(item);
      });
    }

    // 2. 收集所有已缓存的子树节点
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

    // 去重 (因为同一个子目录可能在不同子树里遍历过，虽然理论上不会)
    const uniqueMap = new Map();
    results.forEach(item => {
      uniqueMap.set(item.id || item.relativeUrl, item);
    });

    const searchPool = Array.from(uniqueMap.values());

    // B. 进行模糊搜索匹配与过滤器筛选
    const filtered = searchPool.filter(item => {
      const nameMatch = item.name.toLowerCase().includes(query);
      const pathMatch = item.relativeUrl.toLowerCase().includes(query);
      if (!nameMatch && !pathMatch) return false;

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
        const icon = isFolder ? '📁' : '📄';

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
              <span class="node-icon">${icon}</span>
              <span class="node-name" style="font-weight: 500;">${item.name}</span>
            </div>
            <div class="result-path" title="${item.relativeUrl}">路径: ${parentPath}</div>
          </div>
          <div class="node-right ${isFav ? 'is-fav' : ''}">
            <button class="action-btn open-btn" data-url="${item.webUrl}">打开 🌐</button>
            <button class="action-btn copy-btn" data-url="${item.webUrl}">复制 📋</button>
            <button class="fav-btn ${isFav ? 'active' : ''}">${isFav ? '⭐' : '☆'}</button>
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
    showToast('🚀 已在后台启动子目录同步，可以关闭此窗口...');
    chrome.runtime.sendMessage({
      action: 'sync_subtree',
      folderId: folderId,
      relativeUrl: relativeUrl
    }, (response) => {
      if (chrome.runtime.lastError) {
        console.error('Background sync trigger error:', chrome.runtime.lastError);
        showToast('❌ 触发后台同步失败');
      } else if (response && !response.success) {
        console.error('Background sync failed:', response.error);
        showToast(`❌ 同步失败: ${response.error}`);
      } else {
        showToast('✨ 子树目录同步完成！已完全缓存。');
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
      return encodeURI(decodeURI(targetUrl));
    } catch (e) {
      console.warn('URL decoding/encoding failed:', e);
      return encodeURI(targetUrl);
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
        showToast('📋 链接已成功复制到剪贴板！');
      })
      .catch(err => {
        console.error('Failed to copy text: ', err);
        showToast('❌ 复制失败，请重试');
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
      showToast('⭐ 已添加至快捷收藏');

      // 核心要求：如果是 1 级文件夹，收藏后立即触发对子树进行全量递归同步
      if (item.type === 'folder' && item.level === 1) {
        triggerSubtreeSyncInBackground(item.id, item.relativeUrl);
      }
    } else {
      // 取消收藏
      favorites.splice(index, 1);
      showToast('☆ 已取消收藏');
      
      // 如果被删除的是 1 级目录，同时从 subtreeCache 中移除以释放存储空间
      if (item.type === 'folder' && item.level === 1) {
        delete subtreeCache[item.id];
        await chrome.storage.local.set({ subtree_cache: subtreeCache });
      }
    }

    // 保存至 storage 并重绘界面
    await chrome.storage.local.set({ favorites: favorites });
    renderFavorites();
    renderDirectoryTree();
    
    // 如果在搜索视图下，刷新搜索面板
    const query = searchInput.value.trim().toLowerCase();
    if (query) {
      performSearch(query);
    }
  }

  // ==================== 提示条与 Toast 通用组件 ====================

  function showToast(message) {
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerHTML = `<span>${message}</span>`;
    
    toastContainer.appendChild(toast);
    
    // 动画结束后自动移除 DOM
    setTimeout(() => {
      toast.remove();
    }, 2500);
  }

  function showAlert(msg) {
    alertText.innerText = msg;
    alertBanner.classList.remove('hide');
  }

  function hideAlert() {
    alertBanner.classList.add('hide');
  }

  // 监听本地存储变化，实现实时响应刷新
  chrome.storage.onChanged.addListener(async (changes, namespace) => {
    if (namespace === 'local') {
      if (changes.subtree_cache || changes.favorites || changes.l1_cache || changes.sync_status) {
        await loadDataFromStorage();
        populateL1FilterDropdown();
        renderFavorites();
        renderDirectoryTree();
      }
    }
  });
});
