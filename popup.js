// popup.js

document.addEventListener('DOMContentLoaded', async () => {
  // DOM 元素获取
  const refreshBtn = document.getElementById('refreshBtn');
  const settingsBtn = document.getElementById('settingsBtn');
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
  const syncTimeSpan = document.getElementById('syncTime');
  const toastContainer = document.getElementById('toastContainer');

  // 全局数据状态缓存
  let spConfig = null;
  let favorites = [];
  let l1Cache = null;
  let subtreeCache = {};
  
  // 树状图折叠展开状态映射 (folderId -> boolean)
  let expandedState = {};

  // 1. 初始化检查配置
  await initApp();

  // 2. 绑定页面通用交互事件
  settingsBtn.addEventListener('click', () => chrome.runtime.openOptionsPage());
  alertActionBtn.addEventListener('click', () => chrome.runtime.openOptionsPage());
  
  // 手动同步事件
  refreshBtn.addEventListener('click', async () => {
    refreshBtn.classList.add('loading');
    showToast('🔄 正在同步 1 级目录...');
    
    try {
      const items = await syncLevel1();
      showToast('✨ 1 级目录同步成功！');
      await loadDataFromStorage();
      renderFavorites();
      renderDirectoryTree();
    } catch (err) {
      console.error(err);
      showToast(`❌ 同步失败: ${err.message || err}`);
    } finally {
      refreshBtn.classList.remove('loading');
    }
  });

  // 搜索输入过滤
  searchInput.addEventListener('input', () => {
    const query = searchInput.value.trim().toLowerCase();
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
    searchInput.focus();
  });

  // ==================== 初始化与数据加载 ====================

  async function initApp() {
    const data = await chrome.storage.local.get(['sp_config', 'favorites', 'l1_cache', 'subtree_cache']);
    spConfig = data.sp_config;
    favorites = data.favorites || [];
    l1Cache = data.l1_cache;
    subtreeCache = data.subtree_cache || {};

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
            renderFavorites();
            renderDirectoryTree();
          })
          .catch(err => console.warn('Auto refresh failed:', err));
      }
      
      updateSyncTimeDisplay();
      renderFavorites();
      renderDirectoryTree();
    }
  }

  // 重新从 storage 读取最新数据
  async function loadDataFromStorage() {
    const data = await chrome.storage.local.get(['favorites', 'l1_cache', 'subtree_cache']);
    favorites = data.favorites || [];
    l1Cache = data.l1_cache;
    subtreeCache = data.subtree_cache || {};
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
    favoritesList.innerHTML = '';
    if (favorites.length === 0) {
      favoritesList.className = 'empty-list-placeholder';
      favoritesList.innerText = '暂无收藏。点击目录树中文件夹或文件旁的 ⭐ 即可加入收藏。';
      return;
    }

    favoritesList.className = '';
    favorites.forEach(fav => {
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

    nodeEl.innerHTML = `
      <div class="node-left">
        <span class="${toggleClass}">${toggleIcon}</span>
        <span class="node-icon">${icon}</span>
        <span class="node-name ${isFolder ? 'folder-node' : ''}" title="${item.relativeUrl}">${item.name}</span>
      </div>
      <div class="node-right ${isFav ? 'is-fav' : ''}">
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

    // 3. 展开折叠事件 (如果是文件夹)
    if (isFolder) {
      const toggleNode = () => {
        const nextExpanded = !expandedState[item.id];
        expandedState[item.id] = nextExpanded;
        
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

    } else {
      // 不在缓存中：这只能是 1 级目录
      if (parentItem.level === 1) {
        const isFav = favorites.some(fav => fav.id === parentItem.id);
        const tipEl = document.createElement('div');
        tipEl.className = 'uncached-tip';
        tipEl.style.marginLeft = `${(depth - 1) * 16 + 10}px`;
        
        if (isFav) {
          tipEl.innerHTML = `
            <span>该目录缓存未同步。请点击 <button class="inline-sync-btn">🔄 重新同步子目录</button> 尝试拉取。</span>
          `;
          tipEl.querySelector('.inline-sync-btn').addEventListener('click', async () => {
            const btn = tipEl.querySelector('.inline-sync-btn');
            btn.disabled = true;
            btn.innerText = '⏳ 正在同步...';
            showToast('🚀 正在同步该文件夹下的子目录...');
            try {
              await syncSubtree(parentItem.id, parentItem.relativeUrl);
              showToast('✨ 子树目录同步完成！已完全缓存。');
              await loadDataFromStorage();
              renderDirectoryTree();
            } catch (err) {
              console.error('Subtree sync failed:', err);
              showToast(`⚠️ 同步失败: ${err.message || err}`);
              btn.disabled = false;
              btn.innerText = '🔄 重新同步子目录';
            }
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

    // B. 进行模糊搜索匹配
    const filtered = searchPool.filter(item => {
      const nameMatch = item.name.toLowerCase().includes(query);
      const pathMatch = item.relativeUrl.toLowerCase().includes(query);
      return nameMatch || pathMatch;
    });

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

  // ==================== 打开与复制行为 ====================

  function handleOpen(url) {
    chrome.tabs.create({ url: url });
  }

  function handleCopy(url) {
    navigator.clipboard.writeText(url)
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
        showToast('🚀 正在同步该文件夹下的子目录...');
        syncSubtree(item.id, item.relativeUrl)
          .then(async (count) => {
            showToast('✨ 子树目录同步完成！已完全缓存。');
            // 重新拉取 storage 刷新树的节点
            await loadDataFromStorage();
            renderDirectoryTree();
          })
          .catch((err) => {
            console.error('Subtree sync failed:', err);
            showToast('⚠️ 子目录同步失败，请检查网络登录态。');
          });
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
});
