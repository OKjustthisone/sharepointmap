// content.js

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'show_search_results') {
    showSearchResultsModal(request.query, request.results);
    sendResponse({ success: true });
  }
});

function showSearchResultsModal(query, results) {
  // 1. 移除已存在的 Modal，防止重复堆叠
  const existing = document.getElementById('sp-map-search-overlay');
  if (existing) {
    existing.remove();
  }

  // 2. 创建最外层 Overlay 容器
  const overlay = document.createElement('div');
  overlay.id = 'sp-map-search-overlay';

  // 3. 构建 Modal HTML 内容
  const modal = document.createElement('div');
  modal.className = 'sp-map-modal-card';

  // Header 部分
  const header = document.createElement('div');
  header.className = 'sp-map-modal-header';

  const title = document.createElement('h3');
  title.className = 'sp-map-modal-title';
  title.innerText = `SharePoint Map 搜索结果: "${query}"`;
  header.appendChild(title);

  const closeBtn = document.createElement('button');
  closeBtn.className = 'sp-map-modal-close-btn';
  closeBtn.innerHTML = '✕';
  closeBtn.addEventListener('click', closeModal);
  header.appendChild(closeBtn);

  modal.appendChild(header);

  // Body 部分
  const body = document.createElement('div');
  body.className = 'sp-map-modal-body';

  const info = document.createElement('div');
  info.className = 'sp-map-results-info';
  info.innerText = `找到 ${results.length} 个匹配项`;
  body.appendChild(info);

  if (results.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'sp-map-empty-placeholder';
    empty.innerText = '🔍 未能找到匹配的文件夹或文件。确保插件已配置且数据已同步。';
    body.appendChild(empty);
  } else {
    const list = document.createElement('div');
    list.className = 'sp-map-results-list';

    results.forEach(item => {
      const itemEl = document.createElement('div');
      itemEl.className = 'sp-map-result-item';

      const left = document.createElement('div');
      left.className = 'sp-map-item-left';

      const icon = document.createElement('span');
      icon.className = 'sp-map-item-icon';
      icon.innerText = getItemIcon(item);
      left.appendChild(icon);

      const text = document.createElement('div');
      text.className = 'sp-map-item-text';

      const name = document.createElement('span');
      name.className = 'sp-map-item-name';
      name.innerText = item.name;
      name.title = `点击在浏览器中打开: ${item.name}`;
      
      // 文件或文件夹点击，直接调用 window.open 打开
      name.addEventListener('click', () => {
        window.open(getOnlineViewUrl(item.webUrl, item.name), '_blank');
      });

      text.appendChild(name);

      const path = document.createElement('span');
      path.className = 'sp-map-item-path';
      path.innerText = item.relativeUrl;
      text.appendChild(path);

      left.appendChild(text);
      itemEl.appendChild(left);

      // 右侧操作按钮
      const actions = document.createElement('div');
      actions.className = 'sp-map-item-actions';

      // 1. 打开按钮
      const openBtn = document.createElement('button');
      openBtn.className = 'sp-map-action-btn';
      openBtn.title = '在新标签页中打开';
      openBtn.innerHTML = '🌐';
      openBtn.addEventListener('click', () => {
        window.open(getOnlineViewUrl(item.webUrl, item.name), '_blank');
      });
      actions.appendChild(openBtn);

      // 2. 复制链接按钮
      const copyBtn = document.createElement('button');
      copyBtn.className = 'sp-map-action-btn';
      copyBtn.title = '复制链接';
      copyBtn.innerHTML = '🔗';
      copyBtn.addEventListener('click', () => {
        navigator.clipboard.writeText(getOnlineViewUrl(item.webUrl, item.name))
          .then(() => {
            showToastMessage('📋 链接已成功复制到剪贴板！');
          })
          .catch(() => {
            showToastMessage('❌ 复制失败，请重试');
          });
      });
      actions.appendChild(copyBtn);

      itemEl.appendChild(actions);
      list.appendChild(itemEl);
    });

    body.appendChild(list);
  }

  modal.appendChild(body);
  overlay.appendChild(modal);
  document.body.appendChild(overlay);

  // 点击外部关闭
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) {
      closeModal();
    }
  });

  // 按 Esc 键关闭
  document.addEventListener('keydown', escCloseHandler);

  function closeModal() {
    overlay.remove();
    document.removeEventListener('keydown', escCloseHandler);
  }

  function escCloseHandler(e) {
    if (e.key === 'Escape') {
      closeModal();
    }
  }
}

// 辅助函数：根据文件类型获取图标
function getItemIcon(item) {
  if (item.type === 'folder') return '📁';
  const ext = item.name.split('.').pop().toLowerCase();
  switch (ext) {
    case 'doc':
    case 'docx':
      return '📄';
    case 'xls':
    case 'xlsx':
      return '📊';
    case 'ppt':
    case 'pptx':
      return '📉';
    case 'pdf':
      return '📕';
    case 'opju':
      return '📊';
    default:
      return '📝';
  }
}

// 辅助函数：复制 popup.js 中的在线预览 URL 转码和附件逻辑
function getOnlineViewUrl(url, filename) {
  if (!url) return url;
  
  let targetUrl = url;
  if (!targetUrl.includes('?')) {
    const officeExtensions = [
      'doc', 'docx', 'docm', 'dot', 'dotx',
      'xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx',
      'ppt', 'pptx', 'pps', 'ppsx', 'pptm',
      'pdf'
    ];
    
    const ext = filename.split('.').pop().toLowerCase();
    if (officeExtensions.includes(ext)) {
      targetUrl = `${targetUrl}?web=1`;
    }
  }
  
  try {
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

    const encodedMain = encodeURI(mainUrl).replace(/#/g, '%23');
    return encodedMain + suffix;
  } catch (e) {
    return encodeURI(targetUrl).replace(/#/g, '%23');
  }
}

// 注入浮动 Toast 提示消息函数
function showToastMessage(msg) {
  const existingToast = document.getElementById('sp-map-toast-message');
  if (existingToast) {
    existingToast.remove();
  }

  const toast = document.createElement('div');
  toast.id = 'sp-map-toast-message';
  
  // 设置内联样式，绝对隔离
  Object.assign(toast.style, {
    position: 'fixed',
    bottom: '50px',
    left: '50%',
    transform: 'translateX(-50%)',
    background: 'rgba(11, 12, 19, 0.9)',
    border: '1px solid rgba(0, 240, 255, 0.3)',
    color: '#ffffff',
    padding: '12px 24px',
    borderRadius: '10px',
    fontSize: '14px',
    fontWeight: '500',
    boxShadow: '0 10px 30px rgba(0, 240, 255, 0.1)',
    zIndex: '9999999999',
    pointerEvents: 'none',
    transition: 'all 0.3s ease',
    opacity: '0'
  });

  toast.innerText = msg;
  document.body.appendChild(toast);

  // 触发 CSS 过渡
  setTimeout(() => {
    toast.style.opacity = '1';
    toast.style.transform = 'translate(-50%, -10px)';
  }, 50);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translate(-50%, -30px)';
    setTimeout(() => {
      toast.remove();
    }, 300);
  }, 2000);
}
