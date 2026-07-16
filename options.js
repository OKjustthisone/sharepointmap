// options.js

document.addEventListener('DOMContentLoaded', async () => {
  const configsList = document.getElementById('configsList');
  const formTitle = document.getElementById('formTitle');
  const configNameInput = document.getElementById('configName');
  const siteUrlInput = document.getElementById('siteUrl');
  const libraryNameInput = document.getElementById('libraryName');
  const saveBtn = document.getElementById('saveBtn');
  const cancelBtn = document.getElementById('cancelBtn');
  const statusContainer = document.getElementById('statusContainer');
  const statusTitle = document.getElementById('statusTitle');
  const statusDesc = document.getElementById('statusDesc');

  let spConfigs = [];
  let currentConfigId = '';
  let editingConfigId = null; // null 表示当前是 "添加" 模式

  // 1. 初始化并运行配置迁移，然后加载数据
  await migrateConfigsIfNeeded();
  await loadConfigs();

  async function loadConfigs() {
    const data = await chrome.storage.local.get(['sp_configs', 'current_config_id']);
    spConfigs = data.sp_configs || [];
    currentConfigId = data.current_config_id || '';
    renderConfigsList();
  }

  function renderConfigsList() {
    configsList.innerHTML = '';
    if (spConfigs.length === 0) {
      configsList.innerHTML = `<div style="text-align: center; color: var(--text-tip); padding: 12px; font-size: 13px;">暂无配置站点，请在下方添加。</div>`;
      return;
    }

    spConfigs.forEach(config => {
      const isCurrent = config.id === currentConfigId;
      const itemEl = document.createElement('div');
      itemEl.className = `config-item ${isCurrent ? 'active-config' : ''}`;
      
      // 构建信息部分
      const infoEl = document.createElement('div');
      infoEl.className = 'config-info';
      infoEl.style.cursor = 'pointer';
      infoEl.addEventListener('click', () => handleSelectActive(config.id));

      const titleRow = document.createElement('div');
      titleRow.className = 'config-title-row';

      const nameEl = document.createElement('span');
      nameEl.className = 'config-item-name';
      nameEl.innerText = config.name || '未命名配置';
      titleRow.appendChild(nameEl);

      if (isCurrent) {
        const badge = document.createElement('span');
        badge.className = 'config-active-badge';
        badge.innerText = '当前激活';
        titleRow.appendChild(badge);
      }

      infoEl.appendChild(titleRow);

      const detailsEl = document.createElement('span');
      detailsEl.className = 'config-item-details';
      detailsEl.innerText = `${config.siteUrl} (${config.libraryName})`;
      infoEl.appendChild(detailsEl);

      itemEl.appendChild(infoEl);

      // 操作按钮部分
      const actionsEl = document.createElement('div');
      actionsEl.className = 'config-actions';

      const editBtn = document.createElement('button');
      editBtn.className = 'icon-btn';
      editBtn.title = '编辑配置';
      editBtn.innerHTML = '✏️';
      editBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        startEdit(config);
      });

      const deleteBtn = document.createElement('button');
      deleteBtn.className = 'icon-btn delete-btn';
      deleteBtn.title = '删除配置';
      deleteBtn.innerHTML = '🗑️';
      deleteBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        handleDelete(config.id);
      });

      actionsEl.appendChild(editBtn);
      actionsEl.appendChild(deleteBtn);
      itemEl.appendChild(actionsEl);

      configsList.appendChild(itemEl);
    });
  }

  // 切换激活状态
  async function handleSelectActive(configId) {
    if (configId === currentConfigId) return;
    currentConfigId = configId;
    await chrome.storage.local.set({ current_config_id: configId });
    renderConfigsList();
    showStatus('success', '激活成功', '已切换当前激活的站点。在弹窗窗口中将显示此站点的目录树和收藏夹。');
  }

  // 开始编辑配置
  function startEdit(config) {
    editingConfigId = config.id;
    formTitle.innerText = `编辑配置: ${config.name}`;
    configNameInput.value = config.name || '';
    siteUrlInput.value = config.siteUrl || '';
    libraryNameInput.value = config.libraryName || 'Shared Documents';
    saveBtn.innerText = '更新并测试连接';
    cancelBtn.classList.remove('hide');
    statusContainer.classList.add('hide');
    configNameInput.focus();
  }

  // 取消编辑
  cancelBtn.addEventListener('click', () => {
    resetForm();
  });

  function resetForm() {
    editingConfigId = null;
    formTitle.innerText = '配置新站点连接';
    configNameInput.value = '';
    siteUrlInput.value = '';
    libraryNameInput.value = 'Shared Documents';
    saveBtn.innerText = '保存并测试连接';
    cancelBtn.classList.add('hide');
    statusContainer.classList.add('hide');
  }

  // 删除配置
  async function handleDelete(configId) {
    if (!confirm('确定要删除此站点配置吗？该站点的本地缓存和快捷收藏也将被清除。')) {
      return;
    }

    spConfigs = spConfigs.filter(c => c.id !== configId);
    const updates = { sp_configs: spConfigs };

    // 清除该配置关联的缓存
    await chrome.storage.local.remove([
      `l1_cache_${configId}`,
      `favorites_${configId}`
    ]);

    // 如果删除的是当前激活的，切换到另一个
    if (currentConfigId === configId) {
      currentConfigId = spConfigs.length > 0 ? spConfigs[0].id : '';
      updates.current_config_id = currentConfigId;
    }

    await chrome.storage.local.set(updates);
    renderConfigsList();
    showStatus('success', '删除成功', '已成功删除站点配置及本地缓存。');
    if (editingConfigId === configId) {
      resetForm();
    }
  }

  // 保存（新增/编辑）按钮事件
  saveBtn.addEventListener('click', async () => {
    const configName = configNameInput.value.trim();
    let siteUrl = siteUrlInput.value.trim();
    const libraryName = libraryNameInput.value.trim();

    if (!configName || !siteUrl || !libraryName) {
      showStatus('error', '保存失败', '请完整填写配置别名、SharePoint 站点 URL 和文档库名称。');
      return;
    }

    if (siteUrl.endsWith('/')) {
      siteUrl = siteUrl.slice(0, -1);
    }

    try {
      showStatus('testing', '正在配置权限...', '正在请求对该站点的访问权限，请在浏览器权限弹窗中确认（如适用）...');
      
      const parsedUrl = new URL(siteUrl);
      const siteOrigin = parsedUrl.origin;

      const permissionGranted = await requestHostPermission(siteOrigin);
      if (!permissionGranted) {
        showStatus('error', '权限被拒绝', '未能获得该域名的访问权限，无法发起 API 连接测试。');
        return;
      }

      showStatus('testing', '正在测试连接并拉取缓存...', '正在连接 SharePoint 并同步首屏数据，请稍候...');

      // 临时生成一个临时 ID 进行同步测试，防止写入失败导致原有数据丢失
      const targetId = editingConfigId || `config_${Date.now()}`;
      const tempConfig = { id: targetId, name: configName, siteUrl, libraryName, siteOrigin };

      // 先把临时配置存入 sp_configs 中以供 syncLevel1 读取测试（syncLevel1 中会根据 configId 读取）
      const backupConfigs = [...spConfigs];
      const existingIndex = spConfigs.findIndex(c => c.id === targetId);
      if (existingIndex !== -1) {
        spConfigs[existingIndex] = tempConfig;
      } else {
        spConfigs.push(tempConfig);
      }
      await chrome.storage.local.set({ sp_configs: spConfigs });

      try {
        const items = await syncLevel1(targetId);
        
        // 同步成功，正式保存
        if (!currentConfigId || editingConfigId === null) {
          // 如果没有激活的，或者添加了新配置，自动将其设为激活
          currentConfigId = targetId;
        }
        await chrome.storage.local.set({
          sp_configs: spConfigs,
          current_config_id: currentConfigId
        });

        showStatus('success', '同步成功！', `已成功连接至 "${configName}" 文档库并同步了 ${items.length} 个 1 级项目！`);
        resetForm();
        renderConfigsList();
      } catch (syncErr) {
        console.error('Initial sync failed:', syncErr);
        // 同步失败，回滚
        spConfigs = backupConfigs;
        await chrome.storage.local.set({ sp_configs: spConfigs });
        showStatus('error', '连接失败', `连接或同步失败: ${syncErr.message || syncErr}。请确保文档库名称正确，且已登录网页版。`);
      }
    } catch (err) {
      console.error(err);
      showStatus('error', '配置异常', `发生错误: ${err.message || err}`);
    }
  });

  // 动态申请 Host 权限函数
  function requestHostPermission(origin) {
    return new Promise((resolve) => {
      chrome.permissions.request({
        origins: [origin + '/*']
      }, (granted) => {
        resolve(granted);
      });
    });
  }

  // 辅助函数：显示测试状态
  function showStatus(type, title, desc) {
    statusContainer.className = 'status-panel ' + type;
    statusTitle.innerText = title;
    statusDesc.innerText = desc;
  }
});
