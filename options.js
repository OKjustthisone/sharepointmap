// options.js

document.addEventListener('DOMContentLoaded', () => {
  const siteUrlInput = document.getElementById('siteUrl');
  const libraryNameInput = document.getElementById('libraryName');
  const saveBtn = document.getElementById('saveBtn');
  const statusContainer = document.getElementById('statusContainer');
  const statusTitle = document.getElementById('statusTitle');
  const statusDesc = document.getElementById('statusDesc');

  // 1. 初始化读取已保存的配置
  chrome.storage.local.get('sp_config', (data) => {
    if (data.sp_config) {
      siteUrlInput.value = data.sp_config.siteUrl || '';
      libraryNameInput.value = data.sp_config.libraryName || 'Shared Documents';
    }
  });

  // 2. 保存并测试按钮事件
  saveBtn.addEventListener('click', async () => {
    let siteUrl = siteUrlInput.value.trim();
    const libraryName = libraryNameInput.value.trim();

    if (!siteUrl || !libraryName) {
      showStatus('error', '保存失败', '请完整填写 SharePoint 站点 URL 和文档库名称。');
      return;
    }

    // 格式化 siteUrl，移除末尾斜杠
    if (siteUrl.endsWith('/')) {
      siteUrl = siteUrl.slice(0, -1);
    }

    try {
      showStatus('testing', '正在配置权限...', '正在请求对该站点的访问权限，请在弹窗中确认（如适用）...');

      // 提取站点的 Origin (域名)，用于动态申请 Host 权限
      const parsedUrl = new URL(siteUrl);
      const siteOrigin = parsedUrl.origin;

      // 动态申请主机权限，绕过跨域限制
      const permissionGranted = await requestHostPermission(siteOrigin);
      if (!permissionGranted) {
        showStatus('error', '权限被拒绝', '未能获得该域名的访问权限，无法发起 API 连接测试。');
        return;
      }

      showStatus('testing', '正在测试连接并拉取缓存...', '正在连接 SharePoint 并同步首屏数据，请稍候...');

      // 临时保存配置以供 syncLevel1 读取
      const config = { siteUrl, libraryName, siteOrigin };
      const previousData = await chrome.storage.local.get('sp_config');
      await chrome.storage.local.set({ sp_config: config });

      try {
        const items = await syncLevel1();
        showStatus('success', '连接并同步成功！', `已成功连接至 SharePoint 文档库并同步了 ${items.length} 个 1 级项目！`);
      } catch (syncErr) {
        console.error('Initial sync failed:', syncErr);
        // 回滚配置
        if (previousData.sp_config) {
          await chrome.storage.local.set({ sp_config: previousData.sp_config });
        } else {
          await chrome.storage.local.remove('sp_config');
        }
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
      // 动态申请权限：形如 "https://company.sharepoint.com/*"
      chrome.permissions.request({
        origins: [origin + '/*']
      }, (granted) => {
        resolve(granted);
      });
    });
  }

  // 测试 SharePoint 连通性
  async function testSharePointConnection(siteUrl, libraryName) {
    // 拼接测试接口获取根目录下目录信息
    const testApi = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(libraryName)}')/RootFolder/Folders`;
    try {
      const response = await fetch(testApi, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json'
        },
        credentials: 'include' // 关键：携带当前浏览器的登录 Cookie 会话
      });

      if (response.ok) {
        return true;
      } else {
        console.warn('API connection test response not OK:', response.status, response.statusText);
        return false;
      }
    } catch (error) {
      console.error('API connection fetch failed:', error);
      return false;
    }
  }

  // 辅助函数：显示测试状态
  function showStatus(type, title, desc) {
    statusContainer.className = 'status-panel ' + type;
    statusTitle.innerText = title;
    statusDesc.innerText = desc;
  }
});
