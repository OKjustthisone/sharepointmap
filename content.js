// content.js
// 注入到 SharePoint 页面中的内容脚本，运行在第一方上下文中。
// 负责接收插件前端发送的请求指令，以第一方身份发起 Fetch 请求并携带本地 Cookie，然后将 JSON 结果回传。

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'fetch_via_page') {
    const { url } = request;
    console.log('[SharePoint Map] Content script executing fetch for:', url);

    fetch(url, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json'
      }
    })
    .then(async (response) => {
      const text = await response.text();
      
      if (!response.ok) {
        // 将 HTTP 报错时的详细页面片段带回，便于排查
        const errorSnippet = text.substring(0, 200).replace(/</g, '&lt;').replace(/>/g, '&gt;');
        sendResponse({ 
          success: false, 
          error: `HTTP ${response.status} ${response.statusText}\n内容片段: ${errorSnippet}` 
        });
        return;
      }

      try {
        const data = JSON.parse(text);
        sendResponse({ success: true, data: data });
      } catch (e) {
        // 如果解析 JSON 失败，说明返回的不是 JSON（可能是重定向到了登录页、验证码页或报错页）
        // 抓取前 200 个字符返回给前端以精确定位问题
        const htmlSnippet = text.substring(0, 200).replace(/</g, '&lt;').replace(/>/g, '&gt;');
        sendResponse({ 
          success: false, 
          error: `接口未返回 JSON。可能被重定向到了登录页或安全校验页。\n返回内容前200字: ${htmlSnippet}` 
        });
      }
    })
    .catch((err) => {
      console.error('[SharePoint Map] Fetch error inside page:', err);
      sendResponse({ success: false, error: err.message || err.toString() });
    });

    return true; // 保持异步通道开启
  }
});
