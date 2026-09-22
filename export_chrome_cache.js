/**
 * SharePoint Map 缓存导出脚本
 * 在Chrome浏览器的Console中运行此脚本以导出缓存数据
 * 
 * 使用方法：
 * 1. 打开Chrome浏览器
 * 2. 访问 chrome://extensions/
 * 3. 找到SharePoint Map扩展
 * 4. 点击"详情" -> "检查视图"
 * 5. 在Console中运行此脚本
 */

async function exportSharePointCache() {
  console.log('开始导出SharePoint Map缓存数据...');
  
  try {
    // 获取所有存储数据
    const data = await chrome.storage.local.get(null);
    console.log('获取到存储数据，共有 ' + Object.keys(data).length + ' 个键');
    
    // 创建导出数据结构（复用共享逻辑：仅导出 /sites/IVPT/IVPProjects，URL 已转义）
    const exportData = (typeof buildCacheExportData === 'function')
      ? await buildCacheExportData()
      : {
          exportTime: new Date().toISOString(),
          extensionVersion: chrome.runtime.getManifest().version,
          caches: {},
          configs: data.sp_configs || [],
          currentConfigId: data.current_config_id || ''
        };
    
    // 提取缓存数据（仅在共享逻辑不可用时回退到原始提取）
    let cacheCount = 0;
    if (typeof buildCacheExportData !== 'function') {
      for (const [key, value] of Object.entries(data)) {
        if (key.startsWith('l1_cache') || 
            key.startsWith('subtree_cache') || 
            key.startsWith('favorites') ||
            key === 'sp_config') {
          exportData.caches[key] = value;
          cacheCount++;
        }
      }
    } else {
      cacheCount = Object.keys(exportData.caches).length;
    }
    
    console.log(`提取了 ${cacheCount} 个缓存项（仅限 /sites/IVPT/IVPProjects 收藏目录）`);
    
    // 分析缓存结构，为VBA做准备
    console.log('分析缓存结构...');
    analyzeCacheStructure(exportData.caches);
    
    // 转换为JSON字符串
    const jsonStr = JSON.stringify(exportData, null, 2);
    
    // 优先使用 chrome.downloads 固定文件名并覆盖旧文件
    if (typeof chrome !== 'undefined' && chrome.downloads) {
      const dataUrl = 'data:application/json;charset=utf-8,' + encodeURIComponent(jsonStr);
      chrome.downloads.download({
        url: dataUrl,
        filename: 'my app\\sharepointmap\\exported_cache.json',
        conflictAction: 'overwrite',
        saveAs: false
      }, (downloadId) => {
        if (chrome.runtime.lastError) {
          console.error('导出失败:', chrome.runtime.lastError.message);
        } else {
          console.log('缓存数据导出完成！已覆盖 C:\\Users\\xin.zhou\\Downloads\\my app\\sharepointmap\\exported_cache.json');
        }
      });
    } else {
      // 回退：普通浏览器页面使用链接下载（无法强制覆盖）
      const blob = new Blob([jsonStr], {type: 'application/json'});
      const url = URL.createObjectURL(blob);

      const a = document.createElement('a');
      a.href = url;
      a.download = `sharepoint_cache_${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      console.log('缓存数据导出完成！文件已下载。');
    }
    
    console.log('请将下载的JSON文件保存到指定位置供VBA使用。');
    
    // 显示导出路径建议
    console.log('导出文件名: exported_cache.json（固定文件名，每次导出覆盖旧文件）');
    console.log('导出路径: C:\\Users\\xin.zhou\\Downloads\\my app\\sharepointmap\\exported_cache.json');
    
    // 显示缓存统计信息
    showCacheStatistics(exportData.caches);
    
  } catch (error) {
    console.error('导出失败:', error);
  }
}

/**
 * 分析缓存结构，帮助VBA程序理解数据
 */
function analyzeCacheStructure(caches) {
  console.log('=== 缓存结构分析 ===');
  
  let l1CacheCount = 0;
  let subtreeCacheCount = 0;
  let totalItems = 0;
  
  for (const [key, value] of Object.entries(caches)) {
    console.log(`缓存键: ${key}`);
    
    if (key.startsWith('l1_cache')) {
      l1CacheCount++;
      if (value && value.items) {
        console.log(`  - 包含 ${value.items.length} 个一级项目`);
        totalItems += value.items.length;
        
        // 显示前几个项目的名称
        const sampleItems = value.items.slice(0, 5);
        sampleItems.forEach((item, index) => {
          console.log(`  - 项目 ${index + 1}: ${item.name} (${item.type === 1 ? '文件夹' : '文件'})`);
        });
      }
    } else if (key.startsWith('subtree_cache')) {
      subtreeCacheCount++;
      if (value && value.items) {
        console.log(`  - 包含 ${value.items.length} 个子树项目`);
        totalItems += value.items.length;
      }
    }
  }
  
  console.log(`=== 统计 ===`);
  console.log(`一级缓存数: ${l1CacheCount}`);
  console.log(`子树缓存数: ${subtreeCacheCount}`);
  console.log(`总项目数: ${totalItems}`);
}

/**
 * 显示缓存统计信息
 */
function showCacheStatistics(caches) {
  console.log('=== 缓存统计信息 ===');
  
  const stats = {
    totalCaches: Object.keys(caches).length,
    l1Caches: 0,
    subtreeCaches: 0,
    favoriteCaches: 0,
    otherCaches: 0,
    itemCount: 0
  };
  
  // 分析每个缓存
  for (const [key, value] of Object.entries(caches)) {
    if (key.startsWith('l1_cache')) {
      stats.l1Caches++;
    } else if (key.startsWith('subtree_cache')) {
      stats.subtreeCaches++;
    } else if (key.startsWith('favorites')) {
      stats.favoriteCaches++;
    } else {
      stats.otherCaches++;
    }
    
    // 计数项目
    if (value && value.items && Array.isArray(value.items)) {
      stats.itemCount += value.items.length;
    }
  }
  
  console.log(`总缓存数: ${stats.totalCaches}`);
  console.log(`一级缓存: ${stats.l1Caches}`);
  console.log(`子树缓存: ${stats.subtreeCaches}`);
  console.log(`收藏缓存: ${stats.favoriteCaches}`);
  console.log(`其他缓存: ${stats.otherCaches}`);
  console.log(`总项目数: ${stats.itemCount}`);
  
  // 显示可能包含study ID的文件夹示例
  console.log('\n=== 搜索建议 ===');
  console.log('在VBA中搜索时，可以查找包含以下关键词的文件夹:');
  console.log('- Protocol');
  console.log('- Amendment'); 
  console.log('- 协议');
  console.log('- 修订');
  console.log('- Study ID模式 (如: STUDY-12345, ST12345, etc.)');
}

/**
 * 快速搜索缓存中的study ID
 */
function quickSearchStudyID(studyID) {
  console.log(`快速搜索 Study ID: ${studyID}`);
  
  const results = [];
  
  // 搜索所有缓存
  for (const [key, value] of Object.entries(caches)) {
    if (value && value.items && Array.isArray(value.items)) {
      value.items.forEach(item => {
        if (item.name && item.name.includes(studyID)) {
          results.push({
            cacheKey: key,
            itemName: item.name,
            itemType: item.type === 1 ? '文件夹' : '文件',
            webUrl: item.webUrl || '无URL',
            serverRelativeUrl: item.serverRelativeUrl || '无相对路径'
          });
        }
      });
    }
  }
  
  console.log(`找到 ${results.length} 个匹配项:`);
  results.forEach((result, index) => {
    console.log(`${index + 1}. ${result.itemName} (${result.itemType})`);
    console.log(`   URL: ${result.webUrl}`);
  });
  
  return results;
}

/**
 * 导出指定配置的缓存
 */
async function exportSpecificConfig(configId) {
  console.log(`导出配置 ${configId} 的缓存数据...`);
  
  const cacheKeys = [
    `l1_cache_${configId}`,
    `favorites_${configId}`
  ];
  
  // 查找所有子树缓存
  const allData = await chrome.storage.local.get(null);
  for (const key in allData) {
    if (key.startsWith('subtree_cache_')) {
      cacheKeys.push(key);
    }
  }
  
  const specificData = await chrome.storage.local.get(cacheKeys);
  
  const exportData = {
    exportTime: new Date().toISOString(),
    configId: configId,
    cacheKeys: cacheKeys.filter(key => specificData[key]),
    data: specificData
  };
  
  const jsonStr = JSON.stringify(exportData, null, 2);
  const blob = new Blob([jsonStr], {type: 'application/json'});
  const url = URL.createObjectURL(blob);
  
  const a = document.createElement('a');
  a.href = url;
  a.download = `sharepoint_cache_config_${configId}_${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  
  console.log(`配置 ${configId} 的缓存导出完成！`);
}

// 导出函数供外部调用
window.exportSharePointCache = exportSharePointCache;
window.quickSearchStudyID = quickSearchStudyID;
window.exportSpecificConfig = exportSpecificConfig;

console.log('SharePoint Map缓存导出脚本已加载。');
console.log('可用函数:');
console.log('1. exportSharePointCache() - 导出所有缓存数据');
console.log('2. quickSearchStudyID("YOUR_STUDY_ID") - 快速搜索指定Study ID');
console.log('3. exportSpecificConfig("config_id") - 导出指定配置的缓存');

// 自动执行导出（可选，取消注释以自动运行）
// exportSharePointCache();