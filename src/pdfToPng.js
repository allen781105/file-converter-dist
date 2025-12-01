const fs = require('fs');
const path = require('path');
const { pdf } = require('pdf-to-img');
const { mergePngImages, mergeGridImages } = require('./pptToPng');

/**
 * 将PDF文件转换为PNG图片
 * @param {string} pdfPath - PDF文件路径
 * @param {string} outputDir - 输出目录
 * @param {Object} options - 转换选项
 * @param {number} options.scale - 图片缩放比例，默认2.0
 * @returns {Promise<string[]>} - 返回生成的PNG文件路径数组
 */
async function convertPdfToPng(pdfPath, outputDir, options = {}) {
  const { scale = 2.0 } = options;
  const outputFiles = [];
  
  try {
    // 确保输出目录存在
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    console.log(`开始转换 PDF: ${pdfPath}`);
    console.log(`输出目录: ${outputDir}`);
    
    // 将PDF的每一页转换为PNG
    console.log('正在将 PDF 转换为 PNG 图片...');
    
    // 临时禁用 console.warn 以隐藏可能出现的警告
    const originalWarn = console.warn;
    console.warn = () => {};
    
    const document = await pdf(pdfPath, { scale });
    
    let pageIndex = 1;
    for await (const image of document) {
      const outputPath = path.join(outputDir, `page-${pageIndex}.png`);
      fs.writeFileSync(outputPath, image);
      outputFiles.push(outputPath);
      console.log(`已生成第 ${pageIndex} 页: ${outputPath}`);
      pageIndex++;
    }
    
    console.log(`转换完成！共生成 ${outputFiles.length} 张图片`);
    
    // 恢复 console.warn
    console.warn = originalWarn;
    
    return outputFiles;
    
  } catch (error) {
    console.error('转换PDF时出错:', error);
    throw error;
  }
}

/**
 * 转换PDF并按组合并
 * @param {string} pdfPath - PDF文件路径
 * @param {string} outputDir - 输出目录
 * @param {Object} options - 选项
 * @returns {Promise<Object>} - 返回单张图片和合并图片
 */
async function convertAndMergePdf(pdfPath, outputDir, options = {}) {
  const { 
    scale = 2.0, 
    mergeCount = 0, 
    spacing = 20, 
    aspectRatio = null,
    outputMode = 'long'
  } = options;
  
  try {
    // 1. 先转换为单张PNG
    const singleImages = await convertPdfToPng(pdfPath, outputDir, { scale });
    
    // 2. 如果是单张输出模式，直接返回
    if (outputMode === 'single') {
      return {
        singleImages,
        mergedImages: []
      };
    }
    
    // 3. 根据输出模式进行合并
    const mergedImages = [];
    
    if (outputMode === 'grid-2x2') {
      // 四宫格模式：每4张图片合并为2x2网格
      console.log('使用四宫格(2x2)模式...');
      for (let i = 0; i < singleImages.length; i += 4) {
        const group = singleImages.slice(i, i + 4);
        
        if (group.length > 0) {
          const mergedFileName = `grid-2x2-${Math.floor(i / 4) + 1}.png`;
          const mergedPath = path.join(outputDir, mergedFileName);
          
          await mergeGridImages(group, mergedPath, { 
            spacing, 
            gridCols: 2,
            backgroundColor: '#ffffff'
          });
          mergedImages.push(mergedPath);
        }
      }
    } else if (outputMode === 'grid-3x3') {
      // 九宫格模式：每9张图片合并为3x3网格
      console.log('使用九宫格(3x3)模式...');
      for (let i = 0; i < singleImages.length; i += 9) {
        const group = singleImages.slice(i, i + 9);
        
        if (group.length > 0) {
          const mergedFileName = `grid-3x3-${Math.floor(i / 9) + 1}.png`;
          const mergedPath = path.join(outputDir, mergedFileName);
          
          await mergeGridImages(group, mergedPath, { 
            spacing, 
            gridCols: 3,
            backgroundColor: '#ffffff'
          });
          mergedImages.push(mergedPath);
        }
      }
    } else if (outputMode === 'long') {
      // 长图模式：按指定数量纵向合并
      console.log(`使用长图模式，每${mergeCount}张合并...`);
      if (!mergeCount || mergeCount <= 1) {
        return {
          singleImages,
          mergedImages: []
        };
      }
      
      for (let i = 0; i < singleImages.length; i += mergeCount) {
        const group = singleImages.slice(i, i + mergeCount);
        
        if (group.length > 0) {
          const mergedFileName = `merged-${Math.floor(i / mergeCount) + 1}.png`;
          const mergedPath = path.join(outputDir, mergedFileName);
          
          await mergePngImages(group, mergedPath, { spacing, aspectRatio });
          mergedImages.push(mergedPath);
        }
      }
    }
    
    console.log(`✓ 共生成 ${mergedImages.length} 张合并图片`);
    
    return {
      singleImages,
      mergedImages
    };
    
  } catch (error) {
    console.error('转换并合并失败:', error);
    throw error;
  }
}

module.exports = {
  convertPdfToPng,
  convertAndMergePdf
};

