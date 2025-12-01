const fs = require('fs');
const path = require('path');
const libre = require('libreoffice-convert');
const { promisify } = require('util');
const { pdf } = require('pdf-to-img');
const JSZip = require('jszip');
const { execSync } = require('child_process');

// 配置LibreOffice路径
libre.convertAsync = promisify(libre.convert);

// 配置LibreOffice选项（支持无头模式）
const LIBREOFFICE_OPTIONS = {
  // 强制使用无头模式，避免GUI弹出
  // 这对服务器部署至关重要
  sofficeOptions: [
    '--headless',
    '--invisible',
    '--nocrashreport',
    '--nodefault',
    '--nofirststartwizard',
    '--nolockcheck',
    '--nologo',
    '--norestore'
  ]
};

// 检测soffice路径（支持Mac和Linux）
function getSofficePath() {
  const possiblePaths = [
    // Linux路径（ECS服务器）
    '/usr/bin/soffice',
    '/usr/bin/libreoffice',
    // Mac路径
    '/usr/local/bin/soffice',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    // 通用
    'soffice',
    'libreoffice'
  ];
  
  for (const sofficePath of possiblePaths) {
    try {
      if (fs.existsSync(sofficePath)) {
        return sofficePath;
      }
    } catch (e) {
      // 继续尝试下一个路径
    }
  }
  
  // 尝试从命令行获取
  try {
    const result = execSync('which soffice', { encoding: 'utf8' }).trim();
    if (result) return result;
  } catch (e) {
    // which命令失败
  }
  
  return null;
}

const sofficePath = getSofficePath();
if (sofficePath) {
  console.log(`找到 LibreOffice: ${sofficePath}`);
  // 设置环境变量
  process.env.LIBRE_OFFICE_PATH = sofficePath;
  
  // 检测运行环境
  const isHeadless = !process.env.DISPLAY && process.platform === 'linux';
  if (isHeadless) {
    console.log('检测到无头环境（服务器模式），将使用 --headless 模式');
  } else {
    console.log('检测到图形环境（开发模式）');
  }
} else {
  console.warn('警告: 未找到 LibreOffice，PPT转换功能可能无法使用');
  console.warn('请安装 LibreOffice:');
  console.warn('  Mac: brew install --cask libreoffice');
  console.warn('  Ubuntu/Debian: sudo apt-get install -y libreoffice');
  console.warn('  CentOS/RHEL: sudo yum install -y libreoffice');
}

const libreConvert = promisify(libre.convert);

/**
 * 将PPT文件转换为PNG图片
 * @param {string} pptPath - PPT文件路径
 * @param {string} outputDir - 输出目录
 * @param {Object} options - 转换选项
 * @param {number} options.scale - 图片缩放比例，默认2.0
 * @returns {Promise<string[]>} - 返回生成的PNG文件路径数组
 */
async function convertPptToPng(pptPath, outputDir, options = {}) {
  const { scale = 2.0 } = options;
  const outputFiles = [];
  
  try {
    // 确保输出目录存在
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    console.log(`开始转换 PPT: ${pptPath}`);
    console.log(`输出目录: ${outputDir}`);
    
    // 读取PPT文件
    const pptBuffer = fs.readFileSync(pptPath);
    
    // 先将PPT转换为PDF（使用无头模式选项）
    console.log('步骤 1/2: 正在将 PPT 转换为 PDF...');
    
    // 在服务器环境下，LibreOffice会自动以无头模式运行
    // libreoffice-convert库会读取LIBRE_OFFICE_PATH环境变量
    const pdfBuffer = await libreConvert(pptBuffer, '.pdf', undefined);
    
    // 保存临时PDF文件
    const tempPdfPath = path.join(outputDir, 'temp.pdf');
    fs.writeFileSync(tempPdfPath, pdfBuffer);
    console.log('PDF 转换完成');
    
    // 将PDF的每一页转换为PNG
    console.log('步骤 2/2: 正在将 PDF 转换为 PNG 图片...');
    
    // 临时禁用 console.warn 以隐藏 "Knockout groups not supported" 警告
    const originalWarn = console.warn;
    console.warn = () => {};
    
    const document = await pdf(tempPdfPath, { scale });
    
    let pageIndex = 1;
    for await (const image of document) {
      const outputPath = path.join(outputDir, `slide-${pageIndex}.png`);
      fs.writeFileSync(outputPath, image);
      outputFiles.push(outputPath);
      console.log(`已生成第 ${pageIndex} 页: ${outputPath}`);
      pageIndex++;
    }
    
    console.log(`转换完成！共生成 ${outputFiles.length} 张图片`);
    
    // 恢复 console.warn
    console.warn = originalWarn;
    
    // 清理临时PDF文件
    if (fs.existsSync(tempPdfPath)) {
      fs.unlinkSync(tempPdfPath);
      console.log('已清理临时文件');
    }
    
    return outputFiles;
    
  } catch (error) {
    console.error('转换PPT时出错:', error);
    throw error;
  }
}

/**
 * 使用另一种方法：直接从PPTX提取图片和渲染
 * 这个方法适用于已有截图或预览图的情况
 */
async function extractPptxSlides(pptxPath, outputDir) {
  try {
    // 确保输出目录存在
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    console.log(`开始处理 PPTX: ${pptxPath}`);
    
    // 读取PPTX文件
    const data = fs.readFileSync(pptxPath);
    const zip = await JSZip.loadAsync(data);
    
    // 获取幻灯片数量
    const slideFiles = Object.keys(zip.files).filter(name => 
      name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
    );
    
    console.log(`找到 ${slideFiles.length} 张幻灯片`);
    
    const slideInfo = [];
    for (let i = 0; i < slideFiles.length; i++) {
      slideInfo.push({
        index: i + 1,
        xmlPath: slideFiles[i]
      });
    }
    
    return slideInfo;
    
  } catch (error) {
    console.error('提取PPTX信息时出错:', error);
    throw error;
  }
}

/**
 * 合并多张PNG图片为网格布局
 * @param {string[]} imagePaths - 图片路径数组
 * @param {string} outputPath - 输出路径
 * @param {Object} options - 选项
 * @returns {Promise<string>} - 返回输出路径
 */
async function mergeGridImages(imagePaths, outputPath, options = {}) {
  const sharp = require('sharp');
  const { spacing = 0, backgroundColor = '#ffffff', gridCols = 2 } = options;
  const gridRows = Math.ceil(imagePaths.length / gridCols);
  
  try {
    console.log(`开始合并 ${imagePaths.length} 张图片为 ${gridRows}x${gridCols} 网格...`);
    
    // 读取所有图片的元数据
    const imageMetadata = await Promise.all(
      imagePaths.map(path => sharp(path).metadata())
    );
    
    // 计算每个格子的尺寸（使用最大宽高）
    const cellWidth = Math.max(...imageMetadata.map(meta => meta.width));
    const cellHeight = Math.max(...imageMetadata.map(meta => meta.height));
    

    // 计算画布总尺寸
    const canvasWidth = Math.ceil(cellWidth * gridCols + spacing * (gridCols - 1));
    const canvasHeight = Math.ceil(cellHeight * gridRows + spacing * (gridRows - 1));
    
    console.log(`网格尺寸: ${gridRows}行 x ${gridCols}列`);
    console.log(`每格尺寸: ${cellWidth} x ${cellHeight}`);
    console.log(`画布尺寸: ${canvasWidth} x ${canvasHeight}`);
    
    // 创建背景画布
    let canvas = sharp({
      create: {
        width: canvasWidth,
        height: canvasHeight,
        channels: 4,
        background: backgroundColor
      }
    });
    
    // 准备合成操作
    const compositeOperations = [];
    
    for (let i = 0; i < imagePaths.length; i++) {
      const imagePath = imagePaths[i];
      const row = Math.floor(i / gridCols);
      const col = i % gridCols;
      
      // 计算位置
      const x = col * (cellWidth + spacing);
      const y = row * (cellHeight + spacing);
      
      // 读取并调整图片
      const imageBuffer = await sharp(imagePath, { limitInputPixels: false })
        .resize(cellWidth, cellHeight, {
          fit: 'contain',
          background: backgroundColor
        })
        .toBuffer();
      
      compositeOperations.push({
        input: imageBuffer,
        top: Math.round(y),
        left: Math.round(x)
      });
      
      console.log(`添加第 ${i + 1} 张图片到 (行${row}, 列${col}) = X=${x}, Y=${y}`);
    }
    
    // 执行合成
    await canvas
      .composite(compositeOperations)
      .png()
      .toFile(outputPath);
    
    console.log(`✓ 网格合并完成: ${outputPath}`);
    
    return outputPath;
    
  } catch (error) {
    console.error('网格合并失败:', error);
    throw error;
  }
}

/**
 * 合并多张PNG图片为一张长图
 * @param {string[]} imagePaths - 图片路径数组
 * @param {string} outputPath - 输出路径
 * @param {Object} options - 选项
 * @returns {Promise<string>} - 返回输出路径
 */
async function mergePngImages(imagePaths, outputPath, options = {}) {
  const sharp = require('sharp');
  const { spacing = 0, backgroundColor = '#ffffff', aspectRatio = null } = options;
  
  try {
    console.log(`开始合并 ${imagePaths.length} 张图片...`);
    
    // 读取所有图片的元数据
    const imageMetadata = await Promise.all(
      imagePaths.map(path => sharp(path, { limitInputPixels: false }).metadata())
    );
    
    // 计算合并后的原始尺寸
    const maxWidth = Math.max(...imageMetadata.map(meta => meta.width));
    const totalHeight = imageMetadata.reduce((sum, meta) => sum + meta.height, 0) + 
                        (spacing * (imagePaths.length - 1));
    
    let canvasWidth = maxWidth;
    let canvasHeight = totalHeight;
    
    // 如果指定了长宽比，调整画布尺寸
    if (aspectRatio) {
      const [ratioWidth, ratioHeight] = aspectRatio.split(':').map(Number);
      const targetRatio = ratioWidth / ratioHeight;
      const currentRatio = maxWidth / totalHeight;
      
      console.log(`目标长宽比: ${aspectRatio} (${targetRatio.toFixed(2)})`);
      console.log(`当前长宽比: ${currentRatio.toFixed(2)}`);
      
      if (currentRatio > targetRatio) {
        // 当前太宽，需要增加高度
        canvasHeight = Math.round(maxWidth / targetRatio);
      } else if (currentRatio < targetRatio) {
        // 当前太高，需要增加宽度
        canvasWidth = Math.round(totalHeight * targetRatio);
      }
      
      console.log(`调整后画布尺寸: ${canvasWidth} x ${canvasHeight}`);
    }
    
    console.log(`合并图片尺寸: ${canvasWidth} x ${canvasHeight}`);
    
    // 创建背景画布
    let canvas = sharp({
      create: {
        width: canvasWidth,
        height: canvasHeight,
        channels: 4,
        background: backgroundColor
      }
    });
    
    // 准备合成操作
    const compositeOperations = [];
    
    // 计算内容的居中偏移
    const contentWidth = maxWidth;
    const contentHeight = totalHeight;
    const offsetX = Math.round((canvasWidth - contentWidth) / 2);
    const offsetY = Math.round((canvasHeight - contentHeight) / 2);
    
    console.log(`内容居中偏移: X=${offsetX}, Y=${offsetY}`);
    
    let currentY = offsetY;
    
    for (let i = 0; i < imagePaths.length; i++) {
      const imagePath = imagePaths[i];
      const imageBuffer = await sharp(imagePath, { limitInputPixels: false }).toBuffer();
      const imageWidth = imageMetadata[i].width;
      
      // 每张图片在水平方向也居中
      const imageOffsetX = offsetX + Math.round((maxWidth - imageWidth) / 2);
      
      compositeOperations.push({
        input: imageBuffer,
        top: Math.round(currentY),
        left: Math.round(imageOffsetX)
      });
      
      console.log(`添加第 ${i + 1} 张图片到 X=${imageOffsetX}, Y=${currentY}`);
      currentY += imageMetadata[i].height + spacing;
    }
    
    // 执行合成
    await canvas
      .composite(compositeOperations)
      .png()
      .toFile(outputPath);
    
    console.log(`✓ 合并完成: ${outputPath}`);
    
    return outputPath;
    
  } catch (error) {
    console.error('合并图片失败:', error);
    throw error;
  }
}

/**
 * 转换PPT并按组合并
 * @param {string} pptPath - PPT文件路径
 * @param {string} outputDir - 输出目录
 * @param {Object} options - 选项
 * @returns {Promise<Object>} - 返回单张图片和合并图片
 */
async function convertAndMergePpt(pptPath, outputDir, options = {}) {
  const { 
    scale = 2.0, 
    mergeCount = 0, 
    spacing = 20, 
    aspectRatio = null,
    outputMode = 'long'
  } = options;
  
  try {
    // 1. 先转换为单张PNG
    const singleImages = await convertPptToPng(pptPath, outputDir, { scale });
    
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
  convertPptToPng,
  extractPptxSlides,
  mergePngImages,
  mergeGridImages,
  convertAndMergePpt
};
