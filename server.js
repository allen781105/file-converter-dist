require('dotenv').config();
const path = require('path');
const fs = require('fs');
const express = require('express');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');
const dayjs = require('dayjs');
const JSZip = require('jszip');
const { convertPptToPng, convertAndMergePpt } = require('./src/pptToPng');
const { convertPdfToPng, convertAndMergePdf } = require('./src/pdfToPng');

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT = __dirname;
const UPLOADS_DIR = path.join(ROOT, 'uploads');
const OUTPUTS_DIR = path.join(ROOT, 'outputs');

// 确保目录存在
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
if (!fs.existsSync(OUTPUTS_DIR)) fs.mkdirSync(OUTPUTS_DIR, { recursive: true });

// 静态文件服务
app.use(express.static(path.join(ROOT, 'public')));
app.use('/outputs', express.static(OUTPUTS_DIR));
app.use('/uploads', express.static(UPLOADS_DIR));
app.use(express.json());

// 首页重定向
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Multer存储设置
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => {
    const id = uuidv4();
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, `${dayjs().format('YYYYMMDD-HHmmss')}-${id}${ext}`);
  }
});

// PPT文件上传配置
const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    if (file.originalname.toLowerCase().endsWith('.pptx')) cb(null, true);
    else cb(new Error('只支持 .pptx 文件'));
  },
  limits: { fileSize: 200 * 1024 * 1024 } // 200MB
});

// PDF文件上传配置
const pdfUpload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    if (file.originalname.toLowerCase().endsWith('.pdf')) cb(null, true);
    else cb(new Error('只支持 .pdf 文件'));
  },
  limits: { fileSize: 200 * 1024 * 1024 } // 200MB
});

// PPT转PNG接口
app.post('/api/ppt-to-png', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '未提供PPT文件' });
    }
    
    // 获取参数
    const scale = req.body.scale ? parseFloat(req.body.scale) : 2.0;
    const outputMode = req.body.outputMode || 'single';
    const mergeCount = req.body.mergeCount ? parseInt(req.body.mergeCount) : 0;
    const spacing = req.body.spacing ? parseInt(req.body.spacing) : 20;
    const aspectRatio = req.body.aspectRatio || null;
    
    // 创建输出目录
    const outputId = uuidv4();
    const outputDir = path.join(OUTPUTS_DIR, `png-${outputId}`);
    
    console.log(`开始转换 PPT 到 PNG: ${req.file.originalname}`);
    console.log(`输出模式: ${outputMode}`);
    
    let result;
    
    if (outputMode === 'single') {
      // 单张输出模式
      const imageFiles = await convertPptToPng(req.file.path, outputDir, { scale });
      result = {
        singleImages: imageFiles,
        mergedImages: []
      };
    } else {
      // 其他模式：转换并合并
      result = await convertAndMergePpt(req.file.path, outputDir, { 
        scale, 
        outputMode,
        mergeCount,
        spacing,
        aspectRatio
      });
    }
    
    // 生成可访问的URL列表
    const singleImageUrls = result.singleImages.map(filePath => {
      const relativePath = path.relative(OUTPUTS_DIR, filePath);
      return `/outputs/${relativePath}`;
    });
    
    const mergedImageUrls = result.mergedImages.map(filePath => {
      const relativePath = path.relative(OUTPUTS_DIR, filePath);
      return `/outputs/${relativePath}`;
    });
    
    // 清理临时文件
    fs.unlink(req.file.path, () => {});
    
    console.log(`✓ PPT转PNG完成`);
    console.log(`  - 单张图片: ${result.singleImages.length} 张`);
    console.log(`  - 合并图片: ${result.mergedImages.length} 张`);
    
    res.json({
      success: true,
      singleImages: singleImageUrls,
      mergedImages: mergedImageUrls,
      singleCount: result.singleImages.length,
      mergedCount: result.mergedImages.length,
      outputDir: `png-${outputId}`
    });
  } catch (err) {
    console.error('PPT转PNG失败:', err);
    res.status(500).json({ error: err.message || 'PPT转PNG失败' });
  }
});

// PDF转PNG接口
app.post('/api/pdf-to-png', pdfUpload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '未提供PDF文件' });
    }
    
    // 获取参数
    const scale = req.body.scale ? parseFloat(req.body.scale) : 2.0;
    const outputMode = req.body.outputMode || 'single';
    const mergeCount = req.body.mergeCount ? parseInt(req.body.mergeCount) : 0;
    const spacing = req.body.spacing ? parseInt(req.body.spacing) : 20;
    const aspectRatio = req.body.aspectRatio || null;
    
    // 创建输出目录
    const outputId = uuidv4();
    const outputDir = path.join(OUTPUTS_DIR, `pdf-png-${outputId}`);
    
    console.log(`开始转换 PDF 到 PNG: ${req.file.originalname}`);
    console.log(`输出模式: ${outputMode}`);
    
    let result;
    
    if (outputMode === 'single') {
      // 单张输出模式
      const imageFiles = await convertPdfToPng(req.file.path, outputDir, { scale });
      result = {
        singleImages: imageFiles,
        mergedImages: []
      };
    } else {
      // 其他模式：转换并合并
      result = await convertAndMergePdf(req.file.path, outputDir, { 
        scale, 
        outputMode,
        mergeCount,
        spacing,
        aspectRatio
      });
    }
    
    // 生成可访问的URL列表
    const singleImageUrls = result.singleImages.map(filePath => {
      const relativePath = path.relative(OUTPUTS_DIR, filePath);
      return `/outputs/${relativePath}`;
    });
    
    const mergedImageUrls = result.mergedImages.map(filePath => {
      const relativePath = path.relative(OUTPUTS_DIR, filePath);
      return `/outputs/${relativePath}`;
    });
    
    // 清理临时文件
    fs.unlink(req.file.path, () => {});
    
    console.log(`✓ PDF转PNG完成`);
    console.log(`  - 单张图片: ${result.singleImages.length} 张`);
    console.log(`  - 合并图片: ${result.mergedImages.length} 张`);
    
    res.json({
      success: true,
      singleImages: singleImageUrls,
      mergedImages: mergedImageUrls,
      singleCount: result.singleImages.length,
      mergedCount: result.mergedImages.length,
      outputDir: `pdf-png-${outputId}`
    });
  } catch (err) {
    console.error('PDF转PNG失败:', err);
    res.status(500).json({ error: err.message || 'PDF转PNG失败' });
  }
});

// 批量打包下载接口
app.post('/api/download-zip', async (req, res) => {
  try {
    const { outputDir } = req.body;
    
    if (!outputDir) {
      return res.status(400).json({ error: '缺少outputDir参数' });
    }
    
    // 安全检查：防止路径遍历
    if (outputDir.includes('..') || outputDir.includes('/') || outputDir.includes('\\')) {
      return res.status(400).json({ error: '无效的目录名称' });
    }
    
    const sourceDir = path.join(OUTPUTS_DIR, outputDir);
    if (!fs.existsSync(sourceDir)) {
      return res.status(404).json({ error: '目录不存在' });
    }
    
    const zipFileName = `${outputDir}.zip`;
    const zipFilePath = path.join(OUTPUTS_DIR, zipFileName);
    
    // 如果zip文件已存在，直接返回
    if (fs.existsSync(zipFilePath)) {
      return res.json({ 
        success: true, 
        downloadUrl: `/outputs/${zipFileName}` 
      });
    }
    
    console.log(`开始打包目录: ${sourceDir}`);
    
    const zip = new JSZip();
    const files = fs.readdirSync(sourceDir);
    
    let fileCount = 0;
    for (const file of files) {
      // 只打包图片文件
      if (file.toLowerCase().endsWith('.png') || file.toLowerCase().endsWith('.jpg')) {
        const content = fs.readFileSync(path.join(sourceDir, file));
        zip.file(file, content);
        fileCount++;
      }
    }
    
    if (fileCount === 0) {
      return res.status(400).json({ error: '目录中没有可打包的图片' });
    }
    
    // 生成zip文件
    const content = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(zipFilePath, content);
    
    console.log(`✓ 打包完成: ${zipFileName} (${fileCount} 个文件)`);
    
    res.json({ 
      success: true, 
      downloadUrl: `/outputs/${zipFileName}` 
    });
    
  } catch (err) {
    console.error('打包下载失败:', err);
    res.status(500).json({ error: err.message || '打包下载失败' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running: http://localhost:${PORT}`);
});

