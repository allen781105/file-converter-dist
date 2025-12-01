# PPT/PDF to PNG Converter Service

这是一个独立的微服务项目，专门用于将 PPT 和 PDF 文件转换为高质量的 PNG 图片。

## 👀 效果预览

![首页预览](screenshots/preview.png)

## 🌟 功能特性

- **PPT 转图片**：将 .pptx 文件的每一页转换为图片（依赖 LibreOffice）
- **PDF 转图片**：将 .pdf 文件的每一页转换为图片
- **多种输出模式**：
  - 🖼️ 单张模式：每页单独输出
  - 📜 长图模式：智能拼接为长图（支持自定义数量、间距）
  - 🔲 宫格模式：2x2 或 3x3 智能拼图
- **📦 批量下载**：一键打包 ZIP 下载所有结果
- **✨ 易用界面**：内置简洁美观的 Web 操作界面

---

## 🚀 快速开始

### 前置要求

在运行本项目之前，请确保你的电脑上已安装：
1. **Node.js** (v18 或更高版本)
2. **LibreOffice** (用于 PPT 转换)

### 第一步：安装 LibreOffice

本项目依赖 LibreOffice 的命令行工具进行 PPT 到 PDF 的转换。请根据你的操作系统进行安装：

**🍎 macOS**:
```bash
brew install --cask libreoffice
```

**🐧 Ubuntu / Debian**:
```bash
sudo apt-get update
sudo apt-get install libreoffice fonts-wqy-microhei fonts-noto-cjk
# (推荐安装中文字体以防止转换乱码)
```

**🪟 Windows**:
1. 访问 [LibreOffice 官网](https://www.libreoffice.org/download/download-libreoffice/) 下载安装包并安装。
2. **关键步骤**：确保将 `soffice.exe` 的路径添加到系统的环境变量 `PATH` 中（通常在 `C:\Program Files\LibreOffice\program`）。

---

### 第二步：安装项目依赖

下载本项目后，在项目根目录下运行：

```bash
# 安装所有依赖
npm install
```

### 第三步：启动服务

```bash
# 启动服务器
npm start
```

服务启动后，浏览器会自动打开或请手动访问：
👉 [http://localhost:3000](http://localhost:3000)

---

## 📂 目录结构

```
├── src/                # 核心逻辑
│   ├── pptToPng.js     # PPT 转换与图片合并算法
│   └── pdfToPng.js     # PDF 转换逻辑
├── public/             # 前端界面
├── uploads/            # 上传文件临时存储
├── outputs/            # 转换结果存储
├── server.js           # 服务入口
└── package.json
```

## 🛠️ API 文档

### 1. PPT 转 PNG
`POST /api/ppt-to-png`
- `file`: .pptx 文件
- `outputMode`: `single` | `long` | `grid-2x2` | `grid-3x3`
- `scale`: 图片缩放倍率 (默认 2.0)

### 2. PDF 转 PNG
`POST /api/pdf-to-png`
- `file`: .pdf 文件
- 参数同上

### 3. 批量下载
`POST /api/download-zip`
- `outputDir`: 转换任务 ID 目录
