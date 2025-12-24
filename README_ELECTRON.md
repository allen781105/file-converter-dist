# Electron Build Guide / Electron 构建指南

To run this application as a standalone desktop app or build it for Windows, follow these instructions.
要运行此应用程序或将其打包为 Windows 应用，请遵循以下说明。

## Prerequisites / 前置要求

This project uses `canvas` and `libreoffice-convert`. You **MUST** have the following installed:
本项目使用了 `canvas` 和 `libreoffice-convert`，您 **必须** 安装以下组件：

### 1. Prerequisites (Windows)
### 1. 前置条件 (Windows)

- **Node.js**: Install from [nodejs.org](https://nodejs.org/)
- **Node.js**: 请从 [nodejs.org](https://nodejs.org/) 安装
- **LibreOffice**: Install from [libreoffice.org/download](https://www.libreoffice.org/download/download-libreoffice/)
    - **Crucial**: The default installation path is usually `C:\Program Files\LibreOffice`. If you install it elsewhere, you may need to configure the environment variable.
- **LibreOffice**: 请从 [libreoffice.org/download](https://www.libreoffice.org/download/download-libreoffice/) 下载并安装
    - **关键**: 默认通常安装在 `C:\Program Files\LibreOffice`，程序会自动识别。如果安装在其他位置，可能需要配置环境变量。is in your PATH or the default installation location.
- **macOS**: `brew install --cask libreoffice`

### 2. System Libraries for Canvas (Important for macOS/Linux)
The PDF conversion uses `canvas`, which is a native module requiring system libraries.
PDF 转换使用 `canvas`，这是一个需要系统库的原生模块。

**macOS (via Homebrew):**
```bash
brew install pkg-config cairo pango libpng jpeg giflib librsvg
```

**Windows:**
Usually, prebuilt binaries are downloaded automatically. If you encounter issues, you may need to install the build tools:
通常会自动下载预编译的二进制文件。如果遇到问题，可能需要安装构建工具：
```bash
npm install --global --production windows-build-tools
```

## Running Locally / 本地运行

1. Install dependencies / 安装依赖:
   ```bash
   npm install
   ```
   *If this fails on `canvas` install, make sure you installed the prerequisites above.*
   *如果 `canvas` 安装失败，请确保安装了上述前置要求。*

2. Run in Development Mode / 开发模式运行:
   ```bash
   npm run electron:dev
   ```

## Building for Windows / 打包 Windows 应用

### On Windows / 在 Windows 上
Running the build on a Windows machine is the most reliable way.
在 Windows 机器上运行构建是最可靠的方法。

```bash
npm run electron:build
```
This will create an installer in the `dist` folder.
这将在 `dist` 文件夹中生成安装程序。

### On macOS / 在 macOS 上
You can build for Windows on macOS, but it requires `wine` for some features (like icon changing) and cross-compilation support.
您可以在 macOS 上构建 Windows 应用，但这可能需要安装 `wine` 并支持交叉编译。

1. Ensure dependencies are installed (see above).
2. Run build:
   ```bash
   npm run electron:build
   ```
   *Note: If `canvas` fails to rebuild for the target architecture, you might see errors. It is recommended to build on the target OS.*
   *注意：如果 `canvas` 无法为目标架构重新编译，可能会报错。建议在目标操作系统上进行构建。*

## Troubleshooting / 故障排除

**Error: `node-gyp failed to rebuild ... canvas`**
- **On macOS**: This often happens because the build tools cannot find the installed system libraries (`cairo`, etc.) even if they are installed.
- **Solution**: The most reliable way to build the Windows app is to **run the build command on a Windows machine**.
- **Dev Mode**: If `npm start` works but `npm run electron:dev` fails, you can develop using `npm start` (backend only).

**Important: Switching between Node.js and Electron**
**重要：在 Node.js 和 Electron 之间切换**
Because `canvas` is a native module, it must be compiled for the specific runtime you are using.
由于 `canvas` 是原生模块，必须为您使用的特定运行时进行编译。

- **To run `npm start` (Node.js)**: Run `npm run rebuild:node`
- **To run Electron**: Run `npm run rebuild` (compiles for Electron)

- **要运行 `npm start` (Node.js)**: 请运行 `npm run rebuild:node`
- **要运行 Electron**: 请运行 `npm run rebuild` (编译 Electron 版本)

**错误：`node-gyp failed to rebuild ... canvas`**
- **在 macOS 上**：这通常是因为构建工具找不到已安装的系统库（即使已安装）。
- **解决方案**：打包 Windows 应用最可靠的方法是 **在 Windows 机器上运行构建命令**。
- **开发模式**：如果 `npm start` 可以运行但 `npm run electron:dev` 失败，您可以使用 `npm start` 进行后端开发，或者尝试手动设置 `PKG_CONFIG_PATH`。

**Error: converting PPT fails**
- Fix: Ensure LibreOffice is installed.
- 修复方法：确保已安装 LibreOffice。
