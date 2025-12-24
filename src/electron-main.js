const { app, BrowserWindow } = require('electron');
const path = require('path');

// 引入 server.js，这会自动启动 Express 服务
// 注意：server.js 结尾有 app.listen，所以 require 时就会启动
require('../server.js');

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false, // 简化集成，如果是生产环境通常建议开启隔离
    },
  });

  // 等待一点时间确保服务器启动
  setTimeout(() => {
    win.loadURL('http://localhost:3000');
  }, 1000);

  // 开发环境打开控制台
  // win.webContents.openDevTools();
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
