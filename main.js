const { app, BrowserWindow } = require('electron');

function createWindow () {

  const win = new BrowserWindow({
    width: 800,
    height: 600,
    // resizable: false,
    // frame: false,
    webPreferences: {
      nodeIntegration: true
    }
  });

  win.removeMenu(); // Zaten bir boka yaramıyordu.
  // win.webContents.openDevTools();
  win.loadFile('app/index.html');

}

app.whenReady().then(createWindow);

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

app.on('window-all-closed', () => {
  app.quit();
});
