const {
  app,
  BrowserWindow,
  ipcMain
} = require('electron')
const url = require("url");
const path = require("path");
const fs = require("fs");
const Excel = require('exceljs');

const cacheFile = "test.xlsx";

let appWindow

function initWindow() {
  appWindow = new BrowserWindow({
    width: 1140,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      devTools: true
    }
  });

  // Electron Build Path
  appWindow.loadURL(
    url.format({
      pathname: path.join(__dirname, `/dist/index.html`),
      protocol: "file:",
      slashes: false
    })
  );

  // Initialize the DevTools.
  appWindow.webContents.openDevTools()




  appWindow.on('closed', function () {
    appWindow = null
  })
}


app.on('ready', initWindow)

// Close when all windows are closed.
app.on('window-all-closed', function () {

  // On macOS specific close process
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', function () {

  if (appWindow === null) {
    initWindow()


  }
})

ipcMain.on('asynchronous-message', function (event, arg) {

  var workbook = new Excel.Workbook();
  workbook.xlsx.load(arg.buffer).then(()=>{
    workbook.xlsx.writeFile(cacheFile);
  })

  event.sender.send('asynchronous-reply', 'pong')
})


ipcMain.on('load-cache', function (event, arg) {
  if (fs.existsSync(cacheFile)) {
    const fileBuffer = fs.readFileSync(cacheFile);
    event.sender.send('send-cache-buffer', fileBuffer)
  }
})
