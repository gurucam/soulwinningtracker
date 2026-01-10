const { app, BrowserWindow, shell, ipcMain } = require("electron")
const { autoUpdater } = require("electron-updater")
const path = require("path")

const isDev = !app.isPackaged
let mainWindow

const sendUpdateStatus = (payload) => {
  if (!mainWindow) return
  mainWindow.webContents.send("update-status", payload)
}

const createWindow = () => {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 960,
    minHeight: 600,
    backgroundColor: "#F6EFE3",
    icon: path.join(__dirname, "..", "swt.png"),
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  mainWindow.setMenuBarVisibility(false)
  mainWindow.setMenu(null)

  if (isDev) {
    mainWindow.loadURL("http://localhost:5173")
  } else {
    mainWindow.loadFile(path.join(__dirname, "..", "dist", "index.html"))
  }

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url)
    return { action: "deny" }
  })
}

autoUpdater.autoDownload = true
autoUpdater.autoInstallOnAppQuit = true

autoUpdater.on("checking-for-update", () => sendUpdateStatus({ message: "Checking for updates..." }))
autoUpdater.on("update-available", () =>
  sendUpdateStatus({ message: "Update available. Downloading..." }),
)
autoUpdater.on("update-not-available", () => sendUpdateStatus({ message: "You are up to date." }))
autoUpdater.on("error", (err) =>
  sendUpdateStatus({ message: `Update error: ${err == null ? "unknown" : err.message}` }),
)
autoUpdater.on("download-progress", (progress) => {
  const pct = progress.percent.toFixed(0)
  sendUpdateStatus({ message: `Downloading update... ${pct}%` })
})
autoUpdater.on("update-downloaded", () => {
  sendUpdateStatus({ message: "Update downloaded. Restart to apply.", ready: true })
})

ipcMain.handle("update:check", async () => {
  if (!app.isPackaged) {
    sendUpdateStatus({ message: "Updates are available in the installed app." })
    return { ok: false, error: "not-packaged" }
  }
  try {
    await autoUpdater.checkForUpdates()
    return { ok: true }
  } catch (err) {
    sendUpdateStatus({ message: "Update check failed." })
    return { ok: false, error: String(err) }
  }
})

ipcMain.handle("update:install", () => {
  if (!app.isPackaged) return false
  autoUpdater.quitAndInstall()
  return true
})

ipcMain.handle("open-external", (_event, url) => {
  if (typeof url !== "string" || !url) return false
  shell.openExternal(url)
  return true
})

ipcMain.handle("app:getUserDataPath", () => {
  return app.getPath("userData")
})

app.whenReady().then(() => {
  createWindow()
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit()
  }
  mainWindow = null
})
