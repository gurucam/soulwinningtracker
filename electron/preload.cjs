const { contextBridge, ipcRenderer } = require("electron")
const fs = require("fs")
const path = require("path")

const resolveAppVersion = () => {
  if (process.env.npm_package_version) return process.env.npm_package_version
  try {
    const packagePath = path.join(__dirname, "..", "package.json")
    const raw = fs.readFileSync(packagePath, "utf8")
    const parsed = JSON.parse(raw)
    if (parsed && typeof parsed.version === "string") {
      return parsed.version
    }
  } catch {
    // ignore
  }
  return "0.0.0"
}

contextBridge.exposeInMainWorld("soulwinning", {
  version: resolveAppVersion(),
  checkForUpdates: () => ipcRenderer.invoke("update:check"),
  installUpdate: () => ipcRenderer.invoke("update:install"),
  openExternal: (url) => ipcRenderer.invoke("open-external", url),
  getUserDataPath: () => ipcRenderer.invoke("app:getUserDataPath"),
  onUpdateStatus: (handler) => {
    if (typeof handler !== "function") return undefined
    const wrapped = (_event, payload) => handler(payload)
    ipcRenderer.on("update-status", wrapped)
    return () => ipcRenderer.off("update-status", wrapped)
  },
})
