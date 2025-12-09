const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld("startup", {
  args: process.argv
});

contextBridge.exposeInMainWorld('api', {
  authCheck: (data) => ipcRenderer.send("auth-check", data),
  onAuthResult: (cb) => ipcRenderer.on("auth-result", (_e, res) => cb(res)),
  loadParams: () => ipcRenderer.invoke('params:load'),
  saveParams: (params) => ipcRenderer.send('params:save', params),
  openCSV: () => ipcRenderer.invoke('dialog:openCSV'),
  readFile: (p) => ipcRenderer.invoke('file:read', p),
  runFullProcess: (data) => ipcRenderer.send('deploy:runFullProcess', data),
  onLog: (cb) => ipcRenderer.on('deploy:log', (_, d) => cb(d)),
  onDone: (cb) => ipcRenderer.on('deploy:done', (_, d) => cb(d)),
  openAddUserWindow: (csvPath) => ipcRenderer.send('user:openAddWindow', csvPath),
  onUserInit: (cb) => ipcRenderer.on('user:init', (_, csvPath) => cb(csvPath)),
  saveUser: (data) => ipcRenderer.send('user:save', data),
  onUserSaved: (cb) => ipcRenderer.on('user:saved', (_, res) => cb(res)),
  onCSVUpdated: (cb) => ipcRenderer.on('csv:updated', (_, path) => cb(path)),
  deleteUserFromCSV: (filePath, index) => ipcRenderer.invoke('csv:deleteUser', { filePath, index }),
  onStatus: (cb) => ipcRenderer.on('deploy:status', (_, data) => cb(data))
});