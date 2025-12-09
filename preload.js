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


 $sec = ConvertTo-SecureString "AfsxP6cv07" -AsPlainText -Force 
      $cred = New-Object System.Management.Automation.PSCredential("centrum-air\az.ruziev-su", $sec) 
        # Получаем группы текущего пользователя на домене
        $groupDns = Invoke-Command -ComputerName TASVDC04 -Credential $cred -ScriptBlock {
            (Get-ADUser $env:USERNAME -Properties memberOf).memberOf
        }

        # Приводим к строке для простого поиска
        $groups = $groupDns -join ";"

        # Группы, которые должны присутствовать
        $required1 = "air.tas.Mail.Recipents"
        $required2 = "AIR.TAS.Local.Admins"

        # Проверка вхождения обеих групп (по CN, без зависимости от DN)
        $hasGroup1 = $groups -match $required1
        $hasGroup2 = $groups -match $required2

        if ($hasGroup1 -and $hasGroup2) {
          Write-Output "OK"
        } else {
          Write-Output "NOADMIN"
        }
