const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');
const Store = require('electron-store');
const store = new Store();

let authWindow;
let mainWindow;

// === WINDOW MANAGEMENT ===

function createAuthWindow() {
  authWindow = new BrowserWindow({
    width: 400,
    height: 410,
    resizable: false,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });
  authWindow.loadFile("auth.html");
}

function createMainWindow(credentials) {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      additionalArguments: [
        `--login=${credentials.login}`,
        `--password=${credentials.password}`
      ]
    }
  });
  mainWindow.loadFile("renderer.html");
}

function createAddUserWindow(csvPath) {
  const addWin = new BrowserWindow({
    width: 600,
    height: 800,
    title: 'Добавить пользователя',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  addWin.loadURL(`file://${path.join(__dirname, 'adduser.html')}`);
  addWin.webContents.once('did-finish-load', () => {
    addWin.webContents.send('user:init', csvPath || null);
  });
}

app.whenReady().then(createAuthWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// === HELPERS ===

function runPowerShell(cmd, onData, onErr, onClose) {
  const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', cmd];
  const ps = spawn('powershell.exe', args, { windowsHide: true });

  ps.stdout.on('data', d => onData(d.toString()));
  ps.stderr.on('data', d => onErr(d.toString()));
  ps.on('close', c => onClose(c));

  return ps;
}

// === IPC HANDLERS: AUTH ===
ipcMain.on("auth-check", (event, { login, password }) => {
  const dcHost = store.get('adHost', 'TASVDC04.centrum-air.com');
  const safePass = password.replace(/"/g, '`"');
  const safeLogin = login; 

  const psScript = `
    Add-Type -AssemblyName System.DirectoryServices
    $dcHost = "${dcHost}"
    $fullLogin = "${safeLogin}"
    if ($fullLogin -match "\\\\") { $username = $fullLogin.Split("\\")[1] } 
    elseif ($fullLogin -match "@") { $username = $fullLogin.Split("@")[0] } 
    else { $username = $fullLogin }

    try {
        $ldapPath = "LDAP://$dcHost"
        $entry = New-Object System.DirectoryServices.DirectoryEntry($ldapPath, $fullLogin, "${safePass}")
        $null = $entry.NativeObject
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($entry)
        $searcher.Filter = "(&(objectClass=user)(sAMAccountName=$username))"
        $searcher.PropertiesToLoad.Add("memberOf") | Out-Null
        $result = $searcher.FindOne()

        if ($result -eq $null) {
            Write-Output "ERROR: Auth success, but user object not found."
            return
        }
        $groups = $result.Properties["memberOf"]
        $allGroupsStr = $groups -join ";"
        $required1 = "air.tas.Mail.Recipents"
        $required2 = "AIR.TAS.Local.Admins"

        if (($allGroupsStr -match $required1) -and ($allGroupsStr -match $required2)) {
            Write-Output "OK"
        } else {
            Write-Output "NOADMIN"
        }
    } catch {
        Write-Output "ERROR: $($_.Exception.Message)"
    }
  `;

  const ps = spawn("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", psScript]);

  ps.stdout.on("data", data => {
    const text = data.toString().trim();
    if (text === "OK") {
      event.sender.send("auth-result", { ok: true });
      createMainWindow({ login, password });
      if (authWindow) authWindow.close();
    } else if (text === "NOADMIN") {
      event.sender.send("auth-result", { ok: false, error: "Нет необходимых прав доступа." });
    } else {
      event.sender.send("auth-result", { ok: false, error: text });
    }
  });
});

// === IPC HANDLERS: FILES & CSV ===

ipcMain.handle('dialog:openCSV', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    title: 'Выберите CSV-файл',
    filters: [{ name: 'CSV', extensions: ['csv'] }]
  });
  return canceled ? null : filePaths[0];
});

ipcMain.handle('file:read', async (event, filePath) => {
  try {
    const content = fs.readFileSync(filePath, 'utf8');
    return { ok: true, content };
  } catch (err) {
    return { ok: false, error: err.message };
  }
});

ipcMain.handle('csv:deleteUser', async (event, { filePath, index }) => {
  try {
    const content = fs.readFileSync(filePath, 'utf8').trim().split("\n");
    const header = content[0];
    const rows = content.slice(1);
    rows.splice(index, 1);
    fs.writeFileSync(filePath, [header, ...rows].join("\n") + "\n", 'utf8');
    event.sender.send('csv:updated', filePath);
    return { ok: true };
  } catch (err) { return { ok: false, error: err.message }; }
});

ipcMain.on('user:openAddWindow', (event, csvPath) => createAddUserWindow(csvPath));

// === ИСПРАВЛЕННОЕ СОХРАНЕНИЕ ПОЛЬЗОВАТЕЛЯ ===
ipcMain.on('user:save', (event, { csvPath, user }) => {
  try {
    const header = "FirstName,LastName,MiddleName,Username,Department,Title,Phone,Group1,Group2,Group3";
    const row = `${user.FirstName||''},${user.LastName||''},${user.MiddleName||''},${user.Username||''},${user.Department||''},${user.Title||''},${user.Phone||''},${user.Group1||''},${user.Group2||''},${user.Group3||''}\n`;
    
    let finalPath = csvPath;
    
    // Если путь не передан (первый пользователь в сессии), берем дефолтный на рабочем столе
    if (!csvPath) {
        finalPath = path.join(app.getPath('desktop'), 'NewADUsers.csv');
        
        // ВАЖНО: Если мы начинаем с нуля (csvPath === null), удаляем старый файл, 
        // чтобы не дописывать в "вчерашний" список.
        if (fs.existsSync(finalPath)) {
            try { fs.unlinkSync(finalPath); } catch(e) {}
        }
        
        // Создаем новый файл с заголовком
        fs.writeFileSync(finalPath, header + "\n" + row, 'utf8');
    } else {
        // Если путь передан (это уже 2-й юзер или выбран существующий файл), просто дописываем
        if (fs.existsSync(finalPath)) {
             // Проверка на всякий случай, есть ли заголовок
             const content = fs.readFileSync(finalPath, 'utf8');
             if(!content.trim()) fs.writeFileSync(finalPath, header + "\n", 'utf8');
             fs.appendFileSync(finalPath, row, 'utf8');
        } else {
             // Если файл вдруг удалили
             fs.writeFileSync(finalPath, header + "\n" + row, 'utf8');
        }
    }

    event.reply('user:saved', { ok: true, newPath: finalPath });
    
    const mainWin = BrowserWindow.getAllWindows().find(w => w.title !== 'Добавить пользователя' && w !== authWindow);
    if (mainWin) mainWin.webContents.send('csv:updated', finalPath);
  } catch (err) { event.reply('user:saved', { ok: false, error: err.message }); }
});

ipcMain.handle('params:load', async () => ({
    adHost: store.get('adHost', 'TASVDC04.centrum-air.com'),
    exchHost: store.get('exchHost', 'EXCH0402TAS.centrum-air.com'),
    adminUser: store.get('adminUser', 'CENTRUM-AIR\\'),
    adminPass: store.get('adminPass', '')
}));
ipcMain.on('params:save', (event, params) => { for (const [k,v] of Object.entries(params)) store.set(k,v); });

// === DEPLOY PROCESS (WINRM / INVOKE-COMMAND) ===

ipcMain.on('deploy:runFullProcess', (event, data) => {
  const { localCSV, adHost, exchHost, adminUser, adminPass } = data;

  event.sender.send('deploy:status', { step: 'init', text: 'Initializing WinRM Session...' });

  const psScript = `
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $ErrorActionPreference = 'Stop'
    
    # === PARAMETERS ===
    $remoteHost = "${adHost}"
    $exchHost = "${exchHost}"
    $User = "${adminUser}"
    $Pass = "${adminPass}"
    
    $secPass = $Pass | ConvertTo-SecureString -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($User, $secPass)

    $localFile = "${localCSV}"
    $remoteDestFolder = "C:\\Scripts\\ADNewUsers"
    $remoteCSVName = "NewADUsers.csv"
    $remoteScript = "$remoteDestFolder\\ADNewUsers.ps1"
    $remoteResultFile = "$remoteDestFolder\\NewADUsers_Credentials.csv"

    # ==========================================
    # STEP 1: Connect
    # ==========================================
    Write-Host "[STEP]1: Connecting to AD Controller..."
    try {
        $session = New-PSSession -ComputerName $remoteHost -Credential $cred -ErrorAction Stop
        Write-Host "[OK] Connected"
    } catch {
        Write-Error "[ERROR] Connection Failed: $($_.Exception.Message)"
        exit 1
    }

    # ==========================================
    # STEP 2: Upload
    # ==========================================
    Write-Host "[STEP]2: Uploading CSV..."
    try {
        Invoke-Command -Session $session -ScriptBlock { param($f) if(!(Test-Path $f)) { New-Item -ItemType Directory -Path $f -Force } } -ArgumentList $remoteDestFolder
        Copy-Item -Path $localFile -Destination "$remoteDestFolder\\$remoteCSVName" -ToSession $session -Force
        Write-Host "[OK] CSV Uploaded"
    } catch {
        Write-Error "[ERROR] Upload failed: $($_.Exception.Message)"
        Remove-PSSession $session
        exit 1
    }

    # ==========================================
    # STEP 3: Execute
    # ==========================================
    Write-Host "[STEP]3: Running User Creation Script..."
    try {
        Invoke-Command -Session $session -ScriptBlock {
            param($scriptPath, $resFile)
            if(Test-Path $resFile) { Remove-Item $resFile -Force }
            & $scriptPath
        } -ArgumentList $remoteScript, $remoteResultFile
        Write-Host "[OK] Script Finished"
    } catch {
        Write-Error "[ERROR] Script failed: $($_.Exception.Message)"
        Remove-PSSession $session
        exit 1
    }

    # ==========================================
    # STEP 4: Download
    # ==========================================
    Write-Host "[STEP]4: Downloading Results..."
    $desktopPath = [Environment]::GetFolderPath('Desktop')
    $finalLocalPath = Join-Path $desktopPath "NewADUsers_Final.csv"
    
    try {
        $exists = Invoke-Command -Session $session -ScriptBlock { param($f) Test-Path $f } -ArgumentList $remoteResultFile
        if ($exists) {
            Copy-Item -Path $remoteResultFile -Destination $finalLocalPath -FromSession $session -Force
            Write-Host "[OK] Final CSV copied locally: $finalLocalPath"
        } else {
            Write-Warning "[WARN] Result file not found."
            $finalLocalPath = $localFile
        }
    } catch {
        Write-Error "[ERROR] Download failed: $($_.Exception.Message)"
    }

    # ==========================================
    # STEP 5: Exchange (ADDED DELAY HERE)
    # ==========================================
    # Пишем этот заголовок сразу, чтобы прогресс-бар обновился
    Write-Host "[STEP]5: Connecting to Exchange (Wait 60s for Replication)..."
    
    # === ДОБАВЛЕНА ЗАДЕРЖКА 60 СЕКУНД ===
    Start-Sleep -Seconds 60
    $exchSession = $null
    try {
        $exchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$exchHost/PowerShell/" -Authentication Kerberos -Credential $cred -ErrorAction Stop
        Import-PSSession $exchSession -CommandName Enable-RemoteMailbox -DisableNameChecking | Out-Null
        
        if (Test-Path $finalLocalPath) {
            $users = Import-Csv $finalLocalPath
            foreach ($u in $users) {
                $upn = $u.Username
                if (-not $upn) { continue }
                $routing = "$upn@centrumgroup365.mail.onmicrosoft.com"
                
                Write-Host "[PROGRESS] Enabling mailbox: $upn"
                try {
                    Enable-RemoteMailbox -Identity $upn -RemoteRoutingAddress $routing -ErrorAction Stop
                    Write-Host "[OK] Mailbox enabled"
                } catch {
                    Write-Error ("[ERROR] Failed mailbox for {0}: {1}" -f $upn, $_.Exception.Message)
                }
            }
        }
    } catch {
        Write-Error "[ERROR] Exchange failed: $($_.Exception.Message)"
    } finally {
        if ($exchSession) { Remove-PSSession $exchSession }
    }

    # ==========================================
    # STEP 6: Sync
    # ==========================================
    Write-Host "[STEP]6: Triggering AD Sync..."
    try {
        Invoke-Command -Session $session -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
        Write-Host "[OK] Sync triggered"
    } catch {
        Write-Error "[ERROR] Sync failed: $($_.Exception.Message)"
    }

    Remove-PSSession $session
    Write-Host "[DONE] Process finished"
  `;

  runPowerShell(
    psScript,
    (out) => {
      event.sender.send('deploy:log', { type: 'stdout', text: out });
      if (out.includes('[STEP]')) {
        const stepText = out.match(/\[STEP\](.*)/)?.[1]?.trim();
        event.sender.send('deploy:status', { step: 'progress', text: stepText });
      } else if (out.includes('[OK]')) {
        const okText = out.match(/\[OK\](.*)/)?.[1]?.trim();
        event.sender.send('deploy:status', { step: 'success', text: okText });
        if (okText && okText.includes('Final CSV copied locally')) {
          const match = okText.match(/Final CSV copied locally:\s*(.+)$/);
          if (match) global.finalCSVPath = match[1].trim();
        }
      } else if (out.includes('[WARN]')) {
        const warnText = out.match(/\[WARN\](.*)/)?.[1]?.trim();
        event.sender.send('deploy:status', { step: 'warn', text: warnText });
      } else if (out.includes('[DONE]')) {
        const mainWin = BrowserWindow.getAllWindows().find(w => w !== authWindow);
        const csvMsg = global.finalCSVPath ? `Файл сохранен: ${global.finalCSVPath}` : ``;
        dialog.showMessageBox(mainWin, { type: 'info', title: 'Готово', message: 'Процесс завершен', detail: csvMsg });
      }
    },
    (err) => event.sender.send('deploy:log', { type: 'stderr', text: err }),
    (code) => event.sender.send('deploy:done', { code })
  );
});