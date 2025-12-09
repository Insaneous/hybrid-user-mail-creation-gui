const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');
const Store = require('electron-store');
const store = new Store();
let authWindow;
let mainWindow;

function createAuthWindow() {
  authWindow = new BrowserWindow({
    width: 400,
    height: 410,
    resizable: false,
    webPreferences: {
      preload: __dirname + "/preload.js"
    }
  });
  authWindow.loadFile("auth.html");
}

function createMainWindow(credentials) {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 800,
    webPreferences: {
      preload: __dirname + "/preload.js",
      additionalArguments: [
        `--login=${credentials.login}`,
        `--password=${credentials.password}`
      ]
    }
  });
  mainWindow.loadFile("renderer.html");
}

app.whenReady().then(createAuthWindow);
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

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


function runPowerShell(cmd, onData, onErr, onClose) {
  const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', cmd];
  const ps = spawn('powershell.exe', args, { windowsHide: true });

  ps.stdout.on('data', d => onData(d.toString()));
  ps.stderr.on('data', d => onErr(d.toString()));
  ps.on('close', c => onClose(c));

  return ps;
}

// --- IPC handlers ---
ipcMain.on("auth-check", (event, { login, password }) => {
  const ps = spawn("powershell.exe", [
    "-NoProfile",
    "-Command",
    `
      $sec = ConvertTo-SecureString "${password}" -AsPlainText -Force
      $cred = New-Object System.Management.Automation.PSCredential("${login}", $sec)

      try {
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
      } catch {
        Write-Output "ERROR: $($_.Exception.Message)"
      }
    `
  ]);

  ps.stdout.on("data", data => {
    const text = data.toString().trim();

    if (text === "OK") {
      event.sender.send("auth-result", { ok: true });

      // запускаем основное окно
      createMainWindow({ login, password });
      return;
    }

    if (text === "NOADMIN") {
      event.sender.send("auth-result", { ok: false, error: "Неправильные данные или нет прав." });
      return;
    }

    event.sender.send("auth-result", { ok: false, error: text });
  });
});

ipcMain.handle('dialog:openCSV', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    title: 'Выберите CSV-файл пользователей',
    properties: ['openFile'],
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
    if (!fs.existsSync(filePath)) throw new Error("Файл не найден");

    const content = fs.readFileSync(filePath, 'utf8').trim().split("\n");
    if (content.length <= 1) throw new Error("В файле нет данных");

    const header = content[0];
    const rows = content.slice(1);
    if (index < 0 || index >= rows.length) throw new Error("Неверный индекс строки");

    rows.splice(index, 1); // удаляем выбранную строку
    const newContent = [header, ...rows].join("\n") + "\n";

    fs.writeFileSync(filePath, newContent, 'utf8');

    // уведомим renderer, чтобы обновить предпросмотр
    event.sender.send('csv:updated', filePath);
    return { ok: true };
  } catch (err) {
    console.error("Ошибка удаления пользователя:", err);
    return { ok: false, error: err.message };
  }
});

ipcMain.on('user:openAddWindow', (event, csvPath) => {
  createAddUserWindow(csvPath);
});

ipcMain.on('user:save', (event, { csvPath, user }) => {
  try {
    // добавили новые поля Username и Phone
    const header = "FirstName,LastName,MiddleName,Username,Department,Title,Phone,Group1,Group2,Group3";
    const row = `${user.FirstName || ''},${user.LastName || ''},${user.MiddleName || ''},${user.Username || ''},${user.Department || ''},${user.Title || ''},${user.Phone || ''},${user.Group1 || ''},${user.Group2 || ''},${user.Group3 || ''}\n`;

    let finalPath = csvPath;
    if (!csvPath || !fs.existsSync(csvPath)) {
      // создаём новый CSV
      finalPath = path.join(app.getPath('desktop'), 'NewADUsers.csv');
      fs.writeFileSync(finalPath, header + "\n" + row, 'utf8');
    } else {
      // если в старом CSV нет новых столбцов — проверим заголовок и добавим, если нужно
      const existing = fs.readFileSync(csvPath, 'utf8');
      if (!existing.startsWith("FirstName,LastName,MiddleName,Username,Department,Title,Phone")) {
        // перезаписываем с новым заголовком
        const newContent = header + "\n" + existing.split("\n").slice(1).join("\n");
        fs.writeFileSync(csvPath, newContent, 'utf8');
      }
      fs.appendFileSync(csvPath, row, 'utf8');
    }

    event.reply('user:saved', { ok: true, newPath: finalPath });

    // уведомляем главное окно (renderer) об обновлённом CSV
    const allWindows = BrowserWindow.getAllWindows();
    const mainWin = allWindows.find(w => w.title !== 'Добавить пользователя');
    if (mainWin) {
      mainWin.webContents.send('csv:updated', finalPath);
    }

  } catch (err) {
    event.reply('user:saved', { ok: false, error: err.message });
  }
});

// === Save & Load Parameters ===
ipcMain.handle('params:load', async () => {
  return {
    adHost: store.get('adHost', 'TASVDC04.centrum-air.com'),
    exchHost: store.get('exchHost', 'EXCH0402TAS.centrum-air.com'),
    adminUser: store.get('adminUser', 'CENTRUM-AIR\\'),
    adminPass: store.get('adminPass', '')
  };
});

ipcMain.on('params:save', (event, params) => {
  for (const [key, value] of Object.entries(params)) {
    store.set(key, value);
  }
});

// === Главный процесс запуска AD+Exchange+Sync ===
ipcMain.on('deploy:runFullProcess', (event, data) => {
  const { localCSV, adHost, exchHost, adminUser, adminPass } = data;

  event.sender.send('deploy:status', { step: 'init', text: 'Starting process of creating users...' });

  const psScript = `
$ErrorActionPreference = 'Continue'
$plain = "${adminPass}"
$sec = ConvertTo-SecureString $plain -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential("${adminUser}", $sec)

# ================================
# STEP 1 — Copy CSV
# ================================
Write-Host "[STEP]1: Copy CSV"
$netPath = "\\\\${adHost}\\C$\\Temp"
$localFileName = Split-Path -Leaf "${localCSV}"

try {
    New-PSDrive -Name "Z" -PSProvider FileSystem -Root $netPath -Credential $cred -ErrorAction Stop
    Copy-Item -Path "${localCSV}" -Destination "Z:\\$localFileName" -Force
    Write-Host "[OK] CSV copied to ${adHost}:\\C$\\Temp\\$localFileName"
}
catch {
    Write-Error "[ERROR] Failed to copy CSV: $($_.Exception.Message)"
}
finally {
    Remove-PSDrive -Name "Z" -ErrorAction SilentlyContinue
}

# ================================
# STEP 2 — Run ADNewUsers.ps1
# ================================
Write-Host "[STEP]2: AD user creation on ${adHost}"

try {
    Invoke-Command -ComputerName "${adHost}" -Credential $cred -ScriptBlock {
        & "C:\\Scripts\\ADNewUsers.ps1"
    } -ErrorAction Stop

    Write-Host "[OK] AD user creation completed"
}
catch {
    Write-Error "[ERROR] ADNewUsers.ps1 failed: $($_.Exception.Message)"
}

# ================================
# STEP 3 — Fetch final CSV
# ================================
Write-Host "[STEP]3: Fetch final NewADUsers_Credentials.csv"

$remoteFinalCsv = "NewADUsers_Credentials.csv"
$desktopPath = [Environment]::GetFolderPath('Desktop')
$localCSV_Final = Join-Path $desktopPath "NewADUsers_Final.csv"

try {
    New-PSDrive -Name "Z" -PSProvider FileSystem -Root $netPath -Credential $cred -ErrorAction Stop
    
    if (Test-Path "Z:\\$remoteFinalCsv") {
        Copy-Item -Path "Z:\\$remoteFinalCsv" -Destination $localCSV_Final -Force
        Write-Host "[OK] Final CSV copied locally: $localCSV_Final"
    } else {
        Write-Warning "[WARN] Final CSV not found. Using original CSV."
        $localCSV_Final = "${localCSV}"
    }
}
catch {
    Write-Error "[ERROR] Failed to fetch final CSV: $($_.Exception.Message)"
}
finally {
    Remove-PSDrive -Name "Z" -ErrorAction SilentlyContinue
}

# ================================
# STEP 4 — Enable Remote Mailboxes
# ================================
Write-Host "[STEP]4: Enable Remote Mailboxes"

try {
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://${exchHost}/PowerShell/" -Authentication Kerberos -Credential $cred
    Import-PSSession $session -DisableNameChecking | Out-Null

    $users = Import-Csv -Path $localCSV_Final
    foreach ($u in $users) {
        $upn = $u.Username
        if (-not $upn) {
            Write-Warning "[WARN] Skipping user without Username"
            continue
        }

        $remoteRouting = "$upn@centrumgroup365.mail.onmicrosoft.com"
        Write-Host "[PROGRESS] Enabling mailbox: $upn"

        try {
            Enable-RemoteMailbox -Identity $upn -RemoteRoutingAddress $remoteRouting -ErrorAction Stop
            Write-Host "[OK] Mailbox enabled: $upn"
        } catch {
            Write-Error ("[ERROR] Failed mailbox for {0}: {1}" -f $upn, $_.Exception.Message)
        }
    }

    Write-Host "[OK] Exchange provisioning completed"
}
catch {
    Write-Error "[ERROR] Exchange mailbox stage failed: $($_.Exception.Message)"
}
finally {
    if ($session) { Remove-PSSession $session }
}

# ================================
# STEP 5 — Start Azure Sync
# ================================
Write-Host "[STEP]5: Start Azure AD Sync"

try {
    Invoke-Command -ComputerName "${adHost}" -Credential $cred -ScriptBlock { 
        Start-ADSyncSyncCycle -PolicyType Delta 
    } -ErrorAction Stop

    Write-Host "[OK] Azure AD Connect sync started"
}
catch {
    Write-Error "[ERROR] Azure AD Sync failed: $($_.Exception.Message)"
}

Write-Host "[DONE] Process finished"
`;

  runPowerShell(
    psScript,
    out => {
      event.sender.send('deploy:log', { type: 'stdout', text: out });

      if (out.includes('[STEP]')) {
        const stepText = out.match(/\[STEP\](.*)/)?.[1]?.trim();
        event.sender.send('deploy:status', { step: 'progress', text: stepText });

      } else if (out.includes('[OK]')) {
        const okText = out.match(/\[OK\](.*)/)?.[1]?.trim();
        event.sender.send('deploy:status', { step: 'success', text: okText });

        // 💡 Capture final CSV path from "[OK] Final CSV copied locally"
        if (okText?.includes('Final CSV copied locally')) {
          const match = okText.match(/Final CSV copied locally:\s*(.+)$/);
          if (match) {
            global.finalCSVPath = match[1].trim();
          }
        }

      } else if (out.includes('[WARN]')) {
        const warnText = out.match(/\[WARN\](.*)/)?.[1]?.trim();
        event.sender.send('deploy:status', { step: 'warn', text: warnText });

      } else if (out.includes('[DONE]')) {

        // === 🎉 Show popup window with completion message ===
        const mainWindow = BrowserWindow.getAllWindows()[0];
        const csvMsg = global.finalCSVPath 
          ? `CSV со сгенерированными данными сохранён по пути:\n${global.finalCSVPath}` +
            `\n\nВозможно были ошибки, пожалуйста, проверьте логи в программе.`
          : ``;

        dialog.showMessageBox(mainWindow, {
          type: 'info',
          buttons: ['OK'],
          defaultId: 0,
          title: 'Процесс завершён',
          message: 'Все операции завершены',
          detail: csvMsg
        });
      }
    },
    err => event.sender.send('deploy:log', { type: 'stderr', text: err }),
    code => event.sender.send('deploy:done', { code })
  );
});
