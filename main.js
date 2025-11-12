const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');
const Store = require('electron-store');
const store = new Store();


function createWindow() {
  const win = new BrowserWindow({
    width: 1350,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  win.loadFile('renderer.html');
}

app.whenReady().then(createWindow);
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
$ErrorActionPreference = 'Continue'   ### CHANGED: don't stop on non-critical errors
$plain = "${adminPass}"
$sec = ConvertTo-SecureString $plain -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential("${adminUser}", $sec)

# === 1. Copy CSV and run ADNewUsers.ps1 ===
Write-Host "[STEP]1: Copy CSV"
$netPath = "\\\\${adHost}\\C$\\Temp"
$localFileName = Split-Path -Leaf "${localCSV}"

try {
    New-PSDrive -Name "Z" -PSProvider FileSystem -Root $netPath -Credential $cred -ErrorAction Stop
    Copy-Item -Path "${localCSV}" -Destination "Z:\\$($localFileName)" -Force
    Write-Host "[OK] CSV copied to ${adHost}:\\C$\\Temp\\$($localFileName)"
    Remove-PSDrive -Name "Z" -ErrorAction SilentlyContinue
}
catch {
    Write-Warning "[WARN] Failed to copy CSV: $($_.Exception.Message)"   ### CHANGED
    Remove-PSDrive -Name "Z" -ErrorAction SilentlyContinue
}

try {
    Write-Host "[STEP]2: Run ADNewUsers.ps1 on ${adHost}"
    Invoke-Command -ComputerName "${adHost}" -Credential $cred -ScriptBlock {
        & "C:\\Scripts\\ADNewUsers.ps1" -ErrorAction Continue    ### CHANGED
    } -ErrorAction Continue
    Write-Host "[OK] ADNewUsers.ps1 executed (some warnings possible)."
}
catch {
    Write-Warning "[WARN] ADNewUsers.ps1 had errors: $($_.Exception.Message)"   ### CHANGED
}

# === Copy back final usernames CSV ===
Write-Host "[STEP]3: Fetch final NewADUsers_Credentials.csv from AD Host..."
$remoteFinalCsv = "NewADUsers_Credentials.csv"

# Получаем путь к рабочему столу текущего пользователя
$desktopPath = [Environment]::GetFolderPath('Desktop')
$localCSV_Final = Join-Path $desktopPath "NewADUsers_Final.csv"

try {
    $netPath = "\\\\${adHost}\\C$\\Temp"
    New-PSDrive -Name "Z" -PSProvider FileSystem -Root $netPath -Credential $cred -ErrorAction Stop
    
    if (Test-Path "Z:\\$remoteFinalCsv") {
        Copy-Item -Path "Z:\\$remoteFinalCsv" -Destination $localCSV_Final -Force
        Write-Host "[OK] Final CSV copied locally: $localCSV_Final"
    } else {
        Write-Warning "[WARN] Final CSV not found on ADHost, using original CSV."
        $localCSV_Final = "${localCSV}"
    }
    
    Remove-PSDrive -Name "Z" -ErrorAction SilentlyContinue
} catch {
    Write-Warning "[WARN] Failed to copy back final CSV: $($_.Exception.Message)"
    Remove-PSDrive -Name "Z" -ErrorAction SilentlyContinue
}

# === Enable Remote Mailboxes on Exchange ===
Write-Host "[STEP]4: Connect to Exchange and enable remote mailboxes..."
try {
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://${exchHost}/PowerShell/" -Authentication Kerberos -Credential $cred
    Import-PSSession $session -DisableNameChecking | Out-Null

    $users = Import-Csv -Path $localCSV_Final
    foreach ($u in $users) {
        $upn = $u.Username
        if (-not $upn -or $upn.Trim() -eq "") {
            Write-Warning "Skipping user without Username field."
            continue
        }
        Write-Host "[PROGRESS] Enable remote mailbox for $upn..."
        try {
            $remoteRouting = "$upn@centrumgroup365.mail.onmicrosoft.com"
            Enable-RemoteMailbox -Identity $upn -RemoteRoutingAddress $remoteRouting -ErrorAction Stop
            Write-Host "[OK] Remote mailbox enabled for $upn"
        } catch {
            Write-Warning ("Failed mailbox for " + $upn + ": " + $_.Exception.Message)
            continue   ### CHANGED: skip to next user on error
        }
    }

    Remove-PSSession $session
    Write-Host "[OK] Exchange mailbox provisioning complete."
} catch {
    Write-Warning "[WARN] Exchange mailbox step failed: $($_.Exception.Message)"   ### CHANGED
}

# === Trigger Azure AD Connect sync ===
Write-Host "[STEP]5: Start Azure AD Connect sync..."
try {
    Invoke-Command -ComputerName "${adHost}" -Credential $cred -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta } -ErrorAction Continue
    Write-Host "[OK] Azure AD Connect sync started."
} catch {
    Write-Warning "[WARN] Failed to start Azure AD Connect sync: $($_.Exception.Message)"
}

Write-Host "[DONE] Process completed (some warnings may have occurred)."
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
        event.sender.send('deploy:status', { step: 'done', text: '✅ Completed (with warnings possible).' });

        // === 🎉 Show popup window with completion message ===
        const mainWindow = BrowserWindow.getAllWindows()[0];
        const csvMsg = global.finalCSVPath 
          ? `✅ CSV с данными пользователей сохранён по пути:\n${global.finalCSVPath}`
          : ``;

        dialog.showMessageBox(mainWindow, {
          type: 'info',
          buttons: ['OK'],
          defaultId: 0,
          title: 'Процесс завершён',
          message: 'Все операции завершены!',
          detail: csvMsg
        });
      }
    },
    err => event.sender.send('deploy:log', { type: 'stderr', text: err }),
    code => event.sender.send('deploy:done', { code })
  );
});
