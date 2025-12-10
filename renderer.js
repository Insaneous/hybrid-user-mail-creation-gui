const btnCSV = document.getElementById('btnCSV');
const btnClearCSV = document.getElementById('btnClearCSV');
const btnAddUser = document.getElementById('btnAddUser');
const btnRun = document.getElementById('btnRun');
const csvPathSpan = document.getElementById('csvPath');
const previewDiv = document.getElementById('preview');
const logDiv = document.getElementById('log');
const adHostInput = document.getElementById('adHost');
const exchHostInput = document.getElementById('exchHost');
const adminUserInput = document.getElementById('adminUser');
const adminPassInput = document.getElementById('adminPass');
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');

let csvPath = null;

const args = window.startup.args;

let login = "";
let password = "";

args.forEach(arg => {
  if (arg.startsWith("--login=")) login = arg.replace("--login=", "");
  if (arg.startsWith("--password=")) password = arg.replace("--password=", "");
});

adminUserInput.value = login;
adminPassInput.value = password;

// === Кнопка выбора CSV ===
btnCSV.addEventListener('click', async () => {
  const path = await window.api.openCSV();
  if (!path) return;

  csvPath = path;
  csvPathSpan.textContent = path;

  const res = await window.api.readFile(path);
  if (res.ok) renderCSV(res.content);
  else previewDiv.textContent = "Ошибка чтения файла: " + res.error;
});

// === Кнопка очистки CSV ===
btnClearCSV.addEventListener('click', () => {
  csvPath = null;
  csvPathSpan.textContent = "Файл не выбран (Будет создан новый)";
  previewDiv.innerHTML = "";
});

// === Кнопка добавления пользователя ===
btnAddUser.addEventListener('click', () => {
  // Открываем окно, передавая текущий csvPath (null если новый)
  window.api.openAddUserWindow(csvPath);
});

// === Автоподгрузка CSV после добавления пользователя ===
window.api.onCSVUpdated(async newPath => {
  csvPath = newPath;
  csvPathSpan.textContent = newPath;

  const res = await window.api.readFile(newPath);
  if (res.ok) renderCSV(res.content);
  else previewDiv.textContent = "Ошибка чтения: " + res.error;
});

// === Кнопка запуска процесса ===
const form = document.getElementById('paramsForm');
form.onsubmit = e => {
  e.preventDefault();
  if (!csvPath) {
    alert("Сначала выберите CSV файл или добавьте пользователей вручную.");
    return;
  }

  const adHost = adHostInput.value.trim();
  const exchHost = exchHostInput.value.trim();
  const adminUser = adminUserInput.value.trim();
  const adminPass = adminPassInput.value.trim();

  if (!adHost || !exchHost || !adminUser || !adminPass) {
    alert("Пожалуйста, заполните все параметры подключения.");
    return;
  }

  logDiv.innerHTML = "🚀 Запуск процесса создания пользователей...<br>";

  window.api.runFullProcess({
    localCSV: csvPath,
    adHost,
    exchHost,
    adminUser,
    adminPass
  });
};

// === Цветное оформление логов ===
function formatLog(text) {
  if (!text) return "";

  let color = "";
  let cleanText = text;

  if (text.includes("[ERROR]")) {
    color = "red";
  } else if (text.includes("[WARN]")) {
    color = "#ff4800";
  } else if (text.includes("[OK]")) {
    color = "lightgreen";
  } else if (text.includes("[STEP]")) {
    color = "#4ea3ff";
    cleanText = `<b>${text}</b>`;
  } else if (text.includes("[PROGRESS]")) {
    color = "#5ec5ff";
  }

  return `<div style="color:${color}; white-space:pre-wrap;">${cleanText}</div>`;
}

// === Хранилище строк для фильтрации дублей ===
const seenLogs = new Set();

// === Получение логов ===
window.api.onLog(data => {
  if (!data.text) return;

  const trimmed = data.text.trim();
  if (seenLogs.has(trimmed)) return;
  seenLogs.add(trimmed);

  logDiv.innerHTML += formatLog(trimmed);
  logDiv.scrollTop = logDiv.scrollHeight;
});

// === Статусы (Progress Bar) ===
window.api.onStatus?.((status) => {
  if (!status.text) return;

  const formatted = formatLog(status.text);
  logDiv.innerHTML += formatted;
  logDiv.scrollTop = logDiv.scrollHeight;

  // === ИСПРАВЛЕННЫЙ Прогресс бар ===
  // Эти строки должны совпадать с тем, что отправляет main.js в [STEP]
  const steps = [
    'Connecting to AD Controller',
    'Uploading CSV',
    'Running User Creation Script',
    'Downloading Results',
    'Connecting to Exchange',
    'Triggering AD Sync'
  ];

  if (status.step === "progress" || status.step === "success") {
    // Ищем частичное совпадение
    const idx = steps.findIndex(s => status.text.includes(s)) + 1;
    
    if (idx > 0) {
      const percent = Math.round(idx / steps.length * 100);
      progressBar.style.width = percent + '%';
      progressText.innerHTML = `🔄 Шаг ${idx} / ${steps.length}: ${status.text}`;
    }
  }

  if (status.step === "warn") {
    progressText.innerHTML = `⚠️ ${status.text}`;
  }

  if (status.step === "done") {
    progressBar.style.width = "100%";
    progressText.innerHTML = "✅ Процесс завершён!";
  }
});

// === Когда процесс завершён ===
window.api.onDone(res => {
  logDiv.innerHTML += `<div>✅ Процесс завершён. Код выхода: ${res.code}</div>`;
  logDiv.scrollTop = logDiv.scrollHeight;
  progressBar.style.width = "100%";
  progressText.innerHTML = "✅ Готово";
});

// === Отображение CSV ===
function renderCSV(content) {
  const rows = content.trim().split("\n").map(r => r.split(","));
  const headers = rows[0];
  const dataRows = rows.slice(1);

  let html = `<h3>Просмотр CSV</h3><table><tr>
    <th></th>` + headers.map(h => `<th>${h}</th>`).join('') + `</tr>`;

  dataRows.forEach((r, i) => {
    // Не рендерим пустые строки
    if(r.length <= 1 && !r[0]) return;
    
    html += `<tr>
      <td><button class="delBtn" data-index="${i}">🗑️</button></td>` +
      r.map(c => `<td>${c}</td>`).join('') +
      `</tr>`;
  });

  html += `</table>`;
  previewDiv.innerHTML = html;

  document.querySelectorAll(".delBtn").forEach(btn => {
    btn.addEventListener("click", async e => {
      const idx = e.target.getAttribute("data-index");
      if (confirm("Удалить этого пользователя из CSV?")) {
        await window.api.deleteUserFromCSV(csvPath, parseInt(idx));
      }
    });
  });
}

document.addEventListener('DOMContentLoaded', async () => {
  const params = await window.api.loadParams();
  if(params.adHost) document.getElementById('adHost').value = params.adHost;
  if(params.exchHost) document.getElementById('exchHost').value = params.exchHost;

  document.querySelectorAll('#params input').forEach(inp => {
    inp.addEventListener('input', () => {
      const updated = {
        adHost: document.getElementById('adHost').value,
        exchHost: document.getElementById('exchHost').value,
        adminUser: document.getElementById('adminUser').value,
        adminPass: document.getElementById('adminPass').value
      };
      window.api.saveParams(updated);
    });
  });
});