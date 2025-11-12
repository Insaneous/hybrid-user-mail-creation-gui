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
  csvPathSpan.textContent = "Файл не выбран";
  previewDiv.innerHTML = "";
});

// === Кнопка добавления пользователя ===
btnAddUser.addEventListener('click', () => {
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
btnRun.addEventListener('click', () => {
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
});

// === Получение логов из процесса ===
window.api.onLog(data => {
  logDiv.innerHTML += `<div>${data.text}</div>`;
  logDiv.scrollTop = logDiv.scrollHeight;
});

// === Получение статусов ===
window.api.onStatus?.((status) => {
  logDiv.innerHTML += `<div><b>${status.text}</b></div>`;
  logDiv.scrollTop = logDiv.scrollHeight;

  // === Прогресс ===
  const steps = [
    'Copy CSV',
    'Run ADNewUsers.ps1',
    'Fetch final NewADUsers_Credentials.csv',
    'Connect to Exchange',
    'Start Azure AD Connect sync'
  ];

  if (status.step === 'progress' || status.step === 'success') {
    const matchedStep = steps.findIndex(s => status.text.includes(s)) + 1;
    if (matchedStep > 0) {
      const percent = Math.min((matchedStep / steps.length) * 100, 100);
      progressBar.style.width = percent + '%';
      progressText.textContent = `Шаг ${matchedStep} из ${steps.length}: ${status.text}`;
    }
  }

  if (status.step === 'warn') {
    progressText.textContent = `⚠️ ${status.text}`;
  }

  if (status.step === 'done') {
    progressBar.style.width = '100%';
    progressText.textContent = '✅ Процесс завершён!';
  }
});


// === Когда процесс завершён ===
window.api.onDone(res => {
  logDiv.innerHTML += `<div>✅ Процесс завершён. Код выхода: ${res.code}</div>`;
  logDiv.scrollTop = logDiv.scrollHeight;
});

// === Отображение CSV ===
function renderCSV(content) {
  const rows = content.trim().split("\n").map(r => r.split(","));
  const headers = rows[0];
  const dataRows = rows.slice(1);

  // Добавляем заголовок для кнопки удаления в начало
  let html = `<h3>Просмотр CSV</h3><table><tr>
    <th></th>` + headers.map(h => `<th>${h}</th>`).join('') + `</tr>`;

  dataRows.forEach((r, i) => {
    html += `<tr>
      <td><button class="delBtn" data-index="${i}">🗑️</button></td>` +
      r.map(c => `<td>${c}</td>`).join('') +
      `</tr>`;
  });

  html += `</table>`;
  previewDiv.innerHTML = html;

  // === кнопки удаления ===
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
  // === Загрузить сохранённые параметры ===
  const params = await window.api.loadParams();
  document.getElementById('adHost').value = params.adHost;
  document.getElementById('exchHost').value = params.exchHost;
  document.getElementById('adminUser').value = params.adminUser;
  document.getElementById('adminPass').value = params.adminPass;

  // === Автосохранение при изменении ===
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
