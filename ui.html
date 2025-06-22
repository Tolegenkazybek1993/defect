<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Медицинская система</title>
  <style>
    :root {
      --primary-color: #004080;
      --accent-color: #28a745;
      --danger-color: #dc3545;
      --info-color: #4caf50;
      --warn-color: #ffc107;
      --font-color: #212529;
      --bg-color: #f4f6f9;
    }
    body {
      font-family: 'Segoe UI', sans-serif;
      background: var(--bg-color);
      margin: 0;
      color: var(--font-color);
      display: flex;
      flex-direction: column;
      min-height: 100vh;
    }
    header {
      background: var(--primary-color);
      color: white;
      text-align: center;
      padding: 16px;
    }
    main {
      flex: 1;
      max-width: 880px;
      margin: 20px auto;
      background: white;
      border-radius: 10px;
      padding: 20px;
      box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    }
    .tabs {
      display: flex;
      border-bottom: 2px solid var(--primary-color);
      margin-bottom: 16px;
    }
    .tab {
      flex: 1;
      text-align: center;
      padding: 12px;
      font-weight: bold;
      cursor: pointer;
      border-bottom: 3px solid transparent;
      color: var(--primary-color);
    }
    .tab.active {
      border-color: var(--accent-color);
      color: var(--accent-color);
    }
    .tab-content {
      display: none;
    }
    .tab-content.active {
      display: block;
    }
    input[type="file"], select, input[type="text"], button {
      width: 100%;
      padding: 10px;
      margin: 6px 0 12px;
      font-size: 15px;
      border-radius: 6px;
      border: 1px solid #ccc;
      box-sizing: border-box;
    }
    button {
      background: var(--primary-color);
      color: white;
      font-weight: bold;
      border: none;
      cursor: pointer;
    }
    button:hover:not(:disabled) {
      background: #002855;
    }
    button:disabled {
      background: #aaa;
      cursor: not-allowed;
    }
    .status-message {
      font-weight: bold;
      margin: 8px 0;
    }
    .status-message.success { color: var(--accent-color); }
    .status-message.error { color: var(--danger-color); }
    .status-message.loading::before { content: "\23F3 "; }
    .download-link { margin: 6px 0; }
    .section {
      background: #fdfdfd;
      border-left: 4px solid var(--primary-color);
      padding: 10px;
      border-radius: 6px;
      margin-bottom: 8px;
      font-size: 14px;
      line-height: 1.5;
    }
    .accent { border-left-color: var(--warn-color); }
    .info { border-left-color: var(--info-color); }
    .warning { border-left-color: var(--danger-color); }
    .footer {
      text-align: center;
      padding: 16px;
      font-size: 13px;
      color: #666;
    }
  </style>
</head>
<body>
  <header>
    <h1>Центр Медицинской Аналитики</h1>
    <p>Проверка и Подсказка по МКБ-10</p>
  </header>

  <main>
    <div class="tabs">
      <div class="tab active" data-tab="check">📄 Проверка Excel</div>
      <div class="tab" data-tab="mkb">🔍 Подсказка МКБ-10</div>
    </div>

    <section id="check" class="tab-content active">
      <h2>📄 Загрузка файла</h2>
      <input type="file" id="fileInput" accept=".xlsx" />
      <select id="reportType">
        <option value="A">Отчет A — Платёжная система</option>
        <option value="B">Отчет B — Внешние направления</option>
      </select>
      <button onclick="uploadFile()">🚀 Начать проверку</button>
      <div id="status" class="status-message"></div>
      <div id="downloadLink" class="download-link"></div>
      <div id="adminPanel" style="display:none; font-size:13px; color:#555;">
        🔐 Администрирование через лист <b>«Доступ»</b>
      </div>
    </section>

    <section id="mkb" class="tab-content">
      <h2>🔍 Подсказка по МКБ-10</h2>
      <input type="text" id="mkbInput" placeholder="Например: J45" />
      <button onclick="получить()">🔎 Найти</button>
      <div id="result" class="result-box" style="margin-top:10px; display:none;"></div>
    </section>
  </main>

  <footer class="footer">&copy; 2025 Центр Медицинской Аналитики</footer>

  <script>
    document.querySelectorAll('.tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        tab.classList.add('active');
        document.getElementById(tab.getAttribute('data-tab')).classList.add('active');
      });
    });

    function uploadFile() {
      const fileInput = document.getElementById('fileInput');
      const reportType = document.getElementById('reportType').value;
      const status = document.getElementById('status');
      const downloadLink = document.getElementById('downloadLink');
      if (!fileInput.files.length) {
        alert("Выберите файл .xlsx");
        return;
      }
      status.className = 'status-message loading';
      status.textContent = "Файл обрабатывается...";
      downloadLink.innerHTML = "";
      const reader = new FileReader();
      reader.onload = function(e) {
        const base64 = e.target.result;
        google.script.run
          .withSuccessHandler(url => {
            status.className = 'status-message success';
            status.textContent = "✅ Готово.";
            downloadLink.innerHTML = `<a href="${url}" target="_blank">📥 Скачать</a>`;
          })
          .withFailureHandler(error => {
            status.className = 'status-message error';
            status.textContent = `❌ Ошибка: ${error.message}`;
          })
          .processUploadedFile(base64, fileInput.files[0].name, reportType);
      };
      reader.readAsDataURL(fileInput.files[0]);
    }

    google.script.run.withSuccessHandler(email => {
      if (email === "tolegen.kazybek1993@gmail.com") {
        document.getElementById("adminPanel").style.display = "block";
      }
    }).getCurrentUserEmail();

    function получить() {
      const код = document.getElementById('mkbInput').value.trim().toUpperCase();
      const resultBox = document.getElementById('result');
      resultBox.style.display = 'block';
      resultBox.innerHTML = "⏳ Поиск...";

      if (!код) {
        resultBox.innerHTML = "<b>Введите код диагноза.</b>";
        return;
      }

      google.script.run.withSuccessHandler(function(res) {
        let html = '';
        let остроеГОБМП = false;
        let подозрениеОСМС = false;

        const уникальныеЗаписи = new Set();
        const блоки = [];

        if (!res || res.length === 0) {
          html += `<div class="section warning"><b>❌ Не найдено в таблице 'Правила'</b></div>`;
        } else {
          res.forEach(row => {
            const повод = (row[0] || '').trim();
            const оплата = (row[1] || '').trim();
            const ключ = повод + '|' + оплата;
            if (!уникальныеЗаписи.has(ключ)) {
              уникальныеЗаписи.add(ключ);
              блоки.push(`Повод обращения: <b>${повод}</b><br>Тип оплаты: <b>${оплата}</b>`);
            }
            if (повод.toLowerCase().includes("обострение") && оплата.toLowerCase() === "гобмп") остроеГОБМП = true;
            if (повод.toLowerCase().includes("социально") && оплата.toLowerCase() === "осмс") подозрениеОСМС = true;
          });
          html += `<div class="section accent"><b>📌 Из таблицы:</b><br>${блоки.join('<br><br>')}</div>`;
        }

        if (остроеГОБМП) {
          html += `
            <div class="section warning">
              ⚠️ <b>Обратите внимание:</b><br>
              При поводе обращения <b>"Острое заболевание (состояние)/Обострение хронического заболевания"</b> и источнике финансирования <b>ГОБМП</b> — диагноз может требовать <b>диспансерного наблюдения</b> согласно <b>приказу №149 МЗ РК</b>.<br><br>
              💡 <b>Рекомендуется:</b><br>
              📋 Взять пациента на <b>диспансерный учёт</b><br>
              ${подозрениеОСМС ? "✅ Или использовать повод обращения <b>\"Подозрение на социально-значимое заболевание\"</b> с источником финансирования <b>ОСМС</b>, если пациент не состоит на диспансерном учёте." : "✅ Или использовать другие формы диагноза, допустимые с <b>ОСМС</b> при данном поводе обращения."}
            </div>`;
        }

        if (код.startsWith("Z34") || код === "Z36") {
          html += `
            <div class="section warning">
              ⚠️ <b>Обратите внимание:</b><br>
              Диагнозы <b>Z34</b> и <b>Z36</b> относятся к <b>физиологическому течению беременности</b>.<br><br>
              💡 <b>Рекомендуется:</b><br>
              Назначать услуги по ним только в рамках <b>плановых явок</b>. При осложнениях использовать коды: <b>Z35</b>, <b>O10–O16</b>, <b>O20–O29</b>.
            </div>`;
        }

        resultBox.innerHTML = html;
      }).получитьПодсказку(код);
    }
  </script>
</body>
</html>
