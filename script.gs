function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const user = Session.getActiveUser().getEmail();
  const allowed = JSON.parse(PropertiesService.getScriptProperties().getProperty("allowedEmails") || "[]");

  if (!allowed.includes(user.toLowerCase())) {
    ui.alert("⛔ У вас нет доступа к загрузке файлов. Подсказка МКБ доступна через меню.");
  }

  ui.createMenu("✅ Проверка")
    .addItem("📁 Загрузить Excel", "открытьUI")
    .addItem("▶ Выполнить проверку", "ручнаяПроверка")
    .addItem("🔄 Обновить доступ из таблицы", "обновитьСписокИзТаблицы")
    .addToUi();

  ui.createMenu("📘 Подсказка МКБ")
    .addItem("🔍 Открыть подсказку", "открытьПодсказкуМКБ")
    .addToUi();
}

function открытьПодсказкуМКБ() {
  const html = HtmlService.createHtmlOutputFromFile('mkb_help')
    .setWidth(450)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Подсказка по МКБ');
}

function получитьПодсказку(введенныйКод) {
  const код = normalize(введенныйКод);
  if (!код) return [];

  const лист = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Правила");
  if (!лист) return [];

  const правила = лист.getDataRange().getValues().slice(1);
  const результат = [];

  for (const правило of правила) {
    const rawМКБ = правило[0];
    const повод = правило[1] || "";
    const оплата = правило[2] || "";
    if (!rawМКБ) continue;

    const списокМКБ = rawМКБ.split(",").map(s => normalize(s));
    for (const один of списокМКБ) {
      let совпадает = false;
      if (один.includes("-")) совпадает = isInRange(код, один);
      else if (/^[a-z]\d{2}$/i.test(один)) совпадает = isInRange(код, `${один}-${один}.9`);
      else совпадает = normalize(один) === код;

      if (совпадает) {
        результат.push([повод, оплата]);
        break;
      }
    }
  }
  return результат;
}

function normalize(str) {
  return (str || "").toString().replace(/\s+/g, "").replace(/[‐‑‒–—―]/g, "-").replace(/[\u200B-\u200D\uFEFF]/g, "").toLowerCase().trim();
}

function нормализуйЗаголовок(h) {
  return (h || "").toString().toLowerCase().trim();
}

function открытьUI() {
  const user = Session.getActiveUser().getEmail();
  const allowed = JSON.parse(PropertiesService.getScriptProperties().getProperty("allowedEmails") || "[]");
  if (!user || !allowed.includes(user.toLowerCase())) {
    SpreadsheetApp.getUi().alert("⛔ Доступ запрещён к загрузке файлов.\nПожалуйста, войдите в Google-аккаунт с разрешённой почтой.");
    return;
  }
  const html = HtmlService.createHtmlOutputFromFile('ui')
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Интерфейс загрузки');
}

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail()?.toLowerCase() || "";
}

function инициализироватьДоступ() {
  const начальный = ["tolegen.kazybek1993@gmail.com"];
  PropertiesService.getScriptProperties().setProperty("allowedEmails", JSON.stringify(начальный));
}

function получитьСписокEmailов() {
  const props = PropertiesService.getScriptProperties().getProperty("allowedEmails");
  return JSON.parse(props || "[]");
}

function обновитьСписокИзТаблицы() {
  const email = getCurrentUserEmail();
  if (email !== "tolegen.kazybek1993@gmail.com") {
    throw new Error("⛔ Только администратор может обновлять список.");
  }
  const лист = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Доступ");
  if (!лист) throw new Error("❌ Лист 'Доступ' не найден.");
  const данные = лист.getDataRange().getValues().flat().map(e => (e + "").toLowerCase().trim()).filter(e => e.includes("@"));
  PropertiesService.getScriptProperties().setProperty("allowedEmails", JSON.stringify(данные));
  return данные;
}

function processUploadedFile(base64, filename, format) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64.split(",")[1]), MimeType.MICROSOFT_EXCEL, filename);
  return format === "B" ? обработатьФорматB(blob) : обработатьФорматA(blob);
}

function обработатьФорматA(blob) {
  const base64 = Utilities.base64Encode(blob.getBytes());
  return обработатьExcel(base64, blob.getName());
}

function обработатьФорматB(blob) {
  const file = DriveApp.createFile(blob);
  const converted = Drive.Files.insert({
    title: "Загрузка - " + blob.getName().replace(".xlsx", ""),
    mimeType: MimeType.GOOGLE_SHEETS
  }, blob, { convert: true });

  const spreadsheet = SpreadsheetApp.openById(converted.id);
  const исходныйЛист = spreadsheet.getSheets()[0];
  исходныйЛист.setName("Оригинал");

  const данные = исходныйЛист.getDataRange().getValues();
  const заголовки = данные[0];
  if (!заголовки.includes("Результат проверки")) заголовки.push("Результат проверки");

  const новыеДанные = [заголовки];
  for (let i = 1; i < данные.length; i++) {
    const строка = [...данные[i]];
    while (строка.length < заголовки.length - 1) строка.push("");
    строка.push("");
    новыеДанные.push(строка);
  }

  const проверкаЛист = spreadsheet.insertSheet("Проверка");
  const результат = проверить(новыеДанные, "B");
  проверкаЛист.getRange(1, 1, результат.length, результат[0].length).setValues(результат);
  сформироватьОтчет(результат, spreadsheet);
  return spreadsheet.getUrl();
}

function обработатьExcel(base64, filename) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), MimeType.MICROSOFT_EXCEL, filename);
    const file = DriveApp.createFile(blob);
    const converted = Drive.Files.insert({
      title: "Загрузка - " + filename.replace(".xlsx", ""),
      mimeType: MimeType.GOOGLE_SHEETS
    }, blob, { convert: true });

    const spreadsheet = SpreadsheetApp.openById(converted.id);
    const лист = spreadsheet.getSheets()[0];
    лист.setName("Проверка");

    const данные = лист.getDataRange().getValues();
    const результат = проверить(данные, "A");
    лист.getRange(1, 1, результат.length, результат[0].length).setValues(результат);
    сформироватьОтчет(результат, spreadsheet);
    return spreadsheet.getUrl();
  } catch (e) {
    throw new Error("Ошибка при обработке файла: " + e.message);
  }
}

function ручнаяПроверка() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const лист = ss.getSheetByName("Проверка");
  if (!лист) throw new Error("Нет листа 'Проверка'");
  const данные = лист.getDataRange().getValues();
  const результат = проверить(данные);
  лист.getRange(1, 1, результат.length, результат[0].length).setValues(результат);
  сформироватьОтчет(результат, ss);
  SpreadsheetApp.getUi().alert("✅ Готово! Проверка завершена.");
}

function проверить(данные, формат = "A") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const правила = ss.getSheetByName("Правила").getDataRange().getValues().slice(1);
  const заголовки = данные[0];
  let индексМКБ, индексПовод, индексОплата, индексРезультата;

  if (формат === "B") {
    индексМКБ = заголовки.findIndex(h => нормализуйЗаголовок(h).includes("диагноз") || нормализуйЗаголовок(h).includes("код мкб"));
    индексПовод = заголовки.findIndex(h => нормализуйЗаголовок(h).includes("повод"));
    индексОплата = заголовки.findIndex(h => нормализуйЗаголовок(h).includes("оплата") || нормализуйЗаголовок(h).includes("источник"));
  } else {
    индексМКБ = 14;
    индексПовод = 17;
    индексОплата = 18;
  }

  индексРезультата = заголовки.indexOf("Результат проверки");
  if (индексРезультата === -1) {
    заголовки.push("Результат проверки");
    индексРезультата = заголовки.length - 1;
  }

  for (let i = 1; i < данные.length; i++) {
    const строка = данные[i];
    const кодМКБ = normalize((строка[индексМКБ] || "").split(" ")[0]);
    const повод = normalize(строка[индексПовод]);
    const оплата = normalize(строка[индексОплата]);

    if (!кодМКБ) {
      строка[индексРезультата] = "❌ Нет МКБ-10";
      continue;
    }

    let ok = false;
    for (const правило of правила) {
      const rМКБ = normalize(правило[0]);
      const rПовод = normalize(правило[1]);
      const rОплата = normalize(правило[2]);
      const списокМКБ = rМКБ.split(",").map(s => s.trim());
      for (const один of списокМКБ) {
        let совпадение = false;
        if (один.includes("-")) совпадение = isInRange(кодМКБ, один);
        else if (/^[a-z]\d{2}$/i.test(один)) совпадение = isInRange(кодМКБ, `${один}-${один}.9`);
        else совпадение = normalize(один) === кодМКБ;
        if (совпадение && повод === rПовод && оплата === rОплата) {
          ok = true;
          break;
        }
      }
      if (ok) break;
    }
    строка[индексРезультата] = ok ? "OK" : "❌ Несоответствие";
  }
  return данные;
}

function isInRange(code, range) {
  const parse = s => {
    const m = s.toUpperCase().match(/^([A-Z])(\d{2})(?:\.(\d))?$/);
    return m ? { letter: m[1], major: +m[2], minor: m[3] ? +m[3] : 0 } : null;
  };
  const [start, end] = range.includes("-") ? range.split("-") : [range, range];
  const a = parse(start), b = parse(end), x = parse(code);
  if (!a || !b || !x) return false;
  if (!start.includes(".")) a.minor = 0;
  if (!end.includes(".")) b.minor = 9;
  const idx = c => c.major * 10 + c.minor;
  return a.letter === x.letter && b.letter === x.letter && idx(x) >= idx(a) && idx(x) <= idx(b);
}

function сформироватьОтчет(данные, ss) {
  let отчет = ss.getSheetByName("Отчет");
  if (!отчет) отчет = ss.insertSheet("Отчет");
  else отчет.clear();

  const заголовки = данные[0];
  const индексРезультата = заголовки.indexOf("Результат проверки");

  const формат = заголовки.length >= 25 ? "B" : "A";

  // Поиск индексов по заголовкам
  const индексФИО = заголовки.findIndex(h => {
    const норм = нормализуйЗаголовок(h);
    return норм.includes("фио направителя") || норм.includes("врач направитель");
  });

  const индексСумма = заголовки.findIndex(h =>
    нормализуйЗаголовок(h).includes("цена") ||
    нормализуйЗаголовок(h).includes("сумма")
  );

  const индексИИН = заголовки.findIndex(h => нормализуйЗаголовок(h).includes("иин"));
  const индексФИОПациента = заголовки.findIndex(h => нормализуйЗаголовок(h).includes("фио пациента"));
  const индексМКБ = заголовки.findIndex(h => нормализуйЗаголовок(h).includes("мкб"));
  const индексПовод = заголовки.findIndex(h => нормализуйЗаголовок(h).includes("повод"));
  const индексОплата = заголовки.findIndex(h =>
    нормализуйЗаголовок(h).includes("оплата") ||
    нормализуйЗаголовок(h).includes("источник")
  );

  const ошибки = данные.slice(1).filter(r => r[индексРезультата] !== "OK");
  отчет.getRange("A1:C1").setValues([["ФИО направителя", "Количество дефектов", "Сумма ошибок (₸)"]]);

  const grouped = {};
  for (const r of ошибки) {
    const фио = (r[индексФИО] || "").toString().trim() || "пусто";
    const сумма = parseFloat(r[индексСумма] || 0);
    if (!grouped[фио]) grouped[фио] = { count: 1, sum: сумма };
    else {
      grouped[фио].count++;
      grouped[фио].sum += сумма;
    }
  }

  const summary = Object.entries(grouped).map(([фио, v]) => [фио, v.count, v.sum]);
  if (summary.length) отчет.getRange(2, 1, summary.length, 3).setValues(summary);

  const rowStart = summary.length + 4;
  отчет.getRange(rowStart, 1, 1, 7).setValues([
    ["ФИО пациента", "ИИН", "Код МКБ", "Повод", "Тип оплаты", "Сумма (₸)", "ФИО направителя"]
  ]);

  const детали = ошибки.map(r => [
    r[индексФИОПациента] || "",
    r[индексИИН] || "",
    r[индексМКБ] || "",
    r[индексПовод] || "",
    r[индексОплата] || "",
    parseFloat(r[индексСумма] || 0),
    (r[индексФИО] || "").toString().trim() || "пусто"
  ]);

  if (детали.length) отчет.getRange(rowStart + 1, 1, детали.length, 7).setValues(детали);
}

function doGet() {
  const email = Session.getActiveUser().getEmail();
  const список = JSON.parse(PropertiesService.getScriptProperties().getProperty("allowedEmails") || "[]");

  if (!email) {
    return HtmlService.createHtmlOutput('<h2 style="color:red;">⛔ Пожалуйста, войдите в Google аккаунт.</h2>');
  }

  if (список.includes(email.toLowerCase())) {
    // Пользователь с доступом — возвращаем основной интерфейс
    return HtmlService.createHtmlOutputFromFile('ui');
  } else {
    // Пользователь без доступа — возвращаем только подсказку МКБ
    return HtmlService.createHtmlOutputFromFile('mkb_help');
  }
}
