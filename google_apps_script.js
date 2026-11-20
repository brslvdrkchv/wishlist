/**
 * GOOGLE APPS SCRIPT — БЕКЕНД ДЛЯ ВИШЛИСТА БОРИСЛАВА
 * 
 * ЯК ПІДКЛЮЧИТИ (займе 2 хвилини):
 * 1. Відкрийте https://sheets.new (створиться нова Google Таблиця)
 * 2. Назвіть її, наприклад, "Вишлист Борислава — Бронювання"
 * 3. У верхньому меню оберіть: Розширення (Extensions) -> Apps Script
 * 4. Зітріть все, що там є, вставте цей код та збережіть (Ctrl+S / Cmd+S)
 * 5. Натисніть синю кнопку "Почати розгортання" (Deploy) -> "Нове розгортання" (New deployment)
 * 6. Оберіть тип: "Веб-додаток" (Web app)
 * 7. Вкажіть налаштування:
 *    - Виконувати як: Я (Execute as: Me)
 *    - Хто має доступ: Будь-хто (Who has access: Anyone)
 * 8. Натисніть "Розгорнути" (Deploy), підтвердіть права (Advanced -> Go to ... unsafe)
 * 9. Скопіюйте отриманий URL веб-додатка (закінчується на /exec)
 * 10. Вставте його в index.html у рядок:
 *     type: 'google_sheets',
 *     googleSheetsUrl: 'https://script.google.com/macros/s/.../exec'
 */

const SHEET_NAME = "Бронювання";

// Опціонально: сповіщення в Telegram!
// Якщо вказати токен і Chat ID, вам приходитиме повідомлення щойно хтось займе подарунок!
const TELEGRAM_BOT_TOKEN = ""; // наприклад, "123456789:ABCDE..."
const TELEGRAM_CHAT_ID = "";   // ваш chat_id

function doGet(e) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const state = {};
  
  // Рядок 0 — заголовки [ID, Статус, Ім'я, ClientID, Дата]
  for (let i = 1; i < data.length; i++) {
    const [id, reserved, by, clientId, at] = data[i];
    if (id && reserved) {
      state[id] = {
        reserved: true,
        by: by || "Друг",
        clientId: clientId || "",
        at: at || ""
      };
    }
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", state: state }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = e.postData.contents;
    const body = JSON.parse(raw);
    const sheet = getOrCreateSheet();
    
    if (body.action === "sync_all" && body.state) {
      // Очищаємо старі дані (крім заголовка) і записуємо актуальні
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
      }
      
      const rows = [];
      for (const [id, val] of Object.entries(body.state)) {
        if (val && val.reserved) {
          rows.push([id, true, val.by || "Друг", val.clientId || "", val.at || new Date().toISOString()]);
        }
      }
      
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, 5).setValues(rows);
      }
      
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok", count: rows.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ignored" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["ID Товару", "Заброньовано", "Ім'я друга", "Клієнтський ID", "Дата бронювання"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#e2e8f0");
  }
  return sheet;
}

function sendTelegram(text) {
  if (!TELEGRAM_BOT_TOKEN || !TELEGRAM_CHAT_ID) return;
  try {
    const url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage";
    UrlFetchApp.fetch(url, {
      method: "post",
      payload: {
        chat_id: TELEGRAM_CHAT_ID,
        text: text,
        parse_mode: "HTML"
      },
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log("Telegram send error: " + e);
  }
}
