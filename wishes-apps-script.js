/**
 * Приём пожеланий с приглашения в Google Таблицу.
 *
 * Как подключить (делается один раз, минут пять):
 *
 *  1. Создайте пустую Google Таблицу — она и будет хранилищем.
 *  2. В ней: Расширения → Apps Script.
 *  3. Удалите то, что там есть, вставьте весь этот файл, сохраните.
 *  4. Развернуть → Новое развёртывание → тип «Веб-приложение».
 *       Запуск от имени: от моего имени
 *       У кого есть доступ: у всех
 *     Нажмите «Развернуть» и разрешите доступ, когда Google спросит.
 *     Экран «Приложение не проверено» — это нормально: скрипт ваш
 *     собственный. Дополнительные настройки → перейти → разрешить.
 *  5. Скопируйте выданный адрес вида
 *       https://script.google.com/macros/s/....../exec
 *     и вставьте его в index.html в поле wishesUrl.
 *
 * Пожелания появляются в той же таблице, по строке на каждое.
 *
 * Проверить, что всё развернулось: откройте адрес /exec прямо в браузере.
 * Должно ответить {"ok":true,...}.
 *
 * Если позже меняете код — обязательно делайте НОВОЕ развёртывание,
 * иначе старый адрес продолжит работать по старому коду.
 */

/* Заполните, только если скрипт создан отдельно, а не из самой таблицы
   (Расширения → Apps Script). Id берётся из адреса таблицы, между /d/ и /edit. */
var SHEET_ID = '';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    var data = readFields_(e);

    if (!data.name && !data.wish) {
      // видно в Apps Script → Выполнения, если что-то придёт не в том виде
      console.warn('Пустой запрос. Тело: ' + rawBody_(e).slice(0, 500));
      return json({ ok: false, error: 'empty' });
    }

    var sheet = targetSheet_();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Когда', 'Имя', 'Пожелание']);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(2, 180);
      sheet.setColumnWidth(3, 520);
    }

    sheet.appendRow([new Date(), data.name, data.wish]);
    return json({ ok: true });

  } catch (err) {
    console.error(err);
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* открытие адреса в браузере — проверка, что развёртывание живое */
function doGet() {
  try {
    var sheet = targetSheet_();
    return json({ ok: true, note: 'Форма подключена', rows: Math.max(0, sheet.getLastRow() - 1) });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* Страница шлёт urlencoded, и это попадает в e.parameter. Но если данные
   придут любым другим способом — сырым телом или JSON — разбираем и его,
   чтобы форма не молчала из-за формата. */
function readFields_(e) {
  var p = (e && e.parameter) || {};
  var name = p.name || '';
  var wish = p.wish || '';

  if (!name && !wish) {
    var raw = rawBody_(e);
    if (raw) {
      try {
        var obj = JSON.parse(raw);
        name = obj.name || '';
        wish = obj.wish || '';
      } catch (ignored) {
        name = matchField_(raw, 'name');
        wish = matchField_(raw, 'wish');
      }
    }
  }

  return {
    name: String(name).trim().slice(0, 60),
    wish: String(wish).trim().slice(0, 600)
  };
}

function rawBody_(e) {
  return (e && e.postData && e.postData.contents) || '';
}

/* вытаскивает поле из multipart-тела, если оно всё-таки пришло так */
function matchField_(raw, field) {
  var m = raw.match(new RegExp('name="' + field + '"\\s*\\r?\\n\\r?\\n([\\s\\S]*?)\\r?\\n--'));
  return m ? m[1] : '';
}

function targetSheet_() {
  var book = SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!book) {
    throw new Error('Таблица не найдена. Скрипт создан отдельно от таблицы — впишите её id в SHEET_ID.');
  }
  return book.getSheets()[0];
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
