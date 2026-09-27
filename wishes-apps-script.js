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
 *  5. Скопируйте выданный адрес вида
 *       https://script.google.com/macros/s/....../exec
 *     и вставьте его в index.html в поле wishesUrl.
 *
 * Пожелания появляются в той же таблице, по строке на каждое.
 * Открыть её можно с телефона или компьютера, из неё же всё выгружается.
 *
 * Если позже меняете код — обязательно делайте новое развёртывание,
 * иначе старый адрес продолжит работать по старому коду.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Когда', 'Имя', 'Пожелание']);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(2, 180);
      sheet.setColumnWidth(3, 520);
    }

    var p = (e && e.parameter) || {};
    var name = String(p.name || '').slice(0, 60);
    var wish = String(p.wish || '').slice(0, 600);

    if (!name && !wish) {
      return json({ ok: false, error: 'empty' });
    }

    sheet.appendRow([new Date(), name, wish]);
    return json({ ok: true });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* открытие адреса в браузере — просто чтобы проверить, что всё развернулось */
function doGet() {
  return json({ ok: true, note: 'Форма пожеланий подключена' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
