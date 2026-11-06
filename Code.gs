/************************************************************************
 *  Цифрлық филология сауалнамасы — Google Apps Script backend
 *  Жауаптарды Google-кестеге жазады әрі дашбордқа JSON қайтарады.
 *
 *  ОРНАТУ (бір рет, ~5 минут) — төмендегі нұсқаулықты қараңыз.
 ************************************************************************/

/* 1) Құпия кілт — ойлап тауып, өзгертіңіз. Дәл осыны results.html-дегі
      CONFIG.key ішіне жазасыз. Бұл кілтті бөтен адамға бермеңіз.      */
var SECRET = "anelya-2026";

/* 2) Деректер жазылатын бет атауы (өзгертпесеңіз де болады). */
var SHEET_NAME = "Жауаптар";

/* ---- Жауаптарды қабылдау (сайт осында POST жібереді) ---- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var sh = getSheet_();
    var p = (e && e.parameter) || {};
    sh.appendRow([
      p.ts || new Date().toISOString(),
      p.name || "",
      p.age || "",
      Number(p.score || 0),
      Number(p.total || 0),
      Number(p.percent || 0),
      p.answers || ""
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/* ---- Нәтижелерді беру (дашборд JSONP арқылы осыны оқиды) ---- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback;
  var out;
  if (p.key !== SECRET) {
    out = { ok: false, error: "bad key" };
  } else {
    var sh = getSheet_();
    var values = sh.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < values.length; i++) { // 0-жол — тақырыптар
      var r = values[i];
      if (!r[0] && !r[1]) continue;
      rows.push({
        ts: r[0], name: r[1], age: r[2],
        score: r[3], total: r[4], percent: r[5], answers: r[6]
      });
    }
    out = { ok: true, rows: rows, count: rows.length };
  }
  var json = JSON.stringify(out);
  if (cb) {
    return ContentService
      .createTextOutput(cb + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---- Көмекші: бет табады не жасайды + тақырыптар ---- */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Уақыты", "Аты-жөні", "Жасы", "Ұпай", "Барлығы", "Пайыз", "Жауаптар(JSON)"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ============================================================
   ОРНАТУ НҰСҚАУЛЫҒЫ (бір рет):

   1. drive.google.com → «Создать/Жаңа» → «Google Таблица» ашыңыз.
      Атын, мысалы, «Сауалнама — нәтижелер» деп қойыңыз.
   2. Сол кестенің ішінде: жоғарғы мәзір → «Расширения/Қосымшалар»
      → «Apps Script» басыңыз.
   3. Ашылған терезедегі барлық мәтінді өшіріп, ОСЫ файлдың
      толық мазмұнын көшіріп қойыңыз. Қаласаңыз, 1-жолдағы
      SECRET мәнін өз құпия кілтіңізге өзгертіңіз.
   4. Жоғарыдан «Сохранить/Сақтау» (дискета белгісі) басыңыз.
   5. Оң жақ жоғарыдан «Развернуть/Орналастыру» → «Новое развёртывание».
        • Тип (шестерёнка) → «Веб-приложение / Web app».
        • «Выполнять от имени / Execute as»  → Я / Me.
        • «У кого есть доступ / Who has access» → «Все / Anyone».
        • «Развернуть / Deploy».
   6. Google рұқсат сұрайды: «Разрешить/Allow». Егер «Google не проверил
      приложение» деген ескерту шықса → «Дополнительно/Advanced» →
      «Перейти к проекту (небезопасно)» → «Разрешить». (Бұл — өз скриптіңіз,
      қауіпсіз.)
   7. Шыққан «URL веб-приложения» (…/exec деп бітеді) сілтемесін көшіріңіз.

   ОСЫ СІЛТЕМЕНІ МАҒАН ЖІБЕРІҢІЗ — мен оны сайтқа қосып, жариялаймын.
   (Немесе өзіңіз index.html мен results.html ішіндегі CONFIG.endpoint-ке
    қойыңыз; results.html-дегі CONFIG.key = SECRET болсын.)
   ============================================================ */
