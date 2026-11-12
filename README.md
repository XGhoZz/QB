# QB
const ANSWERS = [
  1, 1, 2, 1, 2, 1, 2, 1, 1, 1, 0, 2, 0, 1, 1, 0, 3, 1, 2, 0,
  3, 0, 1, 3, 0, 1, 2, 1, 1, 0, 1, 2, 1, 1, 2, 0, 1, 1, 1, 0,
  1, 0, 0, 1, 2, 0, 1, 2, 0, 2, 1, 1, 1, 3, 3, 2, 2, 2, 2, 1,
  2, 1, 1, 1, 1, 0, 1, 2, 1, 2, 1, 1, 1, 1, 2, 0, 1, 1, 1, 2,
  2, 1, 1, 1, 2, 1, 2, 1, 1, 2, 1, 1, 0, 1
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const props = PropertiesService.getScriptProperties();

    // 1. لحظة دخول الطالب وبدء الامتحان
    if (data.action === "start") {
      const name = String(data.name || "").trim();
      const section = Number(data.section);
      
      if (name.length < 2 || name.length > 60 || !Number.isInteger(section) || section < 1 || section > 999) {
        return jsonRes({ ok: false, error: "invalid_input" });
      }

      // بنشوف لو الطالب ده (نفس الاسم ونفس السكشن) دخل قبل كده
      const studentKey = "STU_" + name.toLowerCase() + "_" + section;
      if (props.getProperty(studentKey)) {
        // لو دخل قبل كده، بنخليها محاولة تدريبية ومش بنسجلها تاني
        return jsonRes({ ok: true, recorded: false });
      }

      // بنسجل وقت الدخول والبيانات في الشيت فوراً
      const entryTime = new Date();
      sheet.appendRow([entryTime, name, section, "جاري الامتحان..."]);
      const rowIndex = sheet.getLastRow();

      // بنعمل Token للجلسة ونربطه برقم الصف بتاعه في الشيت
      const token = Utilities.getUuid();
      const sessionData = JSON.stringify({
        row: rowIndex,
        startTime: entryTime.getTime(),
        finished: false
      });
      
      props.setProperty("TOK_" + token, sessionData);
      props.setProperty(studentKey, "1");

      return jsonRes({ ok: true, recorded: true, token: token });
    }

    // 2. لحظة إنهاء الطالب للامتحان وتسجيل الدرجة في نفس الصف
    if (data.action === "finish") {
      const token = String(data.token || "");
      const answers = data.answers;
      const rawSession = props.getProperty("TOK_" + token);

      if (!rawSession) return jsonRes({ ok: false, error: "unknown_session" });
      const session = JSON.parse(rawSession);

      if (session.finished) return jsonRes({ ok: false, error: "already_finished" });
      if (!Array.isArray(answers) || answers.length !== ANSWERS.length) {
        return jsonRes({ ok: false, error: "invalid_input" });
      }

      // السيرفر بيحسب الدرجة بنفسه عشان محدش يزوّرها
      let score = 0;
      for (let i = 0; i < ANSWERS.length; i++) {
        if (answers[i] === ANSWERS[i]) score++;
      }

      // بنحدث العمود الرابع (عمود الدرجة) في نفس صف الطالب اللي اتسجل وقت الدخول
      sheet.getRange(session.row, 4).setValue(score);

      session.finished = true;
      props.setProperty("TOK_" + token, JSON.stringify(session));

      return jsonRes({ ok: true, recorded: true, score: score });
    }

    return jsonRes({ ok: false, error: "invalid_action" });
  } catch (err) {
    return jsonRes({ ok: false, error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function jsonRes(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
