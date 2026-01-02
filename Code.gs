/********************  FIX DATE  ********************/
function fixDate(d) {
  if (typeof d === "number") {
    return new Date(Math.round((d - 25569) * 86400 * 1000));
  }
  if (d instanceof Date && !isNaN(d)) return d;
  if (typeof d === "string") {
    let nd = new Date(d.replace(/-/g, "/"));
    if (!isNaN(nd)) return nd;
  }
  return null;
}

/********************  GET ALL EXAM DAYS (FOR HTML)  ********************/
function getExamDays() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("الاختبارات");

  const rows = sh.getRange(3, 1, sh.getLastRow() - 2, 1).getValues(); // العمود A من الاختبارات
  const map = {};

  rows.forEach(r => {
    const d = fixDate(r[0]);
    if (!d) return;
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
    map[key] = true;
  });

  return Object.keys(map).sort();
}

/********************  GET TEACHERS  ********************/
function getTeachers() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("المعلمين");

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();

  return rows.map(r => ({
    name: r[1],
    subject: r[2]
  }));
}

/********************  GET SUBJECTS FROM الاختبارات  ********************/
function getSubjects() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("الاختبارات");

  const rows = sh.getRange(3, 4, sh.getLastRow() - 2, 1).getValues(); // عمود المادة
  const map = {};

  rows.forEach(r => {
    if (r[0]) map[r[0]] = true;
  });

  return Object.keys(map);
}

/********************  READ EXCEPTIONS FOR DAY  ********************/
function getDayExceptions(day) {
  const prop = PropertiesService.getScriptProperties();
  const raw = prop.getProperty("exceptions_" + day);
  return raw ? JSON.parse(raw) : { teachers: [], subjects: [] };
}

/********************  SAVE EXCEPTIONS FOR DAY  ********************/
function saveDayExceptions(day, data) {
  const prop = PropertiesService.getScriptProperties();
  prop.setProperty("exceptions_" + day, JSON.stringify(data));
  return true;
}
function getDayEditor(day) {
  const html = HtmlService.createTemplateFromFile("day");
  
  html.day = day;
  html.teachers = getTeachers();
  html.subjects = getSubjects();
  html.saved = getDayExceptions(day);

  return html.evaluate().getContent();
}
