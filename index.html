/********************  أدوات مساعدة  ********************/

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && headers.length) {
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function fixDate(d) {
  // رقم Serial من Excel/Sheets
  if (typeof d === "number") {
    return new Date(Math.round((d - 25569) * 86400 * 1000));
  }
  if (d instanceof Date && !isNaN(d)) return d;

  if (typeof d === "string") {
    // يدعم 2025/12/31 أو 2025-12-31
    const nd = new Date(d.replace(/-/g, "/"));
    if (!isNaN(nd)) return nd;
  }
  return null;
}

function sameDay(d1, d2) {
  d1 = fixDate(d1);
  d2 = fixDate(d2);
  if (!d1 || !d2) return false;
  return (
    d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate()
  );
}

function dateKey_(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM");
}

function toHours_(val) {
  // الزمن في الاختبارات عندك في العمود H (index 7)
  // ممكن يكون:
  // - رقم (مثل 2.5)
  // - وقت (Date) مثل 02:30
  // - رقم كسير من اليوم (مثل 0.104166) = 2.5 ساعة
  if (val == null || val === "") return 0;

  if (val instanceof Date && !isNaN(val)) {
    return val.getHours() + (val.getMinutes() / 60) + (val.getSeconds() / 3600);
  }

  const num = Number(val);
  if (!isNaN(num)) {
    // إذا الرقم صغير جداً (غالباً كسر يوم)
    if (num > 0 && num < 1) return num * 24;
    return num;
  }

  return 0;
}

/********************  توزيع المراقبين  ********************/
function توزيع_المراقبين() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const teachersSh = ss.getSheetByName("المعلمين");
  const examsSh    = ss.getSheetByName("الاختبارات");
  if (!teachersSh || !examsSh) {
    SpreadsheetApp.getUi().alert("⚠ لازم يكون موجود شيت: المعلمين + الاختبارات فقط.");
    return;
  }

  // مخرجات (تتولد تلقائياً)
  const resultsSh       = getOrCreateSheet_(ss, "النتائج", ["التاريخ","الفترة","اللجنة","المادة","المراقب1","المراقب2"]);
  const teacherReportSh = getOrCreateSheet_(ss, "كشف_المعلمين", ["اسم المعلم","التاريخ","الفترة","المادة","اللجنة","الزمن(ساعات)"]);
  const dailyReportSh   = getOrCreateSheet_(ss, "كشف_اليوم", ["التاريخ","الفترة","اللجنة","المادة","المراقب1","المراقب2","الزمن(ساعات)"]);
  const followSh        = getOrCreateSheet_(ss, "كشف_المتابعة"); // سنبنيه كامل تحت

  // إعدادات اختيارية (لو الشيت موجود)
  const settings = ss.getSheetByName("الإعدادات");
  const avoidSubjectTeacher = settings ? (settings.getRange("B2").getValue() === "yes") : true;   // افتراض: نعم
  const preventConsecutive  = settings ? (settings.getRange("C2").getValue() === "yes") : true;   // افتراض: نعم

  /************* قراءة بيانات المعلمين *************/
  // شيت المعلمين: (A) ID (B) اسم (C) المادة ...
  const teacherRows = teachersSh.getRange(2, 1, teachersSh.getLastRow()-1, 6).getValues();

  const teachers = [];
  const teacherSubject = {}; // اسم -> مادة
  teacherRows.forEach(r => {
    const name = (r[1] || "").toString().trim();
    const subj = (r[2] || "").toString().trim();
    if (!name) return;
    teachers.push(name);
    teacherSubject[name] = subj;
  });

  /************* قيود منع معلمي المواد حسب التاريخ *************/
  // المقصود: منع "معلمي هذه المواد" من المراقبة في هذا اليوم
  const subjectTeacherBan = {
    "22/12": ["اللغة العربية", "الرياضيات"],
    "23/12": ["العلوم", "التربية الاسلامية", "اللغة العربية"],
    "24/12": ["العلوم", "اللغة الانجليزية"],
    "25/12": ["اللغة الانجليزية", "الدراسات"],
    "28/12": ["التربية الاسلامية", "العلوم"],
    "29/12": ["الرياضيات"],
    "30/12": ["الرياضيات", "الدراسات"],
    "31/12": ["الدراسات", "العلوم", "الرياضيات"]
  };

  /************* قيود منع أسماء معلمين *************/
  const teacherBan = {
    "21/12": ["حمد محمد المعولي", "محمد البلوشي"],
    "22/12": ["وائل حسين", "احمد الهطالي", "ماجد الصبحي", "جميل الناعبي"],
    "23/12": ["ساجان", "احمد جمال", "احمد الهطالي"],
    "24/12": ["كريم", "سلامة", "وليد خليف"],
    "25/12": ["وهب", "سعيد سليمان"],
    "28/12": ["احمد انور", "محمد الهادي"],
    "31/12": [
      "خليل","عبدالله محمد","سعيد سليمان","محمد القايدي","هيثم فوزي",
      "يونس الوشاحي","راشد الغيثي","محمد البلوشي","اباصيري","احمد علي",
      "ساجان","سلامة احمد","احمد الهطالي","كريم","احمد محمد انور",
      "عبدالرحمن حسين","محمد الهادي","هاشم حسن","زهير","خالد المحذوري",
      "عبدالحميد البحري","ناصر السريري","حسين ابراهيم","محمد وليد","فيصل السعدي"
    ]
  };

  /************* قراءة الاختبارات *************/
  // حسب ملفك: Header في الصف 2
  // A: التاريخ, B: الفترة, C: اللجنة, D: المادة, E: عدد المراقبين, F: الصف, H: الزمن
  const lastRow = examsSh.getLastRow();
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert("⚠ شيت الاختبارات فاضي.");
    return;
  }
  const examRows = examsSh.getRange(3, 1, lastRow - 2, 8).getValues();

  /************* تتبّع التوازن *************/
  const teacherHours = {};   // اسم -> مجموع ساعات
  const teacherDays  = {};   // اسم -> Set(days)
  const lastPeriod   = {};   // اسم -> آخر فترة
  const lastDay      = {};   // اسم -> آخر يوم راقب فيه (منع مرتين في اليوم)
  teachers.forEach(t => {
    teacherHours[t] = 0;
    teacherDays[t] = {};
  });

  // لتجهيز كشف المتابعة: جميع الأيام الفعلية الموجودة في الاختبارات
  const allDateKeys = [];
  const dateKeySet = {};
  examRows.forEach(ex => {
    const dt = fixDate(ex[0]);
    if (!dt) return;
    const dk = dateKey_(dt);
    if (!dateKeySet[dk]) {
      dateKeySet[dk] = true;
      allDateKeys.push(dk);
    }
  });
  // ترتيب الأيام حسب التاريخ الحقيقي
  allDateKeys.sort((a,b)=>{
    const [da,ma]=a.split("/").map(Number);
    const [db,mb]=b.split("/").map(Number);
    return (ma*100+da) - (mb*100+db);
  });

  // سنملأ المتابعة: name -> dk -> true
  const followMap = {}; // {name:{dk:true}}
  teachers.forEach(n => followMap[n] = {});

  /************* تنفيذ التوزيع *************/
  let rRow = 2, tRow = 2, dRow = 2;

  examRows.forEach(ex => {

    const dateObj   = fixDate(ex[0]);
    if (!dateObj) return;

    const dk        = dateKey_(dateObj);
    const period    = ex[1];
    const committee = ex[2];
    const subject   = (ex[3] || "").toString().trim();
    const num       = Number(ex[4]) || 0;
    const hours     = toHours_(ex[7]); // H

    if (!num || num < 1) return;

    // اختيار المتاحين
    const available = [];

    teachers.forEach(name => {
      const tSubj = (teacherSubject[name] || "").trim();

      // 1) منع معلم المادة من مراقبة مادته (اختياري)
      if (avoidSubjectTeacher && tSubj && tSubj === subject) return;

      // 2) منع فترتين متتاليتين (اختياري)
      if (preventConsecutive && lastPeriod[name] === period) return;

      // 3) منع المعلم مرتين في نفس اليوم (شرطك الأساسي)
      if (lastDay[name] && sameDay(lastDay[name], dateObj)) return;

      // 4) منع معلمي مواد محددة في يوم محدد
      if (subjectTeacherBan[dk] && subjectTeacherBan[dk].includes(tSubj)) return;

      // 5) منع أسماء محددة في يوم محدد
      if (teacherBan[dk] && teacherBan[dk].includes(name)) return;

      available.push(name);
    });

    // فرز حسب الأقل ساعات (توازن)
    available.sort((a,b) => (teacherHours[a] || 0) - (teacherHours[b] || 0));

    if (available.length < num) {
      // إذا ما يكفي، نسجّل في اللوج (ولا نوقف كل التوزيع)
      console.log("Not enough supervisors for committee:", committee, "date:", dk, "period:", period);
      return;
    }

    const selected = available.slice(0, num);

    // تحديث
    selected.forEach(n => {
      teacherHours[n] = (teacherHours[n] || 0) + hours;
      lastPeriod[n] = period;
      lastDay[n] = dateObj;

      followMap[n][dk] = true;
      teacherDays[n][dk] = true;
    });

    // كتابة المخرجات
    const s1 = selected[0] || "";
    const s2 = selected[1] || "";

    resultsSh.getRange(rRow++, 1, 1, 6).setValues([[dateObj, period, committee, subject, s1, s2]]);
    dailyReportSh.getRange(dRow++, 1, 1, 7).setValues([[dateObj, period, committee, subject, s1, s2, hours]]);

    selected.forEach(n => {
      teacherReportSh.getRange(tRow++, 1, 1, 6).setValues([[n, dateObj, period, subject, committee, hours]]);
    });

  });

  /************* بناء كشف المتابعة (✓ وتلوين + إجمالي الساعات + عدد الأيام) *************/
  followSh.clear();

  // Header
  const headers = ["اسم المراقب"].concat(allDateKeys).concat(["إجمالي الساعات", "عدد أيام المراقبة"]);
  followSh.getRange(1,1,1,headers.length).setValues([headers]);

  // Rows
  const out = [];
  teachers.forEach(name => {
    const row = [name];
    allDateKeys.forEach(dk => row.push(followMap[name][dk] ? "✓" : ""));
    const totalHours = Number(teacherHours[name] || 0);
    const dayCount = Object.keys(teacherDays[name] || {}).length;
    row.push(totalHours);
    row.push(dayCount);
    out.push(row);
  });

  if (out.length) {
    followSh.getRange(2,1,out.length,headers.length).setValues(out);
  }

  // تنسيق: تلوين الخلايا التي فيها ✓
  const dataRange = followSh.getRange(2, 2, Math.max(out.length,1), allDateKeys.length);
  const values = dataRange.getValues();
  const bgs = values.map(r => r.map(v => v === "✓" ? "#d9ead3" : "#ffffff"));
  dataRange.setBackgrounds(bgs).setHorizontalAlignment("center");

  // تنسيق إجمالي الساعات رقم عشري
  const totalCol = 1 + allDateKeys.length + 1;
  followSh.getRange(2, totalCol, out.length, 1).setNumberFormat("0.00");

  SpreadsheetApp.getUi().alert("✔ تم التوزيع بنجاح + منع مرتين في اليوم + توازن حسب الساعات + كشف المتابعة.");
}
