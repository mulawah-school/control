/***********************
 * Global state
 ***********************/
let inputWB = null;
let outWB = null;

let teachers = [];
let exams = [];
let settings = { avoidSubjectTeacher:true, preventConsecutive:true };

// استثناءات لكل يوم: key = dd/MM/yyyy
let excSubjectsByDay = {}; // { "22/12/2025": Set(["الرياضيات", ...]) }
let excTeachersByDay = {}; // { "22/12/2025": Set(["وائل حسين", ...]) }

let selectedDay = null;

const $ = (id)=>document.getElementById(id);
function setStatus(msg){ $("status").textContent = msg; }

function normalize(s){
  return (s||"").toString()
    .replace(/\u00A0/g," ")
    .replace(/\s+/g," ")
    .trim();
}

function esc(v){ return (v??"").toString().replace(/[&<>"']/g,s=>({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[s])); }
function tableHtml(rows){
  if(!rows || !rows.length) return "<div class='muted'>لا يوجد</div>";
  const [h,...b]=rows;
  return `<table><thead><tr>${h.map(c=>`<th>${esc(c)}</th>`).join("")}</tr></thead>
  <tbody>${b.map(r=>`<tr>${r.map(c=>`<td>${esc(c)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
}

/***********************
 * Date / Hours parsing
 ***********************/
function excelDateToJS(d){
  if(d instanceof Date && !isNaN(d)) return d;
  if(typeof d === "number" && isFinite(d)){
    const utcDays = Math.floor(d - 25569);
    const utcValue = utcDays * 86400;
    return new Date(utcValue * 1000);
  }
  if(typeof d === "string" && d.trim()){
    const nd = new Date(d.replace(/-/g,"/"));
    if(!isNaN(nd)) return nd;
  }
  return null;
}
function fmtDMY(d){
  const dd = String(d.getDate()).padStart(2,"0");
  const mm = String(d.getMonth()+1).padStart(2,"0");
  const yy = d.getFullYear();
  return `${dd}/${mm}/${yy}`;
}
function fmtDM(d){
  const dd = String(d.getDate()).padStart(2,"0");
  const mm = String(d.getMonth()+1).padStart(2,"0");
  return `${dd}/${mm}`;
}
function parseHours(v){
  if(v===null || v===undefined || v==="") return 0;
  if(v instanceof Date && !isNaN(v)) return v.getHours() + v.getMinutes()/60;

  if(typeof v === "number" && isFinite(v)){
    if(v > 0 && v < 1) return v*24;      // fraction day
    if(v > 1e9) return v/3600000;        // ms
    return v;                            // hours
  }
  const s = v.toString().trim().replace(",",".");
  if(!s) return 0;
  if(/^\d+:\d{2}(:\d{2})?$/.test(s)){
    const p = s.split(":").map(Number);
    return (p[0]||0) + (p[1]||0)/60 + (p[2]||0)/3600;
  }
  const n = Number(s);
  if(!isNaN(n)){
    if(n > 0 && n < 1) return n*24;
    if(n > 1e9) return n/3600000;
    return n;
  }
  return 0;
}

/***********************
 * Read sheets
 ***********************/
function sheetToAoA(wb, name){
  const ws = wb.Sheets[name];
  if(!ws) return null;
  return XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:"" });
}
function findCol(headerRow, title){
  const t = normalize(title);
  for(let i=0;i<headerRow.length;i++){
    if(normalize(headerRow[i]) === t) return i;
  }
  return -1;
}

function readSettings(wb){
  const aoa = sheetToAoA(wb, "الإعدادات");
  if(!aoa || aoa.length < 2) return { avoidSubjectTeacher:true, preventConsecutive:true };

  const b2 = normalize(aoa[1]?.[1]); // B2
  const c2 = normalize(aoa[1]?.[2]); // C2

  return {
    avoidSubjectTeacher: (b2.toLowerCase() === "yes"),
    preventConsecutive: (c2.toLowerCase() === "yes"),
  };
}

function readTeachers(wb){
  const aoa = sheetToAoA(wb, "المعلمين");
  if(!aoa || aoa.length < 2) throw new Error("شيت المعلمين غير موجود أو فارغ");

  const header = aoa[0];
  const cName = findCol(header, "اسم المعلم");
  const cSub  = findCol(header, "المادة");
  if(cName === -1 || cSub === -1) throw new Error("لم أجد أعمدة (اسم المعلم / المادة) في شيت المعلمين");

  const list = [];
  for(let r=1;r<aoa.length;r++){
    const name = normalize(aoa[r][cName]);
    const subj = normalize(aoa[r][cSub]);
    if(!name) continue;
    list.push({ name, subject: subj });
  }
  return list;
}

function readExams(wb){
  const aoa = sheetToAoA(wb, "الاختبارات");
  if(!aoa || aoa.length < 2) throw new Error("شيت الاختبارات غير موجود أو فارغ");

  // قد تكون العناوين ليست بالصف الأول
  let headerRowIndex = 0;
  for(let i=0;i<Math.min(8,aoa.length);i++){
    if(normalize(aoa[i][0]) === "التاريخ") { headerRowIndex = i; break; }
  }
  const header = aoa[headerRowIndex];

  const cDate = findCol(header, "التاريخ");
  const cPer  = findCol(header, "الفترة");
  const cCom  = findCol(header, "اللجنة");
  const cSub  = findCol(header, "المادة");
  const cNum  = findCol(header, "عدد المراقبين");
  let cTime   = findCol(header, "الزمن"); // اختياري

  if([cDate,cPer,cCom,cSub,cNum].some(x=>x===-1)) {
    throw new Error("عناوين الاختبارات ناقصة: التاريخ/الفترة/اللجنة/المادة/عدد المراقبين");
  }

  const rows = [];
  for(let r=headerRowIndex+1;r<aoa.length;r++){
    const date = excelDateToJS(aoa[r][cDate]);
    if(!date) continue;

    const period = normalize(aoa[r][cPer]);
    const committee = normalize(aoa[r][cCom]);
    const subject = normalize(aoa[r][cSub]);
    const needed = Number(aoa[r][cNum]) || 0;

    let hours = 0;
    if(cTime !== -1) hours = parseHours(aoa[r][cTime]);

    // fallback: try near cols (G..J)
    if(!hours){
      for(const idx of [6,7,8,9]){
        if(idx < aoa[r].length){
          const h = parseHours(aoa[r][idx]);
          if(h){ hours = h; break; }
        }
      }
    }

    if(needed < 1) continue;

    rows.push({ date, period, committee, subject, needed, hours });
  }

  rows.sort((a,b)=>{
    const d = a.date - b.date;
    if(d) return d;
    if(a.period !== b.period) return a.period > b.period ? 1 : -1;
    return a.committee > b.committee ? 1 : -1;
  });

  return rows;
}

/***********************
 * UI: days & exceptions
 ***********************/
function ensureDayBuckets(day){
  if(!excSubjectsByDay[day]) excSubjectsByDay[day] = new Set();
  if(!excTeachersByDay[day]) excTeachersByDay[day] = new Set();
}

function buildDaysGrid(days){
  const grid = $("daysGrid");
  grid.innerHTML = "";
  days.forEach(d=>{
    const card = document.createElement("div");
    card.className = "day";
    card.dataset.day = d;

    const subCount = excSubjectsByDay[d] ? excSubjectsByDay[d].size : 0;
    const teaCount = excTeachersByDay[d] ? excTeachersByDay[d].size : 0;

    card.innerHTML = `<div class="d">${esc(d)}</div>
      <div class="s">مواد: ${subCount} | معلمين: ${teaCount}</div>`;

    card.onclick = ()=>selectDay(d);
    grid.appendChild(card);
  });
}

function refreshDayCardCounts(){
  // إعادة بناء الشبكة لتحديث العدادات
  const days = [...new Set(exams.map(x=>fmtDMY(x.date)))].sort((a,b)=>excelDateToJS(a)-excelDateToJS(b));
  buildDaysGrid(days);
  // إعادة تفعيل selection
  if(selectedDay) selectDay(selectedDay, true);
}

function selectDay(day, silent=false){
  selectedDay = day;
  ensureDayBuckets(day);

  document.querySelectorAll(".day").forEach(x=>x.classList.remove("active"));
  const el = document.querySelector(`.day[data-day="${CSS.escape(day)}"]`);
  if(el) el.classList.add("active");

  $("dayTitle").textContent = `استثناءات اليوم: ${day}`;
  $("addSub").disabled = false;
  $("addTea").disabled = false;

  if(!silent) renderTags();
}

function renderTags(){
  if(!selectedDay) return;

  const subs = [...excSubjectsByDay[selectedDay]];
  const teas = [...excTeachersByDay[selectedDay]];

  $("subTags").innerHTML = subs.map(s=>tagHtml(s, "sub")).join("");
  $("teaTags").innerHTML = teas.map(t=>tagHtml(t, "tea")).join("");
}

function tagHtml(text, kind){
  return `<span class="tag">
    ${esc(text)}
    <button data-kind="${kind}" data-text="${esc(text)}" title="حذف">×</button>
  </span>`;
}

function wireTagDeletes(){
  document.addEventListener("click", (e)=>{
    const btn = e.target.closest("button[data-kind]");
    if(!btn) return;

    const kind = btn.dataset.kind;
    const text = normalize(btn.parentElement.childNodes[0].textContent);

    if(!selectedDay) return;

    if(kind === "sub") excSubjectsByDay[selectedDay].delete(text);
    if(kind === "tea") excTeachersByDay[selectedDay].delete(text);

    renderTags();
    refreshDayCardCounts();
  });
}

/***********************
 * Distribution (all days)
 ***********************/
function distributeAll(){
  // تجميع القيود على شكل dd/MM (للتطابق مع سكربتك)
  const subjectTeacherBan = {}; // dm -> array subjects
  const teacherBan = {};        // dm -> array teachers

  Object.keys(excSubjectsByDay).forEach(dmy=>{
    const d = excelDateToJS(dmy) || new Date(dmy.replace(/-/g,"/"));
    const dm = dmy.substring(0,5); // dd/MM
    subjectTeacherBan[dm] = [...excSubjectsByDay[dmy]].map(normalize);
  });

  Object.keys(excTeachersByDay).forEach(dmy=>{
    const dm = dmy.substring(0,5);
    teacherBan[dm] = [...excTeachersByDay[dmy]].map(normalize);
  });

  const teacherHours = {};
  const teacherDaysDur = {};
  const lastPeriod = {};
  const lastDay = {};
  const assignMap = {};

  teachers.forEach(t=>{
    teacherHours[t.name] = 0;
    teacherDaysDur[t.name] = 0;
    assignMap[t.name] = {};
  });

  const results = [];
  const dailyRows = [];
  const teacherRows = [];

  for(const ex of exams){
    const dmy = fmtDMY(ex.date);
    const dm = fmtDM(ex.date);

    const available = [];

    for(const t of teachers){
      const name = t.name;
      const tSub = t.subject;

      // 1) منع معلم المادة من مراقبة مادته
      if(settings.avoidSubjectTeacher && tSub === ex.subject) continue;

      // 2) منع فترتين متتاليتين
      if(settings.preventConsecutive && lastPeriod[name] === ex.period) continue;

      // 3) منع مراقبتين في نفس اليوم
      if(lastDay[name] && fmtDMY(lastDay[name]) === dmy) continue;

      // 4) استثناء معلمي مواد محددة في هذا اليوم
      if(subjectTeacherBan[dm] && subjectTeacherBan[dm].includes(tSub)) continue;

      // 5) استثناء أسماء معلمين في هذا اليوم
      if(teacherBan[dm] && teacherBan[dm].includes(name)) continue;

      available.push(name);
    }

    // عدالة: الأقل ساعات أولاً
    available.sort((a,b)=> (teacherHours[a]||0) - (teacherHours[b]||0));

    if(available.length < ex.needed) {
      continue;
    }

    const selected = available.slice(0, ex.needed);
    const sup1 = selected[0] || "";
    const sup2 = selected[1] || "";

    selected.forEach(n=>{
      teacherHours[n] += (ex.hours || 0);
      teacherDaysDur[n] += (ex.hours || 0) / 24;
      lastPeriod[n] = ex.period;
      lastDay[n] = ex.date;
      assignMap[n][dm] = true;
    });

    results.push([ex.date, ex.period, ex.committee, ex.subject, sup1, sup2]);
    dailyRows.push([ex.date, ex.period, ex.committee, ex.subject, sup1, sup2]);

    selected.forEach(n=>{
      teacherRows.push([n, ex.date, ex.period, ex.subject, ex.committee]);
    });
  }

  return { teacherHours, teacherDaysDur, assignMap, results, dailyRows, teacherRows, subjectTeacherBan, teacherBan };
}

function buildFollowSheet(assignMap, teacherHours, teacherDaysDur){
  const dateSet = {};
  exams.forEach(ex=> dateSet[fmtDMY(ex.date)] = true);
  const dates = Object.keys(dateSet).sort((a,b)=> excelDateToJS(a) - excelDateToJS(b));

  const header = ["اسم المراقب", ...dates, "إجمالي الساعات", "عدد أيام المراقبة", "إجمالي المدة"];
  const rows = [header];

  const colByDM = {};
  dates.forEach((dStr, idx)=> colByDM[dStr.substring(0,5)] = 1 + idx + 1);

  Object.keys(assignMap).forEach(name=>{
    const row = new Array(header.length).fill("");
    row[0] = name;

    let dayCount = 0;
    Object.keys(assignMap[name] || {}).forEach(dm=>{
      const c = colByDM[dm];
      if(!c) return;
      row[c] = "✓";
      dayCount++;
    });

    row[header.length-3] = Number(teacherHours[name] || 0);
    row[header.length-2] = dayCount;
    row[header.length-1] = Number(teacherDaysDur[name] || 0);

    rows.push(row);
  });

  return rows;
}

function buildExceptionsSheets(subjectTeacherBan, teacherBan){
  const excSub = [["التاريخ(dd/MM)","المواد (استثناء معلميها)"]];
  Object.keys(subjectTeacherBan).sort().forEach(dm=>{
    excSub.push([dm, subjectTeacherBan[dm].join(" , ")]);
  });

  const excTea = [["التاريخ(dd/MM)","المعلمون المستثنون"]];
  Object.keys(teacherBan).sort().forEach(dm=>{
    excTea.push([dm, teacherBan[dm].join(" , ")]);
  });

  return { excSub, excTea };
}

function buildOutputWorkbook(dist){
  const wb = XLSX.utils.book_new();

  const resultsAoA = [["التاريخ","الفترة","اللجنة","المادة","مراقب 1","مراقب 2"], ...dist.results];
  const dailyAoA   = [["التاريخ","الفترة","اللجنة","المادة","مراقب 1","مراقب 2"], ...dist.dailyRows];
  const teacherAoA = [["المعلم","التاريخ","الفترة","المادة","اللجنة"], ...dist.teacherRows];
  const followAoA  = buildFollowSheet(dist.assignMap, dist.teacherHours, dist.teacherDaysDur);

  const { excSub, excTea } = buildExceptionsSheets(dist.subjectTeacherBan, dist.teacherBan);

  const sheets = {
    "النتائج": resultsAoA,
    "كشف_اليوم": dailyAoA,
    "كشف_المعلمين": teacherAoA,
    "كشف_المتابعة": followAoA,
    "استثناء_المواد": excSub,
    "استثناء_المعلمين": excTea,
  };

  Object.entries(sheets).forEach(([name, aoa])=>{
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // تنسيق كشف المتابعة
    if(name === "كشف_المتابعة" && aoa.length > 1){
      const lastCol = aoa[0].length - 1;        // إجمالي المدة
      const colHours = aoa[0].length - 3;       // إجمالي الساعات
      for(let r=1; r<aoa.length; r++){
        ws[XLSX.utils.encode_cell({r, c:lastCol})].z = "[h]:mm";
        ws[XLSX.utils.encode_cell({r, c:colHours})].z = "0.00";
      }
      // تلوين خلايا ✓
      // (SheetJS لا يدعم styles في النسخة المجانية بسهولة داخل المتصفح—لكن ✓ موجودة)
    }

    XLSX.utils.book_append_sheet(wb, ws, name);
  });

  return wb;
}

/***********************
 * File input + Init
 ***********************/
async function onFileSelected(file){
  setStatus("جاري قراءة الملف...");
  const buf = await file.arrayBuffer();
  inputWB = XLSX.read(buf, { type:"array", cellDates:true, raw:true });

  settings = readSettings(inputWB);
  teachers = readTeachers(inputWB);
  exams = readExams(inputWB);

  // Days list
  const days = [...new Set(exams.map(x=>fmtDMY(x.date)))]
    .sort((a,b)=> excelDateToJS(a) - excelDateToJS(b));

  // init exceptions buckets
  days.forEach(d=>ensureDayBuckets(d));
  buildDaysGrid(days);

  // select first day
  if(days.length) selectDay(days[0]);

  $("btnRunAll").disabled = false;
  setStatus(`تم تحميل الملف ✅ (معلمين: ${teachers.length} | اختبارات: ${exams.length})`);

  $("addSub").disabled = false;
  $("addTea").disabled = false;
}

function addSubject(){
  if(!selectedDay) return;
  const v = normalize($("subInput").value);
  if(!v) return;
  ensureDayBuckets(selectedDay);
  excSubjectsByDay[selectedDay].add(v);
  $("subInput").value = "";
  renderTags();
  refreshDayCardCounts();
}
function addTeacher(){
  if(!selectedDay) return;
  const v = normalize($("teaInput").value);
  if(!v) return;
  ensureDayBuckets(selectedDay);
  excTeachersByDay[selectedDay].add(v);
  $("teaInput").value = "";
  renderTags();
  refreshDayCardCounts();
}

function runAll(){
  if(!inputWB) return;

  setStatus("جاري توزيع جميع الأيام...");
  const dist = distributeAll();

  outWB = buildOutputWorkbook(dist);

  // Preview results (first 60)
  const prev = [["التاريخ","الفترة","اللجنة","المادة","مراقب 1","مراقب 2"]];
  dist.results.slice(0,60).forEach(r=>{
    const d = r[0] instanceof Date ? fmtDMY(r[0]) : r[0];
    prev.push([d, r[1], r[2], r[3], r[4], r[5]]);
  });
  $("resultsPreview").innerHTML = tableHtml(prev);

  $("btnExport").disabled = false;
  setStatus("تم التوزيع ✅ جاهز للتصدير");
}

function exportXlsx(){
  if(!outWB) return;
  XLSX.writeFile(outWB, "برنامج_المراقبة_نهائي.xlsx");
}

/***********************
 * Wire UI
 ***********************/
$("file").addEventListener("change", (e)=>{
  const f = e.target.files?.[0];
  if(f) onFileSelected(f);
});

$("addSub").addEventListener("click", addSubject);
$("addTea").addEventListener("click", addTeacher);

$("btnRunAll").addEventListener("click", runAll);
$("btnExport").addEventListener("click", exportXlsx);

wireTagDeletes();
