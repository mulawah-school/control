let teachers = [];
let exams = [];
let exceptions = { teachers: {}, subjects: {} };

// ----------------------------
// 1) رفع ملف المعلمين
// ----------------------------
function uploadTeachers() {
    const file = document.getElementById("fileTeachers").files[0];
    if (!file) return alert("اختر ملف المعلمين");

    let reader = new FileReader();
    reader.onload = function(e) {
        let workbook = XLSX.read(e.target.result, { type: "binary" });
        let sheet = workbook.Sheets[workbook.SheetNames[0]];

        teachers = XLSX.utils.sheet_to_json(sheet);

        localStorage.setItem("teachers", JSON.stringify(teachers));

        document.getElementById("teacherStatus").innerHTML =
            "<span style='color:green'>✔ تم رفع ملف المعلمين (" + teachers.length + " معلم)</span>";
    };
    reader.readAsBinaryString(file);
}
function addExam() {
    const ex = {
        date: document.getElementById("examDate").value,
        subject: document.getElementById("examSubject").value,
        period: document.getElementById("examPeriod").value,
        committee: document.getElementById("examCommittee").value,
        needed: Number(document.getElementById("examNeeded").value),
        duration: document.getElementById("examDuration").value
    };

    exams.push(ex);
    localStorage.setItem("exams", JSON.stringify(exams));

    renderExamTable();
}

function renderExamTable() {
    let html = "<tr><th>التاريخ</th><th>اللجنة</th><th>المادة</th><th>الفترة</th><th>عدد</th><th>حذف</th></tr>";

    exams.forEach((e, i) => {
        html += `
        <tr>
        <td>${e.date}</td>
        <td>${e.committee}</td>
        <td>${e.subject}</td>
        <td>${e.period}</td>
        <td>${e.needed}</td>
        <td><button onclick="deleteExam(${i})">🗑</button></td>
        </tr>`;
    });

    document.getElementById("examTable").innerHTML = html;
}
function parseDuration(str) {
    if (!str) return 0;

    if (str.includes(":")) {
        let parts = str.split(":");
        return Number(parts[0]) + Number(parts[1]) / 60;
    }

    return Number(str) || 0;
}
function runDistribution() {

    if (teachers.length === 0) return alert("حمّل ملف المعلمين أولاً");
    if (exams.length === 0) return alert("أضف الاختبارات أولاً");

    let teacherHours = {};
    let assigned = {}; // "خليل|2025-12-22": true

    teachers.forEach(t => teacherHours[t.name] = 0);

    let resultsHTML = "<tr><th>التاريخ</th><th>اللجنة</th><th>المادة</th><th>المراقبين</th><th>واتساب</th></tr>";

    exams.forEach(ex => {

        let day = ex.date;
        if (!day) return;

        let needed = Number(ex.needed);
        let duration = parseDuration(ex.duration);

        let list = teachers
            .map(t => t.name)
            .filter(n => {
                // استثناء معلم
                if (exceptions.teachers[day]?.includes(n)) return false;

                // استثناء مادة
                if (exceptions.subjects[day]?.includes(ex.subject)) return false;

                // لا يراقب في اليوم مرتين
                if (assigned[n + "|" + day]) return false;

                return true;
            });

        // ترتيب حسب الأقل ساعات
        list.sort((a, b) => teacherHours[a] - teacherHours[b]);

        let selected = list.slice(0, needed);

        selected.forEach(n => {
            teacherHours[n] += duration;
            assigned[n + "|" + day] = true;
        });

        resultsHTML += `
        <tr>
            <td>${ex.date}</td>
            <td>${ex.committee}</td>
            <td>${ex.subject}</td>
            <td>${selected.join(" ، ") || "-"}</td>
            <td><button class="whatsapp-btn" onclick="sendWhatsApp('${selected.join(",")}','${ex.date}','${ex.committee}')">📱</button></td>
        </tr>`;
    });

    document.getElementById("resultTable").innerHTML = resultsHTML;
}
function sendWhatsApp(names, date, committee) {
    let msg = `تم تكليفك بالمراقبة يوم ${date} في لجنة ${committee}`;
    window.open(`https://wa.me/?text=${encodeURIComponent(msg)}`);
}
function exportExcel() {
    let table = document.getElementById("resultTable");
    let wb = XLSX.utils.table_to_book(table);
    XLSX.writeFile(wb, "التوزيع.xlsx");
}
function calculateTeacherHours() {
    let hoursMap = {};
    teachers.forEach(t => hoursMap[t.name] = 0);

    let rows = document.querySelectorAll("#resultTable tr");

    rows.forEach(r => {
        let cols = r.querySelectorAll("td");
        if (cols.length === 0) return;

        let names = cols[3].innerText.split("،").map(s => s.trim());
        let duration = cols[5] ? parseDuration(cols[5].innerText) : 0;

        names.forEach(n => {
            if (hoursMap[n] !== undefined) {
                hoursMap[n] += duration;
            }
        });
    });

    return hoursMap;
}

function buildFollowMatrix() {

    if (exams.length === 0) return alert("لا توجد اختبارات");

    // استخراج جميع الأيام بدون تكرار
    let days = [...new Set(exams.map(e => e.date))];
    days.sort();

    // استخراج جميع المعلمين
    let teacherNames = teachers.map(t => t.name);

    // تجهيز جدول فارغ
    let follow = {};
    teacherNames.forEach(n => follow[n] = {});

    // تسجيل ✓ لمن راقب
    exams.forEach(ex => {
        let day = ex.date;

        // نفس المنطق المستخدم في جدول النتائج
        let duration = parseDuration(ex.duration);

        let selected = []; // سنأخذ الأسماء من جدول نتائج التوزيع

        // استخراج النتائج من صفحة HTML
        let resultRows = document.querySelectorAll("#resultTable tr");

        resultRows.forEach(row => {
            let cols = row.querySelectorAll("td");
            if (cols.length === 0) return;

            let rDay = cols[0].innerText.trim();
            let rCommittee = cols[1].innerText.trim();
            let rNames = cols[3].innerText.split("،").map(s => s.trim());

            if (rDay === day) {
                rNames.forEach(n => {
                    if (teacherNames.includes(n)) {
                        follow[n][day] = true;
                    }
                });
            }
        });
    });

    // بناء HTML للجدول
    let html = "<table><tr><th class='follow-header'>اسم المعلم</th>";

    days.forEach(d => {
        html += `<th class='follow-header'>${d}</th>`;
    });

    html += "<th class='follow-header'>الأيام</th>";
    html += "<th class='follow-header'>الساعات</th>";
    html += "</tr>";

    // حساب الساعات
    let teacherHours = calculateTeacherHours();

    // تعبئة الصفوف
    teacherNames.forEach(n => {
        let countDays = 0;

        html += `<tr><td>${n}</td>`;

        days.forEach(d => {
            if (follow[n][d]) {
                html += `<td class='follow-ok'>✓</td>`;
                countDays++;
            } else {
                html += `<td class='follow-empty'></td>`;
            }
        });

        html += `<td>${countDays}</td>`;
        html += `<td>${(teacherHours[n] || 0).toFixed(1)}</td>`;
        html += "</tr>";
    });

    html += "</table>";

    document.getElementById("followMatrix").innerHTML = html;
}
