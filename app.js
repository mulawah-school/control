const API = "https://script.google.com/macros/s/AKfycbzMXIRJPkQlkcnfaz1StwKivZEbghqL4u9XBfX_NnUzLtB24lVbrGHq6NNlbOqx0btTHw/exec";

/************* API CALLER **************/
async function call(action, params = {}) {
  const query = new URLSearchParams({ action, ...params }).toString();
  let res = await fetch(`${API}?${query}`);
  return res.json();
}

/************* توزيع الأيام **************/
async function loadDayDistribution() {
  let day = document.getElementById("daySelect").value;
  let data = await call("getDistribution", {day});

  let html = "";
  data.rows.forEach(r=>{
    html += `
      <tr>
        <td>${r.committee}</td>
        <td>${r.subject}</td>
        <td>${r.supervisors.join(" ، ")}</td>
      </tr>`;
  });

  document.getElementById("distTable").innerHTML = html;
}

async function runSmart(){
  let res = await call("runSmart");
  alert(res.message);
  loadDays();
}

/************* تحميل الأيام **************/
async function loadDays(){
  let days = await call("getDays");
  let sel = document.getElementById("daySelect");
  if (!sel) return;
  sel.innerHTML = days.map(d=>`<option>${d}</option>`).join("");
}

/************* تحميل المعلمين **************/
async function loadTeachers(){
  let data = await call("getTeachers");
  let sel = document.getElementById("teacherSelect");
  if(sel) sel.innerHTML = data.map(t=>`<option>${t.name}</option>`).join("");
}

/************* تحميل المواد **************/
async function loadSubjects(){
  let subs = await call("getSubjects");
  let sel = document.getElementById("subjectSelect");
  if(sel) sel.innerHTML = subs.map(s=>`<option>${s}</option>`).join("");
}

/************* الاستثناءات **************/
async function loadExceptions(){
  let day = document.getElementById("daySelect").value;
  let data = await call("getExceptions", {day});

  let html = "";

  data.teachers.forEach(t=>{
    html += `<li class="list-group-item">👨‍🏫 ${t}</li>`;
  });

  data.subjects.forEach(s=>{
    html += `<li class="list-group-item">📘 ${s}</li>`;
  });

  document.getElementById("excList").innerHTML = html;
}

async function addTeacherException(){
  let day = document.getElementById("daySelect").value;
  let teacher = document.getElementById("teacherSelect").value;

  await call("saveException", {day, teacher});
  loadExceptions();
}

async function addSubjectException(){
  let day = document.getElementById("daySelect").value;
  let subject = document.getElementById("subjectSelect").value;

  await call("saveException", {day, subject});
  loadExceptions();
}

/************* متابعة **************/
async function loadFollow(){
  let data = await call("followMatrix");

  let html = "";

  data.teachers.forEach(t=>{
    html += `
      <tr>
        <td>${t.name}</td>
        <td>${t.hours}</td>
        <td>${t.days.length}</td>
        <td>${t.days.join(" ، ")}</td>
      </tr>`;
  });

  document.getElementById("followTable").innerHTML = html;
}

/************* AUTO LOAD **************/
window.onload = ()=>{
  loadDays();
  loadTeachers();
  loadSubjects();
  loadFollow();
};
