import { sb, USERS, MOTIVATIONS, getMotivation, EMAIL_TO_USERNAME, VIEW_ONLY_USERS } from './config.js';
import { today, fmt, countWorkingDays, addWorkingDays, isOverdue, escHtml, sanitizeUrl } from './utils.js';
import Chart from 'chart.js/auto';
import * as XLSX from 'xlsx';
import emailjs from '@emailjs/browser';


// ═══════════════════════════════════════════════════════
// DATA STORE
// ═══════════════════════════════════════════════════════
function canWrite(){ return !!currentUser && !VIEW_ONLY_USERS.includes(currentUser); }

let DB = { tasks:[], comments:[], attachments:[], recycleBin:[], stepProgress:{}, emailLog:[], dailySummary:{}, routineTasks:[], routineCompletions:[], personalTasks:[] };
let currentUser = null;
let charts = {};

async function loadDB(){
  try {
    // Column lists trimmed to exactly what's mapped below (was select('*') —
    // fetched every column including ones never read, on every login).
    // ff_email_log is also the fastest-growing table (already the largest by
    // far), so it's capped to the most recent 300 rather than fetched in full.
    const [tRes, cRes, eRes, spRes, rbRes, dsRes, rtRes, rcRes, ptRes] = await Promise.all([
      sb.from('ff_tasks').select('id,title,description,assigned_to,assigned_by,assigned_date,deadline_date,completion_date,priority,status,external_link,is_recurring,recurrence_type,urgent_flag,steps,attachments,created_date,updated_date,checked_by,checked_date,is_doc_request,category'),
      sb.from('ff_comments').select('id,task_id,user_id,text,created_date'),
      sb.from('ff_email_log').select('id,type,subject,to_email,from_email,body,task_id,timestamp').order('timestamp',{ascending:false}).limit(300),
      sb.from('ff_step_progress').select('task_id,step_id,done'),
      sb.from('ff_recycle_bin').select('id,title,description,assigned_to,assigned_by,assigned_date,deadline_date,completion_date,priority,status,external_link,is_recurring,recurrence_type,urgent_flag,steps,attachments,created_date,updated_date,deleted_by,deleted_date'),
      sb.from('ff_daily_summary').select('date,employee,completed_count,overdue_count'),
      sb.from('ff_recurring_tasks').select('id,title,frequency,assigned_to,created_by,created_date,reminder_time,active,remarks,link,archived,archived_date,attachments,deadline_date,deadline_time'),
      sb.from('ff_recurring_completions').select('id,task_id,due_date,completed_at,completed_by,status,deadline_date,deadline_time,auto_archive_at'),
      sb.from('ff_personal_tasks').select('id,title,description,owner,deadline_date,deadline_time,recurrence,status,completion_date,created_date,link,attachments')
    ]);
    DB.tasks = (tRes.data||[]).map(r=>({
      id:r.id, title:r.title, description:r.description, assignedTo:r.assigned_to,
      assignedBy:r.assigned_by, assignedDate:r.assigned_date, deadlineDate:r.deadline_date,
      completionDate:r.completion_date, priority:r.priority, status:r.status,
      externalLink:r.external_link, isRecurring:r.is_recurring, recurrenceType:r.recurrence_type,
      urgentFlag:r.urgent_flag, steps:r.steps||[], attachments:r.attachments||[],
      createdDate:r.created_date, updatedDate:r.updated_date,
      checkedBy:r.checked_by||null, checkedDate:r.checked_date||null,
      isDocRequest:r.is_doc_request||false,
      category:r.category||''
    }));
    DB.comments = (cRes.data||[]).map(r=>({
      id:r.id, taskId:r.task_id, userId:r.user_id, text:r.text, createdDate:r.created_date
    }));
    // Fetched newest-first (see the .order() above) so the 300-row cap keeps
    // the most recent emails; reversed back to oldest-first here so DB.emailLog
    // keeps the exact same storage order it always had — renderEmailLog() and
    // every other reader still does its own .reverse() to show newest-first.
    DB.emailLog = (eRes.data||[]).map(r=>({
      id:r.id, type:r.type, subject:r.subject, to:r.to_email, from:r.from_email,
      body:r.body, taskId:r.task_id, timestamp:r.timestamp
    })).reverse();
    DB.stepProgress = {};
    (spRes.data||[]).forEach(r=>{
      if(!DB.stepProgress[r.task_id]) DB.stepProgress[r.task_id]={};
      DB.stepProgress[r.task_id][r.step_id]=r.done;
    });
    DB.recycleBin = (rbRes.data||[]).map(r=>({
      id:r.id, title:r.title, description:r.description, assignedTo:r.assigned_to,
      assignedBy:r.assigned_by, assignedDate:r.assigned_date, deadlineDate:r.deadline_date,
      completionDate:r.completion_date, priority:r.priority, status:r.status,
      externalLink:r.external_link, isRecurring:r.is_recurring, recurrenceType:r.recurrence_type,
      urgentFlag:r.urgent_flag, steps:r.steps||[], attachments:r.attachments||[],
      createdDate:r.created_date, updatedDate:r.updated_date,
      deletedBy:r.deleted_by, deletedDate:r.deleted_date
    }));
    DB.dailySummary={};
    (dsRes.data||[]).forEach(r=>{
      if(!DB.dailySummary[r.date]) DB.dailySummary[r.date]={};
      DB.dailySummary[r.date][r.employee]={completed_count:r.completed_count, overdue_count:r.overdue_count};
    });
    DB.routineTasks=(rtRes.data||[]).map(r=>({
      id:r.id, title:r.title, frequency:r.frequency, assignedTo:r.assigned_to,
      createdBy:r.created_by, createdDate:r.created_date, reminderTime:r.reminder_time||'09:00', active:r.active!==false,
      remarks:r.remarks||'', link:r.link||'', archived:r.archived||false, archivedDate:r.archived_date||null, attachments:r.attachments||[],
      // Pre-existing bug fixed here: deadlineDate/deadlineTime are read
      // throughout the routine-task logic (completion deadlines, exports)
      // but were never populated from the fetched row, so they were always
      // undefined for any routine task loaded after a page refresh.
      deadlineDate:r.deadline_date||'', deadlineTime:r.deadline_time||''
    }));
    DB.routineCompletions=(rcRes.data||[]).map(r=>({
      id:r.id, taskId:r.task_id, dueDate:r.due_date, completedAt:r.completed_at,
      completedBy:r.completed_by, status:r.status||'pending',
      deadlineDate:r.deadline_date||'', deadlineTime:r.deadline_time||'',
      autoArchiveAt:r.auto_archive_at||null
    }));
    DB.personalTasks=(ptRes.data||[]).map(r=>({
      id:r.id, title:r.title, description:r.description||'', owner:r.owner,
      deadlineDate:r.deadline_date, deadlineTime:r.deadline_time||'17:00',
      recurrence:r.recurrence||'none', status:r.status||'Pending',
      completionDate:r.completion_date, createdDate:r.created_date,
      link:r.link||'', attachments:r.attachments||[]
    }));
    return true;
  } catch(e){
    console.error('loadDB error',e);
    showToast('Could not load your data — check your connection and refresh','error');
    return false;
  }
}

function saveDB(){}

async function snapshotDailySummary(){
  if(!canWrite()) return; // view-only users don't write daily stats
  const d=today();
  for(const uid of ['de','mitiksha']){
    const completed=DB.tasks.filter(t=>t.assignedTo===uid&&t.completionDate&&t.completionDate.startsWith(d)).length;
    const overdue=DB.tasks.filter(t=>t.assignedTo===uid&&(t.status==='Overdue'||(t.status!=='Completed'&&t.deadlineDate&&t.deadlineDate<d))).length;
    if(!DB.dailySummary[d]) DB.dailySummary[d]={};
    DB.dailySummary[d][uid]={completed_count:completed, overdue_count:overdue};
    await runDb(sb.from('ff_daily_summary').upsert({date:d, employee:uid, completed_count:completed, overdue_count:overdue}), 'save daily summary', {silent:true});
  }
}

// Central DB-write wrapper: awaits a Supabase query, surfaces any error to the
// user (and console) instead of failing silently, and returns a success flag.
// Usage: const ok = await runDb(sb.from('x').insert(...), 'save task');
async function runDb(query, action, {silent=false}={}){
  try{
    const { error } = await query;
    if(error){
      console.error(`DB error [${action}]:`, error);
      if(!silent) showToast(`Couldn't ${action}: ${error.message||'database error'}`, 'error');
      return false;
    }
    return true;
  }catch(e){
    console.error(`DB exception [${action}]:`, e);
    if(!silent) showToast(`Couldn't ${action} — check your connection`, 'error');
    return false;
  }
}

function taskToRow(t){
  return {
    id:t.id, title:t.title, description:t.description||'', assigned_to:t.assignedTo,
    assigned_by:t.assignedBy, assigned_date:t.assignedDate, deadline_date:t.deadlineDate,
    completion_date:t.completionDate, priority:t.priority, status:t.status,
    external_link:t.externalLink||'', is_recurring:t.isRecurring||false,
    recurrence_type:t.recurrenceType||'', urgent_flag:t.urgentFlag||false,
    steps:t.steps||[], attachments:t.attachments||[],
    created_date:t.createdDate, updated_date:t.updatedDate,
    checked_by:t.checkedBy||null, checked_date:t.checkedDate||null,
    is_doc_request:t.isDocRequest||false,
    category:t.category||''
  };
}

async function dbInsertTask(t){
  return runDb(sb.from('ff_tasks').insert(taskToRow(t)), 'save task');
}
async function dbUpdateTask(t){
  return runDb(sb.from('ff_tasks').update(taskToRow(t)).eq('id',t.id), 'update task');
}
async function dbDeleteTask(id){
  return runDb(sb.from('ff_tasks').delete().eq('id',id), 'delete task');
}
async function dbInsertComment(c){
  return runDb(sb.from('ff_comments').insert({id:c.id, task_id:c.taskId, user_id:c.userId, text:c.text, created_date:c.createdDate}), 'save comment');
}
async function dbInsertEmail(e){
  // Email logging is a background side-effect — record failures quietly, don't nag the user.
  return runDb(sb.from('ff_email_log').insert({id:e.id, type:e.type, subject:e.subject, to_email:e.to, from_email:e.from, body:e.body, task_id:e.taskId||null, timestamp:e.timestamp}), 'log email', {silent:true});
}
async function dbUpsertStepProgress(taskId, stepId, done){
  return runDb(sb.from('ff_step_progress').upsert({task_id:taskId, step_id:stepId, done}), 'save step progress');
}
async function dbInsertRecycleBin(t){
  return runDb(sb.from('ff_recycle_bin').insert({
    id:t.id, title:t.title, description:t.description||'', assigned_to:t.assignedTo,
    assigned_by:t.assignedBy, assigned_date:t.assignedDate, deadline_date:t.deadlineDate,
    completion_date:t.completionDate, priority:t.priority, status:t.status,
    external_link:t.externalLink||'', is_recurring:t.isRecurring||false,
    recurrence_type:t.recurrenceType||'', urgent_flag:t.urgentFlag||false,
    steps:t.steps||[], attachments:t.attachments||[],
    created_date:t.createdDate, updated_date:t.updatedDate,
    deleted_by:t.deletedBy, deleted_date:t.deletedDate
  }), 'archive task');
}
async function dbDeleteRecycleBin(id){
  return runDb(sb.from('ff_recycle_bin').delete().eq('id',id), 'remove from archive');
}


// ═══════════════════════════════════════════════════════
// AUTH
// ═══════════════════════════════════════════════════════
async function login(){
  const uid = document.getElementById('login-user').value;
  const pass = document.getElementById('login-pass').value;
  const email = USERS[uid]?.email;
  if(!email){ showToast('Unknown account','error'); return; }
  if(!pass){ showToast('Enter your password','error'); return; }
  const btn = document.querySelector('#auth-screen .btn-lg');
  if(btn){ btn.disabled = true; btn.textContent = 'Signing in…'; }
  const { data, error } = await sb.auth.signInWithPassword({ email, password: pass });
  if(btn){ btn.disabled = false; btn.textContent = 'Sign In'; }
  if(error || !data?.session){ showToast('Invalid email or password','error'); return; }
  currentUser = uid;
  localStorage.setItem('ffUser', uid); // convenience only: pre-selects the dropdown next time
  initApp();
}

async function forgotPassword(){
  const uid = document.getElementById('login-user').value;
  const email = USERS[uid]?.email;
  if(!email){ showToast('Select an account first','error'); return; }
  const { error } = await sb.auth.resetPasswordForEmail(email, {
    redirectTo: window.location.origin + window.location.pathname
  });
  if(error){ showToast('Failed to send reset email','error'); return; }
  showToast('Password reset link sent to ' + email, 'success');
}
async function logout(){
  await sb.auth.signOut();
  currentUser=null;
  localStorage.removeItem('ffUser');
  document.getElementById('app').classList.add('hidden');
  document.getElementById('auth-screen').classList.remove('hidden');
}

async function initApp(){
  document.getElementById('auth-screen').classList.add('hidden');
  document.getElementById('app').classList.remove('hidden');
  const u = USERS[currentUser];
  document.getElementById('sb-avatar').textContent = u.avatar;
  document.getElementById('sb-avatar').style.background = `linear-gradient(135deg,${u.color},${u.color}cc)`;
  document.getElementById('sb-name').textContent = u.name;
  document.getElementById('sb-role').textContent = u.role;
  buildNav();
  document.getElementById('new-task-btn').style.display = (currentUser==='sumudu'||currentUser==='de')?'':'none';
  document.getElementById('doc-request-btn').style.display = (currentUser==='de'||currentUser==='mitiksha')?'':'none';
  const hasLocal=localStorage.getItem('ffdb');
  document.getElementById('migrate-btn').style.display = hasLocal?'':'none';
  await loadDB();
  await snapshotDailySummary();
  await checkMainTaskRecurrence();
  ensureRoutineCompletions();
  startRoutineReminders();
  setInterval(checkAutoMorningReport, 60000);
  checkAutoMorningReport();
  initCalendar();
  updateNotifBadge();
  navigateTo('dashboard');
  showToast(`Welcome back, ${u.name}! ${getMotivation()}`,'success');
}

// ═══════════════════════════════════════════════════════
// NAVIGATION
// ═══════════════════════════════════════════════════════
const NAV = [
  {id:'dashboard',icon:'\u{1F4CA}',label:'Dashboard',all:true},
  {id:'tasks',icon:'\u{1F4CB}',label:'All Tasks',roles:['sumudu','de','mitiksha','trupal']},
  {id:'my-tasks',icon:'✅',label:'My Tasks',roles:['sumudu','de','mitiksha']},
  {id:'my-requests',icon:'📩',label:'My Requests',roles:['de','mitiksha']},
  {id:'personal-tasks',icon:'📝',label:'Personal Tasks',roles:['sumudu']},
  {id:'schedule',icon:'\u{1F4C5}',label:'Schedule',all:true},
  {id:'performance',icon:'\u{1F3C6}',label:'Performance',all:true},
  {id:'reports',icon:'\u{1F4C8}',label:'Reports',all:true},
  {id:'routine-tasks',icon:'🔄',label:'Routine Tasks',all:true},
  {id:'recycle',icon:'📦',label:'Archived Tasks',roles:['sumudu']},
  {id:'emails',icon:'📧',label:'Email Log',roles:['sumudu','de','mitiksha']},
];

function navLabel(item){
  if(item.id==='my-tasks' && currentUser==='sumudu') return 'Requests';
  if(item.id==='routine-tasks' && currentUser==='sumudu') return 'Team Routines';
  return item.label;
}

function getPendingRequestsCount(){
  if(currentUser!=='sumudu') return 0;
  return DB.tasks.filter(t=>t.isDocRequest && t.assignedTo==='sumudu' && t.status!=='Completed').length;
}

function buildNav(){
  const nav = document.getElementById('sidebar-nav');
  nav.innerHTML = '';
  NAV.forEach(item=>{
    if(!item.all && !(item.roles && item.roles.includes(currentUser))) return;
    const el = document.createElement('div');
    el.className = 'nav-item';
    el.dataset.page = item.id;
    el.setAttribute('role','menuitem');
    el.setAttribute('tabindex','0');
    el.innerHTML = `<span class="icon">${item.icon}</span><span class="nav-label">${navLabel(item)}</span>${item.id==='my-tasks'?'<span class="nav-badge hidden" id="nav-badge-requests"></span>':''}`;
    el.onclick = ()=>{ navigateTo(item.id); closeSidebar(); };
    el.onkeydown = (e)=>{ if(e.key==='Enter'||e.key===' '){e.preventDefault();navigateTo(item.id);closeSidebar();} };
    nav.appendChild(el);
  });
  updateNavBadges();
}

function updateNavBadges(){
  const badge=document.getElementById('nav-badge-requests');
  if(!badge) return;
  const count=getPendingRequestsCount();
  if(count>0){ badge.textContent=count; badge.classList.remove('hidden'); }
  else { badge.classList.add('hidden'); }
}

function navigateTo(pageId){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  const pg = document.getElementById('page-'+pageId);
  if(pg) pg.classList.add('active');
  const nav = document.querySelector(`.nav-item[data-page="${pageId}"]`);
  if(nav) nav.classList.add('active');
  const labels = {dashboard:'Dashboard',tasks:'All Tasks','my-tasks':(currentUser==='sumudu'?'Requests':'My Tasks'),'my-requests':'My Requests','personal-tasks':'Personal Tasks','routine-tasks':(currentUser==='sumudu'?'Team Routines':'Routine Tasks'),performance:'Performance',schedule:'Monthly Schedule',reports:'Reports & Analytics',recycle:'Archived Tasks',emails:'Email Log'};
  document.getElementById('topbar-title').textContent = labels[pageId]||pageId;
  if(pageId==='dashboard') renderDashboard();
  if(pageId==='tasks'){
    renderTasks();
    const hideTrupal=currentUser==='trupal';
    document.getElementById('export-btn-tasks').style.display=hideTrupal?'none':'';
    document.getElementById('import-btn-tasks').style.display=hideTrupal?'none':'';
  }
  if(pageId==='my-tasks') renderMyTasks();
  if(pageId==='performance'){ initPerfFilters(); renderPerformance(); }
  if(pageId==='schedule'){ initScheduleFilter(); renderSchedule(); }
  if(pageId==='recycle') renderRecycle();
  if(pageId==='reports') renderReports();
  if(pageId==='emails') renderEmailLog();
  if(pageId==='my-requests') renderMyRequests();
  if(pageId==='personal-tasks') loadPersonalTasks();
  if(pageId==='routine-tasks') initRoutinePage();
}

// ═══════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════
function renderDashboard(){
  const tasks = ((currentUser==='sumudu'||currentUser==='trupal') ? DB.tasks.filter(t=>t.assignedTo==='de'||t.assignedTo==='mitiksha') : DB.tasks.filter(t=>t.assignedTo===currentUser)).filter(t=>!t.isDocRequest);
  const pending = tasks.filter(t=>t.status==='Pending').length;
  const inprog  = tasks.filter(t=>t.status==='In Progress').length;
  const done    = tasks.filter(t=>t.status==='Completed').length;
  const overdue = tasks.filter(t=>t.status==='Overdue').length;

  // Motivation banner
  const motivHtml=`<div class="dash-motivation"><span class="motiv-icon">✨</span><span>${getMotivation()} ✨</span></div>`;

  const statData = [
    {label:'Total Tasks',value:tasks.length,color:'#4a3aaf',bg:'#5b4dc7',icon:'<svg width="22" height="22" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" viewBox="0 0 24 24"><rect x="3" y="3" width="18" height="18" rx="3"/><path d="M8 7h8M8 12h5M8 17h8"/></svg>'},
    {label:'Pending',value:pending,color:'#c88d0a',bg:'#e5a819',icon:'<svg width="22" height="22" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg>'},
    {label:'Completed',value:done,color:'#16864a',bg:'#21a85e',icon:'<svg width="22" height="22" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-5"/></svg>'},
    {label:'Overdue',value:overdue,color:'#c42828',bg:'#e53e3e',icon:'<svg width="22" height="22" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" viewBox="0 0 24 24"><path d="M12 2L2 20h20L12 2z"/><path d="M12 10v4M12 17h.01"/></svg>'},
  ];

  const statFilters = ['','Pending','Completed','Overdue'];
  const personalTasksCard = currentUser==='sumudu' ? `
    <div class="stat-card clickable" onclick="navigateTo('personal-tasks')">
      <div style="display:flex;align-items:center;gap:14px">
        <div class="stat-icon" style="background:#7c3aed;border-radius:var(--radius);color:#fff"><svg width="22" height="22" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" viewBox="0 0 24 24"><path d="M4 4h11l5 5v11H4z"/><path d="M15 4v5h5"/><path d="M8 13h8M8 17h5"/></svg></div>
        <div class="stat-info"><div class="value" style="color:#7c3aed;font-size:18px">Personal Tasks</div><div class="label">Open →</div></div>
      </div>
    </div>` : '';
  document.getElementById('dash-stats').innerHTML = motivHtml + `<div class="stats-grid">${statData.map((s,i)=>`
    <div class="stat-card clickable" onclick="dashStatClick('${statFilters[i]}')">
      <div style="display:flex;align-items:center;gap:14px">
        <div class="stat-icon" style="background:${s.bg};border-radius:var(--radius);color:#fff">${s.icon}</div>
        <div class="stat-info"><div class="value" style="color:${s.color}">${s.value}</div><div class="label">${s.label}</div></div>
      </div>
    </div>`).join('')}${personalTasksCard}</div>`;

  const isDark = document.documentElement.getAttribute('data-theme')==='dark';
  const textColor = isDark ? '#9499b3' : '#5c5f77';
  const gridColor = isDark ? '#2a2b3d' : '#e4e7ef';

  destroyChart('chart-status');
  charts['chart-status'] = new Chart(document.getElementById('chart-status'),{
    type:'doughnut',
    data:{ labels:['Pending','In Progress','Completed','Overdue'],
           datasets:[{data:[pending,inprog,done,overdue],backgroundColor:['#f59e0b','#3b82f6','#22c55e','#ef4444'],borderWidth:0,borderRadius:4}]},
    options:{ responsive:true, maintainAspectRatio:false, cutout:'72%', plugins:{ legend:{ position:'bottom', labels:{ color:textColor, padding:14, font:{size:12,weight:'500',family:'Inter'}, usePointStyle:true, pointStyleWidth:8 } } } }
  });

  destroyChart('chart-monthly');
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const now = new Date();
  const monthlyDone = months.map((_,m)=> tasks.filter(t=> t.status==='Completed' && t.completionDate && new Date(t.completionDate).getMonth()===m && new Date(t.completionDate).getFullYear()===now.getFullYear()).length);
  const monthlyTotal = months.map((_,m)=> tasks.filter(t=> t.assignedDate && new Date(t.assignedDate).getMonth()===m && new Date(t.assignedDate).getFullYear()===now.getFullYear()).length);
  charts['chart-monthly'] = new Chart(document.getElementById('chart-monthly'),{
    type:'bar',
    data:{ labels:months, datasets:[
      {label:'Assigned',data:monthlyTotal,backgroundColor:isDark?'#6366f1':'#6366f1cc',borderRadius:6,borderSkipped:false},
      {label:'Completed',data:monthlyDone,backgroundColor:isDark?'#22c55e':'#22c55ecc',borderRadius:6,borderSkipped:false}
    ]},
    options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{ labels:{ color:textColor, font:{size:12,weight:'500',family:'Inter'}, usePointStyle:true, pointStyleWidth:8 } } }, scales:{ x:{ticks:{color:textColor,font:{size:11,family:'Inter'}},grid:{display:false}}, y:{ticks:{color:textColor,font:{size:11,family:'Inter'}},grid:{color:gridColor},border:{dash:[4,4]}} } }
  });

  // My Performance / Completion Rate widget
  const myPerfEl=document.getElementById('dash-my-perf');
  if(myPerfEl){
    const myTasks = currentUser==='sumudu' ? DB.tasks : DB.tasks.filter(t=>t.assignedTo===currentUser);
    const myDone = myTasks.filter(t=>t.status==='Completed').length;
    const myPct = myTasks.length ? Math.round((myDone/myTasks.length)*100) : 0;
    const ringColor = myPct>=70?'#22c55e':myPct>=40?'#f59e0b':'#ef4444';
    const circ2=2*Math.PI*48;
    const dash2=(myPct/100)*circ2;
    myPerfEl.innerHTML=`<div class="my-perf-card">
      <h3>Task Completion Rate</h3>
      <div class="my-perf-ring">
        <svg width="120" height="120" viewBox="0 0 120 120">
          <circle cx="60" cy="60" r="48" fill="none" stroke="${isDark?'#2a2b3d':'#e4e7ef'}" stroke-width="10"/>
          <circle cx="60" cy="60" r="48" fill="none" stroke="${ringColor}" stroke-width="10" stroke-dasharray="${dash2} ${circ2}" stroke-linecap="round" style="transition:stroke-dasharray .8s var(--ease)"/>
        </svg>
        <div class="perf-pct" style="color:${ringColor}">${myPct}%</div>
      </div>
      <div class="my-perf-label">${myDone} of ${myTasks.length} tasks completed</div>
    </div>`;
  }

  if(currentUser==='sumudu'){
    const employees = ['de','mitiksha'];
    let html = '<div class="card" style="margin-top:8px"><div class="card-header"><span class="card-title">Employee Overview</span></div><div class="card-body"><div class="perf-grid">';
    employees.forEach(uid=>{
      const u=USERS[uid];
      const etasks=DB.tasks.filter(t=>t.assignedTo===uid);
      const score=calcScore(uid, new Date().getMonth(), new Date().getFullYear());
      const color = score>=80?'#22c55e':score>=60?'#f59e0b':'#ef4444';
      const circ = 2*Math.PI*36;
      const dash = (score/100)*circ;
      html+=`<div class="perf-card">
        <div class="emp-name">${u.name}</div>
        <div class="emp-role">${u.role}</div>
        <div class="score-ring" style="width:100px;height:100px">
          <svg width="100" height="100" viewBox="0 0 100 100">
            <circle cx="50" cy="50" r="36" fill="none" stroke="${isDark?'#2a2b3d':'#e4e7ef'}" stroke-width="8"/>
            <circle cx="50" cy="50" r="36" fill="none" stroke="${color}" stroke-width="8" stroke-dasharray="${dash} ${circ}" stroke-linecap="round" style="transition:stroke-dasharray .8s var(--ease)"/>
          </svg>
          <div class="score-num" style="color:${color}">${score}</div>
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;margin-top:8px;width:100%;text-align:left">
          <div class="score-bar-row"><span class="bar-label">Total</span><span style="font-weight:700;margin-left:auto">${etasks.length}</span></div>
          <div class="score-bar-row"><span class="bar-label">Completed</span><span style="font-weight:700;color:#22c55e;margin-left:auto">${etasks.filter(t=>t.status==='Completed').length}</span></div>
          <div class="score-bar-row"><span class="bar-label">Overdue</span><span style="font-weight:700;color:#ef4444;margin-left:auto">${etasks.filter(t=>t.status==='Overdue').length}</span></div>
        </div>
      </div>`;
    });
    html+='</div></div></div>';
    document.getElementById('dash-emp-perf').innerHTML = html;
  } else {
    document.getElementById('dash-emp-perf').innerHTML='';
  }
}

function destroyChart(id){ if(charts[id]){ charts[id].destroy(); delete charts[id]; } }

// ═══════════════════════════════════════════════════════
// TASK CRUD
// ═══════════════════════════════════════════════════════
function openNewTask(){
  const body = `
  <div class="form-grid">
    <div class="form-field full"><label>Task Title</label><input id="nt-title" placeholder="What needs to be done?"></div>
    <div class="form-field full"><label>Description</label><textarea id="nt-desc" placeholder="Add details..."></textarea></div>
    <div class="form-field"><label>Assign To</label>
      ${currentUser==='sumudu'?'<div id="nt-assign-group"><label style="display:flex;align-items:center;gap:6px;cursor:pointer;margin-bottom:6px"><input type="checkbox" value="de" class="nt-assign-cb"> Dev</label><label style="display:flex;align-items:center;gap:6px;cursor:pointer"><input type="checkbox" value="mitiksha" class="nt-assign-cb"> Mitiksha</label></div>':'<select id="nt-assign"><option value="">— Select —</option><option value="mitiksha">Mitiksha</option></select>'}
    </div>
    <div class="form-field"><label>Urgent?</label>
      <label class="urgent-check">
        <input type="checkbox" id="nt-urgent-flag" onchange="this.parentElement.classList.toggle('checked',this.checked)">
        🔴 Mark as Urgent
      </label>
    </div>
    <div class="form-field"><label>Assigned Date</label><input type="date" id="nt-adate" value="${today()}"></div>
    <div class="form-field"><label>Deadline Date</label><input type="date" id="nt-ddate" value="${today()}"></div>
    <div class="form-field full"><label>External Link</label><input id="nt-link" placeholder="https://drive.google.com/..."></div>
    <div class="form-field full"><label>Recurring</label>
      <select id="nt-recur"><option value="">No Repeat</option><option value="daily">Daily</option><option value="weekly">Weekly</option><option value="monthly">Monthly</option><option value="quarterly">Every Three Months</option><option value="annually">Annually</option></select>
    </div>
    <div class="form-field full">
      <label>Attach Files</label>
      <div class="file-upload-area" id="nt-file-area">
        <input type="file" multiple id="nt-files" onchange="previewFiles(this,'nt-files-list')">
        <div class="upload-icon">\u{1F4CE}</div>
        <p>Click or drag files here</p>
      </div>
      <div class="attached-files" id="nt-files-list"></div>
    </div>
    <div class="form-field full">
      <label>Checklist Steps</label>
      <div class="steps-list" id="nt-steps"></div>
      <div class="add-step-row">
        <input id="nt-step-input" placeholder="Add a step..." onkeydown="if(event.key==='Enter'){addStep('nt-steps','nt-step-input')}">
        <button class="btn btn-secondary btn-sm" onclick="addStep('nt-steps','nt-step-input')">Add</button>
      </div>
    </div>
  </div>`;
  openModal('New Task', body, [
    {label:'Cancel', action:'closeModal()', cls:'btn-secondary'},
    {label:'Create Task', action:'createTask()', cls:'btn-primary'}
  ]);
}

function addStep(listId, inputId){
  const inp = document.getElementById(inputId);
  const val = inp.value.trim();
  if(!val) return;
  const list = document.getElementById(listId);
  const item = document.createElement('div');
  item.className='step-item';
  item.innerHTML=`<input type="checkbox"><span class="step-text">${escHtml(val)}</span><button class="step-del" onclick="this.parentElement.remove()">\u{1F5D1}</button>`;
  list.appendChild(item);
  inp.value='';
  inp.focus();
}

function getSteps(listId){
  const steps=[];
  document.querySelectorAll(`#${listId} .step-item`).forEach(el=>{
    steps.push({ id:'step_'+Date.now()+Math.random(), text:el.querySelector('.step-text').textContent, done:el.querySelector('input').checked });
  });
  return steps;
}

async function createTask(){
  const title = document.getElementById('nt-title').value.trim();
  if(!title){ showToast('Task title is required','error'); return; }
  let assignees=[];
  if(currentUser==='sumudu'){
    document.querySelectorAll('.nt-assign-cb:checked').forEach(cb=>assignees.push(cb.value));
    if(!assignees.length){ showToast('Please select at least one assignee','error'); return; }
  } else {
    if(!document.getElementById('nt-assign').value){ showToast('Please select an assignee','error'); return; }
    assignees=[document.getElementById('nt-assign').value];
  }
  for(const assignee of assignees){
  const task = {
    id: 'task_'+Date.now()+'_'+assignee,
    title,
    description: document.getElementById('nt-desc').value.trim(),
    assignedTo: assignee,
    assignedBy: currentUser,
    assignedDate: document.getElementById('nt-adate').value || today(),
    deadlineDate: document.getElementById('nt-ddate').value || today(),
    completionDate: null,
    priority: document.getElementById('nt-urgent-flag').checked ? 'Urgent' : 'Normal',
    status: 'Pending',
    externalLink: document.getElementById('nt-link').value.trim(),
    isRecurring: document.getElementById('nt-recur').value!=='',
    recurrenceType: document.getElementById('nt-recur').value,
    createdDate: new Date().toISOString(),
    updatedDate: new Date().toISOString(),
    steps: getSteps('nt-steps'),
    attachments: [],
    urgentFlag: document.getElementById('nt-urgent-flag').checked
  };
  if(pendingFiles.length){
    task.attachments = await uploadFilesToStorage(task.id, pendingFiles);
  }
  DB.tasks.push(task);
  await dbInsertTask(task);
  simulateEmail('new_task', task);
  }
  pendingFiles = [];
  closeModal();
  showToast(`Task "${title}" assigned to ${assignees.map(a=>USERS[a].name).join(' & ')}`,'success');
  renderTasks();
  renderDashboard();
  updateNotifBadge();
}

function openDocRequest(){
  const body = `
  <div class="form-grid">
    <div class="form-field full"><label>Document / Request Title</label><input id="dr-title" placeholder="What document do you need?"></div>
    <div class="form-field full"><label>Details</label><textarea id="dr-desc" placeholder="Describe what you need..."></textarea></div>
    <div class="form-field full"><label>Attach Files (optional)</label>
      <div class="file-upload-area" id="dr-file-area">
        <input type="file" multiple id="dr-files" onchange="previewFiles(this,'dr-files-list')">
        <div class="upload-icon">\u{1F4CE}</div>
        <p>Click or drag files here</p>
      </div>
      <div class="attached-files" id="dr-files-list"></div>
    </div>
  </div>`;
  openModal('📩 Request to Manager', body, [
    {label:'Cancel', action:'closeModal()', cls:'btn-secondary'},
    {label:'Send Request', action:'createDocRequest()', cls:'btn-primary'}
  ]);
}

async function createDocRequest(){
  const title = document.getElementById('dr-title').value.trim();
  if(!title){ showToast('Please enter a title','error'); return; }
  const task = {
    id: 'req_'+Date.now(),
    title: '📄 ' + title,
    description: document.getElementById('dr-desc').value.trim(),
    assignedTo: 'sumudu',
    assignedBy: currentUser,
    assignedDate: today(),
    deadlineDate: null,
    completionDate: null,
    priority: 'Normal',
    status: 'Pending',
    externalLink: '',
    isRecurring: false,
    recurrenceType: '',
    urgentFlag: false,
    isDocRequest: true,
    createdDate: new Date().toISOString(),
    updatedDate: new Date().toISOString(),
    steps: [],
    attachments: []
  };
  if(pendingFiles.length){
    task.attachments = await uploadFilesToStorage(task.id, pendingFiles);
  }
  DB.tasks.push(task);
  await dbInsertTask(task);
  pendingFiles = [];
  closeModal();
  simulateEmail('doc_request', task);
  showToast(`Document request sent to Sumudu`,'success');
  renderTasks();
  renderDashboard();
}

let pendingFiles = [];
function previewFiles(input, listId){
  const list = document.getElementById(listId);
  if(!list) return;
  Array.from(input.files).forEach(f=>{
    pendingFiles.push(f);
    const div=document.createElement('div');
    div.className='attached-file';
    div.innerHTML=`<span class="file-icon">\u{1F4C4}</span><span class="file-name">${escHtml(f.name)}</span><button class="file-remove" onclick="this.parentElement.remove()">✕</button>`;
    list.appendChild(div);
  });
}

async function uploadFilesToStorage(taskId, files){
  const uploaded=[];
  for(const f of files){
    const path=`${taskId}/${Date.now()}_${f.name}`;
    const {error}=await sb.storage.from('attachments').upload(path, f);
    if(!error){
      const {data}=sb.storage.from('attachments').getPublicUrl(path);
      uploaded.push({name:f.name, url:data.publicUrl, path, uploadedBy:currentUser});
    }
  }
  return uploaded;
}

// ═══════════════════════════════════════════════════════
// RENDER TASKS TABLE
// ═══════════════════════════════════════════════════════
function getFilteredTasks(){
  let tasks = DB.tasks.filter(t=>!t.isDocRequest);
  if(currentUser==='de'){ tasks=tasks.filter(t=>t.assignedTo==='de'||t.assignedTo==='mitiksha'); }
  if(currentUser==='trupal'){ tasks=tasks.filter(t=>t.assignedTo==='de'||t.assignedTo==='mitiksha'); }
  if(currentUser==='mitiksha'){ tasks=tasks.filter(t=>t.assignedTo==='mitiksha'); }
  const st=document.getElementById('filter-status')?.value;
  const pr=document.getElementById('filter-priority')?.value;
  const as=document.getElementById('filter-assignee')?.value;
  const so=document.getElementById('filter-sort')?.value||'created_desc';
  if(st==='Pending') tasks=tasks.filter(t=>t.status==='Pending'||t.status==='Overdue');
  else if(st) tasks=tasks.filter(t=>t.status===st);
  if(pr) tasks=tasks.filter(t=>t.priority===pr);
  if(as) tasks=tasks.filter(t=>t.assignedTo===as);
  const pOrder={Urgent:0};
  tasks.sort((a,b)=>{
    const aUrg=a.status!=='Completed'&&a.deadlineDate&&a.deadlineDate<today()?1:0;
    const bUrg=b.status!=='Completed'&&b.deadlineDate&&b.deadlineDate<today()?1:0;
    if(aUrg!==bUrg) return bUrg-aUrg;
    if(so==='created_desc') return new Date(b.createdDate)-new Date(a.createdDate);
    if(so==='deadline_asc') return (a.deadlineDate||'').localeCompare(b.deadlineDate||'');
    if(so==='deadline_desc') return (b.deadlineDate||'').localeCompare(a.deadlineDate||'');
    if(so==='priority') return (pOrder[a.priority]??9)-(pOrder[b.priority]??9);
    return 0;
  });
  return tasks;
}

function renderTasks(){
  const tasks=getFilteredTasks();
  const tbody=document.getElementById('tasks-tbody');
  if(!tasks.length){ tbody.innerHTML=`<tr><td colspan="5"><div class="empty-state"><div class="empty-icon">\u{1F4ED}</div><p>No tasks found</p></div></td></tr>`; return; }
  tbody.innerHTML = tasks.map(t=>taskRow(t)).join('');
}

function renderMyTasks(){
  let tasks = currentUser==='de'
    ? DB.tasks.filter(t=>t.assignedTo==='de'||t.assignedTo==='mitiksha')
    : currentUser==='trupal' ? DB.tasks.filter(t=>t.assignedTo==='de'||t.assignedTo==='mitiksha')
    : DB.tasks.filter(t=>t.assignedTo===currentUser);
  const st=document.getElementById('my-filter-status')?.value;
  const pr=document.getElementById('my-filter-priority')?.value;
  if(st==='Pending') tasks=tasks.filter(t=>t.status==='Pending'||t.status==='Overdue');
  else if(st) tasks=tasks.filter(t=>t.status===st);
  if(pr) tasks=tasks.filter(t=>t.priority===pr);
  const tbody=document.getElementById('my-tasks-tbody');
  if(!tasks.length){ tbody.innerHTML=`<tr><td colspan="5"><div class="empty-state"><div class="empty-icon">\u{1F4ED}</div><p>No tasks assigned to you</p></div></td></tr>`; return; }
  tbody.innerHTML=tasks.map(t=>{
    const checkedBadge = t.checkedBy ? `<span class="badge badge-checked" style="margin-top:4px;display:inline-block">✔ Checked by ${USERS[t.checkedBy]?.name||t.checkedBy}</span>` : '';
    return `<tr style="cursor:pointer" onclick="viewTask('${t.id}')">
      <td><div class="task-title-cell"><div class="title ${t.status==='Completed'?'completed':''}">${(t.urgentFlag||(t.deadlineDate&&t.deadlineDate<today()&&t.status!=='Completed'))&&!t.isDocRequest?'<span class="urgent-banner">🔴 Urgent</span> ':''}${escHtml(t.title)}${t.isRecurring?' \u{1F504}':''}</div><div class="desc">${escHtml(t.description||'')}</div>${checkedBadge}</div></td>
      <td>${statusBadge(t.status)}</td>
      <td style="color:${t.deadlineDate&&t.deadlineDate<today()&&t.status!=='Completed'?'var(--danger)':'inherit'};font-weight:500">${t.isDocRequest?'—':fmt(t.deadlineDate)}</td>
      <td>${t.externalLink?`<a href="${sanitizeUrl(t.externalLink)}" target="_blank" rel="noopener" class="link-chip">Open</a>`:'—'}</td>
      <td>
        <div style="display:flex;gap:6px;flex-wrap:wrap" onclick="event.stopPropagation()">
          <button class="btn btn-secondary btn-sm" onclick="viewTask('${t.id}')">View</button>
          ${t.status!=='Completed'&&currentUser!=='trupal'?`<button class="btn btn-success btn-sm" onclick="markComplete('${t.id}')">Done</button>`:''}
          ${currentUser!=='trupal'?`<button class="btn btn-secondary btn-sm" onclick="openRemarks('${t.id}')">Remark</button>`:''}
        </div>
      </td>
    </tr>`;
  }).join('');
}

function renderMyRequests(){
  const requests = DB.tasks.filter(t=>t.isDocRequest && t.assignedBy===currentUser);
  const tbody = document.getElementById('my-requests-tbody');
  if(!requests.length){ tbody.innerHTML=`<tr><td colspan="4"><div class="empty-state"><div class="empty-icon">📩</div><p>No requests sent yet</p></div></td></tr>`; return; }
  tbody.innerHTML = requests.map(t=>{
    const comments = DB.comments.filter(c=>c.taskId===t.id);
    const attachments = (t.attachments||[]).filter(f=>typeof f==='object'&&f.url);
    const statusColor = t.status==='Completed'?'var(--success)':t.status==='In Progress'?'var(--info)':'var(--warning)';
    let responseHtml='';
    if(t.status==='Completed'){
      responseHtml=`<div style="margin-top:8px;font-size:12.5px">`;
      if(comments.length){
        responseHtml+=`<div style="color:var(--text3);font-weight:600;margin-bottom:4px">Manager Remarks:</div>`;
        responseHtml+=comments.map(c=>`<div style="background:var(--surface2);padding:8px 12px;border-radius:var(--radius);margin-bottom:4px;border:1px solid var(--border-subtle)"><strong style="color:var(--text2)">${USERS[c.userId]?.name||c.userId}:</strong> ${escHtml(c.text)}</div>`).join('');
      }
      if(attachments.length){
        responseHtml+=`<div style="color:var(--text3);font-weight:600;margin-top:6px;margin-bottom:4px">Attachments:</div>`;
        responseHtml+=attachments.map(f=>`<a href="${sanitizeUrl(f.url)}" target="_blank" rel="noopener" class="attached-file" style="text-decoration:none;display:inline-flex;margin-right:6px"><span class="file-icon">📄</span><span class="file-name">${escHtml(f.name)}</span><span style="color:var(--primary);font-size:11px;font-weight:600">⬇</span></a>`).join('');
      }
      responseHtml+=`</div>`;
    }
    return `<tr>
      <td><div class="task-title-cell"><div class="title ${t.status==='Completed'?'completed':''}">${escHtml(t.title)}</div><div class="desc">${escHtml(t.description||'')}</div>${responseHtml}</div></td>
      <td><span style="color:${statusColor};font-weight:700">${t.status}</span></td>
      <td style="font-weight:500">${fmt(t.assignedDate)}</td>
      <td><button class="btn btn-secondary btn-sm" onclick="viewTask('${t.id}')">View</button></td>
    </tr>`;
  }).join('');
}

function taskRow(t){
  const checkedBadge = t.checkedBy ? `<span class="badge badge-checked">✔ Checked by ${USERS[t.checkedBy]?.name||t.checkedBy}</span>` : '';
  const canCheck = currentUser==='de' && t.assignedTo==='mitiksha' && !t.checkedBy && t.status==='Completed';
  const desc = t.description || '';
  const joshuaBadge = t.category==='joshua' ? '<span style="background:#7c3aed;color:#fff;padding:2px 8px;border-radius:99px;font-size:10px;font-weight:700;margin-left:6px">Joshua</span>' : '';
  return `<tr style="cursor:pointer" onclick="viewTask('${t.id}')">
    <td style="min-width:260px"><div class="task-title-cell"><div class="title ${t.status==='Completed'?'completed':''}">${(t.urgentFlag||(t.deadlineDate&&t.deadlineDate<today()&&t.status!=='Completed'))&&!t.isDocRequest?'<span class="urgent-banner">🔴 Urgent</span> ':''}${escHtml(t.title)}${t.isRecurring?' \u{1F504}':''}${joshuaBadge}</div><div class="desc" style="font-size:12px;color:var(--text3);margin-top:3px;line-height:1.4">${escHtml(desc)}</div>${checkedBadge}</div></td>
    <td style="font-size:11px;white-space:nowrap;padding:6px 4px"><div class="user-chip" style="font-size:11px;gap:4px"><div class="chip-avatar" style="width:18px;height:18px;font-size:9px">${USERS[t.assignedTo]?.avatar||'?'}</div>${USERS[t.assignedTo]?.name||t.assignedTo}</div></td>
    <td style="font-size:11px;padding:6px 4px">${statusBadge(t.status)}</td>
    <td style="font-size:11px;color:${t.deadlineDate&&t.deadlineDate<today()&&t.status!=='Completed'?'var(--danger)':'inherit'};font-weight:500;white-space:nowrap;padding:6px 4px">${t.isDocRequest?'—':fmt(t.deadlineDate)}</td>
    <td>
      <div style="display:flex;gap:6px;flex-wrap:nowrap;justify-content:flex-end;white-space:nowrap" onclick="event.stopPropagation()">
        <button class="btn btn-secondary btn-sm" style="font-size:10px;padding:3px 6px" onclick="viewTask('${t.id}')">View</button>
        ${currentUser==='sumudu'?`<button class="btn btn-warning btn-sm" style="font-size:10px;padding:3px 6px" onclick="editTask('${t.id}')">Edit</button>`:''}
        ${canCheck?`<button class="btn btn-checked btn-sm" style="font-size:10px;padding:3px 6px" onclick="markChecked('${t.id}')">✔</button>`:''}
        ${t.status!=='Completed'&&currentUser!=='trupal'?`<button class="btn btn-success btn-sm" style="font-size:10px;padding:3px 6px" onclick="markComplete('${t.id}')">Done</button>`:''}
        ${currentUser!=='trupal'?`<button class="btn btn-secondary btn-sm" style="font-size:10px;padding:3px 6px" onclick="openRemarks('${t.id}')">Note</button>`:''}
        ${currentUser==='sumudu'?`<button class="btn btn-danger btn-sm" style="font-size:10px;padding:3px 6px" onclick="archiveTask('${t.id}')">Archive</button>`:''}
        <label style="display:inline-flex;align-items:center;gap:3px;cursor:pointer;font-size:10px;font-weight:600;color:${t.category==='joshua'?'#7c3aed':'var(--text3)'};padding:3px 6px;border-radius:var(--radius-sm);background:${t.category==='joshua'?'#ede9fe':'var(--surface2)'}"><input type="checkbox" ${t.category==='joshua'?'checked':''} onchange="toggleJoshuaCategory('${t.id}')" style="accent-color:#7c3aed">Joshua</label>
      </div>
    </td>
  </tr>`;
}

function statusBadge(s){
  const map={Pending:'pending','In Progress':'inprogress',Completed:'completed',Overdue:'overdue'};
  return `<span class="badge badge-${map[s]||'pending'}">${s}</span>`;
}

async function toggleJoshuaCategory(tid){
  const t=DB.tasks.find(x=>x.id===tid);
  if(!t) return;
  t.category = t.category==='joshua' ? '' : 'joshua';
  t.updatedDate=new Date().toISOString();
  await dbUpdateTask(t);
  renderTasks(); renderMyTasks();
  showToast(t.category==='joshua'?'Marked as Joshua task':'Removed from Joshua','info');
}

async function markChecked(tid){
  const t=DB.tasks.find(x=>x.id===tid);
  if(!t) return;
  t.checkedBy=currentUser;
  t.checkedDate=new Date().toISOString();
  t.updatedDate=new Date().toISOString();
  await dbUpdateTask(t);
  ['mitiksha','de','sumudu'].forEach(uid=>{
    if(uid===currentUser) return;
    const key='seenTasks_'+uid;
    const seen=JSON.parse(localStorage.getItem(key)||'[]').filter(id=>id!==tid);
    localStorage.setItem(key, JSON.stringify(seen));
  });
  showToast(`Task verified by ${USERS[currentUser].name}`,'success');
  renderTasks(); renderMyTasks(); updateNotifBadge();
}

function getSeenTaskIds(){
  try{ return JSON.parse(localStorage.getItem('seenTasks_'+currentUser)||'[]'); }catch(e){ return []; }
}
function setSeenTaskIds(ids){
  localStorage.setItem('seenTasks_'+currentUser, JSON.stringify(ids));
}
function getUnseenTasks(){
  const seen=new Set(getSeenTaskIds());
  let tasks=DB.tasks.filter(t=>{
    if(seen.has(t.id)) return false;
    if(t.assignedTo===currentUser) return true;
    if(currentUser==='sumudu') return true;
    if(currentUser==='de' && t.assignedTo==='mitiksha') return true;
    return false;
  });
  return tasks;
}
function updateNotifBadge(){
  const unseen=getUnseenTasks();
  const badge=document.getElementById('notif-badge');
  if(!badge) return;
  if(unseen.length>0){
    badge.textContent=unseen.length;
    badge.classList.remove('hidden');
  } else {
    badge.classList.add('hidden');
  }
  updateNavBadges();
}
function showNotifications(){
  const unseen=getUnseenTasks();
  if(!unseen.length){
    showToast('No new notifications','info');
    return;
  }
  const list=unseen.map(t=>`<div style="padding:10px 14px;border-bottom:1px solid var(--border-subtle);cursor:pointer" onclick="markNotifSeen('${t.id}')">
    <div style="font-weight:600;font-size:13.5px">${escHtml(t.title)}</div>
    <div style="font-size:12px;color:var(--text3);margin-top:2px">Assigned by ${USERS[t.assignedBy]?.name||t.assignedBy} · ${fmt(t.createdDate)}</div>
  </div>`).join('');
  const body=`<div style="max-height:400px;overflow-y:auto">${list}</div>
    <div style="padding:12px;text-align:right"><button class="btn btn-secondary btn-sm" onclick="markAllNotifSeen()">Mark all read</button></div>`;
  openModal('🔔 New Tasks ('+unseen.length+')', body, [{label:'Close',action:'closeModal()',cls:'btn-secondary'}]);
}
function markNotifSeen(tid){
  const seen=getSeenTaskIds();
  if(!seen.includes(tid)) seen.push(tid);
  setSeenTaskIds(seen);
  updateNotifBadge();
  closeModal();
  viewTask(tid);
}
function markAllNotifSeen(){
  const unseen=getUnseenTasks();
  const seen=getSeenTaskIds();
  unseen.forEach(t=>{ if(!seen.includes(t.id)) seen.push(t.id); });
  setSeenTaskIds(seen);
  updateNotifBadge();
  closeModal();
  showToast('All notifications marked as read','success');
}

async function markComplete(tid){
  if(currentUser==='trupal'){showToast('View only — you cannot complete tasks','error');return;}
  const t=DB.tasks.find(x=>x.id===tid);
  if(!t) return;
  t.status='Completed';
  t.completionDate=new Date().toISOString();
  t.updatedDate=new Date().toISOString();
  await dbUpdateTask(t);
  await snapshotDailySummary();
  if(t.isRecurring && t.recurrenceType){
    await spawnNextMainTask(t);
  }
  showToast('Task completed!','success');
  renderTasks(); renderMyTasks(); renderDashboard();
}

async function spawnNextMainTask(t){
  const now=new Date();
  let appearAt, newDeadline;

  if(t.recurrenceType==='daily'){
    // Reappear same day at 8PM for next working day
    appearAt=new Date(now);
    appearAt.setHours(20,0,0,0);
    if(now>=appearAt) appearAt=new Date(now.getTime()+1000);
    const nextDay=new Date(now);
    nextDay.setDate(nextDay.getDate()+1);
    while(nextDay.getDay()===0||nextDay.getDay()===6) nextDay.setDate(nextDay.getDate()+1);
    newDeadline=nextDay.toISOString().split('T')[0];
  } else if(t.recurrenceType==='weekly'){
    // Reappear next Monday 7AM
    const day=now.getDay();
    const daysToMon=day===0?1:(8-day);
    appearAt=new Date(now);
    appearAt.setDate(now.getDate()+daysToMon);
    appearAt.setHours(7,0,0,0);
    // Deadline = next Friday
    const fri=new Date(appearAt);
    fri.setDate(fri.getDate()+4);
    newDeadline=fri.toISOString().split('T')[0];
  } else if(t.recurrenceType==='monthly'){
    // Reappear on 28th of current month for next month
    const y=now.getMonth()===11&&now.getDate()>=28?now.getFullYear()+1:now.getFullYear();
    const m=now.getDate()>=28?(now.getMonth()===11?0:now.getMonth()+1):now.getMonth();
    appearAt=new Date(y, now.getMonth(), 28, 0, 0, 0);
    if(now>=appearAt) appearAt=new Date(now.getTime()+1000);
    const nextM=m===11?0:m+1;
    const nextY=m===11?y+1:y;
    const lastDay=new Date(nextY, nextM+1, 0).getDate();
    newDeadline=`${nextY}-${String(nextM+1).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;
  } else if(t.recurrenceType==='quarterly'){
    appearAt=new Date(now.getTime()+1000);
    const d=new Date(t.deadlineDate+'T00:00:00');
    d.setMonth(d.getMonth()+3);
    newDeadline=d.toISOString().split('T')[0];
  } else if(t.recurrenceType==='annually'){
    appearAt=new Date(now.getTime()+1000);
    const d=new Date(t.deadlineDate+'T00:00:00');
    d.setFullYear(d.getFullYear()+1);
    newDeadline=d.toISOString().split('T')[0];
  }

  const delay=Math.max(0, appearAt.getTime()-now.getTime());
  const taskData={title:t.title, description:t.description||'', assignedTo:t.assignedTo,
    assignedBy:t.assignedBy, deadlineDate:newDeadline, priority:t.priority,
    externalLink:t.externalLink||'', recurrenceType:t.recurrenceType,
    steps:t.steps||[], urgentFlag:t.urgentFlag||false};

  if(delay<=1000){
    await createRecurringOccurrence(taskData);
  } else {
    setTimeout(async()=>{
      await createRecurringOccurrence(taskData);
      renderTasks(); renderMyTasks(); renderDashboard();
    }, delay);
  }
}

async function createRecurringOccurrence(d){
  const exists=DB.tasks.find(x=>x.isRecurring && x.title===d.title && x.assignedTo===d.assignedTo && x.recurrenceType===d.recurrenceType && x.status!=='Completed');
  if(exists) return;
  const newTask={
    id:'task_'+Date.now()+'_'+Math.random().toString(36).substr(2,5),
    title:d.title, description:d.description, assignedTo:d.assignedTo,
    assignedBy:d.assignedBy, assignedDate:today(), deadlineDate:d.deadlineDate,
    completionDate:null, priority:d.priority, status:'Pending',
    externalLink:d.externalLink, isRecurring:true, recurrenceType:d.recurrenceType,
    createdDate:new Date().toISOString(), updatedDate:new Date().toISOString(),
    steps:d.steps, attachments:[], urgentFlag:d.urgentFlag, isDocRequest:false
  };
  DB.tasks.push(newTask);
  await dbInsertTask(newTask);
}

async function checkMainTaskRecurrence(){
  if(!canWrite()) return; // recurrence spawning is a write; skip for view-only users
  const now=new Date();

  const completedRecurring=DB.tasks.filter(t=>t.status==='Completed' && t.isRecurring && t.recurrenceType && t.completionDate);

  for(const t of completedRecurring){
    const compDate=new Date(t.completionDate);
    // Check if a pending occurrence already exists
    const hasNext=DB.tasks.find(x=>x.isRecurring && x.recurrenceType===t.recurrenceType && x.title===t.title && x.assignedTo===t.assignedTo && x.status!=='Completed' && x.id!==t.id);
    if(hasNext) continue;

    let shouldSpawn=false;
    if(t.recurrenceType==='daily'){
      const reappear=new Date(compDate); reappear.setHours(20,0,0,0);
      if(now>=reappear || now.toISOString().split('T')[0]>compDate.toISOString().split('T')[0]) shouldSpawn=true;
    } else if(t.recurrenceType==='weekly'){
      const day=now.getDay();
      const mon=new Date(now); mon.setDate(now.getDate()-(day===0?6:day-1)); mon.setHours(7,0,0,0);
      if(now>=mon && compDate<mon) shouldSpawn=true;
    } else if(t.recurrenceType==='monthly'){
      const appear28=new Date(now.getFullYear(), now.getMonth(), 28, 0, 0, 0);
      if(now>=appear28 && compDate<appear28) shouldSpawn=true;
      if(now.getMonth()!==compDate.getMonth()||now.getFullYear()!==compDate.getFullYear()) shouldSpawn=true;
    } else if(t.recurrenceType==='quarterly'){
      const compQ=Math.floor(compDate.getMonth()/3);
      const nowQ=Math.floor(now.getMonth()/3);
      if(nowQ!==compQ || now.getFullYear()!==compDate.getFullYear()) shouldSpawn=true;
    }
    if(shouldSpawn) await spawnNextMainTask(t);
  }

}

async function archiveTask(tid){
  if(!confirm('Archive this task?')) return;
  const idx=DB.tasks.findIndex(x=>x.id===tid);
  if(idx===-1) return;
  const t=DB.tasks.splice(idx,1)[0];
  const archived={...t, deletedBy:currentUser, deletedDate:new Date().toISOString()};
  DB.recycleBin.push(archived);
  await dbDeleteTask(tid);
  await dbInsertRecycleBin(archived);
  showToast('Task archived','warning');
  renderTasks(); renderDashboard();
}

// ═══════════════════════════════════════════════════════
// VIEW TASK MODAL
// ═══════════════════════════════════════════════════════
function viewTask(tid){
  const t=DB.tasks.find(x=>x.id===tid);
  if(!t) return;
  const comments=DB.comments.filter(c=>c.taskId===tid);
  const steps=t.steps||[];
  const stepProgress=DB.stepProgress[tid]||{};
  const stepsHtml=steps.length?`
    <div style="margin-top:20px">
      <strong style="font-size:13px;color:var(--text2);text-transform:uppercase;letter-spacing:.06em;font-size:12px">Checklist</strong>
      <div class="steps-list" style="margin-top:10px">
        ${steps.map(s=>`
          <div class="step-item">
            <input type="checkbox" ${stepProgress[s.id]?'checked':''} onchange="toggleStep('${tid}','${s.id}',this.checked)">
            <span class="step-text ${stepProgress[s.id]?'done':''}" id="step-lbl-${s.id}">${escHtml(s.text)}</span>
          </div>`).join('')}
      </div>
    </div>`:'';

  const attachments=t.attachments||[];
  const isTrupal=currentUser==='trupal';
  const attachHtml=attachments.length?`<div class="attached-files">${attachments.map((f,i)=>{
    const canRemove = !isTrupal && (currentUser==='sumudu' || (typeof f==='object' && f.uploadedBy===currentUser));
    const removeBtn = canRemove ? `<button class="file-remove" onclick="event.preventDefault();event.stopPropagation();removeAttachment('${tid}',${i})" title="Remove">✕</button>` : '';
    if(typeof f==='object'&&f.url){
      if(isTrupal){
        return `<div class="attached-file"><span class="file-icon">\u{1F4C4}</span><span class="file-name">${escHtml(f.name)}</span><span style="color:var(--text3);font-size:11px">🔒 View only</span></div>`;
      }
      return `<div class="attached-file"><a href="${sanitizeUrl(f.url)}" target="_blank" rel="noopener" style="display:flex;align-items:center;gap:10px;flex:1;text-decoration:none;color:inherit"><span class="file-icon">\u{1F4C4}</span><span class="file-name">${escHtml(f.name)}</span><span style="color:var(--primary);font-size:12px;font-weight:600">⬇ Open</span></a>${removeBtn}</div>`;
    }
    return `<div class="attached-file"><span class="file-icon">\u{1F4C4}</span><span class="file-name">${escHtml(typeof f==='string'?f:f.name)}</span>${removeBtn}</div>`;
  }).join('')}</div>`:'<p style="color:var(--text3);font-size:13px">No attachments</p>';

  const body=`
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:20px">
      ${statusBadge(t.status)}
      <span class="badge badge-${t.priority.toLowerCase()}">${t.priority}</span>
      ${t.isRecurring?`<span class="badge" style="background:var(--primary-light);color:var(--primary)">\u{1F504} ${t.recurrenceType}</span>`:''}
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;font-size:13.5px;margin-bottom:20px">
      <div><div style="color:var(--text3);font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Assigned To</div><div style="font-weight:600">${USERS[t.assignedTo]?.name}</div></div>
      <div><div style="color:var(--text3);font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Assigned By</div><div style="font-weight:600">${USERS[t.assignedBy]?.name}</div></div>
      <div><div style="color:var(--text3);font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Assigned Date & Time</div><div style="font-weight:500">${t.createdDate?new Date(t.createdDate).toLocaleString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'}):fmt(t.assignedDate)}</div></div>
      <div><div style="color:var(--text3);font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Deadline</div><div style="font-weight:500;color:${t.deadlineDate<today()&&t.status!=='Completed'?'var(--danger)':'inherit'}">${fmt(t.deadlineDate)}</div></div>
      ${t.completionDate?`<div><div style="color:var(--text3);font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px">Completed Date & Time</div><div style="font-weight:500">${new Date(t.completionDate).toLocaleString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'})}</div></div>`:''}
    </div>
    ${t.description?`<div style="margin-bottom:16px;font-size:14px;color:var(--text2);background:var(--surface2);padding:16px 20px;border-radius:var(--radius);border:1px solid var(--border-subtle);line-height:1.7;white-space:pre-wrap;max-height:300px;overflow-y:auto">${escHtml(t.description)}</div>`:''}
    <div style="margin-bottom:16px">
      <div style="color:var(--text3);font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">External Link</div>
      <div style="display:flex;gap:8px;align-items:center">
        <input type="url" id="vt-link-${tid}" value="${escHtml(t.externalLink||'')}" placeholder="https://..." style="flex:1;padding:6px 10px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:13px;background:var(--surface);color:var(--text)">
        <button class="btn btn-primary btn-sm" style="font-size:11px;padding:5px 10px" onclick="saveTaskLink('${tid}')">Save</button>
        ${t.externalLink?`<a href="${sanitizeUrl(t.externalLink)}" target="_blank" rel="noopener" style="font-size:12px;color:var(--primary);white-space:nowrap">🔗 Open</a>`:''}
      </div>
    </div>
    ${stepsHtml}
    <div style="margin-top:20px">
      <div style="color:var(--text3);font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px">📌 Original Task Attachments</div>
      ${attachHtml}
      ${currentUser!=='trupal'?`<div class="file-upload-area" style="margin-top:10px;padding:14px">
        <input type="file" multiple id="view-files" onchange="addAttachmentToTask('${tid}',this)">
        <p style="font-size:12px;color:var(--text3)">+ Attach file</p>
      </div>`:''}
    </div>
    <div class="comments-section">
      <div style="color:var(--text3);font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:14px">Remarks & Comments</div>
      ${comments.length?comments.map(c=>`
        <div class="comment-item">
          <div class="comment-avatar" style="background:linear-gradient(135deg,${USERS[c.userId]?.color||'#6366f1'},${USERS[c.userId]?.color||'#6366f1'}cc)">${USERS[c.userId]?.avatar||'?'}</div>
          <div class="comment-bubble">
            <div class="comment-meta"><strong>${USERS[c.userId]?.name||c.userId}</strong> · ${new Date(c.createdDate).toLocaleString()}</div>
            <div class="comment-text">${escHtml(c.text)}</div>
            ${(c.attachments&&c.attachments.length)?`<div style="margin-top:8px;border-top:1px solid var(--border-subtle);padding-top:6px"><div style="font-size:11px;color:var(--text3);font-weight:600;margin-bottom:4px">📎 Reply Attachments</div>${c.attachments.map(a=>'<div style="margin-top:3px"><a href="'+sanitizeUrl(a.url)+'" target="_blank" style="font-size:12px;color:var(--primary)">📄 '+escHtml(a.name)+'</a></div>').join('')}</div>`:''}
          </div>
        </div>`).join(''):'<p style="color:var(--text3);font-size:13px;margin-bottom:14px">No comments yet.</p>'}
      ${currentUser!=='trupal'?`<div class="comment-input-row"><input id="comment-input" placeholder="Write a remark..." onkeydown="if(event.key==='Enter')addComment('${tid}')"><div style="display:flex;gap:4px;align-items:center"><label class="btn btn-secondary btn-sm" style="cursor:pointer;padding:5px 8px" title="Attach file to reply">📎<input type="file" multiple id="comment-files" style="display:none"></label><button class="btn btn-primary btn-sm" onclick="addComment('${tid}')">Send</button></div></div><div id="comment-files-preview" style="margin-top:4px;font-size:12px;color:var(--text3)"></div>`:'<p style="color:var(--text3);font-size:12px;font-style:italic">View only — comments disabled</p>'}
    </div>
  `;
  const footer=[{label:'Close',action:'closeModal()',cls:'btn-secondary'}];
  if(currentUser!=='trupal' && t.status!=='Completed') footer.push({label:'Mark Complete',action:`markComplete('${tid}');closeModal()`,cls:'btn-success'});
  if(currentUser==='sumudu') footer.push({label:'Archive',action:`archiveTask('${tid}');closeModal()`,cls:'btn-danger'});
  openModal(escHtml(t.title), body, footer);
}

async function toggleStep(tid,sid,checked){
  if(!DB.stepProgress[tid]) DB.stepProgress[tid]={};
  DB.stepProgress[tid][sid]=checked;
  await dbUpsertStepProgress(tid,sid,checked);
  const lbl=document.getElementById('step-lbl-'+sid);
  if(lbl){ lbl.classList.toggle('done',checked); }
}

async function saveTaskLink(tid){
  const t=DB.tasks.find(x=>x.id===tid);
  if(!t) return;
  const linkEl=document.getElementById('vt-link-'+tid);
  if(!linkEl) return;
  t.externalLink=linkEl.value.trim();
  await dbUpdateTask(t);
  showToast('🔗 Link saved!','success');
  viewTask(tid);
}

async function addAttachmentToTask(tid,input){
  const t=DB.tasks.find(x=>x.id===tid);
  if(!t) return;
  if(!t.attachments) t.attachments=[];
  const files=Array.from(input.files);
  const uploaded=await uploadFilesToStorage(tid, files);
  t.attachments.push(...uploaded);
  await dbUpdateTask(t);
  showToast('File attached','success');
  viewTask(tid);
}

async function removeAttachment(tid, index){
  if(!confirm('Remove this attachment?')) return;
  const t=DB.tasks.find(x=>x.id===tid);
  if(!t||!t.attachments) return;
  const f=t.attachments[index];
  if(typeof f==='object'&&f.path){
    await sb.storage.from('attachments').remove([f.path]);
  }
  t.attachments.splice(index,1);
  await dbUpdateTask(t);
  showToast('Attachment removed','warning');
  viewTask(tid);
}

function openRemarks(tid){ viewTask(tid); }

function editTask(tid){
  const t=DB.tasks.find(x=>x.id===tid);
  if(!t) return;
  const body=`
  <div class="form-grid">
    <div class="form-field full"><label>Task Title</label><input id="et-title" value="${escHtml(t.title)}"></div>
    <div class="form-field full"><label>Description</label><textarea id="et-desc">${escHtml(t.description||'')}</textarea></div>
    <div class="form-field"><label>Assign To</label>
      <select id="et-assign"><option value="de" ${t.assignedTo==='de'?'selected':''}>Dev</option><option value="mitiksha" ${t.assignedTo==='mitiksha'?'selected':''}>Mitiksha</option></select>
    </div>
    <div class="form-field"><label>Status</label>
      <select id="et-status"><option value="Pending" ${t.status==='Pending'?'selected':''}>Pending</option><option value="In Progress" ${t.status==='In Progress'?'selected':''}>In Progress</option><option value="Completed" ${t.status==='Completed'?'selected':''}>Completed</option><option value="Overdue" ${t.status==='Overdue'?'selected':''}>Overdue</option></select>
    </div>
    <div class="form-field"><label>Assigned Date</label><input type="date" id="et-adate" value="${t.assignedDate||''}"></div>
    <div class="form-field"><label>Deadline Date</label><input type="date" id="et-ddate" value="${t.deadlineDate||''}"></div>
    <div class="form-field full"><label>External Link</label><input id="et-link" value="${escHtml(t.externalLink||'')}"></div>
    <div class="form-field"><label>Recurring</label>
      <select id="et-recur"><option value="" ${!t.isRecurring?'selected':''}>No Repeat</option><option value="daily" ${t.recurrenceType==='daily'?'selected':''}>Daily</option><option value="weekly" ${t.recurrenceType==='weekly'?'selected':''}>Weekly</option><option value="monthly" ${t.recurrenceType==='monthly'?'selected':''}>Monthly</option></select>
    </div>
    <div class="form-field"><label>Urgent?</label>
      <label class="urgent-check ${t.urgentFlag?'checked':''}">
        <input type="checkbox" id="et-urgent-flag" ${t.urgentFlag?'checked':''} onchange="this.parentElement.classList.toggle('checked',this.checked)">
        🔴 Mark as Urgent
      </label>
    </div>
    <div class="form-field full"><label>Attachments</label>
      <div class="attached-files" id="et-attachments">${(t.attachments||[]).map((f,i)=>{
        const name=typeof f==='object'?f.name:f;
        const link=typeof f==='object'&&f.url?`<a href="${sanitizeUrl(f.url)}" target="_blank" rel="noopener" style="color:var(--primary);font-size:12px;font-weight:600;text-decoration:none">⬇ Open</a>`:'';
        return `<div class="attached-file"><span class="file-icon">📄</span><span class="file-name">${escHtml(name)}</span>${link}<button class="file-remove" onclick="removeAttachmentFromEdit('${tid}',${i})">✕</button></div>`;
      }).join('')}</div>
      <div class="file-upload-area" style="margin-top:10px;padding:14px">
        <input type="file" multiple id="et-files" onchange="previewFiles(this,'et-new-files')">
        <p style="font-size:12px;color:var(--text3)">+ Attach new files</p>
      </div>
      <div class="attached-files" id="et-new-files"></div>
    </div>
  </div>`;
  openModal('Edit Task', body, [
    {label:'Cancel', action:'closeModal()', cls:'btn-secondary'},
    {label:'Save Changes', action:`saveEditTask('${tid}')`, cls:'btn-primary'}
  ]);
}

async function removeAttachmentFromEdit(tid, index){
  const t=DB.tasks.find(x=>x.id===tid);
  if(!t||!t.attachments) return;
  const f=t.attachments[index];
  if(typeof f==='object'&&f.path){
    await sb.storage.from('attachments').remove([f.path]);
  }
  t.attachments.splice(index,1);
  await dbUpdateTask(t);
  showToast('Attachment removed','warning');
  editTask(tid);
}

async function saveEditTask(tid){
  const t=DB.tasks.find(x=>x.id===tid);
  if(!t) return;
  const title=document.getElementById('et-title').value.trim();
  if(!title){ showToast('Task title is required','error'); return; }
  t.title=title;
  t.description=document.getElementById('et-desc').value.trim();
  t.assignedTo=document.getElementById('et-assign').value;
  t.status=document.getElementById('et-status').value;
  t.assignedDate=document.getElementById('et-adate').value;
  t.deadlineDate=document.getElementById('et-ddate').value;
  t.externalLink=document.getElementById('et-link').value.trim();
  const recur=document.getElementById('et-recur').value;
  t.isRecurring=recur!=='';
  t.recurrenceType=recur;
  t.urgentFlag=document.getElementById('et-urgent-flag').checked;
  t.priority=t.urgentFlag?'Urgent':'Normal';
  // If deadline changed to future and task was overdue, move to Pending
  if(t.status==='Overdue' && t.deadlineDate && t.deadlineDate >= today()){
    t.status='Pending';
  }
  if(pendingFiles.length){
    const uploaded=await uploadFilesToStorage(tid, pendingFiles);
    if(!t.attachments) t.attachments=[];
    t.attachments.push(...uploaded);
  }
  if(t.status==='Completed'&&!t.completionDate) t.completionDate=today();
  t.updatedDate=new Date().toISOString();
  await dbUpdateTask(t);
  closeModal();
  showToast('Task updated successfully','success');
  renderTasks(); renderMyTasks(); renderDashboard();
}

async function addComment(tid){
  const inp=document.getElementById('comment-input');
  const text=inp.value.trim();
  const fileInput=document.getElementById('comment-files');
  const files=fileInput?Array.from(fileInput.files):[];
  if(!text && !files.length) return;
  let replyAttachments=[];
  if(files.length){
    for(const file of files){
      if(file.size>10*1024*1024){showToast(file.name+' too large','error');continue;}
      const filePath=`comments/${tid}/${Date.now()}_${file.name}`;
      const {error}=await sb.storage.from('attachments').upload(filePath, file);
      if(error){showToast('Upload failed','error');continue;}
      const {data:urlData}=sb.storage.from('attachments').getPublicUrl(filePath);
      replyAttachments.push({name:file.name, url:urlData.publicUrl, uploadedBy:currentUser, uploadedAt:new Date().toISOString()});
    }
  }
  const c={ id:'c_'+Date.now(), taskId:tid, userId:currentUser, text, createdDate:new Date().toISOString(), attachments:replyAttachments };
  DB.comments.push(c);
  await dbInsertComment(c);
  inp.value='';
  if(fileInput) fileInput.value='';
  const t=DB.tasks.find(x=>x.id===tid);
  if(t && t.status==='Completed' && t.assignedTo!==currentUser){
    const seen=getSeenTaskIds();
    const idx=seen.indexOf(tid);
    if(idx!==-1) seen.splice(idx,1);
    localStorage.setItem('seenTasks_'+t.assignedTo, JSON.stringify(
      JSON.parse(localStorage.getItem('seenTasks_'+t.assignedTo)||'[]').filter(id=>id!==tid)
    ));
  }
  viewTask(tid);
}

// ═══════════════════════════════════════════════════════
// PERFORMANCE
// ═══════════════════════════════════════════════════════
function calcScore(uid, month, year){
  const tasks=DB.tasks.filter(t=>t.assignedTo===uid && t.assignedDate && new Date(t.assignedDate).getMonth()===month && new Date(t.assignedDate).getFullYear()===year);
  if(!tasks.length) return 100;
  let score=100;
  tasks.forEach(t=>{
    if(t.status==='Completed'){
      if(t.completionDate && t.deadlineDate){
        const workDays=countWorkingDays(t.assignedDate, t.completionDate);
        const allowedDays=countWorkingDays(t.assignedDate, t.deadlineDate);
        if(workDays<=allowedDays) score+=2;
        else score-=5;
      }
    } else if(t.status==='Overdue'){ score-=10; }
    else { score-=2; }
  });
  return Math.max(0,Math.min(100,Math.round(score)));
}

function initPerfFilters(){
  const now=new Date();
  document.getElementById('perf-month').value=now.getMonth();
  const ySel=document.getElementById('perf-year');
  if(!ySel.options.length){
    for(let y=now.getFullYear()-2;y<=now.getFullYear();y++){ const o=document.createElement('option'); o.value=y; o.textContent=y; if(y===now.getFullYear())o.selected=true; ySel.appendChild(o); }
  }
}

function renderPerformance(){
  const month=parseInt(document.getElementById('perf-month').value);
  const year=parseInt(document.getElementById('perf-year').value)||new Date().getFullYear();
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const employees = (currentUser==='sumudu'||currentUser==='trupal')?['de','mitiksha']:[currentUser];
  const isDark=document.documentElement.getAttribute('data-theme')==='dark';
  let html='<div class="perf-grid">';
  employees.forEach(uid=>{
    const u=USERS[uid];
    const tasks=DB.tasks.filter(t=>t.assignedTo===uid && t.assignedDate && new Date(t.assignedDate).getMonth()===month && new Date(t.assignedDate).getFullYear()===year);
    const done=tasks.filter(t=>t.status==='Completed').length;
    const late=tasks.filter(t=>t.status==='Completed'&&t.completionDate&&t.completionDate>t.deadlineDate).length;
    const overdue=tasks.filter(t=>t.status==='Overdue').length;
    const ontime=done-late;
    const score=calcScore(uid,month,year);
    const color=score>=80?'#22c55e':score>=60?'#f59e0b':'#ef4444';
    const circ=2*Math.PI*40;
    const dash=(score/100)*circ;
    html+=`<div class="perf-card">
      <div class="emp-name">${u.name}</div>
      <div class="emp-role">${months[month]} ${year}</div>
      <div class="score-ring" style="width:100px;height:100px">
        <svg width="100" height="100" viewBox="0 0 100 100">
          <circle cx="50" cy="50" r="40" fill="none" stroke="${isDark?'#2a2b3d':'#e4e7ef'}" stroke-width="8"/>
          <circle cx="50" cy="50" r="40" fill="none" stroke="${color}" stroke-width="8" stroke-dasharray="${dash} ${circ}" stroke-linecap="round"/>
        </svg>
        <div class="score-num" style="color:${color}">${score}</div>
      </div>
      <div style="font-size:12.5px;text-align:center;margin-bottom:14px;color:var(--text3);font-weight:600">out of 100</div>
      ${bar('Assigned',tasks.length,tasks.length||1,'#6366f1')}
      ${bar('Completed',done,tasks.length||1,'#22c55e')}
      ${bar('On Time',ontime,tasks.length||1,'#3b82f6')}
      ${bar('Late',late,tasks.length||1,'#f59e0b')}
      ${bar('Overdue',overdue,tasks.length||1,'#ef4444')}
    </div>`;
  });
  html+='</div>';
  document.getElementById('perf-content').innerHTML=html;
}

function bar(label,val,total,color){
  const pct=total?Math.round((val/total)*100):0;
  return `<div class="score-bar-row"><span class="bar-label">${label}</span><div class="score-bar-bg"><div class="score-bar-fill" style="width:${pct}%;background:${color}"></div></div><span style="font-size:12px;width:24px;text-align:right;font-weight:600">${val}</span></div>`;
}

function exportPerfExcel(){
  const month=parseInt(document.getElementById('perf-month').value);
  const year=parseInt(document.getElementById('perf-year').value)||new Date().getFullYear();
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const rows=[['Employee','Month','Year','Total','Completed','Late','Overdue','Score']];
  ['de','mitiksha'].forEach(uid=>{
    const tasks=DB.tasks.filter(t=>t.assignedTo===uid && t.assignedDate && new Date(t.assignedDate).getMonth()===month && new Date(t.assignedDate).getFullYear()===year);
    const done=tasks.filter(t=>t.status==='Completed').length;
    const late=tasks.filter(t=>t.status==='Completed'&&t.completionDate&&t.completionDate>t.deadlineDate).length;
    const overdue=tasks.filter(t=>t.status==='Overdue').length;
    rows.push([USERS[uid].name,months[month],year,tasks.length,done,late,overdue,calcScore(uid,month,year)]);
  });
  const ws=XLSX.utils.aoa_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Performance');
  XLSX.writeFile(wb,`Performance_${months[month]}_${year}.xlsx`);
}

// ═══════════════════════════════════════════════════════
// SCHEDULE PAGE
// ═══════════════════════════════════════════════════════
function initScheduleFilter(){
  document.getElementById('schedule-month-filter').value=new Date().getMonth();
}

function renderSchedule(){
  const month=parseInt(document.getElementById('schedule-month-filter').value);
  const months=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const year=new Date().getFullYear();
  const employees = (currentUser==='sumudu'||currentUser==='trupal')?['de','mitiksha']:[currentUser];
  let html='';
  employees.forEach(uid=>{
    const u=USERS[uid];
    const tasks=DB.tasks.filter(t=>t.assignedTo===uid && t.deadlineDate && new Date(t.deadlineDate).getMonth()===month && new Date(t.deadlineDate).getFullYear()===year);
    if(!tasks.length && currentUser!=='sumudu') { html+=`<div class="empty-state"><div class="empty-icon">\u{1F4C5}</div><p>No tasks scheduled for ${months[month]}</p></div>`; return; }
    html+=`<div class="schedule-month"><h3>\u{1F464} ${u.name} — ${months[month]} ${year}</h3>`;
    if(!tasks.length){ html+=`<p style="color:var(--text3);font-size:13.5px;font-weight:500">No tasks this month</p>`; }
    tasks.forEach(t=>{
      const steps=t.steps||[];
      const sp=DB.stepProgress[t.id]||{};
      const done=steps.filter(s=>sp[s.id]).length;
      html+=`<div class="schedule-task-item">
        <div class="schedule-task-header">
          <span class="priority-dot dot-${t.priority.toLowerCase()}"></span>
          <span class="schedule-task-title ${t.status==='Completed'?'completed':''}">${escHtml(t.title)}</span>
          ${statusBadge(t.status)}
          <span style="font-size:12.5px;color:var(--text3);margin-left:auto;font-weight:500">Due: ${fmt(t.deadlineDate)}</span>
        </div>
        ${steps.length?`
          <div class="schedule-steps">
            <div style="font-size:12px;color:var(--text3);margin-bottom:6px;font-weight:600">${done}/${steps.length} steps completed</div>
            ${steps.map(s=>`<div class="schedule-step">
              <input type="checkbox" ${sp[s.id]?'checked':''} onchange="toggleStep('${t.id}','${s.id}',this.checked)">
              <span class="step-label ${sp[s.id]?'done':''}">${escHtml(s.text)}</span>
            </div>`).join('')}
          </div>`:''}
      </div>`;
    });
    html+='</div>';
  });
  document.getElementById('schedule-content').innerHTML=html||`<div class="empty-state"><div class="empty-icon">\u{1F4C5}</div><p>No tasks found for this month</p></div>`;
}

// ═══════════════════════════════════════════════════════
// PERSONAL TASKS
// ═══════════════════════════════════════════════════════
let personalTasksLoaded=false;
function loadPersonalTasks(){
  if(currentUser==='sumudu'){
    document.getElementById('personal-tasks-sumudu').style.display='';
    document.getElementById('personal-tasks-officer').style.display='none';
    if(!personalTasksLoaded){
      const frame=document.getElementById("personal-tasks-frame");
      if(!frame) return;
  const html=atob("PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04IiAvPgo8bWV0YSBuYW1lPSJ2aWV3cG9ydCIgY29udGVudD0id2lkdGg9ZGV2aWNlLXdpZHRoLCBpbml0aWFsLXNjYWxlPTEuMCwgbWF4aW11bS1zY2FsZT0xIiAvPgo8dGl0bGU+RHV0eSBMaXN0IE1hbmFnZXI8L3RpdGxlPgo8bGluayByZWw9Imljb24iIGhyZWY9ImRhdGE6aW1hZ2Uvc3ZnK3htbCw8c3ZnIHhtbG5zPSUyMmh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnJTIyIHZpZXdCb3g9JTIyMCAwIDI0IDI0JTIyPjx0ZXh0IHk9JTIyLjllbSUyMiBmb250LXNpemU9JTIyMjIlMjI+JUUyJTlDJTg1PC90ZXh0Pjwvc3ZnPiI+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi50YWlsd2luZGNzcy5jb20iPjwvc2NyaXB0Pgo8c2NyaXB0IHNyYz0iaHR0cHM6Ly9jZG4uanNkZWxpdnIubmV0L25wbS9Ac3VwYWJhc2Uvc3VwYWJhc2UtanNAMiI+PC9zY3JpcHQ+CjxzY3JpcHQ+CiAgdGFpbHdpbmQuY29uZmlnID0gewogICAgZGFya01vZGU6ICdjbGFzcycsCiAgICB0aGVtZTogewogICAgICBleHRlbmQ6IHsKICAgICAgICBjb2xvcnM6IHsKICAgICAgICAgIHByaW1hcnk6IHsKICAgICAgICAgICAgNTA6ICcjZWVmMmZmJywgMTAwOiAnI2UwZTdmZicsIDIwMDogJyNjN2QyZmUnLCAzMDA6ICcjYTViNGZjJywKICAgICAgICAgICAgNDAwOiAnIzgxOGNmOCcsIDUwMDogJyM2MzY2ZjEnLCA2MDA6ICcjNGY0NmU1JywgNzAwOiAnIzQzMzhjYScsCiAgICAgICAgICAgIDgwMDogJyMzNzMwYTMnLCA5MDA6ICcjMzEyZTgxJwogICAgICAgICAgfQogICAgICAgIH0KICAgICAgfQogICAgfQogIH0KPC9zY3JpcHQ+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi5qc2RlbGl2ci5uZXQvbnBtL3hsc3hAMC4xOC41L2Rpc3QveGxzeC5mdWxsLm1pbi5qcyI+PC9zY3JpcHQ+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi5qc2RlbGl2ci5uZXQvbnBtL21hbW1vdGhAMS44LjAvbWFtbW90aC5icm93c2VyLm1pbi5qcyI+PC9zY3JpcHQ+CjxzY3JpcHQgc3JjPSJodHRwczovL2Nkbi5qc2RlbGl2ci5uZXQvbnBtL2pzcGRmQDIuNS4xL2Rpc3QvanNwZGYudW1kLm1pbi5qcyI+PC9zY3JpcHQ+CjxzdHlsZT4KICBodG1sIHsgLXdlYmtpdC10YXAtaGlnaGxpZ2h0LWNvbG9yOiB0cmFuc3BhcmVudDsgfQogIGJvZHkgewogICAgZm9udC1mYW1pbHk6ICdJbnRlcicsIHN5c3RlbS11aSwgLWFwcGxlLXN5c3RlbSwgc2Fucy1zZXJpZjsKICAgIC13ZWJraXQtZm9udC1zbW9vdGhpbmc6IGFudGlhbGlhc2VkOwogIH0KICA6Oi13ZWJraXQtc2Nyb2xsYmFyIHsgd2lkdGg6IDhweDsgaGVpZ2h0OiA4cHg7IH0KICA6Oi13ZWJraXQtc2Nyb2xsYmFyLXRodW1iIHsgYmFja2dyb3VuZDogI2QxZDVkYjsgYm9yZGVyLXJhZGl1czogOTk5OXB4OyB9CiAgLmRhcmsgOjotd2Via2l0LXNjcm9sbGJhci10aHVtYiB7IGJhY2tncm91bmQ6ICMzNzQxNTE7IH0KICA6Oi13ZWJraXQtc2Nyb2xsYmFyLXRyYWNrIHsgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IH0KCiAgLmxpbmUtY2xhbXAtMiB7CiAgICBkaXNwbGF5OiAtd2Via2l0LWJveDsKICAgIC13ZWJraXQtbGluZS1jbGFtcDogMjsKICAgIC13ZWJraXQtYm94LW9yaWVudDogdmVydGljYWw7CiAgICBvdmVyZmxvdzogaGlkZGVuOwogIH0KCiAgQGtleWZyYW1lcyBmYWRlSW4geyBmcm9tIHsgb3BhY2l0eTogMCB9IHRvIHsgb3BhY2l0eTogMSB9IH0KICBAa2V5ZnJhbWVzIHNsaWRlVXAgeyBmcm9tIHsgb3BhY2l0eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDhweCkgfSB0byB7IG9wYWNpdHk6IDE7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSgwKSB9IH0KICBAa2V5ZnJhbWVzIHNwaW4geyB0byB7IHRyYW5zZm9ybTogcm90YXRlKDM2MGRlZykgfSB9CiAgLmFuaW0tZmFkZSB7IGFuaW1hdGlvbjogZmFkZUluIC4ycyBlYXNlLW91dDsgfQogIC5hbmltLXNsaWRlIHsgYW5pbWF0aW9uOiBzbGlkZVVwIC4ycyBlYXNlLW91dDsgfQogIC5hbmltLXNwaW4geyBhbmltYXRpb246IHNwaW4gMXMgbGluZWFyIGluZmluaXRlOyBkaXNwbGF5OiBpbmxpbmUtZmxleDsgfQoKICAuaWNvbiB7IHdpZHRoOiAxLjI1cmVtOyBoZWlnaHQ6IDEuMjVyZW07IGZsZXgtc2hyaW5rOiAwOyB9CiAgLmljb24tc20geyB3aWR0aDogMXJlbTsgaGVpZ2h0OiAxcmVtOyBmbGV4LXNocmluazogMDsgfQoKICBrYmQgewogICAgZm9udC1mYW1pbHk6IHVpLW1vbm9zcGFjZSwgbW9ub3NwYWNlOwogICAgZm9udC1zaXplOiAwLjc1cmVtOwogIH0KCiAgaW5wdXRbdHlwZT0iY2hlY2tib3giXS5kbG0tY2hlY2tib3ggewogICAgd2lkdGg6IDEuNXJlbTsgaGVpZ2h0OiAxLjVyZW07IGZsZXgtc2hyaW5rOiAwOyBjdXJzb3I6IHBvaW50ZXI7CiAgICBhY2NlbnQtY29sb3I6ICM0ZjQ2ZTU7CiAgfQoKICB0ZXh0YXJlYSwgaW5wdXRbdHlwZT0idGV4dCJdLCBpbnB1dFt0eXBlPSJlbWFpbCJdLCBpbnB1dFt0eXBlPSJwYXNzd29yZCJdIHsKICAgIGZvbnQtZmFtaWx5OiBpbmhlcml0OwogIH0KPC9zdHlsZT4KPC9oZWFkPgo8Ym9keSBjbGFzcz0iYmctZ3JheS01MCB0ZXh0LWdyYXktOTAwIGRhcms6YmctZ3JheS05NTAgZGFyazp0ZXh0LWdyYXktMTAwIHRyYW5zaXRpb24tY29sb3JzIGR1cmF0aW9uLTIwMCI+CiAgPGRpdiBpZD0iYXBwIj48L2Rpdj4KPHNjcmlwdD4KLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgIFBlcnNvbmFsIER1dHkgTGlzdCBNYW5hZ2VyIC0gc2luZ2xlLWZpbGUgY2xpZW50LXNpZGUgYXBwCiAgIERhdGEgaXMgc3RvcmVkIGVudGlyZWx5IGluIHRoaXMgYnJvd3NlcidzIGxvY2FsU3RvcmFnZS4KICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCgovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KLy8gVXRpbGl0aWVzCi8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQpmdW5jdGlvbiB1aWQoKSB7CiAgcmV0dXJuIERhdGUubm93KCkudG9TdHJpbmcoMzYpICsgTWF0aC5yYW5kb20oKS50b1N0cmluZygzNikuc2xpY2UoMiwgOCk7Cn0KCmZ1bmN0aW9uIG5vd0lTTygpIHsKICByZXR1cm4gbmV3IERhdGUoKS50b0lTT1N0cmluZygpOwp9CgpmdW5jdGlvbiBlc2NhcGVIdG1sKHN0cikgewogIGlmIChzdHIgPT09IG51bGwgfHwgc3RyID09PSB1bmRlZmluZWQpIHJldHVybiAnJzsKICByZXR1cm4gU3RyaW5nKHN0cikKICAgIC5yZXBsYWNlKC8mL2csICcmYW1wOycpCiAgICAucmVwbGFjZSgvPC9nLCAnJmx0OycpCiAgICAucmVwbGFjZSgvPi9nLCAnJmd0OycpCiAgICAucmVwbGFjZSgvIi9nLCAnJnF1b3Q7JykKICAgIC5yZXBsYWNlKC8nL2csICcmIzM5OycpOwp9CgpmdW5jdGlvbiBmb3JtYXREYXRlKGlzbykgewogIGlmICghaXNvKSByZXR1cm4gJyc7CiAgY29uc3QgZCA9IG5ldyBEYXRlKGlzbyk7CiAgaWYgKGlzTmFOKGQuZ2V0VGltZSgpKSkgcmV0dXJuICcnOwogIHJldHVybiBkLnRvTG9jYWxlRGF0ZVN0cmluZyh1bmRlZmluZWQsIHsgbW9udGg6ICdzaG9ydCcsIGRheTogJ251bWVyaWMnLCB5ZWFyOiAnbnVtZXJpYycgfSk7Cn0KCmZ1bmN0aW9uIGZvcm1hdERhdGVUaW1lKGlzbykgewogIGlmICghaXNvKSByZXR1cm4gJ+KAlCc7CiAgY29uc3QgZCA9IG5ldyBEYXRlKGlzbyk7CiAgaWYgKGlzTmFOKGQuZ2V0VGltZSgpKSkgcmV0dXJuICfigJQnOwogIHJldHVybiBkLnRvTG9jYWxlU3RyaW5nKHVuZGVmaW5lZCwgeyBtb250aDogJ3Nob3J0JywgZGF5OiAnbnVtZXJpYycsIHllYXI6ICdudW1lcmljJywgaG91cjogJzItZGlnaXQnLCBtaW51dGU6ICcyLWRpZ2l0JyB9KTsKfQoKZnVuY3Rpb24gZGVib3VuY2UoZm4sIGRlbGF5KSB7CiAgbGV0IHRpbWVyID0gbnVsbDsKICByZXR1cm4gZnVuY3Rpb24gKC4uLmFyZ3MpIHsKICAgIGNsZWFyVGltZW91dCh0aW1lcik7CiAgICB0aW1lciA9IHNldFRpbWVvdXQoKCkgPT4gZm4uYXBwbHkodGhpcywgYXJncyksIGRlbGF5KTsKICB9Owp9CgpmdW5jdGlvbiBpc1RydXRoeVZhbHVlKHZhbHVlKSB7CiAgaWYgKHZhbHVlID09PSB1bmRlZmluZWQgfHwgdmFsdWUgPT09IG51bGwpIHJldHVybiBmYWxzZTsKICBjb25zdCBzdHIgPSBTdHJpbmcodmFsdWUpLnRyaW0oKS50b0xvd2VyQ2FzZSgpOwogIHJldHVybiBbJ3llcycsICd5JywgJ3RydWUnLCAnMScsICd1cmdlbnQnXS5pbmNsdWRlcyhzdHIpOwp9CgovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KLy8gSW5saW5lIGljb24gc2V0IChzaW1wbGUgb3JpZ2luYWwgZ2VvbWV0cmljIFNWR3MpCi8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQpjb25zdCBJQ09OUyA9IHsKICBjaGVjazogJzxzdmcgY2xhc3M9Imljb24iIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJjdXJyZW50Q29sb3IiIHN0cm9rZS13aWR0aD0iMi41IiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik01IDEzbDQgNEwxOSA3Ii8+PC9zdmc+JywKICBjbG9jazogJzxzdmcgY2xhc3M9Imljb24iIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJjdXJyZW50Q29sb3IiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48Y2lyY2xlIGN4PSIxMiIgY3k9IjEyIiByPSI5Ii8+PHBhdGggZD0iTTEyIDd2NWwzIDMiLz48L3N2Zz4nLAogIGxvYWRlcjogJzxzdmcgY2xhc3M9Imljb24gYW5pbS1zcGluIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCI+PHBhdGggZD0iTTEyIDNhOSA5IDAgMSAwIDkgOSIgLz48L3N2Zz4nLAogIGFsZXJ0OiAnPHN2ZyBjbGFzcz0iaWNvbiIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik0xMiAyLjUgMjEuNSAyMEgyLjV6Ii8+PHBhdGggZD0iTTEyIDl2NCIvPjxjaXJjbGUgY3g9IjEyIiBjeT0iMTYuMyIgcj0iMC42IiBmaWxsPSJjdXJyZW50Q29sb3IiIHN0cm9rZT0ibm9uZSIvPjwvc3ZnPicsCiAgcGVuY2lsOiAnPHN2ZyBjbGFzcz0iaWNvbiIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik00IDIwaDRMMTkuNSA4LjVhMi4xIDIuMSAwIDAgMC0zLTNMNCAxN3oiLz48cGF0aCBkPSJNMTQgNi41bDMgMyIvPjwvc3ZnPicsCiAgdHJhc2g6ICc8c3ZnIGNsYXNzPSJpY29uIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBhdGggZD0iTTQgN2gxNiIvPjxwYXRoIGQ9Ik05IDdWNC41aDZWNyIvPjxwYXRoIGQ9Ik02IDdsMSAxM2gxMGwxLTEzIi8+PHBhdGggZD0iTTEwIDExdjZNMTQgMTF2NiIvPjwvc3ZnPicsCiAgc2VhcmNoOiAnPHN2ZyBjbGFzcz0iaWNvbiIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxjaXJjbGUgY3g9IjExIiBjeT0iMTEiIHI9IjciLz48cGF0aCBkPSJNMjEgMjFsLTQuNS00LjUiLz48L3N2Zz4nLAogIHg6ICc8c3ZnIGNsYXNzPSJpY29uIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIuNSIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cGF0aCBkPSJNNSA1bDE0IDE0TTE5IDVMNSAxOSIvPjwvc3ZnPicsCiAgcGx1czogJzxzdmcgY2xhc3M9Imljb24iIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJjdXJyZW50Q29sb3IiIHN0cm9rZS13aWR0aD0iMi41IiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik0xMiA1djE0TTUgMTJoMTQiLz48L3N2Zz4nLAogIHN1bjogJzxzdmcgY2xhc3M9Imljb24iIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJjdXJyZW50Q29sb3IiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48Y2lyY2xlIGN4PSIxMiIgY3k9IjEyIiByPSI0Ii8+PHBhdGggZD0iTTEyIDJ2Mk0xMiAyMHYyTTQgMTJIMk0yMiAxMmgtMk00LjkgNC45bDEuNCAxLjRNMTcuNyAxNy43bDEuNCAxLjRNMTkuMSA0LjlsLTEuNCAxLjRNNi4zIDE3LjdsLTEuNCAxLjQiLz48L3N2Zz4nLAogIG1vb246ICc8c3ZnIGNsYXNzPSJpY29uIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBhdGggZD0iTTIxIDEyLjhBOSA5IDAgMSAxIDExLjIgMyA3IDcgMCAwIDAgMjEgMTIuOHoiLz48L3N2Zz4nLAogIHVwbG9hZDogJzxzdmcgY2xhc3M9Imljb24iIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJjdXJyZW50Q29sb3IiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cGF0aCBkPSJNMTIgMTVWNE03IDlsNS01IDUgNSIvPjxwYXRoIGQ9Ik00IDE4LjVoMTYiLz48L3N2Zz4nLAogIGRvd25sb2FkOiAnPHN2ZyBjbGFzcz0iaWNvbiIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik0xMiA0djExTTcgMTBsNSA1IDUtNSIvPjxwYXRoIGQ9Ik00IDE4LjVoMTYiLz48L3N2Zz4nLAogIHNlbmQ6ICc8c3ZnIGNsYXNzPSJpY29uIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBhdGggZD0iTTIxIDMgMTEgMTMiLz48cGF0aCBkPSJNMjEgM2wtNi41IDE4LTQtOC04LTR6Ii8+PC9zdmc+JywKICBsb2dvdXQ6ICc8c3ZnIGNsYXNzPSJpY29uIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBhdGggZD0iTTkgMjFINmEyIDIgMCAwIDEtMi0yVjVhMiAyIDAgMCAxIDItMmgzIi8+PHBhdGggZD0iTTE2IDE3bDUtNS01LTUiLz48cGF0aCBkPSJNMjEgMTJIOSIvPjwvc3ZnPicsCiAgbGlzdDogJzxzdmcgY2xhc3M9Imljb24iIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJjdXJyZW50Q29sb3IiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cGF0aCBkPSJNOSA2aDExTTkgMTJoMTFNOSAxOGgxMSIvPjxjaXJjbGUgY3g9IjQuNSIgY3k9IjYiIHI9IjEiLz48Y2lyY2xlIGN4PSI0LjUiIGN5PSIxMiIgcj0iMSIvPjxjaXJjbGUgY3g9IjQuNSIgY3k9IjE4IiByPSIxIi8+PC9zdmc+JywKICBtZXNzYWdlOiAnPHN2ZyBjbGFzcz0iaWNvbiIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik00IDRoMTZ2MTJIOWwtNSA0eiIvPjwvc3ZnPicsCiAgZXhjZWw6ICc8c3ZnIGNsYXNzPSJpY29uIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBhdGggZD0iTTQgNGgxMWw1IDV2MTFINHoiLz48cGF0aCBkPSJNMTUgNHY1aDUiLz48cGF0aCBkPSJNOCAxMmw0IDZNMTIgMTJsLTQgNiIvPjwvc3ZnPicsCiAgcGRmOiAnPHN2ZyBjbGFzcz0iaWNvbiIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik00IDRoMTFsNSA1djExSDR6Ii8+PHBhdGggZD0iTTE1IDR2NWg1Ii8+PHBhdGggZD0iTTggMTd2LTRoMS41YTEuNSAxLjUgMCAwIDEgMCAzSDhNMTMgMTd2LTRoMS41TTEzIDE1aDEuNU0xNyAxM3Y0bDEuMi0yIi8+PC9zdmc+JywKICBleWU6ICc8c3ZnIGNsYXNzPSJpY29uIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBhdGggZD0iTTIgMTJzMy41LTcgMTAtNyAxMCA3IDEwIDctMy41IDctMTAgNy0xMC03LTEwLTd6Ii8+PGNpcmNsZSBjeD0iMTIiIGN5PSIxMiIgcj0iMyIvPjwvc3ZnPicsCiAgZXllT2ZmOiAnPHN2ZyBjbGFzcz0iaWNvbiIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik0yIDEyczMuNS03IDEwLTdjMiAwIDMuNi41IDUgMS4zTTIyIDEycy0xIDItMyAzLjhNOS45IDkuOWEzIDMgMCAwIDAgNC4yIDQuMiIvPjxwYXRoIGQ9Ik0zIDNsMTggMTgiLz48L3N2Zz4nLAogIG1haWw6ICc8c3ZnIGNsYXNzPSJpY29uIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHJlY3QgeD0iMyIgeT0iNSIgd2lkdGg9IjE4IiBoZWlnaHQ9IjE0IiByeD0iMiIvPjxwYXRoIGQ9Ik0zIDdsOSA2IDktNiIvPjwvc3ZnPicsCiAgbG9jazogJzxzdmcgY2xhc3M9Imljb24iIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJjdXJyZW50Q29sb3IiIHN0cm9rZS13aWR0aD0iMiIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj48cmVjdCB4PSI0IiB5PSIxMSIgd2lkdGg9IjE2IiBoZWlnaHQ9IjkiIHJ4PSIyIi8+PHBhdGggZD0iTTggMTFWN2E0IDQgMCAwIDEgOCAwdjQiLz48L3N2Zz4nLAogIHdpZmlPZmY6ICc8c3ZnIGNsYXNzPSJpY29uLXNtIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBhdGggZD0iTTIgMmwyMCAyMCIvPjxwYXRoIGQ9Ik04LjUgMTYuNWE1IDUgMCAwIDEgNyAwIi8+PHBhdGggZD0iTTUgMTIuNWExMCAxMCAwIDAgMSA0LTIuM00xOSAxMi41YTEwIDEwIDAgMCAwLTIuNS0xLjlNMTIgMjBoLjAxIi8+PC9zdmc+JywKICBiYWNrdXA6ICc8c3ZnIGNsYXNzPSJpY29uIiB2aWV3Qm94PSIwIDAgMjQgMjQiIGZpbGw9Im5vbmUiIHN0cm9rZT0iY3VycmVudENvbG9yIiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBhdGggZD0iTTIxIDEyYTkgOSAwIDEgMS0yLjYtNi40Ii8+PHBhdGggZD0iTTIxIDR2NWgtNSIvPjwvc3ZnPicsCn07CgovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KLy8gU3VwYWJhc2UgY29ubmVjdGlvbgovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KY29uc3QgZGxtU2IgPSB3aW5kb3cuc3VwYWJhc2UuY3JlYXRlQ2xpZW50KAogICdodHRwczovL29oa2t4YXhnaXppdXZkb3Zwbm11LnN1cGFiYXNlLmNvJywKICAnZXlKaGJHY2lPaUpJVXpJMU5pSXNJblI1Y0NJNklrcFhWQ0o5LmV5SnBjM01pT2lKemRYQmhZbUZ6WlNJc0luSmxaaUk2SW05b2EydDRZWGhuYVhwcGRYWmtiM1p3Ym0xMUlpd2ljbTlzWlNJNkltRnViMjRpTENKcFlYUWlPakUzT0RFeE9UUTJOemdzSW1WNGNDSTZNakE1TmpjM01EWTNPSDAuZUxfSHJIcFQyZ2U4U3NzWnpkMVgtU1Qtd3N3ek5TZGJWUFNkNmVTZGxKSScKKTsKCi8vIFN0b3JhZ2UgaGVscGVycyAoU3VwYWJhc2UtYmFja2VkKQpmdW5jdGlvbiB0YXNrVG9Sb3codCl7CiAgcmV0dXJuIHtpZDp0LmlkLCB0YXNrX25hbWU6dC50YXNrTmFtZSwgZGVzY3JpcHRpb246dC5kZXNjcmlwdGlvbnx8JycsIHN0YXR1czp0LnN0YXR1c3x8J3BlbmRpbmcnLAogICAgaXNfdXJnZW50OiEhdC5pc1VyZ2VudCwgbm90ZXM6dC5ub3Rlc3x8JycsIGNvbW1lbnRzOnQuY29tbWVudHN8fFtdLCBkdWVfZGF0ZTp0LmR1ZURhdGV8fCcnLCBjYXRlZ29yeTp0LmNhdGVnb3J5fHwnJywKICAgIGNyZWF0ZWRfYXQ6dC5jcmVhdGVkQXQsIHVwZGF0ZWRfYXQ6dC51cGRhdGVkQXQsIGNvbXBsZXRlZF9hdDp0LmNvbXBsZXRlZEF0LCBhcmNoaXZlZF9hdDp0LmFyY2hpdmVkQXQsIG93bmVyOidzdW11ZHUnfTsKfQpmdW5jdGlvbiByb3dUb1Rhc2socil7CiAgcmV0dXJuIHtpZDpyLmlkLCB0YXNrTmFtZTpyLnRhc2tfbmFtZSwgZGVzY3JpcHRpb246ci5kZXNjcmlwdGlvbnx8JycsIHN0YXR1czpyLnN0YXR1c3x8J3BlbmRpbmcnLAogICAgaXNVcmdlbnQ6ISFyLmlzX3VyZ2VudCwgbm90ZXM6ci5ub3Rlc3x8JycsIGNvbW1lbnRzOnIuY29tbWVudHN8fFtdLCBkdWVEYXRlOnIuZHVlX2RhdGV8fCcnLCBjYXRlZ29yeTpyLmNhdGVnb3J5fHwnJywKICAgIGNyZWF0ZWRBdDpyLmNyZWF0ZWRfYXQsIHVwZGF0ZWRBdDpyLnVwZGF0ZWRfYXQsIGNvbXBsZXRlZEF0OnIuY29tcGxldGVkX2F0LCBhcmNoaXZlZEF0OnIuYXJjaGl2ZWRfYXR9Owp9Cgphc3luYyBmdW5jdGlvbiBsb2FkVGFza3NGcm9tREIoKSB7CiAgY29uc3Qge2RhdGEsZXJyb3J9ID0gYXdhaXQgZGxtU2IuZnJvbSgnZmZfZHV0eV9saXN0Jykuc2VsZWN0KCcqJykuZXEoJ293bmVyJywnc3VtdWR1Jykub3JkZXIoJ2NyZWF0ZWRfYXQnLHthc2NlbmRpbmc6ZmFsc2V9KTsKICBpZihlcnJvcil7Y29uc29sZS5lcnJvcignTG9hZCBlcnJvcjonLGVycm9yKTsgcmV0dXJuIG51bGw7fQogIHJldHVybiBkYXRhID8gZGF0YS5tYXAocm93VG9UYXNrKSA6IG51bGw7Cn0KCmFzeW5jIGZ1bmN0aW9uIHNhdmVUYXNrVG9EQih0YXNrKSB7CiAgYXdhaXQgZGxtU2IuZnJvbSgnZmZfZHV0eV9saXN0JykudXBzZXJ0KHRhc2tUb1Jvdyh0YXNrKSk7Cn0KCmFzeW5jIGZ1bmN0aW9uIGRlbGV0ZVRhc2tGcm9tREIoaWQpIHsKICBhd2FpdCBkbG1TYi5mcm9tKCdmZl9kdXR5X2xpc3QnKS5kZWxldGUoKS5lcSgnaWQnLGlkKTsKfQoKYXN5bmMgZnVuY3Rpb24gc2F2ZUFsbFRhc2tzVG9EQih0YXNrcykgewogIGZvcihjb25zdCB0IG9mIHRhc2tzKXsgYXdhaXQgc2F2ZVRhc2tUb0RCKHQpOyB9Cn0KCmZ1bmN0aW9uIGxvYWRUYXNrcygpIHsKICB0cnkgewogICAgY29uc3QgcmF3ID0gbG9jYWxTdG9yYWdlLmdldEl0ZW0oJ2RsbV90YXNrcycpOwogICAgcmV0dXJuIHJhdyA/IEpTT04ucGFyc2UocmF3KSA6IG51bGw7CiAgfSBjYXRjaCB7IHJldHVybiBudWxsOyB9Cn0KCmZ1bmN0aW9uIHNhdmVUYXNrcyh0YXNrcykgewogIGxvY2FsU3RvcmFnZS5zZXRJdGVtKCdkbG1fdGFza3MnLCBKU09OLnN0cmluZ2lmeSh0YXNrcykpOwp9CgpmdW5jdGlvbiBsb2FkVGhlbWUoKSB7CiAgcmV0dXJuIGxvY2FsU3RvcmFnZS5nZXRJdGVtKCdkbG1fdGhlbWUnKSB8fAogICAgKHdpbmRvdy5tYXRjaE1lZGlhKCcocHJlZmVycy1jb2xvci1zY2hlbWU6IGRhcmspJykubWF0Y2hlcyA/ICdkYXJrJyA6ICdsaWdodCcpOwp9CgpmdW5jdGlvbiBzYXZlVGhlbWUodGhlbWUpIHsKICBsb2NhbFN0b3JhZ2Uuc2V0SXRlbSgnZGxtX3RoZW1lJywgdGhlbWUpOwp9CgpmdW5jdGlvbiBzZWVkU2FtcGxlVGFza3MoKSB7CiAgY29uc3Qgbm93ID0gbm93SVNPKCk7CiAgY29uc3QgaW5GaXZlRGF5cyA9IG5ldyBEYXRlKERhdGUubm93KCkgKyA1ICogMjQgKiA2MCAqIDYwICogMTAwMCkudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgY29uc3QgaW5Ud29EYXlzID0gbmV3IERhdGUoRGF0ZS5ub3coKSArIDIgKiAyNCAqIDYwICogNjAgKiAxMDAwKS50b0lTT1N0cmluZygpLnNsaWNlKDAsIDEwKTsKICByZXR1cm4gWwogICAgewogICAgICBpZDogdWlkKCksIHRhc2tOYW1lOiAnQ2FsbCBzdXBwbGllciBhYm91dCBkZWxheWVkIHNoaXBtZW50JywgZGVzY3JpcHRpb246ICdGb2xsb3cgdXAgb24gdGhlIG9yZGVyIHBsYWNlZCBsYXN0IHdlZWsuJywKICAgICAgc3RhdHVzOiAncGVuZGluZycsIGlzVXJnZW50OiB0cnVlLCBub3RlczogJycsIGNvbW1lbnRzOiBbXSwgZHVlRGF0ZTogaW5Ud29EYXlzLAogICAgICBjcmVhdGVkQXQ6IG5vdywgdXBkYXRlZEF0OiBub3csIGNvbXBsZXRlZEF0OiBudWxsLCBhcmNoaXZlZEF0OiBudWxsLAogICAgfSwKICAgIHsKICAgICAgaWQ6IHVpZCgpLCB0YXNrTmFtZTogJ1ByZXBhcmUgbW9udGhseSByZXBvcnQnLCBkZXNjcmlwdGlvbjogJ1N1bW1hcml6ZSBzYWxlcyBhbmQgZXhwZW5zZXMgZm9yIHRoZSBtb250aC4nLAogICAgICBzdGF0dXM6ICdpbl9wcm9ncmVzcycsIGlzVXJnZW50OiBmYWxzZSwgbm90ZXM6ICcnLCBjb21tZW50czogW10sIGR1ZURhdGU6IGluRml2ZURheXMsCiAgICAgIGNyZWF0ZWRBdDogbm93LCB1cGRhdGVkQXQ6IG5vdywgY29tcGxldGVkQXQ6IG51bGwsIGFyY2hpdmVkQXQ6IG51bGwsCiAgICB9LAogICAgewogICAgICBpZDogdWlkKCksIHRhc2tOYW1lOiAnUmV2aWV3IGludm9pY2VzJywgZGVzY3JpcHRpb246ICdDaGVjayBvdXRzdGFuZGluZyBpbnZvaWNlcyBhbmQgbWFyayBwYWlkIG9uZXMuJywKICAgICAgc3RhdHVzOiAncGVuZGluZycsIGlzVXJnZW50OiBmYWxzZSwgbm90ZXM6ICcnLCBjb21tZW50czogW10sIGR1ZURhdGU6ICcnLAogICAgICBjcmVhdGVkQXQ6IG5vdywgdXBkYXRlZEF0OiBub3csIGNvbXBsZXRlZEF0OiBudWxsLCBhcmNoaXZlZEF0OiBudWxsLAogICAgfSwKICBdOwp9Cjwvc2NyaXB0Pgo8c2NyaXB0PgovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KLy8gQXBwbGljYXRpb24gc3RhdGUKLy8gLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tCmNvbnN0IHN0YXRlID0gewogIHRoZW1lOiBsb2FkVGhlbWUoKSwKICB0YXNrczogW10sCiAgdmlldzogJ2FsbCcsIC8vICdhbGwnIHwgJ3VyZ2VudCcgfCAnb3ZlcmR1ZScgfCAnYXJjaGl2ZScKICBzZWFyY2g6ICcnLAogIHF1aWNrQWRkVXJnZW50OiBmYWxzZSwKICBxdWlja0FkZEpvc2h1YTogZmFsc2UsCiAgZWRpdGluZ1Rhc2tJZDogbnVsbCwKICBsYXN0VG91Y2hlZFRhc2tJZDogbnVsbCwKfTsKCmNvbnN0IEFSQ0hJVkVfUkVURU5USU9OX0RBWVMgPSAzMDsKCmZ1bmN0aW9uIGFwcGx5VGhlbWUoKSB7CiAgZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LmNsYXNzTGlzdC50b2dnbGUoJ2RhcmsnLCBzdGF0ZS50aGVtZSA9PT0gJ2RhcmsnKTsKICBzYXZlVGhlbWUoc3RhdGUudGhlbWUpOwp9CgpmdW5jdGlvbiB0b2dnbGVUaGVtZSgpIHsKICBzdGF0ZS50aGVtZSA9IHN0YXRlLnRoZW1lID09PSAnZGFyaycgPyAnbGlnaHQnIDogJ2RhcmsnOwogIGFwcGx5VGhlbWUoKTsKfQoKLy8gLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tCi8vIFRhc2sgQ1JVRCAob3BlcmF0ZXMgb24gc3RhdGUudGFza3MsIHBlcnNpc3RzIHRvIGxvY2FsU3RvcmFnZSkKLy8gLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tCmZ1bmN0aW9uIHBlcnNpc3RUYXNrcygpIHsKICBzYXZlVGFza3Moc3RhdGUudGFza3MpOwogIHNhdmVBbGxUYXNrc1RvREIoc3RhdGUudGFza3MpOwp9CgpmdW5jdGlvbiBnZXRUYXNrQnlJZChpZCkgewogIHJldHVybiBzdGF0ZS50YXNrcy5maW5kKCh0KSA9PiB0LmlkID09PSBpZCk7Cn0KCmZ1bmN0aW9uIGNyZWF0ZVRhc2soeyB0YXNrTmFtZSwgZGVzY3JpcHRpb24gPSAnJywgc3RhdHVzID0gJ3BlbmRpbmcnLCBpc1VyZ2VudCA9IGZhbHNlLCBub3RlcyA9ICcnLCBkdWVEYXRlID0gJycsIGNhdGVnb3J5ID0gJycgfSkgewogIGNvbnN0IG5vdyA9IG5vd0lTTygpOwogIGNvbnN0IHRhc2sgPSB7CiAgICBpZDogdWlkKCksCiAgICB0YXNrTmFtZTogdGFza05hbWUudHJpbSgpLAogICAgZGVzY3JpcHRpb246IChkZXNjcmlwdGlvbiB8fCAnJykudHJpbSgpLAogICAgc3RhdHVzOiBbJ3BlbmRpbmcnLCAnaW5fcHJvZ3Jlc3MnLCAnY29tcGxldGVkJ10uaW5jbHVkZXMoc3RhdHVzKSA/IHN0YXR1cyA6ICdwZW5kaW5nJywKICAgIGlzVXJnZW50OiAhIWlzVXJnZW50LAogICAgbm90ZXM6IChub3RlcyB8fCAnJykudHJpbSgpLAogICAgZHVlRGF0ZTogKGR1ZURhdGUgfHwgJycpLnRyaW0oKSwKICAgIGNhdGVnb3J5OiAoY2F0ZWdvcnkgfHwgJycpLnRyaW0oKSwKICAgIGNvbW1lbnRzOiBbXSwKICAgIGNyZWF0ZWRBdDogbm93LAogICAgdXBkYXRlZEF0OiBub3csCiAgICBjb21wbGV0ZWRBdDogc3RhdHVzID09PSAnY29tcGxldGVkJyA/IG5vdyA6IG51bGwsCiAgICBhcmNoaXZlZEF0OiBudWxsLAogIH07CiAgc3RhdGUudGFza3MudW5zaGlmdCh0YXNrKTsKICBzYXZlVGFza3Moc3RhdGUudGFza3MpOwogIHNhdmVUYXNrVG9EQih0YXNrKTsKICByZXR1cm4gdGFzazsKfQoKZnVuY3Rpb24gdXBkYXRlVGFzayhpZCwgdXBkYXRlcykgewogIGNvbnN0IHRhc2sgPSBnZXRUYXNrQnlJZChpZCk7CiAgaWYgKCF0YXNrKSByZXR1cm4gbnVsbDsKCiAgaWYgKHVwZGF0ZXMudGFza05hbWUgIT09IHVuZGVmaW5lZCkgdGFzay50YXNrTmFtZSA9IHVwZGF0ZXMudGFza05hbWUudHJpbSgpOwogIGlmICh1cGRhdGVzLmRlc2NyaXB0aW9uICE9PSB1bmRlZmluZWQpIHRhc2suZGVzY3JpcHRpb24gPSB1cGRhdGVzLmRlc2NyaXB0aW9uOwogIGlmICh1cGRhdGVzLm5vdGVzICE9PSB1bmRlZmluZWQpIHRhc2subm90ZXMgPSB1cGRhdGVzLm5vdGVzOwogIGlmICh1cGRhdGVzLmR1ZURhdGUgIT09IHVuZGVmaW5lZCkgdGFzay5kdWVEYXRlID0gdXBkYXRlcy5kdWVEYXRlOwogIGlmICh1cGRhdGVzLmlzVXJnZW50ICE9PSB1bmRlZmluZWQpIHRhc2suaXNVcmdlbnQgPSAhIXVwZGF0ZXMuaXNVcmdlbnQ7CiAgaWYgKHVwZGF0ZXMuY2F0ZWdvcnkgIT09IHVuZGVmaW5lZCkgdGFzay5jYXRlZ29yeSA9IHVwZGF0ZXMuY2F0ZWdvcnk7CgogIGlmICh1cGRhdGVzLnN0YXR1cyAhPT0gdW5kZWZpbmVkICYmIFsncGVuZGluZycsICdpbl9wcm9ncmVzcycsICdjb21wbGV0ZWQnXS5pbmNsdWRlcyh1cGRhdGVzLnN0YXR1cykpIHsKICAgIGNvbnN0IHdhc0NvbXBsZXRlZCA9IHRhc2suc3RhdHVzID09PSAnY29tcGxldGVkJzsKICAgIGNvbnN0IHdpbGxCZUNvbXBsZXRlZCA9IHVwZGF0ZXMuc3RhdHVzID09PSAnY29tcGxldGVkJzsKICAgIHRhc2suc3RhdHVzID0gdXBkYXRlcy5zdGF0dXM7CiAgICBpZiAod2lsbEJlQ29tcGxldGVkICYmICF3YXNDb21wbGV0ZWQpIHsKICAgICAgdGFzay5jb21wbGV0ZWRBdCA9IG5vd0lTTygpOwogICAgICB0YXNrLmFyY2hpdmVkQXQgPSBub3dJU08oKTsgLy8gY29tcGxldGluZyBhIHRhc2sgYXJjaGl2ZXMgaXQKICAgIH0gZWxzZSBpZiAoIXdpbGxCZUNvbXBsZXRlZCAmJiB3YXNDb21wbGV0ZWQpIHsKICAgICAgdGFzay5jb21wbGV0ZWRBdCA9IG51bGw7CiAgICAgIHRhc2suYXJjaGl2ZWRBdCA9IG51bGw7CiAgICB9CiAgfQoKICB0YXNrLnVwZGF0ZWRBdCA9IG5vd0lTTygpOwogIHNhdmVUYXNrcyhzdGF0ZS50YXNrcyk7CiAgc2F2ZVRhc2tUb0RCKHRhc2spOwogIHJldHVybiB0YXNrOwp9CgpmdW5jdGlvbiBkZWxldGVUYXNrKGlkKSB7CiAgc3RhdGUudGFza3MgPSBzdGF0ZS50YXNrcy5maWx0ZXIoKHQpID0+IHQuaWQgIT09IGlkKTsKICBzYXZlVGFza3Moc3RhdGUudGFza3MpOwogIGRlbGV0ZVRhc2tGcm9tREIoaWQpOwp9CgpmdW5jdGlvbiB0b2dnbGVDb21wbGV0ZShpZCkgewogIGNvbnN0IHRhc2sgPSBnZXRUYXNrQnlJZChpZCk7CiAgaWYgKCF0YXNrKSByZXR1cm47CiAgdXBkYXRlVGFzayhpZCwgeyBzdGF0dXM6IHRhc2suc3RhdHVzID09PSAnY29tcGxldGVkJyA/ICdwZW5kaW5nJyA6ICdjb21wbGV0ZWQnIH0pOwp9CgpmdW5jdGlvbiBhcmNoaXZlVGFzayhpZCkgewogIGNvbnN0IHRhc2sgPSBnZXRUYXNrQnlJZChpZCk7CiAgaWYgKCF0YXNrKSByZXR1cm47CiAgdGFzay5zdGF0dXMgPSAnY29tcGxldGVkJzsKICB0YXNrLmNvbXBsZXRlZEF0ID0gdGFzay5jb21wbGV0ZWRBdCB8fCBub3dJU08oKTsKICB0YXNrLmFyY2hpdmVkQXQgPSBub3dJU08oKTsKICB0YXNrLnVwZGF0ZWRBdCA9IG5vd0lTTygpOwogIHBlcnNpc3RUYXNrcygpOwp9CgpmdW5jdGlvbiByZXN0b3JlRnJvbUFyY2hpdmUoaWQpIHsKICBjb25zdCB0YXNrID0gZ2V0VGFza0J5SWQoaWQpOwogIGlmICghdGFzaykgcmV0dXJuOwogIHRhc2suc3RhdHVzID0gJ3BlbmRpbmcnOwogIHRhc2suY29tcGxldGVkQXQgPSBudWxsOwogIHRhc2suYXJjaGl2ZWRBdCA9IG51bGw7CiAgdGFzay51cGRhdGVkQXQgPSBub3dJU08oKTsKICBwZXJzaXN0VGFza3MoKTsKfQoKZnVuY3Rpb24gdG9nZ2xlVXJnZW50KGlkKSB7CiAgY29uc3QgdGFzayA9IGdldFRhc2tCeUlkKGlkKTsKICBpZiAoIXRhc2spIHJldHVybjsKICB1cGRhdGVUYXNrKGlkLCB7IGlzVXJnZW50OiAhdGFzay5pc1VyZ2VudCB9KTsKfQoKZnVuY3Rpb24gYWRkQ29tbWVudCh0YXNrSWQsIGNvbnRlbnQpIHsKICBjb25zdCB0YXNrID0gZ2V0VGFza0J5SWQodGFza0lkKTsKICBpZiAoIXRhc2sgfHwgIWNvbnRlbnQudHJpbSgpKSByZXR1cm4gbnVsbDsKICBjb25zdCBjb21tZW50ID0geyBpZDogdWlkKCksIGNvbnRlbnQ6IGNvbnRlbnQudHJpbSgpLCBjcmVhdGVkQXQ6IG5vd0lTTygpIH07CiAgdGFzay5jb21tZW50cy5wdXNoKGNvbW1lbnQpOwogIHRhc2sudXBkYXRlZEF0ID0gbm93SVNPKCk7CiAgcGVyc2lzdFRhc2tzKCk7CiAgcmV0dXJuIGNvbW1lbnQ7Cn0KCmZ1bmN0aW9uIGRlbGV0ZUNvbW1lbnQodGFza0lkLCBjb21tZW50SWQpIHsKICBjb25zdCB0YXNrID0gZ2V0VGFza0J5SWQodGFza0lkKTsKICBpZiAoIXRhc2spIHJldHVybjsKICB0YXNrLmNvbW1lbnRzID0gdGFzay5jb21tZW50cy5maWx0ZXIoKGMpID0+IGMuaWQgIT09IGNvbW1lbnRJZCk7CiAgdGFzay51cGRhdGVkQXQgPSBub3dJU08oKTsKICBwZXJzaXN0VGFza3MoKTsKfQoKLy8gLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tCi8vIEF1dG8tZGVsZXRlIGFyY2hpdmVkIHRhc2tzIGFmdGVyIEFSQ0hJVkVfUkVURU5USU9OX0RBWVMgZGF5cwovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KZnVuY3Rpb24gcHVyZ2VFeHBpcmVkQXJjaGl2ZWRUYXNrcygpIHsKICBjb25zdCBjdXRvZmYgPSBEYXRlLm5vdygpIC0gQVJDSElWRV9SRVRFTlRJT05fREFZUyAqIDI0ICogNjAgKiA2MCAqIDEwMDA7CiAgY29uc3QgYmVmb3JlID0gc3RhdGUudGFza3MubGVuZ3RoOwogIHN0YXRlLnRhc2tzID0gc3RhdGUudGFza3MuZmlsdGVyKCh0KSA9PiB7CiAgICBpZiAoIXQuYXJjaGl2ZWRBdCkgcmV0dXJuIHRydWU7CiAgICByZXR1cm4gbmV3IERhdGUodC5hcmNoaXZlZEF0KS5nZXRUaW1lKCkgPj0gY3V0b2ZmOwogIH0pOwogIGlmIChzdGF0ZS50YXNrcy5sZW5ndGggIT09IGJlZm9yZSkgc2F2ZVRhc2tzKHN0YXRlLnRhc2tzKTsKfQoKZnVuY3Rpb24gZGF5c1VudGlsUHVyZ2UodGFzaykgewogIGlmICghdGFzay5hcmNoaXZlZEF0KSByZXR1cm4gbnVsbDsKICBjb25zdCBhcmNoaXZlZFRpbWUgPSBuZXcgRGF0ZSh0YXNrLmFyY2hpdmVkQXQpLmdldFRpbWUoKTsKICBjb25zdCBwdXJnZVRpbWUgPSBhcmNoaXZlZFRpbWUgKyBBUkNISVZFX1JFVEVOVElPTl9EQVlTICogMjQgKiA2MCAqIDYwICogMTAwMDsKICBjb25zdCBkYXlzTGVmdCA9IE1hdGguY2VpbCgocHVyZ2VUaW1lIC0gRGF0ZS5ub3coKSkgLyAoMjQgKiA2MCAqIDYwICogMTAwMCkpOwogIHJldHVybiBNYXRoLm1heChkYXlzTGVmdCwgMCk7Cn0KCi8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQovLyBGaWx0ZXJpbmcsIHNvcnRpbmcgJiBzdGF0cwovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KY29uc3QgU1RBVFVTX09SREVSID0geyBwZW5kaW5nOiAwLCBpbl9wcm9ncmVzczogMSwgY29tcGxldGVkOiAyIH07CgpmdW5jdGlvbiBpc092ZXJkdWUodGFzaykgewogIGlmICghdGFzay5kdWVEYXRlIHx8IHRhc2suc3RhdHVzID09PSAnY29tcGxldGVkJyB8fCB0YXNrLmFyY2hpdmVkQXQpIHJldHVybiBmYWxzZTsKICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCkudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7CiAgcmV0dXJuIHRhc2suZHVlRGF0ZSA8IHRvZGF5Owp9CgpmdW5jdGlvbiBpc0R1ZVRvZGF5KHRhc2spIHsKICBpZiAoIXRhc2suZHVlRGF0ZSB8fCB0YXNrLnN0YXR1cyA9PT0gJ2NvbXBsZXRlZCcgfHwgdGFzay5hcmNoaXZlZEF0KSByZXR1cm4gZmFsc2U7CiAgY29uc3QgdG9kYXkgPSBuZXcgRGF0ZSgpLnRvSVNPU3RyaW5nKCkuc2xpY2UoMCwgMTApOwogIHJldHVybiB0YXNrLmR1ZURhdGUgPT09IHRvZGF5Owp9CgpmdW5jdGlvbiBnZXRUYXNrc0ZvclZpZXcodmlldykgewogIC8vICJBbGwgVGFza3MiIHBhbmVsID0gYWN0aXZlIChub24tYXJjaGl2ZWQpIHRhc2tzLCB1cmdlbnQgb3Igbm90LCBhbnkgc3RhdHVzIGV4Y2VwdCBjb21wbGV0ZWQvYXJjaGl2ZWQuCiAgLy8gIlVyZ2VudCIgcGFuZWwgPSBhY3RpdmUgdGFza3MgbWFya2VkIHVyZ2VudC4KICAvLyAiT3ZlcmR1ZSIgcGFuZWwgPSBhY3RpdmUgdGFza3Mgd2hvc2UgZHVlIGRhdGUgaGFzIHBhc3NlZC4KICAvLyAiQXJjaGl2ZSIgcGFuZWwgPSBjb21wbGV0ZWQgLyBhcmNoaXZlZCB0YXNrcy4KICBpZiAodmlldyA9PT0gJ2FyY2hpdmUnKSB7CiAgICByZXR1cm4gc3RhdGUudGFza3MuZmlsdGVyKCh0KSA9PiAhIXQuYXJjaGl2ZWRBdCk7CiAgfQogIGlmICh2aWV3ID09PSAnam9zaHVhJykgewogICAgcmV0dXJuIHN0YXRlLnRhc2tzLmZpbHRlcigodCkgPT4gIXQuYXJjaGl2ZWRBdCAmJiB0LmNhdGVnb3J5ID09PSAnam9zaHVhJyk7CiAgfQogIGlmICh2aWV3ID09PSAndXJnZW50JykgewogICAgcmV0dXJuIHN0YXRlLnRhc2tzLmZpbHRlcigodCkgPT4gIXQuYXJjaGl2ZWRBdCAmJiB0LmlzVXJnZW50KTsKICB9CiAgaWYgKHZpZXcgPT09ICdvdmVyZHVlJykgewogICAgcmV0dXJuIHN0YXRlLnRhc2tzLmZpbHRlcigodCkgPT4gIXQuYXJjaGl2ZWRBdCAmJiBpc092ZXJkdWUodCkpOwogIH0KICAvLyBhbGwgKGRlZmF1bHQpCiAgcmV0dXJuIHN0YXRlLnRhc2tzLmZpbHRlcigodCkgPT4gIXQuYXJjaGl2ZWRBdCk7Cn0KCmZ1bmN0aW9uIGdldEZpbHRlcmVkU29ydGVkVGFza3MoKSB7CiAgbGV0IGxpc3QgPSBnZXRUYXNrc0ZvclZpZXcoc3RhdGUudmlldyk7CgogIGlmIChzdGF0ZS5zZWFyY2gudHJpbSgpKSB7CiAgICBjb25zdCB0ZXJtID0gc3RhdGUuc2VhcmNoLnRyaW0oKS50b0xvd2VyQ2FzZSgpOwogICAgbGlzdCA9IGxpc3QuZmlsdGVyKCh0KSA9PgogICAgICB0LnRhc2tOYW1lLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXModGVybSkgfHwKICAgICAgKHQuZGVzY3JpcHRpb24gfHwgJycpLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXModGVybSkgfHwKICAgICAgKHQubm90ZXMgfHwgJycpLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXModGVybSkKICAgICk7CiAgfQoKICBsaXN0LnNvcnQoKGEsIGIpID0+IHsKICAgIGlmIChzdGF0ZS52aWV3ID09PSAnYXJjaGl2ZScpIHsKICAgICAgcmV0dXJuIG5ldyBEYXRlKGIuYXJjaGl2ZWRBdCkgLSBuZXcgRGF0ZShhLmFyY2hpdmVkQXQpOwogICAgfQogICAgaWYgKHN0YXRlLnZpZXcgPT09ICdvdmVyZHVlJykgewogICAgICAvLyBNb3N0IG92ZXJkdWUgKG9sZGVzdCBkdWUgZGF0ZSkgZmlyc3QKICAgICAgcmV0dXJuIGEuZHVlRGF0ZSA8IGIuZHVlRGF0ZSA/IC0xIDogYS5kdWVEYXRlID4gYi5kdWVEYXRlID8gMSA6IDA7CiAgICB9CiAgICBpZiAoYS5pc1VyZ2VudCAhPT0gYi5pc1VyZ2VudCkgcmV0dXJuIGEuaXNVcmdlbnQgPyAtMSA6IDE7CiAgICBjb25zdCBzYSA9IFNUQVRVU19PUkRFUlthLnN0YXR1c10gPz8gMzsKICAgIGNvbnN0IHNiID0gU1RBVFVTX09SREVSW2Iuc3RhdHVzXSA/PyAzOwogICAgaWYgKHNhICE9PSBzYikgcmV0dXJuIHNhIC0gc2I7CiAgICAvLyBUYXNrcyB3aXRoIGEgZHVlIGRhdGUgY29tZSBmaXJzdCAoc29vbmVzdCBkdWUgZGF0ZSBmaXJzdCk7IG5vIGR1ZSBkYXRlIHNvcnRzIGxhc3QKICAgIGlmIChhLmR1ZURhdGUgJiYgYi5kdWVEYXRlICYmIGEuZHVlRGF0ZSAhPT0gYi5kdWVEYXRlKSByZXR1cm4gYS5kdWVEYXRlIDwgYi5kdWVEYXRlID8gLTEgOiAxOwogICAgaWYgKGEuZHVlRGF0ZSAmJiAhYi5kdWVEYXRlKSByZXR1cm4gLTE7CiAgICBpZiAoIWEuZHVlRGF0ZSAmJiBiLmR1ZURhdGUpIHJldHVybiAxOwogICAgcmV0dXJuIG5ldyBEYXRlKGIuY3JlYXRlZEF0KSAtIG5ldyBEYXRlKGEuY3JlYXRlZEF0KTsKICB9KTsKCiAgcmV0dXJuIGxpc3Q7Cn0KCmZ1bmN0aW9uIGNvbXB1dGVTdGF0cygpIHsKICBjb25zdCBhY3RpdmUgPSBzdGF0ZS50YXNrcy5maWx0ZXIoKHQpID0+ICF0LmFyY2hpdmVkQXQpOwogIGNvbnN0IGFyY2hpdmVkID0gc3RhdGUudGFza3MuZmlsdGVyKCh0KSA9PiAhIXQuYXJjaGl2ZWRBdCk7CiAgY29uc3QgdG90YWwgPSBzdGF0ZS50YXNrcy5sZW5ndGg7CiAgY29uc3QgcGVuZGluZyA9IGFjdGl2ZS5maWx0ZXIoKHQpID0+IHQuc3RhdHVzID09PSAncGVuZGluZycpLmxlbmd0aDsKICBjb25zdCBpblByb2dyZXNzID0gYWN0aXZlLmZpbHRlcigodCkgPT4gdC5zdGF0dXMgPT09ICdpbl9wcm9ncmVzcycpLmxlbmd0aDsKICBjb25zdCBjb21wbGV0ZWQgPSBhcmNoaXZlZC5sZW5ndGg7CiAgY29uc3QgdXJnZW50ID0gYWN0aXZlLmZpbHRlcigodCkgPT4gdC5pc1VyZ2VudCkubGVuZ3RoOwogIGNvbnN0IG92ZXJkdWUgPSBhY3RpdmUuZmlsdGVyKCh0KSA9PiBpc092ZXJkdWUodCkpLmxlbmd0aDsKICBjb25zdCBqb3NodWEgPSBhY3RpdmUuZmlsdGVyKCh0KSA9PiB0LmNhdGVnb3J5ID09PSAnam9zaHVhJykubGVuZ3RoOwogIHJldHVybiB7IHRvdGFsLCBwZW5kaW5nLCBpblByb2dyZXNzLCBjb21wbGV0ZWQsIHVyZ2VudCwgb3ZlcmR1ZSwgam9zaHVhLCBhcmNoaXZlZENvdW50OiBhcmNoaXZlZC5sZW5ndGgsIGFjdGl2ZUNvdW50OiBhY3RpdmUubGVuZ3RoIH07Cn0KPC9zY3JpcHQ+CjxzY3JpcHQ+Ci8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQovLyBUb3AtbGV2ZWwgcmVuZGVyCi8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQpjb25zdCByb290ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2FwcCcpOwoKYXN5bmMgZnVuY3Rpb24gcmVuZGVyKCkgewogIGFwcGx5VGhlbWUoKTsKICBpZiAoc3RhdGUudGFza3MubGVuZ3RoID09PSAwKSB7CiAgICBjb25zdCBkYlRhc2tzID0gYXdhaXQgbG9hZFRhc2tzRnJvbURCKCk7CiAgICBpZihkYlRhc2tzICYmIGRiVGFza3MubGVuZ3RoID4gMCl7CiAgICAgIHN0YXRlLnRhc2tzID0gZGJUYXNrczsKICAgICAgc2F2ZVRhc2tzKHN0YXRlLnRhc2tzKTsKICAgIH0gZWxzZSB7CiAgICAgIGNvbnN0IHN0b3JlZCA9IGxvYWRUYXNrcygpOwogICAgICBpZihzdG9yZWQgJiYgc3RvcmVkLmxlbmd0aCA+IDApewogICAgICAgIHN0YXRlLnRhc2tzID0gc3RvcmVkOwogICAgICAgIHNhdmVBbGxUYXNrc1RvREIoc3RhdGUudGFza3MpOwogICAgICB9IGVsc2UgewogICAgICAgIHN0YXRlLnRhc2tzID0gc2VlZFNhbXBsZVRhc2tzKCk7CiAgICAgICAgc2F2ZVRhc2tzKHN0YXRlLnRhc2tzKTsKICAgICAgICBzYXZlQWxsVGFza3NUb0RCKHN0YXRlLnRhc2tzKTsKICAgICAgfQogICAgfQogIH0KICBwdXJnZUV4cGlyZWRBcmNoaXZlZFRhc2tzKCk7CiAgcmVuZGVyRGFzaGJvYXJkKCk7Cn0KCi8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQovLyBEYXNoYm9hcmQgc2hlbGwKLy8gLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tCmZ1bmN0aW9uIHJlbmRlckRhc2hib2FyZCgpIHsKICByb290LmlubmVySFRNTCA9IGAKICAgIDxkaXYgY2xhc3M9Im1pbi1oLXNjcmVlbiBwYi0xMiI+CiAgICAgIDxoZWFkZXIgY2xhc3M9ImJnLXdoaXRlLzgwIGRhcms6YmctZ3JheS05NTAvODAgYmFja2Ryb3AtYmx1ci1tZCBib3JkZXItYiBib3JkZXItZ3JheS0yMDAgZGFyazpib3JkZXItZ3JheS04MDAiPgogICAgICAgIDxkaXYgY2xhc3M9Im1heC13LTZ4bCBteC1hdXRvIHB4LTQgc206cHgtNiBweS0zIGZsZXggaXRlbXMtY2VudGVyIGp1c3RpZnktYmV0d2VlbiI+CiAgICAgICAgICA8ZGl2IGNsYXNzPSJmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMi41Ij4KICAgICAgICAgICAgPGRpdiBjbGFzcz0idy05IGgtOSByb3VuZGVkLXhsIGJnLXByaW1hcnktNjAwIHRleHQtd2hpdGUgZmxleCBpdGVtcy1jZW50ZXIganVzdGlmeS1jZW50ZXIgc2hhZG93LW1kIHNoYWRvdy1wcmltYXJ5LTYwMC8zMCI+CiAgICAgICAgICAgICAgJHtJQ09OUy5jaGVja30KICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgICA8L2Rpdj4KICAgICAgICAgIDxkaXYgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGdhcC0yIj4KICAgICAgICAgICAgPGJ1dHRvbiBpZD0iYnRuLXRoZW1lIiB0aXRsZT0iVG9nZ2xlIGRhcmsgbW9kZSIgY2xhc3M9InctMTEgaC0xMSBmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWNlbnRlciByb3VuZGVkLXhsIGJnLWdyYXktMTAwIGRhcms6YmctZ3JheS04MDAgaG92ZXI6YmctZ3JheS0yMDAgZGFyazpob3ZlcjpiZy1ncmF5LTcwMCB0cmFuc2l0aW9uLWNvbG9ycyI+CiAgICAgICAgICAgICAgJHtzdGF0ZS50aGVtZSA9PT0gJ2RhcmsnID8gSUNPTlMuc3VuIDogSUNPTlMubW9vbn0KICAgICAgICAgICAgPC9idXR0b24+CiAgICAgICAgICA8L2Rpdj4KICAgICAgICA8L2Rpdj4KCiAgICAgICAgPCEtLSBUb3AgbmF2aWdhdGlvbiBwYW5lbHM6IEFsbCBUYXNrcyAvIFVyZ2VudCAvIE92ZXJkdWUgLyBBcmNoaXZlIC0tPgogICAgICAgIDxuYXYgY2xhc3M9Im1heC13LTZ4bCBteC1hdXRvIHB4LTQgc206cHgtNiI+CiAgICAgICAgICA8ZGl2IGlkPSJ0b3AtbmF2LXBhbmVscyIgY2xhc3M9ImdyaWQgZ3JpZC1jb2xzLTIgc206Z3JpZC1jb2xzLTQgZ2FwLTIgcGItMi41Ij48L2Rpdj4KICAgICAgICA8L25hdj4KICAgICAgPC9oZWFkZXI+CgogICAgICA8bWFpbiBjbGFzcz0ibWF4LXctNnhsIG14LWF1dG8gcHgtNCBzbTpweC02IHB5LTUgc206cHktOCBzcGFjZS15LTUgc206c3BhY2UteS02Ij4KICAgICAgICA8IS0tIFF1aWNrIEFkZCAtLT4KICAgICAgICA8ZGl2IGNsYXNzPSJiZy13aGl0ZSBkYXJrOmJnLWdyYXktOTAwIGJvcmRlciBib3JkZXItZ3JheS0xMDAgZGFyazpib3JkZXItZ3JheS04MDAgcm91bmRlZC0yeGwgc2hhZG93LXNtIHAtNCBzbTpwLTUiPgogICAgICAgICAgPGRpdiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIganVzdGlmeS1iZXR3ZWVuIGZsZXgtd3JhcCBnYXAtMiBtYi0yIj4KICAgICAgICAgICAgPGxhYmVsIGNsYXNzPSJmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiB0ZXh0LXNtIGZvbnQtc2VtaWJvbGQgdGV4dC1ncmF5LTYwMCBkYXJrOnRleHQtZ3JheS0zMDAiPgogICAgICAgICAgICAgICR7SUNPTlMucGx1cy5yZXBsYWNlKCdpY29uJywgJ2ljb24tc20nKX0gPHNwYW4gY2xhc3M9InRleHQtcHJpbWFyeS02MDAgZGFyazp0ZXh0LXByaW1hcnktNDAwIj5RdWljayBBZGQgVGFzazwvc3Bhbj4KICAgICAgICAgICAgPC9sYWJlbD4KICAgICAgICAgICAgPGRpdiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIgZmxleC13cmFwIj4KICAgICAgICAgICAgICA8bGFiZWwgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGdhcC0xLjUgdGV4dC1zbSBmb250LXNlbWlib2xkIHB4LTMgcHktMS41IHJvdW5kZWQtZnVsbCBiZy1ncmF5LTEwMCBkYXJrOmJnLWdyYXktODAwIHRleHQtZ3JheS01MDAgZGFyazp0ZXh0LWdyYXktNDAwIj4KICAgICAgICAgICAgICAgICR7SUNPTlMuY2xvY2sucmVwbGFjZSgnaWNvbicsJ2ljb24tc20nKX0KICAgICAgICAgICAgICAgIDxpbnB1dCB0eXBlPSJkYXRlIiBpZD0icXVpY2stYWRkLWR1ZS1kYXRlIiBjbGFzcz0iYmctdHJhbnNwYXJlbnQgb3V0bGluZS1ub25lIHRleHQtZ3JheS03MDAgZGFyazp0ZXh0LWdyYXktMjAwIiAvPgogICAgICAgICAgICAgIDwvbGFiZWw+CiAgICAgICAgICAgICAgPGxhYmVsIGlkPSJxdWljay1hZGQtdXJnZW50LWxhYmVsIiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIgdGV4dC1zbSBmb250LXNlbWlib2xkIGN1cnNvci1wb2ludGVyIHNlbGVjdC1ub25lIHB4LTMgcHktMS41IHJvdW5kZWQtZnVsbCB0cmFuc2l0aW9uLWNvbG9ycyAkewogICAgICAgICAgICAgICAgc3RhdGUucXVpY2tBZGRVcmdlbnQKICAgICAgICAgICAgICAgICAgPyAnYmctcmVkLTYwMCB0ZXh0LXdoaXRlIHNoYWRvdy1tZCBzaGFkb3ctcmVkLTYwMC8zMCcKICAgICAgICAgICAgICAgICAgOiAnYmctZ3JheS0xMDAgZGFyazpiZy1ncmF5LTgwMCB0ZXh0LWdyYXktNTAwIGRhcms6dGV4dC1ncmF5LTQwMCBob3ZlcjpiZy1ncmF5LTIwMCBkYXJrOmhvdmVyOmJnLWdyYXktNzAwJwogICAgICAgICAgICAgIH0iPgogICAgICAgICAgICAgICAgPGlucHV0IHR5cGU9ImNoZWNrYm94IiBpZD0icXVpY2stYWRkLXVyZ2VudCIgY2xhc3M9ImRsbS1jaGVja2JveCIgJHtzdGF0ZS5xdWlja0FkZFVyZ2VudCA/ICdjaGVja2VkJyA6ICcnfSAvPgogICAgICAgICAgICAgICAgJHtJQ09OUy5hbGVydC5yZXBsYWNlKCdpY29uJywnaWNvbi1zbScpfSBNYXJrIGFzIFVyZ2VudAogICAgICAgICAgICAgIDwvbGFiZWw+CiAgICAgICAgICAgICAgPGxhYmVsIGlkPSJxdWljay1hZGQtam9zaHVhLWxhYmVsIiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIgdGV4dC1zbSBmb250LXNlbWlib2xkIGN1cnNvci1wb2ludGVyIHNlbGVjdC1ub25lIHB4LTMgcHktMS41IHJvdW5kZWQtZnVsbCB0cmFuc2l0aW9uLWNvbG9ycyAkewogICAgICAgICAgICAgICAgc3RhdGUuY3VycmVudFZpZXcgPT09ICdqb3NodWEnCiAgICAgICAgICAgICAgICAgID8gJ2JnLXB1cnBsZS02MDAgdGV4dC13aGl0ZSBzaGFkb3ctbWQgc2hhZG93LXB1cnBsZS02MDAvMzAnCiAgICAgICAgICAgICAgICAgIDogc3RhdGUucXVpY2tBZGRKb3NodWEKICAgICAgICAgICAgICAgICAgICA/ICdiZy1wdXJwbGUtNjAwIHRleHQtd2hpdGUgc2hhZG93LW1kIHNoYWRvdy1wdXJwbGUtNjAwLzMwJwogICAgICAgICAgICAgICAgICAgIDogJ2JnLWdyYXktMTAwIGRhcms6YmctZ3JheS04MDAgdGV4dC1ncmF5LTUwMCBkYXJrOnRleHQtZ3JheS00MDAgaG92ZXI6YmctZ3JheS0yMDAgZGFyazpob3ZlcjpiZy1ncmF5LTcwMCcKICAgICAgICAgICAgICB9IiBzdHlsZT0iJHtzdGF0ZS5jdXJyZW50VmlldyA9PT0gJ2pvc2h1YScgPyAnZGlzcGxheTpub25lJyA6ICcnfSI+CiAgICAgICAgICAgICAgICA8aW5wdXQgdHlwZT0iY2hlY2tib3giIGlkPSJxdWljay1hZGQtam9zaHVhIiBjbGFzcz0iZGxtLWNoZWNrYm94IiAke3N0YXRlLnF1aWNrQWRkSm9zaHVhID8gJ2NoZWNrZWQnIDogJyd9IC8+CiAgICAgICAgICAgICAgICBKb3NodWEKICAgICAgICAgICAgICA8L2xhYmVsPgogICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgIDwvZGl2PgogICAgICAgICAgPGRpdiBjbGFzcz0icmVsYXRpdmUiPgogICAgICAgICAgICA8dGV4dGFyZWEgaWQ9InF1aWNrLWFkZC1pbnB1dCIgcm93cz0iMiIgcGxhY2Vob2xkZXI9IkNhbGwgc3VwcGxpZXIKUHJlcGFyZSBtb250aGx5IHJlcG9ydApSZXZpZXcgaW52b2ljZXMKClR5cGUgYSB0YXNrIGFuZCBwcmVzcyBFbnRlciB0byBhZGQgaXQgaW5zdGFudGx5Li4uIgogICAgICAgICAgICAgIGNsYXNzPSJ3LWZ1bGwgcmVzaXplLW5vbmUgdGV4dC1sZyBweC00IHB5LTMuNSByb3VuZGVkLXhsIGJvcmRlciBib3JkZXItZ3JheS0yMDAgZGFyazpib3JkZXItZ3JheS03MDAgYmctZ3JheS01MCBkYXJrOmJnLWdyYXktODAwIGZvY3VzOmJnLXdoaXRlIGRhcms6Zm9jdXM6YmctZ3JheS05MDAgdHJhbnNpdGlvbi1jb2xvcnMgb3V0bGluZS1ub25lIGZvY3VzOnJpbmctMiBmb2N1czpyaW5nLXByaW1hcnktNTAwIj48L3RleHRhcmVhPgogICAgICAgICAgPC9kaXY+CiAgICAgICAgICA8cCBjbGFzcz0idGV4dC14cyB0ZXh0LWdyYXktNDAwIG10LTEuNSI+CiAgICAgICAgICAgIFByZXNzIDxrYmQgY2xhc3M9InB4LTEuNSBweS0wLjUgcm91bmRlZCBiZy1ncmF5LTEwMCBkYXJrOmJnLWdyYXktODAwIGJvcmRlciBib3JkZXItZ3JheS0yMDAgZGFyazpib3JkZXItZ3JheS03MDAiPkVudGVyPC9rYmQ+IHRvIGFkZCAmbWlkZG90OyA8a2JkIGNsYXNzPSJweC0xLjUgcHktMC41IHJvdW5kZWQgYmctZ3JheS0xMDAgZGFyazpiZy1ncmF5LTgwMCBib3JkZXIgYm9yZGVyLWdyYXktMjAwIGRhcms6Ym9yZGVyLWdyYXktNzAwIj5DdHJsK048L2tiZD4gdG8ganVtcCBoZXJlICZtaWRkb3Q7IGR1ZSBkYXRlIGFwcGxpZXMgdG8gYWxsIHRhc2tzIGFkZGVkIHRvZ2V0aGVyCiAgICAgICAgICA8L3A+CiAgICAgICAgPC9kaXY+CgogICAgICAgIDxkaXYgaWQ9InRvYXN0LWNvbnRhaW5lciI+PC9kaXY+CgogICAgICAgIDwhLS0gU2VhcmNoIC0tPgogICAgICAgIDxkaXYgY2xhc3M9ImJnLXdoaXRlIGRhcms6YmctZ3JheS05MDAgYm9yZGVyIGJvcmRlci1ncmF5LTEwMCBkYXJrOmJvcmRlci1ncmF5LTgwMCByb3VuZGVkLTJ4bCBzaGFkb3ctc20gcC00IHNtOnAtNSI+CiAgICAgICAgICA8ZGl2IGNsYXNzPSJyZWxhdGl2ZSI+CiAgICAgICAgICAgIDxzcGFuIGNsYXNzPSJhYnNvbHV0ZSBsZWZ0LTMgdG9wLTEvMiAtdHJhbnNsYXRlLXktMS8yIHRleHQtZ3JheS00MDAiPiR7SUNPTlMuc2VhcmNofTwvc3Bhbj4KICAgICAgICAgICAgPGlucHV0IGlkPSJzZWFyY2gtaW5wdXQiIHR5cGU9InRleHQiIHBsYWNlaG9sZGVyPSJTZWFyY2ggdGFza3MuLi4gKHByZXNzIC8gdG8gZm9jdXMpIiB2YWx1ZT0iJHtlc2NhcGVIdG1sKHN0YXRlLnNlYXJjaCl9IgogICAgICAgICAgICAgIGNsYXNzPSJ3LWZ1bGwgcGwtMTEgcHItMTAgcHktMyB0ZXh0LWxnIHJvdW5kZWQteGwgYm9yZGVyIGJvcmRlci1ncmF5LTIwMCBkYXJrOmJvcmRlci1ncmF5LTcwMCBiZy1ncmF5LTUwIGRhcms6YmctZ3JheS04MDAgZm9jdXM6Ymctd2hpdGUgZGFyazpmb2N1czpiZy1ncmF5LTkwMCB0cmFuc2l0aW9uLWNvbG9ycyBvdXRsaW5lLW5vbmUgZm9jdXM6cmluZy0yIGZvY3VzOnJpbmctcHJpbWFyeS01MDAiIC8+CiAgICAgICAgICAgIDxidXR0b24gaWQ9InNlYXJjaC1jbGVhciIgY2xhc3M9ImFic29sdXRlIHJpZ2h0LTMgdG9wLTEvMiAtdHJhbnNsYXRlLXktMS8yIHRleHQtZ3JheS00MDAgaG92ZXI6dGV4dC1ncmF5LTYwMCAke3N0YXRlLnNlYXJjaCA/ICcnIDogJ2hpZGRlbid9Ij4ke0lDT05TLnh9PC9idXR0b24+CiAgICAgICAgICA8L2Rpdj4KICAgICAgICA8L2Rpdj4KCiAgICAgICAgPCEtLSBTZWN0aW9uIGhlYWRpbmcgLS0+CiAgICAgICAgPGRpdiBpZD0ic2VjdGlvbi1oZWFkaW5nIj48L2Rpdj4KCiAgICAgICAgPGRpdiBpZD0idGFzay1saXN0LWNvbnRhaW5lciI+PC9kaXY+CgogICAgICAgIDwhLS0gPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09CiAgICAgICAgICAgICBCb3R0b20gdXRpbGl0eSBzZWN0aW9uOiBJbXBvcnQvRXhwb3J0LCBTaG9ydGN1dHMKICAgICAgICAgICAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAtLT4KICAgICAgICA8ZGl2IGNsYXNzPSJncmlkIGdyaWQtY29scy0xIHNtOmdyaWQtY29scy0yIGdhcC00IHB0LTQgYm9yZGVyLXQgYm9yZGVyLWdyYXktMjAwIGRhcms6Ym9yZGVyLWdyYXktODAwIj4KCiAgICAgICAgICA8IS0tIEltcG9ydCAvIEV4cG9ydCAtLT4KICAgICAgICAgIDxkaXYgY2xhc3M9ImJnLXdoaXRlIGRhcms6YmctZ3JheS05MDAgYm9yZGVyIGJvcmRlci1ncmF5LTEwMCBkYXJrOmJvcmRlci1ncmF5LTgwMCByb3VuZGVkLTJ4bCBzaGFkb3ctc20gcC00IHNtOnAtNSBzcGFjZS15LTMiPgogICAgICAgICAgICA8aDMgY2xhc3M9InRleHQtc20gZm9udC1zZW1pYm9sZCB0ZXh0LWdyYXktNjAwIGRhcms6dGV4dC1ncmF5LTMwMCI+SW1wb3J0ICZhbXA7IEV4cG9ydDwvaDM+CiAgICAgICAgICAgIDxkaXYgaWQ9ImRyb3B6b25lIiBjbGFzcz0iYm9yZGVyLTIgYm9yZGVyLWRhc2hlZCByb3VuZGVkLXhsIHAtNCB0ZXh0LWNlbnRlciBjdXJzb3ItcG9pbnRlciB0cmFuc2l0aW9uLWNvbG9ycyBib3JkZXItZ3JheS0yMDAgZGFyazpib3JkZXItZ3JheS03MDAgaG92ZXI6Ym9yZGVyLXByaW1hcnktNDAwIj4KICAgICAgICAgICAgICA8aW5wdXQgaWQ9ImZpbGUtaW5wdXQiIHR5cGU9ImZpbGUiIGFjY2VwdD0iLnhsc3gsLnhscywuY3N2LC5kb2N4IiBjbGFzcz0iaGlkZGVuIiAvPgogICAgICAgICAgICAgIDxkaXYgY2xhc3M9Im14LWF1dG8gbWItMSB0ZXh0LWdyYXktNDAwIGZsZXgganVzdGlmeS1jZW50ZXIiPiR7SUNPTlMudXBsb2FkfTwvZGl2PgogICAgICAgICAgICAgIDxwIGNsYXNzPSJmb250LXNlbWlib2xkIHRleHQteHMgc206dGV4dC1zbSIgaWQ9ImRyb3B6b25lLXRleHQiPkRyYWcgJmFtcDsgZHJvcCwgb3IgY2xpY2sgdG8gYnJvd3NlPC9wPgogICAgICAgICAgICAgIDxwIGNsYXNzPSJ0ZXh0LXhzIHRleHQtZ3JheS00MDAgbXQtMSI+Lnhsc3gsIC54bHMsIC5jc3YsIC5kb2N4PC9wPgogICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPGRpdiBpZD0iaW1wb3J0LW1lc3NhZ2UiPjwvZGl2PgogICAgICAgICAgICA8ZGl2IGNsYXNzPSJncmlkIGdyaWQtY29scy0yIGdhcC0yIj4KICAgICAgICAgICAgICA8YnV0dG9uIGlkPSJleHBvcnQtZXhjZWwiIGNsYXNzPSJmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWNlbnRlciBnYXAtMS41IHB4LTMgcHktMi41IHJvdW5kZWQteGwgYmctZ3JlZW4tNTAgZGFyazpiZy1ncmVlbi05NTAvMzAgdGV4dC1ncmVlbi03MDAgZGFyazp0ZXh0LWdyZWVuLTQwMCBmb250LXNlbWlib2xkIHRleHQteHMgc206dGV4dC1zbSBob3ZlcjpiZy1ncmVlbi0xMDAgZGFyazpob3ZlcjpiZy1ncmVlbi05NTAvNTAgdHJhbnNpdGlvbi1jb2xvcnMiPgogICAgICAgICAgICAgICAgJHtJQ09OUy5leGNlbH0gRXhjZWwKICAgICAgICAgICAgICA8L2J1dHRvbj4KICAgICAgICAgICAgICA8YnV0dG9uIGlkPSJleHBvcnQtcGRmIiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIganVzdGlmeS1jZW50ZXIgZ2FwLTEuNSBweC0zIHB5LTIuNSByb3VuZGVkLXhsIGJnLXJlZC01MCBkYXJrOmJnLXJlZC05NTAvMzAgdGV4dC1yZWQtNzAwIGRhcms6dGV4dC1yZWQtNDAwIGZvbnQtc2VtaWJvbGQgdGV4dC14cyBzbTp0ZXh0LXNtIGhvdmVyOmJnLXJlZC0xMDAgZGFyazpob3ZlcjpiZy1yZWQtOTUwLzUwIHRyYW5zaXRpb24tY29sb3JzIj4KICAgICAgICAgICAgICAgICR7SUNPTlMucGRmfSBQREYKICAgICAgICAgICAgICA8L2J1dHRvbj4KICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgICAgIDxwIGNsYXNzPSJ0ZXh0LXhzIHRleHQtZ3JheS00MDAiPkNTViBpcyBzdXBwb3J0ZWQgb24gaW1wb3J0ICh1c2UgRXhjZWwgZXhwb3J0IGZvciBDU1YtY29tcGF0aWJsZSBvdXRwdXQpLjwvcD4KICAgICAgICAgIDwvZGl2PgoKICAgICAgICAgIDwhLS0gS2V5Ym9hcmQgU2hvcnRjdXRzIC0tPgogICAgICAgICAgPGRpdiBjbGFzcz0iYmctd2hpdGUgZGFyazpiZy1ncmF5LTkwMCBib3JkZXIgYm9yZGVyLWdyYXktMTAwIGRhcms6Ym9yZGVyLWdyYXktODAwIHJvdW5kZWQtMnhsIHNoYWRvdy1zbSBwLTQgc206cC01Ij4KICAgICAgICAgICAgPGgzIGNsYXNzPSJ0ZXh0LXNtIGZvbnQtc2VtaWJvbGQgdGV4dC1ncmF5LTYwMCBkYXJrOnRleHQtZ3JheS0zMDAgbWItMiI+S2V5Ym9hcmQgU2hvcnRjdXRzPC9oMz4KICAgICAgICAgICAgPHVsIGNsYXNzPSJ0ZXh0LXNtIHRleHQtZ3JheS01MDAgZGFyazp0ZXh0LWdyYXktNDAwIHNwYWNlLXktMS41Ij4KICAgICAgICAgICAgICA8bGk+PGtiZCBjbGFzcz0icHgtMS41IHB5LTAuNSByb3VuZGVkIGJnLWdyYXktMTAwIGRhcms6YmctZ3JheS04MDAgYm9yZGVyIGJvcmRlci1ncmF5LTIwMCBkYXJrOmJvcmRlci1ncmF5LTcwMCI+Q3RybCtOPC9rYmQ+IE5ldyB0YXNrIChmb2N1cyBxdWljayBhZGQpPC9saT4KICAgICAgICAgICAgICA8bGk+PGtiZCBjbGFzcz0icHgtMS41IHB5LTAuNSByb3VuZGVkIGJnLWdyYXktMTAwIGRhcms6YmctZ3JheS04MDAgYm9yZGVyIGJvcmRlci1ncmF5LTIwMCBkYXJrOmJvcmRlci1ncmF5LTcwMCI+Q3RybCtVPC9rYmQ+IFRvZ2dsZSB1cmdlbnQgb24gbGFzdCB0YXNrPC9saT4KICAgICAgICAgICAgICA8bGk+PGtiZCBjbGFzcz0icHgtMS41IHB5LTAuNSByb3VuZGVkIGJnLWdyYXktMTAwIGRhcms6YmctZ3JheS04MDAgYm9yZGVyIGJvcmRlci1ncmF5LTIwMCBkYXJrOmJvcmRlci1ncmF5LTcwMCI+Lzwva2JkPiBGb2N1cyBzZWFyY2g8L2xpPgogICAgICAgICAgICAgIDxsaT48a2JkIGNsYXNzPSJweC0xLjUgcHktMC41IHJvdW5kZWQgYmctZ3JheS0xMDAgZGFyazpiZy1ncmF5LTgwMCBib3JkZXIgYm9yZGVyLWdyYXktMjAwIGRhcms6Ym9yZGVyLWdyYXktNzAwIj5FbnRlcjwva2JkPiBBZGQgdGFzayAvIHNhdmU8L2xpPgogICAgICAgICAgICAgIDxsaT48a2JkIGNsYXNzPSJweC0xLjUgcHktMC41IHJvdW5kZWQgYmctZ3JheS0xMDAgZGFyazpiZy1ncmF5LTgwMCBib3JkZXIgYm9yZGVyLWdyYXktMjAwIGRhcms6Ym9yZGVyLWdyYXktNzAwIj5Fc2M8L2tiZD4gQ2xvc2UgZGlhbG9nPC9saT4KICAgICAgICAgICAgPC91bD4KICAgICAgICAgIDwvZGl2PgogICAgICAgIDwvZGl2PgogICAgICA8L21haW4+CgogICAgICA8ZGl2IGlkPSJ0YXNrLW1vZGFsLWNvbnRhaW5lciI+PC9kaXY+CiAgICA8L2Rpdj4KICBgOwoKICByZW5kZXJTdGF0cygpOwogIHJlbmRlclNlY3Rpb25IZWFkaW5nKCk7CiAgcmVuZGVyVGFza0xpc3QoKTsKICB3aXJlRGFzaGJvYXJkRXZlbnRzKCk7Cn0KPC9zY3JpcHQ+CjxzY3JpcHQ+Ci8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQovLyBTdGF0cyBjYXJkcwovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KZnVuY3Rpb24gcmVuZGVyUXVpY2tBZGRVcmdlbnRMYWJlbCgpIHsKICBjb25zdCBsYWJlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdxdWljay1hZGQtdXJnZW50LWxhYmVsJyk7CiAgaWYgKCFsYWJlbCkgcmV0dXJuOwogIGlmIChzdGF0ZS5xdWlja0FkZFVyZ2VudCkgewogICAgbGFiZWwuY2xhc3NMaXN0LnJlbW92ZSgnYmctZ3JheS0xMDAnLCAnZGFyazpiZy1ncmF5LTgwMCcsICd0ZXh0LWdyYXktNTAwJywgJ2Rhcms6dGV4dC1ncmF5LTQwMCcsICdob3ZlcjpiZy1ncmF5LTIwMCcsICdkYXJrOmhvdmVyOmJnLWdyYXktNzAwJyk7CiAgICBsYWJlbC5jbGFzc0xpc3QuYWRkKCdiZy1yZWQtNjAwJywgJ3RleHQtd2hpdGUnLCAnc2hhZG93LW1kJywgJ3NoYWRvdy1yZWQtNjAwLzMwJyk7CiAgfSBlbHNlIHsKICAgIGxhYmVsLmNsYXNzTGlzdC5yZW1vdmUoJ2JnLXJlZC02MDAnLCAndGV4dC13aGl0ZScsICdzaGFkb3ctbWQnLCAnc2hhZG93LXJlZC02MDAvMzAnKTsKICAgIGxhYmVsLmNsYXNzTGlzdC5hZGQoJ2JnLWdyYXktMTAwJywgJ2Rhcms6YmctZ3JheS04MDAnLCAndGV4dC1ncmF5LTUwMCcsICdkYXJrOnRleHQtZ3JheS00MDAnLCAnaG92ZXI6YmctZ3JheS0yMDAnLCAnZGFyazpob3ZlcjpiZy1ncmF5LTcwMCcpOwogIH0KfQoKZnVuY3Rpb24gcmVuZGVyUXVpY2tBZGRKb3NodWFMYWJlbCgpIHsKICBjb25zdCBsYWJlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdxdWljay1hZGQtam9zaHVhLWxhYmVsJyk7CiAgaWYgKCFsYWJlbCkgcmV0dXJuOwogIGlmIChzdGF0ZS5xdWlja0FkZEpvc2h1YSkgewogICAgbGFiZWwuY2xhc3NMaXN0LnJlbW92ZSgnYmctZ3JheS0xMDAnLCAnZGFyazpiZy1ncmF5LTgwMCcsICd0ZXh0LWdyYXktNTAwJywgJ2Rhcms6dGV4dC1ncmF5LTQwMCcsICdob3ZlcjpiZy1ncmF5LTIwMCcsICdkYXJrOmhvdmVyOmJnLWdyYXktNzAwJyk7CiAgICBsYWJlbC5jbGFzc0xpc3QuYWRkKCdiZy1wdXJwbGUtNjAwJywgJ3RleHQtd2hpdGUnLCAnc2hhZG93LW1kJywgJ3NoYWRvdy1wdXJwbGUtNjAwLzMwJyk7CiAgfSBlbHNlIHsKICAgIGxhYmVsLmNsYXNzTGlzdC5yZW1vdmUoJ2JnLXB1cnBsZS02MDAnLCAndGV4dC13aGl0ZScsICdzaGFkb3ctbWQnLCAnc2hhZG93LXB1cnBsZS02MDAvMzAnKTsKICAgIGxhYmVsLmNsYXNzTGlzdC5hZGQoJ2JnLWdyYXktMTAwJywgJ2Rhcms6YmctZ3JheS04MDAnLCAndGV4dC1ncmF5LTUwMCcsICdkYXJrOnRleHQtZ3JheS00MDAnLCAnaG92ZXI6YmctZ3JheS0yMDAnLCAnZGFyazpob3ZlcjpiZy1ncmF5LTcwMCcpOwogIH0KfQoKZnVuY3Rpb24gcmVuZGVyU3RhdHMoKSB7CiAgcmVuZGVyVG9wTmF2UGFuZWxzKCk7Cn0KCi8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQovLyBUb3AgbmF2aWdhdGlvbiBwYW5lbHM6IEFsbCBUYXNrcyAvIFVyZ2VudCAvIE92ZXJkdWUgLyBBcmNoaXZlCi8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQpjb25zdCBWSUVXX1BBTkVMUyA9IFsKICB7IGtleTogJ2FsbCcsIGxhYmVsOiAnQWxsIFRhc2tzJywgaWNvbjogSUNPTlMubGlzdCwgYWN0aXZlQ29sb3I6ICdiZy1wcmltYXJ5LTYwMCB0ZXh0LXdoaXRlIHNoYWRvdy1sZyBzaGFkb3ctcHJpbWFyeS02MDAvMzAnIH0sCiAgeyBrZXk6ICdqb3NodWEnLCBsYWJlbDogJ0pvc2h1YScsIGljb246IElDT05TLmxpc3QsIGFjdGl2ZUNvbG9yOiAnYmctcHVycGxlLTYwMCB0ZXh0LXdoaXRlIHNoYWRvdy1sZyBzaGFkb3ctcHVycGxlLTYwMC8zMCcgfSwKICB7IGtleTogJ3VyZ2VudCcsIGxhYmVsOiAnVXJnZW50JywgaWNvbjogSUNPTlMuYWxlcnQsIGFjdGl2ZUNvbG9yOiAnYmctcmVkLTYwMCB0ZXh0LXdoaXRlIHNoYWRvdy1sZyBzaGFkb3ctcmVkLTYwMC8zMCcgfSwKICB7IGtleTogJ292ZXJkdWUnLCBsYWJlbDogJ092ZXJkdWUnLCBpY29uOiBJQ09OUy5jbG9jaywgYWN0aXZlQ29sb3I6ICdiZy1yZWQtNjAwIHRleHQtd2hpdGUgc2hhZG93LWxnIHNoYWRvdy1yZWQtNjAwLzMwJyB9LAogIHsga2V5OiAnYXJjaGl2ZScsIGxhYmVsOiAnQXJjaGl2ZScsIGljb246IElDT05TLmJhY2t1cCwgYWN0aXZlQ29sb3I6ICdiZy1ncmF5LTcwMCB0ZXh0LXdoaXRlIHNoYWRvdy1sZyBzaGFkb3ctZ3JheS03MDAvMzAgZGFyazpiZy1ncmF5LTYwMCcgfSwKXTsKCmZ1bmN0aW9uIHJlbmRlclRvcE5hdlBhbmVscygpIHsKICBjb25zdCBjb250YWluZXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndG9wLW5hdi1wYW5lbHMnKTsKICBpZiAoIWNvbnRhaW5lcikgcmV0dXJuOwogIGNvbnN0IHN0YXRzID0gY29tcHV0ZVN0YXRzKCk7CiAgY29uc3QgY291bnRzID0geyBhbGw6IHN0YXRzLmFjdGl2ZUNvdW50LCBqb3NodWE6IHN0YXRzLmpvc2h1YSwgdXJnZW50OiBzdGF0cy51cmdlbnQsIG92ZXJkdWU6IHN0YXRzLm92ZXJkdWUsIGFyY2hpdmU6IHN0YXRzLmFyY2hpdmVkQ291bnQgfTsKCiAgY29udGFpbmVyLmlubmVySFRNTCA9IFZJRVdfUEFORUxTLm1hcCgocCkgPT4gewogICAgY29uc3QgaXNBY3RpdmUgPSBzdGF0ZS52aWV3ID09PSBwLmtleTsKICAgIHJldHVybiBgCiAgICAgIDxidXR0b24gZGF0YS12aWV3PSIke3Aua2V5fSIgY2xhc3M9Im5hdi1wYW5lbC1idG4gZmxleCBpdGVtcy1jZW50ZXIganVzdGlmeS1jZW50ZXIgZ2FwLTEuNSBzbTpnYXAtMiBweC0yIHNtOnB4LTQgcHktMS41IHNtOnB5LTIgcm91bmRlZC14bCB0ZXh0LXhzIHNtOnRleHQtc20gZm9udC1ib2xkIHRyYW5zaXRpb24tYWxsICR7CiAgICAgICAgaXNBY3RpdmUgPyBwLmFjdGl2ZUNvbG9yIDogJ2JnLWdyYXktMTAwIGRhcms6YmctZ3JheS04MDAgdGV4dC1ncmF5LTYwMCBkYXJrOnRleHQtZ3JheS0zMDAgaG92ZXI6YmctZ3JheS0yMDAgZGFyazpob3ZlcjpiZy1ncmF5LTcwMCcKICAgICAgfSI+CiAgICAgICAgJHtwLmljb24ucmVwbGFjZSgnaWNvbicsICdpY29uLXNtJyl9CiAgICAgICAgPHNwYW4gY2xhc3M9InRydW5jYXRlIj4ke3AubGFiZWx9PC9zcGFuPgogICAgICAgIDxzcGFuIGNsYXNzPSJpbmxpbmUtZmxleCBpdGVtcy1jZW50ZXIganVzdGlmeS1jZW50ZXIgbWluLXctWzEuMjVyZW1dIGgtNSBweC0xIHJvdW5kZWQtZnVsbCB0ZXh0LXhzIGZvbnQtZXh0cmFib2xkICR7CiAgICAgICAgICBpc0FjdGl2ZSA/ICdiZy13aGl0ZS8yNSB0ZXh0LXdoaXRlJyA6ICdiZy13aGl0ZSBkYXJrOmJnLWdyYXktOTAwIHRleHQtZ3JheS01MDAgZGFyazp0ZXh0LWdyYXktNDAwJwogICAgICAgIH0iPiR7Y291bnRzW3Aua2V5XX08L3NwYW4+CiAgICAgIDwvYnV0dG9uPgogICAgYDsKICB9KS5qb2luKCcnKTsKCiAgY29udGFpbmVyLnF1ZXJ5U2VsZWN0b3JBbGwoJy5uYXYtcGFuZWwtYnRuJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIHN0YXRlLnZpZXcgPSBidG4uZGF0YXNldC52aWV3OwogICAgICBzdGF0ZS5zZWFyY2ggPSAnJzsKICAgICAgcmVuZGVyVG9wTmF2UGFuZWxzKCk7CiAgICAgIHJlbmRlclNlY3Rpb25IZWFkaW5nKCk7CiAgICAgIHJlbmRlclRhc2tMaXN0KCk7CiAgICAgIGNvbnN0IHNlYXJjaElucHV0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NlYXJjaC1pbnB1dCcpOwogICAgICBpZiAoc2VhcmNoSW5wdXQpIHNlYXJjaElucHV0LnZhbHVlID0gJyc7CiAgICAgIGNvbnN0IGNsZWFyQnRuID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NlYXJjaC1jbGVhcicpOwogICAgICBpZiAoY2xlYXJCdG4pIGNsZWFyQnRuLmNsYXNzTGlzdC5hZGQoJ2hpZGRlbicpOwogICAgfSk7CiAgfSk7Cn0KCi8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQovLyBTZWN0aW9uIGhlYWRpbmcgKHNob3duIGFib3ZlIHRoZSB0YXNrIGxpc3QsIGRlc2NyaWJlcyBjdXJyZW50IHBhbmVsKQovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KY29uc3QgU0VDVElPTl9JTkZPID0gewogIGFsbDogeyB0aXRsZTogJ0FsbCBUYXNrcycsIGRlc2M6ICdBbGwgcGVuZGluZyBhbmQgaW4tcHJvZ3Jlc3MgdGFza3MuIFVyZ2VudCBvbmVzIGFyZSBoaWdobGlnaHRlZCBhbmQgc2hvd24gZmlyc3QuJyB9LAogIGpvc2h1YTogeyB0aXRsZTogJ0pvc2h1YScsIGRlc2M6ICdUYXNrcyByZWxhdGVkIHRvIEpvc2h1YS4nIH0sCiAgdXJnZW50OiB7IHRpdGxlOiAnVXJnZW50IFRhc2tzJywgZGVzYzogJ1Rhc2tzIG1hcmtlZCB1cmdlbnQuIFRoZXkgc3RheSBoZXJlIHVudGlsIHVyZ2VuY3kgaXMgcmVtb3ZlZCBvciB0aGUgdGFzayBpcyBhcmNoaXZlZC4nIH0sCiAgb3ZlcmR1ZTogeyB0aXRsZTogJ092ZXJkdWUgVGFza3MnLCBkZXNjOiAnVGFza3Mgd2hvc2UgZHVlIGRhdGUgaGFzIHBhc3NlZCBhbmQgYXJlIG5vdCB5ZXQgY29tcGxldGVkLCBtb3N0IG92ZXJkdWUgZmlyc3QuJyB9LAogIGFyY2hpdmU6IHsgdGl0bGU6ICdBcmNoaXZlJywgZGVzYzogYENvbXBsZXRlZCAvIGFyY2hpdmVkIHRhc2tzLiBBdXRvLWRlbGV0ZWQgJHtBUkNISVZFX1JFVEVOVElPTl9EQVlTfSBkYXlzIGFmdGVyIGFyY2hpdmluZy5gIH0sCn07CgpmdW5jdGlvbiByZW5kZXJTZWN0aW9uSGVhZGluZygpIHsKICBjb25zdCBjb250YWluZXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2VjdGlvbi1oZWFkaW5nJyk7CiAgaWYgKCFjb250YWluZXIpIHJldHVybjsKICBjb25zdCBpbmZvID0gU0VDVElPTl9JTkZPW3N0YXRlLnZpZXddIHx8IFNFQ1RJT05fSU5GTy5hbGw7CiAgY29udGFpbmVyLmlubmVySFRNTCA9IGAKICAgIDxkaXY+CiAgICAgIDxoMiBjbGFzcz0idGV4dC14bCBzbTp0ZXh0LTJ4bCBmb250LWV4dHJhYm9sZCI+JHtpbmZvLnRpdGxlfTwvaDI+CiAgICAgIDxwIGNsYXNzPSJ0ZXh0LXNtIHRleHQtZ3JheS01MDAgZGFyazp0ZXh0LWdyYXktNDAwIG10LTAuNSI+JHtpbmZvLmRlc2N9PC9wPgogICAgPC9kaXY+CiAgYDsKfQoKLy8gLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tCi8vIFRhc2sgbGlzdCAmIGNhcmRzCi8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQpjb25zdCBTVEFUVVNfQ09ORklHID0gewogIHBlbmRpbmc6IHsgbGFiZWw6ICdQZW5kaW5nJywgaWNvbjogSUNPTlMuY2xvY2ssIGNsYXNzTmFtZTogJ2JnLWdyYXktMTAwIHRleHQtZ3JheS02MDAgZGFyazpiZy1ncmF5LTgwMCBkYXJrOnRleHQtZ3JheS0zMDAnIH0sCiAgaW5fcHJvZ3Jlc3M6IHsgbGFiZWw6ICdJbiBQcm9ncmVzcycsIGljb246IElDT05TLmxvYWRlciwgY2xhc3NOYW1lOiAnYmctYmx1ZS01MCB0ZXh0LWJsdWUtNjAwIGRhcms6YmctYmx1ZS05NTAvNDAgZGFyazp0ZXh0LWJsdWUtNDAwJyB9LAogIGNvbXBsZXRlZDogeyBsYWJlbDogJ0NvbXBsZXRlZCcsIGljb246IElDT05TLmNoZWNrLCBjbGFzc05hbWU6ICdiZy1ncmVlbi01MCB0ZXh0LWdyZWVuLTYwMCBkYXJrOmJnLWdyZWVuLTk1MC80MCBkYXJrOnRleHQtZ3JlZW4tNDAwJyB9LAp9OwoKZnVuY3Rpb24gdGFza0NhcmRIdG1sKHRhc2spIHsKICBjb25zdCBzdGF0dXNJbmZvID0gU1RBVFVTX0NPTkZJR1t0YXNrLnN0YXR1c10gfHwgU1RBVFVTX0NPTkZJRy5wZW5kaW5nOwogIGNvbnN0IGlzQ29tcGxldGVkID0gdGFzay5zdGF0dXMgPT09ICdjb21wbGV0ZWQnOwogIGNvbnN0IGlzQXJjaGl2ZWQgPSAhIXRhc2suYXJjaGl2ZWRBdDsKICBjb25zdCBvdmVyZHVlID0gaXNPdmVyZHVlKHRhc2spOwogIGNvbnN0IGR1ZVRvZGF5ID0gaXNEdWVUb2RheSh0YXNrKTsKICBjb25zdCBwdXJnZURheXMgPSBpc0FyY2hpdmVkID8gZGF5c1VudGlsUHVyZ2UodGFzaykgOiBudWxsOwoKICBjb25zdCBjb21wbGV0ZUJ0biA9IGlzQXJjaGl2ZWQgPyBgCiAgICAgICAgPGJ1dHRvbiBkYXRhLWFjdGlvbj0icmVzdG9yZSIgZGF0YS1pZD0iJHt0YXNrLmlkfSIgdGl0bGU9IlJlc3RvcmUgdG8gUGVuZGluZyIKICAgICAgICAgIGNsYXNzPSJmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWNlbnRlciBnYXAtMSBweC0yLjUgc206cHgtMyBoLTggcm91bmRlZC1sZyB0ZXh0LXhzIHNtOnRleHQtc20gZm9udC1zZW1pYm9sZCB0cmFuc2l0aW9uLWNvbG9ycyBiZy1wcmltYXJ5LTUwIHRleHQtcHJpbWFyeS03MDAgZGFyazpiZy1wcmltYXJ5LTk1MC8zMCBkYXJrOnRleHQtcHJpbWFyeS00MDAgaG92ZXI6YmctcHJpbWFyeS0xMDAgZGFyazpob3ZlcjpiZy1wcmltYXJ5LTk1MC81MCI+CiAgICAgICAgICAke0lDT05TLnVwbG9hZC5yZXBsYWNlKCdpY29uJywnaWNvbi1zbScpfSA8c3BhbiBjbGFzcz0iaGlkZGVuIHNtOmlubGluZSI+UmVzdG9yZTwvc3Bhbj4KICAgICAgICA8L2J1dHRvbj5gIDogYAogICAgICAgIDxidXR0b24gZGF0YS1hY3Rpb249InRvZ2dsZS1jb21wbGV0ZSIgZGF0YS1pZD0iJHt0YXNrLmlkfSIgdGl0bGU9Ik1hcmsgYXMgY29tcGxldGUgKG1vdmVzIHRvIEFyY2hpdmUpIgogICAgICAgICAgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGp1c3RpZnktY2VudGVyIGdhcC0xIHB4LTIuNSBzbTpweC0zIGgtOCByb3VuZGVkLWxnIHRleHQteHMgc206dGV4dC1zbSBmb250LXNlbWlib2xkIHRyYW5zaXRpb24tY29sb3JzIGJnLWdyZWVuLTUwIHRleHQtZ3JlZW4tNzAwIGRhcms6YmctZ3JlZW4tOTUwLzMwIGRhcms6dGV4dC1ncmVlbi00MDAgaG92ZXI6YmctZ3JlZW4tMTAwIGRhcms6aG92ZXI6YmctZ3JlZW4tOTUwLzUwIj4KICAgICAgICAgICR7SUNPTlMuY2hlY2sucmVwbGFjZSgnaWNvbicsJ2ljb24tc20nKX0gPHNwYW4gY2xhc3M9ImhpZGRlbiBzbTppbmxpbmUiPkNvbXBsZXRlPC9zcGFuPgogICAgICAgIDwvYnV0dG9uPmA7CgogIHJldHVybiBgCiAgICA8ZGl2IGNsYXNzPSJncm91cCByZWxhdGl2ZSBiZy13aGl0ZSBkYXJrOmJnLWdyYXktOTAwIGJvcmRlciByb3VuZGVkLXhsIHNoYWRvdy1zbSBweC0zIHB5LTIuNSBzbTpweC00IHNtOnB5LTMgZmxleCBpdGVtcy1jZW50ZXIgZmxleC13cmFwIGdhcC0yLjUgc206Z2FwLTMgdHJhbnNpdGlvbi1hbGwgYW5pbS1mYWRlICR7CiAgICAgICh0YXNrLmlzVXJnZW50IHx8IG92ZXJkdWUpICYmICFpc0FyY2hpdmVkCiAgICAgICAgPyAnYm9yZGVyLXJlZC0zMDAgZGFyazpib3JkZXItcmVkLTgwMCByaW5nLTEgcmluZy1yZWQtMTAwIGRhcms6cmluZy1yZWQtOTAwLzQwIGJnLXJlZC01MC80MCBkYXJrOmJnLXJlZC05NTAvMTAnCiAgICAgICAgOiBkdWVUb2RheSAmJiAhaXNBcmNoaXZlZAogICAgICAgICAgPyAnYm9yZGVyLW9yYW5nZS0zMDAgZGFyazpib3JkZXItb3JhbmdlLTgwMCByaW5nLTEgcmluZy1vcmFuZ2UtMTAwIGRhcms6cmluZy1vcmFuZ2UtOTAwLzQwIGJnLW9yYW5nZS01MC80MCBkYXJrOmJnLW9yYW5nZS05NTAvMTAnCiAgICAgICAgICA6ICdib3JkZXItZ3JheS0xMDAgZGFyazpib3JkZXItZ3JheS04MDAnCiAgICB9Ij4KICAgICAgJHshaXNBcmNoaXZlZCA/IGAKICAgICAgPGJ1dHRvbiBkYXRhLWFjdGlvbj0idG9nZ2xlLWNvbXBsZXRlIiBkYXRhLWlkPSIke3Rhc2suaWR9IiB0aXRsZT0iTWFyayBhcyBjb21wbGV0ZSIKICAgICAgICBjbGFzcz0iZmxleC1zaHJpbmstMCB3LTYgaC02IHJvdW5kZWQtZnVsbCBib3JkZXItMiBmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWNlbnRlciB0cmFuc2l0aW9uLWNvbG9ycyBib3JkZXItZ3JheS0zMDAgZGFyazpib3JkZXItZ3JheS02MDAgaG92ZXI6Ym9yZGVyLXByaW1hcnktNTAwIGhvdmVyOmJnLXByaW1hcnktNTAgZGFyazpob3ZlcjpiZy1wcmltYXJ5LTk1MC8zMCI+CiAgICAgIDwvYnV0dG9uPmAgOiBgCiAgICAgIDxkaXYgY2xhc3M9ImZsZXgtc2hyaW5rLTAgdy02IGgtNiByb3VuZGVkLWZ1bGwgYmctZ3JlZW4tNTAwIHRleHQtd2hpdGUgZmxleCBpdGVtcy1jZW50ZXIganVzdGlmeS1jZW50ZXIiPgogICAgICAgICR7SUNPTlMuY2hlY2sucmVwbGFjZSgnaWNvbicsICdpY29uLXNtJyl9CiAgICAgIDwvZGl2PmB9CgogICAgICA8ZGl2IGNsYXNzPSJmbGV4LTEgbWluLXctMCBjdXJzb3ItcG9pbnRlciIgZGF0YS1hY3Rpb249ImVkaXQiIGRhdGEtaWQ9IiR7dGFzay5pZH0iPgogICAgICAgIDxkaXYgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGZsZXgtd3JhcCBnYXAtMS41IG1iLTAuNSI+CiAgICAgICAgICAke3Rhc2suY2F0ZWdvcnk9PT0nam9zaHVhJyA/ICc8c3BhbiBjbGFzcz0iaW5saW5lLWZsZXggaXRlbXMtY2VudGVyIGdhcC0xIHB4LTIgcHktMC41IHJvdW5kZWQtZnVsbCB0ZXh0LXhzIGZvbnQtYm9sZCBiZy1wdXJwbGUtNjAwIHRleHQtd2hpdGUiPkpvc2h1YTwvc3Bhbj4nIDogJyd9CiAgICAgICAgICAke3Rhc2suaXNVcmdlbnQgPyBgPHNwYW4gY2xhc3M9ImlubGluZS1mbGV4IGl0ZW1zLWNlbnRlciBnYXAtMSBweC0yIHB5LTAuNSByb3VuZGVkLWZ1bGwgdGV4dC14cyBmb250LWJvbGQgYmctcmVkLTYwMCB0ZXh0LXdoaXRlIj4ke0lDT05TLmFsZXJ0LnJlcGxhY2UoJ2ljb24nLCdpY29uLXNtJyl9IFVyZ2VudDwvc3Bhbj5gIDogJyd9CiAgICAgICAgICAke3Rhc2suc3RhdHVzICE9PSAncGVuZGluZycgPyBgPHNwYW4gY2xhc3M9ImlubGluZS1mbGV4IGl0ZW1zLWNlbnRlciBnYXAtMSBweC0yIHB5LTAuNSByb3VuZGVkLWZ1bGwgdGV4dC14cyBmb250LXNlbWlib2xkICR7c3RhdHVzSW5mby5jbGFzc05hbWV9Ij4ke3N0YXR1c0luZm8uaWNvbi5yZXBsYWNlKCdpY29uJywnaWNvbi1zbScpfSAke3N0YXR1c0luZm8ubGFiZWx9PC9zcGFuPmAgOiAnJ30KICAgICAgICAgICR7b3ZlcmR1ZSA/IGA8c3BhbiBjbGFzcz0iaW5saW5lLWZsZXggaXRlbXMtY2VudGVyIGdhcC0xIHB4LTIgcHktMC41IHJvdW5kZWQtZnVsbCB0ZXh0LXhzIGZvbnQtYm9sZCBiZy1yZWQtNjAwIHRleHQtd2hpdGUiPiR7SUNPTlMuY2xvY2sucmVwbGFjZSgnaWNvbicsJ2ljb24tc20nKX0gT3ZlcmR1ZTwvc3Bhbj5gIDogJyd9CiAgICAgICAgICAke2R1ZVRvZGF5ID8gYDxzcGFuIGNsYXNzPSJpbmxpbmUtZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTEgcHgtMiBweS0wLjUgcm91bmRlZC1mdWxsIHRleHQteHMgZm9udC1ib2xkIGJnLW9yYW5nZS01MDAgdGV4dC13aGl0ZSI+JHtJQ09OUy5jbG9jay5yZXBsYWNlKCdpY29uJywnaWNvbi1zbScpfSBEdWUgVG9kYXk8L3NwYW4+YCA6ICcnfQogICAgICAgIDwvZGl2PgoKICAgICAgICA8aDMgY2xhc3M9InRleHQtc20gc206dGV4dC1iYXNlIGZvbnQtYm9sZCBsZWFkaW5nLXNudWcgYnJlYWstd29yZHMgd2hpdGVzcGFjZS1ub3JtYWwgJHtpc0NvbXBsZXRlZCA/ICdsaW5lLXRocm91Z2ggdGV4dC1ncmF5LTQwMCcgOiAnJ30iPiR7ZXNjYXBlSHRtbCh0YXNrLnRhc2tOYW1lKX08L2gzPgoKICAgICAgICAke3Rhc2suZGVzY3JpcHRpb24gPyBgPHAgY2xhc3M9InRleHQteHMgc206dGV4dC1zbSB0ZXh0LWdyYXktNjAwIGRhcms6dGV4dC1ncmF5LTMwMCBtdC0wLjUgYnJlYWstd29yZHMgd2hpdGVzcGFjZS1wcmUtd3JhcCAke2lzQ29tcGxldGVkID8gJ2xpbmUtdGhyb3VnaCBvcGFjaXR5LTcwJyA6ICcnfSI+JHtlc2NhcGVIdG1sKHRhc2suZGVzY3JpcHRpb24pfTwvcD5gIDogJyd9CgogICAgICAgIDxkaXYgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGZsZXgtd3JhcCBnYXAteC0yLjUgZ2FwLXktMSBtdC0xIHRleHQteHMgdGV4dC1ncmF5LTQwMCI+CiAgICAgICAgICAke3Rhc2suZHVlRGF0ZSA/IGAKICAgICAgICAgICAgPHNwYW4gY2xhc3M9ImlubGluZS1mbGV4IGl0ZW1zLWNlbnRlciBnYXAtMSBweC0yIHB5LTAuNSByb3VuZGVkLWZ1bGwgZm9udC1zZW1pYm9sZCAkewogICAgICAgICAgICAgIG92ZXJkdWUKICAgICAgICAgICAgICAgID8gJ2JnLXJlZC02MDAgdGV4dC13aGl0ZScKICAgICAgICAgICAgICAgIDogJ2JnLXJlZC0xMDAgdGV4dC1yZWQtNzAwIGRhcms6YmctcmVkLTk1MC81MCBkYXJrOnRleHQtcmVkLTQwMCcKICAgICAgICAgICAgfSI+JHtJQ09OUy5jbG9jay5yZXBsYWNlKCdpY29uJywnaWNvbi1zbScpfSBEdWUgJHtmb3JtYXREYXRlKHRhc2suZHVlRGF0ZSArICdUMDA6MDA6MDAnKX08L3NwYW4+CiAgICAgICAgICBgIDogJyd9CiAgICAgICAgICA8c3Bhbj5DcmVhdGVkICR7Zm9ybWF0RGF0ZSh0YXNrLmNyZWF0ZWRBdCl9PC9zcGFuPgogICAgICAgICAgJHtpc0FyY2hpdmVkID8gYDxzcGFuIGNsYXNzPSJ0ZXh0LWdyYXktNDAwIj4mbWlkZG90OyBBcmNoaXZlZCAke2Zvcm1hdERhdGUodGFzay5hcmNoaXZlZEF0KX08L3NwYW4+YCA6ICcnfQogICAgICAgICAgJHtpc0FyY2hpdmVkICYmIHB1cmdlRGF5cyAhPT0gbnVsbCA/IGA8c3BhbiBjbGFzcz0idGV4dC1hbWJlci01MDAgZGFyazp0ZXh0LWFtYmVyLTQwMCBmb250LW1lZGl1bSI+Jm1pZGRvdDsgQXV0by1kZWxldGVzIGluICR7cHVyZ2VEYXlzfSBkYXkke3B1cmdlRGF5cyA9PT0gMSA/ICcnIDogJ3MnfTwvc3Bhbj5gIDogJyd9CiAgICAgICAgPC9kaXY+CiAgICAgIDwvZGl2PgoKICAgICAgPGRpdiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTEuNSBzbTpnYXAtMiBmbGV4LXNocmluay0wIG1sLWF1dG8gbXQtMiBzbTptdC0wIj4KICAgICAgICAke2NvbXBsZXRlQnRufQogICAgICAgIDxidXR0b24gZGF0YS1hY3Rpb249InRvZ2dsZS11cmdlbnQiIGRhdGEtaWQ9IiR7dGFzay5pZH0iIHRpdGxlPSIke3Rhc2suaXNVcmdlbnQgPyAnUmVtb3ZlIHVyZ2VudCBmbGFnJyA6ICdNYXJrIHVyZ2VudCd9IgogICAgICAgICAgY2xhc3M9InctOCBoLTggZmxleCBpdGVtcy1jZW50ZXIganVzdGlmeS1jZW50ZXIgcm91bmRlZC1sZyB0cmFuc2l0aW9uLWNvbG9ycyAkewogICAgICAgICAgICB0YXNrLmlzVXJnZW50ID8gJ2JnLXJlZC0xMDAgdGV4dC1yZWQtNjAwIGRhcms6YmctcmVkLTk1MC81MCBkYXJrOnRleHQtcmVkLTQwMCcgOiAnYmctZ3JheS0xMDAgZGFyazpiZy1ncmF5LTgwMCB0ZXh0LWdyYXktNDAwIGhvdmVyOnRleHQtcmVkLTUwMCcKICAgICAgICAgIH0iPiR7SUNPTlMuYWxlcnQucmVwbGFjZSgnaWNvbicsJ2ljb24tc20nKX08L2J1dHRvbj4KICAgICAgICA8YnV0dG9uIGRhdGEtYWN0aW9uPSJlZGl0IiBkYXRhLWlkPSIke3Rhc2suaWR9IiB0aXRsZT0iRWRpdCB0YXNrIgogICAgICAgICAgY2xhc3M9InctOCBoLTggZmxleCBpdGVtcy1jZW50ZXIganVzdGlmeS1jZW50ZXIgcm91bmRlZC1sZyBiZy1ncmF5LTEwMCBkYXJrOmJnLWdyYXktODAwIHRleHQtZ3JheS01MDAgaG92ZXI6dGV4dC1wcmltYXJ5LTYwMCB0cmFuc2l0aW9uLWNvbG9ycyI+JHtJQ09OUy5wZW5jaWwucmVwbGFjZSgnaWNvbicsJ2ljb24tc20nKX08L2J1dHRvbj4KICAgICAgICAkeyFpc0FyY2hpdmVkID8gYDxidXR0b24gZGF0YS1hY3Rpb249ImFyY2hpdmUiIGRhdGEtaWQ9IiR7dGFzay5pZH0iIHRpdGxlPSJBcmNoaXZlIHRhc2siCiAgICAgICAgICBjbGFzcz0idy04IGgtOCBmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWNlbnRlciByb3VuZGVkLWxnIGJnLWdyYXktMTAwIGRhcms6YmctZ3JheS04MDAgdGV4dC1ncmF5LTUwMCBob3Zlcjp0ZXh0LWFtYmVyLTYwMCB0cmFuc2l0aW9uLWNvbG9ycyI+JHtJQ09OUy5iYWNrdXAucmVwbGFjZSgnaWNvbicsJ2ljb24tc20nKX08L2J1dHRvbj5gIDogJyd9CiAgICAgICAgPGJ1dHRvbiBkYXRhLWFjdGlvbj0iZGVsZXRlIiBkYXRhLWlkPSIke3Rhc2suaWR9IiB0aXRsZT0iRGVsZXRlIHRhc2siCiAgICAgICAgICBjbGFzcz0idy04IGgtOCBmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWNlbnRlciByb3VuZGVkLWxnIGJnLWdyYXktMTAwIGRhcms6YmctZ3JheS04MDAgdGV4dC1ncmF5LTUwMCBob3Zlcjp0ZXh0LXJlZC02MDAgdHJhbnNpdGlvbi1jb2xvcnMiPiR7SUNPTlMudHJhc2gucmVwbGFjZSgnaWNvbicsJ2ljb24tc20nKX08L2J1dHRvbj4KICAgICAgPC9kaXY+CiAgICA8L2Rpdj4KICBgOwp9CgpmdW5jdGlvbiByZW5kZXJUYXNrTGlzdCgpIHsKICBjb25zdCBjb250YWluZXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndGFzay1saXN0LWNvbnRhaW5lcicpOwogIGNvbnN0IHRhc2tzID0gZ2V0RmlsdGVyZWRTb3J0ZWRUYXNrcygpOwoKICBpZiAodGFza3MubGVuZ3RoID09PSAwKSB7CiAgICBjb25zdCBlbXB0eU1lc3NhZ2VzID0gewogICAgICBhbGw6IHsgdGl0bGU6ICdObyB0YXNrcyB5ZXQnLCBkZXNjOiAnQWRkIGEgdGFzayBhYm92ZSB0byBnZXQgc3RhcnRlZC4nIH0sCiAgICAgIGpvc2h1YTogeyB0aXRsZTogJ05vIEpvc2h1YSB0YXNrcycsIGRlc2M6ICdUYXNrcyBjYXRlZ29yaXNlZCBhcyBKb3NodWEgd2lsbCBzaG93IGhlcmUuJyB9LAogICAgICB1cmdlbnQ6IHsgdGl0bGU6ICdObyB1cmdlbnQgdGFza3MnLCBkZXNjOiAnTWFyayBhIHRhc2sgdXJnZW50IGFuZCBpdCB3aWxsIHNob3cgdXAgaGVyZS4nIH0sCiAgICAgIG92ZXJkdWU6IHsgdGl0bGU6ICdOb3RoaW5nIG92ZXJkdWUnLCBkZXNjOiAnVGFza3MgcGFzdCB0aGVpciBkdWUgZGF0ZSB3aWxsIHNob3cgdXAgaGVyZS4nIH0sCiAgICAgIGFyY2hpdmU6IHsgdGl0bGU6ICdBcmNoaXZlIGlzIGVtcHR5JywgZGVzYzogJ0NvbXBsZXRlZCB0YXNrcyB3aWxsIGFwcGVhciBoZXJlLCB0aGVuIGF1dG8tZGVsZXRlIGFmdGVyIDMwIGRheXMuJyB9LAogICAgfTsKICAgIGNvbnN0IG1zZyA9IGVtcHR5TWVzc2FnZXNbc3RhdGUudmlld10gfHwgZW1wdHlNZXNzYWdlcy5hbGw7CiAgICBjb250YWluZXIuaW5uZXJIVE1MID0gYAogICAgICA8ZGl2IGNsYXNzPSJmbGV4IGZsZXgtY29sIGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWNlbnRlciBweS0xNiB0ZXh0LWNlbnRlciB0ZXh0LWdyYXktNDAwIGJnLXdoaXRlIGRhcms6YmctZ3JheS05MDAgYm9yZGVyIGJvcmRlci1kYXNoZWQgYm9yZGVyLWdyYXktMjAwIGRhcms6Ym9yZGVyLWdyYXktODAwIHJvdW5kZWQtMnhsIj4KICAgICAgICA8ZGl2IGNsYXNzPSJtYi0zIG9wYWNpdHktNTAiIHN0eWxlPSJ3aWR0aDo0OHB4O2hlaWdodDo0OHB4Ij4ke0lDT05TLmxpc3R9PC9kaXY+CiAgICAgICAgPHAgY2xhc3M9InRleHQtbGcgZm9udC1zZW1pYm9sZCI+JHttc2cudGl0bGV9PC9wPgogICAgICAgIDxwIGNsYXNzPSJ0ZXh0LXNtIG10LTEiPiR7bXNnLmRlc2N9PC9wPgogICAgICA8L2Rpdj4KICAgIGA7CiAgICByZXR1cm47CiAgfQoKICBjb250YWluZXIuaW5uZXJIVE1MID0gYDxkaXYgY2xhc3M9InNwYWNlLXktMiI+JHt0YXNrcy5tYXAodGFza0NhcmRIdG1sKS5qb2luKCcnKX08L2Rpdj5gOwp9Cjwvc2NyaXB0Pgo8c2NyaXB0PgovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KLy8gVGFzayBlZGl0IG1vZGFsICh3aXRoIGF1dG8tc2F2ZSwgbm90ZXMsIGNvbW1lbnRzKQovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KY29uc3QgU1RBVFVTX09QVElPTlMgPSBbCiAgeyB2YWx1ZTogJ3BlbmRpbmcnLCBsYWJlbDogJ1BlbmRpbmcnIH0sCiAgeyB2YWx1ZTogJ2luX3Byb2dyZXNzJywgbGFiZWw6ICdJbiBQcm9ncmVzcycgfSwKICB7IHZhbHVlOiAnY29tcGxldGVkJywgbGFiZWw6ICdDb21wbGV0ZWQnIH0sCl07CgpmdW5jdGlvbiBzZXRTYXZlSW5kaWNhdG9yKHN0YXRlTmFtZSkgewogIGNvbnN0IGVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NhdmUtaW5kaWNhdG9yJyk7CiAgaWYgKCFlbCkgcmV0dXJuOwogIGlmIChzdGF0ZU5hbWUgPT09ICdzYXZpbmcnKSB7CiAgICBlbC5pbm5lckhUTUwgPSBgPHNwYW4gY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGdhcC0xLjUiPiR7SUNPTlMubG9hZGVyLnJlcGxhY2UoJ2ljb24nLCdpY29uLXNtJyl9IFNhdmluZ+KApjwvc3Bhbj5gOwogIH0gZWxzZSBpZiAoc3RhdGVOYW1lID09PSAnc2F2ZWQnKSB7CiAgICBlbC5pbm5lckhUTUwgPSBgPHNwYW4gY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGdhcC0xLjUgdGV4dC1ncmVlbi01MDAiPiR7SUNPTlMuY2hlY2sucmVwbGFjZSgnaWNvbicsJ2ljb24tc20nKX0gU2F2ZWQ8L3NwYW4+YDsKICAgIHNldFRpbWVvdXQoKCkgPT4geyBpZiAoZWwpIGVsLnRleHRDb250ZW50ID0gJ0F1dG8tc2F2ZSBlbmFibGVkJzsgfSwgMTUwMCk7CiAgfSBlbHNlIHsKICAgIGVsLnRleHRDb250ZW50ID0gJ0F1dG8tc2F2ZSBlbmFibGVkJzsKICB9Cn0KCmxldCBwZW5kaW5nU2F2ZSA9IG51bGw7IC8vIHsgaWQsIHVwZGF0ZXMgfQpsZXQgcGVuZGluZ1NhdmVUaW1lciA9IG51bGw7CgpmdW5jdGlvbiBxdWV1ZUF1dG9TYXZlKGlkLCB1cGRhdGVzKSB7CiAgcGVuZGluZ1NhdmUgPSB7IGlkLCB1cGRhdGVzIH07CiAgaWYgKHBlbmRpbmdTYXZlVGltZXIpIGNsZWFyVGltZW91dChwZW5kaW5nU2F2ZVRpbWVyKTsKICBwZW5kaW5nU2F2ZVRpbWVyID0gc2V0VGltZW91dChmbHVzaEF1dG9TYXZlLCA2MDApOwp9CgpmdW5jdGlvbiBmbHVzaEF1dG9TYXZlKCkgewogIGlmIChwZW5kaW5nU2F2ZVRpbWVyKSB7IGNsZWFyVGltZW91dChwZW5kaW5nU2F2ZVRpbWVyKTsgcGVuZGluZ1NhdmVUaW1lciA9IG51bGw7IH0KICBpZiAoIXBlbmRpbmdTYXZlKSByZXR1cm47CiAgY29uc3QgeyBpZCwgdXBkYXRlcyB9ID0gcGVuZGluZ1NhdmU7CiAgcGVuZGluZ1NhdmUgPSBudWxsOwogIHVwZGF0ZVRhc2soaWQsIHVwZGF0ZXMpOwogIHNldFNhdmVJbmRpY2F0b3IoJ3NhdmVkJyk7CiAgcmVuZGVyU3RhdHMoKTsKfQoKZnVuY3Rpb24gb3BlblRhc2tNb2RhbCh0YXNrSWQpIHsKICBzdGF0ZS5lZGl0aW5nVGFza0lkID0gdGFza0lkOwogIHJlbmRlck1vZGFsKCk7Cn0KCmZ1bmN0aW9uIGNsb3NlVGFza01vZGFsKCkgewogIGZsdXNoQXV0b1NhdmUoKTsKICBzdGF0ZS5lZGl0aW5nVGFza0lkID0gbnVsbDsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndGFzay1tb2RhbC1jb250YWluZXInKS5pbm5lckhUTUwgPSAnJzsKICByZW5kZXJUYXNrTGlzdCgpOwogIHJlbmRlclN0YXRzKCk7Cn0KCmZ1bmN0aW9uIHJlbmRlck1vZGFsKCkgewogIGNvbnN0IHRhc2sgPSBnZXRUYXNrQnlJZChzdGF0ZS5lZGl0aW5nVGFza0lkKTsKICBjb25zdCBjb250YWluZXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndGFzay1tb2RhbC1jb250YWluZXInKTsKICBpZiAoIXRhc2spIHsgY29udGFpbmVyLmlubmVySFRNTCA9ICcnOyByZXR1cm47IH0KCiAgY29udGFpbmVyLmlubmVySFRNTCA9IGAKICAgIDxkaXYgaWQ9Im1vZGFsLWJhY2tkcm9wIiBjbGFzcz0iZml4ZWQgaW5zZXQtMCB6LTUwIGZsZXggaXRlbXMtY2VudGVyIGp1c3RpZnktY2VudGVyIHAtNCBiZy1ibGFjay81MCBiYWNrZHJvcC1ibHVyLXNtIGFuaW0tZmFkZSI+CiAgICAgIDxkaXYgaWQ9Im1vZGFsLWJveCIgY2xhc3M9ImJnLXdoaXRlIGRhcms6YmctZ3JheS05MDAgcm91bmRlZC0yeGwgc2hhZG93LTJ4bCB3LWZ1bGwgbWF4LXctMnhsIG1heC1oLVs5MHZoXSBvdmVyZmxvdy15LWF1dG8gYW5pbS1zbGlkZSI+CiAgICAgICAgPGRpdiBjbGFzcz0ic3RpY2t5IHRvcC0wIGJnLXdoaXRlIGRhcms6YmctZ3JheS05MDAgYm9yZGVyLWIgYm9yZGVyLWdyYXktMTAwIGRhcms6Ym9yZGVyLWdyYXktODAwIHB4LTUgc206cHgtNiBweS00IGZsZXggaXRlbXMtY2VudGVyIGp1c3RpZnktYmV0d2VlbiB6LTEwIj4KICAgICAgICAgIDxkaXYgaWQ9InNhdmUtaW5kaWNhdG9yIiBjbGFzcz0idGV4dC1zbSBmb250LW1lZGl1bSB0ZXh0LWdyYXktNDAwIj5BdXRvLXNhdmUgZW5hYmxlZDwvZGl2PgogICAgICAgICAgPGJ1dHRvbiBpZD0ibW9kYWwtY2xvc2UiIGNsYXNzPSJ3LTEwIGgtMTAgZmxleCBpdGVtcy1jZW50ZXIganVzdGlmeS1jZW50ZXIgcm91bmRlZC14bCBob3ZlcjpiZy1ncmF5LTEwMCBkYXJrOmhvdmVyOmJnLWdyYXktODAwIj4ke0lDT05TLnh9PC9idXR0b24+CiAgICAgICAgPC9kaXY+CgogICAgICAgIDxkaXYgY2xhc3M9InAtNSBzbTpwLTYgc3BhY2UteS01Ij4KICAgICAgICAgIDxkaXY+CiAgICAgICAgICAgIDxsYWJlbCBjbGFzcz0iYmxvY2sgdGV4dC1zbSBmb250LXNlbWlib2xkIG1iLTEuNSI+VGFzayBOYW1lPC9sYWJlbD4KICAgICAgICAgICAgPGlucHV0IGlkPSJtb2RhbC1uYW1lIiB0eXBlPSJ0ZXh0IiB2YWx1ZT0iJHtlc2NhcGVIdG1sKHRhc2sudGFza05hbWUpfSIKICAgICAgICAgICAgICBjbGFzcz0idy1mdWxsIHRleHQteGwgZm9udC1ib2xkIHB4LTQgcHktMyByb3VuZGVkLXhsIGJvcmRlciBib3JkZXItZ3JheS0yMDAgZGFyazpib3JkZXItZ3JheS03MDAgYmctZ3JheS01MCBkYXJrOmJnLWdyYXktODAwIGZvY3VzOmJnLXdoaXRlIGRhcms6Zm9jdXM6YmctZ3JheS05MDAgb3V0bGluZS1ub25lIGZvY3VzOnJpbmctMiBmb2N1czpyaW5nLXByaW1hcnktNTAwIiAvPgogICAgICAgICAgPC9kaXY+CgogICAgICAgICAgPGRpdj4KICAgICAgICAgICAgPGxhYmVsIGNsYXNzPSJibG9jayB0ZXh0LXNtIGZvbnQtc2VtaWJvbGQgbWItMS41Ij5EZXNjcmlwdGlvbjwvbGFiZWw+CiAgICAgICAgICAgIDx0ZXh0YXJlYSBpZD0ibW9kYWwtZGVzY3JpcHRpb24iIHJvd3M9IjMiIHBsYWNlaG9sZGVyPSJBZGQgbW9yZSBkZXRhaWxzIGFib3V0IHRoaXMgdGFzay4uLiIKICAgICAgICAgICAgICBjbGFzcz0idy1mdWxsIHB4LTQgcHktMyByb3VuZGVkLXhsIGJvcmRlciBib3JkZXItZ3JheS0yMDAgZGFyazpib3JkZXItZ3JheS03MDAgYmctZ3JheS01MCBkYXJrOmJnLWdyYXktODAwIGZvY3VzOmJnLXdoaXRlIGRhcms6Zm9jdXM6YmctZ3JheS05MDAgcmVzaXplLW5vbmUgb3V0bGluZS1ub25lIGZvY3VzOnJpbmctMiBmb2N1czpyaW5nLXByaW1hcnktNTAwIj4ke2VzY2FwZUh0bWwodGFzay5kZXNjcmlwdGlvbil9PC90ZXh0YXJlYT4KICAgICAgICAgIDwvZGl2PgoKICAgICAgICAgIDxkaXYgY2xhc3M9ImZsZXggZmxleC1jb2wgc206ZmxleC1yb3cgZ2FwLTQiPgogICAgICAgICAgICA8ZGl2IGNsYXNzPSJmbGV4LTEiPgogICAgICAgICAgICAgIDxsYWJlbCBjbGFzcz0iYmxvY2sgdGV4dC1zbSBmb250LXNlbWlib2xkIG1iLTEuNSI+U3RhdHVzPC9sYWJlbD4KICAgICAgICAgICAgICA8ZGl2IGNsYXNzPSJmbGV4IGdhcC0yIj4KICAgICAgICAgICAgICAgICR7U1RBVFVTX09QVElPTlMubWFwKChvcHQpID0+IGAKICAgICAgICAgICAgICAgICAgPGJ1dHRvbiBkYXRhLXN0YXR1cz0iJHtvcHQudmFsdWV9IiBjbGFzcz0ibW9kYWwtc3RhdHVzLWJ0biBmbGV4LTEgcHgtMyBweS0zIHJvdW5kZWQteGwgdGV4dC1zbSBmb250LXNlbWlib2xkIHRyYW5zaXRpb24tY29sb3JzICR7CiAgICAgICAgICAgICAgICAgICAgdGFzay5zdGF0dXMgPT09IG9wdC52YWx1ZQogICAgICAgICAgICAgICAgICAgICAgPyAnYmctcHJpbWFyeS02MDAgdGV4dC13aGl0ZSBzaGFkb3ctbWQgc2hhZG93LXByaW1hcnktNjAwLzMwJwogICAgICAgICAgICAgICAgICAgICAgOiAnYmctZ3JheS0xMDAgZGFyazpiZy1ncmF5LTgwMCB0ZXh0LWdyYXktNjAwIGRhcms6dGV4dC1ncmF5LTMwMCBob3ZlcjpiZy1ncmF5LTIwMCBkYXJrOmhvdmVyOmJnLWdyYXktNzAwJwogICAgICAgICAgICAgICAgICB9Ij4ke29wdC5sYWJlbH08L2J1dHRvbj4KICAgICAgICAgICAgICAgIGApLmpvaW4oJycpfQogICAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPGRpdj4KICAgICAgICAgICAgICA8bGFiZWwgY2xhc3M9ImJsb2NrIHRleHQtc20gZm9udC1zZW1pYm9sZCBtYi0xLjUiPlByaW9yaXR5PC9sYWJlbD4KICAgICAgICAgICAgICA8YnV0dG9uIGlkPSJtb2RhbC11cmdlbnQtYnRuIiBjbGFzcz0iaC1bNDZweF0gcHgtNCByb3VuZGVkLXhsIHRleHQtc20gZm9udC1zZW1pYm9sZCB0cmFuc2l0aW9uLWNvbG9ycyBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiAkewogICAgICAgICAgICAgICAgdGFzay5pc1VyZ2VudCA/ICdiZy1yZWQtNjAwIHRleHQtd2hpdGUgc2hhZG93LW1kIHNoYWRvdy1yZWQtNjAwLzMwJyA6ICdiZy1ncmF5LTEwMCBkYXJrOmJnLWdyYXktODAwIHRleHQtZ3JheS02MDAgZGFyazp0ZXh0LWdyYXktMzAwIGhvdmVyOmJnLWdyYXktMjAwIGRhcms6aG92ZXI6YmctZ3JheS03MDAnCiAgICAgICAgICAgICAgfSI+JHtJQ09OUy5hbGVydH0gJHt0YXNrLmlzVXJnZW50ID8gJ1VyZ2VudCcgOiAnTWFyayBVcmdlbnQnfTwvYnV0dG9uPgogICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPGRpdj4KICAgICAgICAgICAgICA8bGFiZWwgY2xhc3M9ImJsb2NrIHRleHQtc20gZm9udC1zZW1pYm9sZCBtYi0xLjUiPkNhdGVnb3J5PC9sYWJlbD4KICAgICAgICAgICAgICA8YnV0dG9uIGlkPSJtb2RhbC1qb3NodWEtYnRuIiBjbGFzcz0iaC1bNDZweF0gcHgtNCByb3VuZGVkLXhsIHRleHQtc20gZm9udC1zZW1pYm9sZCB0cmFuc2l0aW9uLWNvbG9ycyBmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiAkewogICAgICAgICAgICAgICAgdGFzay5jYXRlZ29yeT09PSdqb3NodWEnID8gJ2JnLXB1cnBsZS02MDAgdGV4dC13aGl0ZSBzaGFkb3ctbWQgc2hhZG93LXB1cnBsZS02MDAvMzAnIDogJ2JnLWdyYXktMTAwIGRhcms6YmctZ3JheS04MDAgdGV4dC1ncmF5LTYwMCBkYXJrOnRleHQtZ3JheS0zMDAgaG92ZXI6YmctZ3JheS0yMDAgZGFyazpob3ZlcjpiZy1ncmF5LTcwMCcKICAgICAgICAgICAgICB9Ij4ke3Rhc2suY2F0ZWdvcnk9PT0nam9zaHVhJyA/ICdKb3NodWEg4pyTJyA6ICdTZXQgSm9zaHVhJ308L2J1dHRvbj4KICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgICA8L2Rpdj4KCiAgICAgICAgICA8ZGl2PgogICAgICAgICAgICA8bGFiZWwgY2xhc3M9ImJsb2NrIHRleHQtc20gZm9udC1zZW1pYm9sZCBtYi0xLjUiPkR1ZSBEYXRlPC9sYWJlbD4KICAgICAgICAgICAgPGlucHV0IGlkPSJtb2RhbC1kdWUtZGF0ZSIgdHlwZT0iZGF0ZSIgdmFsdWU9IiR7ZXNjYXBlSHRtbCh0YXNrLmR1ZURhdGUgfHwgJycpfSIKICAgICAgICAgICAgICBjbGFzcz0idy1mdWxsIHB4LTQgcHktMyByb3VuZGVkLXhsIGJvcmRlciBib3JkZXItZ3JheS0yMDAgZGFyazpib3JkZXItZ3JheS03MDAgYmctZ3JheS01MCBkYXJrOmJnLWdyYXktODAwIGZvY3VzOmJnLXdoaXRlIGRhcms6Zm9jdXM6YmctZ3JheS05MDAgb3V0bGluZS1ub25lIGZvY3VzOnJpbmctMiBmb2N1czpyaW5nLXByaW1hcnktNTAwIiAvPgogICAgICAgICAgPC9kaXY+CgogICAgICAgICAgPGRpdj4KICAgICAgICAgICAgPGxhYmVsIGNsYXNzPSJibG9jayB0ZXh0LXNtIGZvbnQtc2VtaWJvbGQgbWItMS41Ij5Ob3RlczwvbGFiZWw+CiAgICAgICAgICAgIDx0ZXh0YXJlYSBpZD0ibW9kYWwtbm90ZXMiIHJvd3M9IjMiIHBsYWNlaG9sZGVyPSJQZXJzb25hbCBub3RlcyBhYm91dCB0aGlzIHRhc2suLi4iCiAgICAgICAgICAgICAgY2xhc3M9InctZnVsbCBweC00IHB5LTMgcm91bmRlZC14bCBib3JkZXIgYm9yZGVyLWdyYXktMjAwIGRhcms6Ym9yZGVyLWdyYXktNzAwIGJnLWdyYXktNTAgZGFyazpiZy1ncmF5LTgwMCBmb2N1czpiZy13aGl0ZSBkYXJrOmZvY3VzOmJnLWdyYXktOTAwIHJlc2l6ZS1ub25lIG91dGxpbmUtbm9uZSBmb2N1czpyaW5nLTIgZm9jdXM6cmluZy1wcmltYXJ5LTUwMCI+JHtlc2NhcGVIdG1sKHRhc2subm90ZXMpfTwvdGV4dGFyZWE+CiAgICAgICAgICA8L2Rpdj4KCiAgICAgICAgICA8ZGl2IGNsYXNzPSJncmlkIGdyaWQtY29scy0zIGdhcC0zIHRleHQtc20gYmctZ3JheS01MCBkYXJrOmJnLWdyYXktODAwLzYwIHJvdW5kZWQteGwgcC0zIHNtOnAtNCI+CiAgICAgICAgICAgIDxkaXY+CiAgICAgICAgICAgICAgPHAgY2xhc3M9InRleHQtZ3JheS00MDAgdGV4dC14cyB1cHBlcmNhc2UgZm9udC1zZW1pYm9sZCBtYi0wLjUiPkNyZWF0ZWQ8L3A+CiAgICAgICAgICAgICAgPHAgY2xhc3M9ImZvbnQtbWVkaXVtIj4ke2Zvcm1hdERhdGVUaW1lKHRhc2suY3JlYXRlZEF0KX08L3A+CiAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgICA8ZGl2PgogICAgICAgICAgICAgIDxwIGNsYXNzPSJ0ZXh0LWdyYXktNDAwIHRleHQteHMgdXBwZXJjYXNlIGZvbnQtc2VtaWJvbGQgbWItMC41Ij5MYXN0IE1vZGlmaWVkPC9wPgogICAgICAgICAgICAgIDxwIGNsYXNzPSJmb250LW1lZGl1bSI+JHtmb3JtYXREYXRlVGltZSh0YXNrLnVwZGF0ZWRBdCl9PC9wPgogICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgPGRpdj4KICAgICAgICAgICAgICA8cCBjbGFzcz0idGV4dC1ncmF5LTQwMCB0ZXh0LXhzIHVwcGVyY2FzZSBmb250LXNlbWlib2xkIG1iLTAuNSI+JHt0YXNrLmFyY2hpdmVkQXQgPyAnQXJjaGl2ZWQnIDogJ0NvbXBsZXRlZCd9PC9wPgogICAgICAgICAgICAgIDxwIGNsYXNzPSJmb250LW1lZGl1bSI+JHt0YXNrLmFyY2hpdmVkQXQgPyBmb3JtYXREYXRlVGltZSh0YXNrLmFyY2hpdmVkQXQpIDogKHRhc2suY29tcGxldGVkQXQgPyBmb3JtYXREYXRlVGltZSh0YXNrLmNvbXBsZXRlZEF0KSA6ICfigJQnKX08L3A+CiAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgPC9kaXY+CgogICAgICAgICAgPGRpdj4KICAgICAgICAgICAgPGxhYmVsIGNsYXNzPSJmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiB0ZXh0LXNtIGZvbnQtc2VtaWJvbGQgbWItMiI+JHtJQ09OUy5tZXNzYWdlLnJlcGxhY2UoJ2ljb24nLCdpY29uLXNtJyl9IENvbW1lbnRzICgke3Rhc2suY29tbWVudHMubGVuZ3RofSk8L2xhYmVsPgogICAgICAgICAgICA8ZGl2IGlkPSJtb2RhbC1jb21tZW50cy1saXN0IiBjbGFzcz0ic3BhY2UteS0yIG1heC1oLTQ4IG92ZXJmbG93LXktYXV0byBtYi0zIHByLTEiPgogICAgICAgICAgICAgICR7dGFzay5jb21tZW50cy5sZW5ndGggPT09IDAKICAgICAgICAgICAgICAgID8gYDxwIGNsYXNzPSJ0ZXh0LXNtIHRleHQtZ3JheS00MDAgaXRhbGljIj5ObyBjb21tZW50cyB5ZXQuPC9wPmAKICAgICAgICAgICAgICAgIDogdGFzay5jb21tZW50cy5tYXAoKGMpID0+IGAKICAgICAgICAgICAgICAgICAgPGRpdiBjbGFzcz0iYmctZ3JheS01MCBkYXJrOmJnLWdyYXktODAwLzYwIHJvdW5kZWQteGwgcHgtMy41IHB5LTIuNSBmbGV4IGl0ZW1zLXN0YXJ0IGp1c3RpZnktYmV0d2VlbiBnYXAtMiI+CiAgICAgICAgICAgICAgICAgICAgPGRpdj4KICAgICAgICAgICAgICAgICAgICAgIDxwIGNsYXNzPSJ0ZXh0LXNtIj4ke2VzY2FwZUh0bWwoYy5jb250ZW50KX08L3A+CiAgICAgICAgICAgICAgICAgICAgICA8cCBjbGFzcz0idGV4dC14cyB0ZXh0LWdyYXktNDAwIG10LTEiPiR7Zm9ybWF0RGF0ZVRpbWUoYy5jcmVhdGVkQXQpfTwvcD4KICAgICAgICAgICAgICAgICAgICA8L2Rpdj4KICAgICAgICAgICAgICAgICAgICA8YnV0dG9uIGRhdGEtYWN0aW9uPSJkZWxldGUtY29tbWVudCIgZGF0YS1jb21tZW50LWlkPSIke2MuaWR9IiBjbGFzcz0idGV4dC1ncmF5LTMwMCBob3Zlcjp0ZXh0LXJlZC01MDAgZmxleC1zaHJpbmstMCI+JHtJQ09OUy54LnJlcGxhY2UoJ2ljb24nLCdpY29uLXNtJyl9PC9idXR0b24+CiAgICAgICAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgICAgICAgYCkuam9pbignJyl9CiAgICAgICAgICAgIDwvZGl2PgogICAgICAgICAgICA8ZGl2IGNsYXNzPSJmbGV4IGdhcC0yIj4KICAgICAgICAgICAgICA8aW5wdXQgaWQ9Im1vZGFsLW5ldy1jb21tZW50IiB0eXBlPSJ0ZXh0IiBwbGFjZWhvbGRlcj0iQWRkIGEgY29tbWVudC4uLiIKICAgICAgICAgICAgICAgIGNsYXNzPSJmbGV4LTEgcHgtNCBweS0yLjUgcm91bmRlZC14bCBib3JkZXIgYm9yZGVyLWdyYXktMjAwIGRhcms6Ym9yZGVyLWdyYXktNzAwIGJnLWdyYXktNTAgZGFyazpiZy1ncmF5LTgwMCBmb2N1czpiZy13aGl0ZSBkYXJrOmZvY3VzOmJnLWdyYXktOTAwIG91dGxpbmUtbm9uZSBmb2N1czpyaW5nLTIgZm9jdXM6cmluZy1wcmltYXJ5LTUwMCIgLz4KICAgICAgICAgICAgICA8YnV0dG9uIGlkPSJtb2RhbC1hZGQtY29tbWVudCIgY2xhc3M9InctMTEgaC0xMSBmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWNlbnRlciByb3VuZGVkLXhsIGJnLXByaW1hcnktNjAwIHRleHQtd2hpdGUiPiR7SUNPTlMuc2VuZH08L2J1dHRvbj4KICAgICAgICAgICAgPC9kaXY+CiAgICAgICAgICA8L2Rpdj4KICAgICAgICA8L2Rpdj4KCiAgICAgICAgPGRpdiBjbGFzcz0ic3RpY2t5IGJvdHRvbS0wIGJnLXdoaXRlIGRhcms6YmctZ3JheS05MDAgYm9yZGVyLXQgYm9yZGVyLWdyYXktMTAwIGRhcms6Ym9yZGVyLWdyYXktODAwIHB4LTUgc206cHgtNiBweS00IGZsZXggaXRlbXMtY2VudGVyIGp1c3RpZnktYmV0d2VlbiI+CiAgICAgICAgICA8YnV0dG9uIGlkPSJtb2RhbC1kZWxldGUiIGNsYXNzPSJmbGV4IGl0ZW1zLWNlbnRlciBnYXAtMiBweC00IHB5LTIuNSByb3VuZGVkLXhsIHRleHQtcmVkLTYwMCBob3ZlcjpiZy1yZWQtNTAgZGFyazpob3ZlcjpiZy1yZWQtOTUwLzQwIGZvbnQtc2VtaWJvbGQiPiR7SUNPTlMudHJhc2h9IERlbGV0ZTwvYnV0dG9uPgogICAgICAgICAgPGJ1dHRvbiBpZD0ibW9kYWwtZG9uZSIgY2xhc3M9ImZsZXggaXRlbXMtY2VudGVyIGdhcC0yIHB4LTYgcHktMi41IHJvdW5kZWQteGwgYmctcHJpbWFyeS02MDAgaG92ZXI6YmctcHJpbWFyeS03MDAgdGV4dC13aGl0ZSBmb250LXNlbWlib2xkIHNoYWRvdy1tZCBzaGFkb3ctcHJpbWFyeS02MDAvMzAiPiR7SUNPTlMuY2hlY2t9IERvbmU8L2J1dHRvbj4KICAgICAgICA8L2Rpdj4KICAgICAgPC9kaXY+CiAgICA8L2Rpdj4KICBgOwoKICB3aXJlTW9kYWxFdmVudHModGFzayk7Cn0KCmZ1bmN0aW9uIHdpcmVNb2RhbEV2ZW50cyh0YXNrKSB7CiAgY29uc3QgbmFtZUlucHV0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ21vZGFsLW5hbWUnKTsKICBjb25zdCBkZXNjSW5wdXQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbW9kYWwtZGVzY3JpcHRpb24nKTsKICBjb25zdCBub3Rlc0lucHV0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ21vZGFsLW5vdGVzJyk7CiAgY29uc3QgZHVlRGF0ZUlucHV0ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ21vZGFsLWR1ZS1kYXRlJyk7CgogIGNvbnN0IHRyaWdnZXJTYXZlID0gKCkgPT4gewogICAgc2V0U2F2ZUluZGljYXRvcignc2F2aW5nJyk7CiAgICBxdWV1ZUF1dG9TYXZlKHRhc2suaWQsIHsKICAgICAgdGFza05hbWU6IG5hbWVJbnB1dC52YWx1ZSwKICAgICAgZGVzY3JpcHRpb246IGRlc2NJbnB1dC52YWx1ZSwKICAgICAgbm90ZXM6IG5vdGVzSW5wdXQudmFsdWUsCiAgICAgIGR1ZURhdGU6IGR1ZURhdGVJbnB1dC52YWx1ZSwKICAgIH0pOwogIH07CgogIG5hbWVJbnB1dC5hZGRFdmVudExpc3RlbmVyKCdpbnB1dCcsIHRyaWdnZXJTYXZlKTsKICBkZXNjSW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignaW5wdXQnLCB0cmlnZ2VyU2F2ZSk7CiAgbm90ZXNJbnB1dC5hZGRFdmVudExpc3RlbmVyKCdpbnB1dCcsIHRyaWdnZXJTYXZlKTsKICBkdWVEYXRlSW5wdXQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgdHJpZ2dlclNhdmUpOwoKICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcubW9kYWwtc3RhdHVzLWJ0bicpLmZvckVhY2goKGJ0bikgPT4gewogICAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICBmbHVzaEF1dG9TYXZlKCk7CiAgICAgIHVwZGF0ZVRhc2sodGFzay5pZCwgeyBzdGF0dXM6IGJ0bi5kYXRhc2V0LnN0YXR1cyB9KTsKICAgICAgcmVuZGVyTW9kYWwoKTsKICAgICAgcmVuZGVyU3RhdHMoKTsKICAgIH0pOwogIH0pOwoKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbW9kYWwtdXJnZW50LWJ0bicpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgZmx1c2hBdXRvU2F2ZSgpOwogICAgdXBkYXRlVGFzayh0YXNrLmlkLCB7IGlzVXJnZW50OiAhdGFzay5pc1VyZ2VudCB9KTsKICAgIHJlbmRlck1vZGFsKCk7CiAgICByZW5kZXJTdGF0cygpOwogIH0pOwoKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbW9kYWwtam9zaHVhLWJ0bicpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgZmx1c2hBdXRvU2F2ZSgpOwogICAgdXBkYXRlVGFzayh0YXNrLmlkLCB7IGNhdGVnb3J5OiB0YXNrLmNhdGVnb3J5ID09PSAnam9zaHVhJyA/ICcnIDogJ2pvc2h1YScgfSk7CiAgICByZW5kZXJNb2RhbCgpOwogICAgcmVuZGVyU3RhdHMoKTsKICAgIHJlbmRlclRhc2tMaXN0KCk7CiAgfSk7CgogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdtb2RhbC1hZGQtY29tbWVudCcpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgY29uc3QgaW5wdXQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbW9kYWwtbmV3LWNvbW1lbnQnKTsKICAgIGlmICghaW5wdXQudmFsdWUudHJpbSgpKSByZXR1cm47CiAgICBmbHVzaEF1dG9TYXZlKCk7CiAgICBhZGRDb21tZW50KHRhc2suaWQsIGlucHV0LnZhbHVlKTsKICAgIHJlbmRlck1vZGFsKCk7CiAgfSk7CgogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdtb2RhbC1uZXctY29tbWVudCcpLmFkZEV2ZW50TGlzdGVuZXIoJ2tleWRvd24nLCAoZSkgPT4gewogICAgaWYgKGUua2V5ID09PSAnRW50ZXInKSB7CiAgICAgIGUucHJldmVudERlZmF1bHQoKTsKICAgICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ21vZGFsLWFkZC1jb21tZW50JykuY2xpY2soKTsKICAgIH0KICB9KTsKCiAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnW2RhdGEtYWN0aW9uPSJkZWxldGUtY29tbWVudCJdJykuZm9yRWFjaCgoYnRuKSA9PiB7CiAgICBidG4uYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoKSA9PiB7CiAgICAgIGZsdXNoQXV0b1NhdmUoKTsKICAgICAgZGVsZXRlQ29tbWVudCh0YXNrLmlkLCBidG4uZGF0YXNldC5jb21tZW50SWQpOwogICAgICByZW5kZXJNb2RhbCgpOwogICAgfSk7CiAgfSk7CgogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdtb2RhbC1kZWxldGUnKS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgIGlmICghd2luZG93LmNvbmZpcm0oYERlbGV0ZSB0YXNrICIke3Rhc2sudGFza05hbWV9Ij8gVGhpcyBjYW5ub3QgYmUgdW5kb25lLmApKSByZXR1cm47CiAgICBmbHVzaEF1dG9TYXZlKCk7CiAgICBkZWxldGVUYXNrKHRhc2suaWQpOwogICAgY2xvc2VUYXNrTW9kYWwoKTsKICB9KTsKCiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ21vZGFsLWNsb3NlJykuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBjbG9zZVRhc2tNb2RhbCk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ21vZGFsLWRvbmUnKS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGNsb3NlVGFza01vZGFsKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbW9kYWwtYmFja2Ryb3AnKS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIChlKSA9PiB7CiAgICBpZiAoZS50YXJnZXQuaWQgPT09ICdtb2RhbC1iYWNrZHJvcCcpIGNsb3NlVGFza01vZGFsKCk7CiAgfSk7Cn0KPC9zY3JpcHQ+CjxzY3JpcHQ+Ci8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQovLyBUb2FzdCBtZXNzYWdlcwovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KZnVuY3Rpb24gc2hvd1RvYXN0KHR5cGUsIHRleHQpIHsKICBjb25zdCBjb250YWluZXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgndG9hc3QtY29udGFpbmVyJyk7CiAgaWYgKCFjb250YWluZXIpIHJldHVybjsKICBjb25zdCBjb2xvcnMgPSB0eXBlID09PSAnc3VjY2VzcycKICAgID8gJ2JnLWdyZWVuLTUwIGRhcms6YmctZ3JlZW4tOTUwLzMwIHRleHQtZ3JlZW4tNzAwIGRhcms6dGV4dC1ncmVlbi00MDAnCiAgICA6ICdiZy1yZWQtNTAgZGFyazpiZy1yZWQtOTUwLzMwIHRleHQtcmVkLTcwMCBkYXJrOnRleHQtcmVkLTQwMCc7CiAgY29uc3QgaWNvbiA9IHR5cGUgPT09ICdzdWNjZXNzJyA/IElDT05TLmNoZWNrIDogSUNPTlMueDsKICBjb250YWluZXIuaW5uZXJIVE1MID0gYAogICAgPGRpdiBjbGFzcz0iZmxleCBpdGVtcy1zdGFydCBnYXAtMiB0ZXh0LXNtIHJvdW5kZWQteGwgcHgtMy41IHB5LTIuNSBhbmltLWZhZGUgJHtjb2xvcnN9Ij4KICAgICAgJHtpY29uLnJlcGxhY2UoJ2ljb24nLCAnaWNvbi1zbScpfTxzcGFuPiR7ZXNjYXBlSHRtbCh0ZXh0KX08L3NwYW4+CiAgICA8L2Rpdj4KICBgOwogIHNldFRpbWVvdXQoKCkgPT4geyBpZiAoY29udGFpbmVyLmZpcnN0RWxlbWVudENoaWxkKSBjb250YWluZXIuaW5uZXJIVE1MID0gJyc7IH0sIDUwMDApOwp9CgpmdW5jdGlvbiBzaG93VW5kb1RvYXN0KHRleHQsIG9uVW5kbykgewogIGNvbnN0IGNvbnRhaW5lciA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0b2FzdC1jb250YWluZXInKTsKICBpZiAoIWNvbnRhaW5lcikgcmV0dXJuOwogIGNvbnRhaW5lci5pbm5lckhUTUwgPSBgCiAgICA8ZGl2IGNsYXNzPSJmbGV4IGl0ZW1zLWNlbnRlciBqdXN0aWZ5LWJldHdlZW4gZ2FwLTMgdGV4dC1zbSByb3VuZGVkLXhsIHB4LTMuNSBweS0yLjUgYW5pbS1mYWRlIGJnLWdyZWVuLTUwIGRhcms6YmctZ3JlZW4tOTUwLzMwIHRleHQtZ3JlZW4tNzAwIGRhcms6dGV4dC1ncmVlbi00MDAiPgogICAgICA8c3BhbiBjbGFzcz0iZmxleCBpdGVtcy1jZW50ZXIgZ2FwLTIiPiR7SUNPTlMuY2hlY2sucmVwbGFjZSgnaWNvbicsICdpY29uLXNtJyl9JHtlc2NhcGVIdG1sKHRleHQpfTwvc3Bhbj4KICAgICAgPGJ1dHRvbiBpZD0idG9hc3QtdW5kby1idG4iIGNsYXNzPSJmb250LWJvbGQgdW5kZXJsaW5lIGZsZXgtc2hyaW5rLTAiPlVuZG88L2J1dHRvbj4KICAgIDwvZGl2PgogIGA7CiAgY29uc3QgdW5kb0J0biA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0b2FzdC11bmRvLWJ0bicpOwogIGlmICh1bmRvQnRuKSB7CiAgICB1bmRvQnRuLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gewogICAgICBjb250YWluZXIuaW5uZXJIVE1MID0gJyc7CiAgICAgIG9uVW5kbygpOwogICAgfSk7CiAgfQogIHNldFRpbWVvdXQoKCkgPT4geyBpZiAoY29udGFpbmVyLmZpcnN0RWxlbWVudENoaWxkKSBjb250YWluZXIuaW5uZXJIVE1MID0gJyc7IH0sIDYwMDApOwp9CgovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KLy8gSW1wb3J0OiBFeGNlbCAvIENTVgovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KZnVuY3Rpb24gaW1wb3J0RXhjZWxGaWxlKGZpbGUpIHsKICBjb25zdCByZWFkZXIgPSBuZXcgRmlsZVJlYWRlcigpOwogIHJlYWRlci5vbmxvYWQgPSAoZSkgPT4gewogICAgdHJ5IHsKICAgICAgY29uc3QgZGF0YSA9IG5ldyBVaW50OEFycmF5KGUudGFyZ2V0LnJlc3VsdCk7CiAgICAgIGNvbnN0IHdvcmtib29rID0gWExTWC5yZWFkKGRhdGEsIHsgdHlwZTogJ2FycmF5JyB9KTsKICAgICAgY29uc3Qgc2hlZXROYW1lID0gd29ya2Jvb2suU2hlZXROYW1lc1swXTsKICAgICAgY29uc3Qgc2hlZXQgPSB3b3JrYm9vay5TaGVldHNbc2hlZXROYW1lXTsKICAgICAgY29uc3Qgcm93cyA9IFhMU1gudXRpbHMuc2hlZXRfdG9fanNvbihzaGVldCwgeyBkZWZ2YWw6ICcnIH0pOwoKICAgICAgaWYgKCFyb3dzLmxlbmd0aCkgewogICAgICAgIHNob3dUb2FzdCgnZXJyb3InLCAnVGhlIGZpbGUgZG9lcyBub3QgY29udGFpbiBhbnkgcm93cy4nKTsKICAgICAgICByZXR1cm47CiAgICAgIH0KCiAgICAgIGxldCBpbXBvcnRlZCA9IDA7CiAgICAgIHJvd3MuZm9yRWFjaCgocm93KSA9PiB7CiAgICAgICAgY29uc3Qga2V5cyA9IE9iamVjdC5rZXlzKHJvdyk7CiAgICAgICAgY29uc3QgbmFtZUtleSA9IGtleXMuZmluZCgoaykgPT4gL15ccyp0YXNrXHMqbmFtZVxzKiQvaS50ZXN0KGspKSB8fCBrZXlzLmZpbmQoKGspID0+IC9eXHMqbmFtZVxzKiQvaS50ZXN0KGspKSB8fCBrZXlzLmZpbmQoKGspID0+IC9eXHMqdGFza1xzKiQvaS50ZXN0KGspKTsKICAgICAgICBjb25zdCBkZXNjS2V5ID0ga2V5cy5maW5kKChrKSA9PiAvXlxzKmRlc2NyaXB0aW9uXHMqJC9pLnRlc3QoaykpOwogICAgICAgIGNvbnN0IHVyZ2VudEtleSA9IGtleXMuZmluZCgoaykgPT4gL15ccyp1cmdlbnRccyokL2kudGVzdChrKSk7CgogICAgICAgIGNvbnN0IHRhc2tOYW1lID0gbmFtZUtleSA/IFN0cmluZyhyb3dbbmFtZUtleV0pLnRyaW0oKSA6ICcnOwogICAgICAgIGlmICghdGFza05hbWUpIHJldHVybjsKCiAgICAgICAgY29uc3QgZGVzY3JpcHRpb24gPSBkZXNjS2V5ID8gU3RyaW5nKHJvd1tkZXNjS2V5XSkudHJpbSgpIDogJyc7CiAgICAgICAgY29uc3QgaXNVcmdlbnQgPSB1cmdlbnRLZXkgPyBpc1RydXRoeVZhbHVlKHJvd1t1cmdlbnRLZXldKSA6IGZhbHNlOwoKICAgICAgICBjcmVhdGVUYXNrKHsgdGFza05hbWUsIGRlc2NyaXB0aW9uLCBpc1VyZ2VudCwgc3RhdHVzOiAncGVuZGluZycgfSk7CiAgICAgICAgaW1wb3J0ZWQrKzsKICAgICAgfSk7CgogICAgICBzaG93VG9hc3QoJ3N1Y2Nlc3MnLCBgSW1wb3J0ZWQgJHtpbXBvcnRlZH0gdGFzayR7aW1wb3J0ZWQgPT09IDEgPyAnJyA6ICdzJ30gc3VjY2Vzc2Z1bGx5LmApOwogICAgICByZW5kZXJTdGF0cygpOwogICAgICByZW5kZXJUYXNrTGlzdCgpOwogICAgfSBjYXRjaCAoZXJyKSB7CiAgICAgIHNob3dUb2FzdCgnZXJyb3InLCAnQ291bGQgbm90IHJlYWQgdGhlIHVwbG9hZGVkIGZpbGUuIE1ha2Ugc3VyZSBpdCBpcyBhIHZhbGlkIEV4Y2VsIG9yIENTViBmaWxlLicpOwogICAgfQogIH07CiAgcmVhZGVyLnJlYWRBc0FycmF5QnVmZmVyKGZpbGUpOwp9CgovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KLy8gSW1wb3J0OiBXb3JkICguZG9jeCkKLy8gLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tCmZ1bmN0aW9uIGltcG9ydFdvcmRGaWxlKGZpbGUpIHsKICBjb25zdCByZWFkZXIgPSBuZXcgRmlsZVJlYWRlcigpOwogIHJlYWRlci5vbmxvYWQgPSAoZSkgPT4gewogICAgbWFtbW90aC5leHRyYWN0UmF3VGV4dCh7IGFycmF5QnVmZmVyOiBlLnRhcmdldC5yZXN1bHQgfSkKICAgICAgLnRoZW4oKHJlc3VsdCkgPT4gewogICAgICAgIGNvbnN0IGxpbmVzID0gcmVzdWx0LnZhbHVlCiAgICAgICAgICAuc3BsaXQoJ1xuJykKICAgICAgICAgIC5tYXAoKGxpbmUpID0+IGxpbmUucmVwbGFjZSgvXltcc+KAolwtKlx1MjAyMlx1MjVDRlx1MjVBQV0rLywgJycpLnRyaW0oKSkKICAgICAgICAgIC5maWx0ZXIoKGxpbmUpID0+IGxpbmUubGVuZ3RoID4gMCk7CgogICAgICAgIGlmICghbGluZXMubGVuZ3RoKSB7CiAgICAgICAgICBzaG93VG9hc3QoJ2Vycm9yJywgJ05vIHRleHQgY29udGVudCB3YXMgZm91bmQgaW4gdGhlIGRvY3VtZW50LicpOwogICAgICAgICAgcmV0dXJuOwogICAgICAgIH0KCiAgICAgICAgbGluZXMuZm9yRWFjaCgobGluZSkgPT4gY3JlYXRlVGFzayh7IHRhc2tOYW1lOiBsaW5lLCBzdGF0dXM6ICdwZW5kaW5nJyB9KSk7CiAgICAgICAgc2hvd1RvYXN0KCdzdWNjZXNzJywgYEltcG9ydGVkICR7bGluZXMubGVuZ3RofSB0YXNrJHtsaW5lcy5sZW5ndGggPT09IDEgPyAnJyA6ICdzJ30gc3VjY2Vzc2Z1bGx5LmApOwogICAgICAgIHJlbmRlclN0YXRzKCk7CiAgICAgICAgcmVuZGVyVGFza0xpc3QoKTsKICAgICAgfSkKICAgICAgLmNhdGNoKCgpID0+IHNob3dUb2FzdCgnZXJyb3InLCAnQ291bGQgbm90IHJlYWQgdGhlIHVwbG9hZGVkIFdvcmQgZG9jdW1lbnQuJykpOwogIH07CiAgcmVhZGVyLnJlYWRBc0FycmF5QnVmZmVyKGZpbGUpOwp9CgpmdW5jdGlvbiBoYW5kbGVJbXBvcnRGaWxlKGZpbGUpIHsKICBpZiAoIWZpbGUpIHJldHVybjsKICBjb25zdCBleHQgPSBmaWxlLm5hbWUuc3BsaXQoJy4nKS5wb3AoKS50b0xvd2VyQ2FzZSgpOwogIGlmIChbJ3hsc3gnLCAneGxzJywgJ2NzdiddLmluY2x1ZGVzKGV4dCkpIHsKICAgIGltcG9ydEV4Y2VsRmlsZShmaWxlKTsKICB9IGVsc2UgaWYgKGV4dCA9PT0gJ2RvY3gnKSB7CiAgICBpbXBvcnRXb3JkRmlsZShmaWxlKTsKICB9IGVsc2UgewogICAgc2hvd1RvYXN0KCdlcnJvcicsICdVbnN1cHBvcnRlZCBmaWxlIHR5cGUuIFVzZSAueGxzeCwgLnhscywgLmNzdiBvciAuZG9jeCcpOwogIH0KfQoKLy8gLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tCi8vIEV4cG9ydDogRXhjZWwgJiBQREYKLy8gLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tCmZ1bmN0aW9uIHN0YXR1c0xhYmVsKHN0YXR1cykgewogIHJldHVybiB7IHBlbmRpbmc6ICdQZW5kaW5nJywgaW5fcHJvZ3Jlc3M6ICdJbiBQcm9ncmVzcycsIGNvbXBsZXRlZDogJ0NvbXBsZXRlZCcgfVtzdGF0dXNdIHx8IHN0YXR1czsKfQoKZnVuY3Rpb24gZXhwb3J0RXhjZWwoKSB7CiAgY29uc3Qgc29ydGVkID0gc3RhdGUudGFza3Muc2xpY2UoKS5zb3J0KChhLCBiKSA9PiB7CiAgICBpZiAoYS5pc1VyZ2VudCAhPT0gYi5pc1VyZ2VudCkgcmV0dXJuIGEuaXNVcmdlbnQgPyAtMSA6IDE7CiAgICBjb25zdCBzYSA9IFNUQVRVU19PUkRFUlthLnN0YXR1c10gPz8gMzsKICAgIGNvbnN0IHNiID0gU1RBVFVTX09SREVSW2Iuc3RhdHVzXSA/PyAzOwogICAgaWYgKHNhICE9PSBzYikgcmV0dXJuIHNhIC0gc2I7CiAgICByZXR1cm4gbmV3IERhdGUoYi5jcmVhdGVkQXQpIC0gbmV3IERhdGUoYS5jcmVhdGVkQXQpOwogIH0pOwoKICBjb25zdCBkYXRhID0gc29ydGVkLm1hcCgodCkgPT4gKHsKICAgICdUYXNrIE5hbWUnOiB0LnRhc2tOYW1lLAogICAgRGVzY3JpcHRpb246IHQuZGVzY3JpcHRpb24sCiAgICBTdGF0dXM6IHN0YXR1c0xhYmVsKHQuc3RhdHVzKSwKICAgIFVyZ2VudDogdC5pc1VyZ2VudCA/ICdZZXMnIDogJ05vJywKICAgICdEdWUgRGF0ZSc6IHQuZHVlRGF0ZSB8fCAnJywKICAgIE5vdGVzOiB0Lm5vdGVzLAogICAgJ0NyZWF0ZWQgRGF0ZSc6IGZvcm1hdERhdGVUaW1lKHQuY3JlYXRlZEF0KSwKICAgICdMYXN0IE1vZGlmaWVkJzogZm9ybWF0RGF0ZVRpbWUodC51cGRhdGVkQXQpLAogICAgJ0NvbXBsZXRpb24gRGF0ZSc6IHQuY29tcGxldGVkQXQgPyBmb3JtYXREYXRlVGltZSh0LmNvbXBsZXRlZEF0KSA6ICcnLAogICAgJ0FyY2hpdmVkIERhdGUnOiB0LmFyY2hpdmVkQXQgPyBmb3JtYXREYXRlVGltZSh0LmFyY2hpdmVkQXQpIDogJycsCiAgfSkpOwoKICBjb25zdCB3b3Jrc2hlZXQgPSBYTFNYLnV0aWxzLmpzb25fdG9fc2hlZXQoZGF0YSk7CiAgd29ya3NoZWV0WychY29scyddID0gW3sgd2NoOiAzMCB9LCB7IHdjaDogNDAgfSwgeyB3Y2g6IDE0IH0sIHsgd2NoOiA4IH0sIHsgd2NoOiAxNCB9LCB7IHdjaDogNDAgfSwgeyB3Y2g6IDIwIH0sIHsgd2NoOiAyMCB9LCB7IHdjaDogMjAgfSwgeyB3Y2g6IDIwIH1dOwoKICBjb25zdCB3b3JrYm9vayA9IFhMU1gudXRpbHMuYm9va19uZXcoKTsKICBYTFNYLnV0aWxzLmJvb2tfYXBwZW5kX3NoZWV0KHdvcmtib29rLCB3b3Jrc2hlZXQsICdUYXNrcycpOwogIFhMU1gud3JpdGVGaWxlKHdvcmtib29rLCAnZHV0eS1saXN0LXRhc2tzLnhsc3gnKTsKfQoKZnVuY3Rpb24gZXhwb3J0UERGKCkgewogIGNvbnN0IHsganNQREYgfSA9IHdpbmRvdy5qc3BkZjsKICBjb25zdCBkb2MgPSBuZXcganNQREYoeyB1bml0OiAncHQnLCBmb3JtYXQ6ICdhNCcgfSk7CiAgY29uc3QgbWFyZ2luID0gNDA7CiAgbGV0IHkgPSBtYXJnaW47CiAgY29uc3QgcGFnZUhlaWdodCA9IGRvYy5pbnRlcm5hbC5wYWdlU2l6ZS5nZXRIZWlnaHQoKTsKICBjb25zdCBwYWdlV2lkdGggPSBkb2MuaW50ZXJuYWwucGFnZVNpemUuZ2V0V2lkdGgoKTsKCiAgY29uc3Qgc29ydGVkID0gc3RhdGUudGFza3Muc2xpY2UoKS5zb3J0KChhLCBiKSA9PiB7CiAgICBpZiAoYS5pc1VyZ2VudCAhPT0gYi5pc1VyZ2VudCkgcmV0dXJuIGEuaXNVcmdlbnQgPyAtMSA6IDE7CiAgICBjb25zdCBzYSA9IFNUQVRVU19PUkRFUlthLnN0YXR1c10gPz8gMzsKICAgIGNvbnN0IHNiID0gU1RBVFVTX09SREVSW2Iuc3RhdHVzXSA/PyAzOwogICAgaWYgKHNhICE9PSBzYikgcmV0dXJuIHNhIC0gc2I7CiAgICByZXR1cm4gbmV3IERhdGUoYi5jcmVhdGVkQXQpIC0gbmV3IERhdGUoYS5jcmVhdGVkQXQpOwogIH0pOwoKICBkb2Muc2V0Rm9udCgnaGVsdmV0aWNhJywgJ2JvbGQnKTsKICBkb2Muc2V0Rm9udFNpemUoMjApOwogIGRvYy50ZXh0KCdEdXR5IExpc3QgLSBUYXNrIFJlcG9ydCcsIHBhZ2VXaWR0aCAvIDIsIHksIHsgYWxpZ246ICdjZW50ZXInIH0pOwogIHkgKz0gMjA7CgogIGRvYy5zZXRGb250KCdoZWx2ZXRpY2EnLCAnbm9ybWFsJyk7CiAgZG9jLnNldEZvbnRTaXplKDEwKTsKICBkb2Muc2V0VGV4dENvbG9yKDEyMCk7CiAgZG9jLnRleHQoYEdlbmVyYXRlZCBvbiAke25ldyBEYXRlKCkudG9Mb2NhbGVTdHJpbmcoKX1gLCBwYWdlV2lkdGggLyAyLCB5LCB7IGFsaWduOiAnY2VudGVyJyB9KTsKICB5ICs9IDI1OwoKICBjb25zdCBzdGF0cyA9IGNvbXB1dGVTdGF0cygpOwogIGRvYy5zZXRUZXh0Q29sb3IoMCk7CiAgZG9jLnNldEZvbnQoJ2hlbHZldGljYScsICdib2xkJyk7CiAgZG9jLnNldEZvbnRTaXplKDExKTsKICBkb2MudGV4dChgVG90YWw6ICR7c3RhdHMudG90YWx9ICAgUGVuZGluZzogJHtzdGF0cy5wZW5kaW5nfSAgIEluIFByb2dyZXNzOiAke3N0YXRzLmluUHJvZ3Jlc3N9ICAgQ29tcGxldGVkOiAke3N0YXRzLmNvbXBsZXRlZH0gICBVcmdlbnQ6ICR7c3RhdHMudXJnZW50fWAsIG1hcmdpbiwgeSk7CiAgeSArPSAyNTsKCiAgY29uc3QgbWF4V2lkdGggPSBwYWdlV2lkdGggLSBtYXJnaW4gKiAyOwoKICBzb3J0ZWQuZm9yRWFjaCgodGFzaywgaW5kZXgpID0+IHsKICAgIGlmICh5ID4gcGFnZUhlaWdodCAtIDgwKSB7IGRvYy5hZGRQYWdlKCk7IHkgPSBtYXJnaW47IH0KCiAgICBkb2Muc2V0Rm9udCgnaGVsdmV0aWNhJywgJ2JvbGQnKTsKICAgIGRvYy5zZXRGb250U2l6ZSgxMyk7CiAgICBkb2Muc2V0VGV4dENvbG9yKDApOwogICAgbGV0IHRpdGxlID0gYCR7aW5kZXggKyAxfS4gJHt0YXNrLnRhc2tOYW1lfWA7CiAgICBpZiAodGFzay5pc1VyZ2VudCkgdGl0bGUgKz0gJyAgIFtVUkdFTlRdJzsKICAgIGNvbnN0IHRpdGxlTGluZXMgPSBkb2Muc3BsaXRUZXh0VG9TaXplKHRpdGxlLCBtYXhXaWR0aCk7CiAgICBkb2MudGV4dCh0aXRsZUxpbmVzLCBtYXJnaW4sIHkpOwogICAgeSArPSB0aXRsZUxpbmVzLmxlbmd0aCAqIDE2OwoKICAgIGRvYy5zZXRGb250KCdoZWx2ZXRpY2EnLCAnbm9ybWFsJyk7CiAgICBkb2Muc2V0Rm9udFNpemUoMTApOwogICAgZG9jLnNldFRleHRDb2xvcig4NSk7CiAgICBkb2MudGV4dChgU3RhdHVzOiAke3N0YXR1c0xhYmVsKHRhc2suc3RhdHVzKX1gLCBtYXJnaW4sIHkpOwogICAgeSArPSAxNDsKCiAgICBpZiAodGFzay5kZXNjcmlwdGlvbikgewogICAgICBkb2Muc2V0VGV4dENvbG9yKDApOwogICAgICBjb25zdCBsaW5lcyA9IGRvYy5zcGxpdFRleHRUb1NpemUoYERlc2NyaXB0aW9uOiAke3Rhc2suZGVzY3JpcHRpb259YCwgbWF4V2lkdGgpOwogICAgICBkb2MudGV4dChsaW5lcywgbWFyZ2luLCB5KTsKICAgICAgeSArPSBsaW5lcy5sZW5ndGggKiAxNDsKICAgIH0KCiAgICBpZiAodGFzay5ub3RlcykgewogICAgICBkb2Muc2V0VGV4dENvbG9yKDApOwogICAgICBjb25zdCBsaW5lcyA9IGRvYy5zcGxpdFRleHRUb1NpemUoYE5vdGVzOiAke3Rhc2subm90ZXN9YCwgbWF4V2lkdGgpOwogICAgICBkb2MudGV4dChsaW5lcywgbWFyZ2luLCB5KTsKICAgICAgeSArPSBsaW5lcy5sZW5ndGggKiAxNDsKICAgIH0KCiAgICBkb2Muc2V0Rm9udFNpemUoOSk7CiAgICBkb2Muc2V0VGV4dENvbG9yKDE0MCk7CiAgICBsZXQgbWV0YSA9IGBDcmVhdGVkOiAke2Zvcm1hdERhdGVUaW1lKHRhc2suY3JlYXRlZEF0KX0gICB8ICAgTGFzdCBNb2RpZmllZDogJHtmb3JtYXREYXRlVGltZSh0YXNrLnVwZGF0ZWRBdCl9YDsKICAgIGlmICh0YXNrLmNvbXBsZXRlZEF0KSBtZXRhICs9IGAgICB8ICAgQ29tcGxldGVkOiAke2Zvcm1hdERhdGVUaW1lKHRhc2suY29tcGxldGVkQXQpfWA7CiAgICBjb25zdCBtZXRhTGluZXMgPSBkb2Muc3BsaXRUZXh0VG9TaXplKG1ldGEsIG1heFdpZHRoKTsKICAgIGRvYy50ZXh0KG1ldGFMaW5lcywgbWFyZ2luLCB5KTsKICAgIHkgKz0gbWV0YUxpbmVzLmxlbmd0aCAqIDEyICsgMTA7CiAgfSk7CgogIGlmIChzb3J0ZWQubGVuZ3RoID09PSAwKSB7CiAgICBkb2Muc2V0Rm9udFNpemUoMTIpOwogICAgZG9jLnNldFRleHRDb2xvcigwKTsKICAgIGRvYy50ZXh0KCdObyB0YXNrcyBmb3VuZC4nLCBwYWdlV2lkdGggLyAyLCB5LCB7IGFsaWduOiAnY2VudGVyJyB9KTsKICB9CgogIGRvYy5zYXZlKCdkdXR5LWxpc3QtdGFza3MucGRmJyk7Cn0KPC9zY3JpcHQ+CjxzY3JpcHQ+Ci8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQovLyBXaXJlIHVwIGRhc2hib2FyZCBldmVudCBsaXN0ZW5lcnMgKGNhbGxlZCBvbmNlIHBlciBkYXNoYm9hcmQgcmVuZGVyKQovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KZnVuY3Rpb24gd2lyZURhc2hib2FyZEV2ZW50cygpIHsKICAvLyBUaGVtZQogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdidG4tdGhlbWUnKS5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgIHRvZ2dsZVRoZW1lKCk7CiAgICByZW5kZXJEYXNoYm9hcmQoKTsKICB9KTsKCiAgLy8gUXVpY2sgYWRkCiAgY29uc3QgcXVpY2tBZGQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncXVpY2stYWRkLWlucHV0Jyk7CiAgY29uc3QgcXVpY2tBZGRVcmdlbnQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncXVpY2stYWRkLXVyZ2VudCcpOwogIGNvbnN0IHF1aWNrQWRkRHVlRGF0ZSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdxdWljay1hZGQtZHVlLWRhdGUnKTsKICBjb25zdCBxdWlja0FkZEpvc2h1YSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdxdWljay1hZGQtam9zaHVhJyk7CiAgcXVpY2tBZGRVcmdlbnQuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgc3RhdGUucXVpY2tBZGRVcmdlbnQgPSBxdWlja0FkZFVyZ2VudC5jaGVja2VkOwogICAgcmVuZGVyUXVpY2tBZGRVcmdlbnRMYWJlbCgpOwogIH0pOwogIHF1aWNrQWRkSm9zaHVhLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgIHN0YXRlLnF1aWNrQWRkSm9zaHVhID0gcXVpY2tBZGRKb3NodWEuY2hlY2tlZDsKICAgIHJlbmRlclF1aWNrQWRkSm9zaHVhTGFiZWwoKTsKICB9KTsKICBxdWlja0FkZC5hZGRFdmVudExpc3RlbmVyKCdrZXlkb3duJywgKGUpID0+IHsKICAgIGlmIChlLmtleSA9PT0gJ0VudGVyJyAmJiAhZS5zaGlmdEtleSkgewogICAgICBlLnByZXZlbnREZWZhdWx0KCk7CiAgICAgIGNvbnN0IHRleHQgPSBxdWlja0FkZC52YWx1ZTsKICAgICAgaWYgKCF0ZXh0LnRyaW0oKSkgcmV0dXJuOwogICAgICBjb25zdCBsaW5lcyA9IHRleHQuc3BsaXQoJ1xuJykuZmlsdGVyKChsKSA9PiBsLnRyaW0oKS5sZW5ndGggPiAwKTsKICAgICAgY29uc3QgaXNKb3NodWEgPSBzdGF0ZS5jdXJyZW50VmlldyA9PT0gJ2pvc2h1YScgfHwgc3RhdGUucXVpY2tBZGRKb3NodWE7CiAgICAgIGxpbmVzLmZvckVhY2goKGxpbmUpID0+IGNyZWF0ZVRhc2soeyB0YXNrTmFtZTogbGluZSwgc3RhdHVzOiAncGVuZGluZycsIGlzVXJnZW50OiBzdGF0ZS5xdWlja0FkZFVyZ2VudCwgZHVlRGF0ZTogcXVpY2tBZGREdWVEYXRlLnZhbHVlLCBjYXRlZ29yeTogaXNKb3NodWEgPyAnam9zaHVhJyA6ICcnIH0pKTsKICAgICAgcXVpY2tBZGQudmFsdWUgPSAnJzsKICAgICAgcmVuZGVyU3RhdHMoKTsKICAgICAgcmVuZGVyVGFza0xpc3QoKTsKICAgIH0KICB9KTsKCiAgLy8gU2VhcmNoCiAgY29uc3Qgc2VhcmNoSW5wdXQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2VhcmNoLWlucHV0Jyk7CiAgY29uc3Qgc2VhcmNoQ2xlYXIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2VhcmNoLWNsZWFyJyk7CiAgY29uc3QgZGVib3VuY2VkU2VhcmNoID0gZGVib3VuY2UoKCkgPT4gcmVuZGVyVGFza0xpc3QoKSwgMjAwKTsKICBzZWFyY2hJbnB1dC5hZGRFdmVudExpc3RlbmVyKCdpbnB1dCcsICgpID0+IHsKICAgIHN0YXRlLnNlYXJjaCA9IHNlYXJjaElucHV0LnZhbHVlOwogICAgc2VhcmNoQ2xlYXIuY2xhc3NMaXN0LnRvZ2dsZSgnaGlkZGVuJywgIXN0YXRlLnNlYXJjaCk7CiAgICBkZWJvdW5jZWRTZWFyY2goKTsKICB9KTsKICBzZWFyY2hDbGVhci5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsICgpID0+IHsKICAgIHN0YXRlLnNlYXJjaCA9ICcnOwogICAgc2VhcmNoSW5wdXQudmFsdWUgPSAnJzsKICAgIHNlYXJjaENsZWFyLmNsYXNzTGlzdC5hZGQoJ2hpZGRlbicpOwogICAgcmVuZGVyVGFza0xpc3QoKTsKICAgIHNlYXJjaElucHV0LmZvY3VzKCk7CiAgfSk7CgogIC8vIFRhc2sgbGlzdCBhY3Rpb25zIChldmVudCBkZWxlZ2F0aW9uKQogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0YXNrLWxpc3QtY29udGFpbmVyJykuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCAoZSkgPT4gewogICAgY29uc3QgYnRuID0gZS50YXJnZXQuY2xvc2VzdCgnW2RhdGEtYWN0aW9uXScpOwogICAgaWYgKCFidG4pIHJldHVybjsKICAgIGNvbnN0IHsgYWN0aW9uLCBpZCB9ID0gYnRuLmRhdGFzZXQ7CiAgICBjb25zdCB0YXNrID0gZ2V0VGFza0J5SWQoaWQpOwogICAgaWYgKCF0YXNrKSByZXR1cm47CgogICAgc3RhdGUubGFzdFRvdWNoZWRUYXNrSWQgPSBpZDsKCiAgICBpZiAoYWN0aW9uID09PSAndG9nZ2xlLWNvbXBsZXRlJykgewogICAgICBjb25zdCB3YXNDb21wbGV0ZWQgPSB0YXNrLnN0YXR1cyA9PT0gJ2NvbXBsZXRlZCc7CiAgICAgIHRvZ2dsZUNvbXBsZXRlKGlkKTsKICAgICAgcmVuZGVyU3RhdHMoKTsKICAgICAgcmVuZGVyVGFza0xpc3QoKTsKICAgICAgaWYgKCF3YXNDb21wbGV0ZWQpIHsKICAgICAgICBzaG93VW5kb1RvYXN0KGAiJHt0YXNrLnRhc2tOYW1lfSIgbWFya2VkIGNvbXBsZXRlIGFuZCBtb3ZlZCB0byBBcmNoaXZlLmAsICgpID0+IHsKICAgICAgICAgIHJlc3RvcmVGcm9tQXJjaGl2ZShpZCk7CiAgICAgICAgICByZW5kZXJTdGF0cygpOwogICAgICAgICAgcmVuZGVyVGFza0xpc3QoKTsKICAgICAgICB9KTsKICAgICAgfQogICAgfSBlbHNlIGlmIChhY3Rpb24gPT09ICdyZXN0b3JlJykgewogICAgICByZXN0b3JlRnJvbUFyY2hpdmUoaWQpOwogICAgICByZW5kZXJTdGF0cygpOwogICAgICByZW5kZXJUYXNrTGlzdCgpOwogICAgICBzaG93VG9hc3QoJ3N1Y2Nlc3MnLCBgIiR7dGFzay50YXNrTmFtZX0iIHJlc3RvcmVkIHRvIFBlbmRpbmcuYCk7CiAgICB9IGVsc2UgaWYgKGFjdGlvbiA9PT0gJ3RvZ2dsZS11cmdlbnQnKSB7CiAgICAgIHRvZ2dsZVVyZ2VudChpZCk7CiAgICAgIHJlbmRlclN0YXRzKCk7CiAgICAgIHJlbmRlclRhc2tMaXN0KCk7CiAgICB9IGVsc2UgaWYgKGFjdGlvbiA9PT0gJ2VkaXQnKSB7CiAgICAgIG9wZW5UYXNrTW9kYWwoaWQpOwogICAgfSBlbHNlIGlmIChhY3Rpb24gPT09ICdkZWxldGUnKSB7CiAgICAgIGlmICh3aW5kb3cuY29uZmlybShgRGVsZXRlIHRhc2sgIiR7dGFzay50YXNrTmFtZX0iPyBUaGlzIGNhbm5vdCBiZSB1bmRvbmUuYCkpIHsKICAgICAgICBkZWxldGVUYXNrKGlkKTsKICAgICAgICByZW5kZXJTdGF0cygpOwogICAgICAgIHJlbmRlclRhc2tMaXN0KCk7CiAgICAgIH0KICAgIH0KICB9KTsKCiAgLy8gRmlsZSBpbXBvcnQ6IGRyb3B6b25lCiAgY29uc3QgZHJvcHpvbmUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZHJvcHpvbmUnKTsKICBjb25zdCBmaWxlSW5wdXQgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnZmlsZS1pbnB1dCcpOwogIGNvbnN0IGRyb3B6b25lVGV4dCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdkcm9wem9uZS10ZXh0Jyk7CgogIGRyb3B6b25lLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgKCkgPT4gZmlsZUlucHV0LmNsaWNrKCkpOwogIGZpbGVJbnB1dC5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICBpZiAoZmlsZUlucHV0LmZpbGVzWzBdKSBoYW5kbGVJbXBvcnRGaWxlKGZpbGVJbnB1dC5maWxlc1swXSk7CiAgICBmaWxlSW5wdXQudmFsdWUgPSAnJzsKICB9KTsKICBkcm9wem9uZS5hZGRFdmVudExpc3RlbmVyKCdkcmFnb3ZlcicsIChlKSA9PiB7CiAgICBlLnByZXZlbnREZWZhdWx0KCk7CiAgICBkcm9wem9uZS5jbGFzc0xpc3QuYWRkKCdib3JkZXItcHJpbWFyeS01MDAnLCAnYmctcHJpbWFyeS01MCcsICdkYXJrOmJnLXByaW1hcnktOTUwLzMwJyk7CiAgfSk7CiAgZHJvcHpvbmUuYWRkRXZlbnRMaXN0ZW5lcignZHJhZ2xlYXZlJywgKCkgPT4gewogICAgZHJvcHpvbmUuY2xhc3NMaXN0LnJlbW92ZSgnYm9yZGVyLXByaW1hcnktNTAwJywgJ2JnLXByaW1hcnktNTAnLCAnZGFyazpiZy1wcmltYXJ5LTk1MC8zMCcpOwogIH0pOwogIGRyb3B6b25lLmFkZEV2ZW50TGlzdGVuZXIoJ2Ryb3AnLCAoZSkgPT4gewogICAgZS5wcmV2ZW50RGVmYXVsdCgpOwogICAgZHJvcHpvbmUuY2xhc3NMaXN0LnJlbW92ZSgnYm9yZGVyLXByaW1hcnktNTAwJywgJ2JnLXByaW1hcnktNTAnLCAnZGFyazpiZy1wcmltYXJ5LTk1MC8zMCcpOwogICAgaWYgKGUuZGF0YVRyYW5zZmVyLmZpbGVzWzBdKSBoYW5kbGVJbXBvcnRGaWxlKGUuZGF0YVRyYW5zZmVyLmZpbGVzWzBdKTsKICB9KTsKCiAgLy8gRXhwb3J0CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2V4cG9ydC1leGNlbCcpLmFkZEV2ZW50TGlzdGVuZXIoJ2NsaWNrJywgZXhwb3J0RXhjZWwpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdleHBvcnQtcGRmJykuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBleHBvcnRQREYpOwp9CgovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KLy8gR2xvYmFsIGtleWJvYXJkIHNob3J0Y3V0cwovLyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0Kd2luZG93LmFkZEV2ZW50TGlzdGVuZXIoJ2tleWRvd24nLCAoZSkgPT4gewogIGlmIChlLmtleSA9PT0gJ0VzY2FwZScpIHsKICAgIGlmIChzdGF0ZS5lZGl0aW5nVGFza0lkICE9PSBudWxsKSBjbG9zZVRhc2tNb2RhbCgpOwogICAgcmV0dXJuOwogIH0KCiAgLy8gQ3RybCtOOiBmb2N1cyBxdWljayBhZGQgKG5ldyB0YXNrKQogIGlmICgoZS5jdHJsS2V5IHx8IGUubWV0YUtleSkgJiYgZS5rZXkudG9Mb3dlckNhc2UoKSA9PT0gJ24nKSB7CiAgICBlLnByZXZlbnREZWZhdWx0KCk7CiAgICBjb25zdCBlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdxdWljay1hZGQtaW5wdXQnKTsKICAgIGlmIChlbCkgZWwuZm9jdXMoKTsKICAgIHJldHVybjsKICB9CgogIC8vIEN0cmwrVTogdG9nZ2xlIHVyZ2VudCBvbiB0aGUgbW9zdCByZWNlbnRseSB0b3VjaGVkIC8gZmlyc3QgdmlzaWJsZSB0YXNrCiAgaWYgKChlLmN0cmxLZXkgfHwgZS5tZXRhS2V5KSAmJiBlLmtleS50b0xvd2VyQ2FzZSgpID09PSAndScpIHsKICAgIGUucHJldmVudERlZmF1bHQoKTsKICAgIGNvbnN0IHRhcmdldElkID0gc3RhdGUubGFzdFRvdWNoZWRUYXNrSWQgJiYgZ2V0VGFza0J5SWQoc3RhdGUubGFzdFRvdWNoZWRUYXNrSWQpCiAgICAgID8gc3RhdGUubGFzdFRvdWNoZWRUYXNrSWQKICAgICAgOiAoZ2V0RmlsdGVyZWRTb3J0ZWRUYXNrcygpWzBdICYmIGdldEZpbHRlcmVkU29ydGVkVGFza3MoKVswXS5pZCk7CiAgICBpZiAodGFyZ2V0SWQpIHsKICAgICAgdG9nZ2xlVXJnZW50KHRhcmdldElkKTsKICAgICAgcmVuZGVyU3RhdHMoKTsKICAgICAgcmVuZGVyVGFza0xpc3QoKTsKICAgICAgY29uc3QgdGFzayA9IGdldFRhc2tCeUlkKHRhcmdldElkKTsKICAgICAgaWYgKHRhc2spIHNob3dUb2FzdCgnc3VjY2VzcycsIGAiJHt0YXNrLnRhc2tOYW1lfSIgdXJnZW50IGZsYWcgJHt0YXNrLmlzVXJnZW50ID8gJ3NldCcgOiAncmVtb3ZlZCd9LmApOwogICAgfQogICAgcmV0dXJuOwogIH0KCiAgY29uc3QgdGFnID0gZG9jdW1lbnQuYWN0aXZlRWxlbWVudCAmJiBkb2N1bWVudC5hY3RpdmVFbGVtZW50LnRhZ05hbWU7CiAgY29uc3QgaXNUeXBpbmcgPSB0YWcgPT09ICdJTlBVVCcgfHwgdGFnID09PSAnVEVYVEFSRUEnOwogIGlmIChpc1R5cGluZykgcmV0dXJuOwoKICBpZiAoZS5rZXkgPT09ICcvJykgewogICAgZS5wcmV2ZW50RGVmYXVsdCgpOwogICAgY29uc3QgZWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2VhcmNoLWlucHV0Jyk7CiAgICBpZiAoZWwpIGVsLmZvY3VzKCk7CiAgfQp9KTsKCi8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQovLyBJbml0Ci8vIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQpkb2N1bWVudC5hZGRFdmVudExpc3RlbmVyKCdET01Db250ZW50TG9hZGVkJywgcmVuZGVyKTsKPC9zY3JpcHQ+CjwvYm9keT4KPC9odG1sPgo=");
  const bytes=Uint8Array.from(html, c=>c.charCodeAt(0));
  const decoded=new TextDecoder("utf-8").decode(bytes);
  frame.srcdoc=decoded;
  personalTasksLoaded=true;
    }
  } else {
    document.getElementById('personal-tasks-sumudu').style.display='none';
    document.getElementById('personal-tasks-officer').style.display='';
    markOverduePersonalTasks();
    renderPersonalTasks();
  }
}

// ═══════════════════════════════════════════════════════
// RECYCLE BIN
// ═══════════════════════════════════════════════════════
function renderRecycle(){
  const tbody=document.getElementById('recycle-tbody');
  if(!DB.recycleBin.length){ tbody.innerHTML=`<tr><td colspan="5"><div class="empty-state"><div class="empty-icon">📦</div><p>No archived tasks</p></div></td></tr>`; return; }
  tbody.innerHTML=DB.recycleBin.map((t,i)=>`
    <tr class="deleted-task-row">
      <td><div class="task-title-cell"><div class="title">${escHtml(t.title)}</div><div class="desc">${escHtml(t.description||'')}</div></div></td>
      <td>${USERS[t.assignedTo]?.name||t.assignedTo}</td>
      <td>${USERS[t.deletedBy]?.name||t.deletedBy}</td>
      <td style="font-weight:500">${t.deletedDate?new Date(t.deletedDate).toLocaleDateString():''}</td>
      <td>
        <div style="display:flex;gap:6px">
          <button class="btn btn-success btn-sm" onclick="restoreTask(${i})">Restore</button>
          <button class="btn btn-danger btn-sm" onclick="permDelete(${i})">Delete</button>
        </div>
      </td>
    </tr>`).join('');
}

async function restoreTask(i){
  const t=DB.recycleBin.splice(i,1)[0];
  const rid=t.id;
  delete t.deletedBy; delete t.deletedDate;
  t.status='Pending';
  DB.tasks.push(t);
  await dbDeleteRecycleBin(rid);
  await dbInsertTask(t);
  showToast('Task restored','success');
  renderRecycle();
}

async function permDelete(i){
  if(!confirm('Permanently delete this task? This cannot be undone.')) return;
  const t=DB.recycleBin.splice(i,1)[0];
  await dbDeleteRecycleBin(t.id);
  showToast('Task permanently deleted','warning');
  renderRecycle();
}

// ═══════════════════════════════════════════════════════
// REPORTS
// ═══════════════════════════════════════════════════════
function renderReports(){
  const isDark=document.documentElement.getAttribute('data-theme')==='dark';
  const textColor=isDark?'#9499b3':'#5c5f77';
  const gridColor=isDark?'#2a2b3d':'#e4e7ef';
  const fontOpts={size:12,weight:'500',family:'Inter'};

  destroyChart('chart-productivity');
  const names=['Dev','Mitiksha'];
  const deTasks=DB.tasks.filter(t=>t.assignedTo==='de');
  const miTasks=DB.tasks.filter(t=>t.assignedTo==='mitiksha');
  charts['chart-productivity']=new Chart(document.getElementById('chart-productivity'),{
    type:'bar',
    data:{ labels:names, datasets:[
      {label:'Assigned',data:[deTasks.length,miTasks.length],backgroundColor:isDark?'#6366f1':'#6366f1cc',borderRadius:6,borderSkipped:false},
      {label:'Completed',data:[deTasks.filter(t=>t.status==='Completed').length,miTasks.filter(t=>t.status==='Completed').length],backgroundColor:isDark?'#22c55e':'#22c55ecc',borderRadius:6,borderSkipped:false}
    ]},
    options:{ responsive:true,maintainAspectRatio:false, plugins:{legend:{labels:{color:textColor,font:fontOpts,usePointStyle:true,pointStyleWidth:8}}}, scales:{x:{ticks:{color:textColor,font:fontOpts},grid:{display:false}},y:{ticks:{color:textColor,font:fontOpts},grid:{color:gridColor},border:{dash:[4,4]}}} }
  });

  destroyChart('chart-ontime');
  const allDone=DB.tasks.filter(t=>t.status==='Completed');
  const onTime=allDone.filter(t=>t.completionDate<=t.deadlineDate).length;
  const late=allDone.length-onTime;
  charts['chart-ontime']=new Chart(document.getElementById('chart-ontime'),{
    type:'doughnut',
    data:{ labels:['On Time','Late'],datasets:[{data:[onTime,late],backgroundColor:['#22c55e','#ef4444'],borderWidth:0,borderRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'70%',plugins:{legend:{labels:{color:textColor,font:fontOpts,usePointStyle:true,pointStyleWidth:8}}}}
  });

  destroyChart('chart-daily');
  const now2=new Date();
  const daysInM=new Date(now2.getFullYear(),now2.getMonth()+1,0).getDate();
  const dayLabels=[];
  const devCompleted=[], devOverdue=[], mitCompleted=[], mitOverdue=[];
  for(let d=1;d<=daysInM;d++){
    const ds=`${now2.getFullYear()}-${String(now2.getMonth()+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    dayLabels.push(d);
    const snap=DB.dailySummary[ds]||{};
    const dSnap=snap['de']||{completed_count:0,overdue_count:0};
    const mSnap=snap['mitiksha']||{completed_count:0,overdue_count:0};
    devCompleted.push(dSnap.completed_count);
    devOverdue.push(dSnap.overdue_count);
    mitCompleted.push(mSnap.completed_count);
    mitOverdue.push(mSnap.overdue_count);
  }
  charts['chart-daily']=new Chart(document.getElementById('chart-daily'),{
    type:'bar',
    data:{ labels:dayLabels, datasets:[
      {label:'Dev — Completed',data:devCompleted,backgroundColor:'#3b82f6',borderRadius:4,borderSkipped:false,stack:'dev'},
      {label:'Dev — Overdue',data:devOverdue,backgroundColor:'#93c5fd',borderRadius:4,borderSkipped:false,stack:'dev'},
      {label:'Mitiksha — Completed',data:mitCompleted,backgroundColor:'#22c55e',borderRadius:4,borderSkipped:false,stack:'mit'},
      {label:'Mitiksha — Overdue',data:mitOverdue,backgroundColor:'#f87171',borderRadius:4,borderSkipped:false,stack:'mit'}
    ]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:textColor,font:fontOpts,usePointStyle:true,pointStyleWidth:8}}},scales:{x:{stacked:true,ticks:{color:textColor,font:fontOpts},grid:{display:false}},y:{stacked:true,ticks:{color:textColor,font:fontOpts,stepSize:1},grid:{color:gridColor},border:{dash:[4,4]}}}}
  });
}

// ═══════════════════════════════════════════════════════
// EXCEL IMPORT/EXPORT
// ═══════════════════════════════════════════════════════
function exportExcel(){
  const tasks=DB.tasks;
  const rows=[['ID','Title','Description','Assigned To','Assigned By','Assigned Date','Deadline','Priority','Status','Completion Date','Link','Recurring']];
  tasks.forEach(t=>rows.push([t.id,t.title,t.description||'',USERS[t.assignedTo]?.name||t.assignedTo,USERS[t.assignedBy]?.name||t.assignedBy,t.assignedDate,t.deadlineDate,t.priority,t.status,t.completionDate||'',t.externalLink||'',t.isRecurring?t.recurrenceType:'No']));
  const ws=XLSX.utils.aoa_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Tasks');
  XLSX.writeFile(wb,'FinanceFlow_Tasks.xlsx');
  showToast('Tasks exported to Excel','success');
}

function importExcel(event){
  const file=event.target.files[0];
  if(!file) return;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'binary'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const data=XLSX.utils.sheet_to_json(ws,{header:1});
      let imported=0;
      data.slice(1).forEach(row=>{
        if(!row[0]) return;
        const assignTo=Object.keys(USERS).find(k=>USERS[k].name.toLowerCase()===String(row[3]||'').toLowerCase())||'de';
        const task={
          id:'task_'+Date.now()+Math.random(),
          title:String(row[0]||'Untitled'),
          description:String(row[1]||''),
          assignedTo:assignTo,
          assignedBy:currentUser,
          assignedDate:row[2]?String(row[2]):today(),
          deadlineDate:row[4]?String(row[4]):today(),
          priority:'Urgent',
          status:'Pending',
          externalLink:'',
          isRecurring:false,
          recurrenceType:'',
          completionDate:null,
          createdDate:new Date().toISOString(),
          updatedDate:new Date().toISOString(),
          steps:[],attachments:[]
        };
        DB.tasks.push(task);
        imported++;
      });
      saveDB();
      showToast(`${imported} tasks imported from Excel`,'success');
      renderTasks(); renderDashboard();
    }catch(err){ showToast('Error reading Excel file','error'); }
  };
  reader.readAsBinaryString(file);
  event.target.value='';
}

// ═══════════════════════════════════════════════════════
// MODAL
// ═══════════════════════════════════════════════════════
function openModal(title,body,btns=[]){
  document.getElementById('modal-title').textContent=title;
  document.getElementById('modal-body').innerHTML=body;
  const footer=document.getElementById('modal-footer');
  footer.innerHTML=btns.map(b=>`<button class="btn ${b.cls}" onclick="${b.action}">${b.label}</button>`).join('');
  document.getElementById('modal-overlay').classList.remove('hidden');
  document.body.style.overflow='hidden';
  pendingFiles=[];
}

function closeModal(e){
  if(e && e.target && e.target!==document.getElementById('modal-overlay')) return;
  document.getElementById('modal-overlay').classList.add('hidden');
  document.body.style.overflow='';
  pendingFiles=[];
}

// ═══════════════════════════════════════════════════════
// TOAST
// ═══════════════════════════════════════════════════════
function showToast(msg,type='info'){
  const c=document.getElementById('toast-container');
  const t=document.createElement('div');
  t.className=`toast ${type}`;
  const icons={success:'✅',error:'❌',warning:'⚠️',info:'ℹ️'};
  t.innerHTML=`<span>${icons[type]||'ℹ️'}</span><span>${msg}</span>`;
  c.appendChild(t);
  setTimeout(()=>{ t.style.animation='toastOut .3s ease forwards'; setTimeout(()=>t.remove(),300); },4000);
}

// ═══════════════════════════════════════════════════════
// DASHBOARD STAT CLICK
// ═══════════════════════════════════════════════════════
function dashStatClick(statusFilter){
  if(currentUser==='sumudu'||currentUser==='de'){
    navigateTo('tasks');
    if(statusFilter){ document.getElementById('filter-status').value=statusFilter; renderTasks(); }
  } else {
    navigateTo('my-tasks');
    if(statusFilter){ document.getElementById('my-filter-status').value=statusFilter; renderMyTasks(); }
  }
}

// ═══════════════════════════════════════════════════════
// EMAILJS SETUP
// ═══════════════════════════════════════════════════════
emailjs.init('cr7Jzs6awXDeCUUNa');
const EMAILJS_SERVICE='service_82awixe';
const EMAILJS_TEMPLATE='template_2wqz252';

function htmlToPlainText(html){
  const tmp = document.createElement('div');
  tmp.innerHTML = html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<th[^>]*>/gi, '')
    .replace(/<\/th>/gi, ': ')
    .replace(/<td[^>]*>/gi, '')
    .replace(/<\/td>/gi, '\n')
    .replace(/<[^>]+>/g, '');
  let text = tmp.textContent || tmp.innerText || '';
  return text.replace(/\n{3,}/g, '\n\n').trim();
}

async function sendRealEmail(to, subject, body){
  try{
    const plainBody = htmlToPlainText(body);
    const res = await emailjs.send(EMAILJS_SERVICE, EMAILJS_TEMPLATE, {
      to_email: to,
      subject: subject,
      message: plainBody,
      body: plainBody,
      from_name: 'FinanceFlow',
      reply_to: 'noreply@financeflow.com'
    });
    console.log('EmailJS success:', res);
    return true;
  }catch(err){
    console.error('EmailJS error:', err);
    showToast('❌ Email error: ' + (err.text || err.message || JSON.stringify(err)), 'error');
    return false;
  }
}

// ═══════════════════════════════════════════════════════
// EMAIL NOTIFICATION SYSTEM
// ═══════════════════════════════════════════════════════
async function simulateEmail(type, task){
  let subject='', to='', body='';
  const assignee=USERS[task.assignedTo];
  const assigner=USERS[task.assignedBy]||USERS[currentUser];

  if(type==='doc_request'){
    to=USERS['sumudu'].email;
    subject=`📄 Document Request from ${USERS[task.assignedBy]?.name}: ${task.title}`;
    body=`<p>Hi <strong>Sumudu</strong>,</p>
      <p><strong>${USERS[task.assignedBy]?.name}</strong> has requested a document.</p>
      <table><tr><th>Request</th><td>${escHtml(task.title)}</td></tr>
      <tr><th>Date</th><td>${fmt(task.assignedDate)}</td></tr></table>
      ${task.description?`<p><strong>Details:</strong> ${escHtml(task.description)}</p>`:''}
      <p style="margin-top:14px;color:#6366f1;font-weight:600">— FinanceFlow Task Management</p>`;
    const email={id:'em_'+Date.now(),type,subject,to,from:USERS[task.assignedBy]?.email||'system@financeflow.com',body,timestamp:new Date().toISOString(),taskId:task.id};
    DB.emailLog.push(email);
    await dbInsertEmail(email);
    showEmailPreview(email);
    const sent=await sendRealEmail(to, subject, body);
    showToast(sent?`📧 Request emailed to Sumudu`:'📧 Request logged (email delivery pending)','success');
    return;
  }

  if(type==='new_task'){
    to=assignee.email;
    subject=`New Task Assigned: ${task.title}`;
    body=`<p>Hi <strong>${assignee.name}</strong>,</p>
      <p>A new task has been assigned to you by <strong>${assigner.name}</strong>.</p>
      <table><tr><th>Task</th><td>${escHtml(task.title)}</td></tr>
      <tr><th>Priority</th><td>${task.priority}${task.urgentFlag?' 🔴':''}</td></tr>
      <tr><th>Assigned</th><td>${fmt(task.assignedDate)}</td></tr>
      <tr><th>Deadline</th><td>${fmt(task.deadlineDate)}</td></tr>
      <tr><th>Status</th><td>${task.status}</td></tr></table>
      ${task.description?`<p><strong>Details:</strong> ${escHtml(task.description)}</p>`:''}
      ${task.externalLink?`<p><strong>Link:</strong> <a href="${sanitizeUrl(task.externalLink)}">${escHtml(task.externalLink)}</a></p>`:''}
      <p style="margin-top:14px;color:#6366f1;font-weight:600">— FinanceFlow Task Management</p>`;
  }

  const email={id:'em_'+Date.now(),type,subject,to,from:assigner?.email||'system@financeflow.com',body,timestamp:new Date().toISOString(),taskId:task.id};
  DB.emailLog.push(email);
  await dbInsertEmail(email);
  showEmailPreview(email);
  const sent=await sendRealEmail(to, subject, body);
  showToast(sent?`📧 Email sent to ${assignee.name}`:`📧 Logged (email delivery pending)`,'success');
}

function showEmailPreview(email){
  const body=`<div class="email-preview">
    <div class="email-header">
      <div class="email-row"><span class="email-label">From</span><span class="email-val">${escHtml(email.from)}</span></div>
      <div class="email-row"><span class="email-label">To</span><span class="email-val">${escHtml(email.to)}</span></div>
      <div class="email-row"><span class="email-label">Subject</span><span class="email-val" style="font-weight:700">${escHtml(email.subject)}</span></div>
      <div class="email-row"><span class="email-label">Date</span><span class="email-val">${new Date(email.timestamp).toLocaleString()}</span></div>
    </div>
    <div class="email-body">${email.body}</div>
  </div>`;
  openModal('📧 Email Preview', body, [{label:'Close',action:'closeModal()',cls:'btn-secondary'}]);
}

async function sendMorningSummary(){
  const employees = currentUser==='sumudu'?['de','mitiksha']:[currentUser];
  for(const uid of employees){
    const u=USERS[uid];
    const tasks=DB.tasks.filter(t=>t.assignedTo===uid && t.status!=='Completed');
    if(!tasks.length) return;
    const subject=`☀️ Good Morning ${u.name} — Your Tasks for Today`;
    const rows=tasks.map(t=>`<tr><td>${escHtml(t.title)}</td><td>${t.priority}${t.urgentFlag?' 🔴':''}</td><td>${t.status}</td><td>${fmt(t.deadlineDate)}</td></tr>`).join('');
    const body=`<p>Hi <strong>${u.name}</strong>,</p>
      <p>Here's your task summary for today. You have <strong>${tasks.length}</strong> pending task(s).</p>
      <table><tr><th>Task</th><th>Priority</th><th>Status</th><th>Deadline</th></tr>${rows}</table>
      <p>Have a productive day! 💪</p>
      <p style="margin-top:14px;color:#6366f1;font-weight:600">— FinanceFlow Task Management</p>`;
    const email={id:'em_'+Date.now()+Math.random(),type:'morning_summary',subject,to:u.email,from:'system@financeflow.com',body,timestamp:new Date().toISOString()};
    DB.emailLog.push(email);
    await dbInsertEmail(email);
    await sendRealEmail(u.email, subject, body);
  }
  showToast('📬 Morning summary emails sent!','success');
  if(document.getElementById('page-emails')?.classList.contains('active')) renderEmailLog();
}

async function sendEODSummary(){
  const employees = currentUser==='sumudu'?['de','mitiksha']:[currentUser];
  for(const uid of employees){
    const u=USERS[uid];
    const allTasks=DB.tasks.filter(t=>t.assignedTo===uid);
    const done=allTasks.filter(t=>t.status==='Completed');
    const pending=allTasks.filter(t=>t.status!=='Completed');
    const overdue=allTasks.filter(t=>t.status==='Overdue');
    const subject=`🌙 End of Day Summary — ${u.name}`;
    const body=`<p>Hi <strong>${u.name}</strong>,</p>
      <p>Here's your end-of-day summary:</p>
      <table><tr><th>Metric</th><th>Count</th></tr>
      <tr><td>Total Tasks</td><td><strong>${allTasks.length}</strong></td></tr>
      <tr><td>Completed</td><td style="color:#22c55e"><strong>${done.length}</strong></td></tr>
      <tr><td>Pending</td><td style="color:#f59e0b"><strong>${pending.length}</strong></td></tr>
      <tr><td>Overdue</td><td style="color:#ef4444"><strong>${overdue.length}</strong></td></tr></table>
      ${overdue.length?`<p style="color:#ef4444"><strong>⚠️ You have ${overdue.length} overdue task(s). Please prioritize them.</strong></p>`:''}
      <p>Great work today! Rest well. 🌟</p>
      <p style="margin-top:14px;color:#6366f1;font-weight:600">— FinanceFlow Task Management</p>`;
    const email={id:'em_'+Date.now()+Math.random(),type:'eod_summary',subject,to:u.email,from:'system@financeflow.com',body,timestamp:new Date().toISOString()};
    DB.emailLog.push(email);
    await dbInsertEmail(email);
    await sendRealEmail(u.email, subject, body);
  }
  showToast('📮 End-of-day summary emails sent!','success');
  if(document.getElementById('page-emails')?.classList.contains('active')) renderEmailLog();
}

function renderEmailLog(){
  const el=document.getElementById('email-log-list');
  if(!el) return;
  const emails=[...(DB.emailLog||[])].reverse();
  if(!emails.length){ el.innerHTML=`<div class="empty-state"><div class="empty-icon">📧</div><p>No emails sent yet</p></div>`; return; }
  el.innerHTML=emails.map(e=>`
    <div class="email-log-item" onclick="viewEmailById('${e.id}')" style="cursor:pointer">
      <div class="email-log-icon">${e.type==='new_task'?'📋':e.type==='morning_summary'||e.type==='morning_report'?'☀️':e.type==='due_warning'?'⚠️':'🌙'}</div>
      <div class="email-log-info">
        <div class="email-log-subject">${escHtml(e.subject)}</div>
        <div class="email-log-to">To: ${escHtml(e.to)}</div>
      </div>
      <div class="email-log-time">${new Date(e.timestamp).toLocaleString()}</div>
    </div>`).join('');
}

function viewEmailById(id){
  const e=(DB.emailLog||[]).find(x=>x.id===id);
  if(e) showEmailPreview(e);
}


let morningReportSent=localStorage.getItem('ffMorningReport_'+today())==='1';
let dueWarningsSent=localStorage.getItem('ffDueWarnings_'+today())==='1';

function checkAutoMorningReport(){
  if(!canWrite()) return; // sending/logging the report is a write; skip for view-only users
  const now=new Date();
  const hour=now.getHours();
  const min=now.getMinutes();
  if(hour===7 && min<15 && !morningReportSent){
    morningReportSent=true;
    localStorage.setItem('ffMorningReport_'+today(),'1');
    sendAutoMorningReport();
  }
  if(hour>=9 && !dueWarningsSent){
    dueWarningsSent=true;
    localStorage.setItem('ffDueWarnings_'+today(),'1');
    sendDueDateWarnings();
  }
}

async function sendAutoMorningReport(){
  for(const uid of ['de','mitiksha']){
    const u=USERS[uid];
    const pending=DB.tasks.filter(t=>t.assignedTo===uid && (t.status==='Pending'||t.status==='In Progress'));
    const due=DB.tasks.filter(t=>t.assignedTo===uid && t.status!=='Completed' && t.deadlineDate && t.deadlineDate<=today());
    if(!pending.length && !due.length) continue;
    const subject=`☀️ Morning Report — ${u.name} (${new Date().toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'})})`;
    let rows='';
    if(due.length){
      rows+=`<tr><td colspan="4" style="background:#fee2e2;color:#dc2626;font-weight:700;padding:8px">⚠️ DUE / OVERDUE TASKS (${due.length})</td></tr>`;
      rows+=due.map(t=>`<tr style="background:#fff5f5"><td>${escHtml(t.title)}</td><td>${t.priority}</td><td style="color:#dc2626;font-weight:600">${t.status}</td><td style="color:#dc2626">${fmt(t.deadlineDate)}</td></tr>`).join('');
    }
    if(pending.length){
      rows+=`<tr><td colspan="4" style="background:#fef9c3;color:#92400e;font-weight:700;padding:8px">📋 PENDING TASKS (${pending.length})</td></tr>`;
      rows+=pending.map(t=>`<tr><td>${escHtml(t.title)}</td><td>${t.priority}</td><td>${t.status}</td><td>${fmt(t.deadlineDate)}</td></tr>`).join('');
    }
    const body=`<p>Good Morning <strong>${u.name}</strong>,</p>
      <p>Here is your daily task report:</p>
      <table><tr><th>Task</th><th>Priority</th><th>Status</th><th>Deadline</th></tr>${rows}</table>
      <p>Have a productive day! 💪</p>
      <p style="margin-top:14px;color:#6366f1;font-weight:600">— FinanceFlow Task Management</p>`;
    const email={id:'em_'+Date.now()+Math.random(),type:'morning_report',subject,to:u.email,from:'system@financeflow.com',body,timestamp:new Date().toISOString()};
    DB.emailLog.push(email);
    await dbInsertEmail(email);
    await sendRealEmail(u.email, subject, body);
  }
}

async function sendDueDateWarnings(){
  const tomorrow=new Date();
  tomorrow.setDate(tomorrow.getDate()+1);
  const tomorrowStr=tomorrow.toISOString().split('T')[0];
  for(const uid of ['de','mitiksha']){
    const u=USERS[uid];
    const approaching=DB.tasks.filter(t=>t.assignedTo===uid && t.status!=='Completed' && t.deadlineDate && (t.deadlineDate===today()||t.deadlineDate===tomorrowStr));
    const overdue=DB.tasks.filter(t=>t.assignedTo===uid && t.status!=='Completed' && t.deadlineDate && t.deadlineDate<today());
    const allWarning=[...overdue,...approaching];
    if(!allWarning.length) continue;
    const subject=`⚠️ Due Date Warning — ${allWarning.length} task(s) need attention`;
    let rows=allWarning.map(t=>{
      const isOD=t.deadlineDate<today();
      return `<tr style="${isOD?'background:#fff5f5':'background:#fffbeb'}"><td>${escHtml(t.title)}</td><td style="color:${isOD?'#dc2626':'#f59e0b'};font-weight:600">${isOD?'OVERDUE':'Due '+fmt(t.deadlineDate)}</td><td>${fmt(t.deadlineDate)}</td></tr>`;
    }).join('');
    const body=`<p>Hi <strong>${u.name}</strong>,</p>
      <p>The following tasks are approaching or have passed their due date:</p>
      <table><tr><th>Task</th><th>Status</th><th>Deadline</th></tr>${rows}</table>
      <p>Please prioritize these tasks.</p>
      <p style="margin-top:14px;color:#6366f1;font-weight:600">— FinanceFlow Task Management</p>`;
    const email={id:'em_'+Date.now()+Math.random(),type:'due_warning',subject,to:u.email,from:'system@financeflow.com',body,timestamp:new Date().toISOString()};
    DB.emailLog.push(email);
    await dbInsertEmail(email);
    await sendRealEmail(u.email, subject, body);
  }
}

// ═══════════════════════════════════════════════════════
// SEARCH
// ═══════════════════════════════════════════════════════
function globalSearch(q){
  if(!q){ navigateTo('tasks'); return; }
  q=q.toLowerCase();
  const results=DB.tasks.filter(t=>
    t.title.toLowerCase().includes(q)||
    (t.description||'').toLowerCase().includes(q)||
    (USERS[t.assignedTo]?.name||'').toLowerCase().includes(q)
  );
  navigateTo('tasks');
  const tbody=document.getElementById('tasks-tbody');
  if(!results.length){ tbody.innerHTML=`<tr><td colspan="8"><div class="empty-state"><div class="empty-icon">\u{1F50D}</div><p>No results for "${escHtml(q)}"</p></div></td></tr>`; return; }
  tbody.innerHTML=results.map(t=>taskRow(t)).join('');
}

// ═══════════════════════════════════════════════════════
// THEME
// ═══════════════════════════════════════════════════════
function toggleTheme(){
  const curr=document.documentElement.getAttribute('data-theme');
  const next=curr==='dark'?'light':'dark';
  document.documentElement.setAttribute('data-theme',next);
  localStorage.setItem('ffTheme',next);
  setTimeout(()=>{ renderDashboard(); },100);
}

// ═══════════════════════════════════════════════════════
// SIDEBAR MOBILE
// ═══════════════════════════════════════════════════════
function openSidebar(){ document.getElementById('sidebar').classList.add('open'); document.getElementById('sidebar-overlay').classList.add('open'); }
function closeSidebar(){ document.getElementById('sidebar').classList.remove('open'); document.getElementById('sidebar-overlay').classList.remove('open'); }

// ═══════════════════════════════════════════════════════
// UTILS
// ═══════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════
// SIDEBAR CALENDAR
// ═══════════════════════════════════════════════════════
let calYear, calMonth;
function initCalendar(){
  const now=new Date();
  calYear=now.getFullYear(); calMonth=now.getMonth();
  renderCalendar();
}
function renderCalendar(){
  const el=document.getElementById('sidebar-calendar');
  if(!el) return;
  const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const days=['Su','Mo','Tu','We','Th','Fr','Sa'];
  const first=new Date(calYear,calMonth,1).getDay();
  const daysInMonth=new Date(calYear,calMonth+1,0).getDate();
  const todayStr=today();
  const calTasks = currentUser==='mitiksha' ? DB.tasks.filter(t=>t.assignedTo==='mitiksha') : DB.tasks;
  const taskDates=new Set(calTasks.map(t=>t.deadlineDate).filter(Boolean));
  let grid=days.map(d=>`<div class="cal-day-name">${d}</div>`).join('');
  for(let i=0;i<first;i++) grid+=`<div class="cal-day empty"></div>`;
  for(let d=1;d<=daysInMonth;d++){
    const dateStr=`${calYear}-${String(calMonth+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const isToday=dateStr===todayStr;
    const hasTask=taskDates.has(dateStr);
    grid+=`<div class="cal-day${isToday?' today':''}${hasTask?' has-task':''}" onclick="showCalDayTasks('${dateStr}')" style="cursor:pointer">${d}</div>`;
  }
  el.innerHTML=`<div class="cal-header">
    <button onclick="calPrev()">&lsaquo;</button>
    <span class="cal-title">${months[calMonth]} ${calYear}</span>
    <button onclick="calNext()">&rsaquo;</button>
  </div><div class="cal-grid">${grid}</div>`;
}
function calPrev(){ calMonth--; if(calMonth<0){calMonth=11;calYear--;} renderCalendar(); }
function calNext(){ calMonth++; if(calMonth>11){calMonth=0;calYear++;} renderCalendar(); }
function showCalDayTasks(dateStr){
  let dayTasks=DB.tasks.filter(t=>t.deadlineDate===dateStr && t.status!=='Completed');
  if(currentUser==='mitiksha') dayTasks=dayTasks.filter(t=>t.assignedTo==='mitiksha');
  const d=new Date(dateStr+'T00:00:00');
  const dateLabel=d.toLocaleDateString('en-GB',{weekday:'long',day:'2-digit',month:'short',year:'numeric'});
  let body='';
  if(!dayTasks.length){
    body='<p style="color:var(--text3);text-align:center;padding:20px">No pending tasks on this date.</p>';
  } else {
    body=dayTasks.map(t=>`<div style="padding:10px 14px;border-bottom:1px solid var(--border-subtle);display:flex;align-items:center;gap:10px;cursor:pointer" onclick="closeModal();viewTask('${t.id}')">
      <div style="flex:1">
        <div style="font-weight:600;font-size:13.5px">${t.urgentFlag?'🔴 ':''}${escHtml(t.title)}</div>
        <div style="font-size:12px;color:var(--text3);margin-top:2px">${USERS[t.assignedTo]?.name||t.assignedTo} · ${statusBadge(t.status)}</div>
      </div>
    </div>`).join('');
  }
  openModal('📅 '+dateLabel, body, [{label:'Close',action:'closeModal()',cls:'btn-secondary'}]);
}

// ═══════════════════════════════════════════════════════
// ROUTINE TASKS SYSTEM
// ═══════════════════════════════════════════════════════
let currentRoutineTab='daily';
let showingRoutineArchive=false;

function initRoutinePage(){
  const isManager=currentUser==='sumudu';
  showingRoutineArchive=false;
  document.getElementById('routine-page-title').textContent='Team Routine Tasks';
  document.getElementById('routine-header-actions').style.display='';
  document.getElementById('routine-user-filter').style.display=isManager?'inline-block':'none';
  document.getElementById('routine-export-btn').style.display=isManager?'inline-block':'none';
  document.getElementById('routine-import-btn').style.display=isManager?'inline-flex':'none';
  document.getElementById('routine-archive-view-btn').style.display='inline-block';
  const addBtn=document.querySelector('#routine-header-actions button[onclick="showAddRoutineTask()"]');
  if(addBtn) addBtn.style.display='inline-block';
  checkAutoArchiveRoutines();
  checkRoutineReappear();
  ensureRoutineCompletions();
  renderRoutineTasks();
}

function switchRoutineTab(tab){
  currentRoutineTab=tab;
  showingRoutineArchive=false;
  ['daily','weekly','monthly','three_months'].forEach(t=>{
    const btn=document.getElementById('routine-tab-'+t);
    if(btn) btn.className='btn btn-sm '+(t===tab?'btn-primary':'btn-secondary');
  });
  renderRoutineTasks();
}

function toggleRoutineArchiveView(){
  showingRoutineArchive=!showingRoutineArchive;
  const btn=document.getElementById('routine-archive-view-btn');
  if(btn) btn.textContent=showingRoutineArchive?'📋 Active Tasks':'📦 Archived';
  renderRoutineTasks();
}

function getRoutineDueDate(frequency){
  const now=new Date();
  const d=now.toISOString().split('T')[0];
  if(frequency==='daily') return d;
  if(frequency==='weekly'){
    const day=now.getDay();
    const mon=new Date(now);
    mon.setDate(now.getDate()-(day===0?6:day-1));
    return mon.toISOString().split('T')[0];
  }
  if(frequency==='monthly') return d.substring(0,8)+'01';
  if(frequency==='three_months'){
    const q=Math.floor(now.getMonth()/3)*3;
    return `${now.getFullYear()}-${String(q+1).padStart(2,'0')}-01`;
  }
  return d.substring(0,8)+'01';
}

function ensureRoutineCompletions(){
  const activeTasks=DB.routineTasks.filter(t=>t.active && !t.archived);
  for(const rt of activeTasks){
    const dueDate=getRoutineDueDate(rt.frequency);
    const exists=DB.routineCompletions.find(c=>c.taskId===rt.id && c.dueDate===dueDate);
    if(!exists){
      const comp={id:'rc_'+Date.now()+'_'+Math.random().toString(36).substr(2,5), taskId:rt.id, dueDate, completedAt:null, completedBy:null, status:'pending', deadlineDate:rt.deadlineDate||'', deadlineTime:rt.deadlineTime||'17:00', autoArchiveAt:null};
      DB.routineCompletions.push(comp);
      if(canWrite()) runDb(sb.from('ff_recurring_completions').insert({id:comp.id, task_id:comp.taskId, due_date:comp.dueDate, completed_at:null, completed_by:null, status:'pending', deadline_date:comp.deadlineDate, deadline_time:comp.deadlineTime, auto_archive_at:null}), 'create routine occurrence', {silent:true});
    }
  }
  markOverdueRoutines();
}

function markOverdueRoutines(){
  const now=new Date();
  for(const comp of DB.routineCompletions){
    if(comp.status!=='pending') continue;
    const rt=DB.routineTasks.find(t=>t.id===comp.taskId);
    if(!rt) continue;
    if(rt.deadlineDate){
      const dl=new Date(rt.deadlineDate+'T'+(rt.deadlineTime||'17:00'));
      if(now>=dl && comp.status==='pending'){
        comp.status='urgent';
      }
    }
  }
}

function canModifyRoutineTask(rt){
  if(currentUser==='sumudu') return true;
  if(currentUser==='de') return true;
  if(currentUser==='mitiksha') return true;
  return false;
}

function canCompleteRoutineTask(rt){
  if(currentUser==='sumudu') return false;
  if(currentUser==='de') return true;
  if(currentUser==='mitiksha' && rt.assignedTo==='mitiksha') return true;
  return false;
}

function checkAutoArchiveRoutines(){
  const now=new Date();
  DB.routineCompletions.forEach(comp=>{
    if(comp.status==='completed' && comp.autoArchiveAt && !comp._archived){
      const archiveTime=new Date(comp.autoArchiveAt);
      if(now>=archiveTime){
        comp._archived=true;
      }
    }
  });
}

function checkRoutineReappear(){
  const now=new Date();
  const activeTasks=DB.routineTasks.filter(t=>t.active && !t.archived);
  for(const rt of activeTasks){
    const hasPending=DB.routineCompletions.find(c=>c.taskId===rt.id && (c.status==='pending'||c.status==='urgent') && !c._archived);
    if(hasPending) continue;
    const lastComp=DB.routineCompletions.filter(c=>c.taskId===rt.id && c.status==='completed').sort((a,b)=>new Date(b.completedAt)-new Date(a.completedAt))[0];
    if(!lastComp) continue;
    let shouldReappear=false;
    let newDeadline='';
    let newDue='';
    let newDeadlineTime=rt.deadlineTime||'17:00';

    if(rt.frequency==='daily'){
      const compDate=new Date(lastComp.completedAt);
      const reappearTime=new Date(compDate);
      reappearTime.setHours(20,0,0,0);
      if(now>=reappearTime || now.toISOString().split('T')[0]>lastComp.dueDate){
        shouldReappear=true;
        const tomorrow=new Date(now); tomorrow.setDate(tomorrow.getDate()+1);
        newDeadline=tomorrow.toISOString().split('T')[0];
        newDue=now.toISOString().split('T')[0];
      }
    } else if(rt.frequency==='weekly'){
      const day=now.getDay();
      const mon=new Date(now);
      mon.setDate(now.getDate()-(day===0?6:day-1));
      mon.setHours(7,0,0,0);
      if(now>=mon && lastComp.dueDate<mon.toISOString().split('T')[0]){
        shouldReappear=true;
        const fri=new Date(mon); fri.setDate(fri.getDate()+4);
        newDeadline=fri.toISOString().split('T')[0];
        newDue=mon.toISOString().split('T')[0];
      }
    } else if(rt.frequency==='monthly'){
      const curYear=now.getFullYear(), curMonth=now.getMonth();
      const day28=new Date(curYear, curMonth, 28, 7, 0, 0);
      const monthStart=now.toISOString().split('T')[0].substring(0,8)+'01';
      if(lastComp.dueDate<monthStart && (now>=day28 || now.getDate()>=28)){
        const lastDay=new Date(curYear, curMonth+1, 0).getDate();
        newDeadline=`${curYear}-${String(curMonth+1).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;
        newDue=monthStart;
        shouldReappear=true;
      }
    } else if(rt.frequency==='three_months'){
      const curQ=Math.floor(now.getMonth()/3);
      const qStart=`${now.getFullYear()}-${String(curQ*3+1).padStart(2,'0')}-01`;
      if(lastComp.dueDate<qStart){
        const m=curQ*3;
        const lastDay=new Date(now.getFullYear(),m+3,0).getDate();
        newDeadline=`${now.getFullYear()}-${String(m+3).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;
        newDue=qStart;
        shouldReappear=true;
      }
    }
    if(shouldReappear && newDue){
      const alreadyExists=DB.routineCompletions.find(c=>c.taskId===rt.id && c.dueDate===newDue && (c.status==='pending'||c.status==='urgent'));
      if(!alreadyExists){
        const comp={id:'rc_'+Date.now()+'_'+Math.random().toString(36).substr(2,5), taskId:rt.id, dueDate:newDue, completedAt:null, completedBy:null, status:'pending', deadlineDate:newDeadline||rt.deadlineDate||'', deadlineTime:newDeadlineTime, autoArchiveAt:null};
        DB.routineCompletions.push(comp);
        if(canWrite()) runDb(sb.from('ff_recurring_completions').insert({id:comp.id, task_id:comp.taskId, due_date:comp.dueDate, completed_at:null, completed_by:null, status:'pending', deadline_date:comp.deadlineDate, deadline_time:comp.deadlineTime, auto_archive_at:null}), 'reappear routine occurrence', {silent:true});
      }
    }
  }
}

function renderRoutineTasks(){
  const el=document.getElementById('routine-tasks-content');
  if(!el) return;
  if(showingRoutineArchive) return renderArchivedRoutines(el);
  const isManager=currentUser==='sumudu';
  const filterUser=isManager?(document.getElementById('routine-user-filter')?.value||'all'):null;

  let tasks=DB.routineTasks.filter(t=>t.active && !t.archived && t.frequency===currentRoutineTab);
  if(currentUser==='de'||currentUser==='mitiksha') tasks=tasks.filter(t=>t.assignedTo===currentUser);
  if(isManager && filterUser!=='all') tasks=tasks.filter(t=>t.assignedTo===filterUser);

  if(!tasks.length){
    el.innerHTML=`<div class="empty-state"><div class="empty-icon">🔄</div><p>No ${currentRoutineTab} routine tasks yet</p></div>`;
    return;
  }

  const dueDate=getRoutineDueDate(currentRoutineTab);
  const periodLabel=currentRoutineTab==='daily'?'Today':currentRoutineTab==='weekly'?'This Week':currentRoutineTab==='three_months'?'This Quarter':'This Month';

  let taskRows=[];
  for(const rt of tasks){
    let comp=DB.routineCompletions.find(c=>c.taskId===rt.id && c.dueDate===dueDate);
    if(!comp) comp=DB.routineCompletions.filter(c=>c.taskId===rt.id && !c._archived).sort((a,b)=>b.dueDate.localeCompare(a.dueDate))[0];
    const status=comp?comp.status:'pending';
    const isDone=status==='completed';
    const isUrgent=status==='urgent';
    const isOD=status==='overdue';
    const dlDate=comp?.deadlineDate||rt.deadlineDate||'';
    const dlTime=comp?.deadlineTime||rt.deadlineTime||'17:00';
    const dlDisplay=dlDate?`${new Date(dlDate+'T'+dlTime).toLocaleDateString('en-GB',{day:'2-digit',month:'short'})} ${dlTime}`:'—';

    let urgentNow=false;
    if(!isDone && dlDate){
      const dl=new Date(dlDate+'T'+dlTime);
      if(new Date()>=dl) urgentNow=true;
    }

    taskRows.push({rt, comp, status, isDone, isUrgent:urgentNow, isOD, dlDisplay, dlDate, dlTime});
  }

  taskRows.sort((a,b)=>{
    if(a.isUrgent && !b.isUrgent) return -1;
    if(!a.isUrgent && b.isUrgent) return 1;
    if(a.isDone && !b.isDone) return 1;
    if(!a.isDone && b.isDone) return -1;
    return 0;
  });

  let html=`<div style="padding:10px 16px;font-size:12px;color:var(--text3);font-weight:600;text-transform:uppercase;letter-spacing:.5px">Period: ${periodLabel} (${fmt(dueDate)})</div>`;
  html+=`<table><thead><tr><th style="width:4%"></th><th style="width:25%">Task</th><th style="font-size:11px">Assigned To</th>
    <th style="font-size:11px">Deadline</th><th style="font-size:11px">Reminder</th><th style="font-size:11px">Status</th><th style="font-size:11px">Actions</th></tr></thead><tbody>`;

  for(const row of taskRows){
    const {rt, comp, status, isDone, isUrgent, dlDisplay}=row;

    let badge='';
    if(isDone) badge='<span style="background:#dcfce7;color:#16a34a;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:600">Done</span>';
    else if(isUrgent) badge='<span style="background:#fee2e2;color:#dc2626;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:600;animation:pulse 1.5s infinite">🔴 Urgent</span>';
    else badge='<span style="background:#fef9c3;color:#92400e;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:600">Pending</span>';

    const canComplete=canCompleteRoutineTask(rt);
    const canMod=canModifyRoutineTask(rt);

    let actionsHtml='<div style="display:flex;gap:4px;flex-wrap:nowrap">';
    if(canComplete && !isDone){
      actionsHtml+=`<button class="btn btn-primary btn-sm" style="font-size:10px;padding:3px 8px" onclick="event.stopPropagation();toggleRoutineTask('${rt.id}','${dueDate}',true)">✅ Complete</button>`;
    }
    if(isDone && canMod){
      actionsHtml+=`<button class="btn btn-warning btn-sm" style="font-size:10px;padding:3px 6px" onclick="event.stopPropagation();archiveRoutineTask('${rt.id}')">📦</button>`;
    }
    if(canMod){
      actionsHtml+=`<button class="btn btn-danger btn-sm" style="font-size:10px;padding:3px 6px" onclick="event.stopPropagation();deleteRoutineTask('${rt.id}')">🗑️</button>`;
    }
    actionsHtml+='</div>';

    const attachCount=(rt.attachments||[]).length;

    html+=`<tr style="opacity:${isDone?'0.6':'1'};cursor:pointer;${isDone?'text-decoration:line-through;':''}" onclick="viewRoutineTask('${rt.id}')">
      <td onclick="event.stopPropagation()">${canComplete?`<input type="checkbox" ${isDone?'checked disabled':''} onchange="toggleRoutineTask('${rt.id}','${dueDate}',this.checked)" style="width:18px;height:18px;cursor:pointer">`:''}</td>
      <td style="font-weight:600"><div>${escHtml(rt.title)}</div><div style="font-size:11px;color:var(--text3);margin-top:2px">${rt.frequency.charAt(0).toUpperCase()+rt.frequency.slice(1)} · by ${USERS[rt.createdBy]?.name||rt.createdBy}${attachCount?' · 📎'+attachCount:''}</div></td>
      <td style="font-size:11px">${USERS[rt.assignedTo]?.name||rt.assignedTo}</td>
      <td style="font-size:11px;${isUrgent?'color:#dc2626;font-weight:700':''}">${dlDisplay}</td>
      <td style="font-size:11px">⏰ ${rt.reminderTime}</td>
      <td>${badge}</td>
      <td onclick="event.stopPropagation()">${actionsHtml}</td>
    </tr>`;
  }
  html+='</tbody></table>';

  const allComps=DB.routineCompletions.filter(c=>c.dueDate===dueDate);
  const relTasks=taskRows;
  const doneCount=relTasks.filter(r=>r.isDone).length;
  const urgentCount=relTasks.filter(r=>r.isUrgent).length;
  html+=`<div style="padding:14px 16px;border-top:1px solid var(--border-subtle);display:flex;gap:20px;font-size:13px;flex-wrap:wrap">
    <span style="color:var(--success)">✅ Completed: <strong>${doneCount}</strong></span>
    <span style="color:var(--danger)">🔴 Urgent: <strong>${urgentCount}</strong></span>
    <span style="color:var(--text3)">📋 Total: <strong>${relTasks.length}</strong></span>
  </div>`;

  el.innerHTML=html;
}

function renderArchivedRoutines(el){
  const completedComps=DB.routineCompletions.filter(c=>c.status==='completed' && c._archived);
  const archivedTasks=DB.routineTasks.filter(t=>t.archived);

  const allArchived=[];
  for(const comp of completedComps){
    const rt=DB.routineTasks.find(t=>t.id===comp.taskId);
    if(rt) allArchived.push({type:'completion', rt, comp});
  }
  for(const rt of archivedTasks){
    allArchived.push({type:'task', rt, comp:null});
  }

  if(!allArchived.length){
    el.innerHTML=`<div class="empty-state"><div class="empty-icon">📦</div><p>No archived routine tasks</p></div>`;
    return;
  }

  let html=`<div style="padding:10px 16px;font-size:12px;color:var(--text3);font-weight:600;text-transform:uppercase;letter-spacing:.5px">Archived Routine Tasks</div>`;
  html+=`<table><thead><tr><th style="width:30%">Task</th><th style="font-size:11px">Assigned To</th><th style="font-size:11px">Frequency</th><th style="font-size:11px">Completed</th><th style="font-size:11px">Actions</th></tr></thead><tbody>`;

  for(const item of allArchived){
    const rt=item.rt;
    const compDate=item.comp?.completedAt?new Date(item.comp.completedAt).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'}):(rt.archivedDate?new Date(rt.archivedDate).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}):'—');
    const isManager=currentUser==='sumudu';
    html+=`<tr style="text-decoration:line-through;opacity:0.7">
      <td style="font-weight:600">${escHtml(rt.title)}</td>
      <td style="font-size:11px">${USERS[rt.assignedTo]?.name||rt.assignedTo}</td>
      <td style="font-size:11px">${rt.frequency}</td>
      <td style="font-size:11px">${compDate}</td>
      <td><div style="display:flex;gap:4px">
        ${item.type==='task'?`<button class="btn btn-secondary btn-sm" style="font-size:10px;padding:3px 6px" onclick="restoreRoutineTask('${rt.id}')">♻️ Restore</button>`:''}
        ${isManager?`<button class="btn btn-danger btn-sm" style="font-size:10px;padding:3px 6px" onclick="${item.type==='task'?`permanentDeleteRoutineTask('${rt.id}')`:`deleteArchivedCompletion('${item.comp?.id}')`}">🗑️</button>`:''}
      </div></td>
    </tr>`;
  }
  html+='</tbody></table>';
  el.innerHTML=html;
}

async function deleteArchivedCompletion(compId){
  if(!confirm('Remove this archived entry?')) return;
  DB.routineCompletions=DB.routineCompletions.filter(c=>c.id!==compId);
  await runDb(sb.from('ff_recurring_completions').delete().eq('id',compId), 'remove archived entry');
  renderRoutineTasks();
  showToast('Archived entry removed','info');
}

async function toggleRoutineTask(taskId, dueDate, checked){
  const rt=DB.routineTasks.find(t=>t.id===taskId);
  if(!rt) return;
  if(!canCompleteRoutineTask(rt)){showToast('You cannot complete this task','error');return;}

  let comp=DB.routineCompletions.find(c=>c.taskId===taskId && c.dueDate===dueDate);
  if(!comp){
    comp={id:'rc_'+Date.now()+'_'+Math.random().toString(36).substr(2,5), taskId, dueDate, completedAt:null, completedBy:null, status:'pending', deadlineDate:rt.deadlineDate||'', deadlineTime:rt.deadlineTime||'17:00', autoArchiveAt:null};
    DB.routineCompletions.push(comp);
  }
  if(checked){
    comp.status='completed';
    comp.completedAt=new Date().toISOString();
    comp.completedBy=currentUser;

    let archiveDelay=0;
    if(rt.frequency==='daily') archiveDelay=3*60*60*1000;
    else if(rt.frequency==='weekly') archiveDelay=2*24*60*60*1000;
    else if(rt.frequency==='monthly') archiveDelay=5*24*60*60*1000;
    else if(rt.frequency==='three_months') archiveDelay=5*24*60*60*1000;
    const archiveAt=new Date(Date.now()+archiveDelay).toISOString();
    comp.autoArchiveAt=archiveAt;

    setTimeout(()=>{
      comp._archived=true;
      renderRoutineTasks();
    }, archiveDelay);

    const ok=await runDb(sb.from('ff_recurring_completions').upsert({id:comp.id, task_id:comp.taskId, due_date:comp.dueDate, completed_at:comp.completedAt, completed_by:comp.completedBy, status:comp.status, deadline_date:comp.deadlineDate, deadline_time:comp.deadlineTime, auto_archive_at:archiveAt}), 'mark routine complete');
    if(!ok){ comp.status='pending'; comp.completedAt=null; comp.completedBy=null; comp.autoArchiveAt=null; renderRoutineTasks(); return; }

    spawnNextRoutineOccurrence(rt);
  } else {
    comp.status='pending';
    comp.completedAt=null;
    comp.completedBy=null;
    comp.autoArchiveAt=null;
    comp._archived=false;
    await runDb(sb.from('ff_recurring_completions').upsert({id:comp.id, task_id:comp.taskId, due_date:comp.dueDate, completed_at:null, completed_by:null, status:'pending', deadline_date:comp.deadlineDate, deadline_time:comp.deadlineTime, auto_archive_at:null}), 'unmark routine task');
  }
  renderRoutineTasks();
  showToast(checked?'✅ Task marked complete':'Task unmarked','success');
}

async function spawnNextRoutineOccurrence(rt){
  let nextDeadline='';
  let nextDue='';
  let appearAt=null;
  const now=new Date();

  if(rt.frequency==='daily'){
    const tomorrow=new Date(now);
    tomorrow.setDate(tomorrow.getDate()+1);
    nextDeadline=tomorrow.toISOString().split('T')[0];
    nextDue=tomorrow.toISOString().split('T')[0];
    appearAt=new Date(now);
    appearAt.setHours(20,0,0,0);
    if(now>=appearAt) appearAt=now;
  } else if(rt.frequency==='weekly'){
    const day=now.getDay();
    const nextMon=new Date(now);
    nextMon.setDate(now.getDate()+(day===0?1:(8-day)));
    nextMon.setHours(7,0,0,0);
    const nextFri=new Date(nextMon);
    nextFri.setDate(nextFri.getDate()+4);
    nextDeadline=nextFri.toISOString().split('T')[0];
    nextDue=nextMon.toISOString().split('T')[0];
    appearAt=nextMon;
  } else if(rt.frequency==='monthly'){
    const curMonth=now.getMonth();
    const curYear=now.getFullYear();
    const appear28=new Date(curYear, curMonth, 28, 7, 0, 0);
    let targetMonth, targetYear;
    if(now.getDate()<28){
      targetMonth=curMonth+1; targetYear=curYear;
    } else {
      targetMonth=curMonth+2; targetYear=curYear;
    }
    if(targetMonth>11){targetMonth-=12; targetYear++;}
    const lastDay=new Date(targetYear, targetMonth+1, 0).getDate();
    nextDeadline=`${targetYear}-${String(targetMonth+1).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;
    nextDue=`${targetYear}-${String(targetMonth+1).padStart(2,'0')}-01`;
    appearAt=now.getDate()<28?appear28:now;
  } else if(rt.frequency==='three_months'){
    const curQ=Math.floor(now.getMonth()/3);
    let nextQ=curQ+1, y=now.getFullYear();
    if(nextQ>3){nextQ=0;y++;}
    const m=nextQ*3;
    const lastDay=new Date(y,m+3,0).getDate();
    nextDue=`${y}-${String(m+1).padStart(2,'0')}-01`;
    nextDeadline=`${y}-${String(m+3).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;
    appearAt=now;
  }

  const alreadyExists=DB.routineCompletions.find(c=>c.taskId===rt.id && c.dueDate===nextDue && (c.status==='pending'||c.status==='urgent'));
  if(!alreadyExists && nextDue){
    const comp={id:'rc_'+Date.now()+'_'+Math.random().toString(36).substr(2,5), taskId:rt.id, dueDate:nextDue, completedAt:null, completedBy:null, status:'pending', deadlineDate:nextDeadline, deadlineTime:rt.deadlineTime||'17:00', autoArchiveAt:null, appearAt:appearAt?appearAt.toISOString():null};
    DB.routineCompletions.push(comp);
    await runDb(sb.from('ff_recurring_completions').insert({id:comp.id, task_id:comp.taskId, due_date:comp.dueDate, completed_at:null, completed_by:null, status:'pending', deadline_date:comp.deadlineDate, deadline_time:comp.deadlineTime, auto_archive_at:null}), 'create routine occurrence', {silent:true});
  }
}

let pendingRoutineFiles=[];

function showAddRoutineTask(){
  pendingRoutineFiles=[];
  const isManager=currentUser==='sumudu';
  const isDev=currentUser==='de';
  let assignField='';
  if(isManager){
    assignField=`<div><label style="font-size:12px;font-weight:600;color:var(--text2)">Assign To</label>
      <select id="rt-assign" class="filter-select" style="width:100%">
        <option value="de">Dev</option><option value="mitiksha">Mitiksha</option>
      </select></div>`;
  } else if(isDev){
    assignField=`<div><label style="font-size:12px;font-weight:600;color:var(--text2)">Assign To</label>
      <select id="rt-assign" class="filter-select" style="width:100%">
        <option value="de">Dev</option><option value="mitiksha">Mitiksha</option>
      </select></div>`;
  } else {
    assignField=`<input type="hidden" id="rt-assign" value="mitiksha">`;
  }

  const todayStr=new Date().toISOString().split('T')[0];

  const body=`<div style="display:flex;flex-direction:column;gap:14px">
    <div><label style="font-size:12px;font-weight:600;color:var(--text2)">Task Title *</label>
      <input type="text" id="rt-title" placeholder="e.g. Check petty cash" style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:14px;background:var(--surface);color:var(--text)"></div>
    ${assignField}
    <div><label style="font-size:12px;font-weight:600;color:var(--text2)">Frequency</label>
      <select id="rt-frequency" class="filter-select" style="width:100%">
        <option value="daily">Daily</option><option value="weekly">Weekly</option><option value="monthly">Monthly</option><option value="three_months">Every 3 Months</option>
      </select></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div><label style="font-size:12px;font-weight:600;color:var(--text2)">Deadline Date *</label>
        <input type="date" id="rt-deadline-date" value="${todayStr}" min="${todayStr}" style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:14px;background:var(--surface);color:var(--text)"></div>
      <div><label style="font-size:12px;font-weight:600;color:var(--text2)">Deadline Time *</label>
        <input type="time" id="rt-deadline-time" value="17:00" style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:14px;background:var(--surface);color:var(--text)"></div>
    </div>
    <div><label style="font-size:12px;font-weight:600;color:var(--text2)">Reminder Time</label>
      <input type="time" id="rt-reminder" value="09:00" style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:14px;background:var(--surface);color:var(--text)"></div>
    <div><label style="font-size:12px;font-weight:600;color:var(--text2)">Link</label>
      <input type="url" id="rt-new-link" placeholder="https://..." style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:14px;background:var(--surface);color:var(--text)"></div>
    <div><label style="font-size:12px;font-weight:600;color:var(--text2)">Remarks</label>
      <textarea id="rt-new-remarks" rows="2" placeholder="Optional remarks…" style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:14px;background:var(--surface);color:var(--text);resize:vertical"></textarea></div>
    <div>
      <label style="font-size:12px;font-weight:600;color:var(--text2)">Attachments</label>
      <div id="rt-new-file-list" style="margin-top:4px;font-size:13px;color:var(--text3)">No files selected</div>
      <label class="btn btn-secondary btn-sm" style="margin-top:6px;cursor:pointer;display:inline-flex;align-items:center;gap:4px">
        📎 Choose File<input type="file" multiple onchange="handleNewRoutineFiles(this)" style="display:none">
      </label>
    </div>
  </div>`;
  openModal('🔄 Add Routine Task', body, [
    {label:'Add Task', action:'saveRoutineTask()', cls:'btn-primary'},
    {label:'Cancel', action:'closeModal()', cls:'btn-secondary'}
  ]);
}

function handleNewRoutineFiles(input){
  pendingRoutineFiles=Array.from(input.files);
  const listEl=document.getElementById('rt-new-file-list');
  if(listEl){
    if(pendingRoutineFiles.length){
      listEl.innerHTML=pendingRoutineFiles.map(f=>`<div style="padding:3px 0">📎 ${escHtml(f.name)} <span style="color:var(--text3)">(${(f.size/1024).toFixed(1)} KB)</span></div>`).join('');
    } else {
      listEl.textContent='No files selected';
    }
  }
}

async function saveRoutineTask(){
  const title=document.getElementById('rt-title')?.value?.trim();
  const assignTo=document.getElementById('rt-assign')?.value||currentUser;
  const frequency=document.getElementById('rt-frequency')?.value;
  const reminder=document.getElementById('rt-reminder')?.value||'09:00';
  const deadlineDate=document.getElementById('rt-deadline-date')?.value;
  const deadlineTime=document.getElementById('rt-deadline-time')?.value||'17:00';
  const link=document.getElementById('rt-new-link')?.value||'';
  const remarks=document.getElementById('rt-new-remarks')?.value||'';
  if(!title){ showToast('Please enter a task title','error'); return; }
  if(!deadlineDate){ showToast('Deadline date is required','error'); return; }
  if(!deadlineTime){ showToast('Deadline time is required','error'); return; }

  const rtId='rt_'+Date.now();
  const attachments=[];

  if(pendingRoutineFiles.length){
    showToast('Uploading files…','info');
    for(const file of pendingRoutineFiles){
      if(file.size>10*1024*1024){ showToast(`${file.name} too large (max 10MB)`,'error'); continue; }
      const filePath=`routine/${rtId}/${Date.now()}_${file.name}`;
      const {error}=await sb.storage.from('attachments').upload(filePath, file);
      if(error){ showToast('Upload failed: '+error.message,'error'); continue; }
      const {data:urlData}=sb.storage.from('attachments').getPublicUrl(filePath);
      attachments.push({name:file.name, url:urlData.publicUrl, uploadedBy:currentUser, uploadedAt:new Date().toISOString()});
    }
  }

  const rt={id:rtId, title, frequency, assignedTo:assignTo, createdBy:currentUser, createdDate:new Date().toISOString(), reminderTime:reminder, active:true, remarks, link, archived:false, archivedDate:null, attachments, deadlineDate, deadlineTime};
  DB.routineTasks.push(rt);
  const ok=await runDb(sb.from('ff_recurring_tasks').insert({id:rt.id, title:rt.title, frequency:rt.frequency, assigned_to:rt.assignedTo, created_by:rt.createdBy, created_date:rt.createdDate, reminder_time:rt.reminderTime, active:true, remarks:rt.remarks, link:rt.link, archived:false, archived_date:null, attachments:rt.attachments, deadline_date:rt.deadlineDate, deadline_time:rt.deadlineTime}), 'add routine task');
  if(!ok){ DB.routineTasks=DB.routineTasks.filter(x=>x.id!==rt.id); return; }
  pendingRoutineFiles=[];
  ensureRoutineCompletions();
  closeModal();
  currentRoutineTab=frequency;
  switchRoutineTab(frequency);
  showToast('✅ Routine task added!','success');
}

function viewRoutineTask(id){
  const rt=DB.routineTasks.find(t=>t.id===id);
  if(!rt) return;
  const dueDate=getRoutineDueDate(rt.frequency);
  const comp=DB.routineCompletions.find(c=>c.taskId===rt.id && c.dueDate===dueDate);
  const status=comp?comp.status:'pending';
  const freqLabel=rt.frequency==='three_months'?'Every 3 Months':rt.frequency.charAt(0).toUpperCase()+rt.frequency.slice(1);

  const dlDate=comp?.deadlineDate||rt.deadlineDate||'';
  const dlTime=comp?.deadlineTime||rt.deadlineTime||'17:00';
  const dlDisplay=dlDate?`${new Date(dlDate+'T'+dlTime).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'})} at ${dlTime}`:'Not set';

  let urgentNow=false;
  if(status!=='completed' && dlDate){
    const dl=new Date(dlDate+'T'+dlTime);
    if(new Date()>=dl) urgentNow=true;
  }

  const statusBadge=status==='completed'?'<span style="background:#dcfce7;color:#16a34a;padding:3px 10px;border-radius:99px;font-size:12px;font-weight:600">Completed</span>'
    :urgentNow?'<span style="background:#fee2e2;color:#dc2626;padding:3px 10px;border-radius:99px;font-size:12px;font-weight:600">🔴 Urgent</span>'
    :'<span style="background:#fef9c3;color:#92400e;padding:3px 10px;border-radius:99px;font-size:12px;font-weight:600">Pending</span>';

  const attachList=(rt.attachments||[]).map(a=>`<div style="display:flex;align-items:center;gap:8px;padding:6px 10px;background:var(--bg-subtle);border-radius:var(--radius-sm);margin-top:4px">
    <span style="font-size:13px">📎</span>
    <a href="${sanitizeUrl(a.url)}" target="_blank" style="font-size:13px;color:var(--primary);flex:1;word-break:break-all">${escHtml(a.name)}</a>
  </div>`).join('');

  const isManager=currentUser==='sumudu';
  const canMod=canModifyRoutineTask(rt);

  const freqField=canMod?`<select id="rt-edit-freq" class="filter-select" style="width:100%;margin-top:2px">
        <option value="daily" ${rt.frequency==='daily'?'selected':''}>Daily</option>
        <option value="weekly" ${rt.frequency==='weekly'?'selected':''}>Weekly</option>
        <option value="monthly" ${rt.frequency==='monthly'?'selected':''}>Monthly</option>
        <option value="three_months" ${rt.frequency==='three_months'?'selected':''}>Every 3 Months</option>
      </select>`:`<div style="font-weight:600;margin-top:2px">${freqLabel}</div>`;

  const reminderField=canMod?`<input type="time" id="rt-edit-reminder" value="${rt.reminderTime}" style="width:100%;padding:4px 8px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:13px;background:var(--surface);color:var(--text);margin-top:2px">`
    :`<div style="font-weight:600;margin-top:2px">⏰ ${rt.reminderTime}</div>`;

  const statusField=canMod?`<select id="rt-edit-status" class="filter-select" style="width:100%;margin-top:2px">
        <option value="pending" ${status==='pending'||status==='urgent'?'selected':''}>Pending</option>
        <option value="completed" ${status==='completed'?'selected':''}>Completed</option>
      </select>`:`<div style="margin-top:4px">${statusBadge}</div>`;

  const body=`<div style="display:flex;flex-direction:column;gap:16px">
    ${canMod?`<div>
      <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">Title ✏️</div>
      <input type="text" id="rt-edit-title" value="${escHtml(rt.title)}" style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:15px;font-weight:600;background:var(--surface);color:var(--text)">
    </div>`:''}
    <div style="background:var(--bg-subtle);padding:12px 16px;border-radius:var(--radius-sm);border-left:4px solid ${urgentNow?'#dc2626':'var(--primary)'}">
      <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px">Deadline ✏️</div>
      <div style="display:flex;gap:10px;align-items:center;margin-top:6px;flex-wrap:wrap">
        <input type="date" id="rt-edit-dl-date" value="${dlDate}" style="padding:6px 10px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:14px;font-weight:600;background:var(--surface);color:${urgentNow?'#dc2626':'var(--text)'}">
        <input type="time" id="rt-edit-dl-time" value="${dlTime}" style="padding:6px 10px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:14px;font-weight:600;background:var(--surface);color:${urgentNow?'#dc2626':'var(--text)'}">
        ${urgentNow?'<span style="font-size:12px;background:#fee2e2;color:#dc2626;padding:2px 8px;border-radius:99px">URGENT</span>':''}
      </div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div>
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px">Assigned To</div>
        <div style="font-weight:600;margin-top:2px">${USERS[rt.assignedTo]?.name||rt.assignedTo}</div>
      </div>
      <div>
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px">Frequency${canMod?' ✏️':''}</div>
        ${freqField}
      </div>
      <div>
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px">Reminder${canMod?' ✏️':''}</div>
        ${reminderField}
      </div>
      <div>
        <div style="font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px">Status${canMod?' ✏️':''}</div>
        ${statusField}
      </div>
    </div>
    <div style="border-top:1px solid var(--border-subtle);padding-top:14px">
      <label style="font-size:12px;font-weight:600;color:var(--text2)">Remarks</label>
      <textarea id="rt-remarks" rows="3" ${canMod?'':'disabled'} style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:14px;background:var(--surface);color:var(--text);resize:vertical;margin-top:4px">${escHtml(rt.remarks||'')}</textarea>
    </div>
    <div>
      <label style="font-size:12px;font-weight:600;color:var(--text2)">Link</label>
      <input type="url" id="rt-link" value="${escHtml(rt.link||'')}" ${canMod?'':'disabled'} placeholder="https://..." style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:var(--radius-sm);font-size:14px;background:var(--surface);color:var(--text);margin-top:4px">
      ${rt.link?`<a href="${sanitizeUrl(rt.link)}" target="_blank" style="color:var(--primary);font-size:12px;margin-top:4px;display:inline-block">🔗 Open link in new tab</a>`:''}
    </div>
    <div>
      <label style="font-size:12px;font-weight:600;color:var(--text2)">Attachments</label>
      <div id="rt-attach-list" style="margin-top:4px">${attachList||'<div style="color:var(--text3);font-size:13px">No attachments yet</div>'}</div>
      ${canMod?`<label class="btn btn-secondary btn-sm" style="margin-top:8px;cursor:pointer;display:inline-flex;align-items:center;gap:4px">
        📎 Attach File<input type="file" id="rt-attach-input" onchange="uploadRoutineAttachment('${rt.id}')" style="display:none">
      </label>`:''}
    </div>
    <div style="font-size:11px;color:var(--text3);border-top:1px solid var(--border-subtle);padding-top:8px">
      Created by ${USERS[rt.createdBy]?.name||rt.createdBy} · ${rt.createdDate?new Date(rt.createdDate).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}):'—'}
    </div>
  </div>`;

  const buttons=[{label:'Save', action:`saveRoutineDetails('${id}')`, cls:'btn-primary'},{label:'Close', action:'closeModal()', cls:'btn-secondary'}];
  openModal('🔄 '+escHtml(rt.title), body, buttons);
}

async function saveRoutineDetails(id){
  const rt=DB.routineTasks.find(t=>t.id===id);
  if(!rt) return;
  const canMod=canModifyRoutineTask(rt);
  if(canMod){
    rt.remarks=document.getElementById('rt-remarks')?.value||'';
    rt.link=document.getElementById('rt-link')?.value||'';
    const newTitle=document.getElementById('rt-edit-title')?.value?.trim();
    if(newTitle && newTitle!==rt.title) rt.title=newTitle;
  }
  const dbUpdate={remarks:rt.remarks, link:rt.link, title:rt.title};

  if(canMod){
    const newFreq=document.getElementById('rt-edit-freq')?.value;
    const newReminder=document.getElementById('rt-edit-reminder')?.value;
    const newStatus=document.getElementById('rt-edit-status')?.value;

    if(newFreq && newFreq!==rt.frequency){
      rt.frequency=newFreq;
      dbUpdate.frequency=newFreq;
    }
    if(newReminder && newReminder!==rt.reminderTime){
      rt.reminderTime=newReminder;
      dbUpdate.reminder_time=newReminder;
    }
    if(newStatus){
      const dueDate=getRoutineDueDate(rt.frequency);
      let comp=DB.routineCompletions.find(c=>c.taskId===id && c.dueDate===dueDate);
      if(!comp){
        comp={id:'rc_'+Date.now()+'_'+Math.random().toString(36).substr(2,5), taskId:id, dueDate, completedAt:null, completedBy:null, status:'pending', deadlineDate:rt.deadlineDate||'', deadlineTime:rt.deadlineTime||'17:00', autoArchiveAt:null};
        DB.routineCompletions.push(comp);
      }
      if(newStatus!==comp.status){
        comp.status=newStatus;
        comp.completedAt=newStatus==='completed'?new Date().toISOString():null;
        comp.completedBy=newStatus==='completed'?currentUser:null;
        if(newStatus==='completed') spawnNextRoutineOccurrence(rt);
        await runDb(sb.from('ff_recurring_completions').upsert({id:comp.id, task_id:comp.taskId, due_date:comp.dueDate, completed_at:comp.completedAt, completed_by:comp.completedBy, status:comp.status, deadline_date:comp.deadlineDate, deadline_time:comp.deadlineTime}), 'sync routine completion', {silent:true});
      }
    }
  }

  const newDlDate=document.getElementById('rt-edit-dl-date')?.value||'';
  const newDlTime=document.getElementById('rt-edit-dl-time')?.value||'17:00';
  if(newDlDate!==rt.deadlineDate||newDlTime!==rt.deadlineTime){
    rt.deadlineDate=newDlDate;
    rt.deadlineTime=newDlTime;
    dbUpdate.deadline_date=newDlDate;
    dbUpdate.deadline_time=newDlTime;
    const dueDate=getRoutineDueDate(rt.frequency);
    const comp=DB.routineCompletions.find(c=>c.taskId===id && c.dueDate===dueDate);
    if(comp){
      comp.deadlineDate=newDlDate;
      comp.deadlineTime=newDlTime;
      await runDb(sb.from('ff_recurring_completions').update({deadline_date:newDlDate, deadline_time:newDlTime}).eq('id',comp.id), 'update routine deadline', {silent:true});
    }
  }

  await runDb(sb.from('ff_recurring_tasks').update(dbUpdate).eq('id',id), 'update routine task');
  closeModal();
  renderRoutineTasks();
  showToast('✅ Saved!','success');
}

async function uploadRoutineAttachment(taskId){
  const input=document.getElementById('rt-attach-input');
  if(!input||!input.files.length) return;
  const file=input.files[0];
  if(file.size>10*1024*1024){ showToast('File too large (max 10MB)','error'); return; }
  const rt=DB.routineTasks.find(t=>t.id===taskId);
  if(!rt) return;
  showToast('Uploading…','info');
  const filePath=`routine/${taskId}/${Date.now()}_${file.name}`;
  const {data,error}=await sb.storage.from('attachments').upload(filePath, file);
  if(error){ showToast('Upload failed: '+error.message,'error'); return; }
  const {data:urlData}=sb.storage.from('attachments').getPublicUrl(filePath);
  const att={name:file.name, url:urlData.publicUrl, uploadedBy:currentUser, uploadedAt:new Date().toISOString()};
  if(!rt.attachments) rt.attachments=[];
  rt.attachments.push(att);
  await runDb(sb.from('ff_recurring_tasks').update({attachments:rt.attachments}).eq('id',taskId), 'save attachment');
  input.value='';
  showToast('📎 File attached!','success');
  viewRoutineTask(taskId);
}

async function archiveRoutineTask(id){
  if(!confirm('Archive this routine task?')) return;
  const rt=DB.routineTasks.find(t=>t.id===id);
  if(!rt) return;
  rt.archived=true;
  rt.archivedDate=new Date().toISOString();
  await runDb(sb.from('ff_recurring_tasks').update({archived:true, archived_date:rt.archivedDate}).eq('id',id), 'archive routine task');
  renderRoutineTasks();
  showToast('📦 Task archived','success');
}

async function restoreRoutineTask(id){
  const rt=DB.routineTasks.find(t=>t.id===id);
  if(!rt) return;
  rt.archived=false;
  rt.archivedDate=null;
  await runDb(sb.from('ff_recurring_tasks').update({archived:false, archived_date:null}).eq('id',id), 'restore routine task');
  ensureRoutineCompletions();
  renderRoutineTasks();
  showToast('♻️ Task restored','success');
}

async function deleteRoutineTask(id){
  if(currentUser!=='sumudu' && currentUser!=='de' && currentUser!=='mitiksha'){showToast('You do not have permission to delete routine tasks','error');return;}
  if(!confirm('Delete this routine task permanently?')) return;
  DB.routineTasks=DB.routineTasks.filter(t=>t.id!==id);
  DB.routineCompletions=DB.routineCompletions.filter(c=>c.taskId!==id);
  await runDb(sb.from('ff_recurring_tasks').delete().eq('id',id), 'delete routine task');
  await runDb(sb.from('ff_recurring_completions').delete().eq('task_id',id), 'delete routine history', {silent:true});
  renderRoutineTasks();
  showToast('Routine task deleted','info');
}

async function permanentDeleteRoutineTask(id){
  if(currentUser!=='sumudu' && currentUser!=='de' && currentUser!=='mitiksha'){showToast('You do not have permission to delete routine tasks','error');return;}
  if(!confirm('Permanently delete this archived task? This cannot be undone.')) return;
  DB.routineTasks=DB.routineTasks.filter(t=>t.id!==id);
  DB.routineCompletions=DB.routineCompletions.filter(c=>c.taskId!==id);
  await runDb(sb.from('ff_recurring_tasks').delete().eq('id',id), 'delete routine task');
  await runDb(sb.from('ff_recurring_completions').delete().eq('task_id',id), 'delete routine history', {silent:true});
  renderRoutineTasks();
  showToast('Task permanently deleted','info');
}

function exportRoutineTasks(){
  const tasks=DB.routineTasks.filter(t=>t.active && !t.archived);
  if(!tasks.length){ showToast('No routine tasks to export','warning'); return; }
  const rows=tasks.map(rt=>{
    const dueDate=getRoutineDueDate(rt.frequency);
    const comp=DB.routineCompletions.find(c=>c.taskId===rt.id && c.dueDate===dueDate);
    return {
      Title:rt.title, Frequency:rt.frequency, 'Assigned To':USERS[rt.assignedTo]?.name||rt.assignedTo,
      'Deadline Date':rt.deadlineDate||'', 'Deadline Time':rt.deadlineTime||'',
      'Reminder Time':rt.reminderTime, Status:comp?comp.status:'pending',
      Remarks:rt.remarks||'', Link:rt.link||'', 'Created By':USERS[rt.createdBy]?.name||rt.createdBy
    };
  });
  const ws=XLSX.utils.json_to_sheet(rows);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Routine Tasks');
  XLSX.writeFile(wb,'RoutineTasks_'+today()+'.xlsx');
  showToast('📥 Exported!','success');
}


async function importRoutineTasks(input){
  const file=input.files[0];
  if(!file) return;
  const data=await file.arrayBuffer();
  const wb=XLSX.read(data);
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws);
  let count=0;
  for(const row of rows){
    const title=(row.Title||row.title||'').trim();
    if(!title) continue;
    const freq=(row.Frequency||row.frequency||'daily').toLowerCase();
    if(!['daily','weekly','monthly'].includes(freq)) continue;
    let assignTo=(row['Assigned To']||row.assigned_to||'').toLowerCase();
    if(assignTo.includes('dev')) assignTo='de';
    else if(assignTo.includes('miti')) assignTo='mitiksha';
    else assignTo='de';
    const reminder=row['Reminder Time']||row.reminder_time||'09:00';
    const rt={id:'rt_'+Date.now()+'_'+count, title, frequency:freq, assignedTo:assignTo, createdBy:'sumudu', createdDate:new Date().toISOString(), reminderTime:reminder, active:true, remarks:row.Remarks||'', link:row.Link||'', archived:false, archivedDate:null};
    DB.routineTasks.push(rt);
    await runDb(sb.from('ff_recurring_tasks').insert({id:rt.id, title:rt.title, frequency:rt.frequency, assigned_to:rt.assignedTo, created_by:rt.createdBy, created_date:rt.createdDate, reminder_time:rt.reminderTime, active:true, remarks:rt.remarks, link:rt.link, archived:false, archived_date:null}), 're-create routine task', {silent:true});
    count++;
  }
  input.value='';
  if(count){
    ensureRoutineCompletions();
    renderRoutineTasks();
    showToast(`📤 Imported ${count} routine task(s)!`,'success');
  } else {
    showToast('No valid tasks found in file','warning');
  }
}

// Routine task reminders — check every minute + auto 5PM EOD
let routineReminderInterval=null;
function startRoutineReminders(){
  if(routineReminderInterval) clearInterval(routineReminderInterval);
  routineReminderInterval=setInterval(()=>{checkRoutineReminders();checkAutoRoutineEOD();checkAutoArchiveRoutines();checkRoutineReappear();renderRoutineTasks();}, 60000);
  checkRoutineReminders();
}

function checkRoutineReminders(){
  if(!currentUser || currentUser==='sumudu') return;
  const now=new Date();
  const hhmm=String(now.getHours()).padStart(2,'0')+':'+String(now.getMinutes()).padStart(2,'0');
  const todayStr=today();
  const shownKey='ffRoutineReminder_'+currentUser+'_'+todayStr;
  const shown=JSON.parse(localStorage.getItem(shownKey)||'[]');

  const myTasks=DB.routineTasks.filter(t=>t.active && !t.archived && t.assignedTo===currentUser);
  for(const rt of myTasks){
    if(shown.includes(rt.id)) continue;
    if(rt.reminderTime!==hhmm) continue;
    const dueDate=getRoutineDueDate(rt.frequency);
    const comp=DB.routineCompletions.find(c=>c.taskId===rt.id && c.dueDate===dueDate);
    if(comp && comp.status==='completed') continue;
    shown.push(rt.id);
    localStorage.setItem(shownKey, JSON.stringify(shown));
    showRoutineReminder(rt);
  }
}

function checkAutoRoutineEOD(){
  const now=new Date();
  const hh=now.getHours(), mm=now.getMinutes();
  if(hh!==17 || mm!==0) return;
  const sentKey='ffRoutineEODSent_'+today();
  if(localStorage.getItem(sentKey)) return;
  localStorage.setItem(sentKey,'1');
  sendRoutineEODEmail();
}

function showRoutineReminder(rt){
  const freqLabel=rt.frequency==='three_months'?'Every 3 Months':rt.frequency.charAt(0).toUpperCase()+rt.frequency.slice(1);
  openModal('⏰ Routine Task Reminder', `
    <div style="text-align:center;padding:20px">
      <div style="font-size:48px;margin-bottom:12px">🔔</div>
      <h3 style="margin:0 0 8px;font-size:18px">${escHtml(rt.title)}</h3>
      <p style="color:var(--text3);font-size:14px">${freqLabel} task — due now</p>
    </div>`, [
    {label:'Go to Routine Tasks', action:"closeModal();navigateTo('routine-tasks')", cls:'btn-primary'},
    {label:'Dismiss', action:'closeModal()', cls:'btn-secondary'}
  ]);
}

async function sendRoutineEODEmail(){
  for(const uid of ['de','mitiksha']){
    const u=USERS[uid];
    const myTasks=DB.routineTasks.filter(t=>t.active && !t.archived && t.assignedTo===uid);
    if(!myTasks.length) continue;
    const todayDue=getRoutineDueDate('daily');
    const weekDue=getRoutineDueDate('weekly');
    const monthDue=getRoutineDueDate('monthly');
    const qDue=getRoutineDueDate('three_months');

    let rows='';
    let doneCount=0, odCount=0;
    for(const rt of myTasks){
      const due=rt.frequency==='daily'?todayDue:rt.frequency==='weekly'?weekDue:rt.frequency==='three_months'?qDue:monthDue;
      const comp=DB.routineCompletions.find(c=>c.taskId===rt.id && c.dueDate===due);
      const st=comp?comp.status:'pending';
      if(st==='completed') doneCount++;
      if(st==='overdue') odCount++;
      const stColor=st==='completed'?'#22c55e':st==='overdue'?'#ef4444':'#f59e0b';
      rows+=`<tr><td>${escHtml(rt.title)}</td><td>${rt.frequency}</td><td style="color:${stColor};font-weight:600">${st.charAt(0).toUpperCase()+st.slice(1)}</td><td>${escHtml(rt.remarks||'—')}</td></tr>`;
    }

    const subject=`🔄 Routine Tasks EOD — ${u.name}`;
    const body=`<p>Hi <strong>${u.name}</strong>,</p>
      <p>Here's your routine tasks summary for today:</p>
      <table><tr><th>Task</th><th>Frequency</th><th>Status</th><th>Remarks</th></tr>${rows}</table>
      <p style="margin-top:12px">✅ Completed: <strong>${doneCount}</strong> | ⚠️ Overdue: <strong>${odCount}</strong> | 📋 Total: <strong>${myTasks.length}</strong></p>
      ${odCount?'<p style="color:#ef4444"><strong>Please complete your overdue routine tasks!</strong></p>':''}
      <p style="margin-top:14px;color:#6366f1;font-weight:600">— FinanceFlow Task Management</p>`;
    const email={id:'em_'+Date.now()+Math.random(),type:'routine_eod',subject,to:u.email,from:'system@financeflow.com',body,timestamp:new Date().toISOString()};
    DB.emailLog.push(email);
    await dbInsertEmail(email);
    await sendRealEmail(u.email, subject, body);
  }
  showToast('🔄 Routine EOD emails sent!','success');
  if(document.getElementById('page-emails')?.classList.contains('active')) renderEmailLog();
}

// ═══════════════════════════════════════════════════════
// KEYBOARD NAV — Escape to close modal/sidebar
// ═══════════════════════════════════════════════════════
document.addEventListener('keydown', e=>{
  if(e.key==='Escape'){
    if(!document.getElementById('modal-overlay').classList.contains('hidden')) closeModal();
    closeSidebar();
  }
});

// ═══════════════════════════════════════════════════════
// SEED DATA
// ═══════════════════════════════════════════════════════
function seedData(){}

// ═══════════════════════════════════════════════════════
// MIGRATE LOCALSTORAGE → SUPABASE
// ═══════════════════════════════════════════════════════
async function migrateLocalToSupabase(){
  try {
    const d=localStorage.getItem('ffdb');
    if(!d){ showToast('No local data to migrate','warning'); return; }
    const local=JSON.parse(d);
    if(!local.tasks||!local.tasks.length){ showToast('No tasks to migrate','warning'); return; }
    let count=0;
    for(const t of local.tasks){
      const row=taskToRow(t);
      const {error}=await sb.from('ff_tasks').upsert(row);
      if(!error) count++;
    }
    for(const c of (local.comments||[])){
      await sb.from('ff_comments').upsert({id:c.id,task_id:c.taskId,user_id:c.userId,text:c.text,created_date:c.createdDate,attachments:c.attachments||[]});
    }
    for(const e of (local.emailLog||[])){
      await sb.from('ff_email_log').upsert({id:e.id,type:e.type,subject:e.subject,to_email:e.to,from_email:e.from,body:e.body,task_id:e.taskId||null,timestamp:e.timestamp});
    }
    const sp=local.stepProgress||{};
    for(const tid of Object.keys(sp)){
      for(const sid of Object.keys(sp[tid])){
        await sb.from('ff_step_progress').upsert({task_id:tid,step_id:sid,done:sp[tid][sid]});
      }
    }
    for(const t of (local.recycleBin||[])){
      await dbInsertRecycleBin(t).catch(()=>{});
    }
    localStorage.removeItem('ffdb');
    localStorage.removeItem('ffdb_v');
    await loadDB();
    showToast(`✅ Migrated ${count} tasks to cloud!`,'success');
    renderDashboard();
  } catch(e){ showToast('Migration error: '+e.message,'error'); console.error(e); }
}

// ═══════════════════════════════════════════════════════
// PERSONAL TASKS (Dev/Mitiksha)
// ═══════════════════════════════════════════════════════
let ptFilter='all';
function switchPTFilter(f){
  ptFilter=f;
  document.querySelectorAll('[id^="pt-tab-"]').forEach(b=>{
    b.className=b.id==='pt-tab-'+f?'btn btn-sm btn-primary':'btn btn-sm btn-secondary';
  });
  renderPersonalTasks();
}
function markOverduePersonalTasks(){
  const now=new Date();
  DB.personalTasks.forEach(t=>{
    if(t.status==='Pending'){
      const dl=new Date(t.deadlineDate+'T'+t.deadlineTime);
      if(dl<now) t.status='Overdue';
    }
  });
}
function renderPersonalTasks(){
  const el=document.getElementById('pt-officer-content');
  if(!el) return;
  let tasks=DB.personalTasks.filter(t=>t.owner===currentUser);
  if(ptFilter!=='all') tasks=tasks.filter(t=>t.status===ptFilter);
  tasks.sort((a,b)=>new Date(a.deadlineDate+'T'+a.deadlineTime)-new Date(b.deadlineDate+'T'+b.deadlineTime));
  if(!tasks.length){el.innerHTML='<p style="text-align:center;color:var(--text3);padding:32px">No tasks found</p>';return;}
  el.innerHTML=tasks.map(t=>{
    const dl=new Date(t.deadlineDate+'T'+t.deadlineTime);
    const statusColor=t.status==='Completed'?'#10b981':t.status==='Overdue'?'#ef4444':'#f59e0b';
    const recLabel=t.recurrence!=='none'?` <span style="background:var(--bg2);padding:2px 8px;border-radius:12px;font-size:11px">${t.recurrence}</span>`:'';
    return `<div style="display:flex;justify-content:space-between;align-items:center;padding:14px 16px;border-bottom:1px solid var(--border);gap:12px">
      <div style="flex:1;min-width:0">
        <div style="font-weight:600;font-size:14px;${t.status==='Completed'?'text-decoration:line-through;opacity:.6':''}">${escHtml(t.title)}${recLabel}</div>
        ${t.description?`<div style="font-size:12px;color:var(--text3);margin-top:2px">${escHtml(t.description)}</div>`:''}
        <div style="font-size:11px;color:var(--text3);margin-top:4px">Due: ${dl.toLocaleDateString('en-GB')} ${dl.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'})}${t.link?` | <a href="${sanitizeUrl(t.link)}" target="_blank" style="color:var(--primary)">Link</a>`:''}</div>
      </div>
      <div style="display:flex;gap:8px;align-items:center;flex-shrink:0">
        <span style="background:${statusColor};color:#fff;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600">${t.status}</span>
        ${t.status!=='Completed'?`<button class="btn btn-primary btn-sm" onclick="markCompletePersonalTask('${t.id}')">Complete</button>`:''}
      </div>
    </div>`;
  }).join('');
}
function showAddPersonalTask(){
  const today=new Date().toISOString().split('T')[0];
  openModal('Add Personal Task',`
    <div class="form-group"><label class="form-label">Title *</label><input class="form-input" id="pt-title" placeholder="Task title"></div>
    <div class="form-group"><label class="form-label">Description</label><textarea class="form-input" id="pt-desc" rows="2" placeholder="Optional description"></textarea></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="form-group"><label class="form-label">Deadline Date *</label><input type="date" class="form-input" id="pt-date" min="${today}" required></div>
      <div class="form-group"><label class="form-label">Deadline Time *</label><input type="time" class="form-input" id="pt-time" value="17:00" required></div>
    </div>
    <div class="form-group"><label class="form-label">Recurrence</label>
      <select class="form-input" id="pt-recurrence">
        <option value="none">None (One-time)</option>
        <option value="daily">Daily</option>
        <option value="weekly">Weekly</option>
        <option value="monthly">Monthly</option>
        <option value="quarterly">Every 3 Months</option>
      </select>
    </div>
    <div class="form-group"><label class="form-label">Link (optional)</label><input type="url" class="form-input" id="pt-link" placeholder="https://..."></div>
    <div style="text-align:right;margin-top:16px">
      <button class="btn btn-secondary" onclick="closeModal()">Cancel</button>
      <button class="btn btn-primary" onclick="savePersonalTask()" style="margin-left:8px">Save Task</button>
    </div>
  `);
}
async function savePersonalTask(){
  const title=document.getElementById('pt-title').value.trim();
  const desc=document.getElementById('pt-desc').value.trim();
  const date=document.getElementById('pt-date').value;
  const time=document.getElementById('pt-time').value;
  const recurrence=document.getElementById('pt-recurrence').value;
  const link=document.getElementById('pt-link').value.trim();
  if(!title){showToast('Title is required','error');return;}
  if(!date){showToast('Deadline date is required','error');return;}
  if(!time){showToast('Deadline time is required','error');return;}
  const task={title,description:desc,owner:currentUser,deadline_date:date,deadline_time:time,recurrence,status:'Pending',link,created_date:new Date().toISOString().split('T')[0],attachments:[]};
  const{data,error}=await sb.from('ff_personal_tasks').insert([task]).select();
  if(error){showToast('Error saving task','error');console.error(error);return;}
  const r=data[0];
  DB.personalTasks.push({id:r.id,title:r.title,description:r.description||'',owner:r.owner,deadlineDate:r.deadline_date,deadlineTime:r.deadline_time||'17:00',recurrence:r.recurrence||'none',status:r.status||'Pending',completionDate:r.completion_date,createdDate:r.created_date,link:r.link||'',attachments:r.attachments||[]});
  closeModal();
  showToast('Task added!','success');
  renderPersonalTasks();
}
async function markCompletePersonalTask(id){
  const t=DB.personalTasks.find(x=>x.id===id);
  if(!t) return;
  t.status='Completed';
  t.completionDate=new Date().toISOString().split('T')[0];
  const ok=await runDb(sb.from('ff_personal_tasks').update({status:'Completed',completion_date:t.completionDate}).eq('id',id), 'complete task');
  if(!ok){ t.status='Pending'; t.completionDate=null; renderPersonalTasks(); return; }
  showToast('Task completed!','success');
  if(t.recurrence!=='none') spawnNextPersonalTask(t);
  renderPersonalTasks();
}
async function spawnNextPersonalTask(t){
  const d=new Date(t.deadlineDate);
  if(t.recurrence==='daily') d.setDate(d.getDate()+1);
  else if(t.recurrence==='weekly') d.setDate(d.getDate()+7);
  else if(t.recurrence==='monthly') d.setMonth(d.getMonth()+1);
  else if(t.recurrence==='quarterly') d.setMonth(d.getMonth()+3);
  const next={title:t.title,description:t.description||'',owner:t.owner,deadline_date:d.toISOString().split('T')[0],deadline_time:t.deadlineTime,recurrence:t.recurrence,status:'Pending',link:t.link||'',created_date:new Date().toISOString().split('T')[0],attachments:[]};
  const{data,error}=await sb.from('ff_personal_tasks').insert([next]).select();
  if(!error&&data&&data[0]){
    const r=data[0];
    DB.personalTasks.push({id:r.id,title:r.title,description:r.description||'',owner:r.owner,deadlineDate:r.deadline_date,deadlineTime:r.deadline_time||'17:00',recurrence:r.recurrence||'none',status:'Pending',completionDate:null,createdDate:r.created_date,link:r.link||'',attachments:r.attachments||[]});
  }
}
function toggleSumuduPTView(showTeam){
  document.getElementById('personal-tasks-sumudu-view').style.display=showTeam?'':'none';
  document.getElementById('personal-tasks-sumudu-own').style.display=showTeam?'none':'';
  if(showTeam) renderPersonalTasksView();
}
function renderPersonalTasksView(){
  const user=document.getElementById('pt-view-user').value;
  const el=document.getElementById('pt-view-content');
  if(!el) return;
  markOverduePersonalTasks();
  let tasks=DB.personalTasks.filter(t=>t.owner===user);
  tasks.sort((a,b)=>new Date(a.deadlineDate+'T'+a.deadlineTime)-new Date(b.deadlineDate+'T'+b.deadlineTime));
  if(!tasks.length){el.innerHTML='<p style="text-align:center;color:var(--text3);padding:32px">No personal tasks</p>';return;}
  const pending=tasks.filter(t=>t.status==='Pending').length;
  const completed=tasks.filter(t=>t.status==='Completed').length;
  const overdue=tasks.filter(t=>t.status==='Overdue').length;
  let html=`<div style="display:flex;gap:16px;margin-bottom:16px;flex-wrap:wrap">
    <div style="background:#f59e0b22;color:#f59e0b;padding:8px 16px;border-radius:8px;font-weight:600">Pending: ${pending}</div>
    <div style="background:#10b98122;color:#10b981;padding:8px 16px;border-radius:8px;font-weight:600">Completed: ${completed}</div>
    <div style="background:#ef444422;color:#ef4444;padding:8px 16px;border-radius:8px;font-weight:600">Overdue: ${overdue}</div>
  </div>`;
  html+=tasks.map(t=>{
    const dl=new Date(t.deadlineDate+'T'+t.deadlineTime);
    const statusColor=t.status==='Completed'?'#10b981':t.status==='Overdue'?'#ef4444':'#f59e0b';
    const recLabel=t.recurrence!=='none'?` <span style="background:var(--bg2);padding:2px 8px;border-radius:12px;font-size:11px">${t.recurrence}</span>`:'';
    return `<div style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;border-bottom:1px solid var(--border)">
      <div style="flex:1"><div style="font-weight:600;font-size:14px;${t.status==='Completed'?'text-decoration:line-through;opacity:.6':''}">${escHtml(t.title)}${recLabel}</div>
      <div style="font-size:11px;color:var(--text3);margin-top:4px">Due: ${dl.toLocaleDateString('en-GB')} ${dl.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'})}</div></div>
      <span style="background:${statusColor};color:#fff;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600">${t.status}</span>
    </div>`;
  }).join('');
  el.innerHTML=html;
}

// ═══════════════════════════════════════════════════════
// INIT — called once by main.js after the DOM is ready.
// ═══════════════════════════════════════════════════════
export async function bootstrapApp(){
  seedData();

  const savedTheme=localStorage.getItem('ffTheme')||'light';
  document.documentElement.setAttribute('data-theme',savedTheme);

  document.getElementById('motivation-auth').textContent=`”${getMotivation()}”`;

  // Pre-select the last-used account in the dropdown (UI convenience only).
  const savedUser=localStorage.getItem('ffUser');
  if(savedUser && USERS[savedUser]){
    document.getElementById('login-user').value=savedUser;
  }

  // Restore a session only if Supabase confirms a valid, non-expired one exists.
  try{
    const { data:{ session } } = await sb.auth.getSession();
    if(session && session.user){
      const uname = EMAIL_TO_USERNAME[(session.user.email||'').toLowerCase()];
      if(uname && USERS[uname]){
        currentUser = uname;
        document.getElementById('login-user').value = uname;
        initApp();
      }
    }
  }catch(e){ console.error('session restore failed', e); }
}

// ═══════════════════════════════════════════════════════
// EXPORTS — functions referenced by inline onclick="" handlers in
// index.html. main.js re-exposes these on window so the existing
// markup keeps working without rewriting every handler to
// addEventListener (a separate, larger follow-up change).
// ═══════════════════════════════════════════════════════
export {
  addAttachmentToTask,
  addComment,
  addStep,
  archiveRoutineTask,
  archiveTask,
  calNext,
  calPrev,
  closeModal,
  closeSidebar,
  createTask,
  dashStatClick,
  deleteArchivedCompletion,
  deleteRoutineTask,
  editTask,
  exportExcel,
  exportPerfExcel,
  exportRoutineTasks,
  forgotPassword,
  handleNewRoutineFiles,
  importExcel,
  importRoutineTasks,
  login,
  logout,
  markAllNotifSeen,
  markChecked,
  markComplete,
  markCompletePersonalTask,
  markNotifSeen,
  migrateLocalToSupabase,
  navigateTo,
  openDocRequest,
  openNewTask,
  openRemarks,
  openSidebar,
  permDelete,
  permanentDeleteRoutineTask,
  previewFiles,
  removeAttachment,
  removeAttachmentFromEdit,
  renderMyTasks,
  renderPerformance,
  renderPersonalTasksView,
  renderRoutineTasks,
  renderSchedule,
  renderTasks,
  toggleJoshuaCategory,
  restoreRoutineTask,
  restoreTask,
  savePersonalTask,
  saveTaskLink,
  sendEODSummary,
  sendMorningSummary,
  sendRoutineEODEmail,
  showAddPersonalTask,
  showAddRoutineTask,
  showCalDayTasks,
  showNotifications,
  switchPTFilter,
  switchRoutineTab,
  toggleRoutineArchiveView,
  toggleRoutineTask,
  toggleStep,
  toggleSumuduPTView,
  toggleTheme,
  uploadRoutineAttachment,
  viewEmailById,
  viewRoutineTask,
  viewTask,
  globalSearch,
  canWrite
};
