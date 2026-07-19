// ═══════════════════════════════════════════════════════
// UTILS — pure functions with no shared app state, safe to import anywhere.
// ═══════════════════════════════════════════════════════

export function today(){ return new Date().toISOString().split('T')[0]; }
export function fmt(d){ if(!d) return '—'; const dt=new Date(d+'T00:00:00'); return dt.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}); }

export function countWorkingDays(startDate, endDate){
  const start=new Date(startDate+'T00:00:00');
  const end=new Date(endDate+'T00:00:00');
  let count=0;
  const d=new Date(start);
  while(d<=end){
    const day=d.getDay();
    if(day!==0 && day!==6) count++;
    d.setDate(d.getDate()+1);
  }
  return count;
}

export function addWorkingDays(startDate, days){
  const d=new Date(startDate+'T00:00:00');
  let added=0;
  while(added<days){
    d.setDate(d.getDate()+1);
    if(d.getDay()!==0 && d.getDay()!==6) added++;
  }
  return d.toISOString().split('T')[0];
}

export function isOverdue(task){
  if(task.status==='Completed') return false;
  if(!task.deadlineDate) return false;
  const now=new Date();
  const dl=new Date(task.deadlineDate+'T23:59:59');
  if(now.getDay()===0||now.getDay()===6) return task.deadlineDate<today();
  return task.deadlineDate < today();
}

export function escHtml(s){ const d=document.createElement('div'); d.textContent=String(s||''); return d.innerHTML; }

// escHtml alone stops tag/attribute injection but NOT dangerous URL schemes
// (e.g. a task's External Link field set to "javascript:alert(1)" would still
// render as a clickable, executing link). Only allow http(s)/mailto/tel; any
// href built from user-controlled data must go through this first.
export function sanitizeUrl(u){
  const s=String(u||'').trim();
  if(!s) return '';
  try{
    const parsed=new URL(s, window.location.href);
    if(['http:','https:','mailto:','tel:'].includes(parsed.protocol)) return escHtml(s);
  }catch(e){ /* not a valid absolute/relative URL */ }
  return '#';
}
