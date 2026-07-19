import { createClient } from '@supabase/supabase-js';

// ═══════════════════════════════════════════════════════
// SUPABASE CONFIG
// ═══════════════════════════════════════════════════════
const SUPABASE_URL = 'https://ohkkxaxgiziuvdovpnmu.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9oa2t4YXhnaXppdXZkb3Zwbm11Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODExOTQ2NzgsImV4cCI6MjA5Njc3MDY3OH0.eL_HrHpT2ge8SssZzd1X-ST-wswzNSdbVPSd6eSdlJI';
export const sb = createClient(SUPABASE_URL, SUPABASE_KEY);

// ═══════════════════════════════════════════════════════
// DATA STORE — static config only. Mutable app state (DB, currentUser,
// charts) lives in app.js since it's reassigned by login/logout/loadDB
// and ES module live-bindings only allow reassignment from the module
// that declared the variable.
// ═══════════════════════════════════════════════════════
// Passwords are NOT stored here — authentication is handled by Supabase Auth.
// This map only holds display info (name/role/email/avatar) keyed by the app's
// short username. Email is used to resolve the account for sign-in.
export const USERS = {
  sumudu: { name:'Sumudu', role:'Finance Manager', email:'sumudu@focusawards.org.uk', avatar:'SU', color:'#6366f1' },
  de:     { name:'Dev',    role:'Officer',         email:'devanand@focusawards.org.uk', avatar:'DV', color:'#22c55e' },
  mitiksha:{ name:'Mitiksha',role:'Officer',       email:'mitiksha@focusawards.org.uk', avatar:'MI', color:'#f59e0b' },
  trupal:{ name:'Trupal', role:'Office Manager', email:'trupal@focusawards.org.uk', avatar:'TR', color:'#8b5cf6' }
};
// Reverse lookup: auth email -> short username, so a restored session resolves to currentUser.
export const EMAIL_TO_USERNAME = Object.fromEntries(Object.entries(USERS).map(([u, d]) => [d.email.toLowerCase(), u]));

// View-only users (Trupal) are read-only. The database enforces this via RLS;
// this client check just avoids firing background writes that would be denied.
export const VIEW_ONLY_USERS = ['trupal'];

export const MOTIVATIONS = [
  "You're doing amazing today!","Every task completed is progress!",
  "Success comes from consistency!","You're the best!",
  "Great things take time — keep going!","Your dedication inspires everyone!",
  "One task at a time — you've got this!","Excellence is a habit — keep building it!",
  "Today is a great day to be productive!","Your hard work makes the difference!",
  "Believe in yourself — results will follow!","Strong teams are built one task at a time!"
];

export function getMotivation(){ return MOTIVATIONS[new Date().getDate() % MOTIVATIONS.length]; }
