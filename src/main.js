import './style.css';
import * as App from './app.js';
import { bootstrapApp } from './app.js';

// The HTML markup still uses inline onclick="functionName(...)" handlers
// (78+ of them). Converting every one to addEventListener is a separate,
// larger follow-up change (it touches every render function and needs
// careful verification) — see the note in app.js's export block. For now,
// re-expose every function app.js exports as a window global so the
// existing markup keeps working unchanged under the new module build.
Object.keys(App).forEach(name => { window[name] = App[name]; });

document.addEventListener('DOMContentLoaded', () => {
  bootstrapApp();
});
