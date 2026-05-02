<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover, user-scalable=yes">
    <title>GW81 Inspector Pro | Admin Control & User Approval</title>
    <!-- External Libraries -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <style>
        /* ========== BASE THEME VARIABLES ========== */
        * { margin: 0; padding: 0; box-sizing: border-box; }
        :root {
            --navy: #060F1E; --navy2: #0A1929; --navy3: #0F2744; --navy4: #163355;
            --blue: #1565C0; --blue2: #1976D2; --blue3: #42A5F5; --blue-light: #E3F2FD;
            --green: #1B5E20; --green2: #2E7D32; --green3: #43A047; --green-light: #E8F5E9;
            --red: #B71C1C; --red2: #C62828; --red3: #EF5350; --red-light: #FFEBEE;
            --amber: #E65100; --amber2: #F57C00; --amber-light: #FFF3E0;
            --cyan: #00838F; --cyan2: #00ACC1; --cyan-light: #E0F7FA;
            --purple: #4527A0; --purple2: #5C35CC; --purple-light: #EDE7F6;
            --gray0: #F8FAFC; --gray1: #F1F5F9; --gray2: #E2E8F0; --gray3: #94A3B8;
            --gray4: #64748B; --gray5: #475569; --gray6: #334155; --gray7: #1E293B;
            --white: #FFFFFF;
            --r: 8px; --r-lg: 12px; --r-xl: 18px;
            --sh: 0 1px 3px rgba(0,0,0,.1), 0 4px 12px rgba(0,0,0,.08);
            --font: 'Inter', system-ui, sans-serif; --mono: 'JetBrains Mono', monospace;
        }
        body.light { --navy: #F1F5F9; --navy2: #FFFFFF; --navy3: #F8FAFC; --gray0: #FFFFFF; --gray1: #F1F5F9; --gray2: #E2E8F0; --gray6: #1E293B; --gray7: #0F172A; --blue2: #1976D2; }
        body.forest { --navy: #0F2A1F; --navy2: #1A3A2A; --navy3: #234F35; --gray0: #E8F5E9; --blue2: #2E7D32; --blue3: #43A047; }
        body.ocean { --navy: #0A1A2F; --navy2: #0F2842; --navy3: #1A3A5C; --blue2: #00838F; --blue3: #00ACC1; }
        body.sunset { --navy: #2D1B2E; --navy2: #3D2640; --navy3: #5A3D5C; --blue2: #E65100; --blue3: #F57C00; --gray0: #FFF3E0; }
        body { font-family: var(--font); background: var(--navy); color: var(--gray6); font-size: 14px; min-height: 100vh; transition: background 0.2s, color 0.2s; }

        /* ========== LOGIN (unchanged) ========== */
        #login-screen { position: fixed; inset: 0; background: radial-gradient(ellipse at 30% 40%, #0F172A, #020617); display: flex; align-items: center; justify-content: center; z-index: 1000; }
        .login-box { background: rgba(17,24,39,0.9); border: 1px solid rgba(59,130,246,0.3); border-radius: 32px; padding: 2rem; width: 420px; max-width: 90vw; backdrop-filter: blur(16px); }
        .login-logo { display: flex; align-items: center; gap: 12px; margin-bottom: 28px; justify-content: center; }
        .login-icon { width: 52px; height: 52px; background: linear-gradient(135deg, var(--blue2), var(--blue3)); border-radius: 14px; display: flex; align-items: center; justify-content: center; font-size: 24px; }
        .login-title { color: #fff; font-size: 22px; font-weight: 700; }
        .login-sub { color: var(--gray3); font-size: 12px; font-family: var(--mono); }
        .login-field input { width: 100%; background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.12); border-radius: 8px; padding: 11px 14px; color: #fff; margin-bottom: 12px; }
        .login-btn { width: 100%; background: linear-gradient(135deg, var(--blue2), var(--blue3)); color: #fff; border: none; padding: 12px; border-radius: 8px; font-weight: 600; cursor: pointer; }
        .demo-btns { display: flex; gap: 8px; margin-top: 20px; flex-wrap: wrap; }
        .demo-btn { flex: 1; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); border-radius: 7px; padding: 7px; font-size: 11px; cursor: pointer; text-align: center; color: rgba(255,255,255,0.7); }
        #login-error { color: #EF9A9A; font-size: 12px; margin-top: 12px; text-align: center; }

        /* ========== NAVIGATION & LAYOUT (responsive) ========== */
        #app-shell { display: flex; min-height: calc(100vh - 58px); opacity: 0; transition: 0.3s; flex-direction: column; }
        #app-shell.ready { opacity: 1; }
        #nav { background: var(--navy2); border-bottom: 1px solid rgba(255,255,255,0.06); padding: 0 12px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; min-height: 58px; position: sticky; top: 0; z-index: 200; }
        .logo-wrap { display: flex; align-items: center; gap: 8px; }
        .logo-icon { width: 32px; height: 32px; background: linear-gradient(135deg, var(--blue2), var(--blue3)); border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 14px; }
        .logo-text { font-size: 14px; font-weight: 700; color: #fff; }
        .logo-form { font-family: var(--mono); font-size: 9px; color: var(--gray3); }
        #nav-tabs { display: flex; gap: 2px; background: rgba(255,255,255,0.04); border-radius: 8px; padding: 3px; flex-wrap: wrap; justify-content: center; }
        .ntab { background: transparent; border: none; color: var(--gray3); padding: 5px 12px; border-radius: 6px; cursor: pointer; font-size: 12px; font-weight: 500; }
        .ntab.active { background: var(--blue2); color: #fff; }
        #nav-right { display: flex; align-items: center; gap: 4px; flex-wrap: wrap; }
        .user-pill { display: flex; align-items: center; gap: 7px; background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.1); border-radius: 20px; padding: 4px 10px; }
        .user-avatar { width: 24px; height: 24px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 10px; font-weight: 700; color: #fff; background: var(--blue2); }
        .nbtn { background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); color: rgba(255,255,255,0.7); padding: 4px 8px; border-radius: 7px; cursor: pointer; font-size: 10px; transition: 0.15s; white-space: nowrap; }
        .nbtn.primary { background: var(--blue2); border-color: var(--blue2); color: #fff; }
        .nbtn.danger { background: rgba(183,28,28,0.15); border-color: rgba(183,28,28,0.3); color: #EF9A9A; }
        @media (max-width: 768px) {
            .main-layout { flex-direction: column; }
            #sidebar { width: 100%; max-height: 160px; border-right: none; border-bottom: 1px solid rgba(255,255,255,0.05); display: flex; flex-wrap: wrap; gap: 6px; padding: 12px; overflow-x: auto; }
            .si { display: inline-flex; width: auto; margin-right: 4px; }
            #main { padding: 12px; }
            .kpi-grid { grid-template-columns: repeat(2, 1fr); }
            .field-grid { grid-template-columns: 1fr; }
            .ic-top { flex-direction: column; }
            .rb-row { align-self: flex-start; }
            .xcol { min-width: 100%; }
            .ai-panel { width: 280px; right: 12px; bottom: 80px; }
        }
        .main-layout { display: flex; flex: 1; overflow: hidden; flex-direction: row; }
        #sidebar { width: 220px; background: var(--navy3); flex-shrink: 0; border-right: 1px solid rgba(255,255,255,0.05); padding: 16px 8px; overflow-y: auto; }
        .si { display: flex; align-items: center; gap: 6px; padding: 6px 8px; border-radius: 7px; cursor: pointer; margin-bottom: 1px; color: #c5d3e8; font-size: 11.5px; }
        .si-badge { font-size: 9px; padding: 2px 6px; border-radius: 8px; font-weight: 700; margin-left: auto; }
        .bp { background: rgba(46,125,50,0.25); color: #81C784; }
        .bf { background: rgba(183,28,28,0.25); color: #EF9A9A; }
        #main { flex: 1; overflow-y: auto; background: var(--gray0); padding: 22px; }
        .page { display: none; }
        .page.active { display: block; }
        .card { background: #fff; border-radius: var(--r-lg); padding: 20px; box-shadow: var(--sh); border: 1px solid var(--gray2); margin-bottom: 18px; }
        .field-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 10px; }
        .fc { background: #fff; border-radius: var(--r-lg); padding: 12px; border: 1px solid var(--gray2); }
        .fc label { font-size: 10px; text-transform: uppercase; color: var(--gray4); font-weight: 600; display: block; margin-bottom: 5px; }
        .fc input, .fc select, .fc textarea { width: 100%; border: 1.5px solid var(--gray2); border-radius: 7px; padding: 7px 9px; font-size: 13px; }

        /* Inspection items */
        .ic { background: #fff; border-radius: var(--r); margin-bottom: 8px; border: 1px solid var(--gray2); border-left: 4px solid var(--gray2); }
        .ic.ok { border-left-color: var(--green2); }
        .ic.nok { border-left-color: var(--red2); }
        .ic-top { display: flex; gap: 10px; padding: 11px 13px; align-items: flex-start; flex-wrap: wrap; }
        .ic-id { font-family: var(--mono); font-size: 10px; font-weight: 700; color: var(--blue2); background: var(--blue-light); padding: 3px 7px; border-radius: 5px; }
        .ic-body { flex: 1; }
        .ic-desc-editable, .ic-crit-editable { font-size: 13px; font-weight: 600; color: var(--gray7); cursor: pointer; display: inline-block; border-bottom: 1px dashed var(--blue2); }
        .ic-crit-editable { font-size: 11.5px; font-weight: normal; display: block; margin-top: 4px; }
        .rb-row { display: flex; gap: 4px; flex-wrap: wrap; }
        .rb { border: 1.5px solid; padding: 4px 10px; border-radius: 6px; cursor: pointer; font-size: 11px; font-weight: 700; background: transparent; }
        .rb.ok-btn { border-color: var(--green2); color: var(--green2); }
        .rb.ok-btn.act { background: var(--green2); color: #fff; }
        .rb.nok-btn { border-color: var(--red2); color: var(--red2); }
        .rb.nok-btn.act { background: var(--red2); color: #fff; }
        .ic-extra { display: none; padding: 0 13px 13px; gap: 10px; flex-wrap: wrap; border-top: 1px solid var(--gray1); }
        .ic-extra.show { display: flex; }
        .xcol { flex: 1; min-width: 130px; }
        .xcol label { font-size: 10px; text-transform: uppercase; color: var(--gray4); font-weight: 600; display: block; }
        .xcol select, .xcol input, .xcol textarea { width: 100%; border: 1.5px solid var(--gray2); border-radius: 7px; padding: 5px 8px; font-size: 12px; }
        .photo-strip { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; }
        .photo-thumb { width: 65px; height: 65px; object-fit: cover; border-radius: 6px; border: 2px solid var(--gray2); cursor: pointer; }
        .add-photo { width: 65px; height: 65px; border: 2px dashed var(--gray2); border-radius: 6px; display: flex; align-items: center; justify-content: center; font-size: 20px; cursor: pointer; }

        /* Dashboard */
        .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 20px; }
        .kpi-card { background: #fff; border-radius: var(--r-lg); padding: 16px; text-align: center; border: 1px solid var(--gray2); }
        .kpi-value { font-size: 28px; font-weight: 800; font-family: var(--mono); }

        /* Round AI Assistant */
        .ai-round-icon { position: fixed; bottom: 24px; right: 24px; width: 56px; height: 56px; background: var(--blue2); border-radius: 50%; display: flex; align-items: center; justify-content: center; cursor: pointer; box-shadow: 0 4px 12px rgba(0,0,0,0.3); z-index: 150; transition: transform 0.2s; }
        .ai-round-icon:hover { transform: scale(1.05); }
        .ai-round-icon span { font-size: 28px; }
        .ai-panel { position: fixed; bottom: 90px; right: 24px; width: 320px; max-width: calc(100vw - 40px); background: var(--navy2); border-radius: 16px; border: 1px solid rgba(255,255,255,0.1); box-shadow: 0 8px 24px rgba(0,0,0,0.3); z-index: 151; backdrop-filter: blur(8px); display: none; flex-direction: column; }
        .ai-panel.open { display: flex; }
        .ai-header { padding: 12px 16px; background: var(--blue2); border-radius: 16px 16px 0 0; display: flex; justify-content: space-between; font-weight: 700; cursor: pointer; }
        .ai-close { background: none; border: none; color: white; font-size: 18px; cursor: pointer; }
        .ai-messages { max-height: 300px; overflow-y: auto; padding: 12px; font-size: 12px; }
        .ai-message { margin-bottom: 8px; padding: 6px 10px; border-radius: 12px; background: #1E293B; color: #e2e8f0; }
        .ai-message.user { background: var(--blue2); text-align: right; }
        .ai-input { display: flex; gap: 8px; padding: 8px 12px; border-top: 1px solid rgba(255,255,255,0.1); }
        .ai-input input { flex: 1; background: #0F172A; border: 1px solid #334155; border-radius: 20px; padding: 8px 12px; color: white; }
        .ai-send { background: var(--blue2); border: none; border-radius: 20px; padding: 8px 12px; cursor: pointer; color: white; }

        .toast { position: fixed; bottom: 20px; left: 20px; background: #1E293B; padding: 8px 16px; border-radius: 40px; transform: translateY(60px); opacity: 0; transition: 0.2s; z-index: 200; color: white; }
        .toast.show { transform: translateY(0); opacity: 1; }
        .status-badge { font-size: 9px; padding: 2px 6px; border-radius: 10px; margin-left: 6px; }
        .status-approved { background: rgba(46,125,50,0.3); color: #81C784; }
        .status-pending { background: rgba(245,158,11,0.3); color: #FFCC80; }
        .status-rejected { background: rgba(183,28,28,0.3); color: #EF9A9A; }
        .theme-selector { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 12px; }
        .theme-btn { background: var(--gray1); border: 1px solid var(--gray2); border-radius: 20px; padding: 4px 12px; cursor: pointer; font-size: 11px; transition: 0.1s; }
        .theme-btn:hover { background: var(--blue2); color: white; }
    </style>
</head>
<body>

<!-- LOGIN SCREEN -->
<div id="login-screen">
    <div class="login-box">
        <div class="login-logo"><div class="login-icon">⚡</div><div><div class="login-title">GW81 Inspector Pro</div><div class="login-sub">Admin User Approval · Multi‑Format Export</div></div></div>
        <div class="login-field"><input id="li-user" placeholder="Username"></div>
        <div class="login-field"><input id="li-pass" type="password" placeholder="Password"></div>
        <button class="login-btn" onclick="doLogin()">Sign In →</button>
        <div class="demo-btns">
            <div class="demo-btn" onclick="quickLogin('admin','admin123')">👑 Admin</div>
            <div class="demo-btn" onclick="quickLogin('inspector','insp123')">🔍 Inspector</div>
            <div class="demo-btn" onclick="quickLogin('viewer','view123')">👁 Viewer</div>
        </div>
        <div id="login-error"></div>
    </div>
</div>

<!-- MAIN APP -->
<nav id="nav" style="display:none">
    <div class="logo-wrap"><div class="logo-icon">⚡</div><div><div class="logo-text">GW81 Inspector Pro</div><div class="logo-form">F05_I60_GAE07039</div></div></div>
    <div id="nav-tabs">
        <button class="ntab active" onclick="showPage('cover')">📋 Cover</button>
        <button class="ntab" onclick="showPage('inspection')">🔍 Inspection</button>
        <button class="ntab" onclick="showPage('dashboard')">📊 Dashboard</button>
        <button class="ntab" id="tab-admin" onclick="showPage('admin')" style="display:none">⚙ Admin</button>
    </div>
    <div id="nav-right">
        <div class="user-pill"><div class="user-avatar" id="user-av"></div><div><div class="user-name" id="user-name"></div><div class="user-role" id="user-role"></div></div></div>
        <button class="nbtn primary" onclick="exportToExcel()">📊 Excel</button>
        <button class="nbtn primary" onclick="exportToPDF()">📄 PDF</button>
        <button class="nbtn primary" onclick="exportToWord()">📝 Word</button>
        <button class="nbtn" onclick="document.getElementById('imp').click()">📂 Import</button>
        <input type="file" id="imp" accept=".json" style="display:none" onchange="importJSON(event)">
        <button class="nbtn" onclick="window.print()">🖨 Print</button>
        <button class="nbtn danger" onclick="doLogout()">Sign Out</button>
    </div>
</nav>

<div id="app-shell" style="display:none">
    <div class="main-layout">
        <div id="sidebar"></div>
        <div id="main">
            <div id="page-cover" class="page active">
                <div class="card"><h2>📋 Inspection Certificate</h2><div class="field-grid"><div class="fc"><label>Manufacturer</label><input id="f-mfr" oninput="saveCover()"></div><div class="fc"><label>Blade Model</label><input id="f-model" value="GW81" oninput="saveCover()"></div><div class="fc"><label>Serial Number</label><input id="f-serial" oninput="saveCover()"></div><div class="fc"><label>Inspection Date</label><input type="date" id="f-date" oninput="saveCover()"></div><div class="fc"><label>Inspector Name</label><input id="f-inspector" oninput="saveCover()"></div><div class="fc"><label>Site</label><input id="f-site" oninput="saveCover()"></div><div class="fc"><label>Control Number</label><input id="f-ctrl" oninput="saveCover()"></div><div class="fc"><label>Approved By</label><input id="f-approved" oninput="saveCover()"></div></div><div class="fc"><label>Observations</label><textarea id="f-notes" rows="2" oninput="saveCover()"></textarea></div></div>
            </div>
            <div id="page-inspection" class="page"><div id="insp-content"></div></div>
            <div id="page-dashboard" class="page"><div class="kpi-grid" id="kpi-grid"></div><div class="card"><canvas id="dashboard-chart" height="180"></canvas></div><div class="card"><h3>Section Summary</h3><table id="summary-table" style="width:100%"><thead><tr><th>Section</th><th>Total</th><th>OK</th><th>NOK</th><th>Status</th></tr></thead><tbody></tbody></table></div><div class="card"><h3>⚠️ Defect Register</h3><table id="defect-table" style="width:100%"><thead><tr><th>Item</th><th>Description</th><th>Severity</th><th>Comment</th></tr></thead><tbody></tbody></table></div></div>
            <div id="page-admin" class="page">
                <div class="card"><h3>👥 User Management & Approval</h3><div id="user-list-admin"></div><div style="margin-top:12px;"><button class="nbtn primary" onclick="showAddUserForm()">+ Add New User</button></div><div id="add-user-form" style="display:none; margin-top:12px; background:var(--gray1); padding:12px; border-radius:12px;"><input id="new-username" placeholder="Username" style="width:100%; margin-bottom:6px;"><input id="new-password" type="password" placeholder="Password" style="width:100%; margin-bottom:6px;"><select id="new-role"><option value="inspector">Inspector</option><option value="viewer">Viewer</option></select><div style="margin-top:8px;"><button class="nbtn primary" onclick="addPendingUser()">Request Approval</button><button class="nbtn" onclick="hideAddUserForm()">Cancel</button></div></div></div>
                <div class="card"><h3>📝 Pending Edit Requests</h3><div id="pending-edits-list" style="background:var(--gray0); border-radius:12px; padding:12px;"></div></div>
                <div class="card"><h3>🎨 Theme Settings</h3><div class="theme-selector"><div class="theme-btn" onclick="setTheme('dark')">🌙 Dark</div><div class="theme-btn" onclick="setTheme('light')">☀️ Light</div><div class="theme-btn" onclick="setTheme('forest')">🌲 Forest</div><div class="theme-btn" onclick="setTheme('ocean')">🌊 Ocean</div><div class="theme-btn" onclick="setTheme('sunset')">🌅 Sunset</div></div></div>
                <div class="card"><h3>📊 Export Reports</h3><div style="display:flex; gap:10px; flex-wrap:wrap"><button class="nbtn primary" onclick="exportToExcel()">📊 Excel (XLSX)</button><button class="nbtn primary" onclick="exportToPDF()">📄 PDF</button><button class="nbtn primary" onclick="exportToWord()">📝 Word (DOCX)</button><button class="nbtn" onclick="exportJSON()">💾 Backup JSON</button><button class="nbtn" onclick="document.getElementById('imp-admin').click()">📂 Import JSON</button><input type="file" id="imp-admin" style="display:none" accept=".json" onchange="importJSON(event)"><button class="nbtn danger" onclick="resetAllData()">🗑 Reset All</button></div></div>
            </div>
        </div>
    </div>
</div>

<!-- ROUND AI ASSISTANT -->
<div class="ai-round-icon" onclick="toggleAIPanel()"><span>🤖</span></div>
<div class="ai-panel" id="ai-panel">
    <div class="ai-header"><span>🤖 AI Assistant</span><button class="ai-close" onclick="toggleAIPanel()">✕</button></div>
    <div class="ai-messages" id="ai-messages"><div class="ai-message">Ask about severity, studs, or inspection criteria.</div></div>
    <div class="ai-input"><input type="text" id="ai-query" placeholder="Ask..."><button class="ai-send" onclick="askAI()">Send</button></div>
</div>
<div id="toast" class="toast"></div>

<script>
    // ==================== BASE DATA ====================
    const BASE_SECTIONS = [
        {num:"0",name:"Documentation",items:[{id:"101",desc:"Check correct identification of the blade",criteria:"External label with origin, model, serial number and weight.",ctrl:"DOC-001"},{id:"102",desc:"Check correct internal identification",criteria:"Label with critical components traceability.",ctrl:"DOC-002"}]},
        {num:"1",name:"Studs / Inserts",items:[
            {id:"10101",desc:"Check traces of rust on the surface of the studs",criteria:"Check visually 100% of studs for signs of oxide",ctrl:"STU-001"},
            {id:"10102",desc:"Check damages on the threads of the studs",criteria:"Check visually 100% of studs for signs of hits or deformations",ctrl:"STU-002"},
            {id:"10103",desc:"Check studs traceability",criteria:"Check if at least one stud has the batch number engraved.",ctrl:"STU-003"},
            {id:"10104",desc:"Check length of the studs",criteria:"Measure with caliper: distance must be per spec.",ctrl:"STU-004",hasMeasure:true,measureUnit:"mm"},
            {id:"10105",desc:"Check tightening torque of studs",criteria:"Check at least 4 studs turning manually. The studs cannot turn freely.",ctrl:"STU-005"},
            {id:"10106",desc:"Check surface of the inserts",criteria:"No rests of consumables, tapes or resin.",ctrl:"STU-006"},
            {id:"10107",desc:"Check zero-degree mark and red stud",criteria:"Zero-degree mark engraved and red stud identified.",ctrl:"STU-007"},
            {id:"10108",desc:"Check plate with the reference to the Zero-degree position",criteria:"Placed next to the zero-degree stud.",ctrl:"STU-008"},
            {id:"10109",desc:"Check rubber ring between studs and inserts",criteria:"100% correctly located.",ctrl:"STU-009"},
            {id:"10110",desc:"Check Inserts protection against rust",criteria:"Plastic caps present.",ctrl:"STU-010"}
        ]},
        {num:"2",name:"Lightning Protection System",items:[
            {id:"20101",desc:"Check lightning card",criteria:"Placed and legible.",ctrl:"LPS-001"},
            {id:"20102",desc:"Check fixing of LPS cable",criteria:"Tense, no hanging.",ctrl:"LPS-002"},
            {id:"20103",desc:"Check terminals of lightning cables",criteria:"Crimping correct.",ctrl:"LPS-003"},
            {id:"20104",desc:"Check LPS plate",criteria:"Fixed and undamaged.",ctrl:"LPS-004"},
            {id:"20105",desc:"Measure terminals conductivity",criteria:"≤50 mΩ",hasMeasure:true,measureUnit:"mΩ",ctrl:"LPS-005"},
            {id:"20106",desc:"Measure conductivity plate",criteria:"≤5 mΩ",hasMeasure:true,measureUnit:"mΩ",ctrl:"LPS-006"},
            {id:"20107",desc:"Check tip receptor",criteria:"Smooth transition.",ctrl:"LPS-007"},
            {id:"20108",desc:"Check side receptors",criteria:"Correct protrusion & paint.",ctrl:"LPS-008"},
            {id:"20109",desc:"Check drainage hole",criteria:"Free from obstruction.",ctrl:"LPS-009"}
        ]}
    ];

    let state = { cover: {}, results: {}, details: {}, photos: {}, sections: null, pendingEdits: [] };
    let currentUser = null;
    let users = JSON.parse(localStorage.getItem('gw81_users')) || [
        { username: 'admin', password: 'admin123', role: 'admin', name: 'Admin', status: 'approved' },
        { username: 'inspector', password: 'insp123', role: 'inspector', name: 'Inspector', status: 'approved' },
        { username: 'viewer', password: 'view123', role: 'viewer', name: 'Viewer', status: 'approved' }
    ];
    users = users.map(u => ({ ...u, status: u.status || 'approved' }));
    let chart = null;

    function saveState() { localStorage.setItem('gw81_pro', JSON.stringify({ cover: state.cover, results: state.results, details: state.details, photos: state.photos, sections: state.sections, pendingEdits: state.pendingEdits })); localStorage.setItem('gw81_users', JSON.stringify(users)); }
    function loadState() { const d = localStorage.getItem('gw81_pro'); if(d) { const p = JSON.parse(d); state.cover = p.cover || {}; state.results = p.results || {}; state.details = p.details || {}; state.photos = p.photos || {}; state.sections = p.sections || null; state.pendingEdits = p.pendingEdits || []; } if(!state.sections) state.sections = JSON.parse(JSON.stringify(BASE_SECTIONS)); }
    function isAdmin() { return currentUser && currentUser.role === 'admin' && currentUser.status === 'approved'; }
    function canEdit() { return currentUser && (currentUser.role === 'admin' || currentUser.role === 'inspector') && currentUser.status === 'approved'; }

    // ========== THEME ==========
    function setTheme(theme) { document.body.className = theme; localStorage.setItem('gw81_theme', theme); toast(`Theme: ${theme}`); }
    const savedTheme = localStorage.getItem('gw81_theme'); if(savedTheme) document.body.className = savedTheme;

    // ========== LOGIN ==========
    function doLogin() {
        const u = document.getElementById('li-user').value.trim();
        const p = document.getElementById('li-pass').value;
        const user = users.find(x => x.username === u && x.password === p);
        if (!user) { document.getElementById('login-error').innerText = 'Invalid credentials'; return; }
        if (!user.status) user.status = 'approved';
        if (user.status !== 'approved') { document.getElementById('login-error').innerText = `Account ${user.status}. Contact admin.`; return; }
        currentUser = user;
        document.getElementById('login-screen').style.display = 'none';
        document.getElementById('nav').style.display = 'flex';
        document.getElementById('app-shell').style.display = 'flex';
        setTimeout(() => document.getElementById('app-shell').classList.add('ready'), 50);
        loadState();
        document.getElementById('user-av').innerText = user.name.slice(0,2).toUpperCase();
        document.getElementById('user-name').innerText = user.name;
        document.getElementById('user-role').innerText = user.role;
        const adminTab = document.getElementById('tab-admin');
        if (isAdmin()) adminTab.style.display = 'inline-block';
        else adminTab.style.display = 'none';
        buildSidebar();
        buildInspection();
        renderCover();
        refreshDashboard();
        attachTabEvents();
        if (isAdmin()) buildAdminPanel();
        toast(`Welcome ${user.name}`);
    }
    function quickLogin(u,p) { document.getElementById('li-user').value = u; document.getElementById('li-pass').value = p; doLogin(); }
    function doLogout() { currentUser = null; document.getElementById('login-screen').style.display = 'flex'; document.getElementById('nav').style.display = 'none'; document.getElementById('app-shell').style.display = 'none'; document.getElementById('app-shell').classList.remove('ready'); }

    // ========== ADMIN PANEL (complete) ==========
    function buildAdminPanel() {
        if (!isAdmin()) return;
        const userDiv = document.getElementById('user-list-admin');
        userDiv.innerHTML = users.map((u, idx) => `
            <div style="display:flex; justify-content:space-between; align-items:center; background:var(--gray0); padding:10px; border-radius:12px; margin-bottom:8px;">
                <div><strong>${u.name}</strong> (${u.role})<br><span style="font-size:10px;">${u.username}</span></div>
                <div><span class="status-badge ${u.status === 'approved' ? 'status-approved' : (u.status === 'pending' ? 'status-pending' : 'status-rejected')}">${u.status.toUpperCase()}</span></div>
                <div>
                    ${u.status === 'pending' ? `<button class="nbtn primary" style="margin-right:4px;" onclick="approveUser(${idx})">Approve</button><button class="nbtn danger" onclick="rejectUser(${idx})">Reject</button>` : ''}
                    ${u.username !== 'admin' ? `<button class="nbtn danger" onclick="deleteUser(${idx})">Delete</button>` : ''}
                </div>
            </div>
        `).join('');
        const pendingDiv = document.getElementById('pending-edits-list');
        if (state.pendingEdits.length === 0) pendingDiv.innerHTML = '<div style="padding:12px; text-align:center;">No pending edit requests</div>';
        else pendingDiv.innerHTML = state.pendingEdits.map(req => `
            <div style="background:var(--gray0); padding:10px; border-radius:8px; margin-bottom:8px;">
                <div><strong>${req.itemId}</strong> — ${req.field}: "${req.oldValue}" → "${req.newValue}"</div>
                <div style="font-size:11px; color:var(--gray4);">Requested by ${req.requester}</div>
                <div style="margin-top:6px;">
                    <button class="nbtn primary" onclick="approveEdit(${req.id})">Approve</button>
                    <button class="nbtn danger" onclick="rejectEdit(${req.id})">Reject</button>
                </div>
            </div>
        `).join('');
    }
    function showAddUserForm() { document.getElementById('add-user-form').style.display = 'block'; }
    function hideAddUserForm() { document.getElementById('add-user-form').style.display = 'none'; }
    function addPendingUser() {
        const u = document.getElementById('new-username').value.trim();
        const p = document.getElementById('new-password').value;
        const r = document.getElementById('new-role').value;
        if (!u || !p) { toast('Username and password required'); return; }
        if (users.find(x => x.username === u)) { toast('Username exists'); return; }
        users.push({ username: u, password: p, role: r, name: u.charAt(0).toUpperCase()+u.slice(1), status: 'pending' });
        saveState();
        buildAdminPanel();
        hideAddUserForm();
        toast(`User ${u} added (pending approval)`);
    }
    function approveUser(idx) { users[idx].status = 'approved'; saveState(); buildAdminPanel(); toast(`User ${users[idx].username} approved`); }
    function rejectUser(idx) { users[idx].status = 'rejected'; saveState(); buildAdminPanel(); toast(`User ${users[idx].username} rejected`); }
    function deleteUser(idx) { if (users[idx].username === 'admin') { toast('Cannot delete admin'); return; } users.splice(idx,1); saveState(); buildAdminPanel(); toast('User deleted'); }

    // ========== EDIT REQUESTS ==========
    function requestEditApproval(sectionIdx, itemIdx, field, newValue, oldValue, itemId) {
        state.pendingEdits.push({ id: Date.now(), sectionIdx, itemIdx, field, newValue, oldValue, itemId, requester: currentUser.name });
        saveState();
        if (isAdmin()) buildAdminPanel();
        toast(`Edit request submitted for ${itemId}. Awaiting admin approval.`);
    }
    function approveEdit(reqId) {
        const req = state.pendingEdits.find(r => r.id === reqId);
        if (req) {
            state.sections[req.sectionIdx].items[req.itemIdx][req.field] = req.newValue;
            state.pendingEdits = state.pendingEdits.filter(r => r.id !== reqId);
            saveState();
            buildInspection();
            if (isAdmin()) buildAdminPanel();
            toast(`Approved: ${req.itemId} ${req.field} updated`);
        }
    }
    function rejectEdit(reqId) {
        state.pendingEdits = state.pendingEdits.filter(r => r.id !== reqId);
        saveState();
        if (isAdmin()) buildAdminPanel();
        toast('Edit request rejected');
    }

    // ========== UI BUILDING ==========
    function buildSidebar() {
        const sidebar = document.getElementById('sidebar');
        sidebar.innerHTML = state.sections.map((s,i)=>`<div class="si" onclick="jumpToSection(${i})"><span>📌 ${s.num}. ${s.name}</span><span class="si-badge" id="sb-${i}">—</span></div>`).join('');
        updateBadges();
    }
    function updateBadges() { state.sections.forEach((sec,i)=>{ let nok = sec.items.filter(it=>state.results[it.id]==='nok').length; let b = document.getElementById(`sb-${i}`); if(b) b.innerText = nok ? `✗${nok}` : '✓'; b.className = `si-badge ${nok?'bf':'bp'}`; }); }
    function jumpToSection(si) { showPage('inspection'); setTimeout(()=>document.getElementById(`sec-${si}`)?.scrollIntoView({behavior:'smooth'}), 100); }

    function buildInspection() {
        const cont = document.getElementById('insp-content'); if(!cont) return; cont.innerHTML = '';
        state.sections.forEach((sec, si) => {
            const secDiv = document.createElement('div'); secDiv.id = `sec-${si}`;
            secDiv.innerHTML = `<div style="background:#1E293B; padding:12px; border-radius:16px; margin-bottom:12px"><h3>Section ${sec.num}: ${sec.name}</h3></div><div id="items-${si}"></div>`;
            cont.appendChild(secDiv);
            const itemsDiv = document.getElementById(`items-${si}`);
            sec.items.forEach(it => {
                const res = state.results[it.id] || '';
                const details = state.details[it.id] || {};
                itemsDiv.innerHTML += `
                    <div class="ic ${res}" id="ic-${it.id}">
                        <div class="ic-top">
                            <div><div class="ic-id">${it.id}</div><div style="font-size:9px; color:#64748B">${it.ctrl||''}</div></div>
                            <div class="ic-body">
                                <div class="ic-desc-editable" onclick="editField('${it.id}','desc','${escapeHtml(it.desc)}')">${escapeHtml(it.desc)}</div>
                                <div class="ic-crit-editable" onclick="editField('${it.id}','criteria','${escapeHtml(it.criteria)}')">${escapeHtml(it.criteria)}</div>
                            </div>
                            <div class="rb-row">
                                <button class="rb ok-btn ${res==='ok'?'act':''}" onclick="setResult('${it.id}','ok')">OK</button>
                                <button class="rb nok-btn ${res==='nok'?'act':''}" onclick="setResult('${it.id}','nok')">NOK</button>
                                <button class="rb ${res==='na'?'act':''}" style="border-color:#64748B;color:#64748B" onclick="setResult('${it.id}','na')">N/A</button>
                            </div>
                        </div>
                        <div class="ic-extra ${res?'show':''}" id="extra-${it.id}">
                            <div class="xcol"><label>Measurement</label><input value="${details.measurement||''}" onchange="updateDetail('${it.id}','measurement',this.value)"></div>
                            <div class="xcol"><label>Severity (L/M/H)</label><select onchange="updateDetail('${it.id}','severity',this.value)"><option value="">—</option><option ${details.severity==='L'?'selected':''} value="L">L — Low (cosmetic, monitor)</option><option ${details.severity==='M'?'selected':''} value="M">M — Medium (repair within 3 months)</option><option ${details.severity==='H'?'selected':''} value="H">H — High (immediate repair)</option></select></div>
                            <div class="xcol wide"><label>Comment</label><textarea onchange="updateDetail('${it.id}','comment',this.value)">${details.comment||''}</textarea></div>
                            <div class="photo-strip" id="photostrip-${it.id}">${renderPhotos(it.id)}${canEdit()?`<div class="add-photo" onclick="document.getElementById('file-${it.id}').click()">+</div><input type="file" id="file-${it.id}" multiple accept="image/*" style="display:none" onchange="attachPhotos('${it.id}',this)">`:''}</div>
                        </div>
                    </div>`;
            });
        });
        updateBadges();
    }
    function renderPhotos(id) { return (state.photos[id]||[]).slice(0,6).map((p,i)=>`<img class="photo-thumb" src="${p.data}" onclick="window.open('${p.data}')"><button style="background:#EF4444; border:none; border-radius:20px; padding:0 4px; margin-left:-20px; margin-top:-8px;" onclick="deletePhoto('${id}',${i})">✕</button>`).join(''); }
    function attachPhotos(id, inp) { const files = Array.from(inp.files); if((state.photos[id]?.length||0)+files.length>6){ toast('Max 6 photos'); return; } files.forEach(f=>{ const rd=new FileReader(); rd.onload=e=>{ if(!state.photos[id]) state.photos[id]=[]; state.photos[id].push({data:e.target.result,name:f.name}); saveState(); refreshPhotoStrip(id); }; rd.readAsDataURL(f); }); inp.value=''; }
    function deletePhoto(id,idx) { state.photos[id].splice(idx,1); saveState(); refreshPhotoStrip(id); }
    function refreshPhotoStrip(id) { const strip = document.getElementById(`photostrip-${id}`); if(strip) strip.innerHTML = renderPhotos(id)+(canEdit()?`<div class="add-photo" onclick="document.getElementById('file-${id}').click()">+</div><input type="file" id="file-${id}" multiple accept="image/*" style="display:none" onchange="attachPhotos('${it.id}',this)">`:''); }
    function setResult(id,res) { state.results[id]=res; if(!state.details[id]) state.details[id]={}; if(res!=='nok') state.details[id].severity=''; saveState(); buildInspection(); refreshDashboard(); toast(`Item ${id} ${res.toUpperCase()}`); }
    function updateDetail(id,field,val) { if(!state.details[id]) state.details[id]={}; state.details[id][field]=val; saveState(); }
    function refreshDashboard() {
        let total=0,ok=0,nok=0; state.sections.forEach(sec=>sec.items.forEach(it=>{ total++; const r=state.results[it.id]; if(r==='ok') ok++; else if(r==='nok') nok++; }));
        document.getElementById('kpi-grid').innerHTML = `<div class="kpi-card"><div class="kpi-value">${total}</div><div>Total</div></div><div class="kpi-card"><div class="kpi-value" style="color:#2E7D32">${ok}</div><div>OK</div></div><div class="kpi-card"><div class="kpi-value" style="color:#C62828">${nok}</div><div>NOK</div></div><div class="kpi-card"><div class="kpi-value">${Math.round((ok+nok)/total*100)}%</div><div>Complete</div></div>`;
        if(chart) chart.destroy(); chart = new Chart(document.getElementById('dashboard-chart'), { type: 'bar', data: { labels: ['OK','NOK'], datasets: [{ data: [ok,nok], backgroundColor: ['#2E7D32','#C62828'] }] } });
        const tbody = document.querySelector('#summary-table tbody'); tbody.innerHTML = state.sections.map(sec=>{ let okc=0,nokc=0; sec.items.forEach(it=>{if(state.results[it.id]==='ok') okc++; else if(state.results[it.id]==='nok') nokc++;}); return `<tr><td style="font-family:monospace">${sec.name}</td><td style="font-family:monospace">${sec.items.length}</td><td style="color:#2E7D32">${okc}</td><td style="color:#C62828">${nokc}</td><td>${nokc>0?'⚠️ Issues':'✅ Pass'}</td></tr>`; }).join('');
        const defectBody = document.querySelector('#defect-table tbody'); const defects=[]; state.sections.forEach(sec=>sec.items.forEach(it=>{ if(state.results[it.id]==='nok') defects.push({id:it.id,desc:it.desc,sev:state.details[it.id]?.severity||'-',comment:state.details[it.id]?.comment||''}); })); defectBody.innerHTML = defects.map(d=>`<tr><td style="font-family:monospace">${d.id}</td><td style="font-family:monospace">${d.desc}</td><td style="font-family:monospace">${d.sev}</td><td style="font-family:monospace">${d.comment}</td></tr>`).join('');
    }
    function editField(itemId, field, oldValue) {
        if (!canEdit()) { toast('Only approved inspectors/admins can edit'); return; }
        const newVal = prompt(`Edit ${field} for item ${itemId}:`, oldValue);
        if (newVal && newVal !== oldValue) {
            const secIdx = state.sections.findIndex(s => s.items.some(i => i.id === itemId));
            const itemIdx = state.sections[secIdx].items.findIndex(i => i.id === itemId);
            if (isAdmin()) {
                state.sections[secIdx].items[itemIdx][field] = newVal;
                saveState();
                buildInspection();
                toast(`${field} updated directly by admin`);
            } else {
                requestEditApproval(secIdx, itemIdx, field, newVal, oldValue, itemId);
            }
        }
    }

    // ========== EXPORT / IMPORT ==========
    function getExportRows() { let rows = [['Item ID','Ctrl','Description','Criteria','Result','Severity','Measurement','Comment','Photos']]; state.sections.forEach(sec=>{ sec.items.forEach(it=>{ rows.push([it.id,it.ctrl||'',it.desc,it.criteria,state.results[it.id]||'Pending',state.details[it.id]?.severity||'',state.details[it.id]?.measurement||'',state.details[it.id]?.comment||'',(state.photos[it.id]||[]).length]); }); }); return rows; }
    function exportToExcel() { const ws = XLSX.utils.aoa_to_sheet(getExportRows()); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Inspection'); XLSX.writeFile(wb, `GW81_${state.cover.ctrl||'report'}.xlsx`); toast('Excel exported'); }
    async function exportToPDF() { const { jsPDF } = window.jspdf; const doc = new jsPDF({ orientation: 'landscape' }); let y=20; doc.setFontSize(14); doc.text(`GW81 Report - ${state.cover.ctrl||'Report'}`,20,y); y+=10; doc.setFontSize(8); const rows=getExportRows(); const headers=rows[0], dataRows=rows.slice(1); if(dataRows.length===0) doc.text('No data',20,y); else { const widths=[15,20,40,45,12,10,20,45,15]; let x=10; doc.setFillColor(200,200,200); doc.rect(x,y,widths.reduce((a,b)=>a+b,0),6,'F'); doc.setTextColor(0,0,0); let xx=x; headers.forEach((h,i)=>{ doc.text(h,xx+1,y+4); xx+=widths[i]; }); y+=8; for(const row of dataRows){ if(y>190){ doc.addPage(); y=20; } xx=x; for(let i=0;i<row.length;i++){ doc.text(String(row[i]),xx+1,y+3); xx+=widths[i]; } y+=6; } } doc.save(`GW81_${state.cover.ctrl||'report'}.pdf`); toast('PDF exported'); }
    function exportToWord() { let html = `<html><head><meta charset="UTF-8"><title>GW81 Report</title></head><body><h1>GW81 Inspection Report</h1><table border="1"><thead><tr><th>Item ID</th><th>Ctrl</th><th>Description</th><th>Criteria</th><th>Result</th><th>Severity</th><th>Measurement</th><th>Comment</th><th>Photos</th></tr></thead><tbody>`; state.sections.forEach(sec=>{ sec.items.forEach(it=>{ html+=`<tr><td>${it.id}</td><td>${it.ctrl||''}</td><td>${escapeHtml(it.desc)}</td><td>${escapeHtml(it.criteria)}</td><td>${state.results[it.id]||'Pending'}</td><td>${state.details[it.id]?.severity||''}</td><td>${state.details[it.id]?.measurement||''}</td><td>${escapeHtml(state.details[it.id]?.comment||'')}</td><td>${(state.photos[it.id]||[]).length}</td></tr>`; }); }); html+=`</tbody></table></body></html>`; const blob = new Blob([html], {type:'application/msword'}); const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `GW81_${state.cover.ctrl||'report'}.doc`; a.click(); toast('Word exported'); }
    function exportJSON() { const data = JSON.stringify({cover:state.cover, results:state.results, details:state.details, sections:state.sections, pendingEdits:state.pendingEdits}, null,2); const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([data])); a.download=`GW81_backup.json`; a.click(); toast('JSON exported'); }
    function importJSON(e) { const f=e.target.files[0]; if(!f) return; const r=new FileReader(); r.onload=ev=>{ try{ const d=JSON.parse(ev.target.result); if(d.sections) state.sections=d.sections; if(d.results) state.results=d.results; if(d.cover) state.cover=d.cover; if(d.pendingEdits) state.pendingEdits=d.pendingEdits; saveState(); renderCover(); buildInspection(); refreshDashboard(); buildSidebar(); if(isAdmin()) buildAdminPanel(); toast('Import success'); }catch(e){ toast('Invalid file'); } }; r.readAsText(f); e.target.value=''; }
    function resetAllData() { if(confirm('Reset all inspection data?')) { state.results={}; state.details={}; state.photos={}; state.pendingEdits=[]; saveState(); buildInspection(); refreshDashboard(); toast('Data reset'); } }

    function renderCover() { const c=state.cover||{}; ['mfr','model','serial','date','inspector','site','ctrl','approved','notes'].forEach(k=>{ let el=document.getElementById(`f-${k}`); if(el&&c[k]!==undefined) el.value=c[k]; }); if(!c.date) document.getElementById('f-date').valueAsDate=new Date(); }
    function saveCover() { ['mfr','model','serial','date','inspector','site','ctrl','approved','notes'].forEach(k=>{ state.cover[k]=document.getElementById(`f-${k}`)?.value||''; }); saveState(); }

    function showPage(page) { document.querySelectorAll('.page').forEach(p=>p.classList.remove('active')); document.getElementById(`page-${page}`).classList.add('active'); document.querySelectorAll('.ntab').forEach(t=>t.classList.remove('active')); if(page==='cover') document.querySelector('.ntab').classList.add('active'); else if(page==='inspection') document.querySelectorAll('.ntab')[1].classList.add('active'); else if(page==='dashboard') document.querySelectorAll('.ntab')[2].classList.add('active'); else if(page==='admin') document.querySelectorAll('.ntab')[3].classList.add('active'); if(page==='dashboard') refreshDashboard(); if(page==='admin' && isAdmin()) buildAdminPanel(); }
    function attachTabEvents() { document.querySelectorAll('.ntab').forEach((btn,idx)=>{ btn.onclick=()=>{ if(idx===0) showPage('cover'); else if(idx===1) showPage('inspection'); else if(idx===2) showPage('dashboard'); else if(idx===3 && isAdmin()) showPage('admin'); }; }); }
    function toast(msg) { const t=document.getElementById('toast'); t.innerText=msg; t.classList.add('show'); setTimeout(()=>t.classList.remove('show'),2500); }
    function escapeHtml(s) { return (s||'').replace(/[&<>]/g, function(m){if(m==='&')return'&amp;';if(m==='<')return'&lt;';if(m==='>')return'&gt;';return m;}); }

    // ========== AI ASSISTANT ==========
    function toggleAIPanel() { document.getElementById('ai-panel').classList.toggle('open'); }
    function askAI() {
        const query = document.getElementById('ai-query').value.trim();
        if(!query) return;
        const msgDiv = document.getElementById('ai-messages');
        msgDiv.innerHTML += `<div class="ai-message user">${escapeHtml(query)}</div>`;
        setTimeout(()=>{
            let reply = "AI: ";
            if(query.toLowerCase().includes('severity')) reply = "Severity: L (Low - cosmetic, monitor), M (Medium - repair within 3 months), H (High - immediate repair).";
            else if(query.toLowerCase().includes('stud')) reply = "Check all studs for rust, thread damage, and ensure at least one stud has a batch number engraved.";
            else reply += "Refer to the editable criteria for each item.";
            msgDiv.innerHTML += `<div class="ai-message">${reply}</div>`;
            msgDiv.scrollTop = msgDiv.scrollHeight;
        }, 300);
        document.getElementById('ai-query').value = '';
    }

    // Expose global functions for HTML onclick
    window.setResult = setResult; window.updateDetail = updateDetail; window.attachPhotos = attachPhotos; window.deletePhoto = deletePhoto;
    window.exportToExcel = exportToExcel; window.exportToPDF = exportToPDF; window.exportToWord = exportToWord; window.exportJSON = exportJSON;
    window.resetAllData = resetAllData; window.jumpToSection = jumpToSection; window.editField = editField; window.toast = toast; window.showPage = showPage;
    window.toggleAIPanel = toggleAIPanel; window.askAI = askAI;
    window.approveEdit = approveEdit; window.rejectEdit = rejectEdit; window.approveUser = approveUser; window.rejectUser = rejectUser;
    window.deleteUser = deleteUser; window.showAddUserForm = showAddUserForm; window.hideAddUserForm = hideAddUserForm; window.addPendingUser = addPendingUser;
    window.setTheme = setTheme;
    loadState();
</script>
</body>
</html>