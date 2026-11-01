<!DOCTYPE html>

<html lang="fr">

<head>

<meta charset="UTF-8">

<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>FieldOps</title>

<link rel="preconnect" href="https://fonts.googleapis.com">

<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;700&family=Baloo+2:wght@700;800&display=swap" rel="stylesheet">



<!-- Firebase (compat SDK — single file, no build step) -->

<script src="https://www.gstatic.com/firebasejs/10.13.0/firebase-app-compat.js"></script>

<script src="https://www.gstatic.com/firebasejs/10.13.0/firebase-auth-compat.js"></script>

<script src="https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore-compat.js"></script>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>



<style>

:root{

  --sand: #FFF6DE;

  --sand-deep: #FDE29E;

  --sky: #00A082;

  --sky-deep: #007A63;

  --sky-ink: #00614E;

  --mustard: #E8A317;

  --mustard-deep: #B9800E;

  --ink: #1F1B0E;

  --ink-soft: #6B5D3D;

  --paper: #FFFFFF;

  --line: rgba(80,64,20,0.26);

  --green: #047857;

  --red: #DC2626;

  --radius: 14px;

  --font-display: 'Inter', sans-serif;

  --font-mono: 'JetBrains Mono', monospace;

}

*{box-sizing:border-box; margin:0; padding:0;}

html{ background: var(--sand); }

body{

  font-family: var(--font-display);

  background:

    radial-gradient(circle at 12% 0%, rgba(0,160,130,0.07), transparent 42%),

    radial-gradient(circle at 88% 18%, rgba(232,163,23,0.06), transparent 40%),

    var(--sand);

  background-attachment: fixed;

  color: var(--ink);

  min-height: 100vh;

  -webkit-font-smoothing: antialiased;

}

button{font-family: inherit; cursor:pointer; border:none;}

input, select{font-family: inherit;}

::selection{background: var(--mustard); color: #fff;}



/* ============ LOGIN SCREEN ============ */

:root{

  --login-navy: #FFF6DE;

  --login-navy-2: #FDE29E;

  --login-amber: #FFC244;

  --login-cream: #FFFFFF;

}

#loginScreen{ min-height: 100vh; }

.login-split{

  min-height: 100vh; display:flex;

}

.login-hero-panel{

  flex: 1 1 46%; min-width: 340px;

  background:

    radial-gradient(circle at 20% 10%, rgba(0,160,130,0.16), transparent 40%),

    linear-gradient(160deg, var(--login-navy), var(--login-navy-2));

  color: #2B2440;

  display:flex; flex-direction:column; justify-content:space-between;

  padding: 40px 44px;

  position: relative; overflow:hidden;

}

.hero-brand{ display:flex; align-items:center; gap:10px; }

.hero-brand .login-mark{ width:36px; height:36px; box-shadow:none; background: var(--login-amber); }

.login-mark{

  width: 36px; height: 36px; border-radius: 10px;

  display:flex; align-items:center; justify-content:center;

}

.login-mark svg{ width:20px; height:20px; }

.hero-brand span{ font-size:16px; font-weight:800; letter-spacing:-0.01em; color:#2B2440; }

.route-map{ flex:1; display:flex; align-items:center; justify-content:center; padding: 10px 0; }

.route-svg{ width: 100%; max-width: 340px; height: auto; }

.pin-dot-outer{ fill: var(--login-amber); }

.pin-dot-inner{ fill: #fff; }

.pin-label{ fill: rgba(43,36,64,0.82); font-family: var(--font-display); font-size: 12.5px; font-weight: 700; }

.pin-label-sub{ fill: rgba(43,36,64,0.45); font-family: var(--font-mono); font-size: 9.5px; }

.scooter-icon path, .scooter-icon circle{ fill: var(--login-amber); }

@media (prefers-reduced-motion: reduce){ #routeMotion{ display:none; } }

.hero-foot{ font-family: var(--font-mono); }

.hero-foot span{ font-size: 22px; font-weight: 700; color: var(--login-amber); display:block; margin-bottom:6px; }

.hero-foot p{ font-size: 12px; color: rgba(43,36,64,0.6); max-width: 280px; }



.login-form-panel{

  flex: 1 1 54%;

  background: var(--login-cream);

  display:flex; align-items:center; justify-content:center;

  padding: 40px 24px;

}

.login-card{

  width: 100%; max-width: 440px;

  background: transparent;

  padding: 0;

}

.login-brand{

  display:flex; flex-direction:column; align-items:flex-start; text-align:left;

  margin-bottom: 26px;

}

.login-brand h1{ font-size: 24px; font-weight: 800; letter-spacing: -0.02em; color: var(--sky-ink); }

.login-brand p{ font-size: 13px; color: var(--ink-soft); margin-top: 4px; font-family: var(--font-mono); }



@media (max-width: 860px){

  .login-split{ flex-direction: column; }

  .login-hero-panel{ min-width: 0; padding: 28px 28px 22px; min-height: 0; }

  .route-map{ display:none; }

  .hero-foot p{ display:none; }

  .hero-foot span{ font-size: 16px; }

}

.profile-grid{

  display:grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px;

  margin-bottom: 20px; max-height: 320px; overflow-y: auto; padding-right: 2px;

}

.profile-btn{

  background: var(--sand-deep);

  border: 1.5px solid var(--line);

  border-radius: var(--radius);

  padding: 16px 10px;

  display:flex; flex-direction:column; align-items:center; gap: 8px;

  transition: all .15s ease;

}

.profile-btn:hover{ border-color: var(--sky); background: var(--paper); transform: translateY(-1px); box-shadow: 0 4px 14px -8px rgba(30,20,60,0.2); }

.profile-btn.selected{ border-color: var(--sky); background: var(--paper); box-shadow: 0 0 0 3px rgba(0,160,130,0.16); }

.profile-avatar{

  width: 40px; height: 40px; border-radius: 50%;

  display:flex; align-items:center; justify-content:center;

  color:#fff; font-weight:700; font-size: 14px;

}

.profile-avatar svg{ width: 20px; height: 20px; }

.profile-btn span{font-size: 13px; font-weight: 600; color: var(--ink);}

.profile-btn small{font-size: 10px; color: var(--ink-soft); text-transform:uppercase; letter-spacing:.04em;}



.login-input{

  width:100%; padding: 13px 14px; border-radius: var(--radius);

  border: 1.5px solid var(--line); background: var(--sand-deep);

  font-size: 14px; margin-bottom: 12px; color: var(--ink);

  transition: border-color .15s, box-shadow .15s;

}

.login-input:focus{ outline:none; border-color: var(--sky); background: var(--paper); box-shadow: 0 0 0 3px rgba(0,160,130,0.14); }

.login-submit{

  width:100%; padding: 14px; border-radius: var(--radius);

  background: var(--mustard); color: #fff; font-weight: 700; font-size: 14px;

  letter-spacing: .02em; transition: background .15s;

}

.login-submit:hover{ background: var(--mustard-deep); }

.login-submit:disabled{ background: var(--line); color: var(--ink-soft); cursor:not-allowed; }

.login-error{

  font-size: 12.5px; color: var(--red); text-align:center; margin-top: 10px; min-height: 16px;

}



/* ============ LOGIN TABS ============ */

.login-tabs{ display:flex; gap:6px; background: var(--sand-deep); border-radius: 10px; padding: 4px; margin-bottom: 20px; }

.ltab{ flex:1; padding: 9px; border-radius: 7px; font-size: 12.5px; font-weight:700; color: var(--ink-soft); background: transparent; }

.ltab.active{ background: var(--sky-deep); color: #fff; box-shadow: 0 4px 14px -6px rgba(0,122,99,0.35); }



/* ============ AGENT PORTAL ============ */

#agentPortal, #nvAgentPortal{

  min-height: 100vh; display:flex; flex-direction:column; align-items:center; justify-content:flex-start;

  gap: 16px; padding: 20px 16px; position: relative; overflow-x:hidden;

  background:

    radial-gradient(circle at 15% 15%, rgba(0,160,130,0.10), transparent 45%),

    radial-gradient(circle at 85% 85%, rgba(232,163,23,0.08), transparent 45%),

    radial-gradient(circle at 1px 1px, rgba(30,20,60,0.05) 1px, transparent 0) 0 0/28px 28px,

    var(--sand);

}

.nv-portal-top-week{

  display:block; position:static; width:100%; max-width:460px; background:#fff; border-radius:16px;

  padding:14px 16px; box-shadow: 0 10px 24px -14px rgba(30,20,60,0.22);

}

.nv-badges-row{ display:flex; flex-wrap:wrap; justify-content:center; gap:10px; width:100%; max-width:460px; }

.portal-bg-doodle{

  display:flex; flex-direction:column; align-items:center; justify-content:center; gap:6px;

  position:static; width:80px; height:80px; border-radius:14px; pointer-events:none;

  box-shadow:none; padding:8px 4px;

}

.portal-bg-doodle .dd-badge{

  width:36px; height:36px; border-radius:50%; background:#fff; display:flex; align-items:center; justify-content:center;

  flex-shrink:0;

}

.portal-bg-doodle .dd-badge svg{ width:20px; height:20px; }

.portal-bg-doodle .dd-label{

  font-family: 'Baloo 2', var(--font-display); font-size:10px; font-weight:800; text-transform:uppercase; letter-spacing:0;

  color: #fff; text-align:center; line-height:1.05;

}

@media (min-width: 1280px){

  #agentPortal, #nvAgentPortal{ flex-direction:row; align-items:center; justify-content:center; padding-top: 150px; overflow:hidden; gap:0; }

  .nv-badges-row{ display:contents; max-width:none; }

  .portal-bg-doodle{

    flex-direction:column; align-items:center; justify-content:center; gap:14px;

    position:absolute; width:130px; height:130px; border-radius:0; pointer-events:none;

    box-shadow:none; padding:14px 8px;

  }

  .portal-bg-doodle .dd-badge{

    width:62px; height:62px; border-radius:50%; background:#fff; display:flex; align-items:center; justify-content:center;

    box-shadow:none; flex-shrink:0;

  }

  .portal-bg-doodle .dd-badge svg{ width:34px; height:34px; }

  .portal-bg-doodle .dd-label{

    font-family: 'Baloo 2', var(--font-display); font-size:17px; font-weight:800; text-transform:uppercase; letter-spacing:0;

    color: #fff; text-align:center; line-height:1.05;

  }

  .portal-bg-doodle.dd-tl{ top:7%; left:5%; }

  .portal-bg-doodle.dd-bl{ bottom:9%; left:8%; }

  .portal-bg-doodle.dd-tr{ top:9%; right:6%; }

  .portal-bg-doodle.dd-br{ bottom:8%; right:5%; }

  .portal-bg-doodle.dd-ml{ top:44%; left:2%; }

  .portal-bg-doodle.dd-mr{ top:46%; right:3%; }

  .nv-portal-top-week{

    display:block; position:absolute; top:24px; left:50%; transform:translateX(-50%);

    width:min(94%, 1100px); background:#fff; border-radius:16px; padding:16px 20px;

    box-shadow: 0 14px 30px -14px rgba(30,20,60,0.25); pointer-events:none; z-index:1;

  }

}

.portal-card{

  width:100%; max-width: 420px; background: var(--paper); border-radius: 20px; padding: 28px;

  box-shadow: 0 30px 80px -24px rgba(30,20,60,0.18); border:1px solid var(--line);

}

.portal-side-art{ display:none; }

@media (min-width: 900px){

  #agentPortal, #nvAgentPortal{ gap: 48px; }

  .portal-side-art{

    display:flex; flex-direction:column; align-items:center; width: 320px; flex-shrink:0;

  }

  .portal-side-art svg{ width:100%; max-width:320px; height:auto; }

  .portal-side-art-quote{

    font-family: var(--font-display); font-style:italic; font-weight:600; font-size:15px;

    color: var(--ink-soft); text-align:center; margin-top:8px; line-height:1.5; max-width:260px;

  }

}

@media (min-width: 720px){

  .portal-card{ max-width: 460px; padding: 34px; }

}

.portal-header{ display:flex; align-items:center; justify-content:space-between; margin-bottom: 20px; }

.portal-header b{ font-size: 16px; display:block; color: var(--ink); }

.portal-header span{ font-size: 12px; color: var(--ink-soft); font-family: var(--font-mono); }

.portal-status{

  text-align:center; padding: 14px; border-radius: 12px; background: var(--sand-deep);

  font-size: 13px; font-weight:700; color: var(--ink-soft); margin-bottom: 18px;

}

.portal-btn{

  width:100%; padding: 22px; border-radius: 14px; background: var(--green); color:#fff;

  font-size: 16px; font-weight: 800; letter-spacing: .02em; margin-bottom: 6px;

}

.portal-btn.is-checkedin{ background: var(--red); }

.portal-history{ margin-top: 4px; }

.portal-history-item{

  display:flex; justify-content:space-between; padding: 9px 0; border-bottom: 1px solid var(--line);

  font-size: 12.5px;

}

.portal-history-item:last-child{ border-bottom:none; }



.shell{ display:flex; min-height: 100vh; }



.sidebar{

  width: 220px; flex-shrink:0;

  background: linear-gradient(180deg, #FAF8FE, #F0EAFC);

  border-right: 1.5px solid var(--line);

  color: var(--ink);

  display:flex; flex-direction:column;

  padding: 22px 14px;

}

.sidebar-brand{

  display:flex; align-items:center; gap: 10px;

  padding: 0 8px 22px; margin-bottom: 18px;

  border-bottom: 1px solid var(--line);

}

.sidebar-brand .mark{

  width:32px; height:32px; border-radius:9px;

  background: linear-gradient(135deg, var(--sky), var(--sky-deep));

  display:flex; align-items:center; justify-content:center;

  font-weight:800; color:#fff; font-size:14px;

  box-shadow: 0 8px 18px -8px rgba(0,160,130,0.55);

}

.sidebar-brand h2{ font-size: 15.5px; font-weight: 800; letter-spacing: -0.01em; color: var(--ink); }

.nav-group{ margin-bottom: 18px; }

.nav-label{

  font-size: 10px; text-transform: uppercase; letter-spacing: .08em;

  color: var(--ink-soft); padding: 0 10px; margin-bottom: 6px; font-weight:600;

}

.nav-item{

  display:flex; align-items:center; gap: 10px;

  width:100%; text-align:left; padding: 9px 10px; border-radius: 9px;

  color: var(--ink-soft); background: transparent; font-size: 13.5px; font-weight: 600;

  transition: all .12s; margin-bottom: 2px;

}

.nav-item:hover{ background: var(--sand-deep); color: var(--ink); }

.nav-item.active{

  background: linear-gradient(135deg, var(--sky), var(--sky-deep));

  color: #fff;

  box-shadow: 0 6px 18px -8px rgba(0,160,130,0.5);

}

.nav-item svg{ width:16px; height:16px; flex-shrink:0; }

.sidebar-footer{ margin-top:auto; padding-top: 14px; border-top: 1px solid var(--line); }

.user-chip{ display:flex; align-items:center; gap:9px; padding: 8px 10px; }

.user-chip .dot{ width:8px; height:8px; border-radius:50%; background: var(--green); flex-shrink:0; box-shadow: 0 0 8px rgba(5,150,105,0.5); }

.user-chip-text{ display:flex; flex-direction:column; }

.user-chip-text b{ font-size: 12.5px; color: var(--ink); }

.user-chip-text span{ font-size: 10.5px; color: var(--ink-soft); }

.logout-btn{

  width:100%; margin-top:10px; padding: 8px; border-radius: 8px;

  background: var(--sand-deep); color: var(--ink-soft);

  font-size: 12.5px; font-weight:600;

}

.logout-btn:hover{ background: var(--line); color: var(--ink); }



.main{ flex:1; overflow-x:auto; max-width:100vw; }

.topbar{

  display:flex; align-items:center; justify-content:space-between;

  padding: 18px 32px; border-bottom: 1px solid var(--line);

  background: rgba(255,255,255,0.85);

  backdrop-filter: blur(14px);

  position: sticky; top: 0; z-index: 5;

}

.topbar h1{ font-size: 19px; font-weight: 800; letter-spacing: -0.01em; color: var(--ink); }

.topbar-meta{ display:flex; align-items:center; gap: 14px; font-size: 12.5px; color: var(--ink-soft); font-family: var(--font-mono); }

.sync-pill{

  display:flex; align-items:center; gap:6px; padding: 5px 10px;

  background: rgba(5,150,105,0.10); color: var(--green); border-radius: 20px; font-weight:600; font-size: 11.5px;

}

.sync-pill .dot{ width:6px; height:6px; border-radius:50%; background: var(--green); }



.view{ padding: 28px 32px; }

.hidden{ display:none !important; }



/* Dashboard cards */

.stat-grid{ display:grid; grid-template-columns: repeat(4,1fr); gap: 16px; margin-bottom: 24px; }

.stat-card{

  background: var(--paper); border: 1px solid var(--line); border-left: 4px solid var(--sky-ink);

  border-radius: 18px; padding: 18px;

  box-shadow: 0 4px 16px -8px rgba(30,20,60,0.12);

  transition: transform .18s ease, box-shadow .18s ease;

}

.stat-card:hover{ transform: translateY(-2px); box-shadow: 0 16px 32px -20px rgba(0,160,130,0.25); }

.stat-card.accent-green{ border-left-color: #059669; }

.stat-card.accent-blue{ border-left-color: #0284C7; }

.stat-card.accent-amber{ border-left-color: #B45309; }

.stat-card.accent-purple{ border-left-color: #00A082; }

.stat-card.accent-red{ border-left-color: #DC2626; }

.stat-card.accent-orange{ border-left-color: #EA580C; }

.stat-card .label{ font-size: 11.5px; color: var(--ink-soft); text-transform:uppercase; letter-spacing:.04em; font-weight:600; margin-bottom: 8px;}

.stat-card .value{ font-size: 28px; font-weight: 800; color: var(--sky-ink); font-family: var(--font-mono); }

.stat-card .sub{ font-size: 12px; color: var(--ink-soft); margin-top: 4px; }



.zone-grid{ display:grid; grid-template-columns: repeat(auto-fill, minmax(210px,1fr)); gap: 14px; }



/* ============ PLANNING NV — zone cards ============ */

.nv-zone-grid{ display:grid; grid-template-columns: repeat(auto-fit, minmax(320px,1fr)); gap: 18px; }

.nv-zone-card{

  border-radius: 20px; overflow:hidden; background: var(--paper);

  border: 1px solid var(--line); box-shadow: 0 12px 32px -20px rgba(30,20,60,0.18);

}

.nv-zone-head{

  padding: 16px 20px; color:#fff; display:flex; align-items:center; justify-content:space-between;

}

.nv-zone-head h3{ font-size: 19px; font-weight: 800; letter-spacing:-0.01em; }

.nv-zone-head span{ font-size: 12.5px; font-weight: 700; background: rgba(255,255,255,0.28); padding: 5px 12px; border-radius: 20px; }

.nv-hotspot-line{ padding: 10px 20px; font-size: 13px; color: var(--ink-soft); border-bottom: 1px solid var(--line); background: var(--sand-deep); line-height:1.6; }

.nv-member-row{

  display:flex; align-items:center; justify-content:space-between; gap: 10px;

  padding: 14px 20px; border-bottom: 1px solid var(--line);

}

.nv-member-row:last-child{ border-bottom:none; }

.nv-member-name{ font-weight: 800; font-size: 16px; flex: 1; color: var(--ink); }

.nv-member-hotspot{ font-size: 12.5px; color: var(--ink-soft); margin-top: 4px; line-height: 1.5; }

.nv-chip{ font-size: 12.5px; font-weight: 700; padding: 6px 13px; border-radius: 20px; white-space:nowrap; }

.nv-chip.shift-jour{ background: rgba(5,150,105,0.18); color: #047857; }

.nv-chip.shift-nuit{ background: rgba(232,163,23,0.18); color: #8A6208; }

.nv-chip.off{ background: rgba(107,114,128,0.12); color: #6B7280; }

.nv-day-strip{ display:flex; flex-direction:column; gap: 10px; }

.nv-day-block{ border-bottom: 1px solid var(--line); padding-bottom: 8px; }

.nv-day-today{ background: rgba(0,160,130,0.06); border-radius: 8px; padding: 8px; margin: -8px -8px 0; }

.nv-today-tag{ background: var(--sky-deep); color:#fff; font-size:9.5px; font-weight:800; padding:2px 7px; border-radius:20px; text-transform:uppercase; letter-spacing:.03em; }

.nv-portal-hero{

  border-radius: 18px; padding: 20px 22px; color:#fff; box-shadow: 0 16px 34px -20px rgba(30,20,60,0.35);

}

.nv-portal-hero-label{ font-size: 12px; font-weight:600; opacity:.85; margin-bottom:4px; }

.nv-portal-hero-zone{ font-size: 24px; font-weight: 800; letter-spacing:-0.01em; margin-bottom: 10px; }

.nv-portal-hero-shift{ font-size: 14px; font-weight: 700; background: rgba(255,255,255,0.22); display:inline-block; padding: 6px 12px; border-radius: 20px; }

.nv-portal-day{ border-bottom: 1px solid var(--line); padding-bottom: 10px; margin-bottom: 10px; }

.nv-portal-day:last-child{ border-bottom:none; margin-bottom:0; padding-bottom:0; }

.nv-time-chip{ font-size: 10.5px; font-weight: 700; padding: 3px 9px; border-radius: 20px; margin-left: 6px; display:inline-block; }

.nv-time-chip.nv-time-jour{ background: rgba(5,150,105,0.18); color: #047857; }

.nv-time-chip.nv-time-nuit{ background: rgba(232,163,23,0.18); color: #8A6208; }

.nv-day-block:last-child{ border-bottom: none; padding-bottom: 0; }

.nv-day-label{ font-size: 12px; font-weight: 800; color: var(--ink); margin-bottom: 5px; text-transform:uppercase; letter-spacing:.03em; }

.nv-day-venues{ display:flex; flex-wrap:wrap; gap: 6px; }

.nv-day-chip{

  font-size: 14px; color: var(--ink); background: var(--sand-deep); border-radius: 8px;

  padding: 8px 12px; line-height: 1.3; display:inline-block; text-decoration:none; font-weight:600;

}

.nv-day-chip-link{ cursor:pointer; }

.nv-day-chip-link:hover{ background: #FDE29E; text-decoration: underline; }

.nv-day-chip.off{ background: rgba(107,114,128,0.10); color: #6B7280; font-weight: 700; }

.nv-today-hero{

  border-radius: 12px; padding: 12px 14px; color:#fff; display:flex; flex-direction:column; gap: 4px;

  box-shadow: 0 10px 24px -14px rgba(30,20,60,0.4);

}

.nv-today-hero-label{ font-size: 11px; font-weight: 800; text-transform:uppercase; letter-spacing:.04em; opacity:.9; }

.nv-today-hero-venues{ font-size: 15px; font-weight: 800; }

.nv-today-hero-venues a{ color:#fff; text-decoration:none; }

.nv-today-hero-venues a:hover{ text-decoration:underline; }

.nv-today-hero.nv-today-hero-off{

  background: var(--sand-deep); color: var(--ink-soft); font-size: 13px; font-weight: 700;

  border-radius: 12px; padding: 12px 14px; text-align:center; box-shadow:none;

}

.nv-fulllist{ padding: 12px 20px 16px; background: var(--sand-deep); border-top: 1px solid var(--line); }

.nv-fulllist-title{ font-size: 10.5px; text-transform:uppercase; letter-spacing:.04em; font-weight:700; color: var(--ink-soft); margin-bottom: 8px; }

.nv-fulllist-row{ font-size: 12px; color: var(--ink); padding: 8px 0; line-height:1.5; border-bottom: 1px solid var(--line); }

.nv-fulllist-row:last-of-type{ border-bottom: none; }

.nv-venue-tag{

  display:inline-flex; align-items:center; gap:5px; background: var(--paper); border:1px solid var(--line);

  border-radius: 6px; padding: 3px 7px; margin: 2px 3px 2px 0; font-size: 11.5px;

}

.nv-venue-tag button{ border:none; background:none; color: var(--red); cursor:pointer; font-size:11px; padding:0; line-height:1; }

.nv-venue-tag a{ color: var(--ink); text-decoration:none; }

.nv-venue-tag a:hover{ text-decoration:underline; }

.nv-venue-links{ display:flex; flex-direction:column; gap:8px; }

.nv-venue-maplink{

  display:block; background: var(--paper); border:1px solid var(--line); border-radius: 10px;

  padding: 10px 14px; font-size: 14px; font-weight: 700; color: var(--ink); text-decoration:none;

}

.nv-venue-maplink:hover{ border-color: var(--sky); }

.nv-mini-link{

  border:none; background:none; color: var(--sky-ink); font-size: 11px; font-weight:700; cursor:pointer; padding:2px 6px 2px 0;

}

.nv-mini-link.nv-mini-danger{ color: var(--red); }

.agent-pin-chip{

  display:inline-block; font-family: var(--font-mono); font-weight:800; font-size:13px;

  background: var(--sand-deep); color: var(--ink); padding: 4px 10px; border-radius: 8px; letter-spacing: 0.02em;

}

.agent-access-btn{

  border:1px solid var(--line); background: var(--paper); color: var(--sky-ink); font-size: 11.5px; font-weight:700;

  cursor:pointer; padding: 5px 10px; border-radius: 8px; display:inline-flex; align-items:center; gap:4px;

  transition: background 0.15s, border-color 0.15s;

}

.agent-access-btn:hover{ background: rgba(0,160,130,0.08); border-color: var(--sky); }

.nv-fulllist-row b{ color: var(--ink-soft); margin-right: 6px; }

.nv-loc-badge{

  display:block; text-align:center; font-size: 11px; font-weight: 700; padding: 6px 10px;

  border-radius: 8px; margin-top: 8px; text-decoration:none;

}

.nv-loc-badge.nv-loc-none{ background: var(--sand-deep); color: var(--ink-soft); font-weight: 500; }

.nv-loc-badge.nv-loc-fresh{ background: rgba(5,150,105,0.18); color: #047857; }

.nv-loc-badge.nv-loc-stale{ background: rgba(180,83,9,0.12); color: #B45309; }

.zone-card{

  background: var(--paper); border: 1px solid var(--line); border-radius: 18px; padding: 16px;

  border-left: 4px solid var(--mustard);

}

.zone-card h3{ font-size: 14px; font-weight:700; margin-bottom: 10px; color: var(--ink); }

.zone-card .count{ font-size: 24px; font-weight: 800; color: var(--sky-ink); font-family: var(--font-mono); }

.zone-card .count-label{ font-size: 11.5px; color: var(--ink-soft); }

.zone-dot{ display:inline-block; width:8px; height:8px; border-radius:50%; margin-right:6px; vertical-align:middle; }



.section-title{ font-size: 15px; font-weight: 800; margin: 26px 0 14px; color: var(--ink); }



.card{ background: var(--paper); border:1.5px solid var(--line); border-radius: 18px; box-shadow: 0 4px 16px -8px rgba(30,20,60,0.12); }



table{ width:100%; border-collapse: collapse; }

th{ text-align:left; font-size: 11px; text-transform:uppercase; letter-spacing:.04em; color: var(--ink-soft); font-weight:700; padding: 12px 16px; border-bottom: 1px solid var(--line); }

td{ padding: 12px 16px; font-size: 13.5px; border-bottom: 1px solid var(--line); color: var(--ink); }

tr:last-child td{ border-bottom:none; }

.badge{ display:inline-block; padding: 3px 9px; border-radius: 20px; font-size: 11px; font-weight: 700; }

.badge.fix{ background: rgba(0,160,130,0.16); color: #00614E; }

.badge.var{ background: rgba(180,83,9,0.20); color: #92400E; }

.badge.jour{ background: rgba(5,150,105,0.18); color: #047857; }

.badge.nuit{ background: rgba(232,163,23,0.18); color: #8A6208; }

.badge.abs-maladie{ background: rgba(220,38,38,0.16); color: #B91C1C; }

.badge.abs-injustifiee{ background: rgba(107,114,128,0.18); color: #374151; }

.badge.abs-conge{ background: rgba(0,160,130,0.16); color: #00614E; }

.badge.abs-ferie{ background: rgba(5,150,105,0.18); color: #047857; }

.badge.abs-double{ background: rgba(180,83,9,0.20); color: #92400E; }

.badge.abs-off{ background: rgba(2,132,199,0.18); color: #0369A1; }



.placeholder-view{

  display:flex; flex-direction:column; align-items:center; justify-content:center;

  padding: 80px 20px; text-align:center; color: var(--ink-soft);

}

.placeholder-view svg{ width:44px; height:44px; opacity:.35; margin-bottom:14px; }

.placeholder-view h3{ font-size: 16px; color: var(--ink); margin-bottom: 6px; }

.placeholder-view p{ font-size: 13px; max-width: 320px; }



@media (max-width: 900px){

  .stat-grid{ grid-template-columns: repeat(2,1fr); }

  .sidebar{ width: 72px; }

  .sidebar-brand h2, .nav-item span, .user-chip-text, .logout-btn span{ display:none; }

  .nav-item{ justify-content:center; }

}



/* ============ VIEW HEADER / BUTTONS ============ */

.view-header{

  display:flex; align-items:center; justify-content:space-between;

  margin-bottom: 18px; gap: 16px;

}

.view-subtitle{ font-size: 13px; color: var(--ink-soft); }

.btn-primary{

  background: linear-gradient(135deg, var(--mustard), var(--mustard-deep));

  color: #fff; font-weight: 700; font-size: 13px;

  padding: 10px 18px; border-radius: var(--radius); white-space: nowrap;

  transition: transform .15s, box-shadow .15s;

  box-shadow: 0 10px 22px -10px rgba(232,163,23,0.45);

}

.btn-primary:hover{ transform: translateY(-1px); box-shadow: 0 14px 26px -10px rgba(232,163,23,0.55); }

.btn-ghost{

  background: transparent; color: var(--ink-soft); font-weight: 600; font-size: 13px;

  padding: 10px 16px; border-radius: var(--radius);

}

.btn-ghost:hover{ background: var(--sand-deep); color: var(--ink); }

.icon-btn{

  width: 28px; height: 28px; border-radius: 7px; background: transparent;

  display:inline-flex; align-items:center; justify-content:center; color: var(--ink-soft);

  transition: all .12s;

}

.icon-btn:hover{ background: var(--sand-deep); color: var(--ink); }

.icon-btn svg{ width: 15px; height: 15px; }

.icon-btn.danger:hover{ background: rgba(220,38,38,0.10); color: var(--red); }

.row-actions{ display:flex; gap: 4px; justify-content:flex-end; }



/* ============ MODAL ============ */

.modal-overlay{

  position: fixed; inset: 0; background: rgba(30,20,60,0.35);

  backdrop-filter: blur(4px);

  display:flex; align-items:center; justify-content:center; z-index: 100;

  padding: 20px;

}

.modal-card{

  width: 100%; max-width: 420px; background: var(--paper);

  border: 1px solid var(--line);

  border-radius: 16px; padding: 28px 26px 22px;

  box-shadow: 0 30px 70px -24px rgba(30,20,60,0.35);

}

.modal-card h3{ font-size: 17px; font-weight: 800; margin-bottom: 18px; color: var(--ink); }

.form-label{ display:block; font-size: 11.5px; font-weight: 700; color: var(--ink-soft); text-transform:uppercase; letter-spacing:.03em; margin-bottom: 6px; margin-top: 12px; }

.form-label:first-of-type{ margin-top: 0; }

.modal-error{ font-size: 12.5px; color: var(--red); min-height: 16px; margin-top: 6px; }

.modal-actions{ display:flex; justify-content:flex-end; gap: 8px; margin-top: 16px; }

.modal-hint{ font-size: 12px; color: var(--ink-soft); background: var(--sand-deep); border-radius: 8px; padding: 8px 10px; margin-top: 10px; line-height: 1.4; }



/* ============ PLANNING ============ */

.planning-header{

  display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap: 14px;

  margin-bottom: 16px;

}

.week-nav{ display:flex; align-items:center; gap: 12px; }

.icon-btn-lg{

  width: 34px; height: 34px; border-radius: 9px; background: var(--paper); border: 1px solid var(--line);

  display:flex; align-items:center; justify-content:center; color: var(--sky-ink);

  transition: all .12s;

}

.icon-btn-lg svg{ width: 16px; height:16px; }

.icon-btn-lg:hover{ background: var(--sand-deep); }

.week-label{ text-align:center; min-width: 190px; }

.week-label b{ display:block; font-size: 14.5px; font-weight: 800; color: var(--ink); }

.week-label span{ font-size: 11.5px; color: var(--ink-soft); font-family: var(--font-mono); }

.planning-actions{ display:flex; gap: 8px; flex-wrap: wrap; }



.planning-tabs{ display:flex; gap: 4px; margin-bottom: 14px; }

.ptab{

  padding: 8px 16px; border-radius: 8px; background: transparent; color: var(--ink-soft);

  font-size: 13px; font-weight: 700;

}

.ptab.active{ background: linear-gradient(135deg, var(--sky), var(--sky-deep)); color: #fff; }

.ptab:not(.active):hover{ background: var(--sand-deep); }



.planning-card{ overflow-x: auto; }

.planning-table{ width: 100%; border-collapse: collapse; min-width: 780px; table-layout: fixed; }

.planning-table th, .planning-table td{ padding: 6px 4px; font-size: 12px; border-bottom: 1px solid var(--line); text-align:center; white-space: nowrap; }

.planning-table th:not(.sticky-col), .planning-table td:not(.sticky-col):not(.zone-cell){ width: 78px; }

.planning-table th.sticky-col, .planning-table td.sticky-col{ width: 140px; }

.planning-table td.zone-cell{ width: 92px; white-space: normal; text-align:left; font-size:11px; }

.planning-table thead th{

  text-transform:uppercase; font-size: 10px; letter-spacing:.05em; color: var(--ink-soft); font-weight: 800;

  background: var(--sand-deep); border-bottom: 2px solid var(--line); padding-top: 12px; padding-bottom: 12px;

}

.planning-table tbody tr:nth-child(even){ background: rgba(30,20,60,0.015); }

.planning-table tbody tr:hover{ background: rgba(0,160,130,0.04); }

.planning-table td.agent-name-cell{ text-align:left; font-weight: 700; color: var(--ink); }

.planning-table td.zone-cell{ text-align:left; font-size: 11.5px; color: var(--ink-soft); }

.agent-name-edit{

  border: none; background: transparent; font-weight: 700; font-size: 12.5px; color: var(--ink);

  width: 100%; padding: 3px 4px; border-radius: 6px; font-family: var(--font-display);

}

.agent-name-edit:hover{ background: var(--sand-deep); }

.agent-name-edit:focus{ outline: none; background: var(--paper); box-shadow: 0 0 0 1.5px var(--sky); }

.sticky-col{ position: sticky; left: 0; background: var(--paper); text-align:left !important; white-space: normal !important; z-index: 2; }

.planning-table tbody tr:nth-child(even) .sticky-col{ background: #F3F1FA; }

.planning-table td.total-cell{

  font-family: var(--font-mono); font-weight: 800; color: var(--sky-ink); font-size: 13px;

}



.shift-pill{

  display:inline-flex; align-items:center; justify-content:center; gap:2px;

  width: 100%; min-width: 78px; padding: 7px 4px; border-radius: 8px;

  font-size: 11px; font-weight: 700; cursor: pointer; border: 1px solid transparent;

  transition: filter .12s, transform .12s;

  letter-spacing: .01em;

}

.shift-pill:hover{ filter: brightness(1.08); transform: translateY(-1px); }

.shift-pill.s-jour{ background: #059669; color: #fff; }

.shift-pill.s-nuit{ background: #B9800E; color: #fff; }

.shift-pill.s-quiet{ background: var(--sand-deep); color: var(--ink-soft); box-shadow: inset 0 0 0 1px var(--line); }

.cell-select{

  width: 100%; padding: 6px 4px; border-radius: 7px; border: 1.5px solid var(--sky); font-size: 11px;

  background: var(--paper); color: var(--ink); font-weight: 700;

}



.alerts-banner{

  background: rgba(180,83,9,0.08); border: 1px solid rgba(180,83,9,0.25); border-radius: 12px;

  padding: 12px 16px; margin-bottom: 16px; font-size: 12.5px; color: #92400E;

}

.alerts-banner b{ display:block; margin-bottom: 4px; font-size: 12.5px; }

.alerts-banner ul{ margin: 4px 0 0 18px; }



@media print{

  .sidebar, .topbar, .planning-header, .planning-tabs, .alerts-banner{ display:none !important; }

  body{ background: #fff; color:#111; }

  .view{ padding: 0; }

  .planning-card{ border:none; }

  .shift-pill{ border: 1px solid #ccc; }

}

</style>

</head>

<body>



<!-- ============ LOGIN ============ -->

<div id="loginScreen">

  <div class="login-split">

    <aside class="login-hero-panel">

      <div class="hero-brand">

        <div class="login-mark">

          <svg viewBox="0 0 24 24" fill="none" stroke="#101D30" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12l2-2 4 4 8-8 2 2"/><path d="M3 18l2-2 4 4 8-8 2 2"/></svg>

        </div>

        <span>FieldOps</span>

      </div>



      <div class="route-map">

        <svg class="route-svg" viewBox="0 0 340 480" xmlns="http://www.w3.org/2000/svg">

          <path id="routePath" d="M50,50 C 170,20 260,80 280,130 S 120,180 90,230 S 280,270 260,330 S 60,370 80,430"

                fill="none" stroke="rgba(234,241,249,0.22)" stroke-width="2.5" stroke-dasharray="1 9" stroke-linecap="round"/>



          <!-- Technopolis : puce / circuit -->

          <g transform="translate(50,50)">

            <circle r="15" fill="#3E8F6F"/>

            <rect x="-6" y="-6" width="12" height="12" rx="2" fill="none" stroke="#fff" stroke-width="1.4"/>

            <path d="M-6,-2 h-3 M-6,2 h-3 M6,-2 h3 M6,2 h3 M-2,-6 v-3 M2,-6 v-3 M-2,6 v3 M2,6 v3" stroke="#fff" stroke-width="1.4" stroke-linecap="round"/>

            <text x="20" y="4" class="pin-label">Technopolis</text>

          </g>



          <!-- Salé - Medina : dôme / croissant -->

          <g transform="translate(280,130)">

            <circle r="15" fill="#D9643A"/>

            <path d="M-6,4 a6,6 0 1,1 12,0 z" fill="#fff"/>

            <circle cx="0" cy="-4" r="2.2" fill="#fff"/>

            <path d="M-8,4 h16" stroke="#fff" stroke-width="1.4" stroke-linecap="round"/>

            <text x="-20" y="-19" text-anchor="end" class="pin-label">Salé - Medina</text>

          </g>



          <!-- Agdal - Carrousel : grande roue -->

          <g transform="translate(90,230)">

            <circle r="15" fill="#D4A24C"/>

            <circle r="6.5" fill="none" stroke="#fff" stroke-width="1.4"/>

            <path d="M0,-6.5 L0,6.5 M-6.5,0 L6.5,0 M-4.6,-4.6 L4.6,4.6 M-4.6,4.6 L4.6,-4.6" stroke="#fff" stroke-width="1.1"/>

            <text x="20" y="4" class="pin-label">Agdal - Carrousel</text>

          </g>



          <!-- Hay Riad - Soussi : villa -->

          <g transform="translate(260,330)">

            <circle r="15" fill="#7A5C8A"/>

            <path d="M-7,2 L0,-6 L7,2 V7 H-7 Z" fill="#fff"/>

            <rect x="-2" y="1" width="4" height="6" fill="#7A5C8A"/>

            <text x="-20" y="-19" text-anchor="end" class="pin-label">Hay Riad - Soussi</text>

          </g>



          <!-- Temara - Oulad Mtaa : palmier -->

          <g transform="translate(80,430)">

            <circle r="15" fill="#3D6B8A"/>

            <path d="M0,7 V-2" stroke="#fff" stroke-width="1.6" stroke-linecap="round"/>

            <path d="M0,-2 Q-7,-6 -8,-1 M0,-2 Q7,-6 8,-1 M0,-2 Q-5,-8 -1,-9 M0,-2 Q5,-8 1,-9" stroke="#fff" stroke-width="1.4" fill="none" stroke-linecap="round"/>

            <text x="20" y="4" class="pin-label">Temara - Oulad Mtaa</text>

          </g>



          <g id="routeMotion" class="scooter-icon">

            <circle r="9" fill="#101D30"/>

            <path d="M-5,2 h10 l-2,-6 h-6 z M-3,2 a3,3 0 1,0 0.1,0 M3,2 a3,3 0 1,0 0.1,0" stroke="#F5A623" stroke-width="1.4" fill="none"/>

            <animateMotion dur="16s" repeatCount="indefinite" rotate="auto"

              path="M50,50 C 170,20 260,80 280,130 S 120,180 90,230 S 280,270 260,330 S 60,370 80,430"/>

          </g>

        </svg>

      </div>



      <div class="hero-foot">

        <span id="heroClock">--:--</span>

        <div id="heroDate" style="font-size:12px; font-weight:700; color:rgba(43,36,64,0.7); text-transform:capitalize; margin-bottom:8px;"></div>

        <p>Gestion opérationnelle terrain — Rabat &amp; environs</p>

        <div id="heroMotivation" style="margin-top:14px; padding-top:14px; border-top:1px solid rgba(43,36,64,0.14); font-style:italic; font-size:12.5px; color:rgba(43,36,64,0.65); max-width:280px; line-height:1.5;"></div>

      </div>

    </aside>



    <main class="login-form-panel">

      <div class="login-card">

        <div class="login-brand">

          <h1>FieldOps</h1>

          <p>Système Central de Gestion RH</p>

        </div>



        <div class="profile-grid" id="profileGrid"></div>



        <input type="password" id="passwordInput" class="login-input" placeholder="Mot de passe ou code PIN..." disabled>

        <button id="loginBtn" class="login-submit" disabled>DÉVERROUILLER</button>

        <div class="login-error" id="loginError"></div>

      </div>

    </main>

  </div>

</div>



<!-- AGENT PORTAL (écran simplifié pour un agent connecté) -->

<div id="agentPortal" class="hidden">

  <div class="nv-badges-row">

  <div class="portal-bg-doodle dd-tl" style="background:#00A082;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="none" stroke="#00A082" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21s7-7.2 7-12a7 7 0 10-14 0c0 4.8 7 12 7 12z"/><circle cx="12" cy="9" r="2.5" fill="#00A082" stroke="none"/></svg></div>

    <div class="dd-label">Zone</div>

  </div>

  <div class="portal-bg-doodle dd-bl" style="background:#E8A317;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="none" stroke="#E8A317" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></svg></div>

    <div class="dd-label">Livraison</div>

  </div>

  <div class="portal-bg-doodle dd-tr" style="background:#DC7A0E;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="#DC7A0E" stroke="none"><path d="M12 2l2.9 6.9 7.6.6-5.8 5 1.8 7.4L12 18l-6.5 3.9 1.8-7.4-5.8-5 7.6-.6z"/></svg></div>

    <div class="dd-label">Bravo</div>

  </div>

  <div class="portal-bg-doodle dd-br" style="background:#0284C7;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="none" stroke="#0284C7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="8" r="3.5"/><path d="M2 21c0-4 3-6.5 7-6.5S16 17 16 21"/><circle cx="18" cy="9" r="2.5"/><path d="M22 21c0-3-1.8-5-4.5-5.3"/></svg></div>

    <div class="dd-label">Équipe</div>

  </div>

  <div class="portal-bg-doodle dd-ml" style="background:#DC2626;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="none" stroke="#DC2626" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg></div>

    <div class="dd-label">Ponctuel</div>

  </div>

  <div class="portal-bg-doodle dd-mr" style="background:#00A082;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="none" stroke="#00A082" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg></div>

    <div class="dd-label">Check-in</div>

  </div>

  </div>

  <div class="portal-side-art">

    <svg viewBox="0 0 320 320" xmlns="http://www.w3.org/2000/svg">

      <defs>

        <radialGradient id="artGlow" cx="50%" cy="35%" r="60%">

          <stop offset="0%" stop-color="#00A082" stop-opacity="0.35"/>

          <stop offset="100%" stop-color="#00A082" stop-opacity="0"/>

        </radialGradient>

        <linearGradient id="artScooter" x1="0" y1="0" x2="1" y2="1">

          <stop offset="0%" stop-color="#4FBFA8"/>

          <stop offset="100%" stop-color="#007A63"/>

        </linearGradient>

      </defs>

      <circle cx="160" cy="120" r="140" fill="url(#artGlow)"/>

      <g opacity="0.5" fill="#4A3B1F">

        <rect x="10" y="230" width="26" height="60" rx="2"/>

        <rect x="42" y="200" width="30" height="90" rx="2"/>

        <rect x="80" y="220" width="24" height="70" rx="2"/>

        <rect x="220" y="210" width="28" height="80" rx="2"/>

        <rect x="254" y="235" width="24" height="55" rx="2"/>

        <rect x="284" y="195" width="30" height="95" rx="2"/>

      </g>

      <line x1="0" y1="290" x2="320" y2="290" stroke="#5C4A28" stroke-width="2"/>

      <g transform="translate(90,150)">

        <ellipse cx="70" cy="140" rx="72" ry="8" fill="#000" opacity="0.25"/>

        <circle cx="18" cy="132" r="20" fill="none" stroke="url(#artScooter)" stroke-width="6"/>

        <circle cx="118" cy="132" r="20" fill="none" stroke="url(#artScooter)" stroke-width="6"/>

        <path d="M18 132 L55 132 L78 80 L100 80" fill="none" stroke="url(#artScooter)" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>

        <path d="M100 80 L118 132" fill="none" stroke="url(#artScooter)" stroke-width="7" stroke-linecap="round"/>

        <rect x="40" y="95" width="34" height="26" rx="6" fill="url(#artScooter)"/>

        <circle cx="95" cy="55" r="14" fill="#E7E9F3"/>

        <path d="M85 62 Q95 42 112 50 L120 78 Q98 90 82 78 Z" fill="url(#artScooter)"/>

        <rect x="60" y="35" width="26" height="18" rx="4" fill="#34D399"/>

      </g>

      <g fill="#4FBFA8" opacity="0.8">

        <circle cx="45" cy="60" r="3"/>

        <circle cx="270" cy="45" r="4"/>

        <circle cx="290" cy="140" r="3"/>

      </g>

      <path d="M250 60 L253 68 L261 68 L254 73 L257 81 L250 76 L243 81 L246 73 L239 68 L247 68 Z" fill="#FBBF24"/>

    </svg>

    <div class="portal-side-art-quote">"Chaque livraison compte.<br>On avance, ensemble."</div>

  </div>

  <div class="portal-card">

    <div class="portal-header">

      <div>

        <div style="font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.04em; color:var(--ink-soft); margin-bottom:2px;">Bienvenue 👋</div>

        <b id="portalAgentName">—</b>

        <span id="portalClock">--:--</span>

      </div>

      <button class="btn-ghost" id="portalLogoutBtn">Déconnexion</button>

    </div>

    <div id="portalMotivation"></div>

    <div class="portal-status" id="portalStatus">Chargement...</div>

    <button class="portal-btn" id="portalCheckBtn">Check-in</button>

    <div class="portal-history">

      <div class="section-title" style="margin-top:22px;">Historique récent</div>

      <div id="portalHistoryList"></div>

    </div>

  </div>

</div>



<!-- PLANNING NV — PORTAIL AGENT (écran simplifié, indépendant, avec localisation) -->

<div id="nvAgentPortal" class="hidden">

  <div id="nvPortalTopWeekWrap" class="nv-portal-top-week hidden"></div>

  <div class="nv-badges-row">

  <div class="portal-bg-doodle dd-tl" style="background:#00A082;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="none" stroke="#00A082" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21s7-7.2 7-12a7 7 0 10-14 0c0 4.8 7 12 7 12z"/><circle cx="12" cy="9" r="2.5" fill="#00A082" stroke="none"/></svg></div>

    <div class="dd-label">Zone</div>

  </div>

  <div class="portal-bg-doodle dd-bl" style="background:#E8A317;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="none" stroke="#E8A317" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></svg></div>

    <div class="dd-label">Livraison</div>

  </div>

  <div class="portal-bg-doodle dd-tr" style="background:#DC7A0E;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="#DC7A0E" stroke="none"><path d="M12 2l2.9 6.9 7.6.6-5.8 5 1.8 7.4L12 18l-6.5 3.9 1.8-7.4-5.8-5 7.6-.6z"/></svg></div>

    <div class="dd-label">Bravo</div>

  </div>

  <div class="portal-bg-doodle dd-br" style="background:#0284C7;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="none" stroke="#0284C7" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="8" r="3.5"/><path d="M2 21c0-4 3-6.5 7-6.5S16 17 16 21"/><circle cx="18" cy="9" r="2.5"/><path d="M22 21c0-3-1.8-5-4.5-5.3"/></svg></div>

    <div class="dd-label">Équipe</div>

  </div>

  <div class="portal-bg-doodle dd-ml" style="background:#DC2626;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="none" stroke="#DC2626" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg></div>

    <div class="dd-label">Ponctuel</div>

  </div>

  <div class="portal-bg-doodle dd-mr" style="background:#00A082;">

    <div class="dd-badge"><svg viewBox="0 0 24 24" fill="none" stroke="#00A082" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg></div>

    <div class="dd-label">Check-in</div>

  </div>

  </div>

  <div class="portal-side-art">

    <svg viewBox="0 0 320 320" xmlns="http://www.w3.org/2000/svg">

      <defs>

        <radialGradient id="artGlow2" cx="50%" cy="35%" r="60%">

          <stop offset="0%" stop-color="#00A082" stop-opacity="0.35"/>

          <stop offset="100%" stop-color="#00A082" stop-opacity="0"/>

        </radialGradient>

        <linearGradient id="artScooter2" x1="0" y1="0" x2="1" y2="1">

          <stop offset="0%" stop-color="#4FBFA8"/>

          <stop offset="100%" stop-color="#007A63"/>

        </linearGradient>

      </defs>

      <circle cx="160" cy="120" r="140" fill="url(#artGlow2)"/>

      <g opacity="0.5" fill="#4A3B1F">

        <rect x="10" y="230" width="26" height="60" rx="2"/>

        <rect x="42" y="200" width="30" height="90" rx="2"/>

        <rect x="80" y="220" width="24" height="70" rx="2"/>

        <rect x="220" y="210" width="28" height="80" rx="2"/>

        <rect x="254" y="235" width="24" height="55" rx="2"/>

        <rect x="284" y="195" width="30" height="95" rx="2"/>

      </g>

      <line x1="0" y1="290" x2="320" y2="290" stroke="#5C4A28" stroke-width="2"/>

      <g transform="translate(90,150)">

        <ellipse cx="70" cy="140" rx="72" ry="8" fill="#000" opacity="0.25"/>

        <circle cx="18" cy="132" r="20" fill="none" stroke="url(#artScooter2)" stroke-width="6"/>

        <circle cx="118" cy="132" r="20" fill="none" stroke="url(#artScooter2)" stroke-width="6"/>

        <path d="M18 132 L55 132 L78 80 L100 80" fill="none" stroke="url(#artScooter2)" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>

        <path d="M100 80 L118 132" fill="none" stroke="url(#artScooter2)" stroke-width="7" stroke-linecap="round"/>

        <rect x="40" y="95" width="34" height="26" rx="6" fill="url(#artScooter2)"/>

        <circle cx="95" cy="55" r="14" fill="#E7E9F3"/>

        <path d="M85 62 Q95 42 112 50 L120 78 Q98 90 82 78 Z" fill="url(#artScooter2)"/>

        <rect x="60" y="35" width="26" height="18" rx="4" fill="#34D399"/>

      </g>

      <g fill="#4FBFA8" opacity="0.8">

        <circle cx="45" cy="60" r="3"/>

        <circle cx="270" cy="45" r="4"/>

        <circle cx="290" cy="140" r="3"/>

      </g>

      <path d="M250 60 L253 68 L261 68 L254 73 L257 81 L250 76 L243 81 L246 73 L239 68 L247 68 Z" fill="#FBBF24"/>

    </svg>

    <div class="portal-side-art-quote">"Ton travail fait la<br>différence, chaque jour."</div>

  </div>

  <div class="portal-card">

    <div class="portal-header">

      <div>

        <div style="font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.04em; color:var(--ink-soft); margin-bottom:2px;">Bienvenue 👋</div>

        <b id="nvPortalAgentName">—</b>

        <span id="nvPortalClock">--:--</span>

      </div>

      <div style="display:flex; align-items:center; gap:8px;">

        <button class="agent-access-btn" id="nvPortalRapportBtn" title="Rapport &amp; remarques" style="font-size:14px; padding:9px 16px; font-weight:800; position:relative;">📝 Rapport<span id="nvRapportReplyBadge" class="hidden" style="position:absolute; top:-6px; right:-6px; background:var(--red); color:#fff; font-size:10px; font-weight:800; min-width:16px; height:16px; border-radius:8px; display:flex; align-items:center; justify-content:center; padding:0 3px;">●</span></button>

        <button class="btn-ghost" id="nvPortalLogoutBtn">Déconnexion</button>

      </div>

    </div>

    <div id="nvPortalContent"></div>

  </div>

</div>



<!-- ============ APP ============ -->

<div id="app" style="display:none;">

  <div class="shell">

    <aside class="sidebar">

      <div class="sidebar-brand">

        <div class="mark">O</div>

        <h2>FieldOps</h2>

      </div>



      <div class="nav-group">

        <div class="nav-label">Opérations</div>

        <button class="nav-item" data-view="dashboard"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/></svg><span>Dashboard</span></button>

        <button class="nav-item" data-view="planningnv"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="17" rx="2"/><path d="M3 9h18M8 3v3M16 3v3"/></svg><span>Planning</span></button>

        <button class="nav-item" data-view="pointage"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg><span>Pointage</span></button>

        <button class="nav-item" data-view="absences"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 19h16M4 19V7a2 2 0 012-2h4l2 2h6a2 2 0 012 2v10"/></svg><span>Absences</span></button>

      </div>



      <div class="nav-group">

        <div class="nav-label">Ressources</div>

        <button class="nav-item" data-view="agents"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 11-2.83 2.83l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 11-2.83-2.83l.06-.06A1.65 1.65 0 004.6 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 112.83-2.83l.06.06A1.65 1.65 0 009 4.6a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09c0 .68.39 1.29 1 1.51.35.14.74.18 1.11.09"/></svg><span>Paramètres</span></button>

        <button class="nav-item" data-view="paie"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="6" width="20" height="13" rx="2"/><circle cx="12" cy="12.5" r="3"/><path d="M6 6V5a2 2 0 012-2h8a2 2 0 012 2v1"/></svg><span>Paie</span></button>

        <button class="nav-item" data-view="rapports" style="position:relative;"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><path d="M14 2v6h6M9 13h6M9 17h6M9 9h1"/></svg><span>Rapports</span><span id="rapportsNavBadge" class="hidden" style="position:absolute; top:4px; right:8px; background:var(--red); color:#fff; font-size:10px; font-weight:800; min-width:18px; height:18px; border-radius:9px; display:flex; align-items:center; justify-content:center; padding:0 4px;"></span></button>

      </div>



      <div class="sidebar-footer">

        <div class="user-chip">

          <div class="dot"></div>

          <div class="user-chip-text">

            <b id="sidebarUserName">—</b>

            <span id="sidebarUserRole">—</span>

          </div>

        </div>

        <button class="logout-btn" id="changePassBtn" style="margin-top:6px;"><span>Changer le mot de passe</span></button>

        <button class="logout-btn" id="logoutBtn"><span>Déconnexion</span></button>

      </div>

    </aside>



    <main class="main">

      <div class="topbar">

        <h1 id="viewTitle">Dashboard</h1>

        <div class="topbar-meta">

          <div class="sync-pill"><span class="dot"></span>Cloud Sync Active</div>

          <span id="clock">--:--</span>

        </div>

      </div>



      <!-- DASHBOARD -->

      <section class="view" id="view-dashboard">

        <div id="dashboardNewRapportsBanner" class="hidden"></div>

        <div class="stat-grid">

          <div class="stat-card">

            <div class="label">Agents actifs</div>

            <div class="value" id="statActive">—</div>

            <div class="sub">enregistrés</div>

          </div>

          <div class="stat-card accent-green">

            <div class="label">En shift JOUR</div>

            <div class="value" id="statJour">—</div>

            <div class="sub">équipe matin</div>

          </div>

          <div class="stat-card accent-blue">

            <div class="label">En shift NUIT</div>

            <div class="value" id="statNuit">—</div>

            <div class="sub">équipe soir</div>

          </div>

          <div class="stat-card accent-amber">

            <div class="label">Zones couvertes</div>

            <div class="value" id="statZones">—</div>

            <div class="sub">sur 5 zones</div>

          </div>

          <div class="stat-card accent-purple">

            <div class="label">OFF aujourd'hui</div>

            <div class="value" id="statOffToday">—</div>

            <div class="sub">agents au repos</div>

          </div>

          <div class="stat-card accent-orange">

            <div class="label">Absences (semaine)</div>

            <div class="value" id="statAbsences">—</div>

            <div class="sub">congé / maladie</div>

          </div>

          <div class="stat-card accent-red">

            <div class="label">Alertes actives</div>

            <div class="value" id="statAlerts">—</div>

            <div class="sub">à vérifier</div>

          </div>

        </div>



        <div class="section-title">Zones</div>

        <div class="zone-grid" id="zoneGrid"></div>



        <div class="section-title">🟢 En service maintenant</div>

        <div class="card" id="liveCard">

          <table>

            <thead><tr><th>Agent</th><th>Zone</th><th>Shift</th><th>Fin</th></tr></thead>

            <tbody id="liveTableBody"></tbody>

          </table>

          <div class="placeholder-view hidden" id="liveEmptyState">

            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/></svg>

            <h3>Personne en service actuellement</h3>

            <p>Génère le planning de la semaine pour voir qui est en service en temps réel.</p>

          </div>

        </div>

      </section>



      <!-- AGENTS -->

      <section class="view hidden" id="view-agents">

        <div class="view-header">

          <p class="view-subtitle">Gérez la liste des agents de terrain et leur configuration de rotation.</p>

          <div style="display:flex; gap:8px;">

            <button class="btn-ghost" id="syncNvAgentsBtn" style="border:1px solid var(--line);">🔄 Synchroniser avec Planning NV</button>

            <button class="btn-ghost" id="genAllPinsBtn" style="border:1px solid var(--line);">🎲 Générer les PIN manquants</button>

            <button class="btn-ghost" id="genAllNicknamesBtn" style="border:1px solid var(--line);">✏️ Surnoms manquants</button>

            <button class="btn-ghost" id="fixDefaultsBtn" style="border:1px solid var(--line);">🎯 Jour/Nuit/OFF par défaut</button>

            <button class="btn-ghost" id="renameZonesBtn" style="border:1px solid var(--line);">🏷️ Renommer les zones</button>

            <button class="btn-primary" id="openAddAgentBtn">+ Ajouter un agent</button>

          </div>

        </div>

        <div class="card" id="agentsCard">

          <table>

            <thead><tr><th>Agent</th><th>PIN</th><th>Email</th><th>Zone</th><th>Type</th><th>Shift fixe</th><th></th></tr></thead>

            <tbody id="agentsTableBody"></tbody>

          </table>

          <div class="placeholder-view hidden" id="agentsEmptyState">

            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="9" cy="8" r="3.5"/><path d="M2 21c0-4 3-6.5 7-6.5S16 17 16 21"/><circle cx="18" cy="9" r="2.5"/><path d="M22 21c0-3-1.8-5-4.5-5.3"/></svg>

            <h3>Aucun agent pour l'instant</h3>

            <p>Cliquez sur "Ajouter un agent" pour commencer à construire votre équipe.</p>

          </div>

        </div>

      </section>



      <!-- PLANNING -->

      <section class="view hidden" id="view-planning">

        <div class="planning-header">

          <div class="week-nav">

            <button class="icon-btn-lg" id="prevWeekBtn" title="Semaine précédente">

              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 18l-6-6 6-6"/></svg>

            </button>

            <div class="week-label">

              <b id="weekLabel">Semaine —</b>

              <span id="weekRange">—</span>

            </div>

            <button class="icon-btn-lg" id="nextWeekBtn" title="Semaine suivante">

              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 18l6-6-6-6"/></svg>

            </button>

          </div>

          <div class="planning-actions">

            <button class="btn-ghost" id="printPlanningBtn">🖨️ Imprimer / PDF</button>

            <button class="btn-ghost" id="exportExcelBtn">⬇️ Excel</button>

            <button class="btn-ghost" id="loadRealWeekBtn" style="border:1px solid var(--line);">📥 Charger planning réel (cette semaine)</button>

            <button class="btn-primary" id="generatePlanningBtn">⚡ Générer Planning</button>

          </div>

        </div>



        <div id="alertsBanner" class="alerts-banner hidden"></div>



        <div class="planning-tabs">

          <button class="ptab active" data-ptab="grid">Vue Agents</button>

          <button class="ptab" data-ptab="zones">Vue Zones</button>

        </div>



        <div class="card planning-card" id="planningGridWrap">

          <datalist id="agentNamesList"></datalist>

          <table class="planning-table" id="planningTable">

            <thead>

              <tr>

                <th class="sticky-col">Agent</th>

                <th>Zone</th>

                <th id="thLun">Lun</th><th id="thMar">Mar</th><th id="thMer">Mer</th><th id="thJeu">Jeu</th><th id="thVen">Ven</th><th id="thSam">Sam</th><th id="thDim">Dim</th>

                <th>Total</th>

                <th></th>

              </tr>

            </thead>

            <tbody id="planningTableBody"></tbody>

          </table>

          <div class="placeholder-view hidden" id="planningEmptyState">

            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="4" width="18" height="17" rx="2"/><path d="M3 9h18M8 3v3M16 3v3"/></svg>

            <h3>Aucun planning pour cette semaine</h3>

            <p>Cliquez sur "Générer Planning" pour créer automatiquement la semaine (agents avec un groupe de planification assigné).</p>

          </div>

        </div>



        <div class="card planning-card hidden" id="zonesGridWrap">

          <table class="planning-table">

            <thead>

              <tr><th class="sticky-col">Zone</th><th id="thZLun">Lun</th><th id="thZMar">Mar</th><th id="thZMer">Mer</th><th id="thZJeu">Jeu</th><th id="thZVen">Ven</th><th id="thZSam">Sam</th><th id="thZDim">Dim</th></tr>

            </thead>

            <tbody id="zonesTableBody"></tbody>

          </table>

        </div>

      </section>



      <!-- PLACEHOLDERS -->

      <section class="view hidden" id="view-pointage">

        <div class="planning-header">

          <div class="week-nav">

            <button class="icon-btn-lg" id="prevDayBtn" title="Jour précédent">

              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 18l-6-6 6-6"/></svg>

            </button>

            <div class="week-label">

              <b id="pointageDayLabel">—</b>

              <span id="pointageDateSub">—</span>

            </div>

            <button class="icon-btn-lg" id="nextDayBtn" title="Jour suivant">

              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 18l6-6-6-6"/></svg>

            </button>

          </div>

          <button class="btn-ghost" id="pointageTodayBtn" style="border:1px solid var(--line);">Aujourd'hui</button>

        </div>

        <div class="card">

          <table>

            <thead><tr><th>Agent</th><th>Check-in</th><th>Check-out</th><th>Statut</th></tr></thead>

            <tbody id="pointageTableBody"></tbody>

          </table>

        </div>

      </section>

      <section class="view hidden" id="view-absences">

        <div class="planning-header">

          <div class="week-nav">

            <button class="icon-btn-lg" id="prevCycleBtn" title="Cycle précédent">

              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 18l-6-6 6-6"/></svg>

            </button>

            <div class="week-label">

              <b id="cycleLabel">—</b>

              <span id="cycleSub">—</span>

            </div>

            <button class="icon-btn-lg" id="nextCycleBtn" title="Cycle suivant">

              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 18l6-6-6-6"/></svg>

            </button>

          </div>

          <div class="planning-actions">

            <button class="btn-ghost" id="cycleTodayBtn" style="border:1px solid var(--line);">Cycle actuel</button>

            <button class="btn-ghost" id="exportAbsencesBtn" style="border:1px solid var(--line);">⬇️ Excel (cycle)</button>

            <button class="btn-ghost" id="exportRangeBtn" style="border:1px solid var(--line);">📅 Excel (plage)</button>

            <button class="btn-primary" id="addAbsenceBtn">+ Déclarer</button>

          </div>

        </div>



        <div class="section-title">Résumé du cycle (base 26 jours)</div>

        <div class="card" style="margin-bottom:24px; overflow-x:auto;">

          <table>

            <thead><tr><th>Agent</th><th>Maladie</th><th>Injustifiée</th><th>Congé</th><th>Total absences</th><th>Jours présence</th><th>Férié</th><th>Doublé</th><th>OFF travaillé</th><th>Total ajusté</th></tr></thead>

            <tbody id="absencesSummaryBody"></tbody>

          </table>

        </div>



        <div class="section-title">Détail du cycle</div>

        <div class="card">

          <table>

            <thead><tr><th>Agent</th><th>Date</th><th>Type</th><th>Note</th><th></th></tr></thead>

            <tbody id="absencesDetailBody"></tbody>

          </table>

          <div class="placeholder-view hidden" id="absencesEmptyState">

            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M4 19h16M4 19V7a2 2 0 012-2h4l2 2h6a2 2 0 012 2v10"/></svg>

            <h3>Aucune absence déclarée ce cycle</h3>

            <p>Clique sur "+ Déclarer" pour ajouter une absence, un jour férié travaillé ou un jour doublé.</p>

          </div>

        </div>

      </section>

      <section class="view hidden" id="view-rapports">

        <div class="view-header">

          <div>

            <div class="view-subtitle">Rapports et remarques envoyés par les agents Planning NV</div>

          </div>

          <div style="display:flex; gap:8px; align-items:center;">

            <select id="rapportsFilterAgent" class="login-input" style="width:auto; margin:0; padding:8px 12px;">

              <option value="">Tous les agents</option>

            </select>

            <select id="rapportsFilterType" class="login-input" style="width:auto; margin:0; padding:8px 12px;">

              <option value="">Tous les types</option>

              <option value="rapport">Rapport</option>

              <option value="urgent">🚨 Urgent</option>

              <option value="remarque">Remarque</option>

            </select>

            <button class="btn-ghost" id="rapportsExportBtn" style="border:1px solid var(--line);">↓ Export CSV</button>

          </div>

        </div>



        <div class="stat-grid" id="rapportsStatGrid" style="margin-bottom:20px;"></div>



        <div class="card" id="rapportsMissingCard" style="padding:16px 20px; margin-bottom:20px;"></div>



        <div class="card">

          <div id="rapportsList" style="padding:8px 0;"></div>

          <div class="placeholder-view hidden" id="rapportsEmptyState">

            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><path d="M14 2v6h6M16 13H8M16 17H8M10 9H8"/></svg>

            <h3>Aucun rapport pour l'instant</h3>

            <p>Les rapports et remarques envoyés par les agents depuis leur portail apparaîtront ici.</p>

          </div>

        </div>

      </section>

      <section class="view hidden" id="view-paie"></section>

      <section class="view hidden" id="view-planningnv">

        <div class="planning-header">

          <div class="week-nav">

            <button class="icon-btn-lg" id="prevWeekNvBtn" title="Semaine précédente">

              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 18l-6-6 6-6"/></svg>

            </button>

            <div class="week-label">

              <b id="weekLabelNv">Semaine —</b>

              <span id="weekRangeNv">—</span>

            </div>

            <button class="icon-btn-lg" id="nextWeekNvBtn" title="Semaine suivante">

              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 18l6-6-6-6"/></svg>

            </button>

          </div>

          <p class="view-subtitle" style="margin:0;">Espace indépendant — aucun lien avec les Agents, Zones ou l'autre Planning.</p>

          <button class="btn-ghost" id="showNvFlashBtn" style="border:1px solid var(--line);">📢 Info équipe</button>

        </div>

        <div class="card hidden" id="nvFlashPanel" style="padding:16px 20px; margin-bottom:16px;"></div>

        <div class="planning-tabs">

          <button class="ptab active" data-nv-view="week">📊 Vue Semaine</button>

          <button class="ptab" data-nv-view="zones">🗺️ Vue Zones</button>

        </div>

        <div class="card planning-card hidden" id="nvWeekGrid"></div>

        <div class="hidden" id="nvZoneRestoreWrap" style="margin-bottom:12px; display:flex; gap:8px; flex-wrap:wrap;">

          <button class="btn-ghost" id="nvExportHotspotsBtn" style="border:1px solid var(--sky); color:var(--sky-ink); font-size:12px;">💾 Sauvegarder les hotspots (copier le texte)</button>

          <button class="btn-ghost" id="nvRestoreHotspotsBtn" style="border:1px solid var(--red); color:var(--red); font-size:12px;">🔄 Restaurer les hotspots par défaut (Agdal + Hay Riad reconstruits)</button>

        </div>

        <div class="nv-zone-grid" id="nvZoneGrid"></div>

      </section>

    </main>

  </div>

</div>



<!-- RANGE EXPORT MODAL -->

<div class="modal-overlay hidden" id="rangeExportOverlay">

  <div class="modal-card">

    <h3>Exporter une plage personnalisée</h3>

    <label class="form-label">Du</label>

    <input type="date" id="rangeFromInput" class="login-input">

    <label class="form-label">Au</label>

    <input type="date" id="rangeToInput" class="login-input">

    <div class="modal-error" id="rangeExportError"></div>

    <div class="modal-actions">

      <button class="btn-ghost" id="cancelRangeExportBtn">Annuler</button>

      <button class="login-submit" id="confirmRangeExportBtn" style="width:auto; padding:12px 22px;">Exporter</button>

    </div>

  </div>

</div>



<!-- RENAME ZONES MODAL -->

<div class="modal-overlay hidden" id="renameZonesOverlay">

  <div class="modal-card">

    <h3>Renommer les zones</h3>

    <p class="modal-hint" style="margin-top:0; margin-bottom:14px;">Ça change juste le nom affiché — la rotation et les couleurs restent les mêmes.</p>

    <div id="zoneRenameInputs"></div>

    <div class="modal-error" id="renameZonesError"></div>

    <div class="modal-actions">

      <button class="btn-ghost" id="cancelRenameZonesBtn">Annuler</button>

      <button class="login-submit" id="saveRenameZonesBtn" style="width:auto; padding:12px 22px;">Enregistrer</button>

    </div>

  </div>

</div>



<!-- ABSENCE MODAL -->

<div class="modal-overlay hidden" id="absenceModalOverlay">

  <div class="modal-card">

    <h3 id="absenceModalTitle">Déclarer une absence / un jour</h3>

    <label class="form-label">Agent</label>

    <select id="absenceAgentInput" class="login-input"></select>



    <label class="form-label">Type</label>

    <select id="absenceTypeInput" class="login-input">

      <option value="maladie">Absence maladie (−1 jour)</option>

      <option value="injustifiee">Absence injustifiée (−1 jour)</option>

      <option value="conge">Congé payé (−1 jour)</option>

      <option value="ferie">Jour férié travaillé (+1 jour)</option>

      <option value="double">Jour doublé (+1 jour)</option>

      <option value="off_travaille">Jour OFF travaillé (+1 jour)</option>

    </select>



    <label class="form-label">Date (début)</label>

    <input type="date" id="absenceDateInput" class="login-input">



    <label class="form-label">Date de fin (optionnel — pour déclarer plusieurs jours d'un coup)</label>

    <input type="date" id="absenceDateEndInput" class="login-input">



    <label class="form-label">Note (optionnel)</label>

    <input type="text" id="absenceNoteInput" class="login-input" placeholder="Ex: certificat médical...">



    <div class="modal-error" id="absenceModalError"></div>



    <div class="modal-actions">

      <button class="btn-ghost" id="cancelAbsenceBtn">Annuler</button>

      <button class="login-submit" id="saveAbsenceBtn" style="width:auto; padding:12px 22px;">Enregistrer</button>

    </div>

  </div>

</div>



<!-- CHANGE PASSWORD MODAL -->

<div class="modal-overlay hidden" id="changePassOverlay">

  <div class="modal-card">

    <h3>Changer le mot de passe</h3>

    <label class="form-label">Mot de passe actuel</label>

    <input type="password" id="oldPassInput" class="login-input" placeholder="Mot de passe actuel">

    <label class="form-label">Nouveau mot de passe</label>

    <input type="password" id="newPassInput" class="login-input" placeholder="Nouveau mot de passe (6 caractères min.)">

    <label class="form-label">Confirmer le nouveau mot de passe</label>

    <input type="password" id="confirmPassInput" class="login-input" placeholder="Confirmer">

    <div class="modal-error" id="changePassError"></div>

    <div class="modal-actions">

      <button class="btn-ghost" id="cancelChangePassBtn">Annuler</button>

      <button class="login-submit" id="saveChangePassBtn" style="width:auto; padding:12px 22px;">Enregistrer</button>

    </div>

  </div>

</div>



<!-- AGENT MODAL -->

<div class="modal-overlay hidden" id="agentModalOverlay">

  <div class="modal-card">

    <h3 id="agentModalTitle">Ajouter un agent</h3>

    <label class="form-label">Nom complet</label>

    <input type="text" id="agentNameInput" class="login-input" placeholder="Ex: Achraf Akhlij">



    <label class="form-label">Surnom (affiché dans le login Agent)</label>

    <input type="text" id="agentNicknameInput" class="login-input" placeholder="Ex: Achraf">



    <label class="form-label">Code PIN (pour se connecter au portail agent)</label>

    <div style="display:flex; gap:8px; align-items:center;">

      <input type="text" id="agentPinModalInput" class="login-input" placeholder="Ex: 1234" inputmode="numeric" maxlength="8" style="margin-bottom:0;">

      <button type="button" class="btn-ghost" id="genPinBtn" style="border:1px solid var(--line); white-space:nowrap; padding:13px 14px;" title="Générer un PIN aléatoire">🎲</button>

    </div>



    <label class="form-label">Email (pour envoyer l'accès et le planning)</label>

    <input type="email" id="agentEmailModalInput" class="login-input" placeholder="Ex: prenom.nom@gmail.com">



    <label class="form-label">Groupe de planification</label>

    <select id="agentGroupInput" class="login-input">

      <option value="">Aucun (assignation manuelle)</option>

      <option value="zakariae">Fixe solo — zone dédiée</option>

      <option value="sale">Fixe équipe — zone dédiée (rotation Matin/Nuit)</option>

      <option value="trio1">Rotation — Groupe 1</option>

      <option value="trio2">Rotation — Groupe 2</option>

      <option value="trio3">Rotation — Groupe 3</option>

    </select>



    <div id="manualZoneGroup">

      <label class="form-label">Zone</label>

      <select id="agentZoneInput" class="login-input"></select>

    </div>



    <div id="fixedSoloZoneGroup" class="hidden">

      <label class="form-label">Zone dédiée (fixe)</label>

      <select id="agentFixedZoneInput" class="login-input"></select>

    </div>



    <p class="modal-hint hidden" id="rotationHint">La zone de ce groupe change chaque semaine selon la rotation automatique (Agdal - Carrousel / Hay Riad - Soussi / Temara - Oulad Mtaa).</p>

    <p class="modal-hint hidden" id="saleHint">Zone fixe : Salé - Medina.</p>



    <label class="form-label">Période habituelle (fixe — ne tourne plus automatiquement)</label>

    <select id="agentDefaultPeriodInput" class="login-input">

      <option value="JOUR">☀️ Jour</option>

      <option value="NUIT">🌙 Nuit</option>

    </select>



    <label class="form-label">Jour OFF habituel (fixe)</label>

    <select id="agentDefaultOffDayInput" class="login-input">

      <option value="0">Lundi</option>

      <option value="1">Mardi</option>

      <option value="2">Mercredi</option>

      <option value="3">Jeudi</option>

    </select>

    <p class="modal-hint">Ces deux réglages restent les mêmes à chaque "Générer Planning", jusqu'à ce que tu les changes ici. Seule la zone (pour les groupes en rotation) change automatiquement chaque semaine.</p>



    <div class="modal-error" id="agentModalError"></div>



    <div class="modal-actions">

      <button class="btn-ghost" id="cancelAgentBtn">Annuler</button>

      <button class="login-submit" id="saveAgentBtn" style="width:auto; padding:12px 22px;">Enregistrer</button>

    </div>

  </div>

</div>



<script>

/* ================= FIREBASE INIT ================= */

const firebaseConfig = {

  apiKey: "AIzaSyAg7ayUgGUc1ZUn0_Zg9BqNDMmDnjSIBLw",

  authDomain: "ops-hq-v2-848f6.firebaseapp.com",

  projectId: "ops-hq-v2-848f6",

  storageBucket: "ops-hq-v2-848f6.firebasestorage.app",

  messagingSenderId: "939361423052",

  appId: "1:939361423052:web:88435f237b1542ae0f3c69"

};

firebase.initializeApp(firebaseConfig);

const auth = firebase.auth();

const db = firebase.firestore();



/* ================= AGENTS DATA ================= */

// L'app démarre vide — les agents sont ajoutés par l'utilisateur et stockés dans Firestore (collection "agents").

let AGENTS = [];

let agentsUnsub = null;

const ZONES = ["Agdal - Carrousel", "Hay Riad - Soussi", "Salé - Medina", "Technopolis", "Temara - Oulad Mtaa"];

// Identité couleur par zone (même code couleur que la carte de la page de login)

const ZONE_COLORS = {

  "Technopolis": "#3E8F6F",

  "Salé - Medina": "#D9643A",

  "Agdal - Carrousel": "#D4A24C",

  "Hay Riad - Soussi": "#7A5C8A",

  "Temara - Oulad Mtaa": "#3D6B8A",

};

// Noms affichés personnalisés (clé = nom interne fixe ci-dessus, valeur = nom choisi par l'utilisateur).

// La rotation/couleurs utilisent toujours le nom interne — seul l'affichage change.

let ZONE_LABELS = {};

let zoneLabelsUnsub = null;

function getZoneLabel(zoneKey){

  return (ZONE_LABELS && ZONE_LABELS[zoneKey]) ? ZONE_LABELS[zoneKey] : zoneKey;

}

function listenZoneLabels(){

  if(zoneLabelsUnsub) zoneLabelsUnsub();

  zoneLabelsUnsub = db.collection('settings').doc('zoneLabels').onSnapshot(doc => {

    ZONE_LABELS = doc.exists ? (doc.data() || {}) : {};

    initDashboard();

    renderAgentsTable();

    populateZoneSelect();

    if(!document.getElementById('view-planning').classList.contains('hidden')) renderPlanningWeek();

  }, err => console.error(err));

}



/* ================= PLANNING GROUPS ================= */

// Les 3 zones qui tournent chaque semaine entre les groupes "rotating"

const ROTATING_ZONES = ["Agdal - Carrousel", "Hay Riad - Soussi", "Temara - Oulad Mtaa"];

const GROUPS = {

  zakariae: { label: "Fixe solo", mode: "fixed_solo", zone: null },   // zone choisie par l'utilisateur (agentFixedZoneInput)

  sale:     { label: "Fixe équipe (Salé)", mode: "fixed_team", zone: "Salé - Medina" },

  trio1:    { label: "Rotation — Groupe 1", mode: "rotating_team", offset: 0 },

  trio2:    { label: "Rotation — Groupe 2", mode: "rotating_team", offset: 1 },

  trio3:    { label: "Rotation — Groupe 3", mode: "rotating_team", offset: 2 },

};

const DAY_KEYS = ["lun","mar","mer","jeu","ven","sam","dim"];

const DAY_LABELS = ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"];

// Catalogue des créneaux horaires possibles

const SHIFT_CATALOG = {

  OFF:    { label: "OFF",         hours: 0, cls: "s-quiet"  },

  M7:     { label: "10:00-17:00", hours: 7, cls: "s-jour"   },

  M8:     { label: "10:00-18:00", hours: 8, cls: "s-jour"   },

  N7:     { label: "18:00-01:00", hours: 7, cls: "s-nuit"   },

  N8:     { label: "18:00-02:00", hours: 8, cls: "s-nuit"   },

  N7B:    { label: "17:00-00:00", hours: 7, cls: "s-nuit"   },

  N8B:    { label: "17:00-01:00", hours: 8, cls: "s-nuit"   },

  N7C:    { label: "19:00-02:00", hours: 7, cls: "s-nuit"   },

  N8C:    { label: "19:00-03:00", hours: 8, cls: "s-nuit"   },

  ABS:    { label: "Absent",      hours: 0, cls: "s-quiet"  },

  MALADE: { label: "Malade",      hours: 0, cls: "s-quiet"  },

  CONGE:  { label: "Congé",       hours: 0, cls: "s-quiet"  },

};



// Login profiles shown on the lock screen (admin + leads)

const PROFILES = [

  { id: "najib", name: "Najib", role: "Harati", initials: "NH", color: "#3A6EA5", authType: "firebase" },

  { id: "oussama_h", name: "Oussama", role: "El Hamali", initials: "OE", color: "#26241D", authType: "firebase" },

];



/* ================= LOGIN ================= */

let selectedProfile = null;



function renderProfiles(){

  const grid = document.getElementById('profileGrid');

  const passInput = document.getElementById('passwordInput');

  const loginBtn = document.getElementById('loginBtn');

  grid.innerHTML = '<div style="grid-column: 1/-1; text-align:center; font-size:12px; color:var(--ink-soft); padding:10px;">Chargement de l\'équipe...</div>';



  const AVATAR_COLORS = ['#1B5C5A', '#D9643A', '#5C8A6B', '#B23B3B', '#7A5C8A', '#3D6B8A', '#D4A24C', '#4A7A6E', '#A0526D', '#6B7A3A', '#8A6A3A', '#5A6B7A', '#9A5A3A'];

  const PERSON_ICON = `<svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8"/></svg>`;



  // Il faut une session (même anonyme) pour lire Firestore selon nos règles de sécurité.

  auth.signInAnonymously()

    .then(() => {

      // L'ancien système (collection "agents", 13 agents historiques) est retiré de la grille de

      // connexion : Planning NV est désormais le seul système actif pour les agents de terrain.

      const nvProfiles = NV_ALL_AGENTS

        .map(fullName => ({ id: 'nv::' + fullName, name: fullName.split(' ')[0], fullName, role: fullName.split(' ').slice(1).join(' '), authType: 'nv-pin' }))

        .sort((a, b) => a.name.localeCompare(b.name));



      const allProfiles = [...PROFILES, ...nvProfiles]

        .map((p, i) => ({ ...p, color: p.color || AVATAR_COLORS[i % AVATAR_COLORS.length] }));



      grid.innerHTML = '';

      allProfiles.forEach(p => {

        const btn = document.createElement('button');

        btn.className = 'profile-btn';

        btn.innerHTML = `<div class="profile-avatar" style="background:${p.color}">${PERSON_ICON}</div><span>${p.name}</span><small>${p.role}</small>`;

        btn.addEventListener('click', () => selectProfile(p, btn));

        grid.appendChild(btn);

      });

      passInput.disabled = false;

    })

    .catch(err => {

      console.error(err);

      grid.innerHTML = '<div style="grid-column: 1/-1; text-align:center; font-size:12px; color:var(--red); padding:10px;">Erreur de chargement. Vérifie que "Anonyme" est activé dans Firebase Authentication.</div>';

    });

}



function selectProfile(profile, btnEl){

  selectedProfile = profile;

  document.querySelectorAll('.profile-btn').forEach(b => b.classList.remove('selected'));

  btnEl.classList.add('selected');

  document.getElementById('loginBtn').disabled = false;

  document.getElementById('passwordInput').placeholder = (profile.authType === 'pin' || profile.authType === 'nv-pin') ? 'Code PIN...' : "Mot de passe d'accès...";

  document.getElementById('passwordInput').value = '';

  document.getElementById('loginError').textContent = '';

  document.getElementById('passwordInput').focus();

}



document.getElementById('loginBtn').addEventListener('click', attemptLogin);

document.getElementById('passwordInput').addEventListener('keydown', e => { if(e.key === 'Enter') attemptLogin(); });



function attemptLogin(){

  const pass = document.getElementById('passwordInput').value.trim();

  const errEl = document.getElementById('loginError');

  if(!selectedProfile){ errEl.textContent = 'Sélectionnez un profil.'; return; }

  if(!pass){ errEl.textContent = 'Entrez le code d\'accès.'; return; }



  if(selectedProfile.authType === 'nv-pin'){

    errEl.textContent = 'Vérification...';

    const fullName = selectedProfile.fullName;

    db.collection('agents').where('name', '==', fullName).limit(1).get().then(snap => {

      if(!snap.empty) return snap;

      // Nom exact introuvable : on retente en tolérant un petit écart (typo/espace) sur les fiches Agents,

      // pour ne pas bloquer le login à cause d'une orthographe légèrement différente.

      return db.collection('agents').get().then(allSnap => {

        let bestDoc = null, bestDist = 3; // tolère jusqu'à 2 caractères d'écart

        allSnap.docs.forEach(d => {

          const n = (d.data().name || '').trim();

          if(!n) return;

          const dist = levenshtein(n.toLowerCase(), fullName.toLowerCase());

          if(dist < bestDist){ bestDist = dist; bestDoc = d; }

        });

        return bestDoc ? { empty: false, docs: [bestDoc] } : { empty: true, docs: [] };

      });

    }).then(snap => {

      if(snap.empty){

        errEl.textContent = "Cet agent n'a pas encore de fiche dans Agents — crée-la d'abord (avec le PIN).";

        return;

      }

      const agentDoc = snap.docs[0];

      const agentPin = agentDoc.data().pin;

      if(!agentPin || agentPin !== pass){

        errEl.textContent = 'Code PIN incorrect.';

        return;

      }

      errEl.textContent = '';

      currentNvAgent = fullName;

      // On repart d'une session propre (anonyme) : si une session admin traînait dans ce navigateur,

      // elle ne doit jamais rester active en même temps que le portail agent.

      auth.signOut()

        .catch(() => {})

        .then(() => auth.signInAnonymously())

        .finally(() => {

          document.getElementById('app').style.display = 'none';

          document.getElementById('loginScreen').classList.add('hidden');

          document.getElementById('agentPortal').classList.add('hidden');

          document.getElementById('nvAgentPortal').classList.remove('hidden');

          listenHotspotsNv();

          renderNvPortal();

          listenNvReplyNotifications();

        });

    }).catch(err => { console.error(err); errEl.textContent = 'Erreur de connexion.'; });

    return;

  }



  if(selectedProfile.authType === 'pin'){

    errEl.textContent = 'Vérification...';

    db.collection('agents').doc(selectedProfile.id).get().then(doc => {

      if(!doc.exists){ errEl.textContent = 'Agent introuvable.'; return; }

      const agent = { id: doc.id, ...doc.data() };



      if(!agent.pin || agent.pin !== pass){

        errEl.textContent = 'Code PIN incorrect.';

        return;

      }

      errEl.textContent = '';

      enterAgentPortal(agent);

    }).catch(err => { console.error(err); errEl.textContent = 'Erreur de connexion.'; });

    return;

  }



  errEl.textContent = 'Connexion...';

  // Admin/manager profiles use Firebase Auth (email/password).

  const email = `${selectedProfile.id}@opshq.ma`;

  auth.signInWithEmailAndPassword(email, pass)

    .then(() => {

      errEl.textContent = '';

      enterApp(selectedProfile);

    })

    .catch(err => {

      errEl.textContent = "Mot de passe incorrect ou compte non configuré.";

      console.error(err);

    });

}



function enterApp(profile){

  document.getElementById('loginScreen').style.display = 'none';

  document.getElementById('app').style.display = 'block';

  document.getElementById('sidebarUserName').textContent = profile.name;

  document.getElementById('sidebarUserRole').textContent = profile.role;

  populateZoneSelect();

  listenAgents();

  listenZoneLabels();

  listenNewRapportsNotifications();

}



/* ================= NOTIFICATIONS: NOUVEAUX RAPPORTS/REMARQUES ================= */

let newRapportsUnsub = null;



function getRapportsLastSeenAt(){

  return parseInt(localStorage.getItem('opsHqRapportsLastSeen') || '0', 10);

}



function markRapportsSeenNow(){

  localStorage.setItem('opsHqRapportsLastSeen', String(Date.now()));

  const badge = document.getElementById('rapportsNavBadge');

  if(badge) badge.classList.add('hidden');

  const banner = document.getElementById('dashboardNewRapportsBanner');

  if(banner) banner.classList.add('hidden');

}



function listenNewRapportsNotifications(){

  if(newRapportsUnsub) newRapportsUnsub();

  // Écoute en continu (pas juste au chargement) — la pastille et la bannière se mettent à jour

  // dès qu'un nouveau rapport/remarque arrive, même si l'admin est déjà en train d'utiliser l'app.

  newRapportsUnsub = db.collection('rapports').orderBy('createdAt', 'desc').limit(50).onSnapshot(snap => {

    const lastSeen = getRapportsLastSeenAt();

    const fresh = snap.docs

      .map(d => d.data())

      .filter(r => r.createdAt && r.createdAt.toMillis && r.createdAt.toMillis() > lastSeen);



    const badge = document.getElementById('rapportsNavBadge');

    const onRapportsPage = !document.getElementById('view-rapports').classList.contains('hidden');

    if(fresh.length === 0 || onRapportsPage){

      if(badge) badge.classList.add('hidden');

    } else if(badge){

      badge.textContent = fresh.length > 9 ? '9+' : String(fresh.length);

      badge.classList.remove('hidden');

    }



    const banner = document.getElementById('dashboardNewRapportsBanner');

    if(!banner) return;

    if(fresh.length === 0 || onRapportsPage){

      banner.classList.add('hidden');

      return;

    }

    const nUrgent = fresh.filter(r => r.urgent).length;

    banner.innerHTML = `

      <div style="display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap; background:${nUrgent ? 'rgba(220,38,38,0.1)' : 'rgba(0,160,130,0.1)'}; border:1px solid ${nUrgent ? 'rgba(220,38,38,0.3)' : 'rgba(0,160,130,0.3)'}; border-radius:14px; padding:14px 18px; margin-bottom:16px;">

        <div style="font-size:13.5px; font-weight:700;">

          ${nUrgent ? '🚨' : '🔔'} ${fresh.length} nouveau(x) rapport(s)/remarque(s)${nUrgent ? ` — dont ${nUrgent} urgent(s) !` : ''}

        </div>

        <button class="btn-ghost" id="dashboardGoToRapportsBtn" style="border:1px solid var(--line); font-size:12px;">Voir les rapports →</button>

      </div>

    `;

    banner.classList.remove('hidden');

    document.getElementById('dashboardGoToRapportsBtn').addEventListener('click', () => {

      document.querySelector('[data-view="rapports"]').click();

    });

  }, err => console.error(err));

}



document.getElementById('logoutBtn').addEventListener('click', () => {

  auth.signOut().finally(() => location.reload());

});



/* ================= CHANGE PASSWORD (self-service) ================= */

document.getElementById('changePassBtn').addEventListener('click', () => {

  document.getElementById('oldPassInput').value = '';

  document.getElementById('newPassInput').value = '';

  document.getElementById('confirmPassInput').value = '';

  document.getElementById('changePassError').textContent = '';

  document.getElementById('changePassOverlay').classList.remove('hidden');

});

document.getElementById('cancelChangePassBtn').addEventListener('click', () => {

  document.getElementById('changePassOverlay').classList.add('hidden');

});

document.getElementById('changePassOverlay').addEventListener('click', e => {

  if(e.target.id === 'changePassOverlay') document.getElementById('changePassOverlay').classList.add('hidden');

});



document.getElementById('saveChangePassBtn').addEventListener('click', () => {

  const oldPass = document.getElementById('oldPassInput').value;

  const newPass = document.getElementById('newPassInput').value;

  const confirmPass = document.getElementById('confirmPassInput').value;

  const errEl = document.getElementById('changePassError');

  const user = auth.currentUser;



  if(!user || !user.email){ errEl.textContent = "Session invalide, reconnecte-toi."; return; }

  if(!oldPass || !newPass || !confirmPass){ errEl.textContent = 'Remplis tous les champs.'; return; }

  if(newPass.length < 6){ errEl.textContent = 'Le nouveau mot de passe doit faire au moins 6 caractères.'; return; }

  if(newPass !== confirmPass){ errEl.textContent = 'Les deux nouveaux mots de passe ne correspondent pas.'; return; }



  errEl.textContent = 'Vérification...';

  const credential = firebase.auth.EmailAuthProvider.credential(user.email, oldPass);

  user.reauthenticateWithCredential(credential)

    .then(() => user.updatePassword(newPass))

    .then(() => {

      errEl.textContent = '';

      document.getElementById('changePassOverlay').classList.add('hidden');

      alert('Mot de passe changé avec succès.');

    })

    .catch(err => {

      console.error(err);

      if(err.code === 'auth/wrong-password'){ errEl.textContent = 'Mot de passe actuel incorrect.'; }

      else { errEl.textContent = 'Erreur lors du changement de mot de passe.'; }

    });

});



/* ================= AGENT PORTAL (Pointage) ================= */

let currentAgentSession = null;



function enterAgentPortal(agent){

  currentAgentSession = agent;

  document.getElementById('loginScreen').style.display = 'none';

  document.getElementById('agentPortal').classList.remove('hidden');

  document.getElementById('portalAgentName').textContent = agent.nickname || agent.name;



  const firstName = (agent.nickname || agent.name).split(' ')[0];

  const MOTIVATIONAL_QUOTES = [

    "Chaque course compte. Bonne journée sur le terrain !",

    "L'équipe compte sur toi aujourd'hui — donne le meilleur !",

    "Un jour à la fois, une livraison à la fois. On avance !",

    "Ton travail fait la différence pour toute l'équipe.",

    "Reste concentré, reste safe — et fais du bon boulot !",

    "Aujourd'hui encore, on assure la couverture. Bravo d'avance !",

  ];

  const quoteIdx = (new Date().getDate() + firstName.length) % MOTIVATIONAL_QUOTES.length;

  document.getElementById('portalMotivation').innerHTML = `

    <div style="background:linear-gradient(135deg,var(--sky),var(--sky-deep)); border-radius:16px; padding:22px 20px; margin-bottom:16px; text-align:center; color:#fff; box-shadow:0 16px 34px -18px rgba(0,160,130,0.55);">

      <div style="width:52px; height:52px; border-radius:50%; background:rgba(255,255,255,0.18); display:flex; align-items:center; justify-content:center; font-size:26px; margin:0 auto 10px;">🛵</div>

      <div style="font-size:17px; font-weight:800; margin-bottom:4px;">Bonjour, ${firstName} 💪</div>

      <div style="font-size:12.5px; opacity:.9; max-width:280px; margin:0 auto; line-height:1.5;">${MOTIVATIONAL_QUOTES[quoteIdx]}</div>

    </div>

  `;



  refreshPortalStatus();

}



document.getElementById('portalLogoutBtn').addEventListener('click', () => {

  auth.signOut().finally(() => location.reload());

});



function todayDateKey(){

  const d = new Date();

  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

}



function refreshPortalStatus(){

  if(!currentAgentSession) return;

  const docId = `${todayDateKey()}_${currentAgentSession.id}`;

  db.collection('pointage').doc(docId).get().then(doc => {

    const btn = document.getElementById('portalCheckBtn');

    const status = document.getElementById('portalStatus');

    btn.disabled = false;

    const data = doc.exists ? doc.data() : null;



    if(!data || !data.checkIn){

      status.textContent = "Pas encore pointé aujourd'hui";

      btn.textContent = 'Check-in';

      btn.classList.remove('is-checkedin');

      btn.dataset.mode = 'in';

    } else if(data.checkIn && !data.checkOut){

      const t = data.checkIn.toDate ? data.checkIn.toDate() : new Date(data.checkIn);

      status.textContent = `En service depuis ${t.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' })}`;

      btn.textContent = 'Check-out';

      btn.classList.add('is-checkedin');

      btn.dataset.mode = 'out';

    } else {

      const tIn = data.checkIn.toDate ? data.checkIn.toDate() : new Date(data.checkIn);

      const tOut = data.checkOut.toDate ? data.checkOut.toDate() : new Date(data.checkOut);

      status.textContent = `Terminé — ${tIn.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' })} → ${tOut.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' })}`;

      btn.textContent = 'Journée terminée';

      btn.classList.remove('is-checkedin');

      btn.disabled = true;

      btn.dataset.mode = 'done';

    }

    renderPortalHistory();

  }).catch(err => console.error(err));

}



document.getElementById('portalCheckBtn').addEventListener('click', () => {

  if(!currentAgentSession) return;

  const btn = document.getElementById('portalCheckBtn');

  const mode = btn.dataset.mode;

  if(mode === 'done') return;

  const docId = `${todayDateKey()}_${currentAgentSession.id}`;

  const field = mode === 'in' ? 'checkIn' : 'checkOut';

  const payload = {

    agentId: currentAgentSession.id,

    agentName: currentAgentSession.name,

    date: todayDateKey(),

    [field]: firebase.firestore.FieldValue.serverTimestamp(),

  };

  db.collection('pointage').doc(docId).set(payload, { merge: true })

    .then(() => refreshPortalStatus())

    .catch(err => console.error(err));

});



function renderPortalHistory(){

  if(!currentAgentSession) return;

  // Un seul filtre d'égalité (agentId) — pas besoin d'index composite ; tri fait côté client.

  db.collection('pointage').where('agentId', '==', currentAgentSession.id).get()

    .then(snap => {

      const docs = snap.docs.map(d => d.data()).sort((a, b) => b.date.localeCompare(a.date)).slice(0, 7);

      const list = document.getElementById('portalHistoryList');

      list.innerHTML = '';

      if(docs.length === 0){

        list.innerHTML = `<div style="text-align:center; padding:20px 10px; color:var(--ink-soft); font-size:12.5px;">Aucun historique pour l'instant — ton premier check-in apparaîtra ici.</div>`;

        return;

      }

      docs.forEach(d => {

        const inT = d.checkIn && d.checkIn.toDate ? d.checkIn.toDate().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) : '—';

        const outT = d.checkOut && d.checkOut.toDate ? d.checkOut.toDate().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) : '—';

        const row = document.createElement('div');

        row.className = 'portal-history-item';

        row.innerHTML = `<span>${d.date}</span><span>${inT} → ${outT}</span>`;

        list.appendChild(row);

      });

    }).catch(err => console.error(err));

}



setInterval(() => {

  if(currentAgentSession) document.getElementById('portalClock').textContent = new Date().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });

}, 1000);



/* ================= NAVIGATION ================= */

document.querySelectorAll('.nav-item[data-view]').forEach(btn => {

  btn.addEventListener('click', () => {

    document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));

    btn.classList.add('active');

    const view = btn.dataset.view;

    document.querySelectorAll('.view').forEach(v => v.classList.add('hidden'));

    const target = document.getElementById('view-' + view);

    target.classList.remove('hidden');

    document.getElementById('viewTitle').textContent = btn.querySelector('span').textContent;

    if(['paie'].includes(view) && !target.dataset.built){

      renderPlaceholder(target, btn.querySelector('span').textContent);

      target.dataset.built = '1';

    }

    if(view === 'planning'){

      renderPlanningWeek();

    }

    if(view === 'pointage'){

      renderPointageDay();

    }

    if(view === 'absences'){

      renderAbsencesView();

    }

    if(view === 'planningnv'){

      listenHotspotsNv();

    }

    if(view === 'rapports'){

      renderRapportsView();

    }

  });

});

document.querySelector('.nav-item[data-view="dashboard"]').classList.add('active');



function renderPlaceholder(el, label){

  el.innerHTML = `

    <div class="placeholder-view">

      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/></svg>

      <h3>${label} — module en construction</h3>

      <p>C'est la fondation. On construit ce module dans la prochaine étape.</p>

    </div>`;

}



/* ================= CLOCK ================= */

function tickHeroClock(){

  const el = document.getElementById('heroClock');

  if(el) el.textContent = new Date().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });

}

setInterval(tickHeroClock, 1000);

tickHeroClock();



const HERO_MOTIVATIONAL_QUOTES = [

  "Chaque course compte. Bonne journée sur le terrain !",

  "L'équipe compte sur vous aujourd'hui — donnez le meilleur !",

  "Un jour à la fois, une livraison à la fois. On avance !",

  "Votre travail fait la différence pour toute l'équipe.",

  "Restez concentrés, restez safe — et faites du bon boulot !",

  "Aujourd'hui encore, on assure la couverture. Bravo d'avance !",

];

function renderHeroDateAndQuote(){

  const dateEl = document.getElementById('heroDate');

  const quoteEl = document.getElementById('heroMotivation');

  const now = new Date();

  if(dateEl) dateEl.textContent = now.toLocaleDateString('fr-FR', { weekday: 'long', day: 'numeric', month: 'long' });

  if(quoteEl){

    const quoteIdx = now.getDate() % HERO_MOTIVATIONAL_QUOTES.length;

    quoteEl.textContent = `"${HERO_MOTIVATIONAL_QUOTES[quoteIdx]}"`;

  }

}

renderHeroDateAndQuote();



function tickClock(){

  const now = new Date();

  document.getElementById('clock').textContent = now.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });

}

setInterval(tickClock, 1000);

setInterval(() => { if(dashboardPlanningData !== undefined) renderCurrentlyWorking(); }, 60000);

tickClock();



/* ================= DATE / WEEK UTILITIES ================= */

function getWeekStart(date){

  const d = new Date(date);

  const day = d.getDay(); // 0=Sun..6=Sat

  const diff = (day === 0 ? -6 : 1 - day);

  d.setDate(d.getDate() + diff);

  d.setHours(0,0,0,0);

  return d;

}

function getISOWeekIndex(date){

  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));

  const dayNum = d.getUTCDay() || 7;

  d.setUTCDate(d.getUTCDate() + 4 - dayNum);

  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));

  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);

}

function weekKeyFor(weekStart){

  const d = new Date(Date.UTC(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate()));

  const dayNum = d.getUTCDay() || 7;

  d.setUTCDate(d.getUTCDate() + 4 - dayNum);

  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));

  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);

  return `${d.getUTCFullYear()}-W${String(weekNo).padStart(2,'0')}`;

}

function formatWeekRange(weekStart){

  const end = new Date(weekStart); end.setDate(end.getDate() + 6);

  const startStr = weekStart.toLocaleDateString('fr-FR', { day:'numeric', month:'long' });

  const endStr = end.toLocaleDateString('fr-FR', { day:'numeric', month:'long', year:'numeric' });

  return `${startStr} → ${endStr}`;

}



/* ================= GROUP-BASED ZONE / PERIOD RESOLUTION ================= */

function groupMembers(groupId){

  return AGENTS.filter(a => a.group === groupId).sort((x,y) => x.id.localeCompare(y.id));

}

function resolveAgentZone(agent, weekIdx){

  if(!agent.group) return agent.zone || '—';

  const g = GROUPS[agent.group];

  if(!g) return agent.zone || '—';

  if(g.mode === 'fixed_solo') return agent.fixedZone || '—';

  if(g.mode === 'fixed_team') return g.zone;

  if(g.mode === 'rotating_team'){

    const idx = (((g.offset + weekIdx) % 3) + 3) % 3;

    return ROTATING_ZONES[idx];

  }

  return '—';

}

function resolveAgentPeriod(agent, weekIdx){

  // Jour/Nuit est FIXE par agent (réglé dans la page Agents), pas de rotation automatique.

  if(agent.defaultPeriod === 'JOUR' || agent.defaultPeriod === 'NUIT') return agent.defaultPeriod;

  if(!agent.group){

    if(agent.shiftType === 'FIX') return agent.fixedPeriod || 'JOUR';

    return 'JOUR';

  }

  // Agent pas encore configuré : valeur de départ raisonnable selon sa position dans le groupe.

  const members = groupMembers(agent.group);

  const idx = members.findIndex(m => m.id === agent.id);

  return idx === 0 ? 'JOUR' : 'NUIT';

}

// Chaque groupe de 3 personnes doit "sauter" un jour différent des autres,

// pour que la répartition des OFF reste équilibrée (4/3/3/3 sur 13 agents).

const GROUP_SKIP_OFFSET = { trio1: 0, trio2: 1, trio3: 2, sale: 3 };



function offDayIndex(agent, weekIdx){

  // OFF est FIXE par agent (réglé dans la page Agents), pas de rotation automatique.

  if(typeof agent.defaultOffDay === 'number' && agent.defaultOffDay >= 0 && agent.defaultOffDay <= 3){

    return agent.defaultOffDay;

  }

  const members = agent.group ? groupMembers(agent.group) : AGENTS.filter(a => !a.group).sort((x,y) => x.id.localeCompare(y.id));

  let idx = members.findIndex(m => m.id === agent.id);

  if(idx < 0) idx = 0;

  const allDays = [0, 1, 2, 3]; // lun..jeu

  if(members.length === 1) return 0;

  const groupOffset = GROUP_SKIP_OFFSET[agent.group] || 0;

  const skipped = groupOffset % 4;

  const availableDays = allDays.filter(d => d !== skipped);

  return availableDays[idx % availableDays.length];

}



/* ================= FIRESTORE: AGENTS LISTENER ================= */

function listenAgents(){

  if(agentsUnsub) agentsUnsub();

  agentsUnsub = db.collection('agents').orderBy('name').onSnapshot(snap => {

    AGENTS = snap.docs.map(doc => ({ id: doc.id, ...doc.data() }));

    initDashboard();

    renderAgentsTable();

    if(!document.getElementById('view-planning').classList.contains('hidden')) renderPlanningWeek();

    seedMissingAgentEmails();

  }, err => {

    console.error('Erreur de lecture des agents:', err);

  });

}



let emailSeedDone = false;

function seedMissingAgentEmails(){

  if(emailSeedDone) return; // une seule passe par session, pour ne pas boucler sur les onSnapshot suivants

  emailSeedDone = true;

  AGENTS.forEach(a => {

    if(!a.email && DEFAULT_NV_EMAILS[a.name]){

      db.collection('agents').doc(a.id).update({ email: DEFAULT_NV_EMAILS[a.name] }).catch(err => console.error(err));

    }

  });

}



/* ================= DASHBOARD ================= */

function initDashboard(){

  const weekIdx = getISOWeekIndex(new Date());

  let jourCount = 0, nuitCount = 0;

  AGENTS.forEach(a => {

    const period = resolveAgentPeriod(a, weekIdx);

    if(period === 'JOUR') jourCount++; else nuitCount++;

  });

  document.getElementById('statActive').textContent = AGENTS.length;

  document.getElementById('statJour').textContent = jourCount;

  document.getElementById('statNuit').textContent = nuitCount;

  const zonesCovered = new Set(AGENTS.map(a => resolveAgentZone(a, weekIdx))).size;

  document.getElementById('statZones').textContent = `${zonesCovered}`;



  const zoneGrid = document.getElementById('zoneGrid');

  zoneGrid.innerHTML = '';

  ZONES.forEach(zone => {

    const count = AGENTS.filter(a => resolveAgentZone(a, weekIdx) === zone).length;

    const color = ZONE_COLORS[zone] || 'var(--mustard)';

    const card = document.createElement('div');

    card.className = 'zone-card';

    card.style.borderLeftColor = color;

    card.innerHTML = `<h3>${getZoneLabel(zone)}</h3><div class="count">${count}</div><div class="count-label">agent${count>1?'s':''} assigné${count>1?'s':''}</div>`;

    zoneGrid.appendChild(card);

  });



  loadDashboardPlanningStats();

}



function loadDashboardPlanningStats(){

  const todayStart = getWeekStart(new Date());

  const key = weekKeyFor(todayStart);

  db.collection('planning').doc(key).get().then(doc => {

    const data = doc.exists ? doc.data() : null;

    dashboardPlanningData = data;

    const statOff = document.getElementById('statOffToday');

    const statAbs = document.getElementById('statAbsences');

    const statAlerts = document.getElementById('statAlerts');



    if(!data || !data.agents){

      statOff.textContent = '—';

      statAbs.textContent = '—';

      statAlerts.textContent = '—';

      renderCurrentlyWorking();

      return;

    }



    const todayIdx = (new Date().getDay() + 6) % 7; // lundi=0 .. dimanche=6

    const todayKey = DAY_KEYS[todayIdx];



    let offToday = 0, absences = 0;

    Object.values(data.agents).forEach(a => {

      if(!a.cells) return;

      if(a.cells[todayKey] === 'OFF') offToday++;

      Object.values(a.cells).forEach(code => { if(code === 'ABS' || code === 'CONGE') absences++; });

    });



    statOff.textContent = offToday;

    statAbs.textContent = absences;

    statAlerts.textContent = computeWarnings(data).length;

    renderCurrentlyWorking();

  }).catch(err => {

    console.error('Erreur stats dashboard:', err);

  });

}



/* ================= LIVE: QUI TRAVAILLE MAINTENANT ================= */

let dashboardPlanningData = null;



function parseShiftLabel(label){

  const [startStr, endStr] = label.split('-');

  const [sh, sm] = startStr.split(':').map(Number);

  const [eh, em] = endStr.split(':').map(Number);

  const startM = sh * 60 + sm;

  const endM = eh * 60 + em;

  const crossesMidnight = endM <= startM;

  return { startM, endM, crossesMidnight, startStr, endStr };

}



function renderCurrentlyWorking(){

  const tbody = document.getElementById('liveTableBody');

  const table = document.querySelector('#liveCard table');

  const emptyState = document.getElementById('liveEmptyState');

  tbody.innerHTML = '';



  const rows = [];

  if(dashboardPlanningData && dashboardPlanningData.agents){

    const now = new Date();

    const nowMinutes = now.getHours() * 60 + now.getMinutes();

    const todayIdx = (now.getDay() + 6) % 7;

    const yestIdx = (todayIdx + 6) % 7;

    const todayKey = DAY_KEYS[todayIdx];

    const yestKey = DAY_KEYS[yestIdx];



    AGENTS.filter(a => a.group).forEach(a => {

      const data = dashboardPlanningData.agents[a.id];

      if(!data || !data.cells) return;



      const todayCode = data.cells[todayKey];

      if(todayCode && SHIFT_CATALOG[todayCode] && todayCode !== 'OFF' && todayCode !== 'ABS' && todayCode !== 'CONGE'){

        const t = parseShiftLabel(SHIFT_CATALOG[todayCode].label);

        const isActive = t.crossesMidnight ? (nowMinutes >= t.startM) : (nowMinutes >= t.startM && nowMinutes < t.endM);

        if(isActive) rows.push({ name: a.name, zone: data.zone, label: SHIFT_CATALOG[todayCode].label, end: t.endStr });

      }



      const yestCode = data.cells[yestKey];

      if(yestCode && SHIFT_CATALOG[yestCode] && yestCode !== 'OFF' && yestCode !== 'ABS' && yestCode !== 'CONGE'){

        const t = parseShiftLabel(SHIFT_CATALOG[yestCode].label);

        if(t.crossesMidnight && nowMinutes < t.endM){

          rows.push({ name: a.name, zone: data.zone, label: SHIFT_CATALOG[yestCode].label, end: t.endStr });

        }

      }

    });

  }



  if(rows.length === 0){

    table.classList.add('hidden');

    emptyState.classList.remove('hidden');

    return;

  }

  table.classList.remove('hidden');

  emptyState.classList.add('hidden');



  rows.forEach(r => {

    const tr = document.createElement('tr');

    tr.innerHTML = `<td><b>${r.name}</b></td><td>${r.zone}</td><td>${r.label}</td><td class="total-cell">${r.end}</td>`;

    tbody.appendChild(tr);

  });

}



function levenshtein(a, b){

  const m = a.length, n = b.length;

  const dp = Array.from({ length: m + 1 }, (_, i) => [i, ...Array(n).fill(0)]);

  for(let j = 0; j <= n; j++) dp[0][j] = j;

  for(let i = 1; i <= m; i++){

    for(let j = 1; j <= n; j++){

      dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);

    }

  }

  return dp[m][n];

}



function findNvNameTypo(name){

  if(NV_ALL_AGENTS.includes(name)) return null;

  let best = null, bestDist = 4; // tolère jusqu'à 3 caractères d'écart

  NV_ALL_AGENTS.forEach(nvName => {

    const d = levenshtein(name.toLowerCase(), nvName.toLowerCase());

    if(d < bestDist){ bestDist = d; best = nvName; }

  });

  return best;

}



function renderAgentsTable(){

  const weekIdx = getISOWeekIndex(new Date());

  const tbody = document.getElementById('agentsTableBody');

  const emptyState = document.getElementById('agentsEmptyState');

  const table = document.querySelector('#agentsCard table');

  tbody.innerHTML = '';



  if(AGENTS.length === 0){

    table.classList.add('hidden');

    emptyState.classList.remove('hidden');

    return;

  }

  table.classList.remove('hidden');

  emptyState.classList.add('hidden');



  AGENTS.forEach(a => {

    const period = resolveAgentPeriod(a, weekIdx);

    const zone = resolveAgentZone(a, weekIdx);

    const groupLabel = a.group ? GROUPS[a.group].label : 'Manuel';

    const isNvAgent = NV_ALL_AGENTS.includes(a.name);

    const typoMatch = !isNvAgent ? findNvNameTypo(a.name) : null;

    const tr = document.createElement('tr');

    tr.innerHTML = `

      <td><b>${a.name}</b>${typoMatch ? `<div style="font-size:10.5px; color:var(--red); font-weight:600; margin-top:2px;">⚠️ Nom différent de "${typoMatch}" (Planning NV) — PIN et email ne marcheront pas tant que ce n'est pas corrigé</div>` : ''}</td>

      <td>${a.pin ? `<span class="agent-pin-chip">${a.pin}</span>` : '<span style="color:var(--ink-soft); font-weight:400;">—</span>'}</td>

      <td style="font-size:12px; min-width:170px;">

        ${a.email ? `<div style="color:var(--ink); font-weight:600; margin-bottom:6px;">✉️ ${a.email}</div>` : '<div style="color:var(--ink-soft); margin-bottom:6px;">Pas d\'email</div>'}

        ${a.email ? `<div style="display:flex; gap:6px; flex-wrap:wrap;">

          <button class="agent-access-btn" data-send-access="${a.id}">🔑 Accès</button>

        </div>` : ''}

      </td>

      <td><span class="zone-dot" style="background:${ZONE_COLORS[zone] || 'var(--ink-soft)'}"></span>${getZoneLabel(zone)}</td>

      <td><span class="badge ${a.group ? 'fix' : 'var'}">${groupLabel}</span></td>

      <td><span class="badge ${period === 'JOUR' ? 'jour' : 'nuit'}">${period}</span></td>

      <td>

        <div class="row-actions">

          <button class="icon-btn" data-edit="${a.id}" title="Modifier"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 013 3L7 19l-4 1 1-4z"/></svg></button>

          <button class="icon-btn danger" data-delete="${a.id}" title="Supprimer"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2m3 0v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6h14z"/></svg></button>

        </div>

      </td>

    `;

    tbody.appendChild(tr);

  });



  tbody.querySelectorAll('[data-edit]').forEach(btn => btn.addEventListener('click', () => openAgentModal(btn.dataset.edit)));

  tbody.querySelectorAll('[data-delete]').forEach(btn => btn.addEventListener('click', () => deleteAgent(btn.dataset.delete)));

  tbody.querySelectorAll('[data-send-access]').forEach(btn => btn.addEventListener('click', () => {

    const a = AGENTS.find(x => x.id === btn.dataset.sendAccess);

    if(a) sendAgentAccessEmail(a);

  }));

}



/* ================= AGENT MODAL (Add / Edit) ================= */

function populateZoneSelect(){

  document.getElementById('agentZoneInput').innerHTML = ZONES.map(z => `<option value="${z}">${getZoneLabel(z)}</option>`).join('');

  document.getElementById('agentFixedZoneInput').innerHTML = ZONES.map(z => `<option value="${z}">${getZoneLabel(z)}</option>`).join('');

}



let editingAgentId = null;



function openAgentModal(agentId = null){

  editingAgentId = agentId;

  const overlay = document.getElementById('agentModalOverlay');

  const title = document.getElementById('agentModalTitle');

  const nameInput = document.getElementById('agentNameInput');

  const nicknameInput = document.getElementById('agentNicknameInput');

  const pinInput = document.getElementById('agentPinModalInput');

  const emailInput = document.getElementById('agentEmailModalInput');

  const groupInput = document.getElementById('agentGroupInput');

  const zoneInput = document.getElementById('agentZoneInput');

  const fixedZoneInput = document.getElementById('agentFixedZoneInput');

  const defaultPeriodInput = document.getElementById('agentDefaultPeriodInput');

  const defaultOffDayInput = document.getElementById('agentDefaultOffDayInput');

  document.getElementById('agentModalError').textContent = '';



  if(agentId){

    const agent = AGENTS.find(a => a.id === agentId);

    title.textContent = 'Modifier l\'agent';

    nameInput.value = agent.name;

    nicknameInput.value = agent.nickname || '';

    pinInput.value = agent.pin || '';

    emailInput.value = agent.email || DEFAULT_NV_EMAILS[agent.name] || '';

    groupInput.value = agent.group || '';

    if(agent.zone) zoneInput.value = agent.zone;

    if(agent.fixedZone) fixedZoneInput.value = agent.fixedZone;

    defaultPeriodInput.value = agent.defaultPeriod || resolveAgentPeriod(agent, getISOWeekIndex(new Date()));

    defaultOffDayInput.value = typeof agent.defaultOffDay === 'number' ? String(agent.defaultOffDay) : String(offDayIndex(agent, getISOWeekIndex(new Date())));

  } else {

    title.textContent = 'Ajouter un agent';

    nameInput.value = '';

    nicknameInput.value = '';

    pinInput.value = generateRandomPin();

    emailInput.value = '';

    groupInput.value = '';

    zoneInput.selectedIndex = 0;

    fixedZoneInput.selectedIndex = 0;

    defaultPeriodInput.value = 'JOUR';

    defaultOffDayInput.value = '0';

  }

  toggleGroupFieldsVisibility();

  overlay.classList.remove('hidden');

  nameInput.focus();

}



function closeAgentModal(){

  document.getElementById('agentModalOverlay').classList.add('hidden');

  editingAgentId = null;

}



function toggleGroupFieldsVisibility(){

  const group = document.getElementById('agentGroupInput').value;

  document.getElementById('manualZoneGroup').classList.toggle('hidden', group !== '');

  document.getElementById('fixedSoloZoneGroup').classList.toggle('hidden', group !== 'zakariae');

  document.getElementById('rotationHint').classList.toggle('hidden', !(group === 'trio1' || group === 'trio2' || group === 'trio3'));

  document.getElementById('saleHint').classList.toggle('hidden', group !== 'sale');

}



document.getElementById('openAddAgentBtn').addEventListener('click', () => openAgentModal(null));



document.getElementById('syncNvAgentsBtn').addEventListener('click', () => {

  const roster = NV_ALL_AGENTS; // les 12 agents actuels de Planning NV — source de référence

  const existing = AGENTS.map(a => ({ id: a.id, name: (a.name || '').trim() }));

  const existingExactNames = new Set(existing.map(a => a.name.toLowerCase()));



  const missing = [];       // à créer directement (aucune fiche proche trouvée)

  const nearMatches = [];   // fiche proche existante (probable typo) — à corriger à la main, pas de création auto



  roster.forEach(rosterName => {

    if(existingExactNames.has(rosterName.toLowerCase())) return; // déjà OK

    let closest = null, closestDist = 4; // tolère jusqu'à 3 caractères d'écart pour repérer un typo probable

    existing.forEach(a => {

      const dist = levenshtein(a.name.toLowerCase(), rosterName.toLowerCase());

      if(dist < closestDist){ closestDist = dist; closest = a; }

    });

    if(closest) nearMatches.push({ rosterName, existingName: closest.name });

    else missing.push(rosterName);

  });



  if(missing.length === 0 && nearMatches.length === 0){

    alert('Les 12 agents de Planning NV ont déjà tous une fiche correspondante dans Agents. Rien à faire.');

    return;

  }



  let msg = '';

  if(missing.length) msg += `${missing.length} agent(s) manquant(s) vont être créés dans Agents avec un PIN généré :\n${missing.join('\n')}\n\n`;

  if(nearMatches.length) msg += `⚠️ ${nearMatches.length} fiche(s) existante(s) ressemble(nt) à un agent de Planning NV mais le nom ne correspond pas exactement — à corriger toi-même dans Agents (pas de création automatique pour éviter un doublon) :\n${nearMatches.map(m => `"${m.existingName}" → devrait être "${m.rosterName}"`).join('\n')}\n\n`;

  msg += 'Continuer ?';

  if(!confirm(msg)) return;



  if(missing.length === 0){

    alert('Rien créé automatiquement (voir les correspondances approximatives à corriger toi-même).');

    return;

  }



  const btn = document.getElementById('syncNvAgentsBtn');

  btn.disabled = true;

  const batch = db.batch();

  const pinReport = [];

  missing.forEach(name => {

    const ref = db.collection('agents').doc();

    const pin = generateRandomPin();

    pinReport.push(`${name} : ${pin}`);

    batch.set(ref, {

      name,

      nickname: name.split(' ')[0],

      group: null,

      zone: null,

      fixedZone: null,

      pin,

      defaultPeriod: 'JOUR',

      defaultOffDay: 0,

    });

  });

  batch.commit()

    .then(() => {

      alert(`${missing.length} agent(s) créé(s) avec un PIN généré automatiquement :\n\n${pinReport.join('\n')}\n\nNote ces codes pour les donner à chaque agent.`);

      btn.disabled = false;

    })

    .catch(err => { console.error(err); btn.disabled = false; alert('Erreur pendant la synchronisation.'); });

});



/* ================= PIN GENERATION ================= */

function generateRandomPin(){

  return String(Math.floor(1000 + Math.random() * 9000)); // 4 chiffres

}



document.getElementById('genPinBtn').addEventListener('click', () => {

  document.getElementById('agentPinModalInput').value = generateRandomPin();

});



document.getElementById('genAllPinsBtn').addEventListener('click', () => {

  const missing = AGENTS.filter(a => !a.pin);

  if(missing.length === 0){ alert('Tous les agents ont déjà un PIN.'); return; }

  const btn = document.getElementById('genAllPinsBtn');

  btn.disabled = true;

  btn.textContent = 'Génération...';

  const batch = db.batch();

  const pinReport = [];

  missing.forEach(a => {

    const pin = generateRandomPin();

    pinReport.push(`${a.name} : ${pin}`);

    batch.update(db.collection('agents').doc(a.id), { pin });

  });

  batch.commit()

    .then(() => {

      btn.disabled = false;

      btn.textContent = '🎲 Générer les PIN manquants';

      alert(`PIN générés pour ${missing.length} agent(s) :\n\n${pinReport.join('\n')}\n\nNote ces codes pour les donner à chaque agent.`);

    })

    .catch(err => { console.error(err); btn.disabled = false; btn.textContent = '🎲 Générer les PIN manquants'; alert('Erreur pendant la génération.'); });

});



document.getElementById('genAllNicknamesBtn').addEventListener('click', () => {

  const missing = AGENTS.filter(a => !a.nickname);

  if(missing.length === 0){ alert('Tous les agents ont déjà un surnom.'); return; }

  const batch = db.batch();

  missing.forEach(a => {

    batch.update(db.collection('agents').doc(a.id), { nickname: a.name.split(' ').slice(-1)[0] });

  });

  batch.commit()

    .then(() => alert(`Surnom généré pour ${missing.length} agent(s) (premier prénom par défaut). Tu peux les modifier un par un si besoin.`))

    .catch(err => { console.error(err); alert('Erreur pendant la génération.'); });

});



document.getElementById('fixDefaultsBtn').addEventListener('click', () => {

  const missing = AGENTS.filter(a => a.group && (a.defaultPeriod === undefined || a.defaultOffDay === undefined));

  if(missing.length === 0){ alert('Tous les agents ont déjà une période et un OFF par défaut.'); return; }

  const weekIdx0 = 0; // valeurs de départ neutres, indépendantes de la semaine actuelle

  const batch = db.batch();

  missing.forEach(a => {

    batch.update(db.collection('agents').doc(a.id), {

      defaultPeriod: resolveAgentPeriod(a, weekIdx0),

      defaultOffDay: offDayIndex(a, weekIdx0),

    });

  });

  batch.commit()

    .then(() => alert(`Période et OFF par défaut réglés pour ${missing.length} agent(s). Tu peux les ajuster un par un dans "Modifier".`))

    .catch(err => { console.error(err); alert('Erreur pendant la génération.'); });

});



/* ================= RENAME ZONES ================= */

document.getElementById('renameZonesBtn').addEventListener('click', () => {

  const container = document.getElementById('zoneRenameInputs');

  container.innerHTML = ZONES.map(z => `

    <label class="form-label">${z}</label>

    <input type="text" class="login-input zone-rename-input" data-zone-key="${z}" value="${getZoneLabel(z)}">

  `).join('');

  document.getElementById('renameZonesError').textContent = '';

  document.getElementById('renameZonesOverlay').classList.remove('hidden');

});

document.getElementById('cancelRenameZonesBtn').addEventListener('click', () => {

  document.getElementById('renameZonesOverlay').classList.add('hidden');

});

document.getElementById('renameZonesOverlay').addEventListener('click', e => {

  if(e.target.id === 'renameZonesOverlay') document.getElementById('renameZonesOverlay').classList.add('hidden');

});

document.getElementById('saveRenameZonesBtn').addEventListener('click', () => {

  const inputs = document.querySelectorAll('.zone-rename-input');

  const payload = {};

  inputs.forEach(input => {

    const val = input.value.trim();

    payload[input.dataset.zoneKey] = val || input.dataset.zoneKey;

  });

  db.collection('settings').doc('zoneLabels').set(payload, { merge: true })

    .then(() => document.getElementById('renameZonesOverlay').classList.add('hidden'))

    .catch(err => { document.getElementById('renameZonesError').textContent = 'Erreur lors de l\'enregistrement.'; console.error(err); });

});



document.getElementById('cancelAgentBtn').addEventListener('click', closeAgentModal);

document.getElementById('agentModalOverlay').addEventListener('click', e => { if(e.target.id === 'agentModalOverlay') closeAgentModal(); });

document.getElementById('agentGroupInput').addEventListener('change', toggleGroupFieldsVisibility);



document.getElementById('saveAgentBtn').addEventListener('click', () => {

  const name = document.getElementById('agentNameInput').value.trim();

  const nickname = document.getElementById('agentNicknameInput').value.trim();

  const group = document.getElementById('agentGroupInput').value || null;

  const zone = document.getElementById('agentZoneInput').value;

  const fixedZone = document.getElementById('agentFixedZoneInput').value;

  const pin = document.getElementById('agentPinModalInput').value.trim();

  const email = document.getElementById('agentEmailModalInput').value.trim();

  const defaultPeriod = document.getElementById('agentDefaultPeriodInput').value;

  const defaultOffDay = parseInt(document.getElementById('agentDefaultOffDayInput').value, 10);

  const errEl = document.getElementById('agentModalError');



  if(!name){ errEl.textContent = 'Entrez le nom de l\'agent.'; return; }

  if(email && !/^\S+@\S+\.\S+$/.test(email)){ errEl.textContent = 'Email invalide.'; return; }



  const payload = {

    name, group, pin: pin || null, email: email || null,

    nickname: nickname || name.split(' ').slice(-1)[0],

    zone: group ? null : zone,

    fixedZone: group === 'zakariae' ? fixedZone : null,

    defaultPeriod, defaultOffDay,

  };

  const promise = editingAgentId

    ? db.collection('agents').doc(editingAgentId).update(payload)

    : db.collection('agents').add(payload);



  promise.then(() => closeAgentModal())

    .catch(err => { errEl.textContent = 'Erreur lors de l\'enregistrement.'; console.error(err); });

});



function deleteAgent(agentId){

  const agent = AGENTS.find(a => a.id === agentId);

  if(!confirm(`Supprimer ${agent ? agent.name : 'cet agent'} ?`)) return;

  db.collection('agents').doc(agentId).delete().catch(err => console.error(err));

}



/* ================= PLANNING MODULE ================= */

let currentWeekStart = getWeekStart(new Date());

let planningUnsub = null;

let currentPlanningData = null;



function updateWeekHeader(){

  const weekIdx = getISOWeekIndex(currentWeekStart);

  document.getElementById('weekLabel').textContent = `Semaine ${weekIdx}`;

  document.getElementById('weekRange').textContent = formatWeekRange(currentWeekStart);

  updateDayHeaders();

}



const DAY_HEADER_IDS = ['thLun', 'thMar', 'thMer', 'thJeu', 'thVen', 'thSam', 'thDim'];

const DAY_HEADER_SHORT = ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'];

const ZONE_DAY_HEADER_IDS = ['thZLun', 'thZMar', 'thZMer', 'thZJeu', 'thZVen', 'thZSam', 'thZDim'];

function updateDayHeaders(){

  DAY_HEADER_IDS.forEach((id, i) => {

    const d = new Date(currentWeekStart); d.setDate(d.getDate() + i);

    const el = document.getElementById(id);

    if(el) el.textContent = `${DAY_HEADER_SHORT[i]} ${d.getDate()}`;

  });

  ZONE_DAY_HEADER_IDS.forEach((id, i) => {

    const d = new Date(currentWeekStart); d.setDate(d.getDate() + i);

    const el = document.getElementById(id);

    if(el) el.textContent = `${DAY_HEADER_SHORT[i]} ${d.getDate()}`;

  });

}



document.getElementById('prevWeekBtn').addEventListener('click', () => {

  const d = new Date(currentWeekStart); d.setDate(d.getDate() - 7);

  currentWeekStart = d;

  renderPlanningWeek();

});

document.getElementById('nextWeekBtn').addEventListener('click', () => {

  const d = new Date(currentWeekStart); d.setDate(d.getDate() + 7);

  currentWeekStart = d;

  renderPlanningWeek();

});



function renderPlanningWeek(){

  updateWeekHeader();

  const key = weekKeyFor(currentWeekStart);

  if(planningUnsub) planningUnsub();

  planningUnsub = db.collection('planning').doc(key).onSnapshot(doc => {

    currentPlanningData = doc.exists ? doc.data() : null;

    renderPlanningTable();

    renderZonesView();

    renderAlerts();

  }, err => console.error(err));

}



document.getElementById('generatePlanningBtn').addEventListener('click', () => {

  if(isEditingCurrentWeek() && !confirm('⚠️ Tu es sur la semaine EN COURS. Générer va recalculer tout le planning actif. Continuer ?')) return;

  if(currentPlanningData && currentPlanningData.agents && Object.keys(currentPlanningData.agents).length > 0){

    if(!confirm('Un planning existe déjà pour cette semaine. Le régénérer va écraser les modifications manuelles. Continuer ?')) return;

  }

  const weekIdx = getISOWeekIndex(currentWeekStart);

  const agentsData = {};

  AGENTS.forEach(a => {

    if(!a.group) return;

    const zone = resolveAgentZone(a, weekIdx);

    const period = resolveAgentPeriod(a, weekIdx);

    const offIdx = offDayIndex(a, weekIdx);

    const cells = {};

    DAY_KEYS.forEach((day, i) => {

      if(i === offIdx){ cells[day] = 'OFF'; return; }

      const isPremium = (day === 'ven' || day === 'sam');

      const hours = isPremium ? 8 : 7;

      cells[day] = period === 'JOUR' ? (hours === 8 ? 'M8' : 'M7') : (hours === 8 ? 'N8' : 'N7');

    });

    agentsData[a.id] = { zone, cells };

  });

  const key = weekKeyFor(currentWeekStart);

  db.collection('planning').doc(key).set({ agents: agentsData, generatedAt: firebase.firestore.FieldValue.serverTimestamp() })

    .catch(err => console.error(err));

});



/* ================= LOAD REAL WEEK (planning exact fourni par Najib) ================= */

function normalizeName(str){

  return (str || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z]/g, '');

}

const REAL_WEEK_DATA = [

  { match: "yahyarhallah",     zone: "Temara - Oulad Mtaa", cells: ["OFF","N7","N7","N7","N8","N8","N7"] },

  { match: "aymanechlifane",   zone: "Temara - Oulad Mtaa", cells: ["M7","OFF","M7","M7","M8","M8","M7"] },

  { match: "mehdiaitjebour",   zone: "Temara - Oulad Mtaa", cells: ["N7","N7","OFF","N7","N8","N8","N7"] },

  { match: "achrafakhlij",     zone: "Hay Riad - Soussi",   cells: ["OFF","N7","N7","N7","N8","N8","N7"] },

  { match: "islamelmardi",     zone: "Hay Riad - Soussi",   cells: ["N7","OFF","N7","N7","N8","N8","N7"] },

  { match: "oussamaelatrach",  zone: "Hay Riad - Soussi",   cells: ["M7","M7","M7","OFF","M8","M8","M7"] },

  { match: "youssefderka",     zone: "Agdal - Carrousel",   cells: ["OFF","N7","N7","N7","N8","N8","N7"] },

  { match: "mouradcherradi",   zone: "Agdal - Carrousel",   cells: ["N7","OFF","N7","N7","N8","N8","N7"] },

  { match: "ayoubabssi",       zone: "Agdal - Carrousel",   cells: ["M7","M7","OFF","M7","M8","M8","M7"] },

  { match: "ayoubbenmalek",    zone: "Salé - Medina",       cells: ["N7","N7","N7","OFF","N8","N8","N7"] },

  { match: "hassanazgoune",    zone: "Salé - Medina",       cells: ["OFF","M7","M7","M7","M8","M8","M7"] },

  { match: "charafloummihi",   zone: "Salé - Medina",       cells: ["N7","OFF","N7","N7","N8","N8","N7"] },

  { match: "zakariaebenali",   zone: "Technopolis",         cells: ["M7","M7","OFF","M7","M8","M8","M7"] },

];



document.getElementById('loadRealWeekBtn').addEventListener('click', () => {

  if(AGENTS.length === 0){ alert('Ajoutez d\'abord les agents (page Agents) avant de charger ce planning.'); return; }

  if(currentPlanningData && currentPlanningData.agents && Object.keys(currentPlanningData.agents).length > 0){

    if(!confirm('Un planning existe déjà pour cette semaine. Le remplacer par le planning réel fourni. Continuer ?')) return;

  }

  const agentsData = {};

  let matchedCount = 0;

  AGENTS.forEach(a => {

    const norm = normalizeName(a.name);

    const entry = REAL_WEEK_DATA.find(r => norm.includes(r.match) || r.match.includes(norm.slice(0, 8)));

    if(!entry) return;

    matchedCount++;

    const cells = {};

    DAY_KEYS.forEach((day, i) => { cells[day] = entry.cells[i]; });

    agentsData[a.id] = { zone: entry.zone, cells };

  });

  const key = weekKeyFor(currentWeekStart);

  db.collection('planning').doc(key).set({ agents: agentsData, generatedAt: firebase.firestore.FieldValue.serverTimestamp() })

    .then(() => alert(`Planning réel chargé pour ${matchedCount} agent(s).`))

    .catch(err => console.error(err));

});



function hexToRgba(hex, alpha){

  const h = hex.replace('#', '');

  const r = parseInt(h.substring(0, 2), 16);

  const g = parseInt(h.substring(2, 4), 16);

  const b = parseInt(h.substring(4, 6), 16);

  return `rgba(${r},${g},${b},${alpha})`;

}



// Retourne les infos d'affichage d'un code de shift, qu'il soit dans le catalogue

// standard (M7, N8, OFF...) ou un horaire personnalisé encodé "C-HHMM-HHMM".

function getShiftInfo(code){

  if(SHIFT_CATALOG[code]) return SHIFT_CATALOG[code];

  if(code && code.startsWith('C-')){

    const parts = code.slice(2).split('-');

    if(parts.length === 2 && parts[0].length === 4 && parts[1].length === 4){

      const startLabel = parts[0].slice(0, 2) + ':' + parts[0].slice(2);

      const endLabel = parts[1].slice(0, 2) + ':' + parts[1].slice(2);

      const startHour = parseInt(parts[0].slice(0, 2));

      const startM = startHour * 60 + parseInt(parts[0].slice(2));

      let endM = parseInt(parts[1].slice(0, 2)) * 60 + parseInt(parts[1].slice(2));

      let hours = (endM - startM) / 60;

      if(hours <= 0) hours += 24;

      const cls = (startHour >= 10 && startHour < 17) ? 's-jour' : 's-nuit';

      return { label: `${startLabel}-${endLabel}`, hours: Math.round(hours * 10) / 10, cls };

    }

  }

  return { label: code || 'OFF', hours: 0, cls: 's-quiet' };

}

function encodeCustomShift(startStr, endStr){

  return `C-${startStr.replace(':', '')}-${endStr.replace(':', '')}`;

}



function populateAgentNamesDatalist(){

  const list = document.getElementById('agentNamesList');

  if(!list) return;

  list.innerHTML = AGENTS.map(a => `<option value="${a.name}">`).join('');

}



function moveAgentRow(agentId, direction, orderedList){

  const ids = orderedList.map(a => a.id);

  const idx = ids.indexOf(agentId);

  const targetIdx = idx + direction;

  if(targetIdx < 0 || targetIdx >= ids.length) return;

  [ids[idx], ids[targetIdx]] = [ids[targetIdx], ids[idx]];

  const batch = db.batch();

  ids.forEach((id, i) => batch.update(db.collection('agents').doc(id), { order: i }));

  batch.commit().catch(err => console.error(err));

}



function renderPlanningTable(){

  populateAgentNamesDatalist();

  const tbody = document.getElementById('planningTableBody');

  const table = document.getElementById('planningTable');

  const emptyState = document.getElementById('planningEmptyState');

  tbody.innerHTML = '';



  let groupedAgents = AGENTS.filter(a => a.group);

  if(!currentPlanningData || groupedAgents.length === 0){

    table.classList.add('hidden');

    emptyState.classList.remove('hidden');

    return;

  }

  table.classList.remove('hidden');

  emptyState.classList.add('hidden');



  // Trie par Zone (agents de la même zone regroupés côte à côte), puis par nom

  groupedAgents = [...groupedAgents].sort((a, b) => {

    const oa = typeof a.order === 'number' ? a.order : null;

    const ob = typeof b.order === 'number' ? b.order : null;

    if(oa !== null && ob !== null) return oa - ob;

    if(oa !== null) return -1;

    if(ob !== null) return 1;

    const za = (currentPlanningData.agents && currentPlanningData.agents[a.id] && currentPlanningData.agents[a.id].zone) || '';

    const zb = (currentPlanningData.agents && currentPlanningData.agents[b.id] && currentPlanningData.agents[b.id].zone) || '';

    if(za !== zb) return za.localeCompare(zb);

    return a.name.localeCompare(b.name);

  });



  groupedAgents.forEach(a => {

    const data = currentPlanningData.agents && currentPlanningData.agents[a.id];

    const tr = document.createElement('tr');

    const zoneColor = data && ZONE_COLORS[data.zone];

    const zoneBorder = zoneColor ? `border-bottom:2px solid ${zoneColor};` : '';

    let totalHours = 0;

    let dayCellsHtml = '';

    DAY_KEYS.forEach(day => {

      const code = data && data.cells ? (data.cells[day] || 'OFF') : 'OFF';

      const info = getShiftInfo(code);

      totalHours += info.hours;

      dayCellsHtml += `<td style="${zoneBorder}"><button class="shift-pill ${info.cls}" data-agent="${a.id}" data-day="${day}" data-code="${code}">${info.label}</button></td>`;

    });

    const zoneOptions = ZONES.map(z => `<option value="${z}" ${data && data.zone === z ? 'selected' : ''}>${getZoneLabel(z)}</option>`).join('');

    tr.innerHTML = `

      <td class="sticky-col agent-name-cell" style="${zoneColor ? `border-left:3px solid ${zoneColor}; ${zoneBorder}` : ''}"><input type="text" class="agent-name-edit" list="agentNamesList" data-agent-name="${a.id}" value="${a.name}" title="Modifier le nom"></td>

      <td class="zone-cell" style="${zoneBorder}"><select class="cell-select zone-select" data-agent="${a.id}" style="border-color: var(--line); font-weight:600;">${zoneOptions}</select></td>

      ${dayCellsHtml}

      <td class="total-cell" style="${zoneBorder}">${totalHours}h</td>

      <td style="${zoneBorder}; white-space:nowrap;">

        <button class="icon-btn" data-toggle-period="${a.id}" title="Inverser Matin/Nuit toute la semaine">🔄</button>

        <button class="icon-btn" data-move-up="${a.id}" title="Monter">▲</button>

        <button class="icon-btn" data-move-down="${a.id}" title="Descendre">▼</button>

      </td>

    `;

    tbody.appendChild(tr);

  });



  tbody.querySelectorAll('[data-move-up]').forEach(btn => btn.addEventListener('click', () => moveAgentRow(btn.dataset.moveUp, -1, groupedAgents)));

  tbody.querySelectorAll('[data-move-down]').forEach(btn => btn.addEventListener('click', () => moveAgentRow(btn.dataset.moveDown, 1, groupedAgents)));



  tbody.querySelectorAll('.agent-name-edit').forEach(input => {

    input.addEventListener('focus', () => {

      input.dataset.original = input.value;

      input.value = '';

      input.dispatchEvent(new Event('input'));

    });

    input.addEventListener('blur', () => {

      const newName = input.value.trim();

      const agentId = input.dataset.agentName;

      const original = input.dataset.original || (AGENTS.find(a => a.id === agentId) || {}).name || '';



      if(!newName || newName === original){

        input.value = original;

        return;

      }

      input.value = newName;

      db.collection('agents').doc(agentId).update({ name: newName }).catch(err => console.error(err));

    });

    input.addEventListener('keydown', e => { if(e.key === 'Enter') input.blur(); });

  });



  tbody.querySelectorAll('.zone-select').forEach(sel => {

    sel.addEventListener('change', () => {

      if(isEditingCurrentWeek() && !confirm('Tu modifies le planning de la semaine EN COURS. Continuer ?')){

        renderPlanningTable();

        return;

      }

      const key = weekKeyFor(currentWeekStart);

      db.collection('planning').doc(key).set({

        agents: { [sel.dataset.agent]: { zone: sel.value } }

      }, { merge: true }).catch(err => console.error(err));

    });

  });

  tbody.querySelectorAll('[data-toggle-period]').forEach(btn => {

    btn.addEventListener('click', () => togglePeriod(btn.dataset.togglePeriod));

  });

  tbody.querySelectorAll('.shift-pill').forEach(btn => btn.addEventListener('click', () => openCellEditor(btn)));

}



// La semaine "en cours" = celle qui contient la date d'aujourd'hui.

function isEditingCurrentWeek(){

  return weekKeyFor(currentWeekStart) === weekKeyFor(getWeekStart(new Date()));

}



function togglePeriod(agentId){

  if(!currentPlanningData || !currentPlanningData.agents || !currentPlanningData.agents[agentId]) return;

  if(isEditingCurrentWeek() && !confirm('Tu modifies le planning de la semaine EN COURS. Continuer ?')) return;

  const data = currentPlanningData.agents[agentId];

  const swap = { M7: 'N7', M8: 'N8', N7: 'M7', N8: 'M8', N7B: 'M7', N8B: 'M8', N7C: 'M7', N8C: 'M8' };

  const newCells = {};

  DAY_KEYS.forEach(day => {

    const code = data.cells[day];

    newCells[day] = swap[code] || code; // OFF/CONGE/ABS restent inchangés

  });

  const key = weekKeyFor(currentWeekStart);

  db.collection('planning').doc(key).set({

    agents: { [agentId]: { cells: newCells } }

  }, { merge: true }).catch(err => console.error(err));

}



function openCellEditor(pillBtn){

  const agentId = pillBtn.dataset.agent;

  const day = pillBtn.dataset.day;

  const currentCode = pillBtn.dataset.code;

  const td = pillBtn.closest('td');



  const select = document.createElement('select');

  select.className = 'cell-select';

  Object.keys(SHIFT_CATALOG).forEach(code => {

    const opt = document.createElement('option');

    opt.value = code;

    opt.textContent = SHIFT_CATALOG[code].label;

    if(code === currentCode) opt.selected = true;

    select.appendChild(opt);

  });

  const customOpt = document.createElement('option');

  customOpt.value = '__custom__';

  customOpt.textContent = '🕓 Horaire personnalisé...';

  if(currentCode && currentCode.startsWith('C-')) customOpt.selected = true;

  select.appendChild(customOpt);



  td.innerHTML = '';

  td.appendChild(select);

  select.focus();



  const commitCode = (newCode) => {

    if(isEditingCurrentWeek() && !confirm('Tu modifies le planning de la semaine EN COURS. Continuer ?')){

      renderPlanningTable();

      return;

    }

    const key = weekKeyFor(currentWeekStart);

    db.collection('planning').doc(key).set({

      agents: { [agentId]: { cells: { [day]: newCode } } }

    }, { merge: true }).catch(err => console.error(err));

  };



  select.addEventListener('change', () => {

    if(select.value === '__custom__'){

      showCustomTimeInputs(td, agentId, day, currentCode);

      return;

    }

    commitCode(select.value);

    renderPlanningTable();

  });

  // Note : pas de handler "clic en dehors" — ça causait un bug où il fallait cliquer

  // deux fois pour éditer une autre cellule (le handler recréait le tableau au mauvais

  // moment et détachait le bouton qu'on venait de cliquer). Le tableau se remet à jour

  // tout seul dès qu'une modification est enregistrée (écoute Firestore en temps réel).

}



function showCustomTimeInputs(td, agentId, day, currentCode){

  let startVal = '10:00', endVal = '17:00';

  if(currentCode && currentCode.startsWith('C-')){

    const parts = currentCode.slice(2).split('-');

    if(parts.length === 2){

      startVal = parts[0].slice(0, 2) + ':' + parts[0].slice(2);

      endVal = parts[1].slice(0, 2) + ':' + parts[1].slice(2);

    }

  }

  td.innerHTML = `

    <div style="display:flex; flex-direction:column; gap:3px; align-items:stretch;">

      <input type="time" id="customStartInput" value="${startVal}" style="font-size:11px; padding:3px; border-radius:6px; border:1.5px solid var(--sky);">

      <input type="time" id="customEndInput" value="${endVal}" style="font-size:11px; padding:3px; border-radius:6px; border:1.5px solid var(--sky);">

      <button id="confirmCustomTime" class="btn-primary" style="padding:4px; font-size:10px;">OK</button>

    </div>

  `;

  document.getElementById('confirmCustomTime').addEventListener('click', () => {

    const start = document.getElementById('customStartInput').value;

    const end = document.getElementById('customEndInput').value;

    if(!start || !end) return;

    if(isEditingCurrentWeek() && !confirm('Tu modifies le planning de la semaine EN COURS. Continuer ?')){

      renderPlanningTable();

      return;

    }

    const newCode = encodeCustomShift(start, end);

    const key = weekKeyFor(currentWeekStart);

    db.collection('planning').doc(key).set({

      agents: { [agentId]: { cells: { [day]: newCode } } }

    }, { merge: true }).then(() => renderPlanningTable()).catch(err => console.error(err));

  });

}



function renderZonesView(){

  const tbody = document.getElementById('zonesTableBody');

  tbody.innerHTML = '';

  if(!currentPlanningData){ return; }

  ZONES.forEach(zone => {

    let cellsHtml = `<td class="sticky-col agent-name-cell">${getZoneLabel(zone)}</td>`;

    DAY_KEYS.forEach(day => {

      const agentsHere = AGENTS.filter(a => {

        const d = currentPlanningData.agents && currentPlanningData.agents[a.id];

        return d && d.zone === zone && d.cells && d.cells[day] && d.cells[day] !== 'OFF';

      }).map(a => {

        const code = currentPlanningData.agents[a.id].cells[day];

        let icon = '🌙';

        if(code.startsWith('M')) icon = '☀️';

        else if(code.startsWith('C-')){

          const startHour = parseInt(code.slice(2, 4));

          icon = (startHour >= 10 && startHour < 17) ? '☀️' : '🌙';

        }

        return `${icon} ${a.nickname || a.name.split(' ').slice(-1)[0]}`;

      });

      cellsHtml += `<td style="text-align:left; white-space:normal;">${agentsHere.length ? agentsHere.join('<br>') : '<span style="color:var(--ink-soft)">—</span>'}</td>`;

    });

    const tr = document.createElement('tr');

    tr.innerHTML = cellsHtml;

    tbody.appendChild(tr);

  });

}



function computeWarnings(planningData){

  const warnings = [];

  if(!planningData || !planningData.agents) return warnings;

  AGENTS.filter(a => a.group).forEach(a => {

    const data = planningData.agents[a.id];

    if(!data || !data.cells) return;

    const codes = DAY_KEYS.map(d => data.cells[d] || 'OFF');

    const workDays = codes.filter(c => c !== 'OFF').length;

    const offDays = codes.filter(c => c === 'OFF').length;

    const eightHDays = codes.filter(c => c === 'M8' || c === 'N8' || c === 'N8B' || c === 'N8C').length;

    if(workDays > 6) warnings.push(`${a.name} : plus de 6 jours travaillés cette semaine.`);

    if(offDays === 0) warnings.push(`${a.name} : aucun jour OFF cette semaine.`);

    if(offDays > 1) warnings.push(`${a.name} : plus d'un jour OFF cette semaine.`);

    if(eightHDays > 2) warnings.push(`${a.name} : plus de 2 jours à 8h cette semaine.`);

  });

  return warnings;

}



function renderAlerts(){

  const banner = document.getElementById('alertsBanner');

  if(!currentPlanningData || !currentPlanningData.agents){ banner.classList.add('hidden'); return; }

  const warnings = computeWarnings(currentPlanningData);

  if(warnings.length === 0){ banner.classList.add('hidden'); return; }

  banner.classList.remove('hidden');

  banner.innerHTML = `<b>⚠️ ${warnings.length} alerte(s) — vérifiez le planning (ceci ne bloque rien)</b><ul>${warnings.map(w => `<li>${w}</li>`).join('')}</ul>`;

}



document.querySelectorAll('.ptab').forEach(tab => {

  tab.addEventListener('click', () => {

    document.querySelectorAll('.ptab').forEach(t => t.classList.remove('active'));

    tab.classList.add('active');

    const isZones = tab.dataset.ptab === 'zones';

    document.getElementById('planningGridWrap').classList.toggle('hidden', isZones);

    document.getElementById('zonesGridWrap').classList.toggle('hidden', !isZones);

  });

});



document.getElementById('printPlanningBtn').addEventListener('click', () => window.print());



document.getElementById('exportExcelBtn').addEventListener('click', () => {

  if(!currentPlanningData){ alert('Générez d\'abord le planning de cette semaine.'); return; }

  const rows = [['Agent', 'Zone', ...DAY_LABELS, 'Total heures']];

  AGENTS.filter(a => a.group).forEach(a => {

    const data = currentPlanningData.agents && currentPlanningData.agents[a.id];

    let total = 0;

    const dayVals = DAY_KEYS.map(day => {

      const code = data && data.cells ? (data.cells[day] || 'OFF') : 'OFF';

      const info = getShiftInfo(code);

      total += info.hours;

      return info.label;

    });

    rows.push([a.name, data ? data.zone : '—', ...dayVals, `${total}h`]);

  });

  const ws = XLSX.utils.aoa_to_sheet(rows);

  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, 'Planning');

  XLSX.writeFile(wb, `planning-${weekKeyFor(currentWeekStart)}.xlsx`);

});



/* ================= AUTH STATE (persist session on refresh) ================= */

auth.onAuthStateChanged(user => {

  const nvPortalVisible = !document.getElementById('nvAgentPortal').classList.contains('hidden');

  if(user && user.email && !nvPortalVisible){

    const uid = user.email.split('@opshq.ma')[0];

    const profile = PROFILES.find(p => p.id === uid);

    if(profile) enterApp(profile);

  }

  // Les sessions anonymes (agents) n'entrent jamais automatiquement — le PIN reste obligatoire à chaque fois.

});



/* ================= ADMIN: POINTAGE DU JOUR ================= */

let pointageCurrentDate = new Date();



function dateKeyFor(d){

  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

}



document.getElementById('prevDayBtn').addEventListener('click', () => {

  pointageCurrentDate.setDate(pointageCurrentDate.getDate() - 1);

  pointageCurrentDate = new Date(pointageCurrentDate);

  renderPointageDay();

});

document.getElementById('nextDayBtn').addEventListener('click', () => {

  pointageCurrentDate.setDate(pointageCurrentDate.getDate() + 1);

  pointageCurrentDate = new Date(pointageCurrentDate);

  renderPointageDay();

});

document.getElementById('pointageTodayBtn').addEventListener('click', () => {

  pointageCurrentDate = new Date();

  renderPointageDay();

});



function renderPointageDay(){

  const dateKey = dateKeyFor(pointageCurrentDate);

  document.getElementById('pointageDayLabel').textContent = pointageCurrentDate.toLocaleDateString('fr-FR', { weekday: 'long', day: 'numeric', month: 'long' });

  document.getElementById('pointageDateSub').textContent = dateKey;



  // Deux sources : l'ancien système "pointage" (par agentId, check-in/check-out manuels) ET

  // les check-ins réels des agents Planning NV (collection nvCheckins, par nom) — un agent

  // Planning NV n'a jamais de doc dans "pointage" puisqu'il ne passe pas par cet ancien portail.

  Promise.all([

    db.collection('pointage').where('date', '==', dateKey).get(),

    db.collection('nvCheckins').where('date', '==', dateKey).get(),

  ]).then(([pointageSnap, nvSnap]) => {

    const byAgent = {};

    pointageSnap.docs.forEach(doc => { byAgent[doc.data().agentId] = doc.data(); });

    const nvByName = {};

    nvSnap.docs.forEach(doc => { nvByName[doc.data().agentName] = doc.data(); });



    const tbody = document.getElementById('pointageTableBody');

    tbody.innerHTML = '';

    const sortedAgents = [...AGENTS].sort((a, b) => a.name.localeCompare(b.name));



    sortedAgents.forEach(a => {

      const nvCheckin = nvByName[a.name];

      let inT, outT, statusHtml;

      if(nvCheckin){

        // Agent Planning NV : check-in réel via le portail (pas de check-out — flux simplifié).

        inT = nvCheckin.checkInAt && nvCheckin.checkInAt.toDate ? nvCheckin.checkInAt.toDate().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) : '—';

        outT = '—';

        statusHtml = `<span class="badge jour">✅ Check-in fait</span>`;

      } else {

        const d = byAgent[a.id];

        inT = d && d.checkIn && d.checkIn.toDate ? d.checkIn.toDate().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) : '—';

        outT = d && d.checkOut && d.checkOut.toDate ? d.checkOut.toDate().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }) : '—';

        if(!d || !d.checkIn) statusHtml = `<span class="badge" style="background:rgba(196,89,58,0.14); color:var(--red);">Absent</span>`;

        else if(d.checkIn && !d.checkOut) statusHtml = `<span class="badge jour">En service</span>`;

        else statusHtml = `<span class="badge fix">Terminé</span>`;

      }



      const tr = document.createElement('tr');

      tr.innerHTML = `<td><b>${a.name}</b></td><td>${inT}</td><td>${outT}</td><td>${statusHtml}</td>`;

      tbody.appendChild(tr);

    });

  }).catch(err => console.error(err));

}



/* ================= ABSENCES (cycle 21 → 21, base 26 jours) ================= */

const ABSENCE_TYPES = {

  maladie:     { label: 'Absence maladie',        delta: -1, cls: 'abs-maladie' },

  injustifiee: { label: 'Absence injustifiée',    delta: -1, cls: 'abs-injustifiee' },

  conge:       { label: 'Congé payé',             delta: -1, cls: 'abs-conge' },

  ferie:       { label: 'Jour férié travaillé',   delta: +1, cls: 'abs-ferie' },

  double:      { label: 'Jour doublé',            delta: +1, cls: 'abs-double' },

  off_travaille: { label: 'Jour OFF travaillé',   delta: +1, cls: 'abs-off' },

};

const CYCLE_BASE_DAYS = 26;

let currentCycleRef = new Date();

let editingAbsenceId = null;



function getCycleFor(date){

  const d = new Date(date);

  let year = d.getFullYear(), month = d.getMonth();

  if(d.getDate() < 21){

    month -= 1;

    if(month < 0){ month = 11; year -= 1; }

  }

  const start = new Date(year, month, 21);

  const end = new Date(year, month + 1, 20);

  const key = `${start.getFullYear()}-${String(start.getMonth() + 1).padStart(2, '0')}`;

  return { start, end, key };

}

function formatCycleLabel(cycle){

  const startStr = cycle.start.toLocaleDateString('fr-FR', { day: 'numeric', month: 'long' });

  const endStr = cycle.end.toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' });

  return `${startStr} → ${endStr}`;

}



document.getElementById('prevCycleBtn').addEventListener('click', () => {

  currentCycleRef.setMonth(currentCycleRef.getMonth() - 1);

  currentCycleRef = new Date(currentCycleRef);

  renderAbsencesView();

});

document.getElementById('nextCycleBtn').addEventListener('click', () => {

  currentCycleRef.setMonth(currentCycleRef.getMonth() + 1);

  currentCycleRef = new Date(currentCycleRef);

  renderAbsencesView();

});

document.getElementById('cycleTodayBtn').addEventListener('click', () => {

  currentCycleRef = new Date();

  renderAbsencesView();

});



let currentCycleRecords = [];



/* ================= RAPPORTS (rapports & remarques envoyés par les agents Planning NV) ================= */

let rapportsUnsub = null;

let rapportsFilterBuilt = false;



function renderRapportsView(){

  markRapportsSeenNow();

  if(!rapportsFilterBuilt){

    const agentSelect = document.getElementById('rapportsFilterAgent');

    NV_ALL_AGENTS.forEach(name => {

      const opt = document.createElement('option');

      opt.value = name; opt.textContent = name;

      agentSelect.appendChild(opt);

    });

    document.getElementById('rapportsFilterAgent').addEventListener('change', renderRapportsList);

    document.getElementById('rapportsFilterType').addEventListener('change', renderRapportsList);

    document.getElementById('rapportsExportBtn').addEventListener('click', exportRapportsCsv);

    rapportsFilterBuilt = true;

  }



  if(rapportsUnsub) rapportsUnsub();

  rapportsUnsub = db.collection('rapports').orderBy('createdAt', 'desc').limit(200).onSnapshot(snap => {

    RAPPORTS_CACHE = snap.docs.map(d => ({ id: d.id, ...d.data() }));

    renderRapportsStats();

    renderRapportsMissing();

    renderRapportsList();

  }, err => console.error(err));

}



function renderRapportsStats(){

  const grid = document.getElementById('rapportsStatGrid');

  const total = RAPPORTS_CACHE.length;

  const nRapports = RAPPORTS_CACHE.filter(r => r.type === 'rapport').length;

  const nRemarques = RAPPORTS_CACHE.filter(r => r.type === 'remarque').length;

  const nUrgent = RAPPORTS_CACHE.filter(r => r.urgent).length;

  const weekAgo = Date.now() - 7 * 86400000;

  const thisWeek = RAPPORTS_CACHE.filter(r => r.createdAt && r.createdAt.toMillis && r.createdAt.toMillis() >= weekAgo).length;

  const counts = {};

  RAPPORTS_CACHE.forEach(r => { counts[r.agentName] = (counts[r.agentName] || 0) + 1; });

  const topAgent = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];



  grid.innerHTML = `

    <div class="stat-card" style="border-left-color:var(--red);">

      <div class="label">🚨 Urgents</div>

      <div class="value" style="color:${nUrgent ? 'var(--red)' : 'inherit'};">${nUrgent}</div>

      <div class="sub">Nécessitent une attention</div>

    </div>

    <div class="stat-card accent-blue">

      <div class="label">Total envoyés</div>

      <div class="value">${total}</div>

      <div class="sub">${nRapports} rapports · ${nRemarques} remarques</div>

    </div>

    <div class="stat-card accent-green">

      <div class="label">Cette semaine</div>

      <div class="value">${thisWeek}</div>

      <div class="sub">Sur les 7 derniers jours</div>

    </div>

    <div class="stat-card accent-purple">

      <div class="label">Agent le plus actif</div>

      <div class="value" style="font-size:18px;">${topAgent ? topAgent[0].split(' ')[0] : '—'}</div>

      <div class="sub">${topAgent ? topAgent[1] + ' envoi(s)' : 'Aucun envoi'}</div>

    </div>

  `;

}



function renderRapportsMissing(){

  const card = document.getElementById('rapportsMissingCard');

  const weekAgo = Date.now() - 7 * 86400000;

  const activeNames = new Set(RAPPORTS_CACHE.filter(r => r.createdAt && r.createdAt.toMillis && r.createdAt.toMillis() >= weekAgo).map(r => r.agentName));

  const missing = NV_ALL_AGENTS.filter(name => !activeNames.has(name));



  if(missing.length === 0){

    card.innerHTML = `<div style="display:flex; align-items:center; gap:8px; font-size:13px; color:var(--green);">✅ Tous les agents ont envoyé au moins un rapport ou une remarque cette semaine.</div>`;

    return;

  }

  card.innerHTML = `

    <div style="font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:.04em; color:var(--ink-soft); margin-bottom:8px;">⚠️ Aucun rapport cette semaine</div>

    <div style="display:flex; flex-wrap:wrap; gap:6px;">

      ${missing.map(name => `<span class="badge" style="background:rgba(180,83,9,0.14); color:#B45309;">${name}</span>`).join('')}

    </div>

  `;

}



function exportRapportsCsv(){

  const rows = [['Date', 'Agent', 'Type', 'Urgent', 'Texte']];

  RAPPORTS_CACHE.forEach(r => {

    const when = r.createdAt && r.createdAt.toDate ? r.createdAt.toDate().toLocaleString('fr-FR') : (r.date || '');

    rows.push([when, r.agentName, r.type, r.urgent ? 'Oui' : 'Non', (r.text || '').replace(/\n/g, ' ')]);

  });

  const csv = rows.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');

  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });

  const url = URL.createObjectURL(blob);

  const a = document.createElement('a');

  a.href = url; a.download = `rapports_${dateKeyFor(new Date())}.csv`;

  a.click();

  URL.revokeObjectURL(url);

}



let RAPPORTS_CACHE = [];



function renderRapportsList(){

  const agentFilter = document.getElementById('rapportsFilterAgent').value;

  const typeFilter = document.getElementById('rapportsFilterType').value;

  const list = document.getElementById('rapportsList');

  const emptyState = document.getElementById('rapportsEmptyState');



  const filtered = RAPPORTS_CACHE.filter(r =>

    (!agentFilter || r.agentName === agentFilter) &&

    (!typeFilter || r.type === typeFilter)

  ).sort((a, b) => (b.urgent ? 1 : 0) - (a.urgent ? 1 : 0));



  if(filtered.length === 0){

    list.innerHTML = '';

    emptyState.classList.remove('hidden');

    return;

  }

  emptyState.classList.add('hidden');



  list.innerHTML = filtered.map(r => {

    const when = r.createdAt && r.createdAt.toDate ? r.createdAt.toDate().toLocaleString('fr-FR', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' }) : (r.date || '');

    const typeLabel = r.type === 'remarque' ? '💬 Remarque' : r.type === 'urgent' ? '🚨 Urgent' : '📩 Rapport';

    const typeColor = r.type === 'remarque' ? '#FBBF24' : r.type === 'urgent' ? 'var(--red)' : '#38BDF8';

    const seenLabel = r.seenAt

      ? `<span style="font-size:11px; color:var(--green); font-weight:700;">✅ Lu</span>`

      : `<button class="nv-mini-link" data-mark-seen="${r.id}" style="font-size:11px;">Marquer comme lu</button>`;

    const replyBlock = r.adminReply

      ? `<div style="margin-top:8px; padding:8px 12px; background:rgba(0,160,130,0.08); border-radius:10px; border-left:3px solid var(--sky);">

          <div style="font-size:10px; font-weight:700; text-transform:uppercase; color:var(--sky-ink); margin-bottom:2px;">Ta réponse</div>

          <div style="font-size:12.5px; color:var(--ink); white-space:pre-wrap;">${r.adminReply}</div>

        </div>`

      : `<div style="margin-top:8px; display:flex; gap:6px;">

          <input type="text" class="login-input" data-reply-input="${r.id}" placeholder="Répondre à ${r.agentName.split(' ')[0]}..." style="flex:1; padding:8px 10px; margin:0; font-size:12.5px;" />

          <button class="btn-ghost" data-send-reply="${r.id}" style="border:1px solid var(--line); font-size:12px; padding:8px 12px;">Envoyer</button>

        </div>`;

    return `

      <div style="padding:14px 18px; border-bottom:1px solid var(--line); ${r.urgent ? 'background:rgba(220,38,38,0.06); border-left:4px solid var(--red);' : ''}">

        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; flex-wrap:wrap; gap:6px;">

          <span style="display:flex; align-items:center; gap:8px;">

            ${r.urgent ? '<span style="font-size:11px; font-weight:800; color:#fff; background:var(--red); padding:2px 8px; border-radius:20px;">🚨 URGENT</span>' : ''}

            <span style="font-weight:700; font-size:13.5px;">${r.agentName}</span>

          </span>

          <span style="display:flex; align-items:center; gap:8px;">

            <span style="font-size:11px; font-weight:700; color:${typeColor};">${typeLabel}</span>

            <span style="font-size:11px; color:var(--ink-soft);">${when}</span>

            ${seenLabel}

            <button class="icon-btn danger" data-del-rapport="${r.id}" title="Supprimer"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2m3 0v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6h14z"/></svg></button>

          </span>

        </div>

        <div style="font-size:13px; color:var(--ink); line-height:1.5; white-space:pre-wrap;">${r.text}</div>

        ${replyBlock}

      </div>

    `;

  }).join('');



  list.querySelectorAll('[data-del-rapport]').forEach(btn => {

    btn.addEventListener('click', () => {

      if(!confirm('Supprimer cet envoi ?')) return;

      db.collection('rapports').doc(btn.dataset.delRapport).delete().catch(err => console.error(err));

    });

  });

  list.querySelectorAll('[data-mark-seen]').forEach(btn => {

    btn.addEventListener('click', () => {

      db.collection('rapports').doc(btn.dataset.markSeen).set({ seenAt: firebase.firestore.FieldValue.serverTimestamp() }, { merge: true }).catch(err => console.error(err));

    });

  });

  list.querySelectorAll('[data-send-reply]').forEach(btn => {

    btn.addEventListener('click', () => {

      const id = btn.dataset.sendReply;

      const input = list.querySelector(`[data-reply-input="${id}"]`);

      const text = input.value.trim();

      if(!text) return;

      db.collection('rapports').doc(id).set({

        adminReply: text,

        repliedAt: firebase.firestore.FieldValue.serverTimestamp(),

        seenAt: firebase.firestore.FieldValue.serverTimestamp(),

      }, { merge: true }).catch(err => console.error(err));

    });

  });

}



function renderAbsencesView(){

  const cycle = getCycleFor(currentCycleRef);

  document.getElementById('cycleLabel').textContent = `Cycle ${cycle.key}`;

  document.getElementById('cycleSub').textContent = formatCycleLabel(cycle);



  db.collection('absences').where('cycleKey', '==', cycle.key).get().then(snap => {

    currentCycleRecords = snap.docs.map(d => ({ id: d.id, ...d.data() }));

    renderAbsencesSummary();

    renderAbsencesDetail();

  }).catch(err => console.error(err));

}



function renderAbsencesSummary(){

  const tbody = document.getElementById('absencesSummaryBody');

  tbody.innerHTML = '';

  const sortedAgents = [...AGENTS].sort((a, b) => a.name.localeCompare(b.name));

  const typeKeys = Object.keys(ABSENCE_TYPES);



  sortedAgents.forEach(a => {

    const recs = currentCycleRecords.filter(r => r.agentId === a.id);

    const counts = {};

    typeKeys.forEach(k => counts[k] = 0);

    recs.forEach(r => { if(counts[r.type] !== undefined) counts[r.type]++; });

    const delta = typeKeys.reduce((sum, k) => sum + counts[k] * ABSENCE_TYPES[k].delta, 0);

    const total = CYCLE_BASE_DAYS + delta;

    const totalAbsences = (counts.maladie || 0) + (counts.injustifiee || 0) + (counts.conge || 0);

    const presence = CYCLE_BASE_DAYS - totalAbsences;



    const tr = document.createElement('tr');

    tr.innerHTML = `

      <td><b>${a.name}</b></td>

      <td>${counts.maladie || ''}</td>

      <td>${counts.injustifiee || ''}</td>

      <td>${counts.conge || ''}</td>

      <td style="font-weight:700; color:var(--red);">${totalAbsences || ''}</td>

      <td style="font-weight:700; color:var(--green);">${presence}j</td>

      <td>${counts.ferie || ''}</td>

      <td>${counts.double || ''}</td>

      <td>${counts.off_travaille || ''}</td>

      <td class="total-cell">${total}j</td>

    `;

    tbody.appendChild(tr);

  });

}



function renderAbsencesDetail(){

  const tbody = document.getElementById('absencesDetailBody');

  const table = tbody.closest('table');

  const emptyState = document.getElementById('absencesEmptyState');

  tbody.innerHTML = '';



  if(currentCycleRecords.length === 0){

    table.classList.add('hidden');

    emptyState.classList.remove('hidden');

    return;

  }

  table.classList.remove('hidden');

  emptyState.classList.add('hidden');



  const sorted = [...currentCycleRecords].sort((a, b) => {

    const nameCmp = (a.agentName || '').localeCompare(b.agentName || '');

    if(nameCmp !== 0) return nameCmp;

    const typeCmp = (a.type || '').localeCompare(b.type || '');

    if(typeCmp !== 0) return typeCmp;

    return (a.date || '').localeCompare(b.date || '');

  });



  // Regroupe les jours consécutifs (même agent + même type + dates qui se suivent) en une seule ligne "plage".

  const groups = [];

  sorted.forEach(r => {

    const last = groups[groups.length - 1];

    const oneDayAfter = last && last.records[last.records.length - 1].date &&

      new Date(r.date) - new Date(last.records[last.records.length - 1].date) === 86400000;

    if(last && last.agentId === r.agentId && last.type === r.type && oneDayAfter){

      last.records.push(r);

    } else {

      groups.push({ agentId: r.agentId, agentName: r.agentName, type: r.type, records: [r] });

    }

  });



  groups.forEach((g, gi) => {

    const info = ABSENCE_TYPES[g.type] || { label: g.type, cls: '' };

    const n = g.records.length;

    const groupId = `absgrp-${gi}`;



    if(n === 1){

      const r = g.records[0];

      const tr = document.createElement('tr');

      tr.innerHTML = `

        <td><b>${r.agentName}</b></td>

        <td>${r.date}</td>

        <td><span class="badge ${info.cls}">${info.label}</span></td>

        <td>${r.note || '—'}</td>

        <td>

          <div class="row-actions">

            <button class="icon-btn" data-edit-abs="${r.id}" title="Modifier"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 013 3L7 19l-4 1 1-4z"/></svg></button>

            <button class="icon-btn danger" data-delete-abs="${r.id}" title="Supprimer"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2m3 0v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6h14z"/></svg></button>

          </div>

        </td>

      `;

      tbody.appendChild(tr);

      return;

    }



    // Ligne "plage" (repliée par défaut) + lignes individuelles cachées, dépliables.

    const groupTr = document.createElement('tr');

    groupTr.innerHTML = `

      <td><b>${g.agentName}</b></td>

      <td>${g.records[0].date} → ${g.records[n - 1].date} <span style="color:var(--ink-soft); font-weight:400;">(${n} jours)</span></td>

      <td><span class="badge ${info.cls}">${info.label}</span></td>

      <td>—</td>

      <td>

        <div class="row-actions">

          <button class="btn-ghost" data-toggle-group="${groupId}" style="padding:4px 10px; font-size:12px;">▾ Voir les jours</button>

          <button class="icon-btn danger" data-delete-group="${groupId}" title="Supprimer toute la plage"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2m3 0v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6h14z"/></svg></button>

        </div>

      </td>

    `;

    tbody.appendChild(groupTr);



    g.records.forEach(r => {

      const subTr = document.createElement('tr');

      subTr.className = 'abs-group-row hidden';

      subTr.dataset.group = groupId;

      subTr.innerHTML = `

        <td style="padding-left:24px; color:var(--ink-soft);">↳ ${r.agentName}</td>

        <td>${r.date}</td>

        <td><span class="badge ${info.cls}">${info.label}</span></td>

        <td>${r.note || '—'}</td>

        <td>

          <div class="row-actions">

            <button class="icon-btn" data-edit-abs="${r.id}" title="Modifier"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 013 3L7 19l-4 1 1-4z"/></svg></button>

            <button class="icon-btn danger" data-delete-abs="${r.id}" title="Supprimer ce jour uniquement"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2m3 0v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6h14z"/></svg></button>

          </div>

        </td>

      `;

      tbody.appendChild(subTr);

    });

  });



  tbody.querySelectorAll('[data-edit-abs]').forEach(btn => btn.addEventListener('click', () => openAbsenceModal(btn.dataset.editAbs)));

  tbody.querySelectorAll('[data-delete-abs]').forEach(btn => btn.addEventListener('click', () => deleteAbsence(btn.dataset.deleteAbs)));

  tbody.querySelectorAll('[data-toggle-group]').forEach(btn => btn.addEventListener('click', () => {

    const gid = btn.dataset.toggleGroup;

    const rows = tbody.querySelectorAll(`.abs-group-row[data-group="${gid}"]`);

    const nowHidden = rows[0] && rows[0].classList.contains('hidden');

    rows.forEach(r => r.classList.toggle('hidden', !nowHidden));

    btn.textContent = nowHidden ? '▴ Cacher les jours' : '▾ Voir les jours';

  }));

  tbody.querySelectorAll('[data-delete-group]').forEach(btn => btn.addEventListener('click', () => {

    const gid = btn.dataset.deleteGroup;

    const idx = parseInt(gid.split('-')[1], 10);

    const g = groups[idx];

    if(!confirm(`Supprimer toute la plage : ${g.agentName} — ${g.records[0].date} → ${g.records[g.records.length - 1].date} (${g.records.length} jours) ?`)) return;

    const batch = db.batch();

    g.records.forEach(r => batch.delete(db.collection('absences').doc(r.id)));

    batch.commit().then(() => renderAbsencesView()).catch(err => console.error(err));

  }));

}



function populateAbsenceAgentSelect(){

  const sel = document.getElementById('absenceAgentInput');

  const sortedAgents = [...AGENTS].sort((a, b) => a.name.localeCompare(b.name));

  sel.innerHTML = '<option value="" disabled selected>-- Choisir un agent --</option>' +

    sortedAgents.map(a => `<option value="${a.id}">${a.name}</option>`).join('');

}



function openAbsenceModal(absenceId = null){

  editingAbsenceId = absenceId;

  populateAbsenceAgentSelect();

  const title = document.getElementById('absenceModalTitle');

  const endDateInput = document.getElementById('absenceDateEndInput');

  document.getElementById('absenceModalError').textContent = '';



  if(absenceId){

    const rec = currentCycleRecords.find(r => r.id === absenceId);

    title.textContent = 'Modifier';

    document.getElementById('absenceAgentInput').value = rec.agentId;

    document.getElementById('absenceTypeInput').value = rec.type;

    document.getElementById('absenceDateInput').value = rec.date;

    document.getElementById('absenceNoteInput').value = rec.note || '';

    endDateInput.value = '';

    endDateInput.disabled = true; // pas de plage en mode "modifier" (une seule entrée à la fois)

  } else {

    title.textContent = 'Déclarer une absence / un jour';

    document.getElementById('absenceAgentInput').value = '';

    document.getElementById('absenceTypeInput').value = 'maladie';

    document.getElementById('absenceDateInput').value = new Date().toISOString().slice(0, 10);

    document.getElementById('absenceNoteInput').value = '';

    endDateInput.value = '';

    endDateInput.disabled = false;

  }

  document.getElementById('absenceModalOverlay').classList.remove('hidden');

}



document.getElementById('addAbsenceBtn').addEventListener('click', () => openAbsenceModal(null));

document.getElementById('cancelAbsenceBtn').addEventListener('click', () => document.getElementById('absenceModalOverlay').classList.add('hidden'));

document.getElementById('absenceModalOverlay').addEventListener('click', e => { if(e.target.id === 'absenceModalOverlay') document.getElementById('absenceModalOverlay').classList.add('hidden'); });



document.getElementById('saveAbsenceBtn').addEventListener('click', () => {

  const agentId = document.getElementById('absenceAgentInput').value;

  const type = document.getElementById('absenceTypeInput').value;

  const date = document.getElementById('absenceDateInput').value;

  const dateEnd = document.getElementById('absenceDateEndInput').value;

  const note = document.getElementById('absenceNoteInput').value.trim();

  const errEl = document.getElementById('absenceModalError');

  errEl.textContent = '';



  if(!agentId){ errEl.textContent = 'Choisis un agent.'; return; }

  if(!date){ errEl.textContent = 'Choisis une date.'; return; }



  const agent = AGENTS.find(a => a.id === agentId);

  const agentLabel = agent ? agent.name : '(agent inconnu)';

  const typeLabel = (ABSENCE_TYPES[type] && ABSENCE_TYPES[type].label) || type;



  // Plage de dates (sans édition existante) : on crée une entrée par jour.

  if(!editingAbsenceId && dateEnd && dateEnd > date){

    const dayCount = Math.round((new Date(dateEnd) - new Date(date)) / 86400000) + 1;

    if(dayCount > 62){ errEl.textContent = 'La plage est trop longue (max ~2 mois).'; return; }

    const dayStrs = [];

    let dd = new Date(date);

    for(let i = 0; i < dayCount; i++){ dayStrs.push(dd.toISOString().slice(0, 10)); dd.setDate(dd.getDate() + 1); }



    db.collection('absences').where('agentId', '==', agentId).where('date', 'in', dayStrs.slice(0, 10)).get()

      .then(snap => {

        // Vérifie aussi le reste de la plage si > 10 jours (limite Firestore "in")

        const checks = [];

        for(let i = 10; i < dayStrs.length; i += 10){

          checks.push(db.collection('absences').where('agentId', '==', agentId).where('date', 'in', dayStrs.slice(i, i + 10)).get());

        }

        return Promise.all([Promise.resolve(snap), ...checks]);

      })

      .then(snaps => {

        const dupes = snaps.flatMap(s => s.docs.map(d => d.data().date));

        if(dupes.length > 0){

          errEl.textContent = `${agentLabel} a déjà une entrée le ${dupes[0]}. Corrige ou supprime-la d'abord.`;

          return;

        }

        if(!confirm(`Confirmer : ${agentLabel} — ${dayStrs[0]} → ${dayStrs[dayStrs.length - 1]} (${dayCount} jours) — ${typeLabel} ?`)) return;

        const batch = db.batch();

        dayStrs.forEach(dStr => {

          const cycle = getCycleFor(new Date(dStr));

          const ref = db.collection('absences').doc();

          batch.set(ref, { agentId, agentName: agent ? agent.name : '', type, date: dStr, note: note || null, cycleKey: cycle.key });

        });

        batch.commit().then(() => {

          document.getElementById('absenceModalOverlay').classList.add('hidden');

          renderAbsencesView();

        }).catch(err => { errEl.textContent = 'Erreur lors de l\'enregistrement.'; console.error(err); });

      }).catch(err => { errEl.textContent = 'Erreur lors de la vérification.'; console.error(err); });

    return;

  }



  const cycle = getCycleFor(new Date(date));

  const payload = { agentId, agentName: agent ? agent.name : '', type, date, note: note || null, cycleKey: cycle.key };



  const checkDupe = editingAbsenceId

    ? Promise.resolve(null)

    : db.collection('absences').where('agentId', '==', agentId).where('date', '==', date).get()

        .then(snap => snap.empty ? null : snap.docs[0]);



  checkDupe.then(dupeDoc => {

    if(dupeDoc){

      errEl.textContent = `${agentLabel} a déjà une entrée le ${date}. Corrige ou supprime-la d'abord.`;

      return;

    }

    if(!confirm(`Confirmer : ${agentLabel} — ${date} — ${typeLabel} ?`)) return;

    const promise = editingAbsenceId

      ? db.collection('absences').doc(editingAbsenceId).update(payload)

      : db.collection('absences').add(payload);



    promise.then(() => {

      document.getElementById('absenceModalOverlay').classList.add('hidden');

      renderAbsencesView();

    }).catch(err => { errEl.textContent = 'Erreur lors de l\'enregistrement.'; console.error(err); });

  }).catch(err => { errEl.textContent = 'Erreur lors de la vérification.'; console.error(err); });

});



function deleteAbsence(id){

  if(!confirm('Supprimer cette entrée ?')) return;

  db.collection('absences').doc(id).delete().then(() => renderAbsencesView()).catch(err => console.error(err));

}



function buildAbsenceExportRows(records, baseDays){

  const rows = [['Agent', 'Jours de base', 'Maladie', 'Injustifiée', 'Congé', 'Total absences', 'Jours présence', 'Férié travaillé', 'Doublé', 'OFF travaillé', 'Total ajusté']];

  const sortedAgents = [...AGENTS].sort((a, b) => a.name.localeCompare(b.name));

  const typeKeys = Object.keys(ABSENCE_TYPES);

  sortedAgents.forEach(a => {

    const recs = records.filter(r => r.agentId === a.id);

    const counts = {};

    typeKeys.forEach(k => counts[k] = 0);

    recs.forEach(r => { if(counts[r.type] !== undefined) counts[r.type]++; });

    const delta = typeKeys.reduce((sum, k) => sum + counts[k] * ABSENCE_TYPES[k].delta, 0);

    const totalAbsences = (counts.maladie || 0) + (counts.injustifiee || 0) + (counts.conge || 0);

    rows.push([a.name, baseDays, counts.maladie, counts.injustifiee, counts.conge, totalAbsences, baseDays - totalAbsences, counts.ferie, counts.double, counts.off_travaille, baseDays + delta]);

  });

  return rows;

}



function buildAbsenceDetailRows(records){

  const rows = [['Agent', 'Date', 'Type', 'Note']];

  const sorted = [...records].sort((a, b) => {

    const nameCmp = (a.agentName || '').localeCompare(b.agentName || '');

    if(nameCmp !== 0) return nameCmp;

    return (a.date || '').localeCompare(b.date || '');

  });

  sorted.forEach(r => {

    const info = ABSENCE_TYPES[r.type] || { label: r.type };

    rows.push([r.agentName, r.date, info.label, r.note || '']);

  });

  return rows;

}



document.getElementById('exportAbsencesBtn').addEventListener('click', () => {

  const cycle = getCycleFor(currentCycleRef);

  const summaryRows = buildAbsenceExportRows(currentCycleRecords, CYCLE_BASE_DAYS);

  const detailRows = buildAbsenceDetailRows(currentCycleRecords);

  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryRows), 'Résumé');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(detailRows), 'Détail');

  XLSX.writeFile(wb, `absences-${cycle.key}.xlsx`);

});



/* ================= EXPORT: PLAGE PERSONNALISÉE ================= */

document.getElementById('exportRangeBtn').addEventListener('click', () => {

  document.getElementById('rangeFromInput').value = '';

  document.getElementById('rangeToInput').value = '';

  document.getElementById('rangeExportError').textContent = '';

  document.getElementById('rangeExportOverlay').classList.remove('hidden');

});

document.getElementById('cancelRangeExportBtn').addEventListener('click', () => {

  document.getElementById('rangeExportOverlay').classList.add('hidden');

});

document.getElementById('rangeExportOverlay').addEventListener('click', e => {

  if(e.target.id === 'rangeExportOverlay') document.getElementById('rangeExportOverlay').classList.add('hidden');

});



document.getElementById('confirmRangeExportBtn').addEventListener('click', () => {

  const from = document.getElementById('rangeFromInput').value;

  const to = document.getElementById('rangeToInput').value;

  const errEl = document.getElementById('rangeExportError');

  if(!from || !to){ errEl.textContent = 'Choisis les deux dates.'; return; }

  if(from > to){ errEl.textContent = 'La date de début doit être avant la date de fin.'; return; }



  errEl.textContent = 'Génération...';

  // Filtres d'inégalité sur un seul champ ("date") — pas d'index composite nécessaire.

  db.collection('absences').where('date', '>=', from).where('date', '<=', to).get()

    .then(snap => {

      const records = snap.docs.map(d => ({ id: d.id, ...d.data() }));

      // Nombre de jours calendaires dans la plage choisie, utilisé comme "jours de base".

      const dayCount = Math.round((new Date(to) - new Date(from)) / 86400000) + 1;

      const summaryRows = buildAbsenceExportRows(records, dayCount);

      const detailRows = buildAbsenceDetailRows(records);

      const wb = XLSX.utils.book_new();

      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryRows), 'Résumé');

      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(detailRows), 'Détail');

      XLSX.writeFile(wb, `absences-${from}_${to}.xlsx`);

      document.getElementById('rangeExportOverlay').classList.add('hidden');

    })

    .catch(err => { console.error(err); errEl.textContent = 'Erreur lors de la génération.'; });

});



/* ================= PLANNING NV (système indépendant : 12 agents, 4 zones, hotspots, 3 shifts) ================= */

let currentWeekStartNv = getWeekStart(new Date());

let NV_OVERRIDES = {};

let nvOverridesUnsub = null;



// Chaque zone a une liste de "slots" (dans le même ordre que les membres de l'équipe).

// Chaque slot a sa propre liste de hotspots individuels — l'agent qui occupe ce slot

// cette semaine change de hotspot chaque JOUR (rotation quotidienne dans sa liste).

const DEFAULT_HOTSPOTS_NV = {

  "Salé-Rabat": [

    { venues: ["KFC Rabat Ville", "McDonald's Rabat Ville", "KFC Bab El Had", "Océan"] },

    { venues: ["Carrefour Hepar Salé", "Primo Salé", "Burger King Salé", "Centre Salé"] },

    { venues: ["McDonald's Bouregreg", "Burger King Bouregreg", "Marjane Hepar Bouregreg", "Pizza Hut Bouregreg"] },

  ],

  "Agdal": [

    { venues: ["Av. Fal Ould Oumeir", "Pizza Hut Agdal"] },

    { venues: ["Mall Rabat Center", "Mall Carrousel Rabat", "McDonald's Gare Rabat Agdal"] },

    { venues: ["Carrefour Gourmet", "Le Marché Dar Essalam", "Hay Nahda Rabat", "Polyvalente", "Centre commercial Jasmin", "Maymana Souissi", "Carrefour Market Souissi", "snack monalisa", "Mega Mall Rabat"] },

  ],

  "Hay Riad": [

    { venues: ["Marjane Hyper Hay Riad", "McDonald's Hay Riad", "Pizza Hut Marjane Hay Riad", "KFC Marjane Hay Riad"] },

    { venues: ["Carrefour Mehdi Ben Barka", "Av. Annakhil", "Blue Berry Rabat", "Carrefour Market Riad", "Kool Smoothie Almelia", "Marjane Market A.V BENBARKA", "Mahaj Ryad Center"] },

    { venues: ["Prestigia", "Agha Acham Prestigia", "Grillade Al Andalous prestigia", "Marjane Market Prestigia Hay Riad", "L'Atelier Burger Hay riad", "Yoka Sushi Hay Riad", "La Grillardière Hay riad", "Oh My Bun : Buntastic Burgers"] },

  ],

  "Temara": [

    { venues: ["Oulad Mtaa", "Primo Massira 1", "Polyvalente"] },

    { venues: ["Harhoura", "Quick", "Agha Cham", "Paul Harhoura", "Wifak"] },

    { venues: ["McDonald's Temara", "Chawarma Ibn Acham", "Primo", "Michwat Al Akhawayn", "Qrouch"] },

  ],

  "Technopolis": [

    { venues: ["Mcdonald's Arjat", "Carrefour Technopolis", "Marjane Technopolis", "Shawarma Ostorat Sham", "S Tacos", "Primos Chicken", "Yours Brunch", "Boca Mia", "Sushi ARIJ", "Gusto Chicken", "Bocadillos Tanjaoui", "Mahaj Alfawakih", "Vida Souss", "Crepti", "Bari bana", "Shamiyat", "Grill vibe", "175 kg"] },

  ],

};

// État en direct (Firestore) — permet d'ajouter des hotspots depuis l'interface, sans toucher au code.

// NB: chaque slot est un OBJET {venues:[...]} et non un tableau brut, car Firestore interdit

// les tableaux imbriqués directement (array of arrays).

let HOTSPOTS_NV = JSON.parse(JSON.stringify(DEFAULT_HOTSPOTS_NV));

let hotspotsNvUnsub = null;

let nvLocationsUnsub = null;

let NV_LOCATIONS = {};



function isValidHotspotShapeNv(data){

  if(!data || typeof data !== 'object') return false;

  const zoneEntries = Object.entries(data).filter(([k]) => !k.startsWith('_'));

  if(zoneEntries.length === 0) return false;

  return zoneEntries.every(([, slots]) => Array.isArray(slots) && slots.every(s => s && Array.isArray(s.venues)));

}



function listenHotspotsNv(){

  listenNvZoneLabels();

  listenNvRowOrder();

  listenNvVenueLinks();

  listenNvOverrides();

  if(hotspotsNvUnsub) hotspotsNvUnsub();

  hotspotsNvUnsub = db.collection('planningNV').doc('hotspots').onSnapshot(doc => {

    const data = doc.exists ? doc.data() : null;

    if(isValidHotspotShapeNv(data)){

      // Fusionne toute zone présente dans le code (ex: nouvelle zone ajoutée) mais absente

      // des données déjà enregistrées sur Firestore, sans écraser ce qui existe déjà.

      const merged = { ...data };

      let needsWrite = false;

      Object.keys(DEFAULT_HOTSPOTS_NV).forEach(zone => {

        if(!merged[zone]){

          merged[zone] = DEFAULT_HOTSPOTS_NV[zone];

          needsWrite = true;

        }

      });

      if(needsWrite){

        db.collection('planningNV').doc('hotspots').set(merged, { merge: true }).catch(err => console.error(err));

      }

      // D'éventuelles clés internes commençant par "_" (anciens marqueurs de correction) ne sont

      // jamais traitées comme des zones, sinon elles casseraient toutes les boucles sur

      // `Object.keys(HOTSPOTS_NV)` ailleurs dans le code.

      const zonesOnly = {};

      Object.keys(merged).forEach(k => { if(!k.startsWith('_')) zonesOnly[k] = merged[k]; });

      HOTSPOTS_NV = zonesOnly;

    } else {

      // Absent, vide, ou format invalide : on n'écrase JAMAIS automatiquement avec les valeurs

      // par défaut du code (elles sont volontairement minimales et ne doivent pas remplacer les

      // hotspots réels ajoutés au fil du temps par l'admin) — on initialise seulement si le

      // document n'existe pas du tout.

      if(!doc.exists){

        db.collection('planningNV').doc('hotspots').set(DEFAULT_HOTSPOTS_NV).catch(err => console.error(err));

        HOTSPOTS_NV = JSON.parse(JSON.stringify(DEFAULT_HOTSPOTS_NV));

      } else {

        console.error('Document hotspots présent mais de forme invalide — HOTSPOTS_NV conservé tel quel pour éviter toute perte de données.', data);

      }

    }

    refreshNvViews();

  }, err => console.error(err));



  if(nvLocationsUnsub) nvLocationsUnsub();

  nvLocationsUnsub = db.collection('planningNV').doc('locations').onSnapshot(doc => {

    NV_LOCATIONS = doc.exists ? (doc.data() || {}) : {};

    refreshNvViews();

  }, err => console.error(err));

}



function addHotspotNv(zone, slotIdx){

  const value = prompt(`Ajouter un lieu au hotspot #${slotIdx + 1} de "${zone}" :`);

  if(!value || !value.trim()) return;

  const zoneSlots = (HOTSPOTS_NV[zone] || []).map(s => ({ ...s, venues: [...s.venues] }));

  if(!zoneSlots[slotIdx]) zoneSlots[slotIdx] = { venues: [] };

  zoneSlots[slotIdx].venues.push(value.trim());

  const updated = { ...HOTSPOTS_NV, [zone]: zoneSlots };

  db.collection('planningNV').doc('hotspots').set(updated).catch(err => console.error(err));

}

// Liens Google Maps personnalisés par lieu — si absent, on utilise une recherche Maps automatique.

let NV_VENUE_LINKS = {};

let nvVenueLinksUnsub = null;



function listenNvVenueLinks(){

  if(nvVenueLinksUnsub) nvVenueLinksUnsub();

  nvVenueLinksUnsub = db.collection('planningNV').doc('venueLinks').onSnapshot(doc => {

    NV_VENUE_LINKS = (doc.exists && doc.data()) || {};

    refreshNvViews();

  }, err => console.error(err));

}



function getVenueMapUrl(venueName){

  return NV_VENUE_LINKS[venueName] || `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(venueName + ' Rabat')}`;

}



function setVenueLinkNv(venueName){

  const current = NV_VENUE_LINKS[venueName] || '';

  const value = prompt(`Lien Google Maps pour "${venueName}" (laisse vide pour revenir à la recherche automatique) :`, current);

  if(value === null) return;

  const updated = { ...NV_VENUE_LINKS };

  if(value.trim()) updated[venueName] = value.trim();

  else delete updated[venueName];

  db.collection('planningNV').doc('venueLinks').set(updated).catch(err => console.error(err));

}



function deleteHotspotNv(zone, slotIdx, venueIdx){

  if(!confirm('Supprimer ce lieu ?')) return;

  const zoneSlots = (HOTSPOTS_NV[zone] || []).map(s => ({ ...s, venues: [...s.venues] }));

  if(!zoneSlots[slotIdx]) return;

  zoneSlots[slotIdx].venues.splice(venueIdx, 1);

  const updated = { ...HOTSPOTS_NV, [zone]: zoneSlots };

  db.collection('planningNV').doc('hotspots').set(updated).catch(err => console.error(err));

}



function editVenueNameNv(zone, slotIdx, venueIdx){

  const zoneSlots = (HOTSPOTS_NV[zone] || []).map(s => ({ ...s, venues: [...s.venues] }));

  if(!zoneSlots[slotIdx] || !zoneSlots[slotIdx].venues[venueIdx]) return;

  const oldName = zoneSlots[slotIdx].venues[venueIdx];

  const value = prompt('Corriger le nom du lieu :', oldName);

  if(!value || !value.trim() || value.trim() === oldName) return;

  const newName = value.trim();

  zoneSlots[slotIdx].venues[venueIdx] = newName;

  const updated = { ...HOTSPOTS_NV, [zone]: zoneSlots };

  db.collection('planningNV').doc('hotspots').set(updated).catch(err => console.error(err));



  // Si un lien Maps personnalisé existait pour l'ancien nom, on le transfère au nouveau nom.

  if(NV_VENUE_LINKS[oldName]){

    const updatedLinks = { ...NV_VENUE_LINKS };

    updatedLinks[newName] = updatedLinks[oldName];

    delete updatedLinks[oldName];

    db.collection('planningNV').doc('venueLinks').set(updatedLinks).catch(err => console.error(err));

  }

}



function addHotspotSlotNv(zone){

  const value = prompt(`Nom du premier lieu pour ce nouveau hotspot à "${getNvZoneLabel(zone)}" :`);

  if(!value || !value.trim()) return;

  const zoneSlots = (HOTSPOTS_NV[zone] || []).map(s => ({ ...s, venues: [...s.venues] }));

  zoneSlots.push({ venues: [value.trim()] });

  const updated = { ...HOTSPOTS_NV, [zone]: zoneSlots };

  db.collection('planningNV').doc('hotspots').set(updated).catch(err => console.error(err));

}



function deleteHotspotSlotNv(zone, slotIdx){

  if(!confirm(`Supprimer le hotspot #${slotIdx + 1} en entier (tous ses lieux) ?`)) return;

  const zoneSlots = (HOTSPOTS_NV[zone] || []).map(s => ({ ...s, venues: [...s.venues] }));

  zoneSlots.splice(slotIdx, 1);

  const updated = { ...HOTSPOTS_NV, [zone]: zoneSlots };

  db.collection('planningNV').doc('hotspots').set(updated).catch(err => console.error(err));

}



function moveHotspotSlotNv(zone, slotIdx){

  const otherZones = Object.keys(HOTSPOTS_NV).filter(z => z !== zone);

  const choice = prompt(`Déplacer le hotspot #${slotIdx + 1} de "${getNvZoneLabel(zone)}" vers quelle zone ?\n\n${otherZones.map((z, i) => `${i + 1}. ${getNvZoneLabel(z)}`).join('\n')}\n\nTape le numéro :`);

  if(!choice) return;

  const targetZone = otherZones[parseInt(choice, 10) - 1];

  if(!targetZone){ alert('Choix invalide.'); return; }



  const sourceSlots = (HOTSPOTS_NV[zone] || []).map(s => ({ ...s, venues: [...s.venues] }));

  const movedSlot = sourceSlots[slotIdx];

  if(!movedSlot) return;

  sourceSlots.splice(slotIdx, 1);

  const targetSlots = (HOTSPOTS_NV[targetZone] || []).map(s => ({ ...s, venues: [...s.venues] }));

  targetSlots.push(movedSlot);



  const updated = { ...HOTSPOTS_NV, [zone]: sourceSlots, [targetZone]: targetSlots };

  db.collection('planningNV').doc('hotspots').set(updated).catch(err => console.error(err));

}



function renameHotspotSlotNv(zone, slotIdx){

  const zoneSlots = (HOTSPOTS_NV[zone] || []).map(s => ({ ...s, venues: [...s.venues] }));

  if(!zoneSlots[slotIdx]) return;

  const current = zoneSlots[slotIdx].label || `#${slotIdx + 1}`;

  const value = prompt(`Nom de ce hotspot :`, current);

  if(value === null) return;

  zoneSlots[slotIdx].label = value.trim() || null;

  const updated = { ...HOTSPOTS_NV, [zone]: zoneSlots };

  db.collection('planningNV').doc('hotspots').set(updated).catch(err => console.error(err));

}



// Noms d'affichage personnalisés pour les zones Planning NV (ne change pas la clé interne / la logique).

let NV_ZONE_LABELS = {};

let nvZoneLabelsUnsub = null;



function getNvZoneLabel(zone){

  return NV_ZONE_LABELS[zone] || zone;

}



function listenNvZoneLabels(){

  if(nvZoneLabelsUnsub) nvZoneLabelsUnsub();

  nvZoneLabelsUnsub = db.collection('planningNV').doc('zoneLabels').onSnapshot(doc => {

    NV_ZONE_LABELS = (doc.exists && doc.data()) || {};

    refreshNvViews();

  }, err => console.error(err));

}



// Ordre d'affichage personnalisé des agents dans Vue Semaine (persistant, indépendant de la semaine).

let NV_ROW_ORDER = [];

let nvRowOrderUnsub = null;



function listenNvRowOrder(){

  if(nvRowOrderUnsub) nvRowOrderUnsub();

  nvRowOrderUnsub = db.collection('planningNV').doc('rowOrder').onSnapshot(doc => {

    NV_ROW_ORDER = (doc.exists && doc.data() && doc.data().order) || [];

    refreshNvViews();

  }, err => console.error(err));

}



function moveNvAgentRow(name, direction, currentOrderNames){

  const order = NV_ROW_ORDER.length ? [...NV_ROW_ORDER] : [...currentOrderNames];

  // S'assure que tous les agents actuellement affichés sont présents dans l'ordre stocké.

  currentOrderNames.forEach(n => { if(!order.includes(n)) order.push(n); });

  const idx = order.indexOf(name);

  const swapIdx = idx + direction;

  if(idx < 0 || swapIdx < 0 || swapIdx >= order.length) return;

  [order[idx], order[swapIdx]] = [order[swapIdx], order[idx]];

  db.collection('planningNV').doc('rowOrder').set({ order }).catch(err => console.error(err));

}



function renameNvZone(zone){

  const value = prompt(`Nouveau nom affiché pour "${getNvZoneLabel(zone)}" :`, getNvZoneLabel(zone));

  if(!value || !value.trim()) return;

  const updated = { ...NV_ZONE_LABELS, [zone]: value.trim() };

  db.collection('planningNV').doc('zoneLabels').set(updated).catch(err => console.error(err));

}



const SHIFTS_NV = {

  S1: { name: "Matin", label7: "11:00-18:00", label8: "11:00-19:00" },

  S2: { name: "Après-midi", label7: "16:00-23:00", label8: "16:00-00:00" },

  S3: { name: "Nuit", label7: "19:00-02:00", label8: "19:00-03:00" },

};

const SHIFT_KEYS_NV = ['S1', 'S2', 'S3'];

const NV_DAY_LABELS = ['Lun', 'Mar', 'Mer', 'Jeu'];



// 4 groupes de 3 personnes — chacun tourne chaque semaine sur les 4 zones :

// Temara → Agdal → Salé-Rabat → Hay Riad → (retour à Temara), et ainsi de suite.

const ROTATING_TEAMS_NV = [

  ["Charaf Loummihi", "Mourad Cherradi", "Islam Elmardi"],

  ["Youssef Derkaoui", "Ayoub Ben Malek", "Aymane Chlifane"],

  ["Yahya Rhallah", "Oussama El Atrach", "Hassan Azgoune"],

  ["Ayoub Abssi", "Achraf Akhlij", "Mehdi Ait Jebour"],

];

const ROTATING_ZONES_NV = ["Temara", "Agdal", "Salé-Rabat", "Hay Riad"];



function updateWeekHeaderNv(){

  const weekIdx = getISOWeekIndex(currentWeekStartNv);

  document.getElementById('weekLabelNv').textContent = `Semaine ${weekIdx}`;

  document.getElementById('weekRangeNv').textContent = formatWeekRange(currentWeekStartNv);

}



document.getElementById('prevWeekNvBtn').addEventListener('click', () => {

  const d = new Date(currentWeekStartNv); d.setDate(d.getDate() - 7);

  currentWeekStartNv = d;

  listenNvOverrides();

});

document.getElementById('nextWeekNvBtn').addEventListener('click', () => {

  const d = new Date(currentWeekStartNv); d.setDate(d.getDate() + 7);

  currentWeekStartNv = d;

  listenNvOverrides();

});



function listenNvOverrides(){

  const weekIdx = getISOWeekIndex(currentWeekStartNv);

  if(nvOverridesUnsub) nvOverridesUnsub();

  nvOverridesUnsub = db.collection('planningNV').doc('overrides_W' + weekIdx).onSnapshot(doc => {

    NV_OVERRIDES = (doc.exists && doc.data()) || {};

    refreshNvViews();

  }, err => console.error(err));

}



function formatHourNv(h){

  return `${String(((h % 24) + 24) % 24).padStart(2, '0')}:00`;

}



// Agents avec un horaire fixe non-standard par défaut (coupure, mi-temps...), différent des 3 shifts

// classiques (Matin/Après-midi/Nuit). Reste modifiable via override normal (jour ou semaine) si besoin.

const NV_FIXED_SHIFTS_NV = {

  "Zakariae Benali": {

    name: "Coupure",

    label7: "11:00-15:00 · 19:00-22:00",

    label8: "11:00-15:00 · 19:00-23:00",

    isJour: true,

  },

};



function resolveNvShiftInfo(override, memberIdx, weekIdx, agentName){

  if(override && override.customShift){

    const c = override.customShift;

    return { name: c.name, label7: c.label7, label8: c.label8, isJour: c.isJour, shiftKey: null };

  }

  if(override && override.shiftKey){

    const s = SHIFTS_NV[override.shiftKey];

    return { name: s.name, label7: s.label7, label8: s.label8, isJour: SHIFT_KEYS_NV.indexOf(override.shiftKey) === 0, shiftKey: override.shiftKey };

  }

  if(agentName && NV_FIXED_SHIFTS_NV[agentName]){

    return { ...NV_FIXED_SHIFTS_NV[agentName], shiftKey: null };

  }

  const shiftKey = SHIFT_KEYS_NV[(memberIdx + weekIdx) % SHIFT_KEYS_NV.length];

  const s = SHIFTS_NV[shiftKey];

  return { name: s.name, label7: s.label7, label8: s.label8, isJour: SHIFT_KEYS_NV.indexOf(shiftKey) === 0, shiftKey };

}



// Rotation du hotspot en cycle fixe de 3 jours (hotspot 1 → 2 → 3 → 1 → 2 → 3...),

// même si la zone contient plus de 3 hotspots configurés (seuls les 3 premiers tournent).

function resolveNvHotspotIndex(slots, memberIdx, dayIdx, offDayIdx){

  const cycleLen = Math.min(3, slots.length);

  if(cycleLen === 0) return 0;

  // Le jour OFF ne doit pas "consommer" une place dans le cycle 1→2→3, sinon la rotation saute

  // un cran après le repos (ex: Lun=1, Mar=OFF, Mer=3 au lieu de 2). On compte uniquement les

  // jours réellement travaillés depuis le début de la semaine pour faire avancer le cycle.

  const workingDayIdx = (typeof offDayIdx === 'number' && dayIdx > offDayIdx) ? dayIdx - 1 : dayIdx;

  return (memberIdx + workingDayIdx) % cycleLen;

}



// Comme resolveNvHotspotIndex, mais respecte un éventuel override manuel du hotspot pour CE jour précis.

function resolveNvHotspotIndexForAgent(override, slots, memberIdx, dayIdx, offDayIdx){

  const forced = override && override.hotspotOverrides && override.hotspotOverrides[dayIdx];

  if(typeof forced === 'number' && slots[forced]) return forced;

  return resolveNvHotspotIndex(slots, memberIdx, dayIdx, offDayIdx);

}



function applyNvShiftChoice(agentName, weekIdx, shiftKey, dayIdx){

  const current = NV_OVERRIDES[agentName] || {};

  let updatedAgent;

  if(dayIdx === null || dayIdx === undefined){

    updatedAgent = { ...current, shiftKey, customShift: null };

  } else {

    const dayOverrides = { ...(current.dayOverrides || {}), [dayIdx]: { shiftKey } };

    updatedAgent = { ...current, dayOverrides };

  }

  const updated = { ...NV_OVERRIDES, [agentName]: updatedAgent };

  db.collection('planningNV').doc('overrides_W' + weekIdx).set(updated).catch(err => console.error(err));

}



function clearNvDayOverride(agentName, weekIdx, dayIdx){

  const current = NV_OVERRIDES[agentName] || {};

  const dayOverrides = { ...(current.dayOverrides || {}) };

  delete dayOverrides[dayIdx];

  const updated = { ...NV_OVERRIDES, [agentName]: { ...current, dayOverrides } };

  db.collection('planningNV').doc('overrides_W' + weekIdx).set(updated).catch(err => console.error(err));

}



function applyNvCustomHour(agentName, weekIdx, startH, startM, dayIdx){

  const startLabel = `${String(startH).padStart(2, '0')}:${String(startM).padStart(2, '0')}`;

  const endLabel7 = formatHourNv(startH + 7);

  const endLabel8 = formatHourNv(startH + 8);

  const customShift = {

    name: `Personnalisé (${startLabel})`,

    label7: `${startLabel}-${endLabel7}`,

    label8: `${startLabel}-${endLabel8}`,

    isJour: startH < 15,

  };

  const current = NV_OVERRIDES[agentName] || {};

  let updatedAgent;

  if(dayIdx === null || dayIdx === undefined){

    updatedAgent = { ...current, customShift, shiftKey: null };

  } else {

    const dayOverrides = { ...(current.dayOverrides || {}), [dayIdx]: { customShift } };

    updatedAgent = { ...current, dayOverrides };

  }

  const updated = { ...NV_OVERRIDES, [agentName]: updatedAgent };

  db.collection('planningNV').doc('overrides_W' + weekIdx).set(updated).catch(err => console.error(err));

}



function editNvShift(agentName, dayIdx){

  const weekIdx = getISOWeekIndex(currentWeekStartNv);

  const hourOptions = [];

  for(let h = 6; h <= 23; h++) hourOptions.push(h);

  const dayNames = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'];

  const isDaySpecific = dayIdx !== null && dayIdx !== undefined;

  const hasDayOverride = isDaySpecific && NV_OVERRIDES[agentName] && NV_OVERRIDES[agentName].dayOverrides && NV_OVERRIDES[agentName].dayOverrides[dayIdx];



  const overlay = document.createElement('div');

  overlay.className = 'modal-overlay';

  overlay.innerHTML = `

    <div class="modal-card" style="max-width:460px; max-height:85vh; display:flex; flex-direction:column;">

      <h3>⏱️ Shift de "${agentName}"${isDaySpecific ? ` — ${dayNames[dayIdx]} uniquement` : ''}</h3>

      <p class="modal-hint" style="margin-top:0;">${isDaySpecific ? `Ce changement ne s'applique qu'au ${dayNames[dayIdx]}, pas au reste de la semaine.` : 'Ce changement s\'applique par défaut à tous les jours de la semaine.'}</p>

      ${hasDayOverride ? `<button type="button" class="btn-ghost" id="nvShiftClearDayBtn" style="text-align:left; margin-bottom:10px; color:var(--red);">↩️ Revenir au shift par défaut pour ce jour</button>` : ''}

      <div style="display:flex; flex-direction:column; gap:8px; margin-bottom:14px;" id="nvShiftPresets">

        <button type="button" class="btn-ghost" data-preset="S1" style="text-align:left;">${SHIFTS_NV.S1.name} — ${SHIFTS_NV.S1.label7}</button>

        <button type="button" class="btn-ghost" data-preset="S2" style="text-align:left;">${SHIFTS_NV.S2.name} — ${SHIFTS_NV.S2.label7}</button>

        <button type="button" class="btn-ghost" data-preset="S3" style="text-align:left;">${SHIFTS_NV.S3.name} — ${SHIFTS_NV.S3.label7}</button>

      </div>

      <p class="modal-hint" style="margin-bottom:8px;">Ou choisis une heure de début personnalisée :</p>

      <div style="display:grid; grid-template-columns:repeat(4, 1fr); gap:6px; overflow-y:auto; padding-bottom:4px;" id="nvShiftHourGrid">

        ${hourOptions.map(h => `<button type="button" class="btn-ghost" data-hour="${h}" style="padding:8px 4px; font-family:var(--font-mono); font-size:13px;">${String(h).padStart(2,'0')}:00</button>`).join('')}

      </div>

      <p class="modal-hint" style="margin:12px 0 6px;">Ou tape une heure exacte (ex: 14:30) :</p>

      <div style="display:flex; gap:8px;">

        <input type="text" id="nvShiftCustomInput" placeholder="HH:MM" class="login-input" style="flex:1; padding:8px 10px; font-family:var(--font-mono);" />

        <button type="button" class="btn-primary" id="nvShiftCustomApply" style="padding:8px 14px;">OK</button>

      </div>

      <div class="modal-error" id="nvShiftCustomError"></div>

      <div class="modal-actions"><button class="btn-ghost" id="nvShiftCloseBtn">Annuler</button></div>

    </div>

  `;

  document.body.appendChild(overlay);

  const close = () => overlay.remove();

  overlay.addEventListener('click', e => { if(e.target === overlay) close(); });

  overlay.querySelector('#nvShiftCloseBtn').addEventListener('click', close);

  if(hasDayOverride){

    overlay.querySelector('#nvShiftClearDayBtn').addEventListener('click', () => { clearNvDayOverride(agentName, weekIdx, dayIdx); close(); });

  }

  overlay.querySelectorAll('[data-preset]').forEach(btn => {

    btn.addEventListener('click', () => { applyNvShiftChoice(agentName, weekIdx, btn.dataset.preset, dayIdx); close(); });

  });

  overlay.querySelectorAll('[data-hour]').forEach(btn => {

    btn.addEventListener('click', () => { applyNvCustomHour(agentName, weekIdx, parseInt(btn.dataset.hour, 10), 0, dayIdx); close(); });

  });

  overlay.querySelector('#nvShiftCustomApply').addEventListener('click', () => {

    const raw = overlay.querySelector('#nvShiftCustomInput').value.trim();

    const match = raw.match(/^(\d{1,2})[:h](\d{2})$/i);

    const errorEl = overlay.querySelector('#nvShiftCustomError');

    if(!match){ errorEl.textContent = 'Format invalide. Exemple : 14:30'; return; }

    const startH = parseInt(match[1], 10);

    const startM = parseInt(match[2], 10);

    if(startH > 23 || startM > 59){ errorEl.textContent = 'Heure invalide.'; return; }

    applyNvCustomHour(agentName, weekIdx, startH, startM, dayIdx);

    close();

  });

}



function applyNvHotspotChoice(agentName, weekIdx, dayIdx, slotIdx){

  const current = NV_OVERRIDES[agentName] || {};

  const hotspotOverrides = { ...(current.hotspotOverrides || {}), [dayIdx]: slotIdx };

  const updated = { ...NV_OVERRIDES, [agentName]: { ...current, hotspotOverrides } };

  db.collection('planningNV').doc('overrides_W' + weekIdx).set(updated).catch(err => console.error(err));

}



function clearNvHotspotOverride(agentName, weekIdx, dayIdx){

  const current = NV_OVERRIDES[agentName] || {};

  const hotspotOverrides = { ...(current.hotspotOverrides || {}) };

  delete hotspotOverrides[dayIdx];

  const updated = { ...NV_OVERRIDES, [agentName]: { ...current, hotspotOverrides } };

  db.collection('planningNV').doc('overrides_W' + weekIdx).set(updated).catch(err => console.error(err));

}



function editNvHotspotForDay(agentName, dayIdx, zone){

  const weekIdx = getISOWeekIndex(currentWeekStartNv);

  const dayNames = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'];

  const slots = HOTSPOTS_NV[zone] || [];

  const current = NV_OVERRIDES[agentName] || {};

  const hasOverride = current.hotspotOverrides && typeof current.hotspotOverrides[dayIdx] === 'number';



  const overlay = document.createElement('div');

  overlay.className = 'modal-overlay';

  overlay.innerHTML = `

    <div class="modal-card" style="max-width:440px; max-height:85vh; display:flex; flex-direction:column; overflow-y:auto;">

      <h3>📍 Hotspot de "${agentName}" — ${dayNames[dayIdx]}</h3>

      <p class="modal-hint" style="margin-top:0;">Ce choix ne s'applique qu'au ${dayNames[dayIdx]}, pas au reste de la semaine.</p>

      ${hasOverride ? `<button type="button" class="btn-ghost" id="nvHotspotClearBtn" style="text-align:left; margin-bottom:10px; color:var(--red);">↩️ Revenir à la rotation automatique pour ce jour</button>` : ''}

      <div style="display:flex; flex-direction:column; gap:8px;">

        ${slots.map((s, i) => {

          const label = s.label || `Hotspot ${i + 1}`;

          const venues = (s.venues || []).join(', ');

          return `<button type="button" class="btn-ghost" data-slot="${i}" style="text-align:left;"><b>${label}</b><br><span style="font-size:11.5px; color:var(--ink-soft);">${venues}</span></button>`;

        }).join('')}

      </div>

      <div class="modal-actions"><button class="btn-ghost" id="nvHotspotCloseBtn">Annuler</button></div>

    </div>

  `;

  document.body.appendChild(overlay);

  const close = () => overlay.remove();

  overlay.addEventListener('click', e => { if(e.target === overlay) close(); });

  overlay.querySelector('#nvHotspotCloseBtn').addEventListener('click', close);

  if(hasOverride){

    overlay.querySelector('#nvHotspotClearBtn').addEventListener('click', () => { clearNvHotspotOverride(agentName, weekIdx, dayIdx); close(); });

  }

  overlay.querySelectorAll('[data-slot]').forEach(btn => {

    btn.addEventListener('click', () => { applyNvHotspotChoice(agentName, weekIdx, dayIdx, parseInt(btn.dataset.slot, 10)); close(); });

  });

}



function applyNvOffDay(agentName, weekIdx, offDayIdx){

  const updated = { ...NV_OVERRIDES, [agentName]: { ...(NV_OVERRIDES[agentName] || {}), offDayIdx } };

  db.collection('planningNV').doc('overrides_W' + weekIdx).set(updated).catch(err => console.error(err));

}



function editNvOffDay(agentName){

  const weekIdx = getISOWeekIndex(currentWeekStartNv);

  const dayNames = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi'];

  const overlay = document.createElement('div');

  overlay.className = 'modal-overlay';

  overlay.innerHTML = `

    <div class="modal-card" style="max-width:400px;">

      <h3>🛌 Jour OFF de "${agentName}"</h3>

      <p class="modal-hint" style="margin-top:0;">Uniquement Lundi à Jeudi.</p>

      <div style="display:flex; flex-direction:column; gap:8px; margin-bottom:14px;">

        ${dayNames.map((d, i) => `<button type="button" class="btn-ghost" data-off-idx="${i}" style="text-align:left;">${d}</button>`).join('')}

      </div>

      <div class="modal-actions"><button class="btn-ghost" id="nvOffCloseBtn">Annuler</button></div>

    </div>

  `;

  document.body.appendChild(overlay);

  const close = () => overlay.remove();

  overlay.addEventListener('click', e => { if(e.target === overlay) close(); });

  overlay.querySelector('#nvOffCloseBtn').addEventListener('click', close);

  overlay.querySelectorAll('[data-off-idx]').forEach(btn => {

    btn.addEventListener('click', () => { applyNvOffDay(agentName, weekIdx, parseInt(btn.dataset.offIdx, 10)); close(); });

  });

}



function applyNvZoneChoice(agentName, weekIdx, zoneOverride){

  const updated = { ...NV_OVERRIDES, [agentName]: { ...(NV_OVERRIDES[agentName] || {}), zoneOverride } };

  db.collection('planningNV').doc('overrides_W' + weekIdx).set(updated).catch(err => console.error(err));

}



function editNvZone(agentName){

  const weekIdx = getISOWeekIndex(currentWeekStartNv);

  const zoneNames = Object.keys(NV_ZONE_COLORS);

  const overlay = document.createElement('div');

  overlay.className = 'modal-overlay';

  overlay.innerHTML = `

    <div class="modal-card" style="max-width:400px;">

      <h3>📍 Zone de "${agentName}"</h3>

      <p class="modal-hint" style="margin-top:0;">Choisis une zone pour cette semaine, ou reviens à la rotation automatique.</p>

      <div style="display:flex; flex-direction:column; gap:8px; margin-bottom:14px;">

        ${zoneNames.map(z => `<button type="button" class="btn-ghost" data-zone="${z}" style="text-align:left;">${getNvZoneLabel(z)}</button>`).join('')}

        <button type="button" class="btn-ghost" data-zone="" style="text-align:left; color:var(--ink-soft);">↩️ Rotation automatique</button>

      </div>

      <div class="modal-actions"><button class="btn-ghost" id="nvZoneCloseBtn">Annuler</button></div>

    </div>

  `;

  document.body.appendChild(overlay);

  const close = () => overlay.remove();

  overlay.addEventListener('click', e => { if(e.target === overlay) close(); });

  overlay.querySelector('#nvZoneCloseBtn').addEventListener('click', close);

  overlay.querySelectorAll('[data-zone]').forEach(btn => {

    btn.addEventListener('click', () => { applyNvZoneChoice(agentName, weekIdx, btn.dataset.zone || null); close(); });

  });

}



const NV_FIXED_ZONE_AGENTS = {

  "Technopolis": ["Zakariae Benali"],

};



function buildZoneAssignmentsNv(weekIdx){

  const zones = { "Salé-Rabat": [], "Agdal": [], "Hay Riad": [], "Temara": [], "Technopolis": [] };

  ROTATING_TEAMS_NV.forEach((team, teamIdx) => {

    const zoneIdx = ((teamIdx + weekIdx) % 4 + 4) % 4;

    const zone = ROTATING_ZONES_NV[zoneIdx];

    zones[zone].push(...team);

  });

  // Agents à zone fixe (ex: Zakariae Benali à Technopolis) — ne tournent pas avec les 4 groupes.

  Object.entries(NV_FIXED_ZONE_AGENTS).forEach(([zone, names]) => {

    if(zones[zone]) zones[zone].push(...names);

  });

  // Overrides manuels de zone pour la semaine en cours : on déplace l'agent de sa zone

  // calculée automatiquement vers la zone choisie par l'admin.

  Object.entries(NV_OVERRIDES).forEach(([name, ov]) => {

    if(!ov || !ov.zoneOverride || !zones[ov.zoneOverride]) return;

    Object.keys(zones).forEach(z => {

      const idx = zones[z].indexOf(name);

      if(idx >= 0) zones[z].splice(idx, 1);

    });

    zones[ov.zoneOverride].push(name);

  });

  return zones;

}



const NV_ZONE_COLORS = {

  "Salé-Rabat": "#FB7185",

  "Agdal": "#FBBF24",

  "Hay Riad": "#A78BFA",

  "Temara": "#38BDF8",

  "Technopolis": "#34D399",

};

const NV_ZONE_GRADIENTS = {

  "Salé-Rabat": "linear-gradient(135deg,#FB7185,#E11D48)",

  "Agdal": "linear-gradient(135deg,#FBBF24,#D97706)",

  "Hay Riad": "linear-gradient(135deg,#A78BFA,#7C3AED)",

  "Temara": "linear-gradient(135deg,#38BDF8,#0284C7)",

  "Technopolis": "linear-gradient(135deg,#34D399,#059669)",

};



function renderPlanningNV(){

  try{

    renderPlanningNVInner();

  } catch(err){

    console.error('renderPlanningNV error:', err);

    document.getElementById('nvZoneGrid').innerHTML = '<p style="color:var(--red); padding:20px;">Erreur d\'affichage — réessaie de recharger la page.</p>';

  }

}



// À appeler après toute mise à jour de données (hotspots, overrides, liens...) : rafraîchit

// automatiquement la vue admin ET le portail agent si l'un ou l'autre est actuellement affiché,

// pour que les changements soient visibles immédiatement des deux côtés sans recharger la page.

function refreshNvViews(){

  renderPlanningNV();

  const portalVisible = !document.getElementById('nvAgentPortal').classList.contains('hidden');

  if(portalVisible && currentNvAgent){

    renderNvPortal();

  }

}



let nvPlanningView = 'week';



function renderPlanningNVInner(){

  updateWeekHeaderNv();

  document.getElementById('nvWeekGrid').classList.toggle('hidden', nvPlanningView !== 'week');

  document.getElementById('nvZoneGrid').classList.toggle('hidden', nvPlanningView !== 'zones');

  document.getElementById('nvZoneRestoreWrap').classList.toggle('hidden', nvPlanningView !== 'zones');

  if(nvPlanningView === 'week'){

    renderNvWeekGridView();

  } else {

    renderNvZonesView();

  }

}



document.querySelectorAll('#view-planningnv .planning-tabs [data-nv-view]').forEach(btn => {

  btn.addEventListener('click', () => {

    nvPlanningView = btn.dataset.nvView;

    document.querySelectorAll('#view-planningnv .planning-tabs [data-nv-view]').forEach(b => b.classList.toggle('active', b === btn));

    renderPlanningNVInner();

  });

});



document.getElementById('nvExportHotspotsBtn').addEventListener('click', () => {

  let text = 'HOTSPOTS FieldOps — sauvegarde du ' + new Date().toLocaleString('fr-FR') + '\n\n';

  Object.keys(HOTSPOTS_NV).forEach(zone => {

    text += `=== ${zone} ===\n`;

    (HOTSPOTS_NV[zone] || []).forEach((slot, i) => {

      text += `Hotspot ${i + 1}${slot.label ? ' (' + slot.label + ')' : ''} : ${(slot.venues || []).join(', ')}\n`;

    });

    text += '\n';

  });

  if(navigator.clipboard && navigator.clipboard.writeText){

    navigator.clipboard.writeText(text).then(() => {

      alert('Copié ! Colle ce texte dans une note, un WhatsApp à toi-même, ou un fichier — garde-le comme sauvegarde. Fais ça de temps en temps, surtout après avoir ajouté des lieux.');

    }).catch(() => {

      prompt('Copie ce texte manuellement (Ctrl+A puis Ctrl+C) :', text);

    });

  } else {

    prompt('Copie ce texte manuellement (Ctrl+A puis Ctrl+C) :', text);

  }

});



document.getElementById('nvRestoreHotspotsBtn').addEventListener('click', () => {

  if(!confirm(`Ceci va remplacer les hotspots de TOUTES les zones par les valeurs par défaut du code (Agdal et Hay Riad reconstruits à partir de tes captures d'écran ; Salé-Rabat, Temara et Technopolis restent tels quels dans le code).\n\n⚠️ Tout ajout fait depuis Vue Zones qui ne serait pas dans ces valeurs par défaut sera perdu.\n\nContinuer ?`)) return;

  db.collection('planningNV').doc('hotspots').set(DEFAULT_HOTSPOTS_NV)

    .then(() => alert('Hotspots restaurés.'))

    .catch(err => { console.error(err); alert('Erreur pendant la restauration.'); });

});



function renderNvWeekGridView(){

  const weekIdx = getISOWeekIndex(currentWeekStartNv);

  const zones = buildZoneAssignmentsNv(weekIdx);

  const container = document.getElementById('nvWeekGrid');

  const ALL_DAY_LABELS_NV = ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'];

  const todayIdx = (new Date().getDay() + 6) % 7;



  const rows = [];

  Object.keys(HOTSPOTS_NV).forEach(zone => {

    const members = zones[zone];

    const slots = HOTSPOTS_NV[zone] || [];

    const color = NV_ZONE_COLORS[zone];



    members.forEach((name, memberIdx) => {

      const override = NV_OVERRIDES[name] || {};

      const shiftInfo = resolveNvShiftInfo(override, memberIdx, weekIdx, name);

      const offDayIdx = (typeof override.offDayIdx === 'number') ? override.offDayIdx : (memberIdx + weekIdx) % 4;

      const isJour = shiftInfo.isJour;

      const initials = name.split(' ').map(w => w[0]).slice(0, 2).join('').toUpperCase();



      // Total heures cette semaine : 8h Ven/Sam, 7h les autres jours travaillés, 0h le jour OFF.

      let totalHeures = 0;

      ALL_DAY_LABELS_NV.forEach((_, dayIdx) => {

        if(dayIdx === offDayIdx) return;

        const isWeekendShort = dayIdx >= 4 && dayIdx < 6; // Ven, Sam

        totalHeures += isWeekendShort ? 8 : 7;

      });

      const hoursOk = totalHeures === 44;



      const dayCells = ALL_DAY_LABELS_NV.map((dayLabel, dayIdx) => {

        const isToday = dayIdx === todayIdx;

        if(dayIdx === offDayIdx){

          return `<td style="${isToday ? 'background:rgba(0,160,130,0.06);' : ''}"><span class="cell-pill" style="background:var(--sand-deep); color:var(--ink-soft); display:block; font-size:10.5px;">OFF</span></td>`;

        }

        let hotspotLine = '';

        if(slots.length){

          const hotspotOverride = override.hotspotOverrides && override.hotspotOverrides[dayIdx];

          const slotIdx = (typeof hotspotOverride === 'number' && slots[hotspotOverride]) ? hotspotOverride : resolveNvHotspotIndex(slots, memberIdx, dayIdx, offDayIdx);

          const venues = (slots[slotIdx] && slots[slotIdx].venues) || [];

          const label = venues.length ? venues[0] + (venues.length > 1 ? ` +${venues.length - 1}` : '') : '—';

          hotspotLine = `<div data-cell-edit-hotspot="${name}" data-cell-edit-hotspot-day="${dayIdx}" data-cell-edit-hotspot-zone="${zone}" title="${typeof hotspotOverride === 'number' ? 'Hotspot modifié pour ce jour — cliquer pour changer' : 'Cliquer pour changer le hotspot de ce jour'}" style="cursor:pointer; font-size:9px; color:${typeof hotspotOverride === 'number' ? 'var(--sky-ink)' : 'var(--ink-soft)'}; font-weight:${typeof hotspotOverride === 'number' ? '700' : '400'}; margin-top:2px; max-width:78px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; margin-left:auto; margin-right:auto;">${typeof hotspotOverride === 'number' ? '✏️ ' : '📍 '}${label}</div>`;

        }

        // Un shift modifié pour CE jour précis prime sur le shift par défaut de la semaine.

        const dayOverride = (override.dayOverrides && override.dayOverrides[dayIdx]) || null;

        const dayShiftInfo = dayOverride ? resolveNvShiftInfo(dayOverride, memberIdx, weekIdx) : shiftInfo;

        const dayIsJour = dayShiftInfo.isJour;

        const isWeekendShortDay = dayIdx >= 4 && dayIdx < 6; // Ven, Sam

        const cellTimeLabel = isWeekendShortDay ? dayShiftInfo.label8 : dayShiftInfo.label7;

        return `<td style="${isToday ? 'background:rgba(0,160,130,0.06);' : ''}">

          <span class="cell-pill ${dayIsJour ? 'jour' : 'nuit'}" data-cell-edit-shift="${name}" data-cell-edit-day="${dayIdx}" title="${dayOverride ? 'Heure modifiée pour ce jour — cliquer pour changer' : 'Cliquer pour changer l\'heure de ce jour'}" style="display:block; cursor:pointer; position:relative; background:${dayIsJour ? 'rgba(4,120,87,0.14)' : 'rgba(184,128,14,0.16)'}; color:${dayIsJour ? '#047857' : '#8A6208'}; border:1.5px solid ${dayOverride ? 'var(--sky)' : (dayIsJour ? 'rgba(4,120,87,0.35)' : 'rgba(184,128,14,0.4)')}; font-size:12px; font-weight:800; font-family:var(--font-mono); letter-spacing:-0.02em; padding:3px 3px; border-radius:6px;">${dayOverride ? '✏️ ' : (dayIsJour ? '☀️ ' : '🌙 ')}${cellTimeLabel}</span>

          ${hotspotLine}

        </td>`;

      }).join('');



      const rowHtml = `

        <tr>

          <td class="sticky-col agent-name-cell">

            <div style="display:flex; align-items:center; gap:8px;">

              <div style="display:flex; flex-direction:column; gap:1px;">

                <button type="button" class="nv-mini-link" data-move-up="${name}" title="Monter" style="padding:0; line-height:1; font-size:11px;">▲</button>

                <button type="button" class="nv-mini-link" data-move-down="${name}" title="Descendre" style="padding:0; line-height:1; font-size:11px;">▼</button>

              </div>

              <span class="agent-row-avatar" style="display:flex; align-items:center; gap:8px;">

                <span class="avatar" style="width:26px; height:26px; border-radius:50%; background:${color}; color:#fff; display:inline-flex; align-items:center; justify-content:center; font-size:10px; font-weight:800; flex-shrink:0;">${initials}</span>

                ${name}

              </span>

            </div>

          </td>

          <td class="zone-cell">

            ${getNvZoneLabel(zone)}${override.zoneOverride ? ' <span title="Zone modifiée manuellement">✏️</span>' : ''}

            <div style="display:flex; gap:4px; margin-top:3px; flex-wrap:wrap;">

              <button class="nv-mini-link" data-grid-edit-zone="${name}" style="padding:0; font-size:9px;">Zone</button>

              <button class="nv-mini-link" data-grid-edit-shift="${name}" style="padding:0; font-size:9px;">Shift</button>

              <button class="nv-mini-link" data-grid-edit-off="${name}" style="padding:0; font-size:9px;">OFF</button>

            </div>

          </td>

          ${dayCells}

          <td style="font-weight:800; font-family:var(--font-mono); color:${hoursOk ? 'var(--green)' : 'var(--red)'};" title="${hoursOk ? 'Conforme aux 44h/semaine' : 'Ne totalise pas 44h/semaine !'}">

            ${totalHeures}h${!hoursOk ? ' ⚠️' : ''}

          </td>

        </tr>

      `;

      rows.push({ name, html: rowHtml });

    });

  });



  // Applique l'ordre d'affichage personnalisé (persistant), agents non classés à la fin dans l'ordre naturel.

  const allNames = rows.map(r => r.name);

  const orderedNames = [

    ...NV_ROW_ORDER.filter(n => allNames.includes(n)),

    ...allNames.filter(n => !NV_ROW_ORDER.includes(n)),

  ];

  const rowsByName = Object.fromEntries(rows.map(r => [r.name, r.html]));

  const bodyRows = orderedNames.map(n => rowsByName[n]).join('');



  container.innerHTML = `

    <table class="planning-table">

      <thead><tr><th class="sticky-col">Agent</th><th>Zone</th>${ALL_DAY_LABELS_NV.map((d, i) => `<th style="${i === todayIdx ? 'color:var(--sky-ink);' : ''}">${d}</th>`).join('')}<th>Total</th></tr></thead>

      <tbody>${bodyRows}</tbody>

    </table>

  `;



  container.querySelectorAll('[data-grid-edit-zone]').forEach(btn => {

    btn.addEventListener('click', () => { editNvZone(btn.dataset.gridEditZone); });

  });

  container.querySelectorAll('[data-grid-edit-shift]').forEach(btn => {

    btn.addEventListener('click', () => { editNvShift(btn.dataset.gridEditShift); });

  });

  container.querySelectorAll('[data-cell-edit-shift]').forEach(el => {

    el.addEventListener('click', () => { editNvShift(el.dataset.cellEditShift, parseInt(el.dataset.cellEditDay, 10)); });

  });

  container.querySelectorAll('[data-cell-edit-hotspot]').forEach(el => {

    el.addEventListener('click', () => { editNvHotspotForDay(el.dataset.cellEditHotspot, parseInt(el.dataset.cellEditHotspotDay, 10), el.dataset.cellEditHotspotZone); });

  });

  container.querySelectorAll('[data-grid-edit-off]').forEach(btn => {

    btn.addEventListener('click', () => { editNvOffDay(btn.dataset.gridEditOff); });

  });

  container.querySelectorAll('[data-move-up]').forEach(btn => {

    btn.addEventListener('click', () => { moveNvAgentRow(btn.dataset.moveUp, -1, orderedNames); });

  });

  container.querySelectorAll('[data-move-down]').forEach(btn => {

    btn.addEventListener('click', () => { moveNvAgentRow(btn.dataset.moveDown, 1, orderedNames); });

  });

}



function renderNvZonesView(){

  const weekIdx = getISOWeekIndex(currentWeekStartNv);

  const zones = buildZoneAssignmentsNv(weekIdx);

  const grid = document.getElementById('nvZoneGrid');

  grid.innerHTML = '';

  const ALL_DAY_LABELS_NV = ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'];



  Object.keys(HOTSPOTS_NV).forEach(zone => {

    const members = zones[zone];

    const slots = HOTSPOTS_NV[zone] || [];

    const color = NV_ZONE_COLORS[zone];

    const gradient = NV_ZONE_GRADIENTS[zone] || color;

    const todayIdx = (new Date().getDay() + 6) % 7; // 0=Lun..6=Dim



    const card = document.createElement('div');

    card.className = 'nv-zone-card';

    card.style.borderLeft = `4px solid ${color}`;



    let membersHtml = '';

    members.forEach((name, memberIdx) => {

      const override = NV_OVERRIDES[name] || {};

      const shiftInfo = resolveNvShiftInfo(override, memberIdx, weekIdx, name);

      const offDayIdx = (typeof override.offDayIdx === 'number') ? override.offDayIdx : (memberIdx + weekIdx) % 4; // 0..3 = Lun..Jeu

      const shiftLabel = `${shiftInfo.name} · ${shiftInfo.label7} / ${shiftInfo.label8} (Ven-Sam)`;

      const isJour = shiftInfo.isJour;



      // Hotspot du jour (aujourd'hui) mis en avant — même calcul que le détail jour par jour ci-dessous.

      let heroHtml;

      if(todayIdx === offDayIdx){

        heroHtml = `<div class="nv-today-hero nv-today-hero-off">😴 OFF aujourd'hui</div>`;

      } else if(slots.length){

        const todaySlotIdx = resolveNvHotspotIndexForAgent(override, slots, memberIdx, todayIdx, offDayIdx);

        const todayVenues = (slots[todaySlotIdx] && slots[todaySlotIdx].venues) || [];

        const todaySlotLabel = (slots[todaySlotIdx] && slots[todaySlotIdx].label) || `Hotspot ${todaySlotIdx + 1}`;

        const todayVenueHtml = todayVenues.length

          ? todayVenues.map(v => `<a href="${getVenueMapUrl(v)}" target="_blank" rel="noopener">📍 ${v}</a>`).join('<span style="opacity:.5;">·</span>')

          : '<span style="opacity:.7;">Aucun lieu défini</span>';

        heroHtml = `<div class="nv-today-hero" style="background:${gradient};"><span class="nv-today-hero-label">Aujourd'hui · ${todaySlotLabel}</span><span class="nv-today-hero-venues">${todayVenueHtml}</span></div>`;

      } else {

        heroHtml = `<div class="nv-today-hero nv-today-hero-off">Aucun hotspot configuré</div>`;

      }



      let dayChips = '';

      ALL_DAY_LABELS_NV.forEach((dayLabel, dayIdx) => {

        if(dayIdx === offDayIdx){

          dayChips += `<div class="nv-day-block"><div class="nv-day-label">${dayLabel}</div><span class="nv-day-chip off">OFF</span></div>`;

        } else if(slots.length){

          // Le hotspot ENTIER (zone/slot) tourne chaque jour entre les membres de l'équipe :

          // aujourd'hui c'est Charaf qui y est, demain ce sera Islam, etc.

          const slotIdx = resolveNvHotspotIndexForAgent(override, slots, memberIdx, dayIdx, offDayIdx);

          const venues = (slots[slotIdx] && slots[slotIdx].venues) || [];

          const venueChips = venues.length

            ? venues.map(v => `<a class="nv-day-chip nv-day-chip-link" href="${getVenueMapUrl(v)}" target="_blank" rel="noopener">📍 ${v}</a>`).join('')

            : `<span class="nv-day-chip">—</span>`;

          const slotLabel = (slots[slotIdx] && slots[slotIdx].label) || `Hotspot ${slotIdx + 1}`;

          dayChips += `<div class="nv-day-block"><div class="nv-day-label">${dayLabel} · ${slotLabel}</div><div class="nv-day-venues">${venueChips}</div></div>`;

        } else {

          dayChips += `<div class="nv-day-block"><div class="nv-day-label">${dayLabel}</div><span class="nv-day-chip">—</span></div>`;

        }

      });



      membersHtml += `

        <div class="nv-member-row" style="flex-direction:column; align-items:stretch;">

          <div style="display:flex; align-items:center; justify-content:space-between; gap:10px; margin-bottom:8px;">

            <div class="nv-member-name">${name}</div>

            <div style="display:flex; align-items:center; gap:6px;">

              <button class="nv-chip ${isJour ? 'shift-jour' : 'shift-nuit'}" data-edit-shift="${name}" style="border:none; cursor:pointer;" title="Modifier le shift">${shiftLabel} ✏️</button>

            </div>

          </div>

          ${heroHtml}

          <div style="text-align:right; margin:8px 0 6px;">

            <button class="nv-mini-link" data-edit-off="${name}">✏️ Changer le jour OFF (actuellement ${NV_DAY_LABELS[offDayIdx]})</button>

          </div>

          <div class="nv-day-strip">${dayChips}</div>

        </div>

      `;

    });



    let fullListHtml = slots.map((slot, i) => {

      const venuesHtml = (slot.venues || []).map((v, vi) => `

        <span class="nv-venue-tag"><a href="${getVenueMapUrl(v)}" target="_blank" rel="noopener">📍 ${v}</a> <button data-edit-venue-name="${zone}::${i}::${vi}" title="Modifier le nom">✏️</button> <button data-edit-link="${v.replace(/"/g, '&quot;')}" title="Lien Maps personnalisé">🔗</button> <button data-del-venue="${zone}::${i}::${vi}" title="Supprimer">✕</button></span>

      `).join('');

      const slotLabel = slot.label || `#${i + 1}`;

      return `

        <div class="nv-fulllist-row">

          <div><b>${slotLabel}</b> <button class="nv-mini-link" data-rename-slot="${zone}::${i}" style="padding:0; font-size:10px;">✏️</button> ${venuesHtml}</div>

          <div style="margin-top:4px;">

            <button class="nv-mini-link" data-add-venue="${zone}::${i}">+ lieu</button>

            <button class="nv-mini-link" data-move-slot="${zone}::${i}">➡️ déplacer vers une autre zone</button>

            <button class="nv-mini-link nv-mini-danger" data-del-slot="${zone}::${i}">✕ supprimer hotspot ${slotLabel}</button>

          </div>

        </div>

      `;

    }).join('');



    card.innerHTML = `

      <div class="nv-zone-head" style="background:${gradient};">

        <div style="display:flex; align-items:center; gap:8px;">

          <h3>${getNvZoneLabel(zone)}</h3>

          <button class="icon-btn" data-rename-zone="${zone}" title="Renommer" style="background:rgba(255,255,255,0.22); color:#fff; width:26px; height:26px;">✏️</button>

        </div>

        <span>${members.length} agent${members.length > 1 ? 's' : ''}</span>

      </div>

      ${membersHtml}

      <div class="nv-fulllist">

        <div class="nv-fulllist-title">Tous les hotspots de la zone</div>

        ${fullListHtml}

        <button class="btn-ghost" data-add-slot="${zone}" style="border:1px solid var(--line); margin-top:8px; width:100%; font-size:12px; padding:8px;">+ Nouveau hotspot</button>

      </div>

      </div>

    `;

    grid.appendChild(card);

  });



  grid.querySelectorAll('[data-rename-zone]').forEach(btn => {

    btn.addEventListener('click', () => renameNvZone(btn.dataset.renameZone));

  });

  grid.querySelectorAll('[data-add-slot]').forEach(btn => {

    btn.addEventListener('click', () => addHotspotSlotNv(btn.dataset.addSlot));

  });

  grid.querySelectorAll('[data-del-slot]').forEach(btn => {

    btn.addEventListener('click', () => {

      const [zone, slotIdx] = btn.dataset.delSlot.split('::');

      deleteHotspotSlotNv(zone, parseInt(slotIdx, 10));

    });

  });

  grid.querySelectorAll('[data-move-slot]').forEach(btn => {

    btn.addEventListener('click', () => {

      const [zone, slotIdx] = btn.dataset.moveSlot.split('::');

      moveHotspotSlotNv(zone, parseInt(slotIdx, 10));

    });

  });

  grid.querySelectorAll('[data-rename-slot]').forEach(btn => {

    btn.addEventListener('click', () => {

      const [zone, slotIdx] = btn.dataset.renameSlot.split('::');

      renameHotspotSlotNv(zone, parseInt(slotIdx, 10));

    });

  });

  grid.querySelectorAll('[data-add-venue]').forEach(btn => {

    btn.addEventListener('click', () => {

      const [zone, slotIdx] = btn.dataset.addVenue.split('::');

      addHotspotNv(zone, parseInt(slotIdx, 10));

    });

  });

  grid.querySelectorAll('[data-del-venue]').forEach(btn => {

    btn.addEventListener('click', () => {

      const [zone, slotIdx, venueIdx] = btn.dataset.delVenue.split('::');

      deleteHotspotNv(zone, parseInt(slotIdx, 10), parseInt(venueIdx, 10));

    });

  });

  grid.querySelectorAll('[data-edit-link]').forEach(btn => {

    btn.addEventListener('click', () => setVenueLinkNv(btn.dataset.editLink));

  });

  grid.querySelectorAll('[data-edit-venue-name]').forEach(btn => {

    btn.addEventListener('click', () => {

      const [zone, slotIdx, venueIdx] = btn.dataset.editVenueName.split('::');

      editVenueNameNv(zone, parseInt(slotIdx, 10), parseInt(venueIdx, 10));

    });

  });

  grid.querySelectorAll('[data-edit-shift]').forEach(btn => {

    btn.addEventListener('click', () => editNvShift(btn.dataset.editShift));

  });

  grid.querySelectorAll('[data-edit-off]').forEach(btn => {

    btn.addEventListener('click', () => editNvOffDay(btn.dataset.editOff));

  });

  grid.querySelectorAll('[data-view-route]').forEach(btn => {

    btn.addEventListener('click', () => viewNvAgentRoute(btn.dataset.viewRoute));

  });

}



function viewNvAgentRoute(agentName){

  const dateKey = dateKeyFor(new Date());

  const overlay = document.createElement('div');

  overlay.className = 'modal-overlay';

  overlay.innerHTML = `

    <div class="modal-card" style="max-width:440px; max-height:80vh; display:flex; flex-direction:column;">

      <h3>🧭 Trajet du jour — ${agentName}</h3>

      <div style="font-size:12px; color:var(--ink-soft); margin-bottom:10px;">${dateKey}</div>

      <div id="nvRouteList" style="overflow-y:auto; flex:1; min-height:60px;">Chargement…</div>

      <div class="modal-actions"><button class="btn-ghost" id="nvRouteCloseBtn">Fermer</button></div>

    </div>

  `;

  document.body.appendChild(overlay);

  const close = () => overlay.remove();

  overlay.addEventListener('click', e => { if(e.target === overlay) close(); });

  overlay.querySelector('#nvRouteCloseBtn').addEventListener('click', close);



  db.collection('nvRoutePings').where('agentName', '==', agentName).where('date', '==', dateKey).get()

    .then(snap => {

      const pings = snap.docs.map(d => d.data()).sort((a, b) => a.ts - b.ts);

      const list = overlay.querySelector('#nvRouteList');

      if(pings.length === 0){

        list.innerHTML = `<div style="text-align:center; padding:20px; color:var(--ink-soft); font-size:12.5px;">Aucune position enregistrée aujourd'hui pour cet agent.</div>`;

        return;

      }

      list.innerHTML = pings.slice().reverse().map(p => {

        const t = new Date(p.ts).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });

        const mapUrl = `https://www.google.com/maps?q=${p.lat},${p.lng}`;

        return `<div style="display:flex; justify-content:space-between; align-items:center; padding:8px 0; border-bottom:1px solid var(--line); font-size:12.5px;">

          <span style="color:var(--ink-soft);">${t}</span>

          <a href="${mapUrl}" target="_blank" rel="noopener" style="color:var(--sky-ink); font-weight:600;">📍 Voir sur la carte</a>

        </div>`;

      }).join('') + `<div style="text-align:center; padding-top:10px; font-size:11px; color:var(--ink-soft);">${pings.length} position${pings.length > 1 ? 's' : ''} enregistrée${pings.length > 1 ? 's' : ''}</div>`;

    }).catch(err => {

      console.error(err);

      overlay.querySelector('#nvRouteList').innerHTML = `<div style="color:var(--red); text-align:center; padding:20px;">Erreur de chargement.</div>`;

    });

}



/* ================= PLANNING NV — accès agent (PIN + localisation live) ================= */

const NV_ALL_AGENTS = [

  ...ROTATING_TEAMS_NV[0], ...ROTATING_TEAMS_NV[1], ...ROTATING_TEAMS_NV[2], ...ROTATING_TEAMS_NV[3],

  ...Object.values(NV_FIXED_ZONE_AGENTS).flat(),

];



let NV_PINS = {};

let currentNvAgent = null;

let nvGeoIntervalId = null;

let nvUiIntervalId = null;



function stableNvPin(name){

  // PIN dérivé du nom de l'agent, toujours identique (pas de hasard) — évite tout décalage

  // entre ce que l'admin affiche et ce qui est réellement enregistré côté serveur.

  let hash = 0;

  for(let i = 0; i < name.length; i++){

    hash = (hash * 31 + name.charCodeAt(i)) >>> 0;

  }

  return String(1000 + (hash % 9000));

}



/* ================= EMAIL (EmailJS — envoi manuel, pas d'automatisation serveur) ================= */

let NV_EMAILS = {};

const OPS_HQ_URL = 'https://najibharati-spec.github.io/FieldOps/';



const DEFAULT_NV_EMAILS = {

  "Charaf Loummihi": "loummihic@gmail.com",

  "Mourad Cherradi": "raxmou96@gmail.com",

  "Islam Elmardi": "islamelmardi.1993@gmail.com",

  "Youssef Derkaoui": "altuslabcom@gmail.com",

  "Ayoub Ben Malek": "jadalaeayoub@gmail.com",

  "Aymane Chlifane": "aymanechlihfane639@gmail.com",

  "Yahya Rhallah": "yahyarhallab00@gmail.com",

  "Oussama El Atrach": "oussamaelat@gmail.com",

  "Hassan Azgoune": "hassan.azgoune95@gmail.com",

  "Ayoub Abssi": "ayoubabssi05@gmail.com",

  "Achraf Akhlij": "achraf.akhlij2005@gmail.com",

  "Mehdi Ait Jebour": "mehdiaitjebour@gmail.com",

};



function ensureNvEmails(){

  return db.collection('planningNV').doc('emails').get().then(doc => {

    const stored = (doc.exists && doc.data()) || {};

    const emails = {};

    let needsWrite = false;

    NV_ALL_AGENTS.forEach(name => {

      if(stored[name]){

        emails[name] = stored[name];

      } else if(DEFAULT_NV_EMAILS[name]){

        emails[name] = DEFAULT_NV_EMAILS[name];

        needsWrite = true;

      } else {

        emails[name] = stored[name] || '';

      }

    });

    NV_EMAILS = emails;

    if(needsWrite){

      db.collection('planningNV').doc('emails').set(emails).catch(err => console.error(err));

    }

  });

}



function sendNvEmail(toEmail, toName, subject, message){

  if(!toEmail){

    alert('Aucun email enregistré pour cet agent.');

    return Promise.reject(new Error('no email'));

  }

  const gmailUrl = `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(toEmail)}&su=${encodeURIComponent(subject)}&body=${encodeURIComponent(message)}`;

  window.open(gmailUrl, '_blank');

  return Promise.resolve();

}



function sendAgentAccessEmail(agent){

  if(!agent.email){ alert(`${agent.name} n'a pas d'email enregistré.`); return; }

  if(!agent.pin){ alert(`${agent.name} n'a pas encore de PIN.`); return; }

  const subject = 'Bienvenue sur FieldOps — ton accès';

  const message = `Bonjour ${agent.name.split(' ')[0]},



Ton accès à l'espace Planning FieldOps est prêt.



Lien de connexion : ${OPS_HQ_URL}

Ton code PIN : ${agent.pin}



Connecte-toi, sélectionne ton profil, fais ton check-in et active ta localisation pour voir ton planning de la semaine.



Bonne journée sur le terrain !

L'équipe FieldOps`;

  sendNvEmail(agent.email, agent.name, subject, message).catch(() => {});

}



function sendNvWelcomeEmail(name){

  const agent = AGENTS.find(a => a.name === name);

  if(!agent){ alert(`${name} n'a pas encore de fiche dans la page Agents.`); return; }

  sendAgentAccessEmail(agent);

}



function ensureNvPins(){

  return db.collection('planningNV').doc('pins').get().then(doc => {

    const stored = (doc.exists && doc.data()) || {};

    const pins = {};

    let needsWrite = false;

    NV_ALL_AGENTS.forEach(name => {

      if(stored[name]){

        pins[name] = stored[name];

      } else {

        pins[name] = stableNvPin(name);

        needsWrite = true;

      }

    });

    NV_PINS = pins;

    if(needsWrite){

      db.collection('planningNV').doc('pins').set(pins).catch(err => console.error(err));

    }

  });

}



document.getElementById('showNvFlashBtn').addEventListener('click', () => {

  const panel = document.getElementById('nvFlashPanel');

  if(!panel.classList.contains('hidden')){ panel.classList.add('hidden'); return; }

  panel.innerHTML = '<p style="color:var(--ink-soft); font-size:13px;">Chargement...</p>';

  panel.classList.remove('hidden');

  db.collection('planningNV').doc('flashInfo').get().then(doc => {

    const current = (doc.exists && doc.data().message) || '';

    panel.innerHTML = `

      <div style="font-size:11px; text-transform:uppercase; letter-spacing:.04em; font-weight:700; color:var(--ink-soft); margin-bottom:8px;">📢 Info équipe — visible par tous les agents dans leur portail</div>

      <textarea id="nvFlashInput" class="login-input" rows="3" style="resize:vertical;" placeholder="Ex: Forte demande prévue à Agdal ce soir entre 20h et 22h, restez proches de Carrefour Gourmet.">${current}</textarea>

      <div style="display:flex; gap:8px;">

        <button class="btn-primary" id="nvFlashSaveBtn" style="flex:1;">Enregistrer</button>

        ${current ? '<button class="btn-ghost" id="nvFlashClearBtn" style="border:1px solid var(--line);">Effacer</button>' : ''}

      </div>

    `;

    document.getElementById('nvFlashSaveBtn').addEventListener('click', () => {

      const message = document.getElementById('nvFlashInput').value.trim();

      db.collection('planningNV').doc('flashInfo').set({ message, updatedAt: firebase.firestore.FieldValue.serverTimestamp() })

        .then(() => { panel.classList.add('hidden'); })

        .catch(err => { console.error(err); alert('Erreur lors de l\'enregistrement.'); });

    });

    const clearBtn = document.getElementById('nvFlashClearBtn');

    if(clearBtn) clearBtn.addEventListener('click', () => {

      db.collection('planningNV').doc('flashInfo').set({ message: '', updatedAt: firebase.firestore.FieldValue.serverTimestamp() })

        .then(() => { panel.classList.add('hidden'); })

        .catch(err => console.error(err));

    });

  }).catch(err => console.error(err));

});



document.getElementById('nvPortalLogoutBtn').addEventListener('click', () => {

  stopNvLocationTracking();

  auth.signOut().finally(() => location.reload());

});



document.getElementById('nvPortalRapportBtn').addEventListener('click', () => { openNvRapportModal(); });



function findNvAgentSlot(agentName){

  const weekIdx = getISOWeekIndex(currentWeekStartNv);

  const zones = buildZoneAssignmentsNv(weekIdx);

  for(const zone of Object.keys(zones)){

    const idx = zones[zone].indexOf(agentName);

    if(idx >= 0) return { zone, memberIdx: idx };

  }

  return null;

}



function renderNvPortal(){

  document.getElementById('nvPortalAgentName').textContent = currentNvAgent;

  const content = document.getElementById('nvPortalContent');

  const dateKey = dateKeyFor(new Date());

  const checkinId = `${dateKey}_${currentNvAgent}`;



  const firstName = currentNvAgent.split(' ')[0];

  const MOTIVATIONAL_QUOTES = [

    "Chaque course compte. Bonne journée sur le terrain !",

    "L'équipe compte sur toi aujourd'hui — donne le meilleur !",

    "Un jour à la fois, une livraison à la fois. On avance !",

    "Ton travail fait la différence pour toute l'équipe.",

    "Reste concentré, reste safe — et fais du bon boulot !",

    "Aujourd'hui encore, on assure la couverture. Bravo d'avance !",

  ];

  const quoteIdx = (new Date().getDate() + firstName.length) % MOTIVATIONAL_QUOTES.length;

  const motivationHtml = `

    <div style="background:linear-gradient(135deg,var(--sky),var(--sky-deep)); border-radius:16px; padding:22px 20px; margin-bottom:16px; text-align:center; color:#fff; box-shadow:0 16px 34px -18px rgba(0,160,130,0.55);">

      <div style="width:52px; height:52px; border-radius:50%; background:rgba(255,255,255,0.18); display:flex; align-items:center; justify-content:center; font-size:26px; margin:0 auto 10px;">🛵</div>

      <div style="font-size:17px; font-weight:800; margin-bottom:4px;">Bonjour, ${firstName} 💪</div>

      <div style="font-size:12.5px; opacity:.9; max-width:280px; margin:0 auto; line-height:1.5;">${MOTIVATIONAL_QUOTES[quoteIdx]}</div>

      <div id="nvStreakBadge" style="margin-top:12px;"></div>

    </div>

  `;



  db.collection('nvCheckins').doc(checkinId).get().then(checkinDoc => {

    const isCheckedIn = checkinDoc.exists;



    if(!isCheckedIn){

      content.innerHTML = `

        ${motivationHtml}

        <div style="text-align:center; padding:28px 18px; background:var(--sand-deep); border-radius:14px; border:1.5px solid var(--line);">

          <div style="width:64px; height:64px; border-radius:50%; background:rgba(0,160,130,0.14); display:flex; align-items:center; justify-content:center; font-size:30px; margin:0 auto 12px;">✅</div>

          <div style="font-size:15px; font-weight:800; margin-bottom:6px;">Commence ton shift</div>

          <div style="font-size:12.5px; color:var(--ink-soft); max-width:260px; margin:0 auto 18px; line-height:1.5;">Fais ton check-in pour accéder à ton planning, ta zone et ton hotspot du jour.</div>

          <button class="portal-btn" id="nvCheckinBtn" style="width:auto; padding:13px 30px;">✅ Check-in</button>

        </div>

      `;

      document.getElementById('nvCheckinBtn').addEventListener('click', () => {

        const btn = document.getElementById('nvCheckinBtn');

        btn.disabled = true;

        db.collection('nvCheckins').doc(checkinId).set({

          agentName: currentNvAgent, date: dateKey, checkInAt: firebase.firestore.FieldValue.serverTimestamp(),

        }).then(() => renderNvPortal()).catch(err => { console.error(err); btn.disabled = false; });

      });

      return;

    }



    renderNvPortalFull(content, motivationHtml, dateKey);

  }).catch(err => {

    console.error(err);

    content.innerHTML = `<p style="padding:20px; color:var(--red);">Erreur de chargement.</p>`;

  }).finally(() => renderNvStreakBadge());

}



function renderNvStreakBadge(){

  const el = document.getElementById('nvStreakBadge');

  if(!el) return;

  // Une seule requête sur toute la collection : permet de calculer le streak de CHAQUE agent

  // (pour le classement) sans multiplier les lectures Firestore.

  db.collection('nvCheckins').get().then(snap => {

    const byAgent = {};

    snap.docs.forEach(d => {

      const data = d.data();

      if(!byAgent[data.agentName]) byAgent[data.agentName] = new Set();

      byAgent[data.agentName].add(data.date);

    });



    function computeStreak(dates){

      const todayKey = dateKeyFor(new Date());

      let streak = 0;

      let cursor = new Date();

      if(!dates.has(todayKey)) cursor.setDate(cursor.getDate() - 1);

      while(dates.has(dateKeyFor(cursor))){

        streak++;

        cursor.setDate(cursor.getDate() - 1);

      }

      return streak;

    }



    const myStreak = computeStreak(byAgent[currentNvAgent] || new Set());



    const ranking = NV_ALL_AGENTS

      .map(name => ({ name, streak: computeStreak(byAgent[name] || new Set()) }))

      .sort((a, b) => b.streak - a.streak);

    const myRank = ranking.findIndex(r => r.name === currentNvAgent) + 1;

    const rankTag = myStreak >= 1 ? ` · <b>#${myRank}</b>/${ranking.length}` : '';



    if(myStreak >= 2){

      el.innerHTML = `<span style="display:inline-flex; align-items:center; gap:5px; background:rgba(255,255,255,0.18); padding:5px 12px; border-radius:20px; font-size:12px; font-weight:700;">🔥 ${myStreak} jours d'affilée${rankTag}</span>`;

    } else if(myStreak === 1 && myRank <= 3){

      el.innerHTML = `<span style="display:inline-flex; align-items:center; gap:5px; background:rgba(255,255,255,0.18); padding:5px 12px; border-radius:20px; font-size:12px; font-weight:700;">🏆 #${myRank}/${ranking.length} de l'équipe</span>`;

    }

  }).catch(err => console.error(err));

}



function renderNvFlashInfoBox(){

  const box = document.getElementById('nvFlashInfoBox');

  if(!box) return;

  db.collection('planningNV').doc('flashInfo').get().then(doc => {

    const message = doc.exists ? (doc.data().message || '') : '';

    if(!message){ box.innerHTML = ''; return; }

    box.innerHTML = `

      <div style="background:rgba(251,191,36,0.12); border:1px solid rgba(251,191,36,0.35); border-radius:14px; padding:14px 16px; margin-bottom:16px;">

        <div style="font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:.04em; color:#B45309; margin-bottom:6px;">📢 Info équipe</div>

        <div style="font-size:13px; color:var(--ink); line-height:1.5; white-space:pre-wrap;">${message}</div>

      </div>

    `;

  }).catch(err => console.error(err));

}



function renderNvMyReports(){

  const box = document.getElementById('nvMyReportsBox');

  if(!box) return;

  // Pas de orderBy() ici volontairement : évite de dépendre d'un index composite Firestore

  // (agentName + createdAt) qui pourrait ne pas exister — on trie côté client à la place.

  db.collection('rapports').where('agentName', '==', currentNvAgent).limit(20).get().then(snap => {

    if(snap.empty){ box.innerHTML = ''; return; }

    const docs = snap.docs

      .slice()

      .sort((a, b) => {

        const ta = a.data().createdAt && a.data().createdAt.toMillis ? a.data().createdAt.toMillis() : 0;

        const tb = b.data().createdAt && b.data().createdAt.toMillis ? b.data().createdAt.toMillis() : 0;

        return tb - ta;

      })

      .slice(0, 5);

    const items = docs.map(d => {

      const r = d.data();

      const when = r.createdAt && r.createdAt.toDate ? r.createdAt.toDate().toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit' }) : (r.date || '');

      const typeLabel = r.type === 'remarque' ? '💬' : r.type === 'urgent' ? '🚨' : '📩';

      const urgentTag = r.urgent ? ' <span style="color:var(--red); font-weight:800;">URGENT</span>' : '';

      const seenTag = r.seenAt ? ' <span style="color:var(--green); font-weight:700;">✅ Lu</span>' : '';

      const replyBlock = r.adminReply

        ? `<div style="margin-top:4px; padding:6px 10px; background:rgba(0,160,130,0.08); border-radius:8px; border-left:3px solid var(--sky); font-size:11.5px; color:var(--ink);">💬 <b>Réponse :</b> ${r.adminReply}</div>`

        : '';

      return `<div style="padding:8px 0; border-bottom:1px solid var(--line); font-size:12px;">

        <span style="color:var(--ink-soft);">${when}</span> ${typeLabel}${urgentTag}${seenTag} <span style="color:var(--ink);">${(r.text || '').slice(0, 60)}${(r.text || '').length > 60 ? '…' : ''}</span>

        ${replyBlock}

      </div>`;

    }).join('');

    box.innerHTML = `

      <div style="font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.04em; color:var(--ink-soft); margin-bottom:6px;">Mes derniers envois</div>

      ${items}

    `;

    // L'agent vient de voir ses réponses : on marque tout comme lu pour la pastille de

    // notification, et on retire la pastille immédiatement.

    const repliedIds = snap.docs.filter(d => d.data().adminReply).map(d => d.id);

    if(repliedIds.length){

      localStorage.setItem('nvRepliesSeen_' + currentNvAgent, JSON.stringify(repliedIds));

    }

    const badge = document.getElementById('nvRapportReplyBadge');

    if(badge) badge.classList.add('hidden');

  }).catch(err => console.error(err));

}



let nvReplyNotifUnsub = null;

function listenNvReplyNotifications(){

  if(nvReplyNotifUnsub) nvReplyNotifUnsub();

  // Écoute en continu (pas juste à l'ouverture de la modale) : la pastille rouge sur le bouton

  // "Rapport" s'allume dès qu'une réponse admin arrive sur un de mes envois, même si je n'ai

  // jamais ouvert la modale pour la voir.

  nvReplyNotifUnsub = db.collection('rapports').where('agentName', '==', currentNvAgent).onSnapshot(snap => {

    const seen = new Set(JSON.parse(localStorage.getItem('nvRepliesSeen_' + currentNvAgent) || '[]'));

    const hasUnseenReply = snap.docs.some(d => d.data().adminReply && !seen.has(d.id));

    const badge = document.getElementById('nvRapportReplyBadge');

    if(badge) badge.classList.toggle('hidden', !hasUnseenReply);

  }, err => console.error(err));

}





function renderNvPortalFull(content, motivationHtml, dateKey){

  const weekIdx = getISOWeekIndex(currentWeekStartNv);

  const slot = findNvAgentSlot(currentNvAgent);

  if(!slot){ content.innerHTML = motivationHtml + '<p style="padding:20px;">Aucune affectation cette semaine.</p>'; return; }



  const { zone, memberIdx } = slot;

  const override = NV_OVERRIDES[currentNvAgent] || {};

  const todayIdx = (new Date().getDay() + 6) % 7; // 0=Lun..6=Dim

  const weekShiftInfo = resolveNvShiftInfo(override, memberIdx, weekIdx, currentNvAgent);

  const todayOverride = (override.dayOverrides && override.dayOverrides[todayIdx]) || null;

  const shiftInfo = todayOverride ? resolveNvShiftInfo(todayOverride, memberIdx, weekIdx) : weekShiftInfo;

  const offDayIdx = (typeof override.offDayIdx === 'number') ? override.offDayIdx : (memberIdx + weekIdx) % 4;

  const shiftLabel7 = shiftInfo.label7;

  const shiftLabel8 = shiftInfo.label8;

  const shiftName = shiftInfo.name;

  const isJourShift = shiftInfo.isJour;

  const slots = HOTSPOTS_NV[zone] || [];

  const zoneColor = NV_ZONE_COLORS[zone] || 'var(--sky)';

  const ALL_DAY_LABELS_NV = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'];



  window.__nvTodayShiftInfo = (todayIdx !== offDayIdx) ? shiftInfo : null;



  // Génère le bloc d'un jour (nom, heure, hotspot + tous les lieux) — réutilisé pour la carte

  // "aujourd'hui" mise en avant ET pour chaque ligne de "Voir ma semaine complète".

  function buildNvDayBlock(dayIdx, dayLabel, featured){

    const isToday = dayIdx === todayIdx;

    const todayTag = isToday ? ' <span class="nv-today-tag">Aujourd\'hui</span>' : '';

    const isWeekend = dayIdx >= 4; // Ven/Sam/Dim

    const dayOverride = (override.dayOverrides && override.dayOverrides[dayIdx]) || null;

    const dayInfo = dayOverride ? resolveNvShiftInfo(dayOverride, memberIdx, weekIdx) : weekShiftInfo;

    const dayTime = isWeekend && dayIdx < 6 ? dayInfo.label8 : dayInfo.label7; // 8h Ven-Sam, 7h le reste

    const timeChip = dayIdx === offDayIdx ? '' : `<span class="nv-time-chip ${dayInfo.isJour ? 'nv-time-jour' : 'nv-time-nuit'}" style="font-size:13px; padding:5px 12px; margin-left:0;">🕐 ${dayTime}</span>`;

    const dayNameLine = `<div class="nv-day-label" style="font-size:${featured ? '19px' : '15px'}; margin-bottom:8px; ${featured ? 'text-align:center; text-transform:uppercase;' : ''}">${dayLabel}${featured ? '' : todayTag}</div>`;

    const wrapClass = featured ? 'nv-portal-day nv-day-today' : `nv-portal-day${isToday ? ' nv-day-today' : ''}`;

    const wrapStyle = featured ? ' style="background:var(--sand-deep); border-radius:16px; padding:18px; border:none; text-align:center;"' : '';



    if(dayIdx === offDayIdx){

      return `<div class="${wrapClass}"${wrapStyle}>${dayNameLine}<span class="nv-day-chip off" style="font-size:13px; padding:5px 12px;">😴 Repos (OFF)</span></div>`;

    }

    if(slots.length){

      const slotIdx = resolveNvHotspotIndexForAgent(override, slots, memberIdx, dayIdx, offDayIdx);

      const slotLabel = (slots[slotIdx] && slots[slotIdx].label) || `Hotspot ${slotIdx + 1}`;

      const metaLine = `<div style="margin-bottom:8px; display:flex; align-items:center; gap:8px; flex-wrap:wrap; ${featured ? 'justify-content:center;' : ''}">${timeChip}<span style="font-size:12px; color:var(--ink-soft); font-weight:600;">📍 ${slotLabel}</span></div>`;

      // Les lieux précis (venues) ne sont montrés QUE pour aujourd'hui — pour les autres jours de

      // "Voir ma semaine complète", seule l'heure et le numéro de hotspot sont visibles (pas les

      // adresses), pour ne pas exposer les hotspots à l'avance ou après coup.

      if(!isToday){

        return `<div class="${wrapClass}"${wrapStyle}>${dayNameLine}${metaLine}</div>`;

      }

      const venues = (slots[slotIdx] && slots[slotIdx].venues) || [];

      const venueChips = venues.length

        ? venues.map(v => `<a class="nv-day-chip nv-day-chip-link" href="${getVenueMapUrl(v)}" target="_blank" rel="noopener">📍 ${v}</a>`).join('')

        : `<span class="nv-day-chip">—</span>`;

      const venuesWrap = featured ? `<div class="nv-day-venues" style="justify-content:center;">${venueChips}</div>` : `<div class="nv-day-venues">${venueChips}</div>`;

      return `<div class="${wrapClass}"${wrapStyle}>${dayNameLine}${metaLine}${venuesWrap}</div>`;

    }

    return `<div class="${wrapClass}"${wrapStyle}>${dayNameLine}<div style="margin-bottom:8px;">${timeChip}</div><span class="nv-day-chip">—</span></div>`;

  }



  const todayBlockHtml = buildNvDayBlock(todayIdx, ALL_DAY_LABELS_NV[todayIdx], true);



  let weekHtml = '';

  ALL_DAY_LABELS_NV.forEach((dayLabel, dayIdx) => {

    weekHtml += buildNvDayBlock(dayIdx, dayLabel, false);

  });



  const firstName = currentNvAgent.split(' ')[0];



  content.innerHTML = `

    ${motivationHtml}

    <div id="nvFlashInfoBox"></div>

    <div class="nv-portal-hero" style="background: linear-gradient(135deg, ${zoneColor}, ${zoneColor}CC);">      <div class="nv-portal-hero-label">Zone cette semaine</div>

      <div class="nv-portal-hero-zone">${getNvZoneLabel(zone)}</div>

    </div>

    <div style="margin-top:14px;">${todayBlockHtml}</div>

    <div class="portal-status" style="margin-top:14px;">

      <button class="btn-ghost" id="nvToggleWeekBtn" style="width:100%; display:flex; align-items:center; justify-content:space-between; padding:0; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:.04em; color:var(--ink-soft);">

        <span>📅 Voir ma semaine complète</span>

        <span id="nvToggleWeekIcon">▾</span>

      </button>

      <div class="nv-day-strip hidden" id="nvWeekStripWrap" style="margin-top:12px;">${weekHtml}</div>

    </div>

  `;



  renderNvFlashInfoBox();



  // Grille horizontale (façon Vue Semaine admin) affichée en arrière-plan tout en haut de la

  // page, uniquement sur grand écran — juste un aperçu visuel, non cliquable.

  const topWeekWrap = document.getElementById('nvPortalTopWeekWrap');

  if(topWeekWrap){

    const topCells = ALL_DAY_LABELS_NV.map((dayLabel, dayIdx) => {

      const isToday = dayIdx === todayIdx;

      if(dayIdx === offDayIdx){

        return `<div style="flex:1; text-align:center;"><div style="font-size:10px; font-weight:700; text-transform:uppercase; color:var(--ink-soft); margin-bottom:4px;">${dayLabel.slice(0,3)}${isToday ? ' •' : ''}</div><span class="cell-pill" style="background:var(--sand-deep); color:var(--ink-soft); display:block; font-size:11px; padding:5px 4px;">OFF</span></div>`;

      }

      const dayOverride = (override.dayOverrides && override.dayOverrides[dayIdx]) || null;

      const dayInfo = dayOverride ? resolveNvShiftInfo(dayOverride, memberIdx, weekIdx) : weekShiftInfo;

      const isWeekend = dayIdx >= 4;

      const dayTime = isWeekend && dayIdx < 6 ? dayInfo.label8 : dayInfo.label7;

      let hotspotLabel = '';

      if(slots.length){

        const slotIdx = resolveNvHotspotIndexForAgent(override, slots, memberIdx, dayIdx, offDayIdx);

        const slotLabel = (slots[slotIdx] && slots[slotIdx].label) || `Hotspot ${slotIdx + 1}`;

        if(isToday){

          const venues = (slots[slotIdx] && slots[slotIdx].venues) || [];

          hotspotLabel = venues.length

            ? `<div style="font-size:9px; color:var(--ink-soft); margin-top:2px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">📍 ${venues[0]}${venues.length > 1 ? '…' : ''}</div>`

            : `<div style="font-size:9px; color:var(--ink-soft); margin-top:2px;">📍 ${slotLabel}</div>`;

        } else {

          hotspotLabel = `<div style="font-size:9px; color:var(--ink-soft); margin-top:2px;">📍 ${slotLabel}</div>`;

        }

      }

      return `<div style="flex:0 0 62px; text-align:center;">

        <div style="font-size:10px; font-weight:700; text-transform:uppercase; color:${isToday ? 'var(--sky-ink)' : 'var(--ink-soft)'}; margin-bottom:4px;">${dayLabel.slice(0,3)}${isToday ? ' •' : ''}</div>

        <span class="cell-pill ${dayInfo.isJour ? 'jour' : 'nuit'}" style="display:block; background:${dayInfo.isJour ? 'rgba(4,120,87,0.14)' : 'rgba(184,128,14,0.16)'}; color:${dayInfo.isJour ? '#047857' : '#8A6208'}; font-size:11.5px; font-weight:800; font-family:var(--font-mono); padding:5px 4px; border-radius:6px;">${dayTime}</span>

        ${hotspotLabel}

      </div>`;

    }).join('');

    topWeekWrap.innerHTML = `

      <div style="display:flex; align-items:center; gap:10px; margin-bottom:10px;">

        <b style="font-size:13px;">${currentNvAgent}</b>

        <span style="font-size:11px; color:var(--ink-soft);">${getNvZoneLabel(zone)}</span>

      </div>

      <div style="display:flex; gap:6px; overflow-x:auto; padding-bottom:2px;">${topCells}</div>

    `;

    topWeekWrap.classList.remove('hidden');

  }



  document.getElementById('nvToggleWeekBtn').addEventListener('click', () => {

    const wrap = document.getElementById('nvWeekStripWrap');

    const icon = document.getElementById('nvToggleWeekIcon');

    const nowHidden = wrap.classList.contains('hidden');

    wrap.classList.toggle('hidden', !nowHidden);

    icon.textContent = nowHidden ? '▴' : '▾';

  });

}



function openNvRapportModal(){

  const overlay = document.createElement('div');

  overlay.className = 'modal-overlay';

  overlay.innerHTML = `

    <div class="modal-card" style="max-width:420px; max-height:85vh; display:flex; flex-direction:column; overflow-y:auto;">

      <h3>📝 Rapport &amp; remarques</h3>

      <div style="background:var(--sand-deep); border-radius:14px; padding:14px; border:1px solid var(--line); margin-top:10px;">

        <select id="nvRapportType" class="login-input" style="margin-bottom:8px;">

          <option value="rapport">📩 Rapport</option>

          <option value="urgent">🚨 Urgent</option>

          <option value="remarque">💬 Remarque</option>

        </select>

        <textarea id="nvRapportText" class="login-input" rows="3" style="resize:vertical; min-height:70px;" placeholder="Écris ton rapport ou ta remarque ici..."></textarea>

        <button class="portal-btn" id="nvSendRapportBtn" style="padding:12px; font-size:14px;">Envoyer</button>

        <div id="nvRapportMsg" style="text-align:center; font-size:12px; margin-top:8px;"></div>

      </div>

      <div id="nvMyReportsBox" style="margin-top:12px;"></div>

      <div class="modal-actions"><button class="btn-ghost" id="nvRapportCloseBtn">Fermer</button></div>

    </div>

  `;

  document.body.appendChild(overlay);

  const close = () => overlay.remove();

  overlay.addEventListener('click', e => { if(e.target === overlay) close(); });

  overlay.querySelector('#nvRapportCloseBtn').addEventListener('click', close);



  renderNvMyReports();



  overlay.querySelector('#nvSendRapportBtn').addEventListener('click', () => {

    const type = overlay.querySelector('#nvRapportType').value;

    const text = overlay.querySelector('#nvRapportText').value.trim();

    const urgent = type === 'urgent';

    const msgEl = overlay.querySelector('#nvRapportMsg');

    if(!text){ msgEl.style.color = 'var(--red)'; msgEl.textContent = 'Écris quelque chose avant d\'envoyer.'; return; }

    const btn = overlay.querySelector('#nvSendRapportBtn');

    btn.disabled = true;

    db.collection('rapports').add({

      agentName: currentNvAgent,

      type,

      text,

      urgent,

      date: dateKeyFor(new Date()),

      createdAt: firebase.firestore.FieldValue.serverTimestamp(),

    }).then(() => {

      overlay.querySelector('#nvRapportText').value = '';

      msgEl.style.color = 'var(--green)';

      msgEl.textContent = 'Envoyé ✅';

      btn.disabled = false;

      setTimeout(() => { msgEl.textContent = ''; }, 3000);

      renderNvMyReports();

    }).catch(err => {

      console.error(err);

      msgEl.style.color = 'var(--red)';

      msgEl.textContent = 'Erreur, réessaie.';

      btn.disabled = false;

    });

  });

}



function updateNvGeoButtonVisibility(){

  const area = document.getElementById('nvGeoArea');

  if(!area) return;

  const isActive = !!nvGeoIntervalId;

  const everActivated = localStorage.getItem('nvGeoEverActivated_' + currentNvAgent) === '1';

  const shiftEnd = getShiftEndDateNv(window.__nvTodayShiftInfo);

  const minsLeft = shiftEnd ? Math.round((shiftEnd.getTime() - Date.now()) / 60000) : null;

  const nearShiftEnd = isActive && minsLeft !== null && minsLeft <= 10 && minsLeft >= 0;



  if(!isActive){

    // Jamais activé (ou désactivé) : le bouton reste visible pour pouvoir démarrer.

    area.innerHTML = `

      <div style="text-align:center; padding:22px 16px; background:var(--sand-deep); border-radius:14px; border:1px solid var(--line);">

        <div style="font-size:30px; margin-bottom:8px;">📍</div>

        <div style="font-size:14px; font-weight:800; margin-bottom:4px;">Localisation non activée</div>

        <div style="font-size:12px; color:var(--ink-soft); max-width:260px; margin:0 auto 16px; line-height:1.5;">Active ta position pour que l'équipe sache où tu es pendant ton shift. Elle s'arrête automatiquement à la fin.</div>

        <button class="portal-btn" id="nvGeoBtn" style="width:auto; padding:12px 24px;">Activer la localisation</button>

      </div>

      <div id="nvGeoStatus" style="text-align:center; font-size:12px; color:var(--ink-soft); margin-top:10px;"></div>

    `;

    document.getElementById('nvGeoBtn').addEventListener('click', toggleNvLocationTracking);

  } else if(nearShiftEnd){

    // Shift bientôt fini : on réaffiche le bouton comme rappel pour désactiver.

    area.innerHTML = `

      <button class="portal-btn" id="nvGeoBtn" style="background:var(--mustard-deep);">⏰ Ton shift se termine dans ${minsLeft} min — Désactiver</button>

      <div id="nvGeoStatus" style="text-align:center; font-size:12px; color:var(--ink-soft); margin-top:10px;">✅ Localisation active</div>

    `;

    document.getElementById('nvGeoBtn').addEventListener('click', toggleNvLocationTracking);

  } else {

    // Actif et loin de la fin du shift : pas besoin d'afficher un bouton, juste un petit indicateur discret.

    area.innerHTML = `<div style="text-align:center; font-size:12px; color:var(--green); padding:6px;">🟢 Localisation active</div>`;

  }

}



function getShiftEndDateNv(shiftInfo){

  if(!shiftInfo) return null;

  const day = new Date().getDay(); // 0=Dim..6=Sam

  const isWeekendShift = (day === 5 || day === 6); // Ven=5, Sam=6

  const label = isWeekendShift ? shiftInfo.label8 : shiftInfo.label7;

  const endStr = label.split('-')[1]; // "01:00" etc.

  const [h, m] = endStr.split(':').map(Number);

  const start = new Date();

  const end = new Date();

  end.setHours(h, m, 0, 0);

  // Si l'heure de fin est "plus petite" que maintenant de nuit (ex: fin 02:00 après minuit), elle tombe demain.

  if(end < start && h < 12) end.setDate(end.getDate() + 1);

  return end;

}



function toggleNvLocationTracking(){

  if(nvGeoIntervalId){

    stopNvLocationTracking();

    localStorage.removeItem('nvGeoActive_' + currentNvAgent);

  } else {

    localStorage.setItem('nvGeoActive_' + currentNvAgent, '1');

    localStorage.setItem('nvGeoEverActivated_' + currentNvAgent, '1');

    startNvLocationTracking();

  }

  updateNvGeoButtonVisibility();

}



function startNvLocationTracking(){

  if(nvGeoIntervalId) return; // déjà actif

  if(!navigator.geolocation){

    const statusEl = document.getElementById('nvGeoStatus');

    if(statusEl) statusEl.textContent = "Ton navigateur ne supporte pas la localisation.";

    return;

  }

  const sendPosition = () => {

    // Arrêt automatique si le shift du jour est terminé.

    const shiftEnd = getShiftEndDateNv(window.__nvTodayShiftInfo);

    if(shiftEnd && Date.now() > shiftEnd.getTime()){

      stopNvLocationTracking();

      localStorage.removeItem('nvGeoActive_' + currentNvAgent);

      updateNvGeoButtonVisibility();

      return;

    }

    navigator.geolocation.getCurrentPosition(pos => {

      const lat = pos.coords.latitude;

      const lng = pos.coords.longitude;

      const now = Date.now();

      db.collection('planningNV').doc('locations').set({

        [currentNvAgent]: { lat, lng, updatedAt: now },

      }, { merge: true }).catch(err => console.error(err));

      // Traçabilité : on garde un historique de chaque position (pas seulement la dernière)

      // pour pouvoir reconstituer le trajet du jour a posteriori.

      db.collection('nvRoutePings').add({

        agentName: currentNvAgent,

        date: dateKeyFor(new Date()),

        lat, lng,

        ts: now,

      }).catch(err => console.error(err));

    }, err => {

      console.error(err);

    }, { enableHighAccuracy: true, timeout: 15000 });

  };

  sendPosition();

  nvGeoIntervalId = setInterval(sendPosition, 150000); // toutes les 2.5 min

}



function stopNvLocationTracking(){

  if(nvGeoIntervalId){ clearInterval(nvGeoIntervalId); nvGeoIntervalId = null; }

}



renderProfiles();

</script>



<!-- ===== PLANNING AGENTS PRO : coller ce bloc juste AVANT </body> ===== -->
<style>
.pro{--y:#FFC244;--ys:rgba(255,194,68,.32);--g:#00A082;--gs:rgba(0,160,130,.15);--c:#1F1B0E;--m:#6B5D3D;--b:rgba(80,64,20,.22);--bad:#DC2626;color:var(--c)}
.pro .mu{color:var(--m);font-size:12px}
:where(.pro) button,:where(.pro) select,:where(.pro) input{font:inherit;color:var(--c);background:#fffdf6;border:1px solid var(--b);border-radius:10px;padding:6px 10px}
:where(.pro) button{cursor:pointer}
.pro .p-row{display:flex;flex-wrap:wrap;gap:6px;margin:10px 0;align-items:center}
.pro .p-row button.a{background:var(--y);border-color:var(--y);font-weight:700}
.pro .p-w{overflow-x:auto;background:#fff;border:1.5px solid var(--b);border-radius:18px}
.pro .p-tl{display:flex;flex-wrap:wrap;align-items:center;gap:12px}
.pro .p-tl input[type=range]{flex:1;min-width:170px;padding:0;accent-color:var(--g)}
.pro .p-tm{font:800 30px/1 "JetBrains Mono",monospace;min-width:92px}
.pro .p-live{border-radius:99px;padding:3px 11px;font-size:12px;font-weight:700;background:var(--gs);color:#007A63}
.pro .p-live.m{background:var(--ys);color:#8a5a00}
.pro .p-grp,.pro .p-ags{display:grid;gap:12px;grid-template-columns:repeat(auto-fill,minmax(250px,1fr))}
.pro .p-ac{background:#fff;border:1.5px solid var(--b);border-left:5px solid var(--b);border-radius:14px;padding:12px 14px}
.pro .p-ac.on{border-left-color:var(--g)}.pro .p-ac.soon{border-left-color:var(--y)}.pro .p-ac.off{border-left-style:dashed;opacity:.7}
.pro .p-h{display:flex;justify-content:space-between;gap:6px;align-items:flex-start}
.pro .p-bd{border-radius:99px;padding:1px 9px;font-size:11px;font-weight:700;white-space:nowrap}
.pro .p-z{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--b);background:#FFF6DE;border-radius:99px;padding:1px 9px 1px 3px;font-size:11.5px;font-weight:600;margin-top:3px}
.pro .p-z i{font-style:normal;border-radius:50%;width:17px;height:17px;font-size:10px;font-weight:800;display:inline-flex;align-items:center;justify-content:center}
.pro .p-bar{height:6px;background:rgba(80,64,20,.12);border-radius:9px;margin:7px 0 3px;overflow:hidden}.pro .p-bar i{display:block;height:100%;background:var(--g)}
.pro .p-gi{min-width:780px;padding:8px 12px}
.pro .p-gr{display:grid;grid-template-columns:120px 1fr;align-items:center;border-top:1px solid var(--b)}.pro .p-gr:first-child{border-top:0}
.pro .p-gn{font-size:12px;font-weight:700;position:sticky;left:0;background:#fff;z-index:2;padding:4px 6px 4px 0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.pro .p-gh{display:grid;font-size:11px;color:var(--m)}
.pro .p-gc{position:relative;height:32px;overflow:hidden;background:repeating-linear-gradient(90deg,transparent 0,transparent calc(100%/var(--nh,17) - 1px),var(--b) calc(100%/var(--nh,17) - 1px),var(--b) calc(100%/var(--nh,17)))}
.pro .p-gb{position:absolute;top:4px;bottom:4px;border:1px solid var(--g);background:var(--gs);border-radius:8px;font-size:11px;padding:2px 7px;white-space:nowrap;overflow:hidden;font-weight:700}
.pro .p-nl{position:absolute;top:0;bottom:0;width:2px;background:var(--bad);z-index:1}
.pro table.p-t{border-collapse:collapse;width:100%;min-width:1350px;font-size:15px}
.pro .p-t th,.pro .p-t td{padding:8px;border-bottom:1px solid var(--b);text-align:left;vertical-align:top;font-size:14px;text-transform:none;letter-spacing:0;color:var(--c)}
.pro .p-t th{background:#2B2B2B;color:#fff;font-size:13px;padding:11px 8px}.pro .p-t th.td{background:var(--y);color:#1F1B0E}
.pro .p-t td:first-child,.pro .p-t th:first-child{position:sticky;left:0;z-index:2;background:#FFF6DE}.pro .p-t th:first-child{background:#2B2B2B}
.pro .p-t td.c{background:rgba(120,105,60,.1)}.pro .p-t td.sm{background:rgba(255,194,68,.5)}.pro .p-t td.sn{background:rgba(0,160,130,.24)}
.pro .p-t td.m{box-shadow:inset 0 0 0 2px var(--g)}.pro .p-t td.o{background:#EDDFB8;color:var(--m);text-align:center;vertical-align:middle;font-weight:700}
.pro .p-tt{display:flex;gap:4px;margin-bottom:4px}.pro .p-tt input{width:50%;min-width:0;padding:5px 4px;font-size:14px;font-weight:600}
.pro .p-t td select{width:100%;font-size:14px;font-weight:700;border-left:6px solid var(--y);padding:6px 8px}
.pro .p-ag{white-space:nowrap}.pro .p-ag b{font-size:15px}.pro .p-ag button{padding:0 7px;margin:3px 3px 0 0;font-size:12px}
.pro .p-bad{color:var(--bad);font-weight:800}.pro .p-t td.p-bad{background:rgba(220,38,38,.14)}.pro .p-alert{background:rgba(220,38,38,.08);border:1px solid rgba(220,38,38,.3);border-radius:12px;padding:10px 14px;margin:10px 0;color:#991b1b;font-size:13px;line-height:1.6}
.pro .p-zr{display:grid;grid-template-columns:1fr 105px 150px 36px 36px;gap:6px;margin-bottom:6px}
.pro .p-tr{display:grid;grid-template-columns:128px 1fr 36px;gap:6px;margin-bottom:6px}.pro .p-tr input{width:100%}.pro .p-dl{display:grid;grid-template-columns:repeat(auto-fit,minmax(90px,1fr));gap:8px}.pro .p-dl input{width:100%}
.pro .p-sg{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px;margin-top:14px}
.pro .p-cv table{border-collapse:collapse;width:100%;min-width:800px}.pro .p-cv td,.pro .p-cv th{text-align:center;padding:5px 3px;font-size:12px;border-bottom:1px solid var(--b)}
.pro .p-cur{box-shadow:inset 0 0 0 2px var(--g)}
.pro .p-lgd{display:flex;flex-wrap:wrap;gap:14px;margin:6px 0;font-size:12px;color:var(--m)}.pro .p-lgd i{display:inline-block;width:12px;height:12px;border-radius:4px;margin-right:5px;vertical-align:-2px;border:1px solid var(--b)}
.pro .p-h b{white-space:nowrap}.pro .p-t td:nth-child(2) select{min-width:96px;border-left-width:1px}
@media(max-width:640px){.pro .p-tm{font-size:24px}.pro .p-ags{grid-template-columns:1fr}.pro .p-t th,.pro .p-t td{padding:6px}}
</style>
<script>
(function(){
var DAYS=["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"],GN=["A","B","C","D"],GC=["#FFC244","#00A082","#3a3a37","#B9800E"],GT=["#1d1d1b","#fff","#fff","#fff"];
var TPL8=[[10,"McDo Bouragrag"],[18,"McDo Bouragrag"],[10,"McDo Temara"],[18,"McDo Temara"],[10,"McDo Hay Riad"],[18,"McDo Hay Riad"],[11,"McDo Agdal"],[18,"McDo Agdal"],[10,"McDo Hay Chaou"],[15,"Le Marché"],[16,"McDo Carrousel"],[12,"McDo Medina"]];
var TPL7=[[10,"McDo Bouragrag"],[18,"McDo Bouragrag"],[11,"McDo Temara"],[18,"McDo Temara"],[10,"McDo Hay Riad"],[18,"McDo Hay Riad"],[11,"McDo Agdal"],[18,"McDo Agdal"],[18,"Le Marché"],[10,"McDo Hay Chaou"],[15,"Le Marché"],[16,"McDo Carrousel"]];
function mkT(a){return a.map(function(x){return{s:x[0],z:x[1]}})}
var PRIZ=["McDo Bouragrag","McDo Temara","McDo Hay Riad","McDo Agdal"];
var NAMES=["Yahya Rhallah","Youssef Derkaoui","Aymane Chlifane","Islam Elmardi","Oussama El Atrach","Mehdi Ait Jebour","Mourad Cherradi","Ayoub Abssi","Hassan Azgoune","Ayoub Ben Malek","Charaf Loummihi","Mohamd Khalyl","Zakariae Benali"];
var ZN=[["McDo Medina",0],["McDo Bouragrag",0],["McDo Hay Chaou",0],["McDo Carrousel",1],["McDo LGV",1],["McDo Agdal",1],["Le Marché",2],["McDo Temara",2],["McDo Hay Riad",2],["Technopolis",3]];
function def(){return{agents:NAMES.map(function(n,i){var o={id:i+1,n:n,off:i%4};if(n==="Oussama El Atrach"){o.fo=3;o.fw=nowInfo().wi+1}return o}),ov:{},offOv:{},week:0,nid:100,zones:ZN.map(function(z){return z[1]===3?{n:z[0],g:3,fix:13}:{n:z[0],g:z[1],pri:PRIZ.indexOf(z[0])>=0}}),priFrom:null,tm:true,tplFrom:null,tpl:{t7:mkT(TPL7),t8:mkT(TPL8)},tplv:3,frz:true,snap:{},len:[7,7,7,7,8,8,7],target:44,auto:true,open:9,close:26,nOpen:3,nClose:3}}
var S=def(),PLAN=[],PLANN=[],PLANV=[],live=true,sel=0,tm=12,flt="all",CW=0,wk=0,VW=0,ref=null,started=false;
var $=function(i){return document.getElementById(i)};
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]})}
var needSave=false,dirty=false;
function load(o){var leg=o.agents&&o.priFrom===undefined,leg2=o.agents&&o.tpl===undefined,leg3=o.agents&&o.tpl!==undefined&&o.tplv!==3,D=def();for(var k in D)if(o[k]===undefined)o[k]=D[k];S=o;
 if(leg){S.priFrom=nowInfo().wi+1;S.zones.forEach(function(z){if(PRIZ.indexOf(z.n)>=0&&z.pri===undefined)z.pri=true});needSave=true}
 if(leg2){S.tplFrom=nowInfo().wi+1;needSave=true}
 if(leg3){S.tpl={t7:mkT(TPL7),t8:mkT(TPL8)};S.tplv=3;needSave=true}
 migrate()}
function migrate(){var z=null,ag=null;S.agents.forEach(function(a){if(a.n==="Oussama El Atrach"&&a.fo===undefined){a.fo=3;a.fw=nowInfo().wi+1}});S.zones.forEach(function(x){if(x.n==="Technopolis")z=x});S.agents.forEach(function(a){if(a.n==="Zakariae Benali")ag=a});
 if(!ag){ag={id:S.nid++,n:"Zakariae Benali",off:S.agents.length%4};S.agents.push(ag)}
 if(!z)S.zones.push({n:"Technopolis",g:3,fix:ag.id});else if(z.fix==null){z.fix=ag.id;z.g=3}}
function persist(){try{localStorage.setItem("pro1",JSON.stringify(S))}catch(e){}if(ref)ref.set({json:JSON.stringify(S)}).catch(function(e){console.error(e)})}
function save(){persist();render()}
function L(d){return S.len[d]}
function fmt(h){h=h%24;return(h<10?"0":"")+h+":00"}
function ft(t){var h=Math.floor(t),m=Math.round((t-h)*60);if(m==60){h++;m=0}h=h%24;return(h<10?"0":"")+h+":"+(m<10?"0":"")+m}
function pt(v,e,s){var p=(v||"0:0").split(":"),x=+p[0]+(+p[1]||0)/60;if(!e){if(x<S.open)x+=24}else{while(x<=s)x+=24}return x}
function dur(x){var m=Math.round(x*60),h=Math.floor(m/60);m=m%60;return h?h+"h"+(m?(m<10?"0":"")+m:""):m+" min"}
function r1(x){return Math.round(x*10)/10}
function chip(z){var q=S.zones[z]||S.zones[0];return"<span class='p-z'><i style='background:"+GC[q.g]+";color:"+GT[q.g]+"'>"+GN[q.g]+"</i>"+esc(q.n)+(q.pri?" ⭐":"")+"</span>"}
function baseOff(a,wi){var o=a.off;(a.oe||[]).forEach(function(e){if(wi>=e[0])o=e[1]});return o}
function offMap(wi){var m={},cnt=[0,0,0,0],list=S.agents.filter(function(a){return ast(a,wi)==="auto"});
 list.forEach(function(a){var k=wi+"-"+a.id;if(S.offOv[k]!==undefined)m[a.id]=S.offOv[k];else if(a.fo!=null&&wi>=(a.fw||0))m[a.id]=a.fo;else if(baseOff(a,wi)<0)m[a.id]=-1;if(m[a.id]>=0&&m[a.id]<4)cnt[m[a.id]]++});
 list.forEach(function(a){if(m[a.id]!==undefined)return;var pref=S.auto?(baseOff(a,wi)+wi)%4:baseOff(a,wi),best=0,bs=1e9;for(var d=0;d<4;d++){var sc=cnt[d]*10+((d-pref+4)%4);if(sc<bs){bs=sc;best=d}}m[a.id]=best;cnt[best]++});
 return m}
function effOff(a,wi){var m=offMap(wi);if(m[a.id]!==undefined)return m[a.id];var k=wi+"-"+a.id;if(S.offOv[k]!==undefined)return S.offOv[k];var o=baseOff(a,wi);return o<0?-1:(S.auto?(o+wi)%4:o)}
function ast(a,wi){var add=a.add!=null?a.add:a.from;if(add!=null&&wi<add)return"hidden";if(a.to!=null&&wi>=a.to)return"hidden";if(a.from!=null&&wi<a.from)return"manual";return"auto"}
function VA(wi){return S.agents.filter(function(a){return ast(a,wi)!=="hidden"})}
function rebal(eff){var i=0;S.agents.forEach(function(a){if(ast(a,eff)==="auto"&&baseOff(a,eff)>=0&&!(a.fo!=null&&eff>=(a.fw||0))){a.oe=(a.oe||[]).filter(function(e){return e[0]<eff});a.oe.push([eff,i%4]);i++}})}
function matchPri(ags,PZ,pZ,hist){var best=null,bc=1e9;
 function dfs(i,used,cost,asg){if(cost>=bc)return;if(i===ags.length){bc=cost;best=asg.slice();return}
  for(var k=0;k<PZ.length;k++){var z=PZ[k];if(used.indexOf(z)>=0)continue;var c=(pZ[ags[i].id]===z?100:0)+(hist[ags[i].id][z]||0);used.push(z);asg.push(z);dfs(i+1,used,cost+c,asg);asg.pop();used.pop()}}
 dfs(0,[],0,[]);var m={};if(best)ags.forEach(function(a,i){m[a.id]=best[i]});return m}
function gen(wi){
 var A=S.agents.filter(function(a){return ast(a,wi)==="auto"}),Mn=S.agents.filter(function(a){return ast(a,wi)==="manual"}),plan=[],OM=offMap(wi),pS={},pZ={},hist={},N=S.zones.length,seed=wi+S.week,NP=S.zones.filter(function(z){return z.pri}).length,PR=NP>0&&(S.priFrom==null||wi>=S.priFrom),TM=!!(S.tm&&S.tpl&&(S.tplFrom==null||wi>=S.tplFrom));
 S.agents.forEach(function(a){hist[a.id]={}});
 for(var d=0;d<7;d++){
  var act=A.filter(function(a){return OM[a.id]!==d}),isFx=function(a){return S.zones.some(function(z){return z.fix===a.id})},fx=act.filter(isFx),rot=act.filter(function(a){return!isFx(a)}),n=rot.length,l=L(d),day={};
  var mn=Mn.filter(function(a){var v=S.ov[wi+"-"+a.id+"-"+d];return v&&(v.s!=null||v.e!=null)});
  if(!act.length&&!mn.length){plan.push(day);continue}
  var O=S.open,C=S.close,M=[12,14,16,18,18,10,12,18],slots=[],ne=Math.max(S.nOpen,PR?NP:0),nc=Math.max(S.nClose,PR?NP:0);
  while((ne>0||nc>0)&&slots.length<n){if(ne>0){slots.push(O);ne--}if(nc>0&&slots.length<n){slots.push(C-l);nc--}}
  for(var m=0;slots.length<n;m++)slots.push(M[m%M.length]);
  slots.sort(function(a,b){return a-b});slots=slots.map(function(v){return Math.max(O,Math.min(v,C-l))});
  var o=n?(d*3+seed*5)%n:0,st={},en={};
  rot.forEach(function(a,j){st[a.id]=slots[(j+o)%n]});
  fx.forEach(function(a){st[a.id]=Math.max(O,Math.min(d%2?11:12,C-l))});
  rot.forEach(function(a){if(st[a.id]===pS[a.id]){for(var k=0;k<rot.length;k++){var b=rot[k];if(b!==a&&st[b.id]!==pS[a.id]&&st[a.id]!==pS[b.id]&&st[b.id]!==st[a.id]){var t=st[a.id];st[a.id]=st[b.id];st[b.id]=t;break}}}});
  var forcedT={};
  if(TM&&n){
   var tl=l>=8?S.tpl.t8:S.tpl.t7,sl=[];for(var tq=0;tq<n;tq++){if(tl.length)sl.push(tl[tq%tl.length])}
   if(sl.length===n){
    var zi=function(nm){var r=0;S.zones.forEach(function(z,i){if(z.n===nm)r=i});return r};
    var tcost=function(a,e){var z=zi(e.z),c=0;if(pZ[a.id]===z)c+=1000;if(pS[a.id]===e.s)c+=600;c+=(hist[a.id][z]||0)*8;return c};
    var asg=rot.slice(),o2=(d*3+seed*5)%n;asg=asg.slice(o2).concat(asg.slice(0,o2));
    var imp=true,it=0;
    while(imp&&it<40){imp=false;it++;for(var ti=0;ti<n;ti++)for(var tj=ti+1;tj<n;tj++){var b1=tcost(asg[ti],sl[ti])+tcost(asg[tj],sl[tj]),b2=tcost(asg[ti],sl[tj])+tcost(asg[tj],sl[ti]);if(b2<b1){var tt=asg[ti];asg[ti]=asg[tj];asg[tj]=tt;imp=true}}}
    asg.forEach(function(a,i){st[a.id]=sl[i].s;forcedT[a.id]=zi(sl[i].z)})
   }
  }
  act.concat(mn).forEach(function(a){var v=S.ov[wi+"-"+a.id+"-"+d]||{};if(v.s!=null)st[a.id]=v.s;else if(st[a.id]===undefined)st[a.id]=v.e-l;en[a.id]=st[a.id]+l;if(v.e!=null)en[a.id]=v.e});
  var ord=act.concat(mn).sort(function(a,b){return st[a.id]-st[b.id]}),ZC=[],CC=[[],[],[],[]];
  var forced={},rs={};rot.forEach(function(a){rs[a.id]=1});
  if(PR&&!TM){var PZ=[];S.zones.forEach(function(z,i){if(z.pri)PZ.push(i)});
   var nz=function(a){return!(S.ov[wi+"-"+a.id+"-"+d]||{}).z!=null&&((S.ov[wi+"-"+a.id+"-"+d]||{}).z==null)};
   var op=ord.filter(function(a){return rs[a.id]&&st[a.id]<=O+.01&&nz(a)}).slice(0,NP),cl=ord.filter(function(a){return rs[a.id]&&en[a.id]>=C-.01&&nz(a)}).slice(0,NP),m1=matchPri(op,PZ,pZ,hist),m2=matchPri(cl,PZ,pZ,hist);
   for(var k1 in m1)forced[k1]=m1[k1];for(var k2 in m2)forced[k2]=m2[k2]}
  for(var k3 in forcedT)forced[k3]=forcedT[k3];
  ord.forEach(function(a,j){
   var v=S.ov[wi+"-"+a.id+"-"+d]||{},s0=st[a.id],e0=en[a.id],best=0,bs=1e9,q;
   var fz=-1;S.zones.forEach(function(zz,zi){if(zz.fix===a.id)fz=zi});
   if(fz>=0)best=fz;else if(forced[a.id]!==undefined)best=forced[a.id];else if(v.z!=null&&v.z<N&&S.zones[v.z].fix==null)best=v.z;
   else for(var z=0;z<N;z++){if(S.zones[z].fix!=null)continue;
    var sc=(z===pZ[a.id]?1e6:0)+(hist[a.id][z]||0)*3+((z-j-d-seed+900)%N)*0.1,g=S.zones[z].g;
    for(q=Math.floor(s0);q<e0;q++){var zn=(ZC[z]&&ZC[z][q])||0,cn=CC[g][q]||0;sc+=zn*3+cn+(cn===0?-20:0);if(PR&&S.zones[z].pri&&!zn&&(q===Math.floor(S.open)||q===Math.ceil(S.close)-1))sc-=100}
    if(sc<bs){bs=sc;best=z}}
   if(!ZC[best])ZC[best]=[];
   for(q=Math.floor(s0);q<e0;q++){ZC[best][q]=(ZC[best][q]||0)+1;var gg=S.zones[best].g;CC[gg][q]=(CC[gg][q]||0)+1}
   hist[a.id][best]=(hist[a.id][best]||0)+1;day[a.id]={s:s0,e:e0,z:best};pZ[a.id]=best;pS[a.id]=s0;
  });
  plan.push(day);
 }
 return plan;
}
function nowInfo(){var n=new Date(),h=n.getHours()+n.getMinutes()/60,b=new Date(n.getTime()-72e5),d=(b.getDay()+6)%7,wi=Math.floor(((b.getTime()-b.getTimezoneOffset()*6e4)/864e5+3)/7);if(h<2)h+=24;return{d:d,t:h,wi:wi}}
function wlab(wi){var a=new Date((wi*7-3)*864e5);a=new Date(a.getTime()+a.getTimezoneOffset()*6e4);var b=new Date(a.getTime()+6*864e5),f=function(x){return(x.getDate()<10?"0":"")+x.getDate()+"/"+(x.getMonth()<9?"0":"")+(x.getMonth()+1)};return"Semaine du "+f(a)+" au "+f(b)}
function dlab(wi,d){var a=new Date((wi*7-3)*864e5);a=new Date(a.getTime()+a.getTimezoneOffset()*6e4+d*864e5);return(a.getDate()<10?"0":"")+a.getDate()+"/"+(a.getMonth()<9?"0":"")+(a.getMonth()+1)}
function tick(){if(!started)return;var i=nowInfo();if(i.wi!==CW){render();return}if(live){sel=i.d;tm=i.t}renderLive()}
function stat(a,d,t){var c=PLAN[d]&&PLAN[d][a.id];if(!c)return"off";return t>=c.s&&t<c.e?"on":t<c.s?"soon":"done"}
function renderLive(){
 var d=sel,t=tm,cnt={on:0,soon:0,done:0,off:0},on=[];
 VA(CW).forEach(function(a){var s=stat(a,d,t);cnt[s]++;if(s==="on")on.push({a:a,c:PLAN[d][a.id]})});
 $("p-lv").className="p-live"+(live?"":" m");$("p-lv").textContent=live?"● En direct":"Mode manuel";
 var tb="";DAYS.forEach(function(x,i){tb+="<button class='"+(i===d?"a":"")+"' onclick='PRO.setD("+i+")'>"+x+" "+dlab(CW,i)+"</button>"});
 $("p-tabs").innerHTML=tb;$("p-wl").textContent=wlab(CW);$("p-tm").textContent=ft(t);
 $("p-sl").min=S.open;$("p-sl").max=S.close;$("p-sl").value=Math.min(Math.max(t,S.open),S.close);
 var sc=function(c,l,v){return"<div class='stat-card "+c+"'><div class='label'>"+l+"</div><div class='value'>"+v+"</div></div>"};
 $("p-st").innerHTML=sc("accent-green","En service",cnt.on)+sc("accent-amber","Pas encore commencé",cnt.soon)+sc("","Terminé",cnt.done)+sc("accent-purple","OFF",cnt.off);
 var g="";
 for(var k=0;k<4;k++){
  var zs=S.zones.filter(function(z){return z.g===k});if(!zs.length)continue;
  var m=on.filter(function(o){return S.zones[o.c.z].g===k}),b;
  if(k===3)b=m.length?"<span class='p-bd' style='background:var(--gs);color:#007A63'>En service</span>":"<span class='p-bd' style='background:rgba(110,110,104,.14);color:#6B5D3D'>Hors service</span>";
  else if(!m.length)b="<span class='p-bd' style='background:rgba(220,38,38,.12);color:#DC2626'>Non couvert</span>";
  else if(m.length<zs.length)b="<span class='p-bd' style='background:var(--ys);color:#8a5a00'>"+m.length+" agent"+(m.length>1?"s":"")+" · couvrent ensemble</span>";
  else b="<span class='p-bd' style='background:var(--gs);color:#007A63'>1 agent / zone</span>";
  g+="<div class='card' style='padding:14px;border-top:4px solid "+GC[k]+"'><div class='p-h'><b>"+(k===3?"Zone fixe":"Groupe "+GN[k])+"</b>"+b+"</div><div class='mu' style='margin:4px 0 6px'>"+zs.map(function(z){return esc(z.n)}).join(" · ")+"</div>"+m.map(function(o){return"<div style='margin-top:4px'>"+esc(o.a.n)+" "+chip(o.c.z)+"</div>"}).join("")+"</div>";
 }
 $("p-grp").innerHTML=g;
 var O=S.open,NH=Math.ceil(S.close-S.open);$("p-gt").style.setProperty("--nh",NH);
 var gt="<div class='p-gr'><div></div><div class='p-gh' style='grid-template-columns:repeat("+NH+",1fr)'>";
 for(var x=O;x<O+NH;x++)gt+="<span>"+(x%24)+"h</span>";gt+="</div></div>";
 var nl=t>=O&&t<=O+NH?"<i class='p-nl' style='left:"+(t-O)/NH*100+"%'></i>":"";
 VA(CW).forEach(function(a){
  var c=PLAN[d][a.id];gt+="<div class='p-gr'><div class='p-gn'>"+esc(a.n)+"</div><div class='p-gc'>";
  if(c)gt+="<div class='p-gb' style='left:"+Math.max(0,(c.s-O)/NH*100)+"%;width:"+(c.e-c.s)/NH*100+"%'>"+esc(S.zones[c.z].n)+" · "+ft(c.s)+"–"+ft(c.e)+"</div>";
  else gt+="<span class='mu' style='padding:8px;position:absolute'>OFF</span>";
  gt+=nl+"</div></div>";
 });
 $("p-gt").innerHTML=gt;
 var F=[["all","Tous",VA(CW).length],["on","🟢 En service",cnt.on],["soon","🟡 Pas commencé",cnt.soon],["done","⚪ Terminé",cnt.done],["off","OFF",cnt.off]],fh="";
 F.forEach(function(f){fh+="<button class='"+(flt===f[0]?"a":"")+"' onclick=\"PRO.setF('"+f[0]+"')\">"+f[1]+" ("+f[2]+")</button>"});
 $("p-fl").innerHTML=fh;
 var order={on:0,soon:1,done:2,off:3},list=VA(CW).map(function(a){return{a:a,s:stat(a,d,t),c:PLAN[d][a.id]}}).filter(function(o){return flt==="all"||o.s===flt});
 list.sort(function(x,y){return order[x.s]-order[y.s]||(x.c&&y.c?x.c.s-y.c.s:0)});
 var ch="",nxP=d==6?PLANN:PLAN,nd=(d+1)%7;
 list.forEach(function(o){
  var a=o.a,c=o.c,w=0;
  for(var i=0;i<d;i++)if(PLAN[i][a.id])w+=PLAN[i][a.id].e-PLAN[i][a.id].s;
  if(c)w+=Math.max(0,Math.min(t,c.e)-c.s);
  var nx=nxP[nd][a.id],bd={on:["En service","var(--gs)","#007A63"],soon:["Pas encore commencé","var(--ys)","#8a5a00"],done:["Terminé","rgba(110,110,104,.14)","#6B5D3D"],off:["OFF","#EDDFB8","#6B5D3D"]}[o.s];
  ch+="<div class='p-ac "+o.s+"'><div class='p-h'><b>"+esc(a.n)+"</b><span class='p-bd' style='background:"+bd[1]+";color:"+bd[2]+"'>"+bd[0]+"</span></div>";
  if(c){ch+=chip(c.z)+"<div class='mu' style='margin-top:4px'>"+ft(c.s)+" → "+ft(c.e)+"</div>";
   if(o.s==="on")ch+="<div class='p-bar'><i style='width:"+Math.round((t-c.s)/(c.e-c.s)*100)+"%'></i></div><div class='mu'>Reste "+dur(c.e-t)+"</div>";
   else if(o.s==="soon")ch+="<div class='mu'>Commence dans "+dur(c.s-t)+"</div>";
   else ch+="<div class='mu'>Fini à "+ft(c.e)+"</div>"}
  else ch+="<div class='mu' style='margin-top:4px'>Jour de repos</div>";
  ch+="<div class='mu' style='margin-top:6px;border-top:1px solid var(--b);padding-top:5px'>Semaine : <b>"+r1(w)+"h</b> / "+S.target+"h<br>Demain : "+(nx?ft(nx.s)+" → "+ft(nx.e)+" · "+esc(S.zones[nx.z].n):"OFF")+"</div></div>";
 });
 $("p-ags").innerHTML=ch||"<div class='mu'>Aucun agent</div>";
}
function renderCv(){
 var v="<tr><th>Jour</th>";
 for(var x=Math.floor(S.open);x<S.close;x++)v+="<th>"+(x%24)+"h</th>";v+="</tr>";
 for(var d=0;d<7;d++){
  v+="<tr><th>"+DAYS[d]+" "+dlab(VW,d)+"</th>";
  for(var x=Math.floor(S.open);x<S.close;x++){
   var n=0,cc=[0,0,0,0];for(var id in PLANV[d]){var c=PLANV[d][id];if(c.s<=x+.01&&x<c.e){n++;cc[S.zones[c.z].g]++}}
   var rush=(x>=12&&x<15)||(x>=18&&x<25),cur=(VW===CW&&d===sel&&Math.floor(tm)===x);
   v+="<td class='"+(cur?"p-cur":"")+"' style='background:rgba(0,160,130,"+Math.min(n/10,1)*(rush?.25:.08)+")'><b>"+n+"</b><br>"+cc.map(function(k,i){return"<span style='"+(k||i>2?"color:#6B5D3D":"color:#DC2626;font-weight:700")+"'>"+k+"</span>"}).join("·")+"</td>";
  }
  v+="</tr>";
 }
 $("p-cv").innerHTML=v;
}
function zopts(z,id){var fx=null;S.zones.forEach(function(q,i){if(q.fix===id)fx=i});return S.zones.map(function(q,i){if(fx!==null?i!==fx:q.fix!=null)return"";return"<option value='"+i+"'"+(i===z?" selected":"")+">"+GN[q.g]+" · "+esc(q.n)+"</option>"}).join("")}
function ovk(id,d){return VW+"-"+id+"-"+d}
function render(){
 if(!started)return;
 CW=nowInfo().wi;VW=CW+wk;PLAN=getPlan(CW);PLANN=getPlan(CW+1);PLANV=wk===0?PLAN:getPlan(VW);
 $("p-wv").textContent=wlab(VW);$("p-ws").textContent=wk===0?"Cette semaine":"";
 var h="<tr><th>Agent</th><th>OFF</th>",TD=VW===CW?nowInfo().d:-1;
 DAYS.forEach(function(x,d){h+="<th"+(d===TD?" class='td'":"")+">"+x+"<br><b style='font-size:15px'>"+dlab(VW,d)+"</b> <span style='opacity:.7;font-weight:500'>("+L(d)+"h)</span></th>"});h+="<th>Total</th></tr>";
 var vn=0;S.agents.forEach(function(a,i){
  var as=ast(a,VW);if(as==="hidden")return;vn++;var tot=0,eo=(S.frz&&as==="auto")?curOff(a):effOff(a,VW);
  h+="<tr><td class='p-ag'><b>"+vn+". "+esc(a.n)+"</b>"+(as==="manual"?" <span class='p-bd' style='background:var(--ys);color:#8a5a00'>horaires manuels</span>":"")+"<br><button onclick='PRO.mv("+i+",-1)'>↑</button><button onclick='PRO.mv("+i+",1)'>↓</button><button onclick='PRO.del("+i+")'>×</button><br><button title='Lun–Jeu 09:00-16:00 · Ven 09:00-17:00 · Sam 10:00-18:00 · Dim 10:00-17:00' onclick='PRO.preset("+a.id+",\"J\")'>☀️ Matin</button><button title='Lun–Jeu 18:00-01:00 · Ven 17:00-01:00 · Sam 18:00-02:00 · Dim 18:00-01:00' onclick='PRO.preset("+a.id+",\"N\")'>🌙 Nuit</button></td><td><select onchange='PRO.setOff("+a.id+",this.value)'>";
  for(var k=0;k<7;k++)h+="<option value='"+k+"'"+(eo===k?" selected":"")+">"+DAYS[k]+"</option>";
  h+="<option value='-1'"+(eo<0?" selected":"")+">Aucun</option></select><button title='OFF fixe chaque semaine' style='margin-top:4px;padding:0 7px;font-size:12px"+(a.fo!=null?";background:var(--y)":"")+"' onclick='PRO.fo("+a.id+")'>📌 "+(a.fo!=null?"fixe":"fixer")+"</button></td>";
  for(var d=0;d<7;d++){
   var c=PLANV[d][a.id];if(!c){if(as==="manual")h+="<td class='c'><div class='p-tt'><input type='time' onchange=\"PRO.setTime("+a.id+","+d+",'s',this.value)\"><input type='time' onchange=\"PRO.setTime("+a.id+","+d+",'e',this.value)\"></div><span class='mu'>+ horaire</span></td>";else h+="<td class='o'>OFF</td>";continue}
   tot+=c.e-c.s;var mod=S.frz?c.m:S.ov[ovk(a.id,d)],sc=c.s<11?" sm":c.s>=16?" sn":" c";
   h+="<td class='"+sc+(mod?" m":"")+"'><div class='p-tt'><input type='time' value='"+ft(c.s)+"' onchange=\"PRO.setTime("+a.id+","+d+",'s',this.value)\"><input type='time' value='"+ft(c.e)+"' onchange=\"PRO.setTime("+a.id+","+d+",'e',this.value)\"></div><select style='border-left-color:"+GC[S.zones[c.z].g]+"' onchange='PRO.setZ("+a.id+","+d+",this.value)'>"+zopts(c.z,a.id)+"</select>"+(mod?"<button style='padding:0 7px;margin-top:4px;font-size:12px' onclick='PRO.clr("+a.id+","+d+")'>↺</button>":"")+"</td>";
  }
  var df=tot-S.target,over=as==="auto"&&Math.abs(df)>.01;
  h+="<td class='"+(over?"p-bad":"")+"'>"+r1(tot)+"h"+(over?" ⚠ "+(df>0?"+":"")+r1(df)+"h":"")+(as==="manual"?"<br><span class='mu'>(manuel)</span>":"")+"</td></tr>";
 });
 $("p-t").innerHTML=h;
 var zg="";for(var gk=0;gk<4;gk++){var zz=[];S.zones.forEach(function(z,i){if(z.g===gk)zz.push(i)});if(zz.length)zg+="<span style='display:inline-flex;flex-wrap:wrap;gap:4px;align-items:center'><b>"+(gk===3?"Zone fixe":"Groupe "+GN[gk])+" :</b> "+zz.map(chip).join(" ")+"</span>"}
 $("p-zg").innerHTML=zg;
 var zl="";S.zones.forEach(function(z,i){zl+="<div class='p-zr'><input value=\""+esc(z.n)+"\" onchange='PRO.zName("+i+",this.value)'><select onchange='PRO.zGrp("+i+",this.value)'>"+[0,1,2,3].map(function(g){return"<option value='"+g+"'"+(z.g===g?" selected":"")+">"+(g===3?"Zone fixe":"Groupe "+GN[g])+"</option>"}).join("")+"</select><select onchange='PRO.zFix("+i+",this.value)'><option value=''>Fixe : personne</option>"+VA(VW).map(function(a){return"<option value='"+a.id+"'"+(z.fix===a.id?" selected":"")+">Fixe : "+esc(a.n)+"</option>"}).join("")+"</select><button title='Prioritaire : un agent à l’ouverture et à la fermeture' onclick='PRO.zPri("+i+")'>"+(z.pri?"⭐":"☆")+"</button><button onclick='PRO.zDel("+i+")'>×</button></div>"});
 $("p-zl").innerHTML=zl;
 $("p-dl").innerHTML=DAYS.map(function(x,d){return"<label class='mu'>"+x+"<input type='number' step='0.5' value='"+S.len[d]+"' onchange='PRO.setLen("+d+",this.value)'></label>"}).join("");
 $("p-fz").checked=!!S.frz;$("p-tg").value=S.target;$("p-au").checked=!!S.auto;$("p-so").value=ft(S.open);$("p-sc").value=ft(S.close);$("p-no").value=S.nOpen;$("p-nc").value=S.nClose;
 var tp="<b>Modèle fixe des journées</b><div class='mu' style='margin:2px 0 8px'>L’heure et la zone de chaque créneau restent toujours pareils : seuls les agents changent chaque jour. S’il y a moins d’agents (jours OFF), on utilise les premiers créneaux de la liste. S’applique à partir de la semaine prochaine.</div><label style='display:block;margin-bottom:10px'><input type='checkbox' "+(S.tm?"checked":"")+" style='width:auto' onchange='PRO.setTm(this.checked)'> Utiliser le modèle fixe</label><div class='p-sg' style='margin-top:0'>";
 [["t7","Jours de 7h (Lun–Jeu, Dim)"],["t8","Jours de 8h (Ven, Sam)"]].forEach(function(T){
  tp+="<div><div style='font-weight:700;margin-bottom:6px'>"+T[1]+" · "+S.tpl[T[0]].length+" créneaux</div>";
  S.tpl[T[0]].forEach(function(e,i){tp+="<div class='p-tr'><input type='time' value='"+ft(e.s)+"' onchange=\"PRO.tS('"+T[0]+"',"+i+",this.value)\"><select onchange=\"PRO.tZ('"+T[0]+"',"+i+",this.value)\">"+S.zones.map(function(z,zi){return"<option value='"+zi+"'"+(z.n===e.z?" selected":"")+">"+GN[z.g]+" · "+esc(z.n)+"</option>"}).join("")+"</select><button onclick=\"PRO.tD('"+T[0]+"',"+i+")\">×</button></div>"});
  tp+="<button onclick=\"PRO.tA('"+T[0]+"')\">+ créneau</button></div>"});
 $("p-tpl").innerHTML=tp+"</div>";
 $("p-alert").innerHTML=alerts(PLANV,VW);$("p-alert2").innerHTML=alerts(PLAN,CW);
 renderCv();tick();for(var sk in S.snap)if(+sk<CW-10)delete S.snap[sk];if(dirty){dirty=false;persist()}
}
function snapRoster(wi){return S.agents.filter(function(a){return ast(a,wi)!=="hidden"}).map(function(a){return a.id+":"+ast(a,wi)}).join(",")}
function toStore(pl){return pl.map(function(day){var o={};for(var id in day){var c=day[id];o[id]={s:c.s,e:c.e,zn:S.zones[c.z]?S.zones[c.z].n:S.zones[0].n}}return o})}
function toView(sp){return sp.map(function(day){var o={};for(var id in day){var c=day[id],z=0;S.zones.forEach(function(q,i){if(q.n===c.zn)z=i});o[id]={s:c.s,e:c.e,z:z,m:c.m}}return o})}
function getPlan(wi){
 if(!S.frz)return gen(wi);
 var sn=S.snap[wi];
 if(!sn){var prev=S.snap[wi-1],key=snapRoster(wi),sp;
  if(prev&&prev.r===key){sp=JSON.parse(JSON.stringify(prev.p));sp.forEach(function(day){for(var id in day)delete day[id].m})}
  else sp=toStore(gen(wi));
  sn=S.snap[wi]={r:key,p:sp};dirty=true}
 return toView(sn.p)}
function snapEdit(wi){getPlan(wi);return S.snap[wi]}
function resnap(from){from=from==null?CW+1:from;for(var k in S.snap)if(+k>=from)delete S.snap[k]}
function dropLater(wi){for(var k in S.snap){if(+k>wi){var hm=false;S.snap[k].p.forEach(function(day){for(var id in day)if(day[id].m)hm=true});if(!hm)delete S.snap[k]}}}
function defZoneName(){var z=S.zones.filter(function(q){return q.fix==null})[0];return(z||S.zones[0]).n}
function curOff(a){for(var d=0;d<7;d++)if(!PLANV[d][a.id])return d;return -1}
function fzWeek(){resnap(VW);for(var k in S.ov)if(k.indexOf(VW+"-")===0)delete S.ov[k];for(var k2 in S.offOv)if(k2.indexOf(VW+"-")===0)delete S.offOv[k2];S.snap[VW]={r:snapRoster(VW),p:toStore(gen(VW))};dirty=true}
function fzTime(id,d,w,v){var sn=snapEdit(VW),c=sn.p[d][id];
 if(w==="s"){var ns=pt(v,0);if(c){var du=c.e-c.s;c.s=ns;c.e=ns+du}else{sn.p[d][id]=c={s:ns,e:ns+L(d),zn:defZoneName()}}}
 else{var ne=pt(v,1,c?c.s:S.open);if(c)c.e=ne;else sn.p[d][id]=c={s:ne-L(d),e:ne,zn:defZoneName()}}
 c.m=true}
function fzZ(id,d,v){var c=snapEdit(VW).p[d][id];if(c){c.zn=S.zones[+v].n;c.m=true}}
function fzClr(id,d){var sn=snapEdit(VW),g=toStore(gen(VW))[d][id];if(g)sn.p[d][id]=g;else delete sn.p[d][id]}
function fzOff(id,v){var a=null;S.agents.forEach(function(x){if(x.id===id)a=x});if(!a||ast(a,VW)!=="auto")return;
 var P=snapEdit(VW).p,old=-1,dd;for(dd=0;dd<7;dd++)if(!P[dd][id]){old=dd;break}v=+v;
 if(v<0){if(old>=0){var ref=null;for(dd=0;dd<7&&!ref;dd++)if(P[dd][id]&&L(dd)===L(old))ref=P[dd][id];P[old][id]=ref?{s:ref.s,e:ref.s+L(old),zn:ref.zn,m:true}:{s:S.open,e:S.open+L(old),zn:defZoneName(),m:true}}}
 else{if(old===v)return;var c=P[v][id];if(old>=0&&c)P[old][id]={s:c.s,e:c.s+L(old),zn:c.zn,m:true};delete P[v][id]}}
var PJ=[[9,16],[9,16],[9,16],[9,16],[9,17],[10,18],[10,17]],PN=[[18,25],[18,25],[18,25],[18,25],[17,25],[18,26],[18,25]];
function fzPreset(id,t){var pre=t==="J"?PJ:PN,sn=snapEdit(VW),a=null;S.agents.forEach(function(x){if(x.id===id)a=x});
 var td=VW===CW?nowInfo().d:0;
 for(var d=0;d<7;d++){var c=sn.p[d][id];
  if(c){c.s=pre[d][0];c.e=pre[d][1];c.m=true}
  else if(a&&ast(a,VW)==="manual"&&d>=td)sn.p[d][id]={s:pre[d][0],e:pre[d][1],zn:defZoneName(),m:true}}}
function total(pl,a){var t=0;pl.forEach(function(d){if(d[a.id])t+=d[a.id].e-d[a.id].s});return t}
function change(fn,id){
 getPlan(VW);var bak=JSON.stringify([S.ov,S.offOv,S.snap]);fn();
 var pl=getPlan(VW),msg=[],a=null;
 S.agents.forEach(function(x){if(x.id===id)a=x});
 if(VW<=CW)msg.push("⚠️ Tu modifies une semaine EN COURS (ou passée).");
 if(a&&ast(a,VW)==="auto"){var t=total(pl,a);if(Math.abs(t-S.target)>.01)msg.push("⚠️ "+a.n+" aura "+r1(t)+"h au lieu de "+S.target+"h ("+(t>S.target?"+":"")+r1(t-S.target)+"h).")}
 if(msg.length&&!confirm(msg.join("\n")+"\n\nValider la modification ?")){var o=JSON.parse(bak);S.ov=o[0];S.offOv=o[1];S.snap=o[2];render();return}
 if(S.frz)dropLater(VW);save();
}
function alerts(pl,wi){var r=[];VA(wi).forEach(function(a){if(ast(a,wi)!=="auto")return;var t=total(pl,a);if(Math.abs(t-S.target)>.01)r.push("<b>"+esc(a.n)+"</b> : "+r1(t)+"h ("+(t>S.target?"+":"")+r1(t-S.target)+"h)")});return r.length?"<div class='p-alert'>⚠️ <b>"+r.length+" agent(s) différent(s) de "+S.target+"h</b> — "+r.join(" · ")+"</div>":""}
var PRO={
 setD:function(d){live=false;sel=d;renderLive()},setT:function(v){live=false;tm=+v;renderLive()},setF:function(f){flt=f;renderLive()},
 live:function(){live=true;tick()},wk:function(k){wk=k===0?0:wk+k;render()},
 add:function(){var v=$("p-nn").value.trim();if(!v)return;var eff=Math.max(VW,CW+1),add=VW>CW?VW:CW;
  if(add<eff&&!confirm("Ajouter "+v+" ?\n\n• Planifié automatiquement à partir de : "+wlab(eff)+"\n• Pour la semaine en cours : tu remplis ses horaires toi-même dans le tableau.\n• Les autres agents ne changent pas cette semaine."))return;
  S.agents.push({id:S.nid++,n:v,off:0,add:add,from:eff});rebal(eff);resnap(eff);$("p-nn").value="";save()},
 del:function(i){var a=S.agents[i],eff=Math.max(VW,CW+1),isNew=a.add!=null&&a.add>=CW;
  var m=isNew?"Supprimer "+a.n+" (ajouté récemment) ?":"Retirer "+a.n+" ?\n\n• Il reste dans la semaine en cours et celles d’avant.\n• Il disparaît à partir de : "+wlab(eff)+"\n• Le planning est régénéré à partir de cette semaine, sans lui.";
  if(!confirm(m))return;
  if(isNew){for(var k in S.ov)if(k.indexOf("-"+a.id+"-")>0)delete S.ov[k];S.agents.splice(i,1)}else a.to=eff;
  rebal(eff);resnap(eff);save()},
 mv:function(i,k){var a=S.agents,j=i+k;while(j>=0&&j<a.length&&ast(a[j],VW)==="hidden")j+=k;if(j<0||j>=a.length)return;var t=a[i];a[i]=a[j];a[j]=t;save()},
 fo:function(id){var a=null;S.agents.forEach(function(x){if(x.id===id)a=x});if(!a)return;if(a.fo!=null){a.fo=null}else{a.fo=effOff(a,VW);a.fw=Math.max(VW,CW+1);alert(a.n+" : OFF fixe chaque "+DAYS[a.fo]+" à partir de "+wlab(a.fw)+".")}rebal(Math.max(VW,CW+1));resnap(Math.max(VW,CW+1));save()},
 preset:function(id,t){change(function(){if(S.frz)return fzPreset(id,t);var pre=t==="J"?PJ:PN;for(var d=0;d<7;d++)if(PLANV[d][id]){var o=S.ov[ovk(id,d)]||(S.ov[ovk(id,d)]={});o.s=pre[d][0];o.e=pre[d][1]}},id)},
 setOff:function(id,v){change(function(){if(S.frz)return fzOff(id,v);S.offOv[VW+"-"+id]=+v},id)},
 setTime:function(id,d,w,v){change(function(){if(S.frz)return fzTime(id,d,w,v);var o=S.ov[ovk(id,d)]||(S.ov[ovk(id,d)]={}),c=PLANV[d][id];if(w==="s")o.s=pt(v,0);else o.e=pt(v,1,c?c.s:(o.s!=null?o.s:S.open))},id)},
 setZ:function(id,d,v){change(function(){if(S.frz)return fzZ(id,d,v);(S.ov[ovk(id,d)]||(S.ov[ovk(id,d)]={})).z=+v},id)},
 clr:function(id,d){change(function(){if(S.frz)return fzClr(id,d);delete S.ov[ovk(id,d)]},id)},
 clrWeek:function(){change(function(){if(S.frz)return fzWeek();for(var k in S.ov)if(k.indexOf(VW+"-")===0)delete S.ov[k];for(var k in S.offOv)if(k.indexOf(VW+"-")===0)delete S.offOv[k]},0)},
 regen:function(){if(S.frz&&VW<=CW&&!confirm("Régénérer la semaine en cours ? Toutes ses heures et zones seront recalculées."))return;S.week++;if(S.frz)fzWeek();else resnap();save()},reset:function(){if(confirm("Tout réinitialiser ?")){S=def();save()}},
 zName:function(i,v){if(!v.trim())return;var old=S.zones[i].n,nn=v.trim();S.zones[i].n=nn;for(var k in S.snap)S.snap[k].p.forEach(function(day){for(var id in day)if(day[id].zn===old)day[id].zn=nn});["t7","t8"].forEach(function(t){S.tpl[t].forEach(function(e){if(e.z===old)e.z=nn})});save()},setTm:function(c){S.tm=c;S.tplFrom=CW+1;resnap();save()},
 tS:function(k,i,v){S.tpl[k][i].s=pt(v,0);S.tplFrom=CW+1;resnap();save()},tZ:function(k,i,v){S.tpl[k][i].z=S.zones[+v].n;S.tplFrom=CW+1;resnap();save()},
 tD:function(k,i){if(S.tpl[k].length>1){S.tpl[k].splice(i,1);S.tplFrom=CW+1;resnap();save()}},tA:function(k){var l=S.tpl[k],e=l[l.length-1];l.push({s:e.s,z:e.z});S.tplFrom=CW+1;resnap();save()},
 zPri:function(i){S.zones[i].pri=!S.zones[i].pri;S.priFrom=CW+1;resnap();save()},zGrp:function(i,v){S.zones[i].g=+v;save()},zFix:function(i,v){var z=S.zones[i];z.fix=v===""?null:+v;if(z.fix!=null)z.g=3;save()},
 zDel:function(i){if(S.zones.length>1){var old=S.zones[i].n;S.zones.splice(i,1);var fb=defZoneName();for(var k in S.snap)S.snap[k].p.forEach(function(day){for(var id in day)if(day[id].zn===old)day[id].zn=fb});["t7","t8"].forEach(function(t){S.tpl[t].forEach(function(e){if(e.z===old)e.z=fb})});save()}},zAdd:function(){S.zones.push({n:"Nouvelle zone",g:0});save()},
 setLen:function(d,v){S.len[d]=Math.max(1,Math.min(17,+v||7));resnap();save()},
 setTarget:function(v){S.target=+v||44;save()},setAuto:function(c){S.auto=c;resnap();save()},setFrz:function(c){S.frz=c;save()},
 setSvc:function(){var a=$("p-so").value.split(":"),b=$("p-sc").value.split(":"),o=+a[0]+(+a[1]||0)/60,c=+b[0]+(+b[1]||0)/60;while(c<=o)c+=24;if(c-o>24)c=o+24;S.open=o;S.close=c;resnap();save()},
 setN:function(){S.nOpen=Math.max(0,+$("p-no").value||0);S.nClose=Math.max(0,+$("p-nc").value||0);resnap();save()},
 _t:{render:render,gen:gen,getPlan:getPlan,S:function(){return S},start:function(){started=true},cw:function(){return CW}}
};
window.PRO=PRO;
var DASH='<div class="card" style="padding:16px"><div class="p-tl"><div class="p-tm" id="p-tm"></div><input type="range" id="p-sl" step="0.25" oninput="PRO.setT(this.value)"><button class="btn-ghost" onclick="PRO.live()">● Maintenant</button><span id="p-lv" class="p-live"></span></div><div class="p-row" id="p-tabs"></div><div class="mu" id="p-wl"></div></div><div class="stat-grid" id="p-st" style="margin-top:16px"></div><div id="p-alert2"></div><div class="section-title">Groupes de zones</div><div class="p-grp" id="p-grp"></div><div class="section-title">Journée complète</div><div class="p-w"><div class="p-gi" id="p-gt"></div></div><div class="section-title">Tous les agents</div><div class="p-row" id="p-fl"></div><div class="p-ags" id="p-ags"></div>';
var PLN='<div class="planning-header"><div class="week-nav"><button class="icon-btn-lg" onclick="PRO.wk(-1)">‹</button><div class="week-label"><b id="p-wv"></b><span id="p-ws"></span></div><button class="icon-btn-lg" onclick="PRO.wk(1)">›</button><button class="btn-ghost" style="border:1px solid var(--b)" onclick="PRO.wk(0)">Cette semaine</button></div><div class="planning-actions"><input id="p-nn" placeholder="Nom de l’agent"><button class="btn-primary" onclick="PRO.add()">+ Ajouter un agent</button><button class="btn-ghost" style="border:1px solid var(--b)" onclick="PRO.regen()">↻ Autre rotation</button><button class="btn-ghost" style="border:1px solid var(--b)" onclick="PRO.clrWeek()">Effacer mes modifications</button><button class="btn-ghost" style="border:1px solid var(--b)" onclick="PRO.reset()">Réinitialiser</button></div></div><div class="mu">Le planning reste figé : rien ne change tout seul. Une nouvelle semaine reprend la semaine précédente, et tu modifies toi-même les heures, les zones et les jours OFF dans le tableau.</div><div class="p-lgd"><span><i style="background:rgba(255,194,68,.5)"></i>Matin (avant 11h)</span><span><i style="background:rgba(120,105,60,.1)"></i>Journée</span><span><i style="background:rgba(0,160,130,.24)"></i>Soir (16h ou après)</span><span><i style="box-shadow:inset 0 0 0 2px #00A082"></i>Modifié à la main (↺ pour annuler)</span></div><div id="p-alert"></div><div class="p-row" id="p-zg" style="gap:8px 14px"></div><div class="p-w"><table class="p-t" id="p-t"></table></div><div class="p-sg"><div class="card" style="padding:14px"><b>Zones</b><div class="mu" style="margin:2px 0 8px">⭐ = zone prioritaire : un agent y est toujours présent à l’ouverture et à la fermeture. Change le nom, le groupe ou ajoute/supprime une zone.</div><div id="p-zl"></div><button onclick="PRO.zAdd()">+ Ajouter une zone</button></div><div class="card" style="padding:14px"><b>Heures</b><div class="mu" style="margin:2px 0 8px">Ouverture/fermeture, agents à l’ouverture et à la fermeture, durée du shift par jour, total par semaine.</div><div class="p-row"><label>Ouverture <input type="time" id="p-so" onchange="PRO.setSvc()"></label><label>Fermeture <input type="time" id="p-sc" onchange="PRO.setSvc()"></label><label>Agents ouverture <input type="number" min="0" id="p-no" style="width:64px" onchange="PRO.setN()"></label><label>Agents fermeture <input type="number" min="0" id="p-nc" style="width:64px" onchange="PRO.setN()"></label></div><div class="p-dl" id="p-dl"></div><div class="p-row"><label>Total / semaine <input id="p-tg" type="number" step="0.5" style="width:80px" onchange="PRO.setTarget(this.value)"></label><label><input type="checkbox" id="p-au" onchange="PRO.setAuto(this.checked)" style="width:auto"> Changer les OFF chaque semaine</label><label style="width:100%"><input type="checkbox" id="p-fz" onchange="PRO.setFrz(this.checked)" style="width:auto"> Planning figé : rien ne change tout seul, c’est moi qui modifie</label></div></div></div><div class="card" id="p-tpl" style="padding:14px;margin-top:14px"></div><div class="section-title">Couverture par heure</div><div class="mu" style="margin-bottom:8px">Total en gras, puis agents dans Groupe A · B · C · D-fixe (0 en rouge = groupe non couvert).</div><div class="p-w p-cv"><table id="p-cv"></table></div>';
function mount(){
 var dash=$("view-dashboard"),main=dash.parentNode,banner=$("dashboardNewRapportsBanner");
 var box=document.createElement("div");box.className="pro";box.innerHTML=DASH;
 var old=document.createElement("div");old.className="hidden";
 Array.prototype.slice.call(dash.children).forEach(function(c){if(c!==banner)old.appendChild(c)});
 dash.appendChild(box);dash.appendChild(old);
 var pl=document.createElement("section");pl.className="view hidden pro";pl.id="view-planningpro";pl.innerHTML=PLN;main.appendChild(pl);
 var nb=document.querySelector('.nav-item[data-view="planningnv"]');
 if(nb){
  var hs=nb.cloneNode(true);hs.querySelector("span").textContent="Hotspots NV";
  nb.dataset.view="planningpro";nb.parentNode.insertBefore(hs,nb.nextSibling);
  hs.addEventListener("click",function(){
   document.querySelectorAll(".nav-item").forEach(function(b){b.classList.remove("active")});hs.classList.add("active");
   document.querySelectorAll(".view").forEach(function(v){v.classList.add("hidden")});
   $("view-planningnv").classList.remove("hidden");$("viewTitle").textContent="Hotspots NV";
   if(window.listenHotspotsNv)window.listenHotspotsNv();
  });
 }
}
function start(){
 if(started)return;
 mount();started=true;
 try{var r=localStorage.getItem("pro1");if(r)load(JSON.parse(r))}catch(e){}
 S.agents=S.agents.filter(function(a){return a.n!=="Achraf Akhlij"});migrate();if(needSave){needSave=false;persist()}
 render();setInterval(tick,30000);
 try{ref=db.collection("planningNV").doc("pro");ref.onSnapshot(function(d){if(d.exists&&d.data().json){try{load(JSON.parse(d.data().json));S.agents=S.agents.filter(function(a){return a.n!=="Achraf Akhlij"})}catch(e){}if(needSave){needSave=false;persist()}render()}else if(!d.exists){ref.set({json:JSON.stringify(S)}).catch(function(e){console.error(e)})}},function(e){console.error(e)})}catch(e){ref=null}
}
var tried=false;
function safeStart(){
 if(tried)return;tried=true;
 try{start()}catch(e){
  console.error(e);var d=document.getElementById("view-dashboard")||document.body,m=document.createElement("div");
  m.style.cssText="background:#fee;color:#b00;padding:12px;margin:12px;border-radius:8px;font:13px monospace";
  m.textContent="Erreur Planning Pro : "+e.message;d.insertBefore(m,d.firstChild);
 }
}
try{auth.onAuthStateChanged(function(u){if(u&&u.email)safeStart()})}catch(e){console.error(e)}
var iv=setInterval(function(){var a=$("app");if(started||tried){clearInterval(iv);return}if(a&&a.style.display==="block")safeStart()},700);
})();
</script>
<!-- ===== FIN PLANNING AGENTS PRO ===== -->

</body>

</html>
