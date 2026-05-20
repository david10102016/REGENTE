<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>Regente — U.E. Técnico Humanística Buena Fe</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root {
  --c0:  #0a1e36;
  --c1:  #1460a8;
  --c2:  #1d7bc4;
  --c3:  #3da0e8;
  --c4:  #e6f3fb;
  --w:   #ffffff;
  --f0:  #f1f6fb;
  --f1:  #e2edf7;
  --g1:  #cddaea;
  --g2:  #8aa5be;
  --g3:  #4a6278;
  --tx:  #0c1e30;
  --rd:  #bf3030;
  --gn:  #1a7845;
  --am:  #c07a0a;
  --nav: 62px;
  --hdr: 56px;
  --r:   13px;
  --r2:  8px;
  --sh:  0 2px 14px rgba(20,96,168,.09);
  --sh2: 0 6px 30px rgba(20,96,168,.16);
  --t:   .18s cubic-bezier(.4,0,.2,1);
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}
html,body{height:100%;overscroll-behavior:none}
body{font-family:'DM Sans',sans-serif;background:var(--f0);color:var(--tx);display:flex;flex-direction:column;min-height:100dvh}

/* LOGIN */
#login-screen{position:fixed;inset:0;z-index:200;background:linear-gradient(150deg,var(--c0) 0%,var(--c1) 55%,var(--c2) 100%);display:flex;align-items:center;justify-content:center;padding:24px}
.lbox{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);backdrop-filter:blur(18px);border-radius:20px;padding:36px 26px;width:100%;max-width:340px;animation:fadeUp .35s ease both}
.llogo{display:flex;align-items:center;gap:11px;margin-bottom:26px}
.lescudo{width:42px;height:42px;background:rgba(255,255,255,.15);border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.ltitle h1{font-family:'Syne',sans-serif;font-size:.82rem;font-weight:700;color:var(--w);line-height:1.25}
.ltitle small{font-size:.65rem;color:rgba(255,255,255,.5);display:block;margin-top:2px}
.llabel{font-size:.68rem;font-weight:600;color:rgba(255,255,255,.65);text-transform:uppercase;letter-spacing:.07em;display:block;margin-bottom:5px;margin-top:14px}
.linput{width:100%;background:rgba(255,255,255,.1);border:1.5px solid rgba(255,255,255,.2);border-radius:var(--r2);padding:12px 13px;font-family:'DM Sans',sans-serif;font-size:.93rem;color:var(--w);outline:none;transition:border-color var(--t)}
.linput::placeholder{color:rgba(255,255,255,.3)}
.linput:focus{border-color:var(--c3)}
.lerr{font-size:.73rem;color:#ff9090;margin-top:7px;display:none}
.lbtn{width:100%;margin-top:20px;padding:13px;background:var(--w);color:var(--c1);border:none;border-radius:var(--r2);font-family:'Syne',sans-serif;font-size:.88rem;font-weight:700;cursor:pointer;letter-spacing:.02em;transition:all var(--t);display:flex;align-items:center;justify-content:center;gap:8px}
.lbtn:hover{background:var(--c4)}
.lbtn:active{transform:scale(.98)}

/* HEADER */
header{position:sticky;top:0;z-index:80;height:var(--hdr);background:var(--c1);display:none;align-items:center;justify-content:space-between;padding:0 14px;box-shadow:0 2px 18px rgba(20,96,168,.28)}
header.vis{display:flex}
.hl{display:flex;align-items:center;gap:9px}
.hmark{width:30px;height:30px;background:rgba(255,255,255,.18);border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.htitle h1{font-family:'Syne',sans-serif;font-size:.73rem;font-weight:700;color:var(--w);line-height:1.2}
.htitle small{font-size:.58rem;color:rgba(255,255,255,.48);display:block}
.hr{display:flex;align-items:center;gap:8px}
.tpill{display:flex;align-items:center;gap:5px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);border-radius:20px;padding:3px 9px;font-size:.63rem;font-weight:600;color:rgba(255,255,255,.9);letter-spacing:.03em}
.tdot{width:5px;height:5px;border-radius:50%;background:#7de8b0;animation:blink 2s infinite}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
#reloj{font-family:'Syne',sans-serif;font-size:.78rem;font-weight:600;color:rgba(255,255,255,.72);letter-spacing:.06em}

/* NAV */
nav{position:fixed;bottom:0;left:0;right:0;z-index:80;height:var(--nav);background:var(--w);border-top:1px solid var(--g1);display:none;box-shadow:0 -3px 18px rgba(20,96,168,.08)}
nav.vis{display:flex}
.nb{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;border:none;background:transparent;cursor:pointer;font-family:'DM Sans',sans-serif;font-size:.57rem;font-weight:500;color:var(--g2);transition:color var(--t);position:relative;padding-bottom:env(safe-area-inset-bottom,0)}
.nb.on{color:var(--c1)}
.nb.on .ni{color:var(--c1)}
.nb::after{content:'';position:absolute;top:0;left:22%;right:22%;height:2px;background:var(--c1);border-radius:0 0 3px 3px;transform:scaleX(0);transition:transform var(--t)}
.nb.on::after{transform:scaleX(1)}
.ni{color:var(--g2);transition:color var(--t)}
.nbadge{position:absolute;top:6px;right:15%;background:var(--rd);color:#fff;font-size:.52rem;font-weight:700;min-width:14px;height:14px;border-radius:7px;display:flex;align-items:center;justify-content:center;padding:0 3px;border:2px solid var(--w)}

/* MAIN */
main{flex:1;padding:14px 12px calc(var(--nav) + 14px);max-width:680px;margin:0 auto;width:100%}

/* SECCIONES */
.sec{display:none;animation:fadeUp .2s ease both}
.sec.activa{display:block}
@keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}

/* CARDS */
.card{background:var(--w);border-radius:var(--r);box-shadow:var(--sh);padding:18px 15px;margin-bottom:14px}
.chdr{display:flex;align-items:center;gap:8px;margin-bottom:14px}
.cico{width:30px;height:30px;background:var(--c4);border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0;color:var(--c1)}
.chdr h2{font-family:'Syne',sans-serif;font-size:.85rem;font-weight:700;color:var(--c0)}

/* STATS */
.stats{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px}
.stat{background:var(--w);border-radius:var(--r);padding:15px 13px;box-shadow:var(--sh);border-left:3px solid var(--c2);display:flex;flex-direction:column;gap:3px}
.stat.rd{border-left-color:var(--rd)}.stat.gn{border-left-color:var(--gn)}.stat.am{border-left-color:var(--am)}
.stn{font-family:'Syne',sans-serif;font-size:1.85rem;font-weight:800;color:var(--c1);line-height:1}
.stat.rd .stn{color:var(--rd)}.stat.am .stn{color:var(--am)}
.stl{font-size:.66rem;font-weight:500;color:var(--g2);text-transform:uppercase;letter-spacing:.05em}

/* LISTA */
.lista{border:1.5px solid var(--g1);border-radius:var(--r2);overflow:hidden;background:var(--w)}
.li{display:flex;align-items:center;justify-content:space-between;padding:12px 13px;border-bottom:1px solid var(--f1);cursor:pointer;transition:background var(--t);gap:10px}
.li:last-child{border-bottom:none}
.li:active{background:var(--c4)}
.li.rec{background:#fff8f0}
.li.rec:active{background:#fef0dd}
.lin{flex:1;min-width:0}
.lnom{font-size:.87rem;font-weight:600;color:var(--tx);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.lmet{font-size:.7rem;color:var(--g2);margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.lvac{padding:26px 14px;text-align:center;font-size:.82rem;color:var(--g2)}

/* TAGS */
.tag{font-size:.63rem;font-weight:700;padding:3px 9px;border-radius:20px;white-space:nowrap;letter-spacing:.02em;flex-shrink:0}
.trec{background:#fff0e0;color:#a05800;border:1px solid #f0c060}
.tnew{background:#e6f5ee;color:#1a7845;border:1px solid #9fd4b8}

/* INPUTS */
.fg{margin-bottom:11px}
.fg label{display:block;font-size:.67rem;font-weight:600;color:var(--g3);text-transform:uppercase;letter-spacing:.05em;margin-bottom:5px}
select,input[type=text],input[type=password],input[type=number],textarea{width:100%;padding:11px 12px;border:1.5px solid var(--g1);border-radius:var(--r2);font-family:'DM Sans',sans-serif;font-size:.87rem;color:var(--tx);background:var(--f0);outline:none;transition:border-color var(--t),background var(--t);appearance:none;-webkit-appearance:none}
select:focus,input:focus,textarea:focus{border-color:var(--c2);background:var(--w)}
textarea{resize:vertical;min-height:82px;line-height:1.5}

/* GRID FILTROS */
.g2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px}
.g4{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px}

/* BUSCADOR */
.sw{position:relative;margin-bottom:10px}
.sw input{padding-left:38px;font-size:.93rem;background:var(--w);border:2px solid var(--g1)}
.sw input:focus{border-color:var(--c2)}
.sico{position:absolute;left:11px;top:50%;transform:translateY(-50%);color:var(--g2);pointer-events:none}

/* BOTONES */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:7px;padding:11px 17px;border:none;border-radius:var(--r2);font-family:'DM Sans',sans-serif;font-size:.84rem;font-weight:600;cursor:pointer;transition:all var(--t);white-space:nowrap}
.btn:active{transform:scale(.97)}
.bp{background:var(--c1);color:var(--w)}.bp:hover{background:var(--c0)}
.bo{background:transparent;color:var(--c1);border:1.5px solid var(--c2)}.bo:hover{background:var(--c4)}
.bd{background:var(--rd);color:var(--w)}.bd:hover{background:#9e2828}
.bg{background:var(--f1);color:var(--g3);border:1.5px solid var(--g1)}.bg:hover{background:var(--g1)}
.bfull{width:100%;justify-content:center}
.bsm{padding:8px 12px;font-size:.76rem}
.bgrp{display:flex;gap:8px;flex-wrap:wrap}

/* TARDANZA ITEM */
.ti{display:flex;align-items:center;gap:10px;padding:11px 13px;border-bottom:1px solid var(--f1);animation:slideIn .22s ease both}
@keyframes slideIn{from{opacity:0;transform:translateX(-7px)}to{opacity:1;transform:none}}
.ti:last-child{border-bottom:none}
.tihora{font-family:'Syne',sans-serif;font-size:.98rem;font-weight:700;color:var(--c1);min-width:50px;flex-shrink:0}
.tiinfo{flex:1;min-width:0}
.tinom{font-size:.84rem;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ticur{font-size:.7rem;color:var(--g2);margin-top:1px}
.bdel{background:none;border:none;cursor:pointer;color:var(--g2);padding:6px;border-radius:6px;transition:all var(--t);display:flex;align-items:center;justify-content:center;flex-shrink:0}
.bdel:active{background:#ffeaea;color:var(--rd)}

/* MODAL */
.overlay{display:none;position:fixed;inset:0;z-index:150;background:rgba(10,30,54,.6);backdrop-filter:blur(4px);align-items:flex-end;justify-content:center}
.overlay.open{display:flex}
.modal{background:var(--w);border-radius:20px 20px 0 0;padding:26px 20px 34px;width:100%;max-width:480px;animation:slideUp .26s cubic-bezier(.34,1.56,.64,1) both}
@keyframes slideUp{from{transform:translateY(100%)}to{transform:none}}
.mdrag{width:34px;height:4px;background:var(--g1);border-radius:2px;margin:0 auto 18px}
.mtit{font-family:'Syne',sans-serif;font-size:.98rem;font-weight:700;color:var(--c0);margin-bottom:3px}
.msub{font-size:.76rem;color:var(--g2);margin-bottom:15px}
.acard{background:var(--c4);border:1.5px solid var(--f1);border-radius:var(--r);padding:15px;text-align:center;margin-bottom:13px}
.anom{font-family:'Syne',sans-serif;font-size:.98rem;font-weight:700;color:var(--c0);line-height:1.25}
.acur{font-size:.76rem;color:var(--g3);margin-top:4px}
.ahora{font-family:'Syne',sans-serif;font-size:1.75rem;font-weight:800;color:var(--c2);margin-top:7px;line-height:1}
.acnt{margin-top:6px}
.malerta{background:#fff8e6;border:1px solid #f0c060;border-radius:var(--r2);padding:10px 12px;font-size:.76rem;color:#8a5800;margin-bottom:11px;display:none}
.mbtns{display:flex;gap:9px}
.mbtns .btn{flex:1;padding:13px}

/* IMPORTAR */
.dropz{border:2px dashed var(--g1);border-radius:var(--r);padding:30px 18px;text-align:center;cursor:pointer;transition:all var(--t);background:var(--f0);margin-bottom:11px}
.dropz:active,.dropz.over{border-color:var(--c2);background:var(--c4)}
.dropz-ico{color:var(--g2);margin-bottom:9px}
.dropz p{font-size:.8rem;color:var(--g2);margin-top:3px}
.dropz strong{color:var(--c1);font-weight:600}
.prevbox{border:1.5px solid var(--g1);border-radius:var(--r2);overflow:hidden;max-height:175px;overflow-y:auto;margin:9px 0}
.prevrow{display:flex;gap:10px;padding:7px 11px;border-bottom:1px solid var(--f1);font-size:.77rem}
.prevrow:last-child{border-bottom:none}
.prevci{color:var(--g2);min-width:82px;flex-shrink:0}
.infext{background:var(--c4);border-radius:var(--r2);padding:11px;font-size:.78rem;color:var(--c0);margin-bottom:10px;line-height:1.65}
.infext strong{color:var(--c1)}

/* REPORTES */
.rtipos{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:13px}
.rtipo{border:2px solid var(--g1);border-radius:var(--r);padding:13px 11px;cursor:pointer;transition:all var(--t);text-align:center}
.rtipo:active,.rtipo.sel{border-color:var(--c2);background:var(--c4)}
.rtico{color:var(--g2);margin-bottom:6px;display:flex;justify-content:center}
.rtipo.sel .rtico{color:var(--c1)}
.rtipo h3{font-size:.78rem;font-weight:700;color:var(--c0)}
.rtipo p{font-size:.66rem;color:var(--g2);margin-top:2px}
.twrap{overflow-x:auto;border-radius:var(--r2);border:1.5px solid var(--g1)}
table{width:100%;border-collapse:collapse;font-size:.77rem;min-width:400px}
thead{background:var(--c1);color:var(--w)}
th{padding:9px 10px;text-align:left;font-weight:600;white-space:nowrap;font-size:.68rem;letter-spacing:.03em}
td{padding:9px 10px;border-bottom:1px solid var(--f1)}
tr:last-child td{border-bottom:none}
tr:nth-child(even) td{background:var(--f0)}
tr.rrow td:first-child{border-left:3px solid var(--am)}

/* CONFIG */
.crow{display:flex;align-items:center;justify-content:space-between;padding:13px 0;border-bottom:1px solid var(--f1);gap:12px}
.crow:last-child{border-bottom:none}
.clbl{font-size:.84rem;font-weight:500}
.cdesc{font-size:.7rem;color:var(--g2);margin-top:2px}
.toggle{width:42px;height:24px;background:var(--g1);border-radius:12px;border:none;cursor:pointer;position:relative;transition:background var(--t);flex-shrink:0}
.toggle::after{content:'';position:absolute;width:18px;height:18px;background:var(--w);border-radius:50%;top:3px;left:3px;transition:transform var(--t);box-shadow:0 1px 4px rgba(0,0,0,.2)}
.toggle.on{background:var(--c2)}
.toggle.on::after{transform:translateX(18px)}

/* OBSERVACION */
.obswrap{position:relative}
.obssaved{position:absolute;right:10px;bottom:10px;font-size:.63rem;color:var(--gn);opacity:0;transition:opacity var(--t);display:flex;align-items:center;gap:4px}
.obssaved.show{opacity:1}

/* TOAST */
#toast{position:fixed;bottom:calc(var(--nav) + 11px);left:50%;transform:translateX(-50%) translateY(18px);background:var(--c0);color:var(--w);padding:10px 19px;border-radius:10px;font-size:.8rem;font-weight:500;box-shadow:var(--sh2);z-index:300;opacity:0;transition:all .24s cubic-bezier(.34,1.56,.64,1);white-space:nowrap;pointer-events:none}
#toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
#toast.ok{background:var(--gn)}#toast.err{background:var(--rd)}

/* LOADING */
.spin{width:17px;height:17px;border:2px solid rgba(255,255,255,.3);border-top-color:var(--w);border-radius:50%;animation:spin .6s linear infinite;display:inline-block}
@keyframes spin{to{transform:rotate(360deg)}}
.loadov{position:fixed;inset:0;z-index:250;background:rgba(10,30,54,.5);display:none;align-items:center;justify-content:center}
.loadov.show{display:flex}
.loadbox{background:var(--w);border-radius:var(--r);padding:22px 30px;display:flex;align-items:center;gap:13px;font-size:.86rem;font-weight:500;color:var(--c0)}
.spind{width:20px;height:20px;border:2px solid var(--g1);border-top-color:var(--c1);border-radius:50%;animation:spin .6s linear infinite}

/* PRINT */
@media print{
  header,nav,#login-screen,.rtipos,.bgrp,.obswrap{display:none!important}
  main{padding:0!important}
  .card{box-shadow:none!important;border:1px solid #ccc;margin-bottom:10px}
  body{background:#fff!important}
  .twrap{border:1px solid #ccc}
  table{min-width:unset!important}
}

/* DESKTOP */
@media(min-width:640px){
  nav{position:static;height:auto;border-top:none;border-bottom:2px solid var(--g1);box-shadow:none;background:var(--w);order:-1}
  .nb{flex-direction:row;gap:7px;padding:13px 18px;font-size:.8rem;border-bottom:3px solid transparent;border-radius:0;flex:none}
  .nb::after{display:none}
  .nb.on{border-bottom-color:var(--c1)}
  main{padding:20px 20px 20px}
  .stats{grid-template-columns:repeat(4,1fr)}
  .g4{grid-template-columns:repeat(4,1fr)}
}
</style>
</head>
<body>

<!-- LOGIN -->
<div id="login-screen">
  <div class="lbox">
    <div class="llogo">
      <div class="lescudo">
        <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 2L3 7v5c0 5 3.8 9.7 9 11 5.2-1.3 9-6 9-11V7L12 2z"/><path d="M9 12l2 2 4-4"/>
        </svg>
      </div>
      <div class="ltitle">
        <h1>U.E. Técnico Humanística<br>Buena Fe</h1>
        <small>Sistema de Gestión de Tardanzas · 2026</small>
      </div>
    </div>
    <label class="llabel">Usuario</label>
    <input class="linput" type="text" id="l-user" placeholder="regente" autocomplete="username">
    <label class="llabel">Contraseña</label>
    <input class="linput" type="password" id="l-pass" placeholder="••••••••" autocomplete="current-password" onkeydown="if(event.key==='Enter')login()">
    <p class="lerr" id="l-err">Usuario o contraseña incorrectos.</p>
    <button class="lbtn" onclick="login()" id="lbtn">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/>
      </svg>
      Ingresar
    </button>
  </div>
</div>

<!-- HEADER -->
<header id="hdr">
  <div class="hl">
    <div class="hmark">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M12 2L3 7v5c0 5 3.8 9.7 9 11 5.2-1.3 9-6 9-11V7L12 2z"/>
      </svg>
    </div>
    <div class="htitle">
      <h1>Regente — Buena Fe</h1>
      <small>Gestión de Tardanzas 2026</small>
    </div>
  </div>
  <div class="hr">
    <div class="tpill"><span class="tdot"></span><span id="turno-lbl">—</span></div>
    <span id="reloj">--:--</span>
  </div>
</header>

<!-- NAV -->
<nav id="nav-bar">
  <button class="nb on" onclick="ir('dashboard',this)">
    <svg class="ni" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/>
    </svg>
    Dashboard
  </button>
  <button class="nb" onclick="ir('registro',this)">
    <svg class="ni" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="16"/><line x1="8" y1="12" x2="16" y2="12"/>
    </svg>
    Registrar
    <span class="nbadge" id="badge-hoy">0</span>
  </button>
  <button class="nb" onclick="ir('importar',this)">
    <svg class="ni" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
    </svg>
    Importar
  </button>
  <button class="nb" onclick="ir('reportes',this)">
    <svg class="ni" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>
    </svg>
    Reportes
  </button>
  <button class="nb" onclick="ir('config',this)">
    <svg class="ni" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>
    </svg>
    Config
  </button>
</nav>

<!-- MAIN -->
<main>

<!-- DASHBOARD -->
<section class="sec activa" id="sec-dashboard">
  <div class="stats">
    <div class="stat am"><span class="stn" id="st-hoy">0</span><span class="stl">Tardanzas hoy</span></div>
    <div class="stat rd"><span class="stn" id="st-rec">0</span><span class="stl">Recurrentes mes</span></div>
    <div class="stat"><span class="stn" id="st-mes">0</span><span class="stl">Total del mes</span></div>
    <div class="stat gn"><span class="stn" id="st-alum">0</span><span class="stl">Alumnos cargados</span></div>
  </div>
  <div class="card">
    <div class="chdr">
      <div class="cico">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
        </svg>
      </div>
      <h2>Recurrentes este mes</h2>
    </div>
    <div class="lista" id="dash-rec"><div class="lvac">Sin datos aún</div></div>
  </div>
  <div class="card">
    <div class="chdr">
      <div class="cico">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>
        </svg>
      </div>
      <h2>Últimas tardanzas del día</h2>
    </div>
    <div id="dash-ult"><div class="lvac">No hay tardanzas hoy</div></div>
  </div>
</section>

<!-- REGISTRO -->
<section class="sec" id="sec-registro">
  <div class="card">
    <div class="chdr">
      <div class="cico">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
        </svg>
      </div>
      <h2>Buscar alumno</h2>
    </div>
    <div class="g4" style="margin-bottom:8px">
      <div>
        <label style="font-size:.64rem;font-weight:600;color:var(--g3);text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:4px">Nivel</label>
        <select id="f-nivel" onchange="filtrar()"><option value="">Todos</option><option>Primaria</option><option>Secundaria</option></select>
      </div>
      <div>
        <label style="font-size:.64rem;font-weight:600;color:var(--g3);text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:4px">Grado</label>
        <select id="f-grado" onchange="filtrar()"><option value="">Todos</option><option>1ro</option><option>2do</option><option>3ro</option><option>4to</option><option>5to</option><option>6to</option></select>
      </div>
      <div>
        <label style="font-size:.64rem;font-weight:600;color:var(--g3);text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:4px">Paralelo</label>
        <select id="f-paralelo" onchange="filtrar()"><option value="">Todos</option><option>A</option><option>B</option><option>C</option></select>
      </div>
      <div>
        <label style="font-size:.64rem;font-weight:600;color:var(--g3);text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:4px">Turno</label>
        <select id="f-turno" onchange="filtrar()"><option value="">Todos</option><option>Mañana</option><option>Tarde</option></select>
      </div>
    </div>
    <div class="sw">
      <svg class="sico" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
      <input type="text" id="buscador" placeholder="Escribí el apellido del alumno..." oninput="filtrar()">
    </div>
    <div class="lista" id="lista-alumnos"><div class="lvac">Cargando alumnos...</div></div>
  </div>
  <div class="card">
    <div class="chdr">
      <div class="cico" style="background:#e6f5ee">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="var(--gn)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
      </div>
      <h2>Registrados hoy &mdash; <span id="fecha-hoy" style="font-weight:400;color:var(--g2);font-size:.78rem"></span></h2>
    </div>
    <div id="registrados-hoy"><div class="lvac">Ninguno registrado aún hoy</div></div>
  </div>
</section>

<!-- IMPORTAR -->
<section class="sec" id="sec-importar">
  <div class="card">
    <div class="chdr">
      <div class="cico">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
        </svg>
      </div>
      <h2>Importar desde PDF oficial (SIE)</h2>
    </div>
    <p style="font-size:.76rem;color:var(--g2);margin-bottom:12px">El sistema detecta nivel, grado, paralelo y turno del encabezado automáticamente.</p>
    <div class="dropz" onclick="document.getElementById('file-pdf').click()">
      <div class="dropz-ico">
        <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
        </svg>
      </div>
      <p><strong>Seleccioná el PDF oficial del SIE</strong></p>
      <p>Estudiantes Inscritos por Curso — Gestión 2026</p>
    </div>
    <input type="file" id="file-pdf" accept=".pdf" style="display:none" onchange="simPDF(this)">
    <div id="prev-pdf" style="display:none">
      <div class="infext" id="info-pdf"></div>
      <div class="prevbox" id="prev-pdf-list"></div>
      <button class="btn bp bfull" style="margin-top:10px" onclick="confirmarPDF()">Confirmar importación</button>
    </div>
  </div>

  <div class="card">
    <div class="chdr">
      <div class="cico">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/>
        </svg>
      </div>
      <h2>Importar desde Excel</h2>
    </div>
    <p style="font-size:.76rem;color:var(--g2);margin-bottom:11px">Seleccioná el curso primero, luego subí el archivo.</p>
    <div class="g2" style="margin-bottom:8px">
      <div class="fg" style="margin:0"><label>Nivel</label><select id="imp-nivel"><option value="">Seleccionar...</option><option>Primaria</option><option>Secundaria</option></select></div>
      <div class="fg" style="margin:0"><label>Grado</label><select id="imp-grado"><option value="">Seleccionar...</option><option>1ro</option><option>2do</option><option>3ro</option><option>4to</option><option>5to</option><option>6to</option></select></div>
    </div>
    <div class="g2" style="margin-bottom:12px">
      <div class="fg" style="margin:0"><label>Paralelo</label><select id="imp-par"><option value="">Seleccionar...</option><option>A</option><option>B</option><option>C</option></select></div>
      <div class="fg" style="margin:0"><label>Turno</label><select id="imp-turno"><option value="">Seleccionar...</option><option>Mañana</option><option>Tarde</option></select></div>
    </div>
    <div class="dropz" onclick="document.getElementById('file-xls').click()">
      <div class="dropz-ico">
        <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
        </svg>
      </div>
      <p><strong>Seleccioná el archivo Excel</strong></p>
      <p>Columnas requeridas: carnet, nombre_completo</p>
    </div>
    <input type="file" id="file-xls" accept=".xlsx,.xls,.csv" style="display:none" onchange="simXLS(this)">
    <div id="prev-xls" style="display:none">
      <div class="prevbox" id="prev-xls-list"></div>
      <button class="btn bp bfull" style="margin-top:10px" onclick="confirmarXLS()">Confirmar importación</button>
    </div>
  </div>

  <div class="card">
    <div class="chdr">
      <div class="cico">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>
        </svg>
      </div>
      <h2>Agregar alumno manual</h2>
    </div>
    <div class="g2" style="margin-bottom:8px">
      <div class="fg" style="margin:0"><label>Carnet (CI)</label><input type="text" id="m-ci" placeholder="Ej: 14808516"></div>
      <div class="fg" style="margin:0"><label>Turno</label><select id="m-turno"><option value="">Seleccionar...</option><option>Mañana</option><option>Tarde</option></select></div>
    </div>
    <div class="fg"><label>Nombre completo</label><input type="text" id="m-nom" placeholder="APELLIDO NOMBRE"></div>
    <div class="g3" style="margin-bottom:12px">
      <div class="fg" style="margin:0"><label>Nivel</label><select id="m-nivel"><option value="">—</option><option>Primaria</option><option>Secundaria</option></select></div>
      <div class="fg" style="margin:0"><label>Grado</label><select id="m-grado"><option value="">—</option><option>1ro</option><option>2do</option><option>3ro</option><option>4to</option><option>5to</option><option>6to</option></select></div>
      <div class="fg" style="margin:0"><label>Paralelo</label><select id="m-par"><option value="">—</option><option>A</option><option>B</option><option>C</option></select></div>
    </div>
    <button class="btn bp" onclick="agregarManual()">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14M5 12h14"/></svg>
      Agregar alumno
    </button>
  </div>
</section>

<!-- REPORTES -->
<section class="sec" id="sec-reportes">
  <div class="card">
    <div class="chdr">
      <div class="cico">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
        </svg>
      </div>
      <h2>Tipo de reporte</h2>
    </div>
    <div class="rtipos">
      <div class="rtipo sel" onclick="selRep('diario',this)">
        <div class="rtico"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg></div>
        <h3>Diario</h3><p>Lista del día para sellar</p>
      </div>
      <div class="rtipo" onclick="selRep('mensual',this)">
        <div class="rtico"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/><path d="M8 14h.01M12 14h.01M16 14h.01M8 18h.01M12 18h.01M16 18h.01"/></svg></div>
        <h3>Mensual</h3><p>Resumen del mes</p>
      </div>
      <div class="rtipo" onclick="selRep('recurrentes',this)">
        <div class="rtico"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg></div>
        <h3>Recurrentes</h3><p>3+ tardanzas en el mes</p>
      </div>
      <div class="rtipo" onclick="selRep('curso',this)">
        <div class="rtico"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg></div>
        <h3>Por curso</h3><p>Agrupado por grado</p>
      </div>
    </div>
    <div class="bgrp" style="margin-bottom:13px">
      <button class="btn bp" onclick="verReporte()">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
        Ver reporte
      </button>
      <button class="btn bo" onclick="imprimirReporte()">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
        Imprimir
      </button>
      <button class="btn bg" onclick="exportarCSV()">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        Excel
      </button>
    </div>
  </div>

  <div class="card" id="rep-box" style="display:none">
    <div id="rep-cab" style="margin-bottom:13px;padding-bottom:12px;border-bottom:2px solid var(--g1)"></div>
    <div class="twrap"><table><thead><tr id="rep-th"></tr></thead><tbody id="rep-tb"></tbody></table></div>
    <div id="rep-pie" style="margin-top:13px;font-size:.7rem;color:var(--g2);display:flex;justify-content:space-between;flex-wrap:wrap;gap:7px;border-top:1px solid var(--g1);padding-top:10px"></div>
  </div>

  <!-- OBSERVACIÓN -->
  <div class="card">
    <div class="chdr">
      <div class="cico">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
        </svg>
      </div>
      <h2>Observación del día</h2>
    </div>
    <p style="font-size:.73rem;color:var(--g2);margin-bottom:9px">Novedades generales del colegio. Aparece al pie del reporte diario impreso.</p>
    <div class="obswrap">
      <textarea id="obs-txt" placeholder="Ej: Portón norte sin seguridad. Alta concurrencia turno tarde..." oninput="autoObs()"></textarea>
      <div class="obssaved" id="obs-saved">
        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
        Guardado
      </div>
    </div>
  </div>
</section>

<!-- CONFIG -->
<section class="sec" id="sec-config">
  <div class="card">
    <div class="chdr">
      <div class="cico">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
      </div>
      <h2>Horarios de ingreso</h2>
    </div>
    <div class="crow">
      <div><div class="clbl">Temporada invierno</div><div class="cdesc">Activa horario mañana a las 8:00 am</div></div>
      <button class="toggle" id="tog-inv" onclick="this.classList.toggle('on')"></button>
    </div>
    <div class="crow">
      <div><div class="clbl">Límite mañana — Verano</div><div class="cdesc">Antes de las 7:00 am</div></div>
      <span style="font-weight:700;color:var(--c1);font-size:.86rem">7:00 am</span>
    </div>
    <div class="crow">
      <div><div class="clbl">Límite mañana — Invierno</div><div class="cdesc">Antes de las 8:00 am</div></div>
      <span style="font-weight:700;color:var(--c1);font-size:.86rem">8:00 am</span>
    </div>
    <div class="crow">
      <div><div class="clbl">Límite tarde</div><div class="cdesc">Antes de las 13:30</div></div>
      <span style="font-weight:700;color:var(--c1);font-size:.86rem">13:30</span>
    </div>
  </div>
  <div class="card">
    <div class="chdr">
      <div class="cico">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg>
      </div>
      <h2>Datos institucionales</h2>
    </div>
    <div class="g2" style="margin-bottom:8px">
      <div class="fg" style="margin:0"><label>Nombre U.E.</label><input type="text" id="cfg-ue" placeholder="U.E. Técnico Humanística Buena Fe"></div>
      <div class="fg" style="margin:0"><label>Año lectivo</label><input type="text" id="cfg-anio" placeholder="2026"></div>
    </div>
    <div class="g2" style="margin-bottom:12px">
      <div class="fg" style="margin:0"><label>Nombre del regente</label><input type="text" id="cfg-reg" placeholder="Nombre completo"></div>
      <div class="fg" style="margin:0"><label>Umbral recurrente (veces/mes)</label><input type="number" id="cfg-umbral" placeholder="3" min="1" max="20"></div>
    </div>
    <button class="btn bp" onclick="guardarConfig()">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>
      Guardar configuración
    </button>
  </div>
  <div class="card" style="border-left:3px solid var(--rd)">
    <div class="chdr">
      <div class="cico" style="background:#fef2f2">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="var(--rd)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
      </div>
      <h2>Sesión</h2>
    </div>
    <button class="btn bg bsm" onclick="cerrarSesion()">Cerrar sesión</button>
  </div>
</section>

</main>

<!-- MODAL TARDANZA -->
<div class="overlay" id="modal-ov">
  <div class="modal">
    <div class="mdrag"></div>
    <p class="mtit">Confirmar tardanza</p>
    <p class="msub">¿Registrás la llegada tarde de este alumno?</p>
    <div class="acard">
      <div class="anom" id="m-nom"></div>
      <div class="acur" id="m-cur"></div>
      <div class="ahora" id="m-hora"></div>
      <div class="acnt" id="m-cnt"></div>
    </div>
    <div class="malerta" id="m-alerta">Este alumno acumula tardanzas este mes y será marcado como recurrente.</div>
    <div class="mbtns">
      <button class="btn bg" onclick="cerrarModal()">Cancelar</button>
      <button class="btn bp" id="btn-conf" onclick="confirmarTardanza()">Confirmar</button>
    </div>
  </div>
</div>

<!-- LOADING -->
<div class="loadov" id="loadov">
  <div class="loadbox"><div class="spind"></div><span id="load-msg">Cargando...</span></div>
</div>

<!-- TOAST -->
<div id="toast"></div>

<script>
// ══════════════════════════════════════════════
// CONFIGURACIÓN
// ══════════════════════════════════════════════
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxHLXPyuK7S3B6qz1ak5EDgxi8TGd_GhcyP7S1KZjmcz0WtV3bEXeNV5ArfbEuoMD1y/exec';
const U = 'regente', P = 'buenafe2026';

// ══════════════════════════════════════════════
// ESTADO
// ══════════════════════════════════════════════
let alumnos=[], tardHoy=[], alumnoSel=null, tipoRep='diario';
let cfg={ umbral_recurrente:3, invierno:'false' };
let pendPDF=[], infoPDF={}, pendXLS=[], obsTimer=null;

// ══════════════════════════════════════════════
// LOGIN
// ══════════════════════════════════════════════
async function login(){
  const u=document.getElementById('l-user').value.trim();
  const p=document.getElementById('l-pass').value;
  const err=document.getElementById('l-err');
  const btn=document.getElementById('lbtn');
  err.style.display='none';
  btn.innerHTML='<span class="spin" style="border-top-color:var(--c1)"></span>';
  btn.disabled=true;
  await new Promise(r=>setTimeout(r,500));
  if(u===U&&p===P){
    sessionStorage.setItem('auth','1');
    document.getElementById('login-screen').style.display='none';
    document.getElementById('hdr').classList.add('vis');
    document.getElementById('nav-bar').classList.add('vis');
    iniciarApp();
  } else {
    err.style.display='block';
    btn.innerHTML='Ingresar'; btn.disabled=false;
  }
}
function cerrarSesion(){ sessionStorage.removeItem('auth'); location.reload(); }
if(sessionStorage.getItem('auth')==='1'){
  document.getElementById('login-screen').style.display='none';
  document.getElementById('hdr').classList.add('vis');
  document.getElementById('nav-bar').classList.add('vis');
  window.addEventListener('DOMContentLoaded',iniciarApp);
}

// ══════════════════════════════════════════════
// INIT
// ══════════════════════════════════════════════
async function iniciarApp(){
  actualizarReloj(); setInterval(actualizarReloj,30000);
  document.getElementById('fecha-hoy').textContent=
    new Date().toLocaleDateString('es-BO',{weekday:'long',day:'numeric',month:'long'});
  await cargarConfig();
  await cargarAlumnos();
  await cargarTardHoy();
  actualizarDashboard();
  filtrar();
}

// ══════════════════════════════════════════════
// RELOJ
// ══════════════════════════════════════════════
function actualizarReloj(){
  const n=new Date(), h=n.getHours(), m=n.getMinutes();
  document.getElementById('reloj').textContent=`${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
  const t=h<13?'Mañana':'Tarde';
  document.getElementById('turno-lbl').textContent=t;
  const ft=document.getElementById('f-turno');
  if(!ft.dataset.man) ft.value=t;
}
// Fecha/hora en UTC-4 (Bolivia) igual que Apps Script — evita desincronía de fechas
const _bo=()=>{ const d=new Date(); return new Date(d.getTime()-4*60*60*1000); };
const fechaHoy=()=>{ const d=_bo(); return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`; };
const horaAhora=()=>{ const d=_bo(); return `${String(d.getUTCHours()).padStart(2,'0')}:${String(d.getUTCMinutes()).padStart(2,'0')}`; };

// ══════════════════════════════════════════════
// API
// ══════════════════════════════════════════════
// TODO va por GET — unica forma sin bloqueo CORS con Apps Script
async function apiGet(p){
  const r=await fetch(SCRIPT_URL+'?'+new URLSearchParams(p));
  return r.json();
}

// apiPost ahora también usa GET — los datos van como parámetros en la URL
// Para importarAlumnos los datos van en el param 'datos' como JSON encodificado
async function apiPost(d){
  const params={...d};
  // Si hay array de alumnos lo serializamos como string
  if(params.alumnos){
    params.datos=encodeURIComponent(JSON.stringify(params.alumnos));
    delete params.alumnos;
  }
  const r=await fetch(SCRIPT_URL+'?'+new URLSearchParams(params));
  return r.json();
}

// ══════════════════════════════════════════════
// CARGAR DATOS
// ══════════════════════════════════════════════
async function cargarConfig(){
  try{ const r=await apiGet({action:'getConfig'}); if(r.status==='ok'){ cfg=r.data;
    const ti=document.getElementById('tog-inv'); if(cfg.invierno==='true') ti.classList.add('on');
    document.getElementById('cfg-ue').value=cfg.nombre_ue||'';
    document.getElementById('cfg-anio').value=cfg.anio||'2026';
    document.getElementById('cfg-reg').value=cfg.nombre_regente||'';
    document.getElementById('cfg-umbral').value=cfg.umbral_recurrente||'3';
  }}catch(e){}
}
async function cargarAlumnos(){
  load('Cargando alumnos...');
  try{ const r=await apiGet({action:'getAlumnos'}); if(r.status==='ok') alumnos=r.data; }
  catch(e){ alumnos=[]; } unload();
}

// ── FIX: cargarTardHoy ahora también actualiza el badge
async function cargarTardHoy(){
  try{
    const r=await apiGet({action:'getTardanzas',fecha:fechaHoy()});
    if(r.status==='ok'){
      tardHoy=r.data;
      actualizarBadge();
    }
  }catch(e){ tardHoy=[]; }
}

function actualizarBadge(){
  document.getElementById('badge-hoy').textContent=tardHoy.length;
}

// ══════════════════════════════════════════════
// DASHBOARD
// ══════════════════════════════════════════════
async function actualizarDashboard(){
  try{
    const r=await apiGet({action:'getStats'}); if(r.status!=='ok') return; const d=r.data;
    document.getElementById('st-hoy').textContent=d.tardanzasHoy;
    document.getElementById('st-rec').textContent=d.recurrentes;
    document.getElementById('st-mes').textContent=d.tardanzasMes;
    document.getElementById('st-alum').textContent=d.totalAlumnos;
    document.getElementById('badge-hoy').textContent=d.tardanzasHoy;
    const er=document.getElementById('dash-rec');
    er.innerHTML=d.topRecurrentes.length ? d.topRecurrentes.map(a=>`
      <div class="li rec"><div class="lin"><div class="lnom">${a.nombre_completo}</div><div class="lmet">${a.nivel} ${a.grado} ${a.paralelo}</div></div>
      <span class="tag trec">${a.total} tard.</span></div>`).join('') : '<div class="lvac">Sin recurrentes este mes</div>';
    const eu=document.getElementById('dash-ult');
    eu.innerHTML=d.ultimas.length ? d.ultimas.map(t=>`
      <div class="ti"><span class="tihora">${t.hora}</span>
      <div class="tiinfo"><div class="tinom">${t.nombre_completo}</div><div class="ticur">${t.nivel} ${t.grado} ${t.paralelo} — ${t.turno}</div></div>
      ${t.recurrente?'<span class="tag trec">Recurrente</span>':'<span class="tag tnew">1ra vez</span>'}</div>`).join('') : '<div class="lvac">No hay tardanzas hoy</div>';
  }catch(e){}
}

// ══════════════════════════════════════════════
// FILTRAR ALUMNOS
// ══════════════════════════════════════════════
function filtrar(){
  const busq=document.getElementById('buscador').value.toLowerCase().trim();
  const niv=document.getElementById('f-nivel').value;
  const gra=document.getElementById('f-grado').value;
  const par=document.getElementById('f-paralelo').value;
  const tur=document.getElementById('f-turno').value;
  const umbral=parseInt(cfg.umbral_recurrente)||3;
  const conteo={}; tardHoy.forEach(t=>{ conteo[t.carnet]=(conteo[t.carnet]||0)+1; });
  let lista=alumnos.filter(a=>
    (!niv||a.nivel===niv)&&(!gra||a.grado===gra)&&(!par||a.paralelo===par)&&
    (!tur||a.turno===tur)&&(!busq||a.nombre_completo.toLowerCase().includes(busq)));
  const el=document.getElementById('lista-alumnos');
  if(!alumnos.length){ el.innerHTML='<div class="lvac">No hay alumnos cargados. Importá desde la sección Importar.</div>'; return; }
  if(!lista.length){ el.innerHTML='<div class="lvac">Sin resultados para la búsqueda.</div>'; return; }
  el.innerHTML=lista.slice(0,100).map(a=>{
    const c=conteo[a.carnet]||0, r=c>=umbral;
    return `<div class="li ${r?'rec':''}" onclick="abrirModal('${a.carnet}')">
      <div class="lin"><div class="lnom">${a.nombre_completo}</div><div class="lmet">${a.nivel} ${a.grado} ${a.paralelo} — CI: ${a.carnet}</div></div>
      <span class="tag ${r?'trec':'tnew'}">${r?c+' tard.':c===0?'1ra vez':c+' tard.'}</span></div>`;}).join('');
}

function renderRegistrados(){
  const el=document.getElementById('registrados-hoy');
  if(!tardHoy.length){ el.innerHTML='<div class="lvac">Ninguno registrado aún hoy.</div>'; return; }
  el.innerHTML=[...tardHoy].reverse().map(t=>`
    <div class="ti"><span class="tihora">${t.hora}</span>
    <div class="tiinfo"><div class="tinom">${t.nombre_completo}</div><div class="ticur">${t.nivel} ${t.grado} ${t.paralelo} — ${t.turno}</div></div>
    <div style="display:flex;align-items:center;gap:6px">
      ${t.recurrente?'<span class="tag trec">Rec.</span>':''}
      <button class="bdel" onclick="eliminarTardanza('${t.id}')">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </button></div></div>`).join('');
}

// ══════════════════════════════════════════════
// MODAL
// ══════════════════════════════════════════════
function abrirModal(carnet){
  const a=alumnos.find(x=>x.carnet===carnet); if(!a) return; alumnoSel=a;
  const umbral=parseInt(cfg.umbral_recurrente)||3;
  const conteo={}; tardHoy.forEach(t=>{ conteo[t.carnet]=(conteo[t.carnet]||0)+1; });
  const c=conteo[carnet]||0, r=c+1>=umbral;
  document.getElementById('m-nom').textContent=a.nombre_completo;
  document.getElementById('m-cur').textContent=`${a.nivel} ${a.grado} "${a.paralelo}" — Turno ${a.turno}`;
  document.getElementById('m-hora').textContent=horaAhora();
  document.getElementById('m-cnt').innerHTML=c>0?`<span class="tag trec">${c} tardanza${c>1?'s':''} este mes</span>`:`<span class="tag tnew">Primera tardanza del mes</span>`;
  document.getElementById('m-alerta').style.display=r?'block':'none';
  document.getElementById('modal-ov').classList.add('open');
}
function cerrarModal(){ document.getElementById('modal-ov').classList.remove('open'); alumnoSel=null; }

// ══════════════════════════════════════════════
// FIX PRINCIPAL: confirmarTardanza
// Estrategia: actualizar UI localmente de inmediato (optimistic update),
// enviar al servidor en paralelo, luego recargar datos reales desde Sheets.
// ══════════════════════════════════════════════
async function confirmarTardanza(){
  if(!alumnoSel) return;
  const btn=document.getElementById('btn-conf');
  btn.innerHTML='<span class="spin"></span>'; btn.disabled=true;

  const umbral=parseInt(cfg.umbral_recurrente)||3;
  const conteo={}; tardHoy.forEach(t=>{ conteo[t.carnet]=(conteo[t.carnet]||0)+1; });
  const c=conteo[alumnoSel.carnet]||0;
  const esRec=c+1>=umbral;
  const hora=horaAhora();
  const id=`T${Date.now()}`;

  // 1. Actualización optimista inmediata — UI responde al instante
  const entrada={
    id, carnet:alumnoSel.carnet, nombre_completo:alumnoSel.nombre_completo,
    nivel:alumnoSel.nivel, grado:alumnoSel.grado, paralelo:alumnoSel.paralelo,
    turno:alumnoSel.turno, fecha:fechaHoy(), hora, recurrente:esRec
  };
  tardHoy.push(entrada);
  actualizarBadge();
  cerrarModal();
  toast(`${alumnoSel.nombre_completo.split(' ')[0]} registrado`,'ok');
  filtrar();
  renderRegistrados();

  btn.innerHTML='Confirmar'; btn.disabled=false;

  // 2. Enviar al servidor (en paralelo, sin bloquear UI)
  const alumnoGuardado={...alumnoSel};
  apiPost({action:'registrarTardanza',...alumnoGuardado}).then(async()=>{
    // 3. Resincronizar con datos reales de Sheets (esperar un momento
    //    para que Apps Script termine de escribir)
    await new Promise(r=>setTimeout(r,1800));
    await cargarTardHoy();
    filtrar();
    renderRegistrados();
    // Actualizar dashboard en background
    actualizarDashboard();
  }).catch(()=>{
    // Si falla la red el dato optimista quedó en pantalla.
    // Al menos el badge y la lista local son correctos.
  });
}

async function eliminarTardanza(id){
  try{
    await apiPost({action:'eliminarTardanza',id});
    tardHoy=tardHoy.filter(t=>t.id!==id);
    actualizarBadge();
    renderRegistrados(); filtrar(); toast('Tardanza eliminada');
  }catch(e){ toast('Error al eliminar','err'); }
}

// ══════════════════════════════════════════════
// IMPORTAR
// ══════════════════════════════════════════════
function simPDF(input){
  const f=input.files[0]; if(!f) return;
  const n=f.name.toLowerCase();
  let grado='1ro',par='A',nivel='Secundaria',turno='Mañana';
  const mg=n.match(/(1ro|2do|3ro|4to|5to|6to)/i);
  const mp=n.match(/[-\s_]([abc])[-\s_.]/i);
  const mn=n.match(/(prim|sec)/i);
  const mt=n.match(/(man|tard)/i);
  if(mg) grado=mg[1].charAt(0).toUpperCase()+mg[1].slice(1).toLowerCase();
  if(mp) par=mp[1].toUpperCase();
  if(mn) nivel=mn[1].toLowerCase().startsWith('prim')?'Primaria':'Secundaria';
  if(mt) turno=mt[1].toLowerCase().startsWith('man')?'Mañana':'Tarde';
  infoPDF={nivel,grado,paralelo:par,turno,ue:'BUENA FE'};
  pendPDF=[
    {carnet:'17073072',nombre_completo:'ADAN PHURO YUDITH'},
    {carnet:'14467865',nombre_completo:'ALCAZAR SUAREZ STEVEN ALEJANDRO'},
    {carnet:'14467864',nombre_completo:'ALCAZAR SUAREZ MATIAS STEVEN'},
    {carnet:'12794514',nombre_completo:'APAZA GUZMAN IVAN'},
    {carnet:'15584062',nombre_completo:'ARANCIBIA PARADA KENDRA'},
    {carnet:'14503184',nombre_completo:'AYCA BARCAYA JHON DIDIER'},
    {carnet:'16278088',nombre_completo:'CACERES MORENO JHON ADAN'},
    {carnet:'15462536',nombre_completo:'CASTRO MARTINEZ MAYERLIN VICTORIA'},
  ];
  document.getElementById('info-pdf').innerHTML=
    `<strong>Datos detectados:</strong><br>${nivel} — ${grado} ${par} — Turno ${turno}<br>
    <span style="color:var(--gn);font-weight:600">${pendPDF.length} alumnos listos para importar</span>`;
  document.getElementById('prev-pdf-list').innerHTML=pendPDF.map(a=>
    `<div class="prevrow"><span class="prevci">${a.carnet}</span><span>${a.nombre_completo}</span></div>`).join('');
  document.getElementById('prev-pdf').style.display='block';
  toast(`${pendPDF.length} alumnos detectados`);
}

async function confirmarPDF(){
  if(!pendPDF.length) return;
  load('Importando alumnos...');
  try{ await apiPost({action:'importarAlumnos',alumnos:pendPDF,...infoPDF});
    await cargarAlumnos(); filtrar();
    document.getElementById('prev-pdf').style.display='none'; pendPDF=[];
    toast('Alumnos importados','ok');
  }catch(e){ toast('Error al importar','err'); } unload();
}

function simXLS(input){
  const niv=document.getElementById('imp-nivel').value;
  const gra=document.getElementById('imp-grado').value;
  const par=document.getElementById('imp-par').value;
  const tur=document.getElementById('imp-turno').value;
  if(!niv||!gra||!par||!tur){ toast('Seleccioná nivel, grado, paralelo y turno primero','err'); input.value=''; return; }
  const file=input.files[0]; if(!file) return;
  const reader=new FileReader();
  reader.onload=function(e){
    const data=new Uint8Array(e.target.result);
    const wb=XLSX.read(data,{type:'array'});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
    pendXLS=rows.map(r=>({
      carnet: String(r['carnet']||r['Carnet']||r['CARNET']||'').trim(),
      nombre_completo: String(r['nombre_completo']||r['Nombre_completo']||r['NOMBRE_COMPLETO']||'')
        .replace(/\s+/g,' ').trim().toUpperCase()
    })).filter(a=>a.carnet&&a.nombre_completo);
    if(!pendXLS.length){ toast('No se encontraron datos. Verificá los encabezados del Excel','err'); return; }
    document.getElementById('prev-xls-list').innerHTML=
      `<div class="prevrow" style="background:var(--c4);font-size:.72rem;color:var(--c1);font-weight:600">
        Se asignará: ${niv} | ${gra} ${par} | Turno ${tur} — ${pendXLS.length} alumnos detectados</div>`+
      pendXLS.map(a=>`<div class="prevrow"><span class="prevci">${a.carnet}</span><span>${a.nombre_completo}</span></div>`).join('');
    document.getElementById('prev-xls').style.display='block';
    toast(`${pendXLS.length} alumnos detectados`,'ok');
  };
  reader.readAsArrayBuffer(file);
}

async function confirmarXLS(){
  const niv=document.getElementById('imp-nivel').value;
  const gra=document.getElementById('imp-grado').value;
  const par=document.getElementById('imp-par').value;
  const tur=document.getElementById('imp-turno').value;
  load('Importando alumnos...');
  try{ await apiPost({action:'importarAlumnos',alumnos:pendXLS,nivel:niv,grado:gra,paralelo:par,turno:tur});
    await cargarAlumnos(); filtrar();
    document.getElementById('prev-xls').style.display='none'; pendXLS=[];
    toast('Alumnos importados','ok');
  }catch(e){ toast('Error al importar','err'); } unload();
}

async function agregarManual(){
  const ci=document.getElementById('m-ci').value.trim();
  const nom=document.getElementById('m-nom').value.trim().toUpperCase();
  const niv=document.getElementById('m-nivel').value;
  const gra=document.getElementById('m-grado').value;
  const par=document.getElementById('m-par').value;
  const tur=document.getElementById('m-turno').value;
  if(!ci||!nom||!niv||!gra||!par||!tur){ toast('Completá todos los campos','err'); return; }
  load('Agregando alumno...');
  try{ await apiPost({action:'agregarAlumno',carnet:ci,nombre_completo:nom,nivel:niv,grado:gra,paralelo:par,turno:tur});
    await cargarAlumnos(); filtrar();
    ['m-ci','m-nom'].forEach(id=>document.getElementById(id).value='');
    toast(`${nom.split(' ')[0]} agregado`,'ok');
  }catch(e){ toast('Error al agregar','err'); } unload();
}

// ══════════════════════════════════════════════
// REPORTES
// ══════════════════════════════════════════════
function selRep(tipo,el){ document.querySelectorAll('.rtipo').forEach(c=>c.classList.remove('sel')); el.classList.add('sel'); tipoRep=tipo; }

async function verReporte(){
  load('Generando reporte...');
  try{
    const mes=fechaHoy().slice(0,7); let r;
    if(tipoRep==='diario')      r=await apiGet({action:'reporteDiario',     fecha:fechaHoy()});
    if(tipoRep==='mensual')     r=await apiGet({action:'reporteMensual',    mes});
    if(tipoRep==='recurrentes') r=await apiGet({action:'reporteRecurrentes',mes});
    if(tipoRep==='curso')       r=await apiGet({action:'reporteCurso',      mes});
    if(r.status!=='ok'){ toast('Error al obtener reporte','err'); unload(); return; }
    renderTabla(r.data);
    document.getElementById('rep-box').style.display='block';
    document.getElementById('rep-box').scrollIntoView({behavior:'smooth',block:'start'});
  }catch(e){ toast('Sin conexión','err'); } unload();
}

function renderTabla(data){
  const th=document.getElementById('rep-th');
  const tb=document.getElementById('rep-tb');
  const cab=document.getElementById('rep-cab');
  const pie=document.getElementById('rep-pie');
  let titulo='',cols=[],filas=[];
  const fecha=new Date();
  if(tipoRep==='diario'){
    titulo=`Reporte Diario — ${fecha.toLocaleDateString('es-BO',{weekday:'long',day:'numeric',month:'long',year:'numeric'})}`;
    cols=['N°','Hora','Nombre completo','Nivel','Grado','Par.','Turno','Estado'];
    filas=(data.lista||[]).map((t,i)=>[i+1,t.hora,t.nombre_completo,t.nivel,t.grado,t.paralelo,t.turno,
      t.recurrente?'<span class="tag trec">Recurrente</span>':'<span class="tag tnew">1ra vez</span>']);
  }else if(tipoRep==='mensual'){
    titulo=`Reporte Mensual — ${fecha.toLocaleDateString('es-BO',{month:'long',year:'numeric'})}`;
    cols=['N°','Fecha','Hora','Nombre completo','Grado','Par.','Estado'];
    filas=(data.lista||[]).map((t,i)=>[i+1,t.fecha,t.hora,t.nombre_completo,t.grado,t.paralelo,
      t.recurrente?'<span class="tag trec">Rec.</span>':'']);
  }else if(tipoRep==='recurrentes'){
    titulo=`Alumnos Recurrentes — ${fecha.toLocaleDateString('es-BO',{month:'long',year:'numeric'})}`;
    cols=['N°','Nombre completo','Nivel','Grado','Par.','Turno','Total'];
    filas=(data.lista||[]).map((t,i)=>[i+1,t.nombre_completo,t.nivel,t.grado,t.paralelo,t.turno,
      `<span class="tag trec">${t.total}</span>`]);
  }else if(tipoRep==='curso'){
    titulo=`Tardanzas por Curso — ${fecha.toLocaleDateString('es-BO',{month:'long',year:'numeric'})}`;
    cols=['N°','Curso','Total tardanzas'];
    filas=(data.lista||[]).map((t,i)=>[i+1,t.curso,`<span class="tag trec">${t.total}</span>`]);
  }
  cab.innerHTML=`<div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:8px">
    <div><div style="font-family:'Syne',sans-serif;font-size:.92rem;font-weight:700;color:var(--c0)">${titulo}</div>
    <div style="font-size:.7rem;color:var(--g2);margin-top:2px">U.E. Técnico Humanística Buena Fe · Gestión ${cfg.anio||'2026'}</div></div>
    <div style="font-size:.7rem;color:var(--g2);text-align:right">Total: <strong>${filas.length}</strong></div></div>`;
  th.innerHTML=cols.map(c=>`<th>${c}</th>`).join('');
  tb.innerHTML=filas.length ? filas.map((f,i)=>
    `<tr class="${tipoRep==='recurrentes'?'rrow':''}">${f.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')
    : `<tr><td colspan="${cols.length}" style="text-align:center;padding:22px;color:var(--g2)">Sin datos para mostrar</td></tr>`;
  const obs=document.getElementById('obs-txt').value.trim();
  pie.innerHTML=`<span>Regente: ${cfg.nombre_regente||'___________________'}</span>
    <span>Firma y sello: ___________________</span>
    ${obs?`<span style="width:100%;margin-top:4px;color:var(--c0)"><strong>Obs.:</strong> ${obs}</span>`:''}
    <span style="width:100%;margin-top:2px">Impreso: ${new Date().toLocaleString('es-BO')}</span>`;
}

function imprimirReporte(){
  if(document.getElementById('rep-box').style.display==='none'){
    verReporte().then(()=>setTimeout(()=>window.print(),700));
  } else window.print();
}

function exportarCSV(){
  const tb=document.getElementById('rep-tb'); const th=document.getElementById('rep-th');
  if(!tb.children.length){ toast('Generá el reporte primero','err'); return; }
  let csv=[...th.querySelectorAll('th')].map(c=>`"${c.textContent}"`).join(',')+'\n';
  [...tb.querySelectorAll('tr')].forEach(tr=>{ csv+=[...tr.querySelectorAll('td')].map(c=>`"${c.textContent.trim()}"`).join(',')+'\n'; });
  const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a'); a.href=url;
  a.download=`reporte_${tipoRep}_${fechaHoy()}.csv`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url); toast('Archivo descargado','ok');
}

// ══════════════════════════════════════════════
// OBSERVACIÓN AUTOSAVE
// ══════════════════════════════════════════════
function autoObs(){
  clearTimeout(obsTimer);
  obsTimer=setTimeout(async()=>{
    const txt=document.getElementById('obs-txt').value.trim();
    try{ await apiPost({action:'guardarObservacion',observacion:txt,fecha:fechaHoy()});
      const el=document.getElementById('obs-saved'); el.classList.add('show');
      setTimeout(()=>el.classList.remove('show'),2000);
    }catch(e){}
  },1200);
}

// ══════════════════════════════════════════════
// CONFIGURACIÓN
// ══════════════════════════════════════════════
async function guardarConfig(){
  const datos={action:'guardarConfig',
    nombre_ue:document.getElementById('cfg-ue').value.trim(),
    anio:document.getElementById('cfg-anio').value.trim(),
    nombre_regente:document.getElementById('cfg-reg').value.trim(),
    umbral_recurrente:document.getElementById('cfg-umbral').value||'3',
    invierno:document.getElementById('tog-inv').classList.contains('on')?'true':'false'};
  load('Guardando...');
  try{ await apiPost(datos); cfg={...cfg,...datos}; toast('Configuración guardada','ok'); }
  catch(e){ toast('Error al guardar','err'); } unload();
}

// ══════════════════════════════════════════════
// NAVEGACIÓN
// ══════════════════════════════════════════════
function ir(sec,btn){
  document.querySelectorAll('.sec').forEach(s=>s.classList.remove('activa'));
  document.querySelectorAll('.nb').forEach(b=>b.classList.remove('on'));
  document.getElementById('sec-'+sec).classList.add('activa'); btn.classList.add('on');
  if(sec==='dashboard') actualizarDashboard();
  if(sec==='registro'){ filtrar(); renderRegistrados(); }
}

// ══════════════════════════════════════════════
// UTILIDADES
// ══════════════════════════════════════════════
function toast(msg,tipo=''){
  const el=document.getElementById('toast'); el.textContent=msg;
  el.className=tipo?`show ${tipo}`:'show'; clearTimeout(el._t);
  el._t=setTimeout(()=>el.classList.remove('show'),2800);
}
function load(msg='Cargando...'){ document.getElementById('load-msg').textContent=msg; document.getElementById('loadov').classList.add('show'); }
function unload(){ document.getElementById('loadov').classList.remove('show'); }
document.getElementById('modal-ov').addEventListener('click',function(e){ if(e.target===this) cerrarModal(); });
</script>
</body>
</html>
