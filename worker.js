/**
 * MDV Global Delivery Intelligence — Cloudflare Worker
 *
 * Routes:
 *  GET /api/shipments  → fetches from Airtable using AIRTABLE_API_KEY secret
 *  GET /*              → serves the static dashboard HTML
 */

const AIRTABLE_BASE_ID  = 'appcwvcnVuHI9TMdA';
const AIRTABLE_TABLE_ID = 'tbl6oiAmXfM1SFzT8';

const FIELD_IDS = [
  'fldl6RM0TkX4GFrZF', // Shipment ID (SH-xxx)
  'fldLwZ53Zc4MO59OU', // Shipment Reference (LI ref)
  'fldLnyv6M7vjg5sEg', // Shipping Destination (linked record)
  'fldQUXBo2smwtjFgX', // Shipment Status (single select)
  'fldBdrLERHs7kK6L9', // Current ETA to Warehouse (formula → date string)
  'fldfPbJDbTWy4ZHec', // Confirmed Freight Type (lookup)
  'fldB1m2fWiZobPVeM', // Total Units on Shipment (rollup)
  'fldJtOiIWBFRbrDpD', // Total Units Delivered (rollup)
];

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type': 'application/json',
};

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // ── CORS preflight ──────────────────────────────────────────────────
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS });
    }

    // ── API route: fetch shipments from Airtable ────────────────────────
    if (url.pathname === '/api/shipments') {
      return handleShipments(env);
    }

    // ── Everything else: serve the dashboard HTML ───────────────────────
    return new Response(getDashboardHTML(), {
      headers: { 'Content-Type': 'text/html; charset=utf-8' },
    });
  },
};

// ── Airtable fetcher ────────────────────────────────────────────────────────
async function handleShipments(env) {
  const apiKey = env.AIRTABLE_API_KEY;

  if (!apiKey) {
    return json({ error: 'AIRTABLE_API_KEY secret is not set. Add it in Cloudflare Dashboard → Workers → Your Worker → Settings → Variables and Secrets.' }, 500);
  }

  try {
    const allRecords = [];
    let offset = null;

    do {
      const params = new URLSearchParams();
      FIELD_IDS.forEach(id => params.append('fields[]', id));
      params.set('pageSize', '100');
      if (offset) params.set('offset', offset);

      const resp = await fetch(
        `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${AIRTABLE_TABLE_ID}?${params}`,
        { headers: { Authorization: `Bearer ${apiKey}` } }
      );

      if (!resp.ok) {
        const text = await resp.text();
        throw new Error(`Airtable ${resp.status}: ${text}`);
      }

      const data = await resp.json();
      allRecords.push(...data.records);
      offset = data.offset || null;
    } while (offset);

    const shipments = allRecords.map(rec => {
      const f = rec.cellValuesByFieldId;

      // Destination: linked record array → first name
      const destLinks = f['fldLnyv6M7vjg5sEg'] || [];
      const dest = Array.isArray(destLinks) && destLinks.length > 0
        ? (destLinks[0].name || '').trim()
        : '';

      // Freight: lookup returns nested valuesByLinkedRecordId structure
      let freight = '';
      const fr = f['fldfPbJDbTWy4ZHec'];
      if (fr) {
        if (typeof fr === 'string') {
          freight = fr;
        } else if (fr.valuesByLinkedRecordId) {
          const vals = Object.values(fr.valuesByLinkedRecordId);
          if (vals.length && Array.isArray(vals[0]) && vals[0].length) {
            freight = vals[0][0].name || '';
          }
        }
      }

      // Status: singleSelect → .name
      const statusRaw = f['fldQUXBo2smwtjFgX'];
      const status = statusRaw && typeof statusRaw === 'object'
        ? statusRaw.name || ''
        : statusRaw || '';

      return {
        shId:      f['fldl6RM0TkX4GFrZF'] || '',
        liRef:     f['fldLwZ53Zc4MO59OU'] || '',
        dest,
        status,
        eta:       f['fldBdrLERHs7kK6L9'] || null,
        freight,
        unitsDue:  parseFloat(f['fldB1m2fWiZobPVeM'] || 0),
        unitsRcvd: parseFloat(f['fldJtOiIWBFRbrDpD'] || 0),
      };
    });

    return json({ shipments, total: shipments.length });
  } catch (err) {
    return json({ error: err.message }, 500);
  }
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: CORS });
}

// ── Dashboard HTML (inlined so one file deploys cleanly) ────────────────────
function getDashboardHTML() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Manière De Voir — Global Delivery Intelligence</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,500;1,300;1,400&family=Montserrat:wght@300;400;500;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"><\/script>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --black:#0a0a0a;--white:#fafaf8;--warm:#f4f2ed;--warm2:#eceae3;
  --gold:#b8a06a;--muted:#888880;--border:#dddad2;--border-dark:#c8c5bc;
  --danger:#8b2020;--danger-bg:#fdf4f4;--danger-border:#e8c0c0;
  --success:#1a5c3a;--success-bg:#f2faf6;--success-border:#b0d4c0;
  --warning:#7a5a1a;--warning-bg:#fdf8f0;--warning-border:#d4c0a0;
  --info:#2a4a70;--info-bg:#f0f4fa;--info-border:#c0d0e8;
  --purple:#4a3a70;--purple-bg:#f6f4fa;--purple-border:#c8c0d8;
}
html{font-size:14px}
body{background:var(--white);color:var(--black);font-family:'Montserrat',sans-serif;font-weight:400;line-height:1.5;min-height:100vh;-webkit-font-smoothing:antialiased}
.page{max-width:1320px;margin:0 auto;padding:40px 40px 80px}
.header{display:flex;align-items:flex-end;justify-content:space-between;padding-bottom:24px;border-bottom:1px solid var(--black);margin-bottom:36px}
.wordmark{font-family:'Cormorant Garamond',serif;font-size:30px;font-weight:300;letter-spacing:0.22em;text-transform:uppercase;color:var(--black);line-height:1}
.wordmark em{font-style:italic;font-weight:300}
.sub-title{font-size:9px;letter-spacing:0.32em;text-transform:uppercase;color:var(--muted);margin-top:7px}
.header-actions{display:flex;align-items:center;gap:10px}
.live-indicator{display:none;align-items:center;gap:6px;font-size:9px;letter-spacing:0.18em;text-transform:uppercase;color:var(--muted);padding:6px 10px;border:1px solid var(--border)}
.live-dot{width:6px;height:6px;border-radius:50%;background:var(--success);animation:blink 2.5s ease-in-out infinite}
@keyframes blink{0%,100%{opacity:1}50%{opacity:0.35}}
.btn{font-family:'Montserrat',sans-serif;font-size:9px;font-weight:500;letter-spacing:0.22em;text-transform:uppercase;padding:9px 20px;border:1px solid var(--black);background:transparent;color:var(--black);cursor:pointer;transition:background 0.18s,color 0.18s}
.btn:hover{background:var(--black);color:var(--white)}
.btn:disabled{opacity:0.3;cursor:not-allowed}
.btn-primary{background:var(--black);color:var(--white)}
.btn-primary:hover{background:#2a2a2a}
.loading-screen{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:50vh;gap:20px}
.loading-word{font-family:'Cormorant Garamond',serif;font-size:22px;font-weight:300;letter-spacing:0.35em;text-transform:uppercase;color:var(--black);animation:fade-pulse 1.6s ease-in-out infinite}
@keyframes fade-pulse{0%,100%{opacity:0.35}50%{opacity:1}}
.loading-sub{font-size:9px;letter-spacing:0.24em;text-transform:uppercase;color:var(--muted)}
.error-box{padding:16px 20px;border:1px solid var(--danger-border);background:var(--danger-bg);color:var(--danger);font-size:12px;line-height:1.6;margin-bottom:24px}
.kpi-strip{display:grid;grid-template-columns:repeat(5,1fr);border:1px solid var(--border);margin-bottom:36px}
.kpi-cell{padding:22px 20px;border-right:1px solid var(--border)}
.kpi-cell:last-child{border-right:none}
.kpi-label{font-size:8px;letter-spacing:0.28em;text-transform:uppercase;color:var(--muted);margin-bottom:10px;font-weight:500}
.kpi-value{font-family:'Cormorant Garamond',serif;font-size:36px;font-weight:300;line-height:1;color:var(--black)}
.kpi-value.alert{color:var(--danger)}
.kpi-sub{font-size:9px;color:var(--muted);margin-top:6px;letter-spacing:0.06em}
.region-nav{display:flex;border-bottom:1px solid var(--border);margin-bottom:28px}
.r-tab{font-family:'Montserrat',sans-serif;font-size:9px;font-weight:500;letter-spacing:0.22em;text-transform:uppercase;padding:11px 20px;border:none;border-bottom:2px solid transparent;background:transparent;color:var(--muted);cursor:pointer;display:flex;align-items:center;gap:7px;margin-bottom:-1px;transition:color 0.15s;white-space:nowrap}
.r-tab:hover{color:var(--black)}
.r-tab.active{color:var(--black);border-bottom-color:var(--black)}
.r-tab-count{font-size:8px;letter-spacing:0.05em;padding:2px 6px;background:var(--warm);font-weight:500}
.r-tab.active .r-tab-count{background:var(--black);color:var(--white)}
.r-tab-alert{color:var(--danger);font-size:9px}
.region-metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--border);border:1px solid var(--border);margin-bottom:20px}
.rm-cell{background:var(--white);padding:18px 20px}
.rm-label{font-size:8px;letter-spacing:0.26em;text-transform:uppercase;color:var(--muted);margin-bottom:7px;font-weight:500}
.rm-value{font-family:'Cormorant Garamond',serif;font-size:28px;font-weight:300;line-height:1.1}
.progress-bar{height:2px;background:var(--border);margin-top:9px}
.progress-fill{height:100%;background:var(--black);transition:width 0.4s ease}
.status-row{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px}
.status-chip{font-size:8px;letter-spacing:0.14em;text-transform:uppercase;padding:3px 9px;border:1px solid var(--border);color:var(--muted);font-weight:500}
.chip-danger{border-color:var(--danger-border);color:var(--danger);background:var(--danger-bg)}
.overdue-alert{display:flex;align-items:flex-start;gap:12px;padding:12px 16px;background:var(--danger-bg);border:1px solid var(--danger-border);margin-bottom:20px;font-size:10px;color:var(--danger);line-height:1.6;letter-spacing:0.04em}
.alert-icon{font-size:13px;flex-shrink:0;margin-top:1px}
.table-container{border:1px solid var(--border)}
.table-topbar{display:flex;align-items:center;justify-content:space-between;padding:13px 18px;border-bottom:1px solid var(--border);background:var(--warm)}
.table-caption{font-size:9px;letter-spacing:0.26em;text-transform:uppercase;color:var(--black);font-weight:500}
.overdue-count-badge{font-size:8px;letter-spacing:0.14em;text-transform:uppercase;padding:4px 10px;background:var(--danger);color:white;font-weight:500}
.table-scroll{overflow-x:auto}
table{width:100%;border-collapse:collapse;min-width:920px}
thead th{font-family:'Montserrat',sans-serif;font-size:8px;font-weight:500;letter-spacing:0.22em;text-transform:uppercase;color:var(--muted);padding:11px 16px;text-align:left;border-bottom:1px solid var(--border);background:var(--warm);white-space:nowrap}
tbody td{padding:10px 16px;font-size:11px;border-bottom:1px solid var(--border);color:var(--black);vertical-align:middle}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover td{background:var(--warm)}
.row-overdue td{background:var(--danger-bg)}
.row-overdue:hover td{background:#fae8e8}
.week-divider td{background:var(--warm2);font-size:8px;letter-spacing:0.22em;text-transform:uppercase;color:var(--muted);font-weight:500;padding:6px 16px;border-bottom:1px solid var(--border)}
.li-ref{font-size:10px;color:var(--gold);letter-spacing:0.06em;font-weight:600;font-family:'Montserrat',sans-serif}
.sh-ref{font-size:10px;color:var(--muted);letter-spacing:0.04em;font-family:'Montserrat',sans-serif}
.mono{font-family:'Montserrat',sans-serif;font-size:10px;letter-spacing:0.02em}
.neg{color:var(--danger)}.pos{color:var(--success)}.muted{color:var(--muted)}
.pill{display:inline-flex;align-items:center;gap:4px;font-size:8px;letter-spacing:0.14em;text-transform:uppercase;padding:3px 8px;border:1px solid var(--border);color:var(--muted);font-weight:500;white-space:nowrap}
.pill-air{border-color:var(--info-border);color:var(--info);background:var(--info-bg)}
.pill-sea{border-color:var(--success-border);color:var(--success);background:var(--success-bg)}
.pill-land{border-color:var(--warning-border);color:var(--warning);background:var(--warning-bg)}
.pill-complete{border-color:var(--success-border);color:var(--success);background:var(--success-bg)}
.pill-short{border-color:var(--warning-border);color:var(--warning);background:var(--warning-bg)}
.pill-pending{border-color:var(--purple-border);color:var(--purple);background:var(--purple-bg)}
.pill-overdue{border-color:var(--danger-border);color:var(--danger);background:var(--danger-bg)}
.pill-received{border-color:var(--success-border);color:var(--success);background:var(--success-bg)}
.pill-shipped{border-color:var(--purple-border);color:var(--purple);background:var(--purple-bg)}
.pill-quoting{border-color:var(--border);color:var(--muted);background:var(--warm)}
.mini-progress{display:flex;align-items:center;gap:8px}
.mini-track{width:44px;height:2px;background:var(--border);flex-shrink:0}
.mini-fill{height:100%;background:var(--black)}
tfoot .total-row td{background:var(--warm2)!important;font-size:10px;letter-spacing:0.1em;font-weight:600;border-top:1px solid var(--border-dark);padding:11px 16px}
.footer{margin-top:48px;padding-top:20px;border-top:1px solid var(--border);display:flex;align-items:center;justify-content:space-between}
.footer-brand{font-family:'Cormorant Garamond',serif;font-size:12px;font-weight:300;letter-spacing:0.2em;text-transform:uppercase;color:var(--muted)}
.footer-note{font-size:8px;letter-spacing:0.16em;text-transform:uppercase;color:var(--muted)}
</style>
</head>
<body>
<div class="page" id="app">
  <header class="header">
    <div>
      <div class="wordmark">Mani&egrave;re <em>de</em> Voir</div>
      <div class="sub-title">Global Delivery Intelligence &nbsp;&middot;&nbsp; MDV 2026</div>
    </div>
    <div class="header-actions">
      <div class="live-indicator" id="live-indicator">
        <span class="live-dot"></span>
        <span id="sync-time">Live</span>
      </div>
      <button class="btn" id="btn-refresh" onclick="loadData()" disabled>&#8635; Refresh</button>
      <button class="btn btn-primary" id="btn-export" onclick="exportXLSX()" disabled>&#8595; Export</button>
    </div>
  </header>

  <div class="loading-screen" id="loading-screen">
    <div class="loading-word">Loading</div>
    <div class="loading-sub" id="loading-sub">Connecting to Airtable</div>
  </div>

  <div class="error-box" id="error-box" style="display:none"></div>

  <div id="main-content" style="display:none">
    <div class="kpi-strip" id="kpi-strip"></div>
    <nav class="region-nav" id="region-nav"></nav>
    <div id="region-panel"></div>
  </div>

  <footer class="footer">
    <div class="footer-brand">Mani&egrave;re <em>de</em> Voir</div>
    <div class="footer-note" id="footer-note">Global Delivery Intelligence Platform</div>
  </footer>
</div>

<script>
const REGIONS=['UK','EU','US','Dubai','Flannels'];
const FLAGS={UK:'🇬🇧',EU:'🇪🇺',US:'🇺🇸',Dubai:'🇦🇪',Flannels:'🏬'};
const TODAY=new Date();TODAY.setHours(0,0,0,0);
let allData=[],byRegion={},activeTab='UK';

async function loadData(){
  show('loading-screen');hide('main-content');hide('error-box');
  setDisabled('btn-refresh',true);setDisabled('btn-export',true);
  setText('loading-sub','Fetching live shipment data');
  try{
    const resp=await fetch('/api/shipments');
    if(!resp.ok){const e=await resp.json().catch(()=>({error:resp.statusText}));throw new Error(e.error||'HTTP '+resp.status);}
    const json=await resp.json();
    if(json.error)throw new Error(json.error);
    allData=(json.shipments||[]).map(r=>{
      const dest=(r.dest||'').trim().replace(/\s+$/,'');
      const eta=r.eta?new Date(r.eta):null;if(eta)eta.setHours(0,0,0,0);
      const unitsDue=parseFloat(r.unitsDue||0),unitsRcvd=parseFloat(r.unitsRcvd||0);
      const days=eta?Math.round((eta-TODAY)/86400000):null;
      const overdue=!!(eta&&days<0&&unitsRcvd===0);
      let deliveryStatus='Pending';
      if(unitsDue>0&&unitsRcvd>=unitsDue)deliveryStatus='Complete';
      else if(unitsRcvd>0&&unitsRcvd<unitsDue)deliveryStatus='Short';
      else if(overdue)deliveryStatus='Overdue';
      const freightClean=(r.freight||'').replace(/[🛫🚢🚚✈⚓]/g,'').trim();
      return{...r,dest,eta,days,unitsDue,unitsRcvd,overdue,deliveryStatus,freightClean,disc:unitsRcvd-unitsDue,pct:unitsDue>0?unitsRcvd/unitsDue:0};
    });
    allData.sort((a,b)=>(!a.eta&&!b.eta)?0:(!a.eta?1:(!b.eta?-1:a.eta-b.eta)));
    byRegion={};REGIONS.forEach(r=>byRegion[r]=[]);
    allData.forEach(s=>{
      if(byRegion[s.dest]!==undefined)byRegion[s.dest].push(s);
      else if(s.dest.toLowerCase().includes('flannels'))byRegion['Flannels'].push(s);
    });
    const now=new Date();
    setText('sync-time','Live · '+now.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'}));
    document.getElementById('live-indicator').style.display='flex';
    setText('footer-note','Last synced '+now.toLocaleDateString('en-GB',{day:'2-digit',month:'long',year:'numeric'})+' · '+now.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'}));
    setDisabled('btn-refresh',false);setDisabled('btn-export',false);
    hide('loading-screen');show('main-content');render();
  }catch(e){
    hide('loading-screen');
    document.getElementById('error-box').textContent='Failed to load: '+e.message;
    show('error-box');setDisabled('btn-refresh',false);console.error(e);
  }
}

function render(){
  const total=allData.length,totalDue=allData.reduce((s,x)=>s+x.unitsDue,0),totalRcvd=allData.reduce((s,x)=>s+x.unitsRcvd,0),totalOd=allData.filter(x=>x.overdue).length;
  document.getElementById('kpi-strip').innerHTML=
    kpi('Total Shipments',total,'across all regions')+
    kpi('Units Due',fmt(totalDue),'global inbound')+
    kpi('Units Received',fmt(totalRcvd),pct(totalRcvd/Math.max(1,totalDue))+' fulfilment')+
    kpi('Overdue',totalOd,totalOd>0?'require immediate action':'all on track',totalOd>0)+
    kpi('On-Time Rate',pct((total-totalOd)/Math.max(1,total)),'shipment performance');
  document.getElementById('region-nav').innerHTML=REGIONS.map(r=>{
    const cnt=(byRegion[r]||[]).length,hasOd=(byRegion[r]||[]).some(x=>x.overdue);
    return '<button class="r-tab'+(r===activeTab?' active':'')+'" onclick="switchTab(\''+r+'\')">'+
      (FLAGS[r]||'')+'&nbsp;'+r+' <span class="r-tab-count">'+cnt+'</span>'+(hasOd?'<span class="r-tab-alert">●</span>':'')+
    '</button>';
  }).join('');
  renderRegion();
}

function kpi(label,val,sub,alert){
  return '<div class="kpi-cell"><div class="kpi-label">'+label+'</div><div class="kpi-value'+(alert?' alert':'')+'">'+val+'</div>'+(sub?'<div class="kpi-sub">'+sub+'</div>':'')+'</div>';
}

function switchTab(r){
  activeTab=r;
  document.querySelectorAll('.r-tab').forEach(t=>{
    const on=t.textContent.includes(FLAGS[r]||r)&&t.textContent.trim().includes(r);
    t.classList.toggle('active',on);
    const c=t.querySelector('.r-tab-count');
    if(c){c.style.background=on?'var(--black)':'';c.style.color=on?'var(--white)':'';}
  });
  renderRegion();
}

function renderRegion(){
  const list=byRegion[activeTab]||[];
  const due=list.reduce((s,x)=>s+x.unitsDue,0),rcvd=list.reduce((s,x)=>s+x.unitsRcvd,0);
  const fp=due>0?rcvd/due:0;
  const od=list.filter(x=>x.overdue);
  const complete=list.filter(x=>x.deliveryStatus==='Complete').length;
  const pending=list.filter(x=>x.deliveryStatus==='Pending').length;
  const short=list.filter(x=>x.deliveryStatus==='Short').length;
  let h='<div class="region-metrics">'+
    '<div class="rm-cell"><div class="rm-label">Units Due</div><div class="rm-value">'+fmt(due)+'</div><div class="progress-bar"><div class="progress-fill" style="width:100%"></div></div></div>'+
    '<div class="rm-cell"><div class="rm-label">Units Received</div><div class="rm-value">'+fmt(rcvd)+'</div><div class="progress-bar"><div class="progress-fill" style="width:'+Math.min(100,fp*100).toFixed(1)+'%"></div></div></div>'+
    '<div class="rm-cell"><div class="rm-label">Fulfilment Rate</div><div class="rm-value">'+pct(fp)+'</div><div class="progress-bar"><div class="progress-fill" style="width:'+Math.min(100,fp*100).toFixed(1)+'%"></div></div></div>'+
    '<div class="rm-cell"><div class="rm-label">Shipment Status</div><div class="status-row">'+
      '<span class="status-chip" style="border-color:var(--success-border);color:var(--success);background:var(--success-bg)">&#10003; '+complete+' Complete</span>'+
      (short>0?'<span class="status-chip" style="border-color:var(--warning-border);color:var(--warning);background:var(--warning-bg)">~ '+short+' Short</span>':'')+
      '<span class="status-chip">&#8226; '+pending+' Pending</span>'+
      (od.length>0?'<span class="status-chip chip-danger">&#9888; '+od.length+' Overdue</span>':'')+
    '</div></div></div>';
  if(od.length>0){
    h+='<div class="overdue-alert"><span class="alert-icon">&#9888;</span><div><strong>'+od.length+' overdue shipment'+(od.length>1?'s':'')+' require immediate attention.</strong> '+od.map(s=>s.shId+' ('+s.liRef+')').join(', ')+'.</div></div>';
  }
  h+='<div class="table-container"><div class="table-topbar"><span class="table-caption">'+(FLAGS[activeTab]||'')+' '+activeTab+' &mdash; '+list.length+' Shipments &nbsp;&middot;&nbsp; Sorted by ETA</span>'+(od.length>0?'<span class="overdue-count-badge">'+od.length+' Overdue</span>':'')+'</div>'+
    '<div class="table-scroll"><table><thead><tr>'+
    '<th>LI Reference</th><th>SH Ref</th><th>Freight</th><th>ETA</th><th>Days</th>'+
    '<th>Units Due</th><th>Units Rcvd</th><th>Discrepancy</th><th>Fulfilment</th><th>Status</th><th>Shipment Status</th>'+
    '</tr></thead><tbody>';
  let tDue=0,tRcvd=0,prevWeek=null;
  list.forEach(s=>{
    tDue+=s.unitsDue;tRcvd+=s.unitsRcvd;
    if(s.eta){const wk=weekLabel(s.eta);if(wk!==prevWeek){prevWeek=wk;h+='<tr class="week-divider"><td colspan="11">'+wk+'</td></tr>';}}
    const etaStr=s.eta?s.eta.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'2-digit'}):'&mdash;';
    const dayStr=s.days!==null?(s.days<0?Math.abs(s.days)+'d late':s.days+'d'):'&mdash;';
    const fl=(s.freightClean||'').toLowerCase();
    const fCls=fl.includes('air')?'pill-air':fl.includes('sea')?'pill-sea':fl.includes('land')?'pill-land':'';
    const fIcon=fl.includes('air')?'&#9992;':fl.includes('sea')?'&#9875;':fl.includes('land')?'&#10145;':'';
    const stCls={Complete:'pill-complete',Short:'pill-short',Pending:'pill-pending',Overdue:'pill-overdue'}[s.deliveryStatus]||'pill-pending';
    const stIcon={Complete:'&#10003;',Short:'~',Pending:'&middot;',Overdue:'&#9888;'}[s.deliveryStatus]||'&middot;';
    const shipCls=s.status==='RECEIVED'?'pill-received':s.status==='SHIPPED'?'pill-shipped':'pill-quoting';
    const disc=s.disc!==0?'<span class="mono '+(s.disc<0?'neg':'pos')+'">'+(s.disc>0?'+':'')+fmt(s.disc)+'</span>':'<span class="mono muted">&mdash;</span>';
    h+='<tr class="'+(s.overdue?'row-overdue':'')+'">'+
      '<td><span class="li-ref">'+(s.liRef||'&mdash;')+'</span></td>'+
      '<td><span class="sh-ref">'+(s.shId||'&mdash;')+'</span></td>'+
      '<td><span class="pill '+fCls+'">'+fIcon+' '+(s.freightClean||'&mdash;')+'</span></td>'+
      '<td class="mono'+(s.overdue?' neg':'')+'">'+etaStr+'</td>'+
      '<td class="mono'+(s.overdue?' neg':'')+'">'+dayStr+'</td>'+
      '<td class="mono">'+fmt(s.unitsDue)+'</td>'+
      '<td class="mono">'+fmt(s.unitsRcvd)+'</td>'+
      '<td>'+disc+'</td>'+
      '<td><div class="mini-progress"><div class="mini-track"><div class="mini-fill" style="width:'+Math.min(100,s.pct*100).toFixed(0)+'%"></div></div><span class="mono">'+pct(s.pct)+'</span></div></td>'+
      '<td><span class="pill '+stCls+'">'+stIcon+' '+s.deliveryStatus+'</span></td>'+
      '<td><span class="pill '+shipCls+'" style="font-size:7px">'+(s.status||'&mdash;')+'</span></td>'+
    '</tr>';
  });
  h+='</tbody><tfoot><tr class="total-row"><td colspan="5" style="letter-spacing:0.18em;text-transform:uppercase">Totals</td><td class="mono">'+fmt(tDue)+'</td><td class="mono">'+fmt(tRcvd)+'</td><td class="mono '+(tRcvd-tDue<0?'neg':'')+'">'+
    (tRcvd-tDue!==0?(tRcvd-tDue>0?'+':'')+fmt(tRcvd-tDue):'&mdash;')+'</td><td class="mono">'+pct(tRcvd/Math.max(1,tDue))+'</td><td colspan="2"></td></tr></tfoot></table></div></div>';
  document.getElementById('region-panel').innerHTML=h;
}

function exportXLSX(){
  if(!window.XLSX||!allData.length)return;
  const wb=XLSX.utils.book_new();
  const raw=[['LI Reference','SH ID','Destination','Freight','ETA','Units Due','Units Received','Discrepancy','Fulfilment %','Delivery Status','Shipment Status']];
  allData.forEach(s=>raw.push([s.liRef,s.shId,s.dest,s.freightClean,s.eta?s.eta.toLocaleDateString('en-GB'):'',s.unitsDue,s.unitsRcvd,s.disc,Math.round(s.pct*100)+'%',s.deliveryStatus,s.status]));
  const wsRaw=XLSX.utils.aoa_to_sheet(raw);wsRaw['!cols']=[14,10,10,10,12,10,12,12,12,12,14].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb,wsRaw,'Raw Data');
  REGIONS.forEach(region=>{
    const l=byRegion[region]||[];
    const rows=[['Manière De Voir — '+region+' Delivery Dashboard'],['Generated: '+new Date().toLocaleDateString('en-GB',{weekday:'long',day:'2-digit',month:'long',year:'numeric'})],[],
      ['LI Reference','SH Ref','Freight','ETA','Days','Units Due','Units Received','Discrepancy','Fulfilment %','Delivery Status']];
    l.forEach(s=>rows.push([s.liRef,s.shId,s.freightClean,s.eta?s.eta.toLocaleDateString('en-GB'):'',s.days??'',s.unitsDue,s.unitsRcvd,s.disc,Math.round(s.pct*100)+'%',s.deliveryStatus]));
    const d=l.reduce((a,x)=>a+x.unitsDue,0),r=l.reduce((a,x)=>a+x.unitsRcvd,0);
    rows.push(['Totals','','','','',d,r,r-d,Math.round(r/Math.max(1,d)*100)+'%','']);
    const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=[14,10,10,12,8,10,12,12,12,14].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb,ws,region);
  });
  const sumRows=[['Manière De Voir — Global Delivery Intelligence Summary'],['Generated: '+new Date().toLocaleDateString('en-GB',{weekday:'long',day:'2-digit',month:'long',year:'numeric'})],[],
    ['Region','Shipments','Units Due','Units Received','Fulfilment %','Overdue','Complete','Short','Pending']];
  REGIONS.forEach(r=>{
    const l=byRegion[r]||[],d=l.reduce((a,x)=>a+x.unitsDue,0),rv=l.reduce((a,x)=>a+x.unitsRcvd,0);
    sumRows.push([r,l.length,d,rv,Math.round(rv/Math.max(1,d)*100)+'%',l.filter(x=>x.overdue).length,l.filter(x=>x.deliveryStatus==='Complete').length,l.filter(x=>x.deliveryStatus==='Short').length,l.filter(x=>x.deliveryStatus==='Pending').length]);
  });
  const d=allData.reduce((a,x)=>a+x.unitsDue,0),r=allData.reduce((a,x)=>a+x.unitsRcvd,0);
  sumRows.push(['GLOBAL TOTAL',allData.length,d,r,Math.round(r/Math.max(1,d)*100)+'%',allData.filter(x=>x.overdue).length,allData.filter(x=>x.deliveryStatus==='Complete').length,allData.filter(x=>x.deliveryStatus==='Short').length,allData.filter(x=>x.deliveryStatus==='Pending').length]);
  const wsSum=XLSX.utils.aoa_to_sheet(sumRows);wsSum['!cols']=[16,10,12,14,12,10,10,10,10].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb,wsSum,'Summary');
  XLSX.writeFile(wb,'MDV_Delivery_Intelligence_'+new Date().toISOString().slice(0,10)+'.xlsx');
}

function fmt(n){return Math.round(n||0).toLocaleString('en-GB');}
function pct(n){return Math.round((n||0)*100)+'%';}
function show(id){document.getElementById(id).style.display='block';}
function hide(id){document.getElementById(id).style.display='none';}
function setText(id,v){document.getElementById(id).textContent=v;}
function setDisabled(id,v){document.getElementById(id).disabled=v;}
function weekLabel(date){
  const d=new Date(date),day=d.getDay();
  const mon=new Date(d);mon.setDate(d.getDate()-((day+6)%7));
  const sun=new Date(mon);sun.setDate(mon.getDate()+6);
  const o={day:'2-digit',month:'short'};
  return 'Week of '+mon.toLocaleDateString('en-GB',o)+' – '+sun.toLocaleDateString('en-GB',o);
}
loadData();
<\/script>
</body>
</html>`;
}
