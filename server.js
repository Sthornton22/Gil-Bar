/* ============================================================================
   GIL-BAR PROPOSAL ENGINE — Render web service
   ----------------------------------------------------------------------------
   Same process as the offline single-file builder, byte-for-byte core:
     drop PDFs -> classify (selection pkg / unit cut sheet / drawing set)
     -> recognize (line-accurate text; any manufacturer's Model/Qty/Tag labels)
     -> design-basis resolution (cut-sheet "Models:" line, or Unit Load FULL
        vs Default) + sister-family comparison trap
     -> scope-gate default-EXCLUDE (by-others in type OR status)
     -> verbatim-precedent scoring (exact frame > series > manufacturer)
     -> styled Gil-Bar letterhead .docx (blank pricing, 20-clause T&C,
        red [CONFIRM] for anything the documents don't contain)
   No pricing is ever generated. Drafts only.
   ========================================================================== */
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const express = require('express');
const multer = require('multer');
const fflate = require('fflate');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');

/* ---- docx UMD loaded via the browser-branch factory (verified pattern) ---- */
const dm = { exports: {} };
(new Function('exports', 'module', 'self', 'globalThis', 'window',
  fs.readFileSync(path.join(__dirname, 'node_modules/docx/build/index.umd.js'), 'utf8')
))(dm.exports, dm, undefined, undefined, undefined);
const docxLib = dm.exports;

/* ---- assets ---- */
const PAYLOAD = JSON.parse(fs.readFileSync(path.join(__dirname, 'payload.json'), 'utf8'));
const TC = JSON.parse(fs.readFileSync(path.join(__dirname, 'tc.json'), 'utf8'));
const WM = fs.readFileSync(path.join(__dirname, 'gilbar_wordmark1.png'));
const WM_B64 = fs.readFileSync(path.join(__dirname, 'wm.b64'), 'utf8').trim();
const CORE_SRC = fs.readFileSync(path.join(__dirname, 'app_core.js'), 'utf8');

/* ---- per-request engine: fresh job state, no cross-request bleed ---- */
function newEngine() {
  const w = { PAYLOAD: PAYLOAD };
  (new Function('window', CORE_SRC))(w);
  w.GilBar.setDocx(docxLib);
  return w.GilBar;
}

/* ---- PDF reading with line reconstruction (verified in the browser tool) ---- */
function linesFromTokens(toks) {
  const sorted = toks.slice().sort((a, b) => (a.y - b.y) || (a.x - b.x));
  const lines = []; let cur = [], curY = null; const TOL = 3.5;
  sorted.forEach(o => {
    if (curY === null || Math.abs(o.y - curY) <= TOL) { cur.push(o); curY = (curY === null ? o.y : curY); }
    else { lines.push(cur.sort((a, b) => a.x - b.x).map(x => x.t).join(' ')); cur = [o]; curY = o.y; }
  });
  if (cur.length) lines.push(cur.sort((a, b) => a.x - b.x).map(x => x.t).join(' '));
  return lines.join('\n');
}
async function readPdf(buffer, name) {
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;
  const pagesText = []; let tokens = []; let readable = 0;
  for (let i = 1; i <= pdf.numPages; i++) {
    const pg = await pdf.getPage(i);
    const vp = pg.getViewport({ scale: 1 });
    const tc = await pg.getTextContent();
    const tk = [];
    tc.items.forEach(it => {
      const s = (it.str || '').trim(); if (!s) return;
      readable += s.length;
      tk.push({ t: s, x: it.transform[4], y: vp.height - it.transform[5] });
    });
    tokens = tokens.concat(tk);
    pagesText.push(linesFromTokens(tk));
  }
  return { tokens, pagesText, readable, pages: pdf.numPages, name };
}

/* ---- classification (same rules as the offline tool) ---- */
function classify(r) {
  const all = r.pagesText.join(' ');
  const hasModel = /\b(?:Unit\s+)?Model(?:\s*(?:No\.?|Number))?\s*[:=]\s*[A-Z]{2,5}-?\d/i.test(all);
  const hasQty = /\b(?:Qty|Quantity)\s*[:=]\s*\d/i.test(all);
  const perfTell = /Performance Data|Unit Electrical Data|Selected Options|Entering Conditions|Capacity/i.test(all);
  if (hasModel && (hasQty || perfTell)) return 'selection';
  if (/Dimensional Data|Dimensional Tables/i.test(all) && /Models?:\s*[A-Z]{2,5}/.test(all)) return 'dwg';
  return 'drawing';
}
const classifyLabel = c => c === 'selection' ? 'selection package'
  : c === 'dwg' ? 'unit cut sheet (design-basis hint)' : 'drawing set';

/* ---- the pipeline for one request's file set ---- */
const XLSX = require('xlsx');
async function runPipeline(files) {
  const pdfFiles = files.filter(f => /\.pdf$/i.test(f.originalname));
  const sheetFiles = files.filter(f => /\.(xlsx|xls|csv)$/i.test(f.originalname));

  const reads = [];
  for (const f of pdfFiles) reads.push(await readPdf(f.buffer, f.originalname));
  let costSheets = [];
  const costDiag = [];
  for (const f of sheetFiles) {
    const wb = XLSX.read(f.buffer, { type: 'buffer' });
    costSheets = costSheets.concat(wb.SheetNames.map(n => ({
      name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' })
    })));
    costDiag.push({ name: f.originalname, kind: 'cost sheet (' + wb.SheetNames.length + ' tab' + (wb.SheetNames.length > 1 ? 's' : '') + '; pricing columns skipped by design)', pages: wb.SheetNames.length, readable: 1, scrambled: false });
  }

  const G = newEngine();
  const sel = reads.filter(r => classify(r) === 'selection');
  const dwg = reads.filter(r => classify(r) === 'dwg');
  const draw = reads.filter(r => classify(r) === 'drawing');

  let pages = []; sel.forEach(r => { pages = pages.concat(r.pagesText); });
  const dwgText = dwg.map(r => r.pagesText.join('\n')).join('\n');

  if (costSheets.length || sel.length) {
    G.ingestJob({ selectionPages: pages, dwgText: dwgText, costSheets: costSheets });
  } else {
    let toks = []; draw.forEach(r => { toks = toks.concat(r.tokens); });
    G.ingestDrawingTokens(toks, '');
  }

  const diag = costDiag.concat(reads.map(r => ({
    name: r.name, kind: classifyLabel(classify(r)), pages: r.pages,
    readable: r.readable, scrambled: r.readable < 40
  })));

  /* build the docx via the verified builder + fontTable patch */
  let docBuf = null;
  if (G.job.items.length) {
    const doc = G.buildDoc(WM, TC);
    const raw = await docxLib.Packer.toBuffer(doc);
    const zip = fflate.unzipSync(new Uint8Array(raw));
    if (zip['word/fontTable.xml']) delete zip['word/fontTable.xml'];
    if (zip['[Content_Types].xml']) {
      let ct = fflate.strFromU8(zip['[Content_Types].xml']);
      ct = ct.replace(/<Override[^>]*fontTable\.xml[^>]*\/>/g, '');
      zip['[Content_Types].xml'] = fflate.strToU8(ct);
    }
    if (zip['word/_rels/document.xml.rels']) {
      let rl = fflate.strFromU8(zip['word/_rels/document.xml.rels']);
      rl = rl.replace(/<Relationship[^>]*fontTable\.xml[^>]*\/>/g, '');
      zip['word/_rels/document.xml.rels'] = fflate.strToU8(rl);
    }
    docBuf = Buffer.from(fflate.zipSync(zip));
  }
  return { job: G.job, diag, docBuf };
}

/* ---- in-memory store for downloads + history (ephemeral on redeploys) ---- */
const STORE = new Map();   // id -> {name, buf, at}
const HISTORY = [];        // newest first: {id, project, items, flags, at}

/* ================================ web ==================================== */
const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 60 * 1024 * 1024, files: 12 } });

const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

const PAGE_TOP = title => `<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>${esc(title)}</title>
<style>
:root{--rust:#C0272D;--rust-d:#8A1A1E;--ink:#1b1b1e;--sub:#6b6b73;--navy:#1F3D5C;--paper:#faf8f5;--card:#fff;--line:#e7e2db;--zebra:#f6f3ef;--amber:#b7791f;--amber-bg:#fdf6e3}
*{box-sizing:border-box}html,body{margin:0;background:var(--paper);color:var(--ink);font-family:"Archivo","Segoe UI",system-ui,sans-serif;line-height:1.5}
.mono{font-family:"IBM Plex Mono",ui-monospace,Menlo,Consolas,monospace}
.wrap{max-width:940px;margin:0 auto;padding:26px 22px 80px}
header.top{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:12px 0 16px;border-bottom:3px solid var(--rust)}
header.top .brand{display:flex;align-items:center;gap:14px}header.top img{height:46px}
header.top nav a{color:var(--navy);text-decoration:none;font-weight:700;font-size:13px;margin-left:16px}
header.top nav a:hover{color:var(--rust)}
.eyebrow{font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:var(--rust);font-weight:700;margin-top:20px}
h1{font-size:25px;margin:6px 0 2px}.lede{color:var(--sub);font-size:14px;margin:0 0 18px;max-width:62ch}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px;margin:14px 0}
.card h2{font-size:13px;letter-spacing:.1em;text-transform:uppercase;color:var(--navy);margin:0 0 12px}
.btn{appearance:none;border:0;border-radius:8px;padding:11px 18px;font:inherit;font-weight:700;background:var(--rust);color:#fff;cursor:pointer;text-decoration:none;display:inline-block}
.btn:hover{background:var(--rust-d)}
.grid{width:100%;border-collapse:collapse;font-size:13px;margin-top:6px}
.grid th{background:var(--rust);color:#fff;text-align:left;padding:8px 10px;font-size:11px;letter-spacing:.06em}
.grid th:first-child{background:var(--rust-d)}
.grid td{border-bottom:1px solid var(--line);padding:7px 9px}.grid tr:nth-child(even) td{background:var(--zebra)}
.tag{color:var(--rust);font-weight:700}
.excl{font-size:13px;color:var(--sub);margin:5px 0;padding-left:14px;border-left:3px solid var(--line)}.excl b{color:var(--ink)}
.flag{font-size:13px;color:var(--rust);margin:6px 0;padding-left:14px;border-left:3px solid var(--rust)}
.note{font-size:12px;color:var(--sub);margin-top:8px}
.drop{border:2px dashed var(--line);border-radius:12px;padding:34px 20px;text-align:center;background:#fff}
.drop.drag{border-color:var(--rust);background:#fff8f7}
.chips{display:flex;flex-wrap:wrap;gap:8px;justify-content:center;margin-top:14px}
.chip{display:inline-flex;align-items:center;gap:8px;background:#fff;border:1px solid var(--line);border-radius:20px;padding:6px 12px;font-size:12px}
.basis{color:var(--navy);font-weight:600;font-size:13px;margin:0 0 10px}
footer.foot{margin-top:28px;font-size:12px;color:var(--sub);border-top:1px solid var(--line);padding-top:14px}
.hist td{font-size:13px}
.spin{display:none;margin-top:14px;color:var(--sub);font-size:13px}
.mfrs{margin-top:40px;text-align:center}
.mfrs h2{font-size:13px;letter-spacing:.1em;text-transform:uppercase;color:var(--navy);margin:0 0 14px}
.mfrgrid{display:flex;flex-wrap:wrap;gap:8px;justify-content:center}
.mfr{display:inline-flex;align-items:center;gap:7px;background:#fff;border:1px solid var(--line);border-radius:20px;padding:5px 12px 5px 6px;font-size:12px;color:var(--ink)}
.mfr .mlogo{width:22px;height:22px;border-radius:50%;object-fit:contain;background:#fff;border:1px solid var(--line)}
.mfr .mlogo.wide{width:auto;max-width:44px;height:22px;border-radius:6px;border:1px solid var(--line);background:#fff;padding:2px}
.mfr .mmono{width:22px;height:22px;border-radius:50%;background:var(--rust);color:#fff;display:inline-flex;align-items:center;justify-content:center;font-size:10px;font-weight:800}
.hero{text-align:center;margin:26px 0 8px}
.hero img{height:64px}
.hero h1{font-size:34px;margin:18px 0 10px;color:var(--ink);font-weight:800}
.hero h1 em{color:var(--rust);font-style:italic}
.hero .lede{color:var(--sub);font-size:14px;max-width:62ch;margin:0 auto 26px}
.slots{display:flex;gap:14px;justify-content:center;flex-wrap:wrap;margin-bottom:16px}
.slot{flex:1 1 300px;max-width:380px;background:#fff;border:1.5px dashed var(--line);border-radius:12px;padding:22px 18px;text-align:center;cursor:pointer}
.slot.drag{border-color:var(--rust);background:#fdf4f4}
.slot b{display:block;font-size:14px;color:var(--ink);margin-bottom:6px}
.slot b .req{color:var(--rust)}
.slot .sub{font-size:12px;color:var(--sub);line-height:1.5}
.slot .chips{margin-top:10px}
.genrow{text-align:center;margin:6px 0 4px}
</style></head><body><div class="wrap">
<header class="top"><div class="brand"><img src="/static/wordmark.png" alt="Gil-Bar — An Ambient Company"></div>
<nav><a href="/">New proposal</a><a href="/history">History</a></nav></header>`;
const PAGE_BOTTOM = `<footer class="foot">Drafts only — every proposal is completed by an engineer before it is sent. No pricing is ever generated.</footer>
</div></body></html>`;

app.get('/static/wordmark.png', (req, res) => { res.type('png').send(WM); });

app.get('/', (req, res) => {
  res.send(PAGE_TOP('Gil-Bar Proposal Engine') + `
<div class="hero">
  <img src="/static/wordmark.png" alt="Gil-Bar">
  <h1>The Gil-Bar <em>Proposal Engine</em></h1>
  <p class="lede">Drop in the job files — engineer, project, scope, includes, and QA are read off the drawings automatically and flagged if anything is uncertain.</p>
</div>
<form id="f" method="post" action="/generate" enctype="multipart/form-data">
  <div class="slots">
    <div class="slot" id="slotA">
      <b>Mechanical schedule <span class="req">*</span></b>
      <div class="sub">PDF of the drawing set with the schedules<br>(or an extracted equipment CSV)</div>
      <input type="file" id="fA" name="files" accept=".pdf,.csv" multiple style="display:none">
      <div class="chips" id="chipsA"></div>
    </div>
    <div class="slot" id="slotB">
      <b>Cost sheet / factory selection</b>
      <div class="sub">PDF or Excel — recommended;<br>source of truth for models &amp; includes</div>
      <input type="file" id="fB" name="files" accept=".pdf,.xlsx,.xls,.csv" multiple style="display:none">
      <div class="chips" id="chipsB"></div>
    </div>
  </div>
  <div class="genrow">
    <button type="submit" class="btn" id="go" disabled>Generate draft proposal</button>
    <div class="spin" id="spin">Reading schedules, scope-gating, building the draft, running QA — about 30–60 seconds…</div>
  </div>
</form>
<div class="mfrs"><h2>Manufacturers we represent</h2><div class="mfrgrid" id="mfrGrid"></div></div>
<script>
function wire(slotId, inputId, chipsId){
  var slot=document.getElementById(slotId), input=document.getElementById(inputId), chips=document.getElementById(chipsId);
  var staged=new DataTransfer();
  function render(){chips.innerHTML='';for(var i=0;i<staged.files.length;i++){var s=document.createElement('span');s.className='chip';s.textContent=staged.files[i].name;chips.appendChild(s);}input.files=staged.files;update();}
  function add(list){for(var i=0;i<list.length;i++){var f=list[i];if(!/\\.(pdf|xlsx|xls|csv)$/i.test(f.name))continue;var dup=false;for(var j=0;j<staged.files.length;j++)if(staged.files[j].name===f.name&&staged.files[j].size===f.size)dup=true;if(!dup)staged.items.add(f);}render();}
  slot.addEventListener('click',function(e){if(e.target.tagName!=='INPUT')input.click();});
  input.addEventListener('change',function(){add(input.files);});
  ['dragenter','dragover'].forEach(function(ev){slot.addEventListener(ev,function(e){e.preventDefault();e.stopPropagation();slot.classList.add('drag');});});
  ['dragleave','drop'].forEach(function(ev){slot.addEventListener(ev,function(e){e.preventDefault();e.stopPropagation();slot.classList.remove('drag');});});
  slot.addEventListener('drop',function(e){add(e.dataTransfer.files);});
  return function(){return staged.files.length;};
}
var go=document.getElementById('go');
var counts=[];
function update(){var n=0;counts.forEach(function(c){n+=c();});go.disabled=!n;}
counts.push(wire('slotA','fA','chipsA'));
counts.push(wire('slotB','fB','chipsB'));
document.getElementById('f').addEventListener('submit',function(){document.getElementById('spin').style.display='block';go.disabled=true;});
var MFRS=[
 {n:'AAON',d:'aaon.com'},{n:'Samsung',d:'samsunghvac.com'},{n:'York / JCI',d:'johnsoncontrols.com'},
 {n:'Hitachi',d:'hitachiaircon.com'},{n:'Lennox',d:'lennox.com',u:'https://www.lennox.com/application/themes/lennox/assets/global/lennox_logo.svg'},{n:'ClimateMaster',d:'climatemaster.com'},
 {n:'ClimaCool',d:'climacoolcorp.com'},{n:'IEC',d:'iec-okc.com'},{n:'Dadanco',d:'dadanco.com',u:'https://www.dadanco.com/Themes/DadancoTheme/Content/dadanco_logo-01.svg'},
 {n:'MultiStack',d:'multistack.com'},{n:'Mammoth',d:'mammoth-inc.com'},{n:'Temtrol',d:'nortekair.com'},
 {n:'AboveAir',d:'aboveair.com'},{n:'Omega',d:'omega-heatpump.com',u:'https://omega-heatpump.com/assets/images/newspxomegalogo.png'},{n:'United CoolAir',d:'unitedcoolair.com'},
 {n:'Friedrich',d:'friedrich.com'},{n:'Smardt',d:'smardt.com',u:'https://www.michiganair.com/user_area/content_media/raw/smardt.jpg'},{n:'Motivair',d:'motivaircorp.com'},
 {n:'SEMCO',d:'semcohvac.com',u:'https://cmswa.com/wp-content/uploads/2017/01/flktgroup-semco-logo-website-small.png'},{n:'Dectron',d:'dectron.com'},{n:'Armstrong',d:'armstrongfluidtechnology.com'},
 {n:'ClimateCraft',d:'climatecraft.com'},{n:'Data Aire',d:'dataaire.com',u:'https://www.dataaire.com/wp-content/uploads/logo_horizontal_grey_355x69.png'},{n:'Valent',d:'valentair.com'},
 {n:'Sigma',d:'sigmaproducts.com',u:'https://sigmaproducts.com/assets/images/newspxsigmalogo.png'},{n:'RAE',d:'rae-coils.com',u:'https://rae-coils.com/wp-content/uploads/2023/04/NEW-RAE-Coils-logo.png'},{n:'Brasch',d:'braschmfg.com',u:'https://braschmfg.com/wp-content/uploads/2017/11/Brasch_Long_Logo4.png'},
 {n:'Haakon',d:'haakon.com'},{n:'Innovent',d:'innoventair.com'},{n:'BasX',d:'basxsolutions.com'},
 {n:'Governair',d:'nortekair.com'},{n:'Venmar',d:'venmarces.com'}
];
(function(){
 var g=document.getElementById('mfrGrid'); if(!g) return;
 MFRS.forEach(function(m){
  var el=document.createElement('span'); el.className='mfr';
  var img=document.createElement('img'); img.className='mlogo';
  img.alt=m.n; img.loading='lazy'; img.referrerPolicy='no-referrer';
  img.src=m.u || ('https://www.google.com/s2/favicons?domain='+m.d+'&sz=128');
  if(m.u) img.classList.add('wide');
  var mono=document.createElement('span'); mono.className='mmono'; mono.style.display='none';
  mono.textContent=m.n.replace(/[^A-Za-z ]/g,'').split(/[\\s/]+/).map(function(w){return w[0];}).join('').slice(0,2).toUpperCase();
  img.addEventListener('error',function(){ img.style.display='none'; mono.style.display='inline-flex'; });
  el.appendChild(img); el.appendChild(mono); el.appendChild(document.createTextNode(m.n));
  g.appendChild(el);
 });
})();
</script>` + PAGE_BOTTOM);
});

app.post('/generate', upload.array('files', 12), async (req, res) => {
  try {
    const files = (req.files || []).filter(f => /\.(pdf|xlsx|xls|csv)$/i.test(f.originalname));
    if (!files.length) return res.status(400).send(PAGE_TOP('No files') + '<div class="card">No readable files received (PDF, XLSX, XLS, or CSV). <a href="/">Back</a></div>' + PAGE_BOTTOM);
    const { job, diag, docBuf } = await runPipeline(files);

    let id = null;
    if (docBuf) {
      id = crypto.randomBytes(8).toString('hex');
      const fname = (job.header.projectName || 'GilBar').replace(/[^A-Za-z0-9]+/g, '_').replace(/^_|_$/g, '') + '_Proposal.docx';
      STORE.set(id, { name: fname, buf: docBuf, at: new Date() });
      HISTORY.unshift({ id, project: job.header.projectName || '(unnamed)', items: job.items.length, flags: job.flags.length, at: new Date() });
      if (HISTORY.length > 200) HISTORY.pop();
    }

    const basis = (job.flags || []).find(f => /^Design-basis /.test(f));
    const flags = (job.flags || []).filter(f => !/^Design-basis /.test(f));
    res.send(PAGE_TOP('Draft — ' + (job.header.projectName || 'proposal')) + `
<div class="eyebrow">Draft result</div>
<h1>${esc(job.header.projectName || '[CONFIRM — Project Name]')}</h1>
${basis ? `<p class="basis">${esc(basis)}</p>` : ''}
<div class="card"><h2>Quoted</h2>
${job.items.length ? `<table class="grid"><thead><tr><th>TAG</th><th>MANUFACTURER</th><th>MODEL</th><th>VOLTAGE</th><th>QTY</th></tr></thead><tbody>
${job.items.map(it => `<tr><td class="tag">${esc(it.tag || it.model)}</td><td>${esc(it.manufacturer || '[CONFIRM]')}</td><td>${esc(it.model)}</td><td>${esc(it.voltage || '[CONFIRM]')}</td><td>${esc(it.qty || '[CONFIRM]')}</td></tr>`).join('')}
</tbody></table>` : '<div class="note">Nothing quoted — no confirmed Gil-Bar line was recognized in these documents.</div>'}
</div>
<div class="card"><h2>Recognized — not quoted</h2>
${job.excluded.length ? job.excluded.map(e => `<div class="excl"><b>${esc(e.model)}</b> — ${esc(e.reason)}</div>`).join('') : '<div class="note">Nothing excluded.</div>'}
</div>
<div class="card"><h2>Flags / to confirm</h2>
${flags.length ? flags.map(f => `<div class="flag">${esc(f)}</div>`).join('') : '<div class="note">No open flags.</div>'}
</div>
<div class="card"><h2>Files read</h2>
${diag.map(d => `<div class="note"><b>${esc(d.name)}</b> — ${esc(d.kind)}, ${d.pages} page(s), ${d.readable} readable chars${d.scrambled ? ' <span style="color:var(--rust)">(scrambled/image-only — OCR needed; not readable as text)</span>' : ''}</div>`).join('')}
</div>
${id ? `<a class="btn" href="/download/${id}">Download proposal .docx</a>` : ''}
<a class="btn" style="background:#fff;color:var(--rust);border:1px solid var(--rust);margin-left:10px" href="/">Start another</a>
` + PAGE_BOTTOM);
  } catch (e) {
    console.error(e);
    res.status(500).send(PAGE_TOP('Error') + '<div class="card">Could not process these files. <a href="/">Back</a></div>' + PAGE_BOTTOM);
  }
});

app.get('/download/:id', (req, res) => {
  const rec = STORE.get(req.params.id);
  if (!rec) return res.status(404).send('Not found (drafts are kept in memory and clear on redeploy).');
  res.setHeader('Content-Disposition', `attachment; filename="${rec.name}"`);
  res.type('application/vnd.openxmlformats-officedocument.wordprocessingml.document').send(rec.buf);
});

app.get('/history', (req, res) => {
  res.send(PAGE_TOP('History') + `
<div class="eyebrow">History</div><h1>Recent drafts</h1>
<p class="lede">Drafts generated by this running instance. This list is in-memory and clears when the service restarts or redeploys.</p>
<div class="card">${HISTORY.length ? `<table class="grid hist"><thead><tr><th>WHEN</th><th>PROJECT</th><th>ITEMS</th><th>FLAGS</th><th></th></tr></thead><tbody>
${HISTORY.map(h => `<tr><td>${esc(h.at.toISOString().replace('T', ' ').slice(0, 16))}</td><td>${esc(h.project)}</td><td>${h.items}</td><td>${h.flags}</td><td>${STORE.has(h.id) ? `<a href="/download/${h.id}">download</a>` : ''}</td></tr>`).join('')}
</tbody></table>` : '<div class="note">Nothing yet.</div>'}</div>` + PAGE_BOTTOM);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Gil-Bar Proposal Engine listening on :' + PORT));
