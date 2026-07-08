/* ============================================================================
   GIL-BAR PROPOSAL BUILDER — app core
   Baked workbook lives in window.PAYLOAD; libraries (docx/xlsx/pdfjs) inlined.
   Two ingest paths:
     (A) Drawing set  — rasterize/OCR, column-aware recognize, scope-gate
     (B) Selection pkg — clean text, design-basis resolution + sister-family trap
   No pricing ever. Tier-1 header fields never guessed (render [CONFIRM ...]).
   ========================================================================== */
(function () {
  'use strict';
  var P = window.PAYLOAD || {};
  var D = null; // docx, set on init

  /* ------------------------------------------------------------- state */
  var job = {
    header: {
      projectName: '', projectAddress: '', mepEngineer: '', architect: '',
      submittedTo: '', proposalNo: '', proposalDate: '', validUntil: '',
      drawingsDate: '', entityName: 'Gil-Bar Industries'
    },
    items: [],      // quoted items: {tag,manufacturer,model,voltage,qty,series,includes,notIncluded,_amberVolt}
    excluded: [],   // recognized-but-not-quoted: {model,series,reason}
    flags: []       // red confirm/flag strings
  };
  window.__job = job;

  /* ------------------------------------------------------------- helpers */
  function up(s) { return String(s == null ? '' : s).toUpperCase().trim(); }
  function seriesOf(model) {
    var m = up(model).match(/^([A-Z]{2,4})[- ]?\d/);
    return m ? m[1] : up(model).replace(/[^A-Z].*$/, '');
  }
  function mapRowForSeries(series) {
    var rows = P.manufacturerMap || [], S = up(series), i, hit = null;
    for (i = 0; i < rows.length; i++) {
      var pfx = up(rows[i].prefix).split(/[\/, (]/)[0];
      if (pfx === S) { hit = rows[i]; break; }
    }
    if (!hit) for (i = 0; i < rows.length; i++) {
      if (S && up(rows[i].prefix).indexOf(S) === 0) { hit = rows[i]; break; }
    }
    return hit;
  }
  function isByOthers(typeStr) {
    var t = up(typeStr);
    // "HEAT PUMP" / "SOURCE HEAT PUMP" is IN-scope equipment — never let the
    // "PUMP" by-others keyword catch it (the classic false positive).
    var pumpsOnly = t.replace(/HEAT\s+PUMP/g, 'HEAT-UNIT');
    var kw = ['GRILLE', 'REGISTER', 'DIFFUSER', 'LOUVER', 'VAV', 'FAN-POWERED',
      'FAN POWERED', 'DUCT HEATER', 'EXHAUST FAN', 'ENERGY RECOVERY VENT'];
    for (var i = 0; i < kw.length; i++) if (t.indexOf(kw[i]) >= 0) return true;
    // standalone pumps (circulating/condensate/glycol), not heat pumps
    if (/\bPUMPS?\b/.test(pumpsOnly)) return true;
    // standalone ERV (energy recovery ventilator) but not "energy recovery unit/wheel"
    if (/\bERV\b/.test(t) || /ENERGY RECOVERY VENTILATOR/.test(t)) return true;
    return false;
  }

  /* ---- sister-family comparison pairs (documented in the trap map) ---- */
  var COMPARISON_PAIRS = [{
    a: 'SY', b: 'SC', family: 'ClimateMaster',
    note: 'SC and SY are distinct ClimateMaster lines selected side-by-side ' +
      '(SC = higher-total/latent, SY = higher-sensible/low-leaving-air). ' +
      'Only the design-basis series is quoted.'
  }];
  function comparisonPartner(series) {
    var S = up(series);
    for (var i = 0; i < COMPARISON_PAIRS.length; i++) {
      var p = COMPARISON_PAIRS[i];
      if (p.a === S) return { partner: p.b, family: p.family, note: p.note };
      if (p.b === S) return { partner: p.a, family: p.family, note: p.note };
    }
    return null;
  }

  /* =====================================================================
     PATH A — DRAWING SET  (row/column-aware recognize + scope-gate)
     Tokens: [{t, x, y}] as produced by the pdf.js text pass or OCR pass.
     ===================================================================== */
  function recognizePositioned(tokens) {
    var out = [], seen = {};
    // model-looking token: 2-4 letters, dash/space, 2-4 digits (+opt suffix)
    var MODEL = /^[A-Z]{2,4}[- ]?\d{2,4}[A-Z0-9-]*$/;
    for (var i = 0; i < tokens.length; i++) {
      var tk = tokens[i], t = up(tk.t);
      if (!MODEL.test(t)) continue;
      var series = seriesOf(t);
      var mr = mapRowForSeries(series);
      if (!mr) continue;                          // unknown series: not Gil-Bar, skip
      var key = t + '|' + Math.round(tk.y);
      if (seen[key]) continue; seen[key] = 1;
      // voltage: search same-row band AND same-column band, join fragments
      var volt = findVoltage(tokens, tk);
      var qty = findQty(tokens, tk);
      var tag = findTag(tokens, tk);
      out.push({
        tag: tag || '', manufacturer: mr.manufacturer, model: t.replace(/\s+/g, ''),
        series: series, type: mr.type, status: mr.status,
        voltage: volt, qty: qty, _tagFrame: !tag
      });
    }
    return out;
  }
  function bandTokens(tokens, tk) {
    var ytol = 8, xtol = Math.max(60, String(tk.t).length * 9);
    var row = [], col = [];
    for (var i = 0; i < tokens.length; i++) {
      var o = tokens[i];
      if (Math.abs(o.y - tk.y) <= ytol) row.push(o);
      if (Math.abs(o.x - tk.x) <= xtol) col.push(o);
    }
    row.sort(function (a, b) { return a.x - b.x; });
    col.sort(function (a, b) { return a.y - b.y; });
    return { row: row, col: col };
  }
  function joinBand(arr) { return arr.map(function (o) { return o.t; }).join(' '); }
  function findVoltage(tokens, tk) {
    var b = bandTokens(tokens, tk);
    var re = /(\d{3})\s*[\/-]\s*(\d)\s*[\/-]\s*(\d{2})/;
    var m = joinBand(b.row).match(re) || joinBand(b.col).match(re);
    return m ? (m[1] + ' / ' + m[2] + ' / ' + m[3]) : '';
  }
  function findQty(tokens, tk) {
    // standalone small integer in the same row or column, to the right/below
    var b = bandTokens(tokens, tk), best = '', i, o;
    for (i = 0; i < b.row.length; i++) {
      o = b.row[i]; var mm = String(o.t).match(/^\s*(\d{1,2})\s*$/);
      if (mm && +mm[1] >= 1 && +mm[1] <= 99 && o.x > tk.x) best = mm[1];
    }
    if (!best) for (i = 0; i < b.col.length; i++) {
      o = b.col[i]; var m2 = String(o.t).match(/^\s*(\d{1,2})\s*$/);
      if (m2 && +m2[1] >= 1 && +m2[1] <= 99 && o.y > tk.y) { best = m2[1]; break; }
    }
    return best;
  }
  function findTag(tokens, tk) {
    // tag-looking token (letters+dashes+digits, not a pure model) near the model
    var b = bandTokens(tokens, tk);
    var TAG = /^[A-Z]{1,4}-\d[\dA-Z.\-]*$/;
    for (var i = 0; i < b.row.length; i++) {
      var t = up(b.row[i].t);
      if (t !== up(tk.t) && TAG.test(t) && !/^[A-Z]{2,4}-?\d{2,4}$/.test(t)) return t;
    }
    return '';
  }

  /* scope-gate: default-EXCLUDE. Quote only Confirmed + not by-others. */
  function scopeGate(recognized) {
    var quoted = [], excluded = [];
    for (var i = 0; i < recognized.length; i++) {
      var r = recognized[i];
      var confirmed = /confirm/i.test(r.status);
      // "by others" can live in the STATUS text as well as the type
      // (e.g. "Confirmed (other trade — by others)") — either one excludes.
      var byOth = isByOthers(r.type) || /by[\s-]?others/i.test(r.status || '');
      if (confirmed && !byOth) quoted.push(r);
      else excluded.push({
        model: r.model, series: r.series,
        reason: byOth ? ('By-others (' + (r.type || r.status) + ') — recognized, not quoted')
          : ('Status "' + (r.status || 'unknown') + '" — recognized, confirm before quoting')
      });
    }
    return { quoted: quoted, excluded: excluded };
  }

  /* voltage self-heal from Manufacturer-Map hints (amber = verify). */
  function resolveFlags(items) {
    var f = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (!it.voltage) {
        var mr = mapRowForSeries(it.series);
        if (mr && mr.voltHints && mr.voltHints.length === 1) {
          it.voltage = mr.voltHints[0].replace(/\//g, ' / ');
          it._amberVolt = true;
          f.push('FLAG — ' + it.model + ' voltage auto-filled from Manufacturer Map (' +
            it.voltage + '); VERIFY against the schedule.');
        } else {
          f.push('FLAG — ' + it.model + ' voltage missing; enter from the schedule (never guess).');
        }
      }
      if (!it.qty) f.push('FLAG — ' + it.model + ' quantity missing; enter from the takeoff.');
    }
    return f;
  }

  /* =====================================================================
     PATH B — SELECTION / SUBMITTAL PACKAGE
     ===================================================================== */
  function readSelectionPackage(pages) {
    var rows = [];
    var MODEL_RE = /\b(?:Unit\s+)?Model(?:\s*(?:No\.?|Number))?\s*[:=]\s*([A-Z]{2,5}-?\d{2,4}[A-Z0-9\-]*)/ig;
    for (var p = 0; p < pages.length; p++) {
      var t = String(pages[p] || '');
      // Find EVERY model occurrence on the page (schedule-style submittals can
      // list several units per page). Each match gets its own text window from
      // just before it to the next match, so voltage/qty/tag pair to the right unit.
      var matches = [], m;
      MODEL_RE.lastIndex = 0;
      while ((m = MODEL_RE.exec(t)) !== null) matches.push({ idx: m.index, model: up(m[1]).replace(/\s+/g, '') });
      for (var i = 0; i < matches.length; i++) {
        // NON-OVERLAPPING windows: this window runs from this match to the next
        // match (or +1400 chars). A tag/mark printed just BEFORE the model label
        // is caught by scanning the slice between the previous window's end and
        // this model — but qty/voltage are only taken from the forward window,
        // so a neighbor unit's values can never leak in.
        var start = matches[i].idx;
        var end = (i + 1 < matches.length) ? matches[i + 1].idx : Math.min(t.length, matches[i].idx + 1400);
        var w = t.slice(start, end);
        var pre = t.slice(i > 0 ? matches[i - 1].idx : Math.max(0, start - 160), start);
        // keep only the LAST line of the pre-slice (the model's own line lead-in)
        var preLine = pre.slice(pre.lastIndexOf('\n') + 1);
        var volts = extractVoltage(w);
        var mQty = w.match(/\b(?:Qty|Quantity)\s*[:=]\s*(\d{1,3})\b/i);
        var TAG_RE = /\b(?:Unit\s+)?(?:Tag|Mark)\s*[:=]\s*(?!Qty\b|Quantity\b|Model\b)([A-Z0-9][A-Z0-9.\-]+)\b/i;
        var mTag = w.match(TAG_RE) || preLine.match(TAG_RE);
        var mLoad = w.match(/Unit Load:\s*(FULL|Default|[A-Za-z]+)/i);
        rows.push({
          tag: mTag ? up(mTag[1]) : '', manufacturer: '', model: matches[i].model,
          series: seriesOf(matches[i].model), voltage: volts, qty: mQty ? mQty[1] : '',
          unitLoad: mLoad ? up(mLoad[1]) : '', _fromSelection: true
        });
      }
    }
    return rows;
  }

  // Voltage extraction with PLAUSIBILITY VALIDATION: only accept real HVAC
  // nameplate combinations (standard voltages, 1/3 phase, 50/60 Hz). Iterates
  // candidates so a stray number pattern can't win over a real voltage.
  var VOLT_SET = { 110:1,115:1,120:1,200:1,208:1,220:1,230:1,240:1,265:1,277:1,380:1,400:1,415:1,460:1,480:1,575:1,600:1 };
  function validVolt(v, ph, hz) {
    return VOLT_SET[+v] && (ph === '1' || ph === '3') && (hz === '50' || hz === '60');
  }
  function extractVoltage(t) {
    var m, re;
    // 1) labeled "Voltage: 460/3/60" (most reliable)
    re = /(?:Voltage|Volts?|Elec(?:trical)?|V\/PH\/HZ)\s*[:=]?\s*(\d{3})\s*[\/-]\s*(\d)\s*[\/-]\s*(\d{2})/ig;
    while ((m = re.exec(t)) !== null) if (validVolt(m[1], m[2], m[3])) return m[1] + ' / ' + m[2] + ' / ' + m[3];
    // 2) bare "460-3-60" / "460/3/60" (V/PH/HZ)
    re = /\b(\d{3})\s*[\/-]\s*(\d)\s*[\/-]\s*(\d{2})\b/g;
    while ((m = re.exec(t)) !== null) if (validVolt(m[1], m[2], m[3])) return m[1] + ' / ' + m[2] + ' / ' + m[3];
    // 3) ClimateMaster size line "460/60/3" (V/HZ/PH) -> normalize to V/PH/HZ
    re = /\b(\d{3})\/(\d{2})\/(\d)\b/g;
    while ((m = re.exec(t)) !== null) if (validVolt(m[1], m[3], m[2])) return m[1] + ' / ' + m[3] + ' / ' + m[2];
    return '';
  }

  // Project name from a selection title line, tolerant of formats:
  // a "Project:"/"Job:" label anywhere on page 1-3, else the ClimateMaster
  // "QT...;0-<name> Climate Control Group" title line. Page text is
  // line-reconstructed, so the capture stops at the line end; a trailing
  // merged label (Date/Model/Tag/...) is stripped defensively anyway.
  function readProjectName(pages) {
    for (var p = 0; p < pages.length && p < 3; p++) {
      var s = String(pages[p] || '');
      var mL = s.match(/\b(?:Project(?:\s*Name)?|Job(?:\s*Name)?)\s*[:=]\s*([^\n]{2,80})/i);
      if (mL) {
        return mL[1]
          .replace(/\s+(?:Date|Model|Unit|Tag|Mark|Qty|Quantity|Engineer|Architect|Address|Location|Contractor|Submitted|Rev(?:ision)?)\b.*$/i, '')
          .replace(/\s{2,}.*$/, '').trim();
      }
      var m = s.match(/QT\d+;\d+\s*-\s*(.+?)\s+Climate Control Group/i);
      if (m) return m[1].trim();
      var m2 = s.match(/QT\d+;\d+\s*-\s*(.+)/);
      if (m2) return m2[1].replace(/\s+Climate Control Group.*$/i, '').replace(/\n.*$/s, '').trim();
    }
    return '';
  }

  // Decide the design-basis series from the document itself, no user input:
  //   1) any "Models: XX" line on a submitted unit cut sheet (DWG) wins;
  //   2) else the series whose sheets read "Unit Load: FULL" (selected case)
  //      vs the sister family reading "Default" (untouched comparison);
  //   3) else, if only one series present, that one.
  // Returns {series, reason} or {series:'', reason} if genuinely undecidable.
  function detectDesignBasis(rows, dwgText) {
    var seriesSeen = [];
    rows.forEach(function (r) { if (seriesSeen.indexOf(r.series) < 0) seriesSeen.push(r.series); });
    // (1) DWG "Models: SY" line
    if (dwgText) {
      var md = String(dwgText).match(/Models?:\s*\n?\s*([A-Z]{2,4})\b/);
      if (md && seriesSeen.indexOf(up(md[1])) >= 0)
        return { series: up(md[1]), reason: 'submitted unit cut sheet lists Models: ' + up(md[1]) };
    }
    // (2) Unit Load FULL vs Default: exactly ONE series is marked FULL, that
    // series is never also marked Default, and at least one OTHER series exists
    // (marked Default or unmarked). Any ambiguity -> no auto-pick.
    var full = {}, dflt = {};
    rows.forEach(function (r) {
      if (r.unitLoad === 'FULL') full[r.series] = 1;
      else if (r.unitLoad === 'DEFAULT') dflt[r.series] = 1;
    });
    var fullSeries = Object.keys(full);
    if (fullSeries.length === 1 && !dflt[fullSeries[0]] && seriesSeen.length > 1)
      return { series: fullSeries[0], reason: 'selection marks ' + fullSeries[0] + ' as Unit Load: FULL (selected case) vs the sister family at Default' };
    // (3) single series
    if (seriesSeen.length === 1) return { series: seriesSeen[0], reason: 'only one series present' };
    return { series: '', reason: '' };
  }
  function resolveDesignBasis(rows, designBasisSeries) {
    var flags = [], seriesSeen = [];
    rows.forEach(function (r) { if (seriesSeen.indexOf(r.series) < 0) seriesSeen.push(r.series); });
    var basis = up(designBasisSeries || '');
    if (!basis) {
      var pairPresent = null;
      COMPARISON_PAIRS.forEach(function (pp) {
        if (seriesSeen.indexOf(pp.a) >= 0 && seriesSeen.indexOf(pp.b) >= 0) pairPresent = pp;
      });
      if (pairPresent) {
        flags.push('FLAG — Selection shows both ' + pairPresent.a + ' and ' + pairPresent.b +
          ' side-by-side (' + pairPresent.family + '). Design-basis not confirmed — pick the ' +
          'quoted series from the schedule design-basis note; the other family is the comparison ' +
          'side and is NOT quoted.');
        return {
          quoted: [], flags: flags, excluded: rows.map(function (r) {
            return { model: r.model, series: r.series, reason: 'Comparison/selection side — design-basis unconfirmed' };
          })
        };
      }
      basis = seriesSeen[0] || '';
    }
    var partner = comparisonPartner(basis), quoted = [], excluded = [];
    rows.forEach(function (r) {
      if (r.series === basis) {
        var mr = mapRowForSeries(r.series);
        quoted.push({
          tag: r.tag, manufacturer: (mr && mr.manufacturer) || r.manufacturer || '',
          model: r.model, series: r.series, type: mr ? mr.type : '', status: mr ? mr.status : '',
          voltage: r.voltage, qty: r.qty, _tagFrame: !r.tag
        });
      } else if (partner && r.series === partner.partner) {
        excluded.push({
          model: r.model, series: r.series,
          reason: 'Comparison side of the selection run — ' + partner.family + ' ' +
            partner.partner + ' is the sister family to the ' + basis + ' design-basis; recognized but NOT quoted.'
        });
      } else {
        excluded.push({ model: r.model, series: r.series, reason: 'Not in the ' + basis + ' design-basis — recognized, confirm scope.' });
      }
    });
    if (partner && excluded.some(function (e) { return e.series === partner.partner; })) {
      flags.push('FLAG — ' + partner.note + ' Quoted design-basis: ' + basis +
        '. Comparison side left off: ' + partner.family + ' ' + partner.partner + '.');
    }
    // any-manufacturer safety: a quoted series not in the Manufacturer Map is
    // surfaced, never silently blank-filled — the user confirms the line.
    quoted.forEach(function (q) {
      if (!q.manufacturer || !mapRowForSeries(q.series)) {
        flags.push('FLAG — ' + q.model + ' (' + q.series + ') is not in the Gil-Bar Manufacturer Map. ' +
          'Manufacturer/scope not auto-resolved — confirm the line and whether Gil-Bar quotes it before issuing.');
      }
    });
    return { quoted: quoted, excluded: excluded, flags: flags };
  }
  // Nearest banked precedent, scored: an exact frame match in the Model Key
  // (e.g. quoted SY-036 vs a row keyed "SY-036 / ...") beats a generic series
  // match ("SY-series"), which beats same-manufacturer-only. `models` is the
  // list of quoted model strings for exact-frame scoring.
  function verbatimBasis(quotedSeries, manufacturer, jobQuoteNo, models) {
    var lib = P.verbatim || [], S = up(quotedSeries), M = up(manufacturer), Q = up(jobQuoteNo || '');
    var own = null, i;
    for (i = 0; i < lib.length; i++) if (Q && up(lib[i].source).indexOf(Q) >= 0) { own = lib[i]; break; }
    if (own) return { banked: true, row: own, flag: null };
    var mset = (models || []).map(up);
    var best = null, bestScore = -1;
    for (i = 0; i < lib.length; i++) {
      var r = lib[i];
      if (up(r.manufacturer) !== M) continue;
      var key = up(r.modelKey), score = 0;
      for (var j = 0; j < mset.length; j++) if (mset[j] && key.indexOf(mset[j]) >= 0) score += 4;
      if (S && key.indexOf(S) >= 0) score += 1;
      if (score > bestScore) { bestScore = score; best = r; }
    }
    var near = (best && bestScore > 0) ? best : null;
    if (!near) for (i = 0; i < lib.length; i++) if (up(lib[i].manufacturer) === M) { near = lib[i]; break; }
    var flag = 'FLAG — No banked verbatim Includes row cites this job\u2019s quote (' +
      (jobQuoteNo || '[CONFIRM — Quote No.]') + '). ' + (near
        ? 'Includes/Not-Included are carried from the nearest same-config banked precedent (' +
          near.source + '). Confirm against the issued proposal / order acknowledgment, then bank ' +
          'this job\u2019s wording verbatim.'
        : 'No same-series precedent banked — obtain the issued proposal for scope; do not reconstruct from a submittal or memory.');
    return { banked: false, row: near, flag: flag };
  }

  /* =====================================================================
     VERBATIM INCLUDES FETCH (drawing path, per quoted item)
     ===================================================================== */
  function attachIncludes(item, jobQuoteNo) {
    var vb = verbatimBasis(item.series, item.manufacturer, jobQuoteNo, [item.model]);
    if (vb.row) { item.includes = vb.row.includes.slice(); item.notIncluded = vb.row.notIncluded.slice(); item._basisSource = vb.row.source; }
    return vb.flag;
  }

  /* =====================================================================
     ORCHESTRATION
     ===================================================================== */
  function ingestDrawingTokens(tokens, jobQuoteNo) {
    var rec = recognizePositioned(tokens);
    var g = scopeGate(rec);
    job.items = g.quoted; job.excluded = g.excluded; job.flags = [];
    job.flags = job.flags.concat(resolveFlags(job.items));
    job.items.forEach(function (it) { var f = attachIncludes(it, jobQuoteNo); if (f) job.flags.push(f); });
    return job;
  }
  /* =====================================================================
     COST SHEET (Excel/CSV, parsed to arrays-of-arrays by the host)
     The cost sheet is the SOURCE OF TRUTH for the quoted set: which models,
     what tags, what quantities. Pricing columns are detected and SKIPPED —
     their values are never read into memory. Option text is NOT converted
     into Includes (that would bypass the verbatim-banked rule).
     ===================================================================== */
  var MODEL_TOKEN = /^[A-Z]{2,5}-?\d{2,4}[A-Z0-9\-]*$/;
  var PRICE_HDR = /COST|PRICE|SELL|NET|TOTAL|MARGIN|\$/i;
  function readCostSheetAOA(sheets) {
    // sheets: [{name, aoa}] where aoa = array of row-arrays of cell values
    var rows = [], headerFound = false;
    (sheets || []).forEach(function (sh) {
      var aoa = sh.aoa || [];
      // find a header row in the first 12 rows: a cell equal/starting "MODEL"
      var hdrIdx = -1, cols = {};
      for (var r = 0; r < Math.min(12, aoa.length); r++) {
        for (var c = 0; c < (aoa[r] || []).length; c++) {
          var v = up(aoa[r][c]);
          if (/^MODEL(\s*(NO\.?|NUMBER))?$/.test(v)) { hdrIdx = r; break; }
        }
        if (hdrIdx >= 0) break;
      }
      if (hdrIdx >= 0) {
        headerFound = true;
        (aoa[hdrIdx] || []).forEach(function (h, c) {
          var v = up(h);
          if (PRICE_HDR.test(v)) return;                     // pricing: never mapped
          if (/^MODEL/.test(v)) cols.model = c;
          else if (/^(TAG|MARK)/.test(v)) cols.tag = c;
          else if (/^(QTY|QUANTITY)/.test(v)) cols.qty = c;
          else if (/VOLT|V\/PH\/HZ|ELECTRICAL/.test(v)) cols.volt = c;
          else if (/^(MFR|MANUFACTURER|BRAND)/.test(v)) cols.mfr = c;
        });
        for (var r2 = hdrIdx + 1; r2 < aoa.length; r2++) {
          var row = aoa[r2] || [];
          var model = up(row[cols.model]);
          if (!MODEL_TOKEN.test(model)) continue;
          var qty = row[cols.qty] != null ? String(row[cols.qty]).trim() : '';
          if (!/^\d{1,3}$/.test(qty)) qty = '';
          rows.push({
            tag: cols.tag != null ? up(row[cols.tag]) : '',
            model: model, series: seriesOf(model),
            voltage: cols.volt != null ? extractVoltage(String(row[cols.volt] || '')) : '',
            qty: qty,
            mfrText: cols.mfr != null ? String(row[cols.mfr] || '').trim() : '',
            _fromCostSheet: true
          });
        }
      } else {
        // no header row: heuristic scan — any strict model token; qty = a small
        // integer in the same row; tag = a tag-shaped cell in the same row.
        for (var r3 = 0; r3 < aoa.length; r3++) {
          var rr = aoa[r3] || [];
          for (var c3 = 0; c3 < rr.length; c3++) {
            var mv = up(rr[c3]);
            if (!MODEL_TOKEN.test(mv)) continue;
            if (!mapRowForSeries(seriesOf(mv)) && !/\d{3}/.test(mv)) continue; // damp noise
            var qty2 = '', tag2 = '', volt2 = '';
            for (var c4 = 0; c4 < rr.length; c4++) {
              if (c4 === c3) continue;
              var cv = String(rr[c4] == null ? '' : rr[c4]).trim();
              if (!qty2 && /^\d{1,3}$/.test(cv) && +cv >= 1 && +cv <= 999) qty2 = cv;
              if (!tag2 && /^[A-Z]{1,4}-\d[\dA-Z.\-]*$/.test(up(cv)) && up(cv) !== mv) tag2 = up(cv);
              if (!volt2) volt2 = extractVoltage(cv);
            }
            rows.push({ tag: tag2, model: mv, series: seriesOf(mv), voltage: volt2, qty: qty2, mfrText: '', _fromCostSheet: true, _heuristic: true });
            break; // one model per row
          }
        }
      }
    });
    // dedupe identical model+tag rows (repeated across sheets/tabs)
    var seen = {}, out = [];
    rows.forEach(function (r) { var k = r.model + '|' + r.tag; if (!seen[k]) { seen[k] = 1; out.push(r); } });
    return { rows: out, headerFound: headerFound };
  }

  /* Unified ingest: cost sheet (authoritative when present) + selection pages
     (corroboration + comparison-family detection) + cut-sheet text. */
  function ingestJob(opts) {
    opts = opts || {};
    var selPages = opts.selectionPages || [];
    var dwgText = opts.dwgText || '';
    var cost = (opts.costSheets && opts.costSheets.length) ? readCostSheetAOA(opts.costSheets) : null;

    if (!cost || !cost.rows.length) {
      if (cost && !cost.rows.length && !selPages.length) {
        job.items = []; job.excluded = []; job.flags = [
          'FLAG — No equipment rows were recognized on the cost sheet (no Model column or model-shaped cells found). Check the file.'];
        return job;
      }
      return ingestSelection(selPages, dwgText); // no usable cost sheet: existing path
    }

    var flags = [], excluded = [], quoted = [];
    flags.push('Quoted set taken from the cost sheet (' + cost.rows.length + ' line(s)). Pricing columns were skipped by design and never read.');
    if (!cost.headerFound) flags.push('FLAG — Cost sheet had no recognizable header row; models/qty/tags were mapped heuristically. Verify every row.');

    // recognize + scope-gate each cost-sheet line
    cost.rows.forEach(function (r) {
      var mr = mapRowForSeries(r.series);
      var typ = mr ? mr.type : '', status = mr ? mr.status : '';
      var byOth = isByOthers(typ) || /by[\s-]?others/i.test(status || '');
      if (byOth) {
        excluded.push({ model: r.model, series: r.series, reason: 'On the cost sheet but a by-others line (' + (typ || status) + ') — recognized, NOT quoted. Confirm why it was cost-sheeted.' });
        flags.push('FLAG — ' + r.model + ' appears on the cost sheet but is a by-others type (' + (typ || status) + '). Left off the proposal — confirm.');
        return;
      }
      quoted.push({
        tag: r.tag, manufacturer: (mr && mr.manufacturer) || r.mfrText || '',
        model: r.model, series: r.series, type: typ, status: status,
        voltage: r.voltage, qty: r.qty, _tagFrame: !r.tag
      });
      if (!mr) flags.push('FLAG — ' + r.model + ' (' + r.series + ') is not in the Gil-Bar Manufacturer Map. Manufacturer/scope not auto-resolved — confirm the line before issuing.');
    });

    // reconcile against the selection package, if present
    if (selPages.length) {
      var selRows = readSelectionPackage(selPages);
      var quotedSeries = {}; quoted.forEach(function (q) { quotedSeries[q.series] = 1; });
      var byModelSel = {};
      selRows.forEach(function (s) { if (!byModelSel[s.model]) byModelSel[s.model] = s; });
      quoted.forEach(function (q) {
        var s = byModelSel[q.model];
        if (!s) { flags.push('FLAG — ' + q.model + ' is on the cost sheet but has no sheet in the selection package. Confirm the selection matches what was priced.'); return; }
        if (!q.voltage && s.voltage) q.voltage = s.voltage;
        else if (q.voltage && s.voltage && q.voltage !== s.voltage)
          flags.push('FLAG — ' + q.model + ' voltage differs: cost sheet ' + q.voltage + ' vs selection ' + s.voltage + '. Resolve before issuing.');
        if (q.qty && s.qty && q.qty !== s.qty)
          flags.push('FLAG — ' + q.model + ' quantity differs: cost sheet ' + q.qty + ' vs selection ' + s.qty + '. Cost sheet kept; confirm.');
      });
      selRows.forEach(function (s) {
        if (byModelSelUsed(quoted, s.model)) return;
        var partner = null;
        Object.keys(quotedSeries).forEach(function (qs) { var p = comparisonPartner(qs); if (p && p.partner === s.series) partner = p; });
        if (partner) excluded.push({ model: s.model, series: s.series, reason: 'Comparison side of the selection run — sister family to the cost-sheeted design-basis; recognized but NOT quoted.' });
        else excluded.push({ model: s.model, series: s.series, reason: 'In the selection package but not on the cost sheet — recognized, confirm scope.' });
      });
      var pn = readProjectName(selPages);
      if (pn && !job.header.projectName) job.header.projectName = pn;
    }

    // collapse duplicate models (sum qty, accumulate distinct tags)
    var byModel = {};
    quoted.forEach(function (q) {
      if (!byModel[q.model]) byModel[q.model] = Object.assign({}, q, { qty: q.qty || '', _tags: q.tag ? [q.tag] : [] });
      else {
        var e = byModel[q.model];
        if (q.qty && e.qty) e.qty = String((+e.qty || 0) + (+q.qty || 0));
        else if (q.qty && !e.qty) e.qty = q.qty;
        if (q.tag && e._tags.indexOf(q.tag) < 0) e._tags.push(q.tag);
        if (!e.voltage && q.voltage) e.voltage = q.voltage;
      }
    });
    job.items = Object.keys(byModel).map(function (k) {
      var it = byModel[k];
      if (it._tags && it._tags.length) it.tag = it._tags.join(', ');
      return it;
    });
    job.excluded = excluded; job.flags = flags;

    if (job.items.length) {
      var vb = verbatimBasis(job.items[0].series, job.items[0].manufacturer, job.header.proposalNo,
        job.items.map(function (x) { return x.model; }));
      if (vb.flag) job.flags.push(vb.flag);
      if (vb.row) {
        job.items.forEach(function (it) { it.includes = vb.row.includes.slice(); it.notIncluded = vb.row.notIncluded.slice(); it._basisSource = vb.row.source; });
        job.flags.push('BASIS (confirm, not verbatim) — nearest precedent: ' + vb.row.itemTitle + ' [' + vb.row.modelKey + '] from ' + vb.row.source + '. Cost-sheet option text is NOT auto-converted to Includes (verbatim rule).');
      }
    }
    job.flags = job.flags.concat(resolveFlags(job.items));
    return job;
  }
  function byModelSelUsed(quoted, model) {
    for (var i = 0; i < quoted.length; i++) if (quoted[i].model === model) return true;
    return false;
  }

  function ingestSelection(pages, dwgText) {
    var raw = readSelectionPackage(pages);
    // auto-detect the design-basis from the document (no user input)
    var db = detectDesignBasis(raw, dwgText);
    var r = resolveDesignBasis(raw, db.series);
    if (db.series && db.reason)
      r.flags.unshift('Design-basis ' + db.series + ' resolved from the document (' + db.reason + ').');
    // auto-fill the project name from the selection title
    var pn = readProjectName(pages);
    if (pn && !job.header.projectName) job.header.projectName = pn;
    // collapse duplicate quoted models, sum qty
    var byModel = {};
    r.quoted.forEach(function (q) {
      if (!byModel[q.model]) byModel[q.model] = Object.assign({}, q, { qty: q.qty || '', _tags: q.tag ? [q.tag] : [] });
      else {
        var e = byModel[q.model];
        if (q.qty && e.qty) e.qty = String((+e.qty || 0) + (+q.qty || 0));
        else if (q.qty && !e.qty) e.qty = q.qty;
        // keep every DISTINCT tag; sheets for the same model may carry different tags
        if (q.tag && e._tags.indexOf(q.tag) < 0) e._tags.push(q.tag);
        if (!e.voltage && q.voltage) e.voltage = q.voltage;
      }
    });
    job.items = Object.keys(byModel).map(function (k) {
      var it = byModel[k];
      if (it._tags && it._tags.length) it.tag = it._tags.join(', ');
      return it;
    });
    job.excluded = r.excluded; job.flags = r.flags.slice();
    if (job.items.length) {
      var vb = verbatimBasis(job.items[0].series, job.items[0].manufacturer, job.header.proposalNo,
        job.items.map(function (x) { return x.model; }));
      if (vb.flag) job.flags.push(vb.flag);
      if (vb.row) {
        job.items.forEach(function (it) { it.includes = vb.row.includes.slice(); it.notIncluded = vb.row.notIncluded.slice(); it._basisSource = vb.row.source; });
        job.flags.push('BASIS (confirm, not verbatim) — nearest precedent: ' + vb.row.itemTitle + ' [' + vb.row.modelKey + '] from ' + vb.row.source + '.');
      }
    }
    job.flags = job.flags.concat(resolveFlags(job.items));
    return job;
  }

  /* =====================================================================
     DOCX BUILD (styled Gil-Bar letterhead reconstruction)
     ===================================================================== */
  var RED = 'C0272D', REDD = '8A1A1E', REDLT = 'C16B6B', TAGBG = 'F7D2D4', ZEBRA = 'F5F7F9', GREY = '555555', SERIF = 'Georgia', CW = 9360;
  function buildDoc(wmBytes, TC) {
    var d = D;
    var NONE = { style: d.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    var noB = { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE };
    var grid = {
      top: { style: d.BorderStyle.SINGLE, size: 2, color: 'DDDDDD' }, bottom: { style: d.BorderStyle.SINGLE, size: 2, color: 'DDDDDD' },
      left: { style: d.BorderStyle.SINGLE, size: 2, color: 'DDDDDD' }, right: { style: d.BorderStyle.SINGLE, size: 2, color: 'DDDDDD' },
      insideHorizontal: { style: d.BorderStyle.SINGLE, size: 2, color: 'EEEEEE' }, insideVertical: { style: d.BorderStyle.SINGLE, size: 2, color: 'EEEEEE' }
    };
    function redRule(a, b) {
      return new d.Paragraph({ spacing: { before: a == null ? 120 : a, after: b == null ? 120 : b }, border: { bottom: { style: d.BorderStyle.SINGLE, size: 12, color: RED, space: 1 } }, children: [new d.TextRun({ text: '', font: SERIF, size: 2 })] });
    }
    function body(text, o) { o = o || {}; return new d.Paragraph({ spacing: { after: o.after == null ? 80 : o.after, line: 264 }, alignment: o.align, children: [new d.TextRun({ text: text, font: SERIF, size: o.size || 18, bold: o.bold, italics: o.italics, color: o.color || '000000' })] }); }
    function subLabel(text) { return new d.Paragraph({ spacing: { before: 140, after: 40 }, children: [new d.TextRun({ text: text, font: SERIF, size: 18, bold: true, color: RED })] }); }
    function bullet(text, color) { return new d.Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 30, line: 252 }, children: [new d.TextRun({ text: text, font: SERIF, size: 17, color: color || '000000' })] }); }
    function bandCell(label, value, dark) {
      return new d.TableCell({ width: { size: CW / 2, type: d.WidthType.DXA }, shading: { fill: dark ? REDD : RED, type: d.ShadingType.CLEAR, color: 'auto' }, margins: { top: 60, bottom: 120, left: 160, right: 160 }, borders: noB, children: [new d.Paragraph({ spacing: { after: 20 }, children: [new d.TextRun({ text: label, font: SERIF, size: 18, bold: true, color: 'FFFFFF' })] }), new d.Paragraph({ children: [new d.TextRun({ text: value || '', font: SERIF, size: 18, italics: true, color: 'FFFFFF' })] })] });
    }
    function bandTitle(text, dark) { return new d.TableCell({ width: { size: CW / 2, type: d.WidthType.DXA }, shading: { fill: dark ? REDD : RED, type: d.ShadingType.CLEAR, color: 'auto' }, margins: { top: 120, bottom: 40, left: 160, right: 160 }, borders: noB, children: [new d.Paragraph({ children: [new d.TextRun({ text: text, font: SERIF, size: 20, bold: true, color: 'FFFFFF' })] })] }); }
    function specTable(rows) {
      var cols = [1500, 2800, 2660, 1500, 900], head = ['TAG', 'MANUFACTURER', 'MODEL', 'VOLTAGE', 'QTY'];
      var hr = new d.TableRow({ tableHeader: true, children: head.map(function (h, i) { return new d.TableCell({ width: { size: cols[i], type: d.WidthType.DXA }, shading: { fill: i === 0 ? RED : REDD, type: d.ShadingType.CLEAR, color: 'auto' }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, verticalAlign: d.VerticalAlign.CENTER, children: [new d.Paragraph({ alignment: i === 0 ? d.AlignmentType.CENTER : d.AlignmentType.LEFT, children: [new d.TextRun({ text: h, font: SERIF, size: 16, bold: true, color: 'FFFFFF' })] })] }); }) });
      var dr = rows.map(function (r, ri) { return new d.TableRow({ children: r.map(function (c, i) { return new d.TableCell({ width: { size: cols[i], type: d.WidthType.DXA }, shading: i === 0 ? { fill: TAGBG, type: d.ShadingType.CLEAR, color: 'auto' } : (ri % 2 === 1 ? { fill: ZEBRA, type: d.ShadingType.CLEAR, color: 'auto' } : undefined), margins: { top: 50, bottom: 50, left: 100, right: 100 }, verticalAlign: d.VerticalAlign.CENTER, children: [new d.Paragraph({ alignment: i === 0 ? d.AlignmentType.CENTER : d.AlignmentType.LEFT, children: [new d.TextRun({ text: String(c == null ? '' : c), font: SERIF, size: 16, bold: i === 0, color: i === 0 ? REDLT : '000000' })] })] }); }) }); });
      return new d.Table({ width: { size: CW, type: d.WidthType.DXA }, columnWidths: cols, borders: grid, rows: [hr].concat(dr) });
    }
    function pricing(items) {
      var cols = [7060, 2300];
      var hdr = new d.TableRow({ tableHeader: true, children: [new d.TableCell({ width: { size: cols[0], type: d.WidthType.DXA }, shading: { fill: RED, type: d.ShadingType.CLEAR, color: 'auto' }, margins: { top: 60, bottom: 60, left: 120, right: 120 }, children: [new d.Paragraph({ children: [new d.TextRun({ text: 'ITEM / DESCRIPTION', font: SERIF, size: 16, bold: true, color: 'FFFFFF' })] })] }), new d.TableCell({ width: { size: cols[1], type: d.WidthType.DXA }, shading: { fill: RED, type: d.ShadingType.CLEAR, color: 'auto' }, margins: { top: 60, bottom: 60, left: 120, right: 120 }, children: [new d.Paragraph({ alignment: d.AlignmentType.CENTER, children: [new d.TextRun({ text: 'PRICE', font: SERIF, size: 16, bold: true, color: 'FFFFFF' })] })] })] });
      var rows = items.map(function (it) { return new d.TableRow({ children: [new d.TableCell({ width: { size: cols[0], type: d.WidthType.DXA }, margins: { top: 50, bottom: 50, left: 120, right: 120 }, children: [new d.Paragraph({ children: [new d.TextRun({ text: (it.label ? it.label.replace(/:$/, '') : 'Item') + ' \u2014 ' + it.title, font: SERIF, size: 16 })] })] }), new d.TableCell({ width: { size: cols[1], type: d.WidthType.DXA }, margins: { top: 50, bottom: 50, left: 120, right: 120 }, children: [new d.Paragraph({ alignment: d.AlignmentType.CENTER, children: [new d.TextRun({ text: '$____________', font: SERIF, size: 16 })] })] })] }); });
      var tot = new d.TableRow({ children: [new d.TableCell({ width: { size: cols[0], type: d.WidthType.DXA }, shading: { fill: RED, type: d.ShadingType.CLEAR, color: 'auto' }, margins: { top: 50, bottom: 50, left: 120, right: 120 }, children: [new d.Paragraph({ children: [new d.TextRun({ text: 'TOTAL PRICE (excluding taxes)', font: SERIF, size: 16, bold: true, color: 'FFFFFF' })] })] }), new d.TableCell({ width: { size: cols[1], type: d.WidthType.DXA }, shading: { fill: RED, type: d.ShadingType.CLEAR, color: 'auto' }, margins: { top: 50, bottom: 50, left: 120, right: 120 }, children: [new d.Paragraph({ alignment: d.AlignmentType.CENTER, children: [new d.TextRun({ text: '$____________', font: SERIF, size: 16, bold: true, color: 'FFFFFF' })] })] })] });
      return new d.Table({ width: { size: CW, type: d.WidthType.DXA }, columnWidths: cols, borders: grid, rows: [hdr].concat(rows).concat([tot]) });
    }
    function sig(label) { return new d.Paragraph({ spacing: { before: 300, after: 0 }, border: { top: { style: d.BorderStyle.SINGLE, size: 4, color: '999999', space: 2 } }, children: [new d.TextRun({ text: label, font: SERIF, size: 15, color: GREY })] }); }
    function tcP(c) { var runs = [new d.TextRun({ text: c.n + '. ', font: SERIF, size: 15, bold: true })]; if (c.lead) runs.push(new d.TextRun({ text: c.lead + ' ', font: SERIF, size: 15, bold: true })); runs.push(new d.TextRun({ text: c.body, font: SERIF, size: 15 })); return new d.Paragraph({ spacing: { after: 100, line: 240 }, alignment: d.AlignmentType.JUSTIFIED, children: runs }); }

    var H = job.header;
    // assign item labels A, B, ... and titles from type
    var LET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    var itemsForDoc = [];
    // one Item per equipment type
    var byType = {};
    job.items.forEach(function (it) { var ttl = it.type || 'Equipment'; (byType[ttl] = byType[ttl] || []).push(it); });
    Object.keys(byType).forEach(function (ttl, idx) {
      itemsForDoc.push({ label: 'Item ' + LET[idx] + ':', title: ttl, rows: byType[ttl].map(function (it) { return [it.tag || it.model, it.manufacturer, it.model, it.voltage, it.qty]; }), includes: byType[ttl][0].includes || [], notIncluded: byType[ttl][0].notIncluded || [] });
    });

    var docHeader = new d.Header({ children: [new d.Table({ width: { size: CW, type: d.WidthType.DXA }, columnWidths: [3600, 5760], borders: noB, rows: [new d.TableRow({ children: [new d.TableCell({ width: { size: 3600, type: d.WidthType.DXA }, borders: noB, verticalAlign: d.VerticalAlign.CENTER, children: [new d.Paragraph({ children: [new d.ImageRun({ type: 'png', data: wmBytes, transformation: { width: 133, height: 63 } })] })] }), new d.TableCell({ width: { size: 5760, type: d.WidthType.DXA }, borders: noB, verticalAlign: d.VerticalAlign.CENTER, children: [new d.Paragraph({ alignment: d.AlignmentType.RIGHT, spacing: { after: 10 }, children: [new d.TextRun({ text: 'HVAC EQUIPMENT PROPOSAL', font: SERIF, size: 20, bold: true, color: RED })] }), new d.Paragraph({ alignment: d.AlignmentType.RIGHT, children: [new d.TextRun({ text: '498 7th Avenue, 14th Floor, New York, NY 10018', font: SERIF, size: 13, color: GREY })] }), new d.Paragraph({ alignment: d.AlignmentType.RIGHT, children: [new d.TextRun({ text: 'Proposal No: ', font: SERIF, size: 13, color: GREY }), new d.TextRun({ text: H.proposalNo || '[CONFIRM — Quote No.]', font: SERIF, size: 13, color: RED, bold: true }), new d.TextRun({ text: '  |  Date: ', font: SERIF, size: 13, color: GREY }), new d.TextRun({ text: H.proposalDate || '[CONFIRM — Date]', font: SERIF, size: 13, color: RED, bold: true })] })] })] })] }), redRule(80, 40)] });
    var docFooter = new d.Footer({ children: [new d.Paragraph({ border: { top: { style: d.BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } }, tabStops: [{ type: d.TabStopType.RIGHT, position: d.TabStopPosition.MAX }], children: [new d.TextRun({ text: 'Gil-Bar  |  An Ambient Company  |  498 7th Avenue, 14th Floor, New York, NY 10018  |  Confidential', font: SERIF, size: 13, color: GREY }), new d.TextRun({ text: '\tPage ', font: SERIF, size: 13, color: GREY }), new d.TextRun({ children: [d.PageNumber.CURRENT], font: SERIF, size: 13, color: GREY })] })] });

    var children = [];
    children.push(new d.Table({ width: { size: CW, type: d.WidthType.DXA }, columnWidths: [CW / 2, CW / 2], borders: noB, rows: [new d.TableRow({ children: [bandTitle('PROJECT INFORMATION', false), bandTitle('PROPOSAL DETAILS', true)] }), new d.TableRow({ children: [bandCell('Project Name:', H.projectName || '[CONFIRM — Project Name]'), bandCell('Submitted To:', H.submittedTo || '[CONFIRM — Contractor / Client]', true)] }), new d.TableRow({ children: [bandCell('Project Address:', H.projectAddress || '[per project record]'), bandCell('Proposal Date:', H.proposalDate || '[CONFIRM — Date]', true)] }), new d.TableRow({ children: [bandCell('MEP Engineer:', H.mepEngineer || '[CONFIRM — Engineer]'), bandCell('Valid Until:', H.validUntil || '[CONFIRM — Date + 30]', true)] }), new d.TableRow({ children: [bandCell('Architect:', H.architect || '[per project record]'), bandCell('Mechanical Drawings Date:', H.drawingsDate || '[CONFIRM — Drawings Dated]', true)] })] }));
    children.push(redRule(160, 120));
    children.push(body('Gil-Bar is pleased to submit our proposal for ' + (H.projectName || '[PROJECT NAME]') + ', per the mechanical drawings and specifications. All equipment is furnished per the scope below.', { after: 120 }));
    children.push(redRule(120, 140));
    itemsForDoc.forEach(function (it, idx) {
      children.push(new d.Paragraph({ spacing: { after: 100 }, children: [new d.TextRun({ text: it.label + '   ', font: SERIF, size: 22, bold: true, color: RED }), new d.TextRun({ text: it.title, font: SERIF, size: 22, color: '000000' })] }));
      children.push(specTable(it.rows));
      if (it.includes && it.includes.length) { children.push(subLabel('Includes:')); it.includes.forEach(function (x) { children.push(bullet(x)); }); }
      if (it.notIncluded && it.notIncluded.length) { children.push(subLabel('Not Included:')); it.notIncluded.forEach(function (x) { children.push(bullet(x)); }); }
      if (idx < itemsForDoc.length - 1) children.push(redRule(160, 140));
    });
    // flags block
    if (job.flags && job.flags.length) {
      children.push(redRule(160, 120));
      children.push(new d.Paragraph({ spacing: { after: 20 }, children: [new d.TextRun({ text: 'Flags / To Confirm:', font: SERIF, size: 18, bold: true, color: RED })] }));
      job.flags.forEach(function (x) { children.push(bullet(x, RED)); });
    }
    children.push(redRule(160, 120));
    children.push(new d.Paragraph({ spacing: { after: 80 }, children: [new d.TextRun({ text: 'PRICING SUMMARY', font: SERIF, size: 20, bold: true, color: RED })] }));
    children.push(pricing(itemsForDoc));
    children.push(new d.Paragraph({ spacing: { before: 140, after: 60 }, alignment: d.AlignmentType.JUSTIFIED, children: [new d.TextRun({ text: 'NOTE: This quotation does NOT include any potential tariffs that may be imposed. Gil-Bar reserves the right to amend this proposal due to notification from vendors of pricing escalations resulting from any government-imposed tariffs or duties.', font: SERIF, size: 15, italics: true, color: GREY })] }));
    children.push(new d.Paragraph({ spacing: { after: 120 }, children: [new d.TextRun({ text: 'Units ship FOB factory, freight allowed to first stop.', font: SERIF, size: 15, italics: true, color: GREY })] }));
    children.push(body('Thank you for considering Gil-Bar for this project, we look forward to working with you.', { after: 160 }));
    children.push(redRule(120, 140));
    children.push(new d.Paragraph({ spacing: { after: 60 }, children: [new d.TextRun({ text: 'ACCEPTANCE', font: SERIF, size: 20, bold: true, color: RED })] }));
    children.push(body('The undersigned agrees to purchase the above equipment scope as per this proposal.', { after: 120 }));
    children.push(new d.Table({ width: { size: CW, type: d.WidthType.DXA }, columnWidths: [CW / 2, CW / 2], borders: noB, rows: [new d.TableRow({ children: [new d.TableCell({ width: { size: CW / 2, type: d.WidthType.DXA }, borders: noB, margins: { right: 240 }, children: [sig('Authorized Signature \u2014 ' + (job.header.entityName || 'Gil-Bar Industries')), sig('Printed Name'), sig('Title'), sig('Date')] }), new d.TableCell({ width: { size: CW / 2, type: d.WidthType.DXA }, borders: noB, margins: { left: 240 }, children: [sig('Accepted By \u2014 [Client / Contractor]'), sig('Printed Name'), sig('Title'), sig('Company'), sig('Date')] })] })] }));
    children.push(new d.Paragraph({ children: [new d.PageBreak()] }));
    children.push(new d.Paragraph({ spacing: { after: 120 }, border: { bottom: { style: d.BorderStyle.SINGLE, size: 8, color: RED, space: 4 } }, children: [new d.TextRun({ text: 'TERMS AND CONDITIONS OF SALE', font: SERIF, size: 20, bold: true, color: RED })] }));
    TC.forEach(function (c) { children.push(tcP(c)); });

    return new d.Document({ styles: { default: { document: { run: { font: SERIF, size: 18 } } } }, numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: 'bullet', text: '\u2022', alignment: d.AlignmentType.LEFT, style: { paragraph: { indent: { left: 460, hanging: 240 } } } }] }] }, sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1980, right: 1440, bottom: 1080, left: 1440 } } }, headers: { default: docHeader }, footers: { default: docFooter }, children: children }] });
  }

  /* ------------------------------------------------------------- exports */
  window.GilBar = {
    setDocx: function (dx) { D = dx; },
    job: job,
    seriesOf: seriesOf, mapRowForSeries: mapRowForSeries, comparisonPartner: comparisonPartner,
    recognizePositioned: recognizePositioned, scopeGate: scopeGate, resolveFlags: resolveFlags,
    readSelectionPackage: readSelectionPackage, resolveDesignBasis: resolveDesignBasis, verbatimBasis: verbatimBasis,
    readProjectName: readProjectName, detectDesignBasis: detectDesignBasis,
    readCostSheetAOA: readCostSheetAOA, ingestJob: ingestJob,
    ingestDrawingTokens: ingestDrawingTokens, ingestSelection: ingestSelection,
    buildDoc: buildDoc
  };
})();
