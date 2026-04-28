'use strict';

// ─── Get currently filtered records based on active filters ──
function getFilteredRecords() {
  const allDb = getRecords();
  const yr  = (() => { const s=document.getElementById('rec-year-filter'); return s?parseInt(s.value)||0:0; })();
  const db  = filterByYear(allDb, yr);
  const q   = (document.getElementById('rec-search')?.value||'').toLowerCase();
  const sf  = document.getElementById('rec-filter-status')?.value||'';
  const cf  = document.getElementById('rec-filter-charge')?.value||'';
  const dvf = document.getElementById('rec-filter-dv')?.value||'';

  return db.filter(r => {
    const txt = !q ||
      (r.payee||'').toLowerCase().includes(q)       ||
      (r.obrNo||'').toLowerCase().includes(q)       ||
      (r.dvNo||'').toLowerCase().includes(q)        ||
      (r.dvPayee||'').toLowerCase().includes(q)     ||
      (r.particulars||'').toLowerCase().includes(q) ||
      (r.chargeTo||'').toLowerCase().includes(q);
    const matchStatus = !sf || r.action === sf;
    const matchCharge = !cf || r.chargeTo === cf;
    let matchDV = true;
    if (dvf === 'no-dv') { matchDV = !r.dvNo || r.dvNo.trim() === ''; }
    else if (dvf) { const m=(r.dvNo||'').trim().match(/^([A-Za-z]+)/); matchDV=(m?m[1].toUpperCase():'')=== dvf; }
    return txt && matchStatus && matchCharge && matchDV;
  });
}

// ─── Build export label from active filters ───────────────
function getExportLabel() {
  const yr  = (() => { const s=document.getElementById('rec-year-filter'); return s&&s.value!=='0'?s.value:''; })();
  const sf  = document.getElementById('rec-filter-status')?.value||'';
  const cf  = document.getElementById('rec-filter-charge')?.value||'';
  const dvf = document.getElementById('rec-filter-dv')?.value||'';
  const parts = [];
  if (yr)  parts.push(yr);
  if (sf)  parts.push(sf);
  if (cf)  parts.push(cf);
  if (dvf && dvf !== 'no-dv') parts.push(dvf);
  if (dvf === 'no-dv') parts.push('NoDV');
  return parts.length ? '_' + parts.join('_') : '_AllRecords';
}

// ─── CSV Export ───────────────────────────────────────────
function exportCSV() {
  const db = getFilteredRecords();
  if (!db.length) { showToast('⚠️ Walang records na naka-filter', 'error'); return; }

  const label = getExportLabel();
  const h = [
    '#', 'Year', 'Date In', 'OBR/BUR No.', 'Payee',
    'Amount (OBR)', 'Charge To', 'Particulars (OBR)',
    'DV No.', 'DV Payee', 'DV Amount', 'DV Charge To', 'Particulars (DV)',
    'OUP Received By', 'OUP Date', 'Status', 'Location'
  ];
  const rows = db.map((r, i) => [
    i+1, r.year||'', r.dateIn||'', r.obrNo||'', r.payee||'',
    r.amount||0, r.chargeTo||'',
    (r.particulars||'').replace(/,/g, ';').replace(/\n/g, ' '),
    r.dvNo||'', r.dvPayee||'', r.dvAmount||0, r.dvCharge||'',
    (r.dvParticulars||'').replace(/,/g, ';').replace(/\n/g, ' '),
    r.oupReceived||'', r.oupDate||'', r.action||'', locationLabel(r.location),
  ]);
  const csv = [h, ...rows].map(row => row.map(v => `"${v}"`).join(',')).join('\n');
  dlFile(`OBR_Voucher${label}.csv`, 'text/csv;charset=utf-8;', '\uFEFF' + csv);
  showToast(`📄 CSV exported! (${db.length} records)`, 'success');
}

// ─── Excel Export ─────────────────────────────────────────
function exportExcel() {
  const db = getFilteredRecords();
  if (!db.length) { showToast('⚠️ Walang records na naka-filter', 'error'); return; }

  const label = getExportLabel();
  const e = s => (s||'').toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

  // Active filter summary row
  const yr  = document.getElementById('rec-year-filter')?.value||'0';
  const sf  = document.getElementById('rec-filter-status')?.value||'';
  const cf  = document.getElementById('rec-filter-charge')?.value||'';
  const dvf = document.getElementById('rec-filter-dv')?.value||'';
  const filterSummary = [
    yr !== '0' ? `Year: ${yr}` : 'All Years',
    sf ? `Status: ${sf}` : 'All Status',
    cf ? `Charge: ${cf}` : 'All Charges',
    dvf && dvf !== 'no-dv' ? `DV Type: ${dvf}` : dvf === 'no-dv' ? 'Walang DV' : 'All DV Types',
  ].join('  |  ');

  let html = `<table>
    <thead>
      <tr>
        <td colspan="17" style="background:#0d1117;color:#fff;font-size:13px;font-weight:bold;padding:8px 12px">
          OBR &amp; Voucher Tracker — PRMSU Budget Office &nbsp;&nbsp;|&nbsp;&nbsp; ${e(filterSummary)} &nbsp;&nbsp;|&nbsp;&nbsp; ${db.length} records
        </td>
      </tr>
      <tr style="background:#1a2332;color:#fff;font-weight:bold">
        <th>#</th><th>Year</th><th>Date In</th><th>OBR/BUR No.</th>
        <th>Payee</th><th>Amount (OBR)</th><th>Charge To</th><th>Particulars (OBR)</th>
        <th style="background:#1a3a5c">DV No.</th>
        <th style="background:#1a3a5c">DV Payee</th>
        <th style="background:#1a3a5c">DV Amount</th>
        <th style="background:#1a3a5c">DV Charge To</th>
        <th style="background:#1a3a5c">Particulars (DV)</th>
        <th>OUP Received By</th><th>OUP Date</th>
        <th>Status</th><th>Location</th>
      </tr>
    </thead>
    <tbody>`;

  db.forEach((r, i) => {
    const rowBg = i % 2 === 0 ? '' : 'background:#f8fafc';
    html += `<tr style="${rowBg}">
      <td>${i+1}</td>
      <td>${r.year||''}</td>
      <td>${e(r.dateIn)}</td>
      <td><b>${e(r.obrNo)}</b></td>
      <td>${e(r.payee)}</td>
      <td style="text-align:right">${r.amount||0}</td>
      <td>${e(r.chargeTo)}</td>
      <td>${e(r.particulars)}</td>
      <td style="background:#e8f0fe;font-weight:bold;color:#1a6cf0">${e(r.dvNo)}</td>
      <td style="background:#e8f0fe">${e(r.dvPayee)}</td>
      <td style="background:#e8f0fe;text-align:right">${r.dvAmount||''}</td>
      <td style="background:#e8f0fe">${e(r.dvCharge)}</td>
      <td style="background:#e8f0fe">${e(r.dvParticulars)}</td>
      <td>${e(r.oupReceived)}</td>
      <td>${e(r.oupDate)}</td>
      <td>${e(r.action)}</td>
      <td>${e(locationLabel(r.location))}</td>
    </tr>`;
  });

  html += '</tbody></table>';
  dlFile(`OBR_Voucher${label}.xls`, 'application/vnd.ms-excel', html);
  showToast(`📊 Excel exported! (${db.length} records)`, 'success');
}

// ─── Download Helper ──────────────────────────────────────
function dlFile(name, type, content) {
  const blob = new Blob([content], { type });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
