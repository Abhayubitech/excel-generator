import React, { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import './App.css';

// ─── Constants ────────────────────────────────────────────────────────────────
const MONTHS = [
  { name: 'January', val: 1 },  { name: 'February', val: 2 },
  { name: 'March', val: 3 },    { name: 'April', val: 4 },
  { name: 'May', val: 5 },      { name: 'June', val: 6 },
  { name: 'July', val: 7 },     { name: 'August', val: 8 },
  { name: 'September', val: 9 },{ name: 'October', val: 10 },
  { name: 'November', val: 11 },{ name: 'December', val: 12 },
];
const YEARS = Array.from({ length: 11 }, (_, i) => 2024 + i);
const DAYS_OF_WEEK = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const QUICK_TOWNS = ['Gwalior','Bhind','Morena','Shivpuri','Datia','Guna'];
const LS_KEY      = 'psm_gen_v2';        // settings key
const rowsKey     = (y, m) => `psm_rows_${y}_${m}`;

// ─── Pure helpers (no hooks) ──────────────────────────────────────────────────
function recalc(row, rate) {
  const r = rate ?? row.rate ?? 4;
  const bikeAmt = (row.km || 0) * r;
  const total   = (row.hqTa||0)+(row.hqDa||0)+(row.upTa||0)+(row.upDa||0)+bikeAmt+(row.hotel||0)+(row.bus||0);
  return { ...row, rate: r, bikeAmt, total };
}

function buildRow(d, month, year, hq, rate) {
  const dt       = new Date(year, month - 1, d);
  const dayIndex = dt.getDay();
  const dayName  = DAYS_OF_WEEK[dayIndex];
  const isSun    = dayIndex === 0;
  const dateStr  = `${String(d).padStart(2,'0')}/${String(month).padStart(2,'0')}/${year}`;
  if (isSun) {
    return { dayNum:d, dateStr, dayName, town:'', type:'Sunday',
             hqTa:0, hqDa:0, upTa:0, upDa:0, km:0, rate, bikeAmt:0, hotel:0, bus:0, total:0 };
  }
  return { dayNum:d, dateStr, dayName, town:hq, type:'HQ',
           hqTa:175, hqDa:175, upTa:0, upDa:0, km:0, rate, bikeAmt:0, hotel:0, bus:0, total:350 };
}

function applyType(row, type, hq) {
  if (type === 'HQ')       return { ...row, type, hqTa:175, hqDa:175, upTa:0, upDa:0, town: hq };
  if (type === 'Upcountry') return { ...row, type, hqTa:0,   hqDa:0,   upTa:0, upDa:250 };
  return { ...row, type, hqTa:0, hqDa:0, upTa:0, upDa:0, town:'', km:0, hotel:0, bus:0 };
}

function applyTownAndAutoType(row, town, hq) {
  let nextRow = { ...row, town: town };
  if (row.type === 'Sunday' || row.type === 'Leave' || row.type === 'Holiday') {
    return nextRow;
  }
  const cleanedTown = (town || '').trim();
  if (cleanedTown.toLowerCase() === hq.toLowerCase()) {
    nextRow = applyType(nextRow, 'HQ', hq);
  } else if (cleanedTown !== '') {
    nextRow = applyType(nextRow, 'Upcountry', hq);
  }
  return nextRow;
}

function loadSettings() {
  try { return JSON.parse(localStorage.getItem(LS_KEY)) || {}; }
  catch { return {}; }
}

function loadOrBuildRows(month, year, hq, rate) {
  const daysInMonth = new Date(year, month, 0).getDate();
  try {
    const raw   = localStorage.getItem(rowsKey(year, month));
    const saved = raw ? JSON.parse(raw) : null;
    if (saved && saved.length === daysInMonth) return saved;
  } catch { /* fall through */ }
  return Array.from({ length: daysInMonth }, (_, i) => buildRow(i + 1, month, year, hq, rate));
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  // ── lazy-init all state from localStorage ───────────────────────────────────
  const [psmName,     setPsmName]     = useState(() => loadSettings().psmName     || 'Akash Vishoriya');
  const [empId,       setEmpId]       = useState(() => loadSettings().empId       || 'I3785');
  const [joiningDate, setJoiningDate] = useState(() => loadSettings().joiningDate || '25-04-2024');
  const [hq,          setHq]          = useState(() => loadSettings().hq          || 'Gwalior');
  const [mobileNo,    setMobileNo]    = useState(() => loadSettings().mobileNo    || '9074305446');
  const [aseName,     setAseName]     = useState(() => loadSettings().aseName     || '');
  const [ratePerKm,   setRatePerKm]   = useState(() => loadSettings().ratePerKm   || 4);
  const [month,       setMonth]       = useState(() => loadSettings().month       || new Date().getMonth() + 1);
  const [year,        setYear]        = useState(() => loadSettings().year        || new Date().getFullYear());

  const [rows,        setRows]        = useState(() => {
    const s = loadSettings();
    return loadOrBuildRows(
      s.month    || new Date().getMonth() + 1,
      s.year     || new Date().getFullYear(),
      s.hq       || 'Gwalior',
      s.ratePerKm|| 4
    );
  });

  const [bulkTown,  setBulkTown]  = useState('');
  const [toast,     setToast]     = useState(null);
  const [panelOpen, setPanelOpen] = useState(() => !window.matchMedia('(max-width: 767px)').matches);
  const [isMobile,  setIsMobile]  = useState(() =>  window.matchMedia('(max-width: 767px)').matches);

  // range updater
  const [rangeFrom,  setRangeFrom]  = useState(1);
  const [rangeTo,    setRangeTo]    = useState(1);
  const [rangeType,  setRangeType]  = useState('HQ');
  const [rangeTown,  setRangeTown]  = useState('');
  const [rangeKm,    setRangeKm]    = useState('');
  const [rangeHotel, setRangeHotel] = useState('');
  const [rangeBus,   setRangeBus]   = useState('');

  // skip-first-render ref for month/year change effect
  const isFirstRender = useRef(true);

  useEffect(() => {
    const mql = window.matchMedia('(max-width: 767px)');
    const handler = (e) => setIsMobile(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);

  // ── rebuild rows when month/year changes (skip initial mount) ────────────────
  useEffect(() => {
    if (isFirstRender.current) { isFirstRender.current = false; return; }
    setRows(loadOrBuildRows(month, year, hq, ratePerKm));
    setRangeFrom(1);
    setRangeTo(1);
  }, [month, year]); // eslint-disable-line react-hooks/exhaustive-deps

  // ── auto-save rows ───────────────────────────────────────────────────────────
  useEffect(() => {
    if (!rows.length) return;
    try { localStorage.setItem(rowsKey(year, month), JSON.stringify(rows)); } catch(e) {}
  }, [rows, year, month]);

  // ── auto-save settings ───────────────────────────────────────────────────────
  useEffect(() => {
    try {
      localStorage.setItem(LS_KEY, JSON.stringify(
        { psmName, empId, joiningDate, hq, mobileNo, aseName, ratePerKm, month, year }
      ));
    } catch(e) {}
  }, [psmName, empId, joiningDate, hq, mobileNo, aseName, ratePerKm, month, year]);

  // ── toast ────────────────────────────────────────────────────────────────────
  const showToast = useCallback((msg, type = 'success') => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  }, []);

  // ── cell change ──────────────────────────────────────────────────────────────
  const handleCellChange = useCallback((index, field, value) => {
    setRows(prev => {
      const next = [...prev];
      let row = { ...next[index] };
      if (field === 'type')   row = applyType(row, value, hq);
      else if (field === 'town') row = applyTownAndAutoType(row, value, hq);
      else { const n = parseFloat(value); row[field] = isNaN(n) ? 0 : n; }
      next[index] = recalc(row, ratePerKm);
      return next;
    });
  }, [hq, ratePerKm]);

  // ── copy from previous day ────────────────────────────────────────────────────
  const copyFromPrev = useCallback((index) => {
    if (index === 0) return;
    setRows(prev => {
      const src = prev[index - 1];
      if (!src || src.type === 'Sunday') return prev;
      const next = [...prev];
      let row = applyType({ ...next[index] }, src.type, hq);
      row.town  = src.town;
      row.km    = src.km;
      row.hotel = src.hotel;
      row.bus   = src.bus;
      next[index] = recalc(row, ratePerKm);
      return next;
    });
    showToast('Copied from previous day ✓');
  }, [hq, ratePerKm, showToast]);

  // ── quick town ────────────────────────────────────────────────────────────────
  const setQuickTown = useCallback((index, town) => {
    setRows(prev => {
      const next = [...prev];
      let row = applyTownAndAutoType(next[index], town, hq);
      next[index] = recalc(row, ratePerKm);
      return next;
    });
  }, [hq, ratePerKm]);

  // ── bulk town ─────────────────────────────────────────────────────────────────
  const applyBulkTown = useCallback(() => {
    const t = bulkTown.trim();
    if (!t) return;
    setRows(prev => prev.map(r => {
      if (r.type === 'Sunday' || r.type === 'Leave' || r.type === 'Holiday') return r;
      let row = applyTownAndAutoType(r, t, hq);
      return recalc(row, ratePerKm);
    }));
    showToast(`Town set to "${t}" for all working days ✓`);
    setBulkTown('');
  }, [bulkTown, hq, ratePerKm, showToast]);

  // ── range updater ─────────────────────────────────────────────────────────────
  const applyRange = useCallback(() => {
    const from = Math.max(1, parseInt(rangeFrom) || 1);
    const to   = Math.min(rows.length, parseInt(rangeTo) || 1);
    if (from > to) { showToast('Invalid range!', 'error'); return; }
    setRows(prev => prev.map(r => {
      if (r.dayNum < from || r.dayNum > to || r.type === 'Sunday') return r;
      let u = applyType(r, rangeType, hq);
      const t = rangeTown.trim();
      if (t) {
        u = applyTownAndAutoType(u, t, hq);
      }
      if (rangeKm !== '')    u.km    = parseFloat(rangeKm)    || 0;
      if (rangeHotel !== '') u.hotel = parseFloat(rangeHotel) || 0;
      if (rangeBus !== '')   u.bus   = parseFloat(rangeBus)   || 0;
      return recalc(u, ratePerKm);
    }));
    showToast(`Updated days ${from}–${to} ✓`);
  }, [rangeFrom, rangeTo, rangeType, rangeTown, rangeKm, rangeHotel, rangeBus, hq, ratePerKm, rows.length, showToast]);

  // ── totals ────────────────────────────────────────────────────────────────────
  const totals = useMemo(() => rows.reduce(
    (acc, r) => ({
      hqTa:    acc.hqTa    + (r.hqTa    || 0),
      hqDa:    acc.hqDa    + (r.hqDa    || 0),
      upTa:    acc.upTa    + (r.upTa    || 0),
      upDa:    acc.upDa    + (r.upDa    || 0),
      km:      acc.km      + (r.km      || 0),
      bikeAmt: acc.bikeAmt + (r.bikeAmt || 0),
      hotel:   acc.hotel   + (r.hotel   || 0),
      bus:     acc.bus     + (r.bus     || 0),
      total:   acc.total   + (r.total   || 0),
    }),
    { hqTa:0, hqDa:0, upTa:0, upDa:0, km:0, bikeAmt:0, hotel:0, bus:0, total:0 }
  ), [rows]);

  const summaryTA  = totals.hqTa + totals.bikeAmt + totals.upTa;
  const summaryDA  = totals.hqDa + totals.upDa;
  const summaryOE  = totals.hotel + totals.bus;
  const daysInMonth = new Date(year, month, 0).getDate();

  // ── Excel export ──────────────────────────────────────────────────────────────
  const exportToExcel = async () => {
    try {
      showToast('Generating Excel file…');
      const workbook  = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Psm Exp Format', { views: [{ showGridLines: true }] });

      worksheet.columns = [
        {key:'A',width:5.66}, {key:'B',width:14.55}, {key:'C',width:13},
        {key:'D',width:13},   {key:'E',width:16.66}, {key:'F',width:10.33},
        {key:'G',width:13},   {key:'H',width:13},    {key:'I',width:13},
        {key:'J',width:13},   {key:'K',width:13},    {key:'L',width:13},
        {key:'M',width:11.66},{key:'N',width:17.33}, {key:'O',width:16.33},
      ];
      [15.6,15.6,16.2,15.6,18,22.5].forEach((h,i) => { worksheet.getRow(i+1).height = h; });

      const thin    = { top:{style:'thin',color:{argb:'FF808080'}}, left:{style:'thin',color:{argb:'FF808080'}},
                        bottom:{style:'thin',color:{argb:'FF808080'}}, right:{style:'thin',color:{argb:'FF808080'}} };
      const fillG   = { type:'pattern', pattern:'solid', fgColor:{argb:'FF404040'} };
      const fillY   = { type:'pattern', pattern:'solid', fgColor:{argb:'FFFFFF00'} };
      const fillGrn = { type:'pattern', pattern:'solid', fgColor:{argb:'FFC6EFCE'} };
      const fWB12   = { name:'Calibri', size:12, bold:true, color:{argb:'FFFFFFFF'} };
      const fWB11   = { name:'Calibri', size:11, bold:true, color:{argb:'FFFFFFFF'} };
      const fBB11   = { name:'Calibri', size:11, bold:true, color:{argb:'FF000000'} };
      const fBN11   = { name:'Calibri', size:11, bold:false,color:{argb:'FF000000'} };
      const fGB11   = { name:'Calibri', size:11, bold:true, color:{argb:'FF276221'} };
      const aC = { horizontal:'center', vertical:'middle' };
      const aL = { horizontal:'left',   vertical:'middle' };
      const aR = { horizontal:'right',  vertical:'middle' };

      const sc = (cell, font, fill, border, align) => {
        if (font)   cell.font      = font;
        if (fill)   cell.fill      = fill;
        if (border) cell.border    = border;
        if (align)  cell.alignment = align;
      };

      // ── Info rows 1–3
      const mm = month.toString().padStart(2,'0');
      const ld = new Date(year, month, 0).getDate();
      const fmtLast  = `${ld.toString().padStart(2,'0')}-${mm}-${year}`;
      const fmtFirst = `01/${mm}/${year}`;
      const fmtLastS = `${ld.toString().padStart(2,'0')}/${mm}/${year}`;

      for (let r = 1; r <= 3; r++) {
        for (let c = 1; c <= 15; c++) {
          const cell = worksheet.getCell(r, c);
          sc(cell, fWB12, fillG, null, aL);
          cell.border = {
            top:    r===1 ? {style:'medium',color:{argb:'FF000000'}} : {style:'thin',color:{argb:'FF808080'}},
            bottom: r===3 ? {style:'medium',color:{argb:'FF000000'}} : {style:'thin',color:{argb:'FF808080'}},
            left:   c===1 ? {style:'medium',color:{argb:'FF000000'}} : {style:'thin',color:{argb:'FF808080'}},
            right:  c===15? {style:'medium',color:{argb:'FF000000'}} : {style:'thin',color:{argb:'FF808080'}},
          };
        }
      }
      worksheet.getCell('A1').value = `Name of PSM : ${psmName}`;
      worksheet.getCell('E1').value = `Date of Joining : ${joiningDate}`;
      worksheet.getCell('M1').value = `Date : ${fmtLast}`;
      worksheet.getCell('A2').value = `Emp ID : ${empId}`;
      worksheet.getCell('E2').value = `HQ : ${hq}`;
      worksheet.getCell('A3').value = `Mobile No : ${mobileNo}`;
      worksheet.getCell('E3').value = `ASE Name : ${aseName}`;
      worksheet.getCell('M3').value = `Expense from / to : ${fmtFirst} to ${fmtLastS}`;
      ['A1:D1','E1:I1','M1:O1','A2:D2','E2:I2','M2:O2','A3:D3','E3:I3','M3:O3']
        .forEach(r => worksheet.mergeCells(r));

      // ── Row 5 section headers
      const fillOrg = { type:'pattern', pattern:'solid', fgColor:{argb:'FFFFC000'} };
      for (let c = 1; c <= 15; c++) {
        const cell = worksheet.getCell(5, c);
        let fill = null;
        if (c >= 6 && c <= 7)   fill = fillY;
        if (c >= 8 && c <= 14)  fill = fillOrg;
        sc(cell, fBB11, fill, thin, aC);
      }
      worksheet.getCell('F5').value = 'HQ : LOCAL WORKING';
      worksheet.getCell('H5').value = 'UPCOUNTRY WORKING';
      worksheet.mergeCells('F5:G5');
      worksheet.mergeCells('H5:N5');

      // ── Row 6 column headers
      ['S. No','Date','Day','Town','HQ / Upcountry','TA','DA','TA','DA',
       'Total KM','Rate / Km','Bike Amt','Hotel','Bus / Train Amount','Total']
        .forEach((h, i) => { const c = worksheet.getCell(6, i+1); sc(c, fWB11, fillG, thin, aC); c.value = h; });

      // ── Data rows
      const DATA_START = 7;
      rows.forEach((r, idx) => {
        const er   = DATA_START + idx;
        const isSun = r.type === 'Sunday';
        worksheet.getRow(er).height = 18;

        worksheet.getCell(`A${er}`).value = r.dayNum;
        const dc = worksheet.getCell(`B${er}`);
        const [dv,mv,yv] = r.dateStr.split('/');
        dc.value  = new Date(Date.UTC(+yv, +mv-1, +dv)); dc.numFmt = 'dd/mm/yyyy';
        worksheet.getCell(`C${er}`).value = r.dayName;
        worksheet.getCell(`D${er}`).value = isSun ? null : r.town;
        worksheet.getCell(`E${er}`).value = isSun ? null : r.type;
        worksheet.getCell(`F${er}`).value = isSun ? null : r.hqTa;
        worksheet.getCell(`G${er}`).value = isSun ? null : r.hqDa;
        worksheet.getCell(`H${er}`).value = isSun ? null : r.upTa;
        worksheet.getCell(`I${er}`).value = isSun ? null : r.upDa;
        worksheet.getCell(`J${er}`).value = isSun ? null : r.km;
        worksheet.getCell(`K${er}`).value = r.rate;
        worksheet.getCell(`L${er}`).value = { formula: `J${er}*K${er}` };
        worksheet.getCell(`M${er}`).value = isSun ? null : r.hotel;
        worksheet.getCell(`N${er}`).value = isSun ? null : r.bus;
        worksheet.getCell(`O${er}`).value = { formula: `SUM(F${er}:I${er},L${er},M${er},N${er})` };

        for (let col = 1; col <= 15; col++) {
          const cell = worksheet.getCell(er, col);
          if (isSun && col >= 2 && col <= 7) sc(cell, fBB11, fillY, thin, aC);
          else sc(cell, fBN11, null, thin, aC);
        }
      });

      // ── Totals row
      const lastDataRow = DATA_START + rows.length - 1;
      const TR = lastDataRow + 4;
      worksheet.getRow(TR).height = 20;
      const dblBorder = { top:{style:'thin',color:{argb:'FF808080'}}, left:{style:'thin',color:{argb:'FF808080'}},
                          bottom:{style:'double',color:{argb:'FF276221'}}, right:{style:'thin',color:{argb:'FF808080'}} };
      for (let c = 1; c <= 15; c++) sc(worksheet.getCell(TR, c), fGB11, fillGrn, dblBorder, aC);
      worksheet.getCell(`E${TR}`).value = 'TOTAL'; worksheet.getCell(`E${TR}`).alignment = aR;
      ['F','G','H','I','L','M','N','O'].forEach(c => {
        worksheet.getCell(`${c}${TR}`).value = { formula: `SUM(${c}7:${c}${lastDataRow})` };
      });

      // ── Summary rows
      [[TR+1,'TA',`SUM(F${TR},L${TR},H${TR})`],
       [TR+2,'DA',`SUM(G${TR},I${TR})`],
       [TR+3,'OE',`SUM(M${TR}:N${TR})`]].forEach(([row, lbl, formula]) => {
        worksheet.getRow(row).height = 18;
        const h = worksheet.getCell(`H${row}`); h.value = lbl; sc(h, fBB11, fillY, thin, aC);
        const i = worksheet.getCell(`I${row}`); i.value = { formula }; sc(i, fBB11, fillY, thin, aC);
      });

      const monthName = MONTHS.find(m => m.val === month).name.toUpperCase();
      const fileName  = `PSM_EXPENSE_${monthName}_${year}_AKASH.xlsx`;
      const buffer    = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName);
      showToast(`Downloaded ${fileName} ✓`);
    } catch (err) {
      console.error(err);
      showToast('Export failed — check console.', 'error');
    }
  };

  // ── JSX ───────────────────────────────────────────────────────────────────────
  const monthName = MONTHS.find(m => m.val === month)?.name;

  return (
    <div className="app-container">
      <div className="bg-grid" />

      {/* HEADER */}
      <header className="app-header">
        <div className="brand-section">
          <h1>📊 PSM Expense Generator</h1>
          <p>Akash Vishoriya &middot; auto-saved locally</p>
        </div>
        <div className="header-actions">
          <button className="btn-icon" onClick={() => setPanelOpen(o => !o)} title="Toggle settings">⚙️</button>
          <button className="btn-primary" onClick={exportToExcel}>⬇ Download Excel</button>
        </div>
      </header>

      {/* WORKSPACE */}
      <div className={`dashboard-workspace${panelOpen ? '' : ' panel-hidden'}`}>

        {/* CONTROL PANEL */}
        {panelOpen && (
          <aside className="control-panel">
            <button className="control-panel-close-btn" onClick={() => setPanelOpen(false)}>
              ✕ Close Settings
            </button>
            {/* Month / Year */}
            <div className="panel-section">
              <h2>📅 Month &amp; Year</h2>
              <div className="two-col">
                <div className="form-group">
                  <label htmlFor="sel-month">Month</label>
                  <select id="sel-month" className="form-control" value={month}
                    onChange={e => setMonth(+e.target.value)}>
                    {MONTHS.map(m => <option key={m.val} value={m.val}>{m.name}</option>)}
                  </select>
                </div>
                <div className="form-group">
                  <label htmlFor="sel-year">Year</label>
                  <select id="sel-year" className="form-control" value={year}
                    onChange={e => setYear(+e.target.value)}>
                    {YEARS.map(y => <option key={y} value={y}>{y}</option>)}
                  </select>
                </div>
              </div>
            </div>

            {/* Employee info */}
            <div className="panel-section">
              <h2>👤 Employee Info</h2>
              {[
                ['psm-name','PSM Name',    psmName,     setPsmName,   'text'],
                ['emp-id',  'Employee ID', empId,       setEmpId,     'text'],
                ['join-dt', 'Joining Date',joiningDate, setJoiningDate,'text'],
                ['hq-city', 'HQ City',    hq,          setHq,        'text'],
                ['mob-no',  'Mobile',     mobileNo,    setMobileNo,  'tel'],
                ['ase-nm',  'ASE Name',   aseName,     setAseName,   'text'],
              ].map(([id, label, val, setter, type]) => (
                <div className="form-group" key={id}>
                  <label htmlFor={id}>{label}</label>
                  <input id={id} className="form-control" type={type} value={val}
                    onChange={e => setter(e.target.value)} />
                </div>
              ))}
              <div className="form-group">
                <label htmlFor="rate-km">Bike Rate (₹/KM)</label>
                <input id="rate-km" className="form-control" type="number" value={ratePerKm}
                  onChange={e => setRatePerKm(parseFloat(e.target.value) || 0)} />
              </div>
            </div>

            {/* Bulk town */}
            <div className="panel-section">
              <h2>🏙️ Bulk Town Fill</h2>
              <p className="hint-text">Set one town for all working days</p>
              <div className="quick-badges">
                {QUICK_TOWNS.map(t => (
                  <button key={t} className={`badge-btn${bulkTown===t?' active':''}`}
                    onClick={() => setBulkTown(t)}>{t}</button>
                ))}
              </div>
              <div className="row-inline" style={{marginTop:'0.5rem'}}>
                <input className="form-control" type="text" placeholder="Or type a town…"
                  value={bulkTown} onChange={e => setBulkTown(e.target.value)} />
                <button className="btn-secondary" onClick={applyBulkTown}>Apply All</button>
              </div>
            </div>

            {/* Range updater */}
            <div className="panel-section">
              <h2>📆 Range Updater</h2>
              <p className="hint-text">Update multiple days in one go</p>
              <div className="two-col">
                <div className="form-group">
                  <label>From Day</label>
                  <select className="form-control" value={rangeFrom}
                    onChange={e => setRangeFrom(+e.target.value)}>
                    {Array.from({length:daysInMonth},(_,i)=>i+1).map(d=>(
                      <option key={d} value={d}>{d}</option>
                    ))}
                  </select>
                </div>
                <div className="form-group">
                  <label>To Day</label>
                  <select className="form-control" value={rangeTo}
                    onChange={e => setRangeTo(+e.target.value)}>
                    {Array.from({length:daysInMonth},(_,i)=>i+1).map(d=>(
                      <option key={d} value={d}>{d}</option>
                    ))}
                  </select>
                </div>
              </div>
              <div className="form-group">
                <label>Type</label>
                <select className="form-control" value={rangeType}
                  onChange={e => setRangeType(e.target.value)}>
                  <option value="HQ">HQ</option>
                  <option value="Upcountry">Upcountry</option>
                  <option value="Leave">Leave</option>
                  <option value="Holiday">Holiday</option>
                </select>
              </div>
              <div className="form-group">
                <label>Town (blank = keep existing)</label>
                <div className="quick-badges">
                  {QUICK_TOWNS.slice(0,4).map(t=>(
                    <button key={t} className={`badge-btn sm${rangeTown===t?' active':''}`}
                      onClick={()=>setRangeTown(t)}>{t}</button>
                  ))}
                </div>
                <input className="form-control" type="text" placeholder="Town name…"
                  value={rangeTown} onChange={e => setRangeTown(e.target.value)} />
              </div>
              <div className="two-col">
                <div className="form-group">
                  <label>KM</label>
                  <input className="form-control" type="number" placeholder="blank = keep"
                    value={rangeKm} onChange={e => setRangeKm(e.target.value)} />
                </div>
                <div className="form-group">
                  <label>Hotel ₹</label>
                  <input className="form-control" type="number" placeholder="blank = keep"
                    value={rangeHotel} onChange={e => setRangeHotel(e.target.value)} />
                </div>
              </div>
              <div className="form-group">
                <label>Bus/Train ₹</label>
                <input className="form-control" type="number" placeholder="blank = keep"
                  value={rangeBus} onChange={e => setRangeBus(e.target.value)} />
              </div>
              <button className="btn-primary full-width" onClick={applyRange}>
                ✅ Apply to Range
              </button>
            </div>

            {/* Summary */}
            <div className="summary-mini">
              <div className="sum-item"><span>TA</span><strong>₹{summaryTA}</strong></div>
              <div className="sum-item"><span>DA</span><strong>₹{summaryDA}</strong></div>
              <div className="sum-item"><span>OE</span><strong>₹{summaryOE}</strong></div>
              <div className="sum-item total"><span>Grand Total</span><strong>₹{totals.total}</strong></div>
            </div>
          </aside>
        )}

        {/* CONTENT */}
        <div className="preview-card">
          <div className="preview-header">
            <h2>
              {monthName} {year}
              <span className="day-count-badge">{daysInMonth} days</span>
            </h2>
            {!panelOpen && (
              <div className="summary-inline">
                <span>TA <strong>₹{summaryTA}</strong></span>
                <span>DA <strong>₹{summaryDA}</strong></span>
                <span>Total <strong>₹{totals.total}</strong></span>
              </div>
            )}
          </div>

          {isMobile
            ? <MobileCards rows={rows} hq={hq} onChange={handleCellChange}
                onCopy={copyFromPrev} onQuickTown={setQuickTown} />
            : <DesktopTable rows={rows} hq={hq} totals={totals} onChange={handleCellChange}
                onCopy={copyFromPrev} onQuickTown={setQuickTown} />
          }
        </div>
      </div>

      {/* TOAST */}
      {toast && (
        <div className={`toast-msg${toast.type === 'error' ? ' toast-error' : ''}`}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}

// ─── Mobile Cards ─────────────────────────────────────────────────────────────
function MobileCards({ rows, hq, onChange, onCopy, onQuickTown }) {
  return (
    <div className="mobile-cards">
      {rows.map((r, i) => {
        const isSun   = r.type === 'Sunday';
        const isOff   = r.type === 'Leave' || r.type === 'Holiday';
        return (
          <div key={r.dayNum} className={`day-card${isSun?' sunday-card':isOff?' leave-card':''}`}>
            {/* Card header */}
            <div className="card-header">
              <div className="card-date-block">
                <span className="card-daynum">{r.dayNum}</span>
                <div>
                  <div className="card-dayname">{r.dayName}</div>
                  <div className="card-datestr">{r.dateStr}</div>
                </div>
              </div>
              <div className="card-header-right">
                {isSun
                  ? <span className="sunday-badge">🌟 Sunday</span>
                  : <>
                      <button className="copy-btn" onClick={() => onCopy(i)} disabled={i === 0}
                        title="Copy from yesterday">⬆ Copy</button>
                      <span className={`total-badge${r.total > 0 ? ' has-total' : ''}`}>₹{r.total}</span>
                    </>
                }
              </div>
            </div>

            {/* Card body — skip for Sundays */}
            {!isSun && (
              <div className="card-body">
                {/* Type pills */}
                <div className="card-field">
                  <label>Type</label>
                  <div className="type-pills">
                    {['HQ','Upcountry','Leave','Holiday'].map(t => (
                      <button key={t}
                        className={`type-pill${r.type===t?' active-'+t.toLowerCase():''}`}
                        onClick={() => onChange(i, 'type', t)}>{t}</button>
                    ))}
                  </div>
                </div>

                {/* Town */}
                {!isOff && (
                  <div className="card-field">
                    <label>Town</label>
                    <input className="card-input" type="text" value={r.town}
                      onChange={e => onChange(i, 'town', e.target.value)} />
                    <div className="quick-badges compact">
                      {QUICK_TOWNS.map(t => (
                        <button key={t} className={`badge-btn sm${r.town===t?' active':''}`}
                          onClick={() => onQuickTown(i, t)}>{t}</button>
                      ))}
                    </div>
                  </div>
                )}

                {/* Numbers */}
                {!isOff && (
                  <div className="card-fields-grid">
                    <div className="card-field">
                      <label>KM</label>
                      <input className="card-input num" type="number" inputMode="decimal"
                        value={r.km || ''} placeholder="0"
                        onChange={e => onChange(i, 'km', e.target.value)} />
                    </div>
                    <div className="card-field">
                      <label>Hotel ₹</label>
                      <input className="card-input num" type="number" inputMode="decimal"
                        value={r.hotel || ''} placeholder="0"
                        onChange={e => onChange(i, 'hotel', e.target.value)} />
                    </div>
                    <div className="card-field">
                      <label>Bus/Train ₹</label>
                      <input className="card-input num" type="number" inputMode="decimal"
                        value={r.bus || ''} placeholder="0"
                        onChange={e => onChange(i, 'bus', e.target.value)} />
                    </div>
                    <div className="card-field">
                      <label>Bike Amt</label>
                      <div className="card-computed">₹{r.bikeAmt}</div>
                    </div>
                  </div>
                )}

                {/* Allowance summary */}
                <div className="card-amounts">
                  {r.type === 'HQ' && (
                    <span className="amt-tag hq">TA ₹{r.hqTa} · DA ₹{r.hqDa}</span>
                  )}
                  {r.type === 'Upcountry' && (
                    <span className="amt-tag up">UP DA ₹{r.upDa}</span>
                  )}
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ─── Desktop Table ────────────────────────────────────────────────────────────
function DesktopTable({ rows, hq, totals, onChange, onCopy, onQuickTown }) {
  const [townPopup, setTownPopup] = React.useState(null);
  const popupRef = React.useRef(null);

  // Close town popup when clicking outside
  React.useEffect(() => {
    if (townPopup === null) return;
    const handler = (e) => {
      if (popupRef.current && !popupRef.current.contains(e.target)) {
        setTownPopup(null);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [townPopup]);


  return (
    <div className="excel-view-container">
      <table className="excel-mock-table">
        <thead>
          <tr className="excel-section-header">
            <th colSpan={5}></th>
            <th colSpan={2}>HQ Local</th>
            <th colSpan={2}>Upcountry</th>
            <th colSpan={4}>KM / Bike</th>
            <th colSpan={2}>Other</th>
            <th></th>
            <th></th>
          </tr>
          <tr className="excel-column-header">
            {['#','Date','Day','Town','Type','HQ TA','HQ DA','UP TA','UP DA',
              'KM','Rate','Bike','Hotel','Bus','Total',''].map((h,i) => (
              <th key={i}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => {
            const isSun = r.type === 'Sunday';
            const isOff = r.type === 'Leave' || r.type === 'Holiday';
            return (
              <tr key={r.dayNum} className={isSun ? 'row-sunday' : isOff ? 'row-leave' : r.type === 'HQ' ? 'row-hq' : r.type === 'Upcountry' ? 'row-upcountry' : 'row-normal'}>
                <td>{r.dayNum}</td>
                <td className="td-date">{r.dateStr}</td>
                <td>{r.dayName}</td>
                <td className="td-town">
                  {!isSun && !isOff && <>
                    <input className="cell-input" type="text" value={r.town}
                      onFocus={() => setTownPopup(i)}
                      onClick={() => setTownPopup(i)}
                      onChange={e => onChange(i, 'town', e.target.value)} />
                    {townPopup === i && (
                      <div className="town-popup" ref={popupRef}>
                        {QUICK_TOWNS.map(t => (
                          <button key={t} className="badge-btn sm"
                            onMouseDown={(e) => { e.preventDefault(); onQuickTown(i, t); setTownPopup(null); }}>{t}</button>
                        ))}
                      </div>
                    )}
                  </>}
                </td>
                <td>
                  {!isSun && (
                    <select className="cell-select" value={r.type}
                      onChange={e => onChange(i, 'type', e.target.value)}>
                      <option value="HQ">HQ</option>
                      <option value="Upcountry">Upcountry</option>
                      <option value="Leave">Leave</option>
                      <option value="Holiday">Holiday</option>
                    </select>
                  )}
                </td>
                <td className="readonly-cell">{isSun?'':r.hqTa}</td>
                <td className="readonly-cell">{isSun?'':r.hqDa}</td>
                <td className="readonly-cell">{isSun?'':r.upTa}</td>
                <td className="readonly-cell">{isSun?'':r.upDa}</td>
                <td>
                  {!isSun && !isOff && (
                    <input className="cell-input num" type="number" inputMode="decimal"
                      value={r.km||''} placeholder="0"
                      onChange={e => onChange(i, 'km', e.target.value)} />
                  )}
                </td>
                <td className="readonly-cell">{r.rate}</td>
                <td className="formula-cell">{r.bikeAmt}</td>
                <td>
                  {!isSun && !isOff && (
                    <input className="cell-input num" type="number" inputMode="decimal"
                      value={r.hotel||''} placeholder="0"
                      onChange={e => onChange(i, 'hotel', e.target.value)} />
                  )}
                </td>
                <td>
                  {!isSun && !isOff && (
                    <input className="cell-input num" type="number" inputMode="decimal"
                      value={r.bus||''} placeholder="0"
                      onChange={e => onChange(i, 'bus', e.target.value)} />
                  )}
                </td>
                <td className="td-total">{r.total}</td>
                <td className="td-action">
                  {!isSun && i > 0 && (
                    <button className="copy-btn-sm" onClick={() => onCopy(i)}
                      title="Copy from yesterday">⬆</button>
                  )}
                </td>
              </tr>
            );
          })}

          {/* empty spacer rows */}
          <tr><td colSpan={16} style={{height:'8px',background:'transparent'}}></td></tr>
          <tr><td colSpan={16} style={{height:'8px',background:'transparent'}}></td></tr>

          {/* Totals */}
          <tr className="row-total-bottom">
            <td colSpan={5} style={{textAlign:'right',paddingRight:'1rem'}}>TOTAL</td>
            <td>{totals.hqTa}</td><td>{totals.hqDa}</td>
            <td>{totals.upTa}</td><td>{totals.upDa}</td>
            <td>{totals.km}</td><td></td>
            <td>{totals.bikeAmt}</td>
            <td>{totals.hotel}</td><td>{totals.bus}</td>
            <td>{totals.total}</td><td></td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}
