import React, { useState, useEffect, useCallback } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import './App.css';

const MONTHS = [
  { name: 'January', val: 1 },
  { name: 'February', val: 2 },
  { name: 'March', val: 3 },
  { name: 'April', val: 4 },
  { name: 'May', val: 5 },
  { name: 'June', val: 6 },
  { name: 'July', val: 7 },
  { name: 'August', val: 8 },
  { name: 'September', val: 9 },
  { name: 'October', val: 10 },
  { name: 'November', val: 11 },
  { name: 'December', val: 12 }
];

const YEARS = Array.from({ length: 11 }, (_, i) => 2024 + i);

const DAYS_OF_WEEK = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function App() {
  // --- Info Block State ---
  const [psmName, setPsmName] = useState('Akash Vishoriya');
  const [empId, setEmpId] = useState('I3785');
  const [joiningDate, setJoiningDate] = useState('25-04-2024');
  const [hq, setHq] = useState('Gwalior');
  const [mobile, setMobile] = useState('9074305446');
  const [aseName, setAseName] = useState('');
  const [ratePerKm, setRatePerKm] = useState(4);
  
  // --- Time Period State ---
  const [month, setMonth] = useState(4); // April default
  const [year, setYear] = useState(2026); // 2026 default
  
  // --- Data Grid State ---
  const [rows, setRows] = useState([]);
  const [bulkTown, setBulkTown] = useState('');
  const [toast, setToast] = useState('');

  // Helper to trigger floating toast message
  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(''), 4000);
  };

  // Re-generate grid whenever month/year/hq/rate details change
  useEffect(() => {
    const daysInMonth = new Date(year, month, 0).getDate();
    const newRows = [];
    
    for (let d = 1; d <= daysInMonth; d++) {
      const currentDate = new Date(year, month - 1, d);
      const dayIndex = currentDate.getDay();
      const dayName = DAYS_OF_WEEK[dayIndex];
      const isSunday = dayIndex === 0;
      
      const padDay = d.toString().padStart(2, '0');
      const padMonth = month.toString().padStart(2, '0');
      const dateStr = `${padDay}/${padMonth}/${year}`;

      if (isSunday) {
        newRows.push({
          dayNum: d,
          dateStr: dateStr,
          dayName: dayName,
          town: '',
          type: 'Sunday',
          hqTa: 0,
          hqDa: 0,
          upTa: 0,
          upDa: 0,
          km: 0,
          rate: ratePerKm,
          bikeAmt: 0,
          hotel: 0,
          bus: 0,
          total: 0
        });
      } else {
        newRows.push({
          dayNum: d,
          dateStr: dateStr,
          dayName: dayName,
          town: hq,
          type: 'HQ',
          hqTa: 175,
          hqDa: 175,
          upTa: 0,
          upDa: 0,
          km: 0,
          rate: ratePerKm,
          bikeAmt: 0,
          hotel: 0,
          bus: 0,
          total: 350 // TA (175) + DA (175)
        });
      }
    }
    setRows(newRows);
  }, [month, year, hq, ratePerKm]);

  // Handle cell input edits
  const handleCellChange = (index, field, value) => {
    const updatedRows = [...rows];
    const row = { ...updatedRows[index] };

    if (field === 'type') {
      row.type = value;
      if (value === 'HQ') {
        row.hqTa = 175;
        row.hqDa = 175;
        row.upTa = 0;
        row.upDa = 0;
        row.town = hq;
      } else if (value === 'Upcountry') {
        row.hqTa = 0;
        row.hqDa = 0;
        row.upTa = 0;
        row.upDa = 250;
      } else if (value === 'Sunday') {
        row.hqTa = 0;
        row.hqDa = 0;
        row.upTa = 0;
        row.upDa = 0;
        row.town = '';
        row.km = 0;
        row.hotel = 0;
        row.bus = 0;
      } else if (value === 'Leave' || value === 'Holiday') {
        row.hqTa = 0;
        row.hqDa = 0;
        row.upTa = 0;
        row.upDa = 0;
        row.town = '';
        row.km = 0;
        row.hotel = 0;
        row.bus = 0;
      }
    } else {
      // Numerical fields or Town
      if (field === 'town') {
        row.town = value;
      } else {
        const numVal = value === '' ? 0 : parseFloat(value);
        row[field] = isNaN(numVal) ? 0 : numVal;
      }
    }

    // Recalculate derived cells for this row
    row.bikeAmt = (row.km || 0) * (row.rate || 0);
    row.total = (row.hqTa || 0) + (row.hqDa || 0) + (row.upTa || 0) + (row.upDa || 0) + (row.bikeAmt || 0) + (row.hotel || 0) + (row.bus || 0);

    updatedRows[index] = row;
    setRows(updatedRows);
  };

  // Bulk set town names
  const applyBulkTown = () => {
    if (!bulkTown.trim()) return;
    const updatedRows = rows.map(r => {
      if (r.type !== 'Sunday' && r.type !== 'Leave' && r.type !== 'Holiday') {
        const bikeAmt = r.km * r.rate;
        const total = r.hqTa + r.hqDa + r.upTa + r.upDa + bikeAmt + r.hotel + r.bus;
        return { ...r, town: bulkTown.trim(), bikeAmt, total };
      }
      return r;
    });
    setRows(updatedRows);
    showToast(`Updated town to "${bulkTown}" for all active days!`);
  };

  // Calculate live preview totals for the grid bottom
  const getTotals = useCallback(() => {
    return rows.reduce(
      (acc, r) => {
        acc.hqTa += r.hqTa || 0;
        acc.hqDa += r.hqDa || 0;
        acc.upTa += r.upTa || 0;
        acc.upDa += r.upDa || 0;
        acc.km += r.km || 0;
        acc.bikeAmt += r.bikeAmt || 0;
        acc.hotel += r.hotel || 0;
        acc.bus += r.bus || 0;
        acc.total += r.total || 0;
        return acc;
      },
      { hqTa: 0, hqDa: 0, upTa: 0, upDa: 0, km: 0, bikeAmt: 0, hotel: 0, bus: 0, total: 0 }
    );
  }, [rows]);

  const totals = getTotals();
  const summaryTA = totals.hqTa + totals.bikeAmt + totals.upTa;
  const summaryDA = totals.hqDa + totals.upDa;
  const summaryOE = totals.hotel + totals.bus;

  // --- Excel Generation Logic using exceljs ---
  const exportToExcel = async () => {
    try {
      showToast('Generating Excel sheet...');
      
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Psm Exp Format', {
        views: [{ showGridLines: true }]
      });

      // --- 1. Column Widths ---
      worksheet.columns = [
        { key: 'A', width: 5.66 },
        { key: 'B', width: 14.55 },
        { key: 'C', width: 13.0 },
        { key: 'D', width: 13.0 },
        { key: 'E', width: 16.66 },
        { key: 'F', width: 10.33 },
        { key: 'G', width: 13.0 },
        { key: 'H', width: 13.0 },
        { key: 'I', width: 13.0 },
        { key: 'J', width: 13.0 },
        { key: 'K', width: 13.0 },
        { key: 'L', width: 13.0 },
        { key: 'M', width: 11.66 },
        { key: 'N', width: 17.33 },
        { key: 'O', width: 16.33 }
      ];

      // --- 2. Row Heights ---
      worksheet.getRow(1).height = 15.6;
      worksheet.getRow(2).height = 15.6;
      worksheet.getRow(3).height = 16.2;
      worksheet.getRow(4).height = 15.6;
      worksheet.getRow(5).height = 18.0;
      worksheet.getRow(6).height = 22.5;

      // --- 3. Style Palettes ---
      const thinBorder = {
        top: { style: 'thin', color: { argb: 'FF808080' } },
        left: { style: 'thin', color: { argb: 'FF808080' } },
        bottom: { style: 'thin', color: { argb: 'FF808080' } },
        right: { style: 'thin', color: { argb: 'FF808080' } }
      };

      const fillDarkGrey = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF404040' }
      };

      const fillYellow = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' }
      };

      const fontWhiteBold12 = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
      const fontWhiteBold11 = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      const fontBlackBold11 = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
      const fontBlackNorm11 = { name: 'Calibri', size: 11, bold: false, color: { argb: 'FF000000' } };

      const alignCenter = { horizontal: 'center', vertical: 'middle' };
      const alignLeft = { horizontal: 'left', vertical: 'middle' };

      // Helper to apply style to a cell safely
      const styleCell = (cell, font, fill, border, alignment) => {
        if (font) cell.font = font;
        if (fill) cell.fill = fill;
        if (border) cell.border = border;
        if (alignment) cell.alignment = alignment;
      };

      // Helper to merge and style
      const mergeAndStyleRange = (rangeStr, font, fill, alignment, border = thinBorder) => {
        const [startCellName, endCellName] = rangeStr.split(':');
        const startCell = worksheet.getCell(startCellName);
        const endCell = worksheet.getCell(endCellName);
        
        const startRow = startCell.row;
        const startCol = startCell.col;
        const endRow = endCell.row;
        const endCol = endCell.col;

        for (let r = startRow; r <= endRow; r++) {
          for (let c = startCol; c <= endCol; c++) {
            const cell = worksheet.getCell(r, c);
            styleCell(cell, font, fill, border, alignment);
          }
        }
        worksheet.mergeCells(startRow, startCol, endRow, endCol);
      };

      // --- 4. Render Info Block (Rows 1-3) ---
      // Build border wrapping box for outer frame
      for (let r = 1; r <= 3; r++) {
        for (let c = 1; c <= 15; c++) {
          const cell = worksheet.getCell(r, c);
          styleCell(cell, fontWhiteBold12, fillDarkGrey, null, alignLeft);
          cell.border = {
            top: r === 1 ? { style: 'medium', color: { argb: 'FF000000' } } : { style: 'thin', color: { argb: 'FF808080' } },
            bottom: r === 3 ? { style: 'medium', color: { argb: 'FF000000' } } : { style: 'thin', color: { argb: 'FF808080' } },
            left: c === 1 ? { style: 'medium', color: { argb: 'FF000000' } } : { style: 'thin', color: { argb: 'FF808080' } },
            right: c === 15 ? { style: 'medium', color: { argb: 'FF000000' } } : { style: 'thin', color: { argb: 'FF808080' } }
          };
        }
      }

      // Populate text values and merge
      const lastDay = new Date(year, month, 0).getDate();
      const padMonth = month.toString().padStart(2, '0');
      const lastDateStr = `${lastDay.toString().padStart(2, '0')}-${padMonth}-${year}`;
      const firstDateStr = `01/${padMonth}/${year}`;
      const formattedLastDateStr = `${lastDay.toString().padStart(2, '0')}/${padMonth}/${year}`;

      worksheet.getCell('A1').value = `Name of PSM : ${psmName}`;
      worksheet.getCell('E1').value = `Date of Joining : ${joiningDate}`;
      worksheet.getCell('M1').value = `Date : ${lastDateStr}`;
      
      worksheet.getCell('A2').value = `Emp ID : ${empId}`;
      worksheet.getCell('E2').value = `HQ : ${hq}`;

      worksheet.getCell('A3').value = `Mobile No : ${mobile}`;
      worksheet.getCell('E3').value = `ASE Name : ${aseName}`;
      worksheet.getCell('M3').value = `Expense from / to : ${firstDateStr} to ${formattedLastDateStr}`;

      worksheet.mergeCells('A1:D1');
      worksheet.mergeCells('E1:I1');
      worksheet.mergeCells('M1:O1');

      worksheet.mergeCells('A2:D2');
      worksheet.mergeCells('E2:I2');
      worksheet.mergeCells('M2:O2'); // empty but merged

      worksheet.mergeCells('A3:D3');
      worksheet.mergeCells('E3:I3');
      worksheet.mergeCells('M3:O3');

      // --- 5. Render Spacer Row 4 ---
      // (row 4 remains blank, no styling needed)

      // --- 6. Render Section Headers Row 5 ---
      mergeAndStyleRange('F5:G5', fontBlackBold11, fillYellow, alignCenter);
      worksheet.getCell('F5').value = 'HQ : LOCAL WORKING';

      mergeAndStyleRange('H5:N5', fontBlackBold11, fillYellow, alignCenter);
      worksheet.getCell('H5').value = 'UPCOUNTRY WORKING';

      // --- 7. Render Column Headers Row 6 ---
      const headers = [
        'S. No', 'Date', 'Day', 'Town', 'HQ / Upcountry',
        'TA', 'DA', 'TA', 'DA', 'Total KM', 'Rate / Km',
        'Bike Amt', 'Hotel', 'Bus / Train Amount', 'Total'
      ];
      headers.forEach((h, idx) => {
        const colLetter = String.fromCharCode(65 + idx);
        const cell = worksheet.getCell(`${colLetter}6`);
        worksheet.getRow(6).height = 22.5;
        styleCell(cell, fontWhiteBold11, fillDarkGrey, thinBorder, alignCenter);
        cell.value = h;
      });

      // --- 8. Render Data Rows (starting from Row 7) ---
      const DATA_START = 7;
      rows.forEach((r, idx) => {
        const excelRow = DATA_START + idx;
        worksheet.getRow(excelRow).height = 18.0;
        
        const isSunday = r.type === 'Sunday';
        
        // Write cells
        worksheet.getCell(`A${excelRow}`).value = r.dayNum;
        
        // Date Cell (date object for sorting, with formatting)
        const dateCell = worksheet.getCell(`B${excelRow}`);
        const [dayVal, monthVal, yearVal] = r.dateStr.split('/');
        dateCell.value = new Date(Date.UTC(parseInt(yearVal), parseInt(monthVal) - 1, parseInt(dayVal)));
        dateCell.numFmt = 'dd/mm/yyyy';

        worksheet.getCell(`C${excelRow}`).value = r.dayName;
        worksheet.getCell(`D${excelRow}`).value = isSunday ? null : r.town;
        worksheet.getCell(`E${excelRow}`).value = isSunday ? null : r.type;
        
        worksheet.getCell(`F${excelRow}`).value = isSunday ? null : r.hqTa;
        worksheet.getCell(`G${excelRow}`).value = isSunday ? null : r.hqDa;
        worksheet.getCell(`H${excelRow}`).value = isSunday ? null : r.upTa;
        worksheet.getCell(`I${excelRow}`).value = isSunday ? null : r.upDa;
        
        worksheet.getCell(`J${excelRow}`).value = isSunday ? null : r.km;
        worksheet.getCell(`K${excelRow}`).value = r.rate;
        
        // Bike Amt Formula
        worksheet.getCell(`L${excelRow}`).value = { formula: `J${excelRow}*K${excelRow}` };
        
        worksheet.getCell(`M${excelRow}`).value = isSunday ? null : r.hotel;
        worksheet.getCell(`N${excelRow}`).value = isSunday ? null : r.bus;
        
        // Total Formula
        worksheet.getCell(`O${excelRow}`).value = { 
          formula: `SUM(F${excelRow}:I${excelRow},L${excelRow},M${excelRow},N${excelRow})` 
        };

        // Apply styles to all 15 columns in this row
        for (let col = 1; col <= 15; col++) {
          const colLetter = String.fromCharCode(64 + col);
          const cell = worksheet.getCell(`${colLetter}${excelRow}`);
          
          if (isSunday && col >= 2 && col <= 7) {
            // Sunday row columns B to G are solid yellow bold
            styleCell(cell, fontBlackBold11, fillYellow, thinBorder, alignCenter);
          } else {
            // Normal styling
            styleCell(cell, fontBlackNorm11, null, thinBorder, alignCenter);
          }
        }
      });

      const lastDataRow = DATA_START + rows.length - 1;

      // --- 9. Render Blank Spacer Rows (3 blank rows gap) ---
      const totalSpacerStart = lastDataRow + 1;
      for (let r = totalSpacerStart; r < totalSpacerStart + 3; r++) {
        worksheet.getRow(r).height = 15.6;
        for (let col = 1; col <= 15; col++) {
          const colLetter = String.fromCharCode(64 + col);
          const cell = worksheet.getCell(`${colLetter}${r}`);
          styleCell(cell, fontBlackNorm11, null, thinBorder, alignCenter);
        }
      }

      // --- 10. Render Totals Row ---
      const TOTAL_ROW = totalSpacerStart + 3;
      worksheet.getRow(TOTAL_ROW).height = 20.0;
      
      const fillLightGreen = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC6EFCE' }
      };
      const fontGreenBold11 = {
        name: 'Calibri',
        size: 11,
        bold: true,
        color: { argb: 'FF276221' }
      };
      const totalBorder = {
        top: { style: 'thin', color: { argb: 'FF808080' } },
        left: { style: 'thin', color: { argb: 'FF808080' } },
        bottom: { style: 'double', color: { argb: 'FF276221' } },
        right: { style: 'thin', color: { argb: 'FF808080' } }
      };

      // Fill and border all cells in Total row
      for (let col = 1; col <= 15; col++) {
        const colLetter = String.fromCharCode(64 + col);
        const cell = worksheet.getCell(`${colLetter}${TOTAL_ROW}`);
        styleCell(cell, fontGreenBold11, fillLightGreen, totalBorder, alignCenter);
      }

      // Write values / formulas
      const labelCell = worksheet.getCell(`E${TOTAL_ROW}`);
      labelCell.value = 'TOTAL';
      labelCell.alignment = { horizontal: 'right', vertical: 'middle' };

      worksheet.getCell(`F${TOTAL_ROW}`).value = { formula: `SUM(F7:F${lastDataRow})` };
      worksheet.getCell(`G${TOTAL_ROW}`).value = { formula: `SUM(G7:G${lastDataRow})` };
      worksheet.getCell(`H${TOTAL_ROW}`).value = { formula: `SUM(H7:H${lastDataRow})` };
      worksheet.getCell(`I${TOTAL_ROW}`).value = { formula: `SUM(I7:I${lastDataRow})` };
      worksheet.getCell(`L${TOTAL_ROW}`).value = { formula: `SUM(L7:L${lastDataRow})` };
      worksheet.getCell(`M${TOTAL_ROW}`).value = { formula: `SUM(M7:M${lastDataRow})` };
      worksheet.getCell(`N${TOTAL_ROW}`).value = { formula: `SUM(N7:N${lastDataRow})` };
      worksheet.getCell(`O${TOTAL_ROW}`).value = { formula: `SUM(O7:O${lastDataRow})` };

      // --- 11. Render Summary Rows (TA/DA/OE) ---
      const SR1 = TOTAL_ROW + 1;
      const SR2 = TOTAL_ROW + 2;
      const SR3 = TOTAL_ROW + 3;

      const summaryConfigs = [
        { row: SR1, label: 'TA', formula: `SUM(F${TOTAL_ROW},L${TOTAL_ROW},H${TOTAL_ROW})` },
        { row: SR2, label: 'DA', formula: `SUM(G${TOTAL_ROW},I${TOTAL_ROW})` },
        { row: SR3, label: 'OE', formula: `SUM(M${TOTAL_ROW}:N${TOTAL_ROW})` }
      ];

      summaryConfigs.forEach(cfg => {
        worksheet.getRow(cfg.row).height = 18.0;
        
        // H Column: Label
        const cellH = worksheet.getCell(`H${cfg.row}`);
        cellH.value = cfg.label;
        styleCell(cellH, fontBlackBold11, fillYellow, thinBorder, alignCenter);

        // I Column: Formula
        const cellI = worksheet.getCell(`I${cfg.row}`);
        cellI.value = { formula: cfg.formula };
        styleCell(cellI, fontBlackBold11, fillYellow, thinBorder, alignCenter);
      });

      // --- 12. Save File ---
      const monthName = MONTHS.find(m => m.val === month).name.toUpperCase();
      const fileName = `PSM_EXPENSE_${monthName}_${year}_AKASH.xlsx`;

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      saveAs(blob, fileName);
      
      showToast(`Exported successfully: ${fileName}`);
    } catch (err) {
      console.error(err);
      showToast('Error exporting Excel file!');
    }
  };

  return (
    <div className="app-container">
      <div className="bg-grid"></div>
      
      {/* App Header */}
      <header className="app-header">
        <div className="brand-section">
          <h1>PSM Expense Sheet Generator</h1>
          <p>Create professional, pre-formatted travel expense journals for Akash Vishoriya</p>
        </div>
        <div className="export-actions">
          <button className="btn-primary" onClick={exportToExcel}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/>
            </svg>
            Download Excel (.xlsx)
          </button>
        </div>
      </header>

      {/* Main Workspace */}
      <main className="dashboard-workspace">
        
        {/* Left Side: Control Panel */}
        <div className="control-panel">
          
          <div className="panel-section">
            <h2>Select Month & Year</h2>
            <div className="date-selector-grid">
              <div className="form-group">
                <label htmlFor="select-month">Month</label>
                <select 
                  id="select-month" 
                  className="form-control" 
                  value={month} 
                  onChange={(e) => setMonth(parseInt(e.target.value))}
                >
                  {MONTHS.map(m => <option key={m.val} value={m.val}>{m.name}</option>)}
                </select>
              </div>
              <div className="form-group">
                <label htmlFor="select-year">Year</label>
                <select 
                  id="select-year" 
                  className="form-control" 
                  value={year} 
                  onChange={(e) => setYear(parseInt(e.target.value))}
                >
                  {YEARS.map(y => <option key={y} value={y}>{y}</option>)}
                </select>
              </div>
            </div>
          </div>

          <div className="panel-section">
            <h2>Employee Information</h2>
            <div className="form-group">
              <label htmlFor="psm-name">PSM Name</label>
              <input 
                id="psm-name" 
                className="form-control" 
                type="text" 
                value={psmName} 
                onChange={(e) => setPsmName(e.target.value)} 
              />
            </div>
            <div className="form-group">
              <label htmlFor="emp-id">Employee ID</label>
              <input 
                id="emp-id" 
                className="form-control" 
                type="text" 
                value={empId} 
                onChange={(e) => setEmpId(e.target.value)} 
              />
            </div>
            <div className="form-group">
              <label htmlFor="joining-date">Joining Date</label>
              <input 
                id="joining-date" 
                className="form-control" 
                type="text" 
                value={joiningDate} 
                onChange={(e) => setJoiningDate(e.target.value)} 
              />
            </div>
            <div className="form-group">
              <label htmlFor="hq-town">Headquarters (HQ)</label>
              <input 
                id="hq-town" 
                className="form-control" 
                type="text" 
                value={hq} 
                onChange={(e) => setHq(e.target.value)} 
              />
            </div>
            <div className="form-group">
              <label htmlFor="mobile-no">Mobile Number</label>
              <input 
                id="mobile-no" 
                className="form-control" 
                type="text" 
                value={mobile} 
                onChange={(e) => setMobile(e.target.value)} 
              />
            </div>
            <div className="form-group">
              <label htmlFor="ase-name">ASE Name</label>
              <input 
                id="ase-name" 
                className="form-control" 
                type="text" 
                value={aseName} 
                placeholder="N/A"
                onChange={(e) => setAseName(e.target.value)} 
              />
            </div>
          </div>

          <div className="panel-section">
            <h2>Defaults & Bulk Actions</h2>
            <div className="form-group">
              <label htmlFor="rate-per-km">Bike Rate (Per KM)</label>
              <input 
                id="rate-per-km" 
                className="form-control" 
                type="number" 
                value={ratePerKm} 
                onChange={(e) => setRatePerKm(parseFloat(e.target.value) || 0)} 
              />
            </div>
            <div className="form-group" style={{ marginTop: '0.5rem' }}>
              <label htmlFor="bulk-town">Bulk Fill Town Name</label>
              <div style={{ display: 'flex', gap: '0.5rem' }}>
                <input 
                  id="bulk-town" 
                  className="form-control" 
                  type="text" 
                  placeholder="e.g. Morena" 
                  value={bulkTown} 
                  onChange={(e) => setBulkTown(e.target.value)}
                  style={{ flex: 1 }}
                />
                <button className="btn-secondary" style={{ padding: '0 1rem' }} onClick={applyBulkTown}>
                  Apply
                </button>
              </div>
            </div>
          </div>

        </div>

        {/* Right Side: Spreadsheet Preview */}
        <div className="preview-card">
          <div className="preview-header">
            <h2>Interactive Grid Preview</h2>
            <span style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>
              Values edited here are included in the downloaded Excel file.
            </span>
          </div>

          {/* Simulated Excel Grid */}
          <div className="excel-view-container">
            <table className="excel-mock-table">
              
              {/* Header Blocks (Rows 1-3 Info Cards) */}
              <thead className="excel-mock-header-block">
                <tr className="excel-info-row">
                  <td colSpan="4" className="info-label">Name of PSM : {psmName}</td>
                  <td colSpan="5">Date of Joining : {joiningDate}</td>
                  <td colSpan="3"></td>
                  <td colSpan="3">Date : {new Date(year, month, 0).getDate().toString().padStart(2, '0')}-{month.toString().padStart(2, '0')}-{year}</td>
                </tr>
                <tr className="excel-info-row">
                  <td colSpan="4" className="info-label">Emp ID : {empId}</td>
                  <td colSpan="5">HQ : {hq}</td>
                  <td colSpan="6"></td>
                </tr>
                <tr className="excel-info-row">
                  <td colSpan="4" className="info-label">Mobile No : {mobile}</td>
                  <td colSpan="5">ASE Name : {aseName || 'N/A'}</td>
                  <td colSpan="3"></td>
                  <td colSpan="3">
                    Expense from/to: 01/{month.toString().padStart(2, '0')}/{year} to {new Date(year, month, 0).getDate().toString().padStart(2, '0')}/{month.toString().padStart(2, '0')}/{year}
                  </td>
                </tr>
                
                {/* Row 4 Spacer */}
                <tr style={{ height: '15.6px' }}><td colSpan="15" style={{ border: 'none', background: 'transparent' }}></td></tr>
                
                {/* Row 5 Sections */}
                <tr className="excel-section-header">
                  <td colSpan="5" style={{ background: 'transparent', border: 'none' }}></td>
                  <td colSpan="2">HQ : LOCAL WORKING</td>
                  <td colSpan="7">UPCOUNTRY WORKING</td>
                  <td style={{ background: 'transparent', border: 'none' }}></td>
                </tr>

                {/* Row 6 Headers */}
                <tr className="excel-column-header">
                  <th>S. No</th>
                  <th>Date</th>
                  <th>Day</th>
                  <th>Town</th>
                  <th>HQ / Upcountry</th>
                  <th>TA</th>
                  <th>DA</th>
                  <th>TA</th>
                  <th>DA</th>
                  <th>Total KM</th>
                  <th>Rate/Km</th>
                  <th>Bike Amt</th>
                  <th>Hotel</th>
                  <th>Bus / Train</th>
                  <th>Total</th>
                </tr>
              </thead>

              {/* Data Grid Rows */}
              <tbody>
                {rows.map((r, index) => {
                  const isSunday = r.type === 'Sunday';
                  const isLeave = r.type === 'Leave';
                  const isHoliday = r.type === 'Holiday';
                  
                  let rowClass = 'row-normal';
                  if (isSunday) rowClass = 'row-sunday';
                  else if (isLeave) rowClass = 'row-leave';
                  else if (isHoliday) rowClass = 'row-holiday';

                  return (
                    <tr key={index} className={rowClass}>
                      {/* S.No */}
                      <td>{r.dayNum}</td>
                      
                      {/* Date */}
                      <td style={{ minWidth: '100px' }}>{r.dateStr}</td>
                      
                      {/* Day Name */}
                      <td>{r.dayName}</td>
                      
                      {/* Town */}
                      <td>
                        <input
                          className="cell-input"
                          type="text"
                          value={r.town}
                          disabled={isSunday || isLeave || isHoliday}
                          onChange={(e) => handleCellChange(index, 'town', e.target.value)}
                        />
                      </td>

                      {/* Travel Type */}
                      <td>
                        <select
                          className="cell-select"
                          value={r.type}
                          onChange={(e) => handleCellChange(index, 'type', e.target.value)}
                        >
                          <option value="HQ">HQ</option>
                          <option value="Upcountry">Upcountry</option>
                          <option value="Leave">Leave</option>
                          <option value="Holiday">Holiday</option>
                          <option value="Sunday">Sunday</option>
                        </select>
                      </td>

                      {/* HQ TA */}
                      <td className="readonly-cell">{isSunday ? '' : r.hqTa}</td>
                      
                      {/* HQ DA */}
                      <td className="readonly-cell">{isSunday ? '' : r.hqDa}</td>
                      
                      {/* UP TA */}
                      <td>
                        <input
                          className="cell-input"
                          type="number"
                          value={isSunday ? '' : r.upTa}
                          disabled={isSunday || isLeave || isHoliday || r.type === 'HQ'}
                          onChange={(e) => handleCellChange(index, 'upTa', e.target.value)}
                        />
                      </td>

                      {/* UP DA */}
                      <td className="readonly-cell">{isSunday ? '' : r.upDa}</td>

                      {/* KM Travelled */}
                      <td>
                        <input
                          className="cell-input"
                          type="number"
                          value={isSunday ? '' : r.km}
                          disabled={isSunday || isLeave || isHoliday}
                          onChange={(e) => handleCellChange(index, 'km', e.target.value)}
                        />
                      </td>

                      {/* Rate */}
                      <td className="readonly-cell">{r.rate}</td>

                      {/* Bike Amt */}
                      <td className="formula-cell">{isSunday ? '0' : r.bikeAmt}</td>

                      {/* Hotel */}
                      <td>
                        <input
                          className="cell-input"
                          type="number"
                          value={isSunday ? '' : r.hotel}
                          disabled={isSunday || isLeave || isHoliday}
                          onChange={(e) => handleCellChange(index, 'hotel', e.target.value)}
                        />
                      </td>

                      {/* Bus / Train */}
                      <td>
                        <input
                          className="cell-input"
                          type="number"
                          value={isSunday ? '' : r.bus}
                          disabled={isSunday || isLeave || isHoliday}
                          onChange={(e) => handleCellChange(index, 'bus', e.target.value)}
                        />
                      </td>

                      {/* Total */}
                      <td className="formula-cell">{r.total}</td>
                    </tr>
                  );
                })}

                {/* 3 Spacer Rows Gap in Preview */}
                <tr style={{ height: '18px' }}><td colSpan="15" style={{ background: 'rgba(255,255,255,0.01)', borderLeft: 'none', borderRight: 'none' }}></td></tr>
                <tr style={{ height: '18px' }}><td colSpan="15" style={{ background: 'rgba(255,255,255,0.01)', borderLeft: 'none', borderRight: 'none' }}></td></tr>
                <tr style={{ height: '18px' }}><td colSpan="15" style={{ background: 'rgba(255,255,255,0.01)', borderLeft: 'none', borderRight: 'none' }}></td></tr>

                {/* Totals Row */}
                <tr className="row-total-bottom">
                  <td colSpan="5" style={{ textAlign: 'right', paddingRight: '1rem' }}>TOTAL</td>
                  <td>{totals.hqTa}</td>
                  <td>{totals.hqDa}</td>
                  <td>{totals.upTa}</td>
                  <td>{totals.upDa}</td>
                  <td>{totals.km}</td>
                  <td></td>
                  <td>{totals.bikeAmt}</td>
                  <td>{totals.hotel}</td>
                  <td>{totals.bus}</td>
                  <td>{totals.total}</td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Summary badging */}
          <div className="summary-container">
            <div className="summary-card">
              <span className="summary-label">TA</span>
              <span className="summary-value">₹ {summaryTA}</span>
            </div>
            <div className="summary-card">
              <span className="summary-label">DA</span>
              <span className="summary-value">₹ {summaryDA}</span>
            </div>
            <div className="summary-card">
              <span className="summary-label">OE</span>
              <span className="summary-value">₹ {summaryOE}</span>
            </div>
          </div>

        </div>
      </main>

      {/* Footer */}
      <footer className="app-footer">
        PSM Expense Sheet Generator App &bull; Powered by React and ExcelJS
      </footer>

      {/* Floating Status Toast */}
      {toast && (
        <div className="toast-msg">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" style={{ marginRight: '4px' }}>
            <polyline points="20 6 9 17 4 12"></polyline>
          </svg>
          {toast}
        </div>
      )}
    </div>
  );
}

export default App;
