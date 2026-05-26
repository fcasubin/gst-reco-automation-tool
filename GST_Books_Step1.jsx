import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ── Date / Month ───────────────────────────────────────────────────────────────
const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const FY_ORDER = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'];

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const n = Number(v);
  if (!isNaN(n) && n > 40000 && n < 70000) return new Date(Math.round((n-25569)*86400000));
  if (typeof v === 'string') { const d=new Date(v); if (!isNaN(d)) return d; }
  return null;
}
function monthKey(d) {
  const dt = toDate(d);
  if (!dt) return null;
  const dd = dt instanceof Date ? dt : new Date(dt);
  return `${MON[dd.getUTCMonth()]}-${String(dd.getUTCFullYear()).slice(2)}`;
}
function sortFY(keys) {
  return [...new Set(keys)].sort((a,b) => {
    const [ma,ya]=a.split('-'), [mb,yb]=b.split('-');
    const yd=Number(ya)-Number(yb);
    return yd!==0?yd:FY_ORDER.indexOf(ma)-FY_ORDER.indexOf(mb);
  });
}
function fmtINR(n) {
  if (!n || Math.abs(n)<0.005) return '—';
  return (n<0?'−':'')+Math.abs(n).toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2});
}

// ── Column classification ──────────────────────────────────────────────────────
function extractRate(raw) {
  const c = String(raw||'').replace(/\s+/g,'');
  const m = c.match(/[@]?(\d+(?:\.\d+)?)%/);
  return m ? parseFloat(m[1])/100 : null;
}
function classifyCol(raw, allowInput=false) {
  if (!raw) return null;
  const u = String(raw).toUpperCase().replace(/\s+/g,'');  // space-stripped uppercase

  // ── STEP 1: GST tax columns — checked FIRST before any skip logic ──────────
  const isGSTCol = u.includes('CGST')||u.includes('SGST')||u.includes('IGST')||
                   u.includes('UTGST')||u.includes('CESS');

  if (isGSTCol) {
    if (!allowInput && u.includes('INPUT')) return null;
    if (u.includes('CGST'))                      return {role:'TAX',type:'cgst',rate:extractRate(raw)};
    if (u.includes('SGST')||u.includes('UTGST')) return {role:'TAX',type:'sgst',rate:extractRate(raw)};
    if (u.includes('IGST'))                      return {role:'TAX',type:'igst',rate:extractRate(raw)};
    if (u.includes('CESS'))                      return {role:'TAX',type:'cess',rate:extractRate(raw)};
  }

  // ── STEP 2: Structural / skip columns ─────────────────────────────────────
  const skipKeywords = [
    'DATE', 'INVOICE', 'VOUCHER', 'VCH', 'GSTIN', 'CONSIGNEE', 'NARRATION', 'PARTICULARS', 'PARTY', 'ROUNDOFF', 'ROUND'
  ];
  if (skipKeywords.some(kw => u.includes(kw))) {
    return { role: 'SKIP' };
  }
  if (u === 'GROSSTOTAL') return { role: 'SKIP' };
  if (u === 'VALUE') return { role: 'TAXABLE_VALUE' };

  // ── STEP 3: Taxable value / Sales ledger columns ───────────────────────────
  const hasSale = u.includes('SALE') || u.includes('INTERSTATE');
  if (hasSale && (allowInput || !u.includes('RETURN')) && !isGSTCol) {
    return {role:'SALE_LEDGER', type:'taxable', rate:extractRate(raw)};
  }

  return null;
}

// ── Journal sign detection ─────────────────────────────────────────────────────
// Tally Journal Register exports ALL individual-row amounts as ABSOLUTE POSITIVE
// regardless of Dr/Cr side. However the Grand Total row contains the correct
// NET value (positive = net Credit = addition to output tax).
// Algorithm:
//   1. Identify party/balancing columns in each row (column ≈ sum of all others)
//   2. Track which columns co-appear with GST (these are taxable value columns)
//   3. Use Grand Total to solve sign assignment for each row:
//      find +1/-1 for each row such that sum(sign × cgst_value) = GT_cgst
//   4. Apply consistent signs to taxable columns in same rows
function solveRowSigns(rows, cgstColIdx, gtCgst) {
  const vals = rows.map(r => Math.abs(Number(r[cgstColIdx]||0)));
  const n = vals.length;
  if (n === 0) return [];
  if (n === 1) return [gtCgst >= 0 ? 1 : -1];
  if (n > 15) return new Array(n).fill(1); // too many — default to positive

  for (let mask = 0; mask < (1 << n); mask++) {
    const signs = Array.from({length:n}, (_,i) => (mask>>i)&1 ? 1 : -1);
    const total = signs.reduce((s,sg,i) => s + sg*vals[i], 0);
    if (Math.abs(total - gtCgst) < 1) return signs;
  }
  return new Array(n).fill(1); // no valid combo found — treat all as Cr
}

function extractDataAndGrandTotal(rows, hRow) {
  const dataRows = [];
  let grandTotalRow = null;
  const rawRows = rows.slice(hRow + 1);
  for (const row of rawRows) {
    if (!row) continue;
    const partName = String(row[1] ?? '').trim().toLowerCase();
    if (partName === 'grand total') {
      grandTotalRow = row;
      continue;
    }
    // Skip empty lines (must have date in Col A)
    if (!row[0]) continue;
    dataRows.push(row);
  }
  return { dataRows, grandTotalRow };
}

function computeSheetCalculations(regType, sign, colMap, dataRows, grandTotalRow, fileName, clientName) {
  const taxCols = colMap.filter(c => c.cls?.role === 'TAX');
  
  if (regType === 'JOURNAL') {
    if (dataRows.length === 0 || !grandTotalRow) {
      return {
        totals: {taxable:0,cgst:0,sgst:0,igst:0,cess:0},
        monthly: {},
        rowCount: 0,
        taxCols: [],
        journalInfo: { totalCrTaxable: 0, totalDrTaxable: 0, gstRowCount: 0 }
      };
    }
    const firstGSTCol = taxCols[0];
    if (!firstGSTCol) {
      return {
        totals: {taxable:0,cgst:0,sgst:0,igst:0,cess:0},
        monthly: {},
        rowCount: 0,
        taxCols: [],
        journalInfo: { totalCrTaxable: 0, totalDrTaxable: 0, gstRowCount: 0 }
      };
    }

    const gstRows = dataRows.filter(r =>
      taxCols.some(col => Math.abs(Number(r[col.i]||0)) > 0.005)
    );
    if (gstRows.length === 0) {
      return {
        totals: {taxable:0,cgst:0,sgst:0,igst:0,cess:0},
        monthly: {},
        rowCount: 0,
        taxCols: taxCols.map(c=>c.h),
        journalInfo: { totalCrTaxable: 0, totalDrTaxable: 0, gstRowCount: 0 }
      };
    }

    const partyColSet = new Set();
    const taxableColSet = new Set();

    const STRUCTURAL = new Set(
      colMap.filter(c => {
        if (c.cls?.role === 'SKIP') return true;
        const u = c.h.toUpperCase().replace(/\s+/g,'');
        return ['DATE','PARTICULARS','PARTY','CONSIGNEE','VOUCHERTYPE','VOUCHERNO',
                'VOUCHERREF','NARRATION'].some(s => u.startsWith(s) || u===s)
          || u === 'GROSSTOTAL' || u === 'GSTINUIN' || u === 'GSTIN/UIN';
      }).map(c => c.i)
    );
    const GST_COLS = new Set(taxCols.map(c => c.i));

    // Dynamic companion ledger detection using the elimination/balancing method
    for (const row of gstRows) {
      const active = [];
      for (let i = 0; i < row.length; i++) {
        if (STRUCTURAL.has(i)) continue;
        const rawVal = row[i];
        if (typeof rawVal === 'string' && isNaN(Number(rawVal))) continue;
        const v = Math.abs(Number(rawVal||0));
        if (v > 0.005) active.push({ i, v });
      }

      for (const cand of active) {
        const othersSum = active
          .filter(c => c.i !== cand.i)
          .reduce((s,c) => s + c.v, 0);
        if (Math.abs(cand.v - othersSum) < 2) {
          partyColSet.add(cand.i);
        }
      }

      for (const c of active) {
        if (!GST_COLS.has(c.i) && !partyColSet.has(c.i)) {
          taxableColSet.add(c.i);
        }
      }
    }

    let signingCol = taxCols.find(c => Math.abs(Number(grandTotalRow[c.i]||0)) > 0.005);
    if (!signingCol) signingCol = taxCols[0];
    const gtGST = Number(grandTotalRow[signingCol.i]||0);
    const rowSigns = solveRowSigns(gstRows, signingCol.i, gtGST);

    const monthly = {};
    let totalDrTaxable = 0;
    let totalCrTaxable = 0;

    gstRows.forEach((row, idx) => {
      const mk = monthKey(row[0]);
      if (!mk) return;
      const signVal = rowSigns[idx] || 1;
      if (!monthly[mk]) monthly[mk] = {taxable:0,cgst:0,sgst:0,igst:0,cess:0};

      let rowTaxable = 0;
      for (const ti of taxableColSet) {
        if (partyColSet.has(ti)) continue;
        if (colMap[ti]?.cls?.role === 'SKIP') continue;
        const val = Number(row[ti]||0);
        if (!isNaN(val)) rowTaxable += Math.abs(val);
      }
      
      colMap.forEach(col => {
        if ((col.cls?.role === 'SALE_LEDGER' || col.cls?.role === 'TAXABLE_VALUE') && !taxableColSet.has(col.i)) {
          const val = Number(row[col.i]||0);
          if (!isNaN(val)) rowTaxable += Math.abs(val);
        }
      });
      
      if (signVal === 1) {
        totalCrTaxable += rowTaxable; // Credit increases output tax liability in Step 1
      } else {
        totalDrTaxable += rowTaxable; // Debit decreases output tax liability in Step 1
      }

      monthly[mk].taxable += rowTaxable * signVal;

      for (const col of taxCols) {
        const v = Math.abs(Number(row[col.i]||0));
        if (isNaN(v)) continue;
        const signed = v * signVal;
        if (col.cls.type==='cgst') monthly[mk].cgst += signed;
        if (col.cls.type==='sgst') monthly[mk].sgst += signed;
        if (col.cls.type==='igst') monthly[mk].igst += signed;
        if (col.cls.type==='cess') monthly[mk].cess += signed;
      }
    });

    const totals = Object.values(monthly).reduce((a,m)=>({
      taxable:a.taxable+m.taxable,cgst:a.cgst+m.cgst,
      sgst:a.sgst+m.sgst,igst:a.igst+m.igst,cess:a.cess+m.cess,
    }),{taxable:0,cgst:0,sgst:0,igst:0,cess:0});

    return {
      totals,
      monthly,
      rowCount: gstRows.length,
      taxCols: taxCols.map(c=>c.h),
      journalInfo: {
        totalDrTaxable,
        totalCrTaxable,
        gstRowCount: gstRows.length
      }
    };
  } else {
    const monthly = {};
    let rowCount = 0;
    const valColI = colMap.findIndex(c => c.cls?.role === 'TAXABLE_VALUE');
    const saleLedgerCols = colMap.filter(c => c.cls?.role === 'SALE_LEDGER');

    for (const row of dataRows) {
      const mk = monthKey(row[0]);
      if (!mk) continue;
      if (!monthly[mk]) monthly[mk] = {taxable:0,cgst:0,sgst:0,igst:0,cess:0};

      let taxable = 0;
      if (valColI >= 0) {
        const val = Number(row[valColI]??0);
        if (!isNaN(val)) taxable = val;
      } else {
        taxable = saleLedgerCols.reduce((s,col) => {
          const val = Number(row[col.i]??0);
          return s + (isNaN(val) ? 0 : val);
        }, 0);
      }
      let cgst=0,sgst=0,igst=0,cess=0;
      for (const col of taxCols) {
        const v=Number(row[col.i]??0);
        if (!v || isNaN(v)) continue;
        if (col.cls.type==='cgst') cgst+=v;
        if (col.cls.type==='sgst') sgst+=v;
        if (col.cls.type==='igst') igst+=v;
        if (col.cls.type==='cess') cess+=v;
      }
      monthly[mk].taxable += taxable*sign;
      monthly[mk].cgst    += cgst*sign;
      monthly[mk].sgst    += sgst*sign;
      monthly[mk].igst    += igst*sign;
      monthly[mk].cess    += cess*sign;
      rowCount++;
    }

    const totals = Object.values(monthly).reduce((a,m)=>({
      taxable:a.taxable+m.taxable,cgst:a.cgst+m.cgst,
      sgst:a.sgst+m.sgst,igst:a.igst+m.igst,cess:a.cess+m.cess,
    }),{taxable:0,cgst:0,sgst:0,igst:0,cess:0});

    return {
      totals,
      monthly,
      rowCount,
      taxCols: taxCols.map(c=>c.h),
      journalInfo: null
    };
  }
}

function getUnclassifiedDataColumns(sheet) {
  const unclassified = [];
  const colMap = sheet.colMap;
  const isJournal = sheet.regType === 'JOURNAL';
  const STRUCTURAL = new Set(
    colMap.filter(c => {
      if (c.cls?.role === 'SKIP') return true;
      if (c.cls) return false;
      const u = c.h.toUpperCase().replace(/\s+/g,'');
      return ['DATE','PARTICULARS','PARTY','CONSIGNEE','VOUCHERTYPE','VOUCHERNO',
              'VOUCHERREF','NARRATION'].some(s => u.startsWith(s) || u===s)
        || u === 'GROSSTOTAL' || u === 'GSTINUIN' || u === 'GSTIN/UIN';
    }).map(c => c.i)
  );

  colMap.forEach(col => {
    if (col.cls) return;
    if (STRUCTURAL.has(col.i)) return;
    if (isJournal) return;
    
    let hasValues = false;
    let sum = 0;
    let sample = [];
    
    for (const row of sheet.dataRows) {
      const val = Number(row[col.i]);
      if (!isNaN(val) && Math.abs(val) > 0.005) {
        hasValues = true;
        sum += Math.abs(val);
        if (sample.length < 3) {
          sample.push(val);
        }
      }
    }
    
    if (hasValues) {
      unclassified.push({
        i: col.i,
        h: col.h,
        sum,
        sample
      });
    }
  });
  
  return unclassified;
}

function detectRegisterType(rows, sheetName='') {
  const sn = sheetName.toLowerCase();
  if (sn.includes('credit note')) return 'CREDIT_NOTE';
  if (sn.includes('journal'))     return 'JOURNAL';
  if (sn.includes('sales'))      return 'SALES';
  for (let i = 3; i <= 13; i++) {
    const c = String(rows[i]?.[0]||'').trim().toLowerCase();
    if (c.includes('credit note')) return 'CREDIT_NOTE';
    if (c.includes('journal'))     return 'JOURNAL';
    if (c === 'sales register' || (c.includes('sales') && c.includes('register'))) return 'SALES';
  }
  return 'SALES';
}

function parseRegister(rows, fileName, sheetName='') {
  const regType    = detectRegisterType(rows, sheetName);
  const clientName = String(rows[0]?.[0]||'').trim() || 'Unknown';
  const sign       = regType==='CREDIT_NOTE' ? -1 : 1;

  let hRow=-1;
  for (let i=0;i<Math.min(rows.length,15);i++) {
    if (String(rows[i]?.[0]||'').trim().toLowerCase()==='date') { hRow=i; break; }
  }
  if (hRow<0) return null;

  const headers = rows[hRow].map(h=>String(h??'').trim());
  const allowInput = regType !== 'JOURNAL';
  const colMap  = headers.map((h,i)=>({i,h,cls:classifyCol(h, allowInput)}));
  
  const { dataRows, grandTotalRow } = extractDataAndGrandTotal(rows, hRow);

  const calcs = computeSheetCalculations(regType, sign, colMap, dataRows, grandTotalRow, fileName, clientName);
  if (!calcs) return null;

  return {
    id: `${clientName}::${regType}::${fileName}::${sheetName}`,
    dedupKey: `${clientName}::${regType}`,
    fileName,
    sheetName,
    regType,
    typeLabel: regType === 'SALES' ? 'Sales Register' : regType === 'CREDIT_NOTE' ? 'Credit Note Register' : 'Journal Register',
    sign,
    clientName,
    headers,
    dataRows,
    grandTotalRow,
    colMap,
    originalColMap: JSON.parse(JSON.stringify(colMap)),
    enabled: true,
    isDuplicate: false,
    ...calcs
  };
}

// ── Compute monthly ────────────────────────────────────────────────────────────
function computeMonthly(sheets) {
  const agg={};
  for (const s of sheets.filter(s=>s.enabled)) {
    for (const [mk,d] of Object.entries(s.monthly)) {
      if (!agg[mk]) agg[mk]={taxable:0,cgst:0,sgst:0,igst:0,cess:0};
      agg[mk].taxable+=d.taxable; agg[mk].cgst+=d.cgst;
      agg[mk].sgst+=d.sgst; agg[mk].igst+=d.igst; agg[mk].cess+=d.cess;
    }
  }
  for (const d of Object.values(agg)) {
    d.totalTax=d.cgst+d.sgst+d.igst+d.cess;
    d.grossTotal=d.taxable+d.totalTax;
  }
  return {monthly:agg, months:sortFY(Object.keys(agg))};
}

// ── Excel export ───────────────────────────────────────────────────────────────
function doExport(monthly, months, sheets, plSales) {
  const wb  = XLSX.utils.book_new();
  const clients = [...new Set(sheets.map(s=>s.clientName).filter(Boolean))].join(', ');
  
  const startRow = 7;
  const endRow = 6 + months.length;
  
  const aoa=[
    ['GST Output Tax — Monthly Breakup (Books Data)'],
    [`Client: ${clients}`],
    ['Firm: Subin B & Associates, Chartered Accountants'],
    ['Generated: '+new Date().toLocaleDateString('en-IN',{day:'2-digit',month:'short',year:'numeric'})],
    [],
    ['Month','Taxable Value','IGST','CGST','SGST','Cess','Total Tax','Gross Total'],
  ];
  
  months.forEach(m=>{
    const d=monthly[m]||{};
    aoa.push([m,d.taxable||0,d.igst||0,d.cgst||0,d.sgst||0,d.cess||0,d.totalTax||0,d.grossTotal||0]);
  });
  
  // Total row using dynamic Excel SUM formulas
  const totalRow = [
    'TOTAL',
    { f: `SUM(B${startRow}:B${endRow})` },
    { f: `SUM(C${startRow}:C${endRow})` },
    { f: `SUM(D${startRow}:D${endRow})` },
    { f: `SUM(E${startRow}:E${endRow})` },
    { f: `SUM(F${startRow}:F${endRow})` },
    { f: `SUM(G${startRow}:G${endRow})` },
    { f: `SUM(H${startRow}:H${endRow})` }
  ];
  aoa.push(totalRow);
  
  if (plSales) {
    const totalRowIdx = 7 + months.length; // TOTAL row index
    const plSalesIdx = totalRowIdx + 3;
    const booksIdx = totalRowIdx + 4;
    const diffIdx = totalRowIdx + 5;
    
    aoa.push(
      [],
      ['P&L Cross-Check'],
      ['Sales as per P&L', plSales],
      ['Taxable Value (Books)', { f: `B${totalRowIdx}` }],
      ['Difference', { f: `B${plSalesIdx}-B${booksIdx}` }, '', '', '', '', '', { f: `IF(ABS(B${diffIdx})<1,"✓ Match","⚠ Check required")` }]
    );
  }
  
  const ws1=XLSX.utils.aoa_to_sheet(aoa);
  
  // Cell formatting for Monthly Breakup
  for (const key in ws1) {
    if (key[0] === '!') continue;
    const cell = ws1[key];
    if (!cell) continue;
    
    const colChar = key.replace(/[0-9]/g, '');
    const rowNum = parseInt(key.replace(/[^0-9]/g, ''), 10);
    
    const isNumericCol = ['B', 'C', 'D', 'E', 'F', 'G', 'H'].includes(colChar);
    const isNumericRow = rowNum >= 7;
    
    if (isNumericCol && isNumericRow) {
      cell.z = '#,##,##0.00';
      if (typeof cell.v === 'number') {
        cell.t = 'n';
      } else if (cell.f) {
        cell.t = 'n';
      }
    }
  }

  // Column widths for Monthly Breakup
  ws1['!cols'] = [
    { wch: 15 }, // Month (A)
    { wch: 18 }, // Taxable Value (B)
    { wch: 16 }, // IGST (C)
    { wch: 16 }, // CGST (D)
    { wch: 16 }, // SGST (E)
    { wch: 16 }, // Cess (F)
    { wch: 18 }, // Total Tax (G)
    { wch: 20 }, // Gross Total (H)
  ];
  
  // Row heights for Monthly Breakup
  ws1['!rows'] = [];
  ws1['!rows'][0] = { hpt: 26 }; // Title row
  ws1['!rows'][1] = { hpt: 18 };
  ws1['!rows'][2] = { hpt: 18 };
  ws1['!rows'][3] = { hpt: 18 };
  ws1['!rows'][4] = { hpt: 10 }; // blank spacer
  ws1['!rows'][5] = { hpt: 22 }; // Table header
  for (let i = 6; i < 6 + months.length; i++) {
    ws1['!rows'][i] = { hpt: 18 };
  }
  ws1['!rows'][6 + months.length] = { hpt: 22 }; // TOTAL row
  
  XLSX.utils.book_append_sheet(wb,ws1,'Monthly Breakup');
  
  // Source Summary sheet
  const aoa2=[['Source Summary'],[],
    ['File','Register','Client','Rows','Taxable','CGST','SGST','IGST','Cess','Active?'],
    ...sheets.map(s=>[s.fileName,s.typeLabel,s.clientName,s.rowCount,
      s.totals.taxable,s.totals.cgst,s.totals.sgst,s.totals.igst,s.totals.cess,
      s.enabled?'YES':'NO'])];
      
  const ws2=XLSX.utils.aoa_to_sheet(aoa2);
  
  // Cell formatting for Source Summary
  for (const key in ws2) {
    if (key[0] === '!') continue;
    const cell = ws2[key];
    if (!cell) continue;
    
    const colChar = key.replace(/[0-9]/g, '');
    const rowNum = parseInt(key.replace(/[^0-9]/g, ''), 10);
    
    if (['E', 'F', 'G', 'H', 'I'].includes(colChar) && rowNum >= 4) {
      cell.z = '#,##,##0.00';
      if (typeof cell.v === 'number') cell.t = 'n';
    }
    if (colChar === 'D' && rowNum >= 4) {
      cell.z = '#,##0';
      if (typeof cell.v === 'number') cell.t = 'n';
    }
  }

  // Column widths for Source Summary
  ws2['!cols'] = [
    { wch: 30 }, // File
    { wch: 22 }, // Register
    { wch: 22 }, // Client
    { wch: 10 }, // Rows
    { wch: 16 }, // Taxable
    { wch: 14 }, // CGST
    { wch: 14 }, // SGST
    { wch: 14 }, // IGST
    { wch: 14 }, // Cess
    { wch: 10 }, // Active?
  ];

  // Row heights for Source Summary
  ws2['!rows'] = [
    { hpt: 24 }, // Title
    { hpt: 10 }, // spacer
    { hpt: 20 }, // Header
  ];
  for (let i = 3; i < 3 + sheets.length; i++) {
    ws2['!rows'][i] = { hpt: 18 };
  }
  
  XLSX.utils.book_append_sheet(wb,ws2,'Source Summary');
  
  let fyStr = 'FY';
  if (months && months.length > 0) {
    const startYr = months[0].split('-')[1];
    const endYr = months[months.length-1].split('-')[1];
    fyStr = startYr === endYr ? `FY20${startYr}` : `FY20${startYr}-${endYr}`;
  }
  const safeClientName = clients ? clients.replace(/[\\/:*?"<>|]/g, '').trim() : 'UnknownClient';
  const fileName = `GST_Sales_Monthly Summary _${safeClientName}_${fyStr}.xlsx`;
  
  XLSX.writeFile(wb, fileName);
}

// ── Design ─────────────────────────────────────────────────────────────────────
const C={bg:'#06111E',surface:'#0C1A2B',card:'#0F2035',border:'#163050',
  accent:'#00C896',accentLo:'#00C89614',text:'#D8EAF8',muted:'#5B7A96',dim:'#213040',
  red:'#FF6B6B',warn:'#FBBF24',cgst:'#60A5FA',sgst:'#A78BFA',igst:'#FBBF24',
  cess:'#FB923C',taxable:'#34D399',total:'#94A3B8',gross:'#38BDF8'};

const TYPE_STYLE={
  SALES:      {bg:'#041A0A',text:'#4ADE80',bar:'#00C896',label:'Sales Register'},
  CREDIT_NOTE:{bg:'#1A0404',text:'#FCA5A5',bar:'#FF6B6B',label:'Credit Note Register'},
  JOURNAL:    {bg:'#0A0F1A',text:'#93C5FD',bar:'#60A5FA',label:'Journal Register'},
};
const COLS=[
  {key:'taxable',   label:'Taxable Value', color:C.taxable},
  {key:'igst',      label:'IGST',          color:C.igst},
  {key:'cgst',      label:'CGST',          color:C.cgst},
  {key:'sgst',      label:'SGST',          color:C.sgst},
  {key:'cess',      label:'Cess',          color:C.cess},
  {key:'totalTax',  label:'Total Tax',     color:C.total},
  {key:'grossTotal',label:'Gross Total',   color:C.gross},
];

function FileCard({sheet,onToggle,onOpenMapping}) {
  const ts=TYPE_STYLE[sheet.regType]||TYPE_STYLE.SALES;
  const dim=!sheet.enabled;
  const ji=sheet.journalInfo;
  const warnings = getUnclassifiedDataColumns(sheet);
  const hasWarnings = warnings.length > 0;
  
  return (
    <div style={{background:dim?'#060F1B':C.card,border:`1px solid ${dim?'#0F1E30':C.border}`,
      borderRadius:10,padding:'12px 14px 12px 16px',opacity:dim?0.45:1,
      transition:'all .18s',position:'relative',overflow:'hidden'}}>
      {!dim&&<div style={{position:'absolute',left:0,top:0,bottom:0,width:3,
        background:ts.bar,borderRadius:'10px 0 0 10px'}}/>}
      {sheet.isDuplicate&&(
        <div style={{background:'#2A1800',borderRadius:5,padding:'3px 8px',
          fontSize:10,color:'#FBBF24',marginBottom:6}}>
          ⚠ Duplicate — auto-disabled (same register already loaded)
        </div>
      )}
      {hasWarnings && sheet.enabled && (
        <div onClick={onOpenMapping} style={{
          background: '#2A1800', border: '1px solid #92400E', borderRadius: 6,
          padding: '5px 10px', fontSize: 11, color: '#FBBF24', marginBottom: 8,
          cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6,
          transition: 'all 0.15s'
        }}
        title="Click to view column mappings and assign these columns.">
          <span>⚠️</span>
          <span style={{flex: 1}}>
            {warnings.length} unclassified column{warnings.length > 1 ? 's' : ''} ignored!
          </span>
          <span style={{fontSize: 9, textDecoration: 'underline'}}>Fix</span>
        </div>
      )}
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:8}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:10,color:C.muted,marginBottom:2,
            overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
            📄 {sheet.fileName}
          </div>
          <div style={{display:'flex',gap:5,flexWrap:'wrap',alignItems:'center'}}>
            <span style={{fontSize:11,fontWeight:700,padding:'2px 10px',
              borderRadius:20,background:ts.bg,color:ts.text}}>{ts.label}</span>
            <span style={{fontSize:10,padding:'2px 8px',borderRadius:20,
              background:'#0A1928',color:C.muted}}>{sheet.clientName}</span>
          </div>
        </div>
        <div style={{display:'flex',alignItems:'center',gap:8}}>
          {sheet.enabled && (
            <button onClick={onOpenMapping} style={{
              background: '#0B2135', border: `1px solid ${C.border}`,
              borderRadius: 6, padding: '3px 8px', fontSize: 10, color: C.text,
              cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4
            }}>
              ⚙️ Columns
            </button>
          )}
          <label style={{cursor:'pointer'}}>
            <input type="checkbox" checked={sheet.enabled} onChange={onToggle}
              style={{width:15,height:15,accentColor:C.accent,cursor:'pointer'}}/>
          </label>
        </div>
      </div>

      {/* Journal sign breakdown - Aggregated */}
      {ji&&sheet.enabled&&(
        <div style={{background:'#080F1C',borderRadius:6,padding:'8px 12px',marginBottom:8,fontSize:11,display:'flex',flexDirection:'column',gap:4}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
            <span style={{color:C.muted}}>Total Cr (Addition):</span>
            <span style={{color:'#4ADE80',fontWeight:700}}>
              +{ji.totalCrTaxable.toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2})}
            </span>
          </div>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
            <span style={{color:C.muted}}>Total Dr (Reduction):</span>
            <span style={{color:'#FF6B6B',fontWeight:700}}>
              −{ji.totalDrTaxable.toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2})}
            </span>
          </div>
        </div>
      )}

      <div style={{fontSize:11,color:C.muted,marginBottom:8}}>
        {sheet.rowCount.toLocaleString('en-IN')} GST rows
        {sheet.taxCols.length>0&&<span style={{marginLeft:8,color:C.dim}}>
          · {sheet.taxCols.slice(0,3).join(', ')}
        </span>}
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'3px 10px'}}>
        {[['Taxable','taxable',C.taxable],['CGST','cgst',C.cgst],
          ['SGST','sgst',C.sgst],['IGST','igst',C.igst],['Cess','cess',C.cess]].map(([lbl,key,color])=>
          Math.abs(sheet.totals[key]||0)>0.005?(
            <div key={key} style={{fontSize:11,color:C.muted}}>
              {lbl}: <span style={{color,fontWeight:600}}>{fmtINR(sheet.totals[key])}</span>
            </div>
          ):null
        )}
      </div>
    </div>
  );
}

function ColumnMappingModal({ sheet, onClose, onUpdateMapping, onResetMapping }) {
  const [dragOverBin, setDragOverBin] = useState(null);
  const [search, setSearch] = useState('');
  
  const warnings = getUnclassifiedDataColumns(sheet);
  
  const bins = [
    { role: 'SALE_LEDGER', type: null, label: 'Taxable (Sales Ledger)', bg: '#062615', border: '#105B35', text: '#4ADE80', desc: 'Adds to Taxable Value' },
    { role: 'TAXABLE_VALUE', type: null, label: 'Taxable (Value Column)', bg: '#032329', border: '#0A5561', text: '#22D3EE', desc: 'Pre-computed Value' },
    { role: 'TAX', type: 'cgst', label: 'CGST', bg: '#0B1E36', border: '#1C3E69', text: '#60A5FA', desc: 'Central GST' },
    { role: 'TAX', type: 'sgst', label: 'SGST / UTGST', bg: '#1D133D', border: '#3C297B', text: '#A78BFA', desc: 'State/UT GST' },
    { role: 'TAX', type: 'igst', label: 'IGST', bg: '#29200B', border: '#5C481A', text: '#FBBF24', desc: 'Integrated GST' },
    { role: 'TAX', type: 'cess', label: 'Cess', bg: '#2A1406', border: '#5E3113', text: '#FB923C', desc: 'Compensation Cess' },
    { role: 'SKIP', type: null, label: 'Skip / Structural', bg: '#111A24', border: '#213143', text: '#94A3B8', desc: 'Ignored completely' },
  ];

  const getColSampleString = (colIdx) => {
    const vals = [];
    for (const r of sheet.dataRows) {
      const v = Number(r[colIdx]);
      if (!isNaN(v) && Math.abs(v) > 0.005) {
        vals.push(v.toLocaleString('en-IN', { maximumFractionDigits: 2 }));
        if (vals.length >= 3) break;
      }
    }
    return vals.length > 0 ? vals.join(', ') : '—';
  };

  const getRoleLabel = (cls) => {
    if (!cls) return 'Unclassified (Ignored)';
    if (cls.role === 'SALE_LEDGER') return 'Taxable (Sales Ledger)';
    if (cls.role === 'TAXABLE_VALUE') return 'Taxable (Value Column)';
    if (cls.role === 'SKIP') return 'Skip / Structural';
    if (cls.role === 'TAX') {
      const typeLabel = cls.type === 'cgst' ? 'CGST' : cls.type === 'sgst' ? 'SGST' : cls.type === 'igst' ? 'IGST' : 'Cess';
      return `${typeLabel}${cls.rate ? ` @ ${(cls.rate * 100).toFixed(1)}%` : ''}`;
    }
    return 'Unknown';
  };

  const filteredCols = sheet.colMap.filter(c => 
    c.h.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div style={{
      position: 'fixed', left: 0, top: 0, right: 0, bottom: 0,
      background: 'rgba(3, 9, 17, 0.85)', backdropFilter: 'blur(10px)',
      display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000,
      padding: 16, animation: 'fadeIn 0.2s ease'
    }}>
      <style>{`
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes slideUp { from { transform: translateY(20px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
        .bin-hover { border-color: #00C896 !important; box-shadow: 0 0 10px rgba(0, 200, 150, 0.25) !important; }
      `}</style>
      
      <div style={{
        background: '#091523', border: '1px solid #163050', borderRadius: 14,
        width: '100%', maxWidth: 960, maxHeight: '90vh', display: 'flex', flexDirection: 'column',
        boxShadow: '0 20px 40px rgba(0,0,0,0.6)', animation: 'slideUp 0.25s cubic-bezier(0.16, 1, 0.3, 1)'
      }}>
        {/* Modal Header */}
        <div style={{
          padding: '16px 20px', borderBottom: '1px solid #163050',
          display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 10
        }}>
          <div>
            <div style={{display: 'flex', gap: 8, alignItems: 'center', marginBottom: 2}}>
              <h2 style={{margin: 0, fontSize: 16, fontWeight: 800, color: '#D8EAF8'}}>
                ⚙️ Column Mapping &amp; Audit Trail
              </h2>
              <span style={{
                fontSize: 10, fontWeight: 700, padding: '2px 8px', borderRadius: 12,
                background: TYPE_STYLE[sheet.regType]?.bg || '#0F2035',
                color: TYPE_STYLE[sheet.regType]?.text || '#D8EAF8'
              }}>
                {sheet.typeLabel}
              </span>
            </div>
            <p style={{margin: 0, fontSize: 11, color: '#5B7A96'}}>
              {sheet.clientName} · {sheet.fileName} ({sheet.sheetName})
            </p>
          </div>
          <div style={{display: 'flex', gap: 8}}>
            <button onClick={() => onResetMapping(sheet.id)} style={{
              background: '#0B2135', border: '1px solid #163050', borderRadius: 8,
              padding: '6px 12px', fontSize: 12, color: '#5B7A96', cursor: 'pointer', fontWeight: 600
            }}>
              Reset to Auto-Detect
            </button>
            <button onClick={onClose} style={{
              background: '#00C896', color: '#000', border: 'none', borderRadius: 8,
              padding: '6px 16px', fontSize: 12, fontWeight: 700, cursor: 'pointer'
            }}>
              Close &amp; Recalculate
            </button>
          </div>
        </div>

        {/* Warnings Banner */}
        {warnings.length > 0 && (
          <div style={{
            background: '#2A1800', borderBottom: '1px solid #92400E',
            padding: '10px 20px', fontSize: 12, color: '#FBBF24',
            display: 'flex', alignItems: 'center', gap: 8
          }}>
            <span>⚠️</span>
            <span style={{flex: 1}}>
              <strong>{warnings.length} unclassified column(s) contain numerical data but are ignored:</strong>{' '}
              {warnings.map(w => `'${w.h}' (sum: ₹${w.sum.toLocaleString('en-IN')})`).join(', ')}.
              Please drag them to Taxable or GST bins, or map them using the dropdowns below.
            </span>
          </div>
        )}

        {/* Modal Scroll Content */}
        <div style={{ overflowY: 'auto', padding: 20, flex: 1 }}>
          {/* Drop Bins (Drag-and-Drop Zones) */}
          <div style={{ marginBottom: 20 }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: '#5B7A96', textTransform: 'uppercase', marginBottom: 8, letterSpacing: '.05em' }}>
              Drag columns here to classify them
            </div>
            <div style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))',
              gap: 8
            }}>
              {bins.map(bin => {
                const isOver = dragOverBin === `${bin.role}::${bin.type}`;
                return (
                  <div
                    key={`${bin.role}::${bin.type}`}
                    onDragOver={(e) => {
                      e.preventDefault();
                      setDragOverBin(`${bin.role}::${bin.type}`);
                    }}
                    onDragLeave={() => setDragOverBin(null)}
                    onDrop={(e) => {
                      e.preventDefault();
                      setDragOverBin(null);
                      const colIdx = parseInt(e.dataTransfer.getData("colIdx"));
                      if (!isNaN(colIdx)) {
                        onUpdateMapping(sheet.id, colIdx, bin.role, bin.type);
                      }
                    }}
                    style={{
                      background: bin.bg,
                      border: `1px dashed ${bin.border}`,
                      borderRadius: 8,
                      padding: '10px 12px',
                      display: 'flex',
                      flexDirection: 'column',
                      transition: 'all 0.15s',
                      transform: isOver ? 'scale(1.02)' : 'none',
                      borderColor: isOver ? '#00C896' : bin.border,
                      boxShadow: isOver ? '0 0 10px rgba(0,200,150,0.2)' : 'none'
                    }}
                  >
                    <span style={{ fontSize: 12, fontWeight: 700, color: bin.text, marginBottom: 2 }}>{bin.label}</span>
                    <span style={{ fontSize: 10, color: '#5B7A96' }}>{bin.desc}</span>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Columns Table/List */}
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
              <span style={{ fontSize: 11, fontWeight: 700, color: '#5B7A96', textTransform: 'uppercase', letterSpacing: '.05em' }}>
                Column Audit &amp; Configuration
              </span>
              <input
                type="text"
                placeholder="🔍 Search columns..."
                value={search}
                onChange={e => setSearch(e.target.value)}
                style={{
                  background: '#07101C', border: '1px solid #163050', borderRadius: 6,
                  padding: '4px 10px', color: '#D8EAF8', fontSize: 12, outline: 'none', width: 180
                }}
              />
            </div>
            
            <div style={{ border: '1px solid #163050', borderRadius: 8, overflow: 'hidden' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                <thead>
                  <tr style={{ background: '#050D18', borderBottom: '1px solid #163050' }}>
                    <th style={{ padding: '8px 12px', textAlign: 'left', color: '#5B7A96', width: 40 }}>Col</th>
                    <th style={{ padding: '8px 12px', textAlign: 'left', color: '#5B7A96' }}>Original Header</th>
                    <th style={{ padding: '8px 12px', textAlign: 'left', color: '#5B7A96', width: 220 }}>Current Classification</th>
                    <th style={{ padding: '8px 12px', textAlign: 'left', color: '#5B7A96' }}>Sample Values</th>
                    <th style={{ padding: '8px 12px', textAlign: 'right', color: '#5B7A96', width: 160 }}>Manual Re-assign</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredCols.map((col, index) => {
                    const isUnclassified = !col.cls;
                    const hasValues = warnings.some(w => w.i === col.i);
                    const rowBg = hasValues && isUnclassified ? 'rgba(42, 24, 0, 0.15)' : (index % 2 === 0 ? 'transparent' : '#081220');
                    
                    return (
                      <tr
                        key={col.i}
                        draggable="true"
                        onDragStart={(e) => {
                          e.dataTransfer.setData("colIdx", col.i);
                        }}
                        style={{
                          background: rowBg,
                          borderBottom: '1px solid #10243C',
                          cursor: 'grab',
                          transition: 'background 0.15s'
                        }}
                      >
                        <td style={{ padding: '10px 12px', color: '#5B7A96', fontWeight: 600 }}>{String.fromCharCode(65 + col.i)}</td>
                        <td style={{ padding: '10px 12px', fontWeight: 700, color: '#D8EAF8' }}>
                          <span style={{ marginRight: 6 }}>☰</span> {col.h}
                        </td>
                        <td style={{ padding: '10px 12px' }}>
                          <span style={{
                            padding: '3px 8px', borderRadius: 4, fontSize: 10, fontWeight: 600,
                            background: !col.cls ? '#111A24' : (col.cls.role === 'SKIP' ? '#1A232E' : (col.cls.role === 'TAX' ? (col.cls.type === 'cgst' ? '#0A2540' : col.cls.type === 'sgst' ? '#241040' : col.cls.type === 'igst' ? '#3B2D05' : '#3D1B06') : '#083B20')),
                            color: !col.cls ? '#5B7A96' : (col.cls.role === 'SKIP' ? '#94A3B8' : (col.cls.role === 'TAX' ? (col.cls.type === 'cgst' ? '#60A5FA' : col.cls.type === 'sgst' ? '#A78BFA' : col.cls.type === 'igst' ? '#FBBF24' : '#FB923C') : '#4ADE80')),
                            border: `1px solid ${!col.cls ? '#213143' : (col.cls.role === 'SKIP' ? '#2E3D4F' : (col.cls.role === 'TAX' ? (col.cls.type === 'cgst' ? '#1E3E69' : col.cls.type === 'sgst' ? '#3E2A7A' : col.cls.type === 'igst' ? '#5E4A1B' : '#5F3014') : '#105B35'))}`
                          }}>
                            {getRoleLabel(col.cls)}
                          </span>
                          {hasValues && isUnclassified && (
                            <span style={{ marginLeft: 6, color: '#FBBF24', fontSize: 10 }}>⚠️ Warning: has values</span>
                          )}
                        </td>
                        <td style={{ padding: '10px 12px', color: '#5B7A96', fontVariantNumeric: 'tabular-nums' }}>
                          {getColSampleString(col.i)}
                        </td>
                        <td style={{ padding: '10px 12px', textAlign: 'right' }}>
                          <select
                            value={col.cls ? (col.cls.role === 'TAX' ? `TAX::${col.cls.type}` : col.cls.role) : 'UNCLASSIFIED'}
                            onChange={(e) => {
                              const val = e.target.value;
                              if (val === 'UNCLASSIFIED') {
                                onUpdateMapping(sheet.id, col.i, 'UNCLASSIFIED');
                              } else if (val.startsWith('TAX::')) {
                                const type = val.split('::')[1];
                                onUpdateMapping(sheet.id, col.i, 'TAX', type);
                              } else {
                                onUpdateMapping(sheet.id, col.i, val);
                              }
                            }}
                            style={{
                              background: '#07101C', border: '1px solid #163050', borderRadius: 6,
                              padding: '4px 8px', color: '#D8EAF8', fontSize: 11, outline: 'none', cursor: 'pointer'
                            }}
                          >
                            <option value="UNCLASSIFIED">Ignored (Unclassified)</option>
                            <option value="SALE_LEDGER">Taxable (Sales Ledger)</option>
                            <option value="TAXABLE_VALUE">Taxable (Value Column)</option>
                            <option value="TAX::cgst">CGST</option>
                            <option value="TAX::sgst">SGST / UTGST</option>
                            <option value="TAX::igst">IGST</option>
                            <option value="TAX::cess">Cess</option>
                            <option value="SKIP">Skip / Structural</option>
                          </select>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function RegistersSummary({ totals }) {
  const items = [
    { label: 'Total Taxable Value', value: totals.taxable, color: C.taxable, icon: '📈' },
    { label: 'Consolidated IGST', value: totals.igst, color: C.igst, icon: '🌍' },
    { label: 'Consolidated CGST', value: totals.cgst, color: C.cgst, icon: '🏢' },
    { label: 'Consolidated SGST', value: totals.sgst, color: C.sgst, icon: '🏛️' },
    { label: 'Consolidated Cess', value: totals.cess, color: C.cess, icon: '📦' },
    { label: 'Total Output Tax & Gross Supply', value: totals.grossTotal, tax: totals.totalTax, color: C.gross, icon: '📊' },
  ];

  return (
    <div style={{
      background: C.card, border: `1px solid ${C.border}`, borderRadius: 12,
      padding: '16px 20px', marginBottom: 14
    }}>
      <h3 style={{ margin: '0 0 12px', fontSize: 14, fontWeight: 700, color: C.text }}>
        💼 Consolidated Books Summary (Active Registers)
      </h3>
      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))',
        gap: 12
      }}>
        {items.map((item, idx) => {
          const isZero = Math.abs(item.value) < 0.005;
          const isLast = idx === items.length - 1;
          
          return (
            <div key={idx} style={{
              background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10,
              padding: '12px 16px', display: 'flex', alignItems: 'center', justifyContent: 'space-between',
              opacity: isZero ? 0.5 : 1, transition: 'all 0.18s'
            }}>
              <div>
                <div style={{ fontSize: 10, fontWeight: 600, color: C.muted, textTransform: 'uppercase', letterSpacing: '.03em', marginBottom: 4 }}>
                  {item.label}
                </div>
                <div style={{ fontSize: 18, fontWeight: 800, color: item.color, fontVariantNumeric: 'tabular-nums' }}>
                  {fmtINR(item.value)}
                </div>
                {isLast && (
                  <div style={{ fontSize: 10, color: C.muted, marginTop: 4 }}>
                    Tax: <span style={{ color: C.accent }}>{fmtINR(item.tax)}</span> · Gross: <span style={{ color: C.gross }}>{fmtINR(item.value)}</span>
                  </div>
                )}
              </div>
              <div style={{ fontSize: 24, padding: 8, background: '#07101C', borderRadius: 8 }}>
                {item.icon}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function PLReconcile({totalTaxable, plSales, onPLChange}) {
  const diff   = plSales ? plSales - totalTaxable : null;
  const matched = diff !== null && Math.abs(diff) < 1;
  return (
    <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,
      padding:'16px 20px',marginBottom:14}}>
      <h3 style={{margin:'0 0 10px',fontSize:14,fontWeight:700,color:C.text}}>
        🔍 P&amp;L Cross-Check
      </h3>
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(220px,1fr))',gap:12}}>
        <div>
          <label style={{display:'block',fontSize:11,fontWeight:600,color:C.accent,marginBottom:6}}>
            Sales as per P&amp;L / ITR (₹)
          </label>
          <p style={{fontSize:10,color:C.muted,margin:'0 0 6px',lineHeight:1.5}}>
            Enter total sales from the Profit &amp; Loss account to verify against
            calculated taxable value.
          </p>
          <input type="number" value={plSales||''} placeholder="Enter P&L sales total"
            onChange={e=>onPLChange(parseFloat(e.target.value)||0)}
            style={{width:'100%',padding:'8px 12px',background:'#0A1928',
              border:`1px solid ${C.border}`,borderRadius:7,color:C.text,
              fontSize:13,outline:'none',boxSizing:'border-box'}}/>
        </div>
        {plSales>0&&(
          <div style={{background:matched?'#041A0A':'#1A0A04',
            border:`1px solid ${matched?'#166534':'#92400E'}`,
            borderRadius:8,padding:'12px 14px',alignSelf:'start'}}>
            <div style={{fontSize:11,fontWeight:700,marginBottom:8,
              color:matched?'#4ADE80':'#FBBF24'}}>
              {matched?'✓ Match':'⚠ Difference Found'}
            </div>
            {[['P&L Sales',plSales],['Calculated Taxable',totalTaxable],['Difference',diff]].map(([lbl,val])=>(
              <div key={lbl} style={{display:'flex',justifyContent:'space-between',
                fontSize:12,color:C.text,marginBottom:3}}>
                <span style={{color:C.muted}}>{lbl}</span>
                <span style={{fontWeight:600,
                  color:lbl==='Difference'?(Math.abs(diff||0)<1?C.accent:C.red):C.text}}>
                  {fmtINR(val)}
                </span>
              </div>
            ))}
            {!matched&&diff!==null&&(
              <div style={{marginTop:6,fontSize:10,color:'#FBBF24',lineHeight:1.5}}>
                Check: exempt supplies, non-GST income, or journal sign detection.
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

function MonthlyTable({monthly,months,sheets,plSales,onExport}) {
  const hasCess=months.some(m=>Math.abs((monthly[m]||{}).cess||0)>0.005);
  const hasIGST=months.some(m=>Math.abs((monthly[m]||{}).igst||0)>0.005);
  const visCols=COLS.filter(c=>{
    if(c.key==='cess'&&!hasCess)return false;
    if(c.key==='igst'&&!hasIGST)return false;
    return true;
  });
  const totals=months.reduce((acc,m)=>{
    const d=monthly[m]||{};
    visCols.forEach(c=>{acc[c.key]=(acc[c.key]||0)+(d[c.key]||0)});
    return acc;
  },{});
  return (
    <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,overflow:'hidden'}}>
      <div style={{padding:'14px 20px',background:C.surface,borderBottom:`1px solid ${C.border}`,
        display:'flex',justifyContent:'space-between',alignItems:'center',flexWrap:'wrap',gap:8}}>
        <div>
          <h2 style={{margin:0,fontSize:15,fontWeight:700,color:C.text}}>
            📊 Monthly Output Tax Breakup — Books
          </h2>
          <p style={{margin:'3px 0 0',fontSize:11,color:C.muted}}>
            {months.length} months · Journal Dr/Cr auto-detected · Credit Notes negative
          </p>
        </div>
        <button onClick={onExport} style={{background:C.accent,color:'#000',border:'none',
          borderRadius:8,padding:'8px 18px',cursor:'pointer',fontSize:13,fontWeight:700}}>
          ↓ Export to Excel
        </button>
      </div>
      <div style={{overflowX:'auto'}}>
        <table style={{width:'100%',borderCollapse:'collapse',minWidth:600}}>
          <thead>
            <tr style={{background:'#050D18'}}>
              <th style={{padding:'10px 16px',textAlign:'left',fontSize:11,fontWeight:700,
                color:C.muted,borderBottom:`1px solid ${C.border}`,
                letterSpacing:'.06em',textTransform:'uppercase'}}>Month</th>
              {visCols.map(c=>(
                <th key={c.key} style={{padding:'10px 14px',textAlign:'right',fontSize:11,
                  fontWeight:700,color:c.color,borderBottom:`1px solid ${C.border}`,
                  letterSpacing:'.04em',textTransform:'uppercase',whiteSpace:'nowrap'}}>
                  {c.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {months.map((m,i)=>{
              const d=monthly[m]||{};
              return (
                <tr key={m} style={{background:i%2===0?'transparent':'#081524'}}>
                  <td style={{padding:'9px 16px',fontWeight:600,fontSize:13,color:C.text}}>{m}</td>
                  {visCols.map(c=>{
                    const v=d[c.key]||0;
                    return (
                      <td key={c.key} style={{padding:'9px 14px',textAlign:'right',fontSize:13,
                        color:Math.abs(v)>0.005?(v<0?C.red:C.text):C.dim,
                        fontVariantNumeric:'tabular-nums'}}>
                        {fmtINR(v)}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr style={{borderTop:`2px solid ${C.accent}`,background:'#040C18'}}>
              <td style={{padding:'11px 16px',fontWeight:700,fontSize:13,color:C.accent,
                letterSpacing:'.06em'}}>TOTAL</td>
              {visCols.map(c=>(
                <td key={c.key} style={{padding:'11px 14px',textAlign:'right',
                  fontWeight:700,fontSize:13,color:C.accent,fontVariantNumeric:'tabular-nums'}}>
                  {fmtINR(totals[c.key]||0)}
                </td>
              ))}
            </tr>
          </tfoot>
        </table>
      </div>
      <div style={{padding:'10px 20px',background:C.surface,borderTop:`1px solid ${C.border}`,
        fontSize:11,color:C.muted,display:'flex',justifyContent:'space-between',flexWrap:'wrap',gap:6}}>
        <span>
          ⚡ Freight without GST excluded · INPUT CGST/SGST in Credit Notes included ·
          Journal Dr/Cr detected from Grand Total constraint
        </span>
        <span style={{color:C.dim}}>Step 2 → Upload GSTR portal data for reconciliation</span>
      </div>
    </div>
  );
}

export default function App({ onBack }) {
  const [sheets,   setSheets]  = useState([]);
  const [monthly,  setMonthly] = useState({});
  const [months,   setMonths]  = useState([]);
  const [processing,setProc]  = useState(false);
  const [isDrag,   setIsDrag]  = useState(false);
  const [plSales,  setPlSales] = useState(0);
  const [activeModalSheetId, setActiveModalSheetId] = useState(null);
  const fileRef = useRef();

  const recompute = useCallback((list)=>{
    const {monthly:m,months:mo}=computeMonthly(list);
    setMonthly(m); setMonths(mo);
  },[]);

  const handleFiles = useCallback(async(fileList)=>{
    if(!fileList?.length)return;
    setProc(true);
    const newSheets=[];
    for (const f of Array.from(fileList)) {
      if(!/\.xlsx?$/i.test(f.name))continue;
      try {
        const buf=await f.arrayBuffer();
        const wb=XLSX.read(buf,{type:'array',cellDates:false});
        for (const sn of wb.SheetNames) {
          const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:true});
          const parsed=parseRegister(rows,f.name,sn);
          if(parsed)newSheets.push(parsed);
        }
      } catch(e){console.error('Parse error',f.name,e);}
    }
    const seenKeys = new Set();
    const deduped = newSheets.map(s => {
      if (seenKeys.has(s.dedupKey)) return {...s, enabled:false, isDuplicate:true};
      seenKeys.add(s.dedupKey);
      return s;
    });
    setSheets(deduped); recompute(deduped); setProc(false);
  },[sheets,recompute]);

  const toggle=useCallback((id)=>{
    const next=sheets.map(s=>s.id===id?{...s,enabled:!s.enabled}:s);
    setSheets(next); recompute(next);
  },[sheets,recompute]);

  const updateColumnMapping = useCallback((sheetId, colIndex, newRole, type = null) => {
    const next = sheets.map(s => {
      if (s.id !== sheetId) return s;
      
      const newColMap = s.colMap.map(c => {
        if (c.i !== colIndex) return c;
        
        let cls = null;
        if (newRole === 'SALE_LEDGER') {
          cls = { role: 'SALE_LEDGER', type: 'taxable', rate: extractRate(c.h) };
        } else if (newRole === 'TAXABLE_VALUE') {
          cls = { role: 'TAXABLE_VALUE' };
        } else if (newRole === 'TAX') {
          cls = { role: 'TAX', type: type, rate: extractRate(c.h) };
        } else if (newRole === 'SKIP') {
          cls = { role: 'SKIP' };
        }
        return { ...c, cls };
      });

      const calcs = computeSheetCalculations(s.regType, s.sign, newColMap, s.dataRows, s.grandTotalRow, s.fileName, s.clientName);
      
      return {
        ...s,
        colMap: newColMap,
        ...calcs
      };
    });
    
    setSheets(next);
    recompute(next);
  }, [sheets, recompute]);

  const resetColumnMapping = useCallback((sheetId) => {
    const next = sheets.map(s => {
      if (s.id !== sheetId) return s;
      
      const originalColMap = JSON.parse(JSON.stringify(s.originalColMap));
      const calcs = computeSheetCalculations(s.regType, s.sign, originalColMap, s.dataRows, s.grandTotalRow, s.fileName, s.clientName);
      
      return {
        ...s,
        colMap: originalColMap,
        ...calcs
      };
    });
    
    setSheets(next);
    recompute(next);
  }, [sheets, recompute]);

  const totalTaxable = months.reduce((a,m)=>a+(monthly[m]?.taxable||0),0);
  const totalCgst = months.reduce((a,m)=>a+(monthly[m]?.cgst||0),0);
  const totalSgst = months.reduce((a,m)=>a+(monthly[m]?.sgst||0),0);
  const totalIgst = months.reduce((a,m)=>a+(monthly[m]?.igst||0),0);
  const totalCess = months.reduce((a,m)=>a+(monthly[m]?.cess||0),0);
  const totalTax  = totalCgst + totalSgst + totalIgst + totalCess;
  const grossSupply = totalTaxable + totalTax;

  const clients = [...new Set(sheets.map(s=>s.clientName).filter(Boolean))];
  const activeModalSheet = sheets.find(s => s.id === activeModalSheetId);

  return (
    <div style={{fontFamily:"'DM Sans','Trebuchet MS',sans-serif",
      background:C.bg,minHeight:'100vh',color:C.text,padding:16,boxSizing:'border-box'}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');
        *,*::before,*::after{box-sizing:border-box}
        tr:hover>td{background:rgba(0,200,150,.03)!important}
        button:hover{opacity:.85;transform:translateY(-1px)} button{transition:all .15s}
        input[type=number]:focus{border-color:#00C896!important}
      `}</style>

      <div style={{background:C.surface,border:`1px solid ${C.border}`,
        borderLeft:`4px solid ${C.accent}`,borderRadius:12,
        padding:'13px 20px',marginBottom:14,
        display:'flex',justifyContent:'space-between',alignItems:'center',flexWrap:'wrap',gap:8}}>
        <div style={{display:'flex', alignItems:'center', gap:'12px'}}>
          {onBack && (
            <button onClick={onBack} style={{background:'#0F2035', border:`1px solid ${C.border}`, color:C.text, padding:'4px 10px', borderRadius:'6px', cursor:'pointer', fontSize:'12px'}}>
              ← Back
            </button>
          )}
          <div>
            <h1 style={{margin:0,fontSize:17,fontWeight:800,color:C.accent,letterSpacing:'-.01em'}}>
              GST Output Tax — Books Data Extraction
            </h1>
            <p style={{margin:'3px 0 0',fontSize:11,color:C.muted}}>
              Step 1 of 2 · Subin B &amp; Associates, Chartered Accountants
              {clients.length>0&&` · ${clients.join(', ')}`}
            </p>
          </div>
        </div>
        {sheets.length>0&&(
          <div style={{display:'flex',gap:6,alignItems:'center'}}>
            <span style={{fontSize:11,padding:'3px 10px',borderRadius:20,
              background:C.accentLo,color:C.accent,fontWeight:600}}>
              {sheets.filter(s=>s.enabled).length}/{sheets.length} active
            </span>
            {sheets.some(s=>s.isDuplicate)&&(
              <span style={{fontSize:11,padding:'3px 10px',borderRadius:20,
                background:'#2A1800',color:'#FBBF24',fontWeight:600}}>
                {sheets.filter(s=>s.isDuplicate).length} duplicate(s)
              </span>
            )}
            <button onClick={()=>{setSheets([]);setMonthly({});setMonths([]);setPlSales(0);setActiveModalSheetId(null);}}
              style={{fontSize:11,padding:'3px 10px',borderRadius:20,
                background:'#1A0A0A',color:C.red,border:'none',cursor:'pointer'}}>
              Clear All
            </button>
          </div>
        )}
      </div>

      <div onDrop={e=>{e.preventDefault();setIsDrag(false);handleFiles(e.dataTransfer.files);}}
        onDragOver={e=>{e.preventDefault();setIsDrag(true);}}
        onDragLeave={()=>setIsDrag(false)}
        onClick={()=>fileRef.current?.click()}
        style={{background:isDrag?'#081E2E':C.surface,
          border:`2px dashed ${isDrag?C.accent:C.border}`,
          borderRadius:12,padding:'22px 20px',textAlign:'center',
          cursor:'pointer',marginBottom:14,transition:'all .2s'}}>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" multiple hidden
          onChange={e=>handleFiles(e.target.files)}/>
        <div style={{fontSize:24,marginBottom:5}}>📂</div>
        <div style={{fontSize:14,fontWeight:700,color:isDrag?C.accent:C.text}}>
          {processing?'⏳  Processing…':'Drop Tally Register Exports Here'}
        </div>
        <div style={{fontSize:12,color:C.muted,marginTop:4}}>
          Sales Register · Credit Note Register · Journal Register · .xlsx
        </div>
        <div style={{marginTop:6,fontSize:10,color:C.dim}}>
          Tally: Display → Account Books → [Register] → Set period → Ctrl+E → Excel
        </div>
      </div>

      {sheets.length>0&&(
        <div style={{marginBottom:14}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:10}}>
            <span style={{fontSize:11,fontWeight:700,color:C.muted,
              letterSpacing:'.07em',textTransform:'uppercase'}}>Detected Registers</span>
            <span style={{fontSize:11,color:C.dim}}>Uncheck to exclude · Duplicates auto-disabled</span>
          </div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(260px,1fr))',gap:10}}>
            {sheets.map(s=><FileCard key={s.id} sheet={s} onToggle={()=>toggle(s.id)} onOpenMapping={() => setActiveModalSheetId(s.id)}/>)}
          </div>
        </div>
      )}

      {months.length>0&&(
        <RegistersSummary totals={{
          taxable: totalTaxable,
          cgst: totalCgst,
          sgst: totalSgst,
          igst: totalIgst,
          cess: totalCess,
          totalTax: totalTax,
          grossTotal: grossSupply
        }} />
      )}

      {months.length>0&&(
        <PLReconcile totalTaxable={totalTaxable} plSales={plSales} onPLChange={setPlSales}/>
      )}

      {months.length>0&&(
        <MonthlyTable monthly={monthly} months={months} sheets={sheets}
          plSales={plSales} onExport={()=>doExport(monthly,months,sheets,plSales)}/>
      )}

      {sheets.length===0&&!processing&&(
        <div style={{textAlign:'center',padding:'40px 20px',color:C.muted}}>
          <div style={{fontSize:40,marginBottom:10,opacity:.15}}>📊</div>
          <div style={{fontSize:14,fontWeight:600,marginBottom:6}}>
            Upload Tally Register Exports to Begin
          </div>
          <div style={{fontSize:12,maxWidth:520,margin:'0 auto',lineHeight:1.8,color:C.dim}}>
            Sales Register, Credit Note Register and Journal Register exported from Tally Prime.
            Journal Dr/Cr entries automatically detected and signed correctly.
            Duplicates across files auto-disabled. Client name read from file header.
          </div>
        </div>
      )}

      {activeModalSheet && (
        <ColumnMappingModal
          sheet={activeModalSheet}
          onClose={() => setActiveModalSheetId(null)}
          onUpdateMapping={updateColumnMapping}
          onResetMapping={resetColumnMapping}
        />
      )}
    </div>
  );
}
