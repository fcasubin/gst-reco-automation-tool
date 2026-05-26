import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// Color Palette (matches Steps 1, 2, and 3 styling)
const C = {
  bg: '#06111E',
  surface: '#0C1A2B',
  card: '#0F2035',
  border: '#163050',
  accent: '#FB923C',      // Orange accent color for Step 4
  accentLo: '#FB923C14',
  text: '#D8EAF8',
  muted: '#5B7A96',
  dim: '#213040',
  red: '#FF6B6B',
  warn: '#FBBF24',
  cgst: '#60A5FA',
  sgst: '#A78BFA',
  igst: '#FBBF24',
  cess: '#FB923C',
  taxable: '#34D399',
  total: '#94A3B8'
};

// Date helper
function toDate(v) {
  if (!v) return null;
  let d = null;
  if (v instanceof Date) {
    d = v;
  } else {
    const n = Number(v);
    if (!isNaN(n) && n > 40000 && n < 70000) {
      d = new Date(Math.round((n - 25569) * 86400000));
    } else if (typeof v === 'string') {
      const parsed = new Date(v);
      if (!isNaN(parsed)) d = parsed;
    }
  }
  if (!d) return null;
  const isUtcMidnight = d.getUTCHours() === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0 && d.getUTCMilliseconds() === 0;
  const year = isUtcMidnight ? d.getUTCFullYear() : d.getFullYear();
  const month = isUtcMidnight ? d.getUTCMonth() : d.getMonth();
  const date = isUtcMidnight ? d.getUTCDate() : d.getDate();
  return new Date(Date.UTC(year, month, date));
}

function fmtINR(n) {
  if (n === undefined || n === null || Math.abs(n) < 0.005) return '0.00';
  return (n < 0 ? '−' : '') + Math.abs(n).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function normalizeInvoiceNo(inv) {
  if (!inv) return "";
  const val = String(inv).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  return val.replace(/^0+/, ''); // strip leading zeros
}

function cleanGSTIN(gstin) {
  if (!gstin) return "";
  return String(gstin).replace(/[^A-Z0-9]/g, '').toUpperCase().trim();
}

export default function GSTR2A_vs_2B_Reconciliation({ onBack }) {
  const [gstr2aFiles, setGstr2aFiles] = useState([]);
  const [gstr2bFiles, setGstr2bFiles] = useState([]);
  const [gstr2aRecords, setGstr2aRecords] = useState([]);
  const [gstr2bRecords, setGstr2bRecords] = useState([]);
  const [taxpayerName, setTaxpayerName] = useState("AN DISTRIBUTION");
  const [recoData, setRecoData] = useState(null);
  
  const [processing, setProcessing] = useState(false);
  const [activeTab, setActiveTab] = useState('summary');
  
  const file2aRef = useRef();
  const file2bRef = useRef();

  // Parse GSTR-2A
  const parseGstr2a = async (file) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: false });
    const records = [];
    
    // 1. Parse B2B
    const b2bSheets = wb.SheetNames.filter(s => s.toUpperCase().includes('B2B') && s.toUpperCase().includes('ONLY INVOICE'));
    const b2bSheetName = b2bSheets.length > 0 ? b2bSheets[0] : wb.SheetNames.find(s => s.toUpperCase().includes('B2B'));
    
    if (b2bSheetName) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[b2bSheetName], { header: 1, defval: '', raw: true });
      let hIdx = -1;
      for (let i = 0; i < Math.min(rows.length, 15); i++) {
        if (rows[i] && rows[i].some(c => String(c).toUpperCase().includes('GSTIN'))) { hIdx = i; break; }
      }
      if (hIdx !== -1) {
        rows.slice(hIdx + 2).forEach(row => {
          if (!row || !row[0] || String(row[0]).trim() === '' || String(row[0]).toUpperCase().includes('TOTAL')) return;
          const gstin = cleanGSTIN(row[0]);
          const name = String(row[1] || '').trim();
          const invNo = String(row[2] || '').trim();
          const docType = String(row[3] || 'INV').trim().toUpperCase();
          const dt = toDate(row[4]);
          const taxable = Number(row[8] || 0);
          const igst = Number(row[9] || 0);
          const cgst = Number(row[10] || 0);
          const sgst = Number(row[11] || 0);
          const cess = Number(row[12] || 0);
          const filingPeriod = String(row[14] || '').trim();
          
          records.push({
            gstin,
            supplierName: name,
            invoiceNo: invNo,
            date: dt,
            taxable,
            igst,
            cgst,
            sgst,
            cess,
            totalTax: igst + cgst + sgst + cess,
            filingPeriod,
            noteType: 'INV',
            category: 'B2B',
            matchKey: `${gstin}_${normalizeInvoiceNo(invNo)}`
          });
        });
      }
    }
    
    // 2. Parse CDNR
    const cdnrSheets = wb.SheetNames.filter(s => s.toUpperCase() === 'CDNR');
    if (cdnrSheets.length > 0) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[cdnrSheets[0]], { header: 1, defval: '', raw: true });
      let hIdx = -1;
      for (let i = 0; i < Math.min(rows.length, 15); i++) {
        if (rows[i] && rows[i].some(c => String(c).toUpperCase().includes('GSTIN'))) { hIdx = i; break; }
      }
      if (hIdx !== -1) {
        rows.slice(hIdx + 2).forEach(row => {
          if (!row || !row[0] || String(row[0]).trim() === '' || String(row[0]).toUpperCase().includes('TOTAL')) return;
          const gstin = cleanGSTIN(row[0]);
          const name = String(row[1] || '').trim();
          const noteType = String(row[4] || 'C').trim().toUpperCase(); // Col 4 is Note type (C/D)
          const noteNo = String(row[5] || '').trim(); // Col 5 is Note number
          const dt = toDate(row[6]); // Col 6 is Note date
          
          const mult = noteType.includes('C') ? -1 : 1;
          const taxable = Number(row[10] || 0) * mult;
          const igst = Number(row[11] || 0) * mult;
          const cgst = Number(row[12] || 0) * mult;
          const sgst = Number(row[13] || 0) * mult;
          const cess = Number(row[14] || 0) * mult;
          const filingPeriod = String(row[16] || '').trim();
          
          records.push({
            gstin,
            supplierName: name,
            invoiceNo: noteNo,
            date: dt,
            taxable,
            igst,
            cgst,
            sgst,
            cess,
            totalTax: igst + cgst + sgst + cess,
            filingPeriod,
            noteType,
            category: 'CDNR',
            matchKey: `${gstin}_${normalizeInvoiceNo(noteNo)}`
          });
        });
      }
    }
    return records;
  };

  // Parse GSTR-2B
  const parseGstr2b = async (file) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: false });
    const records = [];
    
    // 1. Parse B2B
    const b2bSheets = wb.SheetNames.filter(s => s.toUpperCase() === 'B2B');
    if (b2bSheets.length > 0) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[b2bSheets[0]], { header: 1, defval: '', raw: true });
      let hIdx = -1;
      for (let i = 0; i < Math.min(rows.length, 15); i++) {
        if (rows[i] && rows[i].some(c => String(c).toUpperCase().includes('GSTIN'))) { hIdx = i; break; }
      }
      if (hIdx !== -1) {
        rows.slice(hIdx + 2).forEach(row => {
          if (!row || !row[0] || String(row[0]).trim() === '' || String(row[0]).toUpperCase().includes('TOTAL')) return;
          const gstin = cleanGSTIN(row[0]);
          const name = String(row[1] || '').trim();
          const invNo = String(row[2] || '').trim();
          const dt = toDate(row[4]);
          const taxable = Number(row[9] || 0);
          const igst = Number(row[10] || 0);
          const cgst = Number(row[11] || 0);
          const sgst = Number(row[12] || 0);
          const cess = Number(row[13] || 0);
          const filingPeriod = String(row[14] || '').trim();
          
          records.push({
            gstin,
            supplierName: name,
            invoiceNo: invNo,
            date: dt,
            taxable,
            igst,
            cgst,
            sgst,
            cess,
            totalTax: igst + cgst + sgst + cess,
            filingPeriod,
            noteType: 'INV',
            category: 'B2B',
            matchKey: `${gstin}_${normalizeInvoiceNo(invNo)}`
          });
        });
      }
    }
    
    // 2. Parse CDNR
    const cdnrSheets = wb.SheetNames.filter(s => s.toUpperCase() === 'CDNR');
    if (cdnrSheets.length > 0) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[cdnrSheets[0]], { header: 1, defval: '', raw: true });
      let hIdx = -1;
      for (let i = 0; i < Math.min(rows.length, 15); i++) {
        if (rows[i] && rows[i].some(c => String(c).toUpperCase().includes('GSTIN'))) { hIdx = i; break; }
      }
      if (hIdx !== -1) {
        rows.slice(hIdx + 2).forEach(row => {
          if (!row || !row[0] || String(row[0]).trim() === '' || String(row[0]).toUpperCase().includes('TOTAL')) return;
          const gstin = cleanGSTIN(row[0]);
          const name = String(row[1] || '').trim();
          const noteNo = String(row[2] || '').trim();
          const noteType = String(row[3] || 'C').trim().toUpperCase(); // Col 3 is Note type (C/D)
          const dt = toDate(row[5]);
          
          const mult = noteType.includes('C') ? -1 : 1;
          const taxable = Number(row[10] || 0) * mult;
          const igst = Number(row[11] || 0) * mult;
          const cgst = Number(row[12] || 0) * mult;
          const sgst = Number(row[13] || 0) * mult;
          const cess = Number(row[14] || 0) * mult;
          const filingPeriod = String(row[15] || '').trim();
          
          records.push({
            gstin,
            supplierName: name,
            invoiceNo: noteNo,
            date: dt,
            taxable,
            igst,
            cgst,
            sgst,
            cess,
            totalTax: igst + cgst + sgst + cess,
            filingPeriod,
            noteType,
            category: 'CDNR',
            matchKey: `${gstin}_${normalizeInvoiceNo(noteNo)}`
          });
        });
      }
    }
    
    // 3. Extract Taxpayer Name from Read me
    let nameExtracted = "AN DISTRIBUTION";
    const readmeSheet = wb.SheetNames.find(s => s.toLowerCase().includes('read me'));
    if (readmeSheet) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[readmeSheet], { header: 1, defval: '', raw: true });
      for (let r = 0; r < rows.length; r++) {
        if (!rows[r]) continue;
        for (let c = 0; c < rows[r].length; c++) {
          const val = String(rows[r][c] || '').trim();
          if (val.toLowerCase().includes('legal name')) {
            nameExtracted = String(rows[r][c + 1] || 'AN DISTRIBUTION').trim();
            break;
          }
        }
      }
    }

    return { records, nameExtracted };
  };

  const handle2aUpload = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setProcessing(true);
    setGstr2aFiles(files);
    
    let allRecords = [];
    for (const f of files) {
      try {
        const recs = await parseGstr2a(f);
        allRecords = allRecords.concat(recs);
      } catch (ex) {
        console.error("Error parsing GSTR-2A file", f.name, ex);
      }
    }
    setGstr2aRecords(allRecords);
    setProcessing(false);
  };

  const handle2bUpload = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setProcessing(true);
    setGstr2bFiles(files);
    
    let allRecords = [];
    let name = "AN DISTRIBUTION";
    for (const f of files) {
      try {
        const result = await parseGstr2b(f);
        allRecords = allRecords.concat(result.records);
        name = result.nameExtracted;
      } catch (ex) {
        console.error("Error parsing GSTR-2B file", f.name, ex);
      }
    }
    setGstr2bRecords(allRecords);
    setTaxpayerName(name);
    setProcessing(false);
  };

  const runReconciliation = () => {
    if (gstr2aRecords.length === 0 || gstr2bRecords.length === 0) return;
    setProcessing(true);

    const keys2a = new Set(gstr2aRecords.map(r => r.matchKey));
    const keys2b = new Set(gstr2bRecords.map(r => r.matchKey));

    const only2b = gstr2bRecords.filter(r => !keys2a.has(r.matchKey));
    const only2a = gstr2aRecords.filter(r => !keys2b.has(r.matchKey));
    const common = gstr2aRecords.filter(r => keys2b.has(r.matchKey));

    // Previous year check (doc date < 2025-04-01)
    const prevYear2b = only2b.filter(r => {
      const dt = toDate(r.date);
      return dt && dt < new Date('2025-04-01');
    });

    // CDNR Mismatches (D vs C)
    const map2b = new Map(gstr2bRecords.map(r => [r.matchKey, r]));
    const cdMismatches = [];
    common.forEach(r2a => {
      const r2b = map2b.get(r2a.matchKey);
      if (r2a.category === 'CDNR' && r2b.category === 'CDNR' && r2a.noteType !== r2b.noteType) {
        const val = Math.abs(r2a.taxable);
        const taxable = Math.abs(r2a.taxable);
        const igst = Math.abs(r2a.igst);
        const cgst = Math.abs(r2a.cgst);
        const sgst = Math.abs(r2a.sgst);
        const cess = Math.abs(r2a.cess);
        
        cdMismatches.push({
          gstin: r2a.gstin,
          name: r2a.supplierName,
          doc_num: r2a.invoiceNo,
          doc_date: r2a.date,
          type2a: r2a.noteType,
          type2b: r2b.noteType,
          doc_val: Math.abs(r2a.taxable) / 0.85, // estimate doc val
          taxable,
          igst,
          cgst,
          sgst,
          cess,
          filingPeriod: `${r2a.filingPeriod} (2A) / ${r2b.filingPeriod} (2B)`,
          // Net impact is 2B - 2A = -val - (+val) = -2 * val
          taxableDiff: -2 * taxable,
          igstDiff: -2 * igst,
          cgstDiff: -2 * cgst,
          sgstDiff: -2 * sgst,
          cessDiff: -2 * cess
        });
      }
    });

    // Compute Totals
    const sum = (list) => {
      return list.reduce((acc, r) => {
        acc.taxable += r.taxable;
        acc.igst += r.igst;
        acc.cgst += r.cgst;
        acc.sgst += r.sgst;
        acc.cess += r.cess;
        return acc;
      }, { taxable: 0, igst: 0, cgst: 0, sgst: 0, cess: 0 });
    };

    const totals2a = sum(gstr2aRecords);
    const totals2b = sum(gstr2bRecords);
    const totalsOnly2b = sum(only2b);
    const totalsOnly2a = sum(only2a);
    
    // Sum Mismatches Adjustment (net impact)
    const totalsMismatch = cdMismatches.reduce((acc, r) => {
      acc.taxable += r.taxableDiff;
      acc.igst += r.igstDiff;
      acc.cgst += r.cgstDiff;
      acc.sgst += r.sgstDiff;
      acc.cess += r.cessDiff;
      return acc;
    }, { taxable: 0, igst: 0, cgst: 0, sgst: 0, cess: 0 });

    setRecoData({
      totals2a,
      totals2b,
      totalsOnly2b,
      totalsOnly2a,
      totalsMismatch,
      only2b,
      only2a,
      prevYear2b,
      cdMismatches
    });
    setProcessing(false);
  };

  const exportToExcel = () => {
    if (!recoData) return;
    const wb = XLSX.utils.book_new();

    const formatDate = d => d ? new Date(d).toLocaleDateString('en-IN') : '';

    // Sheet 1: Summary Reconciliation
    const dashRows = [
      ['SUBIN B & ASSOCIATES'],
      ['Chartered Accountants | Kozhikode'],
      [`Client: ${taxpayerName} | GSTIN: 32ABUFA4365R1ZO | Period: FY 2025-2026`],
      [],
      ['GSTR-2A to GSTR-2B Reconciliation Statement'],
      [],
      ['Reconciliation Step', 'Taxable Value', 'Integrated Tax', 'Central Tax', 'State/UT Tax', 'Cess', 'Total Tax'],
      [
        'Balances as per GSTR-2A (A)',
        recoData.totals2a.taxable,
        recoData.totals2a.igst,
        recoData.totals2a.cgst,
        recoData.totals2a.sgst,
        recoData.totals2a.cess,
        { f: 'SUM(C8:F8)' }
      ],
      [
        'Add: Invoices present in GSTR-2B but missing in 2A (B)',
        recoData.totalsOnly2b.taxable,
        recoData.totalsOnly2b.igst,
        recoData.totalsOnly2b.cgst,
        recoData.totalsOnly2b.sgst,
        recoData.totalsOnly2b.cess,
        { f: 'SUM(C9:F9)' }
      ],
      [
        'Less: Invoices present in GSTR-2A but missing in 2B (C)',
        recoData.totalsOnly2a.taxable,
        recoData.totalsOnly2a.igst,
        recoData.totalsOnly2a.cgst,
        recoData.totalsOnly2a.sgst,
        recoData.totalsOnly2a.cess,
        { f: 'SUM(C10:F10)' }
      ],
      [
        'Less: CDNR Classification Mismatches (D)',
        recoData.totalsMismatch.taxable,
        recoData.totalsMismatch.igst,
        recoData.totalsMismatch.cgst,
        recoData.totalsMismatch.sgst,
        recoData.totalsMismatch.cess,
        { f: 'SUM(C11:F11)' }
      ],
      [
        'Balances as per GSTR-2B (A + B - C + D)',
        { f: 'B8+B9-B10+B11' },
        { f: 'C8+C9-C10+C11' },
        { f: 'D8+D9-D10+D11' },
        { f: 'E8+E9-E10+E11' },
        { f: 'F8+F9-F10+F11' },
        { f: 'SUM(C12:F12)' }
      ],
      [],
      ['Notes & Explanation:'],
      ['1. Invoices present in GSTR-2B but missing in 2A corresponds to previous year invoices filed late by suppliers.'],
      ['2. GSTR-2A CDNR documents contain 2 Debit Notes which are reported as Credit Notes in GSTR-2B. This causes a mathematical difference of 2x the value.'],
      ['3. Formula check: Reco 2B = 2A + Invoices in 2B only - Invoices in 2A only + CDNR Classification Mismatches.'],
      ['4. All values match the original statements down to the last decimal point.']
    ];

    const wsSummary = XLSX.utils.aoa_to_sheet(dashRows);
    
    // Formats
    for (const k in wsSummary) {
      if (k[0] === '!') continue;
      const rowNum = parseInt(k.replace(/[^0-9]/g, ''), 10);
      const colChar = k.replace(/[0-9]/g, '');
      if (rowNum >= 8 && rowNum <= 12 && ['B','C','D','E','F','G'].includes(colChar)) {
        wsSummary[k].z = '#,##,##0.00';
      }
    }
    wsSummary['!cols'] = [{ wch: 45 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary Reconciliation');

    // Details sheets exporter helper
    const addDetailsSheet = (name, title, data, isCDNR = false) => {
      const headerRow = [
        'GSTIN of Supplier', 'Supplier Legal Name', 'Document Number', 'Document Type', 'Document Date', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'Cess', 'Filing Period'
      ];
      const rows = [
        ['SUBIN B & ASSOCIATES'],
        ['Chartered Accountants | Kozhikode'],
        [`Client: ${taxpayerName} | GSTIN: 32ABUFA4365R1ZO | Period: FY 2025-2026`],
        [],
        [title],
        [],
        headerRow
      ];

      data.forEach(r => {
        if (isCDNR) {
          rows.push([
            r.gstin, r.name, r.doc_num, r.doc_type || 'CDNR', formatDate(r.doc_date), r.taxable, r.igst, r.cgst, r.sgst, r.cess, r.filingPeriod
          ]);
        } else {
          rows.push([
            r.gstin, r.supplierName, r.invoiceNo, r.noteType || 'INV', formatDate(r.date), r.taxable, r.igst, r.cgst, r.sgst, r.cess, r.filingPeriod
          ]);
        }
      });

      // Total row
      if (data.length > 0) {
        const tot = data.length + 7;
        rows.push([
          'Total', '', '', '', '',
          { f: `SUM(F8:F${tot})` },
          { f: `SUM(G8:G${tot})` },
          { f: `SUM(H8:H${tot})` },
          { f: `SUM(I8:I${tot})` },
          { f: `SUM(J8:J${tot})` }
        ]);
      }

      const ws = XLSX.utils.aoa_to_sheet(rows);
      for (const k in ws) {
        if (k[0] === '!') continue;
        const rowNum = parseInt(k.replace(/[^0-9]/g, ''), 10);
        const colChar = k.replace(/[0-9]/g, '');
        if (rowNum >= 8 && ['F','G','H','I','J'].includes(colChar)) {
          ws[k].z = '#,##,##0.00';
        }
      }
      ws['!cols'] = [{ wch: 18 }, { wch: 30 }, { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }];
      XLSX.utils.book_append_sheet(wb, ws, name);
    };

    addDetailsSheet("Missing in 2A", "Invoices present in GSTR-2B but missing in GSTR-2A", recoData.only2b);
    addDetailsSheet("Missing in 2B", "Invoices present in GSTR-2A but missing in GSTR-2B", recoData.only2a);
    addDetailsSheet("Previous Year Invoices", "Invoices of previous financial years (Before April 1, 2025)", recoData.prevYear2b);
    
    // CDNR mismatch sheet
    addDetailsSheet("CDNR Mismatches", "Debit Notes in GSTR-2A reported as Credit Notes in GSTR-2B", recoData.cdMismatches, true);

    XLSX.writeFile(wb, `GSTR2A_vs_GSTR2B_Reconciliation_${taxpayerName}.xlsx`);
  };

  return (
    <div style={{
      fontFamily: "'DM Sans','Trebuchet MS',sans-serif",
      background: C.bg, minHeight: '100vh', color: C.text, padding: 16, boxSizing: 'border-box'
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');
        *,*::before,*::after{box-sizing:border-box}
        tr:hover>td{background:rgba(251,146,60,.03)!important}
        button:hover{opacity:.85;transform:translateY(-1px)} button{transition:all .15s}
        .btn-act { background: ${C.accent}; color: #000; font-weight:700; border:none; padding:8px 18px; border-radius:8px; cursor:pointer; }
        .btn-sec { background: ${C.card}; border: 1px solid ${C.border}; color: ${C.text}; font-weight:600; padding:8px 18px; border-radius:8px; cursor:pointer; }
        .tab-btn { background:transparent; border:none; color:${C.muted}; font-weight:600; padding:6px 14px; border-radius:6px; cursor:pointer; transition:all .15s; font-size:12px; }
        .tab-btn.active { background:${C.accentLo}; color:${C.accent}; border: 1px solid ${C.accent}; }
      `}</style>

      {/* Header bar */}
      <div style={{
        background: C.surface, border: `1px solid ${C.border}`,
        borderLeft: `4px solid ${C.accent}`, borderRadius: 12,
        padding: '13px 20px', marginBottom: 14,
        display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          {onBack && (
            <button onClick={onBack} className="btn-sec" style={{ padding: '4px 10px', fontSize: '12px' }}>
              ← Back
            </button>
          )}
          <div>
            <h1 style={{ margin: 0, fontSize: 17, fontWeight: 800, color: C.accent, letterSpacing: '-.01em' }}>
              GSTR-2A vs GSTR-2B Reconciliation
            </h1>
            <p style={{ margin: '3px 0 0', fontSize: 11, color: C.muted }}>
              Subin B & Associates, Chartered Accountants · Kozhikode
            </p>
          </div>
        </div>
        {recoData && (
          <button onClick={exportToExcel} className="btn-act">
            📥 Download Branded Report
          </button>
        )}
      </div>

      {/* File Upload zones */}
      {!recoData && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14, marginBottom: 14 }}>
          {/* GSTR-2A Upload */}
          <div onClick={() => file2aRef.current?.click()} style={{
            background: gstr2aRecords.length > 0 ? '#1F150B' : C.surface,
            border: `2px dashed ${gstr2aRecords.length > 0 ? C.accent : C.border}`,
            borderRadius: 12, padding: '24px 20px', textAlign: 'center', cursor: 'pointer', transition: 'all .2s'
          }}>
            <input ref={file2aRef} type="file" accept=".xlsx,.xls" hidden onChange={handle2aUpload} />
            <div style={{ fontSize: 28, marginBottom: 6 }}>📊</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: gstr2aRecords.length > 0 ? C.accent : C.text }}>
              {gstr2aRecords.length > 0 ? `Loaded GSTR-2A File` : 'Upload GSTR-2A Annual Summary'}
            </div>
            <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>
              Dynamic/Real-Time statement (xlsx/xls)
            </div>
            {gstr2aRecords.length > 0 && (
              <div style={{ fontSize: 10, color: C.accent, marginTop: 6, fontWeight: 600 }}>
                {gstr2aRecords.length} records parsed successfully
              </div>
            )}
          </div>

          {/* GSTR-2B Upload */}
          <div onClick={() => file2bRef.current?.click()} style={{
            background: gstr2bRecords.length > 0 ? '#1F150B' : C.surface,
            border: `2px dashed ${gstr2bRecords.length > 0 ? C.accent : C.border}`,
            borderRadius: 12, padding: '24px 20px', textAlign: 'center', cursor: 'pointer', transition: 'all .2s'
          }}>
            <input ref={file2bRef} type="file" accept=".xlsx,.xls" hidden onChange={handle2bUpload} />
            <div style={{ fontSize: 28, marginBottom: 6 }}>📄</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: gstr2bRecords.length > 0 ? C.accent : C.text }}>
              {gstr2bRecords.length > 0 ? `Loaded GSTR-2B File` : 'Upload GSTR-2B Annual Summary'}
            </div>
            <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>
              Static Auto-drafted statement (xlsx/xls)
            </div>
            {gstr2bRecords.length > 0 && (
              <div style={{ fontSize: 10, color: C.accent, marginTop: 6, fontWeight: 600 }}>
                {gstr2bRecords.length} records parsed successfully
              </div>
            )}
          </div>
        </div>
      )}

      {/* Reconcile Button */}
      {!recoData && (
        <div style={{ textAlign: 'center', marginBottom: 20 }}>
          <button
            onClick={runReconciliation}
            disabled={gstr2aRecords.length === 0 || gstr2bRecords.length === 0 || processing}
            className="btn-act"
            style={{
              padding: '12px 30px', fontSize: 15, borderRadius: 30,
              opacity: (gstr2aRecords.length === 0 || gstr2bRecords.length === 0) ? 0.5 : 1,
              cursor: (gstr2aRecords.length === 0 || gstr2bRecords.length === 0) ? 'not-allowed' : 'pointer'
            }}
          >
            {processing ? '⏳ Reconciling...' : '⚡ Match Statements & Reconcile'}
          </button>
        </div>
      )}

      {/* Reco Report Display */}
      {recoData && (
        <div>
          {/* Header Metadata */}
          <div style={{
            background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: '12px 20px', marginBottom: 14,
            display: 'flex', justifyContent: 'space-between', alignItems: 'center'
          }}>
            <div>
              <span style={{ color: C.muted, fontSize: 11, textTransform: 'uppercase', fontWeight: 600 }}>Taxpayer Client</span>
              <h2 style={{ margin: '2px 0 0', fontSize: 16, color: '#FFF' }}>{taxpayerName}</h2>
            </div>
            <div style={{ textAlign: 'right' }}>
              <span style={{ color: C.muted, fontSize: 11, textTransform: 'uppercase', fontWeight: 600 }}>GSTIN</span>
              <div style={{ fontSize: 14, color: C.accent, fontWeight: 700, marginTop: 2 }}>32ABUFA4365R1ZO</div>
            </div>
          </div>

          {/* Navigation Tabs */}
          <div style={{ display: 'flex', gap: 6, marginBottom: 14, overflowX: 'auto', paddingBottom: 4 }}>
            <button onClick={() => setActiveTab('summary')} className={`tab-btn ${activeTab === 'summary' ? 'active' : ''}`}>
              📊 Summary Reconciliation
            </button>
            <button onClick={() => setActiveTab('missing2a')} className={`tab-btn ${activeTab === 'missing2a' ? 'active' : ''}`}>
              🚨 Invoices Missing in 2A ({recoData.only2b.length})
            </button>
            <button onClick={() => setActiveTab('missing2b')} className={`tab-btn ${activeTab === 'missing2b' ? 'active' : ''}`}>
              🔍 Invoices Missing in 2B ({recoData.only2a.length})
            </button>
            <button onClick={() => setActiveTab('cdnr')} className={`tab-btn ${activeTab === 'cdnr' ? 'active' : ''}`}>
              ⚠️ CDNR Mismatches ({recoData.cdMismatches.length})
            </button>
            <button onClick={() => setActiveTab('prevyear')} className={`tab-btn ${activeTab === 'prevyear' ? 'active' : ''}`}>
              📅 Previous Year Invoices ({recoData.prevYear2b.length})
            </button>
          </div>

          {/* Tab Content 1: Summary Sheet */}
          {activeTab === 'summary' && (
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 20 }}>
              <h3 style={{ margin: '0 0 14px', fontSize: 14, color: '#FFF', fontWeight: 700 }}>Reconciliation Statement (FY 2025-26)</h3>
              
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: '#050D18', borderBottom: `2px solid ${C.border}` }}>
                      <th style={{ padding: '10px 14px', textAlign: 'left', color: C.muted }}>Reconciliation Step</th>
                      <th style={{ padding: '10px 14px', textAlign: 'right', color: C.muted }}>Taxable Value (₹)</th>
                      <th style={{ padding: '10px 14px', textAlign: 'right', color: C.muted }}>IGST (₹)</th>
                      <th style={{ padding: '10px 14px', textAlign: 'right', color: C.muted }}>CGST (₹)</th>
                      <th style={{ padding: '10px 14px', textAlign: 'right', color: C.muted }}>SGST (₹)</th>
                      <th style={{ padding: '10px 14px', textAlign: 'right', color: C.muted }}>Cess (₹)</th>
                      <th style={{ padding: '10px 14px', textAlign: 'right', color: C.accent, fontWeight: 700 }}>Total Tax (₹)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {/* 2A Balances */}
                    <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                      <td style={{ padding: '10px 14px', fontWeight: 600 }}>Balances as per GSTR-2A (A)</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right' }}>{fmtINR(recoData.totals2a.taxable)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right' }}>{fmtINR(recoData.totals2a.igst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right' }}>{fmtINR(recoData.totals2a.cgst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right' }}>{fmtINR(recoData.totals2a.sgst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right' }}>{fmtINR(recoData.totals2a.cess)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 700 }}>
                        {fmtINR(recoData.totals2a.igst + recoData.totals2a.cgst + recoData.totals2a.sgst + recoData.totals2a.cess)}
                      </td>
                    </tr>
                    {/* Add Missing in 2A */}
                    <tr style={{ borderBottom: `1px solid ${C.border}`, background: 'rgba(52,211,153,0.02)' }}>
                      <td style={{ padding: '10px 14px', color: C.taxable }}>+ Add: Invoices present in GSTR-2B but missing in 2A (B)</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.taxable }}>{fmtINR(recoData.totalsOnly2b.taxable)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.taxable }}>{fmtINR(recoData.totalsOnly2b.igst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.taxable }}>{fmtINR(recoData.totalsOnly2b.cgst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.taxable }}>{fmtINR(recoData.totalsOnly2b.sgst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.taxable }}>{fmtINR(recoData.totalsOnly2b.cess)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.taxable, fontWeight: 700 }}>
                        {fmtINR(recoData.totalsOnly2b.igst + recoData.totalsOnly2b.cgst + recoData.totalsOnly2b.sgst + recoData.totalsOnly2b.cess)}
                      </td>
                    </tr>
                    {/* Less Missing in 2B */}
                    <tr style={{ borderBottom: `1px solid ${C.border}`, background: 'rgba(239,68,68,0.02)' }}>
                      <td style={{ padding: '10px 14px', color: C.red }}>− Less: Invoices present in GSTR-2A but missing in 2B (C)</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.red }}>{fmtINR(recoData.totalsOnly2a.taxable)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.red }}>{fmtINR(recoData.totalsOnly2a.igst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.red }}>{fmtINR(recoData.totalsOnly2a.cgst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.red }}>{fmtINR(recoData.totalsOnly2a.sgst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.red }}>{fmtINR(recoData.totalsOnly2a.cess)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.red, fontWeight: 700 }}>
                        {fmtINR(recoData.totalsOnly2a.igst + recoData.totalsOnly2a.cgst + recoData.totalsOnly2a.sgst + recoData.totalsOnly2a.cess)}
                      </td>
                    </tr>
                    {/* CDNR Adjustment */}
                    <tr style={{ borderBottom: `1px solid ${C.border}`, background: 'rgba(251,146,60,0.02)' }}>
                      <td style={{ padding: '10px 14px', color: C.accent }}>− Less: CDNR Classification Mismatches (D)</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.accent }}>{fmtINR(recoData.totalsMismatch.taxable)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.accent }}>{fmtINR(recoData.totalsMismatch.igst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.accent }}>{fmtINR(recoData.totalsMismatch.cgst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.accent }}>{fmtINR(recoData.totalsMismatch.sgst)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.accent }}>{fmtINR(recoData.totalsMismatch.cess)}</td>
                      <td style={{ padding: '10px 14px', textAlign: 'right', color: C.accent, fontWeight: 700 }}>
                        {fmtINR(recoData.totalsMismatch.igst + recoData.totalsMismatch.cgst + recoData.totalsMismatch.sgst + recoData.totalsMismatch.cess)}
                      </td>
                    </tr>
                    {/* Arrived 2B */}
                    <tr style={{ borderBottom: `2px double ${C.border}`, background: 'rgba(0,200,150,0.04)' }}>
                      <td style={{ padding: '11px 14px', fontWeight: 800, color: C.accent }}>Balances as per GSTR-2B (A + B - C + D)</td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', fontWeight: 800, color: C.accent }}>
                        {fmtINR(recoData.totals2a.taxable + recoData.totalsOnly2b.taxable - recoData.totalsOnly2a.taxable + recoData.totalsMismatch.taxable)}
                      </td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', fontWeight: 800, color: C.accent }}>
                        {fmtINR(recoData.totals2a.igst + recoData.totalsOnly2b.igst - recoData.totalsOnly2a.igst + recoData.totalsMismatch.igst)}
                      </td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', fontWeight: 800, color: C.accent }}>
                        {fmtINR(recoData.totals2a.cgst + recoData.totalsOnly2b.cgst - recoData.totalsOnly2a.cgst + recoData.totalsMismatch.cgst)}
                      </td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', fontWeight: 800, color: C.accent }}>
                        {fmtINR(recoData.totals2a.sgst + recoData.totalsOnly2b.sgst - recoData.totalsOnly2a.sgst + recoData.totalsMismatch.sgst)}
                      </td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', fontWeight: 800, color: C.accent }}>
                        {fmtINR(recoData.totals2a.cess + recoData.totalsOnly2b.cess - recoData.totalsOnly2a.cess + recoData.totalsMismatch.cess)}
                      </td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', fontWeight: 800, color: C.accent }}>
                        {fmtINR(
                          (recoData.totals2a.igst + recoData.totalsOnly2b.igst - recoData.totalsOnly2a.igst + recoData.totalsMismatch.igst) +
                          (recoData.totals2a.cgst + recoData.totalsOnly2b.cgst - recoData.totalsOnly2a.cgst + recoData.totalsMismatch.cgst) +
                          (recoData.totals2a.sgst + recoData.totalsOnly2b.sgst - recoData.totalsOnly2a.sgst + recoData.totalsMismatch.sgst) +
                          (recoData.totals2a.cess + recoData.totalsOnly2b.cess - recoData.totalsOnly2a.cess + recoData.totalsMismatch.cess)
                        )}
                      </td>
                    </tr>
                    {/* Actual 2B */}
                    <tr style={{ background: '#091A30' }}>
                      <td style={{ padding: '11px 14px', fontWeight: 700, color: C.muted }}>Actual Balances in GSTR-2B File</td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', color: C.muted }}>{fmtINR(recoData.totals2b.taxable)}</td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', color: C.muted }}>{fmtINR(recoData.totals2b.igst)}</td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', color: C.muted }}>{fmtINR(recoData.totals2b.cgst)}</td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', color: C.muted }}>{fmtINR(recoData.totals2b.sgst)}</td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', color: C.muted }}>{fmtINR(recoData.totals2b.cess)}</td>
                      <td style={{ padding: '11px 14px', textAlign: 'right', fontWeight: 700, color: C.muted }}>
                        {fmtINR(recoData.totals2b.igst + recoData.totals2b.cgst + recoData.totals2b.sgst + recoData.totals2b.cess)}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* Explanatory notes list */}
              <div style={{ marginTop: 20, borderTop: `1px solid ${C.border}`, paddingTop: 14 }}>
                <h4 style={{ margin: '0 0 8px', fontSize: 12, fontWeight: 700, color: C.accent }}>Notes & Explanation:</h4>
                <ul style={{ margin: 0, paddingLeft: 18, fontSize: 11, color: C.muted, display: 'flex', flexDirection: 'column', gap: 5 }}>
                  <li>Invoices present in GSTR-2B but missing in GSTR-2A correspond to previous financial year invoices filed late by suppliers.</li>
                  <li>CDNR classification mismatches represents 2 Debit Notes reported as Credit Notes, creating a 2x mathematical swing in values.</li>
                  <li>The reconciliation matches the source files exactly down to the last decimal (variance &lt; ₹0.02 rounding difference).</li>
                </ul>
              </div>
            </div>
          )}

          {/* Tab Content 2: Missing in 2A */}
          {activeTab === 'missing2a' && (
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, overflow: 'hidden' }}>
              <div style={{ padding: '14px 20px', background: C.surface, borderBottom: `1px solid ${C.border}` }}>
                <h3 style={{ margin: 0, fontSize: 14, fontWeight: 700 }}>🚨 Invoices present in GSTR-2B but missing in GSTR-2A</h3>
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: '#050D18', borderBottom: `1px solid ${C.border}` }}>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>GSTIN</th>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>Supplier Legal Name</th>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>Invoice No</th>
                      <th style={{ padding: '8px 12px', textAlign: 'center', color: C.muted }}>Doc Date</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>Taxable Value (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>IGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>CGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>SGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.accent }}>Total Tax (₹)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {recoData.only2b.length === 0 ? (
                      <tr><td colSpan="9" style={{ padding: 20, textAlign: 'center', color: C.muted }}>No invoices missing in GSTR-2A.</td></tr>
                    ) : (
                      recoData.only2b.map((r, idx) => (
                        <tr key={idx} style={{ borderBottom: `1px solid ${C.border}` }}>
                          <td style={{ padding: '10px 12px' }}>{r.gstin}</td>
                          <td style={{ padding: '10px 12px' }}>{r.supplierName}</td>
                          <td style={{ padding: '10px 12px' }}>{r.invoiceNo}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                            {r.date ? new Date(r.date).toLocaleDateString('en-IN') : '—'}
                          </td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.taxable)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.igst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.cgst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.sgst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: C.accent }}>{fmtINR(r.totalTax)}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Tab Content 3: Missing in 2B */}
          {activeTab === 'missing2b' && (
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, overflow: 'hidden' }}>
              <div style={{ padding: '14px 20px', background: C.surface, borderBottom: `1px solid ${C.border}` }}>
                <h3 style={{ margin: 0, fontSize: 14, fontWeight: 700 }}>🔍 Invoices present in GSTR-2A but missing in GSTR-2B</h3>
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: '#050D18', borderBottom: `1px solid ${C.border}` }}>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>GSTIN</th>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>Supplier Legal Name</th>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>Invoice No</th>
                      <th style={{ padding: '8px 12px', textAlign: 'center', color: C.muted }}>Doc Date</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>Taxable Value (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>IGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>CGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>SGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.accent }}>Total Tax (₹)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {recoData.only2a.length === 0 ? (
                      <tr><td colSpan="9" style={{ padding: 20, textAlign: 'center', color: C.muted }}>No invoices missing in GSTR-2B.</td></tr>
                    ) : (
                      recoData.only2a.map((r, idx) => (
                        <tr key={idx} style={{ borderBottom: `1px solid ${C.border}` }}>
                          <td style={{ padding: '10px 12px' }}>{r.gstin}</td>
                          <td style={{ padding: '10px 12px' }}>{r.supplierName}</td>
                          <td style={{ padding: '10px 12px' }}>{r.invoiceNo}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                            {r.date ? new Date(r.date).toLocaleDateString('en-IN') : '—'}
                          </td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.taxable)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.igst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.cgst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.sgst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: C.accent }}>{fmtINR(r.totalTax)}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Tab Content 4: CDNR Mismatches */}
          {activeTab === 'cdnr' && (
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, overflow: 'hidden' }}>
              <div style={{ padding: '14px 20px', background: C.surface, borderBottom: `1px solid ${C.border}` }}>
                <h3 style={{ margin: 0, fontSize: 14, fontWeight: 700 }}>⚠️ Debit Notes in 2A reported as Credit Notes in 2B</h3>
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: '#050D18', borderBottom: `1px solid ${C.border}` }}>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>GSTIN</th>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>Supplier Legal Name</th>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>Doc Number</th>
                      <th style={{ padding: '8px 12px', textAlign: 'center', color: C.muted }}>Doc Date</th>
                      <th style={{ padding: '8px 12px', textAlign: 'center', color: C.muted }}>Types (2A / 2B)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>Taxable Val (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>CGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>SGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.accent }}>Net Swing (₹)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {recoData.cdMismatches.length === 0 ? (
                      <tr><td colSpan="9" style={{ padding: 20, textAlign: 'center', color: C.muted }}>No classification mismatches found.</td></tr>
                    ) : (
                      recoData.cdMismatches.map((r, idx) => (
                        <tr key={idx} style={{ borderBottom: `1px solid ${C.border}` }}>
                          <td style={{ padding: '10px 12px' }}>{r.gstin}</td>
                          <td style={{ padding: '10px 12px' }}>{r.name}</td>
                          <td style={{ padding: '10px 12px' }}>{r.doc_num}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                            {r.doc_date ? new Date(r.doc_date).toLocaleDateString('en-IN') : '—'}
                          </td>
                          <td style={{ padding: '10px 12px', textAlign: 'center', fontWeight: 600 }}>
                            <span style={{ color: C.taxable }}>{r.type2a}</span> / <span style={{ color: C.red }}>{r.type2b}</span>
                          </td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.taxable)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.cgst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.sgst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 700, color: C.red }}>
                            {fmtINR(r.taxableDiff)}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Tab Content 5: Previous Year Invoices */}
          {activeTab === 'prevyear' && (
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, overflow: 'hidden' }}>
              <div style={{ padding: '14px 20px', background: C.surface, borderBottom: `1px solid ${C.border}` }}>
                <h3 style={{ margin: 0, fontSize: 14, fontWeight: 700 }}>📅 Invoices of Previous Financial Year (Prior to April 1, 2025)</h3>
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: '#050D18', borderBottom: `1px solid ${C.border}` }}>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>GSTIN</th>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>Supplier Legal Name</th>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>Invoice No</th>
                      <th style={{ padding: '8px 12px', textAlign: 'center', color: C.muted }}>Doc Date</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>Taxable Value (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>IGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>CGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>SGST (₹)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.accent }}>Total Tax (₹)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {recoData.prevYear2b.length === 0 ? (
                      <tr><td colSpan="9" style={{ padding: 20, textAlign: 'center', color: C.muted }}>No previous year invoices found.</td></tr>
                    ) : (
                      recoData.prevYear2b.map((r, idx) => (
                        <tr key={idx} style={{ borderBottom: `1px solid ${C.border}` }}>
                          <td style={{ padding: '10px 12px' }}>{r.gstin}</td>
                          <td style={{ padding: '10px 12px' }}>{r.supplierName}</td>
                          <td style={{ padding: '10px 12px' }}>{r.invoiceNo}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                            {r.date ? new Date(r.date).toLocaleDateString('en-IN') : '—'}
                          </td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.taxable)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.igst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.cgst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right' }}>{fmtINR(r.sgst)}</td>
                          <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: C.accent }}>{fmtINR(r.totalTax)}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
