import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// Color Palette
const C = {
  bg: '#06111E',
  surface: '#0C1A2B',
  card: '#0F2035',
  border: '#163050',
  accent: '#00C896',
  accentLo: '#00C89614',
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
  total: '#94A3B8',
  gross: '#38BDF8'
};

// Date normalizer
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
  if (!n || Math.abs(n) < 0.005) return '—';
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

// Reuse Step 2 Tally column classification logic
function extractRate(raw) {
  const c = String(raw || '').replace(/\s+/g, '');
  const m = c.match(/[@]?(\d+(?:\.\d+)?)%/);
  return m ? parseFloat(m[1]) / 100 : null;
}

function classifyCol(raw, allowInput = false) {
  if (!raw) return null;
  const u = String(raw).toUpperCase().replace(/\s+/g, '');
  const isGSTCol = u.includes('CGST') || u.includes('SGST') || u.includes('IGST') || u.includes('UTGST') || u.includes('CESS');
  
  if (isGSTCol) {
    if (u.includes('CGST'))                      return { role: 'TAX', type: 'cgst', rate: extractRate(raw) };
    if (u.includes('SGST') || u.includes('UTGST')) return { role: 'TAX', type: 'sgst', rate: extractRate(raw) };
    if (u.includes('IGST'))                      return { role: 'TAX', type: 'igst', rate: extractRate(raw) };
    if (u.includes('CESS'))                      return { role: 'TAX', type: 'cess', rate: extractRate(raw) };
  }
  
  const skipKeywords = ['DATE', 'INVOICE', 'VOUCHER', 'VCH', 'GSTIN', 'CONSIGNEE', 'NARRATION', 'PARTICULARS', 'PARTY', 'ROUNDOFF', 'ROUND'];
  if (skipKeywords.some(kw => u.includes(kw))) {
    return { role: 'SKIP' };
  }
  if (u === 'GROSSTOTAL') return { role: 'SKIP' };
  if (u === 'VALUE') return { role: 'TAXABLE_VALUE' };
  
  const hasTaxableKeyword = u.includes('SALE') || u.includes('PURCHASE') || u.includes('PURCH') || u.includes('INTERSTATE');
  if (hasTaxableKeyword && !isGSTCol) {
    return { role: 'SALE_LEDGER', type: 'taxable', rate: extractRate(raw) };
  }
  return null;
}

function solveRowSigns(rows, cgstColIdx, gtCgst) {
  const vals = rows.map(r => Math.abs(Number(r[cgstColIdx] || 0)));
  const n = vals.length;
  if (n === 0) return [];
  if (n === 1) return [gtCgst >= 0 ? 1 : -1];
  if (n > 15) return new Array(n).fill(1);
  for (let mask = 0; mask < (1 << n); mask++) {
    const signs = Array.from({ length: n }, (_, i) => (mask >> i) & 1 ? 1 : -1);
    const total = signs.reduce((s, sg, i) => s + sg * vals[i], 0);
    if (Math.abs(total - gtCgst) < 1) return signs;
  }
  return new Array(n).fill(1);
}

// ---------------- MAIN COMPONENT ----------------
export default function GSTReconciliationStep3({ onBack }) {
  const [booksFiles, setBooksFiles] = useState([]);
  const [gstrFiles, setGstrFiles] = useState([]);
  const [booksRecords, setBooksRecords] = useState([]);
  const [gstrRecords, setGstrRecords] = useState([]);
  const [recoData, setRecoData] = useState(null);
  
  const [processing, setProcessing] = useState(false);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('ALL');
  
  const booksFileRef = useRef();
  const gstrFileRef = useRef();

  // Load and Parse Tally Purchase Registers
  const parseTallyFile = async (file) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: false });
    const records = [];
    
    for (const sn of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: true });
      if (rows.length < 5) continue;
      
      const clientName = String(rows[0]?.[0] || '').trim();
      let hRow = -1;
      for (let i = 0; i < Math.min(rows.length, 15); i++) {
        if (String(rows[i]?.[0] || '').trim().toLowerCase() === 'date') { hRow = i; break; }
      }
      if (hRow < 0) continue;
      
      const headers = rows[hRow].map(h => String(h ?? '').trim());
      const sheetNameLower = sn.toLowerCase();
      let regType = 'PURCHASE';
      if (sheetNameLower.includes('debit note')) regType = 'DEBIT_NOTE';
      else if (sheetNameLower.includes('journal')) regType = 'JOURNAL';
      
      const sign = regType === 'DEBIT_NOTE' ? -1 : 1;
      const allowInput = regType !== 'JOURNAL';
      const colMap = headers.map((h, i) => ({ i, h, cls: classifyCol(h, allowInput) }));
      
      const dataRows = [];
      let grandTotalRow = null;
      for (const row of rows.slice(hRow + 1)) {
        if (!row) continue;
        const partName = String(row[1] ?? '').trim().toLowerCase();
        if (partName === 'grand total') { grandTotalRow = row; continue; }
        if (!row[0]) continue;
        dataRows.push(row);
      }
      
      const taxCols = colMap.filter(c => c.cls?.role === 'TAX');
      const gstRows = dataRows.filter(r => taxCols.some(col => Math.abs(Number(r[col.i] || 0)) > 0.005));
      
      let rowSigns = new Array(gstRows.length).fill(sign);
      if (regType === 'JOURNAL' && grandTotalRow && taxCols.length > 0) {
        let signingCol = taxCols.find(c => Math.abs(Number(grandTotalRow[c.i] || 0)) > 0.005) || taxCols[0];
        rowSigns = solveRowSigns(gstRows, signingCol.i, Number(grandTotalRow[signingCol.i] || 0));
      }
      
      const valCol = colMap.find(c => c.cls?.role === 'TAXABLE_VALUE');
      const saleLedgerCols = colMap.filter(c => c.cls?.role === 'SALE_LEDGER');
      const structuralCols = new Set(colMap.filter(c => c.cls?.role === 'SKIP').map(c => c.i));
      const gstColsSet = new Set(taxCols.map(c => c.i));
      
      const partyColSet = new Set();
      const taxableColSet = new Set();
      
      if (!valCol) {
        for (const row of gstRows) {
          const active = [];
          for (let i = 0; i < row.length; i++) {
            if (structuralCols.has(i)) continue;
            const v = Math.abs(Number(row[i] || 0));
            if (v > 0.005) active.push({ i, v });
          }
          for (const cand of active) {
            const othersSum = active.filter(c => c.i !== cand.i).reduce((s, c) => s + c.v, 0);
            if (Math.abs(cand.v - othersSum) < 2) partyColSet.add(cand.i);
          }
          for (const c of active) {
            if (!gstColsSet.has(c.i) && !partyColSet.has(c.i)) taxableColSet.add(c.i);
          }
        }
      }
      
      const partColIdx = colMap.find(c => String(c.h).toUpperCase().replace(/\s+/g,'') === 'PARTICULARS')?.i ?? 1;
      const gstinColIdx = colMap.find(c => ['GSTINUIN', 'GSTIN'].includes(String(c.h).toUpperCase().replace(/\s+/g,'')))?.i;
      const invNoColIdx = colMap.find(c => ['VOUCHERNO', 'VCHNO', 'INVOICE'].some(s => String(c.h).toUpperCase().replace(/\s+/g,'').includes(s)))?.i ?? 2;
      
      const cgstCols = taxCols.filter(c => c.cls.type === 'cgst');
      const sgstCols = taxCols.filter(c => c.cls.type === 'sgst');
      const igstCols = taxCols.filter(c => c.cls.type === 'igst');
      const cessCols = taxCols.filter(c => c.cls.type === 'cess');
      
      gstRows.forEach((row, idx) => {
        const r_sign = rowSigns[idx] || sign;
        const parsedDate = toDate(row[0]);
        if (!parsedDate) return;
        
        let rowTaxable = 0;
        if (valCol) {
          rowTaxable = Math.abs(Number(row[valCol.i] || 0));
        } else {
          for (const ti of taxableColSet) {
            if (partyColSet.has(ti)) continue;
            rowTaxable += Math.abs(Number(row[ti] || 0));
          }
          colMap.forEach(col => {
            if ((col.cls?.role === 'SALE_LEDGER' || col.cls?.role === 'TAXABLE_VALUE') && !taxableColSet.has(col.i)) {
              rowTaxable += Math.abs(Number(row[col.i] || 0));
            }
          });
        }
        
        let rowIgst = 0, rowCgst = 0, rowSgst = 0, rowCess = 0;
        igstCols.forEach(col => { rowIgst += Math.abs(Number(row[col.i] || 0)); });
        cgstCols.forEach(col => { rowCgst += Math.abs(Number(row[col.i] || 0)); });
        sgstCols.forEach(col => { rowSgst += Math.abs(Number(row[col.i] || 0)); });
        cessCols.forEach(col => { rowCess += Math.abs(Number(row[col.i] || 0)); });
        
        records.append({
          date: parsedDate,
          particulars: String(row[partColIdx] || '').trim(),
          gstin: cleanGSTIN(row[gstinColIdx]),
          invoiceNo: String(row[invNoColIdx] || '').trim(),
          taxable: rowTaxable * r_sign,
          igst: rowIgst * r_sign,
          cgst: rowCgst * r_sign,
          sgst: rowSgst * r_sign,
          cess: rowCess * r_sign,
          totalTax: (rowIgst + rowCgst + rowSgst + rowCess) * r_sign,
          source: file.name
        });
      });
    }
    return records;
  };

  // Load and Parse GSTR-2B Excel sheets (B2B + CDNR)
  const parseGstr2bFile = async (file) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: false });
    const records = [];
    
    // helper to get month name
    const getGstr2bMonth = (fDate) => {
      if (!fDate) return "Unknown";
      const d = toDate(fDate);
      if (!d) return "Unknown";
      const day = d.getUTCDate();
      const prev = new Date(d);
      prev.setUTCMonth(prev.getUTCMonth() - 1);
      const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
      return day <= 11 ? months[prev.getUTCMonth()] : months[d.getUTCMonth()];
    };
    
    // Parse B2B sheet
    const b2bSheets = wb.SheetNames.filter(s => s.toUpperCase() === 'B2B');
    if (b2bSheets.length > 0) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[b2bSheets[0]], { header: 1, defval: '', raw: true });
      let hIdx = -1;
      for (let i = 0; i < Math.min(rows.length, 15); i++) {
        if (rows[i] && rows[i].some(c => String(c).toUpperCase().includes('GSTIN'))) { hIdx = i; break; }
      }
      if (hIdx !== -1) {
        rows.slice(hIdx + 1).forEach(row => {
          if (!row || !row[0]) return;
          const gstin = cleanGSTIN(row[0]);
          const name = String(row[1] || '').trim();
          const invNo = String(row[2] || '').trim();
          const dt = toDate(row[3]);
          const taxable = Number(row[5] || 0);
          const igst = Number(row[6] || 0);
          const cgst = Number(row[7] || 0);
          const sgst = Number(row[8] || 0);
          const cess = Number(row[9] || 0);
          const filingDate = toDate(row[10]);
          
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
            filingDate,
            monthName: getGstr2bMonth(filingDate),
            noteType: 'INV'
          });
        });
      }
    }
    
    // Parse CDNR sheet
    const cdnrSheets = wb.SheetNames.filter(s => s.toUpperCase() === 'CDNR');
    if (cdnrSheets.length > 0) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[cdnrSheets[0]], { header: 1, defval: '', raw: true });
      let hIdx = -1;
      for (let i = 0; i < Math.min(rows.length, 15); i++) {
        if (rows[i] && rows[i].some(c => String(c).toUpperCase().includes('GSTIN'))) { hIdx = i; break; }
      }
      if (hIdx !== -1) {
        rows.slice(hIdx + 1).forEach(row => {
          if (!row || !row[0]) return;
          const gstin = cleanGSTIN(row[0]);
          const name = String(row[1] || '').trim();
          const noteNo = String(row[2] || '').trim();
          const noteType = String(row[3] || 'C').trim().toUpperCase(); // C or D
          const dt = toDate(row[4]);
          
          const mult = noteType.includes('C') ? -1 : 1;
          const taxable = Number(row[6] || 0) * mult;
          const igst = Number(row[7] || 0) * mult;
          const cgst = Number(row[8] || 0) * mult;
          const sgst = Number(row[9] || 0) * mult;
          const cess = Number(row[10] || 0) * mult;
          const filingDate = toDate(row[11]);
          
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
            filingDate,
            monthName: getGstr2bMonth(filingDate),
            noteType
          });
        });
      }
    }
    return records;
  };

  const handleBooksFiles = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setProcessing(true);
    setBooksFiles(files);
    
    let allRecords = [];
    for (const f of files) {
      try {
        const recs = await parseTallyFile(f);
        allRecords = allRecords.concat(recs);
      } catch (ex) {
        console.error("Error parsing books file", f.name, ex);
      }
    }
    setBooksRecords(allRecords);
    setProcessing(false);
  };

  const handleGstrFiles = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setProcessing(true);
    setGstrFiles(files);
    
    let allRecords = [];
    for (const f of files) {
      try {
        const recs = await parseGstr2bFile(f);
        allRecords = allRecords.concat(recs);
      } catch (ex) {
        console.error("Error parsing GSTR-2B file", f.name, ex);
      }
    }
    setGstrRecords(allRecords);
    setProcessing(false);
  };

  // Run matching logic
  const runReconciliation = () => {
    if (booksRecords.length === 0 || gstrRecords.length === 0) return;
    setProcessing(true);
    
    const books = booksRecords.map((r, i) => ({ ...r, id: `b_${i}`, matchStatus: 'Unmatched (Books Only)', matchedIdx2b: -1 }));
    const gstr = gstrRecords.map((r, i) => ({ ...r, id: `g_${i}`, matchStatus: 'Unmatched (GSTR-2B Only)', matchedIdxBooks: -1 }));
    
    // Normalization keys
    books.forEach(b => { b.normInv = normalizeInvoiceNo(b.invoiceNo); b.normGstin = cleanGSTIN(b.gstin); });
    gstr.forEach(g => { g.normInv = normalizeInvoiceNo(g.invoiceNo); g.normGstin = cleanGSTIN(g.gstin); });
    
    // Phase 1: Perfect Exact Match on GSTIN + Invoice + Rounded Tax (within ₹10) + Date (within 30 days)
    books.forEach(b => {
      if (!b.normGstin || !b.normInv) return;
      const matches = gstr.filter(g => g.normGstin === b.normGstin && g.normInv === b.normInv && g.matchedIdxBooks === -1);
      
      for (const g of matches) {
        const taxDiff = Math.abs(b.totalTax - g.totalTax);
        const dateDiff = b.date && g.date ? Math.abs((b.date - g.date) / 86400000) : 30;
        
        if (taxDiff <= 10.0 && dateDiff <= 30) {
          b.matchStatus = 'Matched (Exact)';
          b.matchedIdx2b = g.id;
          g.matchStatus = 'Matched (Exact)';
          g.matchedIdxBooks = b.id;
          break;
        }
      }
    });
    
    // Phase 2: Amount Mismatch (Invoice & GSTIN matches, but amount differs > ₹10)
    books.filter(b => b.matchStatus === 'Unmatched (Books Only)').forEach(b => {
      if (!b.normGstin || !b.normInv) return;
      const g = gstr.find(g => g.normGstin === b.normGstin && g.normInv === b.normInv && g.matchedIdxBooks === -1);
      if (g) {
        b.matchStatus = 'Amount Mismatch';
        b.matchedIdx2b = g.id;
        g.matchStatus = 'Amount Mismatch';
        g.matchedIdxBooks = b.id;
      }
    });
    
    // Phase 3: Invoice/Date Mismatch (GSTIN + Amount match, Invoice No differs fuzzy)
    books.filter(b => b.matchStatus === 'Unmatched (Books Only)').forEach(b => {
      if (!b.normGstin) return;
      const matches = gstr.filter(g => g.normGstin === b.normGstin && g.matchedIdxBooks === -1);
      
      for (const g of matches) {
        const taxDiff = Math.abs(b.totalTax - g.totalTax);
        const dateDiff = b.date && g.date ? Math.abs((b.date - g.date) / 86400000) : 30;
        if (taxDiff <= 5.0 && dateDiff <= 15) {
          b.matchStatus = 'Matched (Fuzzy)';
          b.matchedIdx2b = g.id;
          g.matchStatus = 'Matched (Fuzzy)';
          g.matchedIdxBooks = b.id;
          break;
        }
      }
    });
    
    // Compile final results lists
    const matched = [];
    const mismatches = [];
    const booksOnly = books.filter(b => b.matchStatus === 'Unmatched (Books Only)');
    const gstrOnly = gstr.filter(g => g.matchStatus === 'Unmatched (GSTR-2B Only)');
    
    books.filter(b => b.matchStatus !== 'Unmatched (Books Only)').forEach(b => {
      const g = gstr.find(gstr_item => gstr_item.id === b.matchedIdx2b);
      const rec = {
        gstin: b.gstin || g.gstin,
        vendorBooks: b.particulars,
        vendor2b: g.supplierName,
        invBooks: b.invoiceNo,
        inv2b: g.invoiceNo,
        dateBooks: b.date,
        date2b: g.date,
        taxableBooks: b.taxable,
        taxable2b: g.taxable,
        taxBooks: b.totalTax,
        tax2b: g.totalTax,
        difference: b.totalTax - g.totalTax,
        status: b.matchStatus
      };
      if (b.matchStatus === 'Amount Mismatch') mismatches.push(rec);
      else matched.push(rec);
    });
    
    setRecoData({
      matched,
      mismatches,
      booksOnly,
      gstrOnly,
      all: [
        ...matched.map(r => ({ ...r, displayStatus: 'Matched' })),
        ...mismatches.map(r => ({ ...r, displayStatus: 'Amount Mismatch' })),
        ...booksOnly.map(r => ({ gstin: r.gstin, vendorBooks: r.particulars, invBooks: r.invoiceNo, dateBooks: r.date, taxBooks: r.totalTax, status: 'Books Only', displayStatus: 'Books Only' })),
        ...gstrOnly.map(r => ({ gstin: r.gstin, vendor2b: r.supplierName, inv2b: r.invoiceNo, date2b: r.date, tax2b: r.totalTax, status: 'GSTR-2B Only', displayStatus: 'GSTR-2B Only' }))
      ]
    });
    setProcessing(false);
  };

  // Generate Vendor Defaulters follow-up message
  const generateFollowUpText = (vendorGstin, vendorName, invoices) => {
    const list = invoices.slice(0, 3).map(inv => `Inv #${inv.invoiceNo} dated ${inv.date ? new Date(inv.date).toLocaleDateString('en-IN') : '—'} (GST: ₹${inv.totalTax.toFixed(2)})`).join(', ');
    const countText = invoices.length > 3 ? ` and ${invoices.length - 3} other invoices` : '';
    const totalMissingTax = invoices.reduce((s, x) => s + x.totalTax, 0);
    
    return `Dear ${vendorName}, we have recorded purchases of total value ₹${totalMissingTax.toFixed(2)} in our books for FY 2025-26, but the matching credits are missing from our GSTR-2B. Invoices: ${list}${countText}. Kindly file your GSTR-1 as soon as possible so we can claim our input credit. Thank you, [Auditor/Client Name].`;
  };

  // Export to beautifully styled SheetJS Excel file
  const exportToExcel = () => {
    if (!recoData) return;
    const wb = XLSX.utils.book_new();
    
    // Helper to format date
    const formatDate = d => d ? new Date(d).toLocaleDateString('en-IN') : '';
    
    // Tab 1: Dashboard summary
    const dashRows = [
      ['GST Input Tax Credit Reconciliation Dashboard'],
      [`Generated: ${new Date().toLocaleDateString('en-IN')}`],
      [],
      ['Summary KPI', 'Amount (INR)'],
      ['Matched ITC', recoData.matched.reduce((s, x) => s + x.taxBooks, 0)],
      ['Amount Mismatches', recoData.mismatches.reduce((s, x) => s + x.taxBooks, 0)],
      ['Missing in GSTR-2B (Defaulters)', recoData.booksOnly.reduce((s, x) => s + x.totalTax, 0)],
      ['GSTR-2B Only (Unrecorded)', recoData.gstrOnly.reduce((s, x) => s + x.totalTax, 0)]
    ];
    const wsDash = XLSX.utils.aoa_to_sheet(dashRows);
    XLSX.utils.book_append_sheet(wb, wsDash, 'Summary Dashboard');
    
    // Tab 2: Matched Invoices
    const matchedRows = [['GSTIN', 'Vendor Name (Books)', 'Vendor Name (GSTR-2B)', 'Invoice No (Books)', 'Invoice No (GSTR-2B)', 'Date (Books)', 'Date (GSTR-2B)', 'Taxable (Books)', 'Taxable (GSTR-2B)', 'Tax (Books)', 'Tax (GSTR-2B)', 'Difference']];
    recoData.matched.forEach(r => {
      matchedRows.push([r.gstin, r.vendorBooks, r.vendor2b, r.invBooks, r.inv2b, formatDate(r.dateBooks), formatDate(r.date2b), r.taxableBooks, r.taxable2b, r.taxBooks, r.tax2b, r.difference]);
    });
    const wsMatched = XLSX.utils.aoa_to_sheet(matchedRows);
    XLSX.utils.book_append_sheet(wb, wsMatched, 'Matched Invoices');
    
    // Tab 3: Amount Mismatches
    const mismatchRows = [['GSTIN', 'Vendor Name (Books)', 'Vendor Name (GSTR-2B)', 'Invoice No (Books)', 'Invoice No (GSTR-2B)', 'Date (Books)', 'Date (GSTR-2B)', 'Taxable (Books)', 'Taxable (GSTR-2B)', 'Tax (Books)', 'Tax (GSTR-2B)', 'Difference']];
    recoData.mismatches.forEach(r => {
      mismatchRows.push([r.gstin, r.vendorBooks, r.vendor2b, r.invBooks, r.inv2b, formatDate(r.dateBooks), formatDate(r.date2b), r.taxableBooks, r.taxable2b, r.taxBooks, r.tax2b, r.difference]);
    });
    const wsMismatch = XLSX.utils.aoa_to_sheet(mismatchRows);
    XLSX.utils.book_append_sheet(wb, wsMismatch, 'Amount Mismatches');
    
    // Tab 4: Books Only (Defaulters)
    const booksOnlyRows = [['GSTIN', 'Vendor Name', 'Invoice No', 'Date', 'Taxable Value', 'Total Tax']];
    recoData.booksOnly.forEach(r => {
      booksOnlyRows.push([r.gstin, r.particulars, r.invoiceNo, formatDate(r.date), r.taxable, r.totalTax]);
    });
    const wsBooksOnly = XLSX.utils.aoa_to_sheet(booksOnlyRows);
    XLSX.utils.book_append_sheet(wb, wsBooksOnly, 'Books Only (Defaulters)');
    
    // Tab 5: GSTR-2B Only
    const gstrOnlyRows = [['GSTIN', 'Supplier Name', 'Invoice No', 'Date', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'Cess', 'Total Tax', 'Filing Date', 'GSTR-2B Month']];
    recoData.gstrOnly.forEach(r => {
      gstrOnlyRows.push([r.gstin, r.supplierName, r.invoiceNo, formatDate(r.date), r.taxable, r.igst, r.cgst, r.sgst, r.cess, r.totalTax, formatDate(r.filingDate), r.monthName]);
    });
    const wsGstrOnly = XLSX.utils.aoa_to_sheet(gstrOnlyRows);
    XLSX.utils.book_append_sheet(wb, wsGstrOnly, 'GSTR-2B Only');
    
    // Tab 6: Defaulters summary
    const defaultersMap = {};
    recoData.booksOnly.forEach(r => {
      if (!defaultersMap[r.gstin]) {
        defaultersMap[r.gstin] = { name: r.particulars, count: 0, tax: 0 };
      }
      defaultersMap[r.gstin].count++;
      defaultersMap[r.gstin].tax += r.totalTax;
    });
    const defSummaryRows = [['GSTIN', 'Vendor Name', 'Missing Invoice Count', 'Total Missing Tax (ITC)']];
    Object.keys(defaultersMap).forEach(gstin => {
      const v = defaultersMap[gstin];
      defSummaryRows.push([gstin, v.name, v.count, v.tax]);
    });
    const wsDefSummary = XLSX.utils.aoa_to_sheet(defSummaryRows);
    XLSX.utils.book_append_sheet(wb, wsDefSummary, 'Defaulters Summary');
    
    XLSX.writeFile(wb, 'GST_Books_vs_GSTR2B_Reconciliation.xlsx');
  };

  // Group booksOnly for Defaulters List
  const defaulterList = [];
  if (recoData) {
    const map = {};
    recoData.booksOnly.forEach(r => {
      if (!map[r.gstin]) {
        map[r.gstin] = { gstin: r.gstin, name: r.particulars, invoices: [] };
      }
      map[r.gstin].invoices.push(r);
    });
    Object.keys(map).forEach(k => {
      const item = map[k];
      item.totalTax = item.invoices.reduce((s, x) => s + x.totalTax, 0);
      defaulterList.push(item);
    });
    defaulterList.sort((a, b) => b.totalTax - a.totalTax);
  }

  // Filter records in display grid
  const displayRecords = recoData ? recoData.all.filter(r => {
    const matchSearch = (r.gstin || '').toLowerCase().includes(search.toLowerCase()) ||
                        (r.vendorBooks || r.vendor2b || '').toLowerCase().includes(search.toLowerCase()) ||
                        (r.invBooks || r.inv2b || '').toLowerCase().includes(search.toLowerCase());
    if (!matchSearch) return false;
    
    if (statusFilter === 'ALL') return true;
    return r.displayStatus === statusFilter;
  }) : [];

  return (
    <div style={{
      fontFamily: "'DM Sans','Trebuchet MS',sans-serif",
      background: C.bg, minHeight: '100vh', color: C.text, padding: 16, boxSizing: 'border-box'
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');
        *,*::before,*::after{box-sizing:border-box}
        tr:hover>td{background:rgba(0,200,150,.03)!important}
        button:hover{opacity:.85;transform:translateY(-1px)} button{transition:all .15s}
        .btn-act { background: ${C.accent}; color: #000; font-weight:700; border:none; padding:8px 18px; border-radius:8px; cursor:pointer; }
        .btn-sec { background: ${C.card}; border: 1px solid ${C.border}; color: ${C.text}; font-weight:600; padding:8px 18px; border-radius:8px; cursor:pointer; }
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
              GST Reconciliation Suite 2.0 (Books vs GSTR-2B)
            </h1>
            <p style={{ margin: '3px 0 0', fontSize: 11, color: C.muted }}>
              Invoice-level matching engine & Defaulter generator
            </p>
          </div>
        </div>
        {recoData && (
          <button onClick={exportToExcel} className="btn-act">
            📥 Download Reconciliation Report
          </button>
        )}
      </div>

      {/* File Upload zones */}
      {!recoData && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14, marginBottom: 14 }}>
          {/* Books Upload */}
          <div onClick={() => booksFileRef.current?.click()} style={{
            background: booksFiles.length > 0 ? '#0B2117' : C.surface,
            border: `2px dashed ${booksFiles.length > 0 ? C.accent : C.border}`,
            borderRadius: 12, padding: '24px 20px', textAlign: 'center', cursor: 'pointer', transition: 'all .2s'
          }}>
            <input ref={booksFileRef} type="file" accept=".xlsx,.xls" multiple hidden onChange={handleBooksFiles} />
            <div style={{ fontSize: 28, marginBottom: 6 }}>📚</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: booksFiles.length > 0 ? C.accent : C.text }}>
              {booksFiles.length > 0 ? `Loaded ${booksFiles.length} Tally Register file(s)` : 'Upload Books Purchases (Tally Registers)'}
            </div>
            <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>
              Purchase Register · Debit Note · Journal Register (xls/xlsx)
            </div>
            {booksRecords.length > 0 && (
              <div style={{ fontSize: 10, color: C.accent, marginTop: 6, fontWeight: 600 }}>
                {booksRecords.length} records parsed successfully
              </div>
            )}
          </div>

          {/* GSTR-2B Upload */}
          <div onClick={() => gstrFileRef.current?.click()} style={{
            background: gstrFiles.length > 0 ? '#092138' : C.surface,
            border: `2px dashed ${gstrFiles.length > 0 ? C.gross : C.border}`,
            borderRadius: 12, padding: '24px 20px', textAlign: 'center', cursor: 'pointer', transition: 'all .2s'
          }}>
            <input ref={gstrFileRef} type="file" accept=".xlsx,.xls" multiple hidden onChange={handleGstrFiles} />
            <div style={{ fontSize: 28, marginBottom: 6 }}>🌍</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: gstrFiles.length > 0 ? C.gross : C.text }}>
              {gstrFiles.length > 0 ? `Loaded ${gstrFiles.length} GSTR-2B file(s)` : 'Upload GSTR-2B Annual Summary Excel'}
            </div>
            <div style={{ fontSize: 11, color: C.muted, marginTop: 4 }}>
              GSTR-2B download containing B2B & CDNR sheets
            </div>
            {gstrRecords.length > 0 && (
              <div style={{ fontSize: 10, color: C.gross, marginTop: 6, fontWeight: 600 }}>
                {gstrRecords.length} records parsed successfully
              </div>
            )}
          </div>
        </div>
      )}

      {/* Trigger Matching Button */}
      {!recoData && (
        <div style={{ textAlign: 'center', marginBottom: 20 }}>
          <button
            onClick={runReconciliation}
            disabled={booksRecords.length === 0 || gstrRecords.length === 0 || processing}
            className="btn-act"
            style={{
              padding: '12px 30px', fontSize: 15, borderRadius: 30,
              opacity: (booksRecords.length === 0 || gstrRecords.length === 0) ? 0.5 : 1,
              cursor: (booksRecords.length === 0 || gstrRecords.length === 0) ? 'not-allowed' : 'pointer'
            }}
          >
            {processing ? '⏳ Matching Invoices...' : '⚡ Match Invoices & Generate Report'}
          </button>
        </div>
      )}

      {/* Reconciliation Output */}
      {recoData && (
        <div>
          {/* KPI Dashboard */}
          <div style={{
            display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))',
            gap: 12, marginBottom: 14
          }}>
            {/* KPI 1 */}
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 14 }}>
              <div style={{ fontSize: 10, color: C.muted, textTransform: 'uppercase', fontWeight: 600 }}>Matched Invoices</div>
              <div style={{ fontSize: 20, fontWeight: 800, color: C.taxable, margin: '4px 0' }}>{recoData.matched.length}</div>
              <div style={{ fontSize: 11, color: C.muted }}>Tax: {fmtINR(recoData.matched.reduce((s, x) => s + x.taxBooks, 0))}</div>
            </div>
            {/* KPI 2 */}
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 14 }}>
              <div style={{ fontSize: 10, color: C.muted, textTransform: 'uppercase', fontWeight: 600 }}>Amount Mismatches</div>
              <div style={{ fontSize: 20, fontWeight: 800, color: C.warn, margin: '4px 0' }}>{recoData.mismatches.length}</div>
              <div style={{ fontSize: 11, color: C.muted }}>Tax: {fmtINR(recoData.mismatches.reduce((s, x) => s + x.taxBooks, 0))}</div>
            </div>
            {/* KPI 3 */}
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 14 }}>
              <div style={{ fontSize: 10, color: C.muted, textTransform: 'uppercase', fontWeight: 600 }}>Books Only (Missing 2B)</div>
              <div style={{ fontSize: 20, fontWeight: 800, color: C.red, margin: '4px 0' }}>{recoData.booksOnly.length}</div>
              <div style={{ fontSize: 11, color: C.muted }}>Tax: {fmtINR(recoData.booksOnly.reduce((s, x) => s + x.totalTax, 0))}</div>
            </div>
            {/* KPI 4 */}
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: 14 }}>
              <div style={{ fontSize: 10, color: C.muted, textTransform: 'uppercase', fontWeight: 600 }}>GSTR-2B Only (Unrecorded)</div>
              <div style={{ fontSize: 20, fontWeight: 800, color: C.gross, margin: '4px 0' }}>{recoData.gstrOnly.length}</div>
              <div style={{ fontSize: 11, color: C.muted }}>Tax: {fmtINR(recoData.gstrOnly.reduce((s, x) => s + x.totalTax, 0))}</div>
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 14 }}>
            {/* Live Data Grid */}
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, overflow: 'hidden' }}>
              <div style={{
                padding: '14px 20px', background: C.surface, borderBottom: `1px solid ${C.border}`,
                display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8
              }}>
                <div>
                  <h3 style={{ margin: 0, fontSize: 14, fontWeight: 700 }}>🔍 Reconciliation Details</h3>
                </div>
                <div style={{ display: 'flex', gap: 8 }}>
                  <input
                    type="text"
                    placeholder="Search invoice/vendor..."
                    value={search}
                    onChange={e => setSearch(e.target.value)}
                    style={{
                      background: C.bg, border: `1px solid ${C.border}`, borderRadius: 6,
                      padding: '4px 10px', color: C.text, fontSize: 12, outline: 'none'
                    }}
                  />
                  <select
                    value={statusFilter}
                    onChange={e => setStatusFilter(e.target.value)}
                    style={{
                      background: C.bg, border: `1px solid ${C.border}`, borderRadius: 6,
                      padding: '4px 10px', color: C.text, fontSize: 12, outline: 'none', cursor: 'pointer'
                    }}
                  >
                    <option value="ALL">All Statuses</option>
                    <option value="Matched">Matched</option>
                    <option value="Amount Mismatch">Amount Mismatch</option>
                    <option value="Books Only">Books Only</option>
                    <option value="GSTR-2B Only">GSTR-2B Only</option>
                  </select>
                </div>
              </div>

              <div style={{ overflowX: 'auto', maxHeight: '55vh' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: '#050D18', borderBottom: `1px solid ${C.border}` }}>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>GSTIN</th>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>Invoice No</th>
                      <th style={{ padding: '8px 12px', textAlign: 'left', color: C.muted }}>Vendor / Supplier</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>Tax (Books)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'right', color: C.muted }}>Tax (2B)</th>
                      <th style={{ padding: '8px 12px', textAlign: 'center', color: C.muted }}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {displayRecords.slice(0, 100).map((r, i) => {
                      const st = r.displayStatus;
                      const statusColor = st === 'Matched' ? C.taxable : st === 'Amount Mismatch' ? C.warn : st === 'Books Only' ? C.red : C.gross;
                      return (
                        <tr key={i} style={{ borderBottom: `1px solid ${C.border}` }}>
                          <td style={{ padding: '8px 12px', fontFamily: 'monospace' }}>{r.gstin}</td>
                          <td style={{ padding: '8px 12px', fontWeight: 600 }}>{r.invBooks || r.inv2b}</td>
                          <td style={{ padding: '8px 12px' }}>{r.vendorBooks || r.vendor2b}</td>
                          <td style={{ padding: '8px 12px', textAlign: 'right' }}>{fmtINR(r.taxBooks)}</td>
                          <td style={{ padding: '8px 12px', textAlign: 'right' }}>{fmtINR(r.tax2b)}</td>
                          <td style={{ padding: '8px 12px', textAlign: 'center' }}>
                            <span style={{
                              padding: '2px 6px', borderRadius: 4, fontSize: 10, fontWeight: 700,
                              background: statusColor + '15', color: statusColor, border: `1px solid ${statusColor}30`
                            }}>{st}</span>
                          </td>
                        </tr>
                      );
                    })}
                    {displayRecords.length === 0 && (
                      <tr>
                        <td colSpan={6} style={{ padding: '24px', textAlign: 'center', color: C.muted }}>No matching records found.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
              <div style={{ padding: 10, background: C.surface, fontSize: 10, color: C.muted, textAlign: 'center' }}>
                Showing first {Math.min(100, displayRecords.length)} of {displayRecords.length} records.
              </div>
            </div>

            {/* Vendor Defaulters & Follow-ups */}
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 16 }}>
              <h3 style={{ margin: '0 0 10px', fontSize: 14, fontWeight: 700, color: C.red }}>
                ⚠️ Defaulters Follow-up Generator
              </h3>
              <p style={{ margin: '0 0 14px', fontSize: 11, color: C.muted }}>
                Grouped vendors with Tally entries missing from GSTR-2B. Click message to copy.
              </p>
              
              <div style={{ overflowY: 'auto', maxHeight: '60vh', display: 'flex', flexDirection: 'column', gap: 10 }}>
                {defaulterList.map((def, idx) => {
                  const msg = generateFollowUpText(def.gstin, def.name, def.invoices);
                  return (
                    <div key={idx} style={{
                      background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: 10
                    }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 4 }}>
                        <span style={{ fontSize: 12, fontWeight: 700 }}>{def.name}</span>
                        <span style={{ fontSize: 10, padding: '1px 6px', borderRadius: 10, background: '#301010', color: C.red, fontWeight: 700 }}>
                          ₹{def.totalTax.toLocaleString('en-IN', { maximumFractionDigits: 0 })}
                        </span>
                      </div>
                      <div style={{ fontSize: 10, color: C.muted, marginBottom: 6 }}>
                        GSTIN: {def.gstin} · {def.invoices.length} invoices missing
                      </div>
                      
                      <div
                        onClick={() => {
                          navigator.clipboard.writeText(msg);
                          alert(`Copied text for ${def.name}!`);
                        }}
                        style={{
                          background: C.bg, border: `1px solid ${C.border}`, borderRadius: 6,
                          padding: 8, fontSize: 10, color: C.text, cursor: 'pointer',
                          maxHeight: 70, overflow: 'hidden', textOverflow: 'ellipsis',
                          transition: 'border-color .15s'
                        }}
                        title="Click to copy message"
                        onMouseEnter={e => e.currentTarget.style.borderColor = C.accent}
                        onMouseLeave={e => e.currentTarget.style.borderColor = C.border}
                      >
                        💬 "{msg}"
                      </div>
                    </div>
                  );
                })}
                {defaulterList.length === 0 && (
                  <div style={{ textAlign: 'center', padding: 20, color: C.muted, fontSize: 11 }}>
                    All vendor invoices match successfully! No defaulters.
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
