import { useState, useMemo, useCallback } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import {
  CreditCard,
  Upload,
  ArrowRightLeft,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Clock,
  Search,
  Download,
  Copy,
  Check,
  RotateCcw,
  FileSpreadsheet,
  Layers,
  Sparkles,
  ArrowLeft,
  ShieldAlert,
  ShieldCheck
} from 'lucide-react';

// Demo datasets for one-click testing (Fully fictional/hypothetical sample data)
const DEMO_UPS_DATA = [
  {
    TRANSACTIONDATE: '2099-04-17 12:18:27',
    CLIENTAPPLICATIONID: 114,
    NAME: 'Alpha University',
    TRANSACTIONPURPOSEID: 267,
    PURPOSENAME: 'Student Exam form Submission',
    AMOUNT: 645,
    CLIENTAPPTRANSCATIONREFERENCENUMBER: '20990417121827889',
    TRANSCATIONREFERENCENUMBER: '9206290301625',
    DESCRIPTION: 'Online Exam Form Submission',
    OTHERPARAMETERS: 'goLang',
    REF1: '2099010000000001',
    REF2: 3360,
    REF3: 71,
    REF4: 159258,
    REF5: 78,
    CLIENTSYNCDATE: '2099-04-17 14:30:00',
    UPSSTATUS: 'reconciled'
  },
  {
    TRANSACTIONDATE: '2099-04-17 14:53:23',
    CLIENTAPPLICATIONID: 114,
    NAME: 'Alpha University',
    TRANSACTIONPURPOSEID: 267,
    PURPOSENAME: 'Student Exam form Submission',
    AMOUNT: 965,
    CLIENTAPPTRANSCATIONREFERENCENUMBER: '20990417025323024',
    TRANSCATIONREFERENCENUMBER: 'NA',
    DESCRIPTION: 'Online Exam Form Submission',
    OTHERPARAMETERS: 'goLang',
    REF1: '2099010000000002',
    REF2: 21876,
    REF3: 45,
    REF4: 159410,
    REF5: 112,
    CLIENTSYNCDATE: null,
    UPSSTATUS: 'initiated'
  },
  {
    TRANSACTIONDATE: '2099-04-17 21:00:10',
    CLIENTAPPLICATIONID: 114,
    NAME: 'Alpha University',
    TRANSACTIONPURPOSEID: 267,
    PURPOSENAME: 'Student Exam form Submission',
    AMOUNT: 785,
    CLIENTAPPTRANSCATIONREFERENCENUMBER: '20990417085945959',
    TRANSCATIONREFERENCENUMBER: 'NA',
    DESCRIPTION: 'Online Exam Form Submission',
    OTHERPARAMETERS: 'goLang',
    REF1: '2099010000000003',
    REF2: 3514,
    REF3: 82,
    REF4: 159820,
    REF5: 94,
    CLIENTSYNCDATE: null,
    UPSSTATUS: 'initiated'
  },
  {
    TRANSACTIONDATE: '2099-04-18 12:36:26',
    CLIENTAPPLICATIONID: 114,
    NAME: 'Alpha University',
    TRANSACTIONPURPOSEID: 267,
    PURPOSENAME: 'Student Exam form Submission',
    AMOUNT: 1440,
    CLIENTAPPTRANSCATIONREFERENCENUMBER: '20990418123626815',
    TRANSCATIONREFERENCENUMBER: 'NA',
    DESCRIPTION: 'Online Exam Form Submission',
    OTHERPARAMETERS: 'goLang',
    REF1: '2099010000000004',
    REF2: 30462,
    REF3: 118,
    REF4: 160112,
    REF5: 412,
    CLIENTSYNCDATE: null,
    UPSSTATUS: 'Failed'
  },
  {
    TRANSACTIONDATE: '2099-04-22 11:12:16',
    CLIENTAPPLICATIONID: 114,
    NAME: 'Alpha University',
    TRANSACTIONPURPOSEID: 267,
    PURPOSENAME: 'Student Exam form Submission',
    AMOUNT: 1365,
    CLIENTAPPTRANSCATIONREFERENCENUMBER: '20990422111216521',
    TRANSCATIONREFERENCENUMBER: 'NA',
    DESCRIPTION: 'Online Exam Form Submission',
    OTHERPARAMETERS: 'goLang',
    REF1: '2099010000000005',
    REF2: 30950,
    REF3: 130,
    REF4: 161278,
    REF5: 498,
    CLIENTSYNCDATE: null,
    UPSSTATUS: 'Failed'
  },
  {
    TRANSACTIONDATE: '2099-04-24 20:43:24',
    CLIENTAPPLICATIONID: 114,
    NAME: 'Alpha University',
    TRANSACTIONPURPOSEID: 213,
    PURPOSENAME: 'Revaluation_New',
    AMOUNT: 1325,
    CLIENTAPPTRANSCATIONREFERENCENUMBER: '20990424084324782',
    TRANSCATIONREFERENCENUMBER: 'NA',
    DESCRIPTION: 'Application for Revaluation',
    OTHERPARAMETERS: 'goLang',
    REF1: '2099010000000006',
    REF2: 18844,
    REF3: 92,
    REF4: 162004,
    REF5: 120,
    CLIENTSYNCDATE: null,
    UPSSTATUS: 'Failed'
  },
  {
    TRANSACTIONDATE: '2099-04-25 09:15:40',
    CLIENTAPPLICATIONID: 114,
    NAME: 'Alpha University',
    TRANSACTIONPURPOSEID: 212,
    PURPOSENAME: 'Recounting',
    AMOUNT: 280,
    CLIENTAPPTRANSCATIONREFERENCENUMBER: '20990425091540112',
    TRANSCATIONREFERENCENUMBER: 'NA',
    DESCRIPTION: 'Application for Recounting',
    OTHERPARAMETERS: 'goLang',
    REF1: '2099010000000007',
    REF2: 19230,
    REF3: 14,
    REF4: 162540,
    REF5: 35,
    CLIENTSYNCDATE: null,
    UPSSTATUS: 'Failed'
  },
  {
    TRANSACTIONDATE: '2099-04-26 16:20:10',
    CLIENTAPPLICATIONID: 114,
    NAME: 'Alpha University',
    TRANSACTIONPURPOSEID: 267,
    PURPOSENAME: 'Student Exam form Submission',
    AMOUNT: 850,
    CLIENTAPPTRANSCATIONREFERENCENUMBER: '20990426162010884',
    TRANSCATIONREFERENCENUMBER: 'NA',
    DESCRIPTION: 'Online Exam Form Submission',
    OTHERPARAMETERS: 'goLang',
    REF1: '2099010000000008',
    REF2: 7412,
    REF3: 50,
    REF4: 163012,
    REF5: 80,
    CLIENTSYNCDATE: null,
    UPSSTATUS: 'initiated'
  }
];

const DEMO_GATEWAY_DATA = [
  // SBI ePay rows (Fictionalized)
  {
    'TRANSACTION REQUEST DATE & TIME': '2099-04-17 21:00:10',
    'TRANSACTION SUCCESS DATE & TIME': '2099-04-17 21:01:11',
    ATRN: '5865985978227',
    'MERCHANT ORDER NO': '20990417085945959',
    'GATEWAY TRACE NUMBER': '121760878998',
    'MERCHANT ORDER AMOUNT': 785,
    STATUS: 'SUCCESS',
    'STATUS DESCRIPTION': 'Transaction Paid Out',
    'GATEWAY NAME': 'State Bank of India',
    'PAY MODE CODE': 'UPI',
    'CIN NUMBER': '10032442099041700027'
  },
  {
    'TRANSACTION REQUEST DATE & TIME': '2099-04-17 14:53:23',
    'TRANSACTION SUCCESS DATE & TIME': '2099-04-17 14:54:10',
    ATRN: '6012498511234',
    'MERCHANT ORDER NO': '20990417025323024',
    'GATEWAY TRACE NUMBER': '610773437993',
    'MERCHANT ORDER AMOUNT': 965,
    STATUS: 'SUCCESS',
    'STATUS DESCRIPTION': 'Transaction Paid Out',
    'GATEWAY NAME': 'State Bank of India',
    'PAY MODE CODE': 'UPI',
    'CIN NUMBER': '10032442099041700041'
  },
  {
    'TRANSACTION REQUEST DATE & TIME': '2099-04-18 12:36:26',
    'TRANSACTION SUCCESS DATE & TIME': '2099-04-18 12:38:02',
    ATRN: '7100452389102',
    'MERCHANT ORDER NO': '20990418123626815',
    'GATEWAY TRACE NUMBER': '121760992340',
    'MERCHANT ORDER AMOUNT': 1440,
    STATUS: 'SUCCESS',
    'STATUS DESCRIPTION': 'Transaction Paid Out',
    'GATEWAY NAME': 'State Bank of India',
    'PAY MODE CODE': 'NetBanking',
    'CIN NUMBER': '10032442099041800088'
  },
  // ATOM rows (Fictionalized)
  {
    'Merchant Name': 'SAMPLE UNIVERSITY',
    'Merchant ID': '600001',
    'Atom Txn ID': '11000353263741',
    'Merchant Txn ID': '20990422111216521',
    Amount: '1365.00',
    'Txn Date': '22-Apr-2099 11:15:30',
    'Txn Status': 'OK',
    'Recon Status': 'RS',
    'Bank Ref No': '606879858976',
    'Auth No.': '145136723367',
    Description: 'TRANSACTION IS SUCCESSFUL',
    Udf2: 'student.alpha@example.com',
    Udf3: '9999900001'
  },
  {
    'Merchant Name': 'SAMPLE UNIVERSITY',
    'Merchant ID': '600001',
    'Atom Txn ID': '11000360880249',
    'Merchant Txn ID': '20990424084324782',
    Amount: '1325.00',
    'Txn Date': '24-Apr-2099 20:47:09',
    'Txn Status': 'OK',
    'Recon Status': 'RS',
    'Bank Ref No': '611414513141',
    'Auth No.': '157791995061',
    Description: 'TRANSACTION IS SUCCESSFUL',
    Udf2: 'student.beta@example.com',
    Udf3: '9999900002'
  },
  {
    'Merchant Name': 'SAMPLE UNIVERSITY',
    'Merchant ID': '600001',
    'Atom Txn ID': '11000361991204',
    'Merchant Txn ID': '20990425091540112',
    Amount: '280.00',
    'Txn Date': '25-Apr-2099 09:18:22',
    'Txn Status': 'FAILED',
    'Recon Status': 'RNS',
    'Bank Ref No': '611599812341',
    'Auth No.': '',
    Description: 'TRANSACTION FAILED / USER CANCELLED',
    Udf2: 'student.gamma@example.com',
    Udf3: '9999900003'
  }
];

export default function PaymentReconciliationPage() {
  // File upload states
  const [upsFile, setUpsFile] = useState(null);
  const [upsRawRows, setUpsRawRows] = useState([]);
  const [upsSheets, setUpsSheets] = useState([]);
  const [selectedUpsSheet, setSelectedUpsSheet] = useState('');
  const [upsWorkbook, setUpsWorkbook] = useState(null);

  const [gatewayFile, setGatewayFile] = useState(null);
  const [gatewayRawRows, setGatewayRawRows] = useState([]);
  const [gatewaySheets, setGatewaySheets] = useState([]);
  const [selectedGatewaySheet, setSelectedGatewaySheet] = useState('');
  const [gatewayWorkbook, setGatewayWorkbook] = useState(null);

  // Gateway mode: 'auto' | 'atom' | 'epay'
  const [gatewayType, setGatewayType] = useState('auto');
  const [detectedGateway, setDetectedGateway] = useState(null);

  // Status & Filtering
  const [statusFilter, setStatusFilter] = useState('ACTION_NEEDED'); // 'ALL' | 'ACTION_NEEDED' | 'RECONCILED' | 'BOTH_FAILED' | 'DISCREPANCY' | 'MISSING_IN_GATEWAY'
  const [purposeFilter, setPurposeFilter] = useState('ALL');
  const [gatewaySourceFilter, setGatewaySourceFilter] = useState('ALL');
  const [searchQuery, setSearchQuery] = useState('');
  const [, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready for upload. Drop your UPS Report and Gateway (ATOM / SBI ePay) file.');
  const [statusType, setStatusType] = useState('info'); // 'info' | 'success' | 'warning' | 'error'

  // Pagination & Density
  const [pageSize, setPageSize] = useState(50);
  const [page, setPage] = useState(0);
  const [tableDensity, setTableDensity] = useState('normal'); // 'compact' | 'normal' | 'comfortable'

  // Clipboard copy state
  const [copiedType, setCopiedType] = useState(null);

  // Helper normalizers
  const normalizeText = (text) => String(text || '').trim().toUpperCase();
  const normalizeRef = (val) => String(val || '').trim().replace(/[^a-zA-Z0-9]/g, '');

  const isSuccessStatus = (statusStr) => {
    const s = normalizeText(statusStr);
    return (
      s === 'SUCCESS' ||
      s === 'COMPLETED' ||
      s === 'RECONCILED' ||
      s.includes('RECONCILED') ||
      s === 'OK' ||
      s === 'PAID' ||
      s === 'PAID OUT' ||
      s === 'TRANSACTION PAID OUT' ||
      s === 'SETTLED' ||
      s === 'RS'
    );
  };

  const isFailedStatus = (statusStr) => {
    const s = normalizeText(statusStr);
    return (
      s === 'FAILED' ||
      s === 'FAIL' ||
      s === 'INITIATED' ||
      s === 'PENDING' ||
      s === 'CANCELLED' ||
      s === 'DECLINED' ||
      s === 'AUTO REVERSAL' ||
      s === 'AUTO REFUND' ||
      s === 'RNS' ||
      s === 'ABORTED' ||
      s === 'TIMEOUT' ||
      s === 'AUTHORIZATION FAILED'
    );
  };

  // Inspect gateway headers to auto-detect gateway provider
  const detectGatewayFromRows = useCallback((rows) => {
    if (!rows || rows.length === 0) return 'UNKNOWN';
    const first = rows[0];
    const keys = Object.keys(first).map((k) => k.toLowerCase().replace(/[^a-z0-9]/g, ''));

    const hasAtom = keys.some(
      (k) =>
        k.includes('atomtxnid') ||
        k.includes('merchanttxnid') ||
        k.includes('atommid') ||
        k.includes('atomtid') ||
        k.includes('discriminator')
    );
    if (hasAtom) return 'ATOM';

    const hasEpay = keys.some(
      (k) =>
        k.includes('atrn') ||
        k.includes('merchantorderno') ||
        k.includes('gatewaytracenumber') ||
        k.includes('cinnumber') ||
        k.includes('paymodecode')
    );
    if (hasEpay) return 'SBI_EPAY';

    return 'UNKNOWN';
  }, []);

  // Parse Excel / CSV into JSON array
  const parseFileBuffer = (buffer) => {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true, defval: '' });
    return wb;
  };

  // Handle UPS file upload
  const handleUpsUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setIsProcessing(true);
    setStatusMsg('Parsing University Payment System (UPS) report...');
    setStatusType('info');

    try {
      const buffer = await file.arrayBuffer();
      const wb = parseFileBuffer(buffer);
      setUpsWorkbook(wb);
      setUpsSheets(wb.SheetNames);
      setUpsFile(file);

      // Default to first sheet or sheet with "UPS" / "Failed"
      let defaultSheet = wb.SheetNames[0];
      const match = wb.SheetNames.find((s) => s.toLowerCase().includes('ups') || s.toLowerCase().includes('failed') || s.toLowerCase().includes('transaction'));
      if (match) defaultSheet = match;
      setSelectedUpsSheet(defaultSheet);

      const rows = XLSX.utils.sheet_to_json(wb.Sheets[defaultSheet], { defval: '' });
      setUpsRawRows(rows);
      setStatusMsg(`Loaded ${rows.length} records from UPS Report (${file.name})!`);
      setStatusType('success');
    } catch (err) {
      console.error('UPS Parse Error:', err);
      setStatusMsg(`Failed to parse UPS report: ${err.message}`);
      setStatusType('error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Handle UPS sheet change
  const handleUpsSheetChange = (sheetName) => {
    if (!upsWorkbook || !sheetName) return;
    setSelectedUpsSheet(sheetName);
    const rows = XLSX.utils.sheet_to_json(upsWorkbook.Sheets[sheetName], { defval: '' });
    setUpsRawRows(rows);
    setStatusMsg(`Loaded ${rows.length} records from UPS Sheet: ${sheetName}`);
    setStatusType('info');
  };

  // Handle Gateway file upload
  const handleGatewayUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setIsProcessing(true);
    setStatusMsg('Parsing Payment Gateway settlement report...');
    setStatusType('info');

    try {
      const buffer = await file.arrayBuffer();
      const wb = parseFileBuffer(buffer);
      setGatewayWorkbook(wb);
      setGatewaySheets(wb.SheetNames);
      setGatewayFile(file);

      let defaultSheet = wb.SheetNames[0];
      const match = wb.SheetNames.find((s) => s.toLowerCase().includes('epay') || s.toLowerCase().includes('atom') || s.toLowerCase().includes('success') || s.toLowerCase().includes('mis'));
      if (match) defaultSheet = match;
      setSelectedGatewaySheet(defaultSheet);

      const rows = XLSX.utils.sheet_to_json(wb.Sheets[defaultSheet], { defval: '' });
      setGatewayRawRows(rows);

      const detected = detectGatewayFromRows(rows);
      setDetectedGateway(detected);
      setStatusMsg(`Loaded ${rows.length} records from Gateway Report (${file.name}) [Detected Provider: ${detected === 'SBI_EPAY' ? 'SBI ePay' : detected === 'ATOM' ? 'ATOM Technologies' : 'Generic Gateway'}]!`);
      setStatusType('success');
    } catch (err) {
      console.error('Gateway Parse Error:', err);
      setStatusMsg(`Failed to parse Gateway report: ${err.message}`);
      setStatusType('error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Handle Gateway sheet change
  const handleGatewaySheetChange = (sheetName) => {
    if (!gatewayWorkbook || !sheetName) return;
    setSelectedGatewaySheet(sheetName);
    const rows = XLSX.utils.sheet_to_json(gatewayWorkbook.Sheets[sheetName], { defval: '' });
    setGatewayRawRows(rows);
    const detected = detectGatewayFromRows(rows);
    setDetectedGateway(detected);
    setStatusMsg(`Loaded ${rows.length} records from Gateway Sheet: ${sheetName} [${detected === 'SBI_EPAY' ? 'SBI ePay' : detected === 'ATOM' ? 'ATOM' : 'Generic'}]`);
    setStatusType('info');
  };

  // Load Demo Data
  const handleLoadDemoData = () => {
    setUpsRawRows(DEMO_UPS_DATA);
    setUpsFile({ name: 'DEMO_UPS_Exam_Reval_Report.xlsx' });
    setUpsSheets(['Sheet1']);
    setSelectedUpsSheet('Sheet1');

    setGatewayRawRows(DEMO_GATEWAY_DATA);
    setGatewayFile({ name: 'DEMO_Combined_Gateways_ATOM_ePay.xlsx' });
    setGatewaySheets(['Settlement']);
    setSelectedGatewaySheet('Settlement');
    setDetectedGateway('SBI_EPAY & ATOM');

    setStatusMsg('Loaded official demonstration datasets with simulated failed UPS & successful gateway transactions!');
    setStatusType('success');
  };

  // Extract cell value helper
  const getVal = (row, ...keys) => {
    for (const k of keys) {
      if (row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') {
        return row[k];
      }
      // Case-insensitive / normalized lookup
      const target = k.toLowerCase().replace(/[^a-z0-9]/g, '');
      for (const actualKey of Object.keys(row)) {
        if (actualKey.toLowerCase().replace(/[^a-z0-9]/g, '') === target) {
          if (row[actualKey] !== undefined && row[actualKey] !== null && String(row[actualKey]).trim() !== '') {
            return row[actualKey];
          }
        }
      }
    }
    return '';
  };

  // Standardize Gateway Record
  const extractGatewayRecord = useCallback((row, forcedType = 'auto') => {
    const hasMerchantTxnId = getVal(row, 'Merchant Txn ID', 'MerchantTxnID', 'atomtxnid') !== '';
    const hasMerchantOrderNo = getVal(row, 'MERCHANT ORDER NO', 'MerchantOrderNo', 'atrn') !== '';

    let isAtom = forcedType === 'atom';

    if (forcedType === 'auto') {
      if (hasMerchantTxnId) isAtom = true;
      else if (hasMerchantOrderNo) isAtom = false;
      else if (row['Atom Txn ID'] || row['Atom TID']) isAtom = true;
      else if (row['ATRN'] || row['GATEWAY TRACE NUMBER']) isAtom = false;
      else isAtom = false; // default fallback
    }

    if (isAtom) {
      return {
        gatewayName: 'ATOM',
        merchantRef: normalizeRef(getVal(row, 'Merchant Txn ID', 'MerchantTxnID', 'Customer Acc. No.', 'Merchant Ref')),
        gatewayTxnId: String(getVal(row, 'Atom Txn ID', 'AtomTxnID')).trim(),
        bankRefNo: String(getVal(row, 'Bank Ref No', 'BankRefNo', 'Auth No.', 'AuthNo')).trim(),
        amount: Number(getVal(row, 'Amount', 'amount', 'Total Chargeable')) || 0,
        gatewayStatus: String(getVal(row, 'Txn Status', 'TxnStatus', 'Recon Status')).trim(),
        description: String(getVal(row, 'Description', 'description')).trim(),
        txnDate: String(getVal(row, 'Txn Date', 'TxnDate', 'Settlement Date')).trim(),
        customerEmail: String(getVal(row, 'Udf2', 'Email', 'email')).trim(),
        customerPhone: String(getVal(row, 'Udf3', 'Phone', 'Mobile', 'mobile')).trim(),
        raw: row
      };
    } else {
      // SBI ePay
      return {
        gatewayName: 'SBI_EPAY',
        merchantRef: normalizeRef(getVal(row, 'MERCHANT ORDER NO', 'MerchantOrderNo', 'ORDER ID', 'Order No', 'Merchant Order Number')),
        gatewayTxnId: String(getVal(row, 'ATRN', 'atrn', 'Gateway Ref')).trim(),
        bankRefNo: String(getVal(row, 'GATEWAY TRACE NUMBER', 'GatewayTraceNumber', 'CIN NUMBER', 'CinNumber')).trim(),
        amount: Number(getVal(row, 'MERCHANT ORDER AMOUNT', 'GATEWAY POSTING AMOUNT', 'PAYOUT AMOUNT', 'Amount')) || 0,
        gatewayStatus: String(getVal(row, 'STATUS', 'status', 'STATUS DESCRIPTION')).trim(),
        description: String(getVal(row, 'STATUS DESCRIPTION', 'StatusDescription', 'DESCRIPTION')).trim(),
        txnDate: String(getVal(row, 'TRANSACTION SUCCESS DATE & TIME', 'TRANSACTION REQUEST DATE & TIME', 'TRANSACTION BOOKING DATE & TIME')).trim(),
        customerEmail: String(getVal(row, 'BILLING EMAIL ADDRESS', 'Email', 'email')).trim(),
        customerPhone: String(getVal(row, 'BILLING MOBILE NUMBER', 'Mobile', 'phone')).trim(),
        raw: row
      };
    }
  }, []);

  // Reconciliation Pipeline
  const { reconciledRecords, kpiSummary, uniquePurposes } = useMemo(() => {
    if (!upsRawRows.length && !gatewayRawRows.length) {
      return {
        reconciledRecords: [],
        kpiSummary: {
          total: 0,
          actionNeededCount: 0,
          actionNeededAmount: 0,
          reconciledCount: 0,
          reconciledAmount: 0,
          bothFailedCount: 0,
          discrepancyCount: 0,
          missingInGatewayCount: 0,
          missingInUpsCount: 0
        },
        uniquePurposes: []
      };
    }

    // 1. Build Map of Gateway Records
    const gatewayMap = new Map();
    const allGatewayRecords = [];

    gatewayRawRows.forEach((row, idx) => {
      const gRec = extractGatewayRecord(row, gatewayType);
      if (gRec.merchantRef) {
        gatewayMap.set(gRec.merchantRef, gRec);
      }
      allGatewayRecords.push({ ...gRec, _origIndex: idx });
    });

    const purposes = new Set();
    const matchedGatewayRefs = new Set();
    const results = [];

    // 2. Iterate UPS Records and Link
    upsRawRows.forEach((row, idx) => {
      const upsTxnRef = normalizeRef(getVal(row, 'CLIENTAPPTRANSCATIONREFERENCENUMBER', 'ClientAppTranscationReferenceNumber', 'Client App Txn Ref', 'Merchant Order No', 'Order Id', 'Txn Ref'));
      const upsBankRef = String(getVal(row, 'TRANSCATIONREFERENCENUMBER', 'TranscationReferenceNumber', 'Bank Ref', 'ATRN')).trim();
      const prn = String(getVal(row, 'REF1', 'ref1', 'PRN', 'prn', 'Register No', 'Candidate Code', 'Student ID')).trim();
      const studentName = String(getVal(row, 'NAME', 'Student Name', 'Candidate Name', 'name')).trim();
      const purpose = String(getVal(row, 'PURPOSENAME', 'PurposeName', 'Purpose', 'Description')).trim() || 'General Fee';
      const amount = Number(getVal(row, 'AMOUNT', 'amount', 'Fee', 'Total')) || 0;
      const upsStatus = String(getVal(row, 'UPSSTATUS', 'upsstatus', 'Status', 'Payment Status')).trim() || 'Pending';
      const txnDate = String(getVal(row, 'TRANSACTIONDATE', 'TransactionDate', 'Date', 'Created At')).trim();

      if (purpose) purposes.add(purpose);

      // Find matching gateway record by merchant ref or secondary bank ref
      let gMatch = gatewayMap.get(upsTxnRef);
      if (!gMatch && upsBankRef && upsBankRef !== 'NA' && upsBankRef !== '0') {
        // Fallback match on bank ref / ATRN
        gMatch = allGatewayRecords.find((g) => g.gatewayTxnId === upsBankRef || g.bankRefNo === upsBankRef);
      }

      if (gMatch) {
        matchedGatewayRefs.add(gMatch.merchantRef);
      }

      const upsIsSuccess = isSuccessStatus(upsStatus);
      const upsIsFailed = isFailedStatus(upsStatus) || !upsIsSuccess;

      let category = 'MISSING_IN_GATEWAY';
      let remarks = 'No transaction found in payment gateway report.';

      if (gMatch) {
        const gatewayIsSuccess = isSuccessStatus(gMatch.gatewayStatus);
        const gatewayIsFailed = isFailedStatus(gMatch.gatewayStatus) || !gatewayIsSuccess;

        if (upsIsFailed && gatewayIsSuccess) {
          category = 'ACTION_NEEDED';
          remarks = 'CRITICAL: Paid in Gateway but marked Failed/Pending in UPS. Fee sync required!';
        } else if (upsIsSuccess && gatewayIsSuccess) {
          category = 'RECONCILED';
          remarks = 'Verified: Payment is fully recorded in both UPS and Gateway.';
        } else if (upsIsSuccess && gatewayIsFailed) {
          category = 'DISCREPANCY';
          remarks = 'WARNING: Marked Complete in UPS, but Gateway reports Failed / Reversed.';
        } else if (upsIsFailed && gatewayIsFailed) {
          category = 'BOTH_FAILED';
          remarks = 'Clean Fail: Both systems confirm incomplete/failed payment. No funds held.';
        } else {
          category = 'DISCREPANCY';
          remarks = `Status divergence: UPS (${upsStatus}) vs Gateway (${gMatch.gatewayStatus}).`;
        }

        // Amount mismatch check
        if (amount > 0 && gMatch.amount > 0 && Math.abs(amount - gMatch.amount) > 1) {
          remarks += ` [Amount Mismatch: UPS ₹${amount} vs Gateway ₹${gMatch.amount}]`;
        }
      }

      results.push({
        id: `ups_${idx}`,
        refNo: upsTxnRef,
        upsBankRef,
        prn,
        studentName,
        purpose,
        amount,
        upsStatus,
        upsDate: txnDate,
        gatewayMatched: !!gMatch,
        gatewayName: gMatch ? (gMatch.gatewayName === 'SBI_EPAY' ? 'SBI ePay' : 'ATOM') : '—',
        gatewayTxnId: gMatch ? gMatch.gatewayTxnId : '—',
        gatewayBankRef: gMatch ? gMatch.bankRefNo : '—',
        gatewayAmount: gMatch ? gMatch.amount : 0,
        gatewayStatus: gMatch ? gMatch.gatewayStatus : 'NOT_FOUND',
        gatewayDesc: gMatch ? gMatch.description : '—',
        gatewayDate: gMatch ? gMatch.txnDate : '—',
        category,
        remarks,
        rawUps: row,
        rawGateway: gMatch ? gMatch.raw : null
      });
    });

    // 3. Check for Gateway records missing in UPS
    allGatewayRecords.forEach((g, idx) => {
      if (!matchedGatewayRefs.has(g.merchantRef)) {
        const gatewayIsSuccess = isSuccessStatus(g.gatewayStatus);
        results.push({
          id: `gateway_${idx}`,
          refNo: g.merchantRef,
          upsBankRef: '—',
          prn: 'UNKNOWN (Gateway Only)',
          studentName: g.customerEmail || '—',
          purpose: 'Unrecorded in UPS',
          amount: g.amount,
          upsStatus: 'MISSING',
          upsDate: '—',
          gatewayMatched: false,
          gatewayName: g.gatewayName === 'SBI_EPAY' ? 'SBI ePay' : 'ATOM',
          gatewayTxnId: g.gatewayTxnId || '—',
          gatewayBankRef: g.bankRefNo || '—',
          gatewayAmount: g.amount,
          gatewayStatus: g.gatewayStatus,
          gatewayDesc: g.description || '—',
          gatewayDate: g.txnDate || '—',
          category: 'MISSING_IN_UPS',
          remarks: gatewayIsSuccess
            ? 'Gateway processed fee, but transaction record is absent from UPS report.'
            : 'Unlinked failed attempt recorded on gateway.',
          rawUps: null,
          rawGateway: g.raw
        });
      }
    });

    // 4. Calculate KPI metrics
    let actionNeededCount = 0;
    let actionNeededAmount = 0;
    let reconciledCount = 0;
    let reconciledAmount = 0;
    let bothFailedCount = 0;
    let discrepancyCount = 0;
    let missingInGatewayCount = 0;
    let missingInUpsCount = 0;

    results.forEach((r) => {
      if (r.category === 'ACTION_NEEDED') {
        actionNeededCount++;
        actionNeededAmount += r.amount || r.gatewayAmount || 0;
      } else if (r.category === 'RECONCILED') {
        reconciledCount++;
        reconciledAmount += r.amount || 0;
      } else if (r.category === 'BOTH_FAILED') {
        bothFailedCount++;
      } else if (r.category === 'DISCREPANCY') {
        discrepancyCount++;
      } else if (r.category === 'MISSING_IN_GATEWAY') {
        missingInGatewayCount++;
      } else if (r.category === 'MISSING_IN_UPS') {
        missingInUpsCount++;
      }
    });

    return {
      reconciledRecords: results,
      kpiSummary: {
        total: results.length,
        actionNeededCount,
        actionNeededAmount,
        reconciledCount,
        reconciledAmount,
        bothFailedCount,
        discrepancyCount,
        missingInGatewayCount,
        missingInUpsCount
      },
      uniquePurposes: Array.from(purposes).sort()
    };
  }, [upsRawRows, gatewayRawRows, gatewayType, extractGatewayRecord]);

  // Filtered display records
  const filteredRecords = useMemo(() => {
    return reconciledRecords.filter((r) => {
      // Category / Status Tab Filter
      if (statusFilter !== 'ALL' && r.category !== statusFilter) {
        return false;
      }

      // Purpose Filter
      if (purposeFilter !== 'ALL' && r.purpose !== purposeFilter) {
        return false;
      }

      // Gateway Source Filter
      if (gatewaySourceFilter !== 'ALL') {
        if (gatewaySourceFilter === 'ATOM' && !r.gatewayName.includes('ATOM')) return false;
        if (gatewaySourceFilter === 'SBI_EPAY' && !r.gatewayName.includes('SBI')) return false;
      }

      // Search Query
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const refMatch = r.refNo.toLowerCase().includes(q);
        const prnMatch = r.prn.toLowerCase().includes(q);
        const nameMatch = r.studentName.toLowerCase().includes(q);
        const bankMatch = r.gatewayBankRef.toLowerCase().includes(q) || r.gatewayTxnId.toLowerCase().includes(q);
        const descMatch = r.remarks.toLowerCase().includes(q);
        if (!refMatch && !prnMatch && !nameMatch && !bankMatch && !descMatch) return false;
      }

      return true;
    });
  }, [reconciledRecords, statusFilter, purposeFilter, gatewaySourceFilter, searchQuery]);

  // Paginated Rows
  const paginatedRows = useMemo(() => {
    if (pageSize === 'ALL') return filteredRecords;
    const start = page * Number(pageSize);
    return filteredRecords.slice(start, start + Number(pageSize));
  }, [filteredRecords, page, pageSize]);

  const totalPages = pageSize === 'ALL' ? 1 : Math.ceil(filteredRecords.length / Number(pageSize));

  // Copy Actionable PRNs
  const handleCopyActionablePrns = () => {
    const actionable = reconciledRecords.filter((r) => r.category === 'ACTION_NEEDED' && r.prn && r.prn !== 'UNKNOWN (Gateway Only)');
    const prnList = Array.from(new Set(actionable.map((r) => r.prn))).join('\n');
    if (!prnList) {
      alert('No Actionable PRNs to copy.');
      return;
    }
    navigator.clipboard.writeText(prnList);
    setCopiedType('prn');
    setTimeout(() => setCopiedType(null), 2500);
  };

  // Copy Actionable Reference Numbers (SQL IN clause format)
  const handleCopyActionableRefs = () => {
    const actionable = reconciledRecords.filter((r) => r.category === 'ACTION_NEEDED');
    const refs = Array.from(new Set(actionable.map((r) => `'${r.refNo}'`))).join(', ');
    if (!refs) {
      alert('No Actionable Reference Numbers to copy.');
      return;
    }
    navigator.clipboard.writeText(refs);
    setCopiedType('ref');
    setTimeout(() => setCopiedType(null), 2500);
  };

  // 1. Export UPS Fix / Sync File (.xlsx)
  const handleDownloadUpsSyncFile = () => {
    const actionable = reconciledRecords.filter((r) => r.category === 'ACTION_NEEDED');
    if (!actionable.length) {
      alert('No Actionable (UPS Failed / Gateway Paid) records to export.');
      return;
    }

    const exportRows = actionable.map((r) => {
      const u = r.rawUps || {};
      return {
        TRANSACTIONDATE: u.TRANSACTIONDATE || r.upsDate,
        CLIENTAPPLICATIONID: u.CLIENTAPPLICATIONID || 114,
        NAME: u.NAME || r.studentName,
        TRANSACTIONPURPOSEID: u.TRANSACTIONPURPOSEID || '',
        PURPOSENAME: u.PURPOSENAME || r.purpose,
        AMOUNT: r.amount || r.gatewayAmount,
        CLIENTAPPTRANSCATIONREFERENCENUMBER: r.refNo,
        TRANSCATIONREFERENCENUMBER: r.gatewayTxnId !== '—' ? r.gatewayTxnId : r.gatewayBankRef,
        DESCRIPTION: u.DESCRIPTION || 'Online Fee Payment',
        OTHERPARAMETERS: u.OTHERPARAMETERS || 'goLang',
        REF1_PRN: r.prn,
        REF2: u.REF2 || '',
        REF3: u.REF3 || '',
        REF4: u.REF4 || '',
        REF5: u.REF5 || '',
        CLIENTSYNCDATE: new Date().toISOString().replace('T', ' ').substring(0, 19),
        CURRENT_UPS_STATUS: r.upsStatus,
        RECOMMENDED_STATUS: 'Success',
        GATEWAY_NAME: r.gatewayName,
        GATEWAY_STATUS: r.gatewayStatus,
        GATEWAY_BANK_REF: r.gatewayBankRef,
        RECONCILIATION_NOTE: 'Paid in Gateway. Update status in UPS to complete candidate registration.'
      };
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportRows);
    ws['!cols'] = Object.keys(exportRows[0] || {}).map((k) => ({ wch: Math.max(k.length + 4, 15) }));
    XLSX.utils.book_append_sheet(wb, ws, 'UPS_Sync_Update_List');

    const fileName = `UPS_Failed_Gateway_Success_Sync_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
    setStatusMsg(`Downloaded UPS Update Sheet (${fileName}) with ${exportRows.length} records!`);
    setStatusType('success');
  };

  // 2. Export Full Comprehensive Reconciliation Report
  const handleDownloadFullReport = () => {
    if (!reconciledRecords.length) {
      alert('No reconciliation data available to export.');
      return;
    }

    const wb = XLSX.utils.book_new();

    // Sheet 1: Summary KPIs
    const summaryRows = [
      ['KANNUR UNIVERSITY - FEE PAYMENT RECONCILIATION AUDIT REPORT', ''],
      ['Generated On', new Date().toLocaleString()],
      ['UPS Source File', upsFile?.name || 'Manual / Demo Data'],
      ['Gateway Source File', gatewayFile?.name || 'Manual / Demo Data'],
      ['', ''],
      ['Metric', 'Count / Value'],
      ['Total Transactions Analyzed', kpiSummary.total],
      ['CRITICAL: UPS Failed / Initiated but Gateway SUCCESS (Action Needed)', kpiSummary.actionNeededCount],
      ['Total Recoverable Fee Amount (₹)', kpiSummary.actionNeededAmount],
      ['Fully Reconciled Transactions (Both Success)', kpiSummary.reconciledCount],
      ['Total Reconciled Amount (₹)', kpiSummary.reconciledAmount],
      ['Both Failed / Incomplete', kpiSummary.bothFailedCount],
      ['Discrepancies / Reversals', kpiSummary.discrepancyCount],
      ['Missing in Gateway (Dropped Attempts)', kpiSummary.missingInGatewayCount],
      ['Missing in UPS (Gateway Unrecorded)', kpiSummary.missingInUpsCount]
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
    wsSummary['!cols'] = [{ wch: 55 }, { wch: 35 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Executive_Summary');

    // Helper for table sheet
    const makeSheet = (records) => {
      const data = records.map((r) => ({
        'Transaction Ref No': r.refNo,
        'Student PRN (REF1)': r.prn,
        'Student / Billing Name': r.studentName,
        'Purpose': r.purpose,
        'Amount (₹)': r.amount || r.gatewayAmount,
        'UPS Status': r.upsStatus,
        'UPS Txn Date': r.upsDate,
        'Gateway Provider': r.gatewayName,
        'Gateway Status': r.gatewayStatus,
        'Gateway Txn ID': r.gatewayTxnId,
        'Bank Ref / Trace No': r.gatewayBankRef,
        'Gateway Date': r.gatewayDate,
        'Reconciliation Category': r.category,
        'Audit Remarks': r.remarks
      }));
      const ws = XLSX.utils.json_to_sheet(data);
      if (data.length) {
        ws['!cols'] = Object.keys(data[0]).map((k) => ({ wch: Math.max(k.length + 3, 16) }));
      }
      return ws;
    };

    // Sheet 2: Action Needed
    const actionList = reconciledRecords.filter((r) => r.category === 'ACTION_NEEDED');
    if (actionList.length) XLSX.utils.book_append_sheet(wb, makeSheet(actionList), 'Action_Needed_Paid');

    // Sheet 3: Fully Reconciled
    const reconciledList = reconciledRecords.filter((r) => r.category === 'RECONCILED');
    if (reconciledList.length) XLSX.utils.book_append_sheet(wb, makeSheet(reconciledList), 'Fully_Reconciled');

    // Sheet 4: Discrepancies
    const discList = reconciledRecords.filter((r) => r.category === 'DISCREPANCY');
    if (discList.length) XLSX.utils.book_append_sheet(wb, makeSheet(discList), 'Discrepancies');

    // Sheet 5: Missing in Gateway
    const missingList = reconciledRecords.filter((r) => r.category === 'MISSING_IN_GATEWAY');
    if (missingList.length) XLSX.utils.book_append_sheet(wb, makeSheet(missingList), 'Missing_In_Gateway');

    // Sheet 6: All Records
    XLSX.utils.book_append_sheet(wb, makeSheet(reconciledRecords), 'All_Records');

    const fileName = `Fee_Payment_Reconciliation_Full_Report_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
    setStatusMsg(`Downloaded Full Audit Report (${fileName}) with all tabs!`);
    setStatusType('success');
  };

  return (
    <div style={styles.container}>
      {/* Top Header */}
      <header style={styles.header}>
        <div style={styles.headerLeft}>
          <Link to="/" style={styles.backLink}>
            <ArrowLeft size={16} />
            <span>Portal</span>
          </Link>
          <div style={styles.divider} />
          <div style={styles.titleBadge}>
            <CreditCard size={20} color="var(--accent)" />
            <h1 style={styles.title}>Fee Payment Reconciliation Engine</h1>
            <span style={styles.badge}>UPS vs ATOM & SBI ePay</span>
          </div>
        </div>
        <div style={styles.headerRight}>
          <button
            type="button"
            style={styles.demoBtn}
            onClick={handleLoadDemoData}
            title="Load sample test datasets from real university scenarios"
          >
            <Sparkles size={14} />
            <span>Load Sample Data</span>
          </button>
          <button
            type="button"
            style={styles.resetBtn}
            onClick={() => {
              setUpsRawRows([]);
              setGatewayRawRows([]);
              setUpsFile(null);
              setGatewayFile(null);
              setStatusMsg('Ready for upload.');
              setStatusType('info');
            }}
            title="Reset and clear uploaded data"
          >
            <RotateCcw size={14} />
            <span>Reset</span>
          </button>
        </div>
      </header>

      {/* Main Content Area */}
      <div style={styles.content}>
        {/* Upload Zone (Dual Dropboxes) */}
        <div style={styles.uploadGrid}>
          {/* Box 1: UPS Report */}
          <div style={styles.uploadCard}>
            <div style={styles.cardHeader}>
              <div style={styles.cardHeaderTitle}>
                <div style={{ ...styles.iconBox, background: 'rgba(23, 107, 135, 0.12)', color: 'var(--accent)' }}>
                  <FileSpreadsheet size={18} />
                </div>
                <div>
                  <h3 style={styles.cardTitle}>1. University Payment System (UPS) Report</h3>
                  <p style={styles.cardSub}>Upload the UPS Transaction Export (.xlsx, .xls, .csv)</p>
                </div>
              </div>
              {upsFile && (
                <span style={styles.filePill}>
                  <CheckCircle2 size={12} color="#10b981" />
                  {upsRawRows.length} Rows
                </span>
              )}
            </div>

            <div style={styles.dropZone}>
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleUpsUpload}
                style={styles.fileInput}
                id="ups-file-input"
              />
              <label htmlFor="ups-file-input" style={styles.dropLabel}>
                <Upload size={24} color="var(--accent)" />
                <span style={{ fontWeight: 600, fontSize: '13px', marginTop: '6px' }}>
                  {upsFile ? upsFile.name : 'Choose or Drag & Drop UPS Report'}
                </span>
                <span style={{ fontSize: '11px', color: 'var(--muted)' }}>
                  Key column: CLIENTAPPTRANSCATIONREFERENCENUMBER, REF1 (PRN), AMOUNT, UPSSTATUS
                </span>
              </label>
            </div>

            {upsSheets.length > 1 && (
              <div style={styles.sheetSelectorRow}>
                <span style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 500 }}>Select Sheet:</span>
                <select
                  value={selectedUpsSheet}
                  onChange={(e) => handleUpsSheetChange(e.target.value)}
                  style={styles.select}
                >
                  {upsSheets.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </div>
            )}
          </div>

          {/* Box 2: Payment Gateway Report */}
          <div style={styles.uploadCard}>
            <div style={styles.cardHeader}>
              <div style={styles.cardHeaderTitle}>
                <div style={{ ...styles.iconBox, background: 'rgba(99, 102, 241, 0.12)', color: '#6366f1' }}>
                  <ArrowRightLeft size={18} />
                </div>
                <div>
                  <h3 style={styles.cardTitle}>2. Payment Gateway Report (ATOM or SBI ePay)</h3>
                  <p style={styles.cardSub}>Upload settlement / MIS report from ATOM or SBI ePay</p>
                </div>
              </div>
              {gatewayFile && (
                <span style={{ ...styles.filePill, background: 'rgba(99, 102, 241, 0.12)', color: '#6366f1' }}>
                  <CheckCircle2 size={12} color="#6366f1" />
                  {gatewayRawRows.length} Rows [{detectedGateway === 'SBI_EPAY' ? 'SBI ePay' : detectedGateway === 'ATOM' ? 'ATOM' : 'Detected'}]
                </span>
              )}
            </div>

            <div style={styles.dropZone}>
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleGatewayUpload}
                style={styles.fileInput}
                id="gateway-file-input"
              />
              <label htmlFor="gateway-file-input" style={styles.dropLabel}>
                <Upload size={24} color="#6366f1" />
                <span style={{ fontWeight: 600, fontSize: '13px', marginTop: '6px' }}>
                  {gatewayFile ? gatewayFile.name : 'Choose or Drag & Drop Gateway Report'}
                </span>
                <span style={{ fontSize: '11px', color: 'var(--muted)' }}>
                  Auto-detects ATOM (Merchant Txn ID) & SBI ePay (MERCHANT ORDER NO)
                </span>
              </label>
            </div>

            <div style={styles.gatewayControlsRow}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 600 }}>Provider Mode:</span>
                <div style={styles.segmentedControl}>
                  <button
                    type="button"
                    style={{ ...styles.segmentBtn, ...(gatewayType === 'auto' ? styles.segmentBtnActive : {}) }}
                    onClick={() => setGatewayType('auto')}
                  >
                    Auto-Detect
                  </button>
                  <button
                    type="button"
                    style={{ ...styles.segmentBtn, ...(gatewayType === 'atom' ? styles.segmentBtnActive : {}) }}
                    onClick={() => setGatewayType('atom')}
                  >
                    ATOM
                  </button>
                  <button
                    type="button"
                    style={{ ...styles.segmentBtn, ...(gatewayType === 'epay' ? styles.segmentBtnActive : {}) }}
                    onClick={() => setGatewayType('epay')}
                  >
                    SBI ePay
                  </button>
                </div>
              </div>

              {gatewaySheets.length > 1 && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Sheet:</span>
                  <select
                    value={selectedGatewaySheet}
                    onChange={(e) => handleGatewaySheetChange(e.target.value)}
                    style={{ ...styles.select, padding: '3px 8px', fontSize: '11px' }}
                  >
                    {gatewaySheets.map((s) => (
                      <option key={s} value={s}>
                        {s}
                      </option>
                    ))}
                  </select>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Live Feedback Status Banner */}
        <div style={{ ...styles.statusBanner, ...styles[`status_${statusType}`] }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            {statusType === 'success' && <CheckCircle2 size={16} color="#10b981" />}
            {statusType === 'warning' && <AlertTriangle size={16} color="#f59e0b" />}
            {statusType === 'error' && <XCircle size={16} color="#ef4444" />}
            {statusType === 'info' && <Clock size={16} color="var(--accent)" />}
            <span style={{ fontSize: '12.5px', fontWeight: 500 }}>{statusMsg}</span>
          </div>
          {reconciledRecords.length > 0 && (
            <div style={{ display: 'flex', gap: '8px' }}>
              <button
                type="button"
                style={styles.quickExportBtn}
                onClick={handleDownloadUpsSyncFile}
                disabled={kpiSummary.actionNeededCount === 0}
                title="Download Excel sheet with only transactions that paid at gateway but failed in UPS"
              >
                <Download size={13} />
                <span>Export UPS Sync List ({kpiSummary.actionNeededCount})</span>
              </button>
              <button
                type="button"
                style={{ ...styles.quickExportBtn, background: 'var(--accent)', color: '#fff' }}
                onClick={handleDownloadFullReport}
                title="Download complete multi-tab reconciliation report"
              >
                <FileSpreadsheet size={13} />
                <span>Full Audit Report</span>
              </button>
            </div>
          )}
        </div>

        {/* KPI Dashboard Cards */}
        {reconciledRecords.length > 0 && (
          <div style={styles.kpiGrid}>
            {/* 1. Action Needed (Critical Alert) */}
            <div
              style={{
                ...styles.kpiCard,
                border: '1.5px solid #ef4444',
                background: statusFilter === 'ACTION_NEEDED' ? 'rgba(239, 68, 68, 0.08)' : 'var(--card-bg)',
                cursor: 'pointer'
              }}
              onClick={() => {
                setStatusFilter('ACTION_NEEDED');
                setPage(0);
              }}
            >
              <div style={styles.kpiHeader}>
                <span style={{ ...styles.kpiTitle, color: '#ef4444', fontWeight: 700 }}>
                  ⚠️ ACTION NEEDED (UPS Failed / Gateway Paid)
                </span>
                <ShieldAlert size={18} color="#ef4444" />
              </div>
              <div style={styles.kpiBody}>
                <div style={{ fontSize: '24px', fontWeight: 800, color: '#ef4444' }}>
                  {kpiSummary.actionNeededCount}
                </div>
                <div style={{ fontSize: '11.5px', color: 'var(--muted)', marginTop: '2px' }}>
                  Recoverable Fee: <strong style={{ color: 'var(--text)' }}>₹{kpiSummary.actionNeededAmount.toLocaleString()}</strong>
                </div>
              </div>
              <div style={styles.kpiFooterNote}>
                Students charged by gateway; fee status missing in UPS!
              </div>
            </div>

            {/* 2. Reconciled Success */}
            <div
              style={{
                ...styles.kpiCard,
                border: '1px solid var(--line)',
                background: statusFilter === 'RECONCILED' ? 'rgba(16, 185, 129, 0.08)' : 'var(--card-bg)',
                cursor: 'pointer'
              }}
              onClick={() => {
                setStatusFilter('RECONCILED');
                setPage(0);
              }}
            >
              <div style={styles.kpiHeader}>
                <span style={{ ...styles.kpiTitle, color: '#10b981' }}>✅ Fully Reconciled</span>
                <ShieldCheck size={18} color="#10b981" />
              </div>
              <div style={styles.kpiBody}>
                <div style={{ fontSize: '24px', fontWeight: 800, color: '#10b981' }}>
                  {kpiSummary.reconciledCount}
                </div>
                <div style={{ fontSize: '11.5px', color: 'var(--muted)', marginTop: '2px' }}>
                  Total Collected: <strong style={{ color: 'var(--text)' }}>₹{kpiSummary.reconciledAmount.toLocaleString()}</strong>
                </div>
              </div>
              <div style={styles.kpiFooterNote}>Recorded & verified in both UPS and Gateway</div>
            </div>

            {/* 3. Both Failed */}
            <div
              style={{
                ...styles.kpiCard,
                border: '1px solid var(--line)',
                background: statusFilter === 'BOTH_FAILED' ? 'rgba(148, 163, 184, 0.12)' : 'var(--card-bg)',
                cursor: 'pointer'
              }}
              onClick={() => {
                setStatusFilter('BOTH_FAILED');
                setPage(0);
              }}
            >
              <div style={styles.kpiHeader}>
                <span style={styles.kpiTitle}>Clean Failures</span>
                <XCircle size={18} color="var(--muted)" />
              </div>
              <div style={styles.kpiBody}>
                <div style={{ fontSize: '24px', fontWeight: 700, color: 'var(--text)' }}>
                  {kpiSummary.bothFailedCount}
                </div>
                <div style={{ fontSize: '11.5px', color: 'var(--muted)', marginTop: '2px' }}>
                  Incomplete attempts (No funds debited)
                </div>
              </div>
              <div style={styles.kpiFooterNote}>Both systems report failed / cancelled</div>
            </div>

            {/* 4. Missing in Gateway */}
            <div
              style={{
                ...styles.kpiCard,
                border: '1px solid var(--line)',
                background: statusFilter === 'MISSING_IN_GATEWAY' ? 'rgba(245, 158, 11, 0.08)' : 'var(--card-bg)',
                cursor: 'pointer'
              }}
              onClick={() => {
                setStatusFilter('MISSING_IN_GATEWAY');
                setPage(0);
              }}
            >
              <div style={styles.kpiHeader}>
                <span style={{ ...styles.kpiTitle, color: '#f59e0b' }}>Missing in Gateway</span>
                <AlertTriangle size={18} color="#f59e0b" />
              </div>
              <div style={styles.kpiBody}>
                <div style={{ fontSize: '24px', fontWeight: 700, color: '#f59e0b' }}>
                  {kpiSummary.missingInGatewayCount}
                </div>
                <div style={{ fontSize: '11.5px', color: 'var(--muted)', marginTop: '2px' }}>
                  Abandoned before payment gateway
                </div>
              </div>
              <div style={styles.kpiFooterNote}>Initiated in UPS but never reached bank</div>
            </div>

            {/* 5. Total Transactions */}
            <div
              style={{
                ...styles.kpiCard,
                border: '1px solid var(--line)',
                background: statusFilter === 'ALL' ? 'rgba(23, 107, 135, 0.08)' : 'var(--card-bg)',
                cursor: 'pointer'
              }}
              onClick={() => {
                setStatusFilter('ALL');
                setPage(0);
              }}
            >
              <div style={styles.kpiHeader}>
                <span style={styles.kpiTitle}>Total Analyzed</span>
                <Layers size={18} color="var(--accent)" />
              </div>
              <div style={styles.kpiBody}>
                <div style={{ fontSize: '24px', fontWeight: 700, color: 'var(--text)' }}>
                  {kpiSummary.total}
                </div>
                <div style={{ fontSize: '11.5px', color: 'var(--muted)', marginTop: '2px' }}>
                  Combined unique payment events
                </div>
              </div>
              <div style={styles.kpiFooterNote}>Click to view all transactions</div>
            </div>
          </div>
        )}

        {/* Toolbar: Category Tabs, Search, Purpose Filter, Action Buttons */}
        {reconciledRecords.length > 0 && (
          <div style={styles.toolbarCard}>
            <div style={styles.tabsRow}>
              <button
                type="button"
                style={{
                  ...styles.tabBtn,
                  ...(statusFilter === 'ACTION_NEEDED' ? styles.tabBtnAlert : {})
                }}
                onClick={() => {
                  setStatusFilter('ACTION_NEEDED');
                  setPage(0);
                }}
              >
                ⚠️ Action Needed ({kpiSummary.actionNeededCount})
              </button>
              <button
                type="button"
                style={{
                  ...styles.tabBtn,
                  ...(statusFilter === 'RECONCILED' ? styles.tabBtnActive : {})
                }}
                onClick={() => {
                  setStatusFilter('RECONCILED');
                  setPage(0);
                }}
              >
                ✅ Reconciled ({kpiSummary.reconciledCount})
              </button>
              <button
                type="button"
                style={{
                  ...styles.tabBtn,
                  ...(statusFilter === 'BOTH_FAILED' ? styles.tabBtnActive : {})
                }}
                onClick={() => {
                  setStatusFilter('BOTH_FAILED');
                  setPage(0);
                }}
              >
                Both Failed ({kpiSummary.bothFailedCount})
              </button>
              <button
                type="button"
                style={{
                  ...styles.tabBtn,
                  ...(statusFilter === 'MISSING_IN_GATEWAY' ? styles.tabBtnActive : {})
                }}
                onClick={() => {
                  setStatusFilter('MISSING_IN_GATEWAY');
                  setPage(0);
                }}
              >
                Missing in Gateway ({kpiSummary.missingInGatewayCount})
              </button>
              <button
                type="button"
                style={{
                  ...styles.tabBtn,
                  ...(statusFilter === 'ALL' ? styles.tabBtnActive : {})
                }}
                onClick={() => {
                  setStatusFilter('ALL');
                  setPage(0);
                }}
              >
                All Records ({kpiSummary.total})
              </button>
            </div>

            <div style={styles.filterControlsRow}>
              {/* Search */}
              <div style={styles.searchBox}>
                <Search size={14} color="var(--muted)" />
                <input
                  type="text"
                  placeholder="Search PRN, Txn Ref, Student Name, Bank Ref..."
                  value={searchQuery}
                  onChange={(e) => {
                    setSearchQuery(e.target.value);
                    setPage(0);
                  }}
                  style={styles.searchInput}
                />
                {searchQuery && (
                  <button
                    type="button"
                    onClick={() => setSearchQuery('')}
                    style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'var(--muted)' }}
                  >
                    ✕
                  </button>
                )}
              </div>

              {/* Purpose Filter */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 600 }}>Purpose:</span>
                <select
                  value={purposeFilter}
                  onChange={(e) => {
                    setPurposeFilter(e.target.value);
                    setPage(0);
                  }}
                  style={styles.select}
                >
                  <option value="ALL">All Fee Purposes</option>
                  {uniquePurposes.map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
              </div>

              {/* Gateway Source Filter */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 600 }}>Gateway:</span>
                <select
                  value={gatewaySourceFilter}
                  onChange={(e) => {
                    setGatewaySourceFilter(e.target.value);
                    setPage(0);
                  }}
                  style={styles.select}
                >
                  <option value="ALL">All Gateways</option>
                  <option value="ATOM">ATOM Only</option>
                  <option value="SBI_EPAY">SBI ePay Only</option>
                </select>
              </div>

              {/* Clipboard Action Buttons */}
              <div style={{ display: 'flex', gap: '6px', marginLeft: 'auto' }}>
                <button
                  type="button"
                  style={styles.clipboardBtn}
                  onClick={handleCopyActionablePrns}
                  title="Copy PRNs of Action Needed students (for bulk candidate update)"
                >
                  {copiedType === 'prn' ? <Check size={13} color="#10b981" /> : <Copy size={13} />}
                  <span>{copiedType === 'prn' ? 'Copied PRNs!' : 'Copy Action PRNs'}</span>
                </button>
                <button
                  type="button"
                  style={styles.clipboardBtn}
                  onClick={handleCopyActionableRefs}
                  title="Copy Transaction Reference Numbers formatted for SQL IN clause"
                >
                  {copiedType === 'ref' ? <Check size={13} color="#10b981" /> : <Copy size={13} />}
                  <span>{copiedType === 'ref' ? 'Copied SQL Refs!' : 'Copy SQL Refs'}</span>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Reconciled Data Table */}
        {reconciledRecords.length > 0 && (
          <div style={styles.tableCard}>
            <div style={styles.tableHeaderBar}>
              <div style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text)' }}>
                Showing {filteredRecords.length} transactions {statusFilter !== 'ALL' && `(Filtered by: ${statusFilter})`}
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                {/* Density Toggle */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Density:</span>
                  <button
                    type="button"
                    style={{ ...styles.densityBtn, ...(tableDensity === 'compact' ? styles.densityBtnActive : {}) }}
                    onClick={() => setTableDensity('compact')}
                  >
                    Compact
                  </button>
                  <button
                    type="button"
                    style={{ ...styles.densityBtn, ...(tableDensity === 'normal' ? styles.densityBtnActive : {}) }}
                    onClick={() => setTableDensity('normal')}
                  >
                    Normal
                  </button>
                </div>

                {/* Rows per page */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Rows:</span>
                  <select
                    value={pageSize}
                    onChange={(e) => {
                      setPageSize(e.target.value === 'ALL' ? 'ALL' : Number(e.target.value));
                      setPage(0);
                    }}
                    style={{ ...styles.select, padding: '2px 6px', fontSize: '11px' }}
                  >
                    <option value={25}>25</option>
                    <option value={50}>50</option>
                    <option value={100}>100</option>
                    <option value="ALL">All</option>
                  </select>
                </div>
              </div>
            </div>

            <div style={styles.tableWrapper}>
              <table style={styles.table}>
                <thead>
                  <tr style={styles.thRow}>
                    <th style={{ ...styles.th, width: '40px' }}>#</th>
                    <th style={styles.th}>Txn Reference No</th>
                    <th style={styles.th}>Student PRN (REF1)</th>
                    <th style={styles.th}>Student / Payer</th>
                    <th style={styles.th}>Purpose</th>
                    <th style={{ ...styles.th, textAlign: 'right' }}>Fee (₹)</th>
                    <th style={styles.th}>UPS Status</th>
                    <th style={styles.th}>Gateway</th>
                    <th style={styles.th}>Gateway Status</th>
                    <th style={styles.th}>Gateway Txn / Bank Ref</th>
                    <th style={styles.th}>Reconciliation Status</th>
                    <th style={styles.th}>Action / Audit Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  {paginatedRows.length === 0 ? (
                    <tr>
                      <td colSpan={12} style={{ textAlign: 'center', padding: '30px', color: 'var(--muted)' }}>
                        No records match the selected filters.
                      </td>
                    </tr>
                  ) : (
                    paginatedRows.map((r, i) => {
                      const rowNum = page * (pageSize === 'ALL' ? 0 : pageSize) + i + 1;
                      const isActionNeeded = r.category === 'ACTION_NEEDED';
                      const isReconciled = r.category === 'RECONCILED';
                      const isBothFailed = r.category === 'BOTH_FAILED';

                      const rowPad = tableDensity === 'compact' ? '6px 10px' : '10px 12px';

                      return (
                        <tr
                          key={r.id}
                          style={{
                            ...styles.tr,
                            background: isActionNeeded
                              ? 'rgba(239, 68, 68, 0.05)'
                              : isReconciled
                              ? 'transparent'
                              : 'transparent'
                          }}
                        >
                          <td style={{ ...styles.td, padding: rowPad, color: 'var(--muted)', fontSize: '11px' }}>
                            {rowNum}
                          </td>
                          <td style={{ ...styles.td, padding: rowPad, fontFamily: 'monospace', fontWeight: 600 }}>
                            {r.refNo}
                          </td>
                          <td style={{ ...styles.td, padding: rowPad, fontFamily: 'monospace', color: 'var(--accent)', fontWeight: 600 }}>
                            {r.prn !== 'UNKNOWN (Gateway Only)' ? r.prn : <span style={{ color: 'var(--muted)' }}>—</span>}
                          </td>
                          <td style={{ ...styles.td, padding: rowPad, fontSize: '12px' }}>
                            {r.studentName || '—'}
                          </td>
                          <td style={{ ...styles.td, padding: rowPad, fontSize: '11.5px' }}>
                            <span style={styles.purposePill}>{r.purpose}</span>
                          </td>
                          <td style={{ ...styles.td, padding: rowPad, textAlign: 'right', fontWeight: 700 }}>
                            ₹{(r.amount || r.gatewayAmount).toLocaleString()}
                          </td>
                          <td style={{ ...styles.td, padding: rowPad }}>
                            <span
                              style={{
                                ...styles.statusBadge,
                                ...(isSuccessStatus(r.upsStatus)
                                  ? styles.badgeSuccess
                                  : isFailedStatus(r.upsStatus)
                                  ? styles.badgeFailed
                                  : styles.badgePending)
                              }}
                            >
                              {r.upsStatus}
                            </span>
                          </td>
                          <td style={{ ...styles.td, padding: rowPad, fontSize: '11.5px', fontWeight: 600 }}>
                            {r.gatewayName}
                          </td>
                          <td style={{ ...styles.td, padding: rowPad }}>
                            <span
                              style={{
                                ...styles.statusBadge,
                                ...(isSuccessStatus(r.gatewayStatus)
                                  ? styles.badgeSuccess
                                  : isFailedStatus(r.gatewayStatus)
                                  ? styles.badgeFailed
                                  : styles.badgeMissing)
                              }}
                            >
                              {r.gatewayStatus}
                            </span>
                          </td>
                          <td style={{ ...styles.td, padding: rowPad, fontSize: '11px', fontFamily: 'monospace', color: 'var(--muted)' }}>
                            <div>{r.gatewayTxnId !== '—' ? r.gatewayTxnId : ''}</div>
                            {r.gatewayBankRef !== '—' && r.gatewayBankRef !== r.gatewayTxnId && (
                              <div style={{ fontSize: '10px', color: 'var(--muted)' }}>Ref: {r.gatewayBankRef}</div>
                            )}
                          </td>
                          <td style={{ ...styles.td, padding: rowPad }}>
                            {isActionNeeded ? (
                              <span style={{ ...styles.catBadge, background: '#ef4444', color: '#fff' }}>
                                ⚠️ Sync Fee in UPS
                              </span>
                            ) : isReconciled ? (
                              <span style={{ ...styles.catBadge, background: 'rgba(16, 185, 129, 0.15)', color: '#10b981' }}>
                                ✅ Verified Paid
                              </span>
                            ) : isBothFailed ? (
                              <span style={{ ...styles.catBadge, background: 'rgba(148, 163, 184, 0.15)', color: 'var(--muted)' }}>
                                Clean Fail
                              </span>
                            ) : (
                              <span style={{ ...styles.catBadge, background: 'rgba(245, 158, 11, 0.15)', color: '#f59e0b' }}>
                                {r.category}
                              </span>
                            )}
                          </td>
                          <td style={{ ...styles.td, padding: rowPad, fontSize: '11.5px', maxWidth: '280px' }}>
                            <span style={{ color: isActionNeeded ? '#ef4444' : 'var(--muted)', fontWeight: isActionNeeded ? 600 : 400 }}>
                              {r.remarks}
                            </span>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls */}
            {pageSize !== 'ALL' && totalPages > 1 && (
              <div style={styles.paginationRow}>
                <div style={{ fontSize: '12px', color: 'var(--muted)' }}>
                  Page {page + 1} of {totalPages} ({filteredRecords.length} records)
                </div>
                <div style={{ display: 'flex', gap: '6px' }}>
                  <button
                    type="button"
                    style={{ ...styles.pageBtn, ...(page === 0 ? styles.pageBtnDisabled : {}) }}
                    onClick={() => setPage((p) => Math.max(0, p - 1))}
                    disabled={page === 0}
                  >
                    Previous
                  </button>
                  {Array.from({ length: Math.min(5, totalPages) }, (_, idx) => {
                    let pNum = page - 2 + idx;
                    if (page < 2) pNum = idx;
                    if (page > totalPages - 3) pNum = totalPages - 5 + idx;
                    if (pNum < 0 || pNum >= totalPages) return null;
                    return (
                      <button
                        key={pNum}
                        type="button"
                        style={{ ...styles.pageBtn, ...(page === pNum ? styles.pageBtnActive : {}) }}
                        onClick={() => setPage(pNum)}
                      >
                        {pNum + 1}
                      </button>
                    );
                  })}
                  <button
                    type="button"
                    style={{ ...styles.pageBtn, ...(page >= totalPages - 1 ? styles.pageBtnDisabled : {}) }}
                    onClick={() => setPage((p) => Math.min(totalPages - 1, p + 1))}
                    disabled={page >= totalPages - 1}
                  >
                    Next
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Empty State Guide */}
        {reconciledRecords.length === 0 && (
          <div style={styles.emptyStateCard}>
            <CreditCard size={48} color="var(--accent)" style={{ margin: '0 auto 16px' }} />
            <h2 style={{ fontSize: '18px', fontWeight: 700, margin: '0 0 8px', color: 'var(--text)' }}>
              Reconcile Kannur University UPS Payments against ATOM & SBI ePay
            </h2>
            <p style={{ fontSize: '13px', color: 'var(--muted)', maxWidth: '640px', margin: '0 auto 20px', lineHeight: 1.6 }}>
              Upload your <strong>UPS Report</strong> on the left and the <strong>Payment Gateway Settlement Report</strong> (from ATOM or SBI ePay) on the right.
              The engine automatically matches transaction reference numbers, identifies students whose accounts were debited but status failed in UPS, and exports instant sync files.
            </p>
            <div style={{ display: 'flex', justifyContent: 'center', gap: '12px' }}>
              <button type="button" style={styles.demoActionBtn} onClick={handleLoadDemoData}>
                <Sparkles size={16} />
                <span>Load Sample Datasets to Explore</span>
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

const styles = {
  container: {
    minHeight: '100vh',
    background: 'var(--bg)',
    color: 'var(--text)',
    display: 'flex',
    flexDirection: 'column',
    fontFamily: 'var(--font-family)'
  },
  header: {
    padding: '12px 24px',
    borderBottom: '1px solid var(--line)',
    background: 'var(--card-bg)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    zIndex: 10
  },
  headerLeft: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px'
  },
  headerRight: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px'
  },
  backLink: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    color: 'var(--muted)',
    textDecoration: 'none',
    fontSize: '13px',
    fontWeight: 500,
    padding: '5px 10px',
    borderRadius: '6px',
    border: '1px solid var(--line)'
  },
  divider: {
    width: '1px',
    height: '20px',
    background: 'var(--line)'
  },
  titleBadge: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px'
  },
  title: {
    fontSize: '16px',
    fontWeight: 700,
    margin: 0,
    color: 'var(--text)'
  },
  badge: {
    fontSize: '11px',
    fontWeight: 600,
    background: 'rgba(23, 107, 135, 0.1)',
    color: 'var(--accent)',
    padding: '2px 8px',
    borderRadius: '12px'
  },
  demoBtn: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    background: 'rgba(99, 102, 241, 0.12)',
    color: '#6366f1',
    border: '1px solid rgba(99, 102, 241, 0.3)',
    padding: '6px 12px',
    borderRadius: '6px',
    fontSize: '12px',
    fontWeight: 600,
    cursor: 'pointer'
  },
  resetBtn: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    background: 'transparent',
    color: 'var(--muted)',
    border: '1px solid var(--line)',
    padding: '6px 12px',
    borderRadius: '6px',
    fontSize: '12px',
    fontWeight: 500,
    cursor: 'pointer'
  },
  content: {
    padding: '20px 24px',
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    maxWidth: '1600px',
    width: '100%',
    margin: '0 auto',
    boxSizing: 'border-box'
  },
  uploadGrid: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '16px'
  },
  uploadCard: {
    background: 'var(--card-bg)',
    border: '1px solid var(--line)',
    borderRadius: '10px',
    padding: '16px',
    display: 'flex',
    flexDirection: 'column',
    gap: '12px'
  },
  cardHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start'
  },
  cardHeaderTitle: {
    display: 'flex',
    gap: '10px',
    alignItems: 'center'
  },
  iconBox: {
    width: '36px',
    height: '36px',
    borderRadius: '8px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center'
  },
  cardTitle: {
    fontSize: '14px',
    fontWeight: 700,
    margin: 0,
    color: 'var(--text)'
  },
  cardSub: {
    fontSize: '11px',
    color: 'var(--muted)',
    margin: '2px 0 0'
  },
  filePill: {
    fontSize: '11px',
    fontWeight: 600,
    background: 'rgba(16, 185, 129, 0.1)',
    color: '#10b981',
    padding: '3px 8px',
    borderRadius: '6px',
    display: 'flex',
    alignItems: 'center',
    gap: '4px'
  },
  dropZone: {
    border: '1.5px dashed var(--line)',
    borderRadius: '8px',
    padding: '16px',
    textAlign: 'center',
    position: 'relative',
    background: 'var(--bg)',
    cursor: 'pointer',
    transition: 'all 0.15s ease'
  },
  fileInput: {
    position: 'absolute',
    top: 0,
    left: 0,
    width: '100%',
    height: '100%',
    opacity: 0,
    cursor: 'pointer'
  },
  dropLabel: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    cursor: 'pointer'
  },
  sheetSelectorRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingTop: '6px',
    borderTop: '1px dashed var(--line)'
  },
  gatewayControlsRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingTop: '6px',
    borderTop: '1px dashed var(--line)'
  },
  segmentedControl: {
    display: 'flex',
    background: 'var(--bg)',
    border: '1px solid var(--line)',
    borderRadius: '6px',
    padding: '2px'
  },
  segmentBtn: {
    background: 'transparent',
    border: 'none',
    padding: '3px 8px',
    fontSize: '11px',
    fontWeight: 500,
    color: 'var(--muted)',
    borderRadius: '4px',
    cursor: 'pointer'
  },
  segmentBtnActive: {
    background: 'var(--card-bg)',
    color: 'var(--text)',
    fontWeight: 700,
    boxShadow: '0 1px 2px rgba(0,0,0,0.06)'
  },
  select: {
    background: 'var(--bg)',
    color: 'var(--text)',
    border: '1px solid var(--line)',
    borderRadius: '6px',
    padding: '4px 10px',
    fontSize: '12px',
    fontFamily: 'inherit'
  },
  statusBanner: {
    padding: '10px 16px',
    borderRadius: '8px',
    border: '1px solid var(--line)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '12px',
    background: 'var(--card-bg)'
  },
  status_info: {
    borderLeft: '4px solid var(--accent)'
  },
  status_success: {
    borderLeft: '4px solid #10b981'
  },
  status_warning: {
    borderLeft: '4px solid #f59e0b'
  },
  status_error: {
    borderLeft: '4px solid #ef4444'
  },
  quickExportBtn: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    background: '#ef4444',
    color: '#fff',
    border: 'none',
    padding: '5px 10px',
    borderRadius: '6px',
    fontSize: '11.5px',
    fontWeight: 600,
    cursor: 'pointer'
  },
  kpiGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(5, 1fr)',
    gap: '12px'
  },
  kpiCard: {
    background: 'var(--card-bg)',
    borderRadius: '8px',
    padding: '14px',
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'space-between',
    minHeight: '100px',
    boxSizing: 'border-box'
  },
  kpiHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center'
  },
  kpiTitle: {
    fontSize: '11.5px',
    fontWeight: 600,
    color: 'var(--muted)'
  },
  kpiBody: {
    margin: '6px 0'
  },
  kpiFooterNote: {
    fontSize: '10px',
    color: 'var(--muted)',
    borderTop: '1px dashed var(--line)',
    paddingTop: '6px'
  },
  toolbarCard: {
    background: 'var(--card-bg)',
    border: '1px solid var(--line)',
    borderRadius: '8px',
    padding: '12px',
    display: 'flex',
    flexDirection: 'column',
    gap: '10px'
  },
  tabsRow: {
    display: 'flex',
    gap: '6px',
    borderBottom: '1px solid var(--line)',
    paddingBottom: '8px',
    overflowX: 'auto'
  },
  tabBtn: {
    background: 'transparent',
    border: 'none',
    padding: '6px 12px',
    borderRadius: '6px',
    fontSize: '12px',
    fontWeight: 600,
    color: 'var(--muted)',
    cursor: 'pointer',
    whiteSpace: 'nowrap'
  },
  tabBtnActive: {
    background: 'var(--accent)',
    color: '#fff'
  },
  tabBtnAlert: {
    background: '#ef4444',
    color: '#fff'
  },
  filterControlsRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    flexWrap: 'wrap'
  },
  searchBox: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    background: 'var(--bg)',
    border: '1px solid var(--line)',
    borderRadius: '6px',
    padding: '4px 10px',
    width: '320px'
  },
  searchInput: {
    background: 'transparent',
    border: 'none',
    outline: 'none',
    color: 'var(--text)',
    fontSize: '12px',
    width: '100%'
  },
  clipboardBtn: {
    display: 'flex',
    alignItems: 'center',
    gap: '5px',
    background: 'var(--bg)',
    color: 'var(--text)',
    border: '1px solid var(--line)',
    padding: '5px 10px',
    borderRadius: '6px',
    fontSize: '11.5px',
    fontWeight: 600,
    cursor: 'pointer'
  },
  tableCard: {
    background: 'var(--card-bg)',
    border: '1px solid var(--line)',
    borderRadius: '8px',
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column'
  },
  tableHeaderBar: {
    padding: '10px 16px',
    background: 'var(--card-bg)',
    borderBottom: '1px solid var(--line)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between'
  },
  densityBtn: {
    background: 'transparent',
    border: '1px solid var(--line)',
    padding: '2px 8px',
    borderRadius: '4px',
    fontSize: '11px',
    color: 'var(--muted)',
    cursor: 'pointer'
  },
  densityBtnActive: {
    background: 'var(--accent)',
    color: '#fff',
    borderColor: 'var(--accent)',
    fontWeight: 600
  },
  tableWrapper: {
    overflowX: 'auto',
    width: '100%'
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    textAlign: 'left'
  },
  thRow: {
    background: 'var(--bg)',
    borderBottom: '1px solid var(--line)'
  },
  th: {
    padding: '8px 12px',
    fontSize: '11px',
    fontWeight: 700,
    color: 'var(--muted)',
    textTransform: 'uppercase',
    letterSpacing: '0.4px',
    whiteSpace: 'nowrap'
  },
  tr: {
    borderBottom: '1px solid var(--line)',
    transition: 'background 0.1s ease'
  },
  td: {
    fontSize: '12px',
    color: 'var(--text)',
    whiteSpace: 'nowrap'
  },
  purposePill: {
    background: 'rgba(23, 107, 135, 0.08)',
    color: 'var(--accent)',
    padding: '2px 6px',
    borderRadius: '4px',
    fontWeight: 500,
    fontSize: '11px'
  },
  statusBadge: {
    fontSize: '10.5px',
    fontWeight: 700,
    padding: '2px 7px',
    borderRadius: '10px',
    display: 'inline-block'
  },
  badgeSuccess: {
    background: 'rgba(16, 185, 129, 0.15)',
    color: '#10b981'
  },
  badgeFailed: {
    background: 'rgba(239, 68, 68, 0.15)',
    color: '#ef4444'
  },
  badgePending: {
    background: 'rgba(245, 158, 11, 0.15)',
    color: '#f59e0b'
  },
  badgeMissing: {
    background: 'rgba(148, 163, 184, 0.15)',
    color: 'var(--muted)'
  },
  catBadge: {
    fontSize: '10.5px',
    fontWeight: 700,
    padding: '3px 8px',
    borderRadius: '6px',
    display: 'inline-block'
  },
  paginationRow: {
    padding: '10px 16px',
    borderTop: '1px solid var(--line)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    background: 'var(--card-bg)'
  },
  pageBtn: {
    background: 'var(--bg)',
    border: '1px solid var(--line)',
    padding: '4px 10px',
    borderRadius: '4px',
    fontSize: '11px',
    color: 'var(--text)',
    cursor: 'pointer'
  },
  pageBtnActive: {
    background: 'var(--accent)',
    color: '#fff',
    borderColor: 'var(--accent)',
    fontWeight: 700
  },
  pageBtnDisabled: {
    opacity: 0.4,
    cursor: 'not-allowed'
  },
  emptyStateCard: {
    background: 'var(--card-bg)',
    border: '1px dashed var(--line)',
    borderRadius: '12px',
    padding: '60px 24px',
    textAlign: 'center',
    margin: '20px 0'
  },
  demoActionBtn: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    background: 'var(--accent)',
    color: '#fff',
    border: 'none',
    padding: '10px 20px',
    borderRadius: '8px',
    fontSize: '13px',
    fontWeight: 600,
    cursor: 'pointer',
    boxShadow: '0 2px 8px rgba(23, 107, 135, 0.25)'
  }
};
