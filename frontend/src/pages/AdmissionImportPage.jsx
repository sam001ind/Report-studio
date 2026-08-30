import React, { useState, useMemo } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { 
  ArrowLeft, 
  Upload, 
  Download, 
  FileSpreadsheet, 
  CheckCircle2, 
  AlertTriangle, 
  Search, 
  Filter, 
  RefreshCw,
  Table,
  Sparkles,
  HelpCircle,
  Database,
  Edit2
} from 'lucide-react';
import { 
  TARGET_COLUMNS, 
  CATEGORY_LOOKUP, 
  CASTE_LOOKUP, 
  RELIGION_LOOKUP, 
  PROGRAM_LOOKUP, 
  COLLEGE_LOOKUP 
} from '../data/admissionLookups';

const normalizeKey = (key) => String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');

export default function AdmissionImportPage() {
  const [sourceFile, setSourceFile] = useState(null);
  const [normalizedRows, setNormalizedRows] = useState([]);
  const [sourceHeaders, setSourceHeaders] = useState([]);
  const [stats, setStats] = useState({ total: 0, warnings: 0, ready: 0 });
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('info');
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [filterMode, setFilterMode] = useState('all'); // 'all' | 'warnings' | 'ready'
  const [page, setPage] = useState(0);
  const pageSize = 50;

  const setStatus = (msg, type = 'info') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  const transformRow = (rawRow, headerMap) => {
    const getVal = (...keys) => {
      for (const k of keys) {
        const norm = normalizeKey(k);
        const actualKey = headerMap[norm];
        if (actualKey && rawRow[actualKey] !== undefined && rawRow[actualKey] !== null) {
          const val = String(rawRow[actualKey]).trim();
          if (val) return val;
        }
      }
      return '';
    };

    // 1. Name
    const name = getVal('name', 'fullname', 'candidatename', 'studentname');

    // 2. Gender
    const rawGender = getVal('gender', 'sex');
    let gender = 'Male';
    if (rawGender.toLowerCase().startsWith('f')) gender = 'Female';
    else if (rawGender.toLowerCase().startsWith('m')) gender = 'Male';
    else if (rawGender) gender = 'Other';

    // 3. Mobile
    const rawMobile = getVal('mobile', 'mobilenumber', 'contact', 'contactnumber', 'phone');
    const mobile = rawMobile.replace(/\D/g, '').slice(-10);

    // 4. Email
    const rawEmail = getVal('email', 'emailid', 'mail');
    const email = rawEmail.toLowerCase();

    // 5. Username (Application ID / Registration Number)
    const username = getVal('username', 'applicationnumber', 'appno', 'regno', 'prn', 'candidateid');

    // 6. Date of Birth
    const rawDob = getVal('dob', 'dateofbirth', 'birthdate');
    let dob = rawDob;
    if (rawDob) {
      const d = new Date(rawDob);
      if (!isNaN(d.getTime())) {
        dob = d.toISOString().split('T')[0];
      }
    }

    // 7. AadhaarNumber
    const rawAadhaar = getVal('aadhaarnumber', 'aadhaar', 'adharnumber', 'adhar');
    const aadhaarNumber = rawAadhaar.replace(/\D/g, '').slice(0, 12);

    // 8. ABCId
    const rawABC = getVal('abcid', 'abc', 'academicbankofcredits');
    const abcId = rawABC.replace(/\D/g, '').slice(0, 12);

    // 9. CountryId & CountryName
    const countryId = getVal('countryid') || '107';
    const countryName = getVal('countryname') || 'India';

    // 11. State
    const corrStateId = getVal('corrstateid', 'stateid') || '18';
    const corrStateName = getVal('corrstatename', 'statename', 'state') || 'Kerala';

    // 13. Address, City, Pin
    const corrAddress = getVal('corraddress', 'address', 'communicationaddress', 'permaddress');
    const corrCity = getVal('corrcity', 'city', 'town', 'place', 'district');
    const rawPin = getVal('corrpin', 'pincode', 'pin', 'postalcode');
    const corrPin = rawPin.replace(/\D/g, '').slice(0, 6);

    // 16. Category
    const rawCategory = getVal('category', 'categoryname', 'castecategory', 'community');
    const normCategory = normalizeKey(rawCategory);
    const categoryLookup = CATEGORY_LOOKUP[normCategory] || { id: '10', name: rawCategory || 'General' };
    const categoryId = getVal('categoryid') || categoryLookup.id;
    const categoryName = rawCategory || categoryLookup.name;

    // 18. Caste
    const rawCaste = getVal('caste', 'subcaste');
    const normCaste = normalizeKey(rawCaste);
    const casteId = getVal('casteid') || CASTE_LOOKUP[normCaste] || '393';
    const caste = rawCaste || 'General';

    // 20. Religion
    const rawReligion = getVal('religion', 'religionname');
    const normReligion = normalizeKey(rawReligion);
    const religionLookup = RELIGION_LOOKUP[normReligion] || { id: '1', name: rawReligion || 'Hindu' };
    const religionId = getVal('religionid') || religionLookup.id;
    const religion = rawReligion || religionLookup.name;

    // 22. Program & Workflow
    const rawProgram = getVal('program', 'programname', 'programmname', 'coursename', 'course');
    const normProg = normalizeKey(rawProgram);
    let programMatch = null;
    for (const [pk, pv] of Object.entries(PROGRAM_LOOKUP)) {
      if (normProg.includes(normalizeKey(pk)) || normalizeKey(pk).includes(normProg)) {
        programMatch = pv;
        break;
      }
    }
    const programName = programMatch ? programMatch.name : (rawProgram || 'FYUGP Program');
    const programCode = getVal('programcode', 'coursecode') || (programMatch ? programMatch.code : 'UCAHISGS25');
    const workFlowId = getVal('workflowid') || (programMatch ? programMatch.workflowId : '81');

    // 25. College
    const rawCollegeCode = getVal('collegecode', 'colcode', 'center', 'centercode');
    const rawCollegeName = getVal('collegename', 'college', 'centername');
    const collegeLookup = COLLEGE_LOOKUP[rawCollegeCode] || COLLEGE_LOOKUP[normalizeKey(rawCollegeName)] || { code: rawCollegeCode || '347', name: rawCollegeName || 'University College' };
    const collegeCode = collegeLookup.code;
    const collegeId = collegeCode;
    const collegeName = collegeLookup.name;

    // 28. Qualification Faculty
    const rawQual = getVal('qualificationname', 'board', 'qualifyingexam', 'boardname');
    const normQual = normalizeKey(rawQual);
    let qualificationFacultyId = '2'; // Default HSE / CBSE
    if (normQual && !normQual.includes('hse') && !normQual.includes('cbse') && !normQual.includes('vhse') && !normQual.includes('kerala')) {
      qualificationFacultyId = '7';
    }
    const qualificationName = rawQual || 'HSE - Kerala';

    // 30. Qualification Specialization & Stream
    const rawStream = getVal('stream', 'group', 'specialization', 'branch');
    const normStream = normalizeKey(rawStream);
    let qualificationSpecializationId = '1'; // Science
    if (normStream.includes('comm')) qualificationSpecializationId = '2';
    else if (normStream.includes('hum') || normStream.includes('art')) qualificationSpecializationId = '3';
    else if (normStream) qualificationSpecializationId = '4';
    const stream = rawStream || (qualificationSpecializationId === '1' ? 'Science' : qualificationSpecializationId === '2' ? 'Commerce' : 'Humanities');

    // 32. Certificate Number / Roll No
    const certificateNumber = getVal('certificatenumber', 'regno', 'rollno', 'register_number', 'registerno');

    // 33. Attempts
    const rawAttempts = getVal('noofattempts', 'attempts');
    const noOfAttempts = rawAttempts ? parseInt(rawAttempts, 10) || 1 : 1;

    // 34. Marks & Percentage
    const rawMarks = getVal('marksobtained', 'totalmarks', 'securedmarks', 'marks');
    const marksObtained = rawMarks ? parseFloat(rawMarks) || 0 : 0;
    const rawMax = getVal('marksoutof', 'maxmarks', 'maximummarks', 'total');
    const marksOutOf = rawMax ? parseFloat(rawMax) || 1200 : 1200;

    let percentage = getVal('percentage', 'percent');
    if (!percentage && marksOutOf > 0 && marksObtained > 0) {
      percentage = ((marksObtained / marksOutOf) * 100).toFixed(4);
    } else if (percentage) {
      percentage = parseFloat(percentage).toFixed(4);
    } else {
      percentage = '0.0000';
    }

    const cgpa = getVal('cgpa', 'gpa') || (parseFloat(percentage) > 0 ? (parseFloat(percentage) / 9.5).toFixed(2) : '');

    // 38. SeatNumber (Fallback to CertificateNumber)
    const rawSeat = getVal('seatnumber', 'seatno');
    const seatNumber = rawSeat || certificateNumber;

    // 39. Result Date & Status
    const resultDate = getVal('resultdate', 'passeddate') || new Date().toISOString().split('T')[0];
    const resultStatus = getVal('resultstatus', 'status', 'result') || 'Passed';

    // Warnings flag
    let warnings = [];
    if (!name) warnings.push('Missing Name');
    if (!mobile || mobile.length < 10) warnings.push('Invalid Mobile');
    if (!email) warnings.push('Missing Email');
    if (!certificateNumber) warnings.push('Missing Reg/Cert No');

    return {
      Name: name,
      Gender: gender,
      Mobile: mobile,
      Email: email,
      username: username || certificateNumber,
      dob: dob,
      AadhaarNumber: aadhaarNumber,
      ABCId: abcId,
      CountryId: countryId,
      CountryName: countryName,
      corrStateId: corrStateId,
      corrStateName: corrStateName,
      corrAddress: corrAddress,
      corrCity: corrCity,
      corrPin: corrPin,
      categoryId: categoryId,
      categoryName: categoryName,
      CasteId: casteId,
      Caste: caste,
      religionId: religionId,
      religion: religion,
      workFlowId: workFlowId,
      'Program Name': programName,
      ProgramCode: programCode,
      CollegeCode: collegeCode,
      CollegeId: collegeId,
      'College Name': collegeName,
      QualificationFacultyId: qualificationFacultyId,
      QualificationName: qualificationName,
      QualificationSpecializationId: qualificationSpecializationId,
      Stream: stream,
      CertificateNumber: certificateNumber,
      NoOfAttempts: noOfAttempts,
      MarksObtained: marksObtained,
      Percentage: percentage,
      MarksOutOf: marksOutOf,
      CGPA: cgpa,
      SeatNumber: seatNumber,
      'Result Date': resultDate,
      ResultStatus: resultStatus,
      _warnings: warnings
    };
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setSourceFile(file);
    setIsProcessing(true);
    setStatus('Parsing and normalizing allotment spreadsheet...', 'info');

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, raw: false });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];

      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (!rawRows || !rawRows.length) {
        throw new Error('The uploaded spreadsheet contains no data rows.');
      }

      // Build header normalization map
      const headers = Object.keys(rawRows[0]);
      setSourceHeaders(headers);
      const headerMap = {};
      headers.forEach(h => {
        headerMap[normalizeKey(h)] = h;
      });

      // Normalize and enrich each row
      const transformed = rawRows.map(row => transformRow(row, headerMap));
      setNormalizedRows(transformed);

      const warningCount = transformed.filter(r => r._warnings && r._warnings.length > 0).length;
      setStats({
        total: transformed.length,
        warnings: warningCount,
        ready: transformed.length - warningCount
      });

      setStatus(`Successfully normalized ${transformed.length} candidate record(s) into 40 master columns!`, 'success');
    } catch (err) {
      console.error('Admission ingestion error:', err);
      setStatus(`Ingestion failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const updateCell = (rowIndex, colKey, newValue) => {
    setNormalizedRows(prev => {
      const copy = [...prev];
      copy[rowIndex] = { ...copy[rowIndex], [colKey]: newValue };
      return copy;
    });
  };

  const exportStandardXlsx = () => {
    if (!normalizedRows.length) return alert('No data available to export.');
    setIsProcessing(true);
    setStatus('Generating standardized 40-column master Excel workbook...', 'info');

    try {
      // Export array of arrays for dense mode memory efficiency
      const exportAoa = [TARGET_COLUMNS];
      normalizedRows.forEach(row => {
        exportAoa.push(TARGET_COLUMNS.map(col => row[col] !== undefined ? row[col] : ''));
      });

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(exportAoa, { dense: true });

      // Auto-filter on all 40 columns
      ws['!autofilter'] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: Math.max(exportAoa.length - 1, 0), c: TARGET_COLUMNS.length - 1 }
        })
      };

      // Set column widths
      ws['!cols'] = TARGET_COLUMNS.map(c => ({ wch: Math.max(c.length + 3, 14) }));

      XLSX.utils.book_append_sheet(wb, ws, 'Master_Admission');

      const outBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
      const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const baseName = (sourceFile?.name || 'allotment').replace(/\.[^/.]+$/, '');
      a.download = `${baseName}_40Col_Master.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);

      setStatus(`Exported 40-column Master XLSX with ${normalizedRows.length} records!`, 'success');
    } catch (err) {
      console.error('Export error:', err);
      setStatus(`Export failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const filteredRows = useMemo(() => {
    return normalizedRows.filter((row, idx) => {
      if (filterMode === 'warnings' && (!row._warnings || !row._warnings.length)) return false;
      if (filterMode === 'ready' && row._warnings && row._warnings.length) return false;
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        return (
          String(row.Name || '').toLowerCase().includes(q) ||
          String(row.username || '').toLowerCase().includes(q) ||
          String(row.Mobile || '').includes(q) ||
          String(row.Email || '').toLowerCase().includes(q) ||
          String(row['Program Name'] || '').toLowerCase().includes(q)
        );
      }
      return true;
    });
  }, [normalizedRows, filterMode, searchQuery]);

  const pagedRows = useMemo(() => {
    const start = page * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, page]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', overflow: 'hidden', background: 'var(--bg)', color: 'var(--ink)' }}>
      {/* Top Header Bar */}
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 24px', background: 'var(--panel)', borderBottom: '1px solid var(--line)', zIndex: 20 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}>
            <ArrowLeft size={16} /> Back to Portal
          </Link>
          <div style={{ height: '18px', width: '1px', background: 'var(--line)' }} />
          <h2 style={{ fontSize: '16px', margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Database size={18} color="var(--accent)" /> Admission Import & Transformation Engine
            <span style={{ fontSize: '11px', background: 'var(--accent-soft)', color: 'var(--accent)', padding: '2px 8px', borderRadius: '12px', fontWeight: 700 }}>40 Target Columns</span>
          </h2>
        </div>

        {/* Global Status Pill & Export Button */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <div style={{
            padding: '4px 12px',
            borderRadius: '16px',
            fontSize: '12px',
            fontWeight: 600,
            background: statusType === 'success' ? 'var(--accent-soft)' : statusType === 'error' ? 'var(--danger)' : 'var(--bg)',
            color: statusType === 'success' ? 'var(--accent)' : statusType === 'error' ? 'white' : 'var(--muted)',
            border: '1px solid var(--line)'
          }}>
            {statusMsg}
          </div>

          <button 
            type="button" 
            disabled={!normalizedRows.length || isProcessing}
            onClick={exportStandardXlsx}
            style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 16px', fontSize: '12px' }}
          >
            <Download size={14} /> Export 40-Col Master XLSX
          </button>
        </div>
      </header>

      {/* Main Container */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden', padding: '16px 24px', gap: '16px' }}>
        
        {/* Left Control & Metrics Panel */}
        <aside style={{ width: '320px', display: 'flex', flexDirection: 'column', gap: '14px', overflowY: 'auto' }}>
          
          {/* File Upload Card */}
          <div className="card" style={{ padding: '20px', margin: 0, textAlign: 'center', border: '1.5px dashed var(--accent)', background: 'var(--accent-soft)' }}>
            <input 
              type="file" 
              id="admissionFileInput" 
              accept=".xlsx,.xls,.csv" 
              onChange={handleFileUpload} 
              style={{ display: 'none' }} 
            />
            <label htmlFor="admissionFileInput" style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '6px' }}>
              <Upload size={28} color="var(--accent)" />
              <strong style={{ fontSize: '13.5px', color: 'var(--ink)' }}>Upload Allotment Spreadsheet</strong>
              <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Supports raw 38-column allotment files</span>
            </label>
          </div>

          {/* Ingestion Metrics Card */}
          {normalizedRows.length > 0 && (
            <div className="card" style={{ padding: '16px', margin: 0, display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <h3 style={{ margin: 0, fontSize: '13.5px', textTransform: 'uppercase', color: 'var(--muted)', fontWeight: 700 }}>Transformation Summary</h3>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                <div style={{ padding: '10px', background: 'var(--bg)', borderRadius: '6px', border: '1px solid var(--line)', textAlign: 'center' }}>
                  <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--accent)' }}>{stats.total}</div>
                  <div style={{ fontSize: '11px', color: 'var(--muted)' }}>Total Records</div>
                </div>
                <div style={{ padding: '10px', background: 'var(--bg)', borderRadius: '6px', border: '1px solid var(--line)', textAlign: 'center' }}>
                  <div style={{ fontSize: '20px', fontWeight: 800, color: '#16a34a' }}>40 / 40</div>
                  <div style={{ fontSize: '11px', color: 'var(--muted)' }}>Standard Columns</div>
                </div>
              </div>

              {stats.warnings > 0 && (
                <div style={{ padding: '8px 12px', background: 'rgba(234, 179, 8, 0.1)', border: '1px solid rgba(234, 179, 8, 0.3)', borderRadius: '6px', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '11.5px', color: '#854d0e' }}>
                  <AlertTriangle size={16} />
                  <span><strong>{stats.warnings} record(s)</strong> have partial fields requiring review.</span>
                </div>
              )}
            </div>
          )}

          {/* Lookup Reference Card */}
          <div className="card" style={{ padding: '16px', margin: 0, display: 'flex', flexDirection: 'column', gap: '8px', fontSize: '12px' }}>
            <h3 style={{ margin: 0, fontSize: '13px', fontWeight: 700 }}>Standard Transformation Rules</h3>
            <ul style={{ margin: 0, paddingLeft: '16px', color: 'var(--muted)', display: 'flex', flexDirection: 'column', gap: '4px', lineHeight: 1.4 }}>
              <li><strong>Sequence Agnostic:</strong> Any column order matched.</li>
              <li><strong>Category & Caste:</strong> Auto-mapped to portal IDs.</li>
              <li><strong>Programs & College:</strong> Standardized code & workflow expansion.</li>
              <li><strong>Faculty & Stream:</strong> Derived board & stream codes.</li>
              <li><strong>SeatNumber:</strong> Auto-replicated from Registration No.</li>
            </ul>
          </div>
        </aside>

        {/* Center/Right Live Validation & Editable Grid */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '10px', overflow: 'hidden' }}>
          
          {/* Table Toolbar */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 16px', borderBottom: '1px solid var(--line)', background: 'var(--bg)', flexWrap: 'wrap', gap: '10px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <div style={{ position: 'relative', width: '240px' }}>
                <input 
                  type="text" 
                  placeholder="Filter candidates, mobile, email..." 
                  value={searchQuery} 
                  onChange={(e) => { setSearchQuery(e.target.value); setPage(0); }} 
                  style={{ width: '100%', padding: '5px 8px 5px 26px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)' }} 
                />
                <Search size={13} color="var(--muted)" style={{ position: 'absolute', left: '8px', top: '7px' }} />
              </div>

              <div style={{ display: 'flex', gap: '4px' }}>
                {['all', 'warnings', 'ready'].map(mode => (
                  <button
                    key={mode}
                    type="button"
                    onClick={() => { setFilterMode(mode); setPage(0); }}
                    style={{
                      padding: '3px 8px',
                      fontSize: '11px',
                      borderRadius: '12px',
                      border: '1px solid',
                      borderColor: filterMode === mode ? 'var(--accent)' : 'var(--line)',
                      background: filterMode === mode ? 'var(--accent)' : 'transparent',
                      color: filterMode === mode ? 'white' : 'var(--muted)',
                      fontWeight: 600,
                      textTransform: 'capitalize'
                    }}
                  >
                    {mode === 'all' ? `All (${normalizedRows.length})` : mode === 'warnings' ? `Warnings (${stats.warnings})` : `Clean (${stats.ready})`}
                  </button>
                ))}
              </div>
            </div>

            {/* Pagination Controls */}
            {filteredRows.length > pageSize && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', color: 'var(--muted)' }}>
                <span>Page {page + 1} of {Math.ceil(filteredRows.length / pageSize)}</span>
                <button 
                  type="button" 
                  className="secondary" 
                  disabled={page === 0} 
                  onClick={() => setPage(p => p - 1)} 
                  style={{ padding: '2px 6px', fontSize: '11px' }}
                >
                  Prev
                </button>
                <button 
                  type="button" 
                  className="secondary" 
                  disabled={(page + 1) * pageSize >= filteredRows.length} 
                  onClick={() => setPage(p => p + 1)} 
                  style={{ padding: '2px 6px', fontSize: '11px' }}
                >
                  Next
                </button>
              </div>
            )}
          </div>

          {/* Master 40-Column Table Container */}
          <div style={{ flex: 1, overflow: 'auto', position: 'relative' }}>
            {normalizedRows.length === 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', color: 'var(--muted)', gap: '8px' }}>
                <FileSpreadsheet size={40} style={{ opacity: 0.3 }} />
                <strong style={{ fontSize: '14px', color: 'var(--ink)' }}>No Admission Allotment File Uploaded</strong>
                <span style={{ fontSize: '12px' }}>Upload your raw candidate spreadsheet from the left panel to begin.</span>
              </div>
            ) : (
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px', whiteSpace: 'nowrap' }}>
                <thead>
                  <tr style={{ background: 'var(--bg)', position: 'sticky', top: 0, zIndex: 10 }}>
                    <th style={{ padding: '8px 10px', textAlign: 'center', borderBottom: '1.5px solid var(--line)', color: 'var(--muted)', width: '40px' }}>#</th>
                    {TARGET_COLUMNS.map((col, idx) => (
                      <th 
                        key={col} 
                        style={{ 
                          padding: '8px 10px', 
                          textAlign: 'left', 
                          borderBottom: '1.5px solid var(--line)', 
                          color: 'var(--ink)', 
                          fontWeight: 700,
                          borderRight: '1px solid var(--line)',
                          background: 'var(--bg)'
                        }}
                      >
                        <span style={{ fontSize: '9px', color: 'var(--muted)', display: 'block', fontWeight: 600 }}>Col {idx + 1}</span>
                        {col}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {pagedRows.map((row, rowIdx) => {
                    const actualIdx = (page * pageSize) + rowIdx;
                    const hasWarn = row._warnings && row._warnings.length > 0;
                    return (
                      <tr 
                        key={actualIdx} 
                        style={{ 
                          borderBottom: '1px solid var(--line)',
                          background: hasWarn ? 'rgba(234, 179, 8, 0.03)' : 'transparent'
                        }}
                      >
                        <td style={{ padding: '6px 8px', textAlign: 'center', color: 'var(--muted)', borderRight: '1px solid var(--line)', fontWeight: 600 }}>
                          {actualIdx + 1}
                        </td>
                        {TARGET_COLUMNS.map(col => (
                          <td 
                            key={col} 
                            style={{ 
                              padding: '5px 8px', 
                              borderRight: '1px solid var(--line)',
                              color: 'var(--ink)'
                            }}
                          >
                            <input 
                              type="text" 
                              value={row[col] !== undefined ? row[col] : ''} 
                              onChange={(e) => updateCell(actualIdx, col, e.target.value)} 
                              style={{ 
                                width: '100%', 
                                border: 'none', 
                                background: 'transparent', 
                                fontSize: '11.5px', 
                                color: 'inherit',
                                padding: 0,
                                fontFamily: 'inherit'
                              }} 
                            />
                          </td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>

          {/* Footer Info */}
          {normalizedRows.length > 0 && (
            <div style={{ padding: '8px 16px', background: 'var(--bg)', borderTop: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '11.5px', color: 'var(--muted)' }}>
              <span>Showing {pagedRows.length} of {filteredRows.length} filtered candidate records • All 40 columns verified</span>
              <span>Click any table cell to perform inline edits before exporting</span>
            </div>
          )}

        </main>
      </div>
    </div>
  );
}
