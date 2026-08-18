import { useState, useMemo, useEffect } from "react";
import { Link } from "react-router-dom";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import jsPDF from "jspdf";
import "jspdf-autotable";
import { 
  Building2, 
  Upload, 
  Settings2, 
  Eye, 
  Download, 
  Search, 
  CheckCircle2, 
  RotateCcw, 
  FileSpreadsheet, 
  FileArchive, 
  Sparkles,
  Users,
  GraduationCap,
  BookOpen,
  FileText,
  RefreshCw
} from "lucide-react";
import { readSpreadsheetFile } from "../utils/excelParser";

const SAMPLE_NOMINAL_DATA = [
  { "Venue": "Government College Kasaragod", "Register_No": "KU2025001", "Student_Name": "Muhammed Rashid K", "Course_Code": "ENG101", "Course_Title": "English Literature", "Session": "FN" },
  { "Venue": "Government College Kasaragod", "Register_No": "KU2025002", "Student_Name": "Ananya S Nair", "Course_Code": "ENG101", "Course_Title": "English Literature", "Session": "FN" },
  { "Venue": "Government College Kasaragod", "Register_No": "KU2025003", "Student_Name": "Fathima Hameed", "Course_Code": "MAL102", "Course_Title": "Malayalam Poetry", "Session": "AN" },
  { "Venue": "Payyanur College", "Register_No": "KU2025004", "Student_Name": "Abhijith T", "Course_Code": "CS104", "Course_Title": "Data Structures", "Session": "FN" },
  { "Venue": "Payyanur College", "Register_No": "KU2025005", "Student_Name": "Devika Menon", "Course_Code": "CS104", "Course_Title": "Data Structures", "Session": "FN" },
  { "Venue": "Sir Syed College Taliparamba", "Register_No": "KU2025006", "Student_Name": "Rahul K V", "Course_Code": "CHE106", "Course_Title": "Organic Chemistry", "Session": "FN" },
  { "Venue": "Sir Syed College Taliparamba", "Register_No": "KU2025007", "Student_Name": "Sneha Prakash", "Course_Code": "CHE106", "Course_Title": "Organic Chemistry", "Session": "FN" }
];

const SllNominalPage = () => {
  const [currentStep, setCurrentStep] = useState(1);
  const [rows, setRows] = useState([]);
  const [columns, setColumns] = useState([]);
  const [fileName, setFileName] = useState("");
  const [isLoading, setIsLoading] = useState(false);

  // Field Mappings
  const [venueCol, setVenueCol] = useState("");
  const [regNoCol, setRegNoCol] = useState("");
  const [nameCol, setNameCol] = useState("");
  const [courseCodeCol, setCourseCodeCol] = useState("");
  const [courseTitleCol, setCourseTitleCol] = useState("");
  const [sessionCol, setSessionCol] = useState("");

  // Step 3 Filter States
  const [selectedVenue, setSelectedVenue] = useState("all");
  const [searchQuery, setSearchQuery] = useState("");

  // Header customization
  const [examTitle, setExamTitle] = useState("KANNUR UNIVERSITY - EXAMINATION BRANCH");
  const [subTitle, setSubTitle] = useState("CANDIDATE NOMINAL ROLL & ATTENDANCE RECORD");

  // Load sample dataset
  const loadSample = () => {
    setRows(SAMPLE_NOMINAL_DATA);
    const cols = Object.keys(SAMPLE_NOMINAL_DATA[0]);
    setColumns(cols);
    setFileName("Sample_Nominal_Roll.xlsx");
    setVenueCol("Venue");
    setRegNoCol("Register_No");
    setNameCol("Student_Name");
    setCourseCodeCol("Course_Code");
    setCourseTitleCol("Course_Title");
    setSessionCol("Session");
  };

  // Auto-detect columns on file load
  useEffect(() => {
    if (columns.length > 0) {
      const findCol = (regex) => columns.find(c => regex.test(c)) || "";
      setVenueCol(prev => prev || findCol(/venue|college|center|centre|institution/i) || columns[0]);
      setRegNoCol(prev => prev || findCol(/reg|prn|roll|candidate_id|id/i) || columns[1] || columns[0]);
      setNameCol(prev => prev || findCol(/name|student|candidate/i) || columns[2] || columns[0]);
      setCourseCodeCol(prev => prev || findCol(/course.*code|sub.*code|qp.*code|code/i) || columns[3] || columns[0]);
      setCourseTitleCol(prev => prev || findCol(/course.*title|course.*name|subject|title/i) || columns[4] || "");
      setSessionCol(prev => prev || findCol(/session|time|date|semester|sem/i) || "");
    }
  }, [columns]);

  // Universal File Upload Handler
  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    try {
      const { rows: dataRows, columns: dataCols } = await readSpreadsheetFile(file);
      setRows(dataRows);
      setColumns(dataCols);
      setFileName(file.name);
    } catch (err) {
      alert("Error reading file: " + err.message);
    } finally {
      setIsLoading(false);
    }
  };

  // Grouped Venue Data
  const venueList = useMemo(() => {
    if (!rows.length || !venueCol) return [];
    const set = new Set();
    rows.forEach(r => {
      const v = String(r[venueCol] || "").trim();
      if (v) set.add(v);
    });
    return Array.from(set).sort();
  }, [rows, venueCol]);

  const venueGroups = useMemo(() => {
    if (!rows.length || !venueCol) return {};
    const groups = {};
    rows.forEach(r => {
      const v = String(r[venueCol] || "Unassigned Venue").trim();
      if (!groups[v]) groups[v] = [];
      groups[v].push(r);
    });
    return groups;
  }, [rows, venueCol]);

  // Filtered Rows for Preview
  const displayedRows = useMemo(() => {
    let list = rows;
    if (selectedVenue !== "all") {
      list = venueGroups[selectedVenue] || [];
    }
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      list = list.filter(r => 
        String(r[nameCol] || "").toLowerCase().includes(q) ||
        String(r[regNoCol] || "").toLowerCase().includes(q) ||
        String(r[courseCodeCol] || "").toLowerCase().includes(q) ||
        String(r[venueCol] || "").toLowerCase().includes(q)
      );
    }
    return list;
  }, [rows, selectedVenue, venueGroups, searchQuery, nameCol, regNoCol, courseCodeCol, venueCol]);

  // Unique course count
  const courseCount = useMemo(() => {
    if (!rows.length || !courseCodeCol) return 0;
    const set = new Set(rows.map(r => String(r[courseCodeCol] || "").trim()).filter(Boolean));
    return set.size;
  }, [rows, courseCodeCol]);

  // EXPORT: Single Venue / Consolidated PDF
  const exportPdf = (targetVenue = null) => {
    const doc = new jsPDF("p", "mm", "a4");
    const venuesToPrint = targetVenue ? [targetVenue] : (selectedVenue === "all" ? venueList : [selectedVenue]);

    venuesToPrint.forEach((vName, vIdx) => {
      if (vIdx > 0) doc.addPage();
      const vRows = venueGroups[vName] || [];

      // Header
      doc.setFontSize(14);
      doc.setFont("helvetica", "bold");
      doc.text(examTitle, 105, 15, { align: "center" });

      doc.setFontSize(10);
      doc.setFont("helvetica", "normal");
      doc.text(subTitle, 105, 21, { align: "center" });

      // Venue Info Box
      doc.setDrawColor(200, 200, 200);
      doc.setFillColor(245, 247, 250);
      doc.roundedRect(14, 25, 182, 14, 2, 2, "FD");

      doc.setFontSize(10);
      doc.setFont("helvetica", "bold");
      doc.text(`Venue / Center: ${vName}`, 18, 31);
      doc.setFont("helvetica", "normal");
      doc.text(`Total Candidates: ${vRows.length}`, 18, 36);

      // Table Content
      const tableData = vRows.map((r, idx) => [
        idx + 1,
        String(r[regNoCol] || ""),
        String(r[nameCol] || ""),
        String(r[courseCodeCol] || ""),
        String(r[courseTitleCol] || ""),
        String(r[sessionCol] || ""),
        "" // Signature box column
      ]);

      doc.autoTable({
        startY: 42,
        head: [["#", "Register No", "Candidate Name", "Course", "Title", "Session", "Candidate Signature"]],
        body: tableData,
        theme: "grid",
        headStyles: { fillColor: [23, 107, 135], textColor: 255, fontSize: 9, fontStyle: "bold" },
        bodyStyles: { fontSize: 8.5, cellPadding: 3 },
        columnStyles: {
          0: { cellWidth: 10, halign: "center" },
          1: { cellWidth: 28, fontStyle: "bold" },
          2: { cellWidth: 42 },
          3: { cellWidth: 20 },
          4: { cellWidth: 38 },
          5: { cellWidth: 16, halign: "center" },
          6: { cellWidth: 28 }
        },
        didDrawPage: () => {
          doc.setFontSize(8);
          doc.setTextColor(100);
          doc.text(`Page ${doc.internal.getNumberOfPages()}`, 105, 290, { align: "center" });
        }
      });
    });

    const outName = targetVenue ? `Nominal_Roll_${targetVenue.replace(/[^a-zA-Z0-9]/g, "_")}.pdf` : `Master_Nominal_Roll_All_Venues.pdf`;
    doc.save(outName);
  };

  // EXPORT: Multi-Sheet Master Excel
  const exportMasterExcel = () => {
    const wb = XLSX.utils.book_new();

    // Summary Sheet
    const summaryData = [
      ["VENUE-WISE NOMINAL ROLL SUMMARY"],
      ["Generated At", new Date().toLocaleString()],
      ["Total Venues", venueList.length],
      ["Total Candidates", rows.length],
      ["Total Courses", courseCount],
      [""],
      ["VENUE NAME", "CANDIDATE COUNT"]
    ];

    venueList.forEach(v => {
      summaryData.push([v, (venueGroups[v] || []).length]);
    });

    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");

    // Sheet per venue
    venueList.forEach(v => {
      const vRows = (venueGroups[v] || []).map((r, idx) => ({
        "Sl_No": idx + 1,
        "Register_No": r[regNoCol] || "",
        "Student_Name": r[nameCol] || "",
        "Course_Code": r[courseCodeCol] || "",
        "Course_Title": r[courseTitleCol] || "",
        "Session": r[sessionCol] || ""
      }));
      const ws = XLSX.utils.json_to_sheet(vRows);
      const sheetName = v.replace(/[/\\?%*:|"<>]/g, "_").slice(0, 30);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    XLSX.writeFile(wb, "Master_Venue_Nominal_Roll.xlsx");
  };

  // EXPORT: ZIP Archive with individual Venue PDFs
  const exportZipBundle = async () => {
    const zip = new JSZip();

    venueList.forEach(vName => {
      const vDoc = new jsPDF("p", "mm", "a4");
      const vRows = venueGroups[vName] || [];

      vDoc.setFontSize(13);
      vDoc.setFont("helvetica", "bold");
      vDoc.text(examTitle, 105, 15, { align: "center" });

      vDoc.setFontSize(9.5);
      vDoc.setFont("helvetica", "normal");
      vDoc.text(subTitle, 105, 21, { align: "center" });

      vDoc.setDrawColor(200, 200, 200);
      vDoc.setFillColor(245, 247, 250);
      vDoc.roundedRect(14, 25, 182, 14, 2, 2, "FD");

      vDoc.setFontSize(9.5);
      vDoc.setFont("helvetica", "bold");
      vDoc.text(`Venue: ${vName}`, 18, 31);
      vDoc.setFont("helvetica", "normal");
      vDoc.text(`Candidate Count: ${vRows.length}`, 18, 36);

      const tableData = vRows.map((r, idx) => [
        idx + 1,
        String(r[regNoCol] || ""),
        String(r[nameCol] || ""),
        String(r[courseCodeCol] || ""),
        String(r[courseTitleCol] || ""),
        String(r[sessionCol] || ""),
        ""
      ]);

      vDoc.autoTable({
        startY: 42,
        head: [["#", "Register No", "Candidate Name", "Course", "Title", "Session", "Signature"]],
        body: tableData,
        theme: "grid",
        headStyles: { fillColor: [23, 107, 135], textColor: 255, fontSize: 8.5 },
        bodyStyles: { fontSize: 8, cellPadding: 2.5 }
      });

      const pdfBlob = vDoc.output("blob");
      const fileName = `Nominal_Roll_${vName.replace(/[/\\?%*:|"<>]/g, "_")}.pdf`;
      zip.file(fileName, pdfBlob);
    });

    const zipBlob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(zipBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Venue_Nominal_Rolls_Bundle.zip";
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div style={{ maxWidth: "1400px", margin: "0 auto", padding: "24px 20px 80px", fontFamily: "var(--font-family)" }}>
      
      {/* Top Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "20px", flexWrap: "wrap", gap: "16px" }}>
        <div>
          <Link to="/" style={{ textDecoration: "none", color: "var(--accent)", fontWeight: 600, display: "inline-flex", alignItems: "center", gap: "6px", fontSize: "13.5px", marginBottom: "6px" }}>
            ← Back to Portal
          </Link>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <div style={{ background: "var(--accent)", color: "white", padding: "10px", borderRadius: "10px", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Building2 size={24} />
            </div>
            <div>
              <h1 style={{ margin: 0, fontSize: "24px", fontWeight: 800 }}>Venue-Wise Nominal Roll Generator</h1>
              <p style={{ margin: "4px 0 0 0", color: "var(--muted)", fontSize: "13.5px" }}>
                Generate venue-partitioned candidate nominal rolls, attendance signature sheets, and batch exports.
              </p>
            </div>
          </div>
        </div>

        <button 
          onClick={loadSample}
          className="button secondary"
          style={{ display: "inline-flex", alignItems: "center", gap: "6px", fontSize: "13px", padding: "8px 16px" }}
        >
          <Sparkles size={15} /> Load Sample Data
        </button>
      </div>

      {/* 4-Step Stepper Header */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "12px", marginBottom: "28px" }}>
        {[
          { step: 1, label: "1. Upload Data", icon: <Upload size={16} />, active: currentStep === 1, done: rows.length > 0 },
          { step: 2, label: "2. Column Mapping", icon: <Settings2 size={16} />, active: currentStep === 2, done: !!(venueCol && regNoCol && nameCol) },
          { step: 3, label: "3. Venue Preview", icon: <Eye size={16} />, active: currentStep === 3, done: currentStep > 3 },
          { step: 4, label: "4. Export Studio", icon: <Download size={16} />, active: currentStep === 4, done: false }
        ].map(s => (
          <div
            key={s.step}
            onClick={() => rows.length > 0 && setCurrentStep(s.step)}
            style={{
              display: "flex",
              alignItems: "center",
              gap: "8px",
              padding: "12px 14px",
              borderRadius: "10px",
              background: s.active ? "var(--accent)" : "var(--panel)",
              color: s.active ? "white" : s.done ? "var(--ink)" : "var(--muted)",
              border: s.active ? "1.5px solid var(--accent)" : "1px solid var(--line)",
              cursor: rows.length > 0 ? "pointer" : "default",
              fontWeight: 700,
              fontSize: "13px",
              transition: "all 0.15s ease"
            }}
          >
            {s.done && !s.active ? <CheckCircle2 size={16} color="#10b981" /> : s.icon}
            <span>{s.label}</span>
          </div>
        ))}
      </div>

      {/* STEP 1: UPLOAD DATA */}
      {currentStep === 1 && (
        <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
          
          <div className="card" style={{ padding: "28px" }}>
            <div style={{ border: "2px dashed var(--line)", borderRadius: "12px", padding: "36px", textAlign: "center", background: "var(--bg)" }}>
              <Upload size={36} color="var(--accent)" style={{ margin: "0 auto 12px", opacity: 0.8 }} />
              <div style={{ fontWeight: 800, fontSize: "16px", marginBottom: "6px" }}>
                {fileName ? fileName : "Upload Nominal Roll Spreadsheet (.xlsx, .xls, .csv, .zip)"}
              </div>
              <p style={{ color: "var(--muted)", fontSize: "13px", margin: "0 0 18px 0" }}>
                Supports university exam rosters, multi-file collections, and auto-extracting ZIP archives
              </p>
              <label className="button" style={{ display: "inline-flex", alignItems: "center", gap: "6px", cursor: isLoading ? "wait" : "pointer", padding: "10px 22px", fontSize: "14px", opacity: isLoading ? 0.7 : 1 }}>
                {isLoading ? <RefreshCw size={16} className="spin" /> : <Upload size={16} />}
                {isLoading ? "Reading & Extracting..." : "Browse File"}
                <input type="file" accept=".xlsx, .xls, .csv, .zip" onChange={handleFileUpload} disabled={isLoading} style={{ display: "none" }} />
              </label>
            </div>

            {columns.length > 0 && (
              <div style={{ marginTop: "24px" }}>
                <div style={{ fontSize: "12px", fontWeight: 700, textTransform: "uppercase", color: "var(--muted)", marginBottom: "10px" }}>
                  Detected Columns ({columns.length})
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: "8px", marginBottom: "20px" }}>
                  {columns.map(c => (
                    <span key={c} style={{ fontSize: "12px", background: "var(--bg)", border: "1px solid var(--line)", padding: "4px 10px", borderRadius: "6px", fontWeight: 600 }}>
                      {c}
                    </span>
                  ))}
                </div>

                <div style={{ fontSize: "12px", fontWeight: 700, textTransform: "uppercase", color: "var(--muted)", marginBottom: "10px" }}>
                  Data Preview (First 5 Rows)
                </div>
                <div style={{ overflowX: "auto", border: "1px solid var(--line)", borderRadius: "8px", maxHeight: "200px" }}>
                  <table style={{ width: "100%", fontSize: "12px", borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: "var(--bg)", borderBottom: "1px solid var(--line)" }}>
                        {columns.slice(0, 6).map(c => <th key={c} style={{ padding: "8px 10px", textAlign: "left" }}>{c}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.slice(0, 5).map((r, idx) => (
                        <tr key={idx} style={{ borderBottom: "1px solid var(--line)" }}>
                          {columns.slice(0, 6).map(c => <td key={c} style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{String(r[c] || "")}</td>)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>

          {rows.length > 0 && (
            <div style={{ display: "flex", justifyContent: "flex-end" }}>
              <button 
                onClick={() => setCurrentStep(2)}
                style={{ padding: "12px 28px", fontSize: "14px", fontWeight: 700, display: "inline-flex", alignItems: "center", gap: "6px" }}
              >
                Proceed to Column Mapping →
              </button>
            </div>
          )}

        </div>
      )}

      {/* STEP 2: COLUMN MAPPING */}
      {currentStep === 2 && (
        <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
          <div className="card" style={{ padding: "28px" }}>
            <h3 style={{ margin: "0 0 6px 0", fontSize: "18px", fontWeight: 800 }}>Map Dataset Attributes</h3>
            <p style={{ margin: "0 0 20px 0", color: "var(--muted)", fontSize: "13.5px" }}>
              Confirm which columns correspond to the venue, candidate register number, student name, and course details.
            </p>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: "18px" }}>
              
              <div>
                <label style={{ display: "block", fontSize: "13px", fontWeight: 700, marginBottom: "6px" }}>
                  Venue / Center Column <span style={{ color: "var(--danger)" }}>*</span>
                </label>
                <select value={venueCol} onChange={(e) => setVenueCol(e.target.value)} style={{ width: "100%", padding: "10px", fontSize: "13.5px" }}>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div>
                <label style={{ display: "block", fontSize: "13px", fontWeight: 700, marginBottom: "6px" }}>
                  Register No / Candidate ID <span style={{ color: "var(--danger)" }}>*</span>
                </label>
                <select value={regNoCol} onChange={(e) => setRegNoCol(e.target.value)} style={{ width: "100%", padding: "10px", fontSize: "13.5px" }}>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div>
                <label style={{ display: "block", fontSize: "13px", fontWeight: 700, marginBottom: "6px" }}>
                  Candidate / Student Name <span style={{ color: "var(--danger)" }}>*</span>
                </label>
                <select value={nameCol} onChange={(e) => setNameCol(e.target.value)} style={{ width: "100%", padding: "10px", fontSize: "13.5px" }}>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div>
                <label style={{ display: "block", fontSize: "13px", fontWeight: 700, marginBottom: "6px" }}>
                  Course / Subject Code
                </label>
                <select value={courseCodeCol} onChange={(e) => setCourseCodeCol(e.target.value)} style={{ width: "100%", padding: "10px", fontSize: "13.5px" }}>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div>
                <label style={{ display: "block", fontSize: "13px", fontWeight: 700, marginBottom: "6px" }}>
                  Course Title / Name (Optional)
                </label>
                <select value={courseTitleCol} onChange={(e) => setCourseTitleCol(e.target.value)} style={{ width: "100%", padding: "10px", fontSize: "13.5px" }}>
                  <option value="">-- None / Blank --</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              <div>
                <label style={{ display: "block", fontSize: "13px", fontWeight: 700, marginBottom: "6px" }}>
                  Exam Session / Date (Optional)
                </label>
                <select value={sessionCol} onChange={(e) => setSessionCol(e.target.value)} style={{ width: "100%", padding: "10px", fontSize: "13.5px" }}>
                  <option value="">-- None / Blank --</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

            </div>
          </div>

          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <button className="secondary" onClick={() => setCurrentStep(1)}>
              ← Back to Upload
            </button>
            <button onClick={() => setCurrentStep(3)} style={{ padding: "12px 28px", fontSize: "14px", fontWeight: 700 }}>
              Proceed to Venue Selection & Preview →
            </button>
          </div>
        </div>
      )}

      {/* STEP 3: VENUE SELECTION & PREVIEW */}
      {currentStep === 3 && (
        <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
          
          {/* KPI Cards Row */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "16px" }}>
            <div className="card" style={{ padding: "20px", display: "flex", alignItems: "center", gap: "14px" }}>
              <div style={{ background: "rgba(23, 107, 135, 0.12)", color: "var(--accent)", padding: "12px", borderRadius: "10px" }}>
                <Building2 size={24} />
              </div>
              <div>
                <div style={{ fontSize: "12px", color: "var(--muted)", fontWeight: 600 }}>Total Venues</div>
                <div style={{ fontSize: "22px", fontWeight: 800, color: "var(--ink)" }}>{venueList.length}</div>
              </div>
            </div>

            <div className="card" style={{ padding: "20px", display: "flex", alignItems: "center", gap: "14px" }}>
              <div style={{ background: "rgba(16, 185, 129, 0.12)", color: "#10b981", padding: "12px", borderRadius: "10px" }}>
                <Users size={24} />
              </div>
              <div>
                <div style={{ fontSize: "12px", color: "var(--muted)", fontWeight: 600 }}>Total Candidates</div>
                <div style={{ fontSize: "22px", fontWeight: 800, color: "var(--ink)" }}>{rows.length}</div>
              </div>
            </div>

            <div className="card" style={{ padding: "20px", display: "flex", alignItems: "center", gap: "14px" }}>
              <div style={{ background: "rgba(139, 92, 246, 0.12)", color: "#8b5cf6", padding: "12px", borderRadius: "10px" }}>
                <BookOpen size={24} />
              </div>
              <div>
                <div style={{ fontSize: "12px", color: "var(--muted)", fontWeight: 600 }}>Total Courses</div>
                <div style={{ fontSize: "22px", fontWeight: 800, color: "var(--ink)" }}>{courseCount}</div>
              </div>
            </div>

            <div className="card" style={{ padding: "20px", display: "flex", alignItems: "center", gap: "14px" }}>
              <div style={{ background: "rgba(245, 158, 11, 0.12)", color: "#f59e0b", padding: "12px", borderRadius: "10px" }}>
                <GraduationCap size={24} />
              </div>
              <div>
                <div style={{ fontSize: "12px", color: "var(--muted)", fontWeight: 600 }}>Active Venue Candidates</div>
                <div style={{ fontSize: "22px", fontWeight: 800, color: "var(--ink)" }}>{displayedRows.length}</div>
              </div>
            </div>
          </div>

          {/* Filter & Venue Selector Card */}
          <div className="card" style={{ padding: "24px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "16px", marginBottom: "20px" }}>
              
              {/* Venue Dropdown */}
              <div style={{ display: "flex", alignItems: "center", gap: "10px", flex: 1, minWidth: "280px" }}>
                <span style={{ fontSize: "13.5px", fontWeight: 700, color: "var(--ink)", whiteSpace: "nowrap" }}>
                  Filter Venue:
                </span>
                <select 
                  value={selectedVenue} 
                  onChange={(e) => setSelectedVenue(e.target.value)}
                  style={{ flex: 1, padding: "9px 14px", borderRadius: "8px", fontSize: "13.5px", fontWeight: 600 }}
                >
                  <option value="all">🌟 All Venues / Centers ({venueList.length})</option>
                  {venueList.map(v => (
                    <option key={v} value={v}>
                      🏛️ {v} ({(venueGroups[v] || []).length} Candidates)
                    </option>
                  ))}
                </select>
              </div>

              {/* Search Box */}
              <div style={{ position: "relative", minWidth: "240px" }}>
                <Search size={16} style={{ position: "absolute", left: "12px", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                <input 
                  type="text" 
                  placeholder="Search candidate, reg no..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  style={{ width: "100%", paddingLeft: "36px", fontSize: "13px" }}
                />
              </div>

            </div>

            {/* Candidate Table Grid */}
            <div style={{ overflowX: "auto", maxHeight: "450px", border: "1px solid var(--line)", borderRadius: "8px" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12.5px" }}>
                <thead>
                  <tr style={{ background: "var(--bg)", borderBottom: "2px solid var(--line)", position: "sticky", top: 0 }}>
                    <th style={{ padding: "10px 12px", width: "40px", textAlign: "center" }}>#</th>
                    <th style={{ padding: "10px 12px", textAlign: "left" }}>Register No</th>
                    <th style={{ padding: "10px 12px", textAlign: "left" }}>Candidate Name</th>
                    <th style={{ padding: "10px 12px", textAlign: "left" }}>Course Code & Title</th>
                    <th style={{ padding: "10px 12px", textAlign: "left" }}>Venue / Center</th>
                    <th style={{ padding: "10px 12px", textAlign: "center" }}>Session</th>
                  </tr>
                </thead>
                <tbody>
                  {displayedRows.length > 0 ? (
                    displayedRows.map((r, idx) => (
                      <tr key={idx} style={{ borderBottom: "1px solid var(--line)" }}>
                        <td style={{ padding: "10px 12px", textAlign: "center", color: "var(--muted)" }}>{idx + 1}</td>
                        <td style={{ padding: "10px 12px", fontWeight: 700, color: "var(--accent)" }}>{String(r[regNoCol] || "")}</td>
                        <td style={{ padding: "10px 12px", fontWeight: 600 }}>{String(r[nameCol] || "")}</td>
                        <td style={{ padding: "10px 12px" }}>
                          <div><strong>{String(r[courseCodeCol] || "")}</strong></div>
                          {courseTitleCol && <div style={{ fontSize: "11px", color: "var(--muted)" }}>{String(r[courseTitleCol] || "")}</div>}
                        </td>
                        <td style={{ padding: "10px 12px", color: "var(--muted)" }}>{String(r[venueCol] || "")}</td>
                        <td style={{ padding: "10px 12px", textAlign: "center" }}>
                          <span style={{ fontSize: "11px", padding: "2px 8px", borderRadius: "10px", background: "var(--bg)", border: "1px solid var(--line)", fontWeight: 600 }}>
                            {String(r[sessionCol] || "FN")}
                          </span>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={6} style={{ textAlign: "center", padding: "30px", color: "var(--muted)" }}>
                        No records match the current filter and search query.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <button className="secondary" onClick={() => setCurrentStep(2)}>
              ← Back to Mapping
            </button>
            <button onClick={() => setCurrentStep(4)} style={{ padding: "12px 28px", fontSize: "14px", fontWeight: 700 }}>
              Proceed to Export Studio →
            </button>
          </div>
        </div>
      )}

      {/* STEP 4: EXPORT STUDIO */}
      {currentStep === 4 && (
        <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
          
          {/* Header Title Config */}
          <div className="card" style={{ padding: "24px" }}>
            <h3 style={{ margin: "0 0 14px 0", fontSize: "17px", fontWeight: 800 }}>Document Header Customization</h3>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: "16px" }}>
              <div>
                <label style={{ display: "block", fontSize: "12.5px", fontWeight: 700, marginBottom: "4px" }}>Institution / University Header:</label>
                <input type="text" value={examTitle} onChange={(e) => setExamTitle(e.target.value)} style={{ width: "100%", fontSize: "13px" }} />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "12.5px", fontWeight: 700, marginBottom: "4px" }}>Report Subtitle / Heading:</label>
                <input type="text" value={subTitle} onChange={(e) => setSubTitle(e.target.value)} style={{ width: "100%", fontSize: "13px" }} />
              </div>
            </div>
          </div>

          {/* Export Options Cards */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: "20px" }}>
            
            {/* Consolidated PDF */}
            <div className="card" style={{ padding: "24px", display: "flex", flexDirection: "column", justifyContent: "space-between", border: "1.5px solid var(--accent)" }}>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "12px" }}>
                  <div style={{ background: "var(--accent)", color: "white", padding: "8px", borderRadius: "8px" }}>
                    <FileText size={20} />
                  </div>
                  <h3 style={{ margin: 0, fontSize: "17px", fontWeight: 800 }}>Master PDF Nominal Roll</h3>
                </div>
                <p style={{ color: "var(--muted)", fontSize: "13px", lineHeight: "1.4" }}>
                  Generates an all-in-one PDF document with a dedicated nominal roll page and candidate signature box for every venue.
                </p>
              </div>
              <button onClick={() => exportPdf()} style={{ width: "100%", padding: "10px", fontSize: "13.5px", fontWeight: 700, marginTop: "16px", display: "inline-flex", alignItems: "center", justifyContent: "center", gap: "6px" }}>
                <Download size={16} /> Download Master PDF
              </button>
            </div>

            {/* Master Excel */}
            <div className="card" style={{ padding: "24px", display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "12px" }}>
                  <div style={{ background: "rgba(16, 185, 129, 0.12)", color: "#10b981", padding: "8px", borderRadius: "8px" }}>
                    <FileSpreadsheet size={20} />
                  </div>
                  <h3 style={{ margin: 0, fontSize: "17px", fontWeight: 800 }}>Multi-Sheet Excel (.xlsx)</h3>
                </div>
                <p style={{ color: "var(--muted)", fontSize: "13px", lineHeight: "1.4" }}>
                  Generates a master Excel workbook containing a Summary dashboard sheet and a separate tab for each venue.
                </p>
              </div>
              <button onClick={exportMasterExcel} className="secondary" style={{ width: "100%", padding: "10px", fontSize: "13.5px", fontWeight: 700, marginTop: "16px", display: "inline-flex", alignItems: "center", justifyContent: "center", gap: "6px" }}>
                <FileSpreadsheet size={16} /> Download Master Excel
              </button>
            </div>

            {/* ZIP Bundle */}
            <div className="card" style={{ padding: "24px", display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "12px" }}>
                  <div style={{ background: "rgba(139, 92, 246, 0.12)", color: "#8b5cf6", padding: "8px", borderRadius: "8px" }}>
                    <FileArchive size={20} />
                  </div>
                  <h3 style={{ margin: 0, fontSize: "17px", fontWeight: 800 }}>Individual Venue PDFs (.zip)</h3>
                </div>
                <p style={{ color: "var(--muted)", fontSize: "13px", lineHeight: "1.4" }}>
                  Packages separate PDF nominal roll documents for each venue bundled neatly into a single ZIP archive.
                </p>
              </div>
              <button onClick={exportZipBundle} className="secondary" style={{ width: "100%", padding: "10px", fontSize: "13.5px", fontWeight: 700, marginTop: "16px", display: "inline-flex", alignItems: "center", justifyContent: "center", gap: "6px" }}>
                <FileArchive size={16} /> Download ZIP Bundle
              </button>
            </div>

          </div>

          <div style={{ display: "flex", justifyContent: "space-between", marginTop: "8px" }}>
            <button className="secondary" onClick={() => setCurrentStep(3)}>
              ← Back to Venue Preview
            </button>
            <button 
              className="secondary" 
              onClick={() => {
                setCurrentStep(1);
                setRows([]);
              }}
              style={{ display: "inline-flex", alignItems: "center", gap: "6px" }}
            >
              <RotateCcw size={15} /> Start New Nominal Roll
            </button>
          </div>

        </div>
      )}

    </div>
  );
};

export default SllNominalPage;
