export const TOOL_GUIDES = {
  'studio': {
    title: 'Report Studio - Extraction & Generation Guide',
    icon: 'LayoutTemplate',
    description: 'Build custom data-driven templates and generate hundreds of printable reports in seconds.',
    sections: [
      {
        heading: '1. Input Data & Templating Engine',
        content: 'Accepts Excel (.xlsx, .xls) and CSV files. Template placeholders use double curly braces {{Column_Name}} corresponding to headers in your data sheet.'
      },
      {
        heading: '2. Conditional & Loop Directives',
        content: 'Supports conditional blocks (e.g. {{#if Condition}}...{{/if}}) and repeating table rows for multi-course student rosters.'
      },
      {
        heading: '3. Output Formats',
        content: 'Generates high-resolution vector PDFs, combined multi-page documents, and printer-ready batches with custom page breaks.'
      }
    ]
  },
  'revaluation': {
    title: 'Revaluation - Extraction Logic & Guide',
    icon: 'TableProperties',
    description: 'Merge multiple application reports with result sheets to generate combined datasets and final PDFs.',
    sections: [
      {
        heading: '1. Input Sheets',
        content: 'Upload the Revaluation Application Excel and the Original Results Master Sheet.'
      },
      {
        heading: '2. Matching & Difference Calculations',
        content: 'Matches records using Register No / Student ID and Subject/Course Code. Calculates (Updated Marks - Original Marks) and detects status changes (Pass/Fail).'
      },
      {
        heading: '3. Output Master & PDFs',
        content: 'Produces combined audit spreadsheets and formatted university notification notification PDFs.'
      }
    ]
  },
  'splitter': {
    title: 'Excel Lot Splitter - Extraction Logic & Guide',
    icon: 'FileDown',
    description: 'Split a large Excel worksheet into smaller lots/chunks and download them bundled in a ZIP archive.',
    sections: [
      {
        heading: '1. Splitting Modes',
        content: 'Choose between "Fixed Row Count" (e.g. 500 rows per file) or "Group by Column" (e.g. Center Code, College, or Program).'
      },
      {
        heading: '2. Header & Style Preservation',
        content: 'Preserves the original source header row in every split file with auto-fitted column widths.'
      },
      {
        heading: '3. ZIP Archiving',
        content: 'All split chunks are generated in-memory and bundled into a single ZIP archive for fast download.'
      }
    ]
  },
  'merger': {
    title: 'Excel Sheet Merger - Extraction Logic & Guide',
    icon: 'FileStack',
    description: 'Combine multiple Excel sheets into a single document with sequence-agnostic header alignment and strict text preservation for 15+ digit transaction reference numbers.',
    sections: [
      {
        heading: '1. Multi-File Ingestion',
        content: 'Upload any number of Excel workbooks or CSV files simultaneously, or upload a ZIP containing all target files.'
      },
      {
        heading: '2. Sequence-Agnostic Header Alignment',
        content: 'Performs sequence-agnostic column union across all worksheets so columns align under matching headers regardless of order.'
      },
      {
        heading: '3. 15–18 Digit Text & Reference Protection',
        content: 'Headers like CLIENTAPPTRANSCATIONREFERENCENUMBER, TRANSCATIONREFERENCENUMBER, DESCRIPTION, OTHERPARAMETERS, REF1 and 12+ digit numbers are explicitly preserved as text (@) to prevent scientific notation and precision loss.'
      },
      {
        heading: '4. Export',
        content: 'Outputs a consolidated single-sheet master XLSX with autofilter and auto-fitted column widths.'
      }
    ]
  },
  'sll-nominal': {
    title: 'Venue-Wise Nominal Roll - Extraction Logic & Guide',
    icon: 'FileText',
    description: 'Generate venue-wise nominal roll sheets and PDFs with merged student registration cells.',
    sections: [
      {
        heading: '1. Input Requirements',
        content: 'Student registration roster with Center, Venue/Hall, Register No, Student Name, Subject Codes.'
      },
      {
        heading: '2. Venue Grouping & Cell Merging',
        content: 'Groups students by Center and Hall. Multi-subject registrations for the same student are merged across adjacent rows.'
      },
      {
        heading: '3. Attendance & Signature Grids',
        content: 'Generates invigilator attendance sheets with signature boxes and barcode header strips.'
      }
    ]
  },
  'qp-statement': {
    title: 'QP Statement Report - Extraction Logic & Guide',
    icon: 'CalendarRange',
    description: 'Compile daily printing lists and venue packing slips for examination question papers.',
    sections: [
      {
        heading: '1. Date & Session Isolation',
        content: 'Filters registrations by Exam Date and Session (Morning / Afternoon).'
      },
      {
        heading: '2. QP Code & Packet Count Aggregation',
        content: 'Calculates total candidate counts per Question Paper code across all participating examination centers.'
      },
      {
        heading: '3. Packing Slips',
        content: 'Generates venue-wise packet packing slips and master printing manifests.'
      }
    ]
  },
  'qp-label': {
    title: 'QP Label Generator - Extraction Logic & Guide',
    icon: 'TableProperties',
    description: 'Generate packet covers and Question Paper envelope labels sorted by center and subject.',
    sections: [
      {
        heading: '1. Input Data',
        content: 'Center allotments, QP Codes, Subject Names, Date, Session, and Packet Sizes.'
      },
      {
        heading: '2. Label Layout Calculation',
        content: 'Computes standard envelope stickers (e.g. 2x4 or 3x8 per A4 sheet) with center codes, subject title, QP code, and student count.'
      },
      {
        heading: '3. Print Output',
        content: 'Direct-to-print formatted sticker grids.'
      }
    ]
  },
  'compare': {
    title: 'Data Comparison & Reconciliation - Logic & Guide',
    icon: 'GitCompare',
    description: 'Fuzzy match & compare Excel datasets, identify partial matches, detect discrepancies, and export audit reports.',
    sections: [
      {
        heading: '1. Primary & Secondary Datasets',
        content: 'Upload Dataset A (e.g. Portal Master) and Dataset B (e.g. College Roster).'
      },
      {
        heading: '2. Matching Algorithms',
        content: 'Supports exact key matching, normalized key matching, and fuzzy Levenshtein string distance comparisons.'
      },
      {
        heading: '3. Audit Categorization',
        content: 'Classifies records into Exact Matches, Discrepant Fields, Present only in A, and Present only in B.'
      }
    ]
  },
  'shortener': {
    title: 'URL Shortener & QR Studio - Logic & Guide',
    icon: 'Link',
    description: 'Create short links, custom aliases, high-res QR codes, and batch-shorten entire Excel roster columns.',
    sections: [
      {
        heading: '1. Link Shortening & Aliases',
        content: 'Creates 6-character unique base62 aliases or custom semantic slugs linked to target URLs.'
      },
      {
        heading: '2. Batch Excel Processing',
        content: 'Upload an Excel sheet with a column of long URLs; appends short links and QR image embeds in batch.'
      },
      {
        heading: '3. Vector QR Generator',
        content: 'Generates high-contrast printable QR codes in PNG and SVG formats.'
      }
    ]
  },
  'image-tools': {
    title: 'Image Tools & Signature Studio - Logic & Guide',
    icon: 'Sparkles',
    description: 'Clean signatures from paper, compress, resize, crop, upscale, watermark, blur sensitive data, and convert formats locally.',
    sections: [
      {
        heading: '1. Paper Signature Cleaner',
        content: 'Applies adaptive luminance thresholding to isolate dark ink from paper textures, creating transparent PNG signatures.'
      },
      {
        heading: '2. Batch Resizing & Compression',
        content: 'Resize and compress batches of student photos/signatures to strict university portal limits (e.g. <50KB).'
      },
      {
        heading: '3. Local Privacy',
        content: 'All image processing runs 100% locally in your browser memory—no files are uploaded to any external server.'
      }
    ]
  },
  'pdf-tools': {
    title: 'PDF Tool Studio - Logic & Guide',
    icon: 'FileText',
    description: 'Merge, split, compress, watermark, rotate, stamp signatures, add page numbers, and convert images to PDF.',
    sections: [
      {
        heading: '1. Client-Side PDF Engine',
        content: 'Powered by pdf-lib to manipulate vector structures without cloud dependencies.'
      },
      {
        heading: '2. Operations Supported',
        content: 'Merge multiple PDFs, split by page ranges, stamp signatures with coordinates, watermark draft pages, and paginate.'
      },
      {
        heading: '3. Batch Processing',
        content: 'Process dozens of documents simultaneously with instant ZIP export.'
      }
    ]
  },
  'admission-import': {
    title: 'Admission Import - Extraction Logic & Guide',
    icon: 'Database',
    description: 'Automate ingestion, sequence-agnostic schema normalization, lookup enrichment, and export to 40-column master XLSX.',
    sections: [
      {
        heading: '1. Sequence-Agnostic Column Mapping',
        content: 'Automatically resolves column aliases (e.g. StudentName, FullName, Name -> StudentName) regardless of column order.'
      },
      {
        heading: '2. Code & Category Lookup Enrichment',
        content: 'Cross-references Category, Caste, Religion, College, and Degree against standardized university lookup dictionaries.'
      },
      {
        heading: '3. Standard 40-Column Master XLSX',
        content: 'Generates compliant 40-column import files with autofilter and data type validations.'
      }
    ]
  },
  'course-master-import': {
    title: 'Course Master Import - Extraction Logic & Guide',
    icon: 'BookOpen',
    description: 'Process course syllabus matrices into 37 standardized master columns with optional DSC 1 / 2 duplication and multi-subject exports.',
    sections: [
      {
        heading: '1. 37-Column Course Master & 10-Column Group Master',
        content: 'Parses course syllabus sheets, building the complete Group Hierarchy and 37 standardized Course Master columns.'
      },
      {
        heading: '2. Marks & Evaluation Calculations',
        content: 'AMMaxMarks = (TH Max + PR Max), AMMinMarks = (TH Min + PR Min) for ESE and CE components. ATMinMarks is fixed to 0.'
      },
      {
        heading: '3. Sorting & Multi-Column Filtering',
        content: 'Interactive Ascending/Descending sorting on all 37 headers and column-specific instant search.'
      }
    ]
  },
  'affiliated-programs': {
    title: 'Affiliated Programme Details - Extraction Logic & Guide',
    icon: 'Building2',
    description: 'Split and explode comma-separated course details, extract course codes, names, programme year, and semester term with dual reporting formats.',
    sections: [
      {
        heading: '1. Format 1: Deduplicated Report (7 Columns)',
        content: 'College Code, College Name, Programme Year, Program Term Name, Course Details, Course Code, Course Name. Filters duplicate courses for a programme for a specific college, year, and term.'
      },
      {
        heading: '2. Format 2: All Rows Exploded Report (9 Columns)',
        content: 'College Code, College Name, Program Code, Program Term, Programme Year, Program Term Name, Course Details, Course Code, Course Name. Retains all exploded rows without deduplication.'
      },
      {
        heading: '3. Regex Transformations',
        content: 'Exploder: r\',\\s*(?=\\()\' | Year: r\'(Year\\s+[IVXLCDM]+)\' | Term: r\'(SEMESTER\\s+[IVXLCDM]+)\' | Code: r\'^\\(([^)]+)\\)\''
      },
      {
        heading: '4. Export Options',
        content: 'Export either format independently or export a combined 2-Sheet Excel workbook containing both Unique_Deduplicated_Courses and All_Exploded_Rows sheets.'
      }
    ]
  },
  'ades-supplementary-calculator': {
    title: 'ADES Supplementary & Improvement Calculator Guide',
    description: 'Specialized examination result calculator for combined Supplementary and Improvement sessions with multi-event baseline carry-forward.',
    sections: [
      {
        heading: '1. Dual-Report Architecture',
        content: 'Upload the current session ADES marks Excel along with one or more previous event ADES reports (earlier regular or supplementary sessions).'
      },
      {
        heading: '2. Improvement Repetition & Carry-Forward Rule',
        content: 'For Improvement candidates, only the ESE-TH component is repeated. All other components (CE-TH, CE-PR, and ESE-PR) are strictly carried forward from the student’s previous event record.'
      },
      {
        heading: '3. Supplementary Repetition & Carry-Forward Rule',
        content: 'For Supplementary candidates, ESE-TH and ESE-PR are repeated. Continuous evaluation marks (CE-TH, CE-PR) are strictly carried forward from the earlier event record.'
      },
      {
        heading: '4. Multi-Event Historical Baseline Index',
        content: 'Supports multiple earlier regular or supplementary event files indexed by PRN + Course Code to automatically resolve prior marks.'
      },
      {
        heading: '5. Configurable Improvement Scoring Policy',
        content: 'Choose between "Current Attempt Marks" (default) or "Best of Both" (higher of current and previous ESE-TH marks).'
      }
    ]
  },
  'ades-result-calculator': {
    title: 'ADES Result Calculator - Calculation & Moderation Guide',
    icon: 'Calculator',
    description: 'Calculates ESE (30% ceiling rule), Continuous Evaluation (CE), and aggregate course marks (35% overall rule) with course-level Moderation Engine across 31 standardized columns.',
    sections: [
      {
        heading: '1. Student-Course Level Grouping',
        content: 'Groups raw assessment rows by (Faculty, Program Term Name, Seat Number, PRN, Course Code, Course Name) to aggregate individual component marks.'
      },
      {
        heading: '2. ESE Component & 30% Pass Rule',
        content: 'Sums Practical (PR) and Theory (TH) ESE marks. ESE - Max = (ESE - PR Max + ESE - TH Max). ESE - Min = ceil(0.30 × ESE - Max). ESE Overall = (ESE - PR Obtained + ESE - TH Obtained). ESE Pass = "Pass" if ESE Overall >= ESE - Min else "Fail".'
      },
      {
        heading: '3. Continuous Evaluation (CE) Component',
        content: 'Sums Practical (PR) and Theory (TH) CE marks. CE - Max = (CE - PR Max + CE - TH Max). CE - Min = 0. CE Overall Marks = (CE - PR Obtained + CE - TH Obtained).'
      },
      {
        heading: '4. Aggregate Course & 35% Overall Pass Rule',
        content: 'Overall Maximum = ESE - Max + CE - Max. Overall Minimum = ceil(0.35 × Overall Maximum). Course Overall Marks = ESE Overall + CE Overall Marks. Overall pass = "Pass" if Course Overall Marks >= Overall Minimum else "Fail".'
      },
      {
        heading: '5. Combined Course Pass/Fail Status',
        content: 'Course Pass/Fail status is "Pass" only if BOTH ESE Pass is "Pass" AND Overall pass is "Pass". If either condition fails, final status is "Fail".'
      },
      {
        heading: '6. Course Moderation Engine (Theory Only vs PR-Only Option)',
        content: 'Set course-specific moderation limits M_limit via interactive matrix or by uploading a Moderation Excel. Moderation is effected only on courses with an ESE Theory (TH) component by default. Courses with solely ESE - PR are excluded unless the user enables the "Allow Moderation on Solely ESE - PR Courses" toggle in the Moderation tab. Formula: Marks Needed = max(ESE Deficit, Overall Deficit). If Marks Needed <= M_limit, student is awarded Moderation Marks = Marks Needed, turning ESE Pass, Overall pass, and Course Pass/Fail to "Pass".'
      },
      {
        heading: '7. Master 31-Column Export',
        content: 'Exports complete formatted 31-column calculation sheets to Excel with sheet name "Output file ", including the Moderation Marks column directly next to Course Pass/Fail.'
      },
      {
        heading: '8. Course-Wise Pass Simulation (+0 to +10 Moderation)',
        content: 'Provides an instant simulation matrix and Excel export for each course showing total student appearances, Normal Pass count (0 Moderation), and simulated pass counts if +1, +2, +3, ..., +10 moderation marks are awarded.'
      }
    ]
  },
  'scheduler': {
    title: 'Timetable Scheduler - Logic & Guide',
    icon: 'CalendarDays',
    description: 'A dedicated tool to isolate structural blocks, map execution dates, and generate sorted venue logs.',
    sections: [
      {
        heading: '1. Block Matrix Scheduling',
        content: 'Isolates structural exam blocks and maps dates to morning and afternoon sessions.'
      },
      {
        heading: '2. Conflict & Clash Detection',
        content: 'Highlights overlapping exam subjects and student registration conflicts automatically.'
      },
      {
        heading: '3. Venue Logs Export',
        content: 'Generates chronological master timetable matrices and center-wise distribution spreadsheets.'
      }
    ]
  }
};
