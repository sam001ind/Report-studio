import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { 
  FileDown, 
  FileStack, 
  LayoutTemplate, 
  TableProperties, 
  CalendarDays, 
  FileText, 
  CalendarRange,
  GitCompare,
  Link as LinkIcon,
  Sparkles,
  Database,
  BookOpen,
  Building2,
  Calculator,
  HelpCircle,
  X,
  ArrowRight
} from 'lucide-react';
import ThemeToggle from '../components/ThemeToggle';
import { TOOL_GUIDES } from '../data/toolGuides';

const LandingPage = () => {
  const navigate = useNavigate();
  const [activeGuideKey, setActiveGuideKey] = useState(null);

  const selectedGuide = activeGuideKey ? TOOL_GUIDES[activeGuideKey] : null;

  return (
    <div style={styles.container}>
      {/* High-Performance Decorative Background Gradients (No GPU Blur Repaint) */}
      <div style={{
        position: 'absolute',
        top: '-15%',
        left: '-10%',
        width: '55vw',
        height: '55vw',
        borderRadius: '50%',
        background: 'radial-gradient(circle, rgba(23, 107, 135, 0.08) 0%, transparent 70%)',
        pointerEvents: 'none',
        zIndex: 0,
        transform: 'translateZ(0)'
      }} />
      <div style={{
        position: 'absolute',
        bottom: '-15%',
        right: '-10%',
        width: '55vw',
        height: '55vw',
        borderRadius: '50%',
        background: 'radial-gradient(circle, rgba(92, 187, 212, 0.08) 0%, transparent 70%)',
        pointerEvents: 'none',
        zIndex: 0,
        transform: 'translateZ(0)'
      }} />

      <nav style={styles.navbar}>
        <div style={styles.logo}>
          <div style={styles.logoMark}>RS</div>
          <span style={styles.logoText}>Report Studio</span>
        </div>
        <div style={{ display: 'flex', gap: '15px', alignItems: 'center' }}>
          <ThemeToggle />
        </div>
      </nav>

      <main style={styles.main}>
        <section style={styles.features}>
          
          {/* 1. Report Studio */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/studio')}
          >
            <div style={styles.featureIcon}><LayoutTemplate size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Report Studio</h3>
            <p style={styles.featureText}>Build custom data-driven templates and generate hundreds of printable reports in seconds.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('studio'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 2. Revaluation */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/revaluation')}
          >
            <div style={styles.featureIcon}><TableProperties size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Revaluation</h3>
            <p style={styles.featureText}>Merge multiple application reports with result sheets to generate combined datasets and final PDFs.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('revaluation'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 3. Excel Lot Splitter */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/splitter')}
          >
            <div style={styles.featureIcon}><FileDown size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Excel Lot Splitter</h3>
            <p style={styles.featureText}>Split a large Excel worksheet into smaller lots/chunks and download them bundled in a ZIP archive.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('splitter'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 4. Excel Sheet Merger */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/merger')}
          >
            <div style={styles.featureIcon}><FileStack size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Excel Sheet Merger</h3>
            <p style={styles.featureText}>Combine multiple Excel sheets into a single document with file origin tracking and custom headers.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('merger'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 5. Venue-Wise Nominal Roll */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/sll-nominal')}
          >
            <div style={styles.featureIcon}><FileText size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Venue-Wise Nominal Roll</h3>
            <p style={styles.featureText}>Generate venue-wise nominal roll sheets and PDFs with merged student registration cells.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('sll-nominal'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>
          
          {/* 6. QP Statement Report */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/qp-statement')}
          >
            <div style={styles.featureIcon}><CalendarRange size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>QP Statement Report</h3>
            <p style={styles.featureText}>Compile daily printing lists and venue packing slips for examination question papers.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('qp-statement'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 7. QP Label Generator */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/qp-label')}
          >
            <div style={styles.featureIcon}><TableProperties size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>QP Label Generator</h3>
            <p style={styles.featureText}>Generate packet covers and Question Paper envelope labels sorted by center and subject.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('qp-label'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 8. Data Comparison & Reconciliation */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/compare')}
          >
            <div style={styles.featureIcon}><GitCompare size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Data Comparison & Reconciliation</h3>
            <p style={styles.featureText}>Fuzzy match & compare Excel datasets, identify partial matches, detect discrepancies, and export audit reports.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('compare'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 9. URL Shortener & QR Studio */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/shortener')}
          >
            <div style={styles.featureIcon}><LinkIcon size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>URL Shortener & QR Studio</h3>
            <p style={styles.featureText}>Create short links, custom aliases, high-res QR codes, and batch-shorten entire Excel roster columns.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('shortener'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 10. Image Tools & Signature Studio */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/image-tools')}
          >
            <div style={styles.featureIcon}><Sparkles size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Image Tools & Signature Studio</h3>
            <p style={styles.featureText}>Clean signatures from paper, compress, resize, crop, upscale, watermark, blur sensitive data, and convert formats locally.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('image-tools'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 11. PDF Tool Studio */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/pdf-tools')}
          >
            <div style={styles.featureIcon}><FileText size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>PDF Tool Studio</h3>
            <p style={styles.featureText}>Merge, split, compress, watermark, rotate, stamp signatures, add page numbers, and convert images to PDF.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('pdf-tools'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 12. Admission Import */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/admission-import')}
          >
            <div style={styles.featureIcon}><Database size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Admission Import</h3>
            <p style={styles.featureText}>Automate ingestion, sequence-agnostic schema normalization, lookup enrichment, and export to 40-column master XLSX.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('admission-import'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 13. Course Master Import */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/course-master-import')}
          >
            <div style={styles.featureIcon}><BookOpen size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Course Master Import</h3>
            <p style={styles.featureText}>Process course syllabus matrices into 37 standardized master columns with optional DSC 1 / 2 duplication and multi-subject exports.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('course-master-import'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 14. Affiliated programme details */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/affiliated-programs')}
          >
            <div style={styles.featureIcon}><Building2 size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Affiliated programme details</h3>
            <p style={styles.featureText}>Split and explode comma-separated course details, extract course codes, names, programme year, and semester term.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('affiliated-programs'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 15. ADES Result calculator */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/ades-result-calculator')}
          >
            <div style={styles.featureIcon}><Calculator size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>ADES Result Calculator (Regular)</h3>
            <p style={styles.featureText}>Aggregate regular student-course assessments with ESE (30% rule), CE, and 35% aggregate pass logic across 30 master columns.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('ades-result-calculator'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 15b. ADES Supplementary & Improvement Calculator */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid #6366f1', background: 'linear-gradient(135deg, rgba(99,102,241,0.08), transparent)'}} 
            onClick={() => navigate('/ades-supplementary-calculator')}
          >
            <div style={styles.featureIcon}><Calculator size={24} color="#6366f1" /></div>
            <h3 style={styles.featureTitle}>ADES Supplementary / Improvement</h3>
            <p style={styles.featureText}>Aggregate supplementary and improvement exams with multi-event historical carry-forward for CE and ESE components.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('ades-supplementary-calculator'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>

          {/* 16. Timetable Scheduler */}
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/scheduler')}
          >
            <div style={styles.featureIcon}><CalendarDays size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Timetable Scheduler</h3>
            <p style={styles.featureText}>A dedicated tool to isolate structural blocks, map execution dates, and generate sorted venue logs.</p>
            <div style={styles.cardFooter}>
              <button 
                type="button" 
                style={styles.guideBtn} 
                onClick={(e) => { e.stopPropagation(); setActiveGuideKey('scheduler'); }}
              >
                <HelpCircle size={13} /> Extraction Logic & Guide
              </button>
            </div>
          </div>
        </section>
      </main>

      {/* Global Extraction Logic & Guide Modal */}
      {selectedGuide && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0,0,0,0.6)',
          backdropFilter: 'blur(3px)',
          zIndex: 2000,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          padding: '20px'
        }} onClick={() => setActiveGuideKey(null)}>
          <div style={{
            background: 'var(--panel)',
            border: '1px solid var(--line)',
            borderRadius: '12px',
            width: '640px',
            maxWidth: '92vw',
            maxHeight: '85vh',
            display: 'flex',
            flexDirection: 'column',
            boxShadow: '0 20px 40px rgba(0,0,0,0.3)',
            overflow: 'hidden'
          }} onClick={(e) => e.stopPropagation()}>
            <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'var(--bg)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <div style={{ padding: '6px', background: 'var(--accent-soft)', borderRadius: '6px', color: 'var(--accent)' }}>
                  <HelpCircle size={18} />
                </div>
                <div>
                  <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--ink)' }}>{selectedGuide.title}</h3>
                  <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>{selectedGuide.description}</span>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setActiveGuideKey(null)}
                style={{ background: 'none', border: 'none', color: 'var(--muted)', cursor: 'pointer', padding: '4px' }}
              >
                <X size={18} />
              </button>
            </div>

            <div style={{ padding: '20px', overflowY: 'auto', fontSize: '12.5px', lineHeight: '1.6', display: 'flex', flexDirection: 'column', gap: '16px' }}>
              {(selectedGuide.sections || []).map((sec, idx) => (
                <div key={idx} style={{ background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px', padding: '12px 14px' }}>
                  <strong style={{ fontSize: '12.5px', color: 'var(--accent)', display: 'block', marginBottom: '4px' }}>
                    {sec.heading}
                  </strong>
                  <div style={{ color: 'var(--ink)' }}>{sec.content}</div>
                </div>
              ))}
            </div>

            <div style={{ padding: '12px 20px', borderTop: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'var(--bg)' }}>
              <button 
                type="button" 
                className="secondary"
                onClick={() => setActiveGuideKey(null)}
                style={{ padding: '6px 14px', fontSize: '12px' }}
              >
                Close
              </button>
              <button 
                type="button" 
                onClick={() => {
                  const key = activeGuideKey;
                  setActiveGuideKey(null);
                  if (key === 'studio') navigate('/studio');
                  else if (key === 'revaluation') navigate('/revaluation');
                  else if (key === 'splitter') navigate('/splitter');
                  else if (key === 'merger') navigate('/merger');
                  else if (key === 'sll-nominal') navigate('/sll-nominal');
                  else if (key === 'qp-statement') navigate('/qp-statement');
                  else if (key === 'qp-label') navigate('/qp-label');
                  else if (key === 'compare') navigate('/compare');
                  else if (key === 'shortener') navigate('/shortener');
                  else if (key === 'image-tools') navigate('/image-tools');
                  else if (key === 'pdf-tools') navigate('/pdf-tools');
                  else if (key === 'admission-import') navigate('/admission-import');
                  else if (key === 'course-master-import') navigate('/course-master-import');
                  else if (key === 'affiliated-programs') navigate('/affiliated-programs');
                  else if (key === 'ades-result-calculator') navigate('/ades-result-calculator');
                  else if (key === 'ades-supplementary-calculator') navigate('/ades-supplementary-calculator');
                  else if (key === 'scheduler') navigate('/scheduler');
                }}
                style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 16px', background: 'var(--accent)', color: 'white', borderRadius: '6px', border: 'none', fontWeight: 600, fontSize: '12px', cursor: 'pointer' }}
              >
                Open Tool <ArrowRight size={13} />
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};

const styles = {
  container: {
    minHeight: '100vh',
    width: '100vw',
    backgroundColor: 'var(--bg)',
    fontFamily: 'var(--font-family)',
    position: 'relative',
    overflowX: 'hidden',
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
    WebkitFontSmoothing: 'antialiased',
    MozOsxFontSmoothing: 'grayscale',
    textRendering: 'optimizeLegibility'
  },
  blob: {
    position: 'absolute',
    width: '55vw',
    height: '55vw',
    borderRadius: '50%',
    zIndex: 0,
    pointerEvents: 'none',
    transform: 'translateZ(0)'
  },
  navbar: {
    position: 'relative',
    zIndex: 10,
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '16px 32px',
    borderBottom: '1px solid var(--line)',
    background: 'var(--panel)',
    backdropFilter: 'blur(8px)',
    contain: 'layout style'
  },
  logo: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px'
  },
  logoMark: {
    backgroundColor: 'var(--accent)',
    color: 'white',
    width: '32px',
    height: '32px',
    borderRadius: '6px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontWeight: 800,
    fontSize: '14px',
  },
  logoText: {
    fontSize: '18px',
    fontWeight: 700,
    color: 'var(--ink)'
  },
  main: {
    position: 'relative',
    zIndex: 10,
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    padding: '24px 32px'
  },
  features: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
    gap: '16px',
    maxWidth: '1300px',
    width: '100%'
  },
  featureCard: {
    backgroundColor: 'var(--panel)',
    border: '1px solid var(--line)',
    borderRadius: '12px',
    padding: '20px',
    transition: 'transform 0.16s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.16s cubic-bezier(0.16, 1, 0.3, 1), border-color 0.16s ease',
    boxShadow: '0 2px 8px rgba(0,0,0,0.03)',
    willChange: 'transform',
    contain: 'paint layout'
  },
  featureIcon: {
    backgroundColor: 'var(--accent-soft)',
    width: '40px',
    height: '40px',
    borderRadius: '10px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    marginBottom: '14px'
  },
  featureTitle: {
    fontSize: '17px',
    fontWeight: 700,
    margin: '0 0 6px 0',
    color: 'var(--ink)'
  },
  featureText: {
    fontSize: '13px',
    color: 'var(--muted)',
    lineHeight: 1.5,
    margin: 0
  },
  cardFooter: {
    marginTop: '14px',
    paddingTop: '10px',
    borderTop: '1px solid var(--line)',
    display: 'flex',
    justifyContent: 'flex-start'
  },
  guideBtn: {
    background: 'var(--bg)',
    border: '1px solid var(--line)',
    color: 'var(--accent)',
    fontSize: '11.5px',
    fontWeight: 600,
    padding: '4px 10px',
    borderRadius: '6px',
    display: 'inline-flex',
    alignItems: 'center',
    gap: '6px',
    cursor: 'pointer',
    transition: 'all 0.15s ease'
  }
};

export default LandingPage;
