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
  Sparkles
} from 'lucide-react';
import ThemeToggle from '../components/ThemeToggle';

const LandingPage = () => {
  const navigate = useNavigate();

  return (
    <div style={styles.container}>
      {/* Decorative Background Gradients */}
      <div style={{...styles.blob, top: '-10%', left: '-10%', background: 'rgba(23, 107, 135, 0.1)'}} />
      <div style={{...styles.blob, bottom: '-10%', right: '-10%', background: 'rgba(92, 187, 212, 0.1)'}} />

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
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/studio')}
          >
            <div style={styles.featureIcon}><LayoutTemplate size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Report Studio</h3>
            <p style={styles.featureText}>Build custom data-driven templates and generate hundreds of printable reports in seconds.</p>
          </div>

          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/revaluation')}
          >
            <div style={styles.featureIcon}><TableProperties size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Revaluation</h3>
            <p style={styles.featureText}>Merge multiple application reports with result sheets to generate combined datasets and final PDFs.</p>
          </div>

          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/splitter')}
          >
            <div style={styles.featureIcon}><FileDown size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Excel Lot Splitter</h3>
            <p style={styles.featureText}>Split a large Excel worksheet into smaller lots/chunks and download them bundled in a ZIP archive.</p>
          </div>

          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/merger')}
          >
            <div style={styles.featureIcon}><FileStack size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Excel Sheet Merger</h3>
            <p style={styles.featureText}>Combine multiple Excel sheets into a single document with file origin tracking and custom headers.</p>
          </div>

          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/sll-nominal')}
          >
            <div style={styles.featureIcon}><FileText size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Venue-Wise Nominal Roll</h3>
            <p style={styles.featureText}>Generate venue-wise nominal roll sheets and PDFs with merged student registration cells.</p>
          </div>
          
          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/qp-statement')}
          >
            <div style={styles.featureIcon}><CalendarRange size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>QP Statement Report</h3>
            <p style={styles.featureText}>Compile daily printing lists and venue packing slips for examination question papers.</p>
          </div>

          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/qp-label')}
          >
            <div style={styles.featureIcon}><TableProperties size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>QP Label Generator</h3>
            <p style={styles.featureText}>Generate packet covers and Question Paper envelope labels sorted by center and subject.</p>
          </div>

          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/compare')}
          >
            <div style={styles.featureIcon}><GitCompare size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Data Comparison & Reconciliation</h3>
            <p style={styles.featureText}>Fuzzy match & compare Excel datasets, identify partial matches, detect discrepancies, and export audit reports.</p>
          </div>

          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/shortener')}
          >
            <div style={styles.featureIcon}><LinkIcon size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>URL Shortener & QR Studio</h3>
            <p style={styles.featureText}>Create short links, custom aliases, high-res QR codes, and batch-shorten entire Excel roster columns.</p>
          </div>

          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), transparent)'}} 
            onClick={() => navigate('/image-tools')}
          >
            <div style={styles.featureIcon}><Sparkles size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Image Tools & Signature Studio</h3>
            <p style={styles.featureText}>Clean signatures from paper, compress, resize, crop, upscale, watermark, blur sensitive data, and convert formats locally.</p>
          </div>

          <div 
            style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
            onClick={() => navigate('/scheduler')}
          >
            <div style={styles.featureIcon}><CalendarDays size={24} color="var(--accent)" /></div>
            <h3 style={styles.featureTitle}>Timetable Scheduler</h3>
            <p style={styles.featureText}>A dedicated tool to isolate structural blocks, map execution dates, and generate sorted venue logs.</p>
          </div>
        </section>
      </main>
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
    flexDirection: 'column'
  },
  blob: {
    position: 'absolute',
    width: '60vw',
    height: '60vw',
    borderRadius: '50%',
    filter: 'blur(100px)',
    zIndex: 0,
    pointerEvents: 'none'
  },
  navbar: {
    position: 'relative',
    zIndex: 10,
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '16px 32px',
    borderBottom: '1px solid var(--line)',
    background: 'rgba(255, 255, 255, 0.7)',
    backdropFilter: 'blur(10px)'
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
    transition: 'all 0.2s ease',
    boxShadow: '0 2px 8px rgba(0,0,0,0.03)',
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
  }
};

export default LandingPage;
