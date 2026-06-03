import React, { useState } from 'react';
import { Link } from 'react-router-dom';

const Sidebar = ({ activePage, setActivePage, stats }) => {
  const [isOpen, setIsOpen] = useState(true);

  return (
    <aside style={{ ...styles.sidebar, width: isOpen ? '260px' : '80px', padding: isOpen ? '24px' : '24px 8px' }}>
      <div style={{ display: 'flex', justifyContent: isOpen ? 'space-between' : 'center', alignItems: 'center', marginBottom: '24px' }}>
        {isOpen && (
          <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '8px', fontSize: '14px' }}>
            ← Back to Portal
          </Link>
        )}
        <button 
          onClick={() => setIsOpen(!isOpen)} 
          style={{ background: 'transparent', border: 'none', color: 'var(--muted)', cursor: 'pointer', fontSize: '16px', padding: '4px' }}
          title="Toggle Sidebar"
        >
          {isOpen ? '◀' : '▶'}
        </button>
      </div>

      <div style={{ ...styles.brand, justifyContent: isOpen ? 'flex-start' : 'center' }}>
        <div style={styles.brandMark}>RS</div>
        {isOpen && (
          <div>
            <h1 style={styles.brandTitle}>Report Studio</h1>
            <p style={styles.brandSub}>React + Python Generator</p>
          </div>
        )}
      </div>

      <nav style={styles.navMenu}>
        <button 
          onClick={() => setActivePage('config')}
          style={{ 
            ...styles.navBtnBase, 
            ...(activePage === 'config' ? styles.navBtnActive : {}),
            justifyContent: isOpen ? 'flex-start' : 'center', 
            padding: isOpen ? '12px 16px' : '12px 0' 
          }}
        >
          <span className="icon" style={{ marginRight: isOpen ? '12px' : 0 }}>📄</span> 
          {isOpen && <span>Report Config</span>}
        </button>
        <button 
          onClick={() => setActivePage('template')}
          style={{ 
            ...styles.navBtnBase, 
            ...(activePage === 'template' ? styles.navBtnActive : {}),
            justifyContent: isOpen ? 'flex-start' : 'center', 
            padding: isOpen ? '12px 16px' : '12px 0' 
          }}
        >
          <span className="icon" style={{ marginRight: isOpen ? '12px' : 0 }}>🎨</span> 
          {isOpen && <span>Template Creator</span>}
        </button>
        <button 
          onClick={() => setActivePage('generate')}
          style={{ 
            ...styles.navBtnBase, 
            ...(activePage === 'generate' ? styles.navBtnActive : {}),
            justifyContent: isOpen ? 'flex-start' : 'center', 
            padding: isOpen ? '12px 16px' : '12px 0' 
          }}
        >
          <span className="icon" style={{ marginRight: isOpen ? '12px' : 0 }}>⚙️</span> 
          {isOpen && <span>Generate Output</span>}
        </button>
      </nav>

      <section style={{ ...styles.panel, padding: isOpen ? '16px' : '16px 8px' }}>
        {isOpen ? (
          <>
            <div style={styles.panelTitle}>Active Dataset</div>
            <div style={styles.miniStats}>
              <div style={styles.statCol}>
                <strong style={styles.statStr}>{stats.rows}</strong>
                <span style={styles.statSpan}>Rows</span>
              </div>
              <div style={styles.statCol}>
                <strong style={styles.statStr}>{stats.cols}</strong>
                <span style={styles.statSpan}>Columns</span>
              </div>
            </div>
          </>
        ) : (
          <div style={{ textAlign: 'center' }}>
            <strong style={{ fontSize: '14px', color: 'var(--accent)' }}>{stats.rows}</strong>
            <div style={{ fontSize: '10px', color: 'var(--muted)' }}>R</div>
          </div>
        )}
      </section>
    </aside>
  );
};

const styles = {
  sidebar: {
    backgroundColor: 'var(--panel)',
    borderRight: '1px solid var(--line)',
    display: 'flex',
    flexDirection: 'column',
    zIndex: 10,
    transition: 'width 0.3s ease, padding 0.3s ease',
  },
  brand: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    marginBottom: '40px',
  },
  brandMark: {
    backgroundColor: 'var(--accent)',
    color: 'white',
    width: '40px',
    height: '40px',
    borderRadius: '8px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontWeight: 800,
    fontSize: '18px',
  },
  brandTitle: { margin: 0, fontSize: '18px' },
  brandSub: { margin: '4px 0 0', fontSize: '12px', color: 'var(--muted)' },
  navMenu: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    marginBottom: '40px',
  },
  navBtnBase: {
    backgroundColor: 'transparent',
    color: 'var(--ink)',
    borderRadius: '8px',
    fontWeight: 500,
    border: 'none',
    cursor: 'pointer',
    transition: 'background 0.2s',
    display: 'flex',
    alignItems: 'center',
  },
  navBtnActive: {
    backgroundColor: 'var(--accent-soft)',
    color: 'var(--accent)',
    fontWeight: 600,
  },
  panel: {
    backgroundColor: 'var(--bg)',
    borderRadius: '8px',
    padding: '16px',
    marginTop: 'auto',
  },
  panelTitle: {
    fontSize: '12px',
    fontWeight: 600,
    color: 'var(--muted)',
    marginBottom: '12px',
    textTransform: 'uppercase',
  },
  miniStats: {
    display: 'flex',
    gap: '16px',
  },
  statCol: {
    display: 'flex',
    flexDirection: 'column',
  },
  statStr: { fontSize: '24px', color: 'var(--accent)' },
  statSpan: { fontSize: '12px', color: 'var(--muted)' }
};

export default Sidebar;
