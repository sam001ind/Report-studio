import React from 'react';

const Sidebar = ({ activePage, setActivePage, stats }) => {
  return (
    <aside style={styles.sidebar}>
      <div style={styles.brand}>
        <div style={styles.brandMark}>RS</div>
        <div>
          <h1 style={styles.brandTitle}>Report Studio</h1>
          <p style={styles.brandSub}>React + Python Generator</p>
        </div>
      </div>

      <nav style={styles.navMenu}>
        <button 
          style={activePage === 'config' ? { ...styles.navBtn, ...styles.navBtnActive } : styles.navBtn}
          onClick={() => setActivePage('config')}
        >
          1. Report Config
        </button>
        <button 
          style={activePage === 'template' ? { ...styles.navBtn, ...styles.navBtnActive } : styles.navBtn}
          onClick={() => setActivePage('template')}
        >
          2. Template Creator
        </button>
        <button 
          style={activePage === 'generate' ? { ...styles.navBtn, ...styles.navBtnActive } : styles.navBtn}
          onClick={() => setActivePage('generate')}
        >
          3. Generate & Export
        </button>
      </nav>

      <section style={styles.panel}>
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
      </section>
    </aside>
  );
};

const styles = {
  sidebar: {
    width: '260px',
    backgroundColor: 'var(--panel)',
    borderRight: '1px solid var(--line)',
    display: 'flex',
    flexDirection: 'column',
    padding: '24px',
    zIndex: 10,
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
  navBtn: {
    backgroundColor: 'transparent',
    color: 'var(--ink)',
    textAlign: 'left',
    padding: '12px 16px',
    borderRadius: '8px',
    fontWeight: 500,
    border: 'none',
    cursor: 'pointer',
    transition: 'background 0.2s',
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
