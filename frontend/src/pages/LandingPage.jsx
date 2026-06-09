import React from 'react';
import { useNavigate } from 'react-router-dom';
import { FileDown, LayoutTemplate, Zap, ArrowRight, TableProperties, CalendarDays, LogIn, LogOut, Shield } from 'lucide-react';
import { useAuth } from '../context/AuthContext';
import ThemeToggle from '../components/ThemeToggle';

const LandingPage = () => {
  const navigate = useNavigate();
  const { user, profile, permissions, signOut } = useAuth();
  
  // Admins see everything. Normal users see what they have permission for.
  const canSeeStudio = profile?.is_admin || permissions?.can_access_studio;
  const canSeeScheduler = profile?.is_admin || permissions?.can_access_scheduler;

  return (
    <div style={styles.container}>
      {/* Decorative Background Gradients */}
      <div style={{...styles.blob, top: '-10%', left: '-10%', background: 'rgba(23, 107, 135, 0.1)'}} />
      <div style={{...styles.blob, bottom: '-10%', right: '-10%', background: 'rgba(92, 187, 212, 0.1)'}} />

      <nav style={styles.navbar}>
        <div style={styles.logo}>
          <div style={styles.logoMark}>SOS</div>
          <span style={styles.logoText}>Sorted Operational Suite</span>
        </div>
        <div style={{ display: 'flex', gap: '15px', alignItems: 'center' }}>
          <ThemeToggle />
          
          {profile?.is_admin && (
            <button style={{...styles.secondaryBtn, padding: '8px 16px'}} onClick={() => navigate('/admin')}>
              <Shield size={18} style={{marginRight: '8px'}} /> Admin
            </button>
          )}
          {user ? (
            <button style={{...styles.secondaryBtn, padding: '8px 16px'}} onClick={() => signOut()}>
              <LogOut size={18} style={{marginRight: '8px'}} /> Sign Out
            </button>
          ) : (
            <button style={{...styles.primaryBtn, padding: '8px 16px'}} onClick={() => navigate('/auth')}>
              <LogIn size={18} style={{marginRight: '8px'}} /> Sign In
            </button>
          )}
        </div>
      </nav>

      <main style={styles.main}>
        <section style={styles.hero}>
          <div style={styles.heroBadge}>✨ Unified Toolkit</div>
          <h1 style={styles.title}>
            Operational Efficiency,<br/>
            <span style={styles.gradientText}>Sorted.</span>
          </h1>
          <p style={styles.subtitle}>
            Access all available modules and services below to manage reports, schedule timetables, and automate your workflow.
          </p>
        </section>

        <section style={styles.features}>
          {canSeeStudio && (
            <div 
              style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
              onClick={() => navigate('/studio')}
            >
              <div style={styles.featureIcon}><LayoutTemplate size={24} color="var(--accent)" /></div>
              <h3 style={styles.featureTitle}>Report Studio</h3>
              <p style={styles.featureText}>Build custom data-driven templates and generate hundreds of printable reports in seconds.</p>
            </div>
          )}

          {canSeeStudio && (
            <div 
              style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
              onClick={() => navigate('/studio', { state: { activePage: 'revaluation' } })}
            >
              <div style={styles.featureIcon}><TableProperties size={24} color="var(--accent)" /></div>
              <h3 style={styles.featureTitle}>Revaluation</h3>
              <p style={styles.featureText}>Merge multiple application reports with result sheets to generate combined datasets and final PDFs.</p>
            </div>
          )}
          
          {canSeeScheduler && (
            <div 
              style={{...styles.featureCard, cursor: 'pointer', border: '1px solid var(--accent)'}} 
              onClick={() => navigate('/scheduler')}
            >
              <div style={styles.featureIcon}><CalendarDays size={24} color="var(--accent)" /></div>
              <h3 style={styles.featureTitle}>Timetable Scheduler</h3>
              <p style={styles.featureText}>A dedicated tool to isolate structural blocks, map execution dates, and generate sorted venue logs.</p>
            </div>
          )}

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
    padding: '24px 48px',
    borderBottom: '1px solid rgba(255,255,255,0.5)',
    backdropFilter: 'blur(10px)'
  },
  logo: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px'
  },
  logoMark: {
    backgroundColor: 'var(--accent)',
    color: 'white',
    width: '36px',
    height: '36px',
    borderRadius: '8px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontWeight: 800,
    fontSize: '16px',
  },
  logoText: {
    fontSize: '20px',
    fontWeight: 700,
    color: 'var(--ink)'
  },
  primaryBtn: {
    backgroundColor: 'var(--accent)',
    color: 'white',
    border: 'none',
    padding: '10px 20px',
    borderRadius: '8px',
    fontWeight: 600,
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    transition: 'opacity 0.2s, transform 0.2s',
    boxShadow: '0 4px 12px rgba(23, 107, 135, 0.2)'
  },
  secondaryBtn: {
    backgroundColor: 'transparent',
    color: 'var(--ink)',
    border: '1px solid var(--line)',
    padding: '15px 32px',
    borderRadius: '8px',
    fontWeight: 600,
    cursor: 'pointer',
    textDecoration: 'none',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'background 0.2s',
  },
  main: {
    position: 'relative',
    zIndex: 10,
    flex: 1,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    padding: '60px 24px'
  },
  hero: {
    textAlign: 'center',
    maxWidth: '800px',
    marginBottom: '80px',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center'
  },
  heroBadge: {
    backgroundColor: 'var(--accent-soft)',
    color: 'var(--accent)',
    padding: '6px 16px',
    borderRadius: '20px',
    fontSize: '14px',
    fontWeight: 600,
    marginBottom: '24px'
  },
  title: {
    fontSize: '64px',
    fontWeight: 800,
    lineHeight: 1.1,
    color: 'var(--ink)',
    margin: '0 0 24px 0',
    letterSpacing: '-1px'
  },
  gradientText: {
    background: 'linear-gradient(90deg, var(--accent) 0%, #2f8e9d 100%)',
    WebkitBackgroundClip: 'text',
    WebkitTextFillColor: 'transparent',
  },
  subtitle: {
    fontSize: '20px',
    color: 'var(--muted)',
    lineHeight: 1.6,
    margin: '0 0 40px 0',
    maxWidth: '600px'
  },
  ctaGroup: {
    display: 'flex',
    gap: '16px',
    alignItems: 'center'
  },
  features: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))',
    gap: '32px',
    maxWidth: '1200px',
    width: '100%'
  },
  featureCard: {
    backgroundColor: 'rgba(255, 255, 255, 0.05)',
    backdropFilter: 'blur(10px)',
    border: '1px solid var(--accent)',
    borderRadius: '16px',
    padding: '32px',
    transition: 'all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1)',
    boxShadow: 'var(--shadow)',
    ':hover': {
      transform: 'translateY(-5px)',
      boxShadow: '0 0 30px rgba(0, 240, 255, 0.4)'
    }
  },
  featureIcon: {
    backgroundColor: 'var(--accent-soft)',
    width: '48px',
    height: '48px',
    borderRadius: '12px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    marginBottom: '20px'
  },
  featureTitle: {
    fontSize: '20px',
    fontWeight: 700,
    margin: '0 0 12px 0',
    color: 'var(--ink)'
  },
  featureText: {
    fontSize: '15px',
    color: 'var(--muted)',
    lineHeight: 1.6,
    margin: 0
  }
};

export default LandingPage;
