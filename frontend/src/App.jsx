import { useEffect, lazy, Suspense } from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import LandingPage from './pages/LandingPage';
import ErrorBoundary from './components/ErrorBoundary';
import { AuthProvider, useAuth } from './context/AuthContext';
import { ThemeProvider } from './context/ThemeContext';
import { startSupabaseKeepAlive } from './utils/supabaseKeepAlive';

// Code-split heavy tool pages on demand to keep initial load lightweight and buttery smooth
const StudioLayout = lazy(() => import('./pages/StudioLayout'));
const SchedulerPage = lazy(() => import('./pages/SchedulerPage'));
const RevaluationPage = lazy(() => import('./pages/RevaluationPage'));
const SplitterPage = lazy(() => import('./pages/SplitterPage'));
const MergerPage = lazy(() => import('./pages/MergerPage'));
const SllNominalPage = lazy(() => import('./pages/SllNominalPage'));
const QpStatementPage = lazy(() => import('./pages/QpStatementPage'));
const QpLabelPage = lazy(() => import('./pages/QpLabelPage'));
const DataComparisonPage = lazy(() => import('./pages/DataComparisonPage'));
const UrlShortenerPage = lazy(() => import('./pages/UrlShortenerPage'));
const ImageToolsPage = lazy(() => import('./pages/ImageToolsPage'));
const PdfToolsPage = lazy(() => import('./pages/PdfToolsPage'));
const AdmissionImportPage = lazy(() => import('./pages/AdmissionImportPage'));
const CourseMasterImportPage = lazy(() => import('./pages/CourseMasterImportPage'));
const AffiliatedProgrammePage = lazy(() => import('./pages/AffiliatedProgrammePage'));
const AdesResultCalculatorPage = lazy(() => import('./pages/AdesResultCalculatorPage'));
const AdesSupplementaryCalculatorPage = lazy(() => import('./pages/AdesSupplementaryCalculatorPage'));
const UrlRedirectHandler = lazy(() => import('./pages/UrlRedirectHandler'));
const AuthPage = lazy(() => import('./pages/AuthPage'));
const AdminDashboard = lazy(() => import('./pages/AdminDashboard'));

const LoadingFallback = () => (
  <div style={{
    minHeight: '100vh',
    width: '100vw',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    background: 'var(--bg)',
    color: 'var(--muted)',
    fontFamily: 'var(--font-family)',
    gap: '12px'
  }}>
    <div style={{
      width: '28px',
      height: '28px',
      border: '2.5px solid var(--line)',
      borderTopColor: 'var(--accent)',
      borderRadius: '50%',
      animation: 'rsSpin 0.7s linear infinite'
    }} />
    <style>{`
      @keyframes rsSpin {
        to { transform: rotate(360deg); }
      }
    `}</style>
    <span style={{ fontSize: '12px', fontWeight: 600 }}>Loading Module...</span>
  </div>
);

const ProtectedRoute = ({ children }) => {
  const { user } = useAuth();
  if (!user) return <Navigate to="/auth" />;
  return children;
};

function App() {
  useEffect(() => {
    const isElectron = !!(window.electronAPI?.isDesktop || /Electron/i.test(navigator.userAgent));
    const isMac = (window.electronAPI?.platform === 'darwin') || /Mac/i.test(navigator.userAgent) || /Mac/i.test(navigator.platform);
    if (isElectron && isMac) {
      document.body.classList.add('is-electron-mac');
      document.documentElement.classList.add('is-electron-mac');
    }
    const stopKeepAlive = startSupabaseKeepAlive(5 * 60 * 1000); // Heartbeat ping every 5 minutes
    return () => stopKeepAlive();
  }, []);

  return (
    <ErrorBoundary>
      <ThemeProvider>
      <AuthProvider>
        <Router>
          <Suspense fallback={<LoadingFallback />}>
            <Routes>
              <Route path="/" element={<LandingPage />} />
              <Route path="/auth" element={<AuthPage />} />
              <Route path="/admin" element={<AdminDashboard />} />
              <Route path="/studio" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <StudioLayout />
                </ProtectedRoute>
              } />
              <Route path="/revaluation" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <RevaluationPage />
                </ProtectedRoute>
              } />
              <Route path="/splitter" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <SplitterPage />
                </ProtectedRoute>
              } />
              <Route path="/excel-splitter" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <SplitterPage />
                </ProtectedRoute>
              } />
              <Route path="/merger" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <MergerPage />
                </ProtectedRoute>
              } />
              <Route path="/excel-merger" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <MergerPage />
                </ProtectedRoute>
              } />
              <Route path="/sll-nominal" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <SllNominalPage />
                </ProtectedRoute>
              } />
              <Route path="/qp-statement" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <QpStatementPage />
                </ProtectedRoute>
              } />
              <Route path="/qp-label" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <QpLabelPage />
                </ProtectedRoute>
              } />
              <Route path="/compare" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <DataComparisonPage />
                </ProtectedRoute>
              } />
              <Route path="/shortener" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <UrlShortenerPage />
                </ProtectedRoute>
              } />
              <Route path="/image-tools" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <ImageToolsPage />
                </ProtectedRoute>
              } />
              <Route path="/images" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <ImageToolsPage />
                </ProtectedRoute>
              } />
              <Route path="/pdf-tools" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <PdfToolsPage />
                </ProtectedRoute>
              } />
              <Route path="/pdf" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <PdfToolsPage />
                </ProtectedRoute>
              } />
              <Route path="/admission-import" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AdmissionImportPage />
                </ProtectedRoute>
              } />
              <Route path="/admission" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AdmissionImportPage />
                </ProtectedRoute>
              } />
              <Route path="/course-master-import" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <CourseMasterImportPage />
                </ProtectedRoute>
              } />
              <Route path="/course-master" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <CourseMasterImportPage />
                </ProtectedRoute>
              } />
              <Route path="/affiliated-programs" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AffiliatedProgrammePage />
                </ProtectedRoute>
              } />
              <Route path="/affiliated-programmes" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AffiliatedProgrammePage />
                </ProtectedRoute>
              } />
              <Route path="/affiliated" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AffiliatedProgrammePage />
                </ProtectedRoute>
              } />
              <Route path="/ades-result-calculator" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AdesResultCalculatorPage />
                </ProtectedRoute>
              } />
              <Route path="/ades-calculator" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AdesResultCalculatorPage />
                </ProtectedRoute>
              } />
              <Route path="/ades" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AdesResultCalculatorPage />
                </ProtectedRoute>
              } />
              <Route path="/ades-supplementary-calculator" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AdesSupplementaryCalculatorPage />
                </ProtectedRoute>
              } />
              <Route path="/ades-supplementary" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AdesSupplementaryCalculatorPage />
                </ProtectedRoute>
              } />
              <Route path="/supplementary-calculator" element={
                <ProtectedRoute requiredPermission="can_access_studio">
                  <AdesSupplementaryCalculatorPage />
                </ProtectedRoute>
              } />
              <Route path="/s/:code" element={<UrlRedirectHandler />} />
              <Route path="/scheduler" element={
                <ProtectedRoute requiredPermission="can_access_scheduler">
                  <SchedulerPage />
                </ProtectedRoute>
              } />
              <Route path="*" element={<Navigate to="/" replace />} />
            </Routes>
          </Suspense>
        </Router>
        </AuthProvider>
      </ThemeProvider>
    </ErrorBoundary>
  );
}

export default App;
