import { useEffect } from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import LandingPage from './pages/LandingPage';
import StudioLayout from './pages/StudioLayout';
import SchedulerPage from './pages/SchedulerPage';
import RevaluationPage from './pages/RevaluationPage';
import SplitterPage from './pages/SplitterPage';
import MergerPage from './pages/MergerPage';
import SllNominalPage from './pages/SllNominalPage';
import QpStatementPage from './pages/QpStatementPage';
import QpLabelPage from './pages/QpLabelPage';
import DataComparisonPage from './pages/DataComparisonPage';
import UrlShortenerPage from './pages/UrlShortenerPage';
import ImageToolsPage from './pages/ImageToolsPage';
import PdfToolsPage from './pages/PdfToolsPage';
import UrlRedirectHandler from './pages/UrlRedirectHandler';
import AuthPage from './pages/AuthPage';
import AdminDashboard from './pages/AdminDashboard';
import ErrorBoundary from './components/ErrorBoundary';
import { AuthProvider, useAuth } from './context/AuthContext';
import { ThemeProvider } from './context/ThemeContext';
import { startSupabaseKeepAlive } from './utils/supabaseKeepAlive';

const ProtectedRoute = ({ children }) => {
  const { user } = useAuth();
  if (!user) return <Navigate to="/auth" />;
  return children;
};

function App() {
  useEffect(() => {
    const stopKeepAlive = startSupabaseKeepAlive(5 * 60 * 1000); // Heartbeat ping every 5 minutes
    return () => stopKeepAlive();
  }, []);

  return (
    <ErrorBoundary>
      <ThemeProvider>
      <AuthProvider>
        <Router>
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
          <Route path="/s/:code" element={<UrlRedirectHandler />} />
          <Route path="/scheduler" element={
            <ProtectedRoute requiredPermission="can_access_scheduler">
              <SchedulerPage />
            </ProtectedRoute>
          } />
          <Route path="*" element={<Navigate to="/" replace />} />
          </Routes>
        </Router>
        </AuthProvider>
      </ThemeProvider>
    </ErrorBoundary>
  );
}

export default App;
