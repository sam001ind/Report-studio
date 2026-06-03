import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import LandingPage from './pages/LandingPage';
import StudioLayout from './pages/StudioLayout';
import SchedulerPage from './pages/SchedulerPage';
import AuthPage from './pages/AuthPage';
import AdminDashboard from './pages/AdminDashboard';
import { AuthProvider, useAuth } from './context/AuthContext';

const ProtectedRoute = ({ children, requiredPermission }) => {
  const { user, permissions, profile } = useAuth();
  
  if (!user) return <Navigate to="/auth" />;
  
  // If a specific permission is required, and they are NOT an admin, check the permission
  if (requiredPermission && profile && !profile.is_admin) {
    if (!permissions || !permissions[requiredPermission]) {
      // User does not have access
      return (
        <div style={{ padding: '40px', color: 'white', textAlign: 'center' }}>
          <h2>Access Denied</h2>
          <p>You do not have permission to access this module.</p>
          <button onClick={() => window.location.href = '/'} style={{ padding: '10px 20px', background: 'var(--accent)', color: 'white', border: 'none', borderRadius: '5px', cursor: 'pointer' }}>Back to Home</button>
        </div>
      );
    }
  }
  
  return children;
};

function App() {
  return (
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
          <Route path="/scheduler" element={
            <ProtectedRoute requiredPermission="can_access_scheduler">
              <SchedulerPage />
            </ProtectedRoute>
          } />
        </Routes>
      </Router>
    </AuthProvider>
  );
}

export default App;
