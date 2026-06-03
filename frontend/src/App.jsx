import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import LandingPage from './pages/LandingPage';
import StudioLayout from './pages/StudioLayout';
import SchedulerPage from './pages/SchedulerPage';

function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<LandingPage />} />
        <Route path="/studio" element={<StudioLayout />} />
        <Route path="/scheduler" element={<SchedulerPage />} />
      </Routes>
    </Router>
  );
}

export default App;
