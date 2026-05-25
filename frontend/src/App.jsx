import { useState } from 'react';
import Sidebar from './components/Sidebar';
import ConfigPage from './pages/ConfigPage';
import TemplatePage from './pages/TemplatePage';
import GeneratePage from './pages/GeneratePage';

function App() {
  const [activePage, setActivePage] = useState('config');
  const [stats, setStats] = useState({ rows: 0, cols: 0 });

  return (
    <div style={{ display: 'flex', height: '100vh', width: '100vw' }}>
      <Sidebar 
        activePage={activePage} 
        setActivePage={setActivePage} 
        stats={stats} 
      />
      
      <main style={{ flex: 1, overflowY: 'auto', position: 'relative' }}>
        {activePage === 'config' && <ConfigPage onDataLoaded={(data) => setStats({ rows: data.rows, cols: data.cols })} />}
        {activePage === 'template' && <TemplatePage />}
        {activePage === 'generate' && <GeneratePage />}
      </main>
    </div>
  );
}

export default App;
