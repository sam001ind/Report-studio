import React, { useState, useEffect } from 'react';
import { useLocation } from 'react-router-dom';
import Sidebar from '../components/Sidebar';
import ConfigPage from './ConfigPage';
import TemplatePage from './TemplatePage';
import GeneratePage from './GeneratePage';

const StudioLayout = () => {
  const location = useLocation();
  const [activePage, setActivePage] = useState(location.state?.activePage || 'config');
  const [stats, setStats] = useState({ rows: 0, cols: 0 });
  const [dataset, setDataset] = useState({ columns: [], rows: [] });

  useEffect(() => {
    if (location.state?.activePage) {
      setActivePage(location.state.activePage);
    }
  }, [location.state]);

  return (
    <div style={{ display: 'flex', height: '100vh', width: '100vw', overflow: 'hidden' }}>
      <Sidebar 
        activePage={activePage} 
        setActivePage={setActivePage} 
        stats={stats} 
      />
      
      <main style={{ flex: 1, display: 'flex', flexDirection: 'column', overflowY: 'auto', position: 'relative' }}>
        {activePage === 'config' && (
          <ConfigPage 
            onDataLoaded={(data) => {
              setStats({ rows: data.rows, cols: data.columns.length });
              setDataset({ columns: data.columns, rows: data.data });
            }}
            dataset={dataset}
            setDataset={setDataset}
            setStats={setStats}
          />
        )}
        {activePage === 'template' && <TemplatePage dataset={dataset} />}
        {activePage === 'generate' && <GeneratePage dataset={dataset} />}
      </main>
    </div>
  );
};

export default StudioLayout;
