import React, { useState, useEffect } from 'react';
import { useLocation } from 'react-router-dom';
import Sidebar from '../components/Sidebar';
import ConfigPage from './ConfigPage';
import TemplatePage from './TemplatePage';
import GeneratePage from './GeneratePage';
import LibraryPage from './LibraryPage';

const StudioLayout = () => {
  const location = useLocation();
  const [activePage, setActivePage] = useState(location.state?.activePage || 'library');
  const [stats, setStats] = useState({ rows: 0, cols: 0 });
  const [dataset, setDataset] = useState({ columns: [], rows: [] });
  
  const [loadedConfig, setLoadedConfig] = useState(null);
  const [loadedTemplate, setLoadedTemplate] = useState(null);

  const handleLoadConfig = (config) => {
    setLoadedConfig(config);
    setActivePage('config');
  };

  const handleLoadTemplate = (template) => {
    setLoadedTemplate(template);
    setActivePage('template');
  };

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
        {activePage === 'library' && (
          <LibraryPage 
            onLoadConfig={handleLoadConfig}
            onLoadTemplate={handleLoadTemplate}
          />
        )}
        {activePage === 'config' && (
          <ConfigPage 
            onDataLoaded={(data) => {
              setStats({ rows: data.rows, cols: data.columns.length });
              setDataset({ columns: data.columns, rows: data.data });
            }}
            dataset={dataset}
            setDataset={setDataset}
            setStats={setStats}
            initialConfig={loadedConfig}
          />
        )}
        {activePage === 'template' && <TemplatePage dataset={dataset} initialTemplate={loadedTemplate} />}
        {activePage === 'generate' && <GeneratePage dataset={dataset} />}
      </main>
    </div>
  );
};

export default StudioLayout;
