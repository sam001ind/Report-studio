import React, { useState } from 'react';

const ConfigPage = ({ onDataLoaded }) => {
  const [uploadStatus, setUploadStatus] = useState('Click or Drag to Upload CSV / Excel / JSON');

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setUploadStatus(`Loaded: ${file.name}`);
    
    // Create FormData and send to Python backend
    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch('http://localhost:8000/api/upload', {
        method: 'POST',
        body: formData,
      });
      const data = await response.json();
      onDataLoaded(data); // Pass to parent App state
    } catch (error) {
      console.error('Error uploading file:', error);
      alert('Failed to upload and parse file. Is the Python backend running?');
    }
  };

  return (
    <div style={styles.page}>
      <h2>Report Configuration</h2>
      <p className="subtitle">Upload data to be processed by the Python backend.</p>

      <div className="card">
        <h3 style={{ marginTop: 0 }}>Step 1: Upload Dataset</h3>
        
        <div style={styles.fileDrop} onClick={() => document.getElementById('fileInput').click()}>
          <h3>{uploadStatus}</h3>
          <p style={{ color: 'var(--muted)', fontSize: '14px' }}>
            Data is sent to FastAPI and processed using Pandas.
          </p>
          <input 
            type="file" 
            id="fileInput" 
            accept=".csv, .xlsx, .xls, .json" 
            style={{ display: 'none' }}
            onChange={handleFileUpload}
          />
        </div>
      </div>
      
      {/* TODO: Add data management UI (filters, combines) mapped to backend endpoints */}
    </div>
  );
};

const styles = {
  page: {
    padding: '40px',
    maxWidth: '1400px',
    margin: '0 auto',
  },
  fileDrop: {
    border: '2px dashed var(--line)',
    borderRadius: '12px',
    padding: '40px',
    textAlign: 'center',
    backgroundColor: '#fafafa',
    cursor: 'pointer',
    transition: 'border-color 0.2s',
  }
};

export default ConfigPage;
