import React, { useEffect, useState } from 'react';
import { supabase } from '../supabaseClient';
import { useAuth } from '../context/AuthContext';

const LibraryPage = ({ onLoadConfig, onLoadTemplate }) => {
  const { user } = useAuth();
  const [configs, setConfigs] = useState([]);
  const [templates, setTemplates] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetchLibrary();
  }, []);

  const fetchLibrary = async () => {
    setLoading(true);
    
    // Fetch Configs
    const { data: configsData, error: configsErr } = await supabase
      .from('configs')
      .select('*')
      .eq('user_id', user.id);

    if (configsErr) {
      console.error("Error fetching configs:", configsErr);
      alert("Error fetching configs: " + configsErr.message);
    } else {
      setConfigs(configsData || []);
    }

    // Fetch Templates
    const { data: templatesData, error: templatesErr } = await supabase
      .from('templates')
      .select('*')
      .eq('user_id', user.id);

    if (templatesErr) {
      console.error("Error fetching templates:", templatesErr);
      alert("Error fetching templates: " + templatesErr.message);
    } else {
      setTemplates(templatesData || []);
    }

    setLoading(false);
  };

  const deleteItem = async (table, id) => {
    if (!window.confirm("Are you sure you want to delete this item?")) return;
    
    const { error } = await supabase.from(table).delete().eq('id', id);
    if (error) {
      alert("Error deleting item: " + error.message);
    } else {
      fetchLibrary();
    }
  };

  return (
    <div style={{ padding: '40px', maxWidth: '1400px', margin: '0 auto', width: '100%' }}>
      <h2>My Library</h2>
      <p className="subtitle">Manage your saved configurations and templates.</p>

      {loading ? (
        <div style={{ textAlign: 'center', padding: '40px', color: 'var(--muted)' }}>Loading library...</div>
      ) : (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))', gap: '32px' }}>
          
          {/* CONFIGS LIST */}
          <div className="card" style={{ padding: '32px' }}>
            <h3 style={{ marginTop: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span className="icon">📄</span> Saved Configurations
            </h3>
            <p style={{ color: 'var(--muted)', fontSize: '13px', marginBottom: '24px' }}>
              Configurations contain your data pipeline rules (Filters, Advanced Calculations). To edit a config, you must upload your original dataset again.
            </p>
            
            {configs.length === 0 ? (
              <div style={{ padding: '24px', textAlign: 'center', background: '#f9fafb', borderRadius: '8px', color: 'var(--muted)' }}>
                No configurations saved yet.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {configs.map(c => (
                  <div key={c.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px', background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px' }}>
                    <div>
                      <strong style={{ display: 'block', marginBottom: '4px' }}>{c.name}</strong>
                    </div>
                    <div style={{ display: 'flex', gap: '8px' }}>
                      <button className="secondary" onClick={() => onLoadConfig(c)} style={{ padding: '6px 12px', fontSize: '12px', borderColor: 'var(--accent)', color: 'var(--accent)' }}>
                        Edit Rules
                      </button>
                      <button className="danger" onClick={() => deleteItem('configs', c.id)} style={{ padding: '6px 12px', fontSize: '12px' }}>
                        Delete
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* TEMPLATES LIST */}
          <div className="card" style={{ padding: '32px' }}>
            <h3 style={{ marginTop: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span className="icon">🎨</span> Saved Templates
            </h3>
            <p style={{ color: 'var(--muted)', fontSize: '13px', marginBottom: '24px' }}>
              Templates define the visual layout of your generated PDF or HTML report.
            </p>
            
            {templates.length === 0 ? (
              <div style={{ padding: '24px', textAlign: 'center', background: '#f9fafb', borderRadius: '8px', color: 'var(--muted)' }}>
                No templates saved yet.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {templates.map(t => (
                  <div key={t.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px', background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px' }}>
                    <div>
                      <strong style={{ display: 'block', marginBottom: '4px' }}>{t.name}</strong>
                      <span style={{ fontSize: '12px', color: 'var(--muted)' }}>
                        {t.layout_data?.pageSize || 'A4'}
                      </span>
                    </div>
                    <div style={{ display: 'flex', gap: '8px' }}>
                      <button className="secondary" onClick={() => onLoadTemplate(t)} style={{ padding: '6px 12px', fontSize: '12px', borderColor: 'var(--accent)', color: 'var(--accent)' }}>
                        Edit Layout
                      </button>
                      <button className="danger" onClick={() => deleteItem('templates', t.id)} style={{ padding: '6px 12px', fontSize: '12px' }}>
                        Delete
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

        </div>
      )}
    </div>
  );
};

export default LibraryPage;
