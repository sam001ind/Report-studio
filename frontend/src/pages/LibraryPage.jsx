import { useEffect, useState, useCallback } from 'react';
import { supabase } from '../supabaseClient';
import { Sparkles, TableProperties, Plus, ArrowRight, FileText, Trash2, FolderOpen } from 'lucide-react';

const LibraryPage = ({ onLoadConfig, onLoadTemplate }) => {
  const [configs, setConfigs] = useState([]);
  const [templates, setTemplates] = useState([]);
  const [loading, setLoading] = useState(true);

  const fetchLibrary = useCallback(async () => {
    setLoading(true);
    
    let authUserId = null;
    try {
      const { data: authData } = await supabase.auth.getUser();
      if (authData?.user?.id) {
        authUserId = authData.user.id;
      }
    } catch {
      // not logged in
    }

    let cloudConfigs = [];
    let cloudTemplates = [];

    if (authUserId) {
      // Fetch Configs
      const { data: configsData } = await supabase
        .from('configs')
        .select('id, name, created_at, config_data')
        .not('config_data->>isDataset', 'eq', 'true')
        .eq('user_id', authUserId);
      if (configsData) cloudConfigs = configsData;

      // Fetch Templates
      const { data: templatesData } = await supabase
        .from('templates')
        .select('*')
        .eq('user_id', authUserId);
      if (templatesData) cloudTemplates = templatesData;
    }

    // Load Local Storage configs & templates
    let localConfigs = [];
    let localTemplates = [];
    try {
      localConfigs = JSON.parse(localStorage.getItem('saved_configs') || '[]');
      localTemplates = JSON.parse(localStorage.getItem('saved_templates') || '[]');
    } catch (e) {
      console.warn(e);
    }

    // Merge without duplicates
    const combinedConfigs = [...cloudConfigs];
    localConfigs.forEach(lc => {
      if (!combinedConfigs.some(c => c.name === lc.name || c.id === lc.id)) {
        combinedConfigs.push(lc);
      }
    });

    const combinedTemplates = [...cloudTemplates];
    localTemplates.forEach(lt => {
      if (!combinedTemplates.some(t => t.name === lt.name || t.id === lt.id)) {
        combinedTemplates.push(lt);
      }
    });

    setConfigs(combinedConfigs);
    setTemplates(combinedTemplates);
    setLoading(false);
  }, []);

  useEffect(() => {
    fetchLibrary();
  }, [fetchLibrary]);

  const deleteItem = async (table, id) => {
    if (!window.confirm('Are you sure you want to delete this item?')) return;
    
    try {
      await supabase.from(table).delete().eq('id', id);
    } catch (e) {
      console.warn(e);
    }

    try {
      const storageKey = table === 'configs' ? 'saved_configs' : 'saved_templates';
      const items = JSON.parse(localStorage.getItem(storageKey) || '[]');
      const updated = items.filter(it => it.id !== id);
      localStorage.setItem(storageKey, JSON.stringify(updated));
    } catch (e) {
      console.warn(e);
    }

    fetchLibrary();
  };

  return (
    <div style={{ padding: '32px 40px', maxWidth: '1400px', margin: '0 auto', width: '100%', overflowY: 'auto', fontFamily: 'var(--font-family)' }}>
      
      {/* Page Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px', flexWrap: 'wrap', gap: '16px' }}>
        <div>
          <h2 style={{ margin: 0, fontSize: '24px', fontWeight: 800, display: 'flex', alignItems: 'center', gap: '10px' }}>
            <FolderOpen size={24} color="var(--accent)" /> Report Template & Pipeline Library
          </h2>
          <p style={{ margin: '4px 0 0', color: 'var(--muted)', fontSize: '13.5px' }}>
            Manage your saved custom report designs, institutional headers, column bindings, and data transformation pipelines.
          </p>
        </div>

        <button 
          className="button"
          onClick={() => onLoadTemplate({ name: 'New Custom Tabular Report' })}
          style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', fontSize: '13.5px', padding: '9px 18px' }}
        >
          <Plus size={16} /> Create New Custom Report Template
        </button>
      </div>

      {/* Hero Action Card */}
      <div className="card" style={{ padding: '24px 28px', marginBottom: '32px', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), #fff)', border: '1.5px solid var(--accent)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px' }}>
          <div>
            <h3 style={{ margin: '0 0 6px 0', display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--accent)', fontSize: '16px' }}>
              <TableProperties size={20} /> Custom Tabular Data & Report Studio
            </h3>
            <p style={{ margin: 0, color: 'var(--muted)', fontSize: '13px', maxWidth: '750px', lineHeight: '1.4' }}>
              Upload any spreadsheet format (.xlsx, .csv, .tsv) to customize table column headers, institutional titles, page orientation, logo, styling, and multi-format exports (.pdf, .xlsx, .csv, .zip).
            </p>
          </div>
          <button 
            className="button"
            onClick={() => onLoadTemplate({ name: 'New Custom Tabular Report' })}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '8px 16px' }}
          >
            Launch Template Studio <ArrowRight size={15} />
          </button>
        </div>
      </div>

      {loading ? (
        <div style={{ textAlign: 'center', padding: '40px', color: 'var(--muted)' }}>Loading library...</div>
      ) : (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(400px, 1fr))', gap: '32px' }}>
          
          {/* SAVED CUSTOM TEMPLATES */}
          <div className="card" style={{ padding: '28px' }}>
            <h3 style={{ marginTop: 0, display: 'flex', alignItems: 'center', gap: '8px', fontSize: '16px' }}>
              <Sparkles size={18} color="var(--accent)" /> Saved Custom Templates ({templates.length})
            </h3>
            <p style={{ color: 'var(--muted)', fontSize: '12.5px', marginBottom: '20px' }}>
              Your saved custom report templates with customized headers, column mappings, page orientation, and styles.
            </p>
            
            {templates.length === 0 ? (
              <div style={{ padding: '28px', textAlign: 'center', background: '#f8fafc', borderRadius: '8px', border: '1px dashed var(--line)', color: 'var(--muted)' }}>
                No custom report templates saved yet. Click &quot;Create New Custom Report Template&quot; to build your first template.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {templates.map(t => (
                  <div key={t.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '14px 16px', background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px' }}>
                    <div>
                      <strong style={{ display: 'block', marginBottom: '3px', fontSize: '13.5px' }}>{t.name}</strong>
                      <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                        {t.layout_data?.createdAt ? new Date(t.layout_data.createdAt).toLocaleDateString() : 'Recent'} • Custom Tabular
                      </span>
                    </div>
                    <div style={{ display: 'flex', gap: '8px' }}>
                      <button className="button secondary" onClick={() => onLoadTemplate(t.layout_data ? { ...t.layout_data, name: t.name } : t)} style={{ padding: '6px 12px', fontSize: '12px', borderColor: 'var(--accent)', color: 'var(--accent)' }}>
                        Open in Studio
                      </button>
                      <button className="button danger" onClick={() => deleteItem('templates', t.id)} style={{ padding: '6px 10px', fontSize: '12px' }}>
                        <Trash2 size={13} />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* SAVED DATA CONFIGURATIONS */}
          <div className="card" style={{ padding: '28px' }}>
            <h3 style={{ marginTop: 0, display: 'flex', alignItems: 'center', gap: '8px', fontSize: '16px' }}>
              <FileText size={18} color="var(--accent)" /> Saved Data Pipeline Configurations ({configs.length})
            </h3>
            <p style={{ color: 'var(--muted)', fontSize: '12.5px', marginBottom: '20px' }}>
              Configurations contain your data pipeline transformation rules (Filter conditions, Concatenations, Advanced Calculations).
            </p>
            
            {configs.length === 0 ? (
              <div style={{ padding: '28px', textAlign: 'center', background: '#f8fafc', borderRadius: '8px', border: '1px dashed var(--line)', color: 'var(--muted)' }}>
                No data pipeline configurations saved yet.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {configs.map(c => (
                  <div key={c.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '14px 16px', background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px' }}>
                    <div>
                      <strong style={{ display: 'block', marginBottom: '3px', fontSize: '13.5px' }}>{c.name}</strong>
                      <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                        {c.config_data?.createdAt ? new Date(c.config_data.createdAt).toLocaleDateString() : 'Recent'}
                      </span>
                    </div>
                    <div style={{ display: 'flex', gap: '8px' }}>
                      <button className="button secondary" onClick={() => onLoadConfig(c)} style={{ padding: '6px 12px', fontSize: '12px', borderColor: 'var(--accent)', color: 'var(--accent)' }}>
                        Edit Rules
                      </button>
                      <button className="button danger" onClick={() => deleteItem('configs', c.id)} style={{ padding: '6px 10px', fontSize: '12px' }}>
                        <Trash2 size={13} />
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
