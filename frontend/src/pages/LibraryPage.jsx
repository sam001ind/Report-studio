import { useEffect, useState, useCallback } from 'react';
import { supabase } from '../supabaseClient';
import { TEMPLATE_ARCHETYPES } from '../utils/templateEngine';
import { Sparkles, FileText, CalendarRange, Tag, TableProperties, ArrowRight } from 'lucide-react';

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
    <div style={{ padding: '40px', maxWidth: '1400px', margin: '0 auto', width: '100%', overflowY: 'auto' }}>
      <h2>Report Template Library & Starter Archetypes</h2>
      <p className="subtitle">Choose from standard university examination templates or manage your saved custom report designs.</p>

      {/* STARTER PRESET ARCHETYPES */}
      <div className="card" style={{ padding: '28px', marginBottom: '32px', background: 'linear-gradient(135deg, rgba(23,107,135,0.06), #fff)', border: '1.5px solid var(--accent)' }}>
        <h3 style={{ marginTop: 0, marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--accent)' }}>
          <Sparkles size={20} /> Built-in Examination & Report Archetypes
        </h3>
        <p style={{ color: 'var(--muted)', fontSize: '13px', marginBottom: '20px' }}>
          Select an archetype to launch directly into the Template Studio with customizable headers, columns, and A4 layouts:
        </p>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: '16px' }}>
          {Object.values(TEMPLATE_ARCHETYPES).map(arch => (
            <div 
              key={arch.id}
              onClick={() => onLoadTemplate({ archetype: arch.id, name: arch.name, config: arch.defaultConfig })}
              style={{
                padding: '18px',
                borderRadius: '10px',
                border: '1px solid var(--line)',
                background: 'white',
                cursor: 'pointer',
                display: 'flex',
                flexDirection: 'column',
                justifyContent: 'space-between',
                boxShadow: '0 2px 8px rgba(0,0,0,0.04)',
                transition: 'all 0.2s ease'
              }}
            >
              <div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                  {arch.id === 'NOMINAL_ROLL' && <FileText size={20} color="var(--accent)" />}
                  {arch.id === 'QP_STATEMENT' && <CalendarRange size={20} color="var(--accent)" />}
                  {arch.id === 'QP_COVER_LABEL' && <Tag size={20} color="var(--accent)" />}
                  {arch.id === 'CUSTOM_TABULAR' && <TableProperties size={20} color="var(--accent)" />}
                  <strong style={{ fontSize: '14.5px' }}>{arch.name}</strong>
                </div>
                <p style={{ margin: '0 0 16px 0', fontSize: '12px', color: 'var(--muted)', lineHeight: '1.4' }}>
                  {arch.description}
                </p>
              </div>

              <button 
                className="button"
                style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '6px', fontSize: '12.5px', padding: '7px 12px', width: '100%' }}
              >
                Customize Template <ArrowRight size={14} />
              </button>
            </div>
          ))}
        </div>
      </div>

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
                      <span style={{ fontSize: '12px', color: 'var(--muted)' }}>
                        {c.config_data?.createdAt ? new Date(c.config_data.createdAt).toLocaleString() : 'Date Unknown'}
                      </span>
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
              <span className="icon">🎨</span> Saved Custom Templates
            </h3>
            <p style={{ color: 'var(--muted)', fontSize: '13px', marginBottom: '24px' }}>
              Your saved custom templates with bound fields and tailored header information.
            </p>
            
            {templates.length === 0 ? (
              <div style={{ padding: '24px', textAlign: 'center', background: '#f9fafb', borderRadius: '8px', color: 'var(--muted)' }}>
                No custom templates saved yet.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {templates.map(t => (
                  <div key={t.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px', background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px' }}>
                    <div>
                      <strong style={{ display: 'block', marginBottom: '4px' }}>{t.name}</strong>
                      <span style={{ fontSize: '12px', color: 'var(--muted)' }}>
                        {t.layout_data?.archetype || 'Custom'} • {t.layout_data?.createdAt ? new Date(t.layout_data.createdAt).toLocaleString() : 'Date Unknown'}
                      </span>
                    </div>
                    <div style={{ display: 'flex', gap: '8px' }}>
                      <button className="secondary" onClick={() => onLoadTemplate(t.layout_data ? { ...t.layout_data, name: t.name } : t)} style={{ padding: '6px 12px', fontSize: '12px', borderColor: 'var(--accent)', color: 'var(--accent)' }}>
                        Open in Studio
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
