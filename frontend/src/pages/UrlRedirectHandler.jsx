import { useEffect, useState } from 'react';
import { useParams, Link } from 'react-router-dom';
import { ExternalLink, RefreshCw, AlertCircle } from 'lucide-react';
import { supabase } from '../supabaseClient';

const STORAGE_KEY = 'rs_short_urls_history';

const UrlRedirectHandler = () => {
  const { code } = useParams();
  const [status, setStatus] = useState('Resolving short link...');
  const [error, setError] = useState('');
  const [targetUrl, setTargetUrl] = useState('');

  useEffect(() => {
    const resolveLink = async () => {
      if (!code) {
        setError('No short link code provided.');
        return;
      }

      // 1. Check localStorage history
      try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved) {
          const list = JSON.parse(saved);
          const found = list.find(i => i.slug === code || i.id === code);
          if (found && found.longUrl) {
            setTargetUrl(found.longUrl);
            setStatus(`Redirecting to ${found.longUrl}...`);
            window.location.href = found.longUrl;
            return;
          }
        }
      } catch {
        // continue
      }

      // 2. Check Supabase configs table
      try {
        const { data, error: sbErr } = await supabase
          .from('configs')
          .select('config')
          .eq('name', `shortlink_${code}`)
          .single();

        if (!sbErr && data?.config?.longUrl) {
          const url = data.config.longUrl;
          setTargetUrl(url);
          setStatus(`Redirecting to ${url}...`);
          window.location.href = url;
          return;
        }
      } catch {
        // continue
      }

      setError(`Could not find a destination for short code "/s/${code}". The link may have expired or was created in another browser session.`);
    };

    resolveLink();
  }, [code]);

  return (
    <div style={{ minHeight: '80vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px', fontFamily: 'var(--font-family)' }}>
      <div className="card" style={{ maxWidth: '480px', width: '100%', padding: '36px', textAlign: 'center' }}>
        
        {!error ? (
          <div>
            <RefreshCw size={36} color="var(--accent)" className="spin" style={{ margin: '0 auto 16px' }} />
            <h2 style={{ fontSize: '20px', fontWeight: 800, marginBottom: '8px' }}>Opening Short Link</h2>
            <p style={{ color: 'var(--muted)', fontSize: '14px', margin: '0 0 20px 0' }}>
              {status}
            </p>
            {targetUrl && (
              <a 
                href={targetUrl} 
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: 'var(--accent)', fontWeight: 600 }}
              >
                Click here if you are not redirected automatically <ExternalLink size={14} />
              </a>
            )}
          </div>
        ) : (
          <div>
            <AlertCircle size={40} color="var(--danger)" style={{ margin: '0 auto 16px' }} />
            <h2 style={{ fontSize: '20px', fontWeight: 800, color: 'var(--danger)', marginBottom: '8px' }}>Link Not Found</h2>
            <p style={{ color: 'var(--muted)', fontSize: '13.5px', lineHeight: '1.5', margin: '0 0 24px 0' }}>
              {error}
            </p>
            <div style={{ display: 'flex', justifyContent: 'center', gap: '12px' }}>
              <Link to="/shortener" className="button" style={{ fontSize: '13.5px' }}>
                Go to URL Shortener
              </Link>
              <Link to="/" className="button secondary" style={{ fontSize: '13.5px' }}>
                Home Portal
              </Link>
            </div>
          </div>
        )}

      </div>
    </div>
  );
};

export default UrlRedirectHandler;
