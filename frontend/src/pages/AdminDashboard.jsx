import { useEffect, useState, useCallback } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { supabase } from '../supabaseClient';

const styles = {
  container: {
    padding: '40px',
    backgroundColor: 'var(--bg-primary)',
    minHeight: '100vh',
    fontFamily: 'Inter, sans-serif',
    color: 'white'
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '30px'
  },
  title: {
    margin: 0,
    fontSize: '28px',
    fontWeight: 'bold',
    color: 'var(--accent-color)'
  },
  subtitle: {
    margin: '5px 0 0 0',
    color: 'var(--text-secondary)'
  },
  card: {
    backgroundColor: 'var(--bg-secondary)',
    borderRadius: '12px',
    border: '1px solid var(--border-color)',
    padding: '24px',
    boxShadow: '0 8px 30px rgba(0,0,0,0.3)'
  },
  btn: {
    backgroundColor: 'var(--bg-secondary)',
    color: 'white',
    border: '1px solid var(--border)',
    padding: '8px 16px',
    borderRadius: '6px',
    cursor: 'pointer',
    transition: 'background 0.2s'
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    textAlign: 'left'
  },
  th: {
    padding: '12px 16px',
    borderBottom: '1px solid var(--border-color)',
    color: 'var(--text-secondary)',
    fontSize: '14px',
    textTransform: 'uppercase'
  },
  td: {
    padding: '16px',
    borderBottom: '1px solid var(--border-color)',
    fontSize: '15px'
  },
  toggleBtn: (active) => ({
    padding: '6px 12px',
    borderRadius: '20px',
    border: 'none',
    cursor: 'pointer',
    fontWeight: '600',
    fontSize: '12px',
    backgroundColor: active ? 'var(--accent-color)' : 'var(--bg-primary)',
    color: active ? 'white' : 'var(--muted)',
    transition: 'all 0.2s'
  })
};

const AdminDashboard = () => {
  const { profile } = useAuth();
  const navigate = useNavigate();
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);

  const fetchUsers = useCallback(async () => {
    setLoading(true);
    // Fetch profiles and permissions
    const { data: profiles, error: profErr } = await supabase.from('user_profiles').select('*');
    const { data: perms, error: permErr } = await supabase.from('user_permissions').select('*');

    if (!profErr && !permErr) {
      // Merge them
      const merged = (profiles || []).map(p => {
        const userPerms = (perms || []).find(pm => pm.user_id === p.id) || {};
        return { ...p, ...userPerms };
      });
      setUsers(merged);
    }
    setLoading(false);
  }, []);

  useEffect(() => {
    // If not loaded yet, just wait.
    if (!profile && loading) return;
    
    // If loaded and not admin, kick them out.
    if (!profile?.is_admin) {
      alert("Unauthorized Access");
      navigate('/');
      return;
    }

    fetchUsers();
  }, [profile, loading, navigate, fetchUsers]);

  const togglePermission = async (userId, field, currentValue) => {
    const newValue = !currentValue;
    
    // Optimistic UI update
    setUsers(users.map(u => u.id === userId ? { ...u, [field]: newValue } : u));

    const { error } = await supabase
      .from('user_permissions')
      .update({ [field]: newValue })
      .eq('user_id', userId);

    if (error) {
      alert("Failed to update permission: " + error.message);
      // Revert on error
      fetchUsers();
    }
  };

  if (loading || !profile?.is_admin) return <div style={styles.container}>Loading Dashboard...</div>;

  return (
    <div style={styles.container}>
      <div style={styles.header}>
        <h1 style={styles.title}>Admin Access Dashboard</h1>
        <button style={styles.btn} onClick={() => navigate('/')}>Back to Portal</button>
      </div>

      <table style={styles.table}>
        <thead>
          <tr>
            <th style={styles.th}>User Email</th>
            <th style={styles.th}>Role</th>
            <th style={styles.th}>Report Studio Access</th>
            <th style={styles.th}>Timetable Access</th>
          </tr>
        </thead>
        <tbody>
          {users.map(u => (
            <tr key={u.id}>
              <td style={styles.td}>{u.email}</td>
              <td style={styles.td}>
                {u.is_admin ? <span style={{color: '#ffc078'}}>Admin</span> : <span style={{color: 'var(--muted)'}}>User</span>}
              </td>
              <td style={styles.td}>
                <button 
                  style={styles.toggleBtn(u.can_access_studio)}
                  onClick={() => togglePermission(u.id, 'can_access_studio', u.can_access_studio)}
                  disabled={u.is_admin} // Optional: Prevent toggling access for admins
                >
                  {u.can_access_studio ? 'Granted' : 'Revoked'}
                </button>
              </td>
              <td style={styles.td}>
                <button 
                  style={styles.toggleBtn(u.can_access_scheduler)}
                  onClick={() => togglePermission(u.id, 'can_access_scheduler', u.can_access_scheduler)}
                  disabled={u.is_admin}
                >
                  {u.can_access_scheduler ? 'Granted' : 'Revoked'}
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default AdminDashboard;
