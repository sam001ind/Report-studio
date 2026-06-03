import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import ThemeToggle from '../components/ThemeToggle';

const styles = {
  container: {
    display: 'flex',
    height: '100vh',
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: 'var(--bg)',
    fontFamily: 'Inter, sans-serif'
  },
  card: {
    backgroundColor: 'var(--panel)',
    padding: '40px',
    borderRadius: '12px',
    border: '1px solid var(--line)',
    boxShadow: '0 10px 30px rgba(0,0,0,0.5)',
    width: '100%',
    maxWidth: '400px',
    textAlign: 'center'
  },
  title: {
    margin: '0 0 10px 0',
    color: 'var(--ink)',
    fontSize: '24px'
  },
  subtitle: {
    color: 'var(--muted)',
    fontSize: '14px',
    margin: '0 0 30px 0'
  },
  input: {
    width: '100%',
    padding: '12px 15px',
    marginBottom: '15px',
    backgroundColor: 'var(--bg)',
    border: '1px solid var(--line)',
    borderRadius: '8px',
    color: 'var(--ink)',
    fontSize: '14px',
    boxSizing: 'border-box'
  },
  button: {
    width: '100%',
    padding: '12px',
    backgroundColor: 'var(--accent)',
    color: 'white',
    border: 'none',
    borderRadius: '8px',
    fontSize: '16px',
    cursor: 'pointer',
    fontWeight: 'bold',
    transition: 'background 0.2s',
    marginBottom: '15px'
  },
  switchText: {
    color: 'var(--muted)',
    fontSize: '14px'
  },
  switchLink: {
    color: 'var(--accent)',
    cursor: 'pointer',
    fontWeight: 'bold',
    marginLeft: '5px'
  },
  error: {
    color: '#ff6b6b',
    fontSize: '14px',
    marginBottom: '15px'
  }
};

const AuthPage = () => {
  const { user, signIn, signUp } = useAuth();
  const navigate = useNavigate();
  const [isLogin, setIsLogin] = useState(true);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');

  // Redirect if already logged in
  useEffect(() => {
    if (user) {
      navigate('/');
    }
  }, [user, navigate]);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setMessage('');
    setLoading(true);

    try {
      const loginEmail = email.trim() === 'admin' ? 'admin@reportstudio.com' : email;

      if (isLogin) {
        const { error } = await signIn(loginEmail, password);
        if (error) throw error;
        navigate('/'); // Immediate redirect
      } else {
        const { error } = await signUp(loginEmail, password);
        if (error) throw error;
        setMessage("Account created! You can now log in.");
        setIsLogin(true);
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={styles.container}>
      <div style={{ position: 'absolute', top: '24px', right: '24px' }}>
        <ThemeToggle />
      </div>
      <div style={styles.card}>
        <h2 style={styles.title}>Report Studio</h2>
        <p style={styles.subtitle}>
          {isLogin ? 'Sign in to access your templates and configs' : 'Create an account to save your templates'}
        </p>

        {error && <div style={styles.error}>{error}</div>}
        {message && <div style={{color: '#51cf66', marginBottom: '15px'}}>{message}</div>}

        <form onSubmit={handleSubmit}>
          <input 
            type="text" 
            placeholder="Email Address or Username" 
            style={styles.input}
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            required
          />
          <input 
            type="password" 
            placeholder="Password" 
            style={styles.input}
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            required
          />
          <button style={styles.button} type="submit" disabled={loading}>
            {loading ? 'Processing...' : (isLogin ? 'Sign In' : 'Sign Up')}
          </button>
        </form>

        <div style={styles.switchText}>
          {isLogin ? "Don't have an account?" : "Already have an account?"}
          <span 
            style={styles.switchLink} 
            onClick={() => {
              setIsLogin(!isLogin);
              setError('');
              setMessage('');
            }}
          >
            {isLogin ? 'Sign Up' : 'Sign In'}
          </span>
        </div>
      </div>
    </div>
  );
};

export default AuthPage;
