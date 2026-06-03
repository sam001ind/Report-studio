import React from 'react';
import { Sun, Moon } from 'lucide-react';
import { useTheme } from '../context/ThemeContext';

const ThemeToggle = ({ style }) => {
  const { isDark, toggleTheme } = useTheme();

  return (
    <button 
      onClick={toggleTheme}
      style={{
        padding: '8px', 
        borderRadius: '50%', 
        border: 'none', 
        background: isDark ? 'rgba(0, 240, 255, 0.1)' : 'transparent',
        cursor: 'pointer',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        transition: 'all 0.3s ease',
        ...style
      }}
      title="Toggle Luminous Theme"
    >
      {isDark ? (
        <Sun size={20} color="var(--accent)" style={{ filter: 'drop-shadow(0 0 5px var(--accent))' }} />
      ) : (
        <Moon size={20} color="var(--ink)" />
      )}
    </button>
  );
};

export default ThemeToggle;
