import { createContext, useContext, useState } from 'react';

const AuthContext = createContext(null);

// eslint-disable-next-line react-refresh/only-export-components
export const useAuth = () => useContext(AuthContext);

export const AuthProvider = ({ children }) => {
  const [user] = useState({
    id: '00000000-0000-0000-0000-000000000000',
    email: 'admin@reportstudio.com'
  });
  
  const [profile] = useState({
    id: '00000000-0000-0000-0000-000000000000',
    email: 'admin@reportstudio.com',
    is_admin: true
  });
  
  const [permissions] = useState({
    user_id: '00000000-0000-0000-0000-000000000000',
    can_access_studio: true,
    can_access_scheduler: true
  });

  const value = {
    user,
    profile,
    permissions,
    signIn: async () => ({ data: { user }, error: null }),
    signUp: async () => ({ data: { user }, error: null }),
    signOut: async () => {},
  };

  return (
    <AuthContext.Provider value={value}>
      {children}
    </AuthContext.Provider>
  );
};
