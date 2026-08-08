import React, { createContext, useContext, useEffect, useState } from 'react';

const AuthContext = createContext();

export const useAuth = () => {
  return useContext(AuthContext);
};

export const AuthProvider = ({ children }) => {
  // Hardcoded active admin session to bypass login entirely
  const [user, setUser] = useState({
    id: '00000000-0000-0000-0000-000000000000',
    email: 'admin@reportstudio.com'
  });
  
  const [profile, setProfile] = useState({
    id: '00000000-0000-0000-0000-000000000000',
    email: 'admin@reportstudio.com',
    is_admin: true
  });
  
  const [permissions, setPermissions] = useState({
    user_id: '00000000-0000-0000-0000-000000000000',
    can_access_studio: true,
    can_access_scheduler: true
  });
  
  const [loading, setLoading] = useState(false);

  const value = {
    user,
    profile,
    permissions,
    signIn: async (email, password) => { return { data: { user }, error: null }; },
    signUp: async (email, password) => { return { data: { user }, error: null }; },
    signOut: async () => { /* No-op to remain permanently logged in */ },
  };

  return (
    <AuthContext.Provider value={value}>
      {!loading && children}
    </AuthContext.Provider>
  );
};
