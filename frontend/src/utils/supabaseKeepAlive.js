import { supabase } from '../supabaseClient';

/**
 * Supabase Keep-Alive Heartbeat Utility
 * Sends a lightweight query to Supabase periodically to prevent the instance from sleeping or pausing.
 */
let heartbeatInterval = null;

export const pingSupabase = async () => {
  try {
    const { error } = await supabase
      .from('configs')
      .select('id')
      .limit(1);

    if (error) {
      // Fallback: ping auth session
      await supabase.auth.getSession();
    }
    
    // Store timestamp of last successful heartbeat
    try {
      localStorage.setItem('rs_supabase_last_ping', new Date().toISOString());
    } catch {
      // Ignore localStorage errors
    }
    return true;
  } catch {
    return false;
  }
};

export const startSupabaseKeepAlive = (intervalMs = 5 * 60 * 1000) => {
  // Run an immediate ping on startup
  pingSupabase();

  // Clear existing interval if already running
  if (heartbeatInterval) {
    clearInterval(heartbeatInterval);
  }

  // Set recurring interval (every 5 minutes)
  heartbeatInterval = setInterval(() => {
    pingSupabase();
  }, intervalMs);

  return () => {
    if (heartbeatInterval) {
      clearInterval(heartbeatInterval);
      heartbeatInterval = null;
    }
  };
};
