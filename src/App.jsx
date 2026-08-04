import React, { useState, useEffect } from 'react';
import Auth from './components/Auth';
import HardwareInventoryApp from './HardwareInventoryApp';
import { Loader2 } from 'lucide-react';

function App() {
  const [user, setUser] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    // Check active session on startup
    window.db.getCurrentUser()
      .then(currentUser => {
        setUser(currentUser);
      })
      .catch(err => {
        console.error("Session check failed:", err);
      })
      .finally(() => {
        setLoading(false);
      });
  }, []);

  if (loading) {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', width: '100vw', height: '100vh', background: '#F6F3EC', color: '#746C5E', fontFamily: 'sans-serif' }}>
        <Loader2 size={32} style={{ animation: 'spin 1s linear infinite', marginBottom: 12, color: '#D9720B' }} />
        <span style={{ fontSize: 14, fontWeight: 500 }}>Connecting to secure portal...</span>
        <style>{`
          @keyframes spin {
            to { transform: rotate(360deg); }
          }
        `}</style>
      </div>
    );
  }

  return (
    <div style={{ width: '100vw', height: '100vh', margin: 0, padding: 0, overflow: 'hidden' }}>
      {user ? (
        <HardwareInventoryApp user={user} onLogout={() => setUser(null)} />
      ) : (
        <Auth onAuthSuccess={(authenticatedUser) => setUser(authenticatedUser)} />
      )}
    </div>
  );
}

export default App;
