import { useEffect, useState } from 'react';
import Sidebar from './components/Sidebar';
import SettingsPage from './components/SettingsPage';
import ProjectWorkspace from './components/ProjectWorkspace';
import { applyDensity, applyTheme, getDensityPreference, getThemePreference } from './theme';

type Page = 'projects' | 'settings';

interface UserInfo {
  name: string;
  email: string;
}

export default function App() {
  const [page, setPage] = useState<Page>('projects');
  const [user, setUser] = useState<UserInfo | null>(null);
  const [loading, setLoading] = useState(true);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);

  useEffect(() => {
    applyTheme(getThemePreference());
    applyDensity(getDensityPreference());
    window.api.auth.getAccount().then((account) => {
      setUser(account);
      setLoading(false);
    });
  }, []);

  const handleLogin = async () => {
    try {
      setLoading(true);
      const account = await window.api.auth.login();
      setUser(account);
    } catch (error: unknown) {
      alert(`Errore login: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setLoading(false);
    }
  };

  const handleLogout = async () => {
    await window.api.auth.logout();
    setUser(null);
  };

  if (loading) {
    return (
      <div className="login-screen">
        <div className="spinner large" />
        <p>Caricamento...</p>
      </div>
    );
  }

  if (!user && page !== 'settings') {
    return (
      <div className="login-screen">
        <div className="login-panel">
          <p className="eyebrow">Sorgetech project control</p>
          <h1>Project Step Manager</h1>
          <p>Accedi con Microsoft per leggere e aggiornare i piani lavori Excel su SharePoint.</p>
          <div className="login-actions">
            <button className="btn btn-primary" onClick={handleLogin}>Accedi con Microsoft</button>
            <button className="btn btn-outline" onClick={() => setPage('settings')}>Impostazioni</button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <>
      <Sidebar
        currentPage={page}
        onPageChange={setPage}
        user={user}
        onLogout={handleLogout}
        collapsed={sidebarCollapsed}
        onCollapsedChange={setSidebarCollapsed}
      />
      <main className="main-content">
        {page === 'projects' && user && <ProjectWorkspace />}
        {page === 'settings' && (
          <div className="main-content-inner narrow">
            <div className="page-header">
              <div>
                <p className="eyebrow">Preferenze locali</p>
                <h1>Impostazioni</h1>
              </div>
            </div>
            <SettingsPage />
          </div>
        )}
      </main>
    </>
  );
}
