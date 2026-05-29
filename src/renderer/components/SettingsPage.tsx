import { useEffect, useState } from 'react';
import { Palette, Settings2, SlidersHorizontal } from 'lucide-react';
import type { AppSettings } from '../types/project';
import {
  DENSITIES,
  getDensityPreference,
  getThemePreference,
  setDensityPreference,
  setThemePreference,
  THEMES,
  type DensityName,
  type ThemeName,
} from '../theme';

export default function SettingsPage() {
  const [settings, setSettings] = useState<AppSettings | null>(null);
  const [azureConfig, setAzureConfig] = useState({ clientId: '', tenantId: '' });
  const [theme, setTheme] = useState<ThemeName>(() => getThemePreference());
  const [density, setDensity] = useState<DensityName>(() => getDensityPreference());
  const [saved, setSaved] = useState(false);
  const [azureSaved, setAzureSaved] = useState(false);
  const [resolvingDrive, setResolvingDrive] = useState(false);

  useEffect(() => {
    window.api.settings.get().then(setSettings);
    window.api.azure.get().then(setAzureConfig);
  }, []);

  const save = async () => {
    if (!settings) return;
    await window.api.settings.set(settings);
    setSaved(true);
    window.setTimeout(() => setSaved(false), 2000);
  };

  const saveAzure = async () => {
    await window.api.azure.set(azureConfig);
    setAzureSaved(true);
    window.setTimeout(() => setAzureSaved(false), 2000);
  };

  const changeTheme = (nextTheme: ThemeName) => {
    setTheme(nextTheme);
    setThemePreference(nextTheme);
  };

  const changeDensity = (nextDensity: DensityName) => {
    setDensity(nextDensity);
    setDensityPreference(nextDensity);
  };

  const resolveDrive = async () => {
    if (!settings) return;
    setResolvingDrive(true);
    try {
      await window.api.settings.set(settings);
      const driveId = await window.api.sharepointAdmin.resolveDrive();
      const updated = {
        ...settings,
        sharepoint: { ...settings.sharepoint, driveId },
      };
      setSettings(updated);
      await window.api.settings.set(updated);
      alert(`Drive ID rilevato: ${driveId}`);
    } catch (error: unknown) {
      alert(`Errore: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setResolvingDrive(false);
    }
  };

  if (!settings) return <p>Caricamento...</p>;

  const updateSharePoint = (field: string, value: string) => {
    setSettings({
      ...settings,
      sharepoint: { ...settings.sharepoint, [field]: value },
    });
  };

  const updateDefaults = (field: string, value: string | number) => {
    setSettings({
      ...settings,
      defaults: { ...settings.defaults, [field]: value },
    });
  };

  return (
    <>
      <section className="form-section personalization-section">
        <div className="section-title with-icon">
          <Palette size={18} />
          <div>
            <h2>Personalizzazione</h2>
            <p>Preferenze locali: cambiano subito l'interfaccia su questa macchina.</p>
          </div>
        </div>

        <div className="settings-block">
          <div className="settings-block-title">
            <h3>Tema</h3>
            <p>Scegli il tono visivo piu comodo per lavorare sui piani.</p>
          </div>
          <div className="theme-options" role="radiogroup" aria-label="Tema applicazione">
            {THEMES.map((item) => (
              <button
                key={item.value}
                className={`theme-option ${item.value}${theme === item.value ? ' active' : ''}`}
                onClick={() => changeTheme(item.value)}
                role="radio"
                aria-checked={theme === item.value}
              >
                <span className="theme-swatch" />
                <span>{item.label}</span>
              </button>
            ))}
          </div>
        </div>

        <div className="settings-block">
          <div className="settings-block-title with-icon">
            <SlidersHorizontal size={16} />
            <div>
              <h3>Densita</h3>
              <p>Usa la vista compatta quando vuoi piu righe visibili.</p>
            </div>
          </div>
          <div className="density-options" role="radiogroup" aria-label="Densita interfaccia">
            {DENSITIES.map((item) => (
              <button
                key={item.value}
                className={`density-option${density === item.value ? ' active' : ''}`}
                onClick={() => changeDensity(item.value)}
                role="radio"
                aria-checked={density === item.value}
              >
                <strong>{item.label}</strong>
                <span>{item.description}</span>
              </button>
            ))}
          </div>
        </div>

        <div className="settings-block">
          <div className="settings-block-title">
            <h3>Preferenze personali</h3>
            <p>Questo valore viene usato dal filtro "Miei" nella vista progetti.</p>
          </div>
          <div className="form-grid">
            <div className="form-group">
              <label>Il mio owner</label>
              <input
                value={settings.defaults.myOwner || ''}
                onChange={(event) => updateDefaults('myOwner', event.target.value)}
                placeholder="es. Sean"
              />
            </div>
          </div>
        </div>
      </section>

      <details className="advanced-settings">
        <summary>
          <Settings2 size={18} />
          <div>
            <strong>Avanzate</strong>
            <span>Azure, SharePoint, cartelle, template, owner globali e automazioni.</span>
          </div>
        </summary>

        <section className="form-section advanced-section">
          <div className="section-title">
            <h2>Azure / Microsoft</h2>
            <p>Stessa logica dell'app offerte: App Registration Entra ID, login interattivo e cache locale protetta.</p>
          </div>
          <div className="form-grid">
            <div className="form-group">
              <label>Application Client ID</label>
              <input
                value={azureConfig.clientId}
                onChange={(event) => setAzureConfig({ ...azureConfig, clientId: event.target.value })}
                placeholder="GUID applicazione"
              />
            </div>
            <div className="form-group">
              <label>Directory Tenant ID</label>
              <input
                value={azureConfig.tenantId}
                onChange={(event) => setAzureConfig({ ...azureConfig, tenantId: event.target.value })}
                placeholder="GUID tenant"
              />
            </div>
          </div>
          <div className="action-bar">
            <button className="btn btn-primary btn-sm" onClick={saveAzure}>Salva Azure</button>
            {azureSaved && <span className="save-confirm">Salvato</span>}
          </div>
        </section>

        <section className="form-section advanced-section">
          <div className="section-title">
            <h2>SharePoint</h2>
            <p>La cartella root contiene gli Excel progetto. Il template e' opzionale e viene usato solo per nuovi progetti.</p>
          </div>
          <div className="form-grid">
            <div className="form-group">
              <label>Modalita</label>
              <select
                value={settings.sharepoint.mode}
                onChange={(event) => updateSharePoint('mode', event.target.value)}
              >
                <option value="site">Site / Drive specifico</option>
                <option value="group">Group Drive Microsoft 365</option>
              </select>
            </div>

            {settings.sharepoint.mode === 'group' ? (
              <div className="form-group">
                <label>Group ID</label>
                <input
                  value={settings.sharepoint.groupId || ''}
                  onChange={(event) => updateSharePoint('groupId', event.target.value)}
                  placeholder="GUID gruppo M365"
                />
              </div>
            ) : (
              <>
                <div className="form-group">
                  <label>Site Hostname</label>
                  <input
                    value={settings.sharepoint.siteHostname || ''}
                    onChange={(event) => updateSharePoint('siteHostname', event.target.value)}
                    placeholder="es. contoso.sharepoint.com"
                  />
                </div>
                <div className="form-group">
                  <label>Site Path</label>
                  <input
                    value={settings.sharepoint.sitePath || ''}
                    onChange={(event) => updateSharePoint('sitePath', event.target.value)}
                    placeholder="es. /sites/NomeSito"
                  />
                </div>
                <div className="form-group full-width">
                  <label>Drive ID</label>
                  <div className="inline-field">
                    <input
                      value={settings.sharepoint.driveId || ''}
                      onChange={(event) => updateSharePoint('driveId', event.target.value)}
                      placeholder="Auto-rilevato o inserito manualmente"
                    />
                    <button className="btn btn-outline btn-sm" onClick={resolveDrive} disabled={resolvingDrive}>
                      {resolvingDrive ? 'Rilevamento...' : 'Rileva Drive'}
                    </button>
                  </div>
                </div>
              </>
            )}

            <div className="form-group">
              <label>Cartella progetti</label>
              <input
                value={settings.sharepoint.rootPath}
                onChange={(event) => updateSharePoint('rootPath', event.target.value)}
                placeholder="/Progetti"
              />
            </div>
            <div className="form-group">
              <label>Template nuovo progetto</label>
              <input
                value={settings.sharepoint.templatePath || ''}
                onChange={(event) => updateSharePoint('templatePath', event.target.value)}
                placeholder="/Template/template_piano_lavori_frontend.xlsx"
              />
            </div>
            <div className="form-group">
              <label>Cartella template</label>
              <input
                value={settings.sharepoint.templateFolderPath || ''}
                onChange={(event) => updateSharePoint('templateFolderPath', event.target.value)}
                placeholder="/Template"
              />
            </div>
            <div className="form-group">
              <label>Cartella archivio</label>
              <input
                value={settings.sharepoint.archiveFolderName || '_Archivio'}
                onChange={(event) => updateSharePoint('archiveFolderName', event.target.value)}
                placeholder="_Archivio"
              />
            </div>
          </div>
        </section>

        <section className="form-section advanced-section">
          <div className="section-title">
            <h2>Nuovi progetti e automazioni</h2>
          </div>
          <div className="form-grid">
            <div className="form-group">
              <label>Prefisso nome cartella</label>
              <input
                value={settings.defaults.projectNamePrefix}
                onChange={(event) => updateDefaults('projectNamePrefix', event.target.value)}
                placeholder="es. PRJ_"
              />
            </div>
            <div className="form-group">
              <label>Righe vuote iniziali</label>
              <input
                type="number"
                min={1}
                max={500}
                value={settings.defaults.initialRows}
                onChange={(event) => updateDefaults('initialRows', Number(event.target.value))}
              />
            </div>
            <div className="form-group full-width">
              <label>Owner ricorrenti</label>
              <textarea
                value={settings.defaults.owners || ''}
                onChange={(event) => updateDefaults('owners', event.target.value)}
                placeholder="Uno per riga o separati da virgola"
              />
            </div>
            <div className="form-group">
              <label>Email report</label>
              <input
                value={settings.defaults.notificationEmail || ''}
                onChange={(event) => updateDefaults('notificationEmail', event.target.value)}
                placeholder="team@sorgetech.it"
              />
            </div>
            <div className="form-group">
              <label>Webhook Teams</label>
              <input
                value={settings.defaults.teamsWebhookUrl || ''}
                onChange={(event) => updateDefaults('teamsWebhookUrl', event.target.value)}
                placeholder="URL Incoming Webhook"
              />
            </div>
          </div>
        </section>
      </details>

      <div className="action-bar sticky-actions">
        <button className="btn btn-primary" onClick={save}>Salva Impostazioni</button>
        {saved && <span className="save-confirm">Salvato</span>}
      </div>
    </>
  );
}
