# Project Step Manager

Desktop Electron app per leggere e modificare piani lavori Excel salvati su SharePoint.

La struttura e' volutamente allineata a `Software compilazione offerte`: React + Vite, main process Electron con IPC, login Microsoft via MSAL, configurazione persistente con `electron-store`, Graph API per SharePoint.

## Funzioni

- Login Microsoft con cache locale protetta.
- Configurazione SharePoint da UI:
  - modalita site/drive o Microsoft 365 group;
  - rilevamento Drive ID;
  - cartella root dei progetti;
  - template Excel opzionale per nuovi progetti.
- Lista dei file `.xlsx` nella cartella configurata.
- Supporto cartelle progetto: il nuovo progetto crea una cartella in SharePoint e inserisce `Piano lavori.xlsx` al suo interno.
- Gestione diretta di cartelle progetto e piu' Excel per cartella, con cancellazione via Graph da UI.
- Lettura dei due formati Excel usati negli esempi:
  - `Piano lavori` con colonne backend/frontend/stato/stime/note;
  - formato storico `Fase/Sottocategoria/Task/BackEnd/FrontEnd/Stato/Note`.
- Editing inline dei task.
- Il campo note dell'app scrive su `Nota 1`.
- Stato modifiche non salvate con bottone `Salva`.
- Creazione nuovo progetto come nuovo `.xlsx` nella cartella SharePoint configurata.

## Permessi Azure

L'App Registration deve consentire almeno:

- `User.Read`
- `Files.ReadWrite.All`
- `Sites.Read.All`

Per uso desktop configurare una redirect URI pubblica/native client compatibile con MSAL interactive login. L'app usa `http://localhost`.

## Sviluppo

```powershell
npm install
npm run dev
```

## Build

```powershell
npm run build
npm run package
```
