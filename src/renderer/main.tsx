import React from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import { installMockApi } from './mock-api';
import './styles/globals.css';

installMockApi();

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
