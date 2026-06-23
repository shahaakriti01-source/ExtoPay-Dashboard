import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import { msalInstance } from './authConfig.js'

async function start() {
  // MSAL must be initialized before any other MSAL API is called.
  await msalInstance.initialize();

  // Consume any leftover redirect-response (e.g. "#code=..." in the URL) so it
  // never gets stuck and confuses subsequent popup-based sign-in attempts.
  await msalInstance.handleRedirectPromise().catch((err) => {
    console.error("MSAL redirect handling error:", err);
  });

  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
  )
}

start();
