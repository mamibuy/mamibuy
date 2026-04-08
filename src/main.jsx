import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'

if (!window.storage) {
  window.storage = {
    async set(key, value) { try { localStorage.setItem(key, value); return { key, value }; } catch { return null; } },
    async get(key) { try { const value = localStorage.getItem(key); return value ? { key, value } : null; } catch { return null; } },
    async delete(key) { try { localStorage.removeItem(key); return { key, deleted: true }; } catch { return null; } },
    async list(prefix = '') { try { return { keys: Object.keys(localStorage).filter(k => k.startsWith(prefix)) }; } catch { return { keys: [] }; } },
  };
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode><App /></React.StrictMode>
)
