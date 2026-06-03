// src/AppContext.js
// Shared state between pages for jumelage features
import React, { createContext, useContext, useState } from 'react';

const AppContext = createContext(null);

export function AppProvider({ children }) {
  // Jumelage A: Objectifs → Réalisation
  const [objectifsFiles,   setObjectifsFiles]   = useState([]); // [{blob, filename, mois, type}]
  const [sharedRecap,      setSharedRecap]       = useState(null); // File object
  const [sharedRecapName,  setSharedRecapName]   = useState('');

  // Jumelage B: Betshop Marchés → Multi-Périodes
  const [betshopForMulti, setBetshopForMulti] = useState(null);

  return (
    <AppContext.Provider value={{
      objectifsFiles, setObjectifsFiles,
      sharedRecap, setSharedRecap,
      sharedRecapName, setSharedRecapName,
      betshopForMulti, setBetshopForMulti,
    }}>
      {children}
    </AppContext.Provider>
  );
}

export function useAppContext() {
  return useContext(AppContext);
}
