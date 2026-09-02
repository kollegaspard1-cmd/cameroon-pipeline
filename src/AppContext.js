// src/AppContext.js
// Shared state between pages for jumelage features
import React, { createContext, useContext, useState } from 'react';
import { T as TRANSLATIONS, COUNTRIES } from './translations';

const AppContext = createContext(null);

export function AppProvider({ children }) {
  // ── Langue et Pays globaux ──────────────────────────────────────────────
  const [lang,    setLang]    = useState('FR');  // 'FR' | 'EN'
  const [country, setCountry] = useState('CM');  // code pays ex: 'CM'

  // Traductions courantes selon la langue active
  const tr = TRANSLATIONS[lang] || TRANSLATIONS['FR'];
  const currentCountry = COUNTRIES.find(c => c.code === country) || COUNTRIES[0];

  // Jumelage A: Objectifs → Réalisation
  const [objectifsFiles,   setObjectifsFiles]   = useState([]);
  const [sharedRecap,      setSharedRecap]       = useState(null);
  const [sharedRecapName,  setSharedRecapName]   = useState('');

  // Jumelage B: Betshop Marchés → Multi-Périodes
  const [betshopForMulti, setBetshopForMulti] = useState(null);

  // Jumelage C: Retail → autres pages
  const [retailFilter, setRetailFilter] = useState(null);
  const [navigateTo,   setNavigateTo]   = useState(null);

  return (
    <AppContext.Provider value={{
      // Langue & Pays
      lang, setLang,
      country, setCountry,
      tr,               // raccourci : tr.tickets, tr.caissiers, etc.
      currentCountry,   // { code, name, flag, currency, prefix }
      COUNTRIES,        // liste complète des pays disponibles
      // Jumelage A
      objectifsFiles, setObjectifsFiles,
      sharedRecap, setSharedRecap,
      sharedRecapName, setSharedRecapName,
      // Jumelage B
      betshopForMulti, setBetshopForMulti,
      // Jumelage C
      retailFilter, setRetailFilter,
      navigateTo, setNavigateTo,
    }}>
      {children}
    </AppContext.Provider>
  );
}

export function useAppContext() {
  return useContext(AppContext);
}