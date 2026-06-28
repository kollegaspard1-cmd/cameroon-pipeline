// src/lib/savedFiles.js
// Shared service for RECAP and LISTE persistence (Firebase + localStorage fallback)

import { db } from './firebase';
import { doc, setDoc, getDoc, deleteDoc } from 'firebase/firestore';

// ── RECAP DES SALLES PAR ZONE ─────────────────────

export async function saveRecap(zoneData, filename) {
  const nbSalles = Object.values(zoneData).flat().length;
  const savedAt  = new Date().toLocaleString('fr-FR');

  // Convert arrays to objects for Firestore
  const zoneDataObj = {};
  Object.entries(zoneData).forEach(([zone, rows]) => {
    zoneDataObj[zone] = rows.map(r => Array.isArray(r) ? { z:r[0], sup:r[1], sd:r[2], st:r[3] } : { z: r.ZONES||r.ZONE||r.z||'', sup: r.SUPER||r.sup||'', sd: r.SALLES||r.SALLE||r.sd||'', st: r.NOMS_TECH||r.NOM_TECH||r.st||'' });
  });

  const payload = { zoneData: zoneDataObj, filename, nbSalles, savedAt };

  // 1. Firebase
  await setDoc(doc(db, 'saved_files', 'recap_salles'), payload);

  // 2. localStorage fallback
  try {
    localStorage.setItem('recap_salles', JSON.stringify({ zoneData, filename, nbSalles, savedAt }));
  } catch(e) {}

  return { filename, nbSalles, savedAt };
}

export async function loadRecapSaved() {
  // Try Firebase first
  try {
    const snap = await getDoc(doc(db, 'saved_files', 'recap_salles'));
    if (snap.exists()) {
      const d = snap.data();
      const zoneData = {};
      Object.entries(d.zoneData).forEach(([zone, rows]) => {
        zoneData[zone] = rows.map(r => [r.z, r.sup, r.sd, r.st]);
      });
      // Sync to localStorage
      try { localStorage.setItem('recap_salles', JSON.stringify({ zoneData, filename: d.filename, nbSalles: d.nbSalles, savedAt: d.savedAt })); } catch(e) {}
      return { zoneData, meta: { filename: d.filename, nbSalles: d.nbSalles, savedAt: d.savedAt, source: 'Firebase' } };
    }
  } catch(e) {}

  // Fallback to localStorage
  try {
    const raw = localStorage.getItem('recap_salles');
    if (raw) {
      const d = JSON.parse(raw);
      return { zoneData: d.zoneData, meta: { filename: d.filename, nbSalles: d.nbSalles, savedAt: d.savedAt, source: 'localStorage' } };
    }
  } catch(e) {}

  return null;
}

export async function deleteRecap() {
  try { await deleteDoc(doc(db, 'saved_files', 'recap_salles')); } catch(e) {}
  try { localStorage.removeItem('recap_salles'); } catch(e) {}
}

// ── LISTE DES SALLES PAR PROPRIETAIRE ────────────

export async function saveListe(propMap, filename, cityMap) {
  const nbProps = Object.keys(propMap).length;
  const savedAt = new Date().toLocaleString('fr-FR');

  // localStorage always works
  try {
    localStorage.setItem('liste_salles', JSON.stringify({ propMap, cityMap: cityMap||{}, filename, nbProps, savedAt }));
  } catch(e) {}

  // Firebase with compatible structure
  try {
    const propsArray = Object.entries(propMap)
      .filter(([p, s]) => p && s && s.length > 0)
      .map(([prop, shops]) => ({
        prop: String(prop).trim(),
        shops: shops.filter(Boolean).map(String),
      }));
    // cityArray : [{shop, city}, ...]
    const cityArray = Object.entries(cityMap||{}).map(([shop, city]) => ({ shop, city }));
    await setDoc(doc(db, 'saved_files', 'liste_salles'), {
      propsArray, cityArray, filename, nbProps, savedAt,
    });
  } catch(e) {
    console.warn('Firebase liste failed, localStorage OK:', e.message);
  }

  return { filename, nbProps, savedAt };
}

export async function loadListe() {
  // Try Firebase first
  try {
    const snap = await getDoc(doc(db, 'saved_files', 'liste_salles'));
    if (snap.exists()) {
      const d = snap.data();
      const propMap = {};
      if (d.propsArray) {
        d.propsArray.forEach(({ prop, shops }) => { if (prop) propMap[prop] = shops || []; });
      }
      if (Object.keys(propMap).length > 0) {
        try { localStorage.setItem('liste_salles', JSON.stringify({ propMap, filename: d.filename, nbProps: d.nbProps, savedAt: d.savedAt })); } catch(e) {}
        // Reconstruire cityMap depuis cityArray
      const cityMap = {};
      if (d.cityArray) d.cityArray.forEach(({shop, city}) => { if (shop) cityMap[shop] = city; });
      return { propMap, cityMap, meta: { filename: d.filename, nbProps: d.nbProps, savedAt: d.savedAt, source: 'Firebase' } };
      }
    }
  } catch(e) {}

  // Fallback to localStorage
  try {
    const raw = localStorage.getItem('liste_salles');
    if (raw) {
      const d = JSON.parse(raw);
      return { propMap: d.propMap, cityMap: d.cityMap||{}, meta: { filename: d.filename, nbProps: d.nbProps, savedAt: d.savedAt, source: 'localStorage' } };
    }
  } catch(e) {}

  return null;
}

export async function deleteListe() {
  try { await deleteDoc(doc(db, 'saved_files', 'liste_salles')); } catch(e) {}
  try { localStorage.removeItem('liste_salles'); } catch(e) {}
}
