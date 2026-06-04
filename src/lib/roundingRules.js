// src/lib/roundingRules.js
// Manages rounding rules per country, stored in Firebase

import { db } from './firebase';
import { doc, getDoc, setDoc, deleteDoc, collection, getDocs } from 'firebase/firestore';

// ── Default rules (Cameroon, Congo) ──────────────────
export const DEFAULT_RULES = [
  { id: 'r1', label: '< 1 000',  threshold: 1000, step: 100,  operator: 'lt' },
  { id: 'r2', label: '≥ 1 000',  threshold: 1000, step: 500,  operator: 'gte' },
];

// Countries that use default rules
export const DEFAULT_COUNTRIES = ['CM', 'CG'];

// ── Apply rules to a value ────────────────────────────
export function applyRules(value, rules) {
  if (!value || isNaN(value)) return null;
  const sorted = [...rules].sort((a, b) => {
    if (a.operator === 'lt' && b.operator === 'gte') return -1;
    return a.threshold - b.threshold;
  });
  for (const rule of sorted) {
    const match = rule.operator === 'lt'  ? value < rule.threshold
                : rule.operator === 'gte' ? value >= rule.threshold
                : rule.operator === 'lte' ? value <= rule.threshold
                : rule.operator === 'gt'  ? value > rule.threshold
                : false;
    if (match) {
      return Math.floor(value / rule.step) * rule.step;
    }
  }
  // Fallback: floor to nearest 100
  return Math.floor(value / 100) * 100;
}

// ── Firebase CRUD ─────────────────────────────────────
export async function saveRulesForCountry(countryCode, rules) {
  const ref = doc(db, 'rounding_rules', countryCode);
  await setDoc(ref, { rules, updatedAt: new Date().toISOString() });
}

export async function loadRulesForCountry(countryCode) {
  const ref = doc(db, 'rounding_rules', countryCode);
  const snap = await getDoc(ref);
  if (snap.exists()) return snap.data().rules;
  return null; // null = use defaults
}

export async function deleteRulesForCountry(countryCode) {
  const ref = doc(db, 'rounding_rules', countryCode);
  await deleteDoc(ref);
}

export async function loadAllCountryRules() {
  const col = collection(db, 'rounding_rules');
  const snap = await getDocs(col);
  const result = {};
  snap.forEach(d => { result[d.id] = d.data().rules; });
  return result;
}
