// src/lib/roundingRules.js
import { db } from './firebase';
import { doc, getDoc, setDoc, deleteDoc } from 'firebase/firestore';

export const DEFAULT_RULES = [
  { id: 'r1', label: '< 1 000',  threshold: 1000, threshold2: null, step: 100,  operator: 'lt' },
  { id: 'r2', label: '>= 1 000', threshold: 1000, threshold2: null, step: 500,  operator: 'gte' },
];

export const DEFAULT_COUNTRIES = ['CM', 'CG'];

export function applyRules(value, rules) {
  if (!value || isNaN(value)) return null;
  const sorted = [...rules].sort((a, b) => a.threshold - b.threshold);
  for (const rule of sorted) {
    let match = false;
    if (rule.operator === 'between') {
      match = value >= rule.threshold && value < (rule.threshold2 || rule.threshold + 1000);
    } else if (rule.operator === 'lt')  match = value < rule.threshold;
    else if (rule.operator === 'gte')   match = value >= rule.threshold;
    else if (rule.operator === 'lte')   match = value <= rule.threshold;
    else if (rule.operator === 'gt')    match = value > rule.threshold;
    if (match) return Math.floor(value / rule.step) * rule.step;
  }
  return Math.floor(value / 100) * 100;
}

export async function saveRulesForCountry(countryCode, rules) {
  const ref = doc(db, 'rounding_rules', countryCode);
  await setDoc(ref, { rules, updatedAt: new Date().toISOString() });
}

export async function loadRulesForCountry(countryCode) {
  const ref = doc(db, 'rounding_rules', countryCode);
  const snap = await getDoc(ref);
  if (snap.exists()) return snap.data().rules;
  return null;
}

export async function deleteRulesForCountry(countryCode) {
  const ref = doc(db, 'rounding_rules', countryCode);
  await deleteDoc(ref);
}
