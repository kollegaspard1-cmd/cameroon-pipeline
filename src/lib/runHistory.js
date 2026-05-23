// src/lib/runHistory.js
// Version sans Firebase Storage (forfait Spark)
import { db } from './firebase';
import {
  collection, addDoc, getDocs, query, orderBy, limit, serverTimestamp
} from 'firebase/firestore';

const RUNS_COL = 'pipeline_runs';

export async function saveRun({ stats, fileNames, outputFiles }) {
  const docRef = await addDoc(collection(db, RUNS_COL), {
    createdAt: serverTimestamp(),
    fileNames,
    stats,
    outputCount: outputFiles.length
  });
  return { runId: docRef.id, uploads: [] };
}

export async function loadRuns(n = 20) {
  const q = query(collection(db, RUNS_COL), orderBy('createdAt', 'desc'), limit(n));
  const snap = await getDocs(q);
  return snap.docs.map(d => ({ id: d.id, ...d.data() }));
}
