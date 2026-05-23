import { initializeApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';

const firebaseConfig = {
  apiKey:            "AIzaSyCjWxtUhJHjxkGI7m5fRJ-gK3llkJZjZg4",
  authDomain:        "daily-retention.firebaseapp.com",
  projectId:         "daily-retention",
  storageBucket:     "daily-retention.firebasestorage.app",
  messagingSenderId: "388993189248",
  appId:             "1:388993189248:web:258444d750ec5522596d65"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);