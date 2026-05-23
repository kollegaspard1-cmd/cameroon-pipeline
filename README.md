# Pipeline CM — Cameroon Daily Campaign Engine

A browser-based web app that runs the Cameroon player segmentation pipeline entirely client-side (no server needed), with Firebase for run history and file storage, and Netlify for hosting.

---

## Architecture

```
Your browser
  ├── Upload Excel (MTD/YTD/Last 3 days) + CSV (transactions)
  ├── Pipeline runs in-browser (JavaScript port of cameroon_pipeline.py)
  ├── Downloads output files as individual files or .zip
  └── Optionally saves run history + files to Firebase
```

No server. No Python. Works entirely offline after the page loads.

---

## Setup

### 1. Install dependencies

```bash
npm install
```

### 2. Create a Firebase project

1. Go to https://console.firebase.google.com
2. Create a new project (e.g. `cameroon-pipeline`)
3. Click **Add app → Web**, register app, copy config
4. Enable **Firestore Database** (start in production mode, then update rules below)
5. Enable **Storage** (start in production mode, then update rules below)

### 3. Configure Firebase

Edit `src/lib/firebase.js` and replace the placeholder values:

```js
const firebaseConfig = {
  apiKey:            "AIza...",
  authDomain:        "your-project.firebaseapp.com",
  projectId:         "your-project-id",
  storageBucket:     "your-project.appspot.com",
  messagingSenderId: "123456789",
  appId:             "1:123:web:abc..."
};
```

### 4. Set Firestore rules (for internal use without auth)

In the Firebase Console → Firestore → Rules:
```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if true;  // ← OK for internal tool, add auth for production
    }
  }
}
```

### 5. Set Storage rules

In Firebase Console → Storage → Rules:
```
rules_version = '2';
service firebase.storage {
  match /b/{bucket}/o {
    match /{allPaths=**} {
      allow read, write: if true;  // ← same note as above
    }
  }
}
```

### 6. Run locally

```bash
npm start
```

Open http://localhost:3000

---

## Deploy to Netlify

### Option A: Netlify CLI

```bash
npm install -g netlify-cli
npm run build
netlify deploy --prod --dir=build
```

### Option B: Netlify UI (drag & drop)

1. Run `npm run build`
2. Go to https://app.netlify.com → drag the `build/` folder onto the deploy zone

### Option C: Connect your Git repo

1. Push this folder to GitHub
2. In Netlify: New site → Import from Git → select repo
3. Build command: `npm run build`
4. Publish directory: `build`
5. Click Deploy

The `netlify.toml` file already handles routing and build config.

---

## Usage

1. Open the app
2. Drop your `Cameroon_mtd_players.xlsx` (must have MTD, YTD, Last 3 days sheets)
3. Drop your `PLAYER_TRANSACTION_REPORT.csv` (must have Account ID, Total Gross columns)
4. Click **Run Pipeline**
5. View stats and download output files individually or as a .zip
6. Check **History** tab to see past runs stored in Firebase

---

## Output files

| File | Description |
|------|-------------|
| `Cameroon_final.xlsx` | MTD sheet — segmented players, all columns |
| `Cameroon_YTD_phones.xlsx` | Phone numbers from YTD (+ stripped) |
| `casino_bonus_cm.csv` | Casino bonus offer players |
| `Daily_Cashback_Template.csv` | Cash offer players (casino, no preference) |
| `spribe_freebets.xlsx` | Aviator freebet players (5 freebets, 7-day validity) |
| `CM_daily_XXX.csv` | Sport bonus players, one file per amount value |

---

## Firebase data structure

### Firestore: `pipeline_runs` collection
```json
{
  "createdAt": "Timestamp",
  "fileNames": { "excel": "...", "csv": "..." },
  "stats": {
    "totalPlayers": 123,
    "phonesExtracted": 456,
    "removedToday": 5,
    "byCampaign": { "casinoBonus": 20, "sportBonus": 80, "cashback": 15, "spribe": 10 }
  },
  "outputCount": 8
}
```

### Storage: `runs/{runId}/{filename}`
All generated output files are stored here per run.
