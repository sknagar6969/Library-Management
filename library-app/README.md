# 📚 Library Manager — Setup & Deployment Guide

## Kya kya features hain?
- ✅ Firebase Realtime Database — Live sync (phone + laptop dono pe ek saath)
- ✅ Koi bhi device se access — sirf URL kholo
- ✅ Data kabhi delete nahi hoga (Google ke servers pe)
- ✅ Excel Download — latest student list
- ✅ Seat pe Edit button — fees update karo
- ✅ Expiry List — kaun expire hone wala hai
- ✅ Real-time live indicator

---

## STEP 1 — Firebase Project Banao (FREE)

1. **https://console.firebase.google.com** pe jao
2. **"Add project"** click karo
3. Project name: `library-manager` → Continue → Continue → Create Project
4. Left side mein **"Realtime Database"** click karo
5. **"Create Database"** → Location: **Asia Southeast (Singapore)** → **"Start in test mode"** → Enable

---

## STEP 2 — Firebase Config Copy Karo

1. Firebase Console mein upar **⚙️ Project Settings** click karo
2. Neeche scroll karo **"Your apps"** section mein
3. **`</>`** (Web app) icon click karo
4. App nickname: `library-web` → Register app
5. Tumhe ek config milega aise:

```js
const firebaseConfig = {
  apiKey: "AIzaSy...",
  authDomain: "library-manager-xxxx.firebaseapp.com",
  databaseURL: "https://library-manager-xxxx-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "library-manager-xxxx",
  storageBucket: "library-manager-xxxx.appspot.com",
  messagingSenderId: "123456789",
  appId: "1:123456789:web:abcdef"
};
```

6. **`src/firebase.js`** file kholo aur apna config paste karo (replace karo `YOUR_API_KEY` etc.)

---

## STEP 3 — Project Run Karo (Local Test)

```bash
# Project folder mein jao
cd library-app

# Dependencies install karo
npm install

# Local test
npm start
# Browser mein http://localhost:3000 khulega
```

---

## STEP 4 — Vercel pe Deploy Karo (FREE Hosting)

### Option A: Vercel CLI (Easy)

```bash
# Vercel install karo
npm install -g vercel

# Build karo
npm run build

# Deploy karo
vercel

# Pehli baar: email se login karega
# Baad mein: URL milega jaise https://library-manager-xyz.vercel.app
```

### Option B: GitHub + Vercel (Best — Auto Deploy)

1. **https://github.com** pe free account banao
2. New repository banao: `library-manager`
3. Code upload karo (ya GitHub Desktop use karo)
4. **https://vercel.com** pe jao → GitHub se login
5. "New Project" → apna repo select karo → Deploy
6. **Done!** URL milega — isko bookmark karo phone pe

---

## STEP 5 — Phone Pe Use Karna

1. Vercel ka URL kholo Chrome mein
2. **3 dots menu → "Add to Home Screen"**
3. Ab phone pe app jaise icon ban jayega!
4. **Ye URL kabhi bhi, kisi bhi device se kaam karega**

---

## Firebase Database Rules (Security)

Firebase Console → Realtime Database → Rules → ye paste karo:

```json
{
  "rules": {
    ".read": true,
    ".write": true
  }
}
```

> ⚠️ Ye basic setup hai. Agar sirf aap use kar rahe ho toh theek hai.
> Security ke liye baad mein password auth add kar sakte hain.

---

## Files ka Structure

```
library-app/
├── public/
│   └── index.html          ← HTML shell
├── src/
│   ├── firebase.js         ← ⚠️ APNA CONFIG YAHAN DALO
│   ├── LibraryManager.jsx  ← Main app
│   └── index.js            ← Entry point
├── package.json
└── README.md
```

---

## Data Backup (Kabhi Bhi)

Firebase Console → Realtime Database → 3 dots → **Export JSON**

Ye JSON file apne paas rakh lo — ye poora backup hai.

---

## Help chahiye?

Koi step mein problem aaye toh Claude se poochho! 🙂
