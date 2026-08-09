// src/firebase.js
// ⚠️  APNA FIREBASE CONFIG YAHAN PASTE KARO
// (Setup guide neeche README mein hai)

import { initializeApp } from "firebase/app";
import { getDatabase } from "firebase/database";
import { getAuth } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyB_X6pjQPrHdaOei3Chz5BieiwHdnAKt3Q",
  authDomain: "library-management-b6bc6.firebaseapp.com",
  databaseURL: "https://library-management-b6bc6-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "library-management-b6bc6",
  storageBucket: "library-management-b6bc6.firebasestorage.app",
  messagingSenderId: "11609878828",
  appId: "1:11609878828:web:87cd3f02ff8e173119bc89",
  measurementId: "G-1SV5LJS2H6"
};

const app = initializeApp(firebaseConfig);

const db = getDatabase(app);
const auth = getAuth(app);

export { db, auth };
