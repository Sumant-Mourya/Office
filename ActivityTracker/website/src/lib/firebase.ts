import { initializeApp, getApps } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";
// import { getAnalytics } from "firebase/analytics";

const firebaseConfig = {
  apiKey: "AIzaSyAXA7xovAVEMz2q3A4pPfuWQuf8M9mdSBk",
  authDomain: "acitivity-tracker-28.firebaseapp.com",
  projectId: "acitivity-tracker-28",
  storageBucket: "acitivity-tracker-28.firebasestorage.app",
  messagingSenderId: "231431292377",
  appId: "1:231431292377:web:9ef2a41bcf1855d27245c8",
  measurementId: "G-DLH8M95W17"
};

// Initialize Firebase
const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApps()[0];
const auth = getAuth(app);
const db = getFirestore(app);

// Analytics is only available in browser environment
// let analytics;
// if (typeof window !== "undefined") {
//   analytics = getAnalytics(app);
// }

export { app, auth, db };
