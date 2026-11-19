// ============================================================================
// STEP 1 OF THE SETUP GUIDE: paste your Firebase web app config here.
// Firebase console > Project settings > General > Your apps > SDK setup and configuration > Config.
// These values are not secret. Your data is protected by Google sign-in and the Firestore
// security rules (firestore.rules), not by hiding this file.
//
// While apiKey still starts with "PASTE", the app runs in DEMO MODE: no login, data kept
// only in this browser. Add ?demo=1 to the address to force demo mode at any time.
// ============================================================================
export const firebaseConfig = {
  apiKey: 'AIzaSyC6Go8OK9QA6_AcYzQxtQhXhwnL9CoHIH8',
  authDomain: 'mkf-vendor-payables.firebaseapp.com',
  projectId: 'mkf-vendor-payables',
  storageBucket: 'mkf-vendor-payables.firebasestorage.app',
  messagingSenderId: '737636206304',
  appId: '1:737636206304:web:e513e523a18b1e8b8ff285',
};

// Screen locks after this many minutes without any tap or key press. The PIN unlocks it.
export const IDLE_LOCK_MINUTES = 10;
