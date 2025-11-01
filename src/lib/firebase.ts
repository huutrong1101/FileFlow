import { initializeApp } from "firebase/app";
import { initializeFirestore, setLogLevel } from "firebase/firestore";

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

export const app = initializeApp(firebaseConfig);

// Ghi log chi tiết để thấy lỗi gốc (PERMISSION_DENIED, UNAUTHENTICATED, v.v.)
setLogLevel("debug");

// ⚠️ Vá mạng/proxy: ép Firestore dùng long-poll/fetch streams ổn định hơn
// Cách 1 (tự dò): tốt cho môi trường dev/Vite + proxy
export const db = initializeFirestore(app, {
  experimentalAutoDetectLongPolling: true,
});

// (Nếu vẫn lỗi, thử Cách 2:)
// export const db = initializeFirestore(app, {
//   experimentalForceLongPolling: true,
//   useFetchStreams: false,
// });

// (Giữ lại để tham chiếu)
// export const db = getFirestore(app);
