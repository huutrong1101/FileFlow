import {
  collection,
  doc,
  getDocs,
  query,
  updateDoc,
  where,
  serverTimestamp,
  writeBatch,
  setDoc,
  deleteDoc,
} from "firebase/firestore";
import { db } from "../lib/firebase";

const USERS = "users";

export async function upsertUsersBulk(
  list: Array<{
    code: string;
    name: string;
    weightPct: number;
    online: boolean;
  }>
) {
  if (!Array.isArray(list) || list.length === 0) return;
  const batch = writeBatch(db);
  const now = serverTimestamp();
  for (const u of list) {
    const id = String(u.code || "").trim(); // id tài liệu = mã NV
    if (!id) continue;
    const ref = doc(db, USERS, id);
    batch.set(
      ref,
      {
        name: u.name ?? "",
        status: u.online ? "online" : "offline",
        weightPct: Number.isFinite(u.weightPct) ? u.weightPct : 0,
        active: true,
        updatedAt: now,
        createdAt: now,
      },
      { merge: true }
    );
  }
  await batch.commit(); // <-- bắt buộc
}

export async function listUsers(activeOnly = true) {
  const ref = collection(db, USERS);
  const q = activeOnly ? query(ref, where("active", "==", true)) : ref;
  const snap = await getDocs(q);
  return snap.docs.map((d) => {
    const data = d.data() as Record<string, unknown>;
    return {
      code: d.id,
      name: String(data.name ?? ""),
      weightPct: Number(data.weightPct ?? 0),
      online: (data.status ?? "online") === "online",
    };
  });
}

export async function updateUserOnline(userCode: string, online: boolean) {
  const ref = doc(db, USERS, userCode);
  await updateDoc(ref, {
    status: online ? "online" : "offline",
    updatedAt: serverTimestamp(),
  });
}

export async function updateUserWeight(userCode: string, weightPct: number) {
  const ref = doc(db, USERS, userCode);
  await updateDoc(ref, { weightPct, updatedAt: serverTimestamp() });
}

export async function upsertUser(u: {
  userCode: string;
  name: string;
  status: "online" | "offline";
  weightPct: number;
  active?: boolean;
}) {
  const ref = doc(db, "users", u.userCode);
  await setDoc(
    ref,
    {
      name: u.name,
      status: u.status,
      weightPct: u.weightPct,
      active: u.active ?? true,
      updatedAt: serverTimestamp(),
      createdAt: serverTimestamp(),
    },
    { merge: true }
  );
}

export async function deleteUser(userCode: string) {
  const ref = doc(db, "users", userCode);
  await deleteDoc(ref);
}
