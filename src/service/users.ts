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
    warehouses?: string[]; // NEW
  }>
) {
  if (!Array.isArray(list) || list.length === 0) return;
  const batch = writeBatch(db);
  const now = serverTimestamp();
  for (const u of list) {
    const id = String(u.code || "").trim();
    if (!id) continue;
    const ref = doc(db, USERS, id);
    batch.set(
      ref,
      {
        name: u.name ?? "",
        status: u.online ? "online" : "offline",
        weightPct: Number.isFinite(u.weightPct) ? u.weightPct : 0,
        active: true,
        warehouses: Array.isArray(u.warehouses) ? u.warehouses : [], // NEW
        updatedAt: now,
        createdAt: now,
      },
      { merge: true }
    );
  }
  await batch.commit();
}

export async function listUsers(activeOnly = true) {
  const ref = collection(db, USERS);
  const q = activeOnly ? query(ref, where("active", "==", true)) : ref;

  const snap = await getDocs(q);
  const arr = snap.docs.map((d) => {
    const data = d.data() as Record<string, unknown>;
    return {
      code: d.id,
      name: String(data.name ?? ""),
      weightPct: Number(data.weightPct ?? 0),
      online: (data.status ?? "online") === "online",
      warehouses: Array.isArray(data.warehouses)
        ? (data.warehouses as string[])
        : [],
      order: Number(data.order ?? 0),
    } as import("../type").User;
  });

  // Nếu không dùng orderBy ở query, sort ở client:
  arr.sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  return arr;
}

export async function updateUserOnline(userCode: string, online: boolean) {
  const ref = doc(db, USERS, userCode);
  await updateDoc(ref, {
    status: online ? "online" : "offline",
    updatedAt: serverTimestamp(),
  });
}

export async function updateUserWarehouses(
  userCode: string,
  warehouses: string[]
) {
  const ref = doc(db, USERS, userCode);
  await updateDoc(ref, {
    warehouses: Array.isArray(warehouses) ? warehouses : [],
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
  warehouses?: string[]; // NEW
}) {
  const ref = doc(db, "users", u.userCode);
  await setDoc(
    ref,
    {
      name: u.name,
      status: u.status,
      weightPct: u.weightPct,
      active: u.active ?? true,
      warehouses: Array.isArray(u.warehouses) ? u.warehouses : [], // NEW
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

export async function saveUsersOrdering(codesInOrder: string[]) {
  if (!Array.isArray(codesInOrder) || codesInOrder.length === 0) return;
  const batch = writeBatch(db);
  const now = serverTimestamp();

  codesInOrder.forEach((code, idx) => {
    const ref = doc(db, USERS, code);
    batch.set(ref, { order: idx, updatedAt: now }, { merge: true });
  });

  await batch.commit();
}
