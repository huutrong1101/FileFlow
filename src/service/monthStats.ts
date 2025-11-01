import {
  collection,
  doc,
  addDoc,
  getDocs,
  serverTimestamp,
  Timestamp,
  getDoc,
  query,
} from "firebase/firestore";
import { db } from "../lib/firebase";

const MONTH_STATS = "monthStats";

export function monthKey(date = new Date()) {
  // "YYYY-MM"
  return new Intl.DateTimeFormat("en-CA", {
    year: "numeric",
    month: "2-digit",
  }).format(date);
}

export async function logAssignmentToday(
  items: Array<{
    userCode: string;
    assignedCount?: number;
    assignedValue?: number;
    meta?: Record<string, unknown> | null;
  }>,
  forDate = new Date()
) {
  const key = monthKey(forDate);
  const monthDoc = doc(db, MONTH_STATS, key);
  const entriesRef = collection(monthDoc, "entries");

  for (const i of items) {
    await addDoc(entriesRef, {
      userCode: i.userCode,
      assignedCount: i.assignedCount ?? 0,
      assignedValue: i.assignedValue ?? 0,
      meta: i.meta ?? null,
      createdAt: serverTimestamp(),
      date: Timestamp.fromDate(forDate),
    });
  }
}

export async function getMonthEntries(yyyyMM: string) {
  const entriesRef = collection(doc(db, MONTH_STATS, yyyyMM), "entries");
  const snap = await getDocs(entriesRef);
  return snap.docs.map((d) => ({
    id: d.id,
    ...(d.data() as Record<string, unknown>),
  }));
}

// Tổng hợp cuối tháng theo userCode
type Entry = {
  userCode: string;
  assignedCount?: number;
  assignedValue?: number;
};

export function aggregateByUser(entries: Entry[]) {
  const map = new Map<
    string,
    { assignedCount: number; assignedValue: number }
  >();
  for (const e of entries) {
    const cur = map.get(e.userCode) ?? { assignedCount: 0, assignedValue: 0 };
    cur.assignedCount += e.assignedCount ?? 0;
    cur.assignedValue += e.assignedValue ?? 0;
    map.set(e.userCode, cur);
  }
  return Array.from(map, ([userCode, v]) => ({ userCode, ...v }));
}

/**
 * Optional tiện ích: cố gắng đọc theo thứ tự:
 * 1) entries (schema hiện tại) -> 2) days -> 3) logs (legacy)
 */
export async function getMonthAssignments(
  mKey: string
): Promise<Array<{ date: string; userCode: string; assignedCount: number }>> {
  // 0) Ưu tiên đọc entries (đúng với schema hiện tại)
  try {
    const entries = await getMonthEntries(mKey);
    if (entries.length) {
      return entries.map((e: Record<string, unknown>) => {
        const dateField = e?.date as { toDate?: () => Date } | undefined;
        return {
          date: dateField?.toDate
            ? dateField.toDate().toISOString().slice(0, 10)
            : "",
          userCode: String(e.userCode || "").toUpperCase(),
          assignedCount: Number(e.assignedCount || 0),
        };
      });
    }
  } catch {
    // Ignore error and try fallback methods
  }

  // 1) Thử subcollection 'days'
  try {
    const daysRef = collection(db, "monthStats", mKey, "days");
    const snap = await getDocs(query(daysRef));
    const out: Array<{
      date: string;
      userCode: string;
      assignedCount: number;
    }> = [];
    if (!snap.empty) {
      snap.forEach((d) => {
        const data = d.data() as Record<string, unknown>;
        const date = d.id; // giả định id = yyyy-MM-dd
        const items = Array.isArray(data?.items) ? data.items : [];
        items.forEach((it: Record<string, unknown>) => {
          out.push({
            date,
            userCode: String(it.userCode || "").toUpperCase(),
            assignedCount: Number(it.assignedCount || 0),
          });
        });
      });
      if (out.length) return out;
    }
  } catch {
    // Ignore error and try fallback methods
  }

  // 2) Fallback đọc field 'logs' của doc tháng
  try {
    const monthDoc = await getDoc(doc(db, "monthStats", mKey));
    if (monthDoc.exists()) {
      const data = monthDoc.data() as Record<string, unknown>;
      const logs = Array.isArray(data?.logs) ? data.logs : [];
      return logs.map((it: Record<string, unknown>) => ({
        date: String(it.date || ""),
        userCode: String(it.userCode || "").toUpperCase(),
        assignedCount: Number(it.assignedCount || 0),
      }));
    }
  } catch {
    // Ignore error and return empty array
  }

  return [];
}
