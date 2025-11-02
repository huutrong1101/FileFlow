import {
  collection,
  doc,
  addDoc,
  getDocs,
  serverTimestamp,
  Timestamp,
  getDoc,
  setDoc,
  query,
} from "firebase/firestore";
import { db } from "../lib/firebase";
import type { MonthAgg } from "../type";

const MONTH_STATS = "monthStats";

function ymd(date = new Date()) {
  return new Intl.DateTimeFormat("en-CA", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date); // 'YYYY-MM-DD'
}

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

/**
 * Tính KỲ VỌNG TRONG NGÀY theo weightPct NGÀY ĐÓ.
 * - online=true và weightPct>0 mới tham gia
 * - expected ~ (weight / tổng weight active) * tổng công việc ngày
 */
export function expectedTodayMap(
  users: Array<{
    userCode?: string;
    code?: string;
    online?: boolean;
    weightPct?: number;
  }>,
  totalToday: number
) {
  const actives = users.filter(
    (u) => (u.online ?? true) && (u.weightPct ?? 0) > 0
  );
  const W = actives.reduce((s, u) => s + (u.weightPct ?? 0), 0) || 1;

  const map: Record<string, number> = {};
  for (const u of actives) {
    const uRecord = u as Record<string, unknown>;
    const code = String(uRecord.userCode ?? uRecord.code ?? "").toUpperCase();
    if (!code) continue;
    map[code] = (totalToday * (u.weightPct ?? 0)) / W;
  }
  return map;
}

// Lưu/đọc aggregate tháng: monthStats/{YYYY-MM}/agg/v1
export async function getMonthAggregate(
  yyyyMM: string
): Promise<MonthAgg | null> {
  const ref = doc(db, "monthStats", yyyyMM, "agg", "v1");
  const snap = await getDoc(ref);
  if (!snap.exists()) return null;
  return snap.data() as MonthAgg;
}

export async function saveMonthAggregate(yyyyMM: string, agg: MonthAgg) {
  const ref = doc(db, "monthStats", yyyyMM, "agg", "v1");
  await setDoc(ref, { ...agg, version: "1" }, { merge: true });
}

/**
 * settleDayAndGetAgg:
 * - Nhận summary PHÂN BỔ HÔM NAY + trạng thái users HÔM NAY (online/weightPct)
 * - Cộng dồn expected (tính theo weightPct hôm nay) & actual
 * - Cập nhật deficit = expectedCum - actualCum
 * - Lưu aggregate tháng -> trả về agg mới
 */
export async function settleDayAndGetAgg(params: {
  forDate?: Date;
  users: Array<{ code: string; online: boolean; weightPct: number }>;
  summary: Array<{ userCode: string; count: number }>;
}) {
  const date = params.forDate ?? new Date();
  const day = ymd(date); // 'YYYY-MM-DD'
  const mKey = monthKey(date); // 'YYYY-MM'
  const users = params.users ?? [];
  const summary = params.summary ?? [];

  const totalToday = summary.reduce((s, x) => s + (x.count || 0), 0);
  const expMap = expectedTodayMap(
    users.map((u) => ({
      code: u.code,
      online: u.online,
      weightPct: u.weightPct,
    })),
    totalToday
  );

  // Lấy agg hiện có hoặc khởi tạo
  const agg: MonthAgg = (await getMonthAggregate(mKey)) ?? {
    expectedCum: {},
    actualCum: {},
    deficit: {},
    lastServedAt: {},
    version: "1",
  };

  // Cộng dồn expected theo users HÔM NAY (kể cả ai 0 để minh bạch)
  for (const u of users) {
    const code = String(u.code).toUpperCase();
    if (!agg.expectedCum[code]) agg.expectedCum[code] = 0;
    agg.expectedCum[code] += expMap[code] ?? 0;
  }

  // Cộng dồn actual theo summary HÔM NAY
  for (const s of summary) {
    const code = String(s.userCode).toUpperCase();
    if (!agg.actualCum[code]) agg.actualCum[code] = 0;
    agg.actualCum[code] += s.count || 0;
    agg.lastServedAt[code] = day;
  }

  // Tính lại deficit
  const allCodes = new Set([
    ...Object.keys(agg.expectedCum),
    ...Object.keys(agg.actualCum),
  ]);

  for (const code of allCodes) {
    const e = agg.expectedCum[code] ?? 0;
    const a = agg.actualCum[code] ?? 0;
    agg.deficit[code] = e - a; // >0 là còn thiếu so với kỳ vọng tháng
  }

  await saveMonthAggregate(mKey, agg);
  return agg;
}

/**
 * Sắp xếp user cho NGÀY SAU dựa theo deficit tháng:
 * - Thiếu nhiều đứng trước
 * - Online đứng trước Offline
 * - Ai phục vụ lâu rồi (lastServedAt cũ) đứng trước
 * - Cuối cùng ổn định theo code
 */
export function reorderUsersByDeficit<
  T extends { code: string; online: boolean }
>(users: T[], agg: MonthAgg) {
  const arr = [...users];
  arr.sort((a, b) => {
    const ca = String(a.code).toUpperCase();
    const cb = String(b.code).toUpperCase();
    const da = agg.deficit[ca] ?? 0;
    const db = agg.deficit[cb] ?? 0;
    if (db !== da) return db - da; // thiếu nhiều trước
    if (a.online !== b.online) return a.online ? -1 : 1; // online trước
    const la = agg.lastServedAt[ca] ?? "";
    const lb = agg.lastServedAt[cb] ?? "";
    if (la !== lb) return la.localeCompare(lb); // phục vụ cũ trước
    return ca.localeCompare(cb);
  });
  return arr;
}
