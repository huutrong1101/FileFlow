import { useEffect, useMemo, useState } from "react";
import {
  Layout,
  Typography,
  Table,
  Card,
  Row,
  Col,
  Space,
  Slider,
  InputNumber,
  Upload,
  message,
  Button,
  Tag,
  Statistic,
  theme,
  Switch,
  Form,
  Modal,
  Input,
  Popconfirm,
  Segmented,
  Tooltip,
  Badge,
  Avatar,
  Empty,
  Popover,
  Select,
} from "antd";
import type { ColumnsType } from "antd/es/table";
import {
  UploadOutlined,
  FileDoneOutlined,
  DownloadOutlined,
  DatabaseOutlined,
  CheckCircleOutlined,
  TeamOutlined,
  PlusOutlined,
  EditOutlined,
  DeleteOutlined,
  ReloadOutlined,
  SearchOutlined,
  ThunderboltOutlined,
  UserOutlined,
  SettingOutlined,
} from "@ant-design/icons";
import * as XLSX from "xlsx";
import {
  listUsers,
  updateUserOnline,
  updateUserWeight,
  upsertUsersBulk,
  upsertUser,
  deleteUser,
  saveUsersOrdering,
} from "./service/users";
import {
  logAssignmentToday,
  monthKey,
  reorderUsersByDeficit,
  settleDayAndGetAgg,
} from "./service/monthStats";
import type {
  AllocationSummary,
  AllocOpts,
  AssignmentItem,
  TaskRow,
  User,
} from "./type";
import { HolderOutlined } from "@ant-design/icons";
import { CSS } from "@dnd-kit/utilities";
import {
  DndContext,
  closestCenter,
  PointerSensor,
  useSensor,
  useSensors,
  type DragEndEvent,
} from "@dnd-kit/core";
import {
  SortableContext,
  useSortable,
  verticalListSortingStrategy,
  arrayMove,
} from "@dnd-kit/sortable";

const { Header, Content } = Layout;
const { Title, Text } = Typography;

import { Tour } from "antd";
import type { TourProps } from "antd";

import { QuestionCircleOutlined } from "@ant-design/icons";

import { useRef } from "react";

import StatsView from "./StatsView";

// ==================== UTILITIES ====================
const stripDiacritics = (s: string) =>
  s
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/đ/gi, "d");

const toKey = (s: string) => stripDiacritics(String(s)).toLowerCase().trim();

const normCode = (s: string) =>
  String(s ?? "")
    .normalize("NFKC")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

function initials(name: string) {
  const parts = name.trim().split(/\s+/);
  if (!parts.length) return "U";
  const pick = (parts[0]?.[0] || "") + (parts[parts.length - 1]?.[0] || "");
  return pick.toUpperCase();
}

// ==================== EXCEL PARSING ====================
async function parseUsersExcel(file: File): Promise<User[]> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows: TaskRow[] = XLSX.utils.sheet_to_json(ws, { defval: "" });

  if (!rows.length) return [];

  const headers = Object.keys(rows[0] || {});
  const finder = (hints: string[]) =>
    headers.find((h) => hints.some((t) => toKey(h).includes(t)));

  const codeKey =
    finder([
      "ma nv",
      "ma_nhan_vien",
      "employee code",
      "employee_code",
      "code",
      "ma nhan vien",
    ]) || headers[0];
  const nameKey =
    finder(["ten", "nhan vien", "ten nhan vien", "name"]) ||
    headers[1] ||
    codeKey;
  const ratioKey = finder([
    "ti le",
    "ty le",
    "percent",
    "ratio",
    "%",
    "ti le phan cong",
    "ty le phan cong",
  ]);
  const onKey = finder([
    "di lam",
    "online",
    "trang thai",
    "status",
    "off",
    "vang",
    "nghi",
  ]);
  // NEW: cột mã kho (có thể là danh sách, phân tách bằng dấu phẩy)
  const whKey = finder(["ma kho", "warehouse", "warehouses", "kho", "kho lam"]);

  const users: User[] = rows
    .map((r, i) => {
      const codeRaw = r[codeKey] ?? `U${i + 1}`;
      const code = normCode(String(codeRaw));
      const name = String(r[nameKey] ?? r[codeKey] ?? `U${i + 1}`).trim();
      const w = ratioKey ? Number(r[ratioKey]) : 100;
      const onlineVal = onKey != null ? String(r[onKey]).trim() : "true";
      const online =
        !/^\s*(off|0|false|nghi|vang)\s*$/i.test(onlineVal) && onlineVal !== "";

      // NEW: chuẩn hóa danh sách mã kho
      const warehouses: string[] = (() => {
        const raw = whKey ? String(r[whKey] ?? "").trim() : "";
        if (!raw) return [];
        return raw
          .split(/[,\s;]+/)
          .filter(Boolean)
          .map((s) => normCode(s))
          .filter((v, idx, arr) => arr.indexOf(v) === idx);
      })();

      return {
        code,
        name,
        weightPct: Number.isFinite(w) ? Math.max(0, w) : 100,
        online,
        warehouses, // NEW
      };
    })
    .filter((u) => u.code);

  return users;
}

function removeJunkColumns(rows: TaskRow[]) {
  const cleaned = rows.map((r) => {
    const o: TaskRow = {};
    Object.keys(r).forEach((k) => {
      if (/^__EMPTY/i.test(k)) return;
      if (!String(k).trim()) return;
      o[k] = r[k];
    });
    return o;
  });

  const headers = Object.keys(cleaned[0] || {});
  return { cleaned, headers };
}

async function parseTasksExcel(
  file: File
): Promise<{ rows: TaskRow[]; headers: string[] }> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rawRows: TaskRow[] = XLSX.utils.sheet_to_json(ws, { defval: "" });

  const { cleaned, headers } = removeJunkColumns(rawRows);
  return { rows: cleaned, headers };
}

// ==================== SORTING ====================
function detectGroupKeys(headers: string[]) {
  const find = (hints: string[]) =>
    headers.find((h) => hints.some((t) => toKey(h).includes(t)));

  const voucherKey = find([
    "ma chung tu",
    "so ct",
    "chung tu",
    "ct",
    "voucher",
  ]);
  const exportKey = find([
    "ma noi xuat",
    "noi xuat",
    "kho xuat",
    "store xuat",
    "export",
  ]);
  const receiveKey = find([
    "ma noi nhan",
    "noi nhan",
    "kho nhan",
    "store nhan",
    "receive",
  ]);
  const stKey = find(["ma st", "st"]);

  return { voucherKey, exportKey, receiveKey, stKey };
}

function sortRowsByGroupKeys(rows: TaskRow[], keys: string[]) {
  if (!rows.length || !keys.length) return rows.slice();

  const sorted = rows.slice().map((r, i) => ({ ...r, __idx__: i }));
  sorted.sort((a, b) => {
    for (const k of keys) {
      const av = String((a as Record<string, unknown>)[k] ?? "");
      const bv = String((b as Record<string, unknown>)[k] ?? "");
      if (av < bv) return -1;
      if (av > bv) return 1;
    }
    return a.__idx__ - b.__idx__;
  });

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  return sorted.map(({ __idx__, ...rest }) => rest);
}

// ==================== ALLOCATION (BLOCK-VOUCHER) ====================
function allocatePreferWarehousesTwoPhase(
  users: User[],
  rows: TaskRow[],
  arg3: string | null | AllocOpts,
  arg4?: string | null,
  arg5?: string | null
): { summary: AllocationSummary[]; assignments: AssignmentItem[] } {
  // ===== Opts =====
  const opts: AllocOpts =
    typeof arg3 === "object" && arg3 !== null
      ? {
          exportKey: arg3.exportKey ?? null,
          voucherKey: arg3.voucherKey ?? null,
          stKey: arg3.stKey ?? null,
        }
      : {
          exportKey: (arg3 as string | null) ?? null,
          voucherKey: arg4 ?? null,
          stKey: arg5 ?? null,
        };

  const { exportKey, voucherKey } = opts;

  // ===== Active users (force numeric weights) =====
  const activeRaw = users.map((u) => ({
    ...u,
    weightPct: Number(String(u.weightPct).replace("%", "").trim()) || 0,
    online: !!u.online,
    warehouses: Array.isArray(u.warehouses) ? u.warehouses : [],
  }));
  const active = activeRaw.filter((u) => u.online && u.weightPct > 0);

  if (!rows.length || !active.length) {
    return {
      summary: users.map((u) => ({
        userCode: u.code,
        userName: u.name,
        weightPct: u.weightPct,
        online: u.online,
        count: 0,
      })),
      assignments: [],
    };
  }

  // ===== Quota Hamilton (theo SỐ DÒNG) =====
  const N = rows.length;
  const totalW = active.reduce((s, u) => s + u.weightPct, 0);
  let quota: number[];
  if (totalW <= 0) {
    const even = Math.floor(N / active.length);
    const rem = N - even * active.length;
    quota = active.map(() => even);
    for (let i = 0; i < rem; i++) quota[i] += 1;
  } else {
    const shares = active.map((u) => (N * u.weightPct) / totalW);
    const base = shares.map((x) => Math.floor(x));
    const fracs = shares.map((x, i) => x - base[i]);
    const baseSum = base.reduce((s, x) => s + x, 0);
    const rem = N - baseSum;
    const idxs = fracs
      .map((f, i) => ({ i, f }))
      .sort((a, b) => (b.f === a.f ? a.i - b.i : b.f - a.f));
    quota = base.slice();
    for (let k = 0; k < rem; k++) quota[idxs[k].i] += 1;
  }
  const hardCap = quota.slice();

  // ===== GIỮ: giới hạn ngoại kho, escalation 2→3→4 =====
  const FOREIGN_LIMIT_BASE = 2;

  // ===== State =====
  const assignments: AssignmentItem[] = [];
  const assignedCount = new Array(active.length).fill(0);
  const deficitLive = quota.slice();
  const orderRank: number[] = active.map((_u, idx) => idx);
  const foreignExportSets: Array<Set<string>> = active.map(
    () => new Set<string>()
  );

  // ===== Helpers =====
  const normCodeLike = (v: unknown) => {
    const s = String(v ?? "")
      .normalize("NFKC")
      .replace(/\p{Diacritic}/gu, "")
      .trim()
      .toUpperCase()
      .replace(/^0+/, "");
    return s;
  };
  const expKeyOf = (expRaw: unknown) => {
    const s = normCodeLike(expRaw);
    return s || "__NO_EXPORT__";
  };
  const isOwner = (iUser: number, expRaw: unknown) => {
    const exp = normCodeLike(expRaw);
    if (!exp) return true; // không có mã xuất => ai cũng coi như owner
    const ws = Array.isArray(active[iUser].warehouses)
      ? active[iUser].warehouses
      : [];
    for (const w of ws) {
      if (exp === normCodeLike(w)) return true;
    }
    return false;
  };
  const foreignNewKeysFor = (iUser: number, expKeys: Set<string>) => {
    const newly = new Set<string>();
    for (const ek of expKeys) {
      if (ek === "__NO_EXPORT__") continue;
      if (!isOwner(iUser, ek) && !foreignExportSets[iUser].has(ek))
        newly.add(ek);
    }
    return newly;
  };
  const underCap = (i: number, blockSize: number) =>
    assignedCount[i] + blockSize <= hardCap[i];

  // pickBalanced: deficitLive desc, util(today) asc, orderRank asc
  const pickBalanced = (pool: number[]) => {
    let best = -1;
    let bestKey: number[] = [];
    for (const i of pool) {
      const deficit = deficitLive[i];
      const util = assignedCount[i] / Math.max(1, hardCap[i]);
      const tie = -orderRank[i];
      const key = [deficit, -util, tie];
      if (
        best === -1 ||
        key[0] > bestKey[0] ||
        (key[0] === bestKey[0] &&
          (key[1] > bestKey[1] ||
            (key[1] === bestKey[1] && key[2] > bestKey[2])))
      ) {
        best = i;
        bestKey = key;
      }
    }
    return best;
  };

  // ======= NHÓM THEO (VOUCHER, EXPORT) =======
  type Group = {
    vKey: string;
    eKey: string;
    idxs: number[];
    firstIdx: number;
    expKeys: Set<string>;
    owners: number[];
  };

  const groupsMap = new Map<string, Group>(); // key = vKey + '|' + eKey
  const getVKey = (i: number) => {
    if (!voucherKey) return `__ROW_${i}`;
    const vRaw = String(rows[i][voucherKey] ?? "").trim();
    return vRaw ? normCodeLike(vRaw) : `__ROW_${i}`;
  };

  for (let i = 0; i < rows.length; i++) {
    const v = getVKey(i);
    const e = exportKey ? expKeyOf(rows[i][exportKey]) : "__NO_EXPORT__";
    const key = `${v}|${e}`;
    if (!groupsMap.has(key)) {
      groupsMap.set(key, {
        vKey: v,
        eKey: e,
        idxs: [i],
        firstIdx: i,
        expKeys: new Set([e]),
        owners: [],
      });
    } else {
      groupsMap.get(key)!.idxs.push(i);
    }
  }
  const groups = Array.from(groupsMap.values());

  // Tính owners theo eKey (một exp duy nhất/nhóm)
  for (const g of groups) {
    const owners: number[] = [];
    for (let i = 0; i < active.length; i++)
      if (isOwner(i, g.eKey)) owners.push(i);
    g.owners = owners;
  }

  // Ưu tiên block lớn trước
  groups.sort(
    (a, b) => b.idxs.length - a.idxs.length || a.firstIdx - b.firstIdx
  );

  // ===== các filter dùng trong nhiều pha =====
  const filterCapAndForeign = (
    pool: number[],
    g: Group,
    blockSize: number,
    limitDistinct: number
  ) => {
    const out: number[] = [];
    for (const i of pool) {
      if (!underCap(i, blockSize)) continue;
      const addKeys = foreignNewKeysFor(i, g.expKeys);
      if (foreignExportSets[i].size + addKeys.size <= limitDistinct)
        out.push(i);
    }
    return out;
  };
  const filterForeignOnly = (
    pool: number[],
    g: Group,
    limitDistinct: number
  ) => {
    const out: number[] = [];
    for (const i of pool) {
      const addKeys = foreignNewKeysFor(i, g.expKeys);
      if (foreignExportSets[i].size + addKeys.size <= limitDistinct)
        out.push(i);
    }
    return out;
  };

  // ===== PHA CHÍNH: limit=2, ưu tiên OWNER tuyệt đối trước khi xét non-owner =====
  const leftovers: Group[] = [];
  for (const g of groups) {
    const blockSize = g.idxs.length;
    const allIdx = [...Array(active.length).keys()];

    // A1: OWNER underCap + foreign<=2
    let cand = filterCapAndForeign(g.owners, g, blockSize, FOREIGN_LIMIT_BASE);
    if (cand.length) {
      const owner = pickBalanced(cand);
      for (const rIdx of g.idxs)
        assignments.push({
          userCode: active[owner].code,
          userName: active[owner].name,
          taskIndex: rIdx,
        });
      assignedCount[owner] += blockSize;
      deficitLive[owner] -= blockSize;
      const addKeys = foreignNewKeysFor(owner, g.expKeys);
      addKeys.forEach((k) => foreignExportSets[owner].add(k));
      continue;
    }

    // A2: KHÔNG có owner còn cap → non-owner underCap + foreign<=2
    cand = filterCapAndForeign(
      allIdx.filter((i) => !g.owners.includes(i)),
      g,
      blockSize,
      FOREIGN_LIMIT_BASE
    );
    if (cand.length) {
      const pick = pickBalanced(cand);
      for (const rIdx of g.idxs)
        assignments.push({
          userCode: active[pick].code,
          userName: active[pick].name,
          taskIndex: rIdx,
        });
      assignedCount[pick] += blockSize;
      deficitLive[pick] -= blockSize;
      const addKeys = foreignNewKeysFor(pick, g.expKeys);
      addKeys.forEach((k) => foreignExportSets[pick].add(k));
      continue;
    }

    // để lại escalation
    leftovers.push(g);
  }

  // ===== LEFTOVERS (limit=2) – overshoot có kiểm soát: 100% có thể overshoot 1 block tổng
  const overshootBudgetBase = active.map((u) => (u.weightPct >= 100 ? 1 : 0));
  const canOverWith = (budgetArr: number[], iUser: number, blockSize: number) =>
    budgetArr[iUser] >= blockSize;
  const weightSortedIdx = [...Array(active.length).keys()].sort(
    (i, j) =>
      active[j].weightPct - active[i].weightPct ||
      deficitLive[j] - deficitLive[i] ||
      orderRank[i] - orderRank[j]
  );

  const stillLeft: Group[] = [];
  for (const g of leftovers) {
    const blockSize = g.idxs.length;

    // L1: OWNER underCap (limit=2) — thử lại lần nữa
    let pool = filterCapAndForeign(g.owners, g, blockSize, FOREIGN_LIMIT_BASE);
    if (pool.length) {
      const owner = pickBalanced(pool);
      for (const rIdx of g.idxs)
        assignments.push({
          userCode: active[owner].code,
          userName: active[owner].name,
          taskIndex: rIdx,
        });
      assignedCount[owner] += blockSize;
      deficitLive[owner] -= blockSize;
      const addKeys = foreignNewKeysFor(owner, g.expKeys);
      addKeys.forEach((k) => foreignExportSets[owner].add(k));
      continue;
    }

    // L2: OWNER overshoot có kiểm soát (limit=2): 50% không overshoot, ≥100% dùng budget 1 block
    pool = filterForeignOnly(g.owners, g, FOREIGN_LIMIT_BASE).concat(
      filterForeignOnly(
        weightSortedIdx.filter((i) => !g.owners.includes(i)),
        g,
        FOREIGN_LIMIT_BASE
      )
    );
    let chosen = -1;
    for (const i of pool) {
      const needOver = Math.max(0, assignedCount[i] + blockSize - hardCap[i]);
      if (needOver === 0) {
        chosen = i;
        break;
      }
      if (
        active[i].weightPct >= 100 &&
        canOverWith(overshootBudgetBase, i, blockSize)
      ) {
        chosen = i;
        break;
      }
    }
    if (chosen !== -1) {
      const needOver = Math.max(
        0,
        assignedCount[chosen] + blockSize - hardCap[chosen]
      );
      if (needOver > 0) overshootBudgetBase[chosen] -= blockSize;
      for (const rIdx of g.idxs)
        assignments.push({
          userCode: active[chosen].code,
          userName: active[chosen].name,
          taskIndex: rIdx,
        });
      assignedCount[chosen] += blockSize;
      deficitLive[chosen] -= blockSize;
      const addKeys = foreignNewKeysFor(chosen, g.expKeys);
      addKeys.forEach((k) => foreignExportSets[chosen].add(k));
      continue;
    }

    // L3: không xong ở limit=2 -> để escalation
    stillLeft.push(g);
  }

  // ===== Escalation: thử limit=3 rồi 4 (vẫn ưu tiên OWNER trước)
  // Không tăng thêm ngân sách overshoot cho nhóm 100% (khóa ở 1 block)
  const escalateAssign = (groupsIn: Group[], limitDistinct: number) => {
    if (!groupsIn.length) return [];
    const budget = overshootBudgetBase.slice(); // KHÔNG cộng thêm

    const notDone: Group[] = [];
    for (const g of groupsIn) {
      const blockSize = g.idxs.length;

      // E1: OWNER underCap + foreign<=limit
      let pool = filterCapAndForeign(g.owners, g, blockSize, limitDistinct);
      if (pool.length) {
        const owner = pickBalanced(pool);
        for (const rIdx of g.idxs)
          assignments.push({
            userCode: active[owner].code,
            userName: active[owner].name,
            taskIndex: rIdx,
          });
        assignedCount[owner] += blockSize;
        deficitLive[owner] -= blockSize;
        const addKeys = foreignNewKeysFor(owner, g.expKeys);
        addKeys.forEach((k) => foreignExportSets[owner].add(k));
        continue;
      }

      // E2: OWNER overshoot có kiểm soát (ưu tiên), rồi mới đến non-owner underCap
      pool = filterForeignOnly(g.owners, g, limitDistinct);
      let chosen = -1;
      for (const i of pool) {
        const needOver = Math.max(0, assignedCount[i] + blockSize - hardCap[i]);
        if (needOver === 0) {
          chosen = i;
          break;
        }
        if (active[i].weightPct >= 100 && budget[i] >= blockSize) {
          chosen = i;
          break;
        }
      }
      if (chosen !== -1) {
        const needOver = Math.max(
          0,
          assignedCount[chosen] + blockSize - hardCap[chosen]
        );
        if (needOver > 0) budget[chosen] -= blockSize;
        for (const rIdx of g.idxs)
          assignments.push({
            userCode: active[chosen].code,
            userName: active[chosen].name,
            taskIndex: rIdx,
          });
        assignedCount[chosen] += blockSize;
        deficitLive[chosen] -= blockSize;
        const addKeys = foreignNewKeysFor(chosen, g.expKeys);
        addKeys.forEach((k) => foreignExportSets[chosen].add(k));
        continue;
      }

      // E3: non-owner underCap + foreign<=limit
      pool = filterCapAndForeign(
        [...Array(active.length).keys()].filter((i) => !g.owners.includes(i)),
        g,
        blockSize,
        limitDistinct
      );
      if (pool.length) {
        const pick = pickBalanced(pool);
        for (const rIdx of g.idxs)
          assignments.push({
            userCode: active[pick].code,
            userName: active[pick].name,
            taskIndex: rIdx,
          });
        assignedCount[pick] += blockSize;
        deficitLive[pick] -= blockSize;
        const addKeys = foreignNewKeysFor(pick, g.expKeys);
        addKeys.forEach((k) => foreignExportSets[pick].add(k));
        continue;
      }

      // E4: non-owner overshoot có kiểm soát (>=100% + còn budget)
      pool = filterForeignOnly(
        [...Array(active.length).keys()].filter((i) => !g.owners.includes(i)),
        g,
        limitDistinct
      );
      chosen = -1;
      for (const i of pool) {
        const needOver = Math.max(0, assignedCount[i] + blockSize - hardCap[i]);
        if (needOver === 0) {
          chosen = i;
          break;
        }
        if (active[i].weightPct >= 100 && budget[i] >= blockSize) {
          chosen = i;
          break;
        }
      }
      if (chosen !== -1) {
        const needOver = Math.max(
          0,
          assignedCount[chosen] + blockSize - hardCap[chosen]
        );
        if (needOver > 0) budget[chosen] -= blockSize;
        for (const rIdx of g.idxs)
          assignments.push({
            userCode: active[chosen].code,
            userName: active[chosen].name,
            taskIndex: rIdx,
          });
        assignedCount[chosen] += blockSize;
        deficitLive[chosen] -= blockSize;
        const addKeys = foreignNewKeysFor(chosen, g.expKeys);
        addKeys.forEach((k) => foreignExportSets[chosen].add(k));
      } else {
        notDone.push(g);
      }
    }
    return notDone;
  };

  let pending = stillLeft;
  pending = escalateAssign(pending, 3);
  pending = escalateAssign(pending, 4);

  // ===== L5: Force-assign — ưu tiên OWNER nhưng chọn cân bằng hơn
  if (pending.length) {
    for (const g of pending) {
      const blockSize = g.idxs.length;

      // ownerFirst: ưu tiên chủ kho trước, rồi mới đến non-owner
      const ownerFirst = [
        ...g.owners,
        ...[...Array(active.length).keys()].filter(
          (i) => !g.owners.includes(i)
        ),
      ];

      const candidates = ownerFirst.slice();
      const scored = candidates.map((i) => ({
        i,
        deficit: deficitLive[i],
        util: assignedCount[i] / Math.max(1, hardCap[i]),
        isHundred: active[i].weightPct >= 100,
      }));

      // 1) Thử người <100% còn underCap
      let pool = scored.filter(
        (s) => !s.isHundred && assignedCount[s.i] + blockSize <= hardCap[s.i]
      );

      if (!pool.length) {
        // 2) Cho 100% gánh nếu bắt buộc, ưu tiên thiếu nhiều & util thấp
        pool = scored;
      }

      pool.sort(
        (a, b) =>
          b.deficit - a.deficit ||
          a.util - b.util ||
          orderRank[a.i] - orderRank[b.i]
      );

      const chosen = pool[0]?.i ?? candidates[0];

      for (const rIdx of g.idxs)
        assignments.push({
          userCode: active[chosen].code,
          userName: active[chosen].name,
          taskIndex: rIdx,
        });
      assignedCount[chosen] += blockSize;
      deficitLive[chosen] -= blockSize;
      const addKeys = foreignNewKeysFor(chosen, g.expKeys);
      addKeys.forEach((k) => foreignExportSets[chosen].add(k));
    }
  }

  // ===== Tổng kết =====
  const countMap = new Map<string, number>();
  for (let i = 0; i < active.length; i++)
    countMap.set(active[i].code, assignedCount[i]);

  const summary: AllocationSummary[] = users.map((u) => ({
    userCode: u.code,
    userName: u.name,
    weightPct: u.weightPct,
    online: u.online,
    count: countMap.get(u.code) ?? 0,
  }));

  return { summary, assignments };
}

// ==================== EXCEL EXPORT ====================
function exportExcelWithAssignments(
  users: User[],
  taskRows: TaskRow[],
  taskHeaders: string[],
  summary: AllocationSummary[],
  opts: {
    exportKey: string | null;
    voucherKey: string | null;
    stKey: string | null;
  }
) {
  if (!summary.length) {
    message.warning("Chưa có kết quả để xuất.");
    return;
  }

  const { assignments } = allocatePreferWarehousesTwoPhase(
    users,
    taskRows,
    opts
  );

  const assignedMap = new Map<number, { code: string; name: string }>();
  assignments.forEach((a) =>
    assignedMap.set(a.taskIndex, { code: a.userCode, name: a.userName })
  );

  const merged = taskRows.map((r, i) => {
    const assigned = assignedMap.get(i) || { code: "", name: "" };
    return {
      ma_nv_phan_cong: assigned.code,
      ten_nv_phan_cong: assigned.name,
      ...r,
    };
  });

  const finalHeaders = ["ma_nv_phan_cong", "ten_nv_phan_cong", ...taskHeaders];
  const rowsForSheet = merged.map((r) => {
    const o: Record<string, unknown> = {};
    for (const h of finalHeaders)
      o[h] = (r as Record<string, unknown>)[h] ?? "";
    return o;
  });

  const ws = XLSX.utils.aoa_to_sheet([finalHeaders]);
  XLSX.utils.sheet_add_json(ws, rowsForSheet, {
    header: finalHeaders,
    skipHeader: true,
    origin: "A2",
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "PhanCong_CongViec");

  const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `phan_cong_${new Date()
    .toISOString()
    .slice(0, 19)
    .replace(/[:T]/g, "-")}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ==================== MAIN COMPONENT ====================
export default function App() {
  const [users, setUsers] = useState<User[]>([]);
  const [taskRows, setTaskRows] = useState<TaskRow[]>([]);
  const [taskHeaders, setTaskHeaders] = useState<string[]>([]);
  const [summary, setSummary] = useState<AllocationSummary[]>([]);
  const [view, setView] = useState<"assign" | "stats">("assign");
  const [exportKey, setExportKey] = useState<string | null>(null);
  const [showTour, setShowTour] = useState(false);
  const [voucherKey, setVoucherKey] = useState<string | null>(null);
  const [stKey, setStKey] = useState<string | null>(null);
  const [testMode, setTestMode] = useState<boolean>(false);

  const refUpload = useRef<HTMLDivElement | null>(null);
  const refAllocateBtn = useRef<HTMLButtonElement | null>(null);
  const refUsersTable = useRef<HTMLDivElement | null>(null);
  const refDownloadBtn = useRef<HTMLButtonElement | null>(null);
  const refSegmented = useRef<HTMLDivElement | null>(null);
  const refSummaryCard = useRef<HTMLDivElement | null>(null);

  // Loading & modal states
  const [loadingUsers, setLoadingUsers] = useState<boolean>(false);
  const [savingUser, setSavingUser] = useState<boolean>(false);
  const [modalOpen, setModalOpen] = useState<boolean>(false);
  const [editing, setEditing] = useState<User | null>(null);
  const [form] = Form.useForm<{
    code: string;
    name: string;
    weightPct: number;
    online: boolean;
    warehouses?: string[];
  }>();

  // Tìm kiếm & lọc
  const [userQuery, setUserQuery] = useState<string>("");
  const [statusFilter, setStatusFilter] = useState<
    "all" | "online" | "offline"
  >("all");

  const {
    token: {
      colorBgContainer,
      colorBorder,
      borderRadius,
      colorTextSecondary,
      colorPrimary,
    },
  } = theme.useToken();

  const filteredUsers = useMemo(() => {
    const q = toKey(userQuery);
    return users.filter((u) => {
      if (statusFilter === "online" && !u.online) return false;
      if (statusFilter === "offline" && u.online) return false;
      if (!q) return true;
      const hay = `${toKey(u.code)} ${toKey(u.name)}`;
      return hay.includes(q);
    });
  }, [users, userQuery, statusFilter]);

  // Load Users from Firestore
  const reloadUsers = async () => {
    setLoadingUsers(true);
    try {
      const data = await listUsers(true);
      console.log({ data });
      setUsers(data);
    } catch (e: unknown) {
      const error = e as Error;
      message.error(`Lỗi tải users từ Firestore: ${error?.message || e}`);
    } finally {
      setLoadingUsers(false);
    }
  };

  useEffect(() => {
    reloadUsers();
  }, []);

  useEffect(() => {
    if (!users.length) {
      message.info("Vui lòng thêm nhân viên ở phần Danh sách NV.");
    }
  }, [users]);

  // ====== COMPACT WEIGHT POPOVER ======
  const WeightPopover: React.FC<{ u: User }> = ({ u }) => {
    const [temp, setTemp] = useState<number>(u.weightPct);
    return (
      <div style={{ width: 260 }}>
        <Space direction="vertical" style={{ width: "100%" }}>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <Text type="secondary">Điều chỉnh tỉ lệ cho</Text>
            <Text strong>{u.code}</Text>
          </div>
          <Slider
            min={0}
            max={300}
            step={5}
            value={temp}
            onChange={(v) => setTemp(Number(v))}
          />
          <InputNumber
            min={0}
            max={300}
            value={temp}
            onChange={(v) => setTemp(Number(v ?? 0))}
            style={{ width: "100%" }}
          />
          <Space style={{ justifyContent: "flex-end", width: "100%" }}>
            <Button onClick={() => setTemp(u.weightPct)} size="small">
              Reset
            </Button>
            <Button
              type="primary"
              size="small"
              onClick={async () => {
                const v = Number(temp);
                setUsers((prev) =>
                  prev.map((x) =>
                    x.code === u.code ? { ...x, weightPct: v } : x
                  )
                );
                try {
                  await updateUserWeight(u.code, v);
                  message.success("Đã cập nhật tỉ lệ.");
                } catch (e: unknown) {
                  const error = e as Error;
                  message.error(
                    `Cập nhật tỉ lệ thất bại: ${error?.message || e}`
                  );
                }
              }}
            >
              Lưu
            </Button>
          </Space>
        </Space>
      </div>
    );
  };

  // ==================== TABLE COLUMNS ====================
  const userCols: ColumnsType<User> = [
    {
      title: "",
      key: "drag",
      width: 36,
      className: "drag-visible",
      render: () => (
        <HolderOutlined style={{ cursor: "grab", color: "#999" }} />
      ),
    },
    {
      title: "Mã NV",
      dataIndex: "code",
      key: "code",
      width: 120,
      fixed: "left",
      render: (text) => <Text strong>{text}</Text>,
      sorter: (a, b) => a.code.localeCompare(b.code),
      showSorterTooltip: false,
    },
    {
      title: "Tên",
      dataIndex: "name",
      key: "name",
      width: 240,
      render: (_text, u) => (
        <Space size={8}>
          <Badge status={u.online ? "success" : "default"} dot>
            <Avatar size="small" icon={<UserOutlined />}>
              {initials(u.name)}
            </Avatar>
          </Badge>
          <Tooltip title={u.name}>
            <Text style={{ maxWidth: 180 }} ellipsis>
              {u.name}
            </Text>
          </Tooltip>
        </Space>
      ),
      sorter: (a, b) => a.name.localeCompare(b.name),
      showSorterTooltip: false,
    },
    {
      title: "Trạng thái",
      dataIndex: "online",
      key: "online",
      width: 120,
      align: "center",
      render: (_, u) => (
        <Switch
          checked={u.online}
          onChange={async (checked) => {
            setUsers((prev) =>
              prev.map((x) =>
                x.code === u.code ? { ...x, online: checked } : x
              )
            );
            try {
              await updateUserOnline(u.code, checked);
            } catch (e: unknown) {
              const error = e as Error;
              message.error(`Cập nhật online thất bại: ${error?.message || e}`);
              setUsers((prev) =>
                prev.map((x) =>
                  x.code === u.code ? { ...x, online: !checked } : x
                )
              );
            }
          }}
        />
      ),
      filters: [
        { text: "Online", value: "online" },
        { text: "Offline", value: "offline" },
      ],
      onFilter: (v, u) => (v === "online" ? u.online : !u.online),
    },
    {
      title: "Tỉ lệ",
      key: "weightPct",
      width: 140,
      align: "center",
      render: (_, u) => {
        const tone =
          u.weightPct >= 150
            ? "magenta"
            : u.weightPct >= 100
            ? "blue"
            : "default";
        const label =
          u.weightPct >= 150 ? "Cao" : u.weightPct >= 100 ? "Chuẩn" : "Thấp";
        return (
          <Popover
            trigger="click"
            placement="left"
            content={<WeightPopover u={u} />}
          >
            <Tag
              color={tone}
              style={{ cursor: "pointer" }}
              icon={<SettingOutlined />}
            >
              <Text strong>{u.weightPct}%</Text>{" "}
              <Text type="secondary">· {label}</Text>
            </Tag>
          </Popover>
        );
      },
      sorter: (a, b) => a.weightPct - b.weightPct,
      showSorterTooltip: false,
    },
    {
      title: "Mã kho",
      key: "warehouses",
      dataIndex: "warehouses",
      width: 220,
      render: (w: string[] | undefined) =>
        w && w.length ? (
          <Space size={[4, 4]} wrap>
            {w.map((x) => (
              <Tag key={x}>{x}</Tag>
            ))}
          </Space>
        ) : (
          <Text type="secondary">—</Text>
        ),
    },
    {
      title: "Thao tác",
      key: "actions",
      fixed: "right",
      width: 140,
      render: (_, u) => (
        <Space size="small">
          <Tooltip title="Sửa">
            <Button
              size="small"
              icon={<EditOutlined />}
              onClick={() => {
                setEditing(u);
                form.setFieldsValue({
                  code: u.code,
                  name: u.name,
                  weightPct: u.weightPct,
                  online: u.online,
                  warehouses: u.warehouses ?? [],
                });
                setModalOpen(true);
              }}
            />
          </Tooltip>
          <Popconfirm
            title={
              <>
                Xóa nhân viên <Text strong>{u.code}</Text>?
              </>
            }
            okText="Xóa"
            cancelText="Hủy"
            onConfirm={async () => {
              try {
                await deleteUser(u.code);
                message.success(`Đã xóa ${u.code}`);
                reloadUsers();
              } catch (e: unknown) {
                const error = e as Error;
                message.error(`Xóa thất bại: ${error?.message || e}`);
              }
            }}
          >
            <Button size="small" danger icon={<DeleteOutlined />} />
          </Popconfirm>
        </Space>
      ),
    },
  ];

  const sumCols: ColumnsType<AllocationSummary> = [
    {
      title: "Mã NV",
      dataIndex: "userCode",
      key: "userCode",
      width: 110,
      render: (text) => <Text strong>{text}</Text>,
    },
    {
      title: "Tên",
      dataIndex: "userName",
      key: "userName",
      render: (text) => <Text>{text}</Text>,
    },
    {
      title: "Trạng thái",
      dataIndex: "online",
      key: "online",
      align: "center",
      render: (v) =>
        v ? (
          <Tag color="green" icon={<CheckCircleOutlined />}>
            ON
          </Tag>
        ) : (
          <Tag color="default">OFF</Tag>
        ),
      width: 90,
    },
    {
      title: "Tỉ lệ (%)",
      dataIndex: "weightPct",
      key: "weightPct",
      align: "center",
      render: (v: number) => (
        <Tag color="blue">
          <Text strong>{v}%</Text>
        </Tag>
      ),
      width: 120,
    },
    {
      title: "Số việc",
      dataIndex: "count",
      key: "count",
      align: "center",
      render: (v: number) => (
        <Text strong style={{ color: colorPrimary, fontSize: 16 }}>
          {v}
        </Text>
      ),
      width: 100,
    },
  ];

  // ==================== EVENT HANDLERS ====================
  const handleUserFile = async (file: File) => {
    try {
      const parsed = await parseUsersExcel(file);
      if (!parsed.length) {
        message.warning("File nhân viên trống hoặc không đọc được.");
        return Upload.LIST_IGNORE;
      }
      await upsertUsersBulk(parsed);
      await reloadUsers();
      message.success(`Đã nạp ${parsed.length} nhân viên vào Firestore.`);
    } catch (e: unknown) {
      const error = e as Error;
      message.error(`Lỗi import: ${error?.message || e}`);
    }
    return Upload.LIST_IGNORE;
  };

  const handleTasksFile = async (file: File) => {
    try {
      const { rows, headers } = await parseTasksExcel(file);
      const { voucherKey, exportKey, receiveKey, stKey } =
        detectGroupKeys(headers);

      // Thứ tự nhóm: Mã chứng từ -> Mã nơi xuất -> Mã nơi nhận
      const keyOrder = [voucherKey, exportKey, receiveKey].filter(
        Boolean
      ) as string[];
      const sortedRows = sortRowsByGroupKeys(rows, keyOrder);

      setTaskRows(sortedRows);
      setTaskHeaders(headers);
      setVoucherKey(voucherKey ?? null);
      setExportKey(exportKey ?? null);
      setStKey(stKey ?? null);

      message.success(`Đã nạp & sắp xếp ${sortedRows.length} dòng công việc.`);
    } catch (e: unknown) {
      const error = e as Error;
      message.error(`Lỗi đọc file công việc: ${error?.message || e}`);
    }
    return Upload.LIST_IGNORE;
  };

  const handleAllocate = async () => {
    if (!taskRows.length)
      return message.warning("Chưa có công việc để phân bổ!");
    if (!users.length) return message.warning("Chưa có danh sách nhân viên!");

    const { summary: resultSummary } = allocatePreferWarehousesTwoPhase(
      users,
      taskRows,
      {
        exportKey: exportKey ?? null,
        voucherKey: voucherKey ?? null,
        stKey: stKey ?? null,
      }
    );

    setSummary(resultSummary);
    if (!resultSummary.length) {
      return message.warning("Không có user online/weight > 0 để phân bổ.");
    }

    // NEW — Test mode: không ghi Firestore, không reorder users
    if (testMode) {
      message.success(
        "Đã phân bổ (Chế độ thử). Không lưu Firestore & không đổi thứ tự NV."
      );
      return;
    }

    // === Nhánh thật: ghi log + settle + reorder như cũ ===
    try {
      await logAssignmentToday(
        resultSummary.map((s) => ({
          userCode: s.userCode,
          assignedCount: s.count,
          meta: { source: "ui-allocate", at: new Date().toISOString() },
        }))
      );

      const agg = await settleDayAndGetAgg({
        forDate: new Date(),
        users,
        summary: resultSummary.map((s) => ({
          userCode: s.userCode,
          count: s.count,
        })),
      });

      const reordered = reorderUsersByDeficit(users, agg);
      setUsers(reordered);
      await saveUsersOrdering(reordered.map((u) => u.code));

      message.success(
        `Đã phân bổ & lưu vào Firestore (tháng ${monthKey()}). Đã sắp xếp ưu tiên ngày mai.`
      );
    } catch (e: unknown) {
      const error = e as Error;
      message.error(`Lưu phân công/aggregate thất bại: ${error?.message || e}`);
    }
  };

  // Save user (create / update)
  const submitUserForm = async () => {
    try {
      const vals = await form.validateFields();
      setSavingUser(true);
      await upsertUser({
        userCode: normCode(vals.code),
        name: vals.name,
        status: vals.online ? "online" : "offline",
        weightPct: Number(vals.weightPct ?? 0),
        active: true,
        warehouses: Array.isArray(vals.warehouses) ? vals.warehouses : [], // NEW
      });
      message.success(
        editing ? "Đã cập nhật nhân viên." : "Đã thêm nhân viên."
      );
      setModalOpen(false);
      setEditing(null);
      form.resetFields();
      await reloadUsers();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } catch (e: any) {
      if (e?.errorFields) return;
      message.error(`Lưu nhân viên thất bại: ${e?.message || e}`);
    } finally {
      setSavingUser(false);
    }
  };

  const SortableRow: React.FC<
    { id: string } & React.HTMLAttributes<HTMLTableRowElement>
  > = ({ id, style, className, children, ...rest }) => {
    const { attributes, listeners, setNodeRef, transform, transition } =
      useSortable({ id });

    const mergedStyle: React.CSSProperties = {
      ...style,
      transform: CSS.Transform.toString(transform),
      transition,
    };

    return (
      <tr
        ref={setNodeRef}
        className={className}
        style={mergedStyle}
        {...attributes}
        {...listeners}
        {...rest}
      >
        {children}
      </tr>
    );
  };

  // Tạo sensors cho dnd-kit
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 5 } })
  );

  // Hàm đổi thứ tự mảng users theo drag result
  const onDragEndUsers = async (event: DragEndEvent) => {
    const { active, over } = event;
    if (!over || active.id === over.id) return;

    const oldIndex = users.findIndex((u) => u.code === String(active.id));
    const newIndex = users.findIndex((u) => u.code === String(over.id));
    if (oldIndex === -1 || newIndex === -1) return;

    const newUsers = arrayMove(users, oldIndex, newIndex);
    setUsers(newUsers);

    // Lưu order lên Firestore
    try {
      await saveUsersOrdering(newUsers.map((u) => u.code));
      // (tuỳ chọn) message.success("Đã lưu thứ tự.");
    } catch (e: unknown) {
      const error = e as Error;
      message.error(`Lưu thứ tự thất bại: ${error?.message || e}`);
    }
  };

  // AntD Table components override
  const components = {
    body: {
      row: (
        props: {
          "data-row-key"?: string;
        } & React.HTMLAttributes<HTMLTableRowElement>
      ) => {
        const record: User | undefined = props["data-row-key"]
          ? users.find((u) => normCode(u.code) === props["data-row-key"])
          : undefined;
        const id = record ? record.code : Math.random().toString();
        return <SortableRow id={id} {...props} />;
      },
    },
  };
  useEffect(() => {
    const FLAG = "APP_TOUR_DONE_V1";
    if (!localStorage.getItem(FLAG)) {
      setShowTour(true);
      localStorage.setItem(FLAG, "1");
    }
  }, []);

  const steps: TourProps["steps"] = [
    {
      title: "Chọn file công việc",
      description: "Kéo/thả hoặc chọn file Excel.",
      target: () => refUpload.current!,
    },
    {
      title: "Phân bổ tự động & Lưu",
      description:
        "Nhấn 'Phân bổ' để phân công theo tỉ lệ và ưu tiên mã kho. Kết quả sẽ được ngay lập tức lưu trữ",
      target: () => refAllocateBtn.current!,
    },
    {
      title: "Danh sách nhân viên",
      description:
        "Kéo-thả để sắp thứ tự ưu tiên, chỉnh trạng thái Online/Offline và Tỉ lệ (Popover). Có thể thực hiện thêm, xóa, sửa thông tin user trong danh sách.",
      target: () => refUsersTable.current!,
    },
    {
      title: "Xuất báo cáo",
      description:
        "Sau khi phân bổ xong, nhấn để tải Excel gộp kết quả (mã NV & tên NV đã gán).",
      target: () => refDownloadBtn.current!,
    },
    {
      title: "Tóm tắt phân bổ",
      description:
        "Bảng này hiển thị KẾT QUẢ phân bổ. 'Số việc' là số dòng công việc được gán cho từng nhân viên.",
      target: () => refSummaryCard.current!,
    },
    {
      title: "Đổi chế độ xem",
      description:
        "Chuyển giữa màn 'Phân công' và 'Thống kê' để xem biểu đồ phân bổ theo tháng.",
      target: () => refSegmented.current!,
    },
  ];

  // ==================== RENDER ====================
  return (
    <>
      <Layout style={{ minHeight: "100vh" }}>
        <Header
          style={{
            background: colorBgContainer,
            borderBottom: `1px solid ${colorBorder}`,
            padding: "0 24px",
            boxShadow: "0 2px 8px rgba(0, 0, 0, 0.04)",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
          }}
        >
          <Space size={12} align="center">
            <ThunderboltOutlined
              style={{ color: colorPrimary, fontSize: 20 }}
            />
            <Title
              level={3}
              style={{ margin: 0, fontSize: 20, fontWeight: 700 }}
            >
              Quản lý phân công công việc
            </Title>
          </Space>

          <Space align="center" size="middle" wrap>
            <Tooltip title="Xem hướng dẫn nhanh">
              <Button
                type="text"
                icon={<QuestionCircleOutlined />}
                onClick={() => setShowTour(true)}
                style={{ paddingInline: 8 }}
              >
                Xem hướng dẫn
              </Button>
            </Tooltip>

            <Segmented
              value={view}
              onChange={(v) => setView(v as "assign" | "stats")}
              options={[
                { label: "Phân công", value: "assign" },
                { label: "Thống kê", value: "stats" },
              ]}
            />
          </Space>
        </Header>

        <Content style={{ padding: "24px", background: "#fafafa" }}>
          {view === "assign" ? (
            <>
              <Row gutter={[16, 16]}>
                <Col xs={24} lg={8}>
                  <Card
                    style={{
                      borderRadius,
                      boxShadow: "0 2px 8px rgba(0, 0, 0, 0.06)",
                      border: `1px solid ${colorBorder}`,
                    }}
                    bodyStyle={{ padding: "24px" }}
                    title={
                      <Space>
                        <FileDoneOutlined style={{ color: colorPrimary }} />
                        <span style={{ fontWeight: 600 }}>
                          Tải file công việc
                        </span>
                      </Space>
                    }
                  >
                    <Space
                      direction="vertical"
                      style={{ width: "100%", gap: 16 }}
                    >
                      <div ref={refUpload}>
                        <Upload.Dragger
                          maxCount={1}
                          accept=".xlsx,.xls"
                          beforeUpload={handleTasksFile}
                          showUploadList={true}
                          style={{
                            borderRadius,
                            background: `${colorPrimary}08`,
                            borderColor: `${colorPrimary}30`,
                          }}
                        >
                          <p
                            className="ant-upload-drag-icon"
                            style={{ color: colorPrimary }}
                          >
                            <UploadOutlined style={{ fontSize: 32 }} />
                          </p>
                          <p
                            className="ant-upload-text"
                            style={{ fontSize: 14, fontWeight: 500 }}
                          >
                            Kéo/thả hoặc bấm để chọn file
                          </p>
                        </Upload.Dragger>
                      </div>

                      <Row gutter={12}>
                        <Col flex="auto">
                          <div
                            style={{
                              background: `${colorPrimary}08`,
                              padding: "12px 16px",
                              borderRadius,
                              border: `1px solid ${colorPrimary}30`,
                            }}
                          >
                            <Statistic
                              title={
                                <Text
                                  style={{
                                    fontSize: 12,
                                    color: colorTextSecondary,
                                  }}
                                >
                                  Số dòng công việc
                                </Text>
                              }
                              value={taskRows.length}
                              valueStyle={{ color: colorPrimary, fontSize: 22 }}
                            />
                          </div>
                        </Col>
                      </Row>

                      {/* Action row: ưu tiên rõ ràng */}
                      <Row gutter={[8, 8]}>
                        {import.meta.env.DEV && (
                          <Col span={24}>
                            <Space
                              align="center"
                              style={{
                                width: "100%",
                                justifyContent: "space-between",
                                padding: "8px 12px",
                                borderRadius,
                                border: `1px dashed ${colorPrimary}40`,
                                background: `${colorPrimary}06`,
                              }}
                            >
                              <Space>
                                <Switch
                                  checked={testMode}
                                  onChange={setTestMode}
                                />
                                <Text strong>Chế độ thử (không lưu)</Text>
                              </Space>
                              <Text type="secondary" style={{ fontSize: 12 }}>
                                Chỉ tính toán & hiển thị kết quả; KHÔNG ghi
                                Firestore và KHÔNG đổi thứ tự NV
                              </Text>
                            </Space>
                          </Col>
                        )}

                        <Col span={24}>
                          <Button
                            ref={refAllocateBtn}
                            type="primary"
                            onClick={handleAllocate}
                            disabled={!taskRows.length || !users.length}
                            block
                            size="large"
                            style={{ borderRadius, fontWeight: 600 }}
                          >
                            {testMode ? "Phân bổ (Chế độ thử)" : "Phân bổ"}
                          </Button>
                        </Col>

                        <Col span={24}>
                          <Button
                            ref={refDownloadBtn}
                            icon={<DownloadOutlined />}
                            onClick={() =>
                              exportExcelWithAssignments(
                                users,
                                taskRows,
                                taskHeaders,
                                summary,
                                {
                                  exportKey: exportKey ?? null,
                                  voucherKey: voucherKey ?? null,
                                  stKey: stKey ?? null,
                                }
                              )
                            }
                            disabled={!summary.length}
                            size="large"
                            block
                            style={{ borderRadius, fontWeight: 600 }}
                          >
                            Tải Excel (đã phân bổ)
                          </Button>
                        </Col>
                      </Row>
                    </Space>
                  </Card>
                </Col>

                {/* Users panel */}
                <Col xs={24} lg={16}>
                  <Card
                    size="small"
                    title={
                      <Space size={6} align="center">
                        <TeamOutlined style={{ color: colorPrimary }} />
                        <span style={{ color: colorPrimary, fontWeight: 600 }}>
                          Danh sách NV
                        </span>
                      </Space>
                    }
                    style={{
                      borderRadius,
                      boxShadow: "0 2px 8px rgba(0, 0, 0, 0.06)",
                      border: `1px solid ${colorBorder}`,
                    }}
                    extra={
                      <Space wrap>
                        <Input
                          allowClear
                          size="small"
                          prefix={<SearchOutlined />}
                          placeholder="Tìm theo mã / tên"
                          style={{ width: 220 }}
                          value={userQuery}
                          onChange={(e) => setUserQuery(e.target.value)}
                        />
                        <Segmented
                          size="small"
                          value={statusFilter}
                          onChange={(v) =>
                            setStatusFilter(v as "all" | "online" | "offline")
                          }
                          options={[
                            { label: "Tất cả", value: "all" },
                            { label: "Online", value: "online" },
                            { label: "Offline", value: "offline" },
                          ]}
                        />
                        <Tooltip title="Tải lại">
                          <Button
                            size="small"
                            icon={<ReloadOutlined />}
                            onClick={reloadUsers}
                          />
                        </Tooltip>
                        <Button
                          size="small"
                          type="primary"
                          icon={<PlusOutlined />}
                          onClick={() => {
                            setEditing(null);
                            form.resetFields();
                            form.setFieldsValue({
                              code: "",
                              name: "",
                              weightPct: 100,
                              online: true,
                              warehouses: [],
                            });
                            setModalOpen(true);
                          }}
                        >
                          Thêm
                        </Button>
                        <Upload
                          beforeUpload={handleUserFile}
                          maxCount={1}
                          accept=".xlsx,.xls"
                        >
                          <Button
                            size="small"
                            icon={<DatabaseOutlined />}
                            style={{ fontSize: 12 }}
                          >
                            Nhập
                          </Button>
                        </Upload>
                      </Space>
                    }
                    bodyStyle={{ padding: "16px" }}
                  >
                    <DndContext
                      sensors={sensors}
                      collisionDetection={closestCenter}
                      onDragEnd={onDragEndUsers}
                    >
                      <SortableContext
                        items={filteredUsers.map((u) => u.code)}
                        strategy={verticalListSortingStrategy}
                      >
                        <div ref={refUsersTable}>
                          <Table<User>
                            rowKey={(u) => normCode(u.code)}
                            dataSource={filteredUsers}
                            columns={userCols}
                            size="middle"
                            pagination={false}
                            scroll={{ y: 360, x: true }}
                            loading={loadingUsers}
                            bordered
                            components={components}
                            locale={{
                              emptyText: (
                                <Empty
                                  description={
                                    <span>
                                      Chưa có nhân viên. Bấm{" "}
                                      <Text strong>Thêm</Text> hoặc{" "}
                                      <Text strong>Nhập</Text> từ Excel.
                                    </span>
                                  }
                                />
                              ),
                            }}
                          />
                        </div>
                      </SortableContext>
                    </DndContext>
                  </Card>
                </Col>
              </Row>

              <Row gutter={[16, 16]} style={{ marginTop: 16 }}>
                <Col xs={24}>
                  <div ref={refSummaryCard}>
                    <Card
                      title={
                        <Space size={6}>
                          <TeamOutlined style={{ color: colorPrimary }} />
                          <span style={{ fontWeight: 600 }}>
                            Tóm tắt phân bổ
                          </span>
                        </Space>
                      }
                      style={{
                        borderRadius,
                        boxShadow: "0 2px 8px rgba(0, 0, 0, 0.06)",
                        border: `1px solid ${colorBorder}`,
                      }}
                      bodyStyle={{ padding: 0 }}
                    >
                      <Table<AllocationSummary>
                        rowKey="userCode"
                        size="small"
                        dataSource={summary}
                        columns={sumCols}
                        pagination={{ pageSize: 10 }}
                        scroll={{ x: true }}
                        bordered
                      />
                    </Card>
                  </div>
                </Col>
              </Row>
            </>
          ) : (
            <StatsView initialUsers={users} />
          )}
        </Content>

        <Modal
          title={editing ? `Sửa nhân viên: ${editing.code}` : "Thêm nhân viên"}
          open={modalOpen}
          onCancel={() => {
            setModalOpen(false);
            setEditing(null);
            form.resetFields();
          }}
          onOk={submitUserForm}
          okText={editing ? "Cập nhật" : "Thêm"}
          confirmLoading={savingUser}
          destroyOnClose
        >
          <Form form={form} layout="vertical">
            <Form.Item
              label="Mã NV"
              name="code"
              rules={[
                { required: true, message: "Vui lòng nhập mã NV" },
                {
                  pattern: /^[A-Za-z0-9_-]+$/,
                  message: "Chỉ chữ/số/gạch dưới/gạch ngang",
                },
              ]}
            >
              <Input placeholder="VD: U001" disabled={!!editing} />
            </Form.Item>
            <Form.Item
              label="Tên"
              name="name"
              rules={[{ required: true, message: "Vui lòng nhập tên" }]}
            >
              <Input placeholder="VD: Nguyễn Văn A" />
            </Form.Item>
            <Form.Item label="Tỉ lệ (%)" name="weightPct" initialValue={100}>
              <InputNumber min={0} max={300} style={{ width: "100%" }} />
            </Form.Item>
            <Form.Item
              label="Online"
              name="online"
              valuePropName="checked"
              initialValue={true}
            >
              <Switch />
            </Form.Item>
            <Form.Item label="Mã kho (có thể nhập nhiều)" name="warehouses">
              <Select
                mode="tags"
                tokenSeparators={[",", ";", " "]}
                placeholder="VD: KHO001, KHO002"
                // Chuẩn hóa input về normCode khi thay đổi
                onChange={(vals) => {
                  const normalized = (vals as string[])
                    .map((s) => normCode(String(s)))
                    .filter(Boolean)
                    .filter((v, idx, arr) => arr.indexOf(v) === idx);
                  form.setFieldsValue({ warehouses: normalized });
                }}
              />
            </Form.Item>
          </Form>
        </Modal>
      </Layout>
      <Tour open={showTour} onClose={() => setShowTour(false)} steps={steps} />
    </>
  );
}
