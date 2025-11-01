import React, { useEffect, useMemo, useState } from "react";
import {
  Card,
  Table,
  Space,
  Typography,
  DatePicker,
  Button,
  Tag,
  theme,
  message,
} from "antd";
import type { ColumnsType } from "antd/es/table";
import { BarChartOutlined, DownloadOutlined } from "@ant-design/icons";
import dayjs, { Dayjs } from "dayjs";
import * as XLSX from "xlsx";
import { listUsers } from "./service/users";
import {
  monthKey as makeMonthKey,
  getMonthEntries,
  aggregateByUser,
} from "./service/monthStats";
import type { User } from "./type";

const { Text } = Typography;
const { MonthPicker } = DatePicker as any;

type StatRow = {
  userCode: string;
  userName: string;
  total: number;
};

type Props = {
  // Optional: nếu App đã có users, truyền xuống để đỡ gọi lại.
  initialUsers?: User[];
};

const StatsView: React.FC<Props> = ({ initialUsers }) => {
  const {
    token: { colorBorder, borderRadius, colorPrimary },
  } = theme.useToken();

  const [month, setMonth] = useState<Dayjs>(dayjs()); // default: tháng hiện tại
  const [loading, setLoading] = useState(false);
  const [users, setUsers] = useState<User[]>(initialUsers ?? []);
  const [rows, setRows] = useState<StatRow[]>([]);

  // map code -> name
  const nameByCode = useMemo(() => {
    const map = new Map<string, string>();
    users.forEach((u) => map.set(u.code, u.name));
    return map;
  }, [users]);

  const reloadUsersIfNeeded = async () => {
    if (initialUsers && initialUsers.length) return;
    try {
      const data = await listUsers(true);
      setUsers(data);
    } catch (e: any) {
      message.error(`Lỗi tải users: ${e?.message || e}`);
    }
  };

  const loadMonth = async (m: Dayjs) => {
    setLoading(true);
    try {
      const mKey = makeMonthKey
        ? makeMonthKey(m.toDate())
        : m.format("YYYY-MM");

      // Đọc entries của tháng rồi tổng hợp theo user
      const entries = await getMonthEntries(mKey);
      const agg = aggregateByUser(entries as any); // [{userCode, assignedCount, assignedValue}]

      const data: StatRow[] = agg
        .map((it) => ({
          userCode: it.userCode,
          userName: nameByCode.get(it.userCode) || it.userCode,
          total: Number(it.assignedCount || 0),
        }))
        .sort((a, b) => b.total - a.total);

      setRows(data);
    } catch (e: any) {
      message.error(`Lỗi tải thống kê tháng: ${e?.message || e}`);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    reloadUsersIfNeeded();
  }, []);

  useEffect(() => {
    loadMonth(month);
  }, [month]);

  const cols: ColumnsType<StatRow> = [
    {
      title: "Mã NV",
      dataIndex: "userCode",
      key: "userCode",
      width: 140,
      render: (t) => <Text strong>{t}</Text>,
      sorter: (a, b) => a.userCode.localeCompare(b.userCode),
      showSorterTooltip: false,
    },
    {
      title: "Tên",
      dataIndex: "userName",
      key: "userName",
      render: (t) => <Text>{t}</Text>,
      sorter: (a, b) => a.userName.localeCompare(b.userName),
      showSorterTooltip: false,
    },
    {
      title: "Tổng số việc (tháng)",
      dataIndex: "total",
      key: "total",
      align: "center",
      width: 180,
      render: (v: number) => (
        <Tag color="blue">
          <Text strong style={{ fontSize: 16 }}>
            {v}
          </Text>
        </Tag>
      ),
      sorter: (a, b) => a.total - b.total,
      showSorterTooltip: false,
    },
  ];

  const exportExcel = () => {
    if (!rows.length) {
      return message.warning("Chưa có dữ liệu để xuất.");
    }
    const sheetData = rows.map((r) => ({
      ma_nhan_vien: r.userCode,
      ten_nhan_vien: r.userName,
      tong_so_viec: r.total,
    }));
    const ws = XLSX.utils.json_to_sheet(sheetData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ThongKeThang");
    const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    const blob = new Blob([out], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `thong_ke_${month.format("YYYY-MM")}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <Card
      title={
        <Space>
          <BarChartOutlined style={{ color: colorPrimary }} />
          <span style={{ fontWeight: 600 }}>Thống kê phân công theo tháng</span>
        </Space>
      }
      style={{
        borderRadius,
        boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
        border: `1px solid ${colorBorder}`,
      }}
      extra={
        <Space>
          <MonthPicker
            allowClear={false}
            value={month}
            onChange={(v: Dayjs) => v && setMonth(v)}
            format="MM-YYYY"
          />
          <Button icon={<DownloadOutlined />} onClick={exportExcel}>
            Xuất Excel
          </Button>
        </Space>
      }
    >
      <Table<StatRow>
        rowKey="userCode"
        dataSource={rows}
        columns={cols}
        loading={loading}
        pagination={{ pageSize: 12, showSizeChanger: true }}
        bordered
      />
    </Card>
  );
};

export default StatsView;
