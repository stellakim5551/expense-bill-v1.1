"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";
import {
  CalendarDays,
  Download,
  FileSpreadsheet,
  FileText,
  Loader2,
  Pencil,
  Plus,
  Save,
  Trash2,
  X,
} from "lucide-react";
import * as XLSX from "xlsx";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";

type Expense = {
  id: number;
  date: string;
  description: string;
  amount: number;
  created_at: string;
};

type ExpenseForm = {
  date: string;
  description: string;
  amount: string;
};

type ExportRange = {
  startDate: string;
  endDate: string;
};

type QuickRange = "today" | "last3" | "last7" | "month" | "custom";

function formatDate(date: Date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function today() {
  return formatDate(new Date());
}

function addDays(date: Date, days: number) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function currentMonth() {
  return today().slice(0, 7);
}

function monthRange() {
  const now = new Date();
  const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  return {
    startDate: formatDate(firstDay),
    endDate: formatDate(lastDay),
  };
}

function rangeFor(type: QuickRange): ExportRange {
  const now = new Date();

  if (type === "today") {
    return { startDate: formatDate(now), endDate: formatDate(now) };
  }

  if (type === "last3") {
    return { startDate: formatDate(addDays(now, -2)), endDate: formatDate(now) };
  }

  if (type === "last7") {
    return { startDate: formatDate(addDays(now, -6)), endDate: formatDate(now) };
  }

  return monthRange();
}

function money(value: number) {
  return new Intl.NumberFormat("zh-CN", {
    style: "currency",
    currency: "CNY",
    maximumFractionDigits: 2,
  }).format(value);
}

const STORAGE_KEY = "smart-expense-tool-expenses";

const emptyForm = (): ExpenseForm => ({
  date: today(),
  description: "",
  amount: "",
});

const quickRanges: Array<{ key: QuickRange; label: string }> = [
  { key: "today", label: "今天" },
  { key: "last3", label: "最近3天" },
  { key: "last7", label: "最近7天" },
  { key: "month", label: "本月" },
  { key: "custom", label: "自定义" },
];

function parseStoredExpenses(value: string | null): Expense[] {
  if (!value) {
    return [];
  }

  try {
    const result = JSON.parse(value) as Expense[];
    return Array.isArray(result)
      ? result.map((item) => ({
          id: Number(item.id),
          date: String(item.date),
          description: String(item.description),
          amount: Number(item.amount),
          created_at: String(item.created_at),
        }))
      : [];
  } catch {
    return [];
  }
}

function loadExpensesFromStorage(): Expense[] {
  return parseStoredExpenses(window.localStorage.getItem(STORAGE_KEY));
}

function saveExpensesToStorage(expenses: Expense[]) {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(expenses));
}

function calculateTotal(expenses: Expense[]) {
  return expenses.reduce((sum, expense) => sum + expense.amount, 0);
}

function formatExportFilename(prefix: string) {
  return `${prefix}_${new Date().toISOString().slice(0, 10)}.xlsx`;
}

export default function HomePage() {
  const [expenses, setExpenses] = useState<Expense[]>([]);
  const [monthlyTotal, setMonthlyTotal] = useState(0);
  const [form, setForm] = useState<ExpenseForm>(emptyForm);
  const [exportRange, setExportRange] = useState<ExportRange>(rangeFor("today"));
  const [activeRange, setActiveRange] = useState<QuickRange>("today");
  const [editingId, setEditingId] = useState<number | null>(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState("");

  const recentExpenses = useMemo(() => expenses.slice(0, 5), [expenses]);

  useEffect(() => {
    const stored = loadExpensesFromStorage();
    setExpenses(stored);
    setMonthlyTotal(calculateTotal(stored));
    setLoading(false);
  }, []);

  function resetForm() {
    setForm(emptyForm());
    setEditingId(null);
  }

  function editExpense(expense: Expense) {
    setEditingId(expense.id);
    setForm({
      date: expense.date,
      description: expense.description,
      amount: String(expense.amount),
    });
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function chooseQuickRange(type: QuickRange) {
    setActiveRange(type);

    if (type !== "custom") {
      setExportRange(rangeFor(type));
    }
  }

  function updateExportRange(value: Partial<ExportRange>) {
    setActiveRange("custom");
    setExportRange((current) => ({ ...current, ...value }));
  }

  async function submitExpense(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setMessage("");
    setSaving(true);

    const payload = {
      date: form.date,
      description: form.description.trim(),
      amount: Number(form.amount),
    };

    if (!payload.description) {
      setMessage("用途不能为空。");
      setSaving(false);
      return;
    }

    if (!payload.date || Number.isNaN(payload.amount) || payload.amount <= 0) {
      setMessage("请输入正确的日期和金额。");
      setSaving(false);
      return;
    }

    const currentExpenses = [...expenses];
    if (editingId !== null) {
      const index = currentExpenses.findIndex((item) => item.id === editingId);
      if (index >= 0) {
        currentExpenses[index] = {
          ...currentExpenses[index],
          date: payload.date,
          description: payload.description,
          amount: payload.amount,
        };
      }
      setMessage("记录已更新。");
    } else {
      const nextId = currentExpenses.length > 0 ? Math.max(...currentExpenses.map((item) => item.id)) + 1 : 1;
      currentExpenses.unshift({
        id: nextId,
        date: payload.date,
        description: payload.description,
        amount: payload.amount,
        created_at: new Date().toISOString(),
      });
      setMessage("记录已保存。");
    }

    saveExpensesToStorage(currentExpenses);
    setExpenses(currentExpenses);
    setMonthlyTotal(calculateTotal(currentExpenses));
    setSaving(false);
    resetForm();
  }

  function removeExpense(expense: Expense) {
    if (!window.confirm(`确认删除「${expense.description}」吗？`)) {
      return;
    }

    const nextExpenses = expenses.filter((item) => item.id !== expense.id);
    saveExpensesToStorage(nextExpenses);
    setExpenses(nextExpenses);
    setMonthlyTotal(calculateTotal(nextExpenses));
    setMessage("记录已删除。");

    if (editingId === expense.id) {
      resetForm();
    }
  }

  function exportToExcel() {
    const filtered = expenses.filter((expense) => {
      return expense.date >= exportRange.startDate && expense.date <= exportRange.endDate;
    });

    if (filtered.length === 0) {
      setMessage("该区间暂无记录，无法导出。");
      return;
    }

    const rows = filtered.map((item) => ({
      日期: item.date,
      用途: item.description,
      金额: item.amount,
    }));

    const worksheet = XLSX.utils.json_to_sheet(rows, { header: ["日期", "用途", "金额"] });
    // 自动计算列宽：根据内容长度（中文乘以 1.2）设置 wch
    const headers = ["日期", "用途", "金额"];
    const cols = headers.map((h, colIndex) => {
      let max = String(h).length;
      for (const r of rows) {
        const val = Object.values(r)[colIndex];
        const len = String(val ?? "").length;
        if (len > max) max = len;
      }
      const wch = Math.ceil(max * 1.2) + 2;
      return { wch };
    });
    worksheet["!cols"] = cols;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "报销单");

    const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = formatExportFilename("报销单");
    anchor.click();
    URL.revokeObjectURL(url);

    setMessage("Excel 已生成，下载将在稍后开始。注意：仅包含所选区间的数据。");
  }

  async function exportToPdf() {
    setMessage("正在生成 PDF...");
    try {
      // 动态加载 html2canvas
      if (!(window as any).html2canvas) {
        await new Promise<void>((resolve, reject) => {
          const s = document.createElement("script");
          s.src = "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js";
          s.onload = () => resolve();
          s.onerror = () => reject(new Error("加载 html2canvas 失败"));
          document.head.appendChild(s);
        });
      }

      const filtered = expenses.filter((expense) => {
        return expense.date >= exportRange.startDate && expense.date <= exportRange.endDate;
      });

      if (filtered.length === 0) {
        setMessage("该区间暂无记录，无法导出。");
        return;
      }

      // 构建隐藏表格
      const container = document.createElement("div");
      container.style.position = "fixed";
      container.style.left = "-9999px";
      container.style.top = "0";
      container.style.background = "white";
      container.style.padding = "16px";
      container.style.width = "800px";
      container.style.fontSize = "14px";
      container.style.fontFamily = "Arial, sans-serif";

      const title = document.createElement("h2");
      title.innerText = `报销单 ${exportRange.startDate} - ${exportRange.endDate}`;
      title.style.fontSize = "18px";
      title.style.marginBottom = "12px";
      title.style.color = "#333";
      container.appendChild(title);

      const table = document.createElement("table");
      table.style.borderCollapse = "collapse";
      table.style.width = "100%";
      table.style.marginBottom = "12px";

      const thead = document.createElement("thead");
      const headerRow = document.createElement("tr");
      ["日期", "用途", "金额"].forEach((h) => {
        const th = document.createElement("th");
        th.innerText = h;
        th.style.border = "1px solid #999";
        th.style.padding = "8px";
        th.style.textAlign = "left";
        th.style.background = "#f0f0f0";
        th.style.fontWeight = "600";
        headerRow.appendChild(th);
      });
      thead.appendChild(headerRow);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      let total = 0;
      filtered.forEach((exp) => {
        const tr = document.createElement("tr");
        [exp.date, exp.description, String(exp.amount)].forEach((c, i) => {
          const td = document.createElement("td");
          td.innerText = String(c);
          td.style.border = "1px solid #ccc";
          td.style.padding = "8px";
          if (i === 2) td.style.textAlign = "right";
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
        total += exp.amount;
      });

      // 总计行
      const totalRow = document.createElement("tr");
      const td1 = document.createElement("td");
      td1.colSpan = 2;
      td1.innerText = "总计";
      td1.style.border = "1px solid #999";
      td1.style.padding = "8px";
      td1.style.fontWeight = "600";
      const td2 = document.createElement("td");
      td2.innerText = String(total);
      td2.style.border = "1px solid #999";
      td2.style.padding = "8px";
      td2.style.fontWeight = "600";
      td2.style.textAlign = "right";
      totalRow.appendChild(td1);
      totalRow.appendChild(td2);
      tbody.appendChild(totalRow);

      table.appendChild(tbody);
      container.appendChild(table);
      document.body.appendChild(container);

      const html2canvas = (window as any).html2canvas;
      const canvas = await html2canvas(container, { scale: 2, logging: false });
      const { default: jsPDF } = await import("jspdf");
      const pdf = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4" });
      const pw = pdf.internal.pageSize.getWidth();
      const ph = pdf.internal.pageSize.getHeight();
      const imgWidth = pw - 40;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      const imgData = canvas.toDataURL("image/png");

      if (imgHeight <= ph - 40) {
        pdf.addImage(imgData, "PNG", 20, 20, imgWidth, imgHeight);
      } else {
        let y = 0;
        const ratio = canvas.width / imgWidth;
        const pageHeightInCanvas = (ph - 40) * ratio;
        while (y < canvas.height) {
          const slice = document.createElement("canvas");
          slice.width = canvas.width;
          slice.height = Math.min(pageHeightInCanvas, canvas.height - y);
          const ctx = slice.getContext("2d");
          ctx?.drawImage(canvas, 0, y, canvas.width, slice.height, 0, 0, canvas.width, slice.height);
          const sliceData = slice.toDataURL("image/png");
          const sliceHeight = (slice.height * imgWidth) / slice.width;
          pdf.addImage(sliceData, "PNG", 20, 20, imgWidth, sliceHeight);
          y += slice.height;
          if (y < canvas.height) pdf.addPage();
        }
      }

      const filename = `报销单_${exportRange.startDate}_${exportRange.endDate}.pdf`;
      pdf.save(filename);
      setMessage("PDF 已生成并下载。");
      document.body.removeChild(container);
    } catch (err) {
      console.error(err);
      setMessage("生成 PDF 出错，请尝试使用 Excel 导出。");
    }
  }

  function exportFile(type: "excel" | "pdf") {
    if (!exportRange.startDate || !exportRange.endDate) {
      setMessage("请先选择开始日期和结束日期。");
      return;
    }

    if (exportRange.startDate > exportRange.endDate) {
      setMessage("开始日期不能晚于结束日期。");
      return;
    }

    if (type === "excel") {
      exportToExcel();
      return;
    }

    exportToPdf();
  }

  return (
    <main className="min-h-screen bg-slate-50 text-slate-950">
      <div className="mx-auto flex min-h-screen w-full max-w-3xl flex-col px-4 py-5 sm:px-6">
        <header className="mb-5">
          <p className="text-sm font-medium text-emerald-700">智能报销助手 V1.1</p>
          <div className="mt-2 flex items-end justify-between gap-4">
            <div>
              <h1 className="text-2xl font-semibold tracking-normal">采购报销记录</h1>
              <p className="mt-1 text-sm text-slate-500">本月 {currentMonth()}</p>
            </div>
            <div className="text-right">
              <p className="text-xs text-slate-500">本月总支出</p>
              <p className="mt-1 text-2xl font-semibold text-emerald-700">
                {money(monthlyTotal)}
              </p>
            </div>
          </div>
        </header>

        <section className="rounded-lg border border-slate-200 bg-white p-4 shadow-sm">
          <div className="mb-4 flex items-center justify-between">
            <h2 className="text-base font-semibold">
              {editingId ? "编辑记录" : "新增记录"}
            </h2>
            {editingId ? (
              <Button type="button" variant="ghost" size="sm" onClick={resetForm}>
                <X />
                取消
              </Button>
            ) : (
              <Plus className="size-5 text-emerald-700" />
            )}
          </div>

          <form className="grid gap-3" onSubmit={submitExpense}>
            <label className="grid gap-1.5">
              <span className="text-sm font-medium text-slate-700">日期</span>
              <Input
                required
                type="date"
                value={form.date}
                className="h-11"
                onChange={(event) =>
                  setForm((value) => ({ ...value, date: event.target.value }))
                }
              />
            </label>

            <label className="grid gap-1.5">
              <span className="text-sm font-medium text-slate-700">用途</span>
              <Input
                required
                value={form.description}
                placeholder="例如：购买包装袋"
                className="h-11"
                onChange={(event) =>
                  setForm((value) => ({
                    ...value,
                    description: event.target.value,
                  }))
                }
              />
            </label>

            <label className="grid gap-1.5">
              <span className="text-sm font-medium text-slate-700">金额</span>
              <Input
                required
                min="0.01"
                step="0.01"
                type="number"
                inputMode="decimal"
                value={form.amount}
                placeholder="例如：120"
                className="h-11"
                onChange={(event) =>
                  setForm((value) => ({ ...value, amount: event.target.value }))
                }
              />
            </label>

            {message ? (
              <p className="rounded-md bg-slate-100 px-3 py-2 text-sm text-slate-700">
                {message}
              </p>
            ) : null}

            <Button className="mt-1 h-11 bg-emerald-700 text-white hover:bg-emerald-800">
              {saving ? <Loader2 className="animate-spin" /> : <Save />}
              {editingId ? "保存修改" : "保存记录"}
            </Button>
          </form>
        </section>

        <section className="mt-4 rounded-lg border border-slate-200 bg-white p-4 shadow-sm">
          <div className="mb-3 flex items-center gap-2">
            <CalendarDays className="size-5 text-emerald-700" />
            <h2 className="text-base font-semibold">导出报销单</h2>
          </div>

          <div className="mb-3 grid grid-cols-3 gap-2 sm:grid-cols-5">
            {quickRanges.map((range) => (
              <Button
                key={range.key}
                type="button"
                variant={activeRange === range.key ? "default" : "outline"}
                className={
                  activeRange === range.key
                    ? "h-9 bg-emerald-700 text-white hover:bg-emerald-800"
                    : "h-9 bg-white"
                }
                onClick={() => chooseQuickRange(range.key)}
              >
                {range.label}
              </Button>
            ))}
          </div>

          <div className="grid gap-3 sm:grid-cols-2">
            <label className="grid gap-1.5">
              <span className="text-sm font-medium text-slate-700">开始日期</span>
              <Input
                required
                type="date"
                value={exportRange.startDate}
                className="h-11"
                onChange={(event) =>
                  updateExportRange({ startDate: event.target.value })
                }
              />
            </label>
            <label className="grid gap-1.5">
              <span className="text-sm font-medium text-slate-700">结束日期</span>
              <Input
                required
                type="date"
                value={exportRange.endDate}
                className="h-11"
                onChange={(event) =>
                  updateExportRange({ endDate: event.target.value })
                }
              />
            </label>
          </div>

          <div className="mt-3 grid grid-cols-2 gap-3">
            <Button
              type="button"
              variant="outline"
              className="h-11 bg-white"
              onClick={() => exportFile("excel")}
            >
              <FileSpreadsheet />
              导出Excel
            </Button>
            <Button
              type="button"
              variant="outline"
              className="h-11 bg-white"
              onClick={() => exportFile("pdf")}
            >
              <FileText />
              导出PDF
            </Button>
          </div>
        </section>

        <section className="mt-6 flex-1">
          <div className="mb-3 flex items-center justify-between">
            <h2 className="text-base font-semibold">历史记录</h2>
            <div className="flex items-center gap-1 text-xs text-slate-500">
              <Download className="size-3.5" />
              {expenses.length} 条
            </div>
          </div>

          {loading ? (
            <div className="rounded-lg border border-slate-200 bg-white p-4 text-sm text-slate-500">
              正在加载记录...
            </div>
          ) : expenses.length === 0 ? (
            <div className="rounded-lg border border-slate-200 bg-white p-5 text-center text-sm text-slate-500">
              暂无记录，保存第一笔支出后会显示在这里。
            </div>
          ) : (
            <div className="grid gap-3 pb-8">
              {recentExpenses.map((expense) => (
                <article
                  key={expense.id}
                  className="rounded-lg border border-slate-200 bg-white p-4 shadow-sm"
                >
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0">
                      <p className="text-sm text-slate-500">{expense.date}</p>
                      <h3 className="mt-1 truncate text-base font-medium">
                        {expense.description}
                      </h3>
                    </div>
                    <p className="shrink-0 text-lg font-semibold text-slate-950">
                      {money(expense.amount)}
                    </p>
                  </div>

                  <div className="mt-3 flex justify-end gap-2">
                    <Button
                      type="button"
                      variant="outline"
                      size="sm"
                      onClick={() => editExpense(expense)}
                    >
                      <Pencil />
                      编辑
                    </Button>
                    <Button
                      type="button"
                      variant="destructive"
                      size="sm"
                      onClick={() => removeExpense(expense)}
                    >
                      <Trash2 />
                      删除
                    </Button>
                  </div>
                </article>
              ))}

              {expenses.length > recentExpenses.length ? (
                <details className="rounded-lg border border-slate-200 bg-white p-4">
                  <summary className="cursor-pointer text-sm font-medium text-slate-700">
                    查看全部记录
                  </summary>
                  <div className="mt-3 grid gap-3">
                    {expenses.slice(5).map((expense) => (
                      <article
                        key={expense.id}
                        className="border-t border-slate-100 pt-3"
                      >
                        <div className="flex items-start justify-between gap-3">
                          <div className="min-w-0">
                            <p className="text-sm text-slate-500">{expense.date}</p>
                            <h3 className="mt-1 truncate text-base font-medium">
                              {expense.description}
                            </h3>
                          </div>
                          <p className="shrink-0 font-semibold">
                            {money(expense.amount)}
                          </p>
                        </div>
                        <div className="mt-3 flex justify-end gap-2">
                          <Button
                            type="button"
                            variant="outline"
                            size="sm"
                            onClick={() => editExpense(expense)}
                          >
                            <Pencil />
                            编辑
                          </Button>
                          <Button
                            type="button"
                            variant="destructive"
                            size="sm"
                            onClick={() => removeExpense(expense)}
                          >
                            <Trash2 />
                            删除
                          </Button>
                        </div>
                      </article>
                    ))}
                  </div>
                </details>
              ) : null}
            </div>
          )}
        </section>
      </div>
    </main>
  );
}
