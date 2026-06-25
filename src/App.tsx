import { useState, useCallback, useRef, useEffect } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";
import {
  detectColumns,
  readTransactions,
  highlightTransactions,
  clearHighlights,
  createSummarySheet,
  exportToCsv,
  type ColumnMap,
} from "./lib/excel";
import { useCurrency, CURRENCIES } from "./hooks/use-currency";
import { buildSummary, CATEGORIES, type Category, type Summary, type Transaction } from "./lib/categorizer";
import { parsePastedText, writeToExcelSheet } from "./lib/csv-parser";
import { parseFileToTransactions, ACCEPTED_MIME_TYPES, SUPPORTED_EXTENSIONS } from "./lib/file-parser";
import { exportToPdf, buildReportHtml } from "./lib/pdf";
import { getLicense, checkLicenseValid } from "./lib/payment";
import PaymentGate from "./components/PaymentGate";
import SubscriptionDashboard from "./components/SubscriptionDashboard";
import { useAppConfig } from "./context/AppConfigContext";
import iconLogo from "@assets/icons/icon-logo.png";
import iconAnalyze from "@assets/icons/icon-analyze.png";
import iconPaste from "@assets/icons/icon-paste.png";
import iconIncome from "@assets/icons/icon-income.png";
import iconExpenses from "@assets/icons/icon-expenses.png";
import iconSavings from "@assets/icons/icon-savings.png";
import iconRate from "@assets/icons/icon-rate.png";
import iconHighlight from "@assets/icons/icon-highlight.png";
import iconExport from "@assets/icons/icon-export.png";
import iconPro from "@assets/icons/icon-pro.png";

declare const Excel: typeof import("@microsoft/office-js").Excel;
declare const Office: typeof import("@microsoft/office-js");

type Step = "idle" | "upload" | "paste" | "importing" | "loading" | "results" | "error" | "subscription";
type ResultTab = "overview" | "categories" | "transactions" | "budget";

const isOfficeAvailable = () =>
  typeof Office !== "undefined" && typeof Excel !== "undefined";

function runExcel<T>(fn: (context: Excel.RequestContext) => Promise<T>): Promise<T> {
  if (!isOfficeAvailable()) {
    return Promise.reject(new Error("Office.js not available — open this add-in inside Excel."));
  }
  return Excel.run(fn);
}

// ── KPI Card ─────────────────────────────────────────────────────────────────
function KpiCard({
  label, value, color, icon, gradientClass, delay = ""
}: {
  label: string; value: string; color: string; icon: string; gradientClass: string; delay?: string;
}) {
  return (
    <div className={`rounded-2xl border p-5 shadow-sm hover:shadow-md transition-all duration-300 hover:-translate-y-0.5 animate-fade-in-up ${gradientClass} ${delay}`}>
      <div className="flex items-center gap-2.5 mb-3">
        <div className="w-9 h-9 rounded-xl bg-white/70 flex items-center justify-center shadow-sm">
          <img src={icon} alt={label} className="w-5 h-5 object-contain" />
        </div>
        <p className="text-sm font-bold text-foreground/70 uppercase tracking-wide">{label}</p>
      </div>
      <p className={`text-3xl font-extrabold tracking-tight ${color}`}>{value}</p>
    </div>
  );
}

// ── Action Button ─────────────────────────────────────────────────────────────
function ActionBtn({
  onClick, disabled, children, variant = "primary", size = "md"
}: {
  onClick: () => void; disabled?: boolean; children: React.ReactNode;
  variant?: "primary" | "secondary" | "ghost"; size?: "md" | "sm";
}) {
  const base = size === "sm"
    ? "inline-flex items-center justify-center gap-2 text-sm font-semibold py-2.5 px-4 rounded-xl transition-all duration-200 disabled:opacity-50 disabled:cursor-not-allowed hover:scale-[1.02] active:scale-95"
    : "flex-1 flex items-center justify-center gap-2 text-base font-semibold py-3.5 px-5 rounded-xl transition-all duration-200 disabled:opacity-50 disabled:cursor-not-allowed hover:scale-[1.01] active:scale-95";
  const styles = variant === "primary"
    ? `${base} bg-gradient-to-r from-blue-600 to-blue-700 text-white hover:from-blue-500 hover:to-blue-600 shadow-md shadow-blue-200`
    : variant === "ghost"
    ? `${base} text-muted-foreground hover:text-foreground hover:bg-muted`
    : `${base} bg-white text-foreground hover:bg-muted/50 border border-border shadow-sm`;
  return (
    <button onClick={onClick} disabled={disabled} className={styles}>
      {children}
    </button>
  );
}

// ── Pro Lock Badge ─────────────────────────────────────────────────────────────
function ProBadge({ onClick }: { onClick: () => void }) {
  return (
    <button
      onClick={onClick}
      className="flex items-center gap-1.5 text-xs font-bold bg-gradient-to-r from-amber-400 to-orange-400 text-white px-3 py-1 rounded-full shadow-sm hover:from-amber-500 hover:to-orange-500 transition-all duration-200 hover:scale-105 active:scale-95 shrink-0"
    >
      🔒 PRO
    </button>
  );
}

// ── Health Score Ring ─────────────────────────────────────────────────────────
function HealthRing({ score }: { score: number }) {
  const color = score >= 75 ? "#22c55e" : score >= 50 ? "#f59e0b" : "#ef4444";
  const bgColor = score >= 75 ? "#f0fdf4" : score >= 50 ? "#fffbeb" : "#fff5f5";
  const label = score >= 75 ? "Excellent" : score >= 60 ? "Good" : score >= 45 ? "Fair" : "Needs Work";
  const r = 38;
  const circ = 2 * Math.PI * r;
  const dash = (score / 100) * circ;
  return (
    <div className="flex items-center gap-5">
      <div className="relative w-24 h-24 shrink-0">
        <svg className="w-24 h-24 -rotate-90" viewBox="0 0 96 96">
          <circle cx="48" cy="48" r={r} fill={bgColor} stroke="#e2e8f0" strokeWidth="8" />
          <circle cx="48" cy="48" r={r} fill="none" stroke={color} strokeWidth="8"
            strokeDasharray={`${dash} ${circ}`} strokeLinecap="round"
            style={{ transition: "stroke-dasharray 0.8s ease" }} />
        </svg>
        <div className="absolute inset-0 flex items-center justify-center">
          <span className="text-2xl font-extrabold" style={{ color }}>{score}</span>
        </div>
      </div>
      <div>
        <p className="text-xl font-bold text-foreground">Financial Health</p>
        <p className="text-base font-bold mt-1" style={{ color }}>{label}</p>
        <p className="text-sm text-muted-foreground mt-0.5">out of 100</p>
      </div>
    </div>
  );
}

// ── Spinner ───────────────────────────────────────────────────────────────────
function Spinner({ size = "md", color = "border-primary" }: { size?: "sm" | "md"; color?: string }) {
  const s = size === "sm" ? "w-4 h-4 border-2" : "w-5 h-5 border-2";
  return <span className={`${s} ${color} border-t-transparent rounded-full animate-spin inline-block shrink-0`} />;
}

// ── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const config = useAppConfig();
  const appName = config.appearance.name;
  const appTagline = config.appearance.tagline;
  const { symbol, setSymbol, fmt, fmtShort } = useCurrency();
  const [showCurrencyPicker, setShowCurrencyPicker] = useState(false);

  const [step, setStep] = useState<Step>("idle");
  const [error, setError] = useState<string>("");
  const [summary, setSummary] = useState<Summary | null>(null);
  const [highlighting, setHighlighting] = useState(false);
  const [highlightDone, setHighlightDone] = useState(false);
  const [clearing, setClearing] = useState(false);
  const [clearDone, setClearDone] = useState(false);
  const [reAnalyzing, setReAnalyzing] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [exportDone, setExportDone] = useState(false);
  const [activeTab, setActiveTab] = useState<ResultTab>("overview");
  const [csvText, setCsvText] = useState("");
  const [csvError, setCsvError] = useState("");
  const [isPro, setIsPro] = useState(false);
  const [showPayment, setShowPayment] = useState(false);
  const [paymentMode, setPaymentMode] = useState<"pay" | "key">("pay");
  const [pendingAction, setPendingAction] = useState<"highlight" | "export" | "csv" | "pdf" | "budget" | "recurring" | null>(null);
  const [txSearch, setTxSearch] = useState("");
  const [txCategoryFilter, setTxCategoryFilter] = useState("All");
  const [txTypeFilter, setTxTypeFilter] = useState<"all" | "credit" | "debit">("all");
  const [budgets, setBudgets] = useState<Record<string, string>>({});
  const [showDuplicates, setShowDuplicates] = useState(false);
  const [showRecurring, setShowRecurring] = useState(false);
  const [showPdfDialog, setShowPdfDialog] = useState(false);
  const [pdfEmail, setPdfEmail] = useState("");
  const [pdfEmailSending, setPdfEmailSending] = useState(false);
  const [pdfEmailSent, setPdfEmailSent] = useState(false);
  const [pdfEmailError, setPdfEmailError] = useState("");
  const [actionError, setActionError] = useState("");
  const [categoryOverrides, setCategoryOverrides] = useState<Record<number, Category>>({});
  const [expandedTxRow, setExpandedTxRow] = useState<number | null>(null);
  const [txNotes, setTxNotes] = useState<Record<number, string>>({});
  const [flaggedRows, setFlaggedRows] = useState<Set<number>>(new Set());
  const [showFlagged, setShowFlagged] = useState(false);
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [uploadError, setUploadError] = useState("");
  const [isDragOver, setIsDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const lowestPlanPrice = config.plans.length > 0 ? Math.min(...config.plans.map((p) => p.price)) : 5;
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    const key = getLicense();
    if (key) checkLicenseValid(key).then((valid) => setIsPro(valid));
  }, []);

  const requirePro = useCallback((action: typeof pendingAction) => {
    if (isPro) return true;
    setPendingAction(action);
    setShowPayment(true);
    return false;
  }, [isPro]);

  const onPaymentUnlocked = useCallback(() => {
    setIsPro(true);
    setShowPayment(false);
    const action = pendingAction;
    setPendingAction(null);
    if (action === "highlight") doHighlight();
    if (action === "export") doExport();
    if (action === "csv") doExportCsv();
    if (action === "pdf") setShowPdfDialog(true);
    if (action === "budget") setActiveTab("budget");
    if (action === "recurring") setShowRecurring(true);
  }, [pendingAction]); // eslint-disable-line react-hooks/exhaustive-deps

  const analyzeSheet = useCallback(async () => {
    setStep("loading"); setError(""); setSummary(null); setExportDone(false);
    try {
      const txns: Transaction[] = await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await ctx.sync();
        const columnMap: ColumnMap | null = await detectColumns(sheet);
        if (!columnMap) throw new Error("Could not find required columns in row 1.\n\nSupported formats:\n• Date, Description, Amount\n• Date, Description, Debit, Credit (with optional Balance)\n\nMake sure your sheet has a header row with these column names.");
        return await readTransactions(sheet, columnMap);
      });
      if (txns.length === 0) throw new Error("No transactions found. Check that the sheet has data rows below the header.");
      setSummary(buildSummary(txns));
      setStep("results");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setStep("error");
    }
  }, []);

  const analyzeUploadedFile = useCallback(async () => {
    if (!uploadFile) return;
    setUploadError("");
    setStep("loading"); setError(""); setSummary(null); setExportDone(false);
    try {
      const txns = await parseFileToTransactions(uploadFile);
      if (txns.length === 0) throw new Error("No transactions found in the file. Check that there is a header row and at least one data row.");
      setSummary(buildSummary(txns));
      setStep("results");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setStep("error");
    }
  }, [uploadFile]);

  const doReAnalyze = useCallback(async () => {
    setReAnalyzing(true); setActionError("");
    try {
      const txns: Transaction[] = await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await ctx.sync();
        const columnMap: ColumnMap | null = await detectColumns(sheet);
        if (!columnMap) throw new Error("Could not find required columns (Date, Description, Amount) in row 1.");
        return await readTransactions(sheet, columnMap);
      });
      if (txns.length === 0) throw new Error("No transactions found in the active sheet.");
      setSummary(buildSummary(txns));
      setExportDone(false); setHighlightDone(false); setClearDone(false);
    } catch (e: unknown) {
      setActionError(e instanceof Error ? e.message : String(e));
    } finally { setReAnalyzing(false); }
  }, []);

  const importAndAnalyzeCsv = useCallback(async () => {
    setCsvError("");
    const parsed = parsePastedText(csvText);
    if (!parsed) { setCsvError("Could not parse. Make sure there's a header row and at least one data row."); return; }
    setStep("importing");
    try {
      await runExcel(async (ctx) => { await writeToExcelSheet(parsed, ctx); });
      setStep("loading");
      const txns: Transaction[] = await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        await ctx.sync();
        const columnMap = await detectColumns(sheet);
        if (!columnMap) throw new Error("Columns could not be mapped after import.");
        return await readTransactions(sheet, columnMap);
      });
      if (txns.length === 0) throw new Error("No transactions were parsed from the pasted data.");
      setSummary(buildSummary(txns));
      setStep("results");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
      setStep("error");
    }
  }, [csvText]);

  const doHighlight = useCallback(async () => {
    if (!summary) return;
    setHighlighting(true); setHighlightDone(false); setActionError("");
    try {
      await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const columnMap = await detectColumns(sheet);
        if (columnMap) await highlightTransactions(sheet, summary.transactions, columnMap);
        else throw new Error("Could not detect columns on the active sheet.");
      });
      setHighlightDone(true);
    } catch (e: unknown) {
      setActionError(e instanceof Error ? e.message : String(e));
    } finally { setHighlighting(false); }
  }, [summary]);

  const doClearHighlights = useCallback(async () => {
    setClearing(true); setClearDone(false); setActionError("");
    try {
      await runExcel(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        await clearHighlights(sheet);
      });
      setClearDone(true); setHighlightDone(false);
      setTimeout(() => setClearDone(false), 2000);
    } catch (e: unknown) {
      setActionError(e instanceof Error ? e.message : String(e));
    } finally { setClearing(false); }
  }, []);

  const doExport = useCallback(async () => {
    if (!summary) return;
    setExporting(true); setActionError("");
    try {
      await runExcel(async (ctx) => { await createSummarySheet(summary, ctx, symbol); });
      setExportDone(true);
    } catch (e: unknown) {
      setActionError(e instanceof Error ? e.message : String(e));
    } finally { setExporting(false); }
  }, [summary]);

  const doExportCsv = useCallback(() => {
    if (!summary) return;
    exportToCsv(summary, txNotes, flaggedRows);
  }, [summary, txNotes, flaggedRows]);

  const doExportPdf = useCallback(() => {
    if (!summary) return;
    exportToPdf(summary, appName, symbol, txNotes, flaggedRows);
  }, [summary, appName, symbol, txNotes, flaggedRows]);

  const doSendReportByEmail = useCallback(async (email: string) => {
    if (!summary) return;
    setPdfEmailSending(true); setPdfEmailError("");
    try {
      const html = buildReportHtml(summary, appName, symbol, txNotes, flaggedRows);
      const apiBase = (import.meta.env.VITE_API_URL ?? "").replace(/\/$/, "");
      const res = await fetch(`${apiBase}/api/send-report`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email, html, appName }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || `Server error (${res.status})`);
      setPdfEmailSent(true);
    } catch (e: unknown) {
      setPdfEmailError(e instanceof Error ? e.message : String(e));
    } finally { setPdfEmailSending(false); }
  }, [summary, appName]);

  const handleHighlight = useCallback(() => { if (requirePro("highlight")) doHighlight(); }, [requirePro, doHighlight]);
  const handleExport = useCallback(() => { if (requirePro("export")) doExport(); }, [requirePro, doExport]);
  const handleExportCsv = useCallback(() => { if (requirePro("csv")) doExportCsv(); }, [requirePro, doExportCsv]);
  const handleExportPdf = useCallback(() => {
    if (requirePro("pdf")) { setPdfEmailSent(false); setPdfEmailError(""); setPdfEmail(""); setShowPdfDialog(true); }
  }, [requirePro]);
  const handleBudgetTab = useCallback(() => { if (requirePro("budget")) setActiveTab("budget"); }, [requirePro]);
  const handleToggleRecurring = useCallback(() => { if (requirePro("recurring")) setShowRecurring((v) => !v); }, [requirePro]);

  const handleRecategorize = useCallback((row: number, newCat: Category) => {
    if (!summary) return;
    const newOverrides = { ...categoryOverrides, [row]: newCat };
    setCategoryOverrides(newOverrides);
    setExpandedTxRow(null);
    const updatedTxns = summary.transactions.map((tx) =>
      newOverrides[tx.row] ? { ...tx, category: newOverrides[tx.row] } : tx
    );
    setSummary(buildSummary(updatedTxns));
  }, [summary, categoryOverrides]);

  const reset = () => {
    setStep("idle"); setSummary(null); setError("");
    setExportDone(false); setActiveTab("overview");
    setCsvText(""); setCsvError("");
    setTxSearch(""); setTxCategoryFilter("All"); setTxTypeFilter("all");
    setBudgets({}); setShowDuplicates(false); setShowRecurring(false);
    setActionError(""); setShowPdfDialog(false); setHighlightDone(false);
    setClearing(false); setClearDone(false);
    setCategoryOverrides({}); setExpandedTxRow(null); setTxNotes({});
    setFlaggedRows(new Set()); setShowFlagged(false);
    setUploadFile(null); setUploadError(""); setIsDragOver(false);
  };

  const openSubscription = () => setStep("subscription");

  const topCategories = summary
    ? Object.entries(summary.byCategory).sort((a, b) => b[1].total - a[1].total)
    : [];
  const maxCatTotal = topCategories.length > 0 ? topCategories[0][1].total : 1;

  const filteredTxns: Transaction[] = summary
    ? summary.transactions.filter((tx) => {
        const matchSearch = txSearch === "" || tx.description.toLowerCase().includes(txSearch.toLowerCase());
        const matchCat = txCategoryFilter === "All" || tx.category.name === txCategoryFilter;
        const matchType = txTypeFilter === "all" || tx.type === txTypeFilter;
        const matchDupe = !showDuplicates || summary.duplicateRows.has(tx.row);
        const matchFlag = !showFlagged || flaggedRows.has(tx.row);
        return matchSearch && matchCat && matchType && matchDupe && matchFlag;
      })
    : [];

  const categoryNames = summary ? ["All", ...Object.keys(summary.byCategory).sort()] : ["All"];

  return (
    <div className="flex flex-col h-screen bg-background text-foreground">

      {/* ── Payment overlay ── */}
      {showPayment && (
        <PaymentGate
          initialMode={paymentMode}
          onUnlocked={onPaymentUnlocked}
          onDismiss={() => { setShowPayment(false); setPendingAction(null); setPaymentMode("pay"); }}
        />
      )}

      {/* ── PDF / Email dialog ── */}
      {showPdfDialog && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm px-4 animate-fade-in">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-md overflow-hidden animate-scale-in">
            <div className="bg-gradient-to-r from-blue-600 to-indigo-600 px-6 py-5 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-white/20 flex items-center justify-center">
                  <svg className="w-5 h-5 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
                    <line x1="9" y1="15" x2="15" y2="15"/><line x1="9" y1="11" x2="15" y2="11"/>
                  </svg>
                </div>
                <div>
                  <p className="text-base font-bold text-white">Export Report</p>
                  <p className="text-xs text-blue-100">PDF or email delivery</p>
                </div>
              </div>
              <button onClick={() => setShowPdfDialog(false)} className="text-white/70 hover:text-white transition-colors" data-testid="button-close-pdf-dialog">
                <svg className="w-6 h-6" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
              </button>
            </div>

            <div className="p-6 space-y-4">
              <button
                onClick={() => { doExportPdf(); setShowPdfDialog(false); }}
                className="w-full flex items-center gap-4 bg-gradient-to-r from-blue-600 to-blue-700 text-white rounded-2xl px-5 py-4 hover:from-blue-500 hover:to-blue-600 transition-all shadow-lg shadow-blue-200 hover:scale-[1.01] active:scale-[0.99]"
                data-testid="button-open-pdf"
              >
                <div className="w-11 h-11 rounded-xl bg-white/20 flex items-center justify-center shrink-0">
                  <svg className="w-6 h-6" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                    <polyline points="6 9 12 15 18 9"/>
                  </svg>
                </div>
                <div className="text-left">
                  <p className="text-base font-bold">Open PDF</p>
                  <p className="text-sm text-blue-100 mt-0.5">Opens print dialog to save as PDF</p>
                </div>
              </button>

              <div className="flex items-center gap-3">
                <div className="flex-1 h-px bg-border" />
                <span className="text-sm text-muted-foreground font-medium px-2">or send by email</span>
                <div className="flex-1 h-px bg-border" />
              </div>

              {pdfEmailSent ? (
                <div className="flex flex-col items-center gap-3 py-6 text-center animate-scale-in">
                  <div className="w-16 h-16 rounded-full bg-green-100 flex items-center justify-center">
                    <svg className="w-8 h-8 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <polyline points="20 6 9 17 4 12"/>
                    </svg>
                  </div>
                  <div>
                    <p className="text-lg font-bold text-green-700">Report sent!</p>
                    <p className="text-sm text-muted-foreground mt-1">Check <span className="font-semibold text-foreground">{pdfEmail}</span> for your report</p>
                  </div>
                  <button onClick={() => { setPdfEmailSent(false); setPdfEmail(""); }} className="text-sm text-primary hover:underline mt-1">
                    Send to another email
                  </button>
                </div>
              ) : (
                <div className="space-y-3">
                  <p className="text-sm font-semibold text-foreground">Send report to email</p>
                  <input
                    type="email"
                    value={pdfEmail}
                    onChange={(e) => { setPdfEmail(e.target.value); setPdfEmailError(""); }}
                    placeholder="your@email.com"
                    className="w-full rounded-xl border-2 border-border bg-muted/30 px-4 py-3 text-base focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50"
                    data-testid="input-pdf-email"
                  />
                  {pdfEmailError && <p className="text-sm text-destructive font-medium">{pdfEmailError}</p>}
                  <button
                    onClick={() => doSendReportByEmail(pdfEmail)}
                    disabled={pdfEmailSending || !pdfEmail.includes("@")}
                    className="w-full flex items-center justify-center gap-2 bg-gradient-to-r from-emerald-600 to-green-600 text-white rounded-xl px-5 py-3 text-base font-bold hover:from-emerald-500 hover:to-green-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all shadow-md shadow-green-200 hover:scale-[1.01] active:scale-[0.99]"
                    data-testid="button-send-report-email"
                  >
                    {pdfEmailSending
                      ? <><Spinner color="border-white" />Sending…</>
                      : <><svg className="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>Send Report</>
                    }
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ── Header ── */}
      <header className="flex items-center gap-4 px-5 py-4 bg-white border-b border-border shadow-sm shrink-0">
        <div className="flex items-center justify-center w-11 h-11 rounded-2xl overflow-hidden shadow-md">
          <img src={iconLogo} alt="App logo" className="w-11 h-11 object-cover" />
        </div>
        <div className="flex-1 min-w-0">
          <div className="font-extrabold text-lg leading-tight text-foreground truncate">{appName}</div>
          <div className="text-sm text-muted-foreground font-medium">Excel Add-in</div>
        </div>
        <div className="flex items-center gap-2.5 shrink-0">
          {/* Currency picker */}
          <div className="relative">
            <button
              onClick={() => setShowCurrencyPicker((v) => !v)}
              data-testid="button-currency-picker"
              className="flex items-center gap-1.5 text-sm font-bold bg-muted text-muted-foreground hover:bg-muted/70 hover:text-foreground px-3 py-1.5 rounded-xl border border-border transition-all hover:scale-105"
              title="Change currency symbol"
            >
              <span className="font-mono">{symbol}</span>
              <svg className="w-3.5 h-3.5 opacity-60" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <polyline points="6 9 12 15 18 9" />
              </svg>
            </button>
            {showCurrencyPicker && (
              <>
                <div className="fixed inset-0 z-40" onClick={() => setShowCurrencyPicker(false)} />
                <div className="absolute right-0 top-full mt-2 z-50 bg-white border border-border rounded-2xl shadow-2xl w-56 overflow-hidden animate-scale-in">
                  <div className="px-4 py-3 border-b border-border bg-muted/30">
                    <p className="text-xs font-bold text-muted-foreground uppercase tracking-wider">Currency Symbol</p>
                  </div>
                  <div className="max-h-64 overflow-y-auto py-1">
                    {CURRENCIES.map((c) => (
                      <button
                        key={c.symbol}
                        data-testid={`button-currency-${c.symbol}`}
                        onClick={() => { setSymbol(c.symbol); setShowCurrencyPicker(false); }}
                        className={`w-full flex items-center gap-3 px-4 py-2.5 text-left hover:bg-muted/50 transition-colors ${symbol === c.symbol ? "bg-primary/5" : ""}`}
                      >
                        <span className={`font-mono text-sm font-bold w-8 shrink-0 ${symbol === c.symbol ? "text-primary" : "text-foreground"}`}>{c.symbol}</span>
                        <span className="text-sm text-muted-foreground truncate">{c.label}</span>
                        {symbol === c.symbol && (
                          <svg className="w-4 h-4 text-primary shrink-0 ml-auto" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                            <polyline points="20 6 9 17 4 12" />
                          </svg>
                        )}
                      </button>
                    ))}
                  </div>
                </div>
              </>
            )}
          </div>

          {isPro && (
            <button
              onClick={openSubscription}
              className="flex items-center gap-1.5 text-xs font-bold bg-gradient-to-r from-amber-400 to-orange-400 text-white px-3 py-1.5 rounded-xl shadow-sm hover:from-amber-500 hover:to-orange-500 transition-all hover:scale-105"
            >
              <img src={iconPro} alt="Pro" className="w-3.5 h-3.5 object-contain" /> PRO
            </button>
          )}
          {step === "results" && (
            <button
              onClick={doReAnalyze}
              disabled={reAnalyzing}
              title="Re-read the active sheet and refresh results"
              data-testid="button-reanalyze"
              className="flex items-center justify-center w-10 h-10 rounded-xl text-muted-foreground hover:text-foreground hover:bg-muted transition-all disabled:opacity-40 hover:scale-105"
            >
              <svg className={`w-5 h-5 ${reAnalyzing ? "animate-spin" : ""}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <polyline points="23 4 23 10 17 10" />
                <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10" />
              </svg>
            </button>
          )}
          {(step === "results" || step === "error" || step === "upload" || step === "paste" || step === "subscription") && (
            <button onClick={reset}
              className="flex items-center gap-1.5 text-sm font-semibold text-muted-foreground hover:text-foreground px-3 py-2 rounded-xl hover:bg-muted transition-all hover:scale-105">
              <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <polyline points="15 18 9 12 15 6" />
              </svg>
              Back
            </button>
          )}
        </div>
      </header>

      {/* ── Body ── */}
      <div className="flex-1 overflow-y-auto">

        {/* ── IDLE ── */}
        {step === "idle" && (
          <div className="flex flex-col min-h-full">
            {/* Hero */}
            <div className="bg-gradient-to-br from-blue-600 via-blue-700 to-indigo-800 px-6 py-10 text-center text-white">
              <div className="flex flex-col items-center gap-4 animate-fade-in-up">
                <div className="w-20 h-20 rounded-3xl overflow-hidden shadow-2xl shadow-blue-900/40 ring-4 ring-white/20">
                  <img src={iconLogo} alt="Bank Statement Analyzer" className="w-20 h-20 object-cover" />
                </div>
                <div>
                  <h1 className="text-3xl font-extrabold tracking-tight mb-2">{appName}</h1>
                  <p className="text-base text-blue-100 max-w-xs mx-auto leading-relaxed">{appTagline}</p>
                </div>
              </div>
            </div>

            {/* Action cards */}
            <div className="px-5 py-6 space-y-4 -mt-4">
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                <button onClick={() => setStep("upload")}
                  className="text-left bg-white border-2 border-border rounded-2xl p-5 hover:border-primary/60 hover:shadow-lg transition-all duration-200 group hover:-translate-y-0.5 animate-fade-in-up">
                  <div className="flex items-center gap-4">
                    <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-blue-100 to-blue-200 flex items-center justify-center shrink-0 group-hover:from-blue-200 group-hover:to-blue-300 transition-all overflow-hidden shadow-sm">
                      <img src={iconAnalyze} alt="Upload statement" className="w-10 h-10 object-contain" />
                    </div>
                    <div className="min-w-0">
                      <p className="text-base font-bold text-foreground">Upload Statement</p>
                      <p className="text-sm text-muted-foreground mt-0.5">.xlsx · .csv · .pdf · .txt</p>
                    </div>
                    <svg className="w-5 h-5 text-muted-foreground ml-auto shrink-0 group-hover:text-primary group-hover:translate-x-0.5 transition-all" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <polyline points="9 18 15 12 9 6" />
                    </svg>
                  </div>
                </button>

                <button onClick={() => setStep("paste")}
                  className="text-left bg-white border-2 border-border rounded-2xl p-5 hover:border-emerald-400/60 hover:shadow-lg transition-all duration-200 group hover:-translate-y-0.5 animate-fade-in-up delay-100">
                  <div className="flex items-center gap-4">
                    <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-emerald-100 to-emerald-200 flex items-center justify-center shrink-0 group-hover:from-emerald-200 group-hover:to-emerald-300 transition-all overflow-hidden shadow-sm">
                      <img src={iconPaste} alt="Paste CSV" className="w-10 h-10 object-contain" />
                    </div>
                    <div className="min-w-0">
                      <p className="text-base font-bold text-foreground">Paste CSV / Text</p>
                      <p className="text-sm text-muted-foreground mt-0.5">From your bank portal</p>
                    </div>
                    <svg className="w-5 h-5 text-muted-foreground ml-auto shrink-0 group-hover:text-emerald-600 group-hover:translate-x-0.5 transition-all" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <polyline points="9 18 15 12 9 6" />
                    </svg>
                  </div>
                </button>
              </div>

              {/* Free vs Pro card */}
              {!isPro && (
                <div className="rounded-2xl border-2 border-amber-200 overflow-hidden shadow-sm animate-fade-in-up delay-200">
                  <div className="bg-gradient-to-r from-amber-500 to-orange-500 px-5 py-3.5 flex items-center justify-between">
                    <div>
                      <p className="text-base font-bold text-white">Upgrade to Pro</p>
                      <p className="text-xs text-amber-100">Unlock all features</p>
                    </div>
                    <button onClick={() => { setPaymentMode("pay"); setShowPayment(true); }}
                      className="text-sm font-bold text-amber-700 bg-white hover:bg-amber-50 px-4 py-2 rounded-xl transition-all shadow-md hover:scale-105 active:scale-95">
                      From ${lowestPlanPrice} USDT
                    </button>
                  </div>
                  <div className="p-4 bg-white">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-y-2 gap-x-4">
                      {[
                        { label: "Analyze transactions", free: true },
                        { label: "Auto-categorize (14 categories)", free: true },
                        { label: "Monthly trend chart", free: true },
                        { label: "Financial health score", free: true },
                        { label: "Search & filter transactions", free: true },
                        { label: "Top merchants breakdown", free: true },
                        { label: "Highlight cells by category", free: false },
                        { label: "Export Excel summary sheet", free: false },
                        { label: "Download CSV with categories", free: false },
                        { label: "Export color-coded PDF report", free: false },
                        { label: "Recurring subscriptions detector", free: false },
                        { label: "Duplicate transaction flags", free: false },
                        { label: "Category budget tracker", free: false },
                      ].map((f) => (
                        <div key={f.label} className="flex items-center gap-2.5">
                          <span className={`text-base shrink-0 ${f.free ? "text-green-600" : "text-muted-foreground/40"}`}>
                            {f.free ? "✓" : "🔒"}
                          </span>
                          <span className={`text-sm ${f.free ? "text-foreground font-medium" : "text-muted-foreground/60"}`}>
                            {f.label}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {!isPro && (
                <div className="text-center animate-fade-in-up delay-300">
                  <button
                    onClick={() => { setPaymentMode("key"); setShowPayment(true); }}
                    className="text-sm text-muted-foreground hover:text-primary transition-colors underline underline-offset-2"
                  >
                    Already have a license key? Enter it here
                  </button>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ── UPLOAD FILE ── */}
        {step === "upload" && (
          <div className="flex flex-col min-h-full p-5 gap-5 animate-fade-in">
            <div>
              <h2 className="text-2xl font-extrabold text-foreground mb-1.5">Upload Bank Statement</h2>
              <p className="text-base text-muted-foreground leading-relaxed">
                Upload your statement in any supported format and we'll extract the transactions automatically.
              </p>
            </div>

            <input
              ref={fileInputRef}
              type="file"
              accept={ACCEPTED_MIME_TYPES + ",.xlsx,.xls,.csv,.txt,.pdf"}
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0] ?? null;
                setUploadFile(f); setUploadError(""); e.target.value = "";
              }}
            />

            {/* Drop zone */}
            <div
              onClick={() => !uploadFile && fileInputRef.current?.click()}
              onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
              onDragLeave={() => setIsDragOver(false)}
              onDrop={(e) => {
                e.preventDefault(); setIsDragOver(false);
                const f = e.dataTransfer.files?.[0] ?? null;
                if (f) { setUploadFile(f); setUploadError(""); }
              }}
              className={`relative flex flex-col items-center justify-center gap-4 rounded-3xl border-3 border-dashed transition-all duration-200 cursor-pointer min-h-[240px] ${
                uploadFile
                  ? "border-primary/40 bg-primary/5 cursor-default"
                  : isDragOver
                  ? "border-primary bg-primary/10 scale-[1.01]"
                  : "border-border bg-muted/30 hover:border-primary/50 hover:bg-muted/60"
              }`}
            >
              {uploadFile ? (
                <div className="flex flex-col items-center gap-3 px-5 py-8 text-center animate-scale-in">
                  <div className="w-16 h-16 rounded-2xl bg-primary/10 flex items-center justify-center shadow-sm">
                    {uploadFile.name.endsWith(".pdf") ? (
                      <svg className="w-8 h-8 text-red-500" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
                        <line x1="9" y1="15" x2="15" y2="15"/>
                      </svg>
                    ) : uploadFile.name.match(/\.(xlsx?|xls)$/i) ? (
                      <svg className="w-8 h-8 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
                        <polyline points="8 13 10 17 16 11"/>
                      </svg>
                    ) : (
                      <svg className="w-8 h-8 text-primary" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
                      </svg>
                    )}
                  </div>
                  <div>
                    <p className="text-base font-bold text-foreground truncate max-w-[260px]">{uploadFile.name}</p>
                    <p className="text-sm text-muted-foreground mt-0.5">{(uploadFile.size / 1024).toFixed(1)} KB</p>
                  </div>
                  <button
                    onClick={(e) => { e.stopPropagation(); setUploadFile(null); setUploadError(""); }}
                    className="text-sm text-muted-foreground hover:text-destructive transition-colors underline underline-offset-2 mt-1"
                  >
                    Remove file
                  </button>
                </div>
              ) : (
                <div className="flex flex-col items-center gap-3 px-5 py-8 text-center">
                  <div className="w-16 h-16 rounded-2xl bg-muted flex items-center justify-center">
                    <svg className="w-8 h-8 text-muted-foreground" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                      <polyline points="17 8 12 3 7 8"/>
                      <line x1="12" y1="3" x2="12" y2="15"/>
                    </svg>
                  </div>
                  <div>
                    <p className="text-base font-semibold text-foreground">Drop your file here</p>
                    <p className="text-sm text-muted-foreground mt-0.5">or click to browse</p>
                  </div>
                  <div className="flex flex-wrap justify-center gap-1.5 mt-1">
                    {SUPPORTED_EXTENSIONS.map((ext) => (
                      <span key={ext} className="text-xs font-bold uppercase tracking-wide bg-white border border-border text-muted-foreground px-2 py-0.5 rounded-lg shadow-sm">
                        {ext.slice(1)}
                      </span>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* Supported formats */}
            <div className="bg-gradient-to-br from-blue-50 to-indigo-50 border border-blue-100 rounded-2xl px-4 py-4 text-sm text-foreground space-y-1.5">
              <p className="font-bold text-blue-800 text-sm uppercase tracking-wide mb-2">Supported formats</p>
              <p>📊 <span className="font-semibold">.xlsx / .xls</span> — Excel workbooks (first sheet used)</p>
              <p>📄 <span className="font-semibold">.csv / .txt</span> — Comma, tab, or pipe-separated</p>
              <p>📋 <span className="font-semibold">.pdf</span> — Bank-exported PDFs with selectable text</p>
            </div>

            {uploadError && (
              <div className="flex items-start gap-3 bg-red-50 border border-red-200 rounded-2xl px-4 py-3.5">
                <svg className="w-5 h-5 text-destructive shrink-0 mt-0.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <p className="text-sm text-destructive font-medium">{uploadError}</p>
              </div>
            )}

            <div className="flex gap-3 shrink-0 pb-4">
              <ActionBtn variant="secondary" onClick={() => { setUploadFile(null); setUploadError(""); fileInputRef.current?.click(); }}>
                Browse…
              </ActionBtn>
              <ActionBtn onClick={analyzeUploadedFile} disabled={!uploadFile}>
                <svg className="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/>
                  <path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/>
                </svg>
                Analyze File
              </ActionBtn>
            </div>
          </div>
        )}

        {/* ── PASTE CSV ── */}
        {step === "paste" && (
          <div className="flex flex-col min-h-full p-5 gap-5 animate-fade-in">
            <div>
              <h2 className="text-2xl font-extrabold text-foreground mb-1.5">Paste Bank Statement</h2>
              <p className="text-base text-muted-foreground leading-relaxed">
                Copy your statement from your bank's portal or CSV export and paste it below.
              </p>
            </div>
            <textarea
              ref={textareaRef}
              value={csvText}
              onChange={(e) => { setCsvText(e.target.value); setCsvError(""); }}
              placeholder={"Format 1 — Debit/Credit columns (recommended):\nDate,Description,Category,Debit,Credit,Balance\n6/1/2026,Opening Balance,,,10000,10000\n6/2/2026,Payroll,Income,,4500,14500\n\nFormat 2 — Single Amount column:\nDate,Description,Amount,Type\n01/06/2026,SALARY JUNE,650000,CR"}
              className="flex-1 min-h-[260px] w-full rounded-2xl border-2 border-border bg-white p-4 text-sm font-mono resize-none focus:outline-none focus:border-primary placeholder:text-muted-foreground/40 transition-colors shadow-sm"
              spellCheck={false}
            />
            {csvError && (
              <div className="flex items-start gap-3 bg-red-50 border border-red-200 rounded-2xl px-4 py-3.5">
                <svg className="w-5 h-5 text-destructive shrink-0 mt-0.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" />
                </svg>
                <p className="text-sm text-destructive font-medium">{csvError}</p>
              </div>
            )}
            {csvText.trim().length > 0 && (() => {
              const parsed = parsePastedText(csvText);
              return parsed ? (
                <p className="text-sm text-muted-foreground">
                  Detected <span className="font-bold text-foreground">{parsed.rawCount} rows</span> ·{" "}
                  <span className="font-bold text-foreground">{parsed.headers.length} columns</span>:{" "}
                  {parsed.headers.join(", ")}
                </p>
              ) : (
                <p className="text-sm text-amber-600 font-medium">⚠ Need at least a header row and one data row.</p>
              );
            })()}
            <div className="flex gap-3 shrink-0 pb-4">
              <ActionBtn variant="secondary" onClick={() => { setCsvText(""); setCsvError(""); }}>Clear</ActionBtn>
              <ActionBtn onClick={importAndAnalyzeCsv} disabled={!csvText.trim()}>
                <svg className="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <polyline points="16 16 12 12 8 16" /><line x1="12" y1="12" x2="12" y2="21" />
                  <path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3" />
                </svg>
                Import &amp; Analyze
              </ActionBtn>
            </div>
          </div>
        )}

        {/* ── LOADING / IMPORTING ── */}
        {(step === "importing" || step === "loading") && (
          <div className="flex flex-col items-center justify-center h-full gap-6 animate-fade-in">
            <div className="relative w-24 h-24">
              <div className="w-24 h-24 border-8 border-primary/20 rounded-full" />
              <div className="absolute inset-0 w-24 h-24 border-8 border-primary border-t-transparent rounded-full animate-spin" />
              <div className="absolute inset-3 w-18 h-18 border-4 border-blue-300/30 rounded-full" />
            </div>
            <div className="text-center">
              <p className="text-2xl font-extrabold text-foreground">
                {step === "importing" ? "Writing to Excel…" : "Analyzing transactions…"}
              </p>
              <p className="text-base text-muted-foreground mt-2">This will only take a moment</p>
            </div>
          </div>
        )}

        {/* ── ERROR ── */}
        {step === "error" && (
          <div className="flex flex-col items-center justify-center min-h-full px-6 py-10 text-center animate-fade-in">
            <div className="w-20 h-20 rounded-3xl bg-red-100 flex items-center justify-center mb-5 shadow-sm">
              <svg className="w-10 h-10 text-destructive" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10" />
                <line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" />
              </svg>
            </div>
            <h3 className="text-2xl font-extrabold text-destructive mb-3">Analysis Failed</h3>
            <p className="text-base text-muted-foreground mb-7 whitespace-pre-wrap leading-relaxed max-w-sm">{error}</p>
            <button onClick={reset}
              className="text-base font-bold bg-muted hover:bg-muted/70 text-foreground px-8 py-3 rounded-2xl transition-all hover:scale-105 active:scale-95 shadow-sm">
              ← Try Again
            </button>
          </div>
        )}

        {/* ── SUBSCRIPTION ── */}
        {step === "subscription" && (
          <div className="animate-fade-in">
            <SubscriptionDashboard onClose={reset} />
          </div>
        )}

        {/* ── RESULTS ── */}
        {step === "results" && summary && (
          <div className="flex flex-col h-full animate-fade-in">

            {/* Action buttons row */}
            <div className="flex gap-2 px-4 pt-4 pb-3 shrink-0 flex-wrap bg-white border-b border-border shadow-sm">
              <ActionBtn variant="secondary" onClick={handleHighlight} disabled={highlighting || clearing} size="sm">
                {highlighting ? <Spinner size="sm" />
                  : highlightDone ? <svg className="w-4 h-4 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12" /></svg>
                  : <img src={iconHighlight} alt="Highlight" className="w-4 h-4 object-contain" />}
                {highlightDone ? "Done!" : "Highlight"}
                {!isPro && <span className="text-xs opacity-60">🔒</span>}
              </ActionBtn>
              <ActionBtn variant="ghost" onClick={doClearHighlights} disabled={clearing || highlighting} size="sm">
                {clearing ? <Spinner size="sm" />
                  : clearDone ? <svg className="w-4 h-4 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12" /></svg>
                  : <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M20 5H9l-7 7 7 7h11a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2z"/><line x1="18" y1="9" x2="12" y2="15"/><line x1="12" y1="9" x2="18" y2="15"/></svg>}
                {clearDone ? "Cleared!" : "Clear"}
              </ActionBtn>
              <ActionBtn onClick={handleExport} disabled={exporting} size="sm">
                {exporting ? <Spinner size="sm" color="border-white" />
                  : exportDone ? <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12" /></svg>
                  : <img src={iconExport} alt="Export" className="w-4 h-4 object-contain" />}
                {exportDone ? "Exported!" : "Export"}
                {!isPro && <span className="text-xs opacity-60">🔒</span>}
              </ActionBtn>
              <ActionBtn variant="secondary" onClick={handleExportCsv} size="sm">
                <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" />
                </svg>
                CSV {!isPro && <span className="text-xs opacity-60">🔒</span>}
              </ActionBtn>
              <ActionBtn variant="secondary" onClick={handleExportPdf} size="sm">
                <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
                  <line x1="9" y1="15" x2="15" y2="15"/><line x1="9" y1="11" x2="15" y2="11"/>
                </svg>
                PDF {!isPro && <span className="text-xs opacity-60">🔒</span>}
              </ActionBtn>
            </div>

            {/* Action error banner */}
            {actionError && (
              <div className="mx-4 mt-3 flex items-start gap-3 bg-red-50 border border-red-200 rounded-2xl px-4 py-3.5 shrink-0 animate-fade-in">
                <svg className="w-5 h-5 text-destructive shrink-0 mt-0.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <p className="text-sm text-destructive font-medium flex-1 leading-relaxed">{actionError}</p>
                <button onClick={() => setActionError("")} className="text-destructive/60 hover:text-destructive shrink-0 transition-colors">
                  <svg className="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                    <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                  </svg>
                </button>
              </div>
            )}

            {/* Upsell banner */}
            {!isPro && (
              <button onClick={() => setShowPayment(true)}
                className="mx-4 mt-3 flex items-center justify-between bg-gradient-to-r from-amber-50 to-orange-50 border-2 border-amber-200 rounded-2xl px-5 py-3.5 hover:border-amber-400 transition-all group shrink-0 hover:scale-[1.005]">
                <div className="text-left">
                  <p className="text-sm font-bold text-amber-800">🚀 Unlock Pro — from ${lowestPlanPrice} USDT</p>
                  <p className="text-xs text-amber-600 mt-0.5">Budgets · CSV export · Recurring · Duplicates · PDF</p>
                </div>
                <svg className="w-5 h-5 text-amber-500 shrink-0 group-hover:translate-x-1 transition-transform" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <polyline points="9 18 15 12 9 6" />
                </svg>
              </button>
            )}

            {/* Tabs */}
            <div className="flex border-b border-border shrink-0 px-2 mt-3 bg-white">
              {(["overview", "categories", "transactions", "budget"] as const).map((tab) => {
                const isProTab = tab === "budget";
                const icons: Record<string, string> = { overview: "📊", categories: "🏷️", transactions: "📋", budget: "💰" };
                return (
                  <button key={tab} onClick={() => { if (tab === "budget") handleBudgetTab(); else setActiveTab(tab); }}
                    className={`flex-1 py-3 text-sm font-bold capitalize transition-all flex items-center justify-center gap-1.5 border-b-2 ${
                      activeTab === tab
                        ? "border-primary text-primary"
                        : "border-transparent text-muted-foreground hover:text-foreground hover:border-border"
                    }`}>
                    <span className="hidden sm:inline">{icons[tab]}</span>
                    {tab}
                    {isProTab && !isPro && <span className="text-xs">🔒</span>}
                  </button>
                );
              })}
            </div>

            {/* Tab content */}
            <div className="flex-1 overflow-y-auto p-4 space-y-4">

              {/* ── OVERVIEW TAB ── */}
              {activeTab === "overview" && (
                <>
                  {/* KPI grid */}
                  <div className="grid grid-cols-2 gap-3">
                    <KpiCard label="Income" value={fmt(summary.totalIncome)} color="text-green-700" icon={iconIncome} gradientClass="kpi-income" delay="delay-100" />
                    <KpiCard label="Expenses" value={fmt(summary.totalExpenses)} color="text-red-600" icon={iconExpenses} gradientClass="kpi-expense" delay="delay-150" />
                    <KpiCard label="Net Savings" value={fmt(summary.net)} color={summary.net >= 0 ? "text-blue-700" : "text-red-600"} icon={iconSavings} gradientClass="kpi-savings" delay="delay-200" />
                    <KpiCard label="Savings Rate" value={`${summary.savingsRate}%`} color={summary.savingsRate >= 20 ? "text-green-700" : summary.savingsRate >= 10 ? "text-yellow-700" : "text-red-600"} icon={iconRate} gradientClass="kpi-rate" delay="delay-300" />
                  </div>

                  {/* Health Score */}
                  <div className="bg-white border border-border rounded-2xl p-5 shadow-sm animate-fade-in-up">
                    <HealthRing score={summary.healthScore} />
                    {summary.healthTips.length > 0 && (
                      <div className="mt-4 space-y-2 pt-4 border-t border-border">
                        {summary.healthTips.map((tip, i) => (
                          <div key={i} className="flex items-start gap-2.5">
                            <span className="text-base shrink-0 mt-0.5">{i === 0 && summary.healthScore >= 60 ? "💡" : "⚠️"}</span>
                            <p className="text-sm text-muted-foreground leading-relaxed">{tip}</p>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  {/* Monthly Trend Chart */}
                  {summary.monthly.length > 1 && (
                    <div className="bg-white border border-border rounded-2xl p-5 shadow-sm animate-fade-in-up">
                      <p className="text-sm font-bold text-muted-foreground uppercase tracking-widest mb-4">Monthly Trend</p>
                      <ResponsiveContainer width="100%" height={160}>
                        <BarChart data={summary.monthly} barGap={3} barCategoryGap="28%">
                          <XAxis dataKey="month" tick={{ fontSize: 12, fontWeight: 600 }} tickLine={false} axisLine={false} />
                          <YAxis hide tickFormatter={(v) => fmtShort(v)} />
                          <Tooltip
                            formatter={(value: number, name: string) => [fmt(value), name === "income" ? "Income" : "Expenses"]}
                            contentStyle={{ fontSize: 13, borderRadius: 12, border: "1px solid #e2e8f0", boxShadow: "0 4px 12px rgba(0,0,0,0.1)" }}
                          />
                          <Bar dataKey="income" fill="#22c55e" radius={[4, 4, 0, 0]} />
                          <Bar dataKey="expenses" fill="#ef4444" radius={[4, 4, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                      <div className="flex items-center justify-center gap-5 mt-3">
                        <div className="flex items-center gap-2"><div className="w-3 h-3 rounded-full bg-green-500" /><span className="text-sm text-muted-foreground font-semibold">Income</span></div>
                        <div className="flex items-center gap-2"><div className="w-3 h-3 rounded-full bg-red-500" /><span className="text-sm text-muted-foreground font-semibold">Expenses</span></div>
                      </div>
                    </div>
                  )}

                  {/* Month-over-Month Comparison */}
                  {summary.momMonths && summary.momChanges.length > 0 && (
                    <div className="bg-white border border-border rounded-2xl p-5 shadow-sm animate-fade-in-up">
                      <div className="flex items-center justify-between mb-4">
                        <p className="text-sm font-bold text-muted-foreground uppercase tracking-widest">Month vs Month</p>
                        <div className="flex items-center gap-1.5">
                          <span className="text-xs font-semibold text-muted-foreground bg-muted px-2.5 py-1 rounded-full">{summary.momMonths[1]}</span>
                          <svg className="w-4 h-4 text-muted-foreground" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="9 18 15 12 9 6"/></svg>
                          <span className="text-xs font-bold text-primary bg-primary/10 px-2.5 py-1 rounded-full">{summary.momMonths[0]}</span>
                        </div>
                      </div>

                      {(() => {
                        const prevTotal = summary.momChanges.reduce((s, c) => s + c.previous, 0);
                        const curTotal = summary.momChanges.reduce((s, c) => s + c.current, 0);
                        const totalDelta = curTotal - prevTotal;
                        const totalPct = prevTotal > 0 ? Math.round((totalDelta / prevTotal) * 100) : 0;
                        const isUp = totalDelta > 0;
                        return (
                          <div className={`flex items-center justify-between rounded-xl px-4 py-3 mb-4 ${isUp ? "bg-red-50 border border-red-100" : "bg-green-50 border border-green-100"}`}>
                            <span className="text-sm font-semibold text-foreground">Total Expenses</span>
                            <div className="flex items-center gap-2">
                              <span className={`text-sm font-extrabold ${isUp ? "text-red-600" : "text-green-600"}`}>{fmt(curTotal)}</span>
                              <span className={`flex items-center gap-0.5 text-xs font-bold px-2 py-0.5 rounded-full ${isUp ? "bg-red-100 text-red-700" : "bg-green-100 text-green-700"}`}>
                                {isUp ? "▲" : "▼"} {Math.abs(totalPct)}%
                              </span>
                            </div>
                          </div>
                        );
                      })()}

                      <div className="space-y-3">
                        {summary.momChanges.slice(0, 6).map((c) => {
                          const isUp = c.delta > 0;
                          const isNew = c.isNew;
                          const isGone = c.isGone;
                          return (
                            <div key={c.category} className="flex items-center gap-3">
                              <span className={`text-xs px-2 py-0.5 rounded-lg shrink-0 ${c.className}`}>{c.category}</span>
                              <div className="flex-1 min-w-0">
                                {isNew ? (
                                  <div className="h-2 bg-muted rounded-full overflow-hidden">
                                    <div className="h-2 rounded-full bg-amber-400" style={{ width: "100%" }} />
                                  </div>
                                ) : isGone ? (
                                  <div className="h-2 bg-muted rounded-full" />
                                ) : (
                                  <div className="relative h-2 bg-muted rounded-full overflow-hidden">
                                    <div className="absolute inset-y-0 left-0 rounded-full" style={{ width: `${Math.min((Math.min(c.previous, c.current) / Math.max(c.previous, c.current)) * 100, 100)}%`, backgroundColor: c.color + "66" }} />
                                    <div className="absolute inset-y-0 rounded-full" style={{
                                      left: isUp ? `${Math.min((c.previous / Math.max(c.previous, c.current)) * 100, 100)}%` : `${Math.min((c.current / Math.max(c.previous, c.current)) * 100, 100)}%`,
                                      width: `${Math.abs(((c.delta) / Math.max(c.previous, c.current))) * 100}%`,
                                      backgroundColor: isUp ? "#ef4444" : "#22c55e",
                                    }} />
                                  </div>
                                )}
                              </div>
                              <div className="flex items-center gap-1.5 shrink-0">
                                {isNew ? (
                                  <span className="text-xs font-bold text-amber-700 bg-amber-100 px-1.5 py-0.5 rounded">NEW</span>
                                ) : isGone ? (
                                  <span className="text-xs font-bold text-muted-foreground bg-muted px-1.5 py-0.5 rounded">GONE</span>
                                ) : (
                                  <span className={`text-xs font-bold flex items-center gap-0.5 ${isUp ? "text-red-600" : "text-green-600"}`}>
                                    {isUp ? "▲" : "▼"}{Math.abs(c.pctChange)}%
                                  </span>
                                )}
                                <span className="text-xs font-semibold text-muted-foreground">{fmtShort(c.current)}</span>
                              </div>
                            </div>
                          );
                        })}
                      </div>

                      {summary.momChanges.length > 6 && (
                        <p className="text-xs text-muted-foreground text-center mt-3">
                          +{summary.momChanges.length - 6} more categories in the Categories tab
                        </p>
                      )}
                    </div>
                  )}

                  {/* Top Spending */}
                  <div className="bg-white border border-border rounded-2xl p-5 shadow-sm animate-fade-in-up">
                    <p className="text-sm font-bold text-muted-foreground uppercase tracking-widest mb-4">Top Spending</p>
                    {topCategories.slice(0, 4).map(([name, info]) => (
                      <div key={name} className="mb-4 last:mb-0">
                        <div className="flex justify-between items-center mb-2">
                          <span className={`text-sm px-2.5 py-0.5 rounded-lg ${info.className}`}>{name}</span>
                          <span className="text-base font-bold">{fmt(info.total)}</span>
                        </div>
                        <div className="h-2.5 bg-muted rounded-full overflow-hidden">
                          <div className="h-2.5 rounded-full transition-all duration-700" style={{ width: `${(info.total / maxCatTotal) * 100}%`, backgroundColor: info.color }} />
                        </div>
                      </div>
                    ))}
                  </div>

                  {/* Duplicates alert */}
                  {summary.duplicateRows.size > 0 && (
                    <div className="flex items-center justify-between bg-amber-50 border border-amber-200 rounded-2xl px-5 py-4 animate-fade-in-up">
                      <div className="flex items-center gap-3">
                        <span className="text-xl">⚠️</span>
                        <div>
                          <p className="text-sm font-bold text-amber-800">{summary.duplicateRows.size} possible duplicates</p>
                          <p className="text-xs text-amber-600">Check the Transactions tab</p>
                        </div>
                      </div>
                      {isPro
                        ? <ActionBtn size="sm" variant="secondary" onClick={() => { setActiveTab("transactions"); setShowDuplicates(true); }}>Review</ActionBtn>
                        : <ProBadge onClick={() => setShowPayment(true)} />
                      }
                    </div>
                  )}

                  <p className="text-sm text-muted-foreground font-medium text-right">
                    {summary.transactions.length} transactions analyzed
                  </p>
                </>
              )}

              {/* ── CATEGORIES TAB ── */}
              {activeTab === "categories" && (
                <div className="space-y-4">
                  {/* Top Merchants */}
                  <div className="bg-white border border-border rounded-2xl p-5 shadow-sm animate-fade-in-up">
                    <p className="text-sm font-bold text-muted-foreground uppercase tracking-widest mb-4">Top Merchants</p>
                    <div className="space-y-3">
                      {summary.topMerchants.slice(0, 5).map((m, i) => (
                        <div key={i} className="flex items-center gap-3">
                          <div className="w-8 h-8 rounded-xl bg-muted flex items-center justify-center shrink-0">
                            <span className="text-sm font-extrabold text-muted-foreground">#{i + 1}</span>
                          </div>
                          <div className="flex-1 min-w-0">
                            <p className="text-sm font-semibold truncate">{m.name}</p>
                            <span className={`text-xs px-1.5 py-0.5 rounded-md ${m.category.className}`}>{m.category.name}</span>
                          </div>
                          <div className="text-right shrink-0">
                            <p className="text-sm font-extrabold text-red-500">{fmtShort(m.total)}</p>
                            <p className="text-xs text-muted-foreground">{m.count}×</p>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Recurring Subscriptions */}
                  <div className="bg-white border border-border rounded-2xl p-5 shadow-sm animate-fade-in-up delay-100">
                    <div className="flex items-center justify-between mb-4">
                      <p className="text-sm font-bold text-muted-foreground uppercase tracking-widest">Recurring / Subscriptions</p>
                      {!isPro
                        ? <ProBadge onClick={() => setShowPayment(true)} />
                        : <ActionBtn size="sm" variant="ghost" onClick={() => setShowRecurring((v) => !v)}>
                            {showRecurring ? "Hide" : "Show"}
                          </ActionBtn>
                      }
                    </div>
                    {isPro && showRecurring && summary.recurring.length > 0 && (
                      <div className="space-y-3">
                        {summary.recurring.map((r, i) => (
                          <div key={i} className="flex items-center gap-3 py-2 border-b border-border last:border-0">
                            <div className="flex-1 min-w-0">
                              <p className="text-sm font-semibold truncate">{r.description}</p>
                              <div className="flex items-center gap-2 mt-0.5">
                                <span className={`text-xs px-1.5 py-0.5 rounded ${r.category.className}`}>{r.category.name}</span>
                                <span className="text-xs text-muted-foreground">{r.count}× detected</span>
                              </div>
                            </div>
                            <div className="text-right shrink-0">
                              <p className="text-sm font-extrabold text-red-500">{fmtShort(r.totalAmount)}</p>
                              <p className="text-xs text-muted-foreground">~{fmtShort(r.avgAmount)}/time</p>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                    {isPro && showRecurring && summary.recurring.length === 0 && (
                      <p className="text-sm text-muted-foreground text-center py-3">No recurring patterns detected.</p>
                    )}
                    {isPro && !showRecurring && (
                      <p className="text-sm text-muted-foreground">
                        {summary.recurring.length} recurring pattern{summary.recurring.length !== 1 ? "s" : ""} found — click Show to review.
                      </p>
                    )}
                    {!isPro && (
                      <p className="text-sm text-muted-foreground">Upgrade to see subscriptions & recurring charges auto-detected.</p>
                    )}
                  </div>

                  {/* Category breakdown */}
                  {topCategories.map(([name, info], idx) => {
                    const medal = idx === 0 ? "🥇" : idx === 1 ? "🥈" : idx === 2 ? "🥉" : null;
                    return (
                      <div key={name} className={`bg-white border rounded-2xl p-5 shadow-sm hover:shadow-md transition-all animate-fade-in-up ${idx < 3 ? "border-amber-200" : "border-border"}`}>
                        <div className="flex items-center justify-between mb-3">
                          <div className="flex items-center gap-2 min-w-0">
                            {medal && <span className="text-xl shrink-0 leading-none">{medal}</span>}
                            <span className={`text-sm px-2.5 py-0.5 rounded-lg shrink-0 ${info.className}`}>{name}</span>
                            <span className="text-xs font-bold px-2 py-0.5 rounded-full bg-muted text-muted-foreground shrink-0">
                              {info.count} tx
                            </span>
                          </div>
                          <span className="text-lg font-extrabold shrink-0">{fmt(info.total)}</span>
                        </div>
                        <div className="h-2.5 bg-muted rounded-full overflow-hidden mb-3">
                          <div className="h-2.5 rounded-full transition-all duration-700" style={{ width: `${(info.total / maxCatTotal) * 100}%`, backgroundColor: info.color }} />
                        </div>
                        <div className="flex items-center justify-between">
                          <p className="text-sm text-muted-foreground font-semibold">avg {fmt(Math.round(info.total / info.count))} / tx</p>
                          {summary.totalExpenses > 0 && (
                            <p className="text-sm text-muted-foreground font-semibold">{Math.round((info.total / summary.totalExpenses) * 100)}% of expenses</p>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* ── TRANSACTIONS TAB ── */}
              {activeTab === "transactions" && (
                <div className="space-y-3">
                  {/* Search & Filters */}
                  <div className="space-y-2.5">
                    <div className="relative">
                      <svg className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                        <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
                      </svg>
                      <input
                        type="text"
                        placeholder="Search transactions…"
                        value={txSearch}
                        onChange={(e) => setTxSearch(e.target.value)}
                        className="w-full rounded-xl border border-border bg-white pl-10 pr-4 py-3 text-sm focus:outline-none focus:border-primary transition-colors shadow-sm"
                      />
                    </div>
                    <div className="grid grid-cols-2 gap-2">
                      <select
                        value={txCategoryFilter}
                        onChange={(e) => setTxCategoryFilter(e.target.value)}
                        className="rounded-xl border border-border bg-white px-3 py-2.5 text-sm focus:outline-none focus:border-primary transition-colors shadow-sm"
                      >
                        {categoryNames.map((c) => <option key={c}>{c}</option>)}
                      </select>
                      <select
                        value={txTypeFilter}
                        onChange={(e) => setTxTypeFilter(e.target.value as "all" | "credit" | "debit")}
                        className="rounded-xl border border-border bg-white px-3 py-2.5 text-sm focus:outline-none focus:border-primary transition-colors shadow-sm"
                      >
                        <option value="all">All Types</option>
                        <option value="credit">Income only</option>
                        <option value="debit">Expenses only</option>
                      </select>
                    </div>
                    {/* Flagged filter */}
                    {flaggedRows.size > 0 && (
                      <div className="flex items-center justify-between bg-white border border-border rounded-xl px-4 py-3 shadow-sm">
                        <span className="text-sm font-semibold text-foreground">🚩 Show flagged only ({flaggedRows.size})</span>
                        <button
                          data-testid="toggle-flagged-filter"
                          onClick={() => setShowFlagged((v) => !v)}
                          className={`relative w-10 h-5 rounded-full transition-colors ${showFlagged ? "bg-red-400" : "bg-muted-foreground/30"}`}
                        >
                          <div className={`absolute top-0.5 w-4 h-4 bg-white rounded-full shadow transition-transform ${showFlagged ? "translate-x-5" : "translate-x-0.5"}`} />
                        </button>
                      </div>
                    )}
                    {/* Duplicate filter */}
                    {summary.duplicateRows.size > 0 && (
                      <div className="flex items-center justify-between bg-white border border-border rounded-xl px-4 py-3 shadow-sm">
                        <span className="text-sm font-semibold text-foreground">Show duplicates only ({summary.duplicateRows.size})</span>
                        {isPro ? (
                          <button
                            onClick={() => setShowDuplicates((v) => !v)}
                            className={`relative w-10 h-5 rounded-full transition-colors ${showDuplicates ? "bg-primary" : "bg-muted-foreground/30"}`}
                          >
                            <div className={`absolute top-0.5 w-4 h-4 bg-white rounded-full shadow transition-transform ${showDuplicates ? "translate-x-5" : "translate-x-0.5"}`} />
                          </button>
                        ) : (
                          <ProBadge onClick={() => setShowPayment(true)} />
                        )}
                      </div>
                    )}
                  </div>

                  {/* Transaction count */}
                  <p className="text-sm text-muted-foreground font-semibold">{filteredTxns.length} of {summary.transactions.length} transactions</p>

                  {/* Transaction list */}
                  {filteredTxns.length === 0 ? (
                    <div className="text-center py-12 text-muted-foreground animate-fade-in">
                      <p className="text-4xl mb-3">🔍</p>
                      <p className="text-base font-semibold">No transactions match</p>
                      <button onClick={() => { setTxSearch(""); setTxCategoryFilter("All"); setTxTypeFilter("all"); setShowDuplicates(false); }}
                        className="text-sm text-primary mt-3 hover:underline">Clear filters</button>
                    </div>
                  ) : (
                    filteredTxns.map((tx, i) => {
                      const isDupe = summary.duplicateRows.has(tx.row);
                      const isExpanded = expandedTxRow === tx.row;
                      const wasOverridden = !!categoryOverrides[tx.row];
                      const isFlagged = flaggedRows.has(tx.row);
                      const toggleFlag = () => setFlaggedRows((prev) => {
                        const next = new Set(prev);
                        if (next.has(tx.row)) next.delete(tx.row); else next.add(tx.row);
                        return next;
                      });
                      return (
                        <div key={i} className={`bg-white border rounded-2xl shadow-sm overflow-hidden transition-all hover:shadow-md ${isFlagged ? "border-red-300 bg-red-50/30" : isDupe && isPro ? "border-amber-300 bg-amber-50/40" : isExpanded ? "border-primary shadow-md" : "border-border"}`}>
                          <div className="flex items-start gap-3 p-4">
                            <div className={`w-2.5 h-2.5 rounded-full shrink-0 mt-1.5 ${tx.type === "credit" ? "bg-green-500" : "bg-red-500"}`} />
                            <div className="flex-1 min-w-0">
                              <p className="text-sm font-semibold truncate">{tx.description}</p>
                              <div className="flex flex-wrap items-center gap-1.5 mt-1.5">
                                <span className="text-xs text-muted-foreground font-medium">{tx.date}</span>
                                <button
                                  data-testid={`btn-recategorize-${tx.row}`}
                                  onClick={() => setExpandedTxRow(isExpanded ? null : tx.row)}
                                  title="Tap to change category"
                                  className={`text-xs px-2 py-0.5 rounded-lg transition-all hover:opacity-80 active:scale-95 ${tx.category.className} ${wasOverridden ? "ring-1 ring-offset-1 ring-primary/40" : ""}`}
                                >
                                  {tx.category.name} {isExpanded ? "▲" : "▼"}
                                </button>
                                {isFlagged && (
                                  <span className="text-xs font-bold bg-red-100 text-red-600 px-2 py-0.5 rounded-lg border border-red-200">🚩 Flagged</span>
                                )}
                                {isDupe && isPro && (
                                  <span className="text-xs font-bold bg-amber-100 text-amber-700 px-2 py-0.5 rounded-lg border border-amber-200">⚠ Duplicate</span>
                                )}
                              </div>
                            </div>
                            <div className="flex flex-col items-end gap-2 shrink-0">
                              <span className={`text-sm font-extrabold ${tx.type === "credit" ? "text-green-600" : "text-red-500"}`}>
                                {tx.type === "credit" ? "+" : "−"}{fmt(tx.amount)}
                              </span>
                              <button
                                data-testid={`btn-flag-${tx.row}`}
                                onClick={toggleFlag}
                                title={isFlagged ? "Remove flag" : "Flag for review"}
                                className={`text-lg leading-none transition-all hover:scale-125 active:scale-95 ${isFlagged ? "opacity-100" : "opacity-20 hover:opacity-60"}`}
                              >
                                🚩
                              </button>
                            </div>
                          </div>
                          {/* Category picker */}
                          {isExpanded && (
                            <div className="border-t border-border px-4 pb-4 pt-3 bg-muted/20 animate-fade-in">
                              <p className="text-xs font-bold text-muted-foreground uppercase tracking-wider mb-2.5">Change category</p>
                              <div className="flex flex-wrap gap-2">
                                {CATEGORIES.map((cat) => (
                                  <button
                                    key={cat.name}
                                    data-testid={`btn-cat-${cat.name.replace(/\s+/g, "-").toLowerCase()}-${tx.row}`}
                                    onClick={() => handleRecategorize(tx.row, cat)}
                                    className={`text-xs px-2.5 py-1.5 rounded-xl font-medium transition-all hover:opacity-80 active:scale-95 ${cat.className} ${tx.category.name === cat.name ? "ring-2 ring-offset-1 ring-foreground/30" : ""}`}
                                  >
                                    {cat.name}
                                  </button>
                                ))}
                                <button
                                  data-testid={`btn-cat-other-${tx.row}`}
                                  onClick={() => handleRecategorize(tx.row, { name: "Other", color: "#64748b", className: "cat-other", keywords: [], type: "any" })}
                                  className={`text-xs px-2.5 py-1.5 rounded-xl font-medium transition-all hover:opacity-80 active:scale-95 cat-other ${tx.category.name === "Other" ? "ring-2 ring-offset-1 ring-foreground/30" : ""}`}
                                >
                                  Other
                                </button>
                              </div>
                              {/* Notes field */}
                              <div className="mt-3">
                                <p className="text-xs font-bold text-muted-foreground uppercase tracking-wider mb-2">Note</p>
                                <input
                                  data-testid={`input-note-${tx.row}`}
                                  type="text"
                                  placeholder="Add a memo for this transaction…"
                                  maxLength={140}
                                  value={txNotes[tx.row] ?? ""}
                                  onChange={(e) => setTxNotes((prev) => ({ ...prev, [tx.row]: e.target.value }))}
                                  className="w-full rounded-xl border border-border bg-white px-3 py-2 text-sm focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50 shadow-sm"
                                />
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })
                  )}
                </div>
              )}

              {/* ── BUDGET TAB (PRO) ── */}
              {activeTab === "budget" && isPro && (
                <div className="space-y-4">
                  <div className="bg-gradient-to-r from-blue-50 to-indigo-50 border border-blue-100 rounded-2xl px-5 py-4 animate-fade-in-up">
                    <p className="text-base font-bold text-primary mb-1">💰 Category Budget Tracker</p>
                    <p className="text-sm text-muted-foreground">Set a monthly budget for each category and track your spending against it.</p>
                  </div>

                  {topCategories
                    .filter(([name]) => {
                      const cat = summary.byCategory[name];
                      return cat && name !== "Salary";
                    })
                    .map(([name, info], idx) => {
                      const budget = Number(budgets[name] || 0);
                      const spent = info.total;
                      const pct = budget > 0 ? Math.min((spent / budget) * 100, 100) : 0;
                      const over = budget > 0 && spent > budget;
                      const overAmt = over ? spent - budget : 0;
                      return (
                        <div key={name} className={`bg-white border rounded-2xl p-5 shadow-sm hover:shadow-md transition-all animate-fade-in-up`} style={{ animationDelay: `${idx * 50}ms` }}>
                          <div className="flex items-center justify-between mb-3">
                            <span className={`text-sm px-2.5 py-0.5 rounded-lg ${info.className}`}>{name}</span>
                            <div className="flex items-center gap-3">
                              <span className="text-base font-extrabold">{fmt(spent)}</span>
                              {over && <span className="text-xs font-bold bg-red-100 text-red-700 px-2 py-0.5 rounded-lg">Over by {fmt(overAmt)}</span>}
                            </div>
                          </div>
                          <div className="flex items-center gap-3 mb-3">
                            <span className="text-sm text-muted-foreground font-medium shrink-0">Budget:</span>
                            <div className="relative flex-1">
                              <span className="absolute left-3 top-1/2 -translate-y-1/2 text-sm text-muted-foreground font-medium">{symbol}</span>
                              <input
                                data-testid={`input-budget-${name.replace(/\s+/g, "-").toLowerCase()}`}
                                type="number"
                                min="0"
                                value={budgets[name] ?? ""}
                                onChange={(e) => setBudgets((prev) => ({ ...prev, [name]: e.target.value }))}
                                placeholder="Set budget…"
                                className="w-full rounded-xl border border-border bg-muted/30 pl-7 pr-3 py-2 text-sm focus:outline-none focus:border-primary transition-colors"
                              />
                            </div>
                          </div>
                          {budget > 0 && (
                            <>
                              <div className="h-3 bg-muted rounded-full overflow-hidden mb-1.5">
                                <div
                                  className={`h-3 rounded-full transition-all duration-700 ${over ? "bg-red-500" : pct > 80 ? "bg-amber-400" : "bg-green-500"}`}
                                  style={{ width: `${pct}%` }}
                                />
                              </div>
                              <p className="text-xs text-muted-foreground font-medium text-right">
                                {pct.toFixed(0)}% of budget used{!over && ` · ${fmt(budget - spent)} remaining`}
                              </p>
                            </>
                          )}
                        </div>
                      );
                    })}
                </div>
              )}

              {activeTab === "budget" && !isPro && (
                <div className="flex flex-col items-center justify-center py-16 text-center animate-scale-in">
                  <div className="w-20 h-20 rounded-3xl bg-gradient-to-br from-amber-100 to-orange-100 flex items-center justify-center mb-5 shadow-sm">
                    <span className="text-3xl">💰</span>
                  </div>
                  <h3 className="text-xl font-extrabold text-foreground mb-2">Budget Tracker is Pro</h3>
                  <p className="text-base text-muted-foreground mb-6 max-w-xs leading-relaxed">Set monthly budgets per category and track your spending in real time.</p>
                  <button onClick={() => { setPaymentMode("pay"); setShowPayment(true); }}
                    className="text-base font-bold bg-gradient-to-r from-amber-500 to-orange-500 text-white px-8 py-3 rounded-2xl shadow-lg shadow-amber-200 hover:from-amber-600 hover:to-orange-600 transition-all hover:scale-105 active:scale-95">
                    Upgrade to Pro
                  </button>
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {/* ── Footer nav (idle only) ── */}
      {step === "idle" && (
        <div className="shrink-0 border-t border-border bg-white px-5 py-3">
          <div className="flex items-center justify-center gap-4 text-xs text-muted-foreground">
            <a href="/eula" className="hover:text-foreground transition-colors">EULA</a>
            <span>·</span>
            <a href="/privacy" className="hover:text-foreground transition-colors">Privacy Policy</a>
            <span>·</span>
            <a href="/support" className="hover:text-foreground transition-colors">Support</a>
          </div>
          <p className="text-center text-xs text-muted-foreground/60 mt-1">{appName} · Excel Add-in</p>
        </div>
      )}
    </div>
  );
}
