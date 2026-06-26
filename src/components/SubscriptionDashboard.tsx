import { useState, useEffect, useCallback } from "react";
import { getLicense, clearLicense, fetchSubscription, activateLicenseKey, type SubscriptionStatus } from "../lib/payment";
import { useAppConfig } from "../context/AppConfigContext";

interface Props {
  onClose: () => void;
}

function Spinner() {
  return <span className="w-5 h-5 border-2 border-primary border-t-transparent rounded-full animate-spin inline-block shrink-0" />;
}

export default function SubscriptionDashboard({ onClose }: Props) {
  const config = useAppConfig();
  const plans = config.plans;

  const [status, setStatus] = useState<SubscriptionStatus | null>(null);
  const [loading, setLoading] = useState(true);
  const [renewInput, setRenewInput] = useState("");
  const [renewing, setRenewing] = useState(false);
  const [renewError, setRenewError] = useState("");
  const [renewSuccess, setRenewSuccess] = useState(false);
  const [showRenewForm, setShowRenewForm] = useState(false);
  const [confirmClear, setConfirmClear] = useState(false);

  const refresh = useCallback(async () => {
    setLoading(true);
    const key = getLicense();
    if (!key) { setStatus(null); setLoading(false); return; }
    const s = await fetchSubscription(key);
    setStatus(s);
    setLoading(false);
  }, []);

  useEffect(() => { refresh(); }, [refresh]);

  const handleRenew = useCallback(async () => {
    setRenewing(true); setRenewError("");
    const result = await activateLicenseKey(renewInput);
    setRenewing(false);
    if (result.success) {
      setRenewSuccess(true);
      setRenewInput(""); setShowRenewForm(false);
      setTimeout(() => { setRenewSuccess(false); refresh(); }, 1800);
    } else {
      setRenewError(result.error ?? "Invalid or expired key.");
    }
  }, [renewInput, refresh]);

  const handleClearLicense = () => {
    if (!confirmClear) { setConfirmClear(true); setTimeout(() => setConfirmClear(false), 4000); return; }
    clearLicense(); setStatus(null); setConfirmClear(false);
  };

  const planLabel = (planId: string | null) => {
    if (!planId) return "Pro";
    const plan = plans.find((p) => p.id === planId);
    return plan?.label ?? planId;
  };

  if (loading) {
    return (
      <div className="flex flex-col items-center justify-center py-20 gap-4">
        <div className="relative w-16 h-16">
          <div className="w-16 h-16 border-4 border-primary/20 rounded-full" />
          <div className="absolute inset-0 w-16 h-16 border-4 border-primary border-t-transparent rounded-full animate-spin" />
        </div>
        <p className="text-base text-muted-foreground font-medium">Loading subscription…</p>
      </div>
    );
  }

  if (!status || !status.valid) {
    return (
      <div className="flex flex-col min-h-full">
        {/* Hero */}
        <div className="bg-gradient-to-br from-slate-600 to-slate-800 px-6 py-10 text-center text-white">
          <div className="w-20 h-20 rounded-3xl bg-white/10 flex items-center justify-center mx-auto mb-4 shadow-lg">
            <svg className="w-10 h-10 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
              <rect x="3" y="11" width="18" height="11" rx="2"/>
              <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
            </svg>
          </div>
          <h2 className="text-2xl font-extrabold tracking-tight mb-2">
            {status?.reason === "expired" ? "Subscription Expired" : "No Active Subscription"}
          </h2>
          <p className="text-base text-slate-200 max-w-xs mx-auto leading-relaxed">
            {status?.reason === "expired"
              ? "Your license has expired. Renew to regain access to all Pro features."
              : "Unlock premium features to export reports, track budgets, and detect recurring charges."}
          </p>
        </div>

        <div className="p-5 space-y-4">
          {status?.reason === "expired" && (
            <div className="bg-amber-50 border-2 border-amber-200 rounded-2xl p-4 animate-fade-in-up">
              <p className="text-sm font-bold text-amber-800 mb-1">⚠ Expired Plan</p>
              <p className="text-sm text-amber-700">Was: <span className="font-semibold">{planLabel(status.planId)}</span></p>
              {status.expiresAt && (
                <p className="text-sm text-amber-600 mt-0.5">
                  Expired {new Date(status.expiresAt).toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" })}
                </p>
              )}
            </div>
          )}

          {!showRenewForm ? (
            <button
              onClick={() => setShowRenewForm(true)}
              className="w-full flex items-center justify-between bg-white border-2 border-border rounded-2xl px-5 py-4 hover:border-primary/40 transition-all group hover:shadow-sm animate-fade-in-up"
            >
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-primary/10 flex items-center justify-center">
                  <svg className="w-5 h-5 text-primary" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/>
                  </svg>
                </div>
                <span className="text-base font-semibold text-foreground">Already have a license key?</span>
              </div>
              <svg className="w-5 h-5 text-muted-foreground group-hover:text-primary group-hover:translate-x-0.5 transition-all" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <polyline points="9 18 15 12 9 6"/>
              </svg>
            </button>
          ) : (
            <div className="bg-white border-2 border-primary/30 rounded-2xl p-5 space-y-3 animate-scale-in">
              <p className="text-base font-bold text-foreground">Enter license key</p>
              <input
                type="text"
                value={renewInput}
                onChange={(e) => { setRenewInput(e.target.value.toUpperCase()); setRenewError(""); }}
                placeholder="BSA-XXXXXXXXXXXXXXXXXXXX"
                className="w-full rounded-xl border-2 border-border bg-muted/30 px-4 py-3 text-base font-mono focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/40 tracking-wider"
                spellCheck={false}
              />
              {renewError && (
                <div className="flex items-center gap-2 bg-red-50 border border-red-200 rounded-xl px-3 py-2.5">
                  <svg className="w-4 h-4 text-destructive shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                  </svg>
                  <p className="text-sm text-destructive font-medium">{renewError}</p>
                </div>
              )}
              <div className="flex gap-2.5">
                <button
                  onClick={() => { setShowRenewForm(false); setRenewInput(""); setRenewError(""); }}
                  className="flex-1 text-base font-semibold py-3 rounded-xl border-2 border-border hover:bg-muted transition-colors"
                >
                  Cancel
                </button>
                <button
                  onClick={handleRenew}
                  disabled={renewing || !renewInput.trim()}
                  className="flex-1 flex items-center justify-center gap-2 bg-gradient-to-r from-blue-600 to-blue-700 text-white text-base font-bold py-3 rounded-xl hover:from-blue-500 hover:to-blue-600 transition-all disabled:opacity-60 shadow-md shadow-blue-200"
                >
                  {renewing ? <><Spinner />Activating…</> : "Activate Key"}
                </button>
              </div>
            </div>
          )}

          <button onClick={onClose}
            className="w-full text-base font-semibold text-muted-foreground hover:text-foreground py-3 rounded-2xl hover:bg-muted transition-colors border border-border">
            ← Back to App
          </button>
        </div>
      </div>
    );
  }

  // ── Active subscription ──────────────────────────────────────────────────
  const daysLeft = status.daysLeft ?? 0;
  const daysTotal = status.daysTotal ?? 1;
  const pct = Math.max(0, Math.min(100, Math.round((daysLeft / daysTotal) * 100)));
  const isExpiringSoon = daysLeft > 0 && daysLeft <= 7;

  return (
    <div className="flex flex-col min-h-full">
      {/* Hero gradient */}
      <div className="bg-gradient-to-br from-blue-600 via-blue-700 to-indigo-800 px-6 py-8 text-white">
        <div className="flex items-center gap-3 mb-4">
          <div className="w-12 h-12 rounded-2xl bg-white/15 flex items-center justify-center shadow-inner">
            <svg className="w-6 h-6 text-amber-300" viewBox="0 0 24 24" fill="currentColor">
              <path d="M12 2l2.4 7.4H22l-6.2 4.5 2.4 7.4L12 17l-6.2 4.3 2.4-7.4L2 9.4h7.6z"/>
            </svg>
          </div>
          <div>
            <p className="text-xs font-bold text-amber-300 uppercase tracking-wider">Pro Active</p>
            <p className="text-xl font-extrabold text-white">{planLabel(status.planId)}</p>
          </div>
          <span className="ml-auto text-xs font-bold bg-white/20 px-3 py-1 rounded-full">Active</span>
        </div>

        {/* Days progress */}
        <div className="space-y-2">
          <div className="flex items-center justify-between text-sm">
            <span className="text-blue-100 font-medium">Time remaining</span>
            <span className={`font-extrabold text-base ${isExpiringSoon ? "text-amber-300" : "text-white"}`}>
              {daysLeft === 0 ? "Expires today" : `${daysLeft} day${daysLeft !== 1 ? "s" : ""} left`}
            </span>
          </div>
          <div className="h-3 bg-white/20 rounded-full overflow-hidden">
            <div className={`h-full rounded-full transition-all duration-700 ${pct > 30 ? "bg-white" : pct > 10 ? "bg-amber-300" : "bg-red-400"}`}
              style={{ width: `${pct}%` }} />
          </div>
          {status.expiresAt && (
            <p className="text-xs text-blue-200">
              Expires {new Date(status.expiresAt).toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" })}
            </p>
          )}
        </div>
      </div>

      <div className="p-5 space-y-4">
        {renewSuccess && (
          <div className="flex items-center gap-3 bg-green-50 border-2 border-green-200 rounded-2xl px-4 py-3.5 animate-scale-in">
            <svg className="w-5 h-5 text-green-600 shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></svg>
            <p className="text-base font-bold text-green-700">License updated successfully!</p>
          </div>
        )}

        {isExpiringSoon && (
          <div className="flex items-start gap-3 bg-amber-50 border-2 border-amber-200 rounded-2xl px-4 py-4 animate-fade-in-up">
            <svg className="w-5 h-5 text-amber-500 shrink-0 mt-0.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
            </svg>
            <div>
              <p className="text-sm font-bold text-amber-800">Expiring soon</p>
              <p className="text-sm text-amber-600 mt-0.5">Your plan expires in {daysLeft} day{daysLeft !== 1 ? "s" : ""}. Renew to keep access.</p>
            </div>
          </div>
        )}

        {/* Active features */}
        <div className="bg-gradient-to-br from-green-50 to-emerald-50 border border-green-200 rounded-2xl p-5 animate-fade-in-up">
          <p className="text-sm font-bold text-green-700 uppercase tracking-wider mb-3">Included in your plan</p>
          <div className="grid grid-cols-1 gap-2">
            {[
              "Analyze & auto-categorize transactions",
              "Export Excel summary sheet",
              "Download CSV with categories",
              "Color-coded PDF report",
              "Detect recurring subscriptions",
              "Category budget tracker",
            ].map((f) => (
              <div key={f} className="flex items-center gap-2.5">
                <svg className="w-4 h-4 text-green-500 shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></svg>
                <span className="text-sm text-green-800 font-medium">{f}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Renew / extend */}
        <div className="bg-white border-2 border-border rounded-2xl overflow-hidden animate-fade-in-up delay-100">
          <button onClick={() => setShowRenewForm((s) => !s)}
            className="w-full flex items-center justify-between px-5 py-4 text-left hover:bg-muted/30 transition-colors">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-primary/10 flex items-center justify-center">
                <svg className="w-5 h-5 text-primary" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/>
                </svg>
              </div>
              <span className="text-base font-semibold text-foreground">Enter a new license key</span>
            </div>
            <svg className={`w-5 h-5 text-muted-foreground transition-transform ${showRenewForm ? "rotate-180" : ""}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <polyline points="6 9 12 15 18 9"/>
            </svg>
          </button>

          {showRenewForm && (
            <div className="px-5 pb-5 pt-3 border-t border-border bg-muted/20 space-y-3">
              <p className="text-sm text-muted-foreground">Enter a new key to extend or switch your plan.</p>
              <input
                type="text"
                value={renewInput}
                onChange={(e) => { setRenewInput(e.target.value.toUpperCase()); setRenewError(""); }}
                placeholder="BSA-XXXXXXXXXXXXXXXXXXXX"
                className="w-full rounded-xl border-2 border-border bg-white px-4 py-3 text-base font-mono focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/40 tracking-wider"
                spellCheck={false}
              />
              {renewError && (
                <p className="text-sm text-destructive bg-red-50 border border-red-200 rounded-xl px-3 py-2">{renewError}</p>
              )}
              <div className="flex gap-2.5">
                <button
                  onClick={() => { setShowRenewForm(false); setRenewInput(""); setRenewError(""); }}
                  className="flex-1 text-base font-semibold py-3 rounded-xl border-2 border-border hover:bg-muted transition-colors"
                >
                  Cancel
                </button>
                <button
                  onClick={handleRenew}
                  disabled={renewing || !renewInput.trim()}
                  className="flex-1 flex items-center justify-center gap-2 bg-gradient-to-r from-blue-600 to-blue-700 text-white text-base font-bold py-3 rounded-xl hover:from-blue-500 hover:to-blue-600 transition-all disabled:opacity-60 shadow-md shadow-blue-200"
                >
                  {renewing ? <><Spinner />Activating…</> : "Activate"}
                </button>
              </div>
            </div>
          )}
        </div>

        <button onClick={onClose}
          className="w-full text-base font-semibold text-muted-foreground hover:text-foreground py-3 rounded-2xl hover:bg-muted transition-colors border border-border">
          ← Back to App
        </button>

        {/* Remove license */}
        <div className="pt-1 border-t border-border text-center">
          <button onClick={handleClearLicense}
            className={`text-sm transition-colors py-2 ${confirmClear ? "text-destructive font-bold" : "text-muted-foreground hover:text-destructive"}`}>
            {confirmClear ? "⚠ Tap again to confirm — this will remove your license" : "Remove license from this device"}
          </button>
        </div>
      </div>
    </div>
  );
}
