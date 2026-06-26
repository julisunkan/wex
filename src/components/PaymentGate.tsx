import { useState, useEffect, useCallback } from "react";
import { fetchPaymentConfig, verifyPayment, setLicense, activateLicenseKey } from "../lib/payment";
import { useAppConfig } from "../context/AppConfigContext";
import iconExport from "@assets/icons/icon-export.png";

interface Props {
  onUnlocked: () => void;
  onDismiss: () => void;
  initialMode?: "pay" | "key";
}

type PayMode = "pay" | "key";
type PayStep = "info" | "verifying" | "error" | "success";

function Spinner({ color = "border-white" }: { color?: string }) {
  return <span className={`w-5 h-5 border-2 ${color} border-t-transparent rounded-full animate-spin inline-block shrink-0`} />;
}

export default function PaymentGate({ onUnlocked, onDismiss, initialMode = "pay" }: Props) {
  const config = useAppConfig();
  const plans = config.plans;

  const [mode, setMode] = useState<PayMode>(initialMode);
  const [selectedPlanId, setSelectedPlanId] = useState<string>(() => plans[0]?.id ?? "monthly");
  const [wallet, setWallet] = useState<string>("");
  const [networkLabel, setNetworkLabel] = useState<string>("TRC-20 (Tron)");
  const [txHash, setTxHash] = useState("");
  const [userEmail, setUserEmail] = useState("");
  const [payStep, setPayStep] = useState<PayStep>("info");
  const [payError, setPayError] = useState("");
  const [copied, setCopied] = useState(false);
  const [generatedKey, setGeneratedKey] = useState("");
  const [keyCopied, setKeyCopied] = useState(false);
  const [licenseInput, setLicenseInput] = useState("");
  const [keyActivating, setKeyActivating] = useState(false);
  const [keyError, setKeyError] = useState("");
  const [keySuccess, setKeySuccess] = useState(false);

  useEffect(() => {
    fetchPaymentConfig().then((cfg) => {
      if (cfg) {
        setWallet(cfg.address);
        const label = cfg.network === "tron" ? "TRC-20 (Tron)"
          : cfg.network === "bsc" ? "BEP-20 (BSC)"
          : cfg.network === "eth" ? "ERC-20 (ETH)"
          : cfg.network;
        setNetworkLabel(label);
      }
    });
  }, []);

  useEffect(() => {
    if (plans.length > 0 && !plans.find((p) => p.id === selectedPlanId)) {
      setSelectedPlanId(plans[0].id);
    }
  }, [plans, selectedPlanId]);

  const selectedPlan = plans.find((p) => p.id === selectedPlanId) ?? plans[0];

  const copyWallet = useCallback(() => {
    navigator.clipboard.writeText(wallet).catch(() => {});
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }, [wallet]);

  const copyGeneratedKey = useCallback(() => {
    navigator.clipboard.writeText(generatedKey).catch(() => {});
    setKeyCopied(true);
    setTimeout(() => setKeyCopied(false), 2000);
  }, [generatedKey]);

  const handleVerifyPayment = useCallback(async () => {
    const hash = txHash.trim();
    if (!hash) { setPayError("Please paste your transaction hash."); return; }
    setPayStep("verifying"); setPayError("");
    const email = userEmail.trim().includes("@") ? userEmail.trim() : undefined;
    const result = await verifyPayment(hash, selectedPlan?.id ?? "monthly", email);
    if (result.success && result.licenseKey) {
      setGeneratedKey(result.licenseKey);
      setLicense(result.licenseKey);
      setPayStep("success");
      setTimeout(onUnlocked, 4000);
    } else {
      setPayError(result.error ?? "Verification failed.");
      setPayStep("error");
    }
  }, [txHash, selectedPlan, onUnlocked, userEmail]);

  const handleActivateKey = useCallback(async () => {
    setKeyActivating(true); setKeyError("");
    const result = await activateLicenseKey(licenseInput);
    setKeyActivating(false);
    if (result.success) { setKeySuccess(true); setTimeout(onUnlocked, 1500); }
    else setKeyError(result.error ?? "Activation failed.");
  }, [licenseInput, onUnlocked]);

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-black/50 backdrop-blur-sm animate-fade-in">
      <div className="w-full max-w-lg bg-background rounded-t-3xl shadow-2xl overflow-hidden animate-fade-in-up">

        {/* ── Gradient header ── */}
        <div className="bg-gradient-to-r from-blue-600 via-blue-700 to-indigo-700 px-6 py-5 flex items-center justify-between">
          <div>
            <p className="text-white font-extrabold text-xl leading-tight">Unlock Premium Features</p>
            <p className="text-blue-100 text-sm mt-0.5">Choose a plan · Cancel anytime</p>
          </div>
          <button onClick={onDismiss} data-testid="button-dismiss-payment"
            className="text-white/70 hover:text-white transition-colors p-2 hover:bg-white/10 rounded-xl">
            <svg className="w-6 h-6" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
              <line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" />
            </svg>
          </button>
        </div>

        {/* ── Mode tabs ── */}
        <div className="flex border-b border-border bg-muted/30">
          {(["pay", "key"] as const).map((m) => (
            <button key={m} onClick={() => setMode(m)}
              className={`flex-1 py-3.5 text-sm font-bold transition-all border-b-2 ${
                mode === m
                  ? "text-primary border-primary bg-background"
                  : "text-muted-foreground border-transparent hover:text-foreground"
              }`}>
              {m === "pay" ? "💳 Pay with USDT" : "🔑 Enter License Key"}
            </button>
          ))}
        </div>

        <div className="p-5 space-y-5 max-h-[75vh] overflow-y-auto">

          {/* ── PAY MODE ── */}
          {mode === "pay" && (
            <>
              {/* Feature cards */}
              <div className="grid grid-cols-2 gap-3">
                {[
                  { icon: iconExport, label: "Export Reports", desc: "Excel · CSV · PDF", grad: "from-indigo-50 to-indigo-100 border-indigo-200" },
                  { label: "Budget Tracker", desc: "Set category budgets", grad: "from-emerald-50 to-emerald-100 border-emerald-200", emoji: "💰" },
                  { label: "Detect Duplicates", desc: "Flag repeat charges", grad: "from-amber-50 to-amber-100 border-amber-200", emoji: "🔍" },
                ].map((f) => (
                  <div key={f.label} className={`bg-gradient-to-br ${f.grad} rounded-2xl p-4 border hover:shadow-sm transition-all`}>
                    {f.icon
                      ? <img src={f.icon} alt={f.label} className="w-9 h-9 object-contain mb-2" />
                      : <span className="text-2xl mb-2 block">{f.emoji}</span>}
                    <p className="text-sm font-bold text-foreground">{f.label}</p>
                    <p className="text-xs text-muted-foreground mt-0.5">{f.desc}</p>
                  </div>
                ))}
              </div>

              {payStep === "success" ? (
                <div className="flex flex-col items-center gap-4 py-4 text-center animate-scale-in">
                  <div className="w-20 h-20 rounded-3xl bg-green-100 flex items-center justify-center shadow-sm">
                    <svg className="w-10 h-10 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <polyline points="20 6 9 17 4 12" />
                    </svg>
                  </div>
                  <div>
                    <p className="text-xl font-extrabold text-green-700">Payment Verified!</p>
                    <p className="text-sm text-muted-foreground mt-1">Unlocking premium features…</p>
                  </div>
                  {generatedKey && (
                    <div className="w-full bg-amber-50 border-2 border-amber-200 rounded-2xl p-4">
                      <p className="text-xs font-bold text-amber-700 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                        <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/>
                        </svg>
                        Your License Key — Save this!
                      </p>
                      <div className="flex items-center gap-2 bg-white border border-amber-200 rounded-xl px-3 py-2.5">
                        <code className="flex-1 text-sm font-mono font-bold text-foreground tracking-wider select-all break-all">{generatedKey}</code>
                        <button onClick={copyGeneratedKey}
                          className="shrink-0 text-sm font-bold px-3 py-1.5 rounded-xl bg-amber-500 text-white hover:bg-amber-600 transition-colors">
                          {keyCopied ? "✓" : "Copy"}
                        </button>
                      </div>
                      <p className="text-xs text-amber-600 mt-2">Keep this safe — you'll need it on other devices.</p>
                    </div>
                  )}
                </div>
              ) : (
                <>
                  {/* Plan selector */}
                  {plans.length > 0 && (
                    <div>
                      <p className="text-sm font-bold text-foreground mb-3">Choose a plan</p>
                      <div className="grid grid-cols-2 gap-2.5">
                        {plans.map((plan) => {
                          const isSelected = selectedPlanId === plan.id;
                          return (
                            <button key={plan.id} onClick={() => setSelectedPlanId(plan.id)}
                              className={`relative rounded-2xl border-2 p-4 text-left transition-all hover:shadow-sm ${
                                isSelected
                                  ? "border-primary bg-primary/5 shadow-sm"
                                  : "border-border bg-white hover:border-primary/40"
                              }`}>
                              {plan.id === "annual" && (
                                <span className="absolute -top-2.5 right-3 text-xs font-bold bg-green-500 text-white px-2 py-0.5 rounded-full shadow">BEST VALUE</span>
                              )}
                              {plan.id === "quarterly" && (
                                <span className="absolute -top-2.5 right-3 text-xs font-bold bg-blue-500 text-white px-2 py-0.5 rounded-full shadow">POPULAR</span>
                              )}
                              <p className={`text-sm font-bold ${isSelected ? "text-primary" : "text-foreground"}`}>{plan.label}</p>
                              <p className={`text-2xl font-extrabold mt-1 ${isSelected ? "text-primary" : "text-foreground"}`}>
                                ${plan.price}
                                <span className="text-sm font-medium text-muted-foreground ml-1">USDT</span>
                              </p>
                              <p className="text-xs text-muted-foreground mt-0.5">{plan.days} days access</p>
                              {isSelected && (
                                <div className="absolute top-3 right-3 w-5 h-5 rounded-full bg-primary flex items-center justify-center shadow-sm">
                                  <svg className="w-3 h-3 text-white" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3"><polyline points="20 6 9 17 4 12"/></svg>
                                </div>
                              )}
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {/* Step 1 — Send USDT */}
                  <div className="bg-muted/30 rounded-2xl p-4 space-y-2 border border-border">
                    <p className="text-sm font-bold text-foreground flex items-center gap-2">
                      <span className="w-6 h-6 rounded-full bg-primary text-primary-foreground text-xs flex items-center justify-center font-bold shrink-0">1</span>
                      Send exactly <span className="text-primary">${selectedPlan?.price ?? "—"} USDT</span> to this wallet
                    </p>
                    <div className="flex items-center gap-2 bg-white border border-border rounded-xl p-3">
                      <p className="flex-1 text-sm font-mono text-foreground break-all leading-relaxed">
                        {wallet || "Loading wallet address…"}
                      </p>
                      {wallet && (
                        <button onClick={copyWallet}
                          className="shrink-0 text-sm font-bold px-3 py-1.5 rounded-xl bg-primary/10 text-primary hover:bg-primary/20 transition-colors whitespace-nowrap">
                          {copied ? "✓ Copied" : "Copy"}
                        </button>
                      )}
                    </div>
                    <p className="text-xs text-muted-foreground">
                      Network: <span className="font-semibold text-foreground">{networkLabel}</span> · Use any crypto wallet or exchange
                    </p>
                  </div>

                  {/* Step 2 — TX hash */}
                  <div className="space-y-2">
                    <p className="text-sm font-bold text-foreground flex items-center gap-2">
                      <span className="w-6 h-6 rounded-full bg-primary text-primary-foreground text-xs flex items-center justify-center font-bold shrink-0">2</span>
                      Paste your transaction hash
                    </p>
                    <input
                      type="text"
                      value={txHash}
                      onChange={(e) => { setTxHash(e.target.value); setPayError(""); setPayStep("info"); }}
                      placeholder="e.g. 3f8a2b1c4d…"
                      className="w-full rounded-xl border-2 border-border bg-white px-4 py-3 text-sm font-mono focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50"
                      data-testid="input-tx-hash"
                    />
                    {(payStep === "error" || payError) && (
                      <div className="flex items-center gap-2 bg-red-50 border border-red-200 rounded-xl px-3 py-2.5">
                        <svg className="w-4 h-4 text-destructive shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                        </svg>
                        <p className="text-sm text-destructive font-medium">{payError}</p>
                      </div>
                    )}
                  </div>

                  {/* Step 3 — Email (optional) */}
                  <div className="space-y-2">
                    <p className="text-sm font-bold text-foreground flex items-center gap-2">
                      <span className="w-6 h-6 rounded-full bg-muted border border-border text-muted-foreground text-xs flex items-center justify-center font-bold shrink-0">3</span>
                      Your email
                      <span className="text-muted-foreground font-normal text-xs">(optional — renewal reminders)</span>
                    </p>
                    <input
                      type="email"
                      value={userEmail}
                      onChange={(e) => setUserEmail(e.target.value)}
                      placeholder="you@example.com"
                      className="w-full rounded-xl border-2 border-border bg-white px-4 py-3 text-sm focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/50"
                      data-testid="input-user-email"
                    />
                  </div>

                  <button
                    onClick={payStep === "verifying" ? undefined : handleVerifyPayment}
                    disabled={payStep === "verifying" || !txHash.trim()}
                    className="w-full flex items-center justify-center gap-2.5 bg-gradient-to-r from-blue-600 to-blue-700 text-white text-base font-bold py-4 rounded-2xl hover:from-blue-500 hover:to-blue-600 transition-all disabled:opacity-60 shadow-lg shadow-blue-200 hover:scale-[1.01] active:scale-[0.99]"
                    data-testid="button-verify-payment"
                  >
                    {payStep === "verifying" ? (
                      <><Spinner />Verifying on blockchain…</>
                    ) : (
                      <>{selectedPlan ? `Verify & Unlock — $${selectedPlan.price} USDT` : "Verify Payment"}</>
                    )}
                  </button>

                  <p className="text-sm text-center text-muted-foreground">
                    Already have a key?{" "}
                    <button onClick={() => setMode("key")} className="text-primary font-semibold hover:underline">
                      Enter license key instead
                    </button>
                  </p>
                </>
              )}
            </>
          )}

          {/* ── KEY MODE ── */}
          {mode === "key" && (
            <>
              {keySuccess ? (
                <div className="flex flex-col items-center gap-4 py-8 text-center animate-scale-in">
                  <div className="w-20 h-20 rounded-3xl bg-green-100 flex items-center justify-center shadow-sm">
                    <svg className="w-10 h-10 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <polyline points="20 6 9 17 4 12" />
                    </svg>
                  </div>
                  <div>
                    <p className="text-xl font-extrabold text-green-700">License Activated!</p>
                    <p className="text-sm text-muted-foreground mt-1">Unlocking premium features…</p>
                  </div>
                </div>
              ) : (
                <>
                  <div className="flex flex-col items-center gap-4 py-4 text-center">
                    <div className="w-20 h-20 rounded-3xl bg-gradient-to-br from-blue-100 to-indigo-100 flex items-center justify-center shadow-sm">
                      <svg className="w-10 h-10 text-blue-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/>
                      </svg>
                    </div>
                    <div>
                      <p className="text-xl font-extrabold text-foreground">Enter Your License Key</p>
                      <p className="text-sm text-muted-foreground mt-1.5 max-w-xs">
                        Enter the license key you received after payment to activate Pro features.
                      </p>
                    </div>
                  </div>

                  <div className="space-y-3">
                    <input
                      type="text"
                      value={licenseInput}
                      onChange={(e) => { setLicenseInput(e.target.value.toUpperCase()); setKeyError(""); }}
                      placeholder="BSA-XXXXXXXXXXXXXXXXXXXX"
                      className="w-full rounded-xl border-2 border-border bg-white px-4 py-3.5 text-base font-mono focus:outline-none focus:border-primary transition-colors placeholder:text-muted-foreground/40 tracking-wider"
                      spellCheck={false}
                      autoCapitalize="characters"
                      data-testid="input-license-key"
                    />
                    {keyError && (
                      <div className="flex items-center gap-2 bg-red-50 border border-red-200 rounded-xl px-3 py-2.5">
                        <svg className="w-4 h-4 text-destructive shrink-0" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                        </svg>
                        <p className="text-sm text-destructive font-medium">{keyError}</p>
                      </div>
                    )}
                  </div>

                  <button
                    onClick={handleActivateKey}
                    disabled={keyActivating || !licenseInput.trim()}
                    className="w-full flex items-center justify-center gap-2.5 bg-gradient-to-r from-blue-600 to-blue-700 text-white text-base font-bold py-4 rounded-2xl hover:from-blue-500 hover:to-blue-600 transition-all disabled:opacity-60 shadow-lg shadow-blue-200 hover:scale-[1.01] active:scale-[0.99]"
                    data-testid="button-activate-key"
                  >
                    {keyActivating ? <><Spinner />Validating…</> : "Activate License Key →"}
                  </button>

                  <p className="text-sm text-center text-muted-foreground">
                    Don't have a key yet?{" "}
                    <button onClick={() => setMode("pay")} className="text-primary font-semibold hover:underline">
                      Pay with USDT
                    </button>
                  </p>
                </>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
