"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Activity, Check, Monitor, Zap, ShieldCheck, XCircle } from "lucide-react";
import { useState, useEffect } from "react";
import { useAppData } from "@/context/AppDataContext";
import { useAuth } from "@/context/AuthContext";
import { useRouter } from "next/navigation";
import Script from "next/script";
import { db } from "@/lib/firebase";
import { collection, addDoc, serverTimestamp } from "firebase/firestore";

const packages = [
  { id: "single", pcs: 1, priceINR: 300, extra: 0, title: "Single PC", desc: "Perfect for individual users.", popular: false },
  { id: "pack-5", pcs: 5, priceINR: 1400, extra: 0, title: "5 PC Package", desc: "Great for small teams.", popular: false },
  { id: "pack-10", pcs: 10, priceINR: 2700, extra: 1, title: "10 PC Package", desc: "For growing businesses.", popular: true },
  { id: "pack-30", pcs: 30, priceINR: 8500, extra: 5, title: "30 PC Package", desc: "For large organizations.", popular: false },
  { id: "pack-60", pcs: 60, priceINR: 16000, extra: 10, title: "60 PC Package", desc: "Enterprise scale deployment.", popular: false },
  { id: "pack-100", pcs: 100, priceINR: 28000, extra: 20, title: "100 PC Package", desc: "Maximum value for massive teams.", popular: false },
];

export default function PricingPage() {
  const { pricing: firebasePricing, loading: configLoading } = useAppData();
  const { user, loading: authLoading, subscriptions } = useAuth();
  const router = useRouter();
  const [currency, setCurrency] = useState("INR");
  const [rate, setRate] = useState(1);
  const [currencyLoading, setCurrencyLoading] = useState(true);
  const [processing, setProcessing] = useState<string | null>(null);
  const [paymentStatus, setPaymentStatus] = useState<'idle' | 'success' | 'error'>('idle');
  const [paymentMessage, setPaymentMessage] = useState<string>('');

  useEffect(() => {
    const fetchLocalization = async () => {
      try {
        // Prevent TypeError from bubbling if adblock/network blocks the request
        const geoRes = await fetch("https://ipapi.co/json/").catch(() => null);
        if (geoRes && geoRes.ok) {
          const geoData = await geoRes.json();
          
          if (geoData.currency) {
            setCurrency(geoData.currency);
            // Fetch exchange rates base INR
            const rateRes = await fetch("https://api.exchangerate-api.com/v4/latest/INR").catch(() => null);
            if (rateRes && rateRes.ok) {
              const rateData = await rateRes.json();
              if (rateData.rates && rateData.rates[geoData.currency]) {
                setRate(rateData.rates[geoData.currency]);
              }
            }
          }
        }
      } catch (error) {
        console.warn("Localization fetch failed silently.");
      } finally {
        setCurrencyLoading(false);
      }
    };
    fetchLocalization();
  }, []);

  const formatPriceLocal = (localPrice: number) => {
    return new Intl.NumberFormat(undefined, {
      style: "currency",
      currency: currency,
      maximumFractionDigits: 0,
    }).format(localPrice);
  };

  const formatPrice = (priceINR: number) => {
    const localPrice = priceINR * rate;
    return formatPriceLocal(localPrice);
  };

  // --- Proration Logic ---
  let highestActiveSub: any = null;
  let remainingDays = 0;
  let discountLocal = 0;

  if (subscriptions && subscriptions.length > 0) {
    const activeSubs = subscriptions.filter(s => s.active);
    if (activeSubs.length > 0) {
      highestActiveSub = activeSubs.sort((a, b) => (b.price || 0) - (a.price || 0))[0];
      if (highestActiveSub.nextPayment) {
        const nextPay = new Date(highestActiveSub.nextPayment);
        const now = new Date();
        const diffTime = nextPay.getTime() - now.getTime();
        remainingDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        if (remainingDays > 0 && highestActiveSub.price) {
          discountLocal = (highestActiveSub.price / 30) * remainingDays;
        }
      }
    }
  }

  const loadRazorpayScript = () => {
    return new Promise((resolve) => {
      if ((window as any).Razorpay) {
        resolve(true);
        return;
      }
      const script = document.createElement("script");
      script.src = "https://checkout.razorpay.com/v1/checkout.js";
      script.onload = () => resolve(true);
      script.onerror = () => resolve(false);
      document.body.appendChild(script);
    });
  };

  const handleCheckout = async (pkg: any, dynamicPriceINR: number, isUpgrade: boolean, finalPriceLocal: number) => {
    if (!user) {
      router.push("/login?redirect=/pricing");
      return;
    }
    
    setProcessing(pkg.id);
    try {
      // Ensure Razorpay script is loaded
      const isLoaded = await loadRazorpayScript();
      if (!isLoaded) {
        throw new Error("Razorpay SDK failed to load. Are you offline?");
      }

      const res = await fetch("/api/razorpay/create-order", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ amount: finalPriceLocal, currency: currency }),
      });
      const order = await res.json();

      if (!order.id || !order.key) {
        throw new Error(order.error || "Failed to create order. Please check Razorpay keys.");
      }

      const options = {
        key: order.key, // Now strictly taking the key from backend response instead of client-side ENV
        amount: order.amount,
        currency: order.currency,
        name: "Activity Tracker",
        description: `${pkg.title} Subscription`,
        order_id: order.id,
        handler: async function (response: any) {
          try {
            const verifyRes = await fetch("/api/razorpay/verify", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                razorpay_order_id: response.razorpay_order_id,
                razorpay_payment_id: response.razorpay_payment_id,
                razorpay_signature: response.razorpay_signature,
                userId: user.uid, // Use Firebase Auth UID for reliable querying
                upgradeFromId: isUpgrade ? highestActiveSub.id : null,
                planDetails: {
                  planId: pkg.id,
                  planName: pkg.title,
                  pcs: pkg.pcs,
                  bonus: pkg.extra,
                  price: finalPriceLocal,
                  currency: currency,
                }
              }),
            });
            const verifyData = await verifyRes.json();
            if (verifyData.success) {
              setPaymentStatus('success');
              setPaymentMessage('Your subscription has been activated successfully!');
            } else {
              setPaymentStatus('error');
              setPaymentMessage(verifyData.error || 'Payment verification failed');
            }
          } catch (error: any) {
            console.error(error);
            setPaymentStatus('error');
            setPaymentMessage(error?.message || 'Payment verification failed');
          }
        },
        prefill: {
          email: user.email,
        },
        theme: {
          color: "#3b82f6",
        },
      };

      const rzp = new (window as any).Razorpay(options);
      rzp.on('payment.failed', async function (response: any){
        console.error("Razorpay Payment Failed:", response.error);
        
        // Use Firestore as a Web alternative for Crashlytics tracking
        try {
          await addDoc(collection(db, "payment_errors_crashlytics"), {
            errorInfo: response.error,
            userId: user.uid,
            userEmail: user.email,
            planId: pkg.id,
            amountFailed: finalPriceLocal,
            currency: currency,
            timestamp: serverTimestamp(),
            userAgent: window.navigator.userAgent
          });
        } catch (dbErr) {
          console.error("Failed to log payment error to Firestore", dbErr);
        }

        setPaymentStatus('error');
        setPaymentMessage(response.error?.description || response.error?.reason || "Payment failed");
      });
      rzp.open();
    } catch (error: any) {
      console.error(error);
      setPaymentStatus('error');
      setPaymentMessage(error?.message || "Checkout failed");
    } finally {
      setProcessing(null);
    }
  };

  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24">
      <Script src="https://checkout.razorpay.com/v1/checkout.js" strategy="lazyOnload" />
      <Navbar />
      
      <section className="relative py-20 px-6">
        <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[1000px] h-[600px] bg-primary/10 blur-[150px] rounded-full pointer-events-none" />
        
        <div className="container mx-auto max-w-6xl relative z-10 text-center mb-20">
          <motion.div
            initial={{ opacity: 0, scale: 0.9 }}
            animate={{ opacity: 1, scale: 1 }}
            transition={{ duration: 0.5 }}
            className="inline-flex items-center gap-2 px-4 py-2 rounded-full bg-white/5 border border-white/10 text-sm font-semibold text-primary mb-8"
          >
            <Zap size={16} className="text-accent" />
            <span>Simple Monthly Pricing</span>
          </motion.div>

          {highestActiveSub && (
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-primary/10 border border-primary/30 p-4 rounded-2xl max-w-xl mx-auto mb-8 text-primary font-medium"
            >
              You currently have the <span className="font-bold text-white">{highestActiveSub.planName}</span> active. 
              {remainingDays > 0 && ` You have ${remainingDays} days remaining, which gives you a prorated discount on higher tier upgrades!`}
            </motion.div>
          )}

          <motion.h1
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, type: "spring", bounce: 0.4 }}
            className="text-5xl md:text-7xl font-black tracking-tight mb-6"
          >
            Pricing designed for <br className="hidden md:block" />
            <span className="text-transparent bg-clip-text bg-gradient-to-r from-blue-400 via-purple-400 to-accent">Growth & Scale</span>
          </motion.h1>

          <motion.p
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, type: "spring", bounce: 0.4, delay: 0.1 }}
            className="text-xl text-muted-foreground max-w-2xl mx-auto"
          >
            Purchase licenses in bulk and receive extra PCs entirely free. Localized pricing applied automatically.
            <br/><br/>
            <span className="text-sm border border-white/10 bg-white/5 px-3 py-1 rounded-lg">This is not an auto-debit subscription. You need to pay manually every month to continue using it.</span>
          </motion.p>
        </div>

        <div className="container mx-auto max-w-7xl relative z-10">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 lg:gap-8">
            {packages.map((pkg, index) => {
              // Determine unique styling based on tier
              const isEntry = pkg.pcs <= 5;
              const isPro = pkg.pcs > 5 && pkg.pcs <= 30;
              const isEnterprise = pkg.pcs >= 60;

              // Dynamically apply firebase pricing if available, otherwise fallback to local pkg.priceINR
              const dynamicPriceINR = firebasePricing && firebasePricing[pkg.id] !== undefined ? firebasePricing[pkg.id] : pkg.priceINR;

              let cardBg = "bg-[#0a0a0a]";
              let borderGlow = "hover:border-white/20 hover:shadow-[0_0_40px_rgba(255,255,255,0.05)]";
              let badgeColor = "bg-white/10 text-white";
              let buttonStyle = "bg-white/5 hover:bg-white/10 text-white border border-white/10";
              let iconColor = "text-muted-foreground";

              if (pkg.popular || isPro) {
                cardBg = "bg-gradient-to-b from-[#111] to-[#0a0a0a]";
                borderGlow = "border-primary/50 shadow-[0_0_50px_rgba(59,130,246,0.15)]";
                badgeColor = "bg-primary text-white shadow-[0_0_20px_rgba(59,130,246,0.5)]";
                buttonStyle = "bg-primary hover:bg-primary/90 text-white shadow-[0_0_20px_rgba(59,130,246,0.3)]";
                iconColor = "text-primary";
              }

              if (isEnterprise) {
                cardBg = "bg-gradient-to-b from-[#150a1f] to-[#0a0a0a]";
                borderGlow = "border-purple-500/50 hover:shadow-[0_0_50px_rgba(168,85,247,0.2)]";
                badgeColor = "bg-purple-500 text-white";
                buttonStyle = "bg-purple-600 hover:bg-purple-500 text-white shadow-[0_0_20px_rgba(168,85,247,0.3)]";
                iconColor = "text-purple-400";
              }

              return (
                <motion.div
                  key={pkg.id}
                  initial={{ opacity: 0, y: 30 }}
                  whileInView={{ opacity: 1, y: 0 }}
                  whileHover={{ y: -8, scale: 1.05, transition: { duration: 0.15, ease: "easeOut" } }}
                  viewport={{ once: true, margin: "-50px" }}
                  transition={{ duration: 0.4, ease: "easeOut", delay: index * 0.05 }}
                  className={`rounded-3xl p-[1px] flex flex-col relative transition-all duration-300 group ${
                    pkg.popular ? "z-10 shadow-2xl" : ""
                  } ${borderGlow}`}
                >
                  {/* Animated Border Gradient */}
                  <div className={`absolute inset-0 rounded-3xl overflow-hidden opacity-50 group-hover:opacity-100 transition-opacity duration-500 ${
                    pkg.popular ? "bg-gradient-to-b from-primary via-accent to-transparent" :
                    isEnterprise ? "bg-gradient-to-b from-purple-500 via-transparent to-transparent" :
                    "bg-gradient-to-b from-white/20 to-transparent"
                  }`} />

                  {/* Inner Card Content */}
                  <div className={`relative h-full flex flex-col p-8 rounded-[23px] ${cardBg} border border-transparent`}>
                    
                    {pkg.popular && (
                      <div className="absolute top-0 right-8 transform -translate-y-1/2">
                        <span className={`px-4 py-1 rounded-full text-xs font-black uppercase tracking-widest ${badgeColor}`}>
                          Most Popular
                        </span>
                      </div>
                    )}
                    
                    {isEnterprise && !pkg.popular && (
                      <div className="absolute top-0 right-8 transform -translate-y-1/2">
                        <span className={`px-4 py-1 rounded-full text-xs font-black uppercase tracking-widest ${badgeColor}`}>
                          Enterprise
                        </span>
                      </div>
                    )}

                    <div className="mb-6">
                      <h3 className="text-2xl font-bold text-white mb-2">{pkg.title}</h3>
                      <p className="text-muted-foreground text-sm leading-relaxed">{pkg.desc}</p>
                    </div>

                    <div className="mb-8 pb-8 border-b border-white/5">
                      {currencyLoading || configLoading ? (
                        <div className="h-14 w-40 bg-white/5 animate-pulse rounded-xl"></div>
                      ) : (
                        <div className="flex flex-col">
                          {(() => {
                            const isUpgrade = highestActiveSub && (dynamicPriceINR * rate) > highestActiveSub.price;
                            const finalPriceLocal = isUpgrade ? Math.max(0, (dynamicPriceINR * rate) - discountLocal) : (dynamicPriceINR * rate);
                            
                            return (
                              <>
                                <div className="flex items-baseline gap-1">
                                  <span className="text-4xl md:text-5xl font-black text-white tracking-tighter">
                                    {formatPriceLocal(finalPriceLocal)}
                                  </span>
                                  <span className="text-lg text-muted-foreground font-medium uppercase">/mo</span>
                                </div>
                                {isUpgrade && discountLocal > 0 && (
                                  <span className="text-xs text-green-400 mt-2 font-bold bg-green-500/10 self-start px-2 py-1 rounded-md">
                                    Includes {formatPriceLocal(discountLocal)} discount for {remainingDays} days left
                                  </span>
                                )}
                                {!isUpgrade && (
                                  <span className="text-xs text-muted-foreground mt-2 font-medium bg-white/5 self-start px-2 py-1 rounded-md">
                                    Monthly payment
                                  </span>
                                )}
                              </>
                            );
                          })()}
                        </div>
                      )}
                    </div>

                    <div className="flex-1 space-y-6 mb-8">
                      {/* PC Count Box */}
                      <div className={`flex items-center gap-4 p-4 rounded-2xl ${pkg.popular ? 'bg-primary/10' : isEnterprise ? 'bg-purple-500/10' : 'bg-white/5'}`}>
                        <div className={`w-12 h-12 rounded-xl flex items-center justify-center ${pkg.popular ? 'bg-primary/20' : isEnterprise ? 'bg-purple-500/20' : 'bg-white/10'}`}>
                          <Monitor className={iconColor} size={24} />
                        </div>
                        <div>
                          <div className="text-sm text-muted-foreground font-medium">Included Licenses</div>
                          <div className="text-xl font-bold text-white">{pkg.pcs} PCs</div>
                        </div>
                      </div>

                      {/* Extra Free PCs Box */}
                      {pkg.extra > 0 && (
                        <div className="flex items-center gap-4 p-4 rounded-2xl bg-green-500/10 border border-green-500/20 relative overflow-hidden group-hover:bg-green-500/15 transition-colors">
                          <div className="absolute -right-4 -top-4 w-24 h-24 bg-green-500/20 blur-2xl rounded-full pointer-events-none" />
                          <div className="w-12 h-12 rounded-xl flex items-center justify-center bg-green-500/20 relative z-10">
                            <Zap className="text-green-400" size={24} />
                          </div>
                          <div className="relative z-10">
                            <div className="text-sm text-green-400/80 font-medium">Bonus Free PCs</div>
                            <div className="text-xl font-bold text-green-400">+{pkg.extra} PCs</div>
                          </div>
                        </div>
                      )}

                      {/* Features List */}
                      <ul className="pt-4 space-y-3">
                        <li className="flex items-center gap-3">
                          <div className={`flex-shrink-0 w-6 h-6 rounded-full flex items-center justify-center ${pkg.popular ? 'bg-primary/20' : 'bg-white/10'}`}>
                            <Check className={iconColor} size={14} />
                          </div>
                          <span className="text-sm text-white/80">Local JSON Backups</span>
                        </li>
                        <li className="flex items-center gap-3">
                          <div className={`flex-shrink-0 w-6 h-6 rounded-full flex items-center justify-center ${pkg.popular ? 'bg-primary/20' : 'bg-white/10'}`}>
                            <Check className={iconColor} size={14} />
                          </div>
                          <span className="text-sm text-white/80">Google Sheets Sync</span>
                        </li>
                        <li className="flex items-center gap-3">
                          <div className={`flex-shrink-0 w-6 h-6 rounded-full flex items-center justify-center ${pkg.popular ? 'bg-primary/20' : 'bg-white/10'}`}>
                            <Check className={iconColor} size={14} />
                          </div>
                          <span className="text-sm text-white/80">Privacy-First Tracking</span>
                        </li>
                      </ul>
                    </div>

                    {(() => {
                      const isUpgrade = highestActiveSub && (dynamicPriceINR * rate) > highestActiveSub.price;
                      let finalPriceLocal = isUpgrade ? Math.max(0, (dynamicPriceINR * rate) - discountLocal) : (dynamicPriceINR * rate);
                      
                      // Enforce Razorpay minimum limits (1 INR, or roughly 0.50 for USD/EUR etc)
                      const minAllowed = currency === "INR" ? 1 : 0.50;
                      if (finalPriceLocal < minAllowed) {
                        finalPriceLocal = minAllowed; // Charge token amount if fully discounted to pass Razorpay validation
                      }

                      const isCurrentHighest = highestActiveSub && pkg.id === highestActiveSub.planId;
                      const buttonText = isCurrentHighest ? "Renew Plan" : (isUpgrade ? `Upgrade to ${pkg.title}` : `Add ${pkg.title} as Add-on`);
                      
                      return (
                        <button
                          onClick={() => handleCheckout(pkg, dynamicPriceINR, isUpgrade, finalPriceLocal)}
                          disabled={processing === pkg.id}
                          className={`w-full py-4 rounded-xl font-bold transition-all flex items-center justify-center gap-2 mt-auto ${buttonStyle} ${processing === pkg.id ? "opacity-50 cursor-not-allowed" : ""}`}
                        >
                          {processing === pkg.id ? "Processing..." : buttonText}
                        </button>
                      );
                    })()}
                  </div>
                </motion.div>
              );
            })}
          </div>
        </div>
      </section>

      {/* Footer */}
      <footer className="border-t border-white/10 py-12 mt-12 text-center text-muted-foreground">
        <div className="container mx-auto px-6 flex flex-col md:flex-row items-center justify-between">
          <div className="flex items-center gap-2 mb-4 md:mb-0">
            <Image src="/logo.png" alt="Logo" width={20} height={20} className="object-contain" />
            <span className="font-bold text-foreground">ActivityTracker</span>
          </div>
          <p className="text-sm">© 2026 ActivityTracker. All rights reserved.</p>
        </div>
      </footer>

      {/* Payment Status Dialog */}
      {paymentStatus !== 'idle' && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
          <motion.div
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="bg-[#111] border border-white/10 p-8 rounded-3xl max-w-md w-full shadow-2xl flex flex-col items-center text-center"
          >
            {paymentStatus === 'success' ? (
              <div className="w-16 h-16 bg-green-500/20 rounded-full flex items-center justify-center mb-6">
                <Check className="text-green-500" size={32} />
              </div>
            ) : (
              <div className="w-16 h-16 bg-red-500/20 rounded-full flex items-center justify-center mb-6">
                <XCircle className="text-red-500" size={32} />
              </div>
            )}
            
            <h3 className="text-2xl font-bold mb-2">
              {paymentStatus === 'success' ? 'Payment Successful' : 'Payment Failed'}
            </h3>
            
            <p className="text-muted-foreground mb-8">
              {paymentMessage}
            </p>
            
            {paymentStatus === 'success' ? (
              <button
                onClick={() => router.push("/subscription")}
                className="w-full py-3 bg-primary hover:bg-primary/90 text-white rounded-xl font-bold transition-all"
              >
                Go to Subscription
              </button>
            ) : (
              <button
                onClick={() => setPaymentStatus('idle')}
                className="w-full py-3 bg-white/10 hover:bg-white/20 text-white rounded-xl font-bold transition-all"
              >
                Try Again
              </button>
            )}
          </motion.div>
        </div>
      )}
    </main>
  );
}
