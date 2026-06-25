"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Activity, Crown, Monitor, Calendar, CreditCard, ChevronRight, ChevronDown, ChevronUp, Trash2, Laptop, AlertTriangle } from "lucide-react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { useState, useEffect } from "react";
import { useAuth } from "@/context/AuthContext";
import { db } from "@/lib/firebase";
import { collection, onSnapshot, doc, deleteDoc, updateDoc } from "firebase/firestore";

export default function SubscriptionPage() {
  const { user, subscriptions, userDocId, loading } = useAuth();
  const [isPCsOpen, setIsPCsOpen] = useState(false);
  const [pcs, setPcs] = useState<any[]>([]);
  const router = useRouter();
  const [showCancelModal, setShowCancelModal] = useState(false);
  const [isCancelling, setIsCancelling] = useState(false);

  const activeSubs = subscriptions ? subscriptions.filter(s => s.active) : [];

  useEffect(() => {
    if (!loading && (!user || activeSubs.length === 0)) {
      router.push("/pricing");
    }
  }, [user, subscriptions, loading, router]);

  useEffect(() => {
    if (!user || !userDocId) return;
    const unsub = onSnapshot(collection(db, "users", userDocId, "pcs"), (snapshot) => {
      const pcList = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      setPcs(pcList);
    }, (error) => {
      console.warn("PC collection snapshot listener closed/error:", error.message);
    });
    return () => unsub();
  }, [user, userDocId]);

  const removePC = async (pcId: string) => {
    if (!user || !userDocId) return;
    try {
      await deleteDoc(doc(db, "users", userDocId, "pcs", pcId));
    } catch (error) {
      console.error("Error removing PC: ", error);
    }
  };

  const [cancelPlanId, setCancelPlanId] = useState<string | null>(null);

  const confirmCancel = async () => {
    if (!user || !userDocId || !cancelPlanId) return;
    setIsCancelling(true);
    try {
      const updatedSubs = subscriptions.map(sub => 
        sub.id === cancelPlanId ? { ...sub, active: false } : sub
      );
      await updateDoc(doc(db, "users", userDocId), {
        subscriptions: updatedSubs,
        subscription: updatedSubs.filter(s => s.active).sort((a: any, b: any) => (b.price || 0) - (a.price || 0))[0] || null
      });
      setShowCancelModal(false);
      setCancelPlanId(null);
    } catch (error) {
      console.error("Error cancelling subscription: ", error);
      alert("Failed to cancel subscription. Please try again.");
    } finally {
      setIsCancelling(false);
    }
  };

  const cancelSubscription = (planId: string) => {
    setCancelPlanId(planId);
    setShowCancelModal(true);
  };

  if (loading || !user || activeSubs.length === 0) {
    return (
      <main className="min-h-screen bg-background flex items-center justify-center pt-24">
        <Navbar />
        <Activity className="animate-spin text-primary" size={48} />
      </main>
    );
  }

  const totalAllowedPcs = activeSubs.reduce((acc, sub) => acc + (sub.pcs || 0) + (sub.bonus || 0), 0);
  const usagePercent = Math.min(100, (pcs.length / (totalAllowedPcs || 1)) * 100);

  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24">
      <Navbar />
      
      <section className="relative py-20 px-6">
        <div className="absolute top-0 right-0 w-[600px] h-[600px] bg-primary/10 blur-[150px] rounded-full pointer-events-none" />
        <div className="absolute bottom-0 left-0 w-[400px] h-[400px] bg-accent/10 blur-[150px] rounded-full pointer-events-none" />
        
        <div className="container mx-auto max-w-4xl relative z-10">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="mb-12 flex flex-col md:flex-row md:items-end justify-between gap-6"
          >
            <div>
              <h1 className="text-4xl md:text-5xl font-extrabold tracking-tight mb-2">My Subscription</h1>
              <p className="text-lg text-muted-foreground">
                Manage your active licenses and billing information.
              </p>
            </div>
            <Link href="/pricing" className="px-6 py-3 rounded-xl bg-white/10 hover:bg-white/20 transition-colors font-medium flex items-center justify-center gap-2">
              Upgrade Plan
            </Link>
          </motion.div>

          <div className="flex flex-col gap-6">
            {activeSubs.map((sub, index) => (
              <motion.div
                key={sub.id || index}
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.5, delay: 0.1 * index }}
                className="glass rounded-3xl overflow-hidden border border-primary/30 shadow-[0_0_50px_rgba(59,130,246,0.1)] relative"
              >
                 {/* Gradient Header */}
                 <div className="bg-gradient-to-r from-primary/20 via-accent/20 to-primary/5 p-8 border-b border-white/10 flex flex-col md:flex-row items-start md:items-center justify-between gap-6">
                    <div className="flex items-center gap-4">
                       <div className="w-16 h-16 rounded-2xl bg-gradient-to-br from-primary to-accent flex items-center justify-center shadow-lg">
                          <Crown size={32} className="text-white" />
                       </div>
                       <div>
                          <div className="text-sm font-bold text-primary uppercase tracking-wider mb-1">Active Plan</div>
                          <h2 className="text-3xl font-extrabold">{sub.planName || "Custom Plan"}</h2>
                          <p className="text-muted-foreground text-sm mt-1">Billed monthly • Auto-renews</p>
                       </div>
                    </div>
                    <div className="text-left md:text-right w-full md:w-auto bg-black/30 p-4 rounded-xl border border-white/5">
                       <div className="text-sm text-muted-foreground mb-1">Next payment</div>
                       <div className="text-2xl font-bold">
                         {new Intl.NumberFormat(undefined, { style: 'currency', currency: sub.currency || 'INR', maximumFractionDigits: 0 }).format(sub.price || 0)}
                       </div>
                       <div className="text-xs text-muted-foreground mt-1">on {sub.nextPayment || "N/A"}</div>

                       {sub.startDate && (
                         <div className="text-[10px] text-muted-foreground mt-1">
                           Started: {new Date(sub.startDate).toLocaleString()}
                         </div>
                       )}
                    </div>
                 </div>

                <div className="p-8">
                   <h3 className="text-xl font-bold mb-6 flex items-center gap-2">
                     <Monitor className="text-primary" /> License Capabilities
                   </h3>

                   <div className="mb-4 flex flex-col sm:flex-row justify-between items-start sm:items-center">
                      <div>
                         <span className="text-3xl font-bold text-foreground">{sub.pcs} PCs</span>
                         <span className="text-muted-foreground ml-2">Limit</span>
                      </div>
                      {(sub.bonus || 0) > 0 && (
                        <span className="mt-2 sm:mt-0 text-sm font-medium text-primary bg-primary/10 px-3 py-1 rounded-full">
                           +{sub.bonus} Extra Bonus included
                        </span>
                      )}
                   </div>

                   <h3 className="text-xl font-bold mb-6 flex items-center gap-2 pt-6 border-t border-white/10">
                     <Activity className="text-primary" /> License Key
                   </h3>

                   <div className="flex flex-col md:flex-row gap-4 mb-8">
                      <div className="flex-1 bg-black/50 border border-white/10 rounded-xl p-4 font-mono text-center tracking-widest text-lg select-all">
                         {sub.licenseKey || "N/A"}
                      </div>
                   </div>

                   <div className="mt-8">
                      <button 
                        onClick={() => cancelSubscription(sub.id!)}
                        className="w-full p-4 rounded-xl border border-red-500/20 bg-red-500/10 hover:bg-red-500/20 text-red-400 font-bold transition-colors flex items-center justify-center gap-2"
                      >
                         Cancel this Plan
                      </button>
                   </div>
                </div>
              </motion.div>
            ))}
          </div>

          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.3 }}
            className="mt-12 glass p-8 rounded-3xl border border-white/10"
          >
            <h3 className="text-2xl font-bold mb-6 flex items-center gap-2">
               <Laptop className="text-primary" /> Combined Device Usage
            </h3>
            
            <div className="mb-4 flex justify-between items-end">
               <div>
                  <span className="text-3xl font-bold text-foreground">{pcs.length}</span>
                  <span className="text-muted-foreground ml-2">/ {totalAllowedPcs} PCs Total</span>
               </div>
            </div>
            
            <div className="w-full bg-black/50 rounded-full h-3 mb-8 overflow-hidden border border-white/5">
               <motion.div 
                 initial={{ width: 0 }}
                 animate={{ width: `${usagePercent}%` }}
                 transition={{ duration: 1, delay: 0.5 }}
                 className="bg-gradient-to-r from-primary to-accent h-full rounded-full"
               />
            </div>

               {/* Connected PCs Accordion */}
               <div className="mb-8 border border-white/10 rounded-xl overflow-hidden bg-black/20">
                  <button 
                    onClick={() => setIsPCsOpen(!isPCsOpen)}
                    className="w-full flex items-center justify-between p-4 hover:bg-white/5 transition-colors"
                  >
                    <div className="flex items-center gap-3">
                      <Laptop className="text-primary" size={20} />
                      <span className="font-bold">Connected Devices ({pcs.length}/{totalAllowedPcs})</span>
                    </div>
                    {isPCsOpen ? <ChevronUp className="text-muted-foreground" size={20} /> : <ChevronDown className="text-muted-foreground" size={20} />}
                  </button>
                  
                  {isPCsOpen && (
                    <div className="border-t border-white/10">
                      {pcs.length === 0 ? (
                        <div className="p-6 text-center text-muted-foreground text-sm">
                          No devices currently connected to this license.
                        </div>
                      ) : (
                        <div className="divide-y divide-white/5">
                          {pcs.map(pc => (
                            <div key={pc.id} className="p-4 flex flex-col md:flex-row md:items-center justify-between gap-4 hover:bg-white/5 transition-colors">
                              <div>
                                <div className="font-bold text-foreground">{pc.name}</div>
                                <div className="text-xs text-muted-foreground flex gap-3 mt-1">
                                  <span>Last active: {pc.lastActive}</span>
                                  <span>IP: {pc.ip}</span>
                                </div>
                              </div>
                              <button 
                                onClick={() => removePC(pc.id)}
                                className="flex items-center justify-center gap-2 px-3 py-2 rounded-lg bg-red-500/10 text-red-400 hover:bg-red-500/20 transition-colors text-sm font-medium border border-red-500/20"
                              >
                                <Trash2 size={16} /> Remove Access
                              </button>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
               </div>

               <div className="mt-8 flex flex-col gap-4">
                  <Link href="/billing-history" className="w-full p-4 rounded-xl border border-white/5 hover:bg-white/5 transition-colors flex items-center justify-between group">
                     <div className="flex items-center gap-3">
                        <Calendar className="text-muted-foreground group-hover:text-primary transition-colors" />
                        <span className="font-medium">Billing History</span>
                     </div>
                     <ChevronRight size={18} className="text-muted-foreground" />
                  </Link>
               </div>
          </motion.div>
        </div>
      </section>

      {/* Footer */}
      <footer className="border-t border-white/10 py-12 mt-auto text-center text-muted-foreground">
        <div className="container mx-auto px-6 flex flex-col md:flex-row items-center justify-between">
          <div className="flex items-center gap-2 mb-4 md:mb-0">
            <Image src="/logo.png" alt="Logo" width={20} height={20} className="object-contain" />
            <span className="font-bold text-foreground">ActivityTracker</span>
          </div>
          <p className="text-sm">© 2026 ActivityTracker. All rights reserved.</p>
        </div>
      </footer>
      {/* Custom Cancel Confirmation Modal */}
      {showCancelModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
          <motion.div
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="bg-card border border-white/10 p-6 rounded-3xl max-w-sm w-full shadow-2xl relative"
          >
            <div className="w-12 h-12 rounded-full bg-red-500/20 text-red-500 flex items-center justify-center mb-4">
              <AlertTriangle size={24} />
            </div>
            <h3 className="text-xl font-bold mb-2">Cancel Subscription</h3>
            <p className="text-muted-foreground text-sm mb-6">
              Are you sure you want to cancel? This will instantly revoke your access and wipe your premium license data.
            </p>
            <div className="flex gap-3 justify-end">
              <button
                onClick={() => setShowCancelModal(false)}
                disabled={isCancelling}
                className="px-4 py-2 bg-white/5 hover:bg-white/10 rounded-xl font-medium transition-colors disabled:opacity-50"
              >
                Go Back
              </button>
              <button
                onClick={confirmCancel}
                disabled={isCancelling}
                className="px-4 py-2 bg-red-500 hover:bg-red-600 text-white rounded-xl font-bold transition-colors shadow-lg shadow-red-500/20 disabled:opacity-50 flex items-center gap-2"
              >
                {isCancelling ? "Cancelling..." : "Confirm Cancel"}
              </button>
            </div>
          </motion.div>
        </div>
      )}
    </main>
  );
}
