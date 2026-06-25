"use client";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Activity, Mail, RefreshCcw, ArrowRight } from "lucide-react";
import { useRouter } from "next/navigation";
import { useState, useEffect } from "react";
import { auth } from "@/lib/firebase";
import { sendEmailVerification } from "firebase/auth";
import { useAuth } from "@/context/AuthContext";

export default function VerifyEmailPage() {
  const router = useRouter();
  const { user } = useAuth();
  const [loading, setLoading] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [countdown, setCountdown] = useState(120);

  useEffect(() => {
    // If user is somehow already verified, push to dashboard
    if (user?.emailVerified) {
      router.push("/");
    }
  }, [user, router]);

  useEffect(() => {
    let timer: NodeJS.Timeout;
    if (countdown > 0) {
      timer = setInterval(() => {
        setCountdown((prev) => prev - 1);
      }, 1000);
    }
    return () => {
      if (timer) clearInterval(timer);
    };
  }, [countdown]);

  const handleResendEmail = async () => {
    if (!auth.currentUser) return;
    setLoading(true);
    setError("");
    setMessage("");
    try {
      await sendEmailVerification(auth.currentUser);
      setMessage("Verification email has been resent! Check your inbox.");
      setCountdown(120);
    } catch (err: any) {
      console.error(err);
      setError(err.message || "Failed to resend email. Please try again later.");
    } finally {
      setLoading(false);
    }
  };

  const handleRefresh = async () => {
    if (!auth.currentUser) return;
    setRefreshing(true);
    setError("");
    setMessage("");
    try {
      await auth.currentUser.reload();
      if (auth.currentUser.emailVerified) {
        setMessage("Email verified successfully! Redirecting...");
        setTimeout(() => {
          router.push("/");
        }, 1500);
      } else {
        setError("Email is still not verified. Please check your inbox and spam folder.");
      }
    } catch (err: any) {
      console.error(err);
      setError(err.message || "Failed to refresh user status.");
    } finally {
      setRefreshing(false);
    }
  };

  return (
    <main className="min-h-screen bg-background text-foreground flex flex-col pt-24 relative overflow-hidden">
      <Navbar />
      
      <div className="absolute top-[20%] left-1/2 -translate-x-1/2 w-[600px] h-[400px] bg-primary/20 blur-[150px] rounded-full pointer-events-none" />

      <section className="flex-1 flex items-center justify-center p-6 relative z-10">
        <motion.div
          initial={{ opacity: 0, scale: 0.95 }}
          animate={{ opacity: 1, scale: 1 }}
          transition={{ duration: 0.5 }}
          className="w-full max-w-md glass p-10 rounded-3xl border border-white/10 shadow-2xl text-center"
        >
          <div className="w-20 h-20 bg-primary/20 rounded-full flex items-center justify-center text-primary mx-auto mb-6">
            <Mail size={40} />
          </div>
          
          <h1 className="text-3xl font-extrabold mb-4">Verify your email</h1>
          <p className="text-muted-foreground mb-8">
            We've sent a verification link to <span className="text-foreground font-bold">{user?.email || "your email address"}</span>. 
            Please click the link in the email to activate your account.
          </p>

          {message && (
            <div className="bg-green-500/10 border border-green-500/20 text-green-400 p-4 rounded-xl text-sm mb-6">
              {message}
            </div>
          )}

          {error && (
            <div className="bg-red-500/10 border border-red-500/20 text-red-400 p-4 rounded-xl text-sm mb-6">
              {error}
            </div>
          )}

          <div className="flex flex-col gap-4">
            <button
              onClick={handleRefresh}
              disabled={refreshing}
              className="w-full bg-primary hover:bg-primary/90 text-white font-bold py-4 rounded-xl transition-all hover:scale-[1.02] active:scale-[0.98] shadow-lg flex items-center justify-center gap-2 disabled:opacity-50 disabled:pointer-events-none"
            >
              {refreshing ? (
                <RefreshCcw size={18} className="animate-spin" />
              ) : (
                <RefreshCcw size={18} />
              )}
              I've Verified (Refresh)
            </button>

            <button
              onClick={handleResendEmail}
              disabled={loading || countdown > 0}
              className="w-full bg-white/5 hover:bg-white/10 text-foreground font-bold py-4 rounded-xl transition-all flex items-center justify-center gap-2 border border-white/10 disabled:opacity-50 disabled:pointer-events-none"
            >
              {loading ? "Sending..." : countdown > 0 ? `Resend in ${countdown}s` : "Resend Verification Email"} 
              {countdown === 0 && !loading && <ArrowRight size={18} />}
            </button>
          </div>
        </motion.div>
      </section>
    </main>
  );
}
