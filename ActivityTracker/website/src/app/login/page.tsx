"use client";

import { motion } from "framer-motion";
import { Activity, Mail, Lock, ArrowRight } from "lucide-react";
import Link from "next/link";
import { useRouter, useSearchParams } from "next/navigation";
import { useState, useEffect } from "react";
import { Suspense } from "react";
import { useAuth } from "@/context/AuthContext";
import { auth } from "@/lib/firebase";
import { signInWithEmailAndPassword, sendPasswordResetEmail, setPersistence, browserSessionPersistence, browserLocalPersistence } from "firebase/auth";

function LoginContent() {
  const router = useRouter();
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");
  const [loading, setLoading] = useState(false);
  const [rememberMe, setRememberMe] = useState(true);
  const [isForgotPassword, setIsForgotPassword] = useState(false);
  
  const searchParams = useSearchParams();
  const redirectUrl = searchParams.get("redirect") || "/";
  const { user, loading: authLoading } = useAuth();

  useEffect(() => {
    if (user && !authLoading) {
      router.replace(redirectUrl);
    }
  }, [user, authLoading, router, redirectUrl]);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setError("");
    setMessage("");
    setLoading(true);

    try {
      await setPersistence(auth, rememberMe ? browserLocalPersistence : browserSessionPersistence);
      await signInWithEmailAndPassword(auth, email, password);
      router.replace(redirectUrl);
    } catch (err: any) {
      if (err.code === "auth/invalid-credential" || err.code === "auth/user-not-found" || err.code === "auth/wrong-password") {
        setError("Invalid email address or password. Please try again.");
      } else {
        setError(err.message || "Invalid credentials");
      }
    } finally {
      setLoading(false);
    }
  };

  const handleForgotPassword = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!email) {
      setError("Please enter your email address first.");
      return;
    }
    setError("");
    setMessage("");
    setLoading(true);

    try {
      await sendPasswordResetEmail(auth, email);
      setMessage("Password reset email sent! Check your inbox.");
      setIsForgotPassword(false);
    } catch (err: any) {
      console.error(err);
      setError(err.message || "Failed to send reset email.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex bg-background">
      {/* Visual Left Side */}
      <motion.div 
        initial={{ x: -100, opacity: 0 }}
        animate={{ x: 0, opacity: 1 }}
        transition={{ duration: 0.6, ease: "easeOut" }}
        className="hidden lg:flex w-1/2 bg-gradient-to-br from-primary/20 via-background to-accent/10 relative overflow-hidden flex-col items-center justify-center p-12 border-r border-white/10"
      >
        <div className="absolute top-[-10%] left-[-10%] w-[60%] h-[60%] bg-primary/30 blur-[150px] rounded-full pointer-events-none" />
        
        <div className="relative z-10 text-center max-w-lg">
           <div className="bg-primary/20 w-24 h-24 rounded-3xl text-primary mx-auto flex items-center justify-center mb-8 shadow-2xl shadow-primary/20">
              <Activity size={48} />
           </div>
           <h1 className="text-5xl font-extrabold mb-6 leading-tight">Welcome Back to <br/>ActivityTracker</h1>
           <p className="text-xl text-muted-foreground">Sign in to manage your subscriptions and analyze your productivity.</p>
        </div>
        
        <div className="absolute bottom-10 left-10 right-10 glass p-6 rounded-2xl">
           <div className="flex items-center gap-4 text-sm text-muted-foreground font-medium">
              <div className="w-10 h-10 rounded-full bg-primary/20 flex items-center justify-center text-primary font-bold">AT</div>
              "This tool radically transformed how our company measures deep work."
           </div>
        </div>
      </motion.div>

      {/* Form Right Side */}
      <div className="w-full lg:w-1/2 flex items-center justify-center p-6 relative">
        <div className="absolute top-0 right-0 w-full h-full overflow-hidden pointer-events-none lg:hidden">
          <div className="absolute top-[-10%] right-[-10%] w-[60%] h-[60%] bg-primary/20 blur-[120px] rounded-full" />
        </div>

        <motion.div
          initial={{ x: 50, opacity: 0 }}
          animate={{ x: 0, opacity: 1 }}
          transition={{ duration: 0.6, delay: 0.2, ease: "easeOut" }}
          className="w-full max-w-md relative z-10"
        >
          <div className="flex justify-center mb-10 lg:hidden">
            <Link href="/" className="flex items-center gap-2 group">
              <div className="bg-primary/20 p-2 rounded-xl text-primary group-hover:scale-110 transition-transform">
                <Activity size={24} />
              </div>
              <span className="text-2xl font-bold tracking-tight">
                Activity<span className="text-primary">Tracker</span>
              </span>
            </Link>
          </div>

          <h2 className="text-3xl font-bold mb-2">{isForgotPassword ? "Reset Password" : "Sign In"}</h2>
          <p className="text-muted-foreground mb-8">{isForgotPassword ? "Enter your email to receive a reset link." : "Enter your credentials to continue."}</p>

          <form onSubmit={isForgotPassword ? handleForgotPassword : handleLogin} className="space-y-5">
            {error && (
              <div className="bg-red-500/10 border border-red-500/20 text-red-500 p-3 rounded-xl text-sm mb-4">
                {error}
              </div>
            )}
            {message && (
              <div className="bg-green-500/10 border border-green-500/20 text-green-500 p-3 rounded-xl text-sm mb-4">
                {message}
              </div>
            )}

            <div>
              <label className="block text-sm font-bold text-foreground mb-2">Email Address</label>
              <div className="relative">
                <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                  <Mail size={18} className="text-muted-foreground" />
                </div>
                <input
                  type="email"
                  required
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  className="w-full bg-white/5 border border-white/10 rounded-xl py-4 pl-12 pr-4 text-foreground focus:outline-none focus:ring-2 focus:ring-primary focus:border-transparent transition-all"
                  placeholder="you@example.com"
                />
              </div>
            </div>

            {!isForgotPassword && (
              <>
                <div>
                  <label className="block text-sm font-bold text-foreground mb-2">Password</label>
                  <div className="relative">
                    <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                      <Lock size={18} className="text-muted-foreground" />
                    </div>
                    <input
                      type="password"
                      required
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      className="w-full bg-white/5 border border-white/10 rounded-xl py-4 pl-12 pr-4 text-foreground focus:outline-none focus:ring-2 focus:ring-primary focus:border-transparent transition-all"
                      placeholder="••••••••"
                    />
                  </div>
                </div>

                <div className="flex items-center justify-between mt-2 mb-8 text-sm">
                  <label className="flex items-center gap-2 cursor-pointer group">
                    <div className={`relative w-5 h-5 rounded bg-white/5 border border-white/10 flex items-center justify-center transition-colors ${rememberMe ? "border-primary bg-primary/20" : "group-hover:border-primary"}`}>
                       <input 
                          type="checkbox" 
                          className="opacity-0 absolute inset-0 w-full h-full cursor-pointer z-10" 
                          checked={rememberMe}
                          onChange={(e) => setRememberMe(e.target.checked)}
                       />
                       {rememberMe && <div className="w-2.5 h-2.5 bg-primary rounded-sm pointer-events-none" />}
                    </div>
                    <span className="text-muted-foreground group-hover:text-foreground transition-colors select-none">Remember me</span>
                  </label>
                  <button type="button" onClick={() => { setIsForgotPassword(true); setError(""); setMessage(""); }} className="text-primary font-medium hover:underline">Forgot password?</button>
                </div>
              </>
            )}

            {isForgotPassword && (
                <div className="flex justify-end mt-2 mb-8 text-sm">
                  <button type="button" onClick={() => { setIsForgotPassword(false); setError(""); setMessage(""); }} className="text-primary font-medium hover:underline">Back to login</button>
                </div>
            )}

            <button
              type="submit"
              disabled={loading}
              className="w-full bg-primary hover:bg-primary/90 text-white font-bold py-4 rounded-xl transition-all hover:scale-[1.02] active:scale-[0.98] shadow-lg shadow-primary/25 flex items-center justify-center gap-2 disabled:opacity-50 disabled:pointer-events-none"
            >
              {loading ? (isForgotPassword ? "Sending..." : "Signing in...") : (isForgotPassword ? "Send Reset Link" : "Sign In")} <ArrowRight size={18} />
            </button>
          </form>

          <p className="text-center mt-10 text-sm text-muted-foreground">
            Don't have an account?{" "}
            <Link href="/signup" className="text-primary font-bold hover:underline">
              Create an account
            </Link>
          </p>
        </motion.div>
      </div>
    </div>
  );
}

export default function LoginPage() {
  return (
    <Suspense fallback={<div className="min-h-screen flex bg-background items-center justify-center"><Activity className="animate-spin text-primary" size={48} /></div>}>
      <LoginContent />
    </Suspense>
  );
}
