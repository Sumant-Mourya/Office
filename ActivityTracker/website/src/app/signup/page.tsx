"use client";

import { motion } from "framer-motion";
import { Activity, Mail, Lock, User, ArrowRight, CheckCircle2 } from "lucide-react";
import Link from "next/link";
import { useRouter, useSearchParams } from "next/navigation";
import { useState, useEffect } from "react";
import { Suspense } from "react";
import { useAuth } from "@/context/AuthContext";
import { auth, db } from "@/lib/firebase";
import { createUserWithEmailAndPassword, sendEmailVerification } from "firebase/auth";
import { doc, setDoc } from "firebase/firestore";

function SignupContent() {
  const router = useRouter();
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [isSigningUp, setIsSigningUp] = useState(false);

  const searchParams = useSearchParams();
  const redirectUrl = searchParams.get("redirect") || "/";
  const { user, loading: authLoading } = useAuth();

  useEffect(() => {
    if (user && !authLoading && !isSigningUp) {
      router.replace(redirectUrl);
    }
  }, [user, authLoading, router, redirectUrl, isSigningUp]);

  const handleSignup = async (e: React.FormEvent) => {
    e.preventDefault();
    setError("");
    
    // Password Validation
    if (password.length < 8) {
      setError("Password must be at least 8 characters long.");
      return;
    }
    if (/^(.)\1+$/.test(password)) {
      setError("Password cannot be repeating characters (e.g., 00000000, 11111111).");
      return;
    }
    if (!/[A-Z]/.test(password)) {
      setError("Password must contain at least one uppercase letter.");
      return;
    }
    if (!/[a-z]/.test(password)) {
      setError("Password must contain at least one lowercase letter.");
      return;
    }
    if (!/\d/.test(password)) {
      setError("Password must contain at least one number.");
      return;
    }
    if (!/[!@#$%^&*(),.?":{}|<>]/.test(password)) {
      setError("Password must contain at least one special character.");
      return;
    }

    setLoading(true);
    setIsSigningUp(true);

    try {
      const userCredential = await createUserWithEmailAndPassword(auth, email, password);
      const user = userCredential.user;

      // Generate 12-digit license key (xxxx-xxxx-xxxx) with digits only
      const generate12DigitId = () => {
        let id = "";
        for (let i = 0; i < 12; i++) {
          id += Math.floor(Math.random() * 10).toString();
        }
        return `${id.slice(0, 4)}-${id.slice(4, 8)}-${id.slice(8, 12)}`;
      };
      
      const customDocId = generate12DigitId();

      // Try to create base user document in Firestore with custom ID
      try {
        await setDoc(doc(db, "users", customDocId), {
          uid: user.uid,
          name,
          email,
          createdAt: new Date().toISOString(),
          licenseKey: customDocId,
          subscription: null // No active subscription by default
        });
      } catch (firestoreErr) {
        console.warn("Firestore document creation failed (possibly due to security rules). User was still created in Auth.", firestoreErr);
      }

      // Send verification email
      await sendEmailVerification(user);

      // Redirect to verify-email page
      router.push("/verify-email");
    } catch (err: any) {
      if (err.code === "auth/email-already-in-use") {
        setError("Email ID already in use. Please use a different email or log in.");
      } else {
        setError(err.message || "Failed to create an account");
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex bg-background">
      
      {/* Form Left Side */}
      <div className="w-full lg:w-1/2 flex items-center justify-center p-6 relative">
        <div className="absolute top-0 left-0 w-full h-full overflow-hidden pointer-events-none lg:hidden">
          <div className="absolute top-[-10%] left-[-10%] w-[60%] h-[60%] bg-accent/20 blur-[120px] rounded-full" />
        </div>

        <motion.div
          initial={{ y: 50, opacity: 0 }}
          animate={{ y: 0, opacity: 1 }}
          transition={{ duration: 0.6, ease: "easeOut" }}
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

          <h2 className="text-3xl font-bold mb-2">Create an Account</h2>
          <p className="text-muted-foreground mb-8">Start tracking your productivity today.</p>

          <form onSubmit={handleSignup} className="space-y-5">
            {error && (
              <div className="bg-red-500/10 border border-red-500/20 text-red-500 p-3 rounded-xl text-sm mb-4">
                {error}
              </div>
            )}

            <div>
              <label className="block text-sm font-bold text-foreground mb-2">Full Name</label>
              <div className="relative">
                <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                  <User size={18} className="text-muted-foreground" />
                </div>
                <input
                  type="text"
                  required
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  className="w-full bg-white/5 border border-white/10 rounded-xl py-4 pl-12 pr-4 text-foreground focus:outline-none focus:ring-2 focus:ring-accent focus:border-transparent transition-all"
                  placeholder="John Doe"
                />
              </div>
            </div>

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
                  className="w-full bg-white/5 border border-white/10 rounded-xl py-4 pl-12 pr-4 text-foreground focus:outline-none focus:ring-2 focus:ring-accent focus:border-transparent transition-all"
                  placeholder="you@example.com"
                />
              </div>
            </div>

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
                  className="w-full bg-white/5 border border-white/10 rounded-xl py-4 pl-12 pr-4 text-foreground focus:outline-none focus:ring-2 focus:ring-accent focus:border-transparent transition-all"
                  placeholder="••••••••"
                />
              </div>
            </div>

            <button
              type="submit"
              disabled={loading}
              className="w-full bg-accent hover:bg-accent/90 text-white font-bold py-4 rounded-xl transition-all hover:scale-[1.02] active:scale-[0.98] shadow-lg shadow-accent/25 flex items-center justify-center gap-2 mt-6 disabled:opacity-50 disabled:pointer-events-none"
            >
              {loading ? "Creating..." : "Sign Up"} <ArrowRight size={18} />
            </button>
          </form>

          <p className="text-center mt-10 text-sm text-muted-foreground">
            Already have an account?{" "}
            <Link href="/login" className="text-accent font-bold hover:underline">
              Sign in
            </Link>
          </p>
        </motion.div>
      </div>

      {/* Visual Right Side */}
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        transition={{ duration: 1 }}
        className="hidden lg:flex w-1/2 bg-gradient-to-bl from-accent/20 via-background to-primary/10 relative overflow-hidden flex-col items-center justify-center p-12 border-l border-white/10"
      >
        <div className="absolute bottom-[-10%] right-[-10%] w-[60%] h-[60%] bg-accent/30 blur-[150px] rounded-full pointer-events-none" />
        
        <motion.div 
          initial={{ scale: 0.9, opacity: 0 }}
          animate={{ scale: 1, opacity: 1 }}
          transition={{ duration: 0.6, delay: 0.4 }}
          className="relative z-10 w-full max-w-lg"
        >
           <h2 className="text-4xl font-extrabold mb-8">Why join us?</h2>
           
           <div className="space-y-6">
              <div className="glass p-6 rounded-2xl flex gap-4 items-center">
                 <div className="bg-green-500/20 p-3 rounded-xl text-green-500">
                    <CheckCircle2 size={24} />
                 </div>
                 <div>
                    <h3 className="font-bold text-lg">Instant Synchronization</h3>
                    <p className="text-sm text-muted-foreground">Your data flows directly to Google Sheets in real-time.</p>
                 </div>
              </div>

              <div className="glass p-6 rounded-2xl flex gap-4 items-center">
                 <div className="bg-green-500/20 p-3 rounded-xl text-green-500">
                    <CheckCircle2 size={24} />
                 </div>
                 <div>
                    <h3 className="font-bold text-lg">Absolute Privacy</h3>
                    <p className="text-sm text-muted-foreground">No keystrokes or screenshots are ever recorded.</p>
                 </div>
              </div>

              <div className="glass p-6 rounded-2xl flex gap-4 items-center">
                 <div className="bg-green-500/20 p-3 rounded-xl text-green-500">
                    <CheckCircle2 size={24} />
                 </div>
                 <div>
                    <h3 className="font-bold text-lg">Powerful App Blocking</h3>
                    <p className="text-sm text-muted-foreground">Gain control of your focus with advanced blocking rules.</p>
                 </div>
              </div>
           </div>
        </motion.div>
      </motion.div>

    </div>
  );
}

export default function SignupPage() {
  return (
    <Suspense fallback={<div className="min-h-screen flex bg-background items-center justify-center"><Activity className="animate-spin text-accent" size={48} /></div>}>
      <SignupContent />
    </Suspense>
  );
}
