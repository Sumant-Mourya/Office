"use client";

import React, { useState, useEffect } from "react";
import Link from "next/link";
import Image from "next/image";
import { Activity, Menu, X, ChevronDown, LogOut } from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";
import { usePathname } from "next/navigation";
import { useAuth } from "@/context/AuthContext";
import { auth } from "@/lib/firebase";
import { signOut } from "firebase/auth";

export default function Navbar() {
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [showLogoutConfirm, setShowLogoutConfirm] = useState(false);
  const [mounted, setMounted] = useState(false);
  const [isClientLoggedIn, setIsClientLoggedIn] = useState(() => {
    if (typeof window !== "undefined") {
      return localStorage.getItem("auth_user") === "true";
    }
    return false;
  });
  const { user, subscription, loading } = useAuth();
  const pathname = usePathname();

  useEffect(() => {
    setMounted(true);
  }, []);

  const showLoggedInUI = user || isClientLoggedIn;

  const handleLogout = () => {
    signOut(auth);
    if (typeof window !== "undefined") {
      localStorage.removeItem("auth_user");
    }
    setIsClientLoggedIn(false);
    setShowLogoutConfirm(false);
    setMobileMenuOpen(false);
  };

  return (
    <>
    <nav className="fixed top-0 w-full z-50 bg-background/40 backdrop-blur-2xl border-b border-white/10 shadow-lg py-4">
      <div className="container mx-auto px-6 flex items-center justify-between">
        <Link href="/" className="flex items-center gap-2 group">
          <div className="p-1 rounded-xl group-hover:scale-110 transition-transform">
            <Image src="/logo.png" alt="ActivityTracker Logo" width={32} height={32} className="w-8 h-8 object-contain" />
          </div>
          <span className="text-xl font-bold tracking-tight">
            Activity<span className="text-primary">Tracker</span>
          </span>
        </Link>

        {/* Desktop Nav */}
        <div className="hidden lg:flex items-center gap-8 text-sm font-semibold">
          {!mounted ? null : (user && subscription?.active) ? (
            <Link href="/subscription" className="flex items-center gap-2 text-foreground/80 hover:text-primary transition-colors py-2 bg-primary/10 px-4 rounded-full border border-primary/20">
               <Activity size={16} className="text-primary" /> Active Subscription
            </Link>
          ) : null}

          {pathname !== "/" && (
            <Link href="/" className={`transition-colors py-2 ${pathname === "/" ? "text-primary font-bold" : "text-foreground hover:text-primary"}`}>
              Home
            </Link>
          )}

          <Link href="/pricing" className={`transition-colors py-2 ${pathname === "/pricing" ? "text-primary font-bold" : "text-foreground/80 hover:text-primary"}`}>
            Pricing
          </Link>
          
          <div className="relative group">
            <button className={`flex items-center gap-1 transition-colors py-2 ${['/features', '/reviews'].includes(pathname) ? "text-primary font-bold" : "text-foreground/80 hover:text-primary"}`}>
              Product <ChevronDown size={14} className="group-hover:rotate-180 transition-transform duration-300" />
            </button>
            <div className="absolute top-full left-0 pt-2 opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-300">
              <div className="bg-card border border-white/10 rounded-xl p-2 min-w-[200px] shadow-2xl">
                <Link href="/features" className="block px-4 py-2 hover:bg-white/5 rounded-lg transition-colors">Features</Link>
                <Link href="/reviews" className="block px-4 py-2 hover:bg-white/5 rounded-lg transition-colors">Customer Reviews</Link>
              </div>
            </div>
          </div>

          <div className="relative group">
            <button className={`flex items-center gap-1 transition-colors py-2 ${['/docs', '/blog', '/faq'].includes(pathname) ? "text-primary font-bold" : "text-foreground/80 hover:text-primary"}`}>
              Resources <ChevronDown size={14} className="group-hover:rotate-180 transition-transform duration-300" />
            </button>
            <div className="absolute top-full left-0 pt-2 opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-300">
              <div className="bg-card border border-white/10 rounded-xl p-2 min-w-[200px] shadow-2xl">
                <Link href="/docs" className="block px-4 py-2 hover:bg-white/5 rounded-lg transition-colors">Documentation</Link>
                <Link href="/faq" className="block px-4 py-2 hover:bg-white/5 rounded-lg transition-colors">FAQ</Link>
              </div>
            </div>
          </div>

          <div className="relative group">
            <button className={`flex items-center gap-1 transition-colors py-2 ${['/about', '/feedback'].includes(pathname) ? "text-primary font-bold" : "text-foreground/80 hover:text-primary"}`}>
              Company <ChevronDown size={14} className="group-hover:rotate-180 transition-transform duration-300" />
            </button>
            <div className="absolute top-full left-0 pt-2 opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-300">
              <div className="bg-card border border-white/10 rounded-xl p-2 min-w-[200px] shadow-2xl">
                <Link href="/about" className="block px-4 py-2 hover:bg-white/5 rounded-lg transition-colors">About Us</Link>
                <Link href="/feedback" className="block px-4 py-2 hover:bg-white/5 rounded-lg transition-colors">Contact</Link>
              </div>
            </div>
          </div>

          <div className="h-5 w-[1px] bg-white/20 mx-2"></div>
          
          {!mounted ? (
            <div className="flex items-center gap-4 opacity-0">
              <div className="w-20 h-10"></div>
            </div>
          ) : showLoggedInUI ? (
            <div className="flex items-center gap-4">
              <button 
                onClick={() => setShowLogoutConfirm(true)}
                className="text-muted-foreground hover:text-red-400 transition-colors p-2 bg-white/5 rounded-full hover:bg-white/10"
                title="Logout"
              >
                <LogOut size={20} />
              </button>
            </div>
          ) : (
            <div className="flex items-center gap-4">
              <Link href="/login" className="text-foreground hover:text-primary transition-colors">
                Login
              </Link>
              <Link
                href="/signup"
                className="bg-primary hover:bg-primary/90 text-white px-6 py-2.5 rounded-full transition-all hover:scale-105 active:scale-95 shadow-[0_0_20px_rgba(59,130,246,0.3)]"
              >
                Get Started
              </Link>
            </div>
          )}
        </div>

        {/* Mobile Menu Toggle */}
        <button
          className="lg:hidden text-foreground"
          onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
        >
          {mobileMenuOpen ? <X size={24} /> : <Menu size={24} />}
        </button>
      </div>

      {/* Mobile Nav */}
      <AnimatePresence>
        {mobileMenuOpen && (
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className="absolute top-full left-0 w-full max-h-[85vh] overflow-y-auto bg-background/95 backdrop-blur-xl border-b border-white/10 p-6 flex flex-col gap-2 lg:hidden shadow-2xl"
          >
            {!mounted ? null : (user && subscription?.active) ? (
              <Link href="/subscription" onClick={() => setMobileMenuOpen(false)} className="flex items-center gap-2 bg-primary/10 border border-primary/20 text-primary font-bold px-4 py-3 rounded-xl mb-2">
                 <Activity size={18} /> Active Subscription
              </Link>
            ) : null}

            {pathname !== "/" && (
              <Link href="/" onClick={() => setMobileMenuOpen(false)} className="px-4 py-2 hover:bg-white/5 rounded-lg font-bold">Home</Link>
            )}

            <Link href="/pricing" onClick={() => setMobileMenuOpen(false)} className="px-4 py-2 hover:bg-white/5 rounded-lg font-bold">Pricing</Link>

            <div className="font-bold text-xs text-muted-foreground uppercase tracking-wider mb-2 mt-2">Product</div>
            <Link href="/features" onClick={() => setMobileMenuOpen(false)} className="px-4 py-2 hover:bg-white/5 rounded-lg">Features</Link>
            <Link href="/reviews" onClick={() => setMobileMenuOpen(false)} className="px-4 py-2 hover:bg-white/5 rounded-lg">Reviews</Link>
            
            <div className="font-bold text-xs text-muted-foreground uppercase tracking-wider mb-2 mt-4">Resources</div>
            <Link href="/docs" onClick={() => setMobileMenuOpen(false)} className="px-4 py-2 hover:bg-white/5 rounded-lg">Documentation</Link>
            <Link href="/faq" onClick={() => setMobileMenuOpen(false)} className="px-4 py-2 hover:bg-white/5 rounded-lg">FAQ</Link>
            
            <div className="font-bold text-xs text-muted-foreground uppercase tracking-wider mb-2 mt-4">Company</div>
            <Link href="/about" onClick={() => setMobileMenuOpen(false)} className="px-4 py-2 hover:bg-white/5 rounded-lg">About Us</Link>
            <Link href="/feedback" onClick={() => setMobileMenuOpen(false)} className="px-4 py-2 hover:bg-white/5 rounded-lg">Contact</Link>
            
            <hr className="border-white/10 my-4" />
            
            {!mounted ? (
              <div className="flex flex-col gap-2 mt-2 opacity-0">
                <div className="h-10"></div>
              </div>
            ) : showLoggedInUI ? (
              <>
                <button onClick={() => setShowLogoutConfirm(true)} className="px-4 py-2 font-bold text-center text-red-400">Logout</button>
              </>
            ) : (
              <>
                <Link href="/login" onClick={() => setMobileMenuOpen(false)} className="px-4 py-2 font-bold text-center">Login</Link>
                <Link href="/signup" onClick={() => setMobileMenuOpen(false)} className="bg-primary text-white px-4 py-3 rounded-xl font-bold text-center mt-2">Get Started</Link>
              </>
            )}
          </motion.div>
        )}
      </AnimatePresence>

    </nav>
    
    {/* Logout Confirmation Modal - Moved outside nav to avoid filter stacking context */}
    <AnimatePresence>
      {showLogoutConfirm && (
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60 backdrop-blur-md p-4"
        >
          <motion.div
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0.95, opacity: 0 }}
            className="bg-card border border-white/10 rounded-3xl p-6 max-w-sm w-full shadow-2xl relative"
          >
            <div className="w-12 h-12 rounded-full bg-red-500/20 text-red-400 flex items-center justify-center mb-4">
              <LogOut size={24} />
            </div>
            <h3 className="text-xl font-bold mb-2">Confirm Logout</h3>
            <p className="text-muted-foreground mb-6">Are you sure you want to sign out of your account?</p>
            
            <div className="flex justify-end gap-3">
              <button
                onClick={() => setShowLogoutConfirm(false)}
                className="px-4 py-2 rounded-xl bg-white/5 hover:bg-white/10 transition-colors font-medium flex-1"
              >
                Cancel
              </button>
              <button
                onClick={handleLogout}
                className="px-4 py-2 rounded-xl bg-red-500 hover:bg-red-600 text-white transition-colors font-medium flex items-center justify-center gap-2 flex-1 shadow-lg shadow-red-500/20"
              >
                Sign Out
              </button>
            </div>
          </motion.div>
        </motion.div>
      )}
    </AnimatePresence>
    </>
  );
}
