"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Activity, ShieldCheck, Zap, BarChart3, ChevronRight, Ban, Filter, Star, LayoutDashboard } from "lucide-react";
import Link from "next/link";
import { useState, useEffect } from "react";
import { useAuth } from "@/context/AuthContext";
import { useAppData } from "@/context/AppDataContext";

export default function Home() {
  const { user, loading: authLoading } = useAuth();
  const { displayRating, displayReviews, loading } = useAppData();
  const [mounted, setMounted] = useState(false);
  const [isClientLoggedIn, setIsClientLoggedIn] = useState(false);

  useEffect(() => {
    setMounted(true);
    setIsClientLoggedIn(localStorage.getItem("auth_user") === "true");
  }, []);

  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden">
      <Navbar />

      {/* Hero Section */}
      <section className="relative pt-32 pb-20 md:pt-48 md:pb-32 px-6">
        <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[800px] h-[400px] bg-primary/20 blur-[120px] rounded-full pointer-events-none" />
        
        <div className="container mx-auto max-w-6xl relative z-10 text-center">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full glass text-sm font-medium text-primary mb-4"
          >
            <Zap size={16} className="text-accent" />
            <span>Activity Tracker 2.0 is Here</span>
          </motion.div>

          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.05 }}
            className="flex justify-center mb-8"
          >
            {loading ? (
              <div className="flex items-center gap-1.5 px-4 py-1.5 rounded-full border border-yellow-500/10 bg-yellow-500/5 animate-pulse">
                <div className="w-4 h-4 rounded-full bg-yellow-500/20"></div>
                <div className="w-8 h-4 rounded bg-yellow-500/20"></div>
                <div className="w-20 h-4 rounded bg-white/10 mx-2"></div>
                <div className="w-24 h-4 rounded bg-white/10"></div>
              </div>
            ) : (
              <Link href="/reviews" className="flex items-center gap-1.5 px-4 py-1.5 rounded-full border border-yellow-500/30 bg-yellow-500/10 hover:bg-yellow-500/20 transition-all hover:scale-105 cursor-pointer">
                <Star size={16} className="fill-yellow-500 text-yellow-500" />
                <span className="text-sm font-bold text-yellow-500">{displayRating}/5</span>
                <span className="text-sm text-foreground/80 font-medium">Average Rating</span>
                <span className="text-xs text-muted-foreground ml-1">({displayReviews}+ reviews)</span>
              </Link>
            )}
          </motion.div>

          <motion.h1
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.1 }}
            className="text-5xl md:text-7xl font-extrabold tracking-tight mb-8 leading-tight"
          >
            Master Your Time with <br className="hidden md:block" />
            <span className="text-gradient">Ultimate Precision</span>
          </motion.h1>

          <motion.p
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.2 }}
            className="text-lg md:text-xl text-muted-foreground max-w-3xl mx-auto mb-10"
          >
            The premium desktop application to track your productivity automatically. Manage your system, analyze your habits, block distractions, and upgrade your workflow.
          </motion.p>

          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.3 }}
            className="flex flex-col sm:flex-row items-center justify-center gap-4"
          >
            {(!mounted ? !isClientLoggedIn : (!user && !authLoading)) && (
              <Link
                href="/signup"
                className="bg-primary hover:bg-primary/90 text-white px-8 py-4 rounded-full font-semibold transition-all hover:scale-105 shadow-[0_0_30px_rgba(59,130,246,0.4)] flex items-center gap-2"
              >
                Start Free Trial <ChevronRight size={20} />
              </Link>
            )}
            <Link
              href="/features"
              className="px-8 py-4 rounded-full font-semibold glass hover:bg-white/10 transition-colors"
            >
              Explore All Features
            </Link>
          </motion.div>
        </div>
      </section>

      {/* About App Section - Brand New Large Section */}
      <section className="py-24 relative overflow-hidden">
        <div className="absolute inset-0 bg-primary/5 [mask-image:linear-gradient(to_bottom,transparent,black,transparent)]" />
        <div className="container mx-auto px-6 max-w-6xl relative z-10">
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-16 items-center">
            <motion.div
              initial={{ opacity: 0, x: -50 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.7 }}
            >
              <h2 className="text-4xl md:text-5xl font-bold mb-6">About ActivityTracker</h2>
              <p className="text-xl text-muted-foreground mb-6 leading-relaxed">
                We believe that understanding your time is the foundation of high performance. Our tracker runs silently in the background, utilizing deep Windows integrations to provide data that is both accurate and absolutely private.
              </p>
              <ul className="space-y-4">
                <li className="flex items-start gap-4">
                  <div className="bg-primary/20 p-2 rounded-lg text-primary mt-1">
                    <ShieldCheck size={20} />
                  </div>
                  <div>
                    <h4 className="font-bold text-lg">No Keystrokes Logged</h4>
                    <p className="text-muted-foreground text-sm">We only check for input intervals to determine idle time, meaning your passwords and chats are perfectly safe.</p>
                  </div>
                </li>
                <li className="flex items-start gap-4">
                  <div className="bg-accent/20 p-2 rounded-lg text-accent mt-1">
                    <BarChart3 size={20} />
                  </div>
                  <div>
                    <h4 className="font-bold text-lg">Own Your Data</h4>
                    <p className="text-muted-foreground text-sm">Everything flows directly into your personal Google Sheets and local JSON files. We never hoard your activity logs.</p>
                  </div>
                </li>
                <li className="flex items-start gap-4">
                  <div className="bg-green-500/20 p-2 rounded-lg text-green-500 mt-1">
                    <Zap size={20} />
                  </div>
                  <div>
                    <h4 className="font-bold text-lg">Zero Performance Hit</h4>
                    <p className="text-muted-foreground text-sm">Engineered with C-level hooks, the application requires minimal RAM and CPU to function.</p>
                  </div>
                </li>
              </ul>
            </motion.div>

            <motion.div
              initial={{ opacity: 0, x: 50 }}
              whileInView={{ opacity: 1, x: 0 }}
              viewport={{ once: true }}
              transition={{ duration: 0.7 }}
              className="relative"
            >
              <div className="glass p-2 rounded-3xl border border-white/10 shadow-2xl relative z-10 overflow-hidden group">
                <div className="absolute inset-0 bg-gradient-to-br from-primary/20 to-accent/20 opacity-0 group-hover:opacity-100 transition-opacity duration-500" />
                <div className="bg-background/80 backdrop-blur-xl rounded-2xl p-6 border border-white/5">
                   <div className="flex items-center justify-between mb-8 border-b border-white/10 pb-4">
                      <div className="flex items-center gap-3">
                         <div className="w-10 h-10 rounded-full bg-primary/20 flex items-center justify-center text-primary">
                            <LayoutDashboard size={20} />
                         </div>
                         <div>
                            <div className="font-bold">Daily Overview</div>
                            <div className="text-xs text-muted-foreground">Local Dashboard</div>
                         </div>
                      </div>
                      <div className="px-3 py-1 rounded-full bg-green-500/20 text-green-400 text-xs font-bold">Syncing Active</div>
                   </div>

                   <div className="space-y-4">
                      <div className="h-4 w-3/4 bg-white/5 rounded-full" />
                      <div className="h-4 w-1/2 bg-white/5 rounded-full" />
                      <div className="h-4 w-5/6 bg-white/5 rounded-full" />
                      <div className="grid grid-cols-3 gap-4 mt-6">
                         <div className="h-20 bg-white/5 rounded-xl" />
                         <div className="h-20 bg-primary/20 rounded-xl border border-primary/30" />
                         <div className="h-20 bg-white/5 rounded-xl" />
                      </div>
                   </div>
                </div>
              </div>
              
              {/* Decorative elements */}
              <div className="absolute -top-10 -right-10 w-32 h-32 bg-primary/30 blur-[50px] rounded-full" />
              <div className="absolute -bottom-10 -left-10 w-40 h-40 bg-accent/30 blur-[60px] rounded-full" />
            </motion.div>
          </div>
        </div>
      </section>

      {/* Features Section (Expanded Bento Grid) */}
      <section className="py-24 relative">
        <div className="container mx-auto px-6 max-w-6xl">
          <div className="text-center mb-16">
            <h2 className="text-3xl md:text-5xl font-bold mb-4">Everything you need</h2>
            <p className="text-muted-foreground text-lg max-w-2xl mx-auto">
              Powerful tools designed to give you complete control over your productivity without compromising your privacy.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <motion.div
              initial={{ opacity: 0, y: 40 }}
              whileInView={{ opacity: 1, y: 0 }}
              whileHover={{ y: 0, scale: 0.95, transition: { duration: 0.1 } }}
              viewport={{ once: true, margin: "-50px" }}
              transition={{ duration: 1.2, ease: "easeOut" }}
              className="md:col-span-2 glass rounded-3xl p-8 hover:bg-white/10 transition-colors hover:shadow-[0_0_30px_rgba(59,130,246,0.1)] group"
            >
              <div className="bg-primary/20 w-12 h-12 rounded-2xl flex items-center justify-center text-primary mb-6 group-hover:scale-110 transition-transform">
                <BarChart3 size={24} />
              </div>
              <h3 className="text-2xl font-bold mb-3">Google Sheets Sync</h3>
              <p className="text-muted-foreground">
                Automatically export your daily tracking data into a formatted Google Sheet. It organizes your active window time, Chrome usage, and idle periods securely via OAuth 2.0.
              </p>
            </motion.div>

            <motion.div
              initial={{ opacity: 0, y: 40 }}
              whileInView={{ opacity: 1, y: 0 }}
              whileHover={{ y: 0, scale: 0.95, transition: { duration: 0.1 } }}
              viewport={{ once: true, margin: "-50px" }}
              transition={{ duration: 1.2, delay: 0.1, ease: "easeOut" }}
              className="glass rounded-3xl p-8 hover:bg-white/10 transition-colors hover:shadow-[0_0_30px_rgba(168,85,247,0.1)] group"
            >
              <div className="bg-accent/20 w-12 h-12 rounded-2xl flex items-center justify-center text-accent mb-6 group-hover:scale-110 transition-transform">
                <Activity size={24} />
              </div>
              <h3 className="text-2xl font-bold mb-3">Live Tracking</h3>
              <p className="text-muted-foreground">
                Monitor your active windows and Chrome tabs in real-time with zero performance impact.
              </p>
            </motion.div>

            <motion.div
              initial={{ opacity: 0, y: 40 }}
              whileInView={{ opacity: 1, y: 0 }}
              whileHover={{ y: 0, scale: 0.95, transition: { duration: 0.1 } }}
              viewport={{ once: true, margin: "-50px" }}
              transition={{ duration: 1.2, delay: 0.2, ease: "easeOut" }}
              className="glass rounded-3xl p-8 hover:bg-white/10 transition-colors hover:shadow-[0_0_30px_rgba(239,68,68,0.1)] group"
            >
              <div className="bg-red-500/20 w-12 h-12 rounded-2xl flex items-center justify-center text-red-500 mb-6 group-hover:scale-110 transition-transform">
                <Ban size={24} />
              </div>
              <h3 className="text-2xl font-bold mb-3">App & Website Blocking</h3>
              <p className="text-muted-foreground">
                Block distracting applications and websites during work hours. Stay focused and boost your productivity effortlessly.
              </p>
            </motion.div>

            <motion.div
              initial={{ opacity: 0, y: 40 }}
              whileInView={{ opacity: 1, y: 0 }}
              whileHover={{ y: 0, scale: 0.95, transition: { duration: 0.1 } }}
              viewport={{ once: true, margin: "-50px" }}
              transition={{ duration: 1.2, delay: 0.3, ease: "easeOut" }}
              className="md:col-span-2 glass rounded-3xl p-8 hover:bg-white/10 transition-colors hover:shadow-[0_0_30px_rgba(168,85,247,0.1)] group"
            >
              <div className="bg-purple-500/20 w-12 h-12 rounded-2xl flex items-center justify-center text-purple-500 mb-6 group-hover:scale-110 transition-transform">
                <Filter size={24} />
              </div>
              <h3 className="text-2xl font-bold mb-3">Include & Exclude Lists</h3>
              <p className="text-muted-foreground">
                Total control over what gets tracked. Exclude personal apps or explicitly include work-only websites for pinpoint tracking accuracy.
              </p>
            </motion.div>

            <motion.div
              initial={{ opacity: 0, y: 40 }}
              whileInView={{ opacity: 1, y: 0 }}
              whileHover={{ y: 0, scale: 0.95, transition: { duration: 0.1 } }}
              viewport={{ once: true, margin: "-50px" }}
              transition={{ duration: 1.2, delay: 0.4, ease: "easeOut" }}
              className="glass rounded-3xl p-8 hover:bg-white/10 transition-colors hover:shadow-[0_0_30px_rgba(34,197,94,0.1)] group"
            >
              <div className="bg-green-500/20 w-12 h-12 rounded-2xl flex items-center justify-center text-green-500 mb-6 group-hover:scale-110 transition-transform">
                <ShieldCheck size={24} />
              </div>
              <h3 className="text-2xl font-bold mb-3">Privacy First</h3>
              <p className="text-muted-foreground">
                Measures intervals without logging keystrokes.
              </p>
            </motion.div>

            <motion.div
              initial={{ opacity: 0, y: 40 }}
              whileInView={{ opacity: 1, y: 0 }}
              whileHover={{ y: 0, scale: 0.95, transition: { duration: 0.1 } }}
              viewport={{ once: true, margin: "-50px" }}
              transition={{ duration: 1.2, delay: 0.5, ease: "easeOut" }}
              className="glass rounded-3xl p-8 hover:bg-white/10 transition-colors hover:shadow-[0_0_30px_rgba(249,115,22,0.1)] group"
            >
              <div className="bg-orange-500/20 w-12 h-12 rounded-2xl flex items-center justify-center text-orange-500 mb-6 group-hover:scale-110 transition-transform">
                <Zap size={24} />
              </div>
              <h3 className="text-2xl font-bold mb-3">Local Backups</h3>
              <p className="text-muted-foreground">
                Working offline? Data is saved locally and syncs back seamlessly.
              </p>
            </motion.div>
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
    </main>
  );
}
