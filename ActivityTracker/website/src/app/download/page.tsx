"use client";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Download, Monitor, CheckCircle2, ShieldCheck, Zap } from "lucide-react";
import Link from "next/link";

export default function DownloadPage() {
  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24 flex flex-col">
      <Navbar />
      
      <section className="relative py-20 px-6 flex-1 flex flex-col justify-center items-center">
        <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[800px] h-[400px] bg-primary/20 blur-[120px] rounded-full pointer-events-none" />
        
        <div className="container mx-auto max-w-4xl relative z-10 text-center">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full glass text-sm font-medium text-primary mb-8"
          >
            <Zap size={16} className="text-accent" />
            <span>Version 2.0 Now Available</span>
          </motion.div>

          <motion.h1
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.1 }}
            className="text-5xl md:text-6xl font-extrabold tracking-tight mb-6"
          >
            Download <span className="text-gradient">ActivityTracker</span>
          </motion.h1>

          <motion.p
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.2 }}
            className="text-lg md:text-xl text-muted-foreground max-w-2xl mx-auto mb-12"
          >
            Get the ultimate Windows desktop client. Track your time, sync to Google Sheets, and master your productivity natively.
          </motion.p>

          <motion.div
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            transition={{ duration: 0.5, delay: 0.3 }}
            className="glass rounded-3xl p-8 md:p-12 max-w-2xl mx-auto"
          >
            <div className="flex flex-col md:flex-row items-center justify-between gap-8 mb-8">
              <div className="flex items-center gap-6">
                <div className="w-20 h-20 rounded-2xl bg-gradient-to-br from-primary to-accent flex items-center justify-center shadow-lg shadow-primary/25">
                  <Monitor size={40} className="text-white" />
                </div>
                <div className="text-left">
                  <h3 className="text-2xl font-bold">Windows Client</h3>
                  <p className="text-muted-foreground text-sm">Windows 10 / 11 (64-bit)</p>
                </div>
              </div>
              <button className="w-full md:w-auto bg-primary hover:bg-primary/90 text-white px-8 py-4 rounded-xl font-semibold transition-all hover:scale-105 active:scale-95 shadow-[0_0_30px_rgba(59,130,246,0.4)] flex items-center justify-center gap-2">
                <Download size={20} /> Download for Windows
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-left border-t border-white/10 pt-8">
              <div className="flex items-center gap-3">
                <CheckCircle2 size={18} className="text-green-400 flex-shrink-0" />
                <span className="text-sm">Auto-updates included</span>
              </div>
              <div className="flex items-center gap-3">
                <ShieldCheck size={18} className="text-green-400 flex-shrink-0" />
                <span className="text-sm">Verified secure installer</span>
              </div>
              <div className="flex items-center gap-3">
                <CheckCircle2 size={18} className="text-green-400 flex-shrink-0" />
                <span className="text-sm">Local JSON backups</span>
              </div>
              <div className="flex items-center gap-3">
                <CheckCircle2 size={18} className="text-green-400 flex-shrink-0" />
                <span className="text-sm">Zero performance impact</span>
              </div>
            </div>
          </motion.div>
        </div>
      </section>

      {/* Footer */}
      <footer className="border-t border-white/10 py-12 mt-auto text-center text-muted-foreground">
        <div className="container mx-auto px-6 flex flex-col md:flex-row items-center justify-between">
          <div className="flex items-center gap-2 mb-4 md:mb-0">
            <Monitor size={20} className="text-primary" />
            <span className="font-bold text-foreground">ActivityTracker</span>
          </div>
          <p className="text-sm">© 2026 ActivityTracker. All rights reserved.</p>
        </div>
      </footer>
    </main>
  );
}
