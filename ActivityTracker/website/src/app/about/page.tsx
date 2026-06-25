"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Activity, ShieldCheck, Heart, Smartphone, ChevronRight } from "lucide-react";
import Link from "next/link";

export default function AboutPage() {
  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24 flex flex-col">
      <Navbar />
      
      <section className="relative py-20 px-6 flex-1 flex flex-col items-center">
        <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[600px] h-[300px] bg-primary/20 blur-[150px] rounded-full pointer-events-none" />
        
        <div className="container mx-auto max-w-3xl relative z-10 text-center">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="mb-12"
          >
            <div className="flex justify-center mb-8">
              <div className="bg-primary/20 p-4 rounded-3xl text-primary">
                <Activity size={48} />
              </div>
            </div>
            <h1 className="text-4xl md:text-5xl font-extrabold tracking-tight mb-6">About Us</h1>
            <p className="text-lg text-muted-foreground leading-relaxed">
              At ActivityTracker, we believe that understanding your time is the first step toward mastering it. 
              Our mission is to provide you with the most accurate, privacy-respecting productivity insights possible, seamlessly integrated into the tools you already use.
            </p>
          </motion.div>

          <div className="space-y-12 text-left">
             <div className="glass p-8 rounded-3xl">
                <h3 className="text-2xl font-bold mb-4">Our Story</h3>
                <p className="text-muted-foreground leading-relaxed mb-4">
                  ActivityTracker began as an internal tool designed to solve a simple problem: how do we track billable hours without invasive employee monitoring software? Traditional trackers take screenshots, record every keystroke, and make professionals feel like they are constantly being watched.
                </p>
                <p className="text-muted-foreground leading-relaxed">
                  We wanted a solution that respected our team's privacy while still delivering hard, undeniable data about where our time was going. By leveraging deep Windows APIs, we built a system that relies on smart heuristics instead of brute-force surveillance. Today, ActivityTracker helps thousands of freelancers and agencies quantify their productivity effortlessly.
                </p>
             </div>

             <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div className="bg-white/5 backdrop-blur-xl border border-white/10 p-8 rounded-3xl">
                   <ShieldCheck size={32} className="text-green-400 mb-4" />
                   <h3 className="text-xl font-bold mb-2">Privacy First</h3>
                   <p className="text-muted-foreground text-sm leading-relaxed">
                     We don't log keystrokes or record screens. We rely entirely on system-level heuristics to determine activity and idle times, ensuring your sensitive data never leaves your desktop.
                   </p>
                </div>
                <div className="bg-white/5 backdrop-blur-xl border border-white/10 p-8 rounded-3xl">
                   <Heart size={32} className="text-red-400 mb-4" />
                   <h3 className="text-xl font-bold mb-2">Built for Professionals</h3>
                   <p className="text-muted-foreground text-sm leading-relaxed">
                     Whether you are a freelancer billing hourly or a developer trying to optimize deep work sessions, our tools are built to give you the data you need without the bloat you don't.
                   </p>
                </div>
             </div>

             <div className="glass p-8 rounded-3xl text-center">
                <div className="w-16 h-16 bg-blue-500/20 text-blue-400 rounded-2xl flex items-center justify-center mx-auto mb-6">
                   <Smartphone size={32} />
                </div>
                <h3 className="text-3xl font-bold mb-4">Our Mobile Apps</h3>
                <p className="text-muted-foreground leading-relaxed mb-8 max-w-xl mx-auto">
                  We don't just build for the desktop. Discover our ecosystem of mobile applications designed with the same privacy-first, power-user philosophy. Check out our flagship dialer and call recorder for Android.
                </p>
                <Link href="/recline" className="inline-flex items-center gap-2 bg-white/10 hover:bg-white/20 text-foreground font-bold px-6 py-3 rounded-xl transition-colors border border-white/10">
                   Discover Recline <ChevronRight size={18} />
                </Link>
             </div>
          </div>
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
    </main>
  );
}
