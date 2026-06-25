"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Activity, Layout, Keyboard, ShieldCheck, Database, Cloud, Zap, TableProperties, ChevronRight, CheckCircle } from "lucide-react";
import Link from "next/link";

export default function FeaturesPage() {
  const topFeatures = [
    {
      icon: <Layout size={40} />,
      title: "Active Window Tracking",
      description: "Automatically detects which application you are actively using, recording the window title and process name with zero manual input required. Say goodbye to manual timers.",
      color: "text-blue-400",
      bg: "bg-blue-400/10",
      border: "border-blue-400/20",
      benefits: ["Zero manual input", "Precise logging", "Multi-monitor support"]
    },
    {
      icon: <TableProperties size={40} />,
      title: "Google Sheets Sync",
      description: "Automatically exports your daily tracking data into a beautifully formatted Google Sheet. Every session is organized by date, application, and idle time for easy reviewing.",
      color: "text-green-400",
      bg: "bg-green-400/10",
      border: "border-green-400/20",
      benefits: ["Automated formatting", "Customizable reports", "Instant sync"]
    },
    {
      icon: <Keyboard size={40} />,
      title: "Privacy-Safe Activity",
      description: "Measures keyboard and mouse activity intervals without logging any actual keystrokes. Your passwords, messages, and personal data remain entirely secure on your machine.",
      color: "text-purple-400",
      bg: "bg-purple-400/10",
      border: "border-purple-400/20",
      benefits: ["No keystroke logging", "Local processing", "Enterprise-grade security"]
    },
    {
      icon: <Zap size={40} />,
      title: "Smart Idle Detection",
      description: "Utilizes native Windows APIs to instantly pause tracking when you step away from your workstation. Ensures your reports accurately reflect true deep work without manual toggling.",
      color: "text-yellow-400",
      bg: "bg-yellow-400/10",
      border: "border-yellow-400/20",
      benefits: ["Automatic pause/resume", "Highly accurate metrics", "Battery efficient"]
    },
    {
      icon: <ShieldCheck size={40} />,
      title: "Auto-Start Integration",
      description: "Forget about manual startups. ActivityTracker seamlessly integrates into your system boot sequence, launching silently in the background so you never miss a minute of tracked time.",
      color: "text-pink-400",
      bg: "bg-pink-400/10",
      border: "border-pink-400/20",
      benefits: ["Silent boot integration", "Fail-safe tracking", "System-tray access"]
    }
  ];

  const gridFeatures = [
    {
      icon: <Activity size={32} />,
      title: "Chrome Tracking",
      description: "Monitors active tabs in Chrome to give you a precise breakdown of website visits.",
      color: "text-red-400",
      bg: "bg-red-400/20",
    },
    {
      icon: <Database size={32} />,
      title: "Local JSON Backups",
      description: "Working offline? Data is saved locally and syncs back when online.",
      color: "text-orange-400",
      bg: "bg-orange-400/20",
    },
    {
      icon: <Cloud size={32} />,
      title: "Google OAuth 2.0",
      description: "Secure, industry-standard authentication via your Google Account.",
      color: "text-cyan-400",
      bg: "bg-cyan-400/20",
    }
  ];

  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24">
      <Navbar />
      
      <section className="relative py-20 px-6">
        <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[800px] h-[500px] bg-primary/20 blur-[150px] rounded-full pointer-events-none" />
        
        <div className="container mx-auto max-w-6xl relative z-10 text-center mb-24">
          <motion.h1
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, type: "spring", bounce: 0.4 }}
            className="text-5xl md:text-7xl font-extrabold tracking-tight mb-6"
          >
            Powerful features.<br />
            <span className="text-gradient">Zero configuration.</span>
          </motion.h1>
          <motion.p
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, type: "spring", bounce: 0.4, delay: 0.1 }}
            className="text-xl text-muted-foreground max-w-2xl mx-auto"
          >
            ActivityTracker bridges the gap between deep analytics and complete privacy. Discover exactly how our engine powers your productivity.
          </motion.p>
        </div>

        <div className="container mx-auto max-w-6xl space-y-32 mb-32">
          {topFeatures.map((feature, index) => (
            <motion.div
              key={index}
              initial={{ opacity: 0, y: 50 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true, margin: "-100px" }}
              transition={{ duration: 0.7, type: "spring", bounce: 0.3 }}
              className={`flex flex-col md:flex-row items-center gap-12 ${index % 2 !== 0 ? 'md:flex-row-reverse' : ''}`}
            >
              <div className="flex-1 space-y-6">
                <div className={`inline-flex items-center justify-center w-16 h-16 rounded-2xl ${feature.bg} ${feature.color} border ${feature.border} shadow-lg`}>
                  {feature.icon}
                </div>
                <h2 className="text-3xl md:text-4xl font-bold">{feature.title}</h2>
                <p className="text-lg text-muted-foreground leading-relaxed">
                  {feature.description}
                </p>
                <ul className="space-y-3 pt-4">
                  {feature.benefits.map((benefit, bIndex) => (
                    <li key={bIndex} className="flex items-center gap-3 text-foreground/90 font-medium">
                      <CheckCircle className={`${feature.color}`} size={20} />
                      {benefit}
                    </li>
                  ))}
                </ul>
              </div>
              <div className="flex-1 w-full">
                <div className={`w-full aspect-video rounded-3xl ${feature.bg} border ${feature.border} flex items-center justify-center relative overflow-hidden group`}>
                   <div className="absolute inset-0 bg-gradient-to-br from-white/5 to-transparent opacity-0 group-hover:opacity-100 transition-opacity duration-700" />
                   {/* Abstract visual representation */}
                   <div className="w-3/4 h-3/4 glass rounded-2xl border border-white/10 shadow-2xl flex flex-col p-4 relative z-10 group-hover:scale-105 transition-transform duration-700">
                      <div className="h-6 w-1/3 bg-white/10 rounded mb-4" />
                      <div className="flex-1 bg-white/5 rounded-lg border border-white/5 flex items-center justify-center">
                         <div className={`${feature.color} opacity-50 scale-150 group-hover:scale-110 transition-transform duration-700`}>
                            {feature.icon}
                         </div>
                      </div>
                   </div>
                </div>
              </div>
            </motion.div>
          ))}
        </div>

        <div className="container mx-auto max-w-6xl">
          <div className="text-center mb-16">
            <h2 className="text-3xl font-bold">More Built-in Tools</h2>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {gridFeatures.map((feature, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 40 }}
                whileInView={{ opacity: 1, y: 0 }}
                whileHover={{ y: -8, scale: 1.02 }}
                viewport={{ once: true, margin: "-50px" }}
                transition={{ duration: 0.6, type: "spring", bounce: 0.4, delay: index * 0.1 }}
                className="bg-white/5 backdrop-blur-2xl border border-white/10 p-8 rounded-3xl transition-all shadow-[0_8px_30px_rgba(255,255,255,0.02)] group hover:bg-white/10"
              >
                <div className={`w-14 h-14 rounded-2xl flex items-center justify-center mb-6 ${feature.bg} ${feature.color} group-hover:scale-110 transition-transform shadow-inner`}>
                  {feature.icon}
                </div>
                <h3 className="text-xl font-bold mb-3 tracking-tight">{feature.title}</h3>
                <p className="text-muted-foreground text-sm leading-relaxed">
                  {feature.description}
                </p>
              </motion.div>
            ))}
          </div>

          <motion.div
             initial={{ opacity: 0, y: 30 }}
             whileInView={{ opacity: 1, y: 0 }}
             viewport={{ once: true }}
             transition={{ duration: 0.6, type: "spring", bounce: 0.4 }}
             className="mt-20 flex justify-center"
          >
             <Link href="/pricing" className="bg-primary text-white font-bold py-4 px-8 rounded-full shadow-lg shadow-primary/30 hover:bg-primary/90 transition-all flex items-center gap-2 hover:scale-105 active:scale-95">
                View Pricing Plans <ChevronRight size={20} />
             </Link>
          </motion.div>
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
