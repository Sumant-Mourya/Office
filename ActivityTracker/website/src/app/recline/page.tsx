"use client";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Phone, Mic, ShieldCheck, Download, Smartphone, Star } from "lucide-react";
import Link from "next/link";

export default function ReclinePage() {
  return (
    <main className="min-h-screen bg-[#e8f6ff] text-slate-900 overflow-x-hidden pt-24 font-sans selection:bg-blue-200">
      {/* We use a specific light theme for this page to match the app */}
      <nav className="fixed top-0 w-full z-50 transition-all duration-300 bg-white/70 backdrop-blur-xl border-b border-blue-100 shadow-sm py-4">
        <div className="container mx-auto px-6 flex items-center justify-between">
          <Link href="/" className="flex items-center gap-2 group">
            <div className="bg-blue-500 text-white p-2 rounded-xl group-hover:scale-110 transition-transform shadow-md">
              <Phone size={24} />
            </div>
            <span className="text-xl font-bold tracking-tight text-slate-800">
              Activity<span className="text-blue-500">Tracker</span>
            </span>
          </Link>
          <div className="flex gap-4">
             <Link href="/about" className="text-sm font-bold text-slate-600 hover:text-blue-500 transition-colors py-2">
                Back to About
             </Link>
          </div>
        </div>
      </nav>
      
      <section className="relative py-20 px-6">
        <div className="container mx-auto max-w-6xl">
          <div className="flex flex-col lg:flex-row items-center gap-16">
            
            <motion.div 
              initial={{ opacity: 0, x: -50 }}
              animate={{ opacity: 1, x: 0 }}
              transition={{ duration: 0.6, ease: "easeOut" }}
              className="lg:w-1/2"
            >
              <div className="inline-flex items-center gap-2 px-4 py-2 rounded-full bg-blue-100 text-blue-700 font-bold text-sm mb-6">
                <Mic size={16} /> Our Mobile App
              </div>
              <h1 className="text-5xl md:text-7xl font-black mb-6 tracking-tight text-slate-900 leading-tight">
                Recline <br/> <span className="text-blue-500">Call Recorder</span>
              </h1>
              <p className="text-xl text-slate-600 mb-10 leading-relaxed font-medium">
                The ultimate dialer and call recording experience for Android. Recline automatically logs your calls with crystal clear audio, perfectly organized and entirely private.
              </p>
              
              <div className="flex flex-col sm:flex-row gap-4">
                <a 
                  href="https://play.google.com/store/apps/details?id=com.amigo.dialer" 
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="bg-slate-900 hover:bg-slate-800 text-white px-8 py-4 rounded-2xl font-bold transition-all hover:scale-105 shadow-xl flex items-center justify-center gap-3"
                >
                  <Download size={24} /> Download on Google Play
                </a>
                <a 
                  href="https://reclinerecorder.netlify.app/" 
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="bg-white border-2 border-blue-200 text-blue-600 hover:bg-blue-50 px-8 py-4 rounded-2xl font-bold transition-all flex items-center justify-center gap-3"
                >
                  Visit Official Website
                </a>
              </div>
              
              <div className="mt-12 flex items-center gap-2 text-slate-500 font-medium">
                <div className="flex text-yellow-400">
                  <Star size={20} className="fill-current" />
                  <Star size={20} className="fill-current" />
                  <Star size={20} className="fill-current" />
                  <Star size={20} className="fill-current" />
                  <Star size={20} className="fill-current" />
                </div>
                <span>Loved by Android users worldwide</span>
              </div>
            </motion.div>

            <motion.div 
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              transition={{ duration: 0.8, delay: 0.2, ease: "easeOut" }}
              className="lg:w-1/2 relative"
            >
              <div className="absolute inset-0 bg-blue-300 blur-[100px] opacity-40 rounded-full" />
              
              {/* Mockup Frame */}
              <div className="relative mx-auto w-[300px] h-[600px] bg-slate-900 rounded-[3rem] border-[8px] border-slate-900 shadow-2xl overflow-hidden flex flex-col">
                 {/* Status Bar */}
                 <div className="h-6 w-full flex items-center justify-between px-6 pt-2 z-10 absolute top-0">
                    <span className="text-white text-[10px] font-bold">12:00</span>
                    <div className="flex gap-1">
                       <div className="w-3 h-2 bg-white rounded-sm" />
                       <div className="w-2 h-2 bg-white rounded-full" />
                    </div>
                 </div>
                 
                 {/* App UI Simulation */}
                 <div className="flex-1 bg-slate-50 pt-16 px-4 flex flex-col">
                    <h2 className="text-2xl font-black text-slate-800 mb-6 px-2">Recents</h2>
                    
                    <div className="space-y-4 flex-1">
                       {[1,2,3,4].map((i) => (
                         <div key={i} className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 flex items-center justify-between">
                            <div className="flex items-center gap-3">
                               <div className="w-10 h-10 rounded-full bg-blue-100 flex items-center justify-center text-blue-600 font-bold">
                                  {i === 1 ? 'M' : i === 2 ? 'S' : i === 3 ? 'A' : 'J'}
                               </div>
                               <div>
                                  <div className="font-bold text-slate-800 text-sm">
                                    {i === 1 ? 'Mom' : i === 2 ? 'Sarah Office' : i === 3 ? 'Alex' : 'John Doe'}
                                  </div>
                                  <div className="text-xs text-slate-500 font-medium mt-0.5 flex items-center gap-1">
                                    <Phone size={10} className="text-green-500" /> Incoming
                                  </div>
                               </div>
                            </div>
                            <div className="w-8 h-8 rounded-full bg-slate-50 flex items-center justify-center text-blue-500">
                               <Mic size={16} />
                            </div>
                         </div>
                       ))}
                    </div>
                    
                    {/* Bottom Nav */}
                    <div className="h-20 bg-white border-t border-slate-100 flex items-center justify-around px-2 -mx-4">
                       <div className="flex flex-col items-center text-slate-400 gap-1"><Star size={20} /><span className="text-[10px] font-bold">Favorites</span></div>
                       <div className="flex flex-col items-center text-blue-600 gap-1"><Phone size={20} /><span className="text-[10px] font-bold">Recents</span></div>
                       <div className="flex flex-col items-center text-slate-400 gap-1"><Smartphone size={20} /><span className="text-[10px] font-bold">Keypad</span></div>
                    </div>
                 </div>
              </div>
            </motion.div>
          </div>
        </div>
      </section>

      <section className="py-24 bg-white">
        <div className="container mx-auto px-6 max-w-5xl">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-10 text-center">
             <div>
                <div className="w-16 h-16 rounded-3xl bg-blue-50 text-blue-500 flex items-center justify-center mx-auto mb-6">
                   <Mic size={32} />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-3">Auto Recording</h3>
                <p className="text-slate-600 font-medium">Never miss a detail. Every call is automatically recorded and saved directly to your local storage.</p>
             </div>
             <div>
                <div className="w-16 h-16 rounded-3xl bg-green-50 text-green-500 flex items-center justify-center mx-auto mb-6">
                   <Phone size={32} />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-3">Smart Dialer</h3>
                <p className="text-slate-600 font-medium">A lightning fast T9 dialer that quickly searches through your contacts while you type.</p>
             </div>
             <div>
                <div className="w-16 h-16 rounded-3xl bg-purple-50 text-purple-500 flex items-center justify-center mx-auto mb-6">
                   <ShieldCheck size={32} />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-3">Total Privacy</h3>
                <p className="text-slate-600 font-medium">Your call logs and recordings never leave your device. We don't upload your data anywhere.</p>
             </div>
          </div>
        </div>
      </section>
    </main>
  );
}
