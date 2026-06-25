"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Activity, Plus, Minus } from "lucide-react";
import { useState } from "react";

export default function FAQPage() {
  const [openIndex, setOpenIndex] = useState<number | null>(0);

  const faqs = [
    {
      question: "Does ActivityTracker record my keystrokes?",
      answer: "No. ActivityTracker is built with a privacy-first approach. It only logs the intervals of when mouse movement or keyboard input occurred to determine if you are active or idle. It does not record the keys you press or the contents of your screen."
    },
    {
      question: "How does the Google Sheets integration work?",
      answer: "The application uses Google OAuth 2.0 to securely authenticate with your Google account. It then uses the Google Sheets API to append daily tracking logs (active window, idle time, total time) directly to a sheet you specify. You remain in complete control of the sheet."
    },
    {
      question: "What happens if my internet connection drops?",
      answer: "ActivityTracker automatically saves your tracking data to a local JSON file on your machine. When your internet connection is restored, the application will automatically sync the pending data to your Google Sheet without losing any information."
    },
    {
      question: "Can I use ActivityTracker on a Mac?",
      answer: "Currently, ActivityTracker is designed specifically for Windows 10 and 11, utilizing native Windows APIs for precise idle detection and active window tracking. A macOS version is planned for the future."
    },
    {
      question: "How much resources does it consume?",
      answer: "The application is highly optimized to run silently in the background. It consumes minimal CPU and memory (typically under 50MB RAM), ensuring zero impact on your computer's performance while you work."
    }
  ];

  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24">
      <Navbar />
      
      <section className="relative py-20 px-6">
        <div className="absolute top-[30%] right-[-10%] w-[40%] h-[40%] bg-accent/20 blur-[150px] rounded-full pointer-events-none" />
        
        <div className="container mx-auto max-w-3xl relative z-10">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="text-center mb-16"
          >
            <h1 className="text-5xl font-extrabold tracking-tight mb-6">Frequently Asked Questions</h1>
            <p className="text-xl text-muted-foreground">
              Everything you need to know about the product and billing.
            </p>
          </motion.div>

          <div className="space-y-4">
            {faqs.map((faq, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                className="glass rounded-2xl overflow-hidden border border-white/5"
              >
                <button
                  className="w-full px-6 py-5 flex items-center justify-between text-left font-bold text-lg hover:bg-white/5 transition-colors"
                  onClick={() => setOpenIndex(openIndex === index ? null : index)}
                >
                  {faq.question}
                  {openIndex === index ? (
                    <Minus className="text-primary flex-shrink-0" />
                  ) : (
                    <Plus className="text-muted-foreground flex-shrink-0" />
                  )}
                </button>
                <div
                  className={`overflow-hidden transition-all duration-300 ${
                    openIndex === index ? "max-h-96" : "max-h-0"
                  }`}
                >
                  <div className="px-6 pb-6 text-muted-foreground leading-relaxed">
                    {faq.answer}
                  </div>
                </div>
              </motion.div>
            ))}
          </div>
          
          <div className="mt-16 text-center">
            <p className="text-muted-foreground mb-4">Still have questions?</p>
            <a href="/feedback" className="inline-flex px-6 py-3 rounded-xl bg-white/10 hover:bg-white/20 transition-colors font-medium text-sm">
              Contact Support
            </a>
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
