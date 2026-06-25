"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { Activity } from "lucide-react";

export default function PrivacyPage() {
  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24 flex flex-col">
      <Navbar />
      
      <section className="relative py-20 px-6 flex-1">
        <div className="container mx-auto max-w-3xl relative z-10">
          <div className="mb-12">
            <h1 className="text-4xl font-extrabold tracking-tight mb-4">Privacy Policy</h1>
            <p className="text-muted-foreground">Last updated: June 10, 2026</p>
          </div>

          <div className="prose prose-invert max-w-none text-muted-foreground leading-relaxed space-y-6">
            <p>
              Your privacy is extremely important to us. It is ActivityTracker's policy to respect your privacy regarding any information we may collect while operating our application.
            </p>

            <h2 className="text-2xl font-bold text-foreground mt-8 mb-4">1. Data Collection</h2>
            <p>
              ActivityTracker operates natively on your desktop. We log system activity intervals (mouse and keyboard input) and active window titles. We <strong>do not</strong> record keystrokes, passwords, or the contents of your screen. 
            </p>

            <h2 className="text-2xl font-bold text-foreground mt-8 mb-4">2. Google Integration</h2>
            <p>
              To sync data to your Google Sheets, the application requires authorization via Google OAuth 2.0. We request the minimum scopes required to append rows to your selected spreadsheet. The OAuth token is stored locally encrypted on your machine and never transmitted to our servers.
            </p>

            <h2 className="text-2xl font-bold text-foreground mt-8 mb-4">3. Local Data Storage</h2>
            <p>
              When offline, the application stores JSON backups of your activity locally in the `%LOCALAPPDATA%\\ActivityTracker` directory. You have complete control over this data and may delete it at any time.
            </p>
            
            <h2 className="text-2xl font-bold text-foreground mt-8 mb-4">4. Contact Us</h2>
            <p>
              If you have any questions about this Privacy Policy, please contact us via our feedback page.
            </p>
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
