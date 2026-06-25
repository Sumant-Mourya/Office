"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { motion, AnimatePresence } from "framer-motion";
import { Activity, BookOpen, Terminal, Code, Settings, ShieldCheck, Zap, Server, Monitor } from "lucide-react";
import Link from "next/link";
import { useState } from "react";

export default function DocsPage() {
  const [activeTab, setActiveTab] = useState("installation");

  const tabs = [
    { id: "installation", label: "Installation", section: "Getting Started" },
    { id: "configuration", label: "Configuration", section: "Getting Started" },
    { id: "oauth", label: "Google OAuth Setup", section: "Getting Started" },
    { id: "active-window", label: "Active Window Tracking", section: "Core Concepts" },
    { id: "idle-detection", label: "Idle Detection", section: "Core Concepts" },
    { id: "data-syncing", label: "Data Syncing", section: "Core Concepts" },
  ];

  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24">
      <Navbar />
      
      <section className="relative py-20 px-6">
        <div className="absolute top-0 right-0 w-[600px] h-[400px] bg-primary/10 blur-[150px] rounded-full pointer-events-none" />
        
        <div className="container mx-auto max-w-6xl relative z-10">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="mb-16"
          >
            <h1 className="text-5xl font-extrabold tracking-tight mb-6">Documentation</h1>
            <p className="text-xl text-muted-foreground max-w-2xl">
              Learn how to install, configure, and get the most out of ActivityTracker.
            </p>
          </motion.div>

          <div className="flex flex-col md:flex-row gap-12">
            {/* Sidebar Navigation */}
            <div className="w-full md:w-64 flex-shrink-0 space-y-6">
              <div>
                <div className="font-bold text-xs text-muted-foreground uppercase tracking-wider mb-4 px-4">Getting Started</div>
                <div className="space-y-1">
                  {tabs.filter(t => t.section === "Getting Started").map(tab => (
                    <button
                      key={tab.id}
                      onClick={() => setActiveTab(tab.id)}
                      className={`w-full text-left px-4 py-2 rounded-lg transition-colors ${
                        activeTab === tab.id ? "bg-primary/20 text-primary font-bold" : "hover:bg-white/5 text-muted-foreground hover:text-foreground"
                      }`}
                    >
                      {tab.label}
                    </button>
                  ))}
                </div>
              </div>
              
              <div>
                <div className="font-bold text-xs text-muted-foreground uppercase tracking-wider mb-4 px-4">Core Concepts</div>
                <div className="space-y-1">
                  {tabs.filter(t => t.section === "Core Concepts").map(tab => (
                    <button
                      key={tab.id}
                      onClick={() => setActiveTab(tab.id)}
                      className={`w-full text-left px-4 py-2 rounded-lg transition-colors ${
                        activeTab === tab.id ? "bg-primary/20 text-primary font-bold" : "hover:bg-white/5 text-muted-foreground hover:text-foreground"
                      }`}
                    >
                      {tab.label}
                    </button>
                  ))}
                </div>
              </div>
            </div>

            {/* Content Area */}
            <div className="flex-1 glass rounded-3xl p-8 md:p-12 min-h-[600px]">
              <AnimatePresence mode="wait">
                <motion.div
                  key={activeTab}
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -10 }}
                  transition={{ duration: 0.3 }}
                  className="prose prose-invert max-w-none"
                >
                  {/* Installation */}
                  {activeTab === "installation" && (
                    <div>
                      <h2 className="text-3xl font-bold flex items-center gap-3 mb-6">
                        <Terminal className="text-primary" /> Installation
                      </h2>
                      <p className="text-muted-foreground mb-6">
                        ActivityTracker is distributed as a single standalone executable. This makes deployment across small or massive organizations extremely straightforward.
                      </p>
                      
                      <h3 className="text-xl font-bold mt-8 mb-4">Prerequisites</h3>
                      <ul className="list-disc list-inside text-muted-foreground space-y-2 mb-6 ml-4">
                        <li>Windows 10 or Windows 11 (64-bit)</li>
                        <li>An active internet connection (for initial setup and syncing)</li>
                      </ul>

                      <h3 className="text-xl font-bold mt-8 mb-4">Deploying the Executable</h3>
                      <div className="bg-black/50 border border-white/10 rounded-xl p-4 mb-6 font-mono text-sm overflow-x-auto text-muted-foreground leading-relaxed">
                        1. Log into your dashboard and download your unique `ActivityTracker.exe`.<br/>
                        2. Place the executable on your target machine.<br/>
                        3. Double click the file. The local configuration server will boot on port 8580.<br/>
                        4. Enter your License Key when prompted on the startup screen.
                      </div>
                    </div>
                  )}

                  {/* Configuration */}
                  {activeTab === "configuration" && (
                    <div>
                      <h2 className="text-3xl font-bold flex items-center gap-3 mb-6">
                        <Settings className="text-primary" /> Configuration
                      </h2>
                      <p className="text-muted-foreground mb-6">
                        Configuration can be managed either via the local browser dashboard (`http://localhost:8580`) or via the `tracker_config.enc` file located in your AppData directory.
                      </p>
                      
                      <h3 className="text-xl font-bold mt-8 mb-4">Local AppData Storage</h3>
                      <p className="text-muted-foreground mb-4">
                        All configuration is stored locally to ensure complete privacy. The default path is:
                      </p>
                      <div className="bg-black/50 border border-white/10 rounded-xl p-4 mb-6 font-mono text-sm text-primary">
                        %LOCALAPPDATA%\ActivityTracker\
                      </div>

                      <h3 className="text-xl font-bold mt-8 mb-4">Important Settings</h3>
                      <ul className="space-y-4">
                         <li className="glass p-4 rounded-xl">
                            <span className="font-bold text-foreground">Idle Timeout</span>
                            <p className="text-sm text-muted-foreground mt-1">Defines how many seconds of zero keyboard/mouse input must pass before the tracker flags the user as "Idle". Default is 300 seconds (5 minutes).</p>
                         </li>
                         <li className="glass p-4 rounded-xl">
                            <span className="font-bold text-foreground">App Blocklist</span>
                            <p className="text-sm text-muted-foreground mt-1">A list of executable names (e.g. `spotify.exe`) that will be forcibly closed if the tracker detects them running during designated work hours.</p>
                         </li>
                      </ul>
                    </div>
                  )}

                  {/* Google OAuth Setup */}
                  {activeTab === "oauth" && (
                    <div>
                      <h2 className="text-3xl font-bold flex items-center gap-3 mb-6">
                        <ShieldCheck className="text-primary" /> Google OAuth Setup
                      </h2>
                      <p className="text-muted-foreground mb-6">
                        To push data to your Google Sheets, the application requires authorization via Google Cloud Console.
                      </p>
                      
                      <ol className="space-y-6 list-decimal list-inside text-muted-foreground">
                         <li>
                            <strong className="text-foreground">Create a Google Cloud Project:</strong> Navigate to the <a href="https://console.cloud.google.com" className="text-primary hover:underline" target="_blank" rel="noreferrer">Google Cloud Console</a> and create a new project.
                         </li>
                         <li>
                            <strong className="text-foreground">Enable the Sheets API:</strong> In the library, search for "Google Sheets API" and enable it for your project.
                         </li>
                         <li>
                            <strong className="text-foreground">Configure OAuth Consent Screen:</strong> Setup the consent screen for an "External" application.
                         </li>
                         <li>
                            <strong className="text-foreground">Create Credentials:</strong> Create an "OAuth client ID" for a "Desktop App". Download the resulting `credentials.json` file.
                         </li>
                         <li>
                            <strong className="text-foreground">Import Credentials:</strong> Upload this JSON file into the ActivityTracker local dashboard to link the accounts.
                         </li>
                      </ol>
                    </div>
                  )}

                  {/* Active Window */}
                  {activeTab === "active-window" && (
                    <div>
                      <h2 className="text-3xl font-bold flex items-center gap-3 mb-6">
                        <Monitor className="text-primary" /> Active Window Tracking
                      </h2>
                      <p className="text-muted-foreground mb-6">
                        Active Window tracking uses Windows UI Automation libraries to poll the topmost, active application exactly once every second.
                      </p>

                      <div className="glass p-6 border-l-4 border-primary rounded-r-xl mb-6 text-sm text-muted-foreground leading-relaxed">
                         <strong>Privacy Note:</strong> ActivityTracker only reads the Title text of the window and the process executable name. It never captures visual screenshots of the window contents.
                      </div>

                      <h3 className="text-xl font-bold mt-8 mb-4">Chrome Integration</h3>
                      <p className="text-muted-foreground">
                        If the active application is Google Chrome, the tracker attempts to isolate the specific active tab name rather than just logging "Google Chrome" universally. This provides deep insights into web-based workflows.
                      </p>
                    </div>
                  )}

                  {/* Idle Detection */}
                  {activeTab === "idle-detection" && (
                    <div>
                      <h2 className="text-3xl font-bold flex items-center gap-3 mb-6">
                        <Zap className="text-primary" /> Idle Detection
                      </h2>
                      <p className="text-muted-foreground mb-6">
                        Idle detection relies on the native Windows `GetLastInputInfo` system call. This is incredibly efficient and 100% privacy safe.
                      </p>

                      <p className="text-muted-foreground mb-6">
                        The API simply returns the number of milliseconds since the last hardware interrupt (keyboard key pressed down, or mouse moved 1 pixel). By comparing this to the "Idle Timeout" configuration, we can perfectly segment Active work time from Idle time without ever knowing *what* keys were typed.
                      </p>
                    </div>
                  )}

                  {/* Data Syncing */}
                  {activeTab === "data-syncing" && (
                    <div>
                      <h2 className="text-3xl font-bold flex items-center gap-3 mb-6">
                        <Server className="text-primary" /> Data Syncing
                      </h2>
                      <p className="text-muted-foreground mb-6">
                        ActivityTracker operates on an Offline-First architecture.
                      </p>

                      <ul className="space-y-4 text-muted-foreground">
                         <li>
                            <strong className="text-foreground">Step 1: In-Memory Buffer</strong> - Tracking events are batched in memory to prevent aggressive disk writing.
                         </li>
                         <li>
                            <strong className="text-foreground">Step 2: Local JSON Sync</strong> - Every 5 minutes, the buffer flushes to a local daily JSON file (e.g. `log_2026-06-11.json`).
                         </li>
                         <li>
                            <strong className="text-foreground">Step 3: Google Sheets Sync</strong> - Once per hour (or upon manual clicking), the engine parses the local JSON, aggregates the intervals, formats them nicely, and appends the final row to your linked Google Sheet.
                         </li>
                      </ul>
                    </div>
                  )}
                </motion.div>
              </AnimatePresence>
            </div>
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
