"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { Activity } from "lucide-react";

export default function TermsPage() {
  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24 flex flex-col">
      <Navbar />
      
      <section className="relative py-20 px-6 flex-1">
        <div className="container mx-auto max-w-3xl relative z-10">
          <div className="mb-12">
            <h1 className="text-4xl font-extrabold tracking-tight mb-4">Terms of Service</h1>
            <p className="text-muted-foreground">Last updated: June 10, 2026</p>
          </div>

          <div className="prose prose-invert max-w-none text-muted-foreground leading-relaxed space-y-6">
            <h2 className="text-2xl font-bold text-foreground mt-8 mb-4">1. Acceptance of Terms</h2>
            <p>
              By accessing and using ActivityTracker, you agree to be bound by these Terms of Service. If you disagree with any part of the terms, you do not have permission to access the service.
            </p>

            <h2 className="text-2xl font-bold text-foreground mt-8 mb-4">2. License to Use</h2>
            <p>
              We grant you a personal, non-exclusive, non-transferable, limited license to use ActivityTracker software for personal and commercial purposes in accordance with these terms.
            </p>

            <h2 className="text-2xl font-bold text-foreground mt-8 mb-4">3. Prohibited Uses</h2>
            <p>
              You agree not to use the software for any unlawful purpose, or to monitor individuals without their explicit consent. ActivityTracker is built for personal productivity management and should not be used as spyware or stalkerware.
            </p>

            <h2 className="text-2xl font-bold text-foreground mt-8 mb-4">4. Subscriptions and Billing</h2>
            <p>
              Certain premium features may require a paid subscription. All billing is handled securely. You may cancel your subscription at any time, but no refunds will be issued for partial billing periods.
            </p>

            <h2 className="text-2xl font-bold text-foreground mt-8 mb-4">5. Limitation of Liability</h2>
            <p>
              In no event shall ActivityTracker, nor its directors, employees, partners, agents, suppliers, or affiliates, be liable for any indirect, incidental, special, consequential or punitive damages, including without limitation, loss of profits, data, use, goodwill, or other intangible losses, resulting from your access to or use of or inability to access or use the Service.
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
