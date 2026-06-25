"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Activity, Download, FileText } from "lucide-react";
import Link from "next/link";

import { useEffect, useState } from "react";
import { useAuth } from "@/context/AuthContext";
import { db } from "@/lib/firebase";
import { collection, query, orderBy, getDocs } from "firebase/firestore";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

export default function BillingHistoryPage() {
  const { user, userDocId } = useAuth();
  const [invoices, setInvoices] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const fetchInvoices = async () => {
      if (!user || !userDocId) return;
      try {
        const q = query(collection(db, "users", userDocId, "invoices"), orderBy("date", "desc"));
        const snapshot = await getDocs(q);
        const invData = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        setInvoices(invData);
      } catch (error) {
        console.error("Error fetching invoices:", error);
      } finally {
        setLoading(false);
      }
    };
    fetchInvoices();
  }, [user, userDocId]);

  const downloadInvoice = (invoice: any) => {
    const doc = new jsPDF();
    
    // Header
    doc.setFontSize(22);
    doc.setTextColor(59, 130, 246); // Primary color
    doc.text("ActivityTracker", 14, 20);
    
    doc.setFontSize(16);
    doc.setTextColor(0, 0, 0);
    doc.text("INVOICE", 14, 30);
    
    doc.setFontSize(10);
    doc.setTextColor(100, 100, 100);
    doc.text(`Invoice ID: ${invoice.id}`, 14, 40);
    doc.text(`Date: ${invoice.date}`, 14, 46);
    doc.text(`Payment ID: ${invoice.paymentId || 'N/A'}`, 14, 52);
    doc.text(`Status: ${invoice.status}`, 14, 58);
    
    // Billed To
    doc.setFontSize(12);
    doc.setTextColor(0, 0, 0);
    doc.text("Billed To:", 14, 70);
    doc.setFontSize(10);
    doc.setTextColor(100, 100, 100);
    doc.text(user?.email || "Customer", 14, 76);

    // Table
    autoTable(doc, {
      startY: 90,
      head: [['Description', 'Quantity (PCs)', 'Amount']],
      body: [
        [`${invoice.plan} Subscription`, invoice.pcs?.toString() || "-", `${invoice.amount} ${invoice.currency}`],
      ],
      theme: 'grid',
      headStyles: { fillColor: [59, 130, 246] },
    });

    // Total
    const finalY = (doc as any).lastAutoTable.finalY || 120;
    doc.setFontSize(12);
    doc.setTextColor(0, 0, 0);
    doc.text(`Total Amount Paid: ${invoice.amount} ${invoice.currency}`, 14, finalY + 15);
    
    // Footer
    doc.setFontSize(10);
    doc.setTextColor(150, 150, 150);
    doc.text("Thank you for your business!", 14, finalY + 35);

    doc.save(`Invoice_${invoice.id}.pdf`);
  };

  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24 flex flex-col">
      <Navbar />
      
      <section className="relative py-20 px-6 flex-1 flex flex-col items-center">
        <div className="absolute top-[10%] left-[-10%] w-[50%] h-[50%] bg-accent/20 blur-[150px] rounded-full pointer-events-none" />
        
        <div className="container mx-auto max-w-4xl relative z-10">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="mb-12"
          >
            <div className="flex items-center gap-3 mb-4">
              <Link href="/subscription" className="text-muted-foreground hover:text-primary transition-colors text-sm">
                 &larr; Back to Subscription
              </Link>
            </div>
            <h1 className="text-4xl md:text-5xl font-extrabold tracking-tight mb-2">Billing History</h1>
            <p className="text-lg text-muted-foreground leading-relaxed">
              View and download your past invoices.
            </p>
          </motion.div>

          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.1 }}
            className="glass rounded-3xl overflow-hidden border border-white/10"
          >
            <div className="overflow-x-auto">
               <table className="w-full text-left">
                  <thead>
                     <tr className="border-b border-white/10 bg-black/40 text-muted-foreground text-sm uppercase tracking-wider">
                        <th className="p-6 font-medium">Invoice</th>
                        <th className="p-6 font-medium">Date</th>
                        <th className="p-6 font-medium">Plan</th>
                        <th className="p-6 font-medium">Amount</th>
                        <th className="p-6 font-medium">Status</th>
                        <th className="p-6 font-medium text-right">Download</th>
                     </tr>
                  </thead>
                  <tbody className="divide-y divide-white/5">
                     {loading ? (
                       <tr>
                         <td colSpan={6} className="p-6 text-center text-muted-foreground">Loading invoices...</td>
                       </tr>
                     ) : invoices.length === 0 ? (
                       <tr>
                         <td colSpan={6} className="p-6 text-center text-muted-foreground">No billing history found.</td>
                       </tr>
                     ) : invoices.map((invoice) => (
                        <tr key={invoice.id} className="hover:bg-white/5 transition-colors">
                           <td className="p-6 font-mono text-sm">{invoice.id}</td>
                           <td className="p-6 text-sm">{invoice.date}</td>
                           <td className="p-6 text-sm text-muted-foreground">{invoice.plan}</td>
                           <td className="p-6 font-bold">{invoice.amount}</td>
                           <td className="p-6">
                              <span className="px-3 py-1 rounded-full bg-green-500/10 text-green-400 text-xs font-bold uppercase tracking-wider">
                                 {invoice.status}
                              </span>
                           </td>
                           <td className="p-6 text-right">
                              <button 
                                onClick={() => downloadInvoice(invoice)}
                                className="inline-flex items-center justify-center w-10 h-10 rounded-xl bg-white/5 hover:bg-white/10 text-primary transition-colors"
                              >
                                 <Download size={18} />
                              </button>
                           </td>
                        </tr>
                     ))}
                  </tbody>
               </table>
            </div>
          </motion.div>
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
