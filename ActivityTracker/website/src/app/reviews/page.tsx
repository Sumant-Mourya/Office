"use client";
import Image from "next/image";

import Navbar from "@/components/Navbar";
import { motion } from "framer-motion";
import { Activity, Star, Quote, Edit3, X } from "lucide-react";
import { useAuth } from "@/context/AuthContext";
import { useAppData } from "@/context/AppDataContext";
import { useRouter } from "next/navigation";
import { useState } from "react";

import { collection, addDoc } from "firebase/firestore";
import { db } from "@/lib/firebase";

export default function ReviewsPage() {
  const { config, reviews, displayRating, displayReviews, loading: configLoading } = useAppData();

  const { user } = useAuth();
  const router = useRouter();
  const [showModal, setShowModal] = useState(false);
  const [newReviewText, setNewReviewText] = useState("");
  const [newRating, setNewRating] = useState(5);

  const handleWriteReview = () => {
    if (!user) {
      router.push("/login?redirect=/reviews");
    } else {
      setShowModal(true);
    }
  };

  const handleSubmitReview = async (e: React.FormEvent) => {
    e.preventDefault();
    if (newReviewText.trim()) {
      const newReview = {
        name: user?.displayName || user?.email?.split('@')[0] || "User",
        review: newReviewText,
        star: newRating
      };
      
      try {
        await addDoc(collection(db, "reviews"), newReview);
        setShowModal(false);
        setNewReviewText("");
        setNewRating(5);
        window.location.reload();
      } catch (err) {
        console.error("Failed to submit review", err);
      }
    }
  };

  return (
    <main className="min-h-screen bg-background text-foreground overflow-x-hidden pt-24">
      <Navbar />
      
      <section className="relative py-20 px-6">
        <div className="absolute top-[-20%] left-[-10%] w-[50%] h-[50%] bg-accent/20 blur-[150px] rounded-full pointer-events-none" />
        
        <div className="container mx-auto max-w-6xl relative z-10 text-center mb-20">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="flex justify-center mb-8"
          >
            {configLoading ? (
              <div className="inline-flex items-center gap-2 px-4 py-2 rounded-full border border-accent/10 bg-accent/5 animate-pulse">
                <div className="w-4 h-4 rounded-full bg-accent/20"></div>
                <div className="w-32 h-4 rounded bg-white/10"></div>
              </div>
            ) : (
              <div className="inline-flex items-center gap-2 px-4 py-2 rounded-full glass text-sm font-medium text-accent">
                <Star size={16} className="fill-accent text-accent" />
                <span>Trusted by {displayReviews}+ users • {displayRating}/5</span>
              </div>
            )}
          </motion.div>
          <motion.h1
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.1 }}
            className="text-5xl md:text-6xl font-extrabold tracking-tight mb-6"
          >
            Loved by users.<br />
            <span className="text-gradient">Driven by data.</span>
          </motion.h1>
          <motion.p
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.2 }}
            className="text-xl text-muted-foreground max-w-2xl mx-auto"
          >
            See how ActivityTracker is helping professionals take control of their time with precise, privacy-first analytics.
          </motion.p>
          <motion.button
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.3 }}
            onClick={handleWriteReview}
            className="mt-8 bg-primary hover:bg-primary/90 text-white px-8 py-4 rounded-full font-semibold transition-all hover:scale-105 shadow-[0_0_30px_rgba(59,130,246,0.4)] flex items-center gap-2 mx-auto"
          >
            <Edit3 size={20} /> Write your own review
          </motion.button>
        </div>

        <div className="container mx-auto max-w-6xl">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8 mb-12">
            {reviews.slice(0, 6).map((review, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, scale: 0.95 }}
                whileInView={{ opacity: 1, scale: 1 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.1 }}
                className="glass p-8 rounded-3xl relative hover:bg-white/5 transition-colors flex flex-col"
              >
                <Quote size={40} className="absolute top-6 right-6 text-white/5" />
                <div className="flex gap-1 mb-6">
                  {[...Array(review.star)].map((_, i) => (
                    <Star key={i} size={18} className="fill-yellow-500 text-yellow-500" />
                  ))}
                  {[...Array(5 - (review.star || 0))].map((_, i) => (
                    <Star key={i} size={18} className="text-white/20" />
                  ))}
                </div>
                <p className="text-foreground/90 leading-relaxed mb-8 relative z-10 flex-1">
                  "{review.review}"
                </p>
                <div className="flex items-center gap-4 mt-auto">
                  <div className="w-10 h-10 rounded-full bg-gradient-to-tr from-primary to-accent flex items-center justify-center text-white font-bold text-sm">
                    {review.name.charAt(0)}
                  </div>
                  <div>
                    <div className="font-bold text-sm">{review.name}</div>
                  </div>
                </div>
              </motion.div>
            ))}
          </div>
          
          <motion.div 
            initial={{ opacity: 0 }}
            whileInView={{ opacity: 1 }}
            viewport={{ once: true }}
            className="text-center bg-white/5 border border-white/10 rounded-3xl py-8 max-w-sm mx-auto shadow-2xl backdrop-blur-md"
          >
            <div className="text-5xl font-black text-white mb-2">{displayReviews}+</div>
            <div className="text-muted-foreground font-medium uppercase tracking-widest text-sm">Total Reviews</div>
            <div className="flex justify-center items-center gap-2 mt-4 text-yellow-500 font-bold">
              <Star size={20} className="fill-yellow-500" />
              <span>{displayRating} Average Rating</span>
            </div>
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

      {/* Write Review Modal */}
      {showModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
          <motion.div
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-card border border-white/10 rounded-3xl p-8 max-w-md w-full shadow-2xl relative"
          >
            <button 
              onClick={() => setShowModal(false)}
              className="absolute top-6 right-6 text-muted-foreground hover:text-white transition-colors"
            >
              <X size={24} />
            </button>
            <h3 className="text-2xl font-bold mb-6">Write a Review</h3>
            <form onSubmit={handleSubmitReview} className="space-y-6">
              <div>
                <label className="block text-sm font-bold text-foreground mb-2">Rating</label>
                <div className="flex gap-2">
                  {[1, 2, 3, 4, 5].map((star) => (
                    <button
                      key={star}
                      type="button"
                      onClick={() => setNewRating(star)}
                      className="transition-transform hover:scale-110"
                    >
                      <Star size={28} className={star <= newRating ? "fill-yellow-500 text-yellow-500" : "text-white/20"} />
                    </button>
                  ))}
                </div>
              </div>
              <div>
                <label className="block text-sm font-bold text-foreground mb-2">Your Review</label>
                <textarea
                  required
                  value={newReviewText}
                  onChange={(e) => setNewReviewText(e.target.value)}
                  className="w-full bg-white/5 border border-white/10 rounded-xl py-3 px-4 text-foreground focus:outline-none focus:ring-2 focus:ring-primary min-h-[120px] resize-none transition-all"
                  placeholder="Tell us what you think about ActivityTracker..."
                />
              </div>
              <button
                type="submit"
                className="w-full bg-primary hover:bg-primary/90 text-white font-bold py-3 rounded-xl transition-all hover:scale-[1.02] shadow-lg shadow-primary/25"
              >
                Submit Review
              </button>
            </form>
          </motion.div>
        </div>
      )}
    </main>
  );
}
