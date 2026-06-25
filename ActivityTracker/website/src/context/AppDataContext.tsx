"use client";

import React, { createContext, useContext, useEffect, useState } from "react";
import { doc, getDoc, setDoc, collection, getDocs, addDoc } from "firebase/firestore";
import { db } from "@/lib/firebase";

interface AppConfig {
  review_increase: number;
  average_rating: number;
}

interface AppPricing {
  [key: string]: number;
}

export interface Review {
  name: string;
  star: number;
  review: string;
}

interface AppDataContextType {
  config: AppConfig | null;
  pricing: AppPricing | null;
  reviews: Review[];
  displayRating: string;
  displayReviews: string;
  loading: boolean;
}

const defaultReviews: Review[] = [
  {
    name: "Alex M.",
    review: "The idle detection via Windows API is a lifesaver. I used to forget to pause other trackers when taking a break. ActivityTracker handles it perfectly and the Google Sheets sync is flawless.",
    star: 5
  },
  {
    name: "Sarah T.",
    review: "I love that it tracks my keyboard and mouse intervals without logging keystrokes. I feel completely secure while knowing exactly how productive my team and I are being.",
    star: 5
  },
  {
    name: "James K.",
    review: "Chrome website tracking gives me the hard truth about how much time I waste on social media versus actual research. The automated local JSON backups give me peace of mind when my internet drops.",
    star: 4
  },
  {
    name: "Emily R.",
    review: "Having all the data pumped directly into Google Sheets means I can build my own pivot tables and custom dashboards. The OAuth 2.0 integration was super simple to set up.",
    star: 5
  },
  {
    name: "Michael B.",
    review: "I set it to auto-start on boot and haven't touched the config since. It runs silently, eats almost zero RAM, and gives me a perfect daily log. Highly recommended.",
    star: 5
  },
  {
    name: "Jessica W.",
    review: "The active window tracking accurately logs when I'm in Photoshop versus Illustrator. It's totally changed how I bill clients because I have hard data on my time.",
    star: 5
  }
];

const AppDataContext = createContext<AppDataContextType>({
  config: null,
  pricing: null,
  reviews: [],
  displayRating: "0.0",
  displayReviews: "0",
  loading: true,
});

export function AppDataProvider({ children }: { children: React.ReactNode }) {
  const [config, setConfig] = useState<AppConfig | null>(null);
  const [pricing, setPricing] = useState<AppPricing | null>(null);
  const [reviews, setReviews] = useState<Review[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let isMounted = true;

    const fetchAppData = async () => {
      try {
        const configRef = doc(db, "app_data", "config");
        const pricingRef = doc(db, "app_data", "pricing");
        const reviewsCollectionRef = collection(db, "reviews");

        const [configSnap, pricingSnap, reviewsQuerySnap] = await Promise.all([
          getDoc(configRef),
          getDoc(pricingRef),
          getDocs(reviewsCollectionRef)
        ]);

        if (isMounted) {
          if (configSnap.exists()) {
            setConfig(configSnap.data() as AppConfig);
          } else {
            // Default fallbacks if doc doesn't exist yet
            const defaultConfig = { review_increase: 0, average_rating: 0 };
            setConfig(defaultConfig);
            // Auto-create document
            try { await setDoc(configRef, defaultConfig); } catch (e) { console.warn("Write permission denied for config"); }
          }

          if (pricingSnap.exists()) {
            setPricing(pricingSnap.data() as AppPricing);
          } else {
            const defaultPricing = {
              "single": 300,
              "pack-5": 1400,
              "pack-10": 2700,
              "pack-30": 8500,
              "pack-60": 16000,
              "pack-100": 28000
            };
            setPricing(defaultPricing);
            // Auto-create document
            try { await setDoc(pricingRef, defaultPricing); } catch (e) { console.warn("Write permission denied for pricing"); }
          }

          let loadedReviews: Review[] = [];
          if (!reviewsQuerySnap.empty) {
            loadedReviews = reviewsQuerySnap.docs.map(doc => doc.data() as Review);
          } else {
            loadedReviews = defaultReviews;
            // Auto-create default reviews documents
            try {
              await Promise.all(defaultReviews.map(rev => addDoc(reviewsCollectionRef, rev)));
            } catch (e) {
              console.warn("Write permission denied for reviews");
            }
          }
          setReviews(loadedReviews);

          setLoading(false);
        }
      } catch (error) {
        console.error("Error fetching app_data from Firebase:", error);
        if (isMounted) setLoading(false);
      }
    };

    fetchAppData();

    return () => {
      isMounted = false;
    };
  }, []);

  const originalCount = reviews.length;
  const originalSum = reviews.reduce((acc, r) => acc + (r.star || 0), 0);
  const originalAverage = originalCount > 0 ? (originalSum / originalCount) : 0;
  const avgOffset = config?.average_rating ? Number(config.average_rating) : 0;
  const revOffset = config?.review_increase ? Number(config.review_increase) : 0;
  
  const displayRating = (originalAverage + avgOffset).toFixed(1);
  const displayReviews = (originalCount + revOffset).toLocaleString();

  return (
    <AppDataContext.Provider value={{ config, pricing, reviews, displayRating, displayReviews, loading }}>
      {children}
    </AppDataContext.Provider>
  );
}

export const useAppData = () => useContext(AppDataContext);
