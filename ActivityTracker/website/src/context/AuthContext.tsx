"use client";

import React, { createContext, useContext, useEffect, useState } from "react";
import { auth, db } from "../lib/firebase";
import { onAuthStateChanged, User } from "firebase/auth";
import { doc, getDoc, onSnapshot, collection, query, where } from "firebase/firestore";

interface Subscription {
  id?: string;
  active: boolean;
  planId?: string;
  planName?: string;
  pcs?: number;
  bonus?: number;
  nextPayment?: string;
  startDate?: string;
  price?: number;
  currency?: string;
  licenseKey?: string;

  lastUpdate?: string;
}

interface AuthContextType {
  user: User | null;
  loading: boolean;
  subscription: Subscription | null;
  subscriptions: Subscription[];
  userDocId: string | null;
}

const AuthContext = createContext<AuthContextType>({
  user: null,
  loading: true,
  subscription: null,
  subscriptions: [],
  userDocId: null,
});

export const useAuth = () => useContext(AuthContext);

export const AuthProvider = ({ children }: { children: React.ReactNode }) => {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [subscription, setSubscription] = useState<Subscription | null>(null);
  const [subscriptions, setSubscriptions] = useState<Subscription[]>([]);
  const [userDocId, setUserDocId] = useState<string | null>(null);

  useEffect(() => {
    let unsubDoc: (() => void) | undefined;

    const unsubscribeAuth = onAuthStateChanged(auth, async (currentUser) => {
      if (currentUser) {
        if (typeof window !== "undefined") localStorage.setItem("auth_user", "true");
        setUser(currentUser);
        // Subscribe to user document changes using query
        const userDocQuery = query(collection(db, "users"), where("uid", "==", currentUser.uid));
        unsubDoc = onSnapshot(userDocQuery, async (querySnapshot) => {
          let foundDocId = null;
          let foundSubs: Subscription[] = [];

          if (!querySnapshot.empty) {
            const docSnap = querySnapshot.docs[0];
            foundDocId = docSnap.id;
            const data = docSnap.data();
            if (data.subscriptions && Array.isArray(data.subscriptions)) {
              foundSubs = data.subscriptions;
            } else if (data.subscription) {
              foundSubs = [data.subscription]; // Fallback for old data
            }
          }

          // Fallback: If no subscription found by uid query, try fetching by document ID = currentUser.uid
          if (foundSubs.length === 0) {
            try {
              const fallbackDocRef = doc(db, "users", currentUser.uid);
              const fallbackSnap = await getDoc(fallbackDocRef);
              if (fallbackSnap.exists()) {
                const fallbackData = fallbackSnap.data();
                if (fallbackData.subscriptions && Array.isArray(fallbackData.subscriptions)) {
                  foundSubs = fallbackData.subscriptions;
                  foundDocId = fallbackSnap.id;
                } else if (fallbackData.subscription) {
                  foundSubs = [fallbackData.subscription];
                  foundDocId = fallbackSnap.id; 
                }
              }
            } catch (err) {
              console.warn("Fallback doc fetch failed:", err);
            }
          }

          const activeSubs = foundSubs.filter(s => s.active);
          const highestActive = activeSubs.sort((a, b) => (b.price || 0) - (a.price || 0))[0] || null;

          setUserDocId(foundDocId);
          setSubscriptions(foundSubs);
          setSubscription(highestActive);
          setLoading(false);
        }, (error) => {
           console.warn("User document snapshot listener closed/error:", error.message);
           setLoading(false);
        });
      } else {
        if (typeof window !== "undefined") localStorage.removeItem("auth_user");
        setUser(null);
        setSubscription(null);
        setSubscriptions([]);
        setUserDocId(null);
        setLoading(false);
        if (unsubDoc) {
          unsubDoc();
          unsubDoc = undefined;
        }
      }
    });

    return () => {
      unsubscribeAuth();
      if (unsubDoc) {
        unsubDoc();
      }
    };
  }, []);

  return (
    <AuthContext.Provider value={{ user, loading, subscription, subscriptions, userDocId }}>
      {children}
    </AuthContext.Provider>
  );
};
