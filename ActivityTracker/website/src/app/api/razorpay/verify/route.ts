import { NextRequest, NextResponse } from "next/server";
import crypto from "crypto";
import { getAdminDb } from "@/lib/firebaseAdmin";

export const dynamic = "force-dynamic";

export async function POST(req: NextRequest) {
  try {
    const {
      razorpay_order_id,
      razorpay_payment_id,
      razorpay_signature,
      userId,
      planDetails, // Contains { planId, planName, pcs, bonus, price, currency }
      upgradeFromId, // The ID of the plan being upgraded (if any)
    } = await req.json();

    const body = razorpay_order_id + "|" + razorpay_payment_id;

    const expectedSignature = crypto
      .createHmac("sha256", process.env.RAZORPAY_KEY_SECRET!)
      .update(body.toString())
      .digest("hex");

    const isAuthentic = expectedSignature === razorpay_signature;

    if (!isAuthentic) {
      return NextResponse.json(
        { error: "Invalid payment signature" },
        { status: 400 }
      );
    }

    // Payment is verified. Update the user's subscription in Firebase.
    let userRef;
    let actualDocId = userId;

    // Try to find the user document by uid
    const adminDb = getAdminDb();
    const usersSnapshot = await adminDb.collection("users").where("uid", "==", userId).limit(1).get();
    if (!usersSnapshot.empty) {
      userRef = usersSnapshot.docs[0].ref;
      actualDocId = usersSnapshot.docs[0].id;
    } else {
      // Fallback: assume userId is the document ID
      userRef = adminDb.collection("users").doc(userId);
    }

    // Calculate next payment date (1 month from now)
    const nextPaymentDate = new Date();
    nextPaymentDate.setMonth(nextPaymentDate.getMonth() + 1);

    const newSubId = `sub_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

    const newSubscription = {
      id: newSubId,
      active: true,
      planId: planDetails.planId,
      planName: planDetails.planName,
      pcs: planDetails.pcs,
      bonus: planDetails.bonus || 0,
      price: planDetails.price,
      currency: planDetails.currency,
      startDate: new Date().toISOString(),
      nextPayment: nextPaymentDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }),
      licenseKey: newSubId, // Use the unique ID as the license key instead of user uid
      lastUpdate: new Date().toISOString(),
    };

    const userDoc = await userRef.get();
    let currentSubscriptions: any[] = [];
    if (userDoc.exists) {
      const data = userDoc.data();
      if (data?.subscriptions && Array.isArray(data.subscriptions)) {
        currentSubscriptions = [...data.subscriptions];
      } else if (data?.subscription) {
        // Migration: keep old subscription but add an ID if it doesn't have one
        currentSubscriptions = [{ ...data.subscription, id: data.subscription.id || 'legacy_sub' }];
      }
    }

    if (upgradeFromId) {
      // Find the old plan and mark it inactive
      const oldPlanIndex = currentSubscriptions.findIndex(s => s.id === upgradeFromId || (upgradeFromId === 'legacy_sub' && !s.id));
      if (oldPlanIndex >= 0) {
        currentSubscriptions[oldPlanIndex].active = false;
        currentSubscriptions[oldPlanIndex].lastUpdate = new Date().toISOString();
      }
    }

    currentSubscriptions.push(newSubscription);

    const subscriptionUpdate = {
      uid: userId, 
      subscriptions: currentSubscriptions,
      // Keep legacy object updated with highest active for fallback
      subscription: currentSubscriptions.filter(s => s.active).sort((a, b) => (b.price || 0) - (a.price || 0))[0] || null,
    };

    await userRef.set(subscriptionUpdate, { merge: true });

    // Save invoice document
    const invoiceRef = userRef.collection("invoices").doc(razorpay_order_id);
    await invoiceRef.set({
      id: razorpay_order_id,
      paymentId: razorpay_payment_id,
      date: new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' }),
      amount: planDetails.price,
      currency: planDetails.currency,
      status: "Paid",
      plan: planDetails.planName,
      pcs: planDetails.pcs,
    });

    return NextResponse.json({ success: true, message: "Payment verified successfully" });
  } catch (error: any) {
    console.error("Error verifying payment:", error);
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
