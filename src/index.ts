// src/index.ts
// ============================================================
// نقطة الدخول لـ Firebase Functions — إنشاء البوت آمن (lazy / مشروط)
// - لا ننشئ البوت إلا إذا وجد BOT_TOKEN.
// - إن لم يكن BOT_TOKEN مهيأ نصدّر endpoint يعيد 503 مع رسالة مفيدة.
// ============================================================
import * as admin from "firebase-admin";
import { webhookCallback } from "grammy";
import { createBot } from "./bot";

// نستخدم require لتجنب اختلاف أنواع الإصدارات بين firebase-functions
const functions: any = require("firebase-functions");

// تهيئة Firebase Admin فقط إذا لم تُهيَّأ أي تطبيق سابق
if (!admin.apps || admin.apps.length === 0) {
  admin.initializeApp();
}

// الحصول على توكن البوت بأمان: process.env أولاً ثم functions.config داخل try/catch
let BOT_TOKEN = process.env.BOT_TOKEN || "";

try {
  if (functions && typeof functions.config === "function") {
    const cfg = functions.config();
    if (cfg && cfg.bot && cfg.bot.token) {
      BOT_TOKEN = cfg.bot.token;
    }
  }
} catch (err) {
  // تجاهل: بيئة حديثة قد لا تدعم functions.config()
}

if (!BOT_TOKEN) {
  console.warn(
    "❌ BOT_TOKEN not configured! Set it via: firebase functions:config:set bot.token=YOUR_TOKEN OR set process.env.BOT_TOKEN"
  );
}

// Helper: functions.region wrapper (fall back to functions if region() not supported)
const functionsRegion = functions && typeof functions.region === "function"
  ? functions.region("us-central1")
  : functions;

// Export webhook handler:
// - إذا توكن موجود: أنشئ البوت وصدّر webhook callback
// - إذا لا توكن: صدّر health-like endpoint يرد 503 ويشرح المطلوب (آمن للاختبار والنشر)
let telegramWebhook: any;
if (BOT_TOKEN) {
  // أنشئ البوت مرة واحدة فقط
  const bot = createBot(BOT_TOKEN);
  if (functionsRegion.runWith) {
    telegramWebhook = functionsRegion
      .runWith({ timeoutSeconds: 60, memory: "512MB" })
      .https.onRequest(webhookCallback(bot, "express"));
  } else {
    telegramWebhook = functionsRegion.https.onRequest(webhookCallback(bot, "express"));
  }
} else {
  // توكن غير مهيأ — نرد برسالة توضيحية بدل محاولة إنشاء بوت فاشلة
  const handler = (req: any, res: any) => {
    res.status(503).json({
      ok: false,
      error: "BOT_TOKEN_NOT_CONFIGURED",
      message:
        "BOT_TOKEN is not configured. Set it via: firebase functions:config:set bot.token=YOUR_TOKEN or set process.env.BOT_TOKEN",
    });
  };
  if (functionsRegion.runWith) {
    telegramWebhook = functionsRegion
      .runWith({ timeoutSeconds: 60, memory: "128MB" })
      .https.onRequest(handler);
  } else {
    telegramWebhook = functionsRegion.https.onRequest(handler);
  }
}

// تصدير الـ webhook
export { telegramWebhook };

// Health-check endpoint (دائماً متاح)
export const healthStatus = (functions && typeof functions.region === "function"

  ? functions.region("us-central1")
  : functions
).https.onRequest((req: any, res: any) => {
  res.json({
    status: "healthy",
    timestamp: new Date().toISOString(),
    bot: BOT_TOKEN ? "configured" : "missing_token",
  });
});