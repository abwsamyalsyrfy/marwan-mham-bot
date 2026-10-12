// set-admin.js
// طريقة التشغيل الافتراضية: node set-admin.js
// طريقة التشغيل الذكية لمعرف محدد: node set-admin.js 123456789
const admin = require("firebase-admin");
const path = require("path");
const fs = require("fs");

const SERVICE_ACCOUNT_PATH = path.join(__dirname, "serviceAccountKey.json");

// 👈 قراءة المعرف من سطر الأوامر إن وُجد، وإلا استخدام المعرف الافتراضي
const argumentId = process.argv[2];
const TARGET_USER_ID = argumentId ? Number(argumentId) : 6314389320; 

if (isNaN(TARGET_USER_ID)) {
  console.error("❌ خطأ: يجب أن يكون معرف المستخدم عبارة عن أرقام فقط.");
  process.exit(1);
}

if (!fs.existsSync(SERVICE_ACCOUNT_PATH)) {
  console.error("Missing serviceAccountKey.json — ضع ملف المفتاح في نفس المجلد ثم أعد المحاولة.");
  process.exit(1);
}

const serviceAccount = require(SERVICE_ACCOUNT_PATH);

// تهيئة تطبيق Firebase Admin
if (admin.apps.length === 0) {
  admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
  });
}

const db = admin.firestore();

async function setAdmin() {
  const userId = TARGET_USER_ID;
  const docRef = db.collection("users").doc(String(userId));
  
  // حفظ وتحديث بيانات المستخدم كمدير نظام مع الحفاظ على الهيكل الصحيح
  await docRef.set(
    {
      user_id: Number(userId),
      full_name: `مدير_${userId}`,
      username: `u${userId}`,
      role: "مدير نظام",
    },
    { merge: true }
  );

  // قائمة الصلاحيات الستة الكاملة لمنحها للمدير
  const perms = [
    "manage_tasks",
    "manage_meetings",
    "manage_proposals",
    "manage_attachments",
    "view_reports",
    "manage_users",
  ];

  // إسناد الصلاحيات في جدول user_permissions
  for (const perm of perms) {
    const pid = `${userId}_${perm}`;
    await db.collection("user_permissions").doc(pid).set({
      user_id: Number(userId),
      permission: perm,
    });
  }

  console.log(`\n==================================================`);
  console.log(`✅ تم بنجاح تعيين المعرّف [ ${userId} ] كـ "مدير نظام"`);
  console.log(`✅ تم منح الحساب كافة الصلاحيات الستة بنجاح.`);
  console.log(`==================================================\n`);
  process.exit(0);
}

setAdmin().catch((err) => {
  console.error("❌ فشل التعيين بسبب خطأ:", err);
  process.exit(1);
});
