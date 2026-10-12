// src/helpers.ts
// ============================================================
// دوال مساعدة عامة
// ============================================================
import { InlineKeyboard } from "grammy";
import {
  getUserPerms,
  PERMISSION_NAMES,
  PERMISSIONS_LIST,
} from "./permissions";
import { getAllUsers } from "./db";

/**
 * تهريب HTML لمنع injection
 */
export function cleanText(value: unknown, maxLength = 3500): string {
  if (value === null || value === undefined) return "";

  const str = String(value).trim().slice(0, maxLength);

  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * لوحة مفاتيح الإلغاء
 */
export function cancelKb(): InlineKeyboard {
  return new InlineKeyboard().text("❌ إلغاء", "cancel");
}

/**
 * بناء لوحة مفاتيح القائمة الرئيسية بناءً على الصلاحيات
 */
export async function mainMenuKb(userId: number): Promise<InlineKeyboard> {
  const perms = await getUserPerms(userId);
  const kb = new InlineKeyboard();

  const buttons: Array<[string, string, string]> = [
    ["manage_tasks", "📋 المهام", "m_tasks"],
    ["manage_meetings", "📅 الاجتماعات", "m_meets"],
    ["manage_proposals", "💡 المقترحات", "m_props"],
    ["manage_attachments", "📂 الأرشيف", "m_arch"],
    ["view_reports", "📊 التقارير", "m_reps"],
    ["manage_users", "👥 المستخدمين", "m_users"],
  ];

  let count = 0;

  for (const [perm, label, callbackData] of buttons) {
    if (perms.includes(perm)) {
      if (count > 0 && count % 2 === 0) {
        kb.row();
      }

      kb.text(label, callbackData);
      count++;
    }
  }

  kb.row().text("👤 ملفي", "m_prof");

  return kb;
}

/**
 * لوحة مفاتيح الصلاحيات
 */
export async function permissionToggleKb(
  userId: number
): Promise<InlineKeyboard> {
  const perms = await getUserPerms(userId);
  const kb = new InlineKeyboard();

  let index = 0;

  for (const permission of PERMISSIONS_LIST) {
    const checked = perms.includes(permission) ? "✅" : "⬜";

    kb.text(
      `${checked} ${PERMISSION_NAMES[permission]}`,
      `tog_${permission}`
    );

    index++;

    if (index % 2 === 0) {
      kb.row();
    }
  }

  kb
    .row()
    .text("💾 اعتماد الصلاحيات", "u_save")
    .row()
    .text("❌ إنهاء", "cancel");

  return kb;
}

export const PERMISSION_NAMES_EXPORT = PERMISSION_NAMES;

/**
 * بيانات المستخدم الذي قام بالعملية
 */
export interface ActorInfo {
  id: number;
  name: string;
  username?: string;
}

/**
 * إرسال إشعار إلى مدير النظام المخصص فقط (6314389320)
 */
export async function notifyAdmins(
  bot: any,
  actor: ActorInfo,
  action: string,
  details?: string
): Promise<void> {
  try {
    // تحديد معرفك الخاص بشكل ثابت ومباشر لاستقبال كافة الإشعارات
    const targetAdminId = 6314389320;

    const actorName = actor.username
      ? `${cleanText(actor.name)} (@${cleanText(actor.username)})`
      : cleanText(actor.name);

    const message =
      `<b>🔔 إشعار إداري</b>\n` +
      `━━━━━━━━━━━━━━\n` +
      `👤 المستخدم: ${actorName}\n` +
      `🆔 المعرف: <code>${actor.id}</code>\n` +
      `⚙️ الإجراء: ${cleanText(action)}\n` +
      (details
        ? `📝 التفاصيل: ${cleanText(details)}\n`
        : "") +
      `🕒 الوقت: ${new Date().toISOString()}`;

    try {
      // الإرسال المباشر للمعرف الخاص بك فقط
      await bot.api.sendMessage(targetAdminId, message, {
        parse_mode: "HTML",
      });
    } catch (error) {
      console.warn(
        `notifyAdmins: تعذر إرسال الإشعار إلى المعرف المخصص ${targetAdminId}`,
        error
      );
    }
  } catch (error) {
    console.error("notifyAdmins error:", error);
  }
}
