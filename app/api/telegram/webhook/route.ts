import { NextRequest, NextResponse } from "next/server";
import { parseWhatsAppReservation } from "@/lib/whatsappParser";
import {
  saveReservationToSheets,
  getPickUpByDate,
  getVendorSummary,
} from "@/lib/googleSheetsManager";

const BOT_TOKEN = process.env.TELEGRAM_BOT_TOKEN!;
const ALLOWED_CHATS = (process.env.TELEGRAM_ALLOWED_CHATS || "")
  .split(",")
  .map((id) => id.trim())
  .filter(Boolean);

export const maxDuration = 60;

const processedUpdates = new Set<number>();

async function sendTelegramMessage(
  chatId: number | string,
  text: string,
  replyMarkup?: object
) {
  const res = await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      chat_id: chatId,
      text,
      ...(replyMarkup ? { reply_markup: replyMarkup } : {}),
    }),
  });

  if (!res.ok) {
    const errBody = await res.text();
    console.error("Error enviando mensaje a Telegram:", errBody);
  }
}

async function answerCallbackQuery(callbackQueryId: string) {
  await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/answerCallbackQuery`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ callback_query_id: callbackQueryId }),
  });
}

function getMainMenuKeyboard() {
  return {
    inline_keyboard: [
      [
        { text: "🚐 Pick-Up Hoy", callback_data: "pickup_today" },
        { text: "📅 Pick-Up Mañana", callback_data: "pickup_tomorrow" },
      ],
      [{ text: "💰 Cuentas Vendedores", callback_data: "vendor_summary" }],
    ],
  };
}

function getDateString(offsetDays = 0): string {
  const now = new Date();
  const drTime = new Date(
    now.toLocaleString("en-US", { timeZone: "America/Santo_Domingo" })
  );
  drTime.setDate(drTime.getDate() + offsetDays);
  const year = drTime.getFullYear();
  const month = String(drTime.getMonth() + 1).padStart(2, "0");
  const day = String(drTime.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

export async function POST(req: NextRequest) {
  try {
    const update = await req.json();

    if (update?.update_id) {
      if (processedUpdates.has(update.update_id)) {
        return NextResponse.json({ ok: true, duplicate: true });
      }
      processedUpdates.add(update.update_id);
    }

    // 1. Manejar botones inline (Callbacks)
    if (update.callback_query) {
      const callbackId = update.callback_query.id;
      const chatId = update.callback_query.message.chat.id;
      const data = update.callback_query.data as string;

      await answerCallbackQuery(callbackId);

      if (ALLOWED_CHATS.length > 0 && !ALLOWED_CHATS.includes(String(chatId))) {
        return NextResponse.json({ ok: true });
      }

      if (data === "pickup_today" || data === "pickup_tomorrow") {
        const offset = data === "pickup_today" ? 0 : 1;
        const targetDate = getDateString(offset);
        const label = data === "pickup_today" ? "HOY" : "MAÑANA";
        const rows = await getPickUpByDate(targetDate);

        if (rows.length === 0) {
          await sendTelegramMessage(
            chatId,
            `📭 No hay pick-ups registrados para ${label} (${targetDate}).`,
            getMainMenuKeyboard()
          );
          return NextResponse.json({ ok: true });
        }

        let msg = `🚐 LISTA DE PICK-UP ${label} (${targetDate})\nTotal reservas: ${rows.length}\n\n`;
        rows.forEach((r, idx) => {
          msg += `${idx + 1}. ⏰ ${r[1] || ""} | 🏨 ${r[2] || ""} (Hab: ${r[7] || "N/A"})\n`;
          msg += `   👤 ${r[5] || ""} (${r[6] || ""} PAX) — 🏝️ ${r[0] || ""}\n`;
          msg += `   💵 Deben: ${r[9] || "PAGADO"} | 🧑‍💼 Rep: ${r[4] || ""}\n\n`;
        });

        await sendTelegramMessage(chatId, msg, getMainMenuKeyboard());
        return NextResponse.json({ ok: true });
      }

      if (data === "vendor_summary") {
        const vendors = await getVendorSummary();

        if (vendors.length === 0) {
          await sendTelegramMessage(
            chatId,
            "📭 Aún no hay cuentas de vendedores registradas.",
            getMainMenuKeyboard()
          );
          return NextResponse.json({ ok: true });
        }

        let msg = `💰 RESUMEN CUENTAS VENDEDORES\n\n`;
        for (const v of vendors) {
          msg += `🧑‍💼 ${v[0]}\n`;
          msg += `   • Reservas: ${v[1]} (${v[2]} PAX)\n`;
          msg += `   • Vendido: $${v[3]} | Abonado: $${v[4]}\n`;
          msg += `   • Pendiente: $${v[5]}\n\n`;
        }

        await sendTelegramMessage(chatId, msg, getMainMenuKeyboard());
        return NextResponse.json({ ok: true });
      }

      return NextResponse.json({ ok: true });
    }

    // 2. Manejar mensajes de texto
    const message = update.message;
    if (!message || !message.text) {
      return NextResponse.json({ ok: true });
    }

    const chatId = message.chat.id;
    const text = message.text.trim();
    const senderName =
      [message.from?.first_name, message.from?.last_name]
        .filter(Boolean)
        .join(" ") || "Sin asignar";

    if (ALLOWED_CHATS.length > 0 && !ALLOWED_CHATS.includes(String(chatId))) {
      await sendTelegramMessage(chatId, "⛔ Chat no autorizado.");
      return NextResponse.json({ ok: true });
    }

    if (text === "/start" || text === "/menu") {
      await sendTelegramMessage(
        chatId,
        "🌴 JAC Tour and Transfers — Panel Operativo\n\nReenvía o pega aquí cualquier mensaje de reserva de WhatsApp para registrarlo automáticamente en Google Sheets, o usa los botones:",
        getMainMenuKeyboard()
      );
      return NextResponse.json({ ok: true });
    }

    await sendTelegramMessage(
      chatId,
      "⏳ Analizando reserva y sincronizando con Google Sheets..."
    );

    const parsed = await parseWhatsAppReservation(text, senderName);

    if (!parsed.isReservation) {
      await sendTelegramMessage(
        chatId,
        "ℹ️ El mensaje no parece contener datos de una reserva turística. Pega un voucher o mensaje con fecha, hotel y excursión.",
        getMainMenuKeyboard()
      );
      return NextResponse.json({ ok: true });
    }

    const saved = await saveReservationToSheets(parsed);

    const estadoTexto =
      parsed.saldoPendiente > 0
        ? `Pendiente: $${parsed.saldoPendiente}`
        : parsed.precioTotal > 0
          ? "Pagado"
          : "Monto por confirmar";

    const confirmationMsg =
      `✅ RESERVA REGISTRADA (${saved.reservaId})\n\n` +
      `📅 Fecha: ${parsed.fechaTour} — ⏰ ${parsed.horaPickUp}\n` +
      `🏝️ Excursión: ${parsed.excursion}\n` +
      `👤 Cliente: ${parsed.nombreCliente} (${saved.totalPax} PAX)\n` +
      `🏨 Pick-Up: ${parsed.hotelPuntoEncuentro} (Hab: ${parsed.habitacion || "N/A"})\n` +
      `💼 Vendedor: ${parsed.vendedor}\n` +
      `💵 Total: $${parsed.precioTotal} ${parsed.moneda} | ${estadoTexto}\n\n` +
      `📂 Guardado en Reservas_Master, ${saved.pickUpTab || "pendiente de asignar a fecha"}, Cuentas_Vendedores y ${saved.vendorTab}.`;

    await sendTelegramMessage(chatId, confirmationMsg, getMainMenuKeyboard());

    return NextResponse.json({ ok: true });
  } catch (error) {
    console.error("Error en webhook Telegram:", error);
    return NextResponse.json(
      { ok: false, error: "Internal Server Error" },
      { status: 500 }
    );
  }
}