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

async function sendTelegramMessage(
  chatId: number | string,
  text: string,
  replyMarkup?: object
) {
  await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      chat_id: chatId,
      text,
      parse_mode: "HTML",
      ...(replyMarkup ? { reply_markup: replyMarkup } : {}),
    }),
  });
}

function getDateOffset(days: number): string {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.toISOString().split("T")[0];
}

export async function POST(req: NextRequest) {
  try {
    const update = await req.json();

    // Manejar botones inline (Callbacks)
    if (update.callback_query) {
      const chatId = update.callback_query.message.chat.id;
      const data = update.callback_query.data as string;

      if (ALLOWED_CHATS.length > 0 && !ALLOWED_CHATS.includes(String(chatId))) {
        return NextResponse.json({ ok: true });
      }

      if (data === "pickup_hoy" || data === "pickup_manana") {
        const targetDate = data === "pickup_hoy" ? getDateOffset(0) : getDateOffset(1);
        const rows = await getPickUpByDate(targetDate);

        if (rows.length === 0) {
          await sendTelegramMessage(
            chatId,
            `📋 No hay pasajeros registrados en <b>PickUp_${targetDate}</b>.`
          );
        } else {
          const list = rows
            .map(
              (r, idx) =>
                `${idx + 1}. ⏰ <b>${r[0]}</b> | 🏨 ${r[1]} (Hab: ${r[2]})\n   👤 ${r[3]} (${r[6]}) — <i>${r[5]}</i>\n   💵 Cobro: <b>${r[7]}</b> | Vendedor: ${r[8]}`
            )
            .join("\n\n");

          await sendTelegramMessage(
            chatId,
            `🚐 <b>LISTA PICK-UP (${targetDate})</b>\nTotal reservas: ${rows.length}\n\n${list}`
          );
        }
      }

      if (data === "cuentas_vendedores") {
        const vendors = await getVendorSummary();
        if (vendors.length === 0) {
          await sendTelegramMessage(chatId, "📊 Aún no hay cuentas de vendedores registradas.");
        } else {
          const summaryText = vendors
            .map(
              (v) =>
                `👤 <b>${v[0]}</b>\n   • Reservas: ${v[1]} (${v[2]} PAX)\n   • Vendido: $${v[3]} | Abonado: $${v[4]}\n   • 🔴 <b>Pendiente: $${v[5]}</b>`
            )
            .join("\n\n");

          await sendTelegramMessage(
            chatId,
            `💰 <b>ESTADO DE CUENTAS — VENDEDORES</b>\n\n${summaryText}`
          );
        }
      }

      return NextResponse.json({ ok: true });
    }

    // Manejar mensajes de texto
    const message = update.message;
    if (!message || !message.text) {
      return NextResponse.json({ ok: true });
    }

    const chatId = message.chat.id;
    const text = message.text.trim();
    const senderName =
      [message.from?.first_name, message.from?.last_name].filter(Boolean).join(" ") ||
      "Staff JAC";

    // Verificar autorización
    if (ALLOWED_CHATS.length > 0 && !ALLOWED_CHATS.includes(String(chatId))) {
      await sendTelegramMessage(
        chatId,
        `⛔ Acceso no autorizado. Tu ID de chat es: <code>${chatId}</code>`
      );
      return NextResponse.json({ ok: true });
    }

    const mainKeyboard = {
      inline_keyboard: [
        [
          { text: "🚐 Pick-Up Hoy", callback_data: "pickup_hoy" },
          { text: "📅 Pick-Up Mañana", callback_data: "pickup_manana" },
        ],
        [{ text: "💰 Cuentas Vendedores", callback_data: "cuentas_vendedores" }],
      ],
    };

    if (text === "/start" || text === "/menu") {
      await sendTelegramMessage(
        chatId,
        `🌴 <b>JAC Tour and Transfers — Sistema Operativo</b>\n\nReenvíame o pega cualquier mensaje de reserva de WhatsApp y lo guardaré automáticamente en Google Sheets.\n\nTambién puedes consultar tus reportes aquí:`,
        mainKeyboard
      );
      return NextResponse.json({ ok: true });
    }

    // Procesar texto como reserva de WhatsApp con Gemini
    await sendTelegramMessage(chatId, "⏳ Analizando reserva y sincronizando con Google Sheets...");

    const parsed = await parseWhatsAppReservation(text, senderName);

    if (!parsed.isReservation) {
      await sendTelegramMessage(
        chatId,
        "⚠️ No detecté datos de una reserva turística en ese mensaje. Pega el texto de la reserva o usa el menú:",
        mainKeyboard
      );
      return NextResponse.json({ ok: true });
    }

    const saved = await saveReservationToSheets(parsed);

    const confirmMsg =
      `✅ <b>RESERVA REGISTRADA (${saved.reservaId})</b>\n\n` +
      `📅 <b>Fecha:</b> ${parsed.fechaTour} — ⏰ ${parsed.horaPickUp}\n` +
      `🏝 <b>Excursión:</b> ${parsed.excursion}\n` +
      `👤 <b>Cliente:</b> ${parsed.nombreCliente} (${saved.totalPax} PAX)\n` +
      `🏨 <b>Pick-Up:</b> ${parsed.hotelPuntoEncuentro} (Hab: ${parsed.habitacion})\n` +
      `💼 <b>Vendedor:</b> ${parsed.vendedor}\n` +
      `💵 <b>Total:</b> $${parsed.precioTotal} ${parsed.moneda} | <b>Pendiente:</b> $${parsed.saldoPendiente}\n\n` +
      `📑 Guardado en <code>Reservas_Master</code>, <code>${saved.pickUpTab}</code> y <code>Cuentas_Vendedores</code>.`;

    await sendTelegramMessage(chatId, confirmMsg, mainKeyboard);
    return NextResponse.json({ ok: true });
  } catch (error) {
    console.error("Error en webhook Telegram:", error);
    return NextResponse.json({ ok: false, error: String(error) }, { status: 500 });
  }
}