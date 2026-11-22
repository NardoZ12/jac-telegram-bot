import { GoogleGenAI, Type, Schema } from "@google/genai";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY! });

export interface ParsedReservation {
  isReservation: boolean;
  fechaTour: string;
  horaPickUp: string;
  nombreCliente: string;
  telefonoCliente: string;
  hotelPuntoEncuentro: string;
  habitacion: string;
  excursion: string;
  adultos: number;
  ninos: number;
  infantes: number;
  vendedor: string;
  precioTotal: number;
  abonado: number;
  saldoPendiente: number;
  moneda: string;
  metodoPago: string;
  notas: string;
}

const reservationSchema: Schema = {
  type: Type.OBJECT,
  properties: {
    isReservation: {
      type: Type.BOOLEAN,
      description: "True si el mensaje contiene datos de una reserva o traslado turístico; false si es saludo o charla casual.",
    },
    fechaTour: {
      type: Type.STRING,
      description:
        "Fecha del tour o traslado en formato YYYY-MM-DD. Si el mensaje no contiene un día, fecha o referencia temporal inequívoca, devuelve exactamente 'Por confirmar'.",
    },
    horaPickUp: {
      type: Type.STRING,
      description: "Hora de recogida (ej. '07:30 AM', '14:00'). Si no se indica, 'Por confirmar'.",
    },
    nombreCliente: {
      type: Type.STRING,
      description: "Nombre completo del pasajero principal.",
    },
    telefonoCliente: {
      type: Type.STRING,
      description: "Teléfono o WhatsApp del cliente si aparece; si no, cadena vacía.",
    },
    hotelPuntoEncuentro: {
      type: Type.STRING,
      description: "Nombre del hotel, resort, Airbnb, aeropuerto o punto de encuentro en Punta Cana / Bávaro.",
    },
    habitacion: {
      type: Type.STRING,
      description: "Número de habitación o villa si aparece; si no, 'N/A'.",
    },
    excursion: {
      type: Type.STRING,
      description: "Nombre de la excursión o servicio (ej. Isla Saona, Buggies Macao, Catalina, Transfer PUJ).",
    },
    adultos: {
      type: Type.INTEGER,
      description: "Cantidad de adultos (mínimo 1 si no se especifica).",
    },
    ninos: {
      type: Type.INTEGER,
      description: "Cantidad de niños (0 si no hay).",
    },
    infantes: {
      type: Type.INTEGER,
      description: "Cantidad de infantes/bebés (0 si no hay).",
    },
    vendedor: {
      type: Type.STRING,
      description: "Nombre del vendedor, agencia o promotor. Si no se indica, 'Sin asignar'.",
    },
    precioTotal: {
      type: Type.NUMBER,
      description: "Monto total de la reserva (0 si no se menciona).",
    },
    abonado: {
      type: Type.NUMBER,
      description: "Monto ya pagado o depositado por el cliente (0 si no pagó nada aún).",
    },
    saldoPendiente: {
      type: Type.NUMBER,
      description: "Monto pendiente por cobrar al abordar o liquidar por el vendedor (precioTotal - abonado).",
    },
    moneda: {
      type: Type.STRING,
      description: "Moneda del cobro: 'USD', 'DOP' o 'EUR'. Por defecto 'USD'.",
    },
    metodoPago: {
      type: Type.STRING,
      description: "Ej: 'Cash al abordar', 'Transferencia', 'Link de pago', 'Por cobrar vendedor', 'Pagado'.",
    },
    notas: {
      type: Type.STRING,
      description: "Cualquier observación adicional (vuelo, idioma, silla de bebé, restricciones).",
    },
  },
  required: [
    "isReservation",
    "fechaTour",
    "horaPickUp",
    "nombreCliente",
    "hotelPuntoEncuentro",
    "excursion",
    "adultos",
    "ninos",
    "infantes",
    "vendedor",
    "precioTotal",
    "abonado",
    "saldoPendiente",
    "moneda",
    "metodoPago",
  ],
};

export async function parseWhatsAppReservation(
  rawText: string,
  senderFallback: string
): Promise<ParsedReservation> {
  const todayStr = getDominicanDate();

  const prompt = `Eres el asistente operativo de "JAC Tour and Transfers" en Punta Cana, República Dominicana.
Analiza el siguiente mensaje copiado de WhatsApp y extrae los datos estructurados de la reserva turística.

Reglas:
- Fecha actual de referencia: ${todayStr}.
- Devuelve fechaTour en formato YYYY-MM-DD únicamente cuando el mensaje indica explícitamente un día o una referencia temporal ("hoy", "mañana", "pasado mañana", día de la semana, fecha con números o mes).
- Si el mensaje NO menciona cuándo se realizará el tour o traslado, fechaTour debe ser exactamente "Por confirmar". Nunca uses la fecha actual como valor predeterminado.
- Si el mensaje no menciona quién es el vendedor, usa: "${senderFallback}".
- Normaliza nombres de hoteles conocidos de Punta Cana / Bávaro / Cap Cana / Uvero Alto si vienen abreviados.
- Si el texto dice "Pendiente: 50" o "Cobra chofer 50", colócalo en saldoPendiente y reflejalo en metodoPago.
- Si el mensaje NO es una reserva, devuelve isReservation: false.

Mensaje de WhatsApp:
"""
${rawText}
"""`;

const candidates = [
    "gemini-3.6-flash",
    "gemini-3.1-flash-lite",
    "gemini-3.7-flash",
    "gemini-3.5-flash",
    "gemini-3.8-flash",
  ];

  let lastError: unknown;

  for (const modelName of candidates) {
    try {
      const response = await ai.models.generateContent({
        model: modelName,
        contents: prompt,
        config: {
          responseMimeType: "application/json",
          responseSchema: reservationSchema,
          temperature: 0.1,
        },
      });

      const jsonText = response.text || "{}";
      const parsed = JSON.parse(jsonText) as ParsedReservation;

      if (!hasDateReference(rawText)) {
        parsed.fechaTour = "Por confirmar";
      } else if (!isValidISODate(parsed.fechaTour)) {
        parsed.fechaTour = "Por confirmar";
      }

      return parsed;
    } catch (err) {
      lastError = err;
      console.warn(`Modelo ${modelName} ocupado, pasando al siguiente...`);
    }
  }

  throw lastError;
}

function getDominicanDate(): string {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Santo_Domingo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const part = (type: Intl.DateTimeFormatPartTypes) =>
    parts.find((item) => item.type === type)?.value || "";

  return `${part("year")}-${part("month")}-${part("day")}`;
}

function hasDateReference(text: string): boolean {
  const normalized = text
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();

  return [
    /\b(hoy|manana|pasado manana|esta noche|siguiente semana)\b/,
    /\b(lunes|martes|miercoles|jueves|viernes|sabado|domingo)\b/,
    /\b(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\b/,
    /\b\d{1,2}\s*[/.-]\s*\d{1,2}(?:\s*[/.-]\s*\d{2,4})?\b/,
    /\b\d{4}-\d{2}-\d{2}\b/,
    /\bel\s+\d{1,2}\b/,
  ].some((pattern) => pattern.test(normalized));
}

function isValidISODate(value: string): boolean {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) return false;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));

  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
}