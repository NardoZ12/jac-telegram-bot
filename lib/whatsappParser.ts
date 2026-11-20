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
      description: "Fecha del tour o traslado en formato YYYY-MM-DD.",
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
  const todayStr = new Date().toISOString().split("T")[0];

  const prompt = `Eres el asistente operativo de "JAC Tour and Transfers" en Punta Cana, República Dominicana.
Analiza el siguiente mensaje copiado de WhatsApp y extrae los datos estructurados de la reserva turística.

Reglas:
- Fecha actual de referencia: ${todayStr}.
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
      return JSON.parse(jsonText) as ParsedReservation;
    } catch (err) {
      lastError = err;
      console.warn(`Modelo ${modelName} ocupado, pasando al siguiente...`);
    }
  }

  throw lastError;
}