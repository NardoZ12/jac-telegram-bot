import { google } from "googleapis";
import { ParsedReservation } from "./whatsappParser";

const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID!;

function getSheetsClient() {
  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n"),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  return google.sheets({ version: "v4", auth });
}

const MASTER_HEADERS = [
  "ID Reserva",
  "Fecha Registro",
  "Fecha Tour",
  "Hora Pick-Up",
  "Cliente",
  "Teléfono",
  "Hotel / Punto",
  "Habitación",
  "Excursión",
  "Adultos",
  "Niños",
  "Infantes",
  "Total PAX",
  "Vendedor",
  "Total",
  "Abonado",
  "Saldo Pendiente",
  "Moneda",
  "Método Pago",
  "Estado",
  "Notas",
];

const PICKUP_HEADERS = [
  "Hora Pick-Up",
  "Hotel / Punto",
  "Habitación",
  "Cliente",
  "Teléfono",
  "Excursión",
  "PAX (A/N/I)",
  "Cobrar al Abordar",
  "Vendedor",
  "Notas",
];

const VENDEDORES_HEADERS = [
  "Vendedor",
  "Total Reservas",
  "Total PAX",
  "Monto Vendido",
  "Total Abonado",
  "Saldo Pendiente",
  "Última Actualización",
];

async function ensureSheetExists(
  sheets: ReturnType<typeof getSheetsClient>,
  title: string,
  headers: string[]
) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets?.some(
    (s) => s.properties?.title === title
  );

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title } } }],
      },
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${title}'!A1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [headers] },
    });
  }
}

export async function saveReservationToSheets(data: ParsedReservation) {
  const sheets = getSheetsClient();
  const now = new Date().toLocaleString("es-DO", {
    timeZone: "America/Santo_Domingo",
  });
  const reservaId = `JAC-${Date.now().toString().slice(-6)}`;
  const totalPax = data.adultos + data.ninos + data.infantes;
  const pickUpTab = `PickUp_${data.fechaTour}`;

  // 1. Asegurar pestañas necesarias
  await ensureSheetExists(sheets, "Reservas_Master", MASTER_HEADERS);
  await ensureSheetExists(sheets, pickUpTab, PICKUP_HEADERS);
  await ensureSheetExists(sheets, "Cuentas_Vendedores", VENDEDORES_HEADERS);

  // 2. Insertar en Reservas_Master
  const masterRow = [
    reservaId,
    now,
    data.fechaTour,
    data.horaPickUp,
    data.nombreCliente,
    data.telefonoCliente || "N/A",
    data.hotelPuntoEncuentro,
    data.habitacion || "N/A",
    data.excursion,
    data.adultos,
    data.ninos,
    data.infantes,
    totalPax,
    data.vendedor,
    data.precioTotal,
    data.abonado,
    data.saldoPendiente,
    data.moneda,
    data.metodoPago,
    data.saldoPendiente > 0 ? "PENDIENTE PAGO" : "PAGADO",
    data.notas || "",
  ];

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Reservas_Master'!A:U",
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [masterRow] },
  });

  // 3. Insertar en la lista diaria PickUp_YYYY-MM-DD
  const pickUpRow = [
    data.horaPickUp,
    data.hotelPuntoEncuentro,
    data.habitacion || "N/A",
    data.nombreCliente,
    data.telefonoCliente || "N/A",
    data.excursion,
    `${data.adultos}A / ${data.ninos}N / ${data.infantes}I`,
    data.saldoPendiente > 0
      ? `${data.saldoPendiente} ${data.moneda}`
      : "PAGADO",
    data.vendedor,
    data.notas || "",
  ];

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${pickUpTab}'!A:J`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [pickUpRow] },
  });

  // 4. Recalcular resumen en Cuentas_Vendedores
  await recalculateVendorAccounts(sheets, now);

  return { reservaId, pickUpTab, totalPax };
}

async function recalculateVendorAccounts(
  sheets: ReturnType<typeof getSheetsClient>,
  timestamp: string
) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Reservas_Master'!A2:U",
  });

  const rows = res.data.values || [];
  const summary: Record<
    string,
    { reservas: number; pax: number; total: number; abonado: number; pendiente: number }
  > = {};

  for (const row of rows) {
    const vendedor = (row[13] || "Sin asignar").trim();
    const pax = Number(row[12]) || 0;
    const total = Number(row[14]) || 0;
    const abonado = Number(row[15]) || 0;
    const pendiente = Number(row[16]) || 0;

    if (!summary[vendedor]) {
      summary[vendedor] = { reservas: 0, pax: 0, total: 0, abonado: 0, pendiente: 0 };
    }
    summary[vendedor].reservas += 1;
    summary[vendedor].pax += pax;
    summary[vendedor].total += total;
    summary[vendedor].abonado += abonado;
    summary[vendedor].pendiente += pendiente;
  }

  const vendorRows = Object.entries(summary).map(([vendedor, s]) => [
    vendedor,
    s.reservas,
    s.pax,
    s.total,
    s.abonado,
    s.pendiente,
    timestamp,
  ]);

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Cuentas_Vendedores'!A2:G",
  });

  if (vendorRows.length > 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Cuentas_Vendedores'!A2",
      valueInputOption: "USER_ENTERED",
      requestBody: { values: vendorRows },
    });
  }
}

export async function getPickUpByDate(dateStr: string) {
  const sheets = getSheetsClient();
  const tabName = `PickUp_${dateStr}`;

  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${tabName}'!A2:J`,
    });
    return res.data.values || [];
  } catch {
    return [];
  }
}

export async function getVendorSummary() {
  const sheets = getSheetsClient();
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Cuentas_Vendedores'!A2:G",
    });
    return res.data.values || [];
  } catch {
    return [];
  }
}