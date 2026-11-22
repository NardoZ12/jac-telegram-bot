import { google } from "googleapis";
import { ParsedReservation } from "./whatsappParser";

const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID!;
const LOGO_URL =
  process.env.COMPANY_LOGO_URL ||
  "https://lh3.googleusercontent.com/d/1_gKtQHqIoFphnTAZHd9ajw8plAmzZZC3";
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
  "SERVICIO",
  "HORA",
  "HOTEL",
  "OPERADOR",
  "REPRESENTANTE",
  "NOMBRE",
  "PAX",
  "HABITACION",
  "TRANSPORT",
  "DEBEN",
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

const VENDOR_ACCOUNT_HEADERS = [
  "ID Reserva",
  "Fecha Registro",
  "Fecha Tour",
  "Cliente",
  "Excursión",
  "Total PAX",
  "Total",
  "Abonado",
  "Saldo Pendiente",
  "Moneda",
  "Estado",
];

function formatDateDDMMYYYY(isoDate: string): string {
  const parts = isoDate.split("-");
  if (parts.length !== 3) return isoDate;
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

async function ensureMasterOrVendorSheet(
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

    if (headers.length > 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${title}'!A1`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [headers] },
      });
    }
  }
}

async function ensureStyledPickUpSheet(
  sheets: ReturnType<typeof getSheetsClient>,
  title: string,
  fechaTourISO: string
) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets?.some(
    (s) => s.properties?.title === title
  );

  if (exists) return;

  // 1. Crear pestaña nueva y obtener su sheetId
  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{ addSheet: { properties: { title } } }],
    },
  });

  const sheetId = addRes.data.replies?.[0]?.addSheet?.properties?.sheetId;
  const fechaFormateada = formatDateDDMMYYYY(fechaTourISO);

  // 2. Escribir cabecera corporativa + encabezados en fila 7
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${title}'!A1:J7`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [
        ["", "", "", "", `=IMAGE("${LOGO_URL}", 1)`, "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "EXPEDIA AND VIATOR SUPLIER", "", "", "", "", "", ""],
        ["", "", `FECHA:${fechaFormateada}`, "C/PASTOR MAXIMO NO. 8", "", "", "", "", "", ""],
        ["", "", "", "CEIBA DEL SALADO. BAVARO PUNTA CANA", "", "", "", "", "", ""],
        ["", "", "", "", "RNC:1-31-477-453", "", "", "", "", ""],
        PICKUP_HEADERS,
      ],
    },
  });

  if (sheetId === undefined) return;

  // 3. Aplicar estilo visual idéntico a ORDEN EDITADA (combinaciones, fuentes, grises, bordes y anchos)
  const colWidths = [160, 85, 290, 140, 175, 220, 110, 125, 115, 135];
  const dimensionRequests = colWidths.map((px, idx) => ({
    updateDimensionProperties: {
      range: {
        sheetId,
        dimension: "COLUMNS",
        startIndex: idx,
        endIndex: idx + 1,
      },
      properties: { pixelSize: px },
      fields: "pixelSize",
    },
  }));

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        // Combinar espacio del logo (E1:G2) y líneas de dirección (D3:F3, D4:F4, D5:F5, E6:F6)
        {
          mergeCells: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 2, startColumnIndex: 4, endColumnIndex: 7 },
            mergeType: "MERGE_ALL",
          },
        },
        {
          mergeCells: {
            range: { sheetId, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 3, endColumnIndex: 6 },
            mergeType: "MERGE_ALL",
          },
        },
        {
          mergeCells: {
            range: { sheetId, startRowIndex: 3, endRowIndex: 4, startColumnIndex: 3, endColumnIndex: 6 },
            mergeType: "MERGE_ALL",
          },
        },
        {
          mergeCells: {
            range: { sheetId, startRowIndex: 4, endRowIndex: 5, startColumnIndex: 3, endColumnIndex: 6 },
            mergeType: "MERGE_ALL",
          },
        },
        {
          mergeCells: {
            range: { sheetId, startRowIndex: 5, endRowIndex: 6, startColumnIndex: 4, endColumnIndex: 6 },
            mergeType: "MERGE_ALL",
          },
        },
        // Altura de filas de cabecera y encabezados
        {
          updateDimensionProperties: {
            range: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: 2 },
            properties: { pixelSize: 42 },
            fields: "pixelSize",
          },
        },
        {
          updateDimensionProperties: {
            range: { sheetId, dimension: "ROWS", startIndex: 6, endIndex: 7 },
            properties: { pixelSize: 34 },
            fields: "pixelSize",
          },
        },
        // Estilo texto cabecera corporativa (Filas 3-6)
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 2, endRowIndex: 6, startColumnIndex: 2, endColumnIndex: 7 },
            cell: {
              userEnteredFormat: {
                horizontalAlignment: "CENTER",
                verticalAlignment: "MIDDLE",
                textFormat: { fontFamily: "Times New Roman", fontSize: 11, bold: true },
              },
            },
            fields: "userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat)",
          },
        },
        // Fila 7: Encabezados con fondo gris (#808080), negrita y bordes
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 6, endRowIndex: 7, startColumnIndex: 0, endColumnIndex: 10 },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.5, green: 0.5, blue: 0.5 },
                horizontalAlignment: "CENTER",
                verticalAlignment: "MIDDLE",
                textFormat: {
                  fontFamily: "Times New Roman",
                  fontSize: 12,
                  bold: true,
                  foregroundColor: { red: 0, green: 0, blue: 0 },
                },
                borders: {
                  top: { style: "SOLID" },
                  bottom: { style: "SOLID" },
                  left: { style: "SOLID" },
                  right: { style: "SOLID" },
                },
              },
            },
            fields: "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat,borders)",
          },
        },
        // Formato base de filas de datos (Filas 8 a 100) con bordes finos y negrita
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 7, endRowIndex: 100, startColumnIndex: 0, endColumnIndex: 10 },
            cell: {
              userEnteredFormat: {
                horizontalAlignment: "CENTER",
                verticalAlignment: "MIDDLE",
                textFormat: { fontFamily: "Arial", fontSize: 11, bold: true },
                borders: {
                  top: { style: "SOLID" },
                  bottom: { style: "SOLID" },
                  left: { style: "SOLID" },
                  right: { style: "SOLID" },
                },
              },
            },
            fields: "userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat,borders)",
          },
        },
        ...dimensionRequests,
      ],
    },
  });
}

export async function saveReservationToSheets(data: ParsedReservation) {
  const sheets = getSheetsClient();
  const now = new Date().toLocaleString("es-DO", {
    timeZone: "America/Santo_Domingo",
  });
  const reservaId = `JAC-${Date.now().toString().slice(-6)}`;
  const totalPax = data.adultos + data.ninos + data.infantes;
  const fechaTour = isValidISODate(data.fechaTour)
    ? data.fechaTour
    : "Por confirmar";
  const pickUpTab =
    fechaTour === "Por confirmar" ? null : `PickUp_${fechaTour}`;

  const estadoPago =
    data.saldoPendiente > 0
      ? "PENDIENTE PAGO"
      : data.precioTotal > 0
        ? "PAGADO"
        : "POR CONFIRMAR";

  // 1. Asegurar pestañas necesarias (PickUp con estilo corporativo del cliente)
  await ensureMasterOrVendorSheet(sheets, "Reservas_Master", MASTER_HEADERS);
  await ensureMasterOrVendorSheet(sheets, "Cuentas_Vendedores", VENDEDORES_HEADERS);
  if (pickUpTab) {
    await ensureStyledPickUpSheet(sheets, pickUpTab, fechaTour);
  }

  // 2. Insertar en Reservas_Master
  const masterRow = [
    reservaId,
    now,
    fechaTour,
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
    estadoPago,
    data.notas || "",
  ];

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Reservas_Master'!A:U",
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [masterRow] },
  });

  // 3. Formatear PAX y DEBEN al estilo exacto del cliente (ej. "3+3" y "315US")
  const paxFormatted =
    data.ninos > 0 || data.infantes > 0
      ? `${data.adultos}+${data.ninos + data.infantes}`
      : `${data.adultos}`;

  const debenFormatted =
    data.saldoPendiente > 0
      ? `${data.saldoPendiente}${data.moneda === "USD" ? "US" : data.moneda}`
      : "";

  // Insertar en la lista diaria PickUp_YYYY-MM-DD (a partir de fila 8)
  const pickUpRow = [
    data.excursion.toUpperCase(),
    data.horaPickUp.replace(/\s*(AM|PM)/i, ""),
    data.hotelPuntoEncuentro.toUpperCase(),
    "JAC",
    data.vendedor.toUpperCase(),
    data.nombreCliente.toUpperCase(),
    paxFormatted,
    data.habitacion || "",
    "QT",
    debenFormatted,
  ];

  if (pickUpTab) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${pickUpTab}'!A8:J`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [pickUpRow] },
    });
  }

  // 4. Recalcular resumen general y hojas de cobro individuales.
  const vendorTab = await recalculateVendorAccounts(sheets, now, data.vendedor);

  return { reservaId, pickUpTab, vendorTab, totalPax };
}

async function recalculateVendorAccounts(
  sheets: ReturnType<typeof getSheetsClient>,
  timestamp: string,
  currentVendor: string
) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Reservas_Master'!A2:U",
  });

  const rows = res.data.values || [];
  const vendors = new Map<
    string,
    {
      reservas: number;
      pax: number;
      total: number;
      abonado: number;
      pendiente: number;
      details: string[][];
    }
  >();

  for (const row of rows) {
    const vendedor = (row[13] || "Sin asignar").trim();
    const pax = Number(row[12]) || 0;
    const total = Number(row[14]) || 0;
    const abonado = Number(row[15]) || 0;
    const pendiente = Number(row[16]) || 0;

    if (!vendors.has(vendedor)) {
      vendors.set(vendedor, {
        reservas: 0,
        pax: 0,
        total: 0,
        abonado: 0,
        pendiente: 0,
        details: [],
      });
    }
    const vendor = vendors.get(vendedor)!;
    vendor.reservas += 1;
    vendor.pax += pax;
    vendor.total += total;
    vendor.abonado += abonado;
    vendor.pendiente += pendiente;
    vendor.details.push([
      row[0] || "",
      row[1] || "",
      row[2] || "Por confirmar",
      row[4] || "",
      row[8] || "",
      String(pax),
      String(total),
      String(abonado),
      String(pendiente),
      row[17] || "USD",
      row[19] || "",
    ]);
  }

  const vendorRows = [...vendors.entries()].map(([vendedor, s]) => [
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

  let currentVendorTab = "";
  for (const [vendedor, account] of vendors) {
    const title = getVendorAccountSheetTitle(vendedor);
    await ensureMasterOrVendorSheet(sheets, title, []);
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SPREADSHEET_ID,
      range: `${quoteSheetRange(title, "A1:K")}`,
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: quoteSheetRange(title, "A1"),
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [
          ["CUENTA DE COBRO DEL VENDEDOR"],
          ["Vendedor", vendedor],
          ["Saldo total", account.pendiente],
          ["Total vendido", account.total],
          ["Total abonado", account.abonado],
          ["Última actualización", timestamp],
          VENDOR_ACCOUNT_HEADERS,
          ...account.details,
        ],
      },
    });
    if (vendedor === currentVendor.trim()) {
      currentVendorTab = title;
    }
  }

  return currentVendorTab || getVendorAccountSheetTitle(currentVendor);
}

export async function getPickUpByDate(dateStr: string) {
  const sheets = getSheetsClient();
  const tabName = `PickUp_${dateStr}`;

  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${tabName}'!A8:J`,
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

function getVendorAccountSheetTitle(vendor: string): string {
  const base = vendor
    .replace(/[\\/*?:[\]]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 72) || "Sin_asignar";
  let hash = 2166136261;
  for (const char of vendor) {
    hash ^= char.charCodeAt(0);
    hash = Math.imul(hash, 16777619);
  }

  return `Cobro_${base}_${(hash >>> 0).toString(16).slice(-6)}`;
}

function quoteSheetRange(title: string, range: string): string {
  return `'${title.replace(/'/g, "''")}'!${range}`;
}