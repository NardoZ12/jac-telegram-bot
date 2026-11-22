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

const VENDOR_DETAIL_HEADERS = [
  "FECHA TOUR",
  "ID RESERVA",
  "CLIENTE",
  "SERVICIO",
  "PAX",
  "MONTO / PAX",
  "TOTAL RESERVA",
  "ABONADO",
  "PENDIENTE",
  "ESTADO PAGO",
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

function formatDateDDMMYYYY(isoDate: string): string {
  if (!isoDate || isoDate.toLowerCase() === "por confirmar") {
    return "POR CONFIRMAR";
  }
  const parts = isoDate.split("-");
  if (parts.length !== 3) return isoDate;
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

function sanitizeSheetTabName(raw: string): string {
  return raw
    .replace(/[\\/?*[\]:]/g, "")
    .trim()
    .slice(0, 28);
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

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${title}'!A1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [headers] },
    });
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

  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{ addSheet: { properties: { title } } }],
    },
  });

  const sheetId = addRes.data.replies?.[0]?.addSheet?.properties?.sheetId;
  const fechaFormateada = formatDateDDMMYYYY(fechaTourISO);

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

async function ensureStyledVendorBillingSheet(
  sheets: ReturnType<typeof getSheetsClient>,
  vendedorName: string
): Promise<string> {
  const cleanName = sanitizeSheetTabName(vendedorName.toUpperCase() || "SIN ASIGNAR");
  const title = `Cobros_${cleanName}`;

  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = meta.data.sheets?.some(
    (s) => s.properties?.title === title
  );

  if (exists) return title;

  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{ addSheet: { properties: { title } } }],
    },
  });

  const sheetId = addRes.data.replies?.[0]?.addSheet?.properties?.sheetId;

  // Cabecera corporativa con Logo Quality Transport + Panel de Resumen Financiero del Vendedor
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${title}'!A1:J7`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [
        ["", "", "", "", `=IMAGE("${LOGO_URL}", 1)`, "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "EXPEDIA AND VIATOR SUPLIER", "", "", "", "", "TOTAL VENDIDO:", "=SUM(G8:G200)"],
        ["", "", `VENDEDOR: ${cleanName}`, "C/PASTOR MAXIMO NO. 8", "", "", "", "", "ABONADO RESERVAS:", "=SUM(H8:H200)"],
        ["", "", "", "CEIBA DEL SALADO. BAVARO PUNTA CANA", "", "", "", "", "ABONO CUENTA GENERAL:", 0],
        ["", "", "", "", "RNC:1-31-477-453", "", "", "", "SALDO PENDIENTE TOTAL:", "=J3-J4-J5"],
        VENDOR_DETAIL_HEADERS,
      ],
    },
  });

  if (sheetId === undefined) return title;

  const colWidths = [130, 120, 220, 165, 80, 130, 145, 130, 135, 150];
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
        // Estilo cuadro de totales del vendedor (I3:J6)
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 2, endRowIndex: 6, startColumnIndex: 8, endColumnIndex: 10 },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.94, green: 0.96, blue: 0.98 },
                horizontalAlignment: "RIGHT",
                verticalAlignment: "MIDDLE",
                textFormat: { fontFamily: "Arial", fontSize: 10, bold: true },
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
        // Formato moneda en J3:J6 y F8:I200
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 2, endRowIndex: 6, startColumnIndex: 9, endColumnIndex: 10 },
            cell: {
              userEnteredFormat: {
                numberFormat: { type: "CURRENCY", pattern: "$#,##0.00" },
              },
            },
            fields: "userEnteredFormat.numberFormat",
          },
        },
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 7, endRowIndex: 200, startColumnIndex: 5, endColumnIndex: 9 },
            cell: {
              userEnteredFormat: {
                numberFormat: { type: "CURRENCY", pattern: "$#,##0.00" },
              },
            },
            fields: "userEnteredFormat.numberFormat",
          },
        },
        // Fila 7: Encabezados grises
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
                  fontSize: 11,
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
        // Bordes y centrado de filas de datos (8 a 200)
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 7, endRowIndex: 200, startColumnIndex: 0, endColumnIndex: 10 },
            cell: {
              userEnteredFormat: {
                horizontalAlignment: "CENTER",
                verticalAlignment: "MIDDLE",
                textFormat: { fontFamily: "Arial", fontSize: 10, bold: true },
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
        // Lista desplegable en columna J (ESTADO PAGO) para marcar PENDIENTE / ABONADO / PAGADO
        {
          setDataValidation: {
            range: { sheetId, startRowIndex: 7, endRowIndex: 200, startColumnIndex: 9, endColumnIndex: 10 },
            rule: {
              condition: {
                type: "ONE_OF_LIST",
                values: [
                  { userEnteredValue: "PENDIENTE" },
                  { userEnteredValue: "ABONADO" },
                  { userEnteredValue: "PAGADO" },
                ],
              },
              showCustomUi: true,
              strict: false,
            },
          },
        },
        ...dimensionRequests,
      ],
    },
  });

  return title;
}

export async function saveReservationToSheets(data: ParsedReservation) {
  const sheets = getSheetsClient();
  const now = new Date().toLocaleString("es-DO", {
    timeZone: "America/Santo_Domingo",
  });
  const reservaId = `JAC-${Date.now().toString().slice(-6)}`;
  const totalPax = data.adultos + data.ninos + data.infantes;

  // Si no trae fecha válida YYYY-MM-DD, enviarlo a "Por confirmar" en vez de la fecha de hoy
  const hasValidDate =
    data.fechaTour && /^\d{4}-\d{2}-\d{2}$/.test(data.fechaTour.trim());
  const fechaTourFinal = hasValidDate ? data.fechaTour.trim() : "Por confirmar";
  const pickUpTab = `PickUp_${fechaTourFinal}`;

  const estadoPago =
    data.saldoPendiente > 0
      ? data.abonado > 0
        ? "ABONADO"
        : "PENDIENTE"
      : data.precioTotal > 0
        ? "PAGADO"
        : "PENDIENTE";

  // 1. Asegurar pestañas necesarias (incluyendo la hoja individual del vendedor)
  await ensureMasterOrVendorSheet(sheets, "Reservas_Master", MASTER_HEADERS);
  await ensureStyledPickUpSheet(sheets, pickUpTab, fechaTourFinal);
  await ensureMasterOrVendorSheet(sheets, "Cuentas_Vendedores", VENDEDORES_HEADERS);
  const vendorBillingTab = await ensureStyledVendorBillingSheet(sheets, data.vendedor);

  // 2. Insertar en Reservas_Master
  const masterRow = [
    reservaId,
    now,
    fechaTourFinal,
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

  // 3. Insertar en la lista diaria PickUp_<Fecha o Por confirmar>
  const paxFormatted =
    data.ninos > 0 || data.infantes > 0
      ? `${data.adultos}+${data.ninos + data.infantes}`
      : `${data.adultos}`;

  const debenFormatted =
    data.saldoPendiente > 0
      ? `${data.saldoPendiente}${data.moneda === "USD" ? "US" : data.moneda}`
      : "";

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

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${pickUpTab}'!A8:J`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [pickUpRow] },
  });

  // 4. Insertar en la pestaña individual de cobros del vendedor (Cobros_<VENDEDOR>)
  const existingVendorRows = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${vendorBillingTab}'!A8:A`,
  });
  const nextRowNumber = 8 + (existingVendorRows.data.values?.length || 0);
  const precioPorPersona =
    totalPax > 0 && data.precioTotal > 0
      ? Number((data.precioTotal / totalPax).toFixed(2))
      : 0;

  const vendorBillingRow = [
    formatDateDDMMYYYY(fechaTourFinal),
    reservaId,
    data.nombreCliente.toUpperCase(),
    data.excursion.toUpperCase(),
    totalPax,
    precioPorPersona,
    `=E${nextRowNumber}*F${nextRowNumber}`,
    data.abonado || 0,
    `=G${nextRowNumber}-H${nextRowNumber}`,
    estadoPago,
  ];

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${vendorBillingTab}'!A${nextRowNumber}:J${nextRowNumber}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [vendorBillingRow] },
  });

  // 5. Recalcular resumen global en Cuentas_Vendedores
  await recalculateVendorAccounts(sheets, now);

  return { reservaId, pickUpTab, vendorBillingTab, totalPax, fechaTourFinal };
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
