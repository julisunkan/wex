import * as XLSX from "xlsx";
import { categorize, categorizeByName, type Transaction } from "./categorizer";
import { parsePastedText } from "./csv-parser";

export type ColumnMap = {
  date: number;
  description: number;
  amount: number | null;
  type: number | null;
  debit: number | null;
  credit: number | null;
  balance: number | null;
  category: number | null;
};

function detectColumnsFromHeaders(headers: string[]): ColumnMap | null {
  const h = headers.map((x) => String(x || "").toLowerCase().trim());
  const find = (...terms: string[]) =>
    h.findIndex((col) => terms.some((t) => col === t || col.includes(t)));

  const date        = find("date", "tran date", "value date", "trans date");
  const description = find("description", "narration", "details", "particulars", "memo", "remark", "transaction");

  if (date === -1 || description === -1) return null;

  const category = find("category", "cat");
  const debit    = find("debit");
  const credit   = find("credit");

  if (debit !== -1 && credit !== -1) {
    const balance = find("balance", "running balance", "bal");
    return {
      date, description,
      amount: null, type: null,
      debit, credit,
      balance: balance === -1 ? null : balance,
      category: category === -1 ? null : category,
    };
  }

  const amount = find("amount", "value");
  const type   = find("type", "dr/cr", "debit/credit", "transaction type", "cr/dr");

  if (amount === -1) return null;

  return {
    date, description,
    amount, type: type === -1 ? null : type,
    debit: null, credit: null, balance: null,
    category: category === -1 ? null : category,
  };
}

function rowsToTransactions(headers: string[], rows: string[][], columnMap: ColumnMap): Transaction[] {
  const transactions: Transaction[] = [];
  const isSplit = columnMap.debit !== null && columnMap.credit !== null;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rawDate = columnMap.date < row.length ? row[columnMap.date] : "";
    const rawDesc = columnMap.description < row.length ? row[columnMap.description] : "";
    if (!rawDesc) continue;

    const date = rawDate ? String(rawDate) : "";
    const description = String(rawDesc);

    let amount = 0;
    let type: "credit" | "debit" = "debit";

    if (isSplit) {
      const rawDebit  = columnMap.debit  !== null && columnMap.debit  < row.length ? row[columnMap.debit]  : "";
      const rawCredit = columnMap.credit !== null && columnMap.credit < row.length ? row[columnMap.credit] : "";
      const debitVal  = rawDebit  !== "" ? Math.abs(parseFloat(rawDebit.replace(/[^0-9.-]/g, "")) || 0)  : 0;
      const creditVal = rawCredit !== "" ? Math.abs(parseFloat(rawCredit.replace(/[^0-9.-]/g, "")) || 0) : 0;
      if (creditVal > 0) { amount = creditVal; type = "credit"; }
      else if (debitVal > 0) { amount = debitVal; type = "debit"; }
      else continue;
    } else {
      const rawAmount = columnMap.amount !== null && columnMap.amount < row.length ? row[columnMap.amount] : "";
      if (!rawAmount) continue;
      const cleaned = rawAmount.replace(/[^0-9.-]/g, "");
      amount = Math.abs(parseFloat(cleaned) || 0);
      if (isNaN(amount) || amount === 0) continue;

      if (columnMap.type !== null && columnMap.type < row.length) {
        const rawType = String(row[columnMap.type] || "").toLowerCase();
        type = rawType.includes("cr") || rawType.includes("credit") ? "credit" : "debit";
      } else {
        const numVal = parseFloat(cleaned);
        type = numVal > 0 ? "credit" : "debit";
      }
    }

    const rawCat = columnMap.category !== null && columnMap.category < row.length
      ? String(row[columnMap.category] ?? "").trim()
      : "";
    const category = (rawCat ? categorizeByName(rawCat) : null) ?? categorize(description, amount, type);
    transactions.push({ row: i + 1, date, description, amount, type, category });
  }

  return transactions;
}

function parsedCsvToTransactions(headers: string[], rows: string[][]): Transaction[] {
  const columnMap = detectColumnsFromHeaders(headers);
  if (!columnMap) throw new Error(
    "Could not find required columns (Date + Description + Amount or Debit/Credit).\n\nMake sure your file has a header row with column names like:\n• Date, Description, Amount\n• Date, Narration, Debit, Credit, Balance"
  );
  return rowsToTransactions(headers, rows, columnMap);
}

async function parseCsvOrTxt(file: File): Promise<Transaction[]> {
  const text = await file.text();
  const parsed = parsePastedText(text);
  if (!parsed) throw new Error("Could not parse the file. Make sure it has a header row and at least one data row.");
  return parsedCsvToTransactions(parsed.headers, parsed.rows);
}

async function parseXlsxOrXls(file: File): Promise<Transaction[]> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error("No sheets found in the Excel file.");
  const sheet = workbook.Sheets[sheetName];

  const data: string[][] = XLSX.utils.sheet_to_json<string[]>(sheet, {
    header: 1,
    defval: "",
    raw: false,
    dateNF: "yyyy-mm-dd",
  });

  if (data.length < 2) throw new Error("The Excel file appears to be empty or has only one row.");

  // Find header row (containing date/description/amount keywords)
  let headerIdx = 0;
  for (let i = 0; i < Math.min(data.length, 10); i++) {
    const joined = data[i].join(" ").toLowerCase();
    if (
      joined.includes("date") || joined.includes("description") ||
      joined.includes("narration") || joined.includes("amount") ||
      joined.includes("debit") || joined.includes("credit")
    ) {
      headerIdx = i;
      break;
    }
  }

  const headers = data[headerIdx].map((h) => String(h ?? "").trim());
  const rows = data
    .slice(headerIdx + 1)
    .filter((row) => row.some((cell) => cell !== ""))
    .filter((row) => {
      const first = String(row[0] ?? "").toLowerCase();
      return !(
        first.includes("opening") || first.includes("closing") ||
        first.includes("total") || first.includes("balance b/f") ||
        first.includes("balance c/f")
      );
    })
    .map((row) => row.map((cell) => String(cell ?? "").trim()));

  if (rows.length === 0) throw new Error("No data rows found in the Excel file after the header.");
  return parsedCsvToTransactions(headers, rows);
}

async function parsePdf(file: File): Promise<Transaction[]> {
  const pdfjsLib = await import("pdfjs-dist");
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
    "pdfjs-dist/build/pdf.worker.mjs",
    import.meta.url
  ).toString();

  const buffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data: buffer });
  const pdf = await loadingTask.promise;

  const lines: string[] = [];
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();

    // Group items by approximate Y position to reconstruct rows
    const itemsByY = new Map<number, { x: number; text: string }[]>();
    for (const item of textContent.items) {
      if (!("str" in item)) continue;
      const str = item.str.trim();
      if (!str) continue;
      const transform = item.transform;
      const y = Math.round(transform[5]);
      const x = Math.round(transform[4]);
      if (!itemsByY.has(y)) itemsByY.set(y, []);
      itemsByY.get(y)!.push({ x, text: str });
    }

    // Sort rows top-to-bottom (higher y = higher on page in PDF coords)
    const sortedYs = Array.from(itemsByY.keys()).sort((a, b) => b - a);
    for (const y of sortedYs) {
      const items = itemsByY.get(y)!.sort((a, b) => a.x - b.x);
      lines.push(items.map((i) => i.text).join("\t"));
    }
  }

  const text = lines.join("\n");
  const parsed = parsePastedText(text);
  if (!parsed) throw new Error(
    "Could not extract a table from the PDF.\n\nMake sure the PDF contains selectable/copyable text (not a scanned image). Try copying text from the PDF first to verify."
  );
  return parsedCsvToTransactions(parsed.headers, parsed.rows);
}

export async function parseFileToTransactions(file: File): Promise<Transaction[]> {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".txt")) return parseCsvOrTxt(file);
  if (name.endsWith(".xlsx") || name.endsWith(".xls")) return parseXlsxOrXls(file);
  if (name.endsWith(".pdf")) return parsePdf(file);
  throw new Error(`Unsupported file format. Please upload a .xlsx, .xls, .csv, .txt, or .pdf file.`);
}

export const SUPPORTED_EXTENSIONS = [".xlsx", ".xls", ".csv", ".txt", ".pdf"];
export const ACCEPTED_MIME_TYPES = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "text/csv",
  "text/plain",
  "application/pdf",
].join(",");
