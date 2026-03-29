import express from "express";
import multer from "multer";
import { parse } from "csv-parse/sync";
import ExcelJS from "exceljs";
import path from "path";
import { createServer as createViteServer } from "vite";
import fs from "fs";

const app = express();
const PORT = 3000;
const upload = multer({ storage: multer.memoryStorage() });

interface TempoTask {
  epic: string;
  hours: number;
}

function parseTempoCsv(csvContent: string): Record<string, TempoTask> {
  const records = parse(csvContent, {
    columns: true,
    skip_empty_lines: true,
    bom: true,
  });

  const tasks: Record<string, TempoTask> = {};

  for (const row of records) {
    const issue = (row["Issue"] || "").trim();
    const worklog = (row["Worklog"] || "").trim();
    const billable = parseFloat(row["Billable"] || "0");
    const subtask = (row["Sub-task"] || "").trim();
    const epic = (row["Epic"] || "").trim();

    // Skip sub-task rows as per python script logic
    if (subtask) continue;

    if (issue && !worklog && row["User"] !== "Total") {
      let finalIssue = issue;
      if (issue === "No Issue") {
        if (epic && epic !== "No Epic") {
          finalIssue = epic;
        } else {
          continue;
        }
      }

      if (!tasks[finalIssue]) {
        tasks[finalIssue] = { epic: epic === "No Epic" ? "" : epic, hours: 0 };
      }
      tasks[finalIssue].hours += billable;
    }
  }

  return tasks;
}

async function generateExcel(tasks: Record<string, TempoTask>): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Estimation");

  const DATA_START_ROW = 4;
  const NUM_FMT = "0.00";
  const PCT_FMT = "0.00%";
  const CURRENCY_FMT = "$#,##0.00";

  // Colors (ARGB)
  const colors = {
    black: "FF000000",
    white: "FFFFFFFF",
    yellow: "FFFFFF00",
    orange: "FFF4B084",
    lightGray: "FFF2F2F2",
    headerGray: "FFD9D9D9",
    blue: "FF4472C4"
  };

  // Borders and Styles
  const mediumBorder = { style: "medium" as const };
  const thinBorder = { style: "thin" as const };

  const centerAlign: Partial<ExcelJS.Alignment> = { horizontal: "center", vertical: "middle", wrapText: true };
  const leftAlign: Partial<ExcelJS.Alignment> = { horizontal: "left", vertical: "middle", wrapText: true };

  // Row 1: Header Info
  worksheet.mergeCells("B1:D1");
  worksheet.getCell("B1").value = "Any scope out of the document is not included in the estimation.";
  worksheet.getCell("B1").font = { italic: true, size: 10 };

  worksheet.mergeCells("E1:M1");
  worksheet.getCell("E1").value = "Estimation by Resource";
  worksheet.getCell("E1").font = { bold: true, size: 12 };
  worksheet.getCell("E1").alignment = centerAlign;
  worksheet.getCell("E1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerGray } };

  worksheet.getCell("N1").value = "Total \nHours";
  worksheet.getCell("N1").alignment = centerAlign;
  worksheet.getCell("N1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.black } };
  worksheet.getCell("N1").font = { color: { argb: colors.white }, bold: true };

  worksheet.getCell("O1").value = "Plan";
  worksheet.mergeCells("O1:P1");
  worksheet.getCell("O1").alignment = centerAlign;
  worksheet.getCell("O1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerGray } };

  worksheet.getCell("Q1").value = "Actual";
  worksheet.mergeCells("Q1:S1");
  worksheet.getCell("Q1").alignment = centerAlign;
  worksheet.getCell("Q1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.yellow } };

  worksheet.getCell("T1").value = "KPIs";
  worksheet.mergeCells("T1:W1");
  worksheet.getCell("T1").alignment = centerAlign;
  worksheet.getCell("T1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.orange } };

  worksheet.getCell("X1").value = "Savings";
  worksheet.getCell("X1").alignment = centerAlign;
  worksheet.getCell("X1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerGray } };

  // Row 2: Main Headers
  const headers = [
    "", "Epic", "Task / User Story", "Acceptance Criteria", "Design", "Frontend", "Backend", 
    "Mobile", "Desktop", "API", "DevOps", "QA", "PM", "Total Hours", "Start Date", "Due Date", 
    "Planned Value", "Hours Spent", "Completed %", "Earned Value", "SPI", "CPI", "EAC", "Hours Left"
  ];
  
  const headerRow = worksheet.getRow(2);
  headers.forEach((h, i) => {
    if (!h) return;
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true };
    cell.alignment = centerAlign;
    cell.border = { top: mediumBorder, bottom: mediumBorder, left: thinBorder, right: thinBorder };
    
    // Specific header colors
    if (i + 1 === 14) { // N
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.black } };
      cell.font = { color: { argb: colors.white }, bold: true };
    } else if (i + 1 >= 17 && i + 1 <= 19) { // Q, R, S
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.yellow } };
    } else if (i + 1 >= 20 && i + 1 <= 23) { // T, U, V, W
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.orange } };
    } else if (i + 1 >= 5 && i + 1 <= 13 || i + 1 === 15 || i + 1 === 16 || i + 1 === 24) {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.headerGray } };
    }
  });

  const sortedTasks = Object.entries(tasks).sort((a, b) => {
    if (a[1].epic !== b[1].epic) return a[1].epic.localeCompare(b[1].epic);
    return a[0].localeCompare(b[0]);
  });

  let currentRow = DATA_START_ROW;
  for (const [taskName, info] of sortedTasks) {
    const row = worksheet.getRow(currentRow);
    row.getCell(2).value = info.epic;
    row.getCell(3).value = taskName;
    row.getCell(18).value = info.hours; // Hours Spent (AC)

    row.getCell(14).value = { formula: `SUM(E${currentRow}:M${currentRow})` }; // Total Hours (BAC)
    row.getCell(17).value = { formula: `IF(P${currentRow},IF(P${currentRow}<=TODAY(),N${currentRow},IF(O${currentRow}<TODAY(),N${currentRow}*NETWORKDAYS(O${currentRow},TODAY())/NETWORKDAYS(O${currentRow},P${currentRow}),0)),0)` }; // PV
    row.getCell(19).value = 0; // Completed % (Input)
    row.getCell(20).value = { formula: `N${currentRow}*S${currentRow}` }; // EV = BAC * %
    row.getCell(21).value = { formula: `IF(Q${currentRow}>0, T${currentRow}/Q${currentRow}, 0)` }; // SPI = EV / PV
    row.getCell(22).value = { formula: `IF(R${currentRow}>0, T${currentRow}/R${currentRow}, 0)` }; // CPI = EV / AC
    row.getCell(23).value = { formula: `IF(V${currentRow}>0, N${currentRow}/V${currentRow}, N${currentRow})` }; // EAC = BAC / CPI
    row.getCell(24).value = { formula: `N${currentRow}-R${currentRow}` }; // Hours Left

    for (let i = 2; i <= 24; i++) {
      const cell = row.getCell(i);
      cell.border = { left: thinBorder, right: thinBorder, bottom: thinBorder };
      
      // Apply column colors to data rows
      if (i === 14) { // N
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.black } };
        cell.font = { color: { argb: colors.white }, bold: true };
      } else if (i >= 17 && i <= 19) { // Q, R, S
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.yellow } };
      } else if (i >= 20 && i <= 23) { // T, U, V, W
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.orange } };
      }

      if ((i >= 5 && i <= 14) || i >= 17) {
        cell.numFmt = i === 19 ? PCT_FMT : NUM_FMT;
        cell.alignment = centerAlign;
      } else {
        cell.alignment = i === 3 ? leftAlign : centerAlign;
      }
    }
    currentRow++;
  }

  const dataEndRow = currentRow - 1;

  // Summary Rows
  const summaryData = [
    { name: "Site Reliability Engineering (5%)", desc: "Monitoring & Maintaining Servers & Infrastructure", rate: 0.05 },
    { name: "Quality Assurance (10%)", desc: "Manual testing of the system by QA Engineer", rate: 0.10 },
    { name: "Project Management (20%)", desc: "Management of the Project on all levels", rate: 0.20 },
    { name: "Communication / Agile Ceremonies (10%)", desc: "Daily Standup, Weekly Call, Sprint Planning", rate: 0.10 }
  ];

  summaryData.forEach((s, idx) => {
    const r = currentRow + idx;
    const row = worksheet.getRow(r);
    row.getCell(2).value = s.name;
    row.getCell(4).value = s.desc;
    
    if (s.name.includes("Project Management")) {
      // PM = 20% of all data rows (E to L)
      row.getCell(13).value = { formula: `ROUNDUP(SUM(E${DATA_START_ROW}:L${dataEndRow})*${s.rate},0)` };
    } else if (s.name.includes("Communication")) {
      // Communication = 10% of sum of data rows + previous summary rows
      for (let col = 5; col <= 12; col++) {
        const colLetter = worksheet.getColumn(col).letter;
        row.getCell(col).value = { formula: `SUM(${colLetter}${DATA_START_ROW}:${colLetter}${r-1})*${s.rate}` };
      }
    }

    row.getCell(14).value = { formula: `SUM(E${r}:M${r})` };
    row.getCell(17).value = { formula: `IF(P${r},IF(P${r}<=TODAY(),N${r},IF(O${r}<TODAY(),N${r}*NETWORKDAYS(O${r},TODAY())/NETWORKDAYS(O${r},P${r}),0)),0)` };
    row.getCell(18).value = 0;
    row.getCell(19).value = 0;
    row.getCell(20).value = { formula: `N${r}*S${r}` };
    row.getCell(21).value = { formula: `IF(Q${r}>0, T${r}/Q${r}, 0)` };
    row.getCell(22).value = { formula: `IF(R${r}>0, T${r}/R${r}, 0)` };
    row.getCell(23).value = { formula: `IF(V${r}>0, N${r}/V${r}, N${r})` };
    row.getCell(24).value = { formula: `N${r}-R${r}` };

    for (let i = 2; i <= 24; i++) {
      const cell = row.getCell(i);
      cell.border = { left: thinBorder, right: thinBorder, bottom: thinBorder };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.lightGray } };
      cell.numFmt = i === 19 ? PCT_FMT : NUM_FMT;
      cell.alignment = i === 4 ? leftAlign : centerAlign;
    }
  });

  const summaryEndRow = currentRow + summaryData.length - 1;
  currentRow += summaryData.length;

  // Subtotal Row
  const subtotalRowIndex = currentRow;
  const subtotalRow = worksheet.getRow(subtotalRowIndex);
  subtotalRow.getCell(2).value = "Subtotal";
  subtotalRow.getCell(2).font = { bold: true, color: { argb: colors.white } };
  worksheet.mergeCells(`B${subtotalRowIndex}:D${subtotalRowIndex}`);
  
  for (let col = 5; col <= 14; col++) {
    const colLetter = worksheet.getColumn(col).letter;
    subtotalRow.getCell(col).value = { formula: `SUM(${colLetter}${DATA_START_ROW}:${colLetter}${subtotalRowIndex - 1})` };
  }
  subtotalRow.getCell(15).value = "Totals";
  subtotalRow.getCell(17).value = { formula: `SUM(Q${DATA_START_ROW}:Q${subtotalRowIndex - 1})` };
  subtotalRow.getCell(18).value = { formula: `SUM(R${DATA_START_ROW}:R${subtotalRowIndex - 1})` };
  subtotalRow.getCell(19).value = { formula: `IF(N${subtotalRowIndex}>0, T${subtotalRowIndex}/N${subtotalRowIndex}, 0)` };
  subtotalRow.getCell(20).value = { formula: `SUM(T${DATA_START_ROW}:T${subtotalRowIndex - 1})` };
  subtotalRow.getCell(21).value = { formula: `IF(Q${subtotalRowIndex}>0, T${subtotalRowIndex}/Q${subtotalRowIndex}, 0)` };
  subtotalRow.getCell(22).value = { formula: `IF(R${subtotalRowIndex}>0, T${subtotalRowIndex}/R${subtotalRowIndex}, 0)` };
  subtotalRow.getCell(23).value = { formula: `IF(V${subtotalRowIndex}>0, N${subtotalRowIndex}/V${subtotalRowIndex}, N${subtotalRowIndex})` };
  subtotalRow.getCell(24).value = { formula: `N${subtotalRowIndex}-R${subtotalRowIndex}` };

  for (let i = 2; i <= 24; i++) {
    const cell = subtotalRow.getCell(i);
    cell.border = { top: mediumBorder, bottom: mediumBorder, left: thinBorder, right: thinBorder };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.black } };
    cell.font = { color: { argb: colors.white }, bold: true };
    cell.numFmt = i === 19 ? PCT_FMT : NUM_FMT;
    cell.alignment = centerAlign;
  }

  currentRow += 2;

  // Cost Section
  const costRows = [
    { name: "Resource Level", val: "0%" },
    { name: "Resource Rate / Hour", val: "" },
    { name: "Resource Cost Subtotal", val: "$0.00" },
    { name: "COST TOTAL", val: "$0.00" },
    { name: "Duration, weeks", val: "0.00" }
  ];

  const resourceLevelRow = currentRow;
  const rateRow = currentRow + 1;
  const costSubtotalRow = currentRow + 2;
  const costTotalRow = currentRow + 3;
  const durationRow = currentRow + 4;

  costRows.forEach((cr, idx) => {
    const r = currentRow + idx;
    const row = worksheet.getRow(r);
    row.getCell(4).value = cr.name;
    row.getCell(4).font = { bold: true };
    row.getCell(4).alignment = { horizontal: "right" };
    row.getCell(4).fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.lightGray } };

    // Set default values for columns E-M
    for (let col = 5; col <= 13; col++) {
      const cell = row.getCell(col);
      cell.border = { left: thinBorder, right: thinBorder, bottom: thinBorder, top: thinBorder };
      
      if (cr.name === "Resource Level") {
        cell.value = 0;
        cell.numFmt = PCT_FMT;
      } else if (cr.name === "Resource Rate / Hour") {
        cell.value = 0;
        cell.numFmt = CURRENCY_FMT;
      } else if (cr.name === "Resource Cost Subtotal") {
        const colLetter = worksheet.getColumn(col).letter;
        // Subtotal row * Rate row
        cell.value = { formula: `${colLetter}${subtotalRowIndex}*${colLetter}${rateRow}` };
        cell.numFmt = CURRENCY_FMT;
      } else if (cr.name === "Duration, weeks") {
        const colLetter = worksheet.getColumn(col).letter;
        // Subtotal row / 40
        cell.value = { formula: `${colLetter}${subtotalRowIndex}/40` };
        cell.numFmt = NUM_FMT;
      }
      cell.alignment = centerAlign;
    }

    if (cr.name === "COST TOTAL") {
      const totalCell = row.getCell(5);
      totalCell.value = { formula: `SUM(E${costSubtotalRow}:M${costSubtotalRow})` };
      totalCell.numFmt = CURRENCY_FMT;
      totalCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.black } };
      totalCell.font = { color: { argb: colors.white }, bold: true };
      totalCell.alignment = centerAlign;
      // Clear other cells in this row for visual consistency
      for (let col = 6; col <= 13; col++) {
        row.getCell(col).value = null;
      }
    }
  });

  // Final column widths
  worksheet.getColumn("B").width = 25;
  worksheet.getColumn("C").width = 50;
  worksheet.getColumn("D").width = 40;
  worksheet.getColumn("N").width = 12;

  return (await workbook.xlsx.writeBuffer()) as Buffer;
}

async function startServer() {
  // API routes
  app.post("/api/convert", upload.single("file"), async (req: any, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: "No file uploaded" });
      }

      const csvContent = req.file.buffer.toString("utf-8");
      const tasks = parseTempoCsv(csvContent);
      
      if (Object.keys(tasks).length === 0) {
        return res.status(400).json({ error: "No valid tasks found in CSV" });
      }

      const excelBuffer = await generateExcel(tasks);

      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename=Delivery_Report.xlsx`);
      res.send(excelBuffer);
    } catch (error) {
      console.error("Conversion error:", error);
      res.status(500).json({ error: "Failed to convert file" });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    if (fs.existsSync(distPath)) {
      app.use(express.static(distPath));
      app.get("*", (req, res) => {
        res.sendFile(path.join(distPath, "index.html"));
      });
    }
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
