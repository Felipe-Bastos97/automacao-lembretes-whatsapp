import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const projectDir = path.dirname(fileURLToPath(import.meta.url));
const outputDir = path.join(projectDir, "outputs", "exemplo");
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Atendimentos");
sheet.showGridLines = false;
sheet.freezePanes.freezeRows(1);

sheet.getRange("A1:E4").values = [
  ["paciente", "responsavel", "telefone", "data", "horario"],
  ["Ana Souza", "Carla Souza", "+55 34 99999-0001", new Date("2026-07-25T00:00:00"), "09:00"],
  ["Lucas Oliveira", "Marcos Oliveira", "+55 34 99999-0002", new Date("2026-07-25T00:00:00"), "10:30"],
  ["Beatriz Lima", "Juliana Lima", "+55 34 99999-0003", new Date("2026-07-26T00:00:00"), "14:00"],
];

sheet.getRange("A1:E1").format = {
  fill: "#075E54",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: "#054C44" },
};
sheet.getRange("A2:E4").format = {
  fill: "#F7FAF9",
  borders: { insideHorizontal: { style: "thin", color: "#D8E6E3" } },
  verticalAlignment: "center",
};
sheet.getRange("C2:C4").format.numberFormat = "@";
sheet.getRange("D2:D4").format.numberFormat = "dd/mm/yyyy";
sheet.getRange("E2:E4").format.numberFormat = "@";
sheet.getRange("A1:E4").format.rowHeight = 22;
sheet.getRange("A:A").format.columnWidth = 22;
sheet.getRange("B:B").format.columnWidth = 22;
sheet.getRange("C:C").format.columnWidth = 26;
sheet.getRange("D:D").format.columnWidth = 22;
sheet.getRange("E:E").format.columnWidth = 16;

const preview = await workbook.render({
  sheetName: "Atendimentos",
  range: "A1:E4",
  scale: 2,
  format: "png",
});
await fs.writeFile(path.join(outputDir, "clientes_exemplo_preview.png"), new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(outputDir, "clientes_exemplo.xlsx"));

const inspect = await workbook.inspect({
  kind: "table",
  range: "Atendimentos!A1:E4",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 6,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);
