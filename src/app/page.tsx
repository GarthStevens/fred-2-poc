'use client'

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { HyperFormula } from 'hyperformula';
import { Workbook } from 'exceljs';
import { useState, useRef } from 'react';

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [hf, setHF] = useState<HyperFormula | null>(null);

  async function handleFileUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    setFile(file);

  }

  async function handleImport() {
    if (!file) return;

    const buffer = await file.arrayBuffer();
    const workbook = new Workbook();
    await workbook.xlsx.load(buffer);

    const sheetsData: { [sheetName: string]: (string | number | null)[][] } = {};

    workbook.worksheets.forEach((worksheet) => {
      const data: (string | number | null)[][] = [];

      worksheet.eachRow((row, rowNumber) => {
        const rowData: (string | number | null)[] = [];
        row.eachCell((cell, colNumber) => {
          if (cell.formula) {
            rowData[colNumber - 1] = `=${cell.formula}`;
          } else {
            rowData[colNumber - 1] = cell.value as string | number | null;
          }
        });
        data[rowNumber - 1] = rowData;
      });

      sheetsData[worksheet.name] = data;
    });

    const hf = HyperFormula.buildFromSheets(sheetsData, {
      licenseKey: 'gpl-v3'
    });

    setHF(hf);
    setFile(null);

    const calculated = hf.getCellValue({ sheet: 2, col: 0, row: 0 });
    console.log('A1 value:', calculated);
  }

  async function handleExport() {
    if (!hf) return;

    // Create a new Excel workbook using ExcelJS
    const workbook = new Workbook();

    // Get all sheet names from HyperFormula
    const sheetNames = hf.getSheetNames();

    // Loop through each sheet
    for (const sheetName of sheetNames) {
      // Get the sheet index for the given sheet name
      const sheetId = hf.getSheetId(sheetName);
      // Retrieve the entire sheet data as a 2D array
      const sheetData = hf.getSheetSerialized(sheetId!);

      const ws = workbook.addWorksheet(sheetName);

      // Loop through the sheet data and write each cell
      sheetData.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          const excelCell = ws.getCell(rowIndex + 1, colIndex + 1);
          if (typeof cell === 'string' && cell.startsWith('=')) {
            // If the cell contains a formula, remove the '=' before setting it for ExcelJS
            excelCell.value = { formula: cell.substring(1) };
          } else {
            excelCell.value = cell;
          }
        });
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'export.xlsx';
    a.click();
  }


  return (
    <div>
      <div className="flex max-w-sm items-center gap-1.5">
        <Input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
        <Button disabled={!file} onClick={handleImport}>Import</Button>
        <Button disabled={!hf} onClick={handleExport}>Export</Button>
      </div>
      <div className="grid w-full max-w-sm items-center gap-1.5">
        <Label htmlFor="entitlement">Common Entitlement</Label>
        <Input id="entitlement" type="text" />
      </div>
    </div>
  );
}
