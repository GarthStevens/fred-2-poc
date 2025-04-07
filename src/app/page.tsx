'use client'

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { HotTable } from '@handsontable/react-wrapper';
import { ExportedCellChange, ExportedChange, HyperFormula } from 'hyperformula';
import { Workbook } from 'exceljs';
import { useRef, useState } from 'react';
import { registerAllModules } from 'handsontable/registry';
import { CellChange, ChangeSource } from "handsontable/common";

import 'handsontable/styles/handsontable.min.css';
import 'handsontable/styles/ht-theme-main.min.css';

registerAllModules();

const config = {
  licenseKey: 'gpl-v3',
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [hfInstance, setHFInstance] = useState<HyperFormula | null>(HyperFormula.buildEmpty(config));
  const [entitlement, setEntitlement] = useState<number>(0);
  const [total, setTotal] = useState<number>(0);
  const hotTableRef = useRef<any>(null);

  function handleCellChange(changes: ExportedChange[]) {
    console.log(changes);

    for (const change of changes) {
      if (change instanceof ExportedCellChange) {
        if (change.address.sheet === 0 && change.address.col === 1 && change.address.row === 1) {
          setTotal(change.newValue as number);
        }
      }
    }
  }

  async function handleFileUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    setFile(file);
  }

  async function handleImport() {
    if (!file) return;

    if (hfInstance) {
      hfInstance.destroy();
      setHFInstance(null);
    }

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

    hf.on('valuesUpdated', handleCellChange);

    setHFInstance(hf);
    setFile(null);

    const entitlement = hf.getCellValue({ sheet: 0, col: 1, row: 0 }) as number;
    setEntitlement(entitlement);

    const total = hf.getCellValue({ sheet: 0, col: 1, row: 1 }) as number;
    setTotal(total);

    setFile(null);
  }

  async function handleExport() {
    if (!hfInstance) return;

    // Create a new Excel workbook using ExcelJS
    const workbook = new Workbook();

    // Get all sheet names from HyperFormula
    const sheetNames = hfInstance.getSheetNames();

    // Loop through each sheet
    for (const sheetName of sheetNames) {
      // Get the sheet index for the given sheet name
      const sheetId = hfInstance.getSheetId(sheetName);
      // Retrieve the entire sheet data as a 2D array
      const sheetData = hfInstance.getSheetSerialized(sheetId!);

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

  function handleEntitlementChange(event: React.ChangeEvent<HTMLInputElement>) {
    const value = event.target.value;
    setEntitlement(Number(value));
    if (hfInstance) {
      hfInstance.setCellContents({ sheet: 0, col: 1, row: 0 }, Number(value));
    }
  }

  function refreshTableData() {
    if (hfInstance && hotTableRef.current) {
      // Load the recalculated sheet data from HyperFo
      const newData = hfInstance.getSheetSerialized(1);
      hotTableRef.current.hotInstance.loadData(newData);
    }
  };

  function afterChange(changes: CellChange[] | null, source: ChangeSource) {
    if (!hfInstance || source === 'loadData' || !changes) return;

    changes.forEach((change) => {
      if (change instanceof ExportedCellChange && change.address.sheet === 1) {
        hfInstance.setCellContents({ sheet: 1, row: change.address.row, col: change.address.col }, [[change.newValue as any]]);
      }
    });

    refreshTableData();
  };

  return (
    <div className="flex flex-col gap-3">
      <div>
        <h1 className="text-2xl font-bold">Fred v2 POC</h1>
      </div>
      <div className="flex max-w-sm items-center gap-1.5">
        <Input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
        <Button disabled={!file} onClick={handleImport}>Import</Button>
        <Button disabled={!hfInstance} onClick={handleExport}>Export</Button>
      </div>

      {
        hfInstance
          ? (
            <>
              <div className="grid w-full max-w-sm items-center gap-1.5">
                <Label htmlFor="entitlement">Common Entitlement</Label>
                <Input id="entitlement" type="number" value={entitlement} onChange={handleEntitlementChange} step={0.1} />
              </div>
              <div className="grid w-full max-w-sm items-center gap-1.5">
                <Label htmlFor="total">Total</Label>
                <Input id="total" type="number" value={total} readOnly />
              </div>
              <hr className="my-2" />
              <div className="grid w-full items-center gap-1.5">
                <Label>Unit Assets</Label>
                <div className="ht-theme-main-dark-auto">
                  <HotTable
                    // ref={hotTableRef}
                    formulas={{ engine: hfInstance,  }}
                    data={[
                      [10.26, null, 'Sum', '=SUM(A:A)'],
                      [20.12, null, 'Average', '=AVERAGE(A:A)'],
                      [30.01, null, 'Median', '=MEDIAN(A:A)'],
                      [40.29, null, 'MAX', '=MAX(A:A)'],
                      [50.18, null, 'MIN', '=MIN(A1:A5)'],
                    ]}
                    rowHeaders
                    colHeaders
                    height="auto"
                    licenseKey="non-commercial-and-evaluation"
                    // afterChange={afterChange}
                  />
                </div>
              </div>
            </>
          )
          : null
      }

    </div>
  );
}
