'use client'

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { HotTable } from '@handsontable/react-wrapper';
import { ExportedCellChange, ExportedChange, HyperFormula } from 'hyperformula';
import { Workbook } from 'exceljs';
import { useEffect, useState } from 'react';
import { registerAllModules } from 'handsontable/registry';
import { CellChange, ChangeSource } from "handsontable/common";
import numeral from 'numeral';

import 'handsontable/styles/handsontable.min.css';
import 'handsontable/styles/ht-theme-main.min.css';

registerAllModules();

const config = {
  licenseKey: 'gpl-v3',
}

export default function Home() {
  const [hfInstance, setHFInstance] = useState<HyperFormula | null>(null);
  const [entitlement, setEntitlement] = useState<number>(0);
  const [total, setTotal] = useState<number>(0);
  const [unitData, setUnitData] = useState<unknown[][]>([]);
  const [commonData, setCommonData] = useState<unknown[][]>([]);
  const [hasBulked, setHasBulked] = useState<boolean>(false);

  function afterUnitChange(changes: CellChange[] | null, source: ChangeSource) {
    if (!hfInstance || source === 'loadData' || !changes) return;

    changes.forEach(([row, col, , newValue]) => {
      hfInstance.setCellContents({ sheet: 1, row, col: col as number }, [[newValue]]);
    });
  }

  function afterCommonChange(changes: CellChange[] | null, source: ChangeSource) {
    if (!hfInstance || source === 'loadData' || !changes) return;

    changes.forEach(([row, col, , newValue]) => {
      hfInstance.setCellContents({ sheet: 2, row, col: col as number }, [[newValue]]);
    });
  }

  useEffect(() => {
    if (!hfInstance) return;
    const onValuesUpdated = (changes: ExportedChange[]) => {
      for (const change of changes) {
        if (change instanceof ExportedCellChange) {
          if (change.address.sheet === 0 && change.address.col === 1 && change.address.row === 1) {
            const newTotal = hfInstance.getCellValue({ sheet: 0, col: 1, row: 1 });
            setTotal(newTotal as number);
          }
        }
      }
    };

    hfInstance.on('valuesUpdated', onValuesUpdated);
    return () => {
      hfInstance.off('valuesUpdated', onValuesUpdated);
    };
  }, [hfInstance]);

  async function handleFileUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
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

    const hf = HyperFormula.buildFromSheets(sheetsData, config);

    hf.on('valuesUpdated', (changes) => {
      for (const change of changes) {
        if (change instanceof ExportedCellChange) {
          if (change.address.sheet === 0 && change.address.col === 1 && change.address.row === 1) {
            const newTotal = hf.getCellValue({ sheet: 0, col: 1, row: 1 });
            setTotal(newTotal as number);
          }
        }
      }
    });

    setHFInstance(hf);
    setUnitData(hf.getSheetSerialized(1));
    setCommonData(hf.getSheetSerialized(2));

    const entitlement = hf.getCellValue({ sheet: 0, col: 1, row: 0 }) as number;
    setEntitlement(entitlement);

    const total = hf.getCellValue({ sheet: 0, col: 1, row: 1 }) as number;
    setTotal(total);
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

  function handleBulkPopulateClick() {
    if (!hfInstance) return;
    console.log('Bulk Populate Clicked');

    hfInstance.removeRows(1, [1, 5])
    hfInstance.removeRows(2, [1, 5])

    const assetCount = 500

    hfInstance.batch(() => {
      for (let i = 1; i <= assetCount; i++) {
        const row = i + 1
        const rate = Math.floor(Math.random() * 1000) + 1;
        const quantity = Math.floor(Math.random() * 10) + 1;
        hfInstance.addRows(1, [i, 1]);
        hfInstance.setCellContents({ sheet: 1, col: 0, row: i }, [[`Unit Asset ${i}`, rate, quantity, `=B${row}*C${row}`]]);
      }

      for (let i = 1; i <= assetCount; i++) {
        const row = i + 1
        const rate = Math.floor(Math.random() * 1000) + 1;
        const quantity = Math.floor(Math.random() * 1000) + 1;
        hfInstance.addRows(2, [i, 1]);
        hfInstance.setCellContents({ sheet: 2, col: 0, row: i }, [[`Common Asset ${i}`, rate, quantity, `=B${row}*C${row}*Global!$B$1`]]);
      }
    })

    setUnitData(hfInstance.getSheetSerialized(1));
    setCommonData(hfInstance.getSheetSerialized(2));

    hfInstance.setCellContents({ sheet: 0, col: 1, row: 1 }, '=SUM(Unit!D2:D10006) + SUM(Common!D2:D10006)');
    setHasBulked(true);
  }

  return (
    <div className="flex justify-center w-full">
      <div className="w-full max-w-6xl px-4 py-6">
        <div className="flex flex-col gap-3">
          <div>
            <h1 className="text-2xl font-bold">Fred v2 POC</h1>
            <p>This POC should validate the following features:</p>
            <ul className="list-disc pl-6">
              <li><strong>Importing:</strong> Allowing the user to upload a master excel document</li>
              <li><strong>Exporting:</strong> Allowing the user to download an excel document that has all of the values populated</li>
              <li><strong>Cross sheet dependencies:</strong> Updating the Common Entitlement should also update the totals for the common assets</li>
              <li><strong>Readonly fields:</strong> The Total field should reflect the sum of the unit asset totals plus the sum of the common asset totals</li>
              <li><strong>Formula support:</strong> Allowing the user to use formulas in the excel document</li>
              <li><strong>Performance:</strong> Allowing the user to use a grid to edit the excel document</li>
            </ul>

            <h2 className="text-lg font-bold mt-4">How to use the POC</h2>
            <ol className="list-decimal pl-6">
              <li>Upload a master excel document: Click the &quot;Choose File&quot; button and select the master excel document</li>
              <li>Update the Common Entitlement</li>
              <li>Update the unit assets</li>
              <li>Update the common assets</li>
              <li>Press the Bulk Populate button (this will populate 500 unit assets and 500 common assets) and then update the unit assets and common assets to check performance</li>
              <li>Download the excel document</li>
            </ol>

            <h2 className="text-lg font-bold mt-4">Notes</h2>
            <ul className="list-disc pl-6">
              <li>The grids are completely editable. this will be locked down in the actual product</li>
              <li>We will need a license to use this grid. There are other options</li>
            </ul>
          </div>

          <hr className="my-2" />

          <div className="flex max-w-sm items-center gap-1.5">
            {!hfInstance && (<>
              <Input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
              <hr className="my-2" />
            </>
            )}
          </div>

          {
            hfInstance
              ? (
                <>
                  {!hasBulked && (
                    <>
                      <div className="grid w-full max-w-sm gap-1.5">
                        <Button onClick={handleBulkPopulateClick}>Bulk Populate</Button>
                      </div>
                      <hr className="my-2" />
                    </>
                  )}

                  <div className="grid w-full max-w-sm gap-1.5">
                    <Label htmlFor="entitlement">Common Entitlement</Label>
                    <Input id="entitlement" type="number" value={entitlement} onChange={handleEntitlementChange} step={0.1} />
                  </div>

                  <hr className="my-2" />

                  <div className="flex gap-4">
                    <div className="grid w-full items-center gap-1.5">
                      <Label>Unit Assets</Label>
                      <div className="ht-theme-main-dark-auto">
                        <HotTable
                          height={300}
                          formulas={{ engine: hfInstance, sheetName: 'Unit' }}
                          data={unitData}
                          rowHeaders
                          licenseKey="non-commercial-and-evaluation"
                          afterChange={afterUnitChange}
                          stretchH="all"
                        />
                      </div>
                    </div>

                    <div className="grid w-full items-center gap-1.5">
                      <Label>Common Assets</Label>
                      <div className="ht-theme-main-dark-auto">
                        <HotTable
                          height={300}
                          formulas={{ engine: hfInstance, sheetName: 'Common' }}
                          data={commonData}
                          rowHeaders
                          licenseKey="non-commercial-and-evaluation"
                          afterChange={afterCommonChange}
                          stretchH="all"
                        />
                      </div>
                    </div>
                  </div>

                  <hr className="my-2" />

                  <div className="grid w-full max-w-sm items-center gap-1.5">
                    <Label htmlFor="total">Total</Label>
                    <Input id="total" type="text" value={numeral(total).format('$0,0.00')} readOnly />
                  </div>

                  <hr className="my-2" />

                  <Button onClick={handleExport} className="w-full max-w-sm">Export</Button>
                </>
              )
              : null
          }
        </div>
      </div>
    </div >
  );
}
