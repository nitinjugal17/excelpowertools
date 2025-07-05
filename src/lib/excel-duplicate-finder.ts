
import * as XLSX from 'xlsx-js-style';
import type { DuplicateReport } from './excel-types';
import { getColumnIndex, parseColumnIdentifier, sanitizeSheetNameForFormula } from './excel-helpers';


/**
 * Finds and marks duplicate rows based on a composite key from multiple columns.
 * This function MODIFIES THE WORKBOOK IN PLACE while preserving existing cell styles.
 * @param workbook The workbook to process.
 * @param sheetNames Array of sheet names to process.
 * @param keyColumnsIdentifiers Comma-separated columns that define a duplicate (e.g., "A,C" or "Name,Email").
 * @param updateColumnIdentifier The column to update for duplicate rows (e.g., "G" or "Status").
 * @param updateConfig Configuration for what value to write in the update column.
 * @param headerRow The 1-indexed row number where headers are located.
 * @param highlightColor Optional. If provided, applies a background fill color to the entire duplicate row.
 * @param conditionalColumnIdentifier Optional. If provided, a row is only marked as duplicate if the cell in this column is empty.
 * @param stripText Optional. Text to remove from a context-derived value before writing.
 * @param contextDelimiter Optional. Delimiter to split a context-derived value.
 * @param contextPartToUse Optional. The part of the split context value to use (1-indexed).
 * @param onProgress Optional callback to report progress.
 * @param cancellationRequestedRef Optional ref to check for cancellation requests.
 * @returns An object containing the report and the modified workbook.
 */
export function findAndMarkDuplicates(
  workbook: XLSX.WorkBook,
  sheetNames: string[],
  keyColumnsIdentifiers: string,
  updateColumnIdentifier: string,
  updateConfig: {
      mode: 'template' | 'context';
      value: string;
  },
  headerRow: number,
  highlightColor?: string,
  conditionalColumnIdentifier?: string,
  stripText?: string,
  contextDelimiter?: string,
  contextPartToUse?: number,
  onProgress?: (status: { sheetName: string; currentSheet: number; totalSheets: number; duplicatesFound: number }) => void,
  cancellationRequestedRef?: React.RefObject<boolean>
): { report: DuplicateReport; workbook: XLSX.WorkBook } {
    const report: DuplicateReport = {
        summary: {},
        updates: [],
        totalDuplicates: 0,
    };
    const headerRowIndex = headerRow - 1;

    for (let i = 0; i < sheetNames.length; i++) {
        const sheetName = sheetNames[i];
        if (cancellationRequestedRef?.current) {
            throw new Error('Cancelled by user.');
        }
        
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) continue;
        
        onProgress?.({ sheetName, currentSheet: i + 1, totalSheets: sheetNames.length, duplicatesFound: report.totalDuplicates });

        report.summary[sheetName] = 0;
        const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

        if (aoa.length <= headerRowIndex) continue;

        const headers = aoa[headerRowIndex].map(h => String(h || ''));
        
        const keyColIndices = keyColumnsIdentifiers.split(',').map(id => getColumnIndex(id, headers)).filter((idx): idx is number => idx !== null);
        const updateColIndex = getColumnIndex(updateColumnIdentifier, headers);
        const conditionalColIndex = conditionalColumnIdentifier ? getColumnIndex(conditionalColumnIdentifier, headers) : null;
        
        if (keyColIndices.length === 0 || updateColIndex === null) {
            throw new Error(`Could not resolve all required columns on sheet "${sheetName}". Check identifiers and header row.`);
        }

        if (conditionalColumnIdentifier && conditionalColIndex === null) {
             throw new Error(`Could not resolve conditional column "${conditionalColumnIdentifier}" on sheet "${sheetName}".`);
        }

        const seen = new Map<string, number>();
        
        // Iterate over all data rows, respecting their original position
        for (let R = headerRow; R < aoa.length; R++) {
            const row = aoa[R];
            if (!row || row.every(cell => cell === null)) continue;
        
            const compositeKey = keyColIndices.map(idx => String(row[idx] ?? '').toLowerCase()).join('~!~');
            if (!compositeKey) continue;

            if (seen.has(compositeKey)) {
                // Conditional Marking Check
                if (conditionalColIndex !== null) {
                    const conditionalCell = row[conditionalColIndex];
                    const isConditionalCellEmpty = conditionalCell === null || conditionalCell === undefined || String(conditionalCell).trim() === '';
                    // If the conditional cell is NOT empty, we skip marking this duplicate.
                    if (!isConditionalCellEmpty) {
                        continue; 
                    }
                }

                report.summary[sheetName]++;
                report.totalDuplicates++;
                const firstInstanceRow = seen.get(compositeKey)!;
                const excelRowIndex = R;

                const rowObject = headers.reduce((obj, header, i) => {
                    if (header) obj[header] = row[i] ?? null;
                    return obj;
                }, {} as Record<string, any>);

                let valueToUpdate: any;
                
                if (updateConfig.mode === 'context') {
                    const originalRowIndex = firstInstanceRow - 1;
                    const originalRowData = aoa[originalRowIndex];
                    const contextColIdentifiers = updateConfig.value;
                    const contextColIndices = contextColIdentifiers.split(',').map(id => getColumnIndex(id, headers)).filter((idx): idx is number => idx !== null);
                    
                    valueToUpdate = undefined;
                    if (originalRowData) {
                        for (const colIdx of contextColIndices) {
                            const contextCellVal = originalRowData[colIdx];
                            if (contextCellVal !== null && contextCellVal !== undefined && String(contextCellVal).trim() !== '') {
                                valueToUpdate = contextCellVal;
                                break;
                            }
                        }
                    }

                    if (stripText && typeof valueToUpdate === 'string') {
                        const regex = new RegExp(stripText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
                        valueToUpdate = valueToUpdate.replace(regex, '').trim();
                    }

                    if (contextDelimiter && typeof valueToUpdate === 'string') {
                        const parts = valueToUpdate.split(contextDelimiter);
                        const indexToUse = contextPartToUse === -1 ? parts.length - 1 : (contextPartToUse || 1) - 1;
                        if (indexToUse >= 0 && indexToUse < parts.length) {
                            valueToUpdate = parts[indexToUse].trim();
                        } else {
                            valueToUpdate = ''; // Default to empty string if part not found
                        }
                    }
                } else { // 'template' mode
                    const templateValue = updateConfig.value;
                    const pureTemplateMatch = templateValue.match(/^{([^{}]+)}$/);

                    if (pureTemplateMatch) {
                        const colIdentifier = pureTemplateMatch[1].trim();
                        const colIdx = getColumnIndex(colIdentifier, headers);
                        if (colIdx !== null) {
                            valueToUpdate = row[colIdx];
                        } else {
                            valueToUpdate = templateValue;
                        }
                    } else if (templateValue.includes('{') && templateValue.includes('}')) {
                        valueToUpdate = templateValue.replace(/{([^{}]+)}/g, (match, colIdentifier) => {
                            const colIdx = getColumnIndex(colIdentifier.trim(), headers);
                            if (colIdx !== null && row[colIdx] !== null && row[colIdx] !== undefined) {
                                return String(row[colIdx]);
                            }
                            return '';
                        });
                    } else {
                        valueToUpdate = templateValue;
                    }
                }
                
                const updateCellAddress = XLSX.utils.encode_cell({ r: excelRowIndex, c: updateColIndex });

                let cellToUpdate = worksheet[updateCellAddress];
                if (!cellToUpdate) {
                    worksheet[updateCellAddress] = {};
                    cellToUpdate = worksheet[updateCellAddress];
                }

                cellToUpdate.v = valueToUpdate;
                if (valueToUpdate === null || valueToUpdate === undefined) {
                    delete cellToUpdate.v;
                    cellToUpdate.t = 'z';
                } else if (typeof valueToUpdate === 'number') {
                    cellToUpdate.t = 'n';
                } else if (typeof valueToUpdate === 'boolean') {
                    cellToUpdate.t = 'b';
                } else if (valueToUpdate instanceof Date) {
                    cellToUpdate.t = 'd';
                } else {
                    cellToUpdate.t = 's';
                }

                if (highlightColor) {
                    const rowRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
                    for (let C = rowRange.s.c; C <= rowRange.e.c; ++C) {
                        const cellAddress = XLSX.utils.encode_cell({ r: excelRowIndex, c: C });
                        const cellToHighlight = worksheet[cellAddress] || { t: 'z', v: undefined };
                        if (!worksheet[cellAddress]) {
                            worksheet[cellAddress] = cellToHighlight;
                        }
            
                        if (!cellToHighlight.s) cellToHighlight.s = {};
                        cellToHighlight.s.fill = {
                            ...(cellToHighlight.s.fill || {}),
                            patternType: 'solid',
                            fgColor: { rgb: highlightColor.replace('#', '') }
                        };
                    }
                }
                
                report.updates.push({
                    sheetName,
                    row: excelRowIndex + 1,
                    firstInstanceRow,
                    updatedAddress: updateCellAddress,
                    updatedValue: valueToUpdate,
                    key: compositeKey,
                    rowData: rowObject,
                });
            } else {
                seen.set(compositeKey, R + 1);
            }
        }

        if (report.summary[sheetName] === 0) {
            delete report.summary[sheetName];
        }
    }

    return { report, workbook };
}


/**
 * Creates a new workbook with a report on found duplicates.
 * If the report is too large, it will be split into multiple sheets.
 * @param report The DuplicateReport generated by `findAndMarkDuplicates`.
 * @param chunkSize The maximum number of data rows per sheet.
 * @returns A new XLSX.WorkBook object containing the report.
 */
export function createDuplicateReportWorkbook(report: DuplicateReport, chunkSize?: number): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  const baseSheetName = "Duplicate_Row_Report";
  const maxRowsPerSheet = (chunkSize && chunkSize > 0) ? chunkSize : 100000;
  
  const summaryData: any[][] = [['Sheet Name', 'Duplicate Rows Marked']];
  Object.entries(report.summary).forEach(([sheetName, count]) => {
    summaryData.push([sheetName, count]);
  });
  summaryData.push([]); // Spacer row
  summaryData.push([{v: 'Total Duplicates Marked', s: {font: {bold: true}}}, {v: report.totalDuplicates, s: {font: {bold: true}}}]);
  summaryData.push([]); // Spacer row

  const locationHeader = [[
      {v: 'Sheet Name', s: {font: {bold: true}}}, 
      {v: 'Row', s: {font: {bold: true}}},
      {v: 'Cell Updated', s: {font: {bold: true}}},
      {v: 'Value Written', s: {font: {bold: true}}},
      {v: 'Duplicate Key', s: {font: {bold: true}}},
    ]];
  
  const details = report.updates;
  const totalDetails = details.length;
  const numSheets = Math.ceil(totalDetails / (maxRowsPerSheet - 1));

  // Add Summary to the first sheet if there are details, or a dedicated sheet if not
  if (numSheets === 0) {
      const ws = XLSX.utils.aoa_to_sheet(summaryData);
      ws['!cols'] = [{ wch: 25 }, { wch: 30 }];
      XLSX.utils.book_append_sheet(wb, ws, "Duplicate_Summary");
      return wb;
  }
  
  for (let i = 0; i < numSheets; i++) {
      const sheetName = numSheets > 1 ? `${baseSheetName}_${i + 1}` : baseSheetName;
      const chunkStart = i * (maxRowsPerSheet - 1);
      const chunkEnd = Math.min(chunkStart + (maxRowsPerSheet - 1), totalDetails);
      const chunk = details.slice(chunkStart, chunkEnd);

      const updatesData = chunk.map(u => [u.sheetName, u.row, u.updatedAddress, u.updatedValue, u.key]);

      let finalAOA: any[][];
      let dataStartRow = 0;

      if (i === 0) { // Add summary to the first report sheet
          finalAOA = [...summaryData, ...locationHeader, ...updatesData];
          dataStartRow = summaryData.length + locationHeader.length;
      } else {
          finalAOA = [...locationHeader, ...updatesData];
          dataStartRow = locationHeader.length;
      }
      
      const ws = XLSX.utils.aoa_to_sheet(finalAOA, {cellDates: true});
      
      chunk.forEach((update, index) => {
        const cellAddress = XLSX.utils.encode_cell({ r: dataStartRow + index, c: 2 });
        if(ws[cellAddress]) {
            ws[cellAddress].l = { Target: `#${sanitizeSheetNameForFormula(update.sheetName)}!${update.updatedAddress}` };
            ws[cellAddress].s = { font: { color: { rgb: "0000FF" }, underline: true } };
        }
      });
      
      ws['!cols'] = [{ wch: 25 }, { wch: 10 }, { wch: 15 }, { wch: 25 }, { wch: 40 }];
      
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  return wb;
}

/**
 * Inserts a summary of duplicate rows into each corresponding sheet of a workbook.
 * @param workbook The workbook to modify.
 * @param report The duplicate report containing details of what was found.
 * @param config Configuration for where and how to insert the report.
 */
export function addDuplicateReportToSheets(
    workbook: XLSX.WorkBook,
    report: DuplicateReport,
    config: {
        insertCol: string;
        insertRow: number;
        primaryContextCol: string;
        fallbackContextCol: string;
        headerRow: number;
    }
) {
    const insertColIdx = parseColumnIdentifier(config.insertCol);
    const startRowIdx = config.insertRow - 1;

    if (insertColIdx === null || startRowIdx < 0) {
        console.error("Invalid column or row for inserting in-sheet duplicate report.");
        return;
    }

    const sheetsWithDuplicates = Object.keys(report.summary);

    sheetsWithDuplicates.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return;

        // Get all data for lookups
        const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
        if (aoa.length < config.headerRow) return; // Not enough rows for a header
        const headers = aoa[config.headerRow - 1].map(h => String(h || ''));

        const primaryContextColIdx = getColumnIndex(config.primaryContextCol, headers);
        const fallbackContextColIdx = config.fallbackContextCol ? getColumnIndex(config.fallbackContextCol, headers) : null;


        const reportHeaders = ['Duplicate Row', 'Context of Original', 'Context of Duplicate'];
        const aoaData: any[][] = [
            reportHeaders.map(h => ({ v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" } } } }))
        ];

        const duplicatesForSheet = report.updates.filter(u => u.sheetName === sheetName);
        if (duplicatesForSheet.length === 0) return;

        duplicatesForSheet.forEach(dup => {
            // Get context for CURRENT duplicate row
            let contextValue = '';
            if (primaryContextColIdx !== null) {
                const primaryValue = dup.rowData[headers[primaryContextColIdx]];
                if (primaryValue !== null && primaryValue !== undefined && String(primaryValue).trim() !== '') {
                    contextValue = String(primaryValue);
                }
            }
            if (!contextValue && fallbackContextColIdx !== null) {
                 const fallbackValue = dup.rowData[headers[fallbackContextColIdx]];
                if (fallbackValue !== null && fallbackValue !== undefined && String(fallbackValue).trim() !== '') {
                    contextValue = String(fallbackValue);
                }
            }

            // Get context for ORIGINAL row
            let originalContextValue = '';
            const originalRowIndex = dup.firstInstanceRow - 1;
            const originalRowDataArray = aoa[originalRowIndex];

            if (originalRowDataArray) {
                if (primaryContextColIdx !== null) {
                    const primaryValue = originalRowDataArray[primaryContextColIdx];
                    if (primaryValue !== null && primaryValue !== undefined && String(primaryValue).trim() !== '') {
                        originalContextValue = String(primaryValue);
                    }
                }
                if (!originalContextValue && fallbackContextColIdx !== null) {
                    const fallbackValue = originalRowDataArray[fallbackContextColIdx];
                    if (fallbackValue !== null && fallbackValue !== undefined && String(fallbackValue).trim() !== '') {
                        originalContextValue = String(fallbackValue);
                    }
                }
            }
            
            // If still no context, default to the original row number text
            const originalAtText = originalContextValue.trim() || `(Original at Row ${dup.firstInstanceRow})`;

            aoaData.push([
                { v: dup.row, t: 'n', l: { Target: `#${sanitizeSheetNameForFormula(sheetName)}!A${dup.row}` }, s: { font: { color: { rgb: "0000FF" }, underline: true } } },
                originalAtText,
                contextValue
            ]);
        });
        
        aoaData.push([]); // Spacer
        aoaData.push([
            { v: 'Total Duplicates', s: { font: { bold: true } } }, 
            duplicatesForSheet.length, 
            ''
        ]);
        
        XLSX.utils.sheet_add_aoa(worksheet, aoaData, { origin: { r: startRowIdx, c: insertColIdx }, cellDates: true });
        
        const currentRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        const endOfNewDataRow = startRowIdx + aoaData.length - 1;
        const endOfNewDataCol = insertColIdx + (aoaData[0]?.length || 1) - 1;
        
        if (endOfNewDataRow > currentRange.e.r) currentRange.e.r = endOfNewDataRow;
        if (endOfNewDataCol > currentRange.e.c) currentRange.e.c = endOfNewDataCol;
        
        worksheet['!ref'] = XLSX.utils.encode_range(currentRange);
    });
}
