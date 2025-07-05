
import * as XLSX from 'xlsx-js-style';
import type { IndexSheetConfig, SplitterCustomHeaderConfig, SplitterCustomColumnConfig } from './excel-types';
import { parseColumnIdentifier, parseSourceColumns, sanitizeSheetName, sanitizeSheetNameForFormula } from './excel-helpers';

export const INTERNAL_NOT_FOUND_KEY = '__INTERNAL_NOT_FOUND_ROWS__';
export const DISPLAY_NOT_FOUND_SHEET_NAME = 'not_found';
const MAX_SHEET_NAME_LENGTH = 31;


export interface GroupingResult {
  groupedRows: { [key: string]: number[] };
  orderedKeys: string[];
}

/**
 * Groups rows from a worksheet by the values in a specified column, preserving the original encounter order of keys.
 * @param worksheet The worksheet to process.
 * @param colIndex The 0-indexed column index to group by.
 * @param headerRowNumber The 1-indexed row number containing the headers.
 * @returns An object containing the grouped rows and an array of keys in the order they first appeared.
 */
export function groupDataByColumn(worksheet: XLSX.WorkSheet, colIndex: number, headerRowNumber: number): GroupingResult {
  const groupedRows: { [key: string]: number[] } = {};
  const orderedKeys: string[] = []; // This will store keys in the order they appear
  if (!worksheet || !worksheet['!ref']) {
    return { groupedRows, orderedKeys };
  }
  
  const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, blankrows: false });
  const headerRowIndex = headerRowNumber - 1;
  const headers = aoa[headerRowIndex]?.map(String) || [];
  
  if (colIndex < 0 || colIndex >= headers.length) {
    throw new Error(`Invalid column index for grouping: ${colIndex}. Sheet has ${headers.length} columns.`);
  }
  
  // We iterate from the row *after* the header row
  for (let R = headerRowNumber; R < aoa.length; ++R) {
    const row = aoa[R];
    if (!row || row.every(cell => cell === null)) continue;
    
    const value = row[colIndex];
    const groupKey = (value === null || value === undefined || String(value).trim() === '') 
      ? INTERNAL_NOT_FOUND_KEY 
      : String(value);

    if (!groupedRows[groupKey]) {
      groupedRows[groupKey] = [];
      orderedKeys.push(groupKey); // Add the key to our ordered list when it's first seen
    }
    groupedRows[groupKey].push(R); // Store 0-indexed row index
  }

  return { groupedRows, orderedKeys };
}


/**
 * Creates a new Excel workbook from grouped data, preserving cell formatting.
 * Each group becomes a new sheet. It copies cell objects directly to maintain styles.
 * @param sourceWorksheet The original worksheet containing all data and formatting.
 * @param groupedRows An object mapping group keys to arrays of 0-indexed row numbers.
 * @param headersForSheet An array of header names defining the desired column order for the new sheets.
 * @param originalHeaderRowIndex The 0-indexed number of the header row in the source worksheet.
 * @param customHeaderConfig Optional configuration for inserting a new custom header.
 * @param indexSheetConfig Optional configuration for creating a hyperlinked index sheet.
 * @param onProgress Optional callback for progress reporting.
 * @param cancellationRequestedRef Optional ref to check for cancellation requests.
 * @param orderedKeys Optional array of keys to define the sheet creation order.
 * @param customColumnConfig Optional configuration for inserting a new custom column.
 * @returns A new XLSX.WorkBook object with split sheets.
 */
export function createWorkbookFromGroupedData(
  sourceWorksheet: XLSX.WorkSheet,
  groupedRows: { [key: string]: number[] },
  headersForSheet: string[],
  originalHeaderRowIndex: number,
  customHeaderConfig?: Omit<SplitterCustomHeaderConfig, 'text'>,
  indexSheetConfig?: IndexSheetConfig,
  customColumnConfig?: SplitterCustomColumnConfig,
  onProgress?: (status: { groupKey: string; currentGroup: number; totalGroups: number; }) => void,
  cancellationRequestedRef?: React.RefObject<boolean>,
  orderedKeys?: string[]
): XLSX.WorkBook {
  const newWorkbook = XLSX.utils.book_new();
  const allocatedSheetNames = new Set<string>();
  const dataSheetNames: string[] = [];

  const sourceAOA: any[][] = XLSX.utils.sheet_to_json(sourceWorksheet, { header: 1, defval: null });
  const originalHeaders = sourceAOA[originalHeaderRowIndex]?.map(String) || [];
  
  // Create a map to handle duplicate header names correctly.
  // Maps header name to an array of its original 0-indexed column positions.
  const originalHeaderMap = new Map<string, number[]>();
  originalHeaders.forEach((h, i) => {
      const headerName = h || ''; // Treat null/undefined headers as empty string
      if (!originalHeaderMap.has(headerName)) {
          originalHeaderMap.set(headerName, []);
      }
      originalHeaderMap.get(headerName)!.push(i);
  });
  
  const groupKeys = orderedKeys || Object.keys(groupedRows);

  for (let i = 0; i < groupKeys.length; i++) {
    const groupKey = groupKeys[i];
    if (cancellationRequestedRef?.current) throw new Error('Cancelled by user.');
    onProgress?.({ groupKey, currentGroup: i + 1, totalGroups: groupKeys.length });

    const rowIndicesInSource = groupedRows[groupKey];
    
    let baseSheetName = (groupKey === INTERNAL_NOT_FOUND_KEY) ? DISPLAY_NOT_FOUND_SHEET_NAME : sanitizeSheetName(groupKey);
    let finalSheetName = baseSheetName;

    let counter = 1;
    while (allocatedSheetNames.has(finalSheetName.toLowerCase())) {
        const suffix = `_${counter}`;
        const baseMaxLength = MAX_SHEET_NAME_LENGTH - suffix.length;
        finalSheetName = `${baseSheetName.substring(0, baseMaxLength)}${suffix}`;
        counter++;
    }
    allocatedSheetNames.add(finalSheetName.toLowerCase());
    dataSheetNames.push(finalSheetName);

    const newSheet: XLSX.WorkSheet = { '!merges': [], '!ref': 'A1' };
    let newSheetRowIndex = 0;

    // 1. Insert Custom Header
    if (customHeaderConfig) {
        let dynamicHeaderText = '';
        if (rowIndicesInSource.length > 0) {
            const firstDataRowForGroupIndex = rowIndicesInSource[0];
            const firstDataRowForGroup = sourceAOA[firstDataRowForGroupIndex];
            if (firstDataRowForGroup) {
                const sourceColIndices = parseSourceColumns(customHeaderConfig.sourceColumnString, originalHeaders);
                const valuesToJoin = sourceColIndices.map(colIdx => (firstDataRowForGroup?.[colIdx] ?? ""));
                dynamicHeaderText = valuesToJoin.join(customHeaderConfig.valueSeparator);
            }
        }
        
        if (newSheetRowIndex < customHeaderConfig.insertBeforeRow - 1) {
            newSheetRowIndex = customHeaderConfig.insertBeforeRow - 1;
        }
        const numCols = Math.max(1, headersForSheet.length);

        if (customHeaderConfig.mergeAndCenter && numCols > 1) {
            newSheet['!merges']!.push({ s: { r: newSheetRowIndex, c: 0 }, e: { r: newSheetRowIndex, c: numCols - 1 } });
        }
        
        const headerCellAddr = XLSX.utils.encode_cell({ r: newSheetRowIndex, c: 0 });
        newSheet[headerCellAddr] = {
            v: dynamicHeaderText,
            t: 's',
            s: {
                font: {
                    bold: customHeaderConfig.styleOptions.bold,
                    italic: customHeaderConfig.styleOptions.italic,
                    underline: customHeaderConfig.styleOptions.underline,
                    name: customHeaderConfig.styleOptions.fontName || 'Calibri',
                    sz: customHeaderConfig.styleOptions.fontSize || 12,
                },
                alignment: {
                    horizontal: customHeaderConfig.styleOptions.alignment || 'center',
                    vertical: 'center'
                }
            }
        };
        newSheetRowIndex++;
    }

    let mainHeadersFinalIndex = newSheetRowIndex;

    // 2. Add Main Headers from headersForSheet, preserving styles and making bold
    const headerOccurrencesForHeaders = new Map<string, number>();
    headersForSheet.forEach((header, newColIdx) => {
        const isCustomCol = customColumnConfig && header === customColumnConfig.name;
        
        const occurrence = headerOccurrencesForHeaders.get(header) || 0;
        const originalIndices = originalHeaderMap.get(header);
        const styleSourceColIdx = originalIndices ? originalIndices[occurrence] : -1;
        headerOccurrencesForHeaders.set(header, occurrence + 1);

        if (isCustomCol) {
            const destAddr = XLSX.utils.encode_cell({ r: newSheetRowIndex, c: newColIdx });
            newSheet[destAddr] = { t: 's', v: header, s: { font: { bold: true } } };
        } else if (styleSourceColIdx !== -1) {
            const sourceAddr = XLSX.utils.encode_cell({ r: originalHeaderRowIndex, c: styleSourceColIdx });
            if (sourceWorksheet[sourceAddr]) {
                const destAddr = XLSX.utils.encode_cell({ r: newSheetRowIndex, c: newColIdx });
                const newCell = JSON.parse(JSON.stringify(sourceWorksheet[sourceAddr])); // Deep copy
                if (!newCell.s) newCell.s = {};
                if (!newCell.s.font) newCell.s.font = {};
                newCell.s.font.bold = true;
                newSheet[destAddr] = newCell;
            } else {
                 const destAddr = XLSX.utils.encode_cell({ r: newSheetRowIndex, c: newColIdx });
                 newSheet[destAddr] = { t:'s', v: header, s: { font: { bold: true }}};
            }
        }
    });
    newSheetRowIndex++;

    // 3. Add Data Rows, preserving styles
    rowIndicesInSource.forEach(sourceRowIdx => {
        const headerOccurrencesForData = new Map<string, number>();
        headersForSheet.forEach((header, newColIdx) => {
            const isCustomCol = customColumnConfig && header === customColumnConfig.name;

            if (isCustomCol) {
                const destAddr = XLSX.utils.encode_cell({ r: newSheetRowIndex, c: newColIdx });
                const finalValue = customColumnConfig.value.replace(/{SheetName}/g, (groupKey === INTERNAL_NOT_FOUND_KEY ? DISPLAY_NOT_FOUND_SHEET_NAME : groupKey));
                newSheet[destAddr] = { t: 's', v: finalValue };
            } else {
                const occurrence = headerOccurrencesForData.get(header) || 0;
                const originalIndices = originalHeaderMap.get(header);
                const originalColIdx = originalIndices ? originalIndices[occurrence] : -1;
                headerOccurrencesForData.set(header, occurrence + 1);

                if (originalColIdx !== -1) {
                    const sourceAddr = XLSX.utils.encode_cell({ r: sourceRowIdx, c: originalColIdx });
                    if (sourceWorksheet[sourceAddr]) {
                        const destAddr = XLSX.utils.encode_cell({ r: newSheetRowIndex, c: newColIdx });
                        newSheet[destAddr] = JSON.parse(JSON.stringify(sourceWorksheet[sourceAddr])); // Deep copy
                    }
                }
            }
        });
        newSheetRowIndex++;
    });

    // 4. Finalize sheet properties
    if (newSheetRowIndex > 0) {
        const endRange = { r: newSheetRowIndex - 1, c: Math.max(0, headersForSheet.length - 1) };
        newSheet['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: endRange });

        if (headersForSheet.length > 0 && newSheetRowIndex > mainHeadersFinalIndex + 1) {
            newSheet['!autofilter'] = { ref: XLSX.utils.encode_range({s: {r: mainHeadersFinalIndex, c: 0}, e: {r: endRange.r, c: endRange.c}}) };
        }

        const aoaForWidth: any[][] = XLSX.utils.sheet_to_json(newSheet, { header: 1, defval: null });
        const colWidths = Array.from({ length: headersForSheet.length }).map((_, colIdx) => {
            let maxLength = 0;
            aoaForWidth.forEach(row => {
                const cellValue = row?.[colIdx];
                const cellTextLength = cellValue !== null && cellValue !== undefined ? String(cellValue).length : 0;
                if (cellTextLength > maxLength) {
                    maxLength = cellTextLength;
                }
            });
            return { wch: Math.max(10, maxLength + 2) };
        });
        newSheet['!cols'] = colWidths;
    }
    
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, finalSheetName);
  }

  // 5. Create Index Sheet if configured
  if (indexSheetConfig) {
      const indexSheetName = sanitizeSheetName(indexSheetConfig.sheetName);
      const indexSheet = XLSX.utils.aoa_to_sheet([[]]);
      
      const headerCol = parseColumnIdentifier(indexSheetConfig.headerCol);
      if (headerCol !== null) {
          const headerAddr = XLSX.utils.encode_cell({ r: indexSheetConfig.headerRow - 1, c: headerCol });
          XLSX.utils.sheet_add_aoa(indexSheet, [[indexSheetConfig.headerText]], { origin: headerAddr });
          if(indexSheet[headerAddr]) indexSheet[headerAddr].s = { font: { bold: true, sz: 14 } };
      }

      const linksCol = parseColumnIdentifier(indexSheetConfig.linksCol);
      if (linksCol !== null) {
          dataSheetNames.sort((a, b) => a.localeCompare(b)).forEach((sheetName, i) => {
              const linkRow = (indexSheetConfig.linksStartRow - 1) + i;
              const linkAddr = XLSX.utils.encode_cell({ r: linkRow, c: linksCol });
              XLSX.utils.sheet_add_aoa(indexSheet, [[sheetName]], { origin: linkAddr });
              if (indexSheet[linkAddr]) {
                  indexSheet[linkAddr].l = { Target: `#${sanitizeSheetNameForFormula(sheetName)}!A1` };
                  indexSheet[linkAddr].s = { font: { color: { rgb: "0000FF" }, underline: true } };
              }
          });
      }
      
      const backLinkCol = parseColumnIdentifier(indexSheetConfig.backLinkCol);
      const backLinkRow = indexSheetConfig.backLinkRow - 1;

      if (backLinkCol !== null && backLinkRow >= 0) {
          dataSheetNames.forEach(sheetName => {
              const sheet = newWorkbook.Sheets[sheetName];
              if (sheet) {
                  const backLinkAddr = XLSX.utils.encode_cell({ r: backLinkRow, c: backLinkCol });
                  if (!sheet[backLinkAddr]) {
                      sheet[backLinkAddr] = { t: 's', v: indexSheetConfig.backLinkText };
                  } else {
                      sheet[backLinkAddr].v = indexSheetConfig.backLinkText;
                  }
                  sheet[backLinkAddr].l = { Target: `#${sanitizeSheetNameForFormula(indexSheetName)}!A1` };
                  sheet[backLinkAddr].s = { font: { color: { rgb: "0000FF" }, underline: true } };
              }
          });
      }
      
      XLSX.utils.book_append_sheet(newWorkbook, indexSheet, indexSheetName);
      newWorkbook.SheetNames = [indexSheetName, ...dataSheetNames.sort((a,b) => a.localeCompare(b))];
  }

  return newWorkbook;
}
