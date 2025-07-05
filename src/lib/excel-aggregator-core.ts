import * as XLSX from 'xlsx-js-style';
import type { AggregationResult, MatchMode } from './excel-types';
import { getColumnIndex, parseSourceColumns, escapeRegex } from './excel-helpers';

/**
 * Aggregates data by searching for terms in specific columns across multiple sheets.
 * @param workbook The workbook to search in.
 * @param sheetNamesToSearch Array of sheet names to process.
 * @param searchColumnIdentifiers The columns to search in (e.g., "A,C" or "1,3").
 * @param valueToKeyMap A map where keys are search values (lowercase) and values are reporting keys.
 * @param headerRow The 1-indexed row number containing headers.
 * @param config Configuration for counting blank cells, aggregation mode, and generating a detailed blank report.
 * @param onProgress Optional callback to report progress.
 * @returns An AggregationResult object with total and per-sheet counts.
 */
export function aggregateData(
  workbook: XLSX.WorkBook,
  sheetNamesToSearch: string[],
  searchColumnIdentifiers: string,
  valueToKeyMap: Map<string, string>,
  headerRow: number,
  config: {
    aggregationMode?: 'valueMatch' | 'keyMatch';
    keyCountColumn?: string;
    discoverNewKeys?: boolean;
    conditionalColumn?: string;
    countBlanksInColumn?: string;
    generateBlankDetails?: boolean;
    blankCountingMode?: 'rowAware' | 'fullColumn';
    matchMode?: MatchMode;
    summaryTitleCell?: string;
  },
  onProgress?: (status: {
    stage: string;
    sheetName: string;
    currentSheet: number;
    totalSheets: number;
    currentTotals: { [key: string]: number };
  }) => void
): AggregationResult {
  console.log(`Starting data aggregation. Mode: ${config.aggregationMode}, Match Mode: ${config.matchMode}`);
  const { 
      aggregationMode,
      keyCountColumn,
      discoverNewKeys,
      conditionalColumn,
      countBlanksInColumn,
      generateBlankDetails,
      blankCountingMode = 'rowAware',
      matchMode = 'whole',
      summaryTitleCell,
   } = config;

  const reportBaseNames = new Set([
    "update_report",
    "cross-sheet summary", // from summarySheetName default
    "key_mappings",
    "aggregation_report"
  ]);

  const sheetsToActuallySearch = sheetNamesToSearch.filter(name => {
      const lowerCaseName = name.toLowerCase();
      for (const baseName of reportBaseNames) {
          if (lowerCaseName.startsWith(baseName)) {
              console.warn(`Skipping sheet "${name}" as it appears to be a report sheet.`);
              return false;
          }
      }
      return true;
  });

  const uniqueReportingKeys = new Set(valueToKeyMap.values());
  const headerRowIndex = headerRow - 1;

  if (aggregationMode === 'keyMatch' && discoverNewKeys && keyCountColumn) {
      console.log('Discovering new keys from key match column:', keyCountColumn);
      sheetsToActuallySearch.forEach((sheetName, index) => {
        onProgress?.({
            stage: 'Discovering Keys',
            sheetName: sheetName,
            currentSheet: index + 1,
            totalSheets: sheetsToActuallySearch.length,
            currentTotals: {}
        });

        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return;

        const data: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
        if (data.length <= headerRow) return;

        const headers = data[headerRowIndex]?.map(h => String(h || '')) || [];
        
        const keyMatchColIdx = getColumnIndex(keyCountColumn, headers);
        if (keyMatchColIdx === null) return;

        const dataStartIndex = headerRow;
        for (let R = dataStartIndex; R < data.length; R++) {
          const row = data[R];
          if (!row || row.every(cell => cell === null)) continue;
          
          const keyCell = row[keyMatchColIdx];
          if (keyCell !== null && keyCell !== undefined) {
            const rawKeyText = String(keyCell).trim();
            if (rawKeyText) {
                uniqueReportingKeys.add(rawKeyText);
                const lowerKeyText = rawKeyText.toLowerCase();
                 if (!valueToKeyMap.has(lowerKeyText)) {
                    valueToKeyMap.set(lowerKeyText, rawKeyText);
                }
            }
          }
        }
      });
      console.log(`Discovery complete. Total unique keys now: ${uniqueReportingKeys.size}`);
  }


  const reportingKeys = [...uniqueReportingKeys].sort((a,b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
  const results: AggregationResult = {
    totalCounts: {},
    perSheetCounts: {},
    reportingKeys: reportingKeys,
    valueToKeyMap: valueToKeyMap,
    processedSheetNames: sheetsToActuallySearch,
    searchColumnIdentifiers: searchColumnIdentifiers,
    aggregationMode: aggregationMode,
    blankCountingMode: blankCountingMode,
    matchMode: matchMode,
    matchingRows: {},
    sheetKeyColumnIndices: {},
    sheetTitles: {},
  };
  
  if (countBlanksInColumn) {
      results.blankCounts = { total: 0, perSheet: {} };
      if (generateBlankDetails) {
        results.blankDetails = [];
      }
  }

  reportingKeys.forEach(key => { results.totalCounts[key] = 0; });

  sheetsToActuallySearch.forEach((sheetName, index) => {
    onProgress?.({ stage: 'Aggregating Data', sheetName, currentSheet: index + 1, totalSheets: sheetsToActuallySearch.length, currentTotals: results.totalCounts });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;

    if (summaryTitleCell && results.sheetTitles) {
        const titleCell = worksheet[summaryTitleCell];
        results.sheetTitles[sheetName] = titleCell?.v ? String(titleCell.v) : sheetName;
    }

    results.perSheetCounts[sheetName] = {};
    results.matchingRows![sheetName] = new Set<number>();
    reportingKeys.forEach(key => { results.perSheetCounts[sheetName][key] = 0; });
    if (results.blankCounts) { results.blankCounts.perSheet[sheetName] = 0; }

    const data: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    if (data.length <= headerRowIndex) return;

    const headers = data[headerRowIndex].map(h => String(h || ''));
    const dataStartIndex = headerRow;
    
    // Resolve blank column index per sheet, as headers might differ
    const sheetSpecificBlankColIdx = countBlanksInColumn ? getColumnIndex(countBlanksInColumn, headers) : null;
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    const lastRow = range.e.r; // 0-indexed last row

    if (sheetSpecificBlankColIdx !== null && blankCountingMode === 'fullColumn') {
        let blankCount = 0;
        for (let R = dataStartIndex; R <= lastRow; ++R) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: sheetSpecificBlankColIdx });
            const cell = worksheet[cellAddress];
            if (!cell || cell.v === null || cell.v === undefined || String(cell.v).trim() === '') {
                blankCount++;
            }
        }
        if (results.blankCounts) {
            results.blankCounts.perSheet[sheetName] = blankCount;
            results.blankCounts.total += blankCount;
            if (generateBlankDetails) {
                for (let R = dataStartIndex; R <= lastRow; ++R) {
                    const cellAddress = XLSX.utils.encode_cell({ r: R, c: sheetSpecificBlankColIdx });
                    const cell = worksheet[cellAddress];
                     if (!cell || cell.v === null || cell.v === undefined || String(cell.v).trim() === '') {
                        const rowData = headers.reduce((obj, header, i) => {
                            obj[header] = data[R]?.[i] ?? null;
                            return obj;
                        }, {} as Record<string, any>);
                        results.blankDetails!.push({
                            sheetName,
                            rowNumber: R + 1,
                            rowData: rowData,
                        });
                    }
                }
            }
        }
    }
    
    const searchColIndices = parseSourceColumns(searchColumnIdentifiers, headers);
    const keyMatchColIdx = aggregationMode === 'keyMatch' && keyCountColumn ? getColumnIndex(keyCountColumn, headers) : null;
    const conditionalColIdx = conditionalColumn ? getColumnIndex(conditionalColumn, headers) : null;
    
    if (keyMatchColIdx !== null && results.sheetKeyColumnIndices) {
        results.sheetKeyColumnIndices[sheetName] = keyMatchColIdx;
    }

    for (let R = dataStartIndex; R < data.length; R++) {
        const row = data[R];
        if (!row || row.every(cell => cell === null)) continue;
        
        let rowHasAnyMatch = false;

        // --- Blank counting logic --- (now only for rowAware)
        if (sheetSpecificBlankColIdx !== null && blankCountingMode === 'rowAware') {
            const cellValue = row[sheetSpecificBlankColIdx];
            const isRowOtherwiseEmpty = row.every((cell, index) => index === sheetSpecificBlankColIdx || cell === null || cell === undefined || String(cell).trim() === '');
            
            if ((cellValue === null || cellValue === undefined || String(cellValue).trim() === '') && !isRowOtherwiseEmpty) {
                if (results.blankCounts) {
                    results.blankCounts.total++;
                    results.blankCounts.perSheet[sheetName] = (results.blankCounts.perSheet[sheetName] || 0) + 1;
                }
                if (generateBlankDetails && results.blankDetails) {
                    const rowObject = headers.reduce((obj, header, i) => {
                        obj[header] = row[i] ?? null;
                        return obj;
                    }, {} as Record<string, any>);
                    results.blankDetails.push({
                        sheetName,
                        rowNumber: R + 1,
                        rowData: rowObject,
                    });
                }
            }
        }


        // --- Key or Value matching logic ---
        if (aggregationMode === 'keyMatch' && keyMatchColIdx !== null) {
            const keyCell = row[keyMatchColIdx];
            if (keyCell !== null && keyCell !== undefined) {
                const keyText = String(keyCell).trim();
                if (keyText) { // Ensure the key is not an empty string
                    const lowerKeyText = keyText.toLowerCase();
            
                    // Use the map to find the final reporting key. This map contains both user-defined and discovered keys.
                    if (valueToKeyMap.has(lowerKeyText)) {
                        const reportingKey = valueToKeyMap.get(lowerKeyText)!;
                        results.totalCounts[reportingKey] = (results.totalCounts[reportingKey] || 0) + 1;
                        results.perSheetCounts[sheetName][reportingKey] = (results.perSheetCounts[sheetName][reportingKey] || 0) + 1;
                        rowHasAnyMatch = true;
                    }
                }
            }
        } else if (aggregationMode === 'valueMatch') {
            if (conditionalColIdx !== null) {
                const conditionalCell = row[conditionalColIdx];
                const isConditionalCellEmpty = conditionalCell === null || conditionalCell === undefined || String(conditionalCell).trim() === '';
                if (!isConditionalCellEmpty) {
                    continue; // Skip keyword search for this row if the conditional column is not empty
                }
            }
            
            const rowMatchScores = new Map<string, number>();

            searchColIndices.forEach(colIdx => {
                const cellValue = row[colIdx];
                if (cellValue !== null && cellValue !== undefined) {
                    const cellText = String(cellValue);
                    
                    for (const [searchTerm, reportingKey] of valueToKeyMap.entries()) {
                        let score = 0;
                        let isMatch = false;

                        if (matchMode === 'loose') {
                            const searchWords = searchTerm.split(/\s+/).filter(Boolean);
                            if (searchWords.length > 0) {
                                const allWordsFound = searchWords.every(word => {
                                    const pattern = `(^|\\P{L})${escapeRegex(word)}(\\P{L}|$)`;
                                    const regex = new RegExp(pattern, 'iu');
                                    return regex.test(cellText);
                                });
                                if (allWordsFound) {
                                    isMatch = true;
                                    score = searchWords.length;
                                }
                            }
                        } else {
                            const escapedSearchTerm = escapeRegex(searchTerm);
                            const pattern = matchMode === 'whole'
                                ? `(^|\\P{L})${escapedSearchTerm}(\\P{L}|$)`
                                : escapedSearchTerm;
                            const searchRegex = new RegExp(pattern, 'giu');
                            
                            const matches = cellText.match(searchRegex);
                            if (matches) {
                                isMatch = true;
                                score = matches.length;
                            }
                        }

                        if (isMatch) {
                            const currentScore = rowMatchScores.get(reportingKey) || 0;
                            rowMatchScores.set(reportingKey, Math.max(currentScore, score));
                        }
                    }
                }
            });

            if (rowMatchScores.size > 0) {
                let winnerKey = '';
                let maxScore = -1;

                for (const [key, score] of rowMatchScores.entries()) {
                    if (score > maxScore) {
                        maxScore = score;
                        winnerKey = key;
                    } else if (score === maxScore) {
                        if (key < winnerKey) {
                            winnerKey = key;
                        }
                    }
                }
                
                if (winnerKey) {
                    results.totalCounts[winnerKey] = (results.totalCounts[winnerKey] || 0) + 1;
                    results.perSheetCounts[sheetName][winnerKey] = (results.perSheetCounts[sheetName][winnerKey] || 0) + 1;
                    rowHasAnyMatch = true;
                }
            }
        }
        
        if (rowHasAnyMatch) {
            results.matchingRows![sheetName].add(R);
        }
    }
    console.log(`Finished sheet "${sheetName}".`);
  });
  console.log('Aggregation complete.');
  return results;
}
