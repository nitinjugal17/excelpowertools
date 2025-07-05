import * as XLSX from 'xlsx-js-style';
import type { AggregationResult, HeaderFormatOptions, MatchMode, SummaryConfig, UpdateResult } from './excel-types';
import { getColumnIndex, parseSourceColumns, sanitizeSheetNameForFormula, escapeRegex, parseColumnIdentifier } from './excel-helpers';
import { aggregateData } from './excel-aggregator-core';

/**
 * Helper function to apply formatting options to a cell style object.
 * @param cell The cell object to apply formatting to.
 * @param options The formatting options.
 */
export function applyFormatting(cell: XLSX.CellObject, options?: HeaderFormatOptions) {
    if (!options) return;
    if (!cell.s) cell.s = {};

    const font: XLSX.Font = cell.s.font || {};
    if (options.bold !== undefined) font.bold = options.bold;
    if (options.italic !== undefined) font.italic = options.italic;
    if (options.underline !== undefined) font.underline = options.underline;
    if (options.fontName) font.name = options.fontName;
    if (options.fontSize) font.sz = options.fontSize;
    if (options.fontColor) font.color = { rgb: options.fontColor.replace('#', '') };
    cell.s.font = font;

    const alignment: XLSX.Alignment = cell.s.alignment || {};
    if (options.horizontalAlignment) alignment.horizontal = options.horizontalAlignment;
    cell.s.alignment = alignment;

    const fill: XLSX.Fill = cell.s.fill || {};
    if (options.fillColor && options.fillColor.trim()) {
        fill.patternType = 'solid';
        fill.fgColor = { rgb: options.fillColor.replace('#', '') };
        cell.s.fill = fill;
    }
}

/**
 * Strips all formulas from the specified sheets within a workbook, replacing them with their last calculated values.
 * This function MODIFIES THE WORKBOOK IN PLACE.
 * @param workbook The workbook object to modify.
 * @param sheetNames The names of the sheets to strip formulas from.
 */
export function stripFormulasInWorkbook(workbook: XLSX.WorkBook, sheetNames: string[]): void {
  console.log('Stripping formulas from sheets:', sheetNames.join(', '));
  sheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet || !worksheet['!ref']) {
      console.warn(`Sheet "${sheetName}" not found or is empty, skipping formula stripping.`);
      return;
    }

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell_address = XLSX.utils.encode_cell({r:R, c:C});
            const cell = worksheet[cell_address];
            
            if (cell?.f) {
                delete cell.f;
                if (cell.t === 'e') {
                    cell.t = 's';
                }
            }
        }
    }
  });
}

export function markMatchingRows(
    workbook: XLSX.WorkBook,
    sheetsToMark: string[],
    rowsToMarkBySheet: { [sheetName: string]: Set<number> },
    columnToMarkIdentifier: string,
    valueToWrite: string,
    headerRow: number
) {
    console.log(`Marking matched rows in column "${columnToMarkIdentifier}" with value "${valueToWrite}".`);
    sheetsToMark.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return;

        const rowsToMark = rowsToMarkBySheet[sheetName];
        if (!rowsToMark || rowsToMark.size === 0) return;
        console.log(`Marking ${rowsToMark.size} rows on sheet "${sheetName}".`);

        const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
        const headerRowIndex = headerRow - 1;
        if (aoa.length <= headerRowIndex) return;
        
        const headers = aoa[headerRowIndex].map(h => String(h || ''));
        const markColIndex = getColumnIndex(columnToMarkIdentifier, headers);

        if (markColIndex === null) {
            console.warn(`Could not find column '${columnToMarkIdentifier}' to mark on sheet '${sheetName}'. Skipping.`);
            return;
        }

        rowsToMark.forEach(rowIndex => {
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: markColIndex });
            
            XLSX.utils.sheet_add_aoa(worksheet, [[valueToWrite]], { origin: cellAddress });

            const cell = worksheet[cellAddress];
            if (cell) {
              if (!cell.s) cell.s = {};
              cell.s.fill = {
                  patternType: 'solid',
                  fgColor: { rgb: 'FFFF00' }
              };
            }
        });
    });
}


export function lookupAndAndUpdate(
  workbook: XLSX.WorkBook,
  sheetNames: string[],
  searchColumnIdentifiers: string,
  updateColumnIdentifier: string,
  headerRow: number,
  valueToKeyMap: Map<string, string>,
  updateOnlyBlanks: boolean,
  matchMode: MatchMode,
  enablePairedRowValidation?: boolean,
  pairedValidationColumns?: string
): UpdateResult {
  console.log(`Starting lookupAndAndUpdate. Update column: "${updateColumnIdentifier}". Update only blanks: ${updateOnlyBlanks}.`);
  const result: UpdateResult = {
    summary: { totalCellsUpdated: 0, sheetsUpdated: [] },
    details: [],
  };

  const sheetsUpdated = new Set<string>();

  sheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;
    console.log(`Processing sheet for update: "${sheetName}"`);

    const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    const headerRowIndex = headerRow - 1;
    if (aoa.length <= headerRowIndex) return;

    const headers = aoa[headerRowIndex].map(h => String(h || ''));
    const searchColIndices = parseSourceColumns(searchColumnIdentifiers, headers);
    const updateColIndex = getColumnIndex(updateColumnIdentifier, headers);
    const dataStartIndex = headerRow;

    if (updateColIndex === null) {
        console.warn(`Update column "${updateColumnIdentifier}" not found on sheet "${sheetName}". Skipping sheet.`);
        return;
    }

    for (let R = dataStartIndex; R < aoa.length; R++) {
      const row = aoa[R];
      if (!row || row.every(c => c === null || c === undefined)) continue;
      
      const rowMatchScores = new Map<string, number>();
      const triggerInfo = new Map<string, { value: string, column: string }>();

      for (const colIdx of searchColIndices) {
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
                  if (!triggerInfo.has(reportingKey)) {
                      triggerInfo.set(reportingKey, { value: cellText, column: headers[colIdx] });
                  }
              }
          }
        }
      }

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
        
        const updateCell = row[updateColIndex];
        const isUpdateCellBlank = updateCell === null || updateCell === undefined || String(updateCell).trim() === '';

        if (updateOnlyBlanks && !isUpdateCellBlank) {
            continue;
        }

        let validationPassed = true;
        if (enablePairedRowValidation && pairedValidationColumns) {
            validationPassed = false;
            const validationColIndices = parseSourceColumns(pairedValidationColumns, headers);
            
            if (validationColIndices.length > 0) {
                const rowAbove = (R > dataStartIndex) ? aoa[R - 1] : null;
                const rowBelow = (R < aoa.length - 1) ? aoa[R + 1] : null;

                if (rowAbove) {
                    const aboveMatches = validationColIndices.every(idx => {
                        return String(row[idx] ?? '').trim().toLowerCase() === String(rowAbove[idx] ?? '').trim().toLowerCase();
                    });
                    if (aboveMatches) validationPassed = true;
                }
                if (!validationPassed && rowBelow) {
                    const belowMatches = validationColIndices.every(idx => {
                        return String(row[idx] ?? '').trim().toLowerCase() === String(rowBelow[idx] ?? '').trim().toLowerCase();
                    });
                    if (belowMatches) validationPassed = true;
                }
            } else {
              validationPassed = true;
            }
        }

        if (!validationPassed) continue;

        const cellAddress = XLSX.utils.encode_cell({ r: R, c: updateColIndex });
        const originalValue = worksheet[cellAddress]?.v;
        console.log(`Updating cell ${sheetName}!${cellAddress}: Old='${originalValue}', New='${winnerKey}'`);
        
        XLSX.utils.sheet_add_aoa(worksheet, [[winnerKey]], { origin: cellAddress });

        const cell = worksheet[cellAddress];
        if (cell) {
          if (!cell.s) cell.s = {};
          cell.s.fill = {
              ...(cell.s.fill || {}),
              patternType: 'solid',
              fgColor: { rgb: 'FFFF00' } // Yellow
          };
        }
        
        result.summary.totalCellsUpdated++;
        sheetsUpdated.add(sheetName);

        const { value: triggerValue, column: triggerColumn } = triggerInfo.get(winnerKey) || { value: '', column: '' };
        result.details.push({
            sheetName: sheetName,
            rowNumber: R + 1,
            cellAddress: cellAddress,
            originalValue: originalValue,
            newValue: winnerKey,
            keyUsed: winnerKey,
            triggerValue: triggerValue || '',
            triggerColumn: triggerColumn || '',
            rowData: headers.reduce((obj, header, i) => {
              obj[header] = row[i] ?? null;
              return obj;
            }, {} as Record<string, any>)
        });
      }
    }
  });

  result.summary.sheetsUpdated = Array.from(sheetsUpdated);
  console.log(`Update complete. Updated ${result.summary.totalCellsUpdated} cells across ${result.summary.sheetsUpdated.length} sheets.`);
  return result;
}

export function findPotentialUpdates(
  workbook: XLSX.WorkBook,
  sheetNames: string[],
  searchColumnIdentifiers: string,
  updateColumnIdentifier: string,
  headerRow: number,
  valueToKeyMap: Map<string, string>,
  updateOnlyBlanks: boolean,
  matchMode: MatchMode,
  enablePairedRowValidation?: boolean,
  pairedValidationColumns?: string
): UpdateResult {
  console.log(`Finding potential updates for column "${updateColumnIdentifier}".`);
  const result: UpdateResult = {
    summary: { totalCellsUpdated: 0, sheetsUpdated: [] },
    details: [],
  };

  const sheetsUpdated = new Set<string>();

  sheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;

    const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    const headerRowIndex = headerRow - 1;
    if (aoa.length <= headerRowIndex) return;

    const headers = aoa[headerRowIndex].map(h => String(h || ''));
    const searchColIndices = parseSourceColumns(searchColumnIdentifiers, headers);
    const updateColIndex = getColumnIndex(updateColumnIdentifier, headers);
    const dataStartIndex = headerRow;

    if (updateColIndex === null) {
        console.warn(`Update column "${updateColumnIdentifier}" not found on sheet "${sheetName}". Skipping sheet.`);
        return;
    }

    for (let R = dataStartIndex; R < aoa.length; R++) {
      const row = aoa[R];
      if (!row || row.every(c => c === null || c === undefined)) continue;

      const rowMatchScores = new Map<string, number>();
      const triggerInfo = new Map<string, { value: string, column: string }>();

      for (const colIdx of searchColIndices) {
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
              rowMatchScores.set(reportingKey, (rowMatchScores.get(reportingKey) || 0) + score);
              if (!triggerInfo.has(reportingKey)) {
                  triggerInfo.set(reportingKey, { value: cellText, column: headers[colIdx] });
              }
            }
          }
        }
      }

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

        const updateCell = row[updateColIndex];
        const isUpdateCellBlank = updateCell === null || updateCell === undefined || String(updateCell).trim() === '';
        
        if (updateOnlyBlanks && !isUpdateCellBlank) {
            continue;
        }

        let validationPassed = true;
        if (enablePairedRowValidation && pairedValidationColumns) {
            validationPassed = false;
            const validationColIndices = parseSourceColumns(pairedValidationColumns, headers);
            
            if (validationColIndices.length > 0) {
                const rowAbove = (R > dataStartIndex) ? aoa[R - 1] : null;
                const rowBelow = (R < aoa.length - 1) ? aoa[R + 1] : null;

                if (rowAbove) {
                    const aboveMatches = validationColIndices.every(idx => {
                        return String(row[idx] ?? '').trim().toLowerCase() === String(rowAbove[idx] ?? '').trim().toLowerCase();
                    });
                    if (aboveMatches) validationPassed = true;
                }
                if (!validationPassed && rowBelow) {
                    const belowMatches = validationColIndices.every(idx => {
                        return String(row[idx] ?? '').trim().toLowerCase() === String(rowBelow[idx] ?? '').trim().toLowerCase();
                    });
                    if (belowMatches) validationPassed = true;
                }
            } else {
              validationPassed = true;
            }
        }

        if (!validationPassed) continue;

        const { value: triggerValue, column: triggerColumn } = triggerInfo.get(winnerKey) || { value: '', column: '' };
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: updateColIndex });
        const originalValue = worksheet[cellAddress]?.v;
        result.summary.totalCellsUpdated++;
        sheetsUpdated.add(sheetName);
        result.details.push({
            sheetName: sheetName,
            rowNumber: R + 1,
            cellAddress: cellAddress,
            originalValue: originalValue,
            newValue: winnerKey,
            keyUsed: winnerKey,
            triggerValue: triggerValue || '',
            triggerColumn: triggerColumn || '',
            rowData: headers.reduce((obj, header, i) => {
              obj[header] = row[i] ?? null;
              return obj;
            }, {} as Record<string, any>)
        });
      }
    }
  });

  result.summary.sheetsUpdated = Array.from(sheetsUpdated);
  console.log(`Found ${result.summary.totalCellsUpdated} potential updates.`);
  return result;
}

export function fillEmptyKeyColumn(
  workbook: XLSX.WorkBook,
  sheetNames: string[],
  searchColumnIdentifiers: string,
  keyColumnIdentifier: string,
  headerRow: number,
  valueToKeyMap: Map<string, string>,
  matchMode: MatchMode
) {
  console.log(`Filling empty key column "${keyColumnIdentifier}" based on search in "${searchColumnIdentifiers}".`);
  sheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;
    console.log(`Processing sheet for key fill: "${sheetName}"`);

    const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    const headerRowIndex = headerRow - 1;
    if (aoa.length <= headerRowIndex) return;
    
    const headers = aoa[headerRowIndex].map(h => String(h || ''));
    const searchColIndices = parseSourceColumns(searchColumnIdentifiers, headers);
    const keyColIndex = getColumnIndex(keyColumnIdentifier, headers);

    if (keyColIndex === null) {
        console.warn(`Key column "${keyColumnIdentifier}" not found on sheet "${sheetName}". Skipping sheet.`);
        return;
    }

    for (let R = headerRow; R < aoa.length; R++) {
      const row = aoa[R];
      if (!row) continue;

      const keyCell = row[keyColIndex];
      const isKeyCellEmpty = keyCell === null || keyCell === undefined || String(keyCell).trim() === '';

      if (isKeyCellEmpty) {
        const rowMatchScores = new Map<string, number>();

        for (const colIdx of searchColIndices) {
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
        }

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
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: keyColIndex });
                XLSX.utils.sheet_add_aoa(worksheet, [[winnerKey]], { origin: cellAddress });
            }
        }
      }
    }
  });
}

function _findAndClearExistingSummary(worksheet: XLSX.WorkSheet, summaryTitle: string) {
    if (!worksheet || !worksheet['!ref']) return;

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    let summaryStart: { r: number, c: number } | null = null;

    for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cellAddr];
            if (cell && cell.v === summaryTitle) {
                summaryStart = { r: R, c: C };
                break;
            }
        }
        if (summaryStart) break;
    }

    if (!summaryStart) {
        console.log(`No existing summary titled "${summaryTitle}" found to clear.`);
        return;
    }

    let endRow = summaryStart.r;
    for (let R = summaryStart.r + 1; R <= range.e.r + 1; R++) {
        const cellAddr = XLSX.utils.encode_cell({ r: R, c: summaryStart.c });
        const cell = worksheet[cellAddr];
        
        if (cell && typeof cell.v === 'string' && cell.v.toLowerCase() === 'total') {
            endRow = R;
            break;
        }
        if (!cell || cell.v === null || cell.v === undefined || String(cell.v).trim() === '') {
             endRow = R - 1;
             break;
        }
        endRow = R;
    }

    console.log(`Clearing existing summary from ${XLSX.utils.encode_cell(summaryStart)} to ${XLSX.utils.encode_cell({ r: endRow, c: summaryStart.c + 1 })}`);
    for (let R = summaryStart.r; R <= endRow; R++) {
        for (let C = summaryStart.c; C <= summaryStart.c + 1; C++) {
            const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
            if (worksheet[cellAddr]) {
                delete worksheet[cellAddr];
            }
        }
    }

    if (worksheet['!merges']) {
        worksheet['!merges'] = worksheet['!merges'].filter(merge => 
            !(merge.s.r === summaryStart!.r && merge.s.c === summaryStart!.c)
        );
    }
}

function _generateInSheetSummaryAOA(
    results: AggregationResult,
    workbook: XLSX.WorkBook,
    sheetName: string,
    config: SummaryConfig,
    options: { dataSource: 'reportingScope' | 'localSheet', generationMode: 'static' | 'formula', showOnlyLocalKeys: boolean },
    editedKeyToOriginalsMap?: Map<string, string[]>
): any[][] {
    const dataToInsert: any[][] = [];
    const { dataSource, generationMode, showOnlyLocalKeys } = options;
    const { headerRow, keyCountColumn, aggregationMode, countBlanksInColumn, blankCountLabel, showBlanksInInSheetSummary, totalRowFormatting, blankRowFormatting, headerFormatting } = config;

    const dynamicTitle = results.sheetTitles?.[sheetName] || config.inSheetSummaryTitle || 'Summary';
    const headerCell: XLSX.CellObject = { v: dynamicTitle, t: 's' };
    applyFormatting(headerCell, headerFormatting);
    const countHeaderCell: XLSX.CellObject = { v: 'Count', t: 's' };
    applyFormatting(countHeaderCell, headerFormatting);
    dataToInsert.push([headerCell, countHeaderCell]);
    
    const getKeysForSheet = () => {
        const allKeysInReport = new Set(results.reportingKeys);
        if (results.blankCounts && results.blankCounts.total > 0 && blankCountLabel) {
            allKeysInReport.add(blankCountLabel);
        }

        if (dataSource === 'reportingScope' || !showOnlyLocalKeys) {
             return Array.from(allKeysInReport);
        }

        const localKeys = new Set<string>();
        const sheetCounts = results.perSheetCounts[sheetName] || {};
        for (const key in sheetCounts) {
            if (sheetCounts[key] > 0) localKeys.add(key);
        }
        if ((results.blankCounts?.perSheet[sheetName] || 0) > 0 && showBlanksInInSheetSummary && blankCountLabel) {
            localKeys.add(blankCountLabel);
        }
        return Array.from(localKeys);
    };

    const keysToDisplay = getKeysForSheet().sort((a,b)=>a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
    
    let dataRowsAdded = 0;
    keysToDisplay.forEach((key) => {
        let valueCell: XLSX.CellObject;
        
        if (generationMode === 'formula' && aggregationMode === 'keyMatch' && keyCountColumn) {
            const keyColIdx = results.sheetKeyColumnIndices?.[sheetName];
            if (keyColIdx !== null && keyColIdx !== undefined) {
                const keyColLetter = XLSX.utils.encode_col(keyColIdx);
                
                const worksheet = results.processedSheetNames.includes(sheetName) ? workbook.Sheets[sheetName] : undefined;
                const range = worksheet ? XLSX.utils.decode_range(worksheet['!ref'] || 'A1') : { e: { r: headerRow } };
                const lastRow = range.e.r + 1;
                const dataStartRowForFormula = headerRow + 1;
                const formulaRange = `${sanitizeSheetNameForFormula(sheetName)}!${keyColLetter}${dataStartRowForFormula}:${keyColLetter}${lastRow}`;

                if (key === blankCountLabel) {
                    const blankColForFormulaIdx = countBlanksInColumn ? getColumnIndex(countBlanksInColumn, Object.keys(results.perSheetCounts[sheetName])) : keyColIdx;
                    if(blankColForFormulaIdx !== null) {
                         const blankColLetter = XLSX.utils.encode_col(blankColForFormulaIdx);
                         const blankFormulaRange = `${sanitizeSheetNameForFormula(sheetName)}!${blankColLetter}${dataStartRowForFormula}:${blankColLetter}${lastRow}`;
                         valueCell = { t: 'n', f: `COUNTBLANK(${blankFormulaRange})` };
                    } else {
                         valueCell = { t: 'n', v: results.blankCounts?.perSheet[sheetName] || 0 };
                    }
                } else {
                    const originalKeys = editedKeyToOriginalsMap?.get(key) || [key];
                    const criteria = "{" + originalKeys.map(k => `"${k.replace(/"/g, '""')}"`).join(",") + "}";
                    valueCell = { t: 'n', f: `SUMPRODUCT(--ISNUMBER(MATCH(${formulaRange},${criteria},0)))` };
                }
            } else {
                valueCell = { t: 'n', v: results.perSheetCounts[sheetName]?.[key] || 0 };
            }
        } else {
            const sheetCounts = results.perSheetCounts[sheetName] || {};
            const blankCount = results.blankCounts?.perSheet[sheetName] || 0;
            const staticCount = sheetCounts[key] || (key === blankCountLabel ? blankCount : 0);
            
            valueCell = { t: 'n', v: staticCount };
        }
        
        if (key === blankCountLabel && !showBlanksInInSheetSummary) {
            return;
        }

        dataRowsAdded++;
        const linkedValueCell: XLSX.CellObject = valueCell;
        applyFormatting(linkedValueCell, { horizontalAlignment: 'right' });

        const keyCell: XLSX.CellObject = { v: key, t: 's' };
        applyFormatting(keyCell, headerFormatting?.horizontalAlignment ? { horizontalAlignment: headerFormatting.horizontalAlignment } : undefined);

        if (key === blankCountLabel) {
            applyFormatting(keyCell, blankRowFormatting);
            applyFormatting(linkedValueCell, blankRowFormatting);
        }

        dataToInsert.push([keyCell, linkedValueCell]);
    });
    
    if (dataRowsAdded > 0 && totalRowFormatting) {
        const insertColLetter = XLSX.utils.encode_col(parseColumnIdentifier(config.insertColumn || "A")! + 1);
        const totalFormula = `SUM(${insertColLetter}${config.insertStartRow! + 1}:${insertColLetter}${config.insertStartRow! + dataToInsert.length -1})`;
        const totalValueCell: XLSX.CellObject = { t: 'n', f: totalFormula };
        
        const totalLabelCell: XLSX.CellObject = {v:'Total', t: 's'};
        applyFormatting(totalLabelCell, totalRowFormatting);
        applyFormatting(totalValueCell, totalRowFormatting);
        dataToInsert.push([totalLabelCell, totalValueCell]);
    }
    
    return dataToInsert;
}

export function insertAggregationResultsIntoSheets(
    workbook: XLSX.WorkBook,
    results: AggregationResult,
    sheetsToInsertInto: string[],
    headerRow: number,
    config: SummaryConfig,
    options: { dataSource: 'reportingScope' | 'localSheet', generationMode: 'static' | 'formula', showOnlyLocalKeys: boolean },
    editedKeyToOriginalsMap?: Map<string, string[]>
) {
    console.log(`Starting insertAggregationResultsIntoSheets for ${sheetsToInsertInto.length} sheets.`);
    const insertColIdx = config.insertColumn ? parseColumnIdentifier(config.insertColumn) : null;
    if (insertColIdx === null) {
        console.error('Invalid insert column for summary.', config.insertColumn);
        return;
    }

    const insertStartRow = config.insertStartRow || 1;

    for (const sheetName of sheetsToInsertInto) {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) continue;
        
        const dynamicTitle = results.sheetTitles?.[sheetName] || config.inSheetSummaryTitle || 'Summary';
        
        if (config.clearExistingInSheetSummary) {
            _findAndClearExistingSummary(worksheet, dynamicTitle);
        }

        console.log(`Inserting summary into sheet: "${sheetName}" at ${config.insertColumn}${insertStartRow}`);
        
        const dataToInsert = _generateInSheetSummaryAOA(results, workbook, sheetName, config, options, editedKeyToOriginalsMap);
        
        if (dataToInsert.length > 1) { // Only insert if there's more than just the header
            XLSX.utils.sheet_add_aoa(worksheet, dataToInsert, { origin: { r: insertStartRow - 1, c: insertColIdx } });
            
            // Auto-fit new columns
            const keyColMaxWidth = dataToInsert.reduce((w, r) => Math.max(w, String(r[0]?.v || '').length), 0);
            const valColMaxWidth = dataToInsert.reduce((w, r) => Math.max(w, String(r[1]?.v || r[1]?.f || '').length), 0);
            if (!worksheet['!cols']) worksheet['!cols'] = [];
            worksheet['!cols'][insertColIdx] = { wch: Math.max(15, keyColMaxWidth + 2) };
            worksheet['!cols'][insertColIdx + 1] = { wch: Math.max(10, valColMaxWidth + 2) };
        }
    }
    console.log("Finished insertAggregationResultsIntoSheets.");
}
