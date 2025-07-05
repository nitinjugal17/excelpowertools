
import * as XLSX from 'xlsx-js-style';
import { getColumnIndex, parseSourceColumns } from './excel-helpers';
import type { AiImputationContext, AiImputationSuggestion } from './excel-types';
import { imputeData } from '@/ai/flows/impute-data-flow';
import type { ImputeDataInput } from '@/ai/flows/impute-data-flow';


/**
 * Orchestrates the process of finding empty cells, gathering context, and calling an AI flow to get suggestions.
 * This function is optimized to process all empty cells in a single sheet with a single AI call.
 * @param file The Excel file to process.
 * @param sheetName The name of the sheet to process.
 * @param targetColumnIdentifier The column to find empty cells in and fill.
 * @param contextColumnIdentifiers Comma-separated list of columns to use for finding similar rows for context.
 * @param headerRowNumber The 1-indexed row number where headers are located.
 * @returns A promise that resolves to an array of AI suggestions.
 */
export async function getAiImputationSuggestions(
    file: File,
    sheetName: string,
    targetColumnIdentifier: string,
    contextColumnIdentifiers: string,
    headerRowNumber: number,
): Promise<AiImputationSuggestion[]> {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellDates: true });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return [];

    const contexts = findEmptyCellsAndGetContextForImputation(worksheet, targetColumnIdentifier, contextColumnIdentifiers, headerRowNumber);
    if (contexts.length === 0) return [];
    
    // Prepare a single batch request for the AI flow
    const batchInput: ImputeDataInput = {
        headers: contexts[0].headers,
        targetColumn: contexts[0].targetColumn,
        rowsToImpute: contexts.map(c => ({
            identifier: c.address, // Use cell address as the unique identifier
            rowData: c.rowData,
        })),
        // A simpler approach for now is to just take examples from the first context.
        exampleRows: contexts[0].exampleRows.slice(0, 10), // Limit examples to prevent overly large prompts
    };
    
    try {
        const batchResult = await imputeData(batchInput);
        
        // Map AI results back to the original contexts
        const suggestionsMap = new Map(batchResult.suggestions.map(s => [s.identifier, s.suggestion]));

        const finalSuggestions: AiImputationSuggestion[] = [];
        for (const context of contexts) {
            const suggestion = suggestionsMap.get(context.address);
            if (suggestion) {
                const contextDisplay = parseSourceColumns(contextColumnIdentifiers, context.headers)
                    .map(idx => context.headers[idx])
                    .map(header => ({
                        label: `Context (${header})`,
                        value: context.rowData[header]
                    }));

                 finalSuggestions.push({
                    sheetName: context.sheetName,
                    address: context.address,
                    row: context.row,
                    suggestion: suggestion,
                    rowData: context.rowData,
                    isChecked: true, // Default to checked
                    context: contextDisplay,
                });
            }
        }
        return finalSuggestions;

    } catch (error) {
        console.error(`AI batch imputation failed for sheet ${sheetName}:`, error);
        return []; // Return empty array on batch failure
    }
}


/**
 * Finds empty cells and gathers context for AI imputation.
 * @param worksheet The worksheet to process.
 * @param targetColumnIdentifier The identifier for the column to fill.
 * @param contextColumnIdentifiers The comma-separated identifiers for columns to use for context matching.
 * @param headerRowNumber The 1-indexed row number where headers are located.
 * @returns An array of context objects for the AI flow.
 */
function findEmptyCellsAndGetContextForImputation(
    worksheet: XLSX.WorkSheet,
    targetColumnIdentifier: string,
    contextColumnIdentifiers: string,
    headerRowNumber: number
): AiImputationContext[] {
    const contexts: AiImputationContext[] = [];
    const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    
    const headerRowIndex = headerRowNumber > 0 ? headerRowNumber - 1 : 0;
    if (aoa.length <= headerRowIndex) return [];
    
    const headers = aoa[headerRowIndex].map(String);
    const targetColIdx = getColumnIndex(targetColumnIdentifier, headers);
    const contextColIndices = parseSourceColumns(contextColumnIdentifiers, headers);

    if (targetColIdx === null || contextColIndices.length === 0) {
        console.error("Could not find target or context columns.");
        return [];
    }

    const contextColNames = contextColIndices.map(i => headers[i]);

    const dataAsObjects = aoa.slice(headerRowIndex + 1).map(row => {
        const obj: Record<string, any> = {};
        headers.forEach((header, index) => {
            obj[header] = row[index];
        });
        return obj;
    });

    for (let i = 0; i < dataAsObjects.length; i++) {
        const rowData = dataAsObjects[i];
        const targetValue = rowData[targetColumnIdentifier];
        const currentRowInAoa = i + headerRowIndex + 1;

        if (targetValue === null || targetValue === undefined || String(targetValue).trim() === '') {
            
            // Find similar rows for context (match on ALL context columns)
            const exampleRows = dataAsObjects.filter((otherRow, otherIndex) => {
                // Exclude the current row and rows where the target is also empty
                if (i === otherIndex || otherRow[targetColumnIdentifier] === null || otherRow[targetColumnIdentifier] === undefined) {
                    return false;
                }
                // Check if all context columns match
                return contextColNames.every(colName => otherRow[colName] === rowData[colName]);
            }).slice(0, 5); // Limit to 5 examples to keep prompt size reasonable

            contexts.push({
                sheetName: worksheet['!sheetName'] || 'Sheet',
                address: XLSX.utils.encode_cell({ r: currentRowInAoa, c: targetColIdx }),
                row: currentRowInAoa + 1, // 1-indexed for display
                rowData,
                headers,
                targetColumn: targetColumnIdentifier,
                exampleRows,
            });
        }
    }
    return contexts;
}


/**
 * Fills in empty cells based on the most common value (mode) from a group of duplicate rows.
 * This is an optimized version that avoids redundant loops.
 * @param file The Excel file.
 * @param sheetName The sheet to process.
 * @param headerRowNumber The 1-indexed header row.
 * @param config Configuration for the manual imputation.
 * @returns A promise resolving to an array of suggestions.
 */
export async function getManualImputationSuggestions(
    file: File,
    sheetName: string,
    headerRowNumber: number,
    config: {
        targetColumn: string;
        keyColumn: string;
        sourceColumns: string; // Now plural
        delimiter?: string;
        partToUse?: number;
    }
): Promise<AiImputationSuggestion[]> {
    const { targetColumn, keyColumn, sourceColumns, delimiter, partToUse } = config;

    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellDates: true });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return [];

    const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    const headerRowIndex = headerRowNumber > 0 ? headerRowNumber - 1 : 0;
    if (aoa.length <= headerRowIndex) return [];
    
    const headers = aoa[headerRowIndex].map(String);
    const targetColIdx = getColumnIndex(targetColumn, headers);
    const keyColIdx = getColumnIndex(keyColumn, headers);
    const sourceColIndices = parseSourceColumns(sourceColumns, headers);
    const sourceColNames = sourceColIndices.map(i => headers[i]);

    if (targetColIdx === null || keyColIdx === null || sourceColIndices.length === 0) {
        throw new Error(`Could not find one or more required columns on sheet "${sheetName}": Target (${targetColumn}), Key (${keyColumn}), or Source(s) (${sourceColumns}).`);
    }

    const dataAsObjects = aoa.map((row, index) => {
        const obj: Record<string, any> = { __originalRowIndex: index };
        headers.forEach((header, i) => { obj[header] = row[i]; });
        return obj;
    });

    const dataRows = dataAsObjects.slice(headerRowIndex + 1);

    // 1. First Pass: Group all possible source values by their key and count occurrences.
    const groupRawValues = new Map<string, { value: string, rowData: Record<string, any> }[]>();
    const keyCounts = new Map<string, number>();
    const keyFirstRow = new Map<string, number>();

    for (const row of dataRows) {
        const rawKey = row[keyColumn];
        if (rawKey === null || rawKey === undefined || String(rawKey).trim() === '' || String(rawKey).trim() === '0') {
            continue;
        }
        const key = String(rawKey).trim().toLowerCase();

        if (!groupRawValues.has(key)) {
            groupRawValues.set(key, []);
            keyFirstRow.set(key, row.__originalRowIndex + 1);
            keyCounts.set(key, 0);
        }
        
        keyCounts.set(key, keyCounts.get(key)! + 1);

        for (const colIdx of sourceColIndices) {
            const headerName = headers[colIdx];
            const val = row[headerName];
            if (val !== null && val !== undefined && String(val).trim() !== '') {
                groupRawValues.get(key)!.push({ value: String(val), rowData: row });
            }
        }
    }
    
    // 2. Second Pass: Calculate the most common value (mode) for each group *that has duplicates*.
    const groupModes = new Map<string, { suggestion: string, originalSource: string, sourceRowData?: Record<string, any> }>();
    for (const [key, rawValueObjects] of groupRawValues.entries()) {
        if ((keyCounts.get(key) || 0) <= 1) {
            continue;
        }

        if (rawValueObjects.length === 0) continue;

        const processedValueFrequencies: Record<string, number> = {};
        const processedToOriginalMap: Record<string, { original: string, rowData: Record<string, any> }> = {};

        rawValueObjects.forEach(({ value: originalValue, rowData }) => {
            let processedValue = originalValue;
            if (delimiter) {
                const parts = originalValue.split(delimiter);
                const indexToUse = partToUse === -1 ? parts.length - 1 : (partToUse ?? 1) - 1;
                processedValue = (indexToUse >= 0 && indexToUse < parts.length) ? parts[indexToUse].trim() : '';
            }
            
            if (processedValue) {
                processedValueFrequencies[processedValue] = (processedValueFrequencies[processedValue] || 0) + 1;
                if (!processedToOriginalMap[processedValue]) {
                    processedToOriginalMap[processedValue] = { original: originalValue, rowData };
                }
            }
        });

        if (Object.keys(processedValueFrequencies).length > 0) {
            const mode = Object.keys(processedValueFrequencies).reduce((a, b) => processedValueFrequencies[a] > processedValueFrequencies[b] ? a : b);
            const originalInfo = processedToOriginalMap[mode];
            groupModes.set(key, { 
                suggestion: mode, 
                originalSource: originalInfo.original,
                sourceRowData: originalInfo.rowData
            });
        }
    }


    // 3. Final Pass: Generate suggestions for rows that need imputation.
    const suggestions: AiImputationSuggestion[] = [];
    const rowsToImpute = dataRows.filter(row => {
        const val = row[targetColumn];
        return val === null || val === undefined || String(val).trim() === '';
    });
    
    for (const row of rowsToImpute) {
        const groupKey = String(row[keyColumn]).trim().toLowerCase();
        const modeInfo = groupModes.get(groupKey);
        const allGroupValueObjects = groupRawValues.get(groupKey);

        if (modeInfo) {
            const firstInstanceRow = keyFirstRow.get(groupKey);
            const keyLabel = firstInstanceRow ? `Key (${keyColumn}) (Original at R${firstInstanceRow})` : `Key (${keyColumn})`;

            const context: AiImputationSuggestion['context'] = [
                { label: keyLabel, value: row[keyColumn] }
            ];

            if (allGroupValueObjects && allGroupValueObjects.length > 0) {
                 const valueDetails = new Map<string, { count: number, rows: number[] }>();

                 allGroupValueObjects.forEach(valObj => {
                     const val = valObj.value;
                     const rowIndex = valObj.rowData.__originalRowIndex + 1;
                     if (!valueDetails.has(val)) {
                         valueDetails.set(val, { count: 0, rows: [] });
                     }
                     const details = valueDetails.get(val)!;
                     details.count++;
                     if (!details.rows.includes(rowIndex)) {
                        details.rows.push(rowIndex);
                     }
                 });
        
                 const uniqueValuesSummary = Array.from(valueDetails.entries())
                     .map(([val, details]) => `${val} (${details.count}x from R${details.rows.join(', R')})`)
                     .join('; ');
                 
                 context.push({ label: `All Source Values in Group`, value: uniqueValuesSummary });
            }

            if (modeInfo.sourceRowData) {
                const sourceRowData = modeInfo.sourceRowData;
                const sourceRowIndex = sourceRowData.__originalRowIndex + 1;
                const formattedData = Object.entries(sourceRowData)
                    .filter(([key, val]) => key !== '__originalRowIndex' && val !== null && val !== undefined && String(val).trim() !== '')
                    .map(([key, val]) => `${key}: "${String(val)}"`)
                    .join(' | ');
            
                if (formattedData) {
                    context.push({
                        label: `Data from Source Row (R${sourceRowIndex})`,
                        value: formattedData
                    });
                }
            }
            

            if (delimiter) {
                let contextValue = modeInfo.originalSource;
                if (modeInfo.sourceRowData) {
                    const sourceRowData = modeInfo.sourceRowData;
                    const sourceColumnsData = sourceColNames.map(name => {
                        const val = sourceRowData[name];
                        return val ? `${name}: "${String(val)}"` : null;
                    }).filter(Boolean).join(' | ');

                    if (sourceColumnsData) {
                        contextValue += ` (Source Data: ${sourceColumnsData})`;
                    }
                }
                context.push({ label: 'Source Value for Mode', value: contextValue });
                
                const parts = modeInfo.originalSource.split(delimiter);
                context.push({ label: `Split Result`, value: JSON.stringify(parts) });
                
                const indexToUse = partToUse === -1 ? parts.length - 1 : (partToUse ?? 1) - 1;
                const partIndexDisplay = partToUse === -1 ? `Last (${indexToUse+1})` : String(partToUse);
                
                context.push({ label: `Selected Part #${partIndexDisplay}`, value: `"${parts[indexToUse]?.trim() || ''}"` });
            }
            
            suggestions.push({
                sheetName: sheetName,
                address: XLSX.utils.encode_cell({ r: row.__originalRowIndex, c: targetColIdx }),
                row: row.__originalRowIndex + 1,
                suggestion: modeInfo.suggestion,
                isChecked: true,
                rowData: row,
                context
            });
        }
    }
    
    return suggestions;
}


/**
 * Fills in empty cells by concatenating values from other columns in the same row.
 * @param file The Excel file.
 * @param sheetName The sheet to process.
 * @param headerRowNumber The 1-indexed header row.
 * @param config Configuration for the concatenation.
 * @returns A promise resolving to an array of suggestions.
 */
export async function getConcatenationSuggestions(
    file: File,
    sheetName: string,
    headerRowNumber: number,
    config: {
        targetColumn: string;
        sourceColumns: string;
        separator: string;
    }
): Promise<AiImputationSuggestion[]> {
    const { targetColumn, sourceColumns, separator } = config;

    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellDates: true });
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return [];

    const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    const headerRowIndex = headerRowNumber > 0 ? headerRowNumber - 1 : 0;
    if (aoa.length <= headerRowIndex) return [];
    
    const headers = aoa[headerRowIndex].map(String);
    const targetColIdx = getColumnIndex(targetColumn, headers);
    const sourceColIndices = parseSourceColumns(sourceColumns, headers);

    if (targetColIdx === null || sourceColIndices.length === 0) {
        throw new Error(`Could not find Target Column (${targetColumn}) or Source Columns (${sourceColumns}) on sheet "${sheetName}".`);
    }

    const suggestions: AiImputationSuggestion[] = [];

    for (let R = headerRowIndex + 1; R < aoa.length; R++) {
        const row = aoa[R];
        if (!row) continue;

        const targetValue = row[targetColIdx];
        if (targetValue === null || targetValue === undefined || String(targetValue).trim() === '') {
            // This cell is empty, so we generate a suggestion.
             const valuesToConcatenate = sourceColIndices.map(colIdx => {
                const val = row[colIdx];
                return { name: headers[colIdx], value: (val !== null && val !== undefined) ? String(val) : '' };
            });

            const suggestion = valuesToConcatenate.map(v => v.value).filter(Boolean).join(separator);

            if (suggestion) { // Only add if the result is not empty
                 const rowData = headers.reduce((obj, header, i) => {
                    obj[header] = row[i];
                    return obj;
                }, {} as Record<string, any>);

                 const context = valuesToConcatenate
                    .filter(v => v.value) // only show sources that contributed
                    .map(v => ({ label: `Source (${v.name})`, value: v.value }));

                suggestions.push({
                    sheetName,
                    address: XLSX.utils.encode_cell({ r: R, c: targetColIdx }),
                    row: R + 1,
                    suggestion,
                    isChecked: true,
                    rowData,
                    context,
                });
            }
        }
    }
    
    return suggestions;
}


/**
 * Applies the approved suggestions to the workbook.
 * This function modifies the workbook in place.
 * @param workbook The workbook to modify.
 * @param suggestions An array of approved suggestions.
 * @returns The modified workbook.
 */
export function applyImputations(workbook: XLSX.WorkBook, suggestions: AiImputationSuggestion[]): XLSX.WorkBook {
    const suggestionsBySheet: Record<string, AiImputationSuggestion[]> = {};

    for (const s of suggestions) {
        if (!suggestionsBySheet[s.sheetName]) {
            suggestionsBySheet[s.sheetName] = [];
        }
        suggestionsBySheet[s.sheetName].push(s);
    }
    
    for (const sheetName in suggestionsBySheet) {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) continue;
        
        const sheetSuggestions = suggestionsBySheet[sheetName];
        sheetSuggestions.forEach(s => {
            const cellAddress = s.address;
            let cell = worksheet[cellAddress];
            if (!cell) {
                worksheet[cellAddress] = { t: 's' };
                cell = worksheet[cellAddress];
            }
            
            cell.v = s.suggestion;
            // Infer type
            if (!isNaN(Number(s.suggestion)) && s.suggestion.trim() !== '') {
                cell.t = 'n';
                cell.v = Number(s.suggestion);
            } else {
                cell.t = 's';
            }

            // Apply highlight to indicate change
            if (!cell.s) cell.s = {};
            // Preserve existing styles by merging
            cell.s.fill = {
                ...(cell.s.fill || {}),
                patternType: 'solid',
                fgColor: { rgb: 'C6EFCE' } // Light green
            };
        });
    }

    return workbook;
}
