
import * as XLSX from 'xlsx-js-style';
import { parseSourceColumns } from './excel-helpers';

/**
 * Merges sheets from a source workbook into a destination workbook.
 * @param sourceWb The source XLSX.WorkBook object.
 * @param destWb The destination XLSX.WorkBook object. This object will be modified.
 * @param sheetsToMerge An array of sheet names to copy from source to destination.
 * @param replaceIfExists If true, sheets with the same name in the destination will be overwritten.
 * @returns The modified destination workbook.
 */
export function mergeSheets(
    sourceWb: XLSX.WorkBook,
    destWb: XLSX.WorkBook,
    sheetsToMerge: string[],
    replaceIfExists: boolean
): XLSX.WorkBook {
    console.log(`Starting sheet merge operation. Sheets to merge: ${sheetsToMerge.join(', ')}. Replace existing: ${replaceIfExists}`);
    sheetsToMerge.forEach(sheetName => {
        const sourceSheet = sourceWb.Sheets[sheetName];
        if (!sourceSheet) {
            console.warn(`Sheet "${sheetName}" not found in the source workbook, skipping.`);
            return;
        }

        const sheetExistsInDest = destWb.SheetNames.some(name => name.toLowerCase() === sheetName.toLowerCase());
        const actualDestSheetName = destWb.SheetNames.find(name => name.toLowerCase() === sheetName.toLowerCase()) || sheetName;

        if (sheetExistsInDest) {
            if (replaceIfExists) {
                console.log(`Replacing existing sheet "${actualDestSheetName}" in destination workbook.`);
                destWb.Sheets[actualDestSheetName] = sourceSheet;
            } else {
                console.log(`Sheet "${actualDestSheetName}" already exists in destination and replace is false. Skipping.`);
            }
        } else {
            console.log(`Appending new sheet "${sheetName}" to destination workbook.`);
            XLSX.utils.book_append_sheet(destWb, sourceSheet, sheetName);
        }
    });
    console.log('Sheet merge operation complete.');
    return destWb;
}


/**
 * Combines data from multiple sheets within a single workbook into one new sheet.
 * Handles disparate headers by creating a superset of all columns.
 * @param workbook The source XLSX.WorkBook object.
 * @param sheetsToCombine An array of sheet names to combine.
 * @param headerRow The 1-indexed row number containing headers.
 * @param newSheetName The name for the new combined sheet.
 * @param addSourceColumn If true, prepends a column with the source sheet name for each row.
 * @param columnsToIgnore Optional comma-separated string of columns to exclude.
 * @returns A new XLSX.WorkBook object containing only the new combined sheet.
 */
export function combineSheets(
    workbook: XLSX.WorkBook,
    sheetsToCombine: string[],
    headerRow: number,
    newSheetName: string,
    addSourceColumn: boolean,
    columnsToIgnore?: string
): XLSX.WorkBook {
    console.log(`Starting combine operation for ${sheetsToCombine.length} sheets into a new sheet named "${newSheetName}".`);
    const combinedData: any[][] = [];
    const headerRowIndex = headerRow - 1;

    // First pass: Collect a superset of all unique headers from all selected sheets.
    console.log('Pass 1: Collecting all unique headers.');
    const allHeadersSet = new Set<string>();
    sheetsToCombine.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return;
        const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
        if (aoa.length > headerRowIndex) {
            const headers = aoa[headerRowIndex].map(h => String(h || '').trim());
            headers.forEach(h => {
                if (h) allHeadersSet.add(h);
            });
        }
    });

    // Filter out ignored columns
    const uniqueHeadersArray = Array.from(allHeadersSet);
    let combinedHeaders = uniqueHeadersArray;

    if (columnsToIgnore) {
        console.log(`Ignoring columns: ${columnsToIgnore}`);
        const ignoredIndices = new Set(parseSourceColumns(columnsToIgnore, uniqueHeadersArray));
        const ignoredHeaders = new Set(Array.from(ignoredIndices).map(idx => uniqueHeadersArray[idx]));
        combinedHeaders = uniqueHeadersArray.filter(h => !ignoredHeaders.has(h));
        console.log('Final headers after filtering:', combinedHeaders);
    }
    
    if (addSourceColumn) {
        combinedHeaders.unshift("Source Sheet");
    }
    console.log(`Found ${combinedHeaders.length} unique headers to use.`, combinedHeaders);
    
    combinedData.push(combinedHeaders);

    // Second pass: Iterate through sheets again and map data to the combined headers.
    console.log('Pass 2: Mapping data from each sheet to the combined header structure.');
    sheetsToCombine.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return;
        console.log(`Combining data from sheet: "${sheetName}"`);

        const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
        if (aoa.length <= headerRowIndex) return;

        const sheetHeaders = aoa[headerRowIndex].map(h => String(h || '').trim());
        const dataRows = aoa.slice(headerRow);

        dataRows.forEach(row => {
            if (row.every(cell => cell === null)) return; // Skip completely empty rows
            
            const newRow = new Array(combinedHeaders.length).fill(null);
            
            if (addSourceColumn) {
                newRow[0] = sheetName;
            }

            sheetHeaders.forEach((header, localIndex) => {
                if (!header) return; // Skip empty header cells
                const combinedIndex = combinedHeaders.indexOf(header);
                if (combinedIndex !== -1) {
                    newRow[combinedIndex] = row[localIndex];
                }
            });
            combinedData.push(newRow);
        });
    });
    console.log(`Finished combining. Total rows in new sheet: ${combinedData.length}.`);

    const newWs = XLSX.utils.aoa_to_sheet(combinedData, {cellDates: true});
    
    // Auto-fit columns
    if (combinedData.length > 0) {
        const colWidths = combinedHeaders.map(h => ({ wch: Math.max(15, h.length + 2) }));
        newWs['!cols'] = colWidths;
    }
    
    if (combinedData.length > 1) {
        newWs['!autofilter'] = { ref: newWs['!ref']! };
    }
    
    // Create a new workbook with only the combined sheet
    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, newWs, newSheetName);

    return newWb;
}
