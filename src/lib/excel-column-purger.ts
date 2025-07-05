
import * as XLSX from 'xlsx-js-style';
import { parseSourceColumns } from './excel-helpers';

/**
 * Removes specified columns from a worksheet, shifting remaining columns left and preserving styles.
 * @param workbook The workbook object. This object will be modified.
 * @param sheetNamesToUpdate An array of sheet names to process.
 * @param columnsToRemoveIdentifiers A comma-separated string of column identifiers (e.g., "A,C,Status").
 * @param headerRow The 1-indexed row number where headers are located.
 * @returns The modified workbook.
 */
export function purgeColumnsFromSheets(
    workbook: XLSX.WorkBook,
    sheetNamesToUpdate: string[],
    columnsToRemoveIdentifiers: string,
    headerRow: number
): XLSX.WorkBook {
    console.log(`Starting column purge for ${sheetNamesToUpdate.length} sheets. Columns to remove: "${columnsToRemoveIdentifiers}"`);
    for (const sheetName of sheetNamesToUpdate) {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
            console.warn(`Sheet "${sheetName}" not found. Skipping purge.`);
            continue;
        }

        console.log(`Purging columns from sheet: ${sheetName}`);
        const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
        const headers = aoa[headerRow - 1]?.map(h => String(h || '')) || [];
        const removeIndices = new Set(parseSourceColumns(columnsToRemoveIdentifiers, headers));

        if (removeIndices.size === 0) {
            console.log(`No matching columns to purge on sheet: ${sheetName}`);
            continue; // Nothing to do for this sheet
        }
        
        console.log(`Found ${removeIndices.size} columns to remove on sheet "${sheetName}":`, Array.from(removeIndices));

        const oldRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        const maxCols = oldRange.e.c + 1;

        // Create a mapping from old column index to new column index.
        // A value of -1 means the column is removed.
        const colMapping: number[] = [];
        let newColIdx = 0;
        for (let i = 0; i < maxCols; i++) {
            if (removeIndices.has(i)) {
                colMapping.push(-1);
            } else {
                colMapping.push(newColIdx++);
            }
        }
        
        const newSheet: XLSX.WorkSheet = { '!ref': 'A1' }; // Start with a minimal ref

        // Copy cells to their new positions
        for (const address in worksheet) {
            if (address[0] === '!') continue; // Handle special keys later

            const decoded = XLSX.utils.decode_cell(address);
            const newC = colMapping[decoded.c];

            if (newC !== -1) {
                const newAddr = XLSX.utils.encode_cell({ r: decoded.r, c: newC });
                newSheet[newAddr] = worksheet[address];
            }
        }
        
        // Update '!ref'
        const newMaxCol = oldRange.e.c - removeIndices.size;
        if (newMaxCol >= 0) {
            newSheet['!ref'] = XLSX.utils.encode_range({ 
                s: oldRange.s, 
                e: { r: oldRange.e.r, c: newMaxCol }
            });
        }

        // Update '!cols' (column widths)
        if (worksheet['!cols']) {
            newSheet['!cols'] = worksheet['!cols'].filter((_, i) => colMapping[i] !== -1);
        }

        // Update '!merges'
        if (worksheet['!merges']) {
            newSheet['!merges'] = [];
            for (const merge of worksheet['!merges']) {
                const newStartC = colMapping[merge.s.c];
                const newEndC = colMapping[merge.e.c];

                // The merge is only valid if both start and end columns are kept
                if (newStartC !== -1 && newEndC !== -1) {
                    newSheet['!merges'].push({
                        s: { r: merge.s.r, c: newStartC },
                        e: { r: merge.e.r, c: newEndC }
                    });
                }
            }
        }

        // Replace the old sheet with the new one
        workbook.Sheets[sheetName] = newSheet;
        console.log(`Finished purging on sheet: ${sheetName}`);
    }

    return workbook;
}
