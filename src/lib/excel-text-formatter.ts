
import * as XLSX from 'xlsx-js-style';
import type { TextFormatConfig } from './excel-types';
import { parseColumnIdentifier } from './excel-helpers';


/**
 * Finds cells containing specific text or matching a regex and applies formatting to them.
 * This function MODIFIES THE WORKBOOK IN PLACE while preserving existing cell styles.
 * @param workbook The workbook to process.
 * @param sheetNames Array of sheet names to process.
 * @param config Configuration for the text search and formatting.
 * @param onProgress Optional callback for progress reporting.
 * @param cancellationRequestedRef Optional ref to check for cancellation requests.
 * @returns An object containing the modified workbook and the count of formatted cells.
 */
export function findAndFormatText(
  workbook: XLSX.WorkBook,
  sheetNames: string[],
  config: TextFormatConfig,
  onProgress?: (status: { sheetName: string; currentSheet: number; totalSheets: number; cellsFormatted: number }) => void,
  cancellationRequestedRef?: React.RefObject<boolean>
): { workbook: XLSX.WorkBook; cellsFormatted: number } {
  let cellsFormatted = 0;
  const { searchText: searchTerms, searchMode, matchCase, matchEntireCell, style, range: searchRange } = config;
  const hasSearchTerms = searchTerms && searchTerms.length > 0;

  if (!style.font && !style.fill && !style.alignment) {
      return { workbook, cellsFormatted };
  }

  let searchRegexes: RegExp[] = [];
  if (hasSearchTerms && searchMode === 'regex') {
    searchRegexes = searchTerms.map(term => {
      try {
        const pattern = matchEntireCell ? `^${term}$` : term;
        const flags = matchCase ? '' : 'i';
        return new RegExp(pattern, flags);
      } catch (e) {
        console.error(`Invalid regex pattern: ${term}`, e);
        return new RegExp('(?!)');
      }
    });
  }

  const textTargets = hasSearchTerms && searchMode === 'text' 
    ? (matchCase ? searchTerms : searchTerms.map(t => t.toLowerCase())) 
    : [];

  for (let i = 0; i < sheetNames.length; i++) {
    const sheetName = sheetNames[i];
    if (cancellationRequestedRef?.current) throw new Error('Cancelled by user.');

    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet || !worksheet['!ref']) continue;
    
    onProgress?.({ sheetName, currentSheet: i + 1, totalSheets: sheetNames.length, cellsFormatted });
    
    const fullRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');

    // Default to the full range of the sheet
    let startR = fullRange.s.r;
    let endR = fullRange.e.r;
    let startC = fullRange.s.c;
    let endC = fullRange.e.c;

    // If a custom range is specified, try to apply it.
    if (searchRange) {
        // Apply row limits, clamping to sheet boundaries.
        startR = Math.max(fullRange.s.r, searchRange.startRow - 1);
        endR = Math.min(fullRange.e.r, searchRange.endRow - 1);
        
        // Attempt to parse user-defined columns.
        const userStartC = parseColumnIdentifier(searchRange.startCol);
        const userEndC = parseColumnIdentifier(searchRange.endCol);
        
        // Only apply column limits if both start and end are validly parsed.
        // This prevents falling back to full sheet width if one is empty/invalid.
        if (userStartC !== null && userEndC !== null) {
            startC = Math.max(fullRange.s.c, userStartC);
            endC = Math.min(fullRange.e.c, userEndC);
        } else {
            // If the user enabled the range but provided invalid columns, 
            // we should not format anything to prevent accidental full-width formatting.
            // We can do this by ensuring the column loop will not run.
            startC = 1;
            endC = 0;
        }
    }

    // Final check to ensure the calculated range is valid before proceeding.
    if (startR > endR || startC > endC) {
        continue; // Skip to the next sheet if the range is invalid.
    }

    for (let R = startR; R <= endR; ++R) {
      for (let C = startC; C <= endC; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = worksheet[cellAddress];

        let isMatch = false;

        if (hasSearchTerms) {
            if (cell && cell.v !== null && cell.v !== undefined) {
              const cellValue = String(cell.v);

              if (searchMode === 'regex') {
                isMatch = searchRegexes.some(regex => regex.test(cellValue));
              } else { // Text mode
                const valueToCompare = matchCase ? cellValue : cellValue.toLowerCase();
                if (matchEntireCell) {
                  isMatch = textTargets.includes(valueToCompare);
                } else {
                  isMatch = textTargets.some(target => valueToCompare.includes(target));
                }
              }
            }
        } else {
            isMatch = true;
        }
          
        if (isMatch) {
            if (!worksheet[cellAddress]) worksheet[cellAddress] = { t: 'z' };
            const targetCell = worksheet[cellAddress];

            if (!targetCell.s) targetCell.s = {};
            
            if (style.font) {
                if (!targetCell.s.font) targetCell.s.font = {};
                const newFont: any = { ...targetCell.s.font };
                if (style.font.bold !== undefined) newFont.bold = style.font.bold;
                if (style.font.italic !== undefined) newFont.italic = style.font.italic;
                if (style.font.underline !== undefined) newFont.underline = style.font.underline;
                if (style.font.name) newFont.name = style.font.name;
                if (style.font.size) newFont.sz = style.font.size;
                if (style.font.color) newFont.color = { rgb: style.font.color.replace('#', '') };
                targetCell.s.font = newFont;
            }

            if (style.alignment) {
                if (!targetCell.s.alignment) targetCell.s.alignment = {};
                const newAlignment = { ...targetCell.s.alignment };
                if (style.alignment.horizontal) newAlignment.horizontal = style.alignment.horizontal;
                if (style.alignment.vertical) newAlignment.vertical = style.alignment.vertical;
                targetCell.s.alignment = newAlignment;
            }

            if (style.fill && style.fill.color) {
                if (!targetCell.s.fill) targetCell.s.fill = {};
                targetCell.s.fill.patternType = 'solid';
                targetCell.s.fill.fgColor = { rgb: style.fill.color.replace('#', '') };
            }
            
            cellsFormatted++;
        }
      }
    }
    onProgress?.({ sheetName, currentSheet: i + 1, totalSheets: sheetNames.length, cellsFormatted });
  }

  return { workbook, cellsFormatted };
}
