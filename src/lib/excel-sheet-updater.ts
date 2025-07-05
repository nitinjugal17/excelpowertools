
import * as XLSX from 'xlsx-js-style';
import type { CustomColumnConfig, CustomHeaderConfig, FormattingConfig, RangeFormattingConfig, HorizontalAlignment, VerticalAlignment, SheetProtectionConfig } from './excel-types';
import { parseColumnIdentifier } from './excel-helpers';

interface Font {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  name?: string;
  sz?: number;
  color?: { rgb: string };
}


/**
 * Formats headers, applies AutoFilter, and optionally inserts a custom header or column into specified sheets.
 * This function now correctly handles sequential operations by recalculating sheet dimensions after each step.
 * @param workbook The workbook to modify.
 * @param sheetNamesToUpdate Array of sheet names to apply changes to.
 * @param formattingConfig Configuration for styling the data column titles row.
 * @param customHeaderConfig Optional configuration for inserting a new custom header.
 * @param customColumnConfig Optional configuration for inserting a new custom column.
 * @param rangeFormattingConfig Optional configuration for formatting a specific range.
 * @param sheetProtectionConfig Optional configuration for protecting sheets.
 * @param onProgress Optional callback for progress reporting.
 * @param cancellationRequestedRef Optional ref to check for cancellation requests.
 * @returns The modified XLSX.WorkBook object.
 */
export function formatAndUpdateSheets(
  workbook: XLSX.WorkBook,
  sheetNamesToUpdate: string[],
  formattingConfig: FormattingConfig,
  customHeaderConfig?: CustomHeaderConfig,
  customColumnConfig?: CustomColumnConfig,
  rangeFormattingConfig?: RangeFormattingConfig,
  sheetProtectionConfig?: SheetProtectionConfig,
  onProgress?: (status: { sheetName: string; currentSheet: number; totalSheets: number; operation: string }) => void,
  cancellationRequestedRef?: React.RefObject<boolean>
): XLSX.WorkBook {
  
  for (let i = 0; i < sheetNamesToUpdate.length; i++) {
    const sheetName = sheetNamesToUpdate[i];
    if (cancellationRequestedRef?.current) throw new Error('Cancelled by user.');

    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    onProgress?.({ sheetName, currentSheet: i + 1, totalSheets: sheetNamesToUpdate.length, operation: "Starting..." });

    // Effective row indices will be adjusted as rows/columns are inserted.
    let effectiveDataTitlesRow = formattingConfig.dataTitlesRowNumber;
    let effectiveNewColHeaderRow = customColumnConfig?.newColumnHeaderRow;
    let effectiveDataStartRow = customColumnConfig?.dataStartRow;

    // --- Step 1: Insert Custom Header Row ---
    if (customHeaderConfig) {
        onProgress?.({ sheetName, currentSheet: i + 1, totalSheets: sheetNamesToUpdate.length, operation: "Inserting custom header" });
        const insertAtIndex = customHeaderConfig.insertBeforeRow - 1;
        
        let range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        
        // Shift all cells down from the insertion point
        for (let R = range.e.r; R >= insertAtIndex; --R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const fromAddr = XLSX.utils.encode_cell({ r: R, c: C });
                const toAddr = XLSX.utils.encode_cell({ r: R + 1, c: C });
                if (worksheet[fromAddr]) {
                    worksheet[toAddr] = worksheet[fromAddr];
                    delete worksheet[fromAddr];
                }
            }
        }
        
        // Update worksheet range and merges
        range.e.r += 1;
        worksheet['!ref'] = XLSX.utils.encode_range(range);
        if (worksheet['!merges']) {
            worksheet['!merges'].forEach(merge => {
                if(merge.s.r >= insertAtIndex) {
                    merge.s.r++;
                    merge.e.r++;
                }
            });
        }
        
        // Insert new header text and apply styles
        const headerAddr = XLSX.utils.encode_cell({ r: insertAtIndex, c: range.s.c });
        XLSX.utils.sheet_add_aoa(worksheet, [[customHeaderConfig.text]], { origin: headerAddr });
        
        const cell = worksheet[headerAddr];
        if (cell) {
            if (!cell.s) cell.s = {};
            const { styleOptions } = customHeaderConfig;
            cell.s.font = {
                bold: styleOptions.bold,
                italic: styleOptions.italic,
                underline: styleOptions.underline,
                sz: styleOptions.fontSize || 12,
                name: styleOptions.fontName || 'Calibri'
            };
            cell.s.alignment = {
                horizontal: customHeaderConfig.mergeAndCenter ? 'center' : (styleOptions.horizontalAlignment || 'general'),
                vertical: styleOptions.verticalAlignment || 'center',
                wrapText: !!styleOptions.wrapText,
                indent: styleOptions.indent || 0,
            };
        }
        
        if (customHeaderConfig.mergeAndCenter) {
            if (!worksheet['!merges']) worksheet['!merges'] = [];
            worksheet['!merges'].push({ s: { r: insertAtIndex, c: range.s.c }, e: { r: insertAtIndex, c: range.e.c } });
        }
        
        // Adjust effective row indices for subsequent steps
        if (insertAtIndex < effectiveDataTitlesRow) effectiveDataTitlesRow++;
        if (effectiveNewColHeaderRow && insertAtIndex < effectiveNewColHeaderRow) effectiveNewColHeaderRow++;
        if (effectiveDataStartRow && insertAtIndex < effectiveDataStartRow) effectiveDataStartRow++;
    }

    // --- Step 2: Insert Custom Column ---
    if (customColumnConfig && effectiveNewColHeaderRow && effectiveDataStartRow) {
      onProgress?.({ sheetName, currentSheet: i + 1, totalSheets: sheetNamesToUpdate.length, operation: "Inserting custom column" });
      const insertColIdx = parseColumnIdentifier(customColumnConfig.insertColumnBefore);
      const sourceColIdx = parseColumnIdentifier(customColumnConfig.sourceDataColumn);

      if (insertColIdx !== null && sourceColIdx !== null) {
        let range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        
        // Shift cells to the right
        for (let R = range.s.r; R <= range.e.r; ++R) {
          for (let C = range.e.c; C >= insertColIdx; --C) {
            const fromAddr = XLSX.utils.encode_cell({ r: R, c: C });
            const toAddr = XLSX.utils.encode_cell({ r: R, c: C + 1 });
            if (worksheet[fromAddr]) {
              worksheet[toAddr] = worksheet[fromAddr];
              delete worksheet[fromAddr];
            }
          }
        }
        
        // Update range and merges
        range.e.c += 1;
        worksheet['!ref'] = XLSX.utils.encode_range(range);
        if (worksheet['!merges']) {
          worksheet['!merges'].forEach(merge => {
            if (merge.s.c >= insertColIdx) {
              merge.s.c++;
              merge.e.c++;
            } else if (merge.e.c >= insertColIdx) {
              merge.e.c++;
            }
          });
        }
        
        const finalSourceColIdx = sourceColIdx < insertColIdx ? sourceColIdx : sourceColIdx + 1;
        
        // Insert new header
        const newHeaderAddr = XLSX.utils.encode_cell({ r: effectiveNewColHeaderRow - 1, c: insertColIdx });
        XLSX.utils.sheet_add_aoa(worksheet, [[customColumnConfig.newColumnName]], { origin: newHeaderAddr });
        
        // Populate new column data
        for (let R = effectiveDataStartRow - 1; R <= range.e.r; ++R) {
          const sourceAddr = XLSX.utils.encode_cell({ r: R, c: finalSourceColIdx });
          const sourceCell = worksheet[sourceAddr];
          const sourceValue = (sourceCell && sourceCell.v !== null) ? String(sourceCell.v) : '';
          let newValue = '';
          
          if (sourceValue) {
            const parts = sourceValue.split(customColumnConfig.textSplitter);
            if (parts.length > 0) {
              const partIndex = customColumnConfig.partToUse === -1 ? parts.length - 1 : customColumnConfig.partToUse - 1;
              if (partIndex >= 0 && partIndex < parts.length) {
                newValue = parts[partIndex]?.trim() || '';
              }
            }
          }
          
          const newCellAddr = XLSX.utils.encode_cell({ r: R, c: insertColIdx });
          XLSX.utils.sheet_add_aoa(worksheet, [[newValue]], { origin: newCellAddr });

          if (customColumnConfig.alignment && customColumnConfig.alignment !== 'general') {
            const cell = worksheet[newCellAddr];
            if (cell) {
              if (!cell.s) cell.s = {};
              cell.s.alignment = { ...cell.s.alignment, horizontal: customColumnConfig.alignment };
            }
          }
        }
      }
    }

    // --- Step 3: Apply Custom Range Formatting ---
    if (rangeFormattingConfig) {
        onProgress?.({ sheetName, currentSheet: i + 1, totalSheets: sheetNamesToUpdate.length, operation: "Applying range formatting" });
        const { startRow, endRow, startCol, endCol, merge, style } = rangeFormattingConfig;
        const startR = startRow - 1;
        const endR = endRow - 1;
        const startC = parseColumnIdentifier(startCol);
        const endC = parseColumnIdentifier(endCol);

        if (startC !== null && endC !== null && startR >= 0 && endR >= startR && endC >= startC) {
            for (let R = startR; R <= endR; ++R) {
                for (let C = startC; C <= endC; ++C) {
                    const addr = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!worksheet[addr]) worksheet[addr] = { t: 'z' };
                    const cell = worksheet[addr];
                    if (!cell.s) cell.s = {};

                    if (style.font) {
                        if (!cell.s.font) cell.s.font = {};
                        const newFont: Font = { ...cell.s.font };
                        if (style.font.bold !== undefined) newFont.bold = style.font.bold;
                        if (style.font.italic !== undefined) newFont.italic = style.font.italic;
                        if (style.font.underline !== undefined) newFont.underline = style.font.underline;
                        if (style.font.name) newFont.name = style.font.name;
                        if (style.font.size) newFont.sz = style.font.size;
                        if (style.font.color) newFont.color = { rgb: style.font.color.replace('#', '') };
                        cell.s.font = newFont;
                    }
                    if (style.alignment) {
                        if (!cell.s.alignment) cell.s.alignment = {};
                        cell.s.alignment = {...cell.s.alignment, ...style.alignment};
                    }
                    if (style.fill.color) {
                        if (!cell.s.fill) cell.s.fill = {};
                        cell.s.fill = { ...cell.s.fill, patternType: 'solid', fgColor: { rgb: style.fill.color.replace('#', '') } };
                    }
                }
            }
            if (merge) {
                if (!worksheet['!merges']) worksheet['!merges'] = [];
                const mergeExists = worksheet['!merges'].some(m => m.s.r === startR && m.s.c === startC && m.e.r === endR && m.e.c === endC);
                if (!mergeExists) worksheet['!merges'].push({ s: { r: startR, c: startC }, e: { r: endR, c: endC } });
            }
        }
    }
    
    // --- Step 4: Format Data Headers & AutoFilter ---
    onProgress?.({ sheetName, currentSheet: i + 1, totalSheets: sheetNamesToUpdate.length, operation: "Formatting headers & table" });
    const finalRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    const finalDataTitlesRowIndex = effectiveDataTitlesRow - 1;

    if (finalDataTitlesRowIndex >= 0 && finalDataTitlesRowIndex <= finalRange.e.r) {
      for (let C = finalRange.s.c; C <= finalRange.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: finalDataTitlesRowIndex, c: C });
        let cell = worksheet[cellAddress];
        if (!cell) {
          worksheet[cellAddress] = { t: 'z' };
          cell = worksheet[cellAddress];
        }
        if (!cell.s) cell.s = {};
        
        if (!cell.s.font) cell.s.font = {};
        const newFont: Font = { ...cell.s.font };
        const { styleOptions } = formattingConfig;
        if(styleOptions.bold !== undefined) newFont.bold = styleOptions.bold;
        if(styleOptions.italic !== undefined) newFont.italic = styleOptions.italic;
        if(styleOptions.underline !== undefined) newFont.underline = styleOptions.underline;
        if(styleOptions.fontName) newFont.name = styleOptions.fontName;
        if(styleOptions.fontSize) newFont.sz = styleOptions.fontSize;
        cell.s.font = newFont;
        
        if (styleOptions.alignment && styleOptions.alignment !== 'general') {
          if (!cell.s.alignment) cell.s.alignment = {};
          cell.s.alignment = { ...cell.s.alignment, horizontal: styleOptions.alignment };
        }
      }
      
      if (finalDataTitlesRowIndex < finalRange.e.r) {
        worksheet['!autofilter'] = { ref: XLSX.utils.encode_range({
          s: { r: finalDataTitlesRowIndex, c: finalRange.s.c },
          e: { r: finalRange.e.r, c: finalRange.e.c } 
        }) };
      }
    }
    
    // --- Step 5: Auto-adjust column widths ---
    const aoaDataForWidth: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
    const numCols = aoaDataForWidth.reduce((max: number, row: any[]) => Math.max(max, row?.length || 0), 0);
    if (numCols > 0) {
      const colWidths = Array.from({ length: numCols as number }).map((_, colIdx) => {
        let maxLength = 0;
        aoaDataForWidth.forEach((row: any[]) => {
          const cellValue = row?.[colIdx];
          const cellTextLength = cellValue !== null && cellValue !== undefined ? String(cellValue).length : 0;
          if (cellTextLength > maxLength) maxLength = cellTextLength;
        });
        return { wch: Math.min(60, Math.max(10, maxLength + 2)) };
      });
      worksheet['!cols'] = colWidths;
    }

    // --- Step 6: Apply Sheet Protection ---
    if (sheetProtectionConfig && sheetProtectionConfig.password) {
        onProgress?.({ sheetName, currentSheet: i + 1, totalSheets: sheetNamesToUpdate.length, operation: "Applying sheet protection" });
        
        // This sets the sheet-level protection options.
        worksheet['!protect'] = {
            password: sheetProtectionConfig.password,
            selectLockedCells: true,
            selectUnlockedCells: true,
        };

        if (sheetProtectionConfig.type === 'range' && sheetProtectionConfig.range) {
            // First, unlock all cells in the sheet.
            const fullRange = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            for (let R = fullRange.s.r; R <= fullRange.e.r; ++R) {
                for (let C = fullRange.s.c; C <= fullRange.e.c; ++C) {
                    const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                    const cell = worksheet[cellAddress] || { t: 'z' };
                    if (!cell.s) cell.s = {};
                    if (!cell.s.protection) cell.s.protection = {};
                    cell.s.protection.locked = false;
                    worksheet[cellAddress] = cell;
                }
            }

            // Then, lock only the cells in the specified range.
            try {
                const protectionRange = XLSX.utils.decode_range(sheetProtectionConfig.range);
                for (let R = protectionRange.s.r; R <= protectionRange.e.r; ++R) {
                    for (let C = protectionRange.s.c; C <= protectionRange.e.c; ++C) {
                       const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                       const cell = worksheet[cellAddress] || { t: 'z' };
                       if (!cell.s) cell.s = {};
                       if (!cell.s.protection) cell.s.protection = {};
                       cell.s.protection.locked = true;
                       worksheet[cellAddress] = cell;
                    }
                }
            } catch (e) {
                console.error(`Invalid protection range "${sheetProtectionConfig.range}" provided. Falling back to full sheet protection.`);
            }
        }
        // If type is 'full', we don't need to do anything with cell-level locking, as all cells are locked by default when the sheet is protected.
    }
  }

  return workbook;
}
