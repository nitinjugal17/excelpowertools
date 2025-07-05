
import * as XLSX from 'xlsx-js-style';
import { getColumnIndex, getUniqueSheetName, sanitizeSheetNameForFormula } from './excel-helpers';
import type { ComparisonReport, SheetComparisonResult } from './excel-types';

/**
 * Compares two workbooks sheet by sheet based on a primary key.
 * @param wbA The first workbook (e.g., original).
 * @param wbB The second workbook (e.g., updated).
 * @param sheetNamesToCompare An array of sheet names that exist in both workbooks to be compared.
 * @param primaryKeyColumns A comma-separated string of column names or letters to use as a composite primary key.
 * @param headerRow The 1-indexed row number where headers are located.
 * @returns A structured report detailing the differences.
 */
export function compareWorkbooks(
  wbA: XLSX.WorkBook,
  wbB: XLSX.WorkBook,
  sheetNamesToCompare: string[],
  primaryKeyColumns: string,
  headerRow: number = 1
): ComparisonReport {
  const report: ComparisonReport = {
    summary: {
      totalSheetsCompared: sheetNamesToCompare.length,
      sheetsWithDifferences: [],
      totalRowsFound: 0, // Will be populated later
    },
    details: {},
    config: {
        headerRow,
        primaryKeyColumns,
    }
  };

  const headerRowIndex = headerRow - 1;

  sheetNamesToCompare.forEach(sheetName => {
    const wsA = wbA.Sheets[sheetName];
    const wsB = wbB.Sheets[sheetName];

    if (!wsA || !wsB) {
      console.warn(`Sheet "${sheetName}" not found in one of the workbooks, skipping.`);
      return;
    }

    const aoaA: any[][] = XLSX.utils.sheet_to_json(wsA, { header: 1, defval: null });
    const aoaB: any[][] = XLSX.utils.sheet_to_json(wsB, { header: 1, defval: null });

    const headersA = aoaA[headerRowIndex]?.map(h => String(h || ''));
    const headersB = aoaB[headerRowIndex]?.map(h => String(h || ''));

    if (!headersA || !headersB) {
        throw new Error(`Header row ${headerRow} not found on sheet "${sheetName}".`);
    }

    const keyColIndicesA = primaryKeyColumns.split(',').map(pk => getColumnIndex(pk, headersA));
    const keyColIndicesB = primaryKeyColumns.split(',').map(pk => getColumnIndex(pk, headersB));

    if (keyColIndicesA.some(i => i === null) || keyColIndicesB.some(i => i === null)) {
      throw new Error(`Primary key column(s) "${primaryKeyColumns}" not found on sheet "${sheetName}" in both files.`);
    }

    const mapA = new Map<string, any[]>();
    for (let i = headerRow; i < aoaA.length; i++) {
        const row = aoaA[i];
        if (!row || row.every(cell => cell === null)) continue;
        const key = keyColIndicesA.map(idx => String(row[idx!] ?? '').trim().toLowerCase()).join('||');
        if (key) mapA.set(key, row);
    }
    
    const mapB = new Map<string, any[]>();
     for (let i = headerRow; i < aoaB.length; i++) {
        const row = aoaB[i];
        if (!row || row.every(cell => cell === null)) continue;
        const key = keyColIndicesB.map(idx => String(row[idx!] ?? '').trim().toLowerCase()).join('||');
        if (key) mapB.set(key, row);
    }
    
    const sheetResult: SheetComparisonResult = {
        summary: { newRows: 0, deletedRows: 0, modifiedRows: 0 },
        new: [],
        deleted: [],
        modified: []
    };

    // Find modified and deleted rows
    for (const [key, rowA] of mapA.entries()) {
        const rowB = mapB.get(key);
        if (rowB) {
            // Key exists in both, check for modifications
            if (JSON.stringify(rowA) !== JSON.stringify(rowB)) { // Quick check first
                const diffs: { colName: string; valueA: any; valueB: any }[] = [];
                const maxCols = Math.max(rowA.length, rowB.length);
                for(let i = 0; i < maxCols; i++) {
                    const header = headersA[i] || headersB[i] || `Column ${i + 1}`;
                    const valueA = rowA[i] ?? null;
                    const valueB = rowB[i] ?? null;

                    if(String(valueA ?? '').trim().toLowerCase() !== String(valueB ?? '').trim().toLowerCase()) { // Compare as strings to handle type differences (e.g. 1 vs '1')
                        diffs.push({ colName: header, valueA: valueA, valueB: valueB });
                    }
                }
                if(diffs.length > 0) {
                    sheetResult.modified.push({ key, rowA, rowB, diffs });
                }
            }
        } else {
            // Key only in A, so it was deleted
            sheetResult.deleted.push(rowA);
        }
    }

    // Find new rows
    for (const [key, rowB] of mapB.entries()) {
        if (!mapA.has(key)) {
            sheetResult.new.push(rowB);
        }
    }
    
    sheetResult.summary.newRows = sheetResult.new.length;
    sheetResult.summary.deletedRows = sheetResult.deleted.length;
    sheetResult.summary.modifiedRows = sheetResult.modified.length;
    
    if (sheetResult.summary.newRows > 0 || sheetResult.summary.deletedRows > 0 || sheetResult.summary.modifiedRows > 0) {
        report.summary.sheetsWithDifferences.push(sheetName);
        report.details[sheetName] = sheetResult;
    }
  });

  return report;
}


/**
 * Merges differences from one workbook into another based on a comparison report.
 * @param wbA The first workbook.
 * @param wbB The second workbook.
 * @param report The comparison report detailing the differences.
 * @param baseFile Determines the direction of the merge ('A' means update A with changes from B).
 * @returns A new, reconciled XLSX.WorkBook object.
 */
export function reconcileWorkbooks(
  wbA: XLSX.WorkBook,
  wbB: XLSX.WorkBook,
  report: ComparisonReport,
  baseFile: 'A' | 'B' = 'A'
): XLSX.WorkBook {
  const reconciledWb = XLSX.utils.book_new();

  const sourceWb = baseFile === 'A' ? wbA : wbB;
  const changesWb = baseFile === 'A' ? wbB : wbA;
  const sourceName = baseFile === 'A' ? 'A' : 'B';

  // 1. Copy sheets that were not compared from the base workbook
  sourceWb.SheetNames.forEach(sheetName => {
    if (!report.summary.sheetsWithDifferences.includes(sheetName)) {
      // Deep copy the sheet to avoid issues with shared objects
      const sheet = JSON.parse(JSON.stringify(sourceWb.Sheets[sheetName]));
      XLSX.utils.book_append_sheet(reconciledWb, sheet, sheetName);
    }
  });

  // 2. Process and reconcile sheets with differences
  report.summary.sheetsWithDifferences.forEach(sheetName => {
    const newWs: XLSX.WorkSheet = {'!merges': [], '!ref': 'A1' };
    const reportDetail = report.details[sheetName];
    
    // Get headers and data from the base file's sheet
    const sourceWs = sourceWb.Sheets[sheetName];
    const sourceAoa: any[][] = XLSX.utils.sheet_to_json(sourceWs, { header: 1, defval: null });
    const headerRow = report.config.headerRow;
    const headerRowIndex = headerRow - 1;
    const headers = sourceAoa[headerRowIndex]?.map(h => String(h || ''));

    // Create a map of deleted and modified rows for quick lookup
    const deletedKeys = new Set(
        (sourceName === 'A' ? reportDetail.deleted : reportDetail.new).map(row => 
            report.config.primaryKeyColumns.split(',').map(pk => String(row[getColumnIndex(pk, headers)!] ?? '').trim().toLowerCase()).join('||')
        )
    );
    const modifiedRowsMap = new Map(reportDetail.modified.map(mod => [mod.key, sourceName === 'A' ? mod.rowB : mod.rowA]));

    let newRowIndex = 0;
    
    // Copy all rows from the base file, applying reconciliation logic
    for (let r = 0; r < sourceAoa.length; r++) {
      const rowData = sourceAoa[r];
      // Copy non-data rows (e.g., titles above header) directly
      if (r < headerRow) {
          rowData.forEach((_, c) => {
              const sourceCellAddr = XLSX.utils.encode_cell({r: r, c: c});
              if(sourceWs[sourceCellAddr]) {
                const destCellAddr = XLSX.utils.encode_cell({r: newRowIndex, c: c});
                newWs[destCellAddr] = JSON.parse(JSON.stringify(sourceWs[sourceCellAddr])); // deep copy cell
              }
          });
          newRowIndex++;
          continue;
      }
      
      const key = report.config.primaryKeyColumns.split(',').map(pk => String(rowData[getColumnIndex(pk, headers)!] ?? '').trim().toLowerCase()).join('||');
      
      if (deletedKeys.has(key)) {
          // Skip deleted rows
          continue;
      }
      
      const modifiedRowData = modifiedRowsMap.get(key);
      const finalRowData = modifiedRowData || rowData;

      // Write the final row data, preserving styles from the base file's row where possible
      finalRowData.forEach((cellVal, c) => {
          const destCellAddr = XLSX.utils.encode_cell({r: newRowIndex, c});
          const sourceCellAddr = XLSX.utils.encode_cell({r: r, c: c});
          if (sourceWs[sourceCellAddr]) {
             newWs[destCellAddr] = JSON.parse(JSON.stringify(sourceWs[sourceCellAddr])); // copy original cell with style
             newWs[destCellAddr].v = cellVal; // overwrite value
          } else {
             // If original cell doesn't exist, create a new one. This is less common.
             XLSX.utils.sheet_add_aoa(newWs, [[cellVal]], {origin: destCellAddr});
          }
      });
      newRowIndex++;
    }

    // Add new rows to the end
    const newRows = sourceName === 'A' ? reportDetail.new : reportDetail.deleted;
    if (newRows.length > 0) {
      const rowsToAdd = newRows.map(row => {
          // ensure row has same length as headers to avoid data shifting
          const fullRow = [...row];
          while(fullRow.length < headers.length) {
              fullRow.push(null);
          }
          return fullRow;
      });
      XLSX.utils.sheet_add_aoa(newWs, rowsToAdd, { origin: -1 }); // -1 appends to the end
    }
    
    // Finalize sheet properties
    const finalRange = XLSX.utils.decode_range(sourceWs['!ref'] || 'A1');
    const finalRowCount = newRowIndex;
    newWs['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: finalRowCount - 1, c: finalRange.e.c } });
    
    if (sourceWs['!cols']) newWs['!cols'] = JSON.parse(JSON.stringify(sourceWs['!cols'])); // Copy column widths
    if (sourceWs['!merges']) newWs['!merges'] = JSON.parse(JSON.stringify(sourceWs['!merges']));

    XLSX.utils.book_append_sheet(reconciledWb, newWs, sheetName);
  });

  return reconciledWb;
}


/**
 * Creates a new workbook visualizing the comparison report.
 * @param report The comparison report from `compareWorkbooks`.
 * @param nameA Name of the first file.
 * @param nameB Name of the second file.
 * @returns A new XLSX.WorkBook object.
 */
export function generateComparisonReportWorkbook(report: ComparisonReport, nameA: string, nameB: string): XLSX.WorkBook {
    const wb = XLSX.utils.book_new();

    // Summary Sheet
    const summaryData: any[][] = [
        [{ v: 'File Comparison Report', s: { font: { bold: true, sz: 16 } } }],
        [],
        ['File A:', nameA],
        ['File B:', nameB],
        [],
        [{ v: 'Summary of Differences', s: { font: { bold: true, sz: 12 } } }],
        ['Sheet Name', 'New Rows', 'Deleted Rows', 'Modified Rows'].map(h => ({ v: h, s: { font: { bold: true } } }))
    ];
    report.summary.sheetsWithDifferences.forEach(sheetName => {
        const summary = report.details[sheetName].summary;
        summaryData.push([sheetName, summary.newRows, summary.deletedRows, summary.modifiedRows]);
    });
    const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
    summaryWs['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 15 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, summaryWs, "Summary");

    // Detail Sheets
    report.summary.sheetsWithDifferences.forEach(sheetName => {
        const detail = report.details[sheetName];
        const detailWsName = getUniqueSheetName(wb, `Compare_${sheetName}`);
        
        const data: any[][] = [];
        
        // Add Modified Rows
        if (detail.modified.length > 0) {
            data.push([{v: `Modified Rows (${detail.modified.length})`, s: { font: { bold: true, sz: 14, color: { rgb: "BF8F00" } } }}]);
            data.push(['Key', 'Column Changed', `Value in ${nameA} (Old)`, `Value in ${nameB} (New)`].map(h => ({v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "FFEB9C" }}}})));
            detail.modified.forEach(mod => {
                mod.diffs.forEach((diff, index) => {
                    data.push([
                        index === 0 ? mod.key : '',
                        diff.colName,
                        diff.valueA,
                        diff.valueB
                    ]);
                });
                data.push([]); // Spacer
            });
        }
        
        // Get headers from first available row data for New/Deleted
        const firstNewRow = detail.new[0];
        const firstDeletedRow = detail.deleted[0];
        let rowHeaders: string[] = [];

        if (firstNewRow || firstDeletedRow) {
            const sampleRow = firstNewRow || firstDeletedRow;
            rowHeaders = sampleRow.map((_, i) => `Column ${i + 1}`);
        } else if (detail.modified.length > 0 && detail.modified[0].rowA) {
             rowHeaders = detail.modified[0].rowA.map((_, i) => `Column ${i + 1}`);
        }
        
        // Add New Rows
        if (detail.new.length > 0) {
            if (data.length > 0) data.push([]); // Spacer if modified rows exist
            data.push([{v: `New Rows (${detail.new.length})`, s: { font: { bold: true, sz: 14, color: { rgb: "385723" } } }}]);
            data.push(rowHeaders.map(h => ({v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "C6EFCE" } } }})));
            detail.new.forEach(row => {
                data.push(row);
            });
        }
        
        // Add Deleted Rows
        if (detail.deleted.length > 0) {
            if (data.length > 0) data.push([]); // Spacer
            data.push([{v: `Deleted Rows (${detail.deleted.length})`, s: { font: { bold: true, sz: 14, color: { rgb: "9C0006" } } }}]);
            data.push(rowHeaders.map(h => ({v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "FFC7CE" } } }})));
            detail.deleted.forEach(row => {
                data.push(row);
            });
        }
        
        const detailWs = XLSX.utils.aoa_to_sheet(data, {cellStyles: true});
        if (data.length > 0) {
            const maxCols = data.reduce((max, row) => Math.max(max, row.length), 0);
            detailWs['!cols'] = Array.from({ length: maxCols }, () => ({ wch: 25 }));
        }
        
        XLSX.utils.book_append_sheet(wb, detailWs, detailWsName);
    });

    return wb;
}
