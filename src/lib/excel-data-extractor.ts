
import * as XLSX from 'xlsx-js-style';
import type { ExtractionConfig, ExtractionReport } from './excel-types';
import { getColumnIndex, getUniqueSheetName, parseSourceColumns } from './excel-helpers';

/**
 * Finds all rows that match a specific value in a lookup column and extracts data from specified return columns.
 * @param workbook The workbook object.
 * @param sheetNames The names of the sheets to search within.
 * @param config Configuration for the lookup and extraction.
 * @param onProgress Optional callback for progress reporting.
 * @returns An ExtractionReport object.
 */
export function findAndExtractData(
  workbook: XLSX.WorkBook,
  sheetNames: string[],
  config: ExtractionConfig,
  onProgress?: (status: { sheetName: string, rowsFound: number, currentSheet: number, totalSheets: number }) => void
): ExtractionReport {
    const { lookupColumn, lookupValue, returnColumns, headerRow } = config;
    console.log(`Starting data extraction. Lookup: "${lookupValue}" in column "${lookupColumn}".`);
    const report: ExtractionReport = {
        summary: { 
            sheetsSearched: sheetNames,
            perSheetSummary: {},
            totalRowsFound: 0,
        },
        details: [],
        config: config
    };

    sheetNames.forEach((sheetName, index) => {
        onProgress?.({ sheetName, rowsFound: report.summary.totalRowsFound, currentSheet: index + 1, totalSheets: sheetNames.length });
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
            console.warn(`Sheet "${sheetName}" not found in the workbook, skipping.`);
            return;
        }
        console.log(`Processing sheet "${sheetName}" for extraction.`);

        const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
        const headerRowIndex = headerRow - 1;
        if (aoa.length <= headerRowIndex) {
            console.warn(`Header row ${headerRow} is out of bounds for sheet "${sheetName}", skipping.`);
            return;
        }

        const headers = aoa[headerRowIndex].map(h => String(h || ''));
        const lookupColIdx = getColumnIndex(lookupColumn, headers);
        
        if (lookupColIdx === null) {
            console.warn(`Lookup column "${lookupColumn}" not found in headers of sheet "${sheetName}", skipping.`);
            return;
        }

        let returnColIndices: number[];
        if (returnColumns.trim() === '*') {
          returnColIndices = headers.map((_, index) => index);
        } else {
          returnColIndices = parseSourceColumns(returnColumns, headers);
        }
        
        if (returnColIndices.length === 0) {
            console.warn(`Could not resolve any return columns: "${returnColumns}" in sheet "${sheetName}", skipping.`);
            return;
        }

        const returnColNames = returnColIndices.map(idx => headers[idx]);
        let rowsFoundInSheet = 0;

        for (let R = headerRow; R < aoa.length; R++) {
            const row = aoa[R];
            if (!row) continue;
            
            const cellValue = row[lookupColIdx];
            
            if (String(cellValue ?? '').trim().toLowerCase() === String(lookupValue).trim().toLowerCase()) {
                const extractedData: Record<string, any> = {
                    "Source Sheet": sheetName
                };
                returnColIndices.forEach((colIdx, i) => {
                    const headerName = returnColNames[i];
                    extractedData[headerName] = row[colIdx] ?? null;
                });
                report.details.push(extractedData as any);
                rowsFoundInSheet++;
            }
        }
        report.summary.perSheetSummary[sheetName] = rowsFoundInSheet;
        report.summary.totalRowsFound += rowsFoundInSheet;
        console.log(`Finished sheet "${sheetName}", found ${rowsFoundInSheet} matching rows.`);
        onProgress?.({ sheetName, rowsFound: report.summary.totalRowsFound, currentSheet: index + 1, totalSheets: sheetNames.length });
    });
    console.log(`Extraction complete. Total rows found: ${report.summary.totalRowsFound}.`);
    return report;
}

/**
 * Creates a new workbook containing the extracted data report.
 * @param report The ExtractionReport object.
 * @param config Configuration options, including chunk size.
 * @returns A new XLSX.WorkBook object.
 */
export function createExtractionReportWorkbook(
    report: ExtractionReport,
    config: { reportChunkSize?: number }
): XLSX.WorkBook {
    console.log(`Generating extraction report workbook with ${report.details.length} rows.`);
    const wb = XLSX.utils.book_new();
    const { details, summary } = report;
    const maxRowsPerSheet = (config.reportChunkSize && config.reportChunkSize > 0) ? config.reportChunkSize : 100000;
    
    if (details.length === 0) {
        const ws = XLSX.utils.aoa_to_sheet([["No matching rows found."]]);
        XLSX.utils.book_append_sheet(wb, ws, "Report");
        return wb;
    }
    
    // Dynamically determine headers from the first data object
    const baseHeaders = Object.keys(details[0]).filter(h => h !== "Source Sheet");
    const headers = summary.sheetsSearched.length > 1 
        ? ["Source Sheet", ...baseHeaders] 
        : baseHeaders;
    
    const headerRow = headers.map(h => ({ v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" as const } } }));
    
    const totalDetails = details.length;
    const numSheets = Math.ceil(totalDetails / (maxRowsPerSheet - 1));
    const baseSheetName = "Extracted_Data";
    console.log(`Report will be split into ${numSheets} sheet(s).`);

    for (let i = 0; i < numSheets; i++) {
        const sheetName = getUniqueSheetName(wb, numSheets > 1 ? `${baseSheetName}_${i + 1}` : baseSheetName);
        console.log(`Creating report sheet: "${sheetName}"`);
        const chunkStart = i * (maxRowsPerSheet - 1);
        const chunkEnd = Math.min(chunkStart + (maxRowsPerSheet - 1), totalDetails);
        const chunk = details.slice(chunkStart, chunkEnd);

        const aoa: any[][] = [headerRow];
        chunk.forEach(rowObject => {
            const rowArray = headers.map(header => rowObject[header]);
            aoa.push(rowArray);
        });

        const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
        
        const colWidths = headers.map(h => ({ wch: Math.max(15, h.length + 2) }));
        ws['!cols'] = colWidths;
        if (aoa.length > 1) {
            ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: headers.length - 1 } }) };
        }

        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    }

    return wb;
}
