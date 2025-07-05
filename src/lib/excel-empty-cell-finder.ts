
import * as XLSX from 'xlsx-js-style';
import type { EmptyCellReport, EmptyCellResult } from './excel-types';
import { getColumnIndex, sanitizeSheetNameForFormula, getUniqueSheetName, parseSourceColumns } from './excel-helpers';


/**
 * Finds empty cells, optionally highlights them, and generates/merges a report within the workbook.
 * This is the main orchestrator for the Empty Cell Finder tool.
 * @param workbook The workbook to process. It will be modified directly.
 * @returns An object containing the (potentially modified) workbook and the report.
 */
export function findAndHighlightEmptyCells(
    workbook: XLSX.WorkBook,
    sheetNamesToCheck: string[],
    columnsToCheck: string,
    columnsToIgnore: string | undefined,
    headerRow: number,
    highlightColor: string | undefined,
    generateReport: boolean,
    reportOptions?: {
        format: 'compact' | 'detailed' | 'summary';
        columnsToInclude: string;
        includeAllData: boolean;
        contextColumnForCompact?: string;
        summaryKeyColumn?: string;
        summaryContextColumn?: string;
        blankKeyLabel?: string;
        chunkSize?: number;
    },
    onProgress?: (status: { sheetName: string; currentSheet: number; totalSheets: number; emptyFound: number }) => void
): { report: EmptyCellReport; workbook: XLSX.WorkBook } {
    const newFoundEmptyCells: EmptyCellResult[] = [];
    
    const reportBaseNames = new Set([
        "empty_cell_summary",
        "empty_cell_links",
        "empty_cell_details",
        "empty_cell_pivot_summary",
        "empty_cell_pivot_data"
    ]);

    const originalSheetOrder = [...workbook.SheetNames];
    const sheetsToActuallyCheck = sheetNamesToCheck.filter(name => {
        const lowerCaseName = name.toLowerCase();
        for (const baseName of reportBaseNames) {
            if (lowerCaseName.startsWith(baseName)) {
                return false;
            }
        }
        return true;
    });

    let totalEmptyFound = 0;
    
    sheetsToActuallyCheck.forEach((sheetName, sheetIndex) => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return;

        onProgress?.({ sheetName, currentSheet: sheetIndex + 1, totalSheets: sheetsToActuallyCheck.length, emptyFound: totalEmptyFound });

        const range = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 }};
        const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
        
        const headerRowIndex = headerRow > 0 ? headerRow - 1 : 0;
        if (aoa.length <= headerRowIndex) return;
        
        const startDataRow = headerRowIndex + 1;
        if (aoa.length <= startDataRow) return;

        const headers = aoa[headerRowIndex]?.map(h => String(h || '')) || [];
        
        const allPossibleCols = Array.from({length: (range.e.c) + 1}, (_, i) => i);
        const colsToCheckIndices = columnsToCheck === '*' ? allPossibleCols : parseSourceColumns(columnsToCheck, headers);
        const colsToIgnoreIndices = new Set(columnsToIgnore ? parseSourceColumns(columnsToIgnore, headers) : []);
        
        const finalColsToScan = colsToCheckIndices.filter(c => !colsToIgnoreIndices.has(c));

        if (finalColsToScan.length === 0) return;

        const contextColIdxForCompact = (generateReport && reportOptions?.format === 'compact' && reportOptions.contextColumnForCompact)
            ? getColumnIndex(reportOptions.contextColumnForCompact, headers)
            : null;

        const summaryKeyColIdx = (generateReport && reportOptions?.format === 'summary' && reportOptions.summaryKeyColumn)
            ? getColumnIndex(reportOptions.summaryKeyColumn, headers)
            : null;
        
        const summaryContextColIdx = (generateReport && reportOptions?.format === 'summary' && reportOptions.summaryContextColumn)
            ? getColumnIndex(reportOptions.summaryContextColumn, headers)
            : null;
        
        const blankKeyLabel = (generateReport && reportOptions?.format === 'summary' && reportOptions?.blankKeyLabel) ? reportOptions.blankKeyLabel : '(Blanks)';
        
        for (let R = startDataRow; R <= range.e.r; ++R) {
            for (const C of finalColsToScan) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = worksheet[cellAddress];
                
                const isCellEmpty = !cell || cell.v === null || cell.v === undefined || String(cell.v).trim() === '';
                
                if (isCellEmpty) {
                    totalEmptyFound++;
                    
                    if (totalEmptyFound % 50 === 0) {
                        onProgress?.({ sheetName, currentSheet: sheetIndex + 1, totalSheets: sheetsToActuallyCheck.length, emptyFound: totalEmptyFound });
                    }

                    if (highlightColor) {
                        if (!worksheet[cellAddress]) {
                            worksheet[cellAddress] = { t: 'z' };
                        }
                        const targetCell = worksheet[cellAddress];
                        if (!targetCell.s) targetCell.s = {};
                        
                        const existingFill = targetCell.s.fill || {};
                        
                        targetCell.s.fill = {
                            ...existingFill,
                            patternType: 'solid',
                            fgColor: { rgb: highlightColor.replace('#', '') }
                        };
                    }

                    if (generateReport) {
                        const result: EmptyCellResult = { 
                            sheetName, 
                            address: cellAddress,
                            row: R + 1,
                        };

                        if (reportOptions?.format === 'compact' && contextColIdxForCompact !== null) {
                            const contextVal = aoa[R]?.[contextColIdxForCompact] ?? null;
                            if (contextVal !== null && contextVal !== undefined) {
                                const contextString = String(contextVal).trim();
                                if (contextString) {
                                    result.contextValue = contextString;
                                }
                            }
                        }

                         if (reportOptions?.format === 'summary' && summaryKeyColIdx !== null && summaryContextColIdx !== null) {
                            result.keyColumnValue = aoa[R]?.[summaryKeyColIdx] ?? blankKeyLabel;
                            result.contextColumnForSummaryValue = aoa[R]?.[summaryContextColIdx] ?? blankKeyLabel;
                        }

                        if (reportOptions?.format === 'detailed') {
                            result.rowData = headers.reduce((obj, header, i) => {
                                if(header && !colsToIgnoreIndices.has(i)) {
                                    obj[header] = aoa[R]?.[i] ?? null;
                                }
                                return obj;
                            }, {} as Record<string, any>);
                        }
                        newFoundEmptyCells.push(result);
                    }
                }
            }
        }
        onProgress?.({ sheetName, currentSheet: sheetIndex + 1, totalSheets: sheetsToActuallyCheck.length, emptyFound: totalEmptyFound });
    });

    const finalReport: EmptyCellReport = {
        summary: {},
        locations: newFoundEmptyCells,
        totalEmpty: totalEmptyFound,
        processedSheetNames: sheetsToActuallyCheck,
    };

    newFoundEmptyCells.forEach(loc => {
        finalReport.summary[loc.sheetName] = (finalReport.summary[loc.sheetName] || 0) + 1;
    });
    
    finalReport.processedSheetNames = [...new Set(Object.keys(finalReport.summary))].sort((a,b) => originalSheetOrder.indexOf(a) - originalSheetOrder.indexOf(b));


    if (generateReport && finalReport.totalEmpty > 0 && reportOptions) {
        if (reportOptions.format === 'detailed') {
            const maxRowsPerSheet = (reportOptions.chunkSize && reportOptions.chunkSize > 0) ? reportOptions.chunkSize : 100000;
            const { summarySheet } = createCompactReportSheets(finalReport, reportOptions);
            const summarySheetBaseName = "Empty_Cell_Summary";
            const detailsSheetBaseName = "Empty_Cell_Details";
            const summarySheetName = getUniqueSheetName(workbook, summarySheetBaseName);
            XLSX.utils.book_append_sheet(workbook, summarySheet, summarySheetName);

            const firstSheet = workbook.Sheets[finalReport.processedSheetNames[0]];
            const firstSheetAOA: any[][] = firstSheet ? XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: null }) : [[]];
            const headers = firstSheetAOA[headerRow - 1]?.map(h => String(h || '')) || [];
            const ignoredColumnIndices = new Set(columnsToIgnore ? parseSourceColumns(columnsToIgnore, headers) : []);
            
            let contextColumns: string[] = [];
            if (reportOptions.includeAllData) {
                contextColumns = headers.filter((_, index) => !ignoredColumnIndices.has(index));
            } else if (reportOptions.columnsToInclude) {
                contextColumns = reportOptions.columnsToInclude.split(',').map(s => s.trim()).filter(Boolean);
            }
            
            const detailHeaders = [ 'Sheet Name', 'Row #', 'Empty Cell', ...contextColumns ]
                .map(h => ({ v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" as const } } }));
            
            const locations = finalReport.locations;
            const totalDetails = locations.length;
            const numSheets = Math.ceil(totalDetails / (maxRowsPerSheet - 1));
            const generatedDetailSheetNames: string[] = [];

            for (let i = 0; i < numSheets; i++) {
                const sheetName = getUniqueSheetName(workbook, numSheets > 1 ? `${detailsSheetBaseName}_${i + 1}` : detailsSheetBaseName);
                generatedDetailSheetNames.push(sheetName);

                const chunkStart = i * (maxRowsPerSheet - 1);
                const chunkEnd = Math.min(chunkStart + (maxRowsPerSheet - 1), totalDetails);
                const chunk = locations.slice(chunkStart, chunkEnd);

                const detailsAOA: any[][] = [detailHeaders];
                chunk.forEach(loc => {
                    const row: any[] = [
                        { v: loc.sheetName },
                        { v: loc.row, t: 'n', l: { Target: `#${sanitizeSheetNameForFormula(loc.sheetName)}!A${loc.row}` }, s: { font: { color: { rgb: "0000FF" }, underline: true } } },
                        { v: loc.address, l: { Target: `#${sanitizeSheetNameForFormula(loc.sheetName)}!${loc.address}` }, s: { font: { color: { rgb: "0000FF" }, underline: true } } },
                    ];
                    contextColumns.forEach(header => {
                        let cellValue = loc.rowData?.[header] ?? null;
                        if (typeof cellValue === 'string' && cellValue.length > 500) {
                            cellValue = cellValue.substring(0, 500) + '... (truncated)';
                        }
                        row.push(cellValue);
                    });
                    detailsAOA.push(row);
                });

                const detailsSheet = XLSX.utils.aoa_to_sheet(detailsAOA, { cellDates: true });
                const colWidths = [
                    { wch: 20 }, { wch: 10 }, { wch: 20 },
                    ...contextColumns.map(h => ({ wch: Math.max(15, h.length + 2) }))
                ];
                detailsSheet['!cols'] = colWidths;
                if (detailsAOA.length > 1) {
                    detailsSheet['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: detailsAOA.length - 1, c: detailHeaders.length - 1 } }) };
                }
                XLSX.utils.book_append_sheet(workbook, detailsSheet, sheetName);
            }
            
            const allSheets = workbook.SheetNames;
            const otherSheets = allSheets.filter(name => name !== summarySheetName && !generatedDetailSheetNames.includes(name));
            workbook.SheetNames = [summarySheetName, ...generatedDetailSheetNames.sort(), ...otherSheets];
        } else {
            // Handle other non-chunked report formats
            let reportSheets: { summarySheet: XLSX.WorkSheet, detailsSheet: XLSX.WorkSheet };
            let summarySheetBaseName: string;
            let detailsSheetBaseName: string;
            if (reportOptions.format === 'summary' && reportOptions.summaryKeyColumn && reportOptions.summaryContextColumn) {
                reportSheets = createSummaryReportSheets(finalReport, {
                    summaryKeyColumn: reportOptions.summaryKeyColumn,
                    summaryContextColumn: reportOptions.summaryContextColumn,
                    blankKeyLabel: reportOptions.blankKeyLabel,
                });
                summarySheetBaseName = "Empty_Cell_Pivot_Summary";
                detailsSheetBaseName = "Empty_Cell_Pivot_Data";
            } else {
                reportSheets = createCompactReportSheets(finalReport, reportOptions);
                summarySheetBaseName = "Empty_Cell_Summary";
                detailsSheetBaseName = "Empty_Cell_Links";
            }
            
            const summarySheetName = getUniqueSheetName(workbook, summarySheetBaseName);
            XLSX.utils.book_append_sheet(workbook, reportSheets.summarySheet, summarySheetName);

            const detailsSheetName = getUniqueSheetName(workbook, detailsSheetBaseName);
            XLSX.utils.book_append_sheet(workbook, reportSheets.detailsSheet, detailsSheetName);
            
            const allSheets = workbook.SheetNames;
            const otherSheets = allSheets.filter(name => name !== summarySheetName && name !== detailsSheetName);
            workbook.SheetNames = [summarySheetName, detailsSheetName, ...otherSheets];
        }
    }
    
    return { report: finalReport, workbook };
}


/**
 * Creates new worksheets for an empty cell report from an EmptyCellReport object. (Compact Format)
 * @param report The report object to generate sheets from.
 * @returns An object containing the summary and details worksheets.
 */
function createCompactReportSheets(
    report: EmptyCellReport,
    reportOptions?: { contextColumnForCompact?: string }
): { summarySheet: XLSX.WorkSheet, detailsSheet: XLSX.WorkSheet } {
    const sheetOrder = report.processedSheetNames || Object.keys(report.summary);

    const summaryAOA: any[][] = [];
    summaryAOA.push([{ v: 'Empty Cell Report Summary', s: { font: { bold: true, sz: 16 } } }]);
    summaryAOA.push([]);
    summaryAOA.push([
        { v: 'Sheet Name', s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" } } },
        { v: 'Total Number of Empty Cells/Rows', s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" } } }
    ]);

    const sheetsForSummary = sheetOrder.filter(sheetName => report.summary[sheetName]);
    sheetsForSummary.forEach((sheetName) => {
        summaryAOA.push([sheetName, report.summary[sheetName] || 0]);
    });
    summaryAOA.push([]);
    summaryAOA.push([
        { v: 'Total', s: { font: { bold: true } } },
        { v: report.totalEmpty, s: { font: { bold: true } } }
    ]);

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryAOA);
    if (!summarySheet['!merges']) summarySheet['!merges'] = [];
    summarySheet['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } });
    summarySheet['!cols'] = [{ wch: 30 }, { wch: 30 }];

    const detailsAOA: any[][] = [];
    detailsAOA.push([{ v: 'Empty Cell Links', s: { font: { bold: true, sz: 16 } } }]);
    detailsAOA.push([]);
    
    const cellsBySheet: { [sheetName: string]: EmptyCellResult[] } = {};
    const sheetsForDetails = sheetOrder.filter(name => report.summary[name] > 0);
    
    report.locations.forEach(loc => {
        if (!cellsBySheet[loc.sheetName]) cellsBySheet[loc.sheetName] = [];
        cellsBySheet[loc.sheetName].push(loc);
    });

    const detailHeaders = sheetsForDetails;
    const maxRows = Math.max(0, ...Object.values(cellsBySheet).map(cells => cells.length));
    const detailHeaderRow = detailHeaders.map(sheetName => ({
        v: `${sheetName} (${report.summary[sheetName] || 0})`,
        s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" } }
    }));
    detailsAOA.push(detailHeaderRow);
    
    for (let i = 0; i < maxRows; i++) {
        const rowData = detailHeaders.map(sheetName => {
            const loc = cellsBySheet[sheetName]?.[i];
            if (loc) {
                const contextVal = loc.contextValue;
                const contextString = (contextVal !== null && contextVal !== undefined) ? String(contextVal).trim() : '';
                
                let displayValue = loc.address;

                if (reportOptions?.contextColumnForCompact && contextString) {
                    displayValue = `${loc.address} (${contextString})`;
                }

                const cellObject: XLSX.CellObject = {
                    v: displayValue,
                    t: 's',
                    l: { Target: `#${sanitizeSheetNameForFormula(sheetName)}!${loc.address}` },
                    s: { font: { color: { rgb: "0000FF" }, underline: true } }
                };
                
                return cellObject;
            }
            return '';
        });
        detailsAOA.push(rowData);
    }
    
    const detailsSheet = XLSX.utils.aoa_to_sheet(detailsAOA);
    if (!detailsSheet['!merges']) detailsSheet['!merges'] = [];
    const lastColForTitleMerge = detailHeaderRow.length > 0 ? detailHeaderRow.length - 1 : 0;
    if (lastColForTitleMerge > 0) {
        detailsSheet['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: lastColForTitleMerge } });
    }
    
    const colWidths = detailHeaderRow.map((headerCell) => {
        const sheetName = String(headerCell.v || '').split(' (')[0];
        const cellsForSheet = cellsBySheet[sheetName] || [];
        const maxLength = cellsForSheet.reduce((max, loc) => {
            const contextString = loc.contextValue ? ` (${loc.contextValue})` : '';
            const fullLength = loc.address.length + contextString.length;
            return Math.max(max, fullLength);
        }, String(headerCell.v || '').length);

        return { wch: Math.max(25, maxLength + 5) };
    });
    detailsSheet['!cols'] = colWidths;
    
    return { summarySheet, detailsSheet };
}

/**
 * Creates new worksheets for a pivot-style summary report of empty cells.
 * @param report The report object containing all found empty cells.
 * @param reportOptions Configuration for the report, including key and context columns.
 * @returns An object containing the summary and details worksheets.
 */
function createSummaryReportSheets(
    report: EmptyCellReport,
    reportOptions: {
        summaryKeyColumn: string;
        summaryContextColumn: string;
        blankKeyLabel?: string;
    }
): { summarySheet: XLSX.WorkSheet, detailsSheet: XLSX.WorkSheet } {
    const blankLabel = reportOptions.blankKeyLabel || '(Blank)';
    
    // 1. Create the detailed list sheet first (Pivot Data)
    const detailHeaders = [
        'Sheet', 'Row', 'Cell/Row',
        reportOptions.summaryKeyColumn,
        reportOptions.summaryContextColumn
    ].map(h => ({ v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" } } }));
    
    const detailsAOA: any[][] = [
        [{ v: 'Empty Cell Pivot Data', s: { font: { bold: true, sz: 16 } } }],
        [],
        detailHeaders
    ];
    report.locations.forEach(loc => {
        detailsAOA.push([
            loc.sheetName,
            { v: loc.row, t: 'n', l: { Target: `#${sanitizeSheetNameForFormula(loc.sheetName)}!A${loc.row}` } },
            loc.address,
            loc.keyColumnValue,
            loc.contextColumnForSummaryValue
        ]);
    });
    const detailsSheet = XLSX.utils.aoa_to_sheet(detailsAOA, { cellDates: true });
    detailsSheet['!cols'] = [{wch: 20}, {wch: 10}, {wch: 15}, {wch: 25}, {wch: 25}];
    if (detailsAOA.length > 3) {
        detailsSheet['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 2, c: 0 }, e: { r: detailsAOA.length - 1, c: detailHeaders.length - 1 } }) };
    }
    if (!detailsSheet['!merges']) detailsSheet['!merges'] = [];
    detailsSheet['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: detailHeaders.length - 1 }});


    // 2. Create the pivot-style summary sheet
    const summaryData: { [key: string]: { [context: string]: number } } = {};
    report.locations.forEach(loc => {
        const key = String(loc.keyColumnValue ?? blankLabel);
        const context = String(loc.contextColumnForSummaryValue ?? blankLabel);
        if (!summaryData[key]) summaryData[key] = {};
        summaryData[key][context] = (summaryData[key][context] || 0) + 1;
    });

    const summaryHeaders = [
        reportOptions.summaryKeyColumn,
        reportOptions.summaryContextColumn,
        'Empty Cell Count'
    ].map(h => ({ v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" } } }));

    const summaryAOA: any[][] = [
        [{ v: 'Empty Cell Pivot Summary', s: { font: { bold: true, sz: 16 } } }],
        [],
        summaryHeaders
    ];
    
    const sortedKeys = Object.keys(summaryData).sort();
    sortedKeys.forEach(key => {
        const sortedContexts = Object.keys(summaryData[key]).sort();
        sortedContexts.forEach(context => {
            summaryAOA.push([
                key,
                context,
                summaryData[key][context]
            ]);
        });
    });

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryAOA);
    summarySheet['!cols'] = [{wch: 30}, {wch: 30}, {wch: 20}];
    if (summaryAOA.length > 3) {
        summarySheet['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 2, c: 0 }, e: { r: summaryAOA.length - 1, c: summaryHeaders.length - 1 } }) };
    }
    if (!summarySheet['!merges']) summarySheet['!merges'] = [];
    summarySheet['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: summaryHeaders.length - 1 }});


    return { summarySheet, detailsSheet };
}
