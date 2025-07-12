
import * as XLSX from 'xlsx-js-style';
import type { AggregationResult, HeaderFormatOptions, SummaryConfig, UpdateResult } from './excel-types';
import { getUniqueSheetName, sanitizeSheetNameForFormula } from './excel-helpers';
import { applyFormatting } from './excel-aggregator-actions';


function _createDataSourceSheet(
    workbook: XLSX.WorkBook,
    results: AggregationResult,
    config: SummaryConfig,
    baseSheetName: string
): { name: string, rows: number } {
    const dataSourceSheetName = getUniqueSheetName(workbook, baseSheetName);
    console.log(`Creating data source sheet: "${dataSourceSheetName}"`);
    const flatDataAOA: any[][] = [['Sheet', 'Key', 'Count']];

    const allSheets = results.processedSheetNames;
    const allKeys = new Set(results.reportingKeys);
    if (results.blankCounts && results.blankCounts.total > 0 && config.blankLabel) {
        allKeys.add(config.blankLabel);
    }

    for (const sheetName of allSheets) {
        for (const key of allKeys) {
            let count = 0;
             if (key === config.blankLabel) {
                count = results.blankCounts?.perSheet[sheetName] || 0;
             } else {
                count = results.perSheetCounts[sheetName]?.[key] || 0;
             }
            if (count > 0) {
                 flatDataAOA.push([sheetName, key, count]);
            }
        }
    }

    const dataWS = XLSX.utils.aoa_to_sheet(flatDataAOA, {cellDates: true});
    dataWS['!cols'] = [{wch: 30}, {wch: 30}, {wch: 15}];
    XLSX.utils.book_append_sheet(workbook, dataWS, dataSourceSheetName);

    if (!workbook.Workbook) workbook.Workbook = { Sheets: [] };
    if (!workbook.Workbook.Sheets) workbook.Workbook.Sheets = [];
    
    const sheetProps = workbook.Workbook.Sheets.find(s => s.name === dataSourceSheetName);
    if (sheetProps) {
        sheetProps.Hidden = 1;
    } else {
        workbook.Workbook.Sheets.push({ name: dataSourceSheetName, Hidden: 1 });
    }
    
    return { name: dataSourceSheetName, rows: flatDataAOA.length };
}


function _generateSummarySheet(
    title: string,
    results: AggregationResult, 
    config: SummaryConfig,
    useFormulas: boolean,
    pivotDataSourceSheetName: string,
    dataSourceRows: number
): XLSX.WorkSheet {
    const reportLayout = config.reportLayout || 'sheetsAsRows';
    const hiddenColsSet = new Set(
        config.columnsToHide?.split(',').map(s => s.trim().toLowerCase()).filter(Boolean) || []
    );

    const sheetData: any[][] = [];
    sheetData.push([{ v: title, s: { font: { bold: true, sz: 16 } } }]);
    sheetData.push([]);

    const dataStartRow = 2;
    const dataEndRow = dataSourceRows > 1 ? dataSourceRows : 2;
    
    const sanitizedSheetName = sanitizeSheetNameForFormula(pivotDataSourceSheetName);
    const formulaRange = {
        sheet: `${sanitizedSheetName}!`,
        countCol: `$C$${dataStartRow}:$C$${dataEndRow}`,
        sheetCol: `$A$${dataStartRow}:$A$${dataEndRow}`,
        keyCol: `$B$${dataStartRow}:$B$${dataEndRow}`,
    };

    if (reportLayout === 'keysAsRows') {
        const allKeys = new Set(results.reportingKeys);
        if (results.blankCounts && results.blankCounts.total > 0 && config.blankLabel) {
            allKeys.add(config.blankLabel);
        }
        const allSheets = new Set(results.processedSheetNames);

        const sortedUniqueKeys = Array.from(allKeys).filter(k => !hiddenColsSet.has(k.toLowerCase())).sort((a,b)=>a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
        const sortedUniqueSheets = Array.from(allSheets).filter(s => !hiddenColsSet.has(s.toLowerCase())).sort((a,b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

        const summaryHeaders = ['Key Name', ...sortedUniqueSheets, 'Total'];
        sheetData.push(summaryHeaders.map(h => ({ v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" as const } } })));

        sortedUniqueKeys.forEach((key, keyIndex) => {
            const rowData: any[] = [key];
            const keyRowInSheet = keyIndex + 4;

            sortedUniqueSheets.forEach((sheetName) => {
                if (useFormulas && pivotDataSourceSheetName) {
                    const sanitizedSheetNameForFormula = sanitizeSheetNameForFormula(sheetName);
                    const formula = `SUMPRODUCT((${formulaRange.sheet}${formulaRange.sheetCol}="${sanitizedSheetNameForFormula.replace(/'/g, "")}")*(${formulaRange.sheet}${formulaRange.keyCol}=$A${keyRowInSheet}),${formulaRange.sheet}${formulaRange.countCol})`;
                    rowData.push({ t: 'n', f: formula });
                } else {
                     let count = 0;
                     if (key === config.blankLabel) {
                        count = results.blankCounts?.perSheet[sheetName] || 0;
                     } else {
                        count = results.perSheetCounts[sheetName]?.[key] || 0;
                     }
                     rowData.push(count);
                }
            });
            sheetData.push(rowData);
        });
        
        sheetData.slice(3).forEach((row, index) => {
            if (index === 0) return;
            const dataRowIndex = index + 3;
            const totalFormula = `SUM(B${dataRowIndex}:${XLSX.utils.encode_col(sortedUniqueSheets.length)}${dataRowIndex})`;
            row.push({ t: 'n', f: totalFormula });
        });

        const totalRow: any[] = [{v: 'Total', s: {font: {bold: true}}}];
        for (let i = 0; i < sortedUniqueSheets.length + 1; i++) {
            const colLetter = XLSX.utils.encode_col(i + 1);
            const totalFormula = `SUM(${colLetter}4:${colLetter}${sheetData.length})`;
            totalRow.push({ t: 'n', f: totalFormula, s: {font: {bold: true}} });
        }
        sheetData.push(totalRow);
        
        const summaryWS = XLSX.utils.aoa_to_sheet(sheetData, {cellDates: true});

        if (config.autoSizeColumns) {
            const keyColWidth = Math.max(20, ...sortedUniqueKeys.map(k => k.length + 2));
            const sheetColWidths = sortedUniqueSheets.map(s => ({ wch: Math.max(12, s.length + 2) }));
            summaryWS['!cols'] = [{ wch: keyColWidth }, ...sheetColWidths, { wch: 12 }];
        }
        if (sheetData.length > 3) {
            summaryWS['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 2, c: 0 }, e: { r: sheetData.length - 2, c: summaryHeaders.length - 1 } }) };
        }
        if (!summaryWS['!merges']) summaryWS['!merges'] = [];
        summaryWS['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: summaryHeaders.length - 1 } });
        return summaryWS;

    } else {
        const sortedSheetNames = [...results.processedSheetNames]
            .filter(name => !hiddenColsSet.has(name.toLowerCase()))
            .sort((a,b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

        const allKeys = new Set(results.reportingKeys);
        if (results.blankCounts && results.blankCounts.total > 0 && config.blankLabel) {
            allKeys.add(config.blankLabel);
        }
        const sortedKeys = Array.from(allKeys)
            .filter(key => !hiddenColsSet.has(key.toLowerCase()))
            .sort((a,b)=>a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
        
        const headers = ['Sheet Name', ...sortedKeys, 'Total'];
        sheetData.push(headers.map(h => ({ v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" as const } } })));

        sortedSheetNames.forEach((sheetName, sheetIdx) => {
            const row: any[] = [sheetName];
            const sheetRowInSheet = sheetIdx + 4;

            sortedKeys.forEach((key, keyIndex) => {
                const keyColInSheet = XLSX.utils.encode_col(keyIndex + 1);
                if (useFormulas && pivotDataSourceSheetName) {
                    const formula = `SUMPRODUCT((${formulaRange.sheet}${formulaRange.sheetCol}=$A${sheetRowInSheet})*(${formulaRange.sheet}${formulaRange.keyCol}=${keyColInSheet}$3),${formulaRange.sheet}${formulaRange.countCol})`;
                    row.push({ t: 'n', f: formula });
                } else {
                    let count = 0;
                     if (key === config.blankLabel) {
                        count = results.blankCounts?.perSheet[sheetName] || 0;
                     } else {
                        count = results.perSheetCounts[sheetName]?.[key] || 0;
                     }
                    row.push(count);
                }
            });
            
            const dataRowForFormula = sheetIdx + 4;
            const totalFormula = `SUM(B${dataRowForFormula}:${XLSX.utils.encode_col(sortedKeys.length)}${dataRowForFormula})`;
            row.push({ t: 'n', f: totalFormula });
            sheetData.push(row);
        });
        
        if (sheetData.length > 3) {
            const totalRowFormulaCells = headers.slice(1).map((_, keyIndex) => ({ t: 'n', f: `SUM(${XLSX.utils.encode_col(keyIndex + 1)}4:${XLSX.utils.encode_col(keyIndex + 1)}${sheetData.length})` }));
            const totalRow = [{v:'Total', s:{font:{bold: true}}}, ...totalRowFormulaCells];
            sheetData.push(totalRow);
        }
        
        const summaryWS = XLSX.utils.aoa_to_sheet(sheetData, {cellDates: true});
        
        if (config.blankRowFormatting && config.blankLabel) {
            const blankKeyIndex = sortedKeys.indexOf(config.blankLabel);
            if (blankKeyIndex !== -1) {
                const blankColIndex = blankKeyIndex + 1;
                for (let R = 3; R < sheetData.length; R++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: R, c: blankColIndex });
                    const cell = summaryWS[cellAddress];
                    if (cell) applyFormatting(cell, config.blankRowFormatting);
                }
            }
        }

        if (config.autoSizeColumns) {
            summaryWS['!cols'] = [
                { wch: 30 },
                ...sortedKeys.map(k => ({ wch: Math.max(12, k.length + 2) })),
                { wch: 12 }
            ];
        }

        if (sheetData.length > 3) {
            summaryWS['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 2, c: 0 }, e: { r: sheetData.length - 2, c: headers.length - 1 } }) };
        }
        if (!summaryWS['!merges']) summaryWS['!merges'] = [];
        summaryWS['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } });
        return summaryWS;
    }
}


export function createAggregationReportWorkbook(
    originalResults: AggregationResult,
    modifiedResults: AggregationResult, 
    config: SummaryConfig, 
    editedKeyToOriginalsMap?: Map<string, string[]>
): XLSX.WorkBook {
    console.log('Creating new aggregation report workbook with pre- and post-edit summaries.');
    const wb = XLSX.utils.book_new();
    addAggregationReportSheetsToWorkbook(wb, originalResults, modifiedResults, config, editedKeyToOriginalsMap);
    return wb;
}

export function addAggregationReportSheetsToWorkbook(
    workbook: XLSX.WorkBook,
    originalResults: AggregationResult,
    modifiedResults: AggregationResult,
    config: SummaryConfig,
    editedKeyToOriginalsMap?: Map<string, string[]>
) {
    console.log("Adding aggregation report sheets to workbook...");

    const preEditSheet = _generateSummarySheet(
        "Pre-Edit Summary",
        originalResults,
        config,
        false, 
        '', 0
    );
    const preEditSheetName = getUniqueSheetName(workbook, "Pre-Edit Summary");
    XLSX.utils.book_append_sheet(workbook, preEditSheet, preEditSheetName);
    
    const {name: dataSourceName, rows: dataSourceRows} = _createDataSourceSheet(workbook, modifiedResults, config, 'Aggregation_Data_Source');
    const postEditSheet = _generateSummarySheet(
        "Post-Edit Summary (Final)",
        modifiedResults,
        config,
        true,
        dataSourceName,
        dataSourceRows
    );
    const postEditSheetName = getUniqueSheetName(workbook, "Post-Edit Summary (Final)");
    XLSX.utils.book_append_sheet(workbook, postEditSheet, postEditSheetName);
    
    if (editedKeyToOriginalsMap) {
        addKeyMappingsSheetToWorkbook(workbook, originalResults, modifiedResults, editedKeyToOriginalsMap);
    }
    
    const blankDetailsSheetName = addBlankDetailsSheet(workbook, modifiedResults);
    
    const allSheetNamesInWorkbook = [...workbook.SheetNames];

    const visibleReportSheets = new Set([
        postEditSheetName,
        preEditSheetName,
        (editedKeyToOriginalsMap ? "Key Mappings" : undefined),
        blankDetailsSheetName
    ].filter(Boolean) as string[]);
    
    const originalSheets = allSheetNamesInWorkbook.filter(
        name => !visibleReportSheets.has(name) && name !== dataSourceName
    );

    const finalSheetOrder = [
        ...Array.from(visibleReportSheets),
        ...originalSheets,
        dataSourceName
    ].filter(Boolean) as string[];

    workbook.SheetNames = finalSheetOrder;

    if (workbook.Workbook && workbook.Workbook.Sheets) {
        const sheetPropsMap = new Map(workbook.Workbook.Sheets.map(s => [s.name, s]));
        const reorderedWbSheets = finalSheetOrder
            .map(name => sheetPropsMap.get(name))
            .filter(Boolean) as any[];
        workbook.Workbook.Sheets = reorderedWbSheets;
    }
}


export function addUpdateReportSheetToWorkbook(workbook: XLSX.WorkBook, updateResult: UpdateResult, chunkSize: number) {
    console.log("Adding update report sheet to workbook.");
    try {
        const baseSheetName = "Update_Report";
        const maxRowsPerSheet = (chunkSize && chunkSize > 0) ? chunkSize : 100000;
        const { details, summary } = updateResult;

        const summaryData: any[][] = [
        [{ v: 'Update Operation Summary', s: { font: { bold: true, sz: 16 } } }],
        [],
        ['Total Cells Updated:', summary.totalCellsUpdated],
        ['Sheets Affected:', summary.sheetsUpdated.join(', ')]
        ];
        
        const allRowDataHeaders = new Set<string>();
        if(details.length > 0) {
            details.forEach(d => {
                if (d.rowData) {
                    Object.keys(d.rowData).forEach(h => allRowDataHeaders.add(h));
                }
            });
        }
        const sortedRowDataHeaders = Array.from(allRowDataHeaders).sort();
        
        const headers = [
            'Sheet Name', 'Row', 'Cell', 'Original Value', 'New Value', 'Key Used', 'Triggering Column', 'Triggering Value', ...sortedRowDataHeaders
        ].map(h => ({v: h, s: {font:{bold: true}, fill: {fgColor: {rgb:"EAEAEA"}, patternType: "solid" as const}}}));
        
        const totalDetails = details.length;
        const numSheets = Math.ceil(totalDetails / (maxRowsPerSheet - 1));

        for (let i = 0; i < numSheets; i++) {
            const sheetName = getUniqueSheetName(workbook, numSheets > 1 ? `${baseSheetName}_${i + 1}` : baseSheetName);
            const chunkStart = i * (maxRowsPerSheet - 1);
            const chunkEnd = Math.min(chunkStart + (maxRowsPerSheet - 1), totalDetails);
            const chunk = details.slice(chunkStart, chunkEnd);

            const rowData = chunk.map(d => {
                const baseData = [
                    d.sheetName,
                    d.rowNumber,
                    { v: d.cellAddress, l: { Target: `#${sanitizeSheetNameForFormula(d.sheetName)}!${d.cellAddress}` }, s: { font: { color: { rgb: "0000FF" }, underline: true } } },
                    d.originalValue,
                    d.newValue,
                    d.keyUsed,
                    d.triggerColumn,
                    d.triggerValue,
                ];
                const extendedData = sortedRowDataHeaders.map(h => d.rowData ? (d.rowData[h] ?? null) : null);
                return [...baseData, ...extendedData];
            });

            let finalAOA = [headers, ...rowData];
            if (i === 0) {
                finalAOA = [...summaryData, [], ...finalAOA];
            }

            const ws = XLSX.utils.aoa_to_sheet(finalAOA, { cellDates: true });
            ws['!cols'] = Array.from({length: headers.length}, () => ({wch: 25}));
            XLSX.utils.book_append_sheet(workbook, ws, sheetName);
        }
    } catch (error) {
        console.error("Error creating update report sheet:", error);
    }
}

function addKeyMappingsSheetToWorkbook(
    workbook: XLSX.WorkBook,
    originalResults: AggregationResult,
    modifiedResults: AggregationResult,
    editedKeyToOriginalsMap: Map<string, string[]>
) {
    console.log("Adding key mappings sheet.");
    const sheetName = getUniqueSheetName(workbook, "Key Mappings");
    const aoa: any[][] = [];
    const headerStyle = { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" as const } };

    aoa.push([
        { v: "Final Key", s: headerStyle },
        { v: "Final Count", s: headerStyle },
        { v: "Contributing Search Terms (Values)", s: headerStyle },
        { v: "Original Mapped Key(s)", s: headerStyle }
    ]);

    const originalToFinalKeyMap = new Map<string, string>();
    editedKeyToOriginalsMap.forEach((originalKeys, finalKey) => {
        originalKeys.forEach(originalKey => {
            originalToFinalKeyMap.set(originalKey, finalKey);
        });
    });

    const finalKeyToSearchTerms = new Map<string, Set<string>>();
    if (originalResults.valueToKeyMap) {
        for (const [value, originalKey] of originalResults.valueToKeyMap.entries()) {
            const finalKey = originalToFinalKeyMap.get(originalKey);
            if (finalKey) {
                if (!finalKeyToSearchTerms.has(finalKey)) {
                    finalKeyToSearchTerms.set(finalKey, new Set());
                }
                finalKeyToSearchTerms.get(finalKey)!.add(value);
            }
        }
    }
    
    const sortedFinalKeys = Array.from(editedKeyToOriginalsMap.keys()).sort((a,b) => a.localeCompare(b, undefined, {numeric: true}));

    sortedFinalKeys.forEach(finalKey => {
        const originalKeys = editedKeyToOriginalsMap.get(finalKey) || [];
        const finalCount = modifiedResults.totalCounts[finalKey] || 0;
        const contributingSearchTerms = finalKeyToSearchTerms.get(finalKey) || new Set();
        
        aoa.push([
            finalKey,
            finalCount,
            Array.from(contributingSearchTerms).sort().join(', '),
            originalKeys.join(', ')
        ]);
    });
    
    if(aoa.length > 1) {
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!cols'] = [{ wch: 40 }, { wch: 15 }, { wch: 80 }, { wch: 60 }];
        XLSX.utils.book_append_sheet(workbook, ws, sheetName);
    }
}

function addBlankDetailsSheet(
    workbook: XLSX.WorkBook,
    results: AggregationResult
): string | undefined {
    if (!results.blankDetails || results.blankDetails.length === 0) {
        return undefined;
    }
    console.log(`Adding blank details sheet with ${results.blankDetails.length} rows.`);

    const sheetName = getUniqueSheetName(workbook, 'Blank_Cell_Details');
    
    const allHeaders = new Set<string>();
    results.blankDetails.forEach(detail => {
        if(detail.rowData) {
            Object.keys(detail.rowData).forEach(key => allHeaders.add(key));
        }
    });
    const sortedHeaders = Array.from(allHeaders).sort();
    
    const headerRow = [
        'Sheet Name', 
        'Row Number', 
        ...sortedHeaders
    ].map(h => ({ v: h, s: { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" as const } } }));

    const aoa: any[][] = [headerRow];

    results.blankDetails.forEach(detail => {
        const rowDataArray = sortedHeaders.map(header => detail.rowData ? (detail.rowData[header] ?? null) : null);
        aoa.push([
            detail.sheetName,
            { v: detail.rowNumber, t: 'n', l: { Target: `#${sanitizeSheetNameForFormula(detail.sheetName)}!A${detail.rowNumber}` }, s: { font: { color: { rgb: "0000FF" }, underline: true } } },
            ...rowDataArray
        ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
    
    ws['!cols'] = [
        { wch: 25 },
        { wch: 10 },
        ...Array.from({ length: sortedHeaders.length }, () => ({ wch: 20 }))
    ];

    if (aoa.length > 1) {
        ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: headerRow.length - 1 } }) };
    }

    XLSX.utils.book_append_sheet(workbook, ws, sheetName);
    return sheetName;
}

