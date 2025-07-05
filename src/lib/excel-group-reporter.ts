
import * as XLSX from 'xlsx-js-style';
import type { AggregationResult, HeaderFormatOptions, SummaryConfig, TableFormattingOptions } from './excel-types';
import { getUniqueSheetName, sanitizeSheetNameForFormula } from './excel-helpers';
import { applyFormatting } from './excel-aggregator-actions';


/**
 * Parses the user-provided group mapping string.
 * @param mappingString The raw string from the textarea (e.g., "Group A: Key1, Key2\nGroup B: Key3").
 * @returns A map where keys are group names and values are arrays of key names.
 */
function parseGroupMappings(mappingString: string): Map<string, string[]> {
    const mappings = new Map<string, string[]>();
    if (!mappingString) return mappings;

    const lines = mappingString.split('\n').filter(line => line.trim() !== '');
    for (const line of lines) {
        const parts = line.split(':');
        if (parts.length === 2) {
            const groupName = parts[0].trim();
            const keys = parts[1].split(',').map(key => key.trim()).filter(Boolean);
            if (groupName && keys.length > 0) {
                mappings.set(groupName, keys);
            }
        }
    }
    return mappings;
}

/**
 * Applies a consistent border and fill style to a range of cells in a worksheet.
 * It avoids overwriting existing fill colors to preserve special formatting.
 * @param worksheet The worksheet to modify.
 * @param startRow The 0-indexed starting row.
 * @param endRow The 0-indexed ending row.
 * @param startCol The 0-indexed starting column.
 * @param endCol The 0-indexed ending column.
 * @param options The styling options for the table.
 */
function _applyTableStyles(
    worksheet: XLSX.WorkSheet,
    startRow: number,
    endRow: number,
    startCol: number,
    endCol: number,
    options: TableFormattingOptions
) {
    if (endRow < startRow || endCol < startCol) return;

    for (let R = startRow; R <= endRow; R++) {
        for (let C = startCol; C <= endCol; C++) {
            const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cellAddr] || { t: 'z' };
            if (!cell.s) cell.s = {};
            
            // Apply fill color only if one doesn't already exist from a more specific style
            if (options.fillColor && (!cell.s.fill || !cell.s.fill.fgColor)) {
                cell.s.fill = {
                    patternType: 'solid',
                    fgColor: { rgb: options.fillColor.replace('#', '') },
                };
            }

            // Apply borders to all cells in the range
            const borderStyle = options.borderStyle || 'thin';
            const borderColor = options.borderColor ? { rgb: options.borderColor.replace('#', '') } : { auto: 1 };
            cell.s.border = {
                top: { style: borderStyle, color: borderColor },
                bottom: { style: borderStyle, color: borderColor },
                left: { style: borderStyle, color: borderColor },
                right: { style: borderStyle, color: borderColor },
            };
            
            worksheet[cellAddr] = cell;
        }
    }
}


/**
 * Generates the Array-of-Arrays data structure for an in-sheet summary.
 * @param results The final aggregation results.
 * @param workbook The original workbook, needed to calculate formula ranges.
 * @param sheetName The name of the sheet for which to generate the summary.
 * @param config The overall summary configuration.
 * @param options Configuration for how to generate this specific summary.
 * @returns A 2D array of cell data for `sheet_add_aoa`.
 */
function _generateInSheetSummaryAOA(
    results: AggregationResult,
    workbook: XLSX.WorkBook,
    sheetName: string,
    config: SummaryConfig,
    options: { dataSource: 'reportingScope' | 'localSheet', showOnlyLocalKeys: boolean }
): any[][] {
    const dataToInsert: any[][] = [];
    const { dataSource, showOnlyLocalKeys } = options;

    const dynamicTitle = results.sheetTitles?.[sheetName] || config.inSheetSummaryTitle || 'Summary';
    const headerCell: XLSX.CellObject = { v: dynamicTitle, t: 's' };
    applyFormatting(headerCell, config.headerFormatting);
    const countHeaderCell: XLSX.CellObject = { v: 'Count', t: 's' };
    applyFormatting(countHeaderCell, config.headerFormatting);
    dataToInsert.push([headerCell, countHeaderCell]);
    
    const getKeysForSheet = () => {
        const allKeysInReport = new Set(results.reportingKeys);
        if (results.blankCounts && results.blankCounts.total > 0 && config.blankCountLabel) {
            allKeysInReport.add(config.blankCountLabel);
        }

        if (dataSource === 'reportingScope' || !showOnlyLocalKeys) {
             return Array.from(allKeysInReport);
        }

        const localKeys = new Set<string>();
        const sheetCounts = results.perSheetCounts[sheetName] || {};
        for (const key in sheetCounts) {
            if (sheetCounts[key] > 0) localKeys.add(key);
        }
        if ((results.blankCounts?.perSheet[sheetName] || 0) > 0 && config.showBlanksInInSheetSummary && config.blankCountLabel) {
            localKeys.add(config.blankCountLabel);
        }
        return Array.from(localKeys);
    };

    const keysToDisplay = getKeysForSheet().sort((a,b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

    let dataRowsAdded = 0;
    keysToDisplay.forEach(key => {
        const sheetCounts = results.perSheetCounts[sheetName] || {};
        const blankCount = results.blankCounts?.perSheet[sheetName] || 0;
        const staticCount = sheetCounts[key] || (key === config.blankCountLabel ? blankCount : 0);
        
        if (key === config.blankCountLabel && !config.showBlanksInInSheetSummary) {
            return;
        }

        dataRowsAdded++;
        const valueCell: XLSX.CellObject = { t: 'n', v: staticCount };
        applyFormatting(valueCell, { horizontalAlignment: 'right' });

        const keyCell: XLSX.CellObject = { v: key, t: 's' };
        
        if (key === config.blankCountLabel) {
            applyFormatting(keyCell, config.blankRowFormatting);
            applyFormatting(valueCell, config.blankRowFormatting);
        }

        dataToInsert.push([keyCell, valueCell]);
    });

    if (dataRowsAdded > 0 && config.totalRowFormatting) {
        const totalCount = Object.values(results.perSheetCounts[sheetName] || {}).reduce((s, c) => s + c, 0);
        const blankCount = results.blankCounts?.perSheet[sheetName] || 0;
        const totalValue = totalCount + blankCount;

        const totalValueCell: XLSX.CellObject = { t: 'n', v: totalValue };
        const totalLabelCell: XLSX.CellObject = {v:'Total', t: 's'};
        applyFormatting(totalLabelCell, config.totalRowFormatting);
        applyFormatting(totalValueCell, config.totalRowFormatting);
        dataToInsert.push([totalLabelCell, totalValueCell]);
    }
    
    return dataToInsert;
}


/**
 * Creates a new workbook with a grouped summary report and a compiled list of in-sheet summaries.
 * @param aggregationResult The final, post-edit aggregation results.
 * @param workbook The original workbook, needed for dynamic title lookups.
 * @param groupMappingString The raw string defining how to group the final keys.
 * @param sheetsToReportOn The list of original sheet names to generate summaries for.
 * @param summaryConfig The configuration for generating summaries.
 * @param inSheetOptions The options specific to how in-sheet summaries should behave.
 * @param groupReportHeaders Optional custom text for the group report headers.
 * @returns A new XLSX.WorkBook object containing the grouped report and compiled summaries.
 */
export function createGroupReportWorkbook(
    aggregationResult: AggregationResult,
    workbook: XLSX.WorkBook,
    groupMappingString: string,
    sheetsToReportOn: string[],
    summaryConfig: SummaryConfig,
    inSheetOptions: { dataSource: 'reportingScope' | 'localSheet', showOnlyLocalKeys: boolean },
    groupReportHeaders?: { groupName: string, keyName: string, count: string }
): XLSX.WorkBook {
    const wb = XLSX.utils.book_new();
    const { groupReportSheetTitle, groupReportMultiSourceTitle, groupReportHeaderFormatting, groupReportDescription, tableFormatting } = summaryConfig;
    const groupReportSheetName = getUniqueSheetName(wb, groupReportSheetTitle || "Group_Report");

    const groupMappings = parseGroupMappings(groupMappingString);
    const reportData: { groupName: string; totalCount: number; keys: { name: string, count: number }[] }[] = [];
    const unmappedKeys = new Set(Object.keys(aggregationResult.totalCounts));

    for (const [groupName, keysInGroup] of groupMappings.entries()) {
        let groupTotal = 0;
        const validKeysInGroup: { name: string, count: number }[] = [];

        keysInGroup.forEach(key => {
            if (aggregationResult.totalCounts[key] !== undefined) {
                const count = aggregationResult.totalCounts[key];
                groupTotal += count;
                validKeysInGroup.push({ name: key, count });
            }
            unmappedKeys.delete(key);
        });

        if (validKeysInGroup.length > 0) {
            reportData.push({ groupName, totalCount: groupTotal, keys: validKeysInGroup });
        }
    }

    const unmappedData: { key: string, count: number }[] = [];
    unmappedKeys.forEach(key => {
        if(aggregationResult.totalCounts[key] > 0) {
            unmappedData.push({ key: key, count: aggregationResult.totalCounts[key] });
        }
    });

    const groupAoa: any[][] = [];
    const groupTotalStyle = { font: { bold: true }, fill: { fgColor: { rgb: "F2F2F2" }, patternType: "solid" as const } };
    
    // Determine main report title
    const reportedSheetTitles = new Set<string>();
    if (summaryConfig.summaryTitleCell) {
        sheetsToReportOn.forEach(sheetName => {
            const ws = workbook.Sheets[sheetName];
            const titleCell = ws?.[summaryConfig.summaryTitleCell!];
            if (titleCell?.v) {
                reportedSheetTitles.add(String(titleCell.v));
            }
        });
    }

    let mainReportTitle: string;
    if (reportedSheetTitles.size === 1) {
        mainReportTitle = Array.from(reportedSheetTitles)[0];
    } else if (reportedSheetTitles.size > 1) {
        mainReportTitle = groupReportMultiSourceTitle || "Grouped Summary Report (Multiple Sources)";
    } else {
        mainReportTitle = groupReportSheetTitle || "Grouped Summary Report";
    }

    groupAoa.push([{ v: mainReportTitle, s: { font: { bold: true, sz: 16 } } }]);
    const merges: XLSX.Range[] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];

    if (groupReportDescription) {
        groupAoa.push([{ v: groupReportDescription, s: { font: { italic: true, sz: 10 } } }]);
        merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: 2 } });
    }

    groupAoa.push([]);

    const hasGroups = reportData.length > 0;
    const hasUnmapped = unmappedData.length > 0;
    
    let tableStartRow = -1;

    if (hasGroups || hasUnmapped) {
        const headerCell1 = { v: groupReportHeaders?.groupName || "Group Name", s: {} };
        const headerCell2 = { v: groupReportHeaders?.keyName || "Key Name", s: {} };
        const headerCell3 = { v: groupReportHeaders?.count || "Count", s: {} };
        
        applyFormatting(headerCell1, groupReportHeaderFormatting);
        applyFormatting(headerCell2, groupReportHeaderFormatting);
        applyFormatting(headerCell3, groupReportHeaderFormatting);

        tableStartRow = groupAoa.length;
        groupAoa.push([headerCell1, headerCell2, headerCell3]);
    }
    
    if (hasGroups) {
        groupAoa.push([{ v: "User-Defined Groups", s: { font: { bold: true, sz: 14 } } }]);
        merges.push({ s: { r: groupAoa.length - 1, c: 0 }, e: { r: groupAoa.length - 1, c: 2 }});

        reportData.sort((a,b) => a.groupName.localeCompare(b.groupName)).forEach(group => {
            group.keys.sort((a, b) => b.count - a.count);
            const groupStartRow = groupAoa.length;
            group.keys.forEach((keyData, index) => {
                const groupCell = index === 0 ? {v: group.groupName, s: {alignment: {vertical: 'top' as const}}} : '';
                groupAoa.push([groupCell, keyData.name, keyData.count]);
            });

            if (group.keys.length > 1) {
                merges.push({ s: { r: groupStartRow, c: 0 }, e: { r: groupStartRow + group.keys.length - 1, c: 0 } });
            }

            groupAoa.push([
                { v: `${group.groupName} Total`, s: groupTotalStyle },
                '',
                { v: group.totalCount, s: groupTotalStyle }
            ]);
            merges.push({ s: { r: groupAoa.length - 1, c: 0 }, e: { r: groupAoa.length - 1, c: 1 } });
        });
    }
    
    if (hasUnmapped) {
        if(hasGroups) groupAoa.push([]);
        
        groupAoa.push([{ v: "Unmapped Keys", s: { font: { bold: true, sz: 14 } } }]);
        merges.push({ s: { r: groupAoa.length - 1, c: 0 }, e: { r: groupAoa.length - 1, c: 2 } });
        
        unmappedData.sort((a, b) => b.count - a.count);

        if (unmappedData.length > 0) {
            const unmappedTotal = unmappedData.reduce((sum, item) => sum + item.count, 0);
            unmappedData.forEach(item => {
                groupAoa.push(['Unmapped', item.key, item.count]);
            });

            groupAoa.push([
                { v: "Unmapped Total", s: groupTotalStyle },
                '',
                { v: unmappedTotal, s: groupTotalStyle }
            ]);
            merges.push({ s: { r: groupAoa.length - 1, c: 0 }, e: { r: groupAoa.length - 1, c: 1 } });
        }
    }
    
    let tableEndRow = groupAoa.length - 1;
    
    if (hasGroups || hasUnmapped) {
        groupAoa.push([]); // Spacer row
        const unmappedTotal = unmappedData.reduce((sum, item) => sum + item.count, 0);
        const groupedTotal = reportData.reduce((sum, group) => sum + group.totalCount, 0);
        const grandTotal = groupedTotal + unmappedTotal;

        const grandTotalStyle = { font: { bold: true, sz: 12 }, fill: { fgColor: { rgb: "D9D9D9" }, patternType: "solid" as const } };
        groupAoa.push([
            { v: "Grand Total", s: grandTotalStyle },
            '',
            { v: grandTotal, s: grandTotalStyle }
        ]);
        merges.push({ s: { r: groupAoa.length - 1, c: 0 }, e: { r: groupAoa.length - 1, c: 1 } });
    }
    
    const groupWs = XLSX.utils.aoa_to_sheet(groupAoa);
    
    // Apply table styles to the main group report
    if (tableFormatting && tableStartRow !== -1) {
        _applyTableStyles(
            groupWs,
            tableStartRow,
            tableEndRow,
            0,
            2,
            tableFormatting
        );
    }
    
    groupWs['!merges'] = merges;
    groupWs['!cols'] = [{ wch: 35 }, { wch: 45 }, { wch: 15 }];
    
    XLSX.utils.book_append_sheet(wb, groupWs, groupReportSheetName);

    // --- Compiled In-Sheet Summaries ---
    const compiledSheetName = getUniqueSheetName(wb, "Compiled_Summaries");
    const compiledWS = XLSX.utils.aoa_to_sheet([[]]);
    let currentCompiledRow = 0;
    
    sheetsToReportOn.forEach(sheetName => {
        const summaryAOA = _generateInSheetSummaryAOA(
            aggregationResult,
            workbook,
            sheetName,
            summaryConfig,
            inSheetOptions
        );

        if (summaryAOA.length > 1) {
            const title = aggregationResult.sheetTitles?.[sheetName] || sheetName;
            const titleCell = { v: `Summary for: ${title}`, s: { font: { bold: true, sz: 14 } } };
            XLSX.utils.sheet_add_aoa(compiledWS, [[titleCell]], { origin: { r: currentCompiledRow, c: 0 } });
            currentCompiledRow += 1;

            XLSX.utils.sheet_add_aoa(compiledWS, summaryAOA, { origin: { r: currentCompiledRow, c: 0 } });
            
            if (tableFormatting) {
                _applyTableStyles(
                    compiledWS,
                    currentCompiledRow,
                    currentCompiledRow + summaryAOA.length - 1,
                    0,
                    1, 
                    tableFormatting
                );
            }

            currentCompiledRow += summaryAOA.length + 2;
        }
    });
    
    if (currentCompiledRow > 0) {
        compiledWS['!cols'] = [{ wch: 40 }, { wch: 15 }];
        const finalRange = XLSX.utils.decode_range(compiledWS['!ref'] || 'A1');
        finalRange.e.c = 1;
        compiledWS['!ref'] = XLSX.utils.encode_range(finalRange);
        XLSX.utils.book_append_sheet(wb, compiledWS, compiledSheetName);
    }
    
    const finalSheetOrder = [groupReportSheetName];
    if (wb.SheetNames.includes(compiledSheetName)) {
        finalSheetOrder.push(compiledSheetName);
    }
    wb.SheetNames = finalSheetOrder;

    return wb;
}
