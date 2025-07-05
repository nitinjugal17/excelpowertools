
import * as XLSX from 'xlsx-js-style';
import { getColumnIndex, getUniqueSheetName } from './excel-helpers';
import type { AggregationType } from './excel-types';

export interface PivotConfig {
    sheetNames: string[];
    rowFields: string[];
    columnFields: string[];
    valueField: string;
    aggregationType: AggregationType;
    headerRow: number;
    outputSheetName: string;
}

function aggregate(values: number[], type: AggregationType): number {
    if (values.length === 0) return 0;
    switch (type) {
        case 'SUM':
            return values.reduce((a, b) => a + b, 0);
        case 'COUNT':
            return values.length;
        case 'AVERAGE':
            return values.reduce((a, b) => a + b, 0) / values.length;
        case 'MIN':
            return Math.min(...values);
        case 'MAX':
            return Math.max(...values);
        default:
            return 0;
    }
}

/**
 * Creates a new workbook with a pivot table generated from the data in the specified sheets.
 * @param workbook The source XLSX.WorkBook object.
 * @param config The configuration for the pivot table.
 * @returns A new XLSX.WorkBook object containing the pivot table.
 */
export function createPivotTableFromWorkbook(
    workbook: XLSX.WorkBook,
    config: PivotConfig
): XLSX.WorkBook {
    const { sheetNames, rowFields, columnFields, valueField, aggregationType, headerRow, outputSheetName } = config;
    const headerRowIndex = headerRow - 1;

    // 1. Data Collection: Combine data from all selected sheets into a single array of objects.
    const allData: Record<string, any>[] = [];
    sheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return;
        const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
        if (aoa.length <= headerRowIndex) return;

        const headers = aoa[headerRowIndex].map(h => String(h || ''));
        const dataRows = aoa.slice(headerRow);
        
        dataRows.forEach(row => {
            if (row.every(cell => cell === null)) return;
            const rowObject: Record<string, any> = {};
            headers.forEach((header, index) => {
                rowObject[header] = row[index];
            });
            allData.push(rowObject);
        });
    });

    if (allData.length === 0) {
        throw new Error("No data found in the selected sheets below the specified header row.");
    }
    
    // 2. Pivot Data Structure: Group raw values before aggregation.
    const pivotData = new Map<string, Map<string, number[]>>();
    allData.forEach(row => {
        const rowKey = rowFields.map(field => String(row[field] ?? '')).join(' | ');
        const colKey = columnFields.length > 0 ? columnFields.map(field => String(row[field] ?? '')).join(' | ') : 'Total';
        const value = parseFloat(row[valueField]);

        if (!isNaN(value)) {
            if (!pivotData.has(rowKey)) pivotData.set(rowKey, new Map());
            const colMap = pivotData.get(rowKey)!;
            if (!colMap.has(colKey)) colMap.set(colKey, []);
            colMap.get(colKey)!.push(value);
        }
    });

    // 3. Aggregation: Calculate the final value for each cell in the pivot table.
    const aggregatedData = new Map<string, Map<string, number>>();
    const uniqueColKeys = new Set<string>();
    for (const [rowKey, colMap] of pivotData.entries()) {
        const aggregatedColMap = new Map<string, number>();
        for (const [colKey, values] of colMap.entries()) {
            uniqueColKeys.add(colKey);
            aggregatedColMap.set(colKey, aggregate(values, aggregationType));
        }
        aggregatedData.set(rowKey, aggregatedColMap);
    }

    // 4. Table Construction: Build the final 2D array for the worksheet.
    const sortedRowKeys = Array.from(aggregatedData.keys()).sort();
    const sortedColKeys = Array.from(uniqueColKeys).sort();
    
    const aoa: any[][] = [];
    const headerStyle = { font: { bold: true }, fill: { fgColor: { rgb: "EAEAEA" }, patternType: "solid" as const } };
    const rowHeaderStyle = { font: { bold: true } };

    // Create header row
    const displayHeader = [ { v: rowFields.join('/'), s: headerStyle }, ...sortedColKeys.map(key => ({ v: key, s: headerStyle })) ];
    aoa.push(displayHeader);

    // Create data rows
    sortedRowKeys.forEach(rowKey => {
        const row = [{ v: rowKey, s: rowHeaderStyle }];
        const colMap = aggregatedData.get(rowKey);
        sortedColKeys.forEach(colKey => {
            const value = colMap?.get(colKey) ?? null;
            row.push(value);
        });
        aoa.push(row);
    });
    
    // 5. Workbook Creation
    const newWs = XLSX.utils.aoa_to_sheet(aoa, {cellDates: true});
    const colWidths = [
        { wch: Math.max(25, rowFields.join('/').length + 5) },
        ...sortedColKeys.map(key => ({ wch: Math.max(15, key.length + 5) }))
    ];
    newWs['!cols'] = colWidths;
    if (aoa.length > 1) {
       newWs['!autofilter'] = { ref: newWs['!ref']! };
    }
    
    const newWb = XLSX.utils.book_new();
    const uniqueSheetName = getUniqueSheetName(newWb, outputSheetName);
    XLSX.utils.book_append_sheet(newWb, newWs, uniqueSheetName);
    
    return newWb;
}
