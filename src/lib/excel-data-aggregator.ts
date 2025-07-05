import * as XLSX from 'xlsx-js-style';
import type { AggregationResult } from './excel-types';

export * from './excel-aggregator-core';
export * from './excel-aggregator-actions';
export * from './excel-aggregator-reports';

/**
 * Recalculates aggregation results based on a map of edited keys.
 * This function rebuilds the data model from scratch to ensure consistency.
 * @param resultToProcess The original AggregationResult from the initial processing.
 * @param editableKeys A map where keys are original keys and values are the new (potentially edited) keys.
 * @param blankLabel The user-defined label for blank keys.
 * @returns A new AggregationResult object reflecting the edits, or null if input is invalid.
 */
export function getModifiedAggregationData(
    resultToProcess: AggregationResult | null,
    editableKeys: Map<string, string>,
    blankLabel: string
): { modifiedResult: AggregationResult, modifiedValueToKeyMap: Map<string, string>, finalBlankLabel: string, editedKeyToOriginalsMap: Map<string, string[]> } | null {
    if (!resultToProcess) return null;
    console.log("Rebuilding aggregation data based on user edits...");

    const keyResolutionMap = new Map<string, string>();

    const allKeysInCurrentResult = new Set(resultToProcess.reportingKeys);
    const originalBlankLabel = blankLabel.trim() || '(Blanks)';
    if (resultToProcess.blankCounts && resultToProcess.blankCounts.total > 0) {
        allKeysInCurrentResult.add(originalBlankLabel);
    }

    allKeysInCurrentResult.forEach(key => {
        const userEdit = editableKeys.get(key);
        if (userEdit && userEdit.trim()) {
            keyResolutionMap.set(key, userEdit.trim());
        } else {
            keyResolutionMap.set(key, key);
        }
    });

    const editedKeyToOriginalsMap = new Map<string, string[]>();
    keyResolutionMap.forEach((newKey, originalKey) => {
        if (!editedKeyToOriginalsMap.has(newKey)) {
            editedKeyToOriginalsMap.set(newKey, []);
        }
        editedKeyToOriginalsMap.get(newKey)!.push(originalKey);
    });

    const newPerSheetCounts: { [sheetName: string]: { [key: string]: number } } = {};
    resultToProcess.processedSheetNames.forEach(sheetName => {
        newPerSheetCounts[sheetName] = {};
    });

    const addCount = (sheetName: string, key: string, count: number) => {
        if (!newPerSheetCounts[sheetName]) newPerSheetCounts[sheetName] = {};
        newPerSheetCounts[sheetName][key] = (newPerSheetCounts[sheetName][key] || 0) + count;
    };

    for (const sheetName in resultToProcess.perSheetCounts) {
        const originalSheetCounts = resultToProcess.perSheetCounts[sheetName] || {};
        for (const originalKey in originalSheetCounts) {
            const count = originalSheetCounts[originalKey];
            const finalKey = keyResolutionMap.get(originalKey) || originalKey;
            addCount(sheetName, finalKey, count);
        }
    }
    
    if (resultToProcess.blankCounts && resultToProcess.blankCounts.total > 0) {
        const finalBlankKey = keyResolutionMap.get(originalBlankLabel) || originalBlankLabel;
        for (const sheetName in resultToProcess.blankCounts.perSheet) {
            const count = resultToProcess.blankCounts.perSheet[sheetName] || 0;
            if (count > 0) {
                 addCount(sheetName, finalBlankKey, count);
            }
        }
    }
    
    const newTotalCounts: { [key: string]: number } = {};
    for (const sheetName in newPerSheetCounts) {
        for (const key in newPerSheetCounts[sheetName]) {
            newTotalCounts[key] = (newTotalCounts[key] || 0) + newPerSheetCounts[sheetName][key];
        }
    }
    
    const modifiedResult: AggregationResult = {
        ...resultToProcess,
        totalCounts: newTotalCounts,
        perSheetCounts: newPerSheetCounts,
        reportingKeys: Object.keys(newTotalCounts).sort((a, b) => a.localeCompare(b)),
        blankCounts: resultToProcess.blankCounts ? {
            total: Object.values(newPerSheetCounts).reduce((total, sheetCounts) => {
                const finalBlankKey = keyResolutionMap.get(originalBlankLabel) || originalBlankLabel;
                return total + (sheetCounts[finalBlankKey] || 0);
            }, 0),
            perSheet: Object.keys(newPerSheetCounts).reduce((acc, sheetName) => {
                const finalBlankKey = keyResolutionMap.get(originalBlankLabel) || originalBlankLabel;
                acc[sheetName] = newPerSheetCounts[sheetName][finalBlankKey] || 0;
                return acc;
            }, {} as Record<string, number>),
        } : undefined,
        blankDetails: resultToProcess.blankDetails,
    };
    
    const modifiedValueToKeyMap = new Map<string, string>();
    if (resultToProcess.valueToKeyMap) {
        for (const [value, originalKey] of resultToProcess.valueToKeyMap.entries()) {
            const finalKey = keyResolutionMap.get(originalKey) || originalKey;
            modifiedValueToKeyMap.set(value, finalKey);
        }
    }
    
    const finalBlankLabel = keyResolutionMap.get(originalBlankLabel) || originalBlankLabel;
    
    console.log("Finished rebuilding aggregation data. Final keys:", modifiedResult.reportingKeys);
    return { modifiedResult, modifiedValueToKeyMap, finalBlankLabel, editedKeyToOriginalsMap };
}
