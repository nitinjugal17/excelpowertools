
import * as XLSX from 'xlsx-js-style';

const MAX_SHEET_NAME_LENGTH = 31;
const INVALID_SHEET_NAME_CHARS_REGEX = /[\\\/\?\*\[\]:]/g;


/**
 * Sanitizes a string to be a valid Excel sheet name.
 * - Removes invalid characters.
 * - Truncates to a maximum length (31 characters).
 * - Ensures the name is not empty.
 * @param name The original string to sanitize.
 * @returns A valid Excel sheet name.
 */
export function sanitizeSheetName(name: string): string {
  if (typeof name !== 'string' || name.trim() === '') {
    return 'Sheet';
  }

  let sanitized = name.replace(INVALID_SHEET_NAME_CHARS_REGEX, '');

  if (sanitized.length > MAX_SHEET_NAME_LENGTH) {
    sanitized = sanitized.substring(0, MAX_SHEET_NAME_LENGTH);
  }

  if (sanitized.trim() === '') {
    return 'Sheet';
  }

  return sanitized;
}

/**
 * Sanitizes a sheet name for use in an Excel formula by quoting it if necessary.
 * @param name The original sheet name.
 * @returns A formula-safe sheet name.
 */
export function sanitizeSheetNameForFormula(name: string): string {
  // If the sheet name contains spaces, hyphens, or other special characters that might break a formula link, wrap it in single quotes.
  // Also, any single quotes within the name need to be doubled up.
  if (/[ \-!@#$%^&*()+={}|[\]\\:";<>,.?/]/g.test(name) || /^\d+$/.test(name) || !isNaN(Number(name))) {
    return `'${name.replace(/'/g, "''")}'`;
  }
  return name;
}

/**
 * Parses a single column identifier (letter or 1-indexed number) into a 0-indexed column number.
 * @param identifier The string to parse (e.g., "A" or "1").
 * @returns A 0-indexed column number or null if invalid.
 */
export function parseColumnIdentifier(identifier: string): number | null {
  if (!identifier || typeof identifier !== 'string') return null;
  const part = identifier.trim().toUpperCase();
  if (!part) return null;
  
  if (/^[A-Z]+$/.test(part)) {
    try {
      return XLSX.utils.decode_col(part);
    } catch {
      return null;
    }
  } else if (/^\d+$/.test(part)) {
    const colIndex = parseInt(part, 10) - 1;
    return colIndex >= 0 ? colIndex : null;
  }
  return null;
}


/**
 * Parses a comma-separated string of column identifiers (letters or 1-indexed numbers)
 * into an array of 0-indexed column numbers. Supports ranges like "A:C".
 * @param columnsString The string to parse (e.g., "A,C,E:G").
 * @param headers An array of header strings from the sheet.
 * @returns An array of unique, sorted, 0-indexed column numbers.
 */
export function parseSourceColumns(columnsString: string, headers?: string[]): number[] {
    const indices = new Set<number>();
    if (!columnsString || typeof columnsString !== 'string') {
        return [];
    }
    const parts = columnsString.split(',').map(p => p.trim());
    for (const part of parts) {
        if (part.includes(':')) {
            const [start, end] = part.split(':').map(p => p.trim());
            const startIdx = getColumnIndex(start, headers);
            const endIdx = getColumnIndex(end, headers);
            if (startIdx !== null && endIdx !== null && startIdx <= endIdx) {
                for (let i = startIdx; i <= endIdx; i++) {
                    indices.add(i);
                }
            }
        } else {
            const colIndex = getColumnIndex(part, headers);
            if (colIndex !== null) {
                indices.add(colIndex);
            }
        }
    }
    return Array.from(indices).sort((a, b) => a - b);
}


/**
 * Internal helper to resolve a column identifier by name or by letter/number
 * @param identifier The column name or letter/number identifier.
 * @param headers An array of header strings from the sheet.
 * @returns The 0-indexed column number or null if not found.
 */
export function getColumnIndex(identifier: string, headers?: string[]): number | null {
    if (!identifier) return null;
    const trimmedIdentifier = identifier.trim();

    // 1. Check by column name (case-insensitive)
    if (headers && headers.length > 0) {
        const lowerCaseHeaders = headers.map(h => String(h || '').toLowerCase());
        const byNameIndex = lowerCaseHeaders.indexOf(trimmedIdentifier.toLowerCase());
        if (byNameIndex !== -1) return byNameIndex;
    }

    // 2. Check by letter or 1-indexed number
    return parseColumnIdentifier(trimmedIdentifier);
}

/**
 * Generates a unique sheet name within a workbook.
 * @param workbook The workbook to check against.
 * @param desiredName The preferred name for the sheet.
 * @returns A unique sheet name.
 */
export function getUniqueSheetName(workbook: XLSX.WorkBook, desiredName: string): string {
    const sanitized = sanitizeSheetName(desiredName);
    let finalName = sanitized;
    const existingSheetNames = new Set((workbook.SheetNames || []).map(name => name.toLowerCase()));

    if (existingSheetNames.has(finalName.toLowerCase())) {
        let counter = 1;
        let newNameAttempt;

        do {
            const suffix = `_${counter}`;
            const baseName = sanitized.substring(0, MAX_SHEET_NAME_LENGTH - suffix.length);
            newNameAttempt = `${baseName}${suffix}`;
            counter++;
        } while (existingSheetNames.has(newNameAttempt.toLowerCase()));

        finalName = newNameAttempt;
    }
    return finalName;
}

/**
 * Escapes a string for use in a regular expression.
 * @param str The string to escape.
 * @returns The escaped string.
 */
export function escapeRegex(str: string): string {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
