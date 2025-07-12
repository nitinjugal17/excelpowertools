
import * as XLSX from 'xlsx-js-style';

const ROWS_PER_PAGE = 100;

interface TableData {
    title: string;
    headers: { text: string; style: string }[];
    rows: { cells: { text: string; style: string }[] }[];
    footers: { cells: { text: string; style: string }[] }[];
}

interface SheetData {
    sheetName: string;
    tables: TableData[];
}

// --- Data Masking and Scrambling ---
/**
 * Creates a "masked" version of the data where numeric values are replaced with random numbers
 * that are close to the original value (+/- 15%).
 * @param originalData The original sheet or table data.
 * @returns A new array with the masked data.
 */
function createMaskedData(originalData: SheetData[]): SheetData[] {
    const mapTables = (tables: TableData[]) => tables.map(table => ({
        ...table,
        rows: table.rows.map(row => ({
            ...row,
            cells: row.cells.map(cell => {
                const originalValue = parseFloat(cell.text.replace(/,/g, ''));
                let newText = cell.text;
                if (!isNaN(originalValue) && isFinite(originalValue)) {
                    // Generate a random variation between -15% and +15%
                    const variation = (Math.random() - 0.5) * 0.3; // -0.15 to +0.15
                    let newValue = originalValue * (1 + variation);
                    
                    // If original was an integer, keep the new one as an integer
                    if (Number.isInteger(originalValue)) {
                        newValue = Math.round(newValue);
                    } else {
                        // Otherwise, keep a similar number of decimal places
                        const decimalPlaces = (originalValue.toString().split('.')[1] || '').length;
                        newValue = parseFloat(newValue.toFixed(decimalPlaces));
                    }
                    newText = newValue.toLocaleString();
                }
                return { ...cell, text: newText };
            })
        })),
         footers: table.footers.map(row => ({
            ...row,
            cells: row.cells.map(cell => {
                 const originalValue = parseFloat(cell.text.replace(/,/g, ''));
                let newText = cell.text;
                if (!isNaN(originalValue) && isFinite(originalValue)) {
                    const variation = (Math.random() - 0.5) * 0.3;
                    let newValue = originalValue * (1 + variation);
                    if (Number.isInteger(originalValue)) {
                        newValue = Math.round(newValue);
                    } else {
                        const decimalPlaces = (originalValue.toString().split('.')[1] || '').length;
                        newValue = parseFloat(newValue.toFixed(decimalPlaces));
                    }
                    newText = newValue.toLocaleString();
                }
                return { ...cell, text: newText };
            })
        }))
    }));

    return originalData.map(sheetData => ({
        ...sheetData,
        tables: mapTables(sheetData.tables)
    }));
}


// --- HTML Generation ---

function parseSheetToTables(worksheet: XLSX.WorkSheet): TableData[] {
    const allTables: TableData[] = [];
    if (!worksheet || !worksheet['!ref']) {
        return allTables;
    }

    const range = XLSX.utils.decode_range(worksheet['!ref']);
    let currentTableRows: { cells: { text: string; style: string }[] }[] = [];
    let currentTableTitle = '';
    let currentHeaderRow: { cells: { text: string; style: string }[] } | null = null;
    let foundContentSinceLastTable = false;

    const processCurrentTable = () => {
        if (!currentHeaderRow && currentTableRows.length > 0) {
            currentHeaderRow = currentTableRows.shift()!;
        }
        
        if (!currentHeaderRow && currentTableRows.length === 0 && !foundContentSinceLastTable) {
            return;
        }

        if (!currentHeaderRow) {
            if (currentTableRows.length > 0) {
                currentHeaderRow = currentTableRows.shift()!;
            } else {
                currentTableTitle = '';
                currentTableRows = [];
                return;
            }
        }

        const headers = currentHeaderRow.cells;
        
        let bodyRows = [...currentTableRows];
        let footerRows: { cells: { text: string; style: string }[] }[] = [];

        const firstFooterIndex = bodyRows.findIndex(row =>
            row.cells.some(cell => {
                const lowerText = cell.text.toLowerCase().trim();
                return lowerText.startsWith('total') || lowerText.startsWith('grand total');
            })
        );

        if (firstFooterIndex !== -1) {
            footerRows = bodyRows.splice(firstFooterIndex);
        }

        allTables.push({
            title: currentTableTitle,
            headers: headers,
            rows: bodyRows,
            footers: footerRows,
        });

        currentTableRows = [];
        currentTableTitle = '';
        currentHeaderRow = null;
        foundContentSinceLastTable = false;
    };

    for (let R = range.s.r; R <= range.e.r; ++R) {
        const rowCells: { text: string; style: string }[] = [];
        let isRowEmpty = true;
        let rowContentForCheck = '';

        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cellAddress];
            const text = cell?.w || cell?.v || '';
            const textStr = String(text);

            if (textStr.trim() !== '') {
                isRowEmpty = false;
                if (!foundContentSinceLastTable) foundContentSinceLastTable = true;
            }
            rowContentForCheck += textStr;

            let styleString = '';
             if (cell?.s) {
                const style = cell.s;
                if (style.font) {
                    if (style.font.bold) styleString += 'font-weight: bold;';
                    if (style.font.italic) styleString += 'font-style: italic;';
                    if (style.font.color?.rgb && style.font.color.rgb !== "000000") styleString += `color: #${style.font.color.rgb};`;
                }
                if (style.fill?.fgColor?.rgb && style.fill.fgColor.rgb !== "FFFFFF") {
                    styleString += `background-color: #${style.fill.fgColor.rgb};`;
                }
                if (style.alignment?.horizontal) {
                    styleString += `text-align: ${style.alignment.horizontal};`;
                }
            }
            rowCells.push({ text: textStr, style: styleString });
        }

        if (isRowEmpty) {
            if (foundContentSinceLastTable) {
                processCurrentTable();
            }
            continue;
        }

        const firstCellText = rowCells[0].text;
        const lowerFirstCellText = firstCellText.toLowerCase().trim();

        if (lowerFirstCellText.startsWith('summary for:')) {
            if (foundContentSinceLastTable) processCurrentTable();
            currentTableTitle = firstCellText;
            foundContentSinceLastTable = true;
            continue;
        }
        
        if (lowerFirstCellText.startsWith('total') || lowerFirstCellText.startsWith('grand total')) {
            currentTableRows.push({ cells: rowCells });
            processCurrentTable();
            continue;
        }
        
        if (!currentHeaderRow && foundContentSinceLastTable) {
            currentHeaderRow = { cells: rowCells };
        } else {
            currentTableRows.push({ cells: rowCells });
        }
    }
    processCurrentTable();

    return allTables;
}

function generateHtmlForTable(tableData: TableData, sheetId: string, tableIndex: number): string {
    const { title, headers, rows, footers } = tableData;
    const totalRows = rows.length;
    const totalPages = Math.ceil(totalRows / ROWS_PER_PAGE);
    const tableId = `${sheetId}-table-${tableIndex}`;

    let tableHtml = `<div id="${tableId}" class="mb-8 sheet-table" data-scrambled="true">`;
    if (title) {
        tableHtml += `<h2 class="text-xl font-semibold mb-2">${title}</h2>`;
    }
    tableHtml += '<div class="overflow-x-auto shadow ring-1 ring-black ring-opacity-5 rounded-lg">';
    tableHtml += '<table class="min-w-full divide-y divide-gray-300 dark:divide-gray-700">';

    tableHtml += '<thead class="bg-gray-100 dark:bg-gray-800"><tr>';
    headers.forEach(header => {
        const headerStyle = `py-3.5 pl-4 pr-3 text-left text-sm font-semibold text-gray-900 dark:text-gray-100 sm:pl-6; ${header.style}`;
        tableHtml += `<th scope="col" style="${headerStyle}">${header.text}</th>`;
    });
    tableHtml += '</tr></thead>';
    
    tableHtml += `<tbody class="divide-y divide-gray-200 dark:divide-gray-700 bg-white dark:bg-gray-900">`;
    // We render an empty body, which will be populated by JavaScript.
    tableHtml += '</tbody>';

    if (footers.length > 0) {
        tableHtml += '<tfoot class="bg-gray-50 dark:bg-gray-800">';
        tableHtml += '</tfoot>';
    }

    tableHtml += '</table></div>';
    
    if (totalPages > 1) {
        tableHtml += `
            <div class="mt-4 flex items-center justify-between">
                <button onclick="changePage('${tableId}', -1)" class="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-md hover:bg-gray-50 dark:bg-gray-800 dark:text-gray-300 dark:border-gray-600 dark:hover:bg-gray-700">Previous</button>
                <span id="${tableId}-page-info" class="text-sm text-gray-700 dark:text-gray-300">Page 1 of ${totalPages}</span>
                <button onclick="changePage('${tableId}', 1)" class="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-md hover:bg-gray-50 dark:bg-gray-800 dark:text-gray-300 dark:border-gray-600 dark:hover:bg-gray-700">Next</button>
            </div>
        `;
    }
    
    tableHtml += '</div>';
    return tableHtml;
}


/**
 * Orchestrates parsing sheets from a workbook and generating the complete HTML page.
 * @param workbook The full workbook object.
 * @param sheetsToExport An array of sheet names to include in the export.
 * @param fileName The original file name for use in the title.
 * @param fullAccessPassword Optional password for full data access.
 * @param maskedAccessPassword Optional password for masked data access.
 * @returns A promise that resolves to the full HTML string.
 */
export async function generateCombinedHtmlPage(
    workbook: XLSX.WorkBook,
    sheetsToExport: string[],
    fileName: string,
    fullAccessPassword?: string,
    maskedAccessPassword?: string
): Promise<string> {
    const allSheetsData = sheetsToExport.map(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const tables = worksheet ? parseSheetToTables(worksheet) : [];
        return { sheetName, tables };
    });

    const combinedHtmlBody = allSheetsData.map(data => {
        const sanitizedSheetName = data.sheetName.replace(/[^a-zA-Z0-9]/g, '_');
        return `
            <div id="${sanitizedSheetName}" class="sheet-content">
                <h1 class="text-3xl font-bold mb-8 border-b pb-4">${data.sheetName}</h1>
                ${data.tables.map((tableData, tableIndex) => generateHtmlForTable(tableData, sanitizedSheetName, tableIndex)).join('')}
            </div>
        `;
    }).join('');
    
    // Determine the view mode and necessary data/UI elements.
    const useEncryption = Boolean(fullAccessPassword || maskedAccessPassword);
    const defaultToMaskedView = useEncryption && !maskedAccessPassword && fullAccessPassword;

    let scriptData: string;
    let footerHintHtml = '';

    if (useEncryption) {
        const encryptedFullData = fullAccessPassword ? await encryptTables(allSheetsData, fullAccessPassword) : null;
        const maskedData = createMaskedData(allSheetsData);
        const encryptedMaskedData = maskedAccessPassword ? await encryptTables(maskedData, maskedAccessPassword) : null;
        
        let defaultViewData = null;
        if (defaultToMaskedView) {
            defaultViewData = maskedData;
            footerHintHtml = `<span class="mx-2">|</span> <span>Press Ctrl/Cmd + U to Unlock Full Access</span>`;
        }
        
        scriptData = getDecryptionScript(encryptedFullData, encryptedMaskedData, defaultViewData);
    } else {
        scriptData = getUnencryptedScript(allSheetsData);
    }
    
    return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${fileName} - Export</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body { -webkit-user-select: none; -ms-user-select: none; user-select: none; }
        @media print {
            body, html { display: none !important; }
        }
    </style>
</head>
<body class="bg-gray-50 dark:bg-gray-950 text-gray-800 dark:text-gray-200" oncontextmenu="return false;">
    <div class="container mx-auto p-4 sm:p-6 lg:p-8">
        ${combinedHtmlBody}
        <footer class="text-center text-xs text-gray-400 mt-8 pt-4 border-t">
            <span>Generated by Excel Power Tools</span>
            <span id="unlock-hint-container">${footerHintHtml}</span>
        </footer>
    </div>
    <script>${scriptData}</script>
</body>
</html>`;
}


// --- Encryption and Script Generation ---

export async function encryptTables(dataToEncrypt: SheetData[], password: string): Promise<string> {
    const salt = window.crypto.getRandomValues(new Uint8Array(16));
    const iv = window.crypto.getRandomValues(new Uint8Array(12));
    const key = await deriveKey(password, salt);
    
    const plaintext = new TextEncoder().encode(JSON.stringify(dataToEncrypt));
    
    const encryptedContent = await window.crypto.subtle.encrypt(
        { name: 'AES-GCM', iv: iv },
        key,
        plaintext
    );

    const data = {
        salt: bufferToBase64(salt),
        iv: bufferToBase64(iv),
        ciphertext: bufferToBase64(encryptedContent)
    };

    return JSON.stringify(data);
}

function deriveKey(password: string, salt: Uint8Array) {
    const enc = new TextEncoder();
    return window.crypto.subtle.importKey(
        'raw',
        enc.encode(password),
        { name: 'PBKDF2' },
        false,
        ['deriveKey']
    ).then(baseKey =>
        window.crypto.subtle.deriveKey(
            { name: 'PBKDF2', salt: salt, iterations: 100000, hash: 'SHA-256' },
            baseKey,
            { name: 'AES-GCM', length: 256 },
            true,
            ['encrypt', 'decrypt']
        )
    );
}

function bufferToBase64(buffer: ArrayBuffer): string {
    return btoa(String.fromCharCode(...new Uint8Array(buffer)));
}


export function getUnencryptedScript(unencryptedData?: SheetData[]): string {
    const unencryptedDataJson = unencryptedData ? JSON.stringify(unencryptedData) : '[]';
    return `
        const pageStates = {};
        function changePage(tableId, direction) {
            if (!pageStates[tableId]) {
                const tableElement = document.getElementById(tableId);
                const allRows = Array.from(tableElement.querySelectorAll('tbody tr'));
                const totalRows = allRows.length;
                const totalPages = Math.ceil(totalRows / ${ROWS_PER_PAGE});
                pageStates[tableId] = { currentPage: 1, totalPages: totalPages, allRows: allRows };
            }
            const state = pageStates[tableId];
            let newPage = state.currentPage + direction;
            if (newPage < 1) newPage = 1;
            if (newPage > state.totalPages) newPage = state.totalPages;
            state.currentPage = newPage;
            
            const start = (newPage - 1) * ${ROWS_PER_PAGE};
            const end = start + ${ROWS_PER_PAGE};

            state.allRows.forEach((row, index) => {
                row.style.display = (index >= start && index < end) ? '' : 'none';
            });
            
            const pageInfo = document.getElementById(tableId + '-page-info');
            if (pageInfo) {
                pageInfo.textContent = 'Page ' + newPage + ' of ' + state.totalPages;
            }
        }
        function renderData(data) {
             const sheets = Array.isArray(data) ? (data.length > 0 && 'sheetName' in data[0] ? data : []) : [];
             sheets.forEach(sheetData => {
                 const sanitizedSheetName = sheetData.sheetName.replace(/[^a-zA-Z0-9]/g, '_');
                 sheetData.tables.forEach((tableData, tableIndex) => {
                    const tableId = \`\${sanitizedSheetName}-table-\${tableIndex}\`;
                    const tableContainer = document.getElementById(tableId);
                    if (!tableContainer) return;
                    tableContainer.removeAttribute('data-scrambled');
                    
                    const tbody = tableContainer.querySelector('tbody');
                    const tfoot = tableContainer.querySelector('tfoot');
                    
                    if (tbody) {
                        tbody.innerHTML = ''; // Clear existing
                        let html = '';
                        tableData.rows.forEach(row => {
                            html += '<tr>';
                            row.cells.forEach((cell, cellIndex) => {
                                const finalStyle = cellIndex === 0 ? \`font-medium text-gray-900 dark:text-gray-100 sm:pl-6; \${cell.style}\` : cell.style;
                                html += \`<td class="px-3 py-4 text-sm text-gray-500 dark:text-gray-400" style="\${finalStyle}">\${cell.text}</td>\`;
                            });
                            html += '</tr>';
                        });
                        tbody.innerHTML = html;
                        if (tableData.rows.length > ${ROWS_PER_PAGE}) {
                             changePage(tableId, 0); // Initialize pagination display
                        }
                    }

                    if (tfoot) {
                        tfoot.innerHTML = '';
                        let footerHtml = '';
                        tableData.footers.forEach(row => {
                            footerHtml += '<tr>';
                            row.cells.forEach((cell) => {
                                footerHtml += \`<td class="px-3 py-3.5 text-left text-sm font-semibold text-gray-900 dark:text-gray-100" style="\${cell.style}">\${cell.text}</td>\`;
                            });
                            footerHtml += '</tr>';
                        });
                        tfoot.innerHTML = footerHtml;
                    }
                 });
             });
        }
        document.addEventListener('keydown', function (e) {
            if ((e.ctrlKey || e.metaKey) && ['p', 's', 'c', 'x'].includes(e.key)) {
                e.preventDefault();
                alert('This action is disabled for security.');
            }
        });
        document.addEventListener('DOMContentLoaded', function() {
            renderData(${unencryptedDataJson});
        });
    `;
}

export function getDecryptionScript(
    encryptedFullDataJson: string | null, 
    encryptedMaskedDataJson: string | null,
    defaultMaskedData: SheetData[] | null
): string {
    const defaultMaskedDataJson = defaultMaskedData ? JSON.stringify(defaultMaskedData) : 'null';
    
    return `
        const encryptedFullData = ${encryptedFullDataJson || 'null'};
        const encryptedMaskedData = ${encryptedMaskedDataJson || 'null'};
        let decryptedData = null;

        function base64ToBuffer(base64) {
            const binaryString = atob(base64);
            const len = binaryString.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            return bytes.buffer;
        }

        async function deriveKey(password, salt) {
            const enc = new TextEncoder();
            const baseKey = await crypto.subtle.importKey('raw', enc.encode(password), { name: 'PBKDF2' }, false, ['deriveKey']);
            return await crypto.subtle.deriveKey(
                { name: 'PBKDF2', salt: salt, iterations: 100000, hash: 'SHA-256' },
                baseKey,
                { name: 'AES-GCM', length: 256 },
                true,
                ['encrypt', 'decrypt']
            );
        }

        async function decryptAndDisplay(password) {
            const prompt = document.getElementById('password-prompt');
            const errorElement = document.getElementById('password-error');
            if(errorElement) errorElement.textContent = 'Decrypting...';

            if (encryptedFullData) {
                try {
                    const decrypted = await decryptData(encryptedFullData, password);
                    decryptedData = decrypted;
                    if(prompt) document.body.removeChild(prompt);
                    const unlockHintContainer = document.getElementById('unlock-hint-container');
                    if(unlockHintContainer) unlockHintContainer.style.display = 'none';
                    renderData(decryptedData);
                    return; // Success
                } catch (e) {
                    console.log('Full access decryption failed, trying masked access...');
                }
            }

            if (encryptedMaskedData) {
                 try {
                    const decrypted = await decryptData(encryptedMaskedData, password);
                    decryptedData = decrypted;
                    if(prompt) document.body.removeChild(prompt);
                    renderData(decryptedData);
                    return; // Success
                } catch (e) {
                    console.log('Masked access decryption failed.');
                }
            }

            if (errorElement) errorElement.textContent = 'Incorrect password. Please try again.';
        }
        
        async function decryptData(encryptedDataObj, password) {
            const salt = base64ToBuffer(encryptedDataObj.salt);
            const iv = base64ToBuffer(encryptedDataObj.iv);
            const ciphertext = base64ToBuffer(encryptedDataObj.ciphertext);
            const key = await deriveKey(password, salt);
            const decryptedContent = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: iv }, key, ciphertext);
            return JSON.parse(new TextDecoder().decode(decryptedContent));
        }
        
        const pageStates = {};
        function changePage(tableId, direction) {
            if (!pageStates[tableId]) {
                 const tableElement = document.getElementById(tableId);
                const allRows = Array.from(tableElement.querySelectorAll('tbody tr'));
                const totalRows = allRows.length;
                const totalPages = Math.ceil(totalRows / ${ROWS_PER_PAGE});
                pageStates[tableId] = { currentPage: 1, totalPages: totalPages, allRows: allRows };
            }
            const state = pageStates[tableId];
            let newPage = state.currentPage + direction;
            if (newPage < 1) newPage = 1;
            if (newPage > state.totalPages) newPage = state.totalPages;
            state.currentPage = newPage;
            
            const start = (newPage - 1) * ${ROWS_PER_PAGE};
            const end = start + ${ROWS_PER_PAGE};
            
            state.allRows.forEach((row, index) => {
                row.style.display = (index >= start && index < end) ? '' : 'none';
            });
            
            const pageInfo = document.getElementById(tableId + '-page-info');
            if (pageInfo) {
                pageInfo.textContent = 'Page ' + newPage + ' of ' + state.totalPages;
            }
        }
        
        function renderData(data) {
             const sheets = Array.isArray(data) ? (data.length > 0 && 'sheetName' in data[0] ? data : []) : [];
             sheets.forEach(sheetData => {
                 const sanitizedSheetName = sheetData.sheetName.replace(/[^a-zA-Z0-9]/g, '_');
                 sheetData.tables.forEach((tableData, tableIndex) => {
                    const tableId = \`\${sanitizedSheetName}-table-\${tableIndex}\`;
                    const tableContainer = document.getElementById(tableId);
                    if (!tableContainer) return;
                    tableContainer.removeAttribute('data-scrambled');
                    
                    const tbody = tableContainer.querySelector('tbody');
                    const tfoot = tableContainer.querySelector('tfoot');
                    
                    if (tbody) {
                        tbody.innerHTML = ''; // Clear existing
                        let html = '';
                        tableData.rows.forEach(row => {
                            html += '<tr>';
                            row.cells.forEach((cell, cellIndex) => {
                                const finalStyle = cellIndex === 0 ? \`font-medium text-gray-900 dark:text-gray-100 sm:pl-6; \${cell.style}\` : cell.style;
                                html += \`<td class="px-3 py-4 text-sm text-gray-500 dark:text-gray-400" style="\${finalStyle}">\${cell.text}</td>\`;
                            });
                            html += '</tr>';
                        });
                        tbody.innerHTML = html;
                         if (tableData.rows.length > ${ROWS_PER_PAGE}) {
                             changePage(tableId, 0); // Initialize pagination display
                        }
                    }

                    if (tfoot) {
                        tfoot.innerHTML = '';
                        let footerHtml = '';
                        tableData.footers.forEach(row => {
                            footerHtml += '<tr>';
                            row.cells.forEach((cell) => {
                                footerHtml += \`<td class="px-3 py-3.5 text-left text-sm font-semibold text-gray-900 dark:text-gray-100" style="\${cell.style}">\${cell.text}</td>\`;
                            });
                            footerHtml += '</tr>';
                        });
                        tfoot.innerHTML = footerHtml;
                    }
                 });
             });
        }
        
        function showPasswordPrompt() {
            if (document.getElementById('password-prompt')) return;
            const promptHtml = \`
                <div id="password-prompt" class="fixed inset-0 bg-gray-900 bg-opacity-75 flex items-center justify-center z-50">
                    <div class="bg-white dark:bg-gray-800 p-8 rounded-lg shadow-xl max-w-sm w-full">
                        <h2 class="text-2xl font-bold mb-4 text-gray-900 dark:text-gray-100">Password Required</h2>
                        <p class="text-gray-600 dark:text-gray-300 mb-6">This content is encrypted. Please enter the password to view.</p>
                        <form id="password-form">
                            <input type="password" id="password-input" class="w-full px-4 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 dark:bg-gray-700 dark:border-gray-600 dark:text-white">
                            <p id="password-error" class="text-red-500 text-sm mt-2 h-4"></p>
                            <button type="submit" class="mt-4 w-full bg-blue-600 text-white font-bold py-2 px-4 rounded-md hover:bg-blue-700 transition duration-300">Unlock</button>
                        </form>
                    </div>
                </div>\`;
            document.body.insertAdjacentHTML('beforeend', promptHtml);
            document.getElementById('password-form').addEventListener('submit', (e) => {
                e.preventDefault();
                const input = document.getElementById('password-input');
                if (input) {
                    decryptAndDisplay(input.value);
                }
            });
            const inputElement = document.getElementById('password-input');
            if (inputElement) {
                inputElement.focus();
            }
        }
        
        document.addEventListener('keydown', function (e) {
            if ((e.ctrlKey || e.metaKey) && ['p', 's', 'c', 'x'].includes(e.key)) {
                e.preventDefault();
                alert('This action is disabled for security.');
            }
            if ((e.ctrlKey || e.metaKey) && e.key === 'u') {
                e.preventDefault();
                showPasswordPrompt();
            }
        });

        window.logDecryptedData = () => {
            if (decryptedData) {
                console.log("Decrypted Data:", decryptedData);
            } else {
                console.log("Data has not been decrypted yet. Please enter the correct password.");
            }
        };

        document.addEventListener('DOMContentLoaded', function() {
            const defaultData = ${defaultMaskedDataJson};
            if (defaultData) {
                renderData(defaultData);
            } else {
                showPasswordPrompt();
            }
        });
    `;
}
