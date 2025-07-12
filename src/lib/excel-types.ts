





















export type HorizontalAlignment = 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous';
export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
export type Alignment = HorizontalAlignment;
export type AggregationType = 'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX';
export type MatchMode = 'whole' | 'partial' | 'loose';
export type BorderStyle = 'thin' | 'medium' | 'thick' | 'double' | 'dotted' | 'dashed';


export interface FormattingConfig {
  dataTitlesRowNumber: number; 
  styleOptions: { 
    bold?: boolean; 
    italic?: boolean; 
    underline?: boolean;
    alignment?: HorizontalAlignment;
    fontName?: string;
    fontSize?: number;
  };
}

export interface CustomHeaderConfig {
  text: string;
  insertBeforeRow: number; 
  mergeAndCenter?: boolean;
  styleOptions: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontName?: string;
    fontSize?: number;
    wrapText?: boolean;
    indent?: number;
    horizontalAlignment?: HorizontalAlignment;
    verticalAlignment?: VerticalAlignment;
  }
}

export interface CustomColumnConfig {
  newColumnName: string;
  newColumnHeaderRow: number;
  insertColumnBefore: string; 
  sourceDataColumn: string; 
  textSplitter: string;
  partToUse: number; 
  dataStartRow: number; 
  alignment?: HorizontalAlignment;
}

export interface RangeFormattingConfig {
  startRow: number;
  endRow: number;
  startCol: string;
  endCol: string;
  merge: boolean;
  style: {
    font: {
      bold: boolean;
      italic: boolean;
      underline: boolean;
      name: string;
      size: number;
      color: string; // hex
    };
    alignment: {
      horizontal: HorizontalAlignment;
      vertical: VerticalAlignment;
    };
    fill: {
      color: string; // hex
    };
  }
}

export interface SheetProtectionConfig {
  password: string;
  type: 'full' | 'range';
  range?: string; // e.g., "A1:D10"
  selectLockedCells?: boolean;
}

export interface CommandDisablingConfig {
  disableCopyPaste: boolean;
  disablePrint: boolean;
}


export interface SplitterCustomHeaderConfig {
  text?: string;
  insertBeforeRow: number;
  mergeAndCenter: boolean;
  valueSeparator: string;
  sourceColumnString: string;
  styleOptions: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    alignment?: HorizontalAlignment;
    fontName?: string;
    fontSize?: number;
  };
}

export interface SplitterCustomColumnConfig {
  name: string;
  value: string;
}

export interface IndexSheetConfig {
  sheetName: string;
  headerText: string;
  headerRow: number;
  headerCol: string;
  linksStartRow: number;
  linksCol: string;
  backLinkText: string;
  backLinkRow: number;
  backLinkCol: string;
}


export interface BlankRowDetail {
    sheetName: string;
    rowNumber: number; // 1-indexed
    rowData: Record<string, any>; // Store as object with header keys
}

export interface AggregationResult {
  totalCounts: { [key: string]: number };
  perSheetCounts: { [sheetName: string]: { [key: string]: number } };
  blankCounts?: {
    total: number;
    perSheet: { [sheetName: string]: number };
  };
  blankDetails?: BlankRowDetail[];
  reportingKeys: string[];
  valueToKeyMap: Map<string, string>;
  processedSheetNames: string[];
  searchColumnIdentifiers: string;
  matchingRows?: { [sheetName: string]: Set<number> };
  sheetKeyColumnIndices?: { [sheetName: string]: number };
  aggregationMode?: 'valueMatch' | 'keyMatch';
  blankCountingMode?: 'rowAware' | 'fullColumn';
  matchMode?: MatchMode;
  sheetTitles?: Record<string, string>;
}

export interface UpdateDetail {
  sheetName: string;
  rowNumber: number;
  cellAddress: string;
  originalValue: any;
  newValue: any;
  keyUsed: string;
  triggerValue: string;
  triggerColumn: string;
  rowData: Record<string, any>;
}

export interface UpdateResult {
    summary: {
        totalCellsUpdated: number;
        sheetsUpdated: string[];
    };
    details: UpdateDetail[];
}

export interface HeaderFormatOptions {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontName?: string;
    fontSize?: number;
    horizontalAlignment?: HorizontalAlignment;
    fillColor?: string; // RGB hex string without #
    fontColor?: string; // RGB hex string without #
}

export interface TableFormattingOptions {
    fillColor?: string;
    borderColor?: string;
    borderStyle?: BorderStyle;
}

export interface SummaryConfig {
    summarySheetName?: string;
    aggregationMode?: 'valueMatch' | 'keyMatch';
    keyCountColumn?: string;
    insertColumn?: string;
    insertStartRow?: number;
    headerFormatting?: HeaderFormatOptions;
    blankRowFormatting?: HeaderFormatOptions;
    totalRowFormatting?: HeaderFormatOptions;
    countBlanksInColumn?: string;
    blankCountLabel?: string;
    chunkSize?: number;
    conditionalValueMatch?: {
      checkColumn: string;
    };
    inSheetSummaryTitle?: string;
    reportLayout?: 'sheetsAsRows' | 'keysAsRows';
    autoSizeColumns?: boolean;
    columnsToHide?: string;
    blankCountingMode?: 'rowAware' | 'fullColumn';
    showBlanksInInSheetSummary?: boolean;
    clearExistingInSheetSummary?: boolean;
    showOnlyLocalKeysInSummary?: boolean;
    summaryTitleCell?: string;
    tableFormatting?: TableFormattingOptions;
    groupReportSheetTitle?: string;
    groupReportMultiSourceTitle?: string;
    groupReportHeaderFormatting?: HeaderFormatOptions;
    groupReportDescription?: string;
}

export interface EmptyCellResult {
  sheetName: string;
  address: string;
  row: number; // 1-indexed row number
  rowData?: Record<string, any>; // Optional: full row data for detailed reports
  contextValue?: string | number | null; // For compact report context
  keyColumnValue?: string | number | null; // For summary report
  contextColumnForSummaryValue?: string | number | null; // For summary report
}

export interface EmptyCellReport {
  summary: { [sheetName: string]: number };
  locations: EmptyCellResult[];
  totalEmpty: number;
  processedSheetNames?: string[];
}

export interface DuplicateUpdateDetail {
  sheetName: string;
  row: number; // 1-indexed Excel row of the duplicate
  firstInstanceRow: number; // 1-indexed row number of the first instance
  updatedAddress: string;
  updatedValue: any;
  key: string;
  rowData: Record<string, any>;
}


export interface DuplicateReport {
  summary: { [sheetName: string]: number }; // sheetName -> count of duplicates found
  updates: DuplicateUpdateDetail[];
  totalDuplicates: number;
}

export interface TextFormatConfig {
  searchText: string[];
  searchMode: 'text' | 'regex';
  matchCase: boolean;
  matchEntireCell: boolean;
  range?: {
    startRow: number;
    endRow: number;
    startCol: string;
    endCol: string;
  };
  style: {
    font?: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      name?: string;
      size?: number;
      color?: string; // hex
    };
    alignment?: {
      horizontal?: HorizontalAlignment;
      vertical?: VerticalAlignment;
    };
    fill?: {
      color?: string; // hex
    };
  }
}

export interface AiImputationContext {
  sheetName: string;
  address: string;
  row: number;
  rowData: Record<string, any>;
  headers: string[];
  targetColumn: string;
  exampleRows: Record<string, any>[];
}

export interface AiImputationSuggestion {
  sheetName: string;
  address: string;
  row: number;
  suggestion: string;
  isChecked: boolean;
  rowData: Record<string, any>;
  context?: {
    label: string;
    value: any;
  }[];
}

export interface ExtractionConfig {
  lookupColumn: string;
  lookupValue: string;
  returnColumns: string;
  headerRow: number;
}

export interface ExtractionReport {
  summary: {
    sheetsSearched: string[];
    perSheetSummary: Record<string, number>;
    totalRowsFound: number;
  };
  details: (Record<string, any> & { "Source Sheet": string })[];
  config: ExtractionConfig;
}

export interface SheetComparisonResult {
  summary: {
    newRows: number;
    deletedRows: number;
    modifiedRows: number;
  };
  new: any[][];
  deleted: any[][];
  modified: {
    key: string;
    rowA: any[];
    rowB: any[];
    diffs: { colName: string; valueA: any; valueB: any }[];
  }[];
}

export interface ComparisonReport {
  summary: {
    totalSheetsCompared: number;
    sheetsWithDifferences: string[];
    totalRowsFound: number;
  };
  details: {
    [sheetName: string]: SheetComparisonResult;
  };
  config: {
    headerRow: number;
    primaryKeyColumns: string;
  };
}
