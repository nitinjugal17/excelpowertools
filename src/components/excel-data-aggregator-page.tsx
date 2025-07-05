
"use client";

import React, { useState, useCallback, ChangeEvent, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { Textarea } from '@/components/ui/textarea';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, Sigma, ListChecks, CheckCircle2, Loader2, ScrollText, BarChart2, Repeat, FileOutput, Settings, BarChartHorizontal, ChevronDown, Lightbulb, Palette, FileSpreadsheet, FileUp, XCircle, PencilLine, FileText, Search, PlayCircle, NotebookTabs, FileCog, Filter, Eraser, WholeWord, Table as TableIcon } from 'lucide-react';
import { aggregateData, addUpdateReportSheetToWorkbook, fillEmptyKeyColumn, insertAggregationResultsIntoSheets, lookupAndAndUpdate, findPotentialUpdates, addAggregationReportSheetsToWorkbook, markMatchingRows, getModifiedAggregationData, stripFormulasInWorkbook } from '@/lib/excel-data-aggregator';
import { createGroupReportWorkbook } from '@/lib/excel-group-reporter';
import type { AggregationResult, UpdateResult, SummaryConfig, HorizontalAlignment, HeaderFormatOptions, MatchMode, TableFormattingOptions, BorderStyle } from '@/lib/excel-types';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { DropdownMenu, DropdownMenuTrigger, DropdownMenuContent, DropdownMenuItem } from '@/components/ui/dropdown-menu';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { useLanguage } from '@/context/language-context';
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from "@/components/ui/tooltip";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow, TableFooter as UiTableFooter } from '@/components/ui/table';
import { ScrollArea } from '@/components/ui/scroll-area';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';
import { parseSourceColumns, getColumnIndex, escapeRegex } from '@/lib/excel-helpers';
import { Markup } from '@/components/ui/markup';


interface SheetSelection {
  [sheetName: string]: boolean;
}

interface ExcelDataAggregatorPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}


export default function ExcelDataAggregatorPage({ onProcessingChange, onFileStateChange }: ExcelDataAggregatorPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<SheetSelection>({});
  const [headerRow, setHeaderRow] = useState<number>(1);
  const [searchColumns, setSearchColumns] = useState<string>('A');
  const [searchTerms, setSearchTerms] = useState<string>('');
  const [mappingFile, setMappingFile] = useState<File | null>(null);
  const [mappingFileContent, setMappingFileContent] = useState<string>('');
  const [matchMode, setMatchMode] = useState<MatchMode>('whole');
  
  const [stripFormulas, setStripFormulas] = useState<boolean>(false);
  const [enableUpdate, setEnableUpdate] = useState<boolean>(false);
  const [updateOnlyBlanks, setUpdateOnlyBlanks] = useState<boolean>(true);
  const [updateColumn, setUpdateColumn] = useState<string>('B');
  const [generateUpdateReport, setGenerateUpdateReport] = useState<boolean>(false);
  const [enableFillKeyColumn, setEnableFillKeyColumn] = useState<boolean>(false);
  const [enablePairedRowValidation, setEnablePairedRowValidation] = useState<boolean>(false);
  const [pairedValidationColumns, setPairedValidationColumns] = useState<string>('');
  
  const [enableMarking, setEnableMarking] = useState<boolean>(false);
  const [markColumn, setMarkColumn] = useState<string>('');
  const [markValue, setMarkValue] = useState<string>('Processed');


  const [aggregationResult, setAggregationResult] = useState<AggregationResult | null>(null);
  const [finalAggregationResult, setFinalAggregationResult] = useState<AggregationResult | null>(null);
  const [editableKeys, setEditableKeys] = useState<Map<string, string>>(new Map());


  const [insertResultsInSheet, setInsertResultsInSheet] = useState<boolean>(false);
  const [clearExistingInSheetSummary, setClearExistingInSheetSummary] = useState<boolean>(true);
  const [inSheetDataSource, setInSheetDataSource] = useState<'reportingScope' | 'localSheet'>('reportingScope');
  const [showOnlyLocalKeys, setShowOnlyLocalKeys] = useState<boolean>(true);
  const [inSheetGenerationMode, setInSheetGenerationMode] = useState<'static' | 'formula'>('static');
  const [inSheetSummaryTitle, setInSheetSummaryTitle] = useState<string>('Summary');
  const [summaryTitleCell, setSummaryTitleCell] = useState<string>('');
  const [tableFormatting, setTableFormatting] = useState<TableFormattingOptions>({
      fillColor: 'F8F9FA',
      borderStyle: 'thin',
      borderColor: 'DEE2E6',
  });
  const [insertColumn, setInsertColumn] = useState<string>('J');
  const [insertStartRow, setInsertStartRow] = useState<number>(1);
  const [summarySheetName, setSummarySheetName] = useState<string>('Cross-Sheet Summary');
  const [summaryHeaderFormatting, setSummaryHeaderFormatting] = useState<HeaderFormatOptions>({
    bold: true,
    italic: false,
    underline: false,
    fontName: 'Calibri',
    fontSize: 12,
    horizontalAlignment: 'center' as HorizontalAlignment,
    fillColor: 'EAEAEA'
  });
  const [enableTotalRowFormatting, setEnableTotalRowFormatting] = useState(true);
  const [enableBlankRowFormatting, setEnableBlankRowFormatting] = useState(true);
  const [showBlanksInInSheetSummary, setShowBlanksInInSheetSummary] = useState(true);
  const [totalRowFormatting, setTotalRowFormatting] = useState<HeaderFormatOptions>({ bold: true, fillColor: 'EAEAEA' });
  const [blankRowFormatting, setBlankRowFormatting] = useState<HeaderFormatOptions>({ bold: true, fillColor: 'FFF2CC', fontColor: '9C6500' });
  const [groupReportHeaderFormatting, setGroupReportHeaderFormatting] = useState<HeaderFormatOptions>({
    bold: true,
    italic: false,
    underline: false,
    fontName: 'Calibri',
    fontSize: 14,
    fillColor: 'D9EAD3'
  });
  const [groupReportHeaders, setGroupReportHeaders] = useState({
    groupName: 'Group Name',
    keyName: 'Key Name',
    count: 'Count',
  });
  const [groupReportSheetTitle, setGroupReportSheetTitle] = useState<string>('Group Summary Report');
  const [groupReportMultiSourceTitle, setGroupReportMultiSourceTitle] = useState<string>('');
  const [groupReportDescription, setGroupReportDescription] = useState<string>('');


  const [aggregationMode, setAggregationMode] = useState<'valueMatch' | 'keyMatch'>('valueMatch');
  const [keyCountColumn, setKeyCountColumn] = useState<string>('D');
  const [discoverNewKeys, setDiscoverNewKeys] = useState<boolean>(true);
  const [blankLabel, setBlankLabel] = useState<string>('(Blanks)');
  const [blankCountingMode, setBlankCountingMode] = useState<'rowAware' | 'fullColumn'>('rowAware');
  const [generateBlankDetails, setGenerateBlankDetails] = useState<boolean>(false);
  const [enableConditionalMatch, setEnableConditionalMatch] = useState<boolean>(false);

  
  const [reportingScope, setReportingScope] = useState<'main' | 'custom'>('main');
  const [customReportSheets, setCustomReportSheets] = useState<SheetSelection>({});


  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const cancellationRequested = useRef(false);
  const [livePreviewData, setLivePreviewData] = useState<Record<string, number> | null>(null);

  // For rule tester
  const [testSheetName, setTestSheetName] = useState<string>('');
  const [testRowNumber, setTestRowNumber] = useState<string>('2');
  const [isTestingRule, setIsTestingRule] = useState<boolean>(false);
  const [testResult, setTestResult] = useState<string | null>(null);
  
  const [outputFormat, setOutputFormat] = useState<'xlsx' | 'xlsm'>('xlsm');
  const [reportChunkSize, setReportChunkSize] = useState<number>(100000);
  const [reportLayout, setReportLayout] = useState<'sheetsAsRows' | 'keysAsRows'>('sheetsAsRows');
  const [autoSizeColumns, setAutoSizeColumns] = useState<boolean>(true);
  const [columnsToHide, setColumnsToHide] = useState<string>('');
  const [lastUpdateResult, setLastUpdateResult] = useState<UpdateResult | null>(null);
  const { toast } = useToast();
  
  // New state for Group Report
  const [groupMappings, setGroupMappings] = useState('');

  useEffect(() => {
    if (onProcessingChange) {
      onProcessingChange(isProcessing);
    }
  }, [isProcessing, onProcessingChange]);

  useEffect(() => {
    onFileStateChange(file !== null);
  }, [file, onFileStateChange]);

  useEffect(() => {
    // This effect now resets all state whenever the file changes, ensuring a clean slate.
    setSheetNames([]);
    setSelectedSheets({});
    setCustomReportSheets({});
    setHeaderRow(1);
    setSearchColumns('A');
    setSearchTerms('');
    setMappingFile(null);
    setMappingFileContent('');
    setMatchMode('whole');
    setStripFormulas(false);
    setEnableUpdate(false);
    setUpdateOnlyBlanks(true);
    setUpdateColumn('B');
    setGenerateUpdateReport(false);
    setEnableFillKeyColumn(false);
    setEnablePairedRowValidation(false);
    setPairedValidationColumns('');
    setEnableMarking(false);
    setMarkColumn('');
    setMarkValue('Processed');
    setAggregationResult(null);
    setFinalAggregationResult(null);
    setEditableKeys(new Map());
    setInsertResultsInSheet(false);
    setClearExistingInSheetSummary(true);
    setInSheetDataSource('reportingScope');
    setShowOnlyLocalKeys(true);
    setInSheetGenerationMode('static');
    setInSheetSummaryTitle('Summary');
    setSummaryTitleCell('');
    setTableFormatting({ fillColor: 'F8F9FA', borderStyle: 'thin', borderColor: 'DEE2E6' });
    setInsertColumn('J');
    setInsertStartRow(1);
    setSummarySheetName('Cross-Sheet Summary');
    setSummaryHeaderFormatting({
      bold: true,
      italic: false,
      underline: false,
      fontName: 'Calibri',
      fontSize: 12,
      horizontalAlignment: 'center' as HorizontalAlignment,
      fillColor: 'EAEAEA'
    });
    setEnableTotalRowFormatting(true);
    setEnableBlankRowFormatting(true);
    setShowBlanksInInSheetSummary(true);
    setTotalRowFormatting({ bold: true, fillColor: 'EAEAEA' });
    setBlankRowFormatting({ bold: true, fillColor: 'FFF2CC', fontColor: '9C6500' });
    setAggregationMode('valueMatch');
    setKeyCountColumn('D');
    setDiscoverNewKeys(true);
    setBlankLabel('(Blanks)');
    setBlankCountingMode('rowAware');
    setGenerateBlankDetails(false);
    setEnableConditionalMatch(false);
    setReportingScope('main');
    setTestSheetName('');
    setTestRowNumber('2');
    setTestResult(null);
    setLivePreviewData(null);
    setReportChunkSize(100000);
    setReportLayout('sheetsAsRows');
    setAutoSizeColumns(true);
    setColumnsToHide('');
    setLastUpdateResult(null);
    setGroupMappings('');
    setGroupReportHeaderFormatting({
        bold: true,
        italic: false,
        underline: false,
        fontName: 'Calibri',
        fontSize: 14,
        fillColor: 'D9EAD3'
    });
    setGroupReportHeaders({
      groupName: 'Group Name',
      keyName: 'Key Name',
      count: 'Count',
    });
    setGroupReportSheetTitle('Group Summary Report');
    setGroupReportMultiSourceTitle('');
    setGroupReportDescription('');

    if (file) {
      const getSheetNamesFromFile = async () => {
        setIsProcessing(true);
        try {
          const arrayBuffer = await file.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'buffer', bookVBA: true, bookFiles: true });
          const names = workbook.SheetNames;
          setSheetNames(names);
          if (names.length > 0) {
            setTestSheetName(names[0]);
          }
          const initialSelection: SheetSelection = {};
          names.forEach(name => {
            initialSelection[name] = true; // Select all by default
          });
          setSelectedSheets(initialSelection);
          setCustomReportSheets(initialSelection);
        } catch (error) {
          console.error("Error reading sheet names:", error);
          toast({ title: t('toast.errorReadingFile') as string, description: t('toast.errorReadingSheets') as string, variant: "destructive" });
          setSheetNames([]);
          setSelectedSheets({});
          setCustomReportSheets({});
        } finally {
          setIsProcessing(false);
        }
      };
      getSheetNamesFromFile();
    }
  }, [file, toast, t]);
  
  useEffect(() => {
      if (inSheetDataSource === 'reportingScope' || aggregationMode !== 'keyMatch') {
          // Force static generation if conditions for formulas are not met.
          if (inSheetGenerationMode === 'formula') {
            setInSheetGenerationMode('static');
          }
      }
  }, [inSheetDataSource, aggregationMode, inSheetGenerationMode]);

  const handleGroupHeaderChange = (field: keyof typeof groupReportHeaders, value: string) => {
    setGroupReportHeaders(prev => ({ ...prev, [field]: value }));
  };

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      if (!selectedFile.name.match(/\.(xlsx|xls|xlsm)$/)) {
        toast({
          title: t('toast.invalidFileType') as string,
          description: t('toast.invalidFileTypeDesc') as string,
          variant: 'destructive',
        });
        setFile(null);
        return;
      }
      setFile(selectedFile);

      const formData = new FormData();
      formData.append('file', selectedFile);
      fetch('/api/upload', {
        method: 'POST',
        body: formData,
      }).catch(error => {
        console.error("Failed to save file to server:", error);
        toast({
            title: t('toast.uploadErrorTitle') as string,
            description: t('toast.uploadErrorDesc') as string,
            variant: "destructive"
        });
      });

    } else {
      setFile(null);
    }
  };

  const handleMappingFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const mappingFile = event.target.files?.[0];
    if (mappingFile) {
      if (!mappingFile.name.match(/\.(json|txt|csv)$/i)) {
        toast({
          title: t('aggregator.toast.invalidMappingFileTitle') as string,
          description: t('aggregator.toast.invalidMappingFileDesc') as string,
          variant: 'destructive',
        });
        setMappingFile(null);
        setMappingFileContent('');
        return;
      }
      setMappingFile(mappingFile);
      
      const formData = new FormData();
      formData.append('file', mappingFile);
      fetch('/api/upload', {
        method: 'POST',
        body: formData,
      }).catch(error => {
        console.error("Failed to save mapping file to server:", error);
        toast({
            title: t('toast.uploadErrorTitle') as string,
            description: t('toast.uploadErrorDesc') as string,
            variant: "destructive"
        });
      });

      const reader = new FileReader();
      reader.onload = (e) => {
        const content = e.target?.result as string;
        setMappingFileContent(content); // For tooltip

        let newSearchTerms = '';
        if (mappingFile.name.toLowerCase().endsWith('.json')) {
          try {
            const jsonObj = JSON.parse(content);
            if (typeof jsonObj === 'object' && jsonObj !== null && !Array.isArray(jsonObj)) {
              newSearchTerms = Object.entries(jsonObj)
                .map(([key, value]) => `${String(key).trim()} : ${String(value).trim()}`)
                .join('\n');
            } else {
               toast({ title: t('aggregator.toast.invalidJsonFormatTitle') as string, description: t('aggregator.toast.invalidJsonFormatDesc') as string, variant: 'destructive' });
               return;
            }
          } catch (error) {
             toast({ title: t('aggregator.toast.jsonParseErrorTitle') as string, description: t('aggregator.toast.jsonParseErrorDesc') as string, variant: 'destructive' });
             return;
          }
        } else { // .txt or .csv
          newSearchTerms = content;
        }
        setSearchTerms(newSearchTerms);
        toast({ title: t('aggregator.toast.mappingFileSuccessTitle') as string, description: t('aggregator.toast.mappingFileSuccessDesc', { fileName: mappingFile.name }) as string });
      };
      reader.onerror = () => {
        toast({ title: t('toast.errorReadingFile') as string, description: t('aggregator.toast.mappingFileError') as string, variant: 'destructive' });
      };
      reader.readAsText(mappingFile);
    }
  };

  const handleClearMappingFile = () => {
    setMappingFile(null);
    setMappingFileContent('');
    setSearchTerms('');
    const input = document.getElementById('mapping-file-upload') as HTMLInputElement;
    if (input) {
      input.value = '';
    }
  };

  const handleSelectAllSheets = (checked: boolean) => {
    const newSelection: SheetSelection = {};
    sheetNames.forEach(name => {
      newSelection[name] = checked;
    });
    setSelectedSheets(newSelection);
  };
  
  const handleSelectAllCustomSheets = (checked: boolean) => {
    const newSelection: SheetSelection = {};
    sheetNames.forEach(name => {
      newSelection[name] = checked;
    });
    setCustomReportSheets(newSelection);
  };

  const handlePartialSelection = (count: number) => {
    const newSelection: SheetSelection = {};
    sheetNames.forEach((name, index) => {
      newSelection[name] = index < count;
    });
    setSelectedSheets(newSelection);
  };

  const handleSheetSelectionChange = (sheetName: string, checked: boolean) => {
    setSelectedSheets(prev => ({ ...prev, [sheetName]: checked }));
  };
  
  const handleCustomSheetSelectionChange = (sheetName: string, checked: boolean) => {
    setCustomReportSheets(prev => ({ ...prev, [sheetName]: checked }));
  };

  const handleCancel = () => {
    cancellationRequested.current = true;
    setProcessingStatus(t('common.cancelling') as string);
  };

  const handleKeyEdit = (originalKey: string, newKey: string) => {
    setEditableKeys(prev => {
        const newMap = new Map(prev);
        newMap.set(originalKey, newKey);
        return newMap;
    });
  };

  const handleTestRule = useCallback(async () => {
      if (!file || !testSheetName || !testRowNumber) {
          toast({ title: t('toast.missingInfo') as string, description: t('aggregator.toast.missingTestInfo') as string, variant: 'destructive' });
          return;
      }
      const rowNum = parseInt(testRowNumber, 10);
      if (isNaN(rowNum) || rowNum < 1) {
          toast({ title: t('toast.missingInfo') as string, description: t('aggregator.toast.invalidTestRow') as string, variant: 'destructive' });
          return;
      }
  
      setIsTestingRule(true);
      setTestResult(t('common.processing') as string);
  
      try {
          const arrayBuffer = await file.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
          const worksheet = workbook.Sheets[testSheetName];
          if (!worksheet) {
              throw new Error(`Sheet "${testSheetName}" not found.`);
          }
          const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
          const headerRowIndex = headerRow - 1;
          const testRowIndex = rowNum - 1;
  
          if (testRowIndex >= aoa.length || headerRowIndex >= aoa.length) {
              throw new Error(`Row ${rowNum} is out of bounds for sheet "${testSheetName}".`);
          }
  
          const headers = aoa[headerRowIndex].map(h => String(h || ''));
          const rowData = aoa[testRowIndex];
          const searchColIndices = parseSourceColumns(searchColumns, headers);
          
          const valueToKeyMap = new Map<string, string>();
          searchTerms.split('\n').map(t => t.trim()).filter(Boolean).forEach(line => {
              const colonIndex = line.indexOf(':');
              if (colonIndex !== -1) {
                  const value = line.substring(0, colonIndex).trim().toLowerCase();
                  const key = line.substring(colonIndex + 1).trim();
                  if (value && key) valueToKeyMap.set(value, key);
              } else {
                  const key = line.trim();
                  if (key) valueToKeyMap.set(key.toLowerCase(), key);
              }
          });
  
            let resultText = t('aggregator.testRules.resultTitle', { sheetName: testSheetName, rowNum }) as string;
            
            if (aggregationMode === 'valueMatch' && enableConditionalMatch && keyCountColumn.trim()) {
                const conditionalColIdx = getColumnIndex(keyCountColumn, headers);
                if (conditionalColIdx !== null) {
                    const conditionalValue = rowData[conditionalColIdx];
                    const isConditionalCellEmpty = conditionalValue === null || conditionalValue === undefined || String(conditionalValue).trim() === '';
                    
                    if (!isConditionalCellEmpty) {
                         resultText += `\n` + t('aggregator.testRules.conditionalSkip', { columnName: keyCountColumn, value: String(conditionalValue) });
                         setTestResult(resultText);
                         setIsTestingRule(false);
                         return;
                    } else {
                         resultText += `\n` + t('aggregator.testRules.conditionalProceed', { columnName: keyCountColumn });
                    }
                }
            }

            resultText += `\n${t('aggregator.testRules.searchColsLabel')} ${searchColIndices.map(i => headers[i] || `Col ${i+1}`).join(', ')}\n`;

            const rowMatchScores = new Map<string, number>();

            for (const colIdx of searchColIndices) {
                const headerName = headers[colIdx] || `Col ${colIdx + 1}`;
                const cellValue = rowData[colIdx];
                const cellText = (cellValue !== null && cellValue !== undefined) ? String(cellValue) : '';

                resultText += `\n` + t('aggregator.testRules.columnValue', { headerName, cellText });

                if (!cellText) {
                    resultText += `\n` + t('aggregator.testRules.noValueInCell');
                    continue;
                }
                
                let matchesInCell = 0;
                for (const [searchTerm, reportingKey] of valueToKeyMap.entries()) {
                    let score = 0;
                    let isMatch = false;

                    if (matchMode === 'loose') {
                        const searchWords = searchTerm.split(/\s+/).filter(Boolean);
                        if (searchWords.length > 0) {
                            const allWordsFound = searchWords.every(word => {
                                const pattern = `(^|\\P{L})${escapeRegex(word)}(\\P{L}|$)`;
                                const regex = new RegExp(pattern, 'iu');
                                return regex.test(cellText);
                            });
                            if (allWordsFound) {
                                isMatch = true;
                                score = searchWords.length;
                            }
                        }
                    } else {
                        const escapedSearchTerm = escapeRegex(searchTerm);
                        const pattern = matchMode === 'whole'
                            ? `(^|\\P{L})${escapedSearchTerm}(\\P{L}|$)`
                            : escapedSearchTerm;
                        const searchRegex = new RegExp(pattern, 'giu');
                        
                        const matches = cellText.match(searchRegex);
                        if (matches) {
                            isMatch = true;
                            score = matches.length;
                        }
                    }

                    if (isMatch) {
                        const currentScore = rowMatchScores.get(reportingKey) || 0;
                        rowMatchScores.set(reportingKey, Math.max(currentScore, score));
                        matchesInCell++;
                        resultText += `\n` + t('aggregator.testRules.keywordMatch', { count: score, keyword: searchTerm, reportingKey });
                    }
                }
                
                if (matchesInCell === 0) {
                    resultText += `\n` + t('aggregator.testRules.noKeywordsMatched');
                }
            }

            if (rowMatchScores.size > 0) {
                let winnerKey = '';
                let maxScore = -1;

                for (const [key, score] of rowMatchScores.entries()) {
                    if (score > maxScore) {
                        maxScore = score;
                        winnerKey = key;
                    } else if (score === maxScore) {
                        if (key < winnerKey) {
                            winnerKey = key;
                        }
                    }
                }

                const breakdownItems = Array.from(rowMatchScores.entries()).map(([key, score]) => `${key} (score: ${score})`);
                resultText += `\n\n` + t('aggregator.testRules.matchBreakdown', { breakdown: breakdownItems.join(', ') });
                resultText += `\n` + t('aggregator.testRules.winner', { winnerKey });
                resultText += `\n` + t('aggregator.testRules.summaryMessage', { key: winnerKey });

            } else {
                resultText += `\n\n` + t('aggregator.testRules.noMatchesFound');
            }

            setTestResult(resultText);
  
      } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          setTestResult(`${t('aggregator.testRules.errorPrefix')} ${message}`);
          toast({ title: t('toast.errorReadingFile') as string, description: message, variant: 'destructive' });
      } finally {
          setIsTestingRule(false);
      }
  }, [file, testSheetName, testRowNumber, headerRow, searchColumns, searchTerms, toast, t, aggregationMode, enableConditionalMatch, keyCountColumn, matchMode]);
  

  const handleProcess = useCallback(async () => {
    const sheetsToSearch = sheetNames.filter(name => selectedSheets[name]);
    
    const valueToKeyMap = new Map<string, string>();
    const lines = searchTerms.split('\n').map(t => t.trim()).filter(Boolean);
    const initialReportingKeys = new Set<string>();

    lines.forEach(line => {
        const colonIndex = line.indexOf(':');
        if (colonIndex !== -1) {
            const value = line.substring(0, colonIndex).trim().toLowerCase();
            const key = line.substring(colonIndex + 1).trim();
            if (value && key) {
                valueToKeyMap.set(value, key);
                initialReportingKeys.add(key);
            }
        } else {
            const key = line.trim();
            if (key) {
                valueToKeyMap.set(key.toLowerCase(), key);
                initialReportingKeys.add(key);
            }
        }
    });

    if (!file || sheetsToSearch.length === 0 || searchColumns.trim() === '' || headerRow < 1) {
      toast({
        title: t('toast.missingInfo') as string,
        description: t('aggregator.toast.missingInfo') as string,
        variant: 'destructive',
      });
      return;
    }

    cancellationRequested.current = false;
    setIsProcessing(true);
    setLivePreviewData(null);
    setProcessingStatus('');
    setAggregationResult(null);
    setFinalAggregationResult(null);
    setEditableKeys(new Map());
    setLastUpdateResult(null);

    const onProgress = (status: { stage: string; sheetName: string; currentSheet: number; totalSheets: number; currentTotals: Record<string, number>}) => {
        if (cancellationRequested.current) {
            throw new Error('Cancelled by user.');
        }
        
        const stagePrefix = `${status.stage} - `;
        const baseMessage = t('aggregator.toast.processingSheet', {
            current: status.currentSheet, 
            total: status.totalSheets, 
            sheetName: status.sheetName, 
            count: ''
        }) as string;

        setProcessingStatus(stagePrefix + baseMessage);

        // Only show live data during the counting phase.
        if (status.stage.startsWith('Aggregating')) {
            setLivePreviewData(status.currentTotals);
        } else {
            setLivePreviewData(null);
        }
    };

    try {
        console.log("Starting data aggregation...");
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer', bookVBA: true, bookFiles: true });
        
        if (stripFormulas) {
            setProcessingStatus(t('aggregator.toast.strippingFormulas') as string);
            stripFormulasInWorkbook(workbook, sheetsToSearch);
        }

        let blankCol;
        let effectiveAggregationMode = aggregationMode;
        if (aggregationMode === 'keyMatch' && keyCountColumn.trim()) {
            blankCol = keyCountColumn;
        } else if (aggregationMode === 'valueMatch' && enableUpdate && updateColumn.trim()) {
            blankCol = updateColumn;
            // If update is enabled in valueMatch, we must check the update column for blanks.
            // But if conditional match is also on, it takes precedence.
            if (!enableConditionalMatch) {
                effectiveAggregationMode = 'valueMatch';
            }
        }


        const aggResult = aggregateData(
            workbook, 
            sheetsToSearch, 
            searchColumns, 
            valueToKeyMap,
            headerRow,
            { 
              aggregationMode: effectiveAggregationMode,
              discoverNewKeys: aggregationMode === 'keyMatch' ? discoverNewKeys : false,
              keyCountColumn: (aggregationMode === 'keyMatch' || (aggregationMode === 'valueMatch' && enableConditionalMatch)) ? keyCountColumn : undefined,
              conditionalColumn: (aggregationMode === 'valueMatch' && enableConditionalMatch) ? keyCountColumn : undefined,
              generateBlankDetails,
              countBlanksInColumn: blankCol,
              blankCountingMode: blankCountingMode,
              matchMode: matchMode,
              summaryTitleCell,
            },
            onProgress
        );
        setAggregationResult(aggResult);

        const initialEditableKeys = new Map<string, string>();
        aggResult.reportingKeys.forEach(key => {
            initialEditableKeys.set(key, key);
        });
        const trimmedBlankLabel = blankLabel.trim() || '(Blanks)';
        if (aggResult.blankCounts && aggResult.blankCounts.total > 0 && trimmedBlankLabel) {
            initialEditableKeys.set(trimmedBlankLabel, trimmedBlankLabel);
        }
        setEditableKeys(initialEditableKeys);

        toast({
            title: t('aggregator.toast.processSuccessTitle') as string,
            description: t('aggregator.toast.processSuccessDesc', { count: initialEditableKeys.size }) as string,
            action: <CheckCircle2 className="text-green-500" />,
        });

    } catch (error) {
      console.error('Error during processing:', error);
      const errorMessage = error instanceof Error ? error.message : 'An unknown error occurred.';
      if (errorMessage !== 'Cancelled by user.') {
          toast({ title: t('toast.errorReadingFile') as string, description: t('aggregator.toast.error', {errorMessage}) as string, variant: 'destructive' });
      } else {
        toast({ title: t('toast.cancelledTitle') as string, description: t('toast.cancelledDesc') as string, variant: 'default' });
      }
    } finally {
      setIsProcessing(false);
      cancellationRequested.current = false;
      setProcessingStatus('');
    }
  }, [file, sheetNames, selectedSheets, headerRow, searchColumns, searchTerms, toast, enableUpdate, updateColumn, aggregationMode, keyCountColumn, blankLabel, generateBlankDetails, discoverNewKeys, t, enableConditionalMatch, stripFormulas, blankCountingMode, matchMode, summaryTitleCell]);

  const modifiedDataForDisplay = useMemo(() => {
    return getModifiedAggregationData(aggregationResult, editableKeys, blankLabel);
  }, [aggregationResult, editableKeys, blankLabel]);

  const handleDownloadFinalWorkbook = useCallback(async () => {
    if (!file || !aggregationResult) {
        toast({ title: t('toast.noDataToDownload') as string, description: t('aggregator.toast.missingInfo') as string, variant: 'destructive' });
        return;
    }
    
    const needsModification = enableUpdate || insertResultsInSheet || enableFillKeyColumn || enableMarking;
    if (!needsModification) {
        toast({ title: t('aggregator.toast.noModificationsTitle') as string, description: t('aggregator.toast.noModificationsDesc') as string, variant: 'default' });
        return;
    }

    setIsProcessing(true);
    setProcessingStatus(t('aggregator.toast.generatingWorkbook') as string);
    setLastUpdateResult(null);

    try {
        const processedData = getModifiedAggregationData(aggregationResult, editableKeys, blankLabel);
        if (!processedData) {
            throw new Error("Could not process final key mappings.");
        }
        
        const arrayBuffer = await file.arrayBuffer();
        let modifiedWorkbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, bookVBA: true, bookFiles: true });
        
        const sheetsToModify = sheetNames.filter(name => selectedSheets[name]);

        // Step 1: Apply all structural changes to the workbook
        if (stripFormulas) {
            setProcessingStatus(t('aggregator.toast.strippingFormulas') as string);
            stripFormulasInWorkbook(modifiedWorkbook, sheetsToModify);
        }

        let updateResult: UpdateResult | null = null;
        if (enableFillKeyColumn && aggregationMode === 'valueMatch' && keyCountColumn.trim()) {
            fillEmptyKeyColumn(modifiedWorkbook, sheetsToModify, searchColumns, keyCountColumn, headerRow, processedData.modifiedValueToKeyMap, matchMode);
        }
        if (enableUpdate) {
            updateResult = lookupAndAndUpdate(
                modifiedWorkbook, sheetsToModify, searchColumns, updateColumn, headerRow, 
                processedData.modifiedValueToKeyMap, updateOnlyBlanks, matchMode, enablePairedRowValidation, pairedValidationColumns
            );
            setLastUpdateResult(updateResult);
        }

        if (enableMarking) {
            const preMarkingScan = aggregateData(modifiedWorkbook, sheetsToModify, searchColumns, processedData.modifiedValueToKeyMap, headerRow, { 
                aggregationMode, keyCountColumn, conditionalColumn: enableConditionalMatch ? keyCountColumn : undefined, matchMode,
            });
            if (preMarkingScan.matchingRows) {
                markMatchingRows(modifiedWorkbook, sheetsToModify, preMarkingScan.matchingRows, markColumn, markValue, headerRow);
            }
        }

        // Step 2: Verification Scan - Rerun aggregation on the *modified* workbook to get final, accurate counts.
        setProcessingStatus(t('aggregator.toast.recalculating') as string);
        
        const sheetsForReporting = reportingScope === 'custom'
            ? sheetNames.filter(name => customReportSheets[name])
            : sheetsToModify;

        const effectiveMode = (enableFillKeyColumn || enableUpdate) ? 'keyMatch' : aggregationMode;
        const effectiveKeyColumn = enableFillKeyColumn ? keyCountColumn : (enableUpdate ? updateColumn : keyCountColumn);

        // Build the correct key map for the final aggregation scan.
        const finalKeyMapForRecalculation = new Map<string, string>();
        if (effectiveMode === 'keyMatch') {
            // In keyMatch mode, we match against the final keys themselves, as they may have just been written.
            processedData.modifiedResult.reportingKeys.forEach(k => finalKeyMapForRecalculation.set(k.toLowerCase(), k));
        } else { // valueMatch mode
            // In valueMatch, we still need the original search term -> final key mapping.
            processedData.modifiedValueToKeyMap.forEach((finalKey, originalSearchTerm) => {
                finalKeyMapForRecalculation.set(originalSearchTerm, finalKey);
            });
        }
        
        // Run the final, definitive aggregation scan on the MODIFIED workbook.
        const finalAggregationResult = aggregateData(
            modifiedWorkbook, 
            sheetsForReporting, 
            searchColumns, 
            finalKeyMapForRecalculation, 
            headerRow, 
            {
                aggregationMode: effectiveMode,
                keyCountColumn: effectiveKeyColumn,
                matchMode: matchMode,
                blankCountingMode: blankCountingMode,
                countBlanksInColumn: effectiveMode === 'keyMatch' ? effectiveKeyColumn : (aggregationMode === 'valueMatch' && enableUpdate ? updateColumn : undefined),
                generateBlankDetails: generateBlankDetails,
                discoverNewKeys: false, // Keys are final at this stage.
            }
        );
        
        setFinalAggregationResult(finalAggregationResult);


        // Step 3: Use the definitive final results to build reports
        const summaryConfig: SummaryConfig = {
            summarySheetName, 
            aggregationMode: effectiveMode,
            keyCountColumn: effectiveKeyColumn,
            insertColumn, insertStartRow, headerFormatting: summaryHeaderFormatting,
            countBlanksInColumn: finalAggregationResult.blankCounts ? effectiveKeyColumn : undefined, 
            blankCountLabel: processedData.finalBlankLabel,
            blankRowFormatting: (insertResultsInSheet && enableBlankRowFormatting) ? blankRowFormatting : undefined,
            totalRowFormatting: (insertResultsInSheet && enableTotalRowFormatting) ? totalRowFormatting : undefined,
            chunkSize: reportChunkSize, inSheetSummaryTitle: inSheetSummaryTitle, reportLayout: reportLayout,
            showBlanksInInSheetSummary: showBlanksInInSheetSummary, autoSizeColumns, columnsToHide, blankCountingMode,
            clearExistingInSheetSummary,
            showOnlyLocalKeysInSummary: showOnlyLocalKeys,
            summaryTitleCell,
            tableFormatting,
        };
        
        if (insertResultsInSheet) {
           insertAggregationResultsIntoSheets(modifiedWorkbook, finalAggregationResult, sheetsToModify, headerRow, summaryConfig, { dataSource: inSheetDataSource, generationMode: inSheetGenerationMode, showOnlyLocalKeys }, processedData.editedKeyToOriginalsMap);
        }
        
        if ((finalAggregationResult.reportingKeys.length > 0 || (finalAggregationResult.blankCounts && finalAggregationResult.blankCounts.total > 0))) {
            addAggregationReportSheetsToWorkbook(modifiedWorkbook, aggregationResult, finalAggregationResult, summaryConfig, processedData.editedKeyToOriginalsMap);
        }
        
        if (generateUpdateReport && updateResult && updateResult.summary.totalCellsUpdated > 0) {
            addUpdateReportSheetToWorkbook(modifiedWorkbook, updateResult, reportChunkSize);
        }
        
        const originalFileName = file.name.substring(0, file.name.lastIndexOf('.'));
        let suffix = '_with_aggregates';
        if(enableMarking) suffix = '_marked';
        else if (enableUpdate && enableFillKeyColumn) suffix = '_fully_updated';
        else if (enableUpdate) suffix = '_updated';
        else if (enableFillKeyColumn) suffix = '_keys_filled';
        
        XLSX.writeFile(modifiedWorkbook, `${originalFileName}${suffix}.${outputFormat}`, { compression: true, bookVBA: true, bookType: outputFormat });
        
        toast({ title: t('toast.downloadSuccess') as string });
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error('Error creating or downloading workbook:', error);
        toast({ 
            title: t('toast.downloadError') as string, 
            description: t('aggregator.toast.error', {errorMessage: errorMessage}) as string, 
            variant: 'destructive' 
        });
    } finally {
        setIsProcessing(false);
    }
  }, [
      file, aggregationResult, editableKeys, blankLabel, toast, enableUpdate, insertResultsInSheet, 
      outputFormat, t, sheetNames, selectedSheets, searchColumns, headerRow, insertColumn, 
      insertStartRow, summarySheetName, aggregationMode, keyCountColumn, summaryHeaderFormatting, 
      enableBlankRowFormatting, enableTotalRowFormatting, blankRowFormatting, totalRowFormatting, 
      updateColumn, generateUpdateReport, enableFillKeyColumn, discoverNewKeys, generateBlankDetails, 
      reportingScope, customReportSheets, reportChunkSize, inSheetDataSource, 
      inSheetGenerationMode, enableConditionalMatch, enableMarking, markColumn, markValue, 
      inSheetSummaryTitle, reportLayout, updateOnlyBlanks, enablePairedRowValidation, 
      pairedValidationColumns, stripFormulas, blankCountingMode, showBlanksInInSheetSummary, 
      matchMode, autoSizeColumns, columnsToHide, clearExistingInSheetSummary, showOnlyLocalKeys,
      summaryTitleCell, tableFormatting
  ]);

  const handleDownloadReport = useCallback(async () => {
    if (!file || !aggregationResult) {
      toast({ title: t('toast.noDataToDownload') as string, description: t('aggregator.toast.missingInfo') as string, variant: 'destructive' });
      return;
    }

    setIsProcessing(true);
    setProcessingStatus(t('aggregator.toast.generatingReport') as string);

    try {
        const processedData = getModifiedAggregationData(aggregationResult, editableKeys, blankLabel);
        if (!processedData) {
            setIsProcessing(false);
            return;
        }
        const { modifiedResult, editedKeyToOriginalsMap } = processedData;

        let reportAggregationResult: AggregationResult;
        
        const initialProcessSheets = new Set(aggregationResult.processedSheetNames);
        const sheetsForReporting = reportingScope === 'custom'
            ? sheetNames.filter(name => customReportSheets[name])
            : sheetNames.filter(name => selectedSheets[name]);
        const reportingSheetsSet = new Set(sheetsForReporting);
        
        const scopesDiffer = initialProcessSheets.size !== reportingSheetsSet.size || !Array.from(initialProcessSheets).every(item => reportingSheetsSet.has(item));

        if (scopesDiffer) {
            setProcessingStatus(t('aggregator.toast.reaggregating') as string);
            const arrayBuffer = await file.arrayBuffer();
            const workbookForReporting = XLSX.read(arrayBuffer, { type: 'buffer', bookVBA: true, bookFiles: true });
            
            if (stripFormulas) {
                setProcessingStatus(t('aggregator.toast.strippingFormulas') as string);
                stripFormulasInWorkbook(workbookForReporting, sheetsForReporting);
            }

            let blankCol;
            if (aggregationMode === 'keyMatch' && keyCountColumn.trim()) {
                blankCol = keyCountColumn;
            } else if (aggregationMode === 'valueMatch' && enableUpdate && updateColumn.trim()) {
                blankCol = updateColumn;
            }
             reportAggregationResult = aggregateData(
                workbookForReporting, 
                sheetsForReporting,
                searchColumns, 
                processedData.modifiedValueToKeyMap,
                headerRow,
                { 
                  aggregationMode: aggregationMode,
                  discoverNewKeys: false, // Keys are already finalized by user edits
                  keyCountColumn: (aggregationMode === 'keyMatch' || (aggregationMode === 'valueMatch' && enableConditionalMatch)) ? keyCountColumn : undefined,
                  conditionalColumn: (aggregationMode === 'valueMatch' && enableConditionalMatch) ? keyCountColumn : undefined,
                  generateBlankDetails,
                  countBlanksInColumn: blankCol,
                  blankCountingMode: blankCountingMode,
                  matchMode: matchMode,
                }
            );
        } else {
             reportAggregationResult = modifiedResult;
        }
        
        const summaryConfig: SummaryConfig = {
            summarySheetName,
            aggregationMode,
            keyCountColumn,
            headerFormatting: summaryHeaderFormatting,
            countBlanksInColumn: aggregationMode === 'keyMatch' ? keyCountColumn : (enableUpdate ? updateColumn : undefined),
            blankCountLabel: processedData.finalBlankLabel,
            blankRowFormatting: enableBlankRowFormatting ? blankRowFormatting : undefined,
            totalRowFormatting: enableTotalRowFormatting ? totalRowFormatting : undefined,
            chunkSize: reportChunkSize,
            reportLayout: reportLayout,
            autoSizeColumns,
            columnsToHide,
            blankCountingMode,
        };

        const reportWb = createAggregationReportWorkbook(aggregationResult, reportAggregationResult, summaryConfig, editedKeyToOriginalsMap);

        if (enableUpdate && generateUpdateReport) {
            setProcessingStatus(t('aggregator.toast.simulatingUpdates') as string);
            const arrayBuffer = await file.arrayBuffer();
            const workbookForReporting = XLSX.read(arrayBuffer, { type: 'buffer', bookVBA: true, bookFiles: true });
            if (stripFormulas) stripFormulasInWorkbook(workbookForReporting, sheetsForReporting);

            const updateResult = findPotentialUpdates(
                workbookForReporting, 
                sheetsForReporting, 
                searchColumns, 
                updateColumn, 
                headerRow, 
                processedData.modifiedValueToKeyMap, 
                updateOnlyBlanks,
                matchMode,
                enablePairedRowValidation,
                pairedValidationColumns
            );
            
            if (updateResult && updateResult.summary.totalCellsUpdated > 0) {
                 addUpdateReportSheetToWorkbook(reportWb, updateResult, reportChunkSize);
            }
        }
        
        const originalFileName = file.name.substring(0, file.name.lastIndexOf('.')) + '_aggregation_report';
        XLSX.writeFile(reportWb, `${originalFileName}.${outputFormat}`, { compression: true, bookType: outputFormat });
        toast({ title: t('toast.downloadSuccess') as string });

    } catch(error) {
        console.error('Error creating report workbook:', error);
        toast({ title: t('toast.downloadError') as string, description: t('aggregator.toast.error', {errorMessage: 'Failed to create the report workbook.'}) as string, variant: 'destructive' });
    } finally {
        setIsProcessing(false);
    }
  }, [file, aggregationResult, editableKeys, blankLabel, toast, t, aggregationMode, keyCountColumn, enableUpdate, updateColumn, summarySheetName, summaryHeaderFormatting, enableBlankRowFormatting, enableTotalRowFormatting, blankRowFormatting, totalRowFormatting, outputFormat, generateUpdateReport, sheetNames, selectedSheets, searchColumns, headerRow, generateBlankDetails, reportingScope, customReportSheets, reportChunkSize, enableConditionalMatch, reportLayout, updateOnlyBlanks, enablePairedRowValidation, pairedValidationColumns, autoSizeColumns, columnsToHide, stripFormulas, blankCountingMode, matchMode]);

  const handleDownloadPdf = useCallback(async () => {
    if (!aggregationResult) {
      toast({ title: t('toast.noDataToDownload') as string, description: t('aggregator.toast.missingInfo') as string, variant: 'destructive' });
      return;
    }
    
    // Dynamically import client-side libraries
    const { default: jsPDF } = await import('jspdf');
    const { default: autoTable } = await import('jspdf-autotable');

    const processedData = getModifiedAggregationData(aggregationResult, editableKeys, blankLabel);
    if (!processedData) return;

    const { modifiedResult } = processedData;

    const doc = new jsPDF();
    let finalY = 22;

    doc.setFontSize(18);
    doc.text(t('aggregator.pdf.title') as string, 14, finalY);
    finalY += 10;
    
    // --- Grand Totals Table ---
    const grandTotalBody: (string | number)[][] = [];
    const sortedKeys = Object.keys(modifiedResult.totalCounts).sort((a,b) => (modifiedResult.totalCounts[b] || 0) - (modifiedResult.totalCounts[a] || 0));

    for (const key of sortedKeys) {
        grandTotalBody.push([String(key), modifiedResult.totalCounts[key] || 0]);
    }

    const grandTotalValue = Object.values(modifiedResult.totalCounts).reduce((sum, count) => sum + count, 0);
    
    autoTable(doc, {
        startY: finalY,
        head: [[t('aggregator.pdf.grandTotalsHeaderKey') as string, t('aggregator.pdf.grandTotalsHeaderCount') as string]],
        body: grandTotalBody,
        foot: [[t('aggregator.pdf.grandTotal') as string, grandTotalValue]],
        showFoot: 'lastPage',
        headStyles: { fillColor: [75, 85, 99] },
        footStyles: { fillColor: [243, 244, 246], textColor: [0, 0, 0], fontStyle: 'bold' },
        theme: 'striped',
        didDrawPage: (data) => {
            finalY = data.cursor?.y || finalY;
        }
    });

    // --- Per-Sheet Breakdown ---
    const perSheetTitleY = finalY + 15;
    if (doc.internal.pageSize.height < perSheetTitleY) {
        doc.addPage();
        finalY = 22;
    } else {
        finalY = perSheetTitleY;
    }

    doc.setFontSize(16);
    doc.text(t('aggregator.pdf.perSheetTitle') as string, 14, finalY);
    finalY += 10;

    const sheetOrder = (modifiedResult.processedSheetNames || []).filter(
        sheetName => {
            const sheetTotal = Object.values(modifiedResult.perSheetCounts[sheetName] || {}).reduce((s, c) => s + c, 0);
            return sheetTotal > 0;
        }
    );

    sheetOrder.forEach(sheetName => {
        const sheetCounts = modifiedResult.perSheetCounts[sheetName] || {};
        
        const sheetTableBody: (string | number)[][] = [];
        const sortedSheetKeys = Object.keys(sheetCounts).sort((a,b) => (sheetCounts[b] || 0) - (sheetCounts[a] || 0));

        for (const key of sortedSheetKeys) {
            const count = sheetCounts[key];
            if (count > 0) {
                 sheetTableBody.push([String(key), count]);
            }
        }
        
        if (sheetTableBody.length > 0) {
            const sheetTotalValue = sheetTableBody.reduce((sum, row) => sum + (row[1] as number), 0);

            if (finalY > 32) {
                 finalY += 5;
            }

            if (doc.internal.pageSize.height < finalY + 10) { 
                 doc.addPage();
                 finalY = 22;
            }
           
            doc.setFontSize(12);
            doc.setFont('helvetica', 'bold');
            doc.text(`${t('aggregator.pdf.sheetLabel')} ${sheetName}`, 14, finalY);
            doc.setFont('helvetica', 'normal');
            finalY += 7;

            autoTable(doc, {
                startY: finalY,
                head: [[t('aggregator.pdf.perSheetHeaderKey') as string, t('aggregator.pdf.perSheetHeaderCount') as string]],
                body: sheetTableBody,
                foot: [[t('aggregator.pdf.perSheetTotal') as string, sheetTotalValue]],
                showFoot: 'lastPage',
                headStyles: { fillColor: [156, 163, 175], fontSize: 10 },
                footStyles: { fillColor: [243, 244, 246], textColor: [0, 0, 0], fontStyle: 'bold', fontSize: 9 },
                theme: 'grid',
                styles: { fontSize: 9 },
                didDrawPage: (data) => {
                    finalY = data.cursor?.y || finalY;
                }
            });
        }
    });

    const fileName = file ? `${file.name.substring(0, file.name.lastIndexOf('.'))}_summary.pdf` : 'aggregation_summary.pdf';
    doc.save(fileName);
    
    toast({ title: t('toast.downloadSuccess') as string, description: t('aggregator.toast.pdfSuccess') as string });
  }, [aggregationResult, editableKeys, blankLabel, t, file, toast]);

    const handleDownloadGroupReport = useCallback(async (resultSource: 'initial' | 'final') => {
        const sourceResult = resultSource === 'initial' ? aggregationResult : finalAggregationResult;
        
        const processedData = getModifiedAggregationData(sourceResult, editableKeys, blankLabel);
        if (!processedData?.modifiedResult || !groupMappings.trim() || !file) {
            toast({ title: t('toast.missingInfo') as string, description: t('aggregator.toast.groupReportToastDesc') as string, variant: 'destructive' });
            return;
        }

        setIsProcessing(true);
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });

            const sheetsToReportOn = reportingScope === 'custom'
                ? sheetNames.filter(name => customReportSheets[name])
                : sheetNames.filter(name => selectedSheets[name]);

            const summaryConfig: SummaryConfig = {
                summarySheetName, 
                aggregationMode,
                keyCountColumn,
                insertColumn, insertStartRow, headerFormatting: summaryHeaderFormatting,
                countBlanksInColumn: aggregationResult?.blankCounts ? (aggregationMode === 'keyMatch' ? keyCountColumn : updateColumn) : undefined, 
                blankCountLabel: processedData.finalBlankLabel,
                blankRowFormatting,
                totalRowFormatting,
                chunkSize: reportChunkSize,
                inSheetSummaryTitle: inSheetSummaryTitle,
                reportLayout: reportLayout,
                showBlanksInInSheetSummary: showBlanksInInSheetSummary,
                autoSizeColumns,
                columnsToHide,
                clearExistingInSheetSummary,
                showOnlyLocalKeysInSummary: showOnlyLocalKeys,
                summaryTitleCell,
                tableFormatting,
                groupReportSheetTitle,
                groupReportMultiSourceTitle,
                groupReportHeaderFormatting,
                groupReportDescription,
            };
            
            const inSheetOptions = {
                dataSource: inSheetDataSource,
                showOnlyLocalKeys
            };

            const reportWb = createGroupReportWorkbook(
                processedData.modifiedResult, 
                workbook,
                groupMappings,
                sheetsToReportOn,
                summaryConfig,
                inSheetOptions,
                groupReportHeaders
            );
            const fileName = resultSource === 'initial' ? "Preliminary_Group_Report.xlsx" : "Final_Group_Report.xlsx";
            XLSX.writeFile(reportWb, fileName, { compression: true, bookType: 'xlsx' });
            toast({ title: t('toast.downloadSuccess') as string });
        } catch (error) {
            const errorMessage = error instanceof Error ? error.message : "Failed to create group report.";
            toast({ title: t('toast.downloadError') as string, description: errorMessage, variant: 'destructive' });
        } finally {
            setIsProcessing(false);
        }

    }, [
        file, aggregationResult, finalAggregationResult, editableKeys, blankLabel, groupMappings, toast, t,
        reportingScope, sheetNames, customReportSheets, selectedSheets, summarySheetName,
        aggregationMode, keyCountColumn, updateColumn, insertColumn, insertStartRow,
        summaryHeaderFormatting, blankRowFormatting, totalRowFormatting, reportChunkSize,
        inSheetSummaryTitle, reportLayout, showBlanksInInSheetSummary, autoSizeColumns,
        columnsToHide, clearExistingInSheetSummary, showOnlyLocalKeys, inSheetDataSource,
        summaryTitleCell, tableFormatting, groupReportSheetTitle, groupReportMultiSourceTitle, groupReportHeaderFormatting, groupReportHeaders, groupReportDescription
    ]);
  
  const allSheetsSelected = sheetNames.length > 0 && sheetNames.every(name => selectedSheets[name]);
  const allCustomSheetsSelected = sheetNames.length > 0 && sheetNames.every(name => customReportSheets[name]);
  const hasResults = aggregationResult !== null;
  const showBlankLabelInput = (aggregationMode === 'keyMatch' && !!keyCountColumn.trim()) || (aggregationMode === 'valueMatch' && enableUpdate && !!updateColumn.trim());
  
  const finalSortedKeysForDisplay = useMemo(() => {
    if (!modifiedDataForDisplay?.modifiedResult?.totalCounts) return [];
    return Object.entries(modifiedDataForDisplay.modifiedResult.totalCounts)
      .sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true, sensitivity: 'base' }));
  }, [modifiedDataForDisplay]);

  const grandTotalCount = useMemo(() => {
    if (!modifiedDataForDisplay?.modifiedResult?.totalCounts) return 0;
    return Object.values(modifiedDataForDisplay.modifiedResult.totalCounts).reduce((a, b) => a + b, 0);
  }, [modifiedDataForDisplay]);

  const canDownloadModifiedWorkbook = enableUpdate || insertResultsInSheet || enableFillKeyColumn || enableMarking;



  return (
    <Card className="w-full max-w-lg md:max-w-xl lg:max-w-2xl xl:max-w-6xl shadow-xl relative">
      {isProcessing && (
        <div className="absolute inset-0 bg-background/80 backdrop-blur-sm flex flex-col items-center justify-center z-10 rounded-lg space-y-4 p-4">
            <div className="flex items-center gap-2 text-muted-foreground">
                <Loader2 className="h-6 w-6 animate-spin" />
                <span className="text-lg font-medium">{processingStatus || t('common.processing')}</span>
            </div>
            {livePreviewData && (
                 <Card className="w-full max-w-md bg-background/90">
                    <CardHeader className="p-3">
                        <CardTitle className="text-base">{t('aggregator.liveSummary.title')}</CardTitle>
                    </CardHeader>
                    <CardContent className="p-3 pt-0 max-h-48 overflow-y-auto">
                        <Table>
                            <TableHeader>
                                <TableRow>
                                    <TableHead>{t('aggregator.liveSummary.keyHeader')}</TableHead>
                                    <TableHead className="text-right">{t('aggregator.liveSummary.countHeader')}</TableHead>
                                </TableRow>
                            </TableHeader>
                            <TableBody>
                                {Object.entries(livePreviewData).map(([key, value]) => (
                                    <TableRow key={key}>
                                        <TableCell className="font-medium">{key}</TableCell>
                                        <TableCell className="text-right">{value}</TableCell>
                                    </TableRow>
                                ))}
                            </TableBody>
                        </Table>
                    </CardContent>
                </Card>
            )}
            <Button variant="destructive" onClick={handleCancel}>
                <XCircle className="mr-2 h-4 w-4"/>
                {t('common.cancel')}
            </Button>
        </div>
      )}
      <CardHeader>
        <div className="flex items-center space-x-2 mb-2">
          <Sigma className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{t('aggregator.title')}</CardTitle>
        </div>
        <CardDescription className="font-body">
          {t('aggregator.description')}
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        
        <div className="space-y-2">
          <Label htmlFor="file-upload-aggregator" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{t('aggregator.uploadStep')}</span>
          </Label>
          <Input
            id="file-upload-aggregator"
            type="file"
            accept=".xlsx, .xls, .xlsm"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isProcessing}
          />
          {file && <p className="text-xs text-muted-foreground font-code">{t('common.selectedFile', {fileName: file.name})}</p>}
        </div>

        {sheetNames.length > 0 && (
          <div className="space-y-3">
              <Label className="flex items-center space-x-2 text-sm font-medium mb-2">
                  <ListChecks className="h-5 w-5" />
                  <span>{t('aggregator.selectSheetsStep')}</span>
              </Label>
              <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
                  <Checkbox
                      id="select-all-sheets-aggregator"
                      checked={allSheetsSelected}
                      onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                      aria-label="Select all sheets"
                      disabled={isProcessing}
                  />
                  <Label htmlFor="select-all-sheets-aggregator" className="text-sm font-medium flex-grow">
                      {t('common.selectAll')} ({t('common.selectedCount', {selected: Object.values(selectedSheets).filter(Boolean).length, total: sheetNames.length})})
                  </Label>
                  {sheetNames.length > 50 && (
                    <DropdownMenu>
                        <DropdownMenuTrigger asChild>
                            <Button variant="outline" size="sm" disabled={isProcessing}>
                                {t('common.partial')}
                                <ChevronDown className="ml-1 h-4 w-4" />
                            </Button>
                        </DropdownMenuTrigger>
                        <DropdownMenuContent>
                            <DropdownMenuItem onSelect={() => handlePartialSelection(50)}>{t('common.first50')}</DropdownMenuItem>
                            {sheetNames.length >= 100 && <DropdownMenuItem onSelect={() => handlePartialSelection(100)}>{t('common.first100')}</DropdownMenuItem>}
                            {sheetNames.length >= 150 && <DropdownMenuItem onSelect={() => handlePartialSelection(150)}>{t('common.first150')}</DropdownMenuItem>}
                        </DropdownMenuContent>
                    </DropdownMenu>
                  )}
              </div>
              <Card className="max-h-48 overflow-y-auto p-3 bg-background">
                  <div className="space-y-2">
                  {sheetNames.map(name => (
                      <div key={name} className="flex items-center space-x-2">
                      <Checkbox
                          id={`sheet-agg-${name}`}
                          checked={selectedSheets[name] || false}
                          onCheckedChange={(checked) => handleSheetSelectionChange(name, checked as boolean)}
                          disabled={isProcessing}
                      />
                      <Label htmlFor={`sheet-agg-${name}`} className="text-sm font-normal">
                          {name}
                      </Label>
                      </div>
                  ))}
                  </div>
              </Card>
          </div>
        )}

        <div className="space-y-2">
            <Label htmlFor="header-row-agg" className="flex items-center space-x-2 text-sm font-medium">
                <FileSpreadsheet className="h-5 w-5" />
                <span>{t('aggregator.headerRowStep')}</span>
            </Label>
            <Input
                id="header-row-agg"
                type="number"
                min="1"
                value={headerRow}
                onChange={(e) => setHeaderRow(parseInt(e.target.value, 10) || 1)}
                disabled={isProcessing || !file}
            />
            <p className="text-xs text-muted-foreground">{t('aggregator.headerRowDesc')}</p>
        </div>
        
        <div className="space-y-2">
          <Label htmlFor="search-columns" className="flex items-center space-x-2 text-sm font-medium">
            <BarChart2 className="h-5 w-5" />
            <span>{t('aggregator.searchColsStep')}</span>
          </Label>
          <Input
            id="search-columns"
            type="text"
            value={searchColumns}
            onChange={(e) => setSearchColumns(e.target.value)}
            disabled={isProcessing || !file}
            placeholder={t('aggregator.searchColsPlaceholder') as string}
          />
          <p className="text-xs text-muted-foreground">
            <Markup text={t('aggregator.searchColsDesc') as string} />
          </p>
        </div>
        
        <div className="space-y-2">
          <div className="flex justify-between items-center">
            <Label htmlFor="search-terms" className="flex items-center space-x-2 text-sm font-medium">
                <ScrollText className="h-5 w-5" />
                <span>{t('aggregator.mappingsStep')}</span>
            </Label>
            <TooltipProvider>
                <Tooltip>
                    <TooltipTrigger asChild>
                        <Label htmlFor="mapping-file-upload" className="cursor-pointer text-sm font-medium text-primary hover:underline flex items-center gap-1">
                            <FileUp className="h-4 w-4" />
                            {mappingFile ? t('aggregator.mappingsFileReplace') : t('aggregator.mappingsFileUpload')}
                        </Label>
                    </TooltipTrigger>
                    <TooltipContent side="top" align="end" className="max-w-md bg-popover text-popover-foreground p-3 space-y-2">
                        <p className="font-semibold">{t('aggregator.mappingsFileTooltipTitle')}</p>
                        <p className="text-xs text-muted-foreground">{t('aggregator.mappingsFileTooltipDesc')}</p>
                        {mappingFileContent && (
                            <Card className="mt-2 max-h-48 overflow-y-auto bg-background">
                              <CardHeader className="p-2">
                                <CardTitle className="text-xs font-code">{mappingFile?.name}</CardTitle>
                              </CardHeader>
                              <CardContent className="p-2 pt-0">
                                  <pre className="text-xs whitespace-pre-wrap font-code">{mappingFileContent}</pre>
                              </CardContent>
                            </Card>
                        )}
                    </TooltipContent>
                </Tooltip>
            </TooltipProvider>
            <Input
                id="mapping-file-upload"
                type="file"
                accept=".json,.txt,.csv"
                onChange={handleMappingFileChange}
                className="hidden"
                disabled={isProcessing || !file}
            />
          </div>
          <Textarea
            id="search-terms"
            value={searchTerms}
            onChange={(e) => setSearchTerms(e.target.value)}
            disabled={isProcessing || !file}
            placeholder={t('aggregator.mappingsPlaceholder') as string}
            rows={5}
          />
          <div className="flex justify-between items-center min-h-[20px]">
            <p className="text-xs text-muted-foreground">
              <Markup text={t('aggregator.mappingsDesc') as string} />
            </p>
             {mappingFile && (
                <Button variant="link" size="sm" onClick={handleClearMappingFile} className="text-destructive hover:text-destructive h-auto p-0 hover:no-underline">
                    {t('aggregator.mappingsFileClear')}
                </Button>
            )}
          </div>
          <div className="space-y-3 pt-2">
            <Label className="text-sm font-medium">{t('aggregator.matchMode.title')}</Label>
            <RadioGroup value={matchMode} onValueChange={(v) => setMatchMode(v as MatchMode)} className="space-y-1">
                <Label htmlFor="match-mode-whole" className="flex items-start space-x-3 p-3 rounded-md border has-[:checked]:border-primary cursor-pointer">
                    <RadioGroupItem value="whole" id="match-mode-whole" className="mt-1" />
                    <div className="grid gap-1.5">
                        <span className="font-normal">{[t('aggregator.matchMode.whole')].flat().join(' ')}</span>
                        <p className="text-xs text-muted-foreground">{[t('aggregator.matchMode.wholeDesc')].flat().join(' ')}</p>
                    </div>
                </Label>
                <Label htmlFor="match-mode-partial" className="flex items-start space-x-3 p-3 rounded-md border has-[:checked]:border-primary cursor-pointer">
                    <RadioGroupItem value="partial" id="match-mode-partial" className="mt-1" />
                     <div className="grid gap-1.5">
                        <span className="font-normal">{[t('aggregator.matchMode.partial')].flat().join(' ')}</span>
                        <p className="text-xs text-muted-foreground">{[t('aggregator.matchMode.partialDesc')].flat().join(' ')}</p>
                    </div>
                </Label>
                <Label htmlFor="match-mode-loose" className="flex items-start space-x-3 p-3 rounded-md border has-[:checked]:border-primary cursor-pointer">
                    <RadioGroupItem value="loose" id="match-mode-loose" className="mt-1" />
                     <div className="grid gap-1.5">
                        <span className="font-normal">{[t('aggregator.matchMode.loose')].flat().join(' ')}</span>
                        <p className="text-xs text-muted-foreground">{[t('aggregator.matchMode.looseDesc')].flat().join(' ')}</p>
                    </div>
                </Label>
            </RadioGroup>
          </div>
        </div>

        {file && (
          <Card className="p-4 bg-secondary/30">
            <CardHeader className="p-0 pb-4">
              <Label className="flex items-center space-x-2 text-md font-semibold">
                <Search className="h-5 w-5" />
                <span>{t('aggregator.testRules.title')}</span>
              </Label>
              <CardDescription>{t('aggregator.testRules.description')}</CardDescription>
            </CardHeader>
            <CardContent className="p-0 space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="space-y-2">
                  <Label htmlFor="test-sheet-name">{t('aggregator.testRules.sheetToTest')}</Label>
                  <Select value={testSheetName} onValueChange={setTestSheetName} disabled={isTestingRule}>
                    <SelectTrigger id="test-sheet-name"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {sheetNames.map(name => <SelectItem key={name} value={name}>{name}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-2">
                  <Label htmlFor="test-row-number">{t('aggregator.testRules.rowToTest')}</Label>
                  <Input 
                    id="test-row-number" 
                    type="number" 
                    min="1" 
                    value={testRowNumber} 
                    onChange={e => setTestRowNumber(e.target.value)} 
                    disabled={isTestingRule}
                  />
                </div>
                <div className="flex items-end">
                  <Button onClick={handleTestRule} disabled={isTestingRule || !testSheetName || !testRowNumber} className="w-full">
                    {isTestingRule ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <PlayCircle className="mr-2 h-4 w-4" />}
                    {t('aggregator.testRules.button')}
                  </Button>
                </div>
              </div>
              {testResult && (
                <Card className="mt-4 bg-background">
                  <CardContent className="p-3">
                    <pre className="text-xs whitespace-pre-wrap font-code">{testResult}</pre>
                  </CardContent>
                </Card>
              )}
            </CardContent>
          </Card>
        )}
        
        <Card className="p-4 bg-secondary/30">
            <CardHeader className="p-0 pb-4">
                <Label className="flex items-center space-x-2 text-md font-semibold">
                    <BarChartHorizontal className="h-5 w-5" />
                    <span>{t('aggregator.summaryConfigStep')}</span>
                </Label>
            </CardHeader>
            <CardContent className="p-0 space-y-4">
                <RadioGroup value={aggregationMode} onValueChange={(v) => setAggregationMode(v as any)} className="space-y-2">
                    <div className="flex items-start space-x-2">
                        <RadioGroupItem value="valueMatch" id="valueMatch" />
                        <div className="grid gap-1.5 w-full">
                            <Label htmlFor="valueMatch" className="font-normal">{t('aggregator.valueMatch')}</Label>
                            <p className="text-xs text-muted-foreground -mt-1">{t('aggregator.valueMatchDesc')}</p>
                            {aggregationMode === 'valueMatch' && (
                                <Card className="p-3 mt-2 space-y-3 bg-background/50">
                                    <div className="flex items-start space-x-3">
                                        <Checkbox
                                            id="enable-conditional-match"
                                            checked={enableConditionalMatch}
                                            onCheckedChange={(checked) => setEnableConditionalMatch(checked as boolean)}
                                            disabled={isProcessing}
                                            className="mt-1"
                                        />
                                        <div className="grid gap-1.5 leading-none">
                                            <Label htmlFor="enable-conditional-match">{t('aggregator.conditionalMatch.label')}</Label>
                                            <p className="text-xs text-muted-foreground">{t('aggregator.conditionalMatch.description')}</p>
                                        </div>
                                    </div>
                                    {enableConditionalMatch && (
                                        <div className="space-y-2 pl-8">
                                            <Label htmlFor="key-count-column-conditional" className="text-sm font-medium">{t('aggregator.conditionalMatch.columnLabel')}</Label>
                                            <Input 
                                                id="key-count-column-conditional"
                                                value={keyCountColumn}
                                                onChange={(e) => setKeyCountColumn(e.target.value)}
                                                placeholder={t('aggregator.keyCountColPlaceholder') as string}
                                                disabled={isProcessing}
                                            />
                                        </div>
                                    )}
                                </Card>
                            )}
                        </div>
                    </div>
                    
                    <div className="flex items-start space-x-2 pt-2">
                        <RadioGroupItem value="keyMatch" id="keyMatch" />
                         <div className="grid gap-1.5 w-full">
                            <Label htmlFor="keyMatch" className="font-normal">{t('aggregator.keyMatch')}</Label>
                            <p className="text-xs text-muted-foreground -mt-1">{t('aggregator.keyMatchDesc')}</p>
                             {aggregationMode === 'keyMatch' && (
                                <div className="space-y-4 pt-2">
                                    <div className="space-y-2">
                                        <Label htmlFor="key-count-column" className="text-sm font-medium">{t('aggregator.keyCountCol')}</Label>
                                        <Input 
                                            id="key-count-column"
                                            value={keyCountColumn}
                                            onChange={(e) => setKeyCountColumn(e.target.value)}
                                            placeholder={t('aggregator.keyCountColPlaceholder') as string}
                                            disabled={isProcessing}
                                        />
                                        <p className="text-xs text-muted-foreground">{t('aggregator.keyCountColDesc')}</p>
                                    </div>
                                     <div className="flex items-start space-x-3 pt-2">
                                        <Checkbox
                                            id="discover-new-keys"
                                            checked={discoverNewKeys}
                                            onCheckedChange={(checked) => setDiscoverNewKeys(checked as boolean)}
                                            disabled={isProcessing}
                                            className="mt-1"
                                        />
                                        <div className="grid gap-1.5 leading-none">
                                            <Label htmlFor="discover-new-keys">{t('aggregator.discoverNewKeys')}</Label>
                                            <p className="text-xs text-muted-foreground">{t('aggregator.discoverNewKeysDesc')}</p>
                                        </div>
                                    </div>
                                </div>
                            )}
                        </div>
                    </div>
                </RadioGroup>

                {showBlankLabelInput && (
                    <div className="space-y-2 pt-4 border-t">
                        <Label htmlFor="blank-label" className="text-sm font-medium">{t('aggregator.blankLabel')}</Label>
                        <Input 
                            id="blank-label"
                            value={blankLabel}
                            onChange={(e) => setBlankLabel(e.target.value)}
                            placeholder={t('aggregator.blankLabelPlaceholder') as string}
                            disabled={isProcessing || !file}
                        />
                        <p className="text-xs text-muted-foreground">{t('aggregator.blankLabelDesc')}</p>
                        
                        <div className="pt-2">
                          <Label className="text-sm font-medium">{t('aggregator.blankCountingMethod.title')}</Label>
                          <RadioGroup value={blankCountingMode} onValueChange={(v) => setBlankCountingMode(v as any)} className="mt-2 space-y-2">
                              <Label htmlFor="blank-mode-row-aware" className="flex items-start space-x-3 p-3 rounded-md border has-[:checked]:border-primary cursor-pointer">
                                  <RadioGroupItem value="rowAware" id="blank-mode-row-aware" className="mt-1" />
                                  <div className="grid gap-1.5">
                                      <span className="font-normal">{t('aggregator.blankCountingMethod.rowAware')}</span>
                                      <p className="text-xs text-muted-foreground">{t('aggregator.blankCountingMethod.rowAwareDesc')}</p>
                                  </div>
                              </Label>
                              <Label htmlFor="blank-mode-full-column" className="flex items-start space-x-3 p-3 rounded-md border has-[:checked]:border-primary cursor-pointer">
                                  <RadioGroupItem value="fullColumn" id="blank-mode-full-column" className="mt-1" />
                                  <div className="grid gap-1.5">
                                      <span className="font-normal">{t('aggregator.blankCountingMethod.fullColumn')}</span>
                                      <p className="text-xs text-muted-foreground">{t('aggregator.blankCountingMethod.fullColumnDesc')}</p>
                                  </div>
                              </Label>
                          </RadioGroup>
                        </div>

                        <div className="flex items-start space-x-3 pt-4">
                            <Checkbox
                                id="generate-blank-details"
                                checked={generateBlankDetails}
                                onCheckedChange={(checked) => setGenerateBlankDetails(checked as boolean)}
                                disabled={isProcessing || !file}
                                className="mt-1"
                            />
                            <div className="grid gap-1.5 leading-none">
                                <Label htmlFor="generate-blank-details">
                                    {t('aggregator.blankDetails')}
                                </Label>
                                <p className="text-xs text-muted-foreground">
                                    {t('aggregator.blankDetailsDesc')}
                                </p>
                            </div>
                        </div>

                    </div>
                )}
            </CardContent>
        </Card>
        
        {sheetNames.length > 0 && (
          <Card className="p-4 bg-secondary/30">
            <CardHeader className="p-0 pb-4">
              <Label className="flex items-center space-x-2 text-md font-semibold">
                <NotebookTabs className="h-5 w-5" />
                <span>{t('aggregator.reportingScope.title')}</span>
              </Label>
              <CardDescription>{t('aggregator.reportingScope.description')}</CardDescription>
            </CardHeader>
            <CardContent className="p-0 space-y-4">
              <RadioGroup value={reportingScope} onValueChange={(v) => setReportingScope(v as any)} className="space-y-2">
                <div className="flex items-center space-x-2">
                  <RadioGroupItem value="main" id="report-scope-main" />
                  <Label htmlFor="report-scope-main" className="font-normal">{t('aggregator.reportingScope.mainSelection', {count: Object.values(selectedSheets).filter(Boolean).length})}</Label>
                </div>
                <div className="flex items-center space-x-2">
                  <RadioGroupItem value="custom" id="report-scope-custom" />
                  <Label htmlFor="report-scope-custom" className="font-normal">{t('aggregator.reportingScope.customSelection')}</Label>
                </div>
              </RadioGroup>

              {reportingScope === 'custom' && (
                <div className="pl-6 pt-2 space-y-3">
                  <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-background">
                    <Checkbox
                      id="select-all-custom-report-sheets"
                      checked={allCustomSheetsSelected}
                      onCheckedChange={(checked) => handleSelectAllCustomSheets(checked as boolean)}
                      disabled={isProcessing}
                    />
                    <Label htmlFor="select-all-custom-report-sheets" className="text-sm font-medium flex-grow">
                      {t('aggregator.reportingScope.selectAll')} ({t('aggregator.reportingScope.selectedCount', {selected: Object.values(customReportSheets).filter(Boolean).length, total: sheetNames.length})})
                    </Label>
                  </div>
                  <Card className="max-h-48 overflow-y-auto p-3 bg-background">
                    <div className="space-y-2">
                    {sheetNames.map(name => (
                        <div key={`custom-${name}`} className="flex items-center space-x-2">
                        <Checkbox
                            id={`custom-sheet-${name}`}
                            checked={customReportSheets[name] || false}
                            onCheckedChange={(checked) => handleCustomSheetSelectionChange(name, checked as boolean)}
                            disabled={isProcessing}
                        />
                        <Label htmlFor={`custom-sheet-${name}`} className="text-sm font-normal">
                            {name}
                        </Label>
                        </div>
                    ))}
                    </div>
                  </Card>
                </div>
              )}
            </CardContent>
          </Card>
        )}

        <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
            <CardHeader className="p-0 pb-4">
                <Label className="flex items-center space-x-2 text-md font-semibold text-primary">
                    <Settings className="h-5 w-5" />
                    <span>{t('aggregator.optionalActionsStep')}</span>
                </Label>
            </CardHeader>
            <CardContent className="p-0">
                <div className="flex items-start space-x-3">
                    <Checkbox
                        id="strip-formulas"
                        checked={stripFormulas}
                        onCheckedChange={(checked) => setStripFormulas(checked as boolean)}
                        disabled={isProcessing || !file}
                        className="mt-1"
                    />
                    <div className="grid gap-1.5 leading-none w-full">
                        <Label htmlFor="strip-formulas" className="text-sm font-medium">
                           {t('aggregator.stripFormulasAction')}
                        </Label>
                         <p className="text-xs text-muted-foreground">{t('aggregator.stripFormulasDesc')}</p>
                    </div>
                </div>

                <div className="flex items-start space-x-3 pt-4 mt-4 border-t">
                    <Checkbox
                        id="enable-update"
                        checked={enableUpdate}
                        onCheckedChange={(checked) => setEnableUpdate(checked as boolean)}
                        disabled={isProcessing || !file}
                        className="mt-1"
                    />
                    <div className="grid gap-1.5 leading-none w-full">
                        <Label htmlFor="enable-update" className="text-sm font-medium">
                           {t('aggregator.updateColAction')}
                        </Label>
                        {enableUpdate && (
                            <div className="space-y-2 pt-2">
                                <Label htmlFor="update-column" className="text-sm font-medium">{t('aggregator.updateCol')}</Label>
                                <Input
                                    id="update-column"
                                    type="text"
                                    value={updateColumn}
                                    onChange={(e) => setUpdateColumn(e.target.value)}
                                    disabled={isProcessing || !file || !enableUpdate}
                                    placeholder={t('aggregator.updateColPlaceholder') as string}
                                />
                                 <p className="text-xs text-muted-foreground">
                                    {t('aggregator.updateColDesc')}
                                 </p>
                                <div className="flex items-start space-x-3 pt-4">
                                    <Checkbox
                                        id="update-only-blanks"
                                        checked={updateOnlyBlanks}
                                        onCheckedChange={(checked) => setUpdateOnlyBlanks(checked as boolean)}
                                        disabled={isProcessing || !file}
                                        className="mt-1"
                                    />
                                    <div className="grid gap-1.5 leading-none">
                                        <Label htmlFor="update-only-blanks" className="text-sm font-medium">{t('aggregator.updateOnlyBlanks')}</Label>
                                        <p className="text-xs text-muted-foreground"><Markup text={t('aggregator.updateOnlyBlanksDesc') as string} /></p>
                                    </div>
                                </div>
                                <Card className="p-3 mt-4 space-y-3 bg-background/50">
                                    <div className="flex items-start space-x-3">
                                        <Checkbox
                                            id="enable-paired-row-validation"
                                            checked={enablePairedRowValidation}
                                            onCheckedChange={(checked) => setEnablePairedRowValidation(checked as boolean)}
                                            disabled={isProcessing}
                                            className="mt-1"
                                        />
                                        <div className="grid gap-1.5 leading-none">
                                            <Label htmlFor="enable-paired-row-validation">{t('aggregator.pairedValidation.title')}</Label>
                                            <p className="text-xs text-muted-foreground">{t('aggregator.pairedValidation.description')}</p>
                                        </div>
                                    </div>
                                    {enablePairedRowValidation && (
                                        <div className="space-y-2 pl-8">
                                            <Label htmlFor="paired-validation-columns" className="text-sm font-medium">{t('aggregator.pairedValidation.columnsLabel')}</Label>
                                            <Input 
                                                id="paired-validation-columns"
                                                value={pairedValidationColumns}
                                                onChange={(e) => setPairedValidationColumns(e.target.value)}
                                                placeholder={t('aggregator.pairedValidation.columnsPlaceholder') as string}
                                                disabled={isProcessing}
                                            />
                                        </div>
                                    )}
                                </Card>
                                <div className="flex items-start space-x-3 pt-4">
                                    <Checkbox
                                        id="generate-update-report"
                                        checked={generateUpdateReport}
                                        onCheckedChange={(checked) => setGenerateUpdateReport(checked as boolean)}
                                        disabled={isProcessing || !file || !enableUpdate}
                                        className="mt-1"
                                    />
                                    <div className="grid gap-1.5 leading-none">
                                        <Label htmlFor="generate-update-report" className="text-sm font-medium">
                                            {t('aggregator.updateReportAction')}
                                        </Label>
                                        <p className="text-xs text-muted-foreground">
                                            <Markup text={t('aggregator.updateReportDesc') as string} />
                                        </p>
                                    </div>
                                </div>
                            </div>
                        )}
                    </div>
                </div>
                 {/* Fill Empty Key Column Action */}
                <div className="flex items-start space-x-3 pt-4 mt-4 border-t">
                    <Checkbox
                        id="enable-fill-key-column"
                        checked={enableFillKeyColumn}
                        onCheckedChange={(checked) => setEnableFillKeyColumn(checked as boolean)}
                        disabled={isProcessing || !file || aggregationMode !== 'valueMatch'}
                        className="mt-1"
                    />
                    <div className="grid gap-1.5 leading-none w-full">
                        <Label htmlFor="enable-fill-key-column" className="text-sm font-medium">{t('aggregator.fillKeyColumnAction')}</Label>
                         <p className="text-xs text-muted-foreground">{t('aggregator.fillKeyColumnDesc')}</p>
                        {enableFillKeyColumn && aggregationMode === 'valueMatch' && (
                           <div className="space-y-2 pt-2">
                                <Label htmlFor="key-count-column-fill" className="text-sm font-medium">{t('aggregator.fillKeyColumnLabel')}</Label>
                                <Input 
                                    id="key-count-column-fill"
                                    value={keyCountColumn}
                                    onChange={(e) => setKeyCountColumn(e.target.value)}
                                    placeholder={t('aggregator.keyCountColPlaceholder') as string}
                                    disabled={isProcessing}
                                />
                            </div>
                        )}
                    </div>
                </div>

                {/* Mark Matched Rows Action */}
                <div className="flex items-start space-x-3 pt-4 mt-4 border-t">
                    <Checkbox
                        id="enable-marking"
                        checked={enableMarking}
                        onCheckedChange={(checked) => setEnableMarking(checked as boolean)}
                        disabled={isProcessing || !file}
                        className="mt-1"
                    />
                    <div className="grid gap-1.5 leading-none w-full">
                        <Label htmlFor="enable-marking" className="text-sm font-medium">
                            {t('aggregator.markMatchedRowsAction')}
                        </Label>
                        <p className="text-xs text-muted-foreground">{t('aggregator.markMatchedRowsDesc')}</p>
                        {enableMarking && (
                            <div className="space-y-4 pt-2">
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                    <div className="space-y-1">
                                        <Label htmlFor="mark-column" className="text-sm font-medium">{t('aggregator.markColumnLabel')}</Label>
                                        <Input id="mark-column" type="text" value={markColumn} onChange={(e) => setMarkColumn(e.target.value)} placeholder="e.g., Status" disabled={isProcessing}/>
                                    </div>
                                    <div className="space-y-1">
                                        <Label htmlFor="mark-value" className="text-sm font-medium">{t('aggregator.markValueLabel')}</Label>
                                        <Input id="mark-value" type="text" value={markValue} onChange={(e) => setMarkValue(e.target.value)} placeholder="e.g., Processed" disabled={isProcessing}/>
                                        <p className="text-xs text-muted-foreground">{t('aggregator.markValueDesc')}</p>
                                    </div>
                                </div>
                            </div>
                        )}
                    </div>
                </div>

                {/* Insert Aggregates Action */}
                <div className="flex items-start space-x-3 pt-4 mt-4 border-t">
                     <Checkbox
                        id="insert-results"
                        checked={insertResultsInSheet}
                        onCheckedChange={(checked) => setInsertResultsInSheet(checked as boolean)}
                        disabled={isProcessing || !file}
                        className="mt-1"
                      />
                    <div className="grid gap-1.5 leading-none w-full">
                        <Label htmlFor="insert-results" className="text-sm font-medium">
                            {t('aggregator.insertSummaryAction')}
                        </Label>
                        {insertResultsInSheet && (
                          <div className="space-y-4 pt-2">
                            <p className="text-xs text-muted-foreground pb-2">
                                {t('aggregator.insertSummaryDesc')}
                            </p>
                            <Card className="p-3 bg-background/50 space-y-4">
                               <h4 className="text-sm font-semibold">{t('aggregator.inSheetSummaryConfig')}</h4>
                               
                               <div className="flex items-start space-x-3 pt-4 border-t">
                                    <Checkbox
                                        id="clear-existing-summary"
                                        checked={clearExistingInSheetSummary}
                                        onCheckedChange={(c) => setClearExistingInSheetSummary(c as boolean)}
                                        disabled={isProcessing}
                                        className="mt-1"
                                    />
                                    <div className="grid gap-1.5 leading-none">
                                        <Label htmlFor="clear-existing-summary" className="font-normal">{t('aggregator.clearExistingSummary')}</Label>
                                        <p className="text-xs text-muted-foreground">{t('aggregator.clearExistingSummaryDesc')}</p>
                                    </div>
                                </div>

                               <div>
                                    <Label className="text-sm font-medium">{t('aggregator.summaryDataSource.title')}</Label>
                                    <RadioGroup value={inSheetDataSource} onValueChange={(v) => setInSheetDataSource(v as any)} className="mt-2 space-y-2">
                                      <div className="flex items-start space-x-2">
                                          <RadioGroupItem value="reportingScope" id="source-reporting-scope"/>
                                          <div className="grid gap-0.5">
                                            <Label htmlFor="source-reporting-scope" className="font-normal">{t('aggregator.summaryDataSource.useReportingScope')}</Label>
                                            <p className="text-xs text-muted-foreground">{t('aggregator.summaryDataSource.useReportingScopeDesc')}</p>
                                          </div>
                                      </div>
                                      <div className="flex items-start space-x-2">
                                          <RadioGroupItem value="localSheet" id="source-local-sheet"/>
                                          <div className="grid gap-1.5 w-full">
                                            <Label htmlFor="source-local-sheet" className="font-normal">{t('aggregator.summaryDataSource.useLocalSheet')}</Label>
                                            <p className="text-xs text-muted-foreground">{t('aggregator.summaryDataSource.useLocalSheetDesc')}</p>
                                            {inSheetDataSource === 'localSheet' && (
                                                <div className="flex items-start space-x-3 pt-4 pl-2">
                                                    <Checkbox
                                                        id="show-only-local-keys"
                                                        checked={showOnlyLocalKeys}
                                                        onCheckedChange={(c) => setShowOnlyLocalKeys(c as boolean)}
                                                        className="mt-1"
                                                    />
                                                    <div className="grid gap-1.5 leading-none">
                                                        <Label htmlFor="show-only-local-keys" className="font-normal">{t('aggregator.summaryDataSource.showOnlyLocalKeys')}</Label>
                                                        <p className="text-xs text-muted-foreground">{t('aggregator.summaryDataSource.showOnlyLocalKeysDesc')}</p>
                                                    </div>
                                                </div>
                                            )}
                                          </div>
                                      </div>
                                    </RadioGroup>
                               </div>
                               <div>
                                    <Label className="text-sm font-medium">{t('aggregator.valueGeneration.title')}</Label>
                                     <RadioGroup value={inSheetGenerationMode} onValueChange={(v) => setInSheetGenerationMode(v as any)} className="mt-2 space-y-2">
                                      <div className="flex items-start space-x-2">
                                          <RadioGroupItem value="static" id="gen-static"/>
                                           <div className="grid gap-0.5">
                                             <Label htmlFor="gen-static" className="font-normal">{t('aggregator.valueGeneration.staticValues')}</Label>
                                             <p className="text-xs text-muted-foreground">{t('aggregator.valueGeneration.staticValuesDesc')}</p>
                                          </div>
                                      </div>
                                        <TooltipProvider>
                                          <Tooltip>
                                            <TooltipTrigger asChild>
                                                <div className="flex items-start space-x-2">
                                                    <RadioGroupItem value="formula" id="gen-formula" disabled={aggregationMode !== 'keyMatch'}/>
                                                     <div className="grid gap-0.5">
                                                      <Label htmlFor="gen-formula" className="font-normal disabled:opacity-50">{t('aggregator.valueGeneration.liveFormulas')}</Label>
                                                      <p className="text-xs text-muted-foreground">{t('aggregator.valueGeneration.liveFormulasDesc')}</p>
                                                    </div>
                                                </div>
                                            </TooltipTrigger>
                                            {aggregationMode !== 'keyMatch' && (
                                                <TooltipContent side="right" align="center">
                                                    <p>{t('aggregator.valueGeneration.liveFormulasTooltip')}</p>
                                                </TooltipContent>
                                            )}
                                          </Tooltip>
                                        </TooltipProvider>
                                    </RadioGroup>
                               </div>
                            </Card>
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                <div className="space-y-1">
                                    <Label htmlFor="in-sheet-summary-title" className="text-sm font-medium">{t('aggregator.inSheetSummaryTitle')}</Label>
                                    <Input id="in-sheet-summary-title" type="text" value={inSheetSummaryTitle} onChange={(e) => setInSheetSummaryTitle(e.target.value)} placeholder="Summary" disabled={isProcessing}/>
                                </div>
                                 <div className="space-y-1">
                                    <Label htmlFor="summary-title-cell" className="text-sm font-medium">{t('aggregator.dynamicTitleCell')}</Label>
                                    <Input id="summary-title-cell" type="text" value={summaryTitleCell} onChange={(e) => setSummaryTitleCell(e.target.value)} placeholder="e.g., A1" disabled={isProcessing}/>
                                </div>
                            </div>
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                <div className="space-y-1">
                                    <Label htmlFor="insert-column" className="text-sm font-medium">{t('aggregator.insertAtCol')}</Label>
                                    <Input id="insert-column" type="text" value={insertColumn} onChange={(e) => setInsertColumn(e.target.value)} placeholder={t('aggregator.insertAtColPlaceholder') as string} disabled={isProcessing}/>
                                </div>
                                <div className="space-y-1">
                                    <Label htmlFor="insert-start-row" className="text-sm font-medium">{t('aggregator.insertAtRow')}</Label>
                                    <Input id="insert-start-row" type="number" min="1" value={insertStartRow} onChange={(e) => setInsertStartRow(parseInt(e.target.value, 10) || 1)} placeholder={t('aggregator.insertAtRowPlaceholder') as string} disabled={isProcessing}/>
                                </div>
                            </div>
                             <Accordion type="single" collapsible className="w-full">
                                <AccordionItem value="appearance-settings">
                                    <AccordionTrigger className="text-sm font-semibold flex items-center gap-2"><TableIcon className="h-4 w-4"/>{[t('aggregator.tableAppearance.title')].flat().join(' ')}</AccordionTrigger>
                                    <AccordionContent className="space-y-4 pt-2">
                                        <div className="space-y-2">
                                            <Label htmlFor="table-fill-color" className="text-sm font-medium">{t('aggregator.tableAppearance.fillColor')}</Label>
                                            <Input id="table-fill-color" type="text" value={tableFormatting.fillColor || ''} onChange={(e) => setTableFormatting(p => ({...p, fillColor: e.target.value.replace('#', '')}))} disabled={isProcessing} placeholder="e.g., F8F9FA"/>
                                        </div>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                            <div className="space-y-2">
                                                <Label htmlFor="table-border-style" className="text-sm font-medium">{t('aggregator.tableAppearance.borderStyle')}</Label>
                                                <Select value={tableFormatting.borderStyle || 'thin'} onValueChange={(v) => setTableFormatting(p => ({...p, borderStyle: v as BorderStyle}))}>
                                                    <SelectTrigger id="table-border-style"><SelectValue /></SelectTrigger>
                                                    <SelectContent>
                                                        <SelectItem value="thin">{t('common.borderStyles.thin')}</SelectItem>
                                                        <SelectItem value="medium">{t('common.borderStyles.medium')}</SelectItem>
                                                        <SelectItem value="thick">{t('common.borderStyles.thick')}</SelectItem>
                                                        <SelectItem value="double">{t('common.borderStyles.double')}</SelectItem>
                                                        <SelectItem value="dotted">{t('common.borderStyles.dotted')}</SelectItem>
                                                        <SelectItem value="dashed">{t('common.borderStyles.dashed')}</SelectItem>
                                                    </SelectContent>
                                                </Select>
                                            </div>
                                             <div className="space-y-2">
                                                <Label htmlFor="table-border-color" className="text-sm font-medium">{t('aggregator.tableAppearance.borderColor')}</Label>
                                                <Input id="table-border-color" type="text" value={tableFormatting.borderColor || ''} onChange={(e) => setTableFormatting(p => ({...p, borderColor: e.target.value.replace('#', '')}))} disabled={isProcessing} placeholder="e.g., DEE2E6"/>
                                            </div>
                                        </div>
                                    </AccordionContent>
                                </AccordionItem>
                             </Accordion>
                            <Card className="p-3 bg-background/50 mt-4">
                                <h4 className="text-sm font-semibold mb-3">{t('aggregator.summaryHeaderFormatting')}</h4>
                                <div className="space-y-4">
                                    <div className="grid grid-cols-2 gap-4">
                                        <div>
                                            <Label htmlFor="summary-header-font-name" className="text-sm font-medium">{t('common.fontName')}</Label>
                                            <Input id="summary-header-font-name" type="text" value={summaryHeaderFormatting.fontName || ''} onChange={(e) => setSummaryHeaderFormatting(p => ({...p, fontName: e.target.value}))} disabled={isProcessing} placeholder="e.g., Calibri"/>
                                        </div>
                                         <div>
                                            <Label htmlFor="summary-header-font-size" className="text-sm font-medium">{t('common.fontSize')}</Label>
                                            <Input id="summary-header-font-size" type="number" min="1" value={summaryHeaderFormatting.fontSize || 12} onChange={(e) => setSummaryHeaderFormatting(p => ({...p, fontSize: parseInt(e.target.value, 10) || 11}))} disabled={isProcessing} />
                                        </div>
                                    </div>
                                    <div className="grid grid-cols-2 gap-4">
                                        <div>
                                           <Label htmlFor="summary-header-h-align" className="text-sm font-medium">{t('common.alignment')}</Label>
                                            <Select value={summaryHeaderFormatting.horizontalAlignment || 'general'} onValueChange={(v) => setSummaryHeaderFormatting(p => ({...p, horizontalAlignment: v as HorizontalAlignment}))} disabled={isProcessing}>
                                                <SelectTrigger id="summary-header-h-align"><SelectValue /></SelectTrigger>
                                                <SelectContent>
                                                    <SelectItem value="general">{t('common.alignments.general')}</SelectItem>
                                                    <SelectItem value="left">{t('common.alignments.left')}</SelectItem>
                                                    <SelectItem value="center">{t('common.alignments.center')}</SelectItem>
                                                    <SelectItem value="right">{t('common.alignments.right')}</SelectItem>
                                                    <SelectItem value="fill">{t('common.alignments.fill')}</SelectItem>
                                                    <SelectItem value="justify">{t('common.alignments.justify')}</SelectItem>
                                                </SelectContent>
                                            </Select>
                                        </div>
                                        <div>
                                          <Label htmlFor="summary-header-fill-color" className="text-sm font-medium">{t('updater.fillColorHex')}</Label>
                                          <Input
                                            id="summary-header-fill-color"
                                            type="text"
                                            value={summaryHeaderFormatting.fillColor || ''}
                                            onChange={(e) => {
                                                const colorValue = e.target.value.replace('#', '');
                                                setSummaryHeaderFormatting(p => ({...p, fillColor: colorValue}));
                                            }}
                                            disabled={isProcessing}
                                            placeholder="e.g., EAEAEA"
                                            />
                                        </div>
                                    </div>
                                    <div className="flex items-center space-x-4 pt-2">
                                       <div className="flex items-center space-x-2"><Checkbox id="summary-format-bold" checked={!!summaryHeaderFormatting.bold} onCheckedChange={(checked) => setSummaryHeaderFormatting(p => ({ ...p, bold: checked as boolean }))} disabled={isProcessing} /><Label htmlFor="summary-format-bold">{t('common.bold')}</Label></div>
                                       <div className="flex items-center space-x-2"><Checkbox id="summary-format-italic" checked={!!summaryHeaderFormatting.italic} onCheckedChange={(checked) => setSummaryHeaderFormatting(p => ({ ...p, italic: checked as boolean }))} disabled={isProcessing} /><Label htmlFor="summary-format-italic">{t('common.italic')}</Label></div>
                                       <div className="flex items-center space-x-2"><Checkbox id="summary-format-underline" checked={!!summaryHeaderFormatting.underline} onCheckedChange={(checked) => setSummaryHeaderFormatting(p => ({ ...p, underline: checked as boolean }))} disabled={isProcessing} /><Label htmlFor="summary-format-underline">{t('common.underline')}</Label></div>
                                    </div>
                                </div>
                            </Card>
                            <Card className="p-3 bg-background/50 mt-4">
                                <h4 className="text-sm font-semibold mb-2">{t('aggregator.specialRowFormatting')}</h4>
                                <p className="text-xs text-muted-foreground mb-4">{t('aggregator.specialRowFormattingDesc')}</p>
                                <div className="space-y-4">
                                    <div className="p-3 border rounded-md">
                                        <div className="flex items-center space-x-2 mb-2">
                                            <Checkbox id="enable-total-row-formatting" checked={enableTotalRowFormatting} onCheckedChange={(c) => setEnableTotalRowFormatting(c as boolean)} />
                                            <Label htmlFor="enable-total-row-formatting" className="font-medium">{t('aggregator.totalRow')}</Label>
                                        </div>
                                        {enableTotalRowFormatting && (
                                            <div className="grid grid-cols-2 gap-4 pl-6">
                                                <div className="flex items-center space-x-2 self-center">
                                                    <Checkbox id="total-row-bold" checked={!!totalRowFormatting.bold} onCheckedChange={(c) => setTotalRowFormatting(p => ({...p, bold: c as boolean}))} />
                                                    <Label htmlFor="total-row-bold" className="font-normal">{t('common.bold')}</Label>
                                                </div>
                                                <div>
                                                    <Label htmlFor="total-row-fill" className="text-xs block">{t('updater.fillColorHex')}</Label>
                                                    <Input id="total-row-fill" value={totalRowFormatting.fillColor || ''} onChange={e => setTotalRowFormatting(p => ({...p, fillColor: e.target.value.replace('#', '')}))} className="mt-1 h-8" />
                                                </div>
                                            </div>
                                        )}
                                    </div>
                                    <div className="p-3 border rounded-md">
                                        <div className="flex items-center space-x-2 mb-2">
                                            <Checkbox id="enable-blank-row-formatting" checked={enableBlankRowFormatting} onCheckedChange={(c) => setEnableBlankRowFormatting(c as boolean)} />
                                            <Label htmlFor="enable-blank-row-formatting" className="font-medium">{t('aggregator.blankRow')}</Label>
                                        </div>
                                        {enableBlankRowFormatting && (
                                            <div className="space-y-4 pl-6">
                                              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                                                  <div className="flex items-center space-x-2 self-center">
                                                      <Checkbox id="blank-row-bold" checked={!!blankRowFormatting.bold} onCheckedChange={(c) => setBlankRowFormatting(p => ({...p, bold: c as boolean}))} />
                                                      <Label htmlFor="blank-row-bold" className="font-normal">{t('common.bold')}</Label>
                                                  </div>
                                                  <div>
                                                      <Label htmlFor="blank-row-fill" className="text-xs block">{t('updater.fillColorHex')}</Label>
                                                      <Input id="blank-row-fill" value={blankRowFormatting.fillColor || ''} onChange={e => setBlankRowFormatting(p => ({...p, fillColor: e.target.value.replace('#', '')}))} className="mt-1 h-8" />
                                                  </div>
                                                  <div>
                                                      <Label htmlFor="blank-row-font-color" className="text-xs block">{t('updater.fontColorHex')}</Label>
                                                      <Input id="blank-row-font-color" value={blankRowFormatting.fontColor || ''} onChange={e => setBlankRowFormatting(p => ({...p, fontColor: e.target.value.replace('#', '')}))} className="mt-1 h-8" />
                                                  </div>
                                              </div>
                                              <div className="flex items-start space-x-3 pt-2">
                                                  <Checkbox
                                                      id="show-blanks-in-insheet"
                                                      checked={showBlanksInInSheetSummary}
                                                      onCheckedChange={(c) => setShowBlanksInInSheetSummary(c as boolean)}
                                                  />
                                                  <div className="grid gap-1.5 leading-none">
                                                      <Label htmlFor="show-blanks-in-insheet" className="font-normal">{t('aggregator.showBlanksInInSheet')}</Label>
                                                      <p className="text-xs text-muted-foreground">{t('aggregator.showBlanksInInSheetDesc')}</p>
                                                  </div>
                                              </div>
                                            </div>
                                        )}
                                    </div>
                                </div>
                            </Card>
                          </div>
                        )}
                    </div>
                </div>
            </CardContent>
        </Card>
        
        <Accordion type="single" collapsible>
            <AccordionItem value="advanced-settings">
                <AccordionTrigger className="text-md font-semibold">{t('aggregator.advancedSettings.title')}</AccordionTrigger>
                <AccordionContent>
                    <Card className="p-4 border-dashed space-y-4">
                       <div className="space-y-2">
                            <Label htmlFor="report-chunk-size" className="text-sm font-medium">{t('aggregator.advancedSettings.maxRows')}</Label>
                            <Input
                                id="report-chunk-size"
                                type="number"
                                min="1000"
                                step="1000"
                                value={reportChunkSize}
                                onChange={(e) => setReportChunkSize(parseInt(e.target.value, 10) || 100000)}
                                disabled={isProcessing}
                            />
                            <p className="text-xs text-muted-foreground">
                                {t('aggregator.advancedSettings.maxRowsDesc')}
                            </p>
                        </div>
                        <div className="space-y-2 pt-4 border-t">
                            <Label className="text-sm font-medium">{t('aggregator.advancedSettings.reportLayout')}</Label>
                            <RadioGroup value={reportLayout} onValueChange={(v) => setReportLayout(v as any)} className="space-y-2">
                                <div className="flex items-center space-x-2">
                                    <RadioGroupItem value="sheetsAsRows" id="layout-sheets-rows" />
                                    <Label htmlFor="layout-sheets-rows" className="font-normal">{t('aggregator.advancedSettings.sheetsAsRows')}</Label>
                                </div>
                                <p className="text-xs text-muted-foreground pl-6 -mt-1">{t('aggregator.advancedSettings.sheetsAsRowsDesc')}</p>
                                <div className="flex items-center space-x-2">
                                    <RadioGroupItem value="keysAsRows" id="layout-keys-rows" />
                                    <Label htmlFor="layout-keys-rows" className="font-normal">{t('aggregator.advancedSettings.keysAsRows')}</Label>
                                </div>
                                <p className="text-xs text-muted-foreground pl-6 -mt-1">{t('aggregator.advancedSettings.keysAsRowsDesc')}</p>
                            </RadioGroup>
                        </div>
                        <div className="space-y-2 pt-4 border-t">
                            <div className="flex items-center space-x-2">
                                <Checkbox
                                    id="auto-size-columns"
                                    checked={autoSizeColumns}
                                    onCheckedChange={(c) => setAutoSizeColumns(c as boolean)}
                                    disabled={isProcessing}
                                />
                                <Label htmlFor="auto-size-columns" className="font-normal">{t('aggregator.advancedSettings.autoSizeColumns')}</Label>
                            </div>
                            <p className="text-xs text-muted-foreground pl-6 -mt-1">{t('aggregator.advancedSettings.autoSizeColumnsDesc')}</p>
                        </div>
                        <div className="space-y-2 pt-4 border-t">
                            <Label htmlFor="columns-to-hide" className="text-sm font-medium">{t('aggregator.advancedSettings.hideColumns')}</Label>
                            <Input
                                id="columns-to-hide"
                                value={columnsToHide}
                                onChange={(e) => setColumnsToHide(e.target.value)}
                                disabled={isProcessing}
                                placeholder={t('aggregator.advancedSettings.hideColumnsPlaceholder') as string}
                            />
                            <p className="text-xs text-muted-foreground">{t('aggregator.advancedSettings.hideColumnsDesc')}</p>
                        </div>

                         <Accordion type="single" collapsible className="w-full pt-4 border-t">
                            <AccordionItem value="group-report-formatting">
                                <AccordionTrigger className="text-sm font-semibold flex items-center gap-2">
                                    <Palette className="h-4 w-4"/>
                                    <span>{t('aggregator.advancedSettings.groupReportFormattingTitle')}</span>
                                </AccordionTrigger>
                                <AccordionContent className="space-y-4 pt-2">
                                    <div className="space-y-2">
                                        <Label htmlFor="group-report-sheet-title" className="text-sm font-medium">{t('aggregator.advancedSettings.groupReportSheetTitle')}</Label>
                                        <Input
                                            id="group-report-sheet-title"
                                            value={groupReportSheetTitle}
                                            onChange={(e) => setGroupReportSheetTitle(e.target.value)}
                                            disabled={isProcessing}
                                            placeholder={t('aggregator.advancedSettings.groupReportSheetTitle') as string}
                                        />
                                        <p className="text-xs text-muted-foreground">{t('aggregator.advancedSettings.groupReportSheetTitleDesc')}</p>
                                    </div>
                                    <div className="space-y-2 pt-4 border-t">
                                        <Label htmlFor="group-report-multi-source-title" className="text-sm font-medium">{t('aggregator.advancedSettings.groupReportMultiSourceTitle')}</Label>
                                        <Input
                                            id="group-report-multi-source-title"
                                            value={groupReportMultiSourceTitle}
                                            onChange={(e) => setGroupReportMultiSourceTitle(e.target.value)}
                                            disabled={isProcessing}
                                            placeholder={t('aggregator.advancedSettings.groupReportMultiSourceTitlePlaceholder') as string}
                                        />
                                        <p className="text-xs text-muted-foreground">{t('aggregator.advancedSettings.groupReportMultiSourceTitleDesc')}</p>
                                    </div>
                                     <div className="space-y-2 pt-4 border-t">
                                        <Label htmlFor="group-report-description" className="text-sm font-medium">{t('aggregator.advancedSettings.groupReportDescription')}</Label>
                                        <Textarea
                                            id="group-report-description"
                                            value={groupReportDescription}
                                            onChange={(e) => setGroupReportDescription(e.target.value)}
                                            disabled={isProcessing}
                                            placeholder={t('aggregator.advancedSettings.groupReportDescriptionPlaceholder') as string}
                                            rows={2}
                                        />
                                    </div>
                                    <div className="space-y-2 pt-4 border-t">
                                      <Label className="text-sm font-medium">{t('aggregator.advancedSettings.groupReportHeadersTitle')}</Label>
                                      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                                          <div>
                                              <Label htmlFor="gr-header-group-name" className="text-xs">{t('aggregator.advancedSettings.groupReportGroupName')}</Label>
                                              <Input id="gr-header-group-name" value={groupReportHeaders.groupName} onChange={(e) => handleGroupHeaderChange('groupName', e.target.value)} disabled={isProcessing} className="h-8 text-xs mt-1"/>
                                          </div>
                                          <div>
                                              <Label htmlFor="gr-header-key-name" className="text-xs">{t('aggregator.advancedSettings.groupReportKeyName')}</Label>
                                              <Input id="gr-header-key-name" value={groupReportHeaders.keyName} onChange={(e) => handleGroupHeaderChange('keyName', e.target.value)} disabled={isProcessing} className="h-8 text-xs mt-1"/>
                                          </div>
                                          <div>
                                              <Label htmlFor="gr-header-count" className="text-xs">{t('aggregator.advancedSettings.groupReportCount')}</Label>
                                              <Input id="gr-header-count" value={groupReportHeaders.count} onChange={(e) => handleGroupHeaderChange('count', e.target.value)} disabled={isProcessing} className="h-8 text-xs mt-1"/>
                                          </div>
                                      </div>
                                    </div>
                                    <div className="grid grid-cols-2 gap-4">
                                      <div>
                                        <Label htmlFor="gr-header-font-name" className="text-sm">{t('common.fontName')}</Label>
                                        <Input id="gr-header-font-name" value={groupReportHeaderFormatting.fontName || ''} onChange={(e) => setGroupReportHeaderFormatting(p => ({...p, fontName: e.target.value}))} disabled={isProcessing} placeholder="e.g., Calibri"/>
                                      </div>
                                      <div>
                                        <Label htmlFor="gr-header-font-size" className="text-sm">{t('common.fontSize')}</Label>
                                        <Input id="gr-header-font-size" type="number" min="1" value={groupReportHeaderFormatting.fontSize || 12} onChange={(e) => setGroupReportHeaderFormatting(p => ({...p, fontSize: parseInt(e.target.value, 10) || 12}))} disabled={isProcessing} />
                                      </div>
                                    </div>
                                    <div>
                                      <Label htmlFor="gr-header-fill-color" className="text-sm">{t('updater.fillColorHex')}</Label>
                                      <Input id="gr-header-fill-color" value={groupReportHeaderFormatting.fillColor || ''} onChange={e => setGroupReportHeaderFormatting(p => ({...p, fillColor: e.target.value.replace('#', '')}))} disabled={isProcessing} placeholder="e.g., D9EAD3" />
                                    </div>
                                    <div className="flex items-center space-x-4 pt-2">
                                       <div className="flex items-center space-x-2"><Checkbox id="gr-format-bold" checked={!!groupReportHeaderFormatting.bold} onCheckedChange={(checked) => setGroupReportHeaderFormatting(p => ({ ...p, bold: checked as boolean }))} disabled={isProcessing} /><Label htmlFor="gr-format-bold" className="font-normal">{t('common.bold')}</Label></div>
                                       <div className="flex items-center space-x-2"><Checkbox id="gr-format-italic" checked={!!groupReportHeaderFormatting.italic} onCheckedChange={(checked) => setGroupReportHeaderFormatting(p => ({ ...p, italic: checked as boolean }))} disabled={isProcessing} /><Label htmlFor="gr-format-italic" className="font-normal">{t('common.italic')}</Label></div>
                                       <div className="flex items-center space-x-2"><Checkbox id="gr-format-underline" checked={!!groupReportHeaderFormatting.underline} onCheckedChange={(checked) => setGroupReportHeaderFormatting(p => ({ ...p, underline: checked as boolean }))} disabled={isProcessing} /><Label htmlFor="gr-format-underline" className="font-normal">{t('common.underline')}</Label></div>
                                    </div>
                                </AccordionContent>
                            </AccordionItem>
                        </Accordion>
                    </Card>
                </AccordionContent>
            </AccordionItem>
        </Accordion>

        <Button
          onClick={handleProcess}
          disabled={isProcessing || !file || searchTerms.trim() === '' || searchColumns.trim() === ''}
          className="w-full bg-primary hover:bg-primary/90 text-primary-foreground"
        >
          {isProcessing && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
          <Sigma className="mr-2 h-5 w-5" />
          {t('aggregator.processBtn')}
        </Button>
      </CardContent>
      {hasResults && modifiedDataForDisplay && (
        <CardFooter className="flex-col space-y-4 items-stretch">
            <Card className="p-4 bg-secondary/30">
                <CardHeader className="p-2 flex justify-between items-center">
                    <CardTitle className="text-lg font-headline">{t('aggregator.results.finalResultsTitle')}</CardTitle>
                    <TooltipProvider>
                      <Tooltip>
                        <TooltipTrigger asChild>
                          <Button
                            onClick={handleDownloadPdf}
                            disabled={isProcessing || !hasResults}
                            variant="outline"
                            size="icon"
                          >
                            <FileText className="h-4 w-4" />
                            <span className="sr-only">{t('aggregator.downloadPdfBtn')}</span>
                          </Button>
                        </TooltipTrigger>
                        <TooltipContent>
                          <p>{t('aggregator.downloadPdfBtn')}</p>
                        </TooltipContent>
                      </Tooltip>
                    </TooltipProvider>
                </CardHeader>
                <CardContent className="p-2 space-y-4 max-h-96 overflow-y-auto">
                    <Table>
                        <TableHeader>
                            <TableRow>
                                <TableHead>{t('aggregator.results.keyHeader')}</TableHead>
                                <TableHead className="text-right">{t('aggregator.results.countHeader')}</TableHead>
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                             {finalSortedKeysForDisplay.map(([key, count]) => (
                                <TableRow key={key}>
                                    <TableCell className="font-medium">{key}</TableCell>
                                    <TableCell className="text-right">{count.toLocaleString()}</TableCell>
                                </TableRow>
                            ))}
                        </TableBody>
                        <UiTableFooter>
                            <TableRow>
                                <TableCell className="font-bold text-primary">{t('aggregator.grandTotals')}</TableCell>
                                <TableCell className="text-right font-bold text-primary">{grandTotalCount.toLocaleString()}</TableCell>
                            </TableRow>
                        </UiTableFooter>
                    </Table>
                </CardContent>
            </Card>

            <Card className="p-4 bg-secondary/30">
                <CardHeader className="p-2 flex justify-between items-center">
                    <div className="space-y-1">
                        <CardTitle className="text-lg font-headline flex items-center">
                            <PencilLine className="mr-3 h-5 w-5"/>
                            {t('aggregator.editKeysTitle')}
                        </CardTitle>
                        <CardDescription>{t('aggregator.editKeysDesc')}</CardDescription>
                    </div>
                </CardHeader>
                <CardContent className="p-2">
                    <ScrollArea className="h-72 border rounded-md">
                    <Table>
                        <TableHeader>
                        <TableRow>
                            <TableHead>{t('aggregator.originalKey')}</TableHead>
                            <TableHead>{t('aggregator.newKey')}</TableHead>
                        </TableRow>
                        </TableHeader>
                        <TableBody>
                        {Array.from(editableKeys.keys())
                            .sort((a,b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }))
                            .map(originalKey => (
                            <TableRow key={originalKey}>
                            <TableCell className="font-code">{originalKey}</TableCell>
                            <TableCell>
                                <Input
                                value={editableKeys.get(originalKey) || ''}
                                onChange={(e) => handleKeyEdit(originalKey, e.target.value)}
                                className="font-code"
                                />
                            </TableCell>
                            </TableRow>
                        ))}
                        </TableBody>
                    </Table>
                    </ScrollArea>
                </CardContent>
            </Card>
            
            <Card className="p-4 bg-secondary/30">
                <CardHeader className="p-2">
                    <CardTitle className="text-lg font-headline flex items-center">
                        <BarChart2 className="mr-3 h-5 w-5"/>
                        {[t('aggregator.createGroupReport')].flat().join(' ')}
                    </CardTitle>
                    <CardDescription>
                        {[t('aggregator.createGroupReportDesc')].flat().join(' ')}
                    </CardDescription>
                </CardHeader>
                <CardContent className="p-2 space-y-2">
                    <Label htmlFor="group-mappings">{[t('aggregator.groupMappings')].flat().join(' ')}</Label>
                    <Textarea
                        id="group-mappings"
                        value={groupMappings}
                        onChange={(e) => setGroupMappings(e.target.value)}
                        placeholder={[t('aggregator.groupMappingsPlaceholder')].flat().join(' ')}
                        rows={4}
                    />
                    <p className="text-xs text-muted-foreground">
                        <Markup text={[t('aggregator.groupMappingsDesc')].flat().join(' ')} />
                    </p>
                </CardContent>
                <CardFooter className="p-2 pt-4 flex-col items-stretch space-y-4">
                     <Button
                        onClick={() => handleDownloadGroupReport('initial')}
                        disabled={isProcessing || !hasResults || !groupMappings.trim()}
                        className="w-full"
                        variant="outline"
                    >
                        <Download className="mr-2 h-5 w-5" />
                        {[t('aggregator.downloadGroupReportBtn')].flat().join(' ')}
                    </Button>
                    <div className="pt-2 border-t">
                        <Button
                            onClick={() => handleDownloadGroupReport('final')}
                            disabled={isProcessing || !finalAggregationResult || !groupMappings.trim()}
                            className="w-full mt-2"
                            variant="secondary"
                        >
                            <Download className="mr-2 h-5 w-5" />
                            {[t('aggregator.downloadFinalGroupReportBtn')].flat().join(' ')}
                        </Button>
                        <p className="text-xs text-muted-foreground mt-2 text-center">
                            {[t('aggregator.finalGroupReportDesc')].flat().join(' ')}
                        </p>
                    </div>
                </CardFooter>
            </Card>

            <Card className="w-full p-4 bg-secondary/30 space-y-4">
                <Label className="text-md font-semibold font-headline">{t('common.outputOptions.title')}</Label>
                <RadioGroup value={outputFormat} onValueChange={(v) => setOutputFormat(v as any)} className="space-y-3">
                    <div>
                        <div className="flex items-center space-x-2">
                            <RadioGroupItem value="xlsx" id="format-xlsx-agg" />
                            <Label htmlFor="format-xlsx-agg" className="font-normal">{t('common.outputOptions.xlsx')}</Label>
                        </div>
                        <p className="text-xs text-muted-foreground pl-6 pt-1">{t('common.outputOptions.xlsxDesc')}</p>
                    </div>
                    <div>
                        <div className="flex items-center space-x-2">
                            <RadioGroupItem value="xlsm" id="format-xlsm-agg" />
                            <Label htmlFor="format-xlsm-agg" className="font-normal">{t('common.outputOptions.xlsm')}</Label>
                        </div>
                        <p className="text-xs text-muted-foreground pl-6 pt-1">{t('common.outputOptions.xlsmDesc')}</p>
                    </div>
                </RadioGroup>
                
                <Alert variant="default" className="mt-2">
                    <Lightbulb className="h-4 w-4" />
                    <AlertDescription>
                        {t('common.outputOptions.recommendation')}
                    </AlertDescription>
                </Alert>
            </Card>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Button
                    onClick={handleDownloadReport}
                    disabled={isProcessing || !hasResults}
                    className="w-full sm:col-span-2"
                    variant="secondary"
                >
                    <FileOutput className="mr-2 h-5 w-5" />
                    {t('aggregator.downloadReportBtn')}
                </Button>
            </div>
            <Button
                onClick={handleDownloadFinalWorkbook}
                disabled={isProcessing || !hasResults || !canDownloadModifiedWorkbook}
                className="w-full bg-accent hover:bg-accent/90 text-accent-foreground"
            >
                <Download className="mr-2 h-5 w-5" />
                {t('aggregator.downloadBtn')}
            </Button>
        </CardFooter>
      )}
      {lastUpdateResult && (
        <CardFooter>
            <Card className="p-4 bg-secondary/30 w-full">
            <CardHeader className="p-2">
                <CardTitle className="text-lg font-headline flex items-center">
                <CheckCircle2 className="mr-3 h-5 w-5 text-green-500" />
                {t('aggregator.verification.title')}
                </CardTitle>
                <CardDescription>
                {t('aggregator.verification.description', { count: lastUpdateResult.summary.totalCellsUpdated, sheets: lastUpdateResult.summary.sheetsUpdated.length })}
                </CardDescription>
            </CardHeader>
            {lastUpdateResult.details.length > 0 && (
                <CardContent className="p-2">
                <ScrollArea className="h-72 border rounded-md">
                    <Table>
                    <TableHeader>
                        <TableRow>
                        <TableHead>{t('aggregator.verification.sheet')}</TableHead>
                        <TableHead>{t('aggregator.verification.cell')}</TableHead>
                        <TableHead>{t('aggregator.verification.original')}</TableHead>
                        <TableHead>{t('aggregator.verification.new')}</TableHead>
                        <TableHead>{t('aggregator.verification.trigger')}</TableHead>
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {lastUpdateResult.details.slice(0, 100).map((update, index) => (
                        <TableRow key={index}>
                            <TableCell>{update.sheetName}</TableCell>
                            <TableCell className="font-code text-xs">{update.cellAddress}</TableCell>
                            <TableCell>{String(update.originalValue ?? '(blank)')}</TableCell>
                            <TableCell className="text-primary font-medium">{String(update.newValue)}</TableCell>
                            <TableCell className="text-muted-foreground text-xs">{`${update.triggerColumn}: "${update.triggerValue}"`}</TableCell>
                        </TableRow>
                        ))}
                    </TableBody>
                    </Table>
                </ScrollArea>
                {lastUpdateResult.details.length > 100 && (
                    <p className="text-xs text-muted-foreground text-center mt-2">
                    {t('aggregator.verification.previewNote', { count: 100 })}
                    </p>
                )}
                </CardContent>
            )}
            </Card>
        </CardFooter>
      )}
    </Card>
  );
}
