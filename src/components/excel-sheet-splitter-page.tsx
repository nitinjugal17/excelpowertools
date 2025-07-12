
"use client";

import React, { useState, useCallback, ChangeEvent, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Split, Download, FileSpreadsheet, ListFilter, Loader2, CheckCircle2, ListOrdered, ArrowUp, ArrowDown, PencilLine, BookMarked, Lightbulb, XCircle, PlusCircle } from 'lucide-react';
import { groupDataByColumn, createWorkbookFromGroupedData, INTERNAL_NOT_FOUND_KEY, DISPLAY_NOT_FOUND_SHEET_NAME, GroupingResult } from '@/lib/excel-sheet-splitter';
import type { SplitterCustomHeaderConfig, Alignment, IndexSheetConfig, SplitterCustomColumnConfig } from '@/lib/excel-types';
import { sanitizeSheetName, parseSourceColumns } from '@/lib/excel-helpers';
import { Checkbox } from '@/components/ui/checkbox';
import { useLanguage } from '@/context/language-context';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Markup } from '@/components/ui/markup';


interface SheetInfo {
  name: string;
  rowCount: number;
}

interface GroupingState {
    groupedRows: { [key: string]: number[] };
    orderedKeys: string[];
}

interface HeaderInfo {
    name: string;
    index: number;
}

interface ExcelSheetSplitterPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function ExcelSheetSplitterPage({ onProcessingChange, onFileStateChange }: ExcelSheetSplitterPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sourceWorksheet, setSourceWorksheet] = useState<XLSX.WorkSheet | null>(null);
  const [columnIdentifier, setColumnIdentifier] = useState<string>(''); // Stores stringified index
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  
  const [availableHeaders, setAvailableHeaders] = useState<HeaderInfo[]>([]);
  const [columnOrder, setColumnOrder] = useState<string[]>([]);
  const [headerRowNumber, setHeaderRowNumber] = useState<number>(1);


  const [processedInfo, setProcessedInfo] = useState<SheetInfo[] | null>(null);
  const [groupingResult, setGroupingResult] = useState<GroupingState | null>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const cancellationRequested = useRef(false);
  const [outputFormat, setOutputFormat] = useState<'xlsx' | 'xlsm'>('xlsm');
  const { toast } = useToast();

  const [enableCustomHeaderInsertion, setEnableCustomHeaderInsertion] = useState<boolean>(false);
  const [customHeaderConfig, setCustomHeaderConfig] = useState<Omit<SplitterCustomHeaderConfig, 'text'>>({
      sourceColumnString: '',
      insertBeforeRow: 1,
      valueSeparator: ' - ',
      mergeAndCenter: true,
      styleOptions: {
          bold: true,
          italic: false,
          underline: false,
          alignment: 'center' as Alignment,
          fontName: 'Calibri',
          fontSize: 12,
      }
  });
  const [generatedCustomHeaderText, setGeneratedCustomHeaderText] = useState<string | null>(null);

  const [enableIndexSheet, setEnableIndexSheet] = useState<boolean>(false);
  const [indexSheetConfig, setIndexSheetConfig] = useState<IndexSheetConfig>({
    sheetName: 'Index',
    headerText: 'Sheet Index',
    headerRow: 1,
    headerCol: 'A',
    linksStartRow: 2,
    linksCol: 'A',
    backLinkText: 'Back to Index',
    backLinkRow: 1,
    backLinkCol: 'A',
  });

  const [enableCustomColumn, setEnableCustomColumn] = useState<boolean>(false);
  const [customColumnConfig, setCustomColumnConfig] = useState<SplitterCustomColumnConfig>({ name: 'New Column', value: '{SheetName}' });
  const prevCustomColumnName = useRef(customColumnConfig.name);

  useEffect(() => {
    if (onProcessingChange) {
      onProcessingChange(isProcessing);
    }
  }, [isProcessing, onProcessingChange]);
  
  useEffect(() => {
    if (onFileStateChange) {
      onFileStateChange(file !== null);
    }
  }, [file, onFileStateChange]);

  useEffect(() => {
    setSheetNames([]);
    setSelectedSheet('');
    setSourceWorksheet(null);
    setProcessedInfo(null);
    setGroupingResult(null);
    setColumnIdentifier('');
    setAvailableHeaders([]);
    setColumnOrder([]);
    setHeaderRowNumber(1);
    setEnableCustomHeaderInsertion(false);
    setEnableIndexSheet(false);
    setEnableCustomColumn(false);
    setCustomColumnConfig({ name: 'New Column', value: '{SheetName}' });
  }, [file]);

  useEffect(() => {
    const fetchHeaders = async () => {
      if (file && selectedSheet && headerRowNumber > 0) {
        setIsProcessing(true);
        try {
          const arrayBuffer = await file.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, cellDates: true, bookVBA: true, bookFiles: true });
          const worksheet = workbook.Sheets[selectedSheet];
          if (worksheet) {
            setSourceWorksheet(worksheet);
            const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
            const headersFromSheet = aoa[headerRowNumber - 1] || [];
            const headersWithIndices = headersFromSheet.map((h, i) => ({ name: String(h || `Column ${i + 1}`), index: i }));

            if (headersWithIndices.length > 0) {
              setAvailableHeaders(headersWithIndices);
              const headerNames = headersWithIndices.map(h => h.name);
              // Reset column order based on new headers, preserving custom column if enabled
              setColumnOrder(() => {
                  const newOrder = [...headerNames];
                  if (enableCustomColumn && customColumnConfig.name.trim() && !newOrder.includes(customColumnConfig.name.trim())) {
                      newOrder.push(customColumnConfig.name.trim());
                  }
                  return newOrder;
              });
              
              const currentIdentifierIsValid = headersWithIndices.some(h => String(h.index) === columnIdentifier);
              if (!currentIdentifierIsValid) {
                setColumnIdentifier('');
              }
            } else {
              setAvailableHeaders([]);
              setColumnOrder([]);
              setColumnIdentifier('');
              toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('splitter.toast.invalidHeader', {headerRowNumber})].flat().join(' '), variant: 'default' });
            }
          } else {
            setSourceWorksheet(null);
            setAvailableHeaders([]);
            setColumnOrder([]);
            setColumnIdentifier('');
          }
        } catch (error) {
          console.error('Error fetching headers:', error);
          toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('splitter.toast.invalidHeader', {headerRowNumber})].flat().join(' '), variant: 'destructive' });
          setSourceWorksheet(null);
          setAvailableHeaders([]);
          setColumnOrder([]);
          setColumnIdentifier('');
        } finally {
          setIsProcessing(false);
        }
      } else {
        setSourceWorksheet(null);
        setAvailableHeaders([]);
        setColumnOrder([]);
      }
    };
    fetchHeaders();
  }, [file, selectedSheet, headerRowNumber, toast, t, columnIdentifier, enableCustomColumn, customColumnConfig.name]);
  
  useEffect(() => {
    const currentName = customColumnConfig.name.trim();
    const oldName = prevCustomColumnName.current.trim();
    
    setColumnOrder(currentOrder => {
      const customColumnInList = currentOrder.includes(oldName);
      
      if (enableCustomColumn) {
        if (!currentName) { // If new name is empty, remove old one if it exists
          return currentOrder.filter(h => h !== oldName);
        }
        if (customColumnInList) { // Name changed, replace it
          return currentOrder.map(h => (h === oldName ? currentName : h));
        } else { // Not in list, add it
          return [...currentOrder, currentName];
        }
      } else { // Disabled
        return currentOrder.filter(h => h !== oldName);
      }
    });

    prevCustomColumnName.current = currentName;
  }, [enableCustomColumn, customColumnConfig.name]);


  const handleFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      if (!selectedFile.name.match(/\.(xlsx|xls|xlsm)$/)) {
        toast({
          title: [t('toast.invalidFileType')].flat().join(' '),
          description: [t('toast.invalidFileTypeDesc')].flat().join(' '),
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
            title: [t('toast.uploadErrorTitle')].flat().join(' '),
            description: [t('toast.uploadErrorDesc')].flat().join(' '),
            variant: "destructive"
        });
      });

    } else {
      setFile(null);
    }
  };
  
  useEffect(() => {
    const getSheetNames = async () => {
        if (file) {
            setIsProcessing(true);
            try {
                const arrayBuffer = await file.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, cellDates: true, bookVBA: true, bookFiles: true });
                setSheetNames(workbook.SheetNames);
                if (workbook.SheetNames.length > 0) {
                    setSelectedSheet(currentSelectedSheet => {
                        if (!workbook.SheetNames.includes(currentSelectedSheet) || !currentSelectedSheet) {
                            return workbook.SheetNames[0];
                        }
                        return currentSelectedSheet;
                    });
                } else {
                    setSelectedSheet('');
                }
            } catch (error) {
                console.error('Error reading file for sheet names:', error);
                toast({
                    title: [t('toast.errorReadingFile')].flat().join(' '),
                    description: [t('toast.errorReadingSheets')].flat().join(' '),
                    variant: 'destructive',
                });
                setSheetNames([]);
                setSelectedSheet('');
            } finally {
                setIsProcessing(false);
            }
        }
    };
    getSheetNames();
  }, [file, toast, t]);


  const moveColumn = (index: number, direction: 'up' | 'down') => {
    const newOrder = [...columnOrder];
    const item = newOrder[index];
    if (direction === 'up' && index > 0) {
      newOrder.splice(index, 1);
      newOrder.splice(index - 1, 0, item);
    } else if (direction === 'down' && index < newOrder.length - 1) {
      newOrder.splice(index, 1);
      newOrder.splice(index + 1, 0, item);
    }
    setColumnOrder(newOrder);
  };

  const handleProcessFile = useCallback(async () => {
    if (!file || !columnIdentifier || !selectedSheet || headerRowNumber < 1 || !sourceWorksheet) {
      toast({
        title: [t('toast.missingInfo')].flat().join(' '),
        description: [t('splitter.toast.missingInfo')].flat().join(' '),
        variant: 'destructive',
      });
      return;
    }
    if (enableCustomColumn && !customColumnConfig.name.trim()) {
      toast({
        title: [t('toast.missingInfo')].flat().join(' '),
        description: [t('splitter.toast.missingCustomColumnName')].flat().join(' '),
        variant: 'destructive',
      });
      return;
    }

    setIsProcessing(true);
    setProcessedInfo(null);
    setGroupingResult(null);

    try {
      const aoaData: any[][] = XLSX.utils.sheet_to_json(sourceWorksheet, { header: 1, defval: null, blankrows: false });
      
      if (enableCustomHeaderInsertion) {
        if (!customHeaderConfig.sourceColumnString.trim() || customHeaderConfig.insertBeforeRow < 1) {
            toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('splitter.toast.invalidCustomHeader')].flat().join(' '), variant: 'destructive'});
            setIsProcessing(false);
            return;
        }
        // Generate an *example* header text for the UI preview from the first data row of the source sheet.
        const firstDataRow = aoaData[headerRowNumber];
        if (firstDataRow) {
            const headers = aoaData[headerRowNumber - 1]?.map(String) || [];
            const sourceColIndices = parseSourceColumns(customHeaderConfig.sourceColumnString, headers);
            const valuesToJoin = sourceColIndices.map(colIdx => (firstDataRow?.[colIdx] ?? ""));
            setGeneratedCustomHeaderText(valuesToJoin.join(customHeaderConfig.valueSeparator));
        } else {
            setGeneratedCustomHeaderText([t('splitter.noDataForHeader')].flat().join(' '));
        }
      } else {
          setGeneratedCustomHeaderText(null);
      }

      const colIndexToGroup = parseInt(columnIdentifier, 10);
       if (isNaN(colIndexToGroup)) {
          toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('splitter.toast.invalidColumnIdentifier')].flat().join(' '), variant: "destructive" });
          setIsProcessing(false);
          return;
      }
      
      const { groupedRows: currentGroupedRows, orderedKeys: currentOrderedKeys } = groupDataByColumn(sourceWorksheet, colIndexToGroup, headerRowNumber);
      setGroupingResult({ groupedRows: currentGroupedRows, orderedKeys: currentOrderedKeys });

      const info = currentOrderedKeys.map(key => ({
        name: key === INTERNAL_NOT_FOUND_KEY ? DISPLAY_NOT_FOUND_SHEET_NAME : sanitizeSheetName(key),
        rowCount: currentGroupedRows[key].length,
      }));
      
      if (info.length === 0) {
         toast({
          title: [t('toast.missingInfo')].flat().join(' '),
          description: [t('splitter.toast.noData', {headerRowNumber, selectedSheet})].flat().join(' '),
          variant: 'default',
        });
        setIsProcessing(false);
        return;
      }
      setProcessedInfo(info);


      toast({
        title: [t('toast.processingComplete')].flat().join(' '),
        description: [t('splitter.toast.success', {count: info.length})].flat().join(' '),
        variant: 'default',
        action: <CheckCircle2 className="text-green-500" />,
      });

    } catch (error) {
      console.error('Error processing file:', error);
      const errorMessage = error instanceof Error ? error.message : [t('splitter.toast.errorUnknown')].flat().join(' ');
      toast({
        title: [t('toast.errorReadingFile')].flat().join(' '),
        description: `${[t('splitter.toast.errorPrefix')].flat().join(' ')} ${errorMessage}`,
        variant: 'destructive',
      });
    } finally {
      setIsProcessing(false);
    }
  }, [file, columnIdentifier, selectedSheet, headerRowNumber, toast, enableCustomHeaderInsertion, customHeaderConfig, sourceWorksheet, t, enableCustomColumn, customColumnConfig.name]);

  const handleCancel = () => {
    cancellationRequested.current = true;
    setProcessingStatus([t('common.cancelling')].flat().join(' '));
  };

  const handleDownload = useCallback(async () => {
    if (!groupingResult || !file || !sourceWorksheet) {
      toast({
        title: [t('toast.noDataToDownload')].flat().join(' '),
        description: [t('splitter.toast.noDownload')].flat().join(' '),
        variant: 'destructive',
      });
      return;
    }
    if (enableCustomColumn && !customColumnConfig.name.trim()) {
      toast({
        title: [t('toast.missingInfo')].flat().join(' '),
        description: [t('splitter.toast.missingCustomColumnName')].flat().join(' '),
        variant: 'destructive',
      });
      return;
    }

    cancellationRequested.current = false;
    setIsProcessing(true);
    setProcessingStatus([t('common.processing')].flat().join(' '));

    try {
      const finalColumnOrder = columnOrder.length > 0 ? columnOrder : availableHeaders.map(h => h.name);
      
      const finalCustomHeaderConfig = enableCustomHeaderInsertion ? customHeaderConfig : undefined;
      const finalIndexSheetConfig = enableIndexSheet ? indexSheetConfig : undefined;
      const finalCustomColumnConfig = enableCustomColumn ? customColumnConfig : undefined;
      
      const onProgress = (status: { groupKey: string, currentGroup: number, totalGroups: number }) => {
        if (cancellationRequested.current) throw new Error('Cancelled by user.');
        setProcessingStatus([t('splitter.toast.creatingSheet', {current: status.currentGroup, total: status.totalGroups, groupKey: status.groupKey})].flat().join(' '));
      };

      const newWorkbook = createWorkbookFromGroupedData(
        sourceWorksheet,
        groupingResult.groupedRows,
        finalColumnOrder,
        headerRowNumber - 1, // Pass 0-indexed header row
        finalCustomHeaderConfig,
        finalIndexSheetConfig,
        finalCustomColumnConfig,
        onProgress,
        cancellationRequested,
        groupingResult.orderedKeys
      );
      const originalFileName = file.name.substring(0, file.name.lastIndexOf('.'));
      XLSX.writeFile(newWorkbook, `${originalFileName}_split_custom.${outputFormat}`, { compression: true, cellStyles: true, bookType: outputFormat });
      toast({
        title: [t('toast.downloadSuccess')].flat().join(' '),
        description: [t('splitter.toast.downloadReady')].flat().join(' '),
      });
    } catch (error) {
      console.error('Error creating or downloading workbook:', error);
      const errorMessage = error instanceof Error ? error.message : [t('splitter.toast.errorDownload')].flat().join(' ');
      
      if (errorMessage !== 'Cancelled by user.') {
        toast({ title: [t('toast.downloadError')].flat().join(' '), description: errorMessage, variant: 'destructive' });
      } else {
        toast({ title: [t('toast.cancelledTitle')].flat().join(' '), description: [t('toast.cancelledDesc')].flat().join(' '), variant: 'default' });
      }
    } finally {
      setIsProcessing(false);
      cancellationRequested.current = false;
      setProcessingStatus('');
    }
  }, [groupingResult, file, toast, columnOrder, availableHeaders, customHeaderConfig, enableCustomHeaderInsertion, enableIndexSheet, indexSheetConfig, sourceWorksheet, headerRowNumber, t, outputFormat, enableCustomColumn, customColumnConfig]);

  return (
    <Card className="w-full max-w-lg md:max-w-xl lg:max-w-2xl xl:max-w-6xl shadow-xl relative">
      {isProcessing && (
        <div className="absolute inset-0 bg-background/80 backdrop-blur-sm flex flex-col items-center justify-center z-10 rounded-lg space-y-4">
            <div className="flex items-center gap-2 text-muted-foreground">
            <Loader2 className="h-6 w-6 animate-spin" />
            <span className="text-lg font-medium">{processingStatus || [t('common.processing')].flat().join(' ')}</span>
            </div>
            <Button variant="destructive" onClick={handleCancel}>
                <XCircle className="mr-2 h-4 w-4"/>
                {[t('common.cancel')].flat().join(' ')}
            </Button>
        </div>
      )}
      <CardHeader>
        <div className="flex items-center space-x-2 mb-2">
          <FileSpreadsheet className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{t('splitter.title')}</CardTitle>
        </div>
        <CardDescription className="font-body">
          {t('splitter.description')}
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{t('splitter.uploadStep')}</span>
          </Label>
          <Input
            id="file-upload"
            type="file"
            accept=".xlsx, .xls, .xlsm"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isProcessing}
          />
          {file && <p className="text-xs text-muted-foreground font-code">{[t('common.selectedFile', {fileName: file.name})].flat().join(' ')}</p>}
        </div>

        {sheetNames.length > 0 && (
          <div className="space-y-2">
            <Label htmlFor="sheet-select" className="flex items-center space-x-2 text-sm font-medium">
              <ListFilter className="h-5 w-5" />
              <span>{t('splitter.selectSheetStep')}</span>
            </Label>
            <Select value={selectedSheet} onValueChange={setSelectedSheet} disabled={isProcessing || sheetNames.length === 0}>
              <SelectTrigger id="sheet-select">
                <SelectValue placeholder={t('common.selectSheet') as string} />
              </SelectTrigger>
              <SelectContent>
                {sheetNames.map((name) => (
                  <SelectItem key={name} value={name}>
                    {name}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        )}

        {file && selectedSheet && (
            <div className="space-y-2">
            <Label htmlFor="header-row-number" className="flex items-center space-x-2 text-sm font-medium">
                <FileSpreadsheet className="h-5 w-5" />
                <span>{t('splitter.headerRowStep')}</span>
            </Label>
            <Input
                id="header-row-number"
                type="number"
                min="1"
                value={String(headerRowNumber)}
                onChange={(e) => {
                    const val = parseInt(e.target.value, 10);
                    setHeaderRowNumber(isNaN(val) || val < 1 ? 1 : val);
                }}
                disabled={isProcessing || !file || !selectedSheet}
                placeholder={t('splitter.identifierInputPlaceholder') as string}
            />
            <p className="text-xs text-muted-foreground">
                {t('splitter.headerRowDesc')}
            </p>
            </div>
        )}

        {file && selectedSheet && (
            <Card className="p-4 mt-6 border-dashed border-primary/50 bg-primary/5">
              <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                <Checkbox
                  id="enable-custom-header"
                  checked={enableCustomHeaderInsertion}
                  onCheckedChange={(checked) => setEnableCustomHeaderInsertion(checked as boolean)}
                  disabled={isProcessing}
                />
                <Label htmlFor="enable-custom-header" className="flex items-center space-x-2 text-md font-semibold text-primary">
                  <PencilLine className="h-5 w-5" />
                  <span>{t('splitter.customHeaderStep')}</span>
                </Label>
              </CardHeader>

              {enableCustomHeaderInsertion && (
                <CardContent className="p-0">
                    <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                      <div className="space-y-1">
                        <Label htmlFor="custom-header-source-cols" className="text-sm font-medium">{t('splitter.sourceCols')}</Label>
                        <Input id="custom-header-source-cols" type="text" value={customHeaderConfig.sourceColumnString} onChange={(e) => setCustomHeaderConfig(prev => ({ ...prev, sourceColumnString: e.target.value }))} disabled={isProcessing} placeholder={t('splitter.sourceColsPlaceholder') as string} />
                        <p className="text-xs text-muted-foreground"><Markup text={t('splitter.sourceColsDesc') as string}/></p>
                      </div>
                       <div className="space-y-1">
                        <Label htmlFor="custom-header-separator" className="text-sm font-medium">{t('splitter.separator')}</Label>
                        <Input id="custom-header-separator" type="text" value={customHeaderConfig.valueSeparator} onChange={(e) => setCustomHeaderConfig(prev => ({...prev, valueSeparator: e.target.value}))} disabled={isProcessing} placeholder={t('splitter.separatorPlaceholder') as string} />
                        <p className="text-xs text-muted-foreground">{t('splitter.separatorDesc')}</p>
                      </div>
                      <div className="space-y-1">
                        <Label htmlFor="custom-header-insert-row" className="text-sm font-medium">{t('splitter.insertRow')}</Label>
                        <Input id="custom-header-insert-row" type="number" min="1" value={customHeaderConfig.insertBeforeRow} onChange={(e) => setCustomHeaderConfig(prev => ({...prev, insertBeforeRow: Math.max(1, parseInt(e.target.value, 10) || 1)}))} disabled={isProcessing} placeholder={t('splitter.insertRowPlaceholder') as string} />
                        <p className="text-xs text-muted-foreground">{t('splitter.insertRowDesc')}</p>
                      </div>
                       <div className="flex items-center space-x-2 pt-2">
                        <Checkbox id="custom-header-merge-center" checked={!!customHeaderConfig.mergeAndCenter} onCheckedChange={(checked) => setCustomHeaderConfig(prev => ({ ...prev, mergeAndCenter: checked as boolean }))} disabled={isProcessing} />
                        <Label htmlFor="custom-header-merge-center" className="text-sm font-normal">{t('splitter.mergeAndCenter')}</Label>
                      </div>
                       <div className="grid grid-cols-2 gap-4 pt-4">
                            <div>
                                <Label htmlFor="custom-header-font-name" className="text-sm font-medium">{t('common.fontName')}</Label>
                                <Input id="custom-header-font-name" type="text" value={customHeaderConfig.styleOptions.fontName || ''} onChange={(e) => setCustomHeaderConfig(prev => ({...prev, styleOptions: {...prev.styleOptions, fontName: e.target.value}}))} disabled={isProcessing} placeholder="e.g., Calibri"/>
                            </div>
                            <div>
                                <Label htmlFor="custom-header-font-size" className="text-sm font-medium">{t('common.fontSize')}</Label>
                                <Input id="custom-header-font-size" type="number" min="1" value={customHeaderConfig.styleOptions.fontSize || 12} onChange={(e) => setCustomHeaderConfig(prev => ({...prev, styleOptions: {...prev.styleOptions, fontSize: parseInt(e.target.value, 10) || 11}}))} disabled={isProcessing} />
                            </div>
                        </div>
                       <div className="grid grid-cols-2 gap-4 pt-4">
                            <div className="flex flex-col space-y-2">
                               <div className="flex items-center space-x-2"><Checkbox id="h-format-bold" checked={!!customHeaderConfig.styleOptions.bold} onCheckedChange={(checked) => setCustomHeaderConfig(prev => ({ ...prev, styleOptions: { ...prev.styleOptions, bold: checked as boolean } }))} disabled={isProcessing} /><Label htmlFor="h-format-bold">{t('common.bold')}</Label></div>
                               <div className="flex items-center space-x-2"><Checkbox id="h-format-italic" checked={!!customHeaderConfig.styleOptions.italic} onCheckedChange={(checked) => setCustomHeaderConfig(prev => ({ ...prev, styleOptions: { ...prev.styleOptions, italic: checked as boolean } }))} disabled={isProcessing} /><Label htmlFor="h-format-italic">{t('common.italic')}</Label></div>
                               <div className="flex items-center space-x-2"><Checkbox id="h-format-underline" checked={!!customHeaderConfig.styleOptions.underline} onCheckedChange={(checked) => setCustomHeaderConfig(prev => ({ ...prev, styleOptions: { ...prev.styleOptions, underline: checked as boolean } }))} disabled={isProcessing} /><Label htmlFor="h-format-underline">{t('common.underline')}</Label></div>
                            </div>
                            <div className="space-y-1">
                                <Label htmlFor="h-alignment-select">{t('common.alignment')}</Label>
                                <Select value={customHeaderConfig.styleOptions.alignment || 'general'} onValueChange={(value) => setCustomHeaderConfig(prev => ({...prev, styleOptions: {...prev.styleOptions, alignment: value as Alignment}}))} disabled={isProcessing}>
                                    <SelectTrigger id="h-alignment-select"><SelectValue /></SelectTrigger>
                                    <SelectContent>
                                        <SelectItem value="general">{t('common.alignments.general')}</SelectItem>
                                        <SelectItem value="left">{t('common.alignments.left')}</SelectItem>
                                        <SelectItem value="center">{t('common.alignments.center')}</SelectItem>
                                        <SelectItem value="centerContinuous">{t('common.alignments.centerContinuous')}</SelectItem>
                                        <SelectItem value="right">{t('common.alignments.right')}</SelectItem>
                                        <SelectItem value="fill">{t('common.alignments.fill')}</SelectItem>
                                        <SelectItem value="justify">{t('common.alignments.justify')}</SelectItem>
                                    </SelectContent>
                                </Select>
                            </div>
                        </div>
                    </div>
                </CardContent>
              )}
            </Card>
        )}

        <Card className="p-4 mt-6 border-dashed border-primary/50 bg-primary/5">
            <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                <Checkbox
                    id="enable-index-sheet"
                    checked={enableIndexSheet}
                    onCheckedChange={(checked) => setEnableIndexSheet(checked as boolean)}
                    disabled={isProcessing}
                />
                <Label htmlFor="enable-index-sheet" className="flex items-center space-x-2 text-md font-semibold text-primary">
                    <BookMarked className="h-5 w-5" />
                    <span>{t('splitter.indexSheetStep')}</span>
                </Label>
            </CardHeader>

            {enableIndexSheet && (
            <CardContent className="p-0">
                <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                    <p className="text-sm text-muted-foreground pb-2">
                        {t('splitter.indexSheetDesc')}
                    </p>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-4">
                        <div className="space-y-1">
                            <Label htmlFor="index-sheet-name" className="text-sm font-medium">{t('splitter.indexSheetName')}</Label>
                            <Input id="index-sheet-name" value={indexSheetConfig.sheetName} onChange={(e) => setIndexSheetConfig(p => ({...p, sheetName: e.target.value}))} disabled={isProcessing} />
                        </div>
                         <div className="space-y-1">
                            <Label htmlFor="index-header-text" className="text-sm font-medium">{t('splitter.indexHeader')}</Label>
                            <Input id="index-header-text" value={indexSheetConfig.headerText} onChange={(e) => setIndexSheetConfig(p => ({...p, headerText: e.target.value}))} disabled={isProcessing} />
                        </div>
                         <div className="space-y-1">
                            <Label htmlFor="index-header-row" className="text-sm font-medium">{t('splitter.headerRow')}</Label>
                            <Input type="number" id="index-header-row" min="1" value={indexSheetConfig.headerRow} onChange={(e) => setIndexSheetConfig(p => ({...p, headerRow: parseInt(e.target.value, 10) || 1}))} disabled={isProcessing} />
                        </div>
                         <div className="space-y-1">
                            <Label htmlFor="index-header-col" className="text-sm font-medium">{t('splitter.headerCol')}</Label>
                            <Input id="index-header-col" value={indexSheetConfig.headerCol} onChange={(e) => setIndexSheetConfig(p => ({...p, headerCol: e.target.value}))} disabled={isProcessing} />
                        </div>
                         <div className="space-y-1">
                            <Label htmlFor="index-links-row" className="text-sm font-medium">{t('splitter.linksStartRow')}</Label>
                            <Input type="number" id="index-links-row" min="1" value={indexSheetConfig.linksStartRow} onChange={(e) => setIndexSheetConfig(p => ({...p, linksStartRow: parseInt(e.target.value, 10) || 1}))} disabled={isProcessing} />
                        </div>
                        <div className="space-y-1">
                            <Label htmlFor="index-links-col" className="text-sm font-medium">{t('splitter.linksCol')}</Label>
                            <Input id="index-links-col" value={indexSheetConfig.linksCol} onChange={(e) => setIndexSheetConfig(p => ({...p, linksCol: e.target.value}))} disabled={isProcessing} />
                        </div>
                        <div className="space-y-1">
                            <Label htmlFor="backlink-text" className="text-sm font-medium">{t('splitter.backLinkText')}</Label>
                            <Input id="backlink-text" value={indexSheetConfig.backLinkText} onChange={(e) => setIndexSheetConfig(p => ({...p, backLinkText: e.target.value}))} disabled={isProcessing} />
                        </div>
                        <div className="space-y-1">
                            <Label htmlFor="backlink-row" className="text-sm font-medium">{t('splitter.backLinkRow')}</Label>
                            <Input type="number" id="backlink-row" min="1" value={indexSheetConfig.backLinkRow} onChange={(e) => setIndexSheetConfig(p => ({...p, backLinkRow: parseInt(e.target.value, 10) || 1}))} disabled={isProcessing} />
                        </div>
                         <div className="space-y-1">
                            <Label htmlFor="backlink-col" className="text-sm font-medium">{t('splitter.backLinkCol')}</Label>
                            <Input id="backlink-col" value={indexSheetConfig.backLinkCol} onChange={(e) => setIndexSheetConfig(p => ({...p, backLinkCol: e.target.value}))} disabled={isProcessing} />
                        </div>
                    </div>
                </div>
            </CardContent>
            )}
        </Card>

        <Card className="p-4 mt-6 border-dashed border-primary/50 bg-primary/5">
            <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
            <Checkbox
                id="enable-custom-column"
                checked={enableCustomColumn}
                onCheckedChange={(checked) => setEnableCustomColumn(checked as boolean)}
                disabled={isProcessing}
            />
            <Label htmlFor="enable-custom-column" className="flex items-center space-x-2 text-md font-semibold text-primary">
                <PlusCircle className="h-5 w-5" />
                <span>{t('splitter.customColumnStep')}</span>
            </Label>
            </CardHeader>
            {enableCustomColumn && (
            <CardContent className="p-0">
                <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                <div className="space-y-1">
                    <Label htmlFor="custom-column-name" className="text-sm font-medium">{t('splitter.customColumnName')}</Label>
                    <Input id="custom-column-name" value={customColumnConfig.name} onChange={(e) => setCustomColumnConfig(p => ({ ...p, name: e.target.value }))} disabled={isProcessing} placeholder={t('splitter.customColumnNamePlaceholder') as string}/>
                </div>
                <div className="space-y-1">
                    <Label htmlFor="custom-column-value" className="text-sm font-medium">{t('splitter.customColumnValue')}</Label>
                    <Input id="custom-column-value" value={customColumnConfig.value} onChange={(e) => setCustomColumnConfig(p => ({ ...p, value: e.target.value }))} disabled={isProcessing} placeholder={t('splitter.customColumnValuePlaceholder') as string} />
                    <p className="text-xs text-muted-foreground">
                        <Markup text={t('splitter.customColumnValueDesc') as string} />
                    </p>
                </div>
                </div>
            </CardContent>
            )}
        </Card>


        {availableHeaders.length > 0 && (
          <div className="space-y-2">
            <Label htmlFor="column-order" className="flex items-center space-x-2 text-sm font-medium">
              <ListOrdered className="h-5 w-5" />
              <span>{t('splitter.columnOrderStep')}</span>
            </Label>
            <Card className="p-3 bg-secondary/20 max-h-60 overflow-y-auto">
              <ul className="space-y-1">
                {columnOrder.map((header, index) => (
                  <li key={`${header}-${index}`} className="flex items-center justify-between p-2 bg-background rounded-md shadow-sm">
                    <span className="text-sm truncate" title={header}>{header}</span>
                    <div className="space-x-1 flex-shrink-0">
                      <Button
                        variant="ghost"
                        size="icon"
                        className="h-7 w-7"
                        onClick={() => moveColumn(index, 'up')}
                        disabled={index === 0 || isProcessing}
                        aria-label={`Move ${header} up`}
                      >
                        <ArrowUp className="h-4 w-4" />
                      </Button>
                      <Button
                        variant="ghost"
                        size="icon"
                        className="h-7 w-7"
                        onClick={() => moveColumn(index, 'down')}
                        disabled={index === columnOrder.length - 1 || isProcessing}
                        aria-label={`Move ${header} down`}
                      >
                        <ArrowDown className="h-4 w-4" />
                      </Button>
                    </div>
                  </li>
                ))}
              </ul>
            </Card>
            <p className="text-xs text-muted-foreground">
              {t('splitter.columnOrderDesc')}
            </p>
          </div>
        )}
        
        <div className="space-y-2">
          <Label htmlFor="column-name" className="flex items-center space-x-2 text-sm font-medium">
             <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="lucide lucide-key-round"><path d="M2 18v3c0 .6.4 1 1 1h4v-3h3v-3h2l1.4-1.4a6.5 6.5 0 1 0-4-4Z"/><circle cx="16.5" cy="7.5" r=".5"/></svg>
            <span>{t('splitter.identifierStep')}</span>
          </Label>
          {availableHeaders.length > 0 ? ( 
            <Select
              value={columnIdentifier}
              onValueChange={(value) => setColumnIdentifier(value)}
              disabled={isProcessing || !file || availableHeaders.length === 0}
            >
              <SelectTrigger id="column-name-select">
                <SelectValue placeholder={t('splitter.identifierPlaceholder') as string} />
              </SelectTrigger>
              <SelectContent>
                {availableHeaders.map((header) => ( 
                  <SelectItem key={`identifier-col-${header.index}`} value={String(header.index)}>
                    {header.name}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          ) : (
            <Input
              id="column-name-input"
              type="text"
              placeholder={t('splitter.identifierInputPlaceholder') as string}
              value={columnIdentifier}
              onChange={(e) => setColumnIdentifier(e.target.value)}
              disabled={isProcessing || !file}
            />
          )}
          <p className="text-xs text-muted-foreground">
            {t('splitter.identifierDesc')}
          </p>
        </div>

        <Button
          onClick={handleProcessFile}
          disabled={isProcessing || !file || !columnIdentifier || !selectedSheet}
          className="w-full bg-primary hover:bg-primary/90 text-primary-foreground"
        >
          {isProcessing && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
          <Split className="mr-2 h-5 w-5" />
          {t('splitter.processBtn')}
        </Button>

        {processedInfo && processedInfo.length > 0 && (
          <div className="mt-6 p-4 border rounded-md bg-secondary/30">
            <h3 className="text-lg font-semibold mb-2 font-headline">{t('splitter.previewTitle')}</h3>
             {generatedCustomHeaderText && (
                <Alert className="mb-4">
                    <Lightbulb className="h-4 w-4" />
                    <AlertDescription>
                        <strong>{t('splitter.headerPreview')}</strong> {generatedCustomHeaderText}
                    </AlertDescription>
                </Alert>
            )}
            <ul className="space-y-1 max-h-40 overflow-y-auto text-sm">
              {processedInfo.map((info) => (
                <li key={info.name} className="flex justify-between">
                  <span>{t('splitter.sheet')} <span className="font-medium">{info.name}</span></span>
                  <span className="text-muted-foreground">{t('splitter.rows')} {info.rowCount}</span>
                </li>
              ))}
            </ul>
          </div>
        )}
      </CardContent>
      {processedInfo && processedInfo.length > 0 && (
        <CardFooter className="flex-col items-stretch space-y-4">
            <div className="w-full p-4 border rounded-md bg-secondary/30 space-y-4">
                <Label className="text-md font-semibold font-headline">{t('common.outputOptions.title')}</Label>
                <RadioGroup value={outputFormat} onValueChange={(v) => setOutputFormat(v as any)} className="space-y-3">
                    <div>
                        <div className="flex items-center space-x-2">
                            <RadioGroupItem value="xlsx" id="format-xlsx-splitter" />
                            <Label htmlFor="format-xlsx-splitter" className="font-normal">{t('common.outputOptions.xlsx')}</Label>
                        </div>
                        <p className="text-xs text-muted-foreground pl-6 pt-1">{t('common.outputOptions.xlsxDesc')}</p>
                    </div>
                    <div>
                        <div className="flex items-center space-x-2">
                            <RadioGroupItem value="xlsm" id="format-xlsm-splitter" />
                            <Label htmlFor="format-xlsm-splitter" className="font-normal">{t('common.outputOptions.xlsm')}</Label>
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
            </div>
          <Button
            onClick={handleDownload}
            disabled={isProcessing || !groupingResult}
            className="w-full bg-accent hover:bg-accent/90 text-accent-foreground"
          >
            {isProcessing && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
            <Download className="mr-2 h-5 w-5" />
            {t('splitter.downloadBtn')}
          </Button>
        </CardFooter>
      )}
    </Card>
  );
}
