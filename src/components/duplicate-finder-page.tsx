
"use client";

import React, { useState, useCallback, ChangeEvent, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, CopyCheck, CheckCircle2, Loader2, ListChecks, FileInput, Columns, Pencil, FileSpreadsheet, Link as LinkIcon, Edit, ChevronDown, Filter, Lightbulb, Palette, TextCursorInput, ScrollText, XCircle, Settings } from 'lucide-react';
import { findAndMarkDuplicates, createDuplicateReportWorkbook, addDuplicateReportToSheets } from '@/lib/excel-duplicate-finder';
import { generateDuplicateFinderVbs } from '@/lib/vbs-generators';
import type { DuplicateReport } from '@/lib/excel-types';
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from '@/components/ui/dropdown-menu';
import { useLanguage } from '@/context/language-context';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';
import { Markup } from '@/components/ui/markup';

interface SheetSelection {
  [sheetName: string]: boolean;
}

const isValidHex = (hex: string) => /^([0-9A-F]{6})$/i.test(hex);

interface DuplicateFinderPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function DuplicateFinderPage({ onProcessingChange, onFileStateChange }: DuplicateFinderPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<SheetSelection>({});
  
  const [headerRow, setHeaderRow] = useState<number>(1);
  const [keyColumns, setKeyColumns] = useState<string>(''); // e.g., A,B or Name,Email
  const [updateColumn, setUpdateColumn] = useState<string>(''); // e.g., C or Status
  
  const [updateMode, setUpdateMode] = useState<'template' | 'context'>('template');
  const [updateValue, setUpdateValue] = useState<string>('DUPLICATE'); // e.g., DUPLICATE or {Status}
  const [contextColumns, setContextColumns] = useState<string>('');
  const [stripText, setStripText] = useState<string>('');
  const [contextDelimiter, setContextDelimiter] = useState<string>('');
  const [contextPartToUse, setContextPartToUse] = useState<number>(1);


  const [enableDuplicateHighlight, setEnableDuplicateHighlight] = useState<boolean>(false);
  const [duplicateHighlightColor, setDuplicateHighlightColor] = useState<string>('FFFF00');

  const [enableConditionalMarking, setEnableConditionalMarking] = useState<boolean>(false);
  const [conditionalColumn, setConditionalColumn] = useState<string>('');

  const [enableInSheetReport, setEnableInSheetReport] = useState<boolean>(false);
  const [reportInsertCol, setReportInsertCol] = useState<string>('J');
  const [reportInsertRow, setReportInsertRow] = useState<number>(1);
  const [primaryContextCol, setPrimaryContextCol] = useState<string>('');
  const [fallbackContextCol, setFallbackContextCol] = useState<string>('');

  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const cancellationRequested = useRef(false);

  const [processedReport, setProcessedReport] = useState<DuplicateReport | null>(null);
  const [modifiedWorkbook, setModifiedWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [outputFormat, setOutputFormat] = useState<'xlsx' | 'xlsm'>('xlsm');
  const { toast } = useToast();
  const [vbscriptPreview, setVbscriptPreview] = useState<string>('');
  const [reportChunkSize, setReportChunkSize] = useState<number>(100000);

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
    // Reset all state when the file changes to ensure a clean slate.
    setSheetNames([]);
    setSelectedSheets({});
    setHeaderRow(1);
    setKeyColumns('');
    setUpdateColumn('');
    setUpdateValue('DUPLICATE');
    setEnableDuplicateHighlight(false);
    setDuplicateHighlightColor('FFFF00');
    setEnableConditionalMarking(false);
    setConditionalColumn('');
    setEnableInSheetReport(false);
    setReportInsertCol('J');
    setReportInsertRow(1);
    setPrimaryContextCol('');
    setFallbackContextCol('');
    setProcessedReport(null);
    setModifiedWorkbook(null);
    setUpdateMode('template');
    setContextColumns('');
    setStripText('');
    setContextDelimiter('');
    setContextPartToUse(1);
    setReportChunkSize(100000);

    if (file) {
      const getSheetNamesFromFile = async () => {
        setIsProcessing(true);
        try {
          const arrayBuffer = await file.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, cellDates: true, bookVBA: true, bookFiles: true });
          setSheetNames(workbook.SheetNames);
          const initialSelection: SheetSelection = {};
          workbook.SheetNames.forEach(name => {
            initialSelection[name] = true; // Select all by default
          });
          setSelectedSheets(initialSelection);
        } catch (error) {
          console.error("Error reading sheet names:", error);
          toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('toast.errorReadingSheets')].flat().join(' '), variant: "destructive" });
        } finally {
          setIsProcessing(false);
        }
      };
      getSheetNamesFromFile();
    }
  }, [file, toast, t]);

  useEffect(() => {
    const sheetsToUpdate = Object.entries(selectedSheets)
      .filter(([,isSelected]) => isSelected)
      .map(([sheetName]) => sheetName);
      
    const script = generateDuplicateFinderVbs(
        sheetsToUpdate,
        keyColumns,
        updateColumn,
        updateValue,
        headerRow,
        enableDuplicateHighlight ? duplicateHighlightColor : undefined,
        enableConditionalMarking ? conditionalColumn : undefined
    );
    setVbscriptPreview(script);

  }, [selectedSheets, keyColumns, updateColumn, updateValue, headerRow, enableDuplicateHighlight, duplicateHighlightColor, enableConditionalMarking, conditionalColumn]);


  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      if (!selectedFile.name.match(/\.(xlsx|xls|xlsm)$/)) {
        toast({ title: [t('toast.invalidFileType')].flat().join(' '), description: [t('toast.invalidFileTypeDesc')].flat().join(' '), variant: 'destructive' });
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

  const handleSelectAllSheets = (checked: boolean) => {
    setSelectedSheets(prev => {
      const newSelection: SheetSelection = {};
      sheetNames.forEach(name => {
        newSelection[name] = checked;
      });
      return newSelection;
    });
  };

  const handlePartialSelection = (count: number) => {
    const newSelection: SheetSelection = {};
    sheetNames.forEach((name, index) => {
      newSelection[name] = index < count;
    });
    setSelectedSheets(newSelection);
  };
  
  const handleCancel = () => {
    cancellationRequested.current = true;
    setProcessingStatus([t('common.cancelling')].flat().join(' '));
  };

  const handleSheetSelectionChange = (sheetName: string, checked: boolean) => {
    setSelectedSheets(prev => ({ ...prev, [sheetName]: checked }));
  };

  const handleProcess = useCallback(async () => {
    const sheetsToProcess = sheetNames.filter(name => selectedSheets[name]);
    if (!file || sheetsToProcess.length === 0 || !keyColumns.trim() || !updateColumn.trim() || headerRow < 1) {
      toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('duplicates.toast.missingInfo')].flat().join(' '), variant: 'destructive' });
      return;
    }
    if (updateMode === 'context' && !contextColumns.trim()) {
      toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('duplicates.toast.missingContextCols')].flat().join(' '), variant: 'destructive' });
      return;
    }
    if (enableDuplicateHighlight && !isValidHex(duplicateHighlightColor)) {
        toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('duplicates.toast.invalidHighlightColor', { hex: duplicateHighlightColor })].flat().join(' '), variant: 'destructive' });
        return;
    }
    if (enableConditionalMarking && !conditionalColumn.trim()) {
        toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('duplicates.toast.missingConditionalCol')].flat().join(' '), variant: 'destructive' });
        return;
    }
    if (enableInSheetReport && (!reportInsertCol.trim() || reportInsertRow < 1 || !primaryContextCol.trim())) {
      toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('duplicates.toast.missingInSheetInfo')].flat().join(' '), variant: 'destructive' });
      return;
    }

    cancellationRequested.current = false;
    setIsProcessing(true);
    setProcessingStatus('');
    setProcessedReport(null);
    setModifiedWorkbook(null);
    
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, cellDates: true, bookVBA: true, bookFiles: true });
      
      const updateConfig = {
        mode: updateMode,
        value: updateMode === 'template' ? updateValue : contextColumns,
      };

      const onProgress = (status: { sheetName: string; currentSheet: number; totalSheets: number; duplicatesFound: number }) => {
        if (cancellationRequested.current) {
            throw new Error('Cancelled by user.');
        }
        setProcessingStatus(
          `${[t('duplicates.toast.processingSheet', {current: status.currentSheet, total: status.totalSheets, sheetName: status.sheetName, count: ''})].flat().join(' ')} | ${[t('duplicates.resultsFound', { count: status.duplicatesFound })].flat().join(' ')}`
        );
      };

      const { report, workbook: newWb } = findAndMarkDuplicates(
        workbook,
        sheetsToProcess,
        keyColumns,
        updateColumn,
        updateConfig,
        headerRow,
        enableDuplicateHighlight ? duplicateHighlightColor : undefined,
        enableConditionalMarking ? conditionalColumn : undefined,
        stripText,
        contextDelimiter,
        contextPartToUse,
        onProgress,
        cancellationRequested
      );

      if (enableInSheetReport && report.totalDuplicates > 0) {
        addDuplicateReportToSheets(newWb, report, {
          insertCol: reportInsertCol,
          insertRow: reportInsertRow,
          primaryContextCol,
          fallbackContextCol,
          headerRow,
        });
      }


      setProcessedReport(report);
      setModifiedWorkbook(newWb);
      
      toast({
        title: [t('toast.processingComplete')].flat().join(' '),
        description: [t('duplicates.toast.success', { count: report.totalDuplicates, sheets: Object.keys(report.summary).length })].flat().join(' '),
        action: <CheckCircle2 className="text-green-500" />,
      });

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : [t('duplicates.toast.error')].flat().join(' ');
      if (errorMessage !== 'Cancelled by user.') {
        toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: errorMessage, variant: 'destructive' });
      } else {
        toast({ title: [t('toast.cancelledTitle')].flat().join(' '), description: [t('toast.cancelledDesc')].flat().join(' '), variant: 'default' });
      }
    } finally {
      setIsProcessing(false);
      cancellationRequested.current = false;
      setProcessingStatus('');
    }
  }, [file, sheetNames, selectedSheets, headerRow, keyColumns, updateColumn, updateValue, toast, enableDuplicateHighlight, duplicateHighlightColor, enableConditionalMarking, conditionalColumn, enableInSheetReport, reportInsertCol, reportInsertRow, primaryContextCol, fallbackContextCol, t, updateMode, contextColumns, stripText, contextDelimiter, contextPartToUse]);

  const handleDownloadReport = useCallback(() => {
    if (!processedReport || !file) {
      toast({ title: [t('toast.noDataToDownload')].flat().join(' '), description: [t('duplicates.toast.noReport')].flat().join(' '), variant: "destructive" });
      return;
    }
    try {
        const reportWb = createDuplicateReportWorkbook(processedReport, reportChunkSize);
        const originalFileName = file.name.substring(0, file.name.lastIndexOf('.'));
        XLSX.writeFile(reportWb, `${originalFileName}_duplicate_report.${outputFormat}`, { compression: true, bookType: outputFormat, cellStyles: true });
        toast({ title: [t('toast.downloadSuccess')].flat().join(' ') });
    } catch (error) {
        toast({ title: [t('toast.downloadError')].flat().join(' '), description: [t('duplicates.toast.downloadReportError')].flat().join(' '), variant: 'destructive' });
    }
  }, [processedReport, file, toast, t, outputFormat, reportChunkSize]);

  const handleDownloadModifiedFile = useCallback(() => {
     if (!modifiedWorkbook || !file) {
       toast({ title: [t('toast.noFileToDownload')].flat().join(' '), description: [t('duplicates.toast.noFile')].flat().join(' '), variant: "destructive" });
       return;
     }
     try {
        const originalFileName = file.name.substring(0, file.name.lastIndexOf('.'));
        XLSX.writeFile(modifiedWorkbook, `${originalFileName}_duplicates_marked.${outputFormat}`, { compression: true, bookType: outputFormat, cellStyles: true });
        toast({ title: [t('toast.downloadSuccess')].flat().join(' ') });
     } catch (error) {
        toast({ title: [t('toast.downloadError')].flat().join(' '), description: [t('duplicates.toast.downloadFileError')].flat().join(' '), variant: 'destructive' });
     }
  }, [modifiedWorkbook, file, toast, t, outputFormat]);

  const allSheetsSelected = sheetNames.length > 0 && sheetNames.every(name => selectedSheets[name]);
  const hasResults = processedReport !== null;

  return (
    <Card className="w-full max-w-lg md:max-w-4xl xl:max-w-6xl shadow-xl relative">
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
          <CopyCheck className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{[t('duplicates.title')].flat().join(' ')}</CardTitle>
        </div>
        <CardDescription className="font-body">
          {[t('duplicates.description')].flat().join(' ')}
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-dupes" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{[t('duplicates.uploadStep')].flat().join(' ')}</span>
          </Label>
          <Input
            id="file-upload-dupes"
            type="file"
            accept=".xlsx, .xls, .xlsm"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isProcessing}
          />
        </div>

        {sheetNames.length > 0 && (
          <div className="space-y-3">
            <Label className="flex items-center space-x-2 text-sm font-medium mb-2">
              <ListChecks className="h-5 w-5" />
              <span>{[t('duplicates.selectSheetsStep')].flat().join(' ')}</span>
            </Label>
            <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
              <Checkbox
                id="select-all-sheets-dupes"
                checked={allSheetsSelected}
                onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                disabled={isProcessing}
              />
              <Label htmlFor="select-all-sheets-dupes" className="text-sm font-medium flex-grow">
                {[t('common.selectAll')].flat().join(' ')} ({[t('common.selectedCount', {selected: Object.values(selectedSheets).filter(Boolean).length, total: sheetNames.length})].flat().join(' ')})
              </Label>
              {sheetNames.length > 50 && (
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <Button variant="outline" size="sm" disabled={isProcessing}>
                        {[t('common.partial')].flat().join(' ')}
                        <ChevronDown className="ml-1 h-4 w-4" />
                    </Button>
                  </DropdownMenuTrigger>
                  <DropdownMenuContent>
                    <DropdownMenuItem onSelect={() => handlePartialSelection(50)}>{[t('common.first50')].flat().join(' ')}</DropdownMenuItem>
                    {sheetNames.length >= 100 && <DropdownMenuItem onSelect={() => handlePartialSelection(100)}>{[t('common.first100')].flat().join(' ')}</DropdownMenuItem>}
                    {sheetNames.length >= 150 && <DropdownMenuItem onSelect={() => handlePartialSelection(150)}>{[t('common.first150')].flat().join(' ')}</DropdownMenuItem>}
                  </DropdownMenuContent>
                </DropdownMenu>
              )}
            </div>
            <Card className="max-h-48 overflow-y-auto p-3 bg-background">
              <div className="space-y-2">
                {sheetNames.map(name => (
                  <div key={name} className="flex items-center space-x-2">
                    <Checkbox
                      id={`sheet-dupe-${name}`}
                      checked={selectedSheets[name] || false}
                      onCheckedChange={(checked) => handleSheetSelectionChange(name, checked as boolean)}
                      disabled={isProcessing}
                    />
                    <Label htmlFor={`sheet-dupe-${name}`} className="text-sm font-normal">{name}</Label>
                  </div>
                ))}
              </div>
            </Card>
          </div>
        )}

        <div className="space-y-2">
            <Label htmlFor="header-row" className="flex items-center space-x-2 text-sm font-medium">
                <FileSpreadsheet className="h-5 w-5" />
                <span>{[t('duplicates.headerRowStep')].flat().join(' ')}</span>
            </Label>
            <Input 
                id="header-row" 
                type="number" 
                min="1" 
                value={headerRow} 
                onChange={(e) => setHeaderRow(parseInt(e.target.value, 10) || 1)} 
                disabled={isProcessing || !file}
            />
            <p className="text-xs text-muted-foreground">{[t('duplicates.headerRowDesc')].flat().join(' ')}</p>
        </div>
        <div className="space-y-2">
            <Label htmlFor="key-columns" className="flex items-center space-x-2 text-sm font-medium">
                <Columns className="h-5 w-5" />
                <span>{[t('duplicates.keyColsStep')].flat().join(' ')}</span>
            </Label>
            <Input 
                id="key-columns" 
                value={keyColumns} 
                onChange={e => setKeyColumns(e.target.value)} 
                disabled={isProcessing || !file}
                placeholder={[t('duplicates.keyColsPlaceholder')].flat().join(' ')}
            />
            <p className="text-xs text-muted-foreground">{[t('duplicates.keyColsDesc')].flat().join(' ')}</p>
        </div>
        <div className="space-y-2">
            <Label htmlFor="update-column" className="flex items-center space-x-2 text-sm font-medium">
                <Pencil className="h-5 w-5" />
                <span>{[t('duplicates.updateColStep')].flat().join(' ')}</span>
            </Label>
            <Input 
                id="update-column" 
                value={updateColumn} 
                onChange={e => setUpdateColumn(e.target.value)} 
                disabled={isProcessing || !file}
                placeholder={[t('duplicates.updateColPlaceholder')].flat().join(' ')}
            />
            <p className="text-xs text-muted-foreground"><Markup text={[t('duplicates.updateColDesc')].flat().join(' ')}/></p>
        </div>
        
        <div className="space-y-2">
          <Label className="flex items-center space-x-2 text-sm font-medium">
              <Pencil className="h-5 w-5" />
              <span>{[t('duplicates.updateMode')].flat().join(' ')}</span>
          </Label>
          <RadioGroup value={updateMode} onValueChange={v => setUpdateMode(v as any)} className="space-y-2 pt-2">
            <Card className="p-4">
              <div className="flex items-start space-x-3">
                  <RadioGroupItem value="template" id="update-mode-template" className="mt-1" />
                  <div className="grid gap-1.5 w-full">
                    <Label htmlFor="update-mode-template">{[t('duplicates.templateMode')].flat().join(' ')}</Label>
                    <p className="text-xs text-muted-foreground">{[t('duplicates.templateModeDesc')].flat().join(' ')}</p>
                    {updateMode === 'template' && (
                        <div className="pt-2">
                            <Input 
                                id="update-value" 
                                value={updateValue} 
                                onChange={e => setUpdateValue(e.target.value)}
                                disabled={isProcessing || !file}
                                placeholder={[t('duplicates.updateValPlaceholder')].flat().join(' ')}
                            />
                            <p className="text-xs text-muted-foreground pt-1"><Markup text={[t('duplicates.updateValDesc')].flat().join(' ')}/></p>
                        </div>
                    )}
                  </div>
              </div>
            </Card>
            <Card className="p-4">
              <div className="flex items-start space-x-3">
                  <RadioGroupItem value="context" id="update-mode-context" className="mt-1" />
                  <div className="grid gap-1.5 w-full">
                    <Label htmlFor="update-mode-context">{[t('duplicates.contextMode')].flat().join(' ')}</Label>
                    <p className="text-xs text-muted-foreground">{[t('duplicates.contextModeDesc')].flat().join(' ')}</p>
                    {updateMode === 'context' && (
                        <div className="pt-2 space-y-4">
                          <div>
                           <Label htmlFor="context-columns" className="text-sm font-medium">{[t('duplicates.contextCols')].flat().join(' ')}</Label>
                            <Input 
                                id="context-columns" 
                                value={contextColumns} 
                                onChange={e => setContextColumns(e.target.value)}
                                disabled={isProcessing || !file}
                                placeholder={[t('duplicates.contextColsPlaceholder')].flat().join(' ')}
                            />
                             <p className="text-xs text-muted-foreground pt-1">{[t('duplicates.contextColsDesc')].flat().join(' ')}</p>
                          </div>
                          <div>
                            <Label htmlFor="strip-text" className="text-sm font-medium">{[t('duplicates.stripText')].flat().join(' ')}</Label>
                            <Input 
                                id="strip-text" 
                                value={stripText} 
                                onChange={e => setStripText(e.target.value)}
                                disabled={isProcessing || !file}
                                placeholder={[t('duplicates.stripTextPlaceholder')].flat().join(' ')}
                            />
                            <p className="text-xs text-muted-foreground pt-1">{[t('duplicates.stripTextDesc')].flat().join(' ')}</p>
                          </div>
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                             <div>
                               <Label htmlFor="context-delimiter" className="text-sm font-medium">{[t('duplicates.contextDelimiter')].flat().join(' ')}</Label>
                                <Input 
                                    id="context-delimiter" 
                                    value={contextDelimiter} 
                                    onChange={e => setContextDelimiter(e.target.value)}
                                    disabled={isProcessing || !file}
                                    placeholder={[t('duplicates.contextDelimiterPlaceholder')].flat().join(' ')}
                                />
                                <p className="text-xs text-muted-foreground pt-1">{[t('duplicates.contextDelimiterDesc')].flat().join(' ')}</p>
                             </div>
                             <div>
                               <Label htmlFor="context-part" className="text-sm font-medium">{[t('duplicates.contextPartToUse')].flat().join(' ')}</Label>
                                <Input 
                                    id="context-part" 
                                    type="number"
                                    value={contextPartToUse} 
                                    onChange={e => setContextPartToUse(parseInt(e.target.value, 10) || 1)}
                                    disabled={isProcessing || !file || !contextDelimiter}
                                />
                                <p className="text-xs text-muted-foreground pt-1">{[t('duplicates.contextPartToUseDesc')].flat().join(' ')}</p>
                             </div>
                          </div>
                        </div>
                    )}
                  </div>
              </div>
            </Card>
          </RadioGroup>
        </div>


        <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
            <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                <Checkbox
                id="enable-duplicate-highlight"
                checked={enableDuplicateHighlight}
                onCheckedChange={(checked) => setEnableDuplicateHighlight(checked as boolean)}
                disabled={isProcessing}
                />
                <Label htmlFor="enable-duplicate-highlight" className="flex items-center space-x-2 text-md font-semibold text-primary">
                <Palette className="h-5 w-5" />
                <span>{[t('duplicates.highlightStep')].flat().join(' ')}</span>
                </Label>
            </CardHeader>
            {enableDuplicateHighlight && (
                <CardContent className="p-0">
                    <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                        <p className="text-xs text-muted-foreground -mt-2 mb-2">{[t('duplicates.highlightDesc')].flat().join(' ')}</p>
                        <div className="space-y-1">
                            <Label htmlFor="duplicate-highlight-color" className="text-sm">{[t('duplicates.highlightColor')].flat().join(' ')}</Label>
                            <Input 
                                id="duplicate-highlight-color" 
                                value={duplicateHighlightColor} 
                                onChange={e => setDuplicateHighlightColor(e.target.value.toUpperCase().replace(/[^0-9A-F]/g, ''))} 
                                placeholder={[t('duplicates.highlightColorPlaceholder')].flat().join(' ')}
                                maxLength={6}
                                disabled={isProcessing} />
                        </div>
                    </div>
                </CardContent>
            )}
        </Card>

        <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
            <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                <Checkbox
                id="enable-conditional-marking"
                checked={enableConditionalMarking}
                onCheckedChange={(checked) => setEnableConditionalMarking(checked as boolean)}
                disabled={isProcessing}
                />
                <Label htmlFor="enable-conditional-marking" className="flex items-center space-x-2 text-md font-semibold text-primary">
                <Filter className="h-5 w-5" />
                <span>{[t('duplicates.conditionalMarkingStep')].flat().join(' ')}</span>
                </Label>
            </CardHeader>
            {enableConditionalMarking && (
                <CardContent className="p-0">
                    <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                        <div className="space-y-1">
                            <Label htmlFor="conditional-column" className="text-sm">{[t('duplicates.conditionalMarkingCol')].flat().join(' ')}</Label>
                            <Input id="conditional-column" value={conditionalColumn} onChange={e => setConditionalColumn(e.target.value)} placeholder={[t('duplicates.conditionalMarkingColPlaceholder')].flat().join(' ')} disabled={isProcessing} />
                            <p className="text-xs text-muted-foreground">{[t('duplicates.conditionalMarkingColDesc')].flat().join(' ')}</p>
                        </div>
                    </div>
                </CardContent>
            )}
        </Card>
        
        <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
            <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                <Checkbox
                id="enable-in-sheet-report"
                checked={enableInSheetReport}
                onCheckedChange={(checked) => setEnableInSheetReport(checked as boolean)}
                disabled={isProcessing}
                />
                <Label htmlFor="enable-in-sheet-report" className="flex items-center space-x-2 text-md font-semibold text-primary">
                <Edit className="h-5 w-5" />
                <span>{[t('duplicates.inSheetReportStep')].flat().join(' ')}</span>
                </Label>
            </CardHeader>
            {enableInSheetReport && (
                <CardContent className="p-0">
                    <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div className="space-y-1">
                                <Label htmlFor="report-insert-col" className="text-sm">{[t('duplicates.insertAtCol')].flat().join(' ')}</Label>
                                <Input id="report-insert-col" value={reportInsertCol} onChange={e => setReportInsertCol(e.target.value)} placeholder={[t('duplicates.insertAtColPlaceholder')].flat().join(' ')} disabled={isProcessing} />
                            </div>
                            <div className="space-y-1">
                                <Label htmlFor="report-insert-row" className="text-sm">{[t('duplicates.insertAtRow')].flat().join(' ')}</Label>
                                <Input id="report-insert-row" type="number" min="1" value={reportInsertRow} onChange={e => setReportInsertRow(parseInt(e.target.value, 10) || 1)} disabled={isProcessing} />
                            </div>
                        </div>
                        <div className="space-y-1">
                            <Label htmlFor="primary-context-col" className="text-sm">{[t('duplicates.primaryContextCol')].flat().join(' ')}</Label>
                            <Input id="primary-context-col" value={primaryContextCol} onChange={e => setPrimaryContextCol(e.target.value)} placeholder={[t('duplicates.primaryContextColPlaceholder')].flat().join(' ')} disabled={isProcessing} />
                             <p className="text-xs text-muted-foreground">{[t('duplicates.primaryContextColDesc')].flat().join(' ')}</p>
                        </div>
                        <div className="space-y-1">
                            <Label htmlFor="fallback-context-col" className="text-sm">{[t('duplicates.fallbackContextCol')].flat().join(' ')}</Label>
                            <Input id="fallback-context-col" value={fallbackContextCol} onChange={e => setFallbackContextCol(e.target.value)} placeholder={[t('duplicates.fallbackContextColPlaceholder')].flat().join(' ')} disabled={isProcessing} />
                             <p className="text-xs text-muted-foreground">{[t('duplicates.fallbackContextColDesc')].flat().join(' ')}</p>
                        </div>
                    </div>
                </CardContent>
            )}
        </Card>
        
        <Accordion type="single" collapsible>
            <AccordionItem value="advanced-settings">
                <AccordionTrigger className="text-md font-semibold">{[t('duplicates.advancedSettings.title')].flat().join(' ')}</AccordionTrigger>
                <AccordionContent>
                    <Card className="p-4 border-dashed">
                       <div className="space-y-2">
                            <Label htmlFor="report-chunk-size" className="text-sm font-medium">{[t('duplicates.advancedSettings.maxRows')].flat().join(' ')}</Label>
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
                                {[t('duplicates.advancedSettings.maxRowsDesc')].flat().join(' ')}
                            </p>
                        </div>
                    </Card>
                </AccordionContent>
            </AccordionItem>
        </Accordion>

        {file && (
          <div className="space-y-2">
            <Label className="flex items-center space-x-2 text-sm font-medium">
              <ScrollText className="h-5 w-5" />
              <span>{[t('updater.vbsPreviewStep')].flat().join(' ')}</span>
            </Label>
            <Card className="bg-secondary/20">
              <CardContent className="p-0">
                <pre className="text-xs p-4 overflow-x-auto bg-gray-800 text-white rounded-md max-h-60">
                  <code>{vbscriptPreview}</code>
                </pre>
              </CardContent>
            </Card>
             <p className="text-xs text-muted-foreground">
              <Markup text={[t('updater.vbsPreviewDesc')].flat().join(' ')} />
            </p>
          </div>
        )}

        <Button onClick={handleProcess} disabled={isProcessing || !file} className="w-full">
          {isProcessing && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
          <CopyCheck className="mr-2 h-5 w-5" />
          {[t('duplicates.processBtn')].flat().join(' ')}
        </Button>
      </CardContent>

      {hasResults && (
        <CardFooter className="flex-col space-y-4 items-stretch">
          <div className="p-4 border rounded-md bg-secondary/30">
            <h3 className="text-lg font-semibold mb-2 font-headline">{[t('duplicates.resultsTitle')].flat().join(' ')}</h3>
            <p>{[t('duplicates.resultsFound', { count: processedReport.totalDuplicates })].flat().join(' ')}</p>
            <ul className="text-sm mt-2 max-h-32 overflow-y-auto">
              {Object.entries(processedReport.summary).map(([sheetName, count]) => (
                <li key={sheetName} className="flex justify-between">
                  <span>{sheetName}:</span>
                  <span className="font-medium">{count}</span>
                </li>
              ))}
            </ul>
          </div>
            <div className="w-full p-4 border rounded-md bg-secondary/30 space-y-4">
                <Label className="text-md font-semibold font-headline">{[t('common.outputOptions.title')].flat().join(' ')}</Label>
                <RadioGroup value={outputFormat} onValueChange={(v) => setOutputFormat(v as any)} className="space-y-3">
                    <div>
                        <div className="flex items-center space-x-2">
                            <RadioGroupItem value="xlsx" id="format-xlsx-dupes" />
                            <Label htmlFor="format-xlsx-dupes" className="font-normal">{[t('common.outputOptions.xlsx')].flat().join(' ')}</Label>
                        </div>
                        <p className="text-xs text-muted-foreground pl-6 pt-1">{[t('common.outputOptions.xlsxDesc')].flat().join(' ')}</p>
                    </div>
                    <div>
                        <div className="flex items-center space-x-2">
                            <RadioGroupItem value="xlsm" id="format-xlsm-dupes" />
                            <Label htmlFor="format-xlsm-dupes" className="font-normal">{[t('common.outputOptions.xlsm')].flat().join(' ')}</Label>
                        </div>
                        <p className="text-xs text-muted-foreground pl-6 pt-1">{[t('common.outputOptions.xlsmDesc')].flat().join(' ')}</p>
                    </div>
                </RadioGroup>
                <Alert variant="default" className="mt-2">
                    <Lightbulb className="h-4 w-4" />
                    <AlertDescription>
                        {[t('common.outputOptions.recommendation')].flat().join(' ')}
                    </AlertDescription>
                </Alert>
            </div>
          <Button onClick={handleDownloadModifiedFile} variant="outline" className="w-full">
              <Download className="mr-2 h-5 w-5" />
              {[t('duplicates.downloadMarkedBtn')].flat().join(' ')}
          </Button>
          <Button onClick={handleDownloadReport} className="w-full bg-accent hover:bg-accent/90 text-accent-foreground">
            <LinkIcon className="mr-2 h-5 w-5" />
            {[t('duplicates.downloadReportBtn')].flat().join(' ')}
          </Button>
        </CardFooter>
      )}
    </Card>
  );
}
