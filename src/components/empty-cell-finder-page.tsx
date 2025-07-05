
"use client";

import React, { useState, useCallback, ChangeEvent, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, FileScan, ListChecks, CheckCircle2, Loader2, Palette, List, Columns, ChevronDown, Lightbulb, ScrollText, XCircle, BarChartHorizontal, FilterX, FileSpreadsheet, Pilcrow, Settings } from 'lucide-react';
import { findAndHighlightEmptyCells } from '@/lib/excel-empty-cell-finder';
import { generateEmptyCellFinderVbs } from '@/lib/vbs-generators';
import type { EmptyCellReport } from '@/lib/excel-types';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger, DropdownMenuSeparator } from '@/components/ui/dropdown-menu';
import { useLanguage } from '@/context/language-context';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';

interface SheetSelection {
  [sheetName: string]: boolean;
}

const isValidHex = (hex: string) => /^([0-9A-F]{6})$/i.test(hex);

const getColorNameFromHex = (hex: string, t: (key: string) => string | string[]): string => {
  const translations: Record<string, string> = {
    'FF0000': 'colors.red',
    '00FF00': 'colors.lime',
    '0000FF': 'colors.blue',
    'FFFF00': 'colors.yellow',
    'FFC0CB': 'colors.pink',
    'FFA500': 'colors.orange',
    'ADD8E6': 'colors.lightBlue',
    '90EE90': 'colors.lightGreen',
    'E6E6FA': 'colors.lavender',
    'FFFFFF': 'colors.white',
    'C0C0C0': 'colors.silver',
    '808080': 'colors.gray',
    '000000': 'colors.black',
  };
  const key = translations[hex.toUpperCase()];
  return key ? [t(key)].flat().join(' ') : [t('finder.customColor')].flat().join(' ');
};

interface EmptyCellFinderPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
}

export default function EmptyCellFinderPage({ onProcessingChange }: EmptyCellFinderPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<SheetSelection>({});
  
  const [headerRow, setHeaderRow] = useState<number>(1);
  const [checkMode, setCheckMode] = useState<'all' | 'specific'>('all');
  const [columnsToCheck, setColumnsToCheck] = useState<string>('');
  const [columnsToIgnore, setColumnsToIgnore] = useState<string>('');

  const [enableHighlight, setEnableHighlight] = useState<boolean>(true);
  const [highlightColor, setHighlightColor] = useState<string>('FFFF00'); // Yellow default
  const [generateReportSheet, setGenerateReportSheet] = useState<boolean>(true);
  const [colorName, setColorName] = useState<string>('Yellow');

  const [reportFormat, setReportFormat] = useState<'compact' | 'detailed' | 'summary'>('compact');
  const [columnsToInclude, setColumnsToInclude] = useState<string>('');
  const [includeAllData, setIncludeAllData] = useState<boolean>(false);
  const [contextColumnForCompact, setContextColumnForCompact] = useState<string>('');
  const [summaryKeyColumn, setSummaryKeyColumn] = useState<string>('');
  const [summaryContextColumn, setSummaryContextColumn] = useState<string>('');
  const [blankKeyLabel, setBlankKeyLabel] = useState<string>('(Blanks)');


  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const cancellationRequested = useRef(false);
  const [processedReport, setProcessedReport] = useState<EmptyCellReport | null>(null);
  const [modifiedWorkbook, setModifiedWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [outputFormat, setOutputFormat] = useState<'xlsx' | 'xlsm'>('xlsm');
  const [vbscriptPreview, setVbscriptPreview] = useState<string>('');
  const { toast } = useToast();
  const [reportChunkSize, setReportChunkSize] = useState<number>(100000);

  useEffect(() => {
    if (onProcessingChange) {
      onProcessingChange(isLoading);
    }
  }, [isLoading, onProcessingChange]);

  useEffect(() => {
    // Reset all state when the file changes to ensure a clean slate.
    setSheetNames([]);
    setSelectedSheets({});
    setHeaderRow(1);
    setCheckMode('all');
    setColumnsToCheck('');
    setColumnsToIgnore('');
    setEnableHighlight(true);
    setHighlightColor('FFFF00');
    setGenerateReportSheet(true);
    setReportFormat('compact');
    setColumnsToInclude('');
    setIncludeAllData(false);
    setContextColumnForCompact('');
    setSummaryKeyColumn('');
    setSummaryContextColumn('');
    setProcessedReport(null);
    setModifiedWorkbook(null);
    setBlankKeyLabel('(Blanks)');
    setReportChunkSize(100000);

    if (file) {
      const getSheetNamesFromFile = async () => {
        setIsLoading(true);
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
          setIsLoading(false);
        }
      };
      getSheetNamesFromFile();
    }
  }, [file, toast, t]);

  useEffect(() => {
    const upperHex = highlightColor.toUpperCase();
    if (isValidHex(upperHex)) {
        setColorName(getColorNameFromHex(upperHex, t));
    } else {
        setColorName('');
    }
  }, [highlightColor, t]);

  useEffect(() => {
    const sheetsToUpdate = Object.entries(selectedSheets)
      .filter(([,isSelected]) => isSelected)
      .map(([sheetName]) => sheetName);
      
    const script = generateEmptyCellFinderVbs(
        sheetsToUpdate,
        checkMode === 'all' ? '*' : columnsToCheck,
        columnsToIgnore,
        headerRow,
        enableHighlight ? highlightColor : undefined
    );
    setVbscriptPreview(script);

  }, [selectedSheets, checkMode, columnsToCheck, columnsToIgnore, headerRow, enableHighlight, highlightColor]);


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
  
  const handleCancel = () => {
    cancellationRequested.current = true;
    setProcessingStatus([t('common.cancelling')].flat().join(' '));
  };

  const handleSelectAllSheets = (checked: boolean) => {
    const newSelection: SheetSelection = {};
    sheetNames.forEach(name => {
      newSelection[name] = checked;
    });
    setSelectedSheets(newSelection);
  };

  const handlePartialSelection = (count: number) => {
    const newSelection: SheetSelection = {};
    sheetNames.forEach((name, index) => {
      newSelection[name] = index < count;
    });
    setSelectedSheets(newSelection);
  };
  
  const handleLastPartialSelection = (count: number) => {
    const newSelection: SheetSelection = {};
    const totalSheets = sheetNames.length;
    if (count > totalSheets) count = totalSheets;
    sheetNames.forEach((name, index) => {
      newSelection[name] = index >= totalSheets - count;
    });
    setSelectedSheets(newSelection);
  };

  const handleSheetSelectionChange = (sheetName: string, checked: boolean) => {
    setSelectedSheets(prev => ({ ...prev, [sheetName]: checked }));
  };

  const handleProcess = useCallback(async () => {
    cancellationRequested.current = false;
    const sheetsToProcess = sheetNames.filter(name => selectedSheets[name]);
    if (!file || sheetsToProcess.length === 0 || (checkMode === 'specific' && !columnsToCheck.trim())) {
      toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('finder.toast.missingInfo')].flat().join(' '), variant: 'destructive' });
      return;
    }
    if (enableHighlight && !isValidHex(highlightColor)) {
      toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('finder.toast.invalidColor', { hex: highlightColor })].flat().join(' '), variant: 'destructive' });
      return;
    }
    if (generateReportSheet && reportFormat === 'summary' && (!summaryKeyColumn || !summaryContextColumn)) {
        toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('finder.toast.missingSummaryCols')].flat().join(' '), variant: 'destructive'});
        return;
    }


    setIsLoading(true);
    setProcessedReport(null);
    setModifiedWorkbook(null);
    setProcessingStatus('');
    
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, cellDates: true, bookVBA: true, bookFiles: true });
      
      const columnParam = checkMode === 'all' ? '*' : columnsToCheck;
      
      const reportOptions = {
        format: reportFormat,
        columnsToInclude: columnsToInclude,
        includeAllData: includeAllData,
        contextColumnForCompact: contextColumnForCompact,
        summaryKeyColumn: summaryKeyColumn,
        summaryContextColumn: summaryContextColumn,
        blankKeyLabel: blankKeyLabel.trim() || '(Blanks)',
        chunkSize: reportChunkSize,
      };

      const onProgress = (status: { sheetName: string; currentSheet: number; totalSheets: number; emptyFound: number }) => {
        if (cancellationRequested.current) {
            throw new Error('Cancelled by user.');
        }
        setProcessingStatus(
            [t('finder.toast.processingSheet', { current: status.currentSheet, total: status.totalSheets, sheetName: status.sheetName, count: status.emptyFound })].flat().join(' ')
        );
      };
      
      const { report, workbook: newWb } = findAndHighlightEmptyCells(
        workbook,
        sheetsToProcess,
        columnParam,
        columnsToIgnore,
        headerRow,
        enableHighlight ? highlightColor : undefined,
        generateReportSheet,
        reportOptions,
        onProgress
      );

      setProcessedReport(report);

      if (report.totalEmpty > 0) {
        setModifiedWorkbook(newWb);
      } else {
        setModifiedWorkbook(null);
      }
      
      toast({
        title: [t('toast.processingComplete')].flat().join(' '),
        description: [t('finder.toast.success', { count: report.totalEmpty, sheets: Object.keys(report.summary).length })].flat().join(' '),
        action: <CheckCircle2 className="text-green-500" />,
      });

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : [t('finder.toast.error')].flat().join(' ');
      if (errorMessage !== 'Cancelled by user.') {
        toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: errorMessage, variant: 'destructive' });
      } else {
        toast({ title: [t('toast.cancelledTitle')].flat().join(' '), description: [t('toast.cancelledDesc')].flat().join(' '), variant: 'default' });
      }
    } finally {
      setIsLoading(false);
      cancellationRequested.current = false;
      setProcessingStatus('');
    }
  }, [file, sheetNames, selectedSheets, checkMode, columnsToCheck, columnsToIgnore, headerRow, enableHighlight, highlightColor, generateReportSheet, toast, reportFormat, columnsToInclude, includeAllData, contextColumnForCompact, summaryKeyColumn, summaryContextColumn, t, blankKeyLabel, reportChunkSize]);

  const handleDownloadUpdatedFile = useCallback(() => {
     if (!modifiedWorkbook || !file) {
       toast({ title: [t('toast.noFileToDownload')].flat().join(' '), description: [t('finder.toast.noFile')].flat().join(' '), variant: "destructive" });
       return;
     }
     try {
        const originalFileName = file.name.substring(0, file.name.lastIndexOf('.'));
        XLSX.writeFile(modifiedWorkbook, `${originalFileName}_updated.${outputFormat}`, { compression: true, bookType: outputFormat, cellStyles: true });
        toast({ title: [t('toast.downloadSuccess')].flat().join(' ') });
     } catch (error) {
        toast({ title: [t('toast.downloadError')].flat().join(' '), description: [t('toast.downloadError')].flat().join(' '), variant: 'destructive' });
     }
  }, [modifiedWorkbook, file, toast, t, outputFormat]);

  const allSheetsSelected = sheetNames.length > 0 && sheetNames.every(name => selectedSheets[name]);
  const hasResults = processedReport !== null;

  return (
    <Card className="w-full max-w-lg md:max-w-4xl xl:max-w-6xl shadow-xl relative">
      {isLoading && (
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
          <FileScan className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{[t('finder.title')].flat().join(' ')}</CardTitle>
        </div>
        <CardDescription className="font-body">
          {[t('finder.description')].flat().join(' ')}
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-finder" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{[t('finder.uploadStep')].flat().join(' ')}</span>
          </Label>
          <Input
            id="file-upload-finder"
            type="file"
            accept=".xlsx, .xls, .xlsm"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isLoading}
          />
        </div>

        {sheetNames.length > 0 && (
          <div className="space-y-3">
            <Label className="flex items-center space-x-2 text-sm font-medium mb-2">
              <ListChecks className="h-5 w-5" />
              <span>{[t('finder.selectSheetsStep')].flat().join(' ')}</span>
            </Label>
            <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
              <Checkbox
                id="select-all-sheets-finder"
                checked={allSheetsSelected}
                onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                disabled={isLoading}
              />
              <Label htmlFor="select-all-sheets-finder" className="text-sm font-medium flex-grow">
                {[t('common.selectAll')].flat().join(' ')} ({[t('common.selectedCount', {selected: Object.values(selectedSheets).filter(Boolean).length, total: sheetNames.length})].flat().join(' ')})
              </Label>
              {sheetNames.length > 50 && (
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <Button variant="outline" size="sm" disabled={isLoading}>
                        {[t('common.partial')].flat().join(' ')}
                        <ChevronDown className="ml-1 h-4 w-4" />
                    </Button>
                  </DropdownMenuTrigger>
                  <DropdownMenuContent>
                    <DropdownMenuItem onSelect={() => handlePartialSelection(50)}>{[t('common.first50')].flat().join(' ')}</DropdownMenuItem>
                    {sheetNames.length >= 100 && <DropdownMenuItem onSelect={() => handlePartialSelection(100)}>{[t('common.first100')].flat().join(' ')}</DropdownMenuItem>}
                    {sheetNames.length >= 150 && <DropdownMenuItem onSelect={() => handlePartialSelection(150)}>{[t('common.first150')].flat().join(' ')}</DropdownMenuItem>}
                    <DropdownMenuSeparator />
                    <DropdownMenuItem onSelect={() => handleLastPartialSelection(50)}>{[t('common.last50')].flat().join(' ')}</DropdownMenuItem>
                    {sheetNames.length >= 100 && <DropdownMenuItem onSelect={() => handleLastPartialSelection(100)}>{[t('common.last100')].flat().join(' ')}</DropdownMenuItem>}
                    {sheetNames.length >= 150 && <DropdownMenuItem onSelect={() => handleLastPartialSelection(150)}>{[t('common.last150')].flat().join(' ')}</DropdownMenuItem>}
                  </DropdownMenuContent>
                </DropdownMenu>
              )}
            </div>
            <Card className="max-h-48 overflow-y-auto p-3 bg-background">
              <div className="space-y-2">
                {sheetNames.map(name => (
                  <div key={name} className="flex items-center space-x-2">
                    <Checkbox
                      id={`sheet-find-${name}`}
                      checked={selectedSheets[name] || false}
                      onCheckedChange={(checked) => handleSheetSelectionChange(name, checked as boolean)}
                      disabled={isLoading}
                    />
                    <Label htmlFor={`sheet-find-${name}`} className="text-sm font-normal">{name}</Label>
                  </div>
                ))}
              </div>
            </Card>
          </div>
        )}

        <div className="space-y-2">
            <Label htmlFor="header-row" className="flex items-center space-x-2 text-sm font-medium">
                <FileSpreadsheet className="h-5 w-5" />
                <span>{[t('finder.headerRowStep')].flat().join(' ')}</span>
            </Label>
            <Input 
                id="header-row" 
                type="number" 
                min="1" 
                value={headerRow} 
                onChange={(e) => setHeaderRow(parseInt(e.target.value, 10) || 1)} 
                disabled={isLoading || !file}
            />
            <p className="text-xs text-muted-foreground">{[t('finder.headerRowDesc')].flat().join(' ')}</p>
        </div>

        <div className="space-y-4">
            <div className="space-y-2">
              <Label htmlFor="column-to-check" className="flex items-center space-x-2 text-sm font-medium">
                <FileScan className="h-5 w-5" />
                <span>{[t('finder.columnToCheckStep')].flat().join(' ')}</span>
              </Label>
              <RadioGroup value={checkMode} onValueChange={(v) => setCheckMode(v as any)} className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <Label className="flex items-center space-x-2 p-2 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer">
                    <RadioGroupItem value="all" id="mode-all-cols" />
                    <span>{[t('finder.checkAllCols')].flat().join(' ')}</span>
                  </Label>
                   <Label className="flex items-center space-x-2 p-2 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer">
                    <RadioGroupItem value="specific" id="mode-specific-cols" />
                    <span>{[t('finder.checkSpecificCols')].flat().join(' ')}</span>
                  </Label>
              </RadioGroup>
              
              {checkMode === 'specific' && (
                <div className="pt-2">
                    <Input
                        id="columns-to-check-input"
                        value={columnsToCheck}
                        onChange={(e) => setColumnsToCheck(e.target.value)}
                        disabled={isLoading || !file}
                        placeholder={[t('finder.columnInputPlaceholder')].flat().join(' ')}
                    />
                    <p className="text-xs text-muted-foreground pt-1">{[t('finder.columnInputDesc')].flat().join(' ')}</p>
                </div>
              )}
            </div>
            <div className="space-y-2">
              <Label htmlFor="columns-to-ignore" className="flex items-center space-x-2 text-sm font-medium">
                  <FilterX className="h-5 w-5" />
                  <span>{[t('finder.ignoreColsStep')].flat().join(' ')}</span>
              </Label>
              <Input
                  id="columns-to-ignore"
                  value={columnsToIgnore}
                  onChange={(e) => setColumnsToIgnore(e.target.value)}
                  disabled={isLoading || !file}
                  placeholder={[t('finder.ignoreColsPlaceholder')].flat().join(' ')}
              />
              <p className="text-xs text-muted-foreground">{[t('finder.ignoreColsDesc')].flat().join(' ')}</p>
            </div>
        </div>

        <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
          <CardHeader className="p-0 pb-4">
            <Label className="flex items-center space-x-2 text-md font-semibold text-primary">
              <Palette className="h-5 w-5" />
              <span>{[t('finder.optionsStep')].flat().join(' ')}</span>
            </Label>
          </CardHeader>
          <CardContent className="p-0 space-y-4">
            <div className="flex items-start space-x-3">
              <Checkbox
                id="enable-highlight"
                checked={enableHighlight}
                onCheckedChange={(checked) => setEnableHighlight(checked as boolean)}
                disabled={isLoading}
                className="mt-1"
              />
              <div className="grid gap-1.5 leading-none w-full">
                <Label htmlFor="enable-highlight">{[t('finder.highlightOption')].flat().join(' ')}</Label>
                <p className="text-xs text-muted-foreground">{[t('finder.highlightOptionDesc')].flat().join(' ')}</p>
                {enableHighlight && (
                  <div className="pt-2 space-y-2">
                    <Label htmlFor="highlight-color" className="text-sm font-medium">{[t('finder.highlightColor')].flat().join(' ')}</Label>
                    <div className="flex items-center gap-2">
                        <div className="flex items-center border rounded-md bg-background focus-within:ring-2 focus-within:ring-ring">
                           <span className="pl-3 text-muted-foreground">#</span>
                           <Input
                             id="highlight-color"
                             value={highlightColor}
                             onChange={(e) => setHighlightColor(e.target.value.toUpperCase().replace(/[^0-9A-F]/g, ''))}
                             disabled={isLoading}
                             placeholder="FFFF00"
                             maxLength={6}
                             className="font-mono w-28 border-0 shadow-none focus-visible:ring-0"
                           />
                        </div>
                        <div 
                            className="h-8 w-8 rounded-md border shrink-0" 
                            style={{ 
                                backgroundColor: isValidHex(highlightColor) ? `#${highlightColor}` : 'transparent',
                            }}
                        ></div>
                    </div>
                    {colorName && isValidHex(highlightColor) && <p className="text-xs text-muted-foreground">{[t('finder.colorPreview', { colorName })].flat().join(' ')}</p>}
                  </div>
                )}
              </div>
            </div>
             <div className="flex items-start space-x-3 pt-4">
              <Checkbox
                id="generate-report"
                checked={generateReportSheet}
                onCheckedChange={(checked) => setGenerateReportSheet(checked as boolean)}
                disabled={isLoading}
                className="mt-1"
              />
              <div className="grid gap-1.5 leading-none w-full">
                <Label htmlFor="generate-report">{[t('finder.reportOption')].flat().join(' ')}</Label>
                <p className="text-xs text-muted-foreground">{[t('finder.reportOptionDesc')].flat().join(' ')}</p>
                {generateReportSheet && (
                  <Card className="p-4 mt-2 space-y-4 bg-background/50">
                    <Label>{[t('finder.reportFormat')].flat().join(' ')}</Label>
                    <RadioGroup value={reportFormat} onValueChange={(v) => setReportFormat(v as any)}>
                      <div className="flex items-start space-x-2">
                        <RadioGroupItem value="compact" id="compact-report" className="mt-1"/>
                        <div className="grid gap-1.5">
                          <Label htmlFor="compact-report" className="font-normal">{[t('finder.compactReport')].flat().join(' ')}</Label>
                          <p className="text-xs text-muted-foreground">{[t('finder.compactReportDesc')].flat().join(' ')}</p>
                          {reportFormat === 'compact' && (
                            <div className="pt-2 space-y-2">
                              <Label htmlFor="context-column-compact" className="text-sm font-medium">{[t('finder.compactContextCol')].flat().join(' ')}</Label>
                              <Input
                                id="context-column-compact"
                                value={contextColumnForCompact}
                                onChange={e => setContextColumnForCompact(e.target.value)}
                                disabled={isLoading}
                                placeholder={[t('finder.compactContextColPlaceholder')].flat().join(' ')}
                              />
                              <p className="text-xs text-muted-foreground">{[t('finder.compactContextColDesc')].flat().join(' ')}</p>
                            </div>
                          )}
                        </div>
                      </div>
                      <div className="flex items-start space-x-2">
                        <RadioGroupItem value="detailed" id="detailed-report" className="mt-1"/>
                        <div className="grid gap-1.5">
                          <Label htmlFor="detailed-report" className="font-normal">{[t('finder.detailedReport')].flat().join(' ')}</Label>
                          <p className="text-xs text-muted-foreground">{[t('finder.detailedReportDesc')].flat().join(' ')}</p>
                        </div>
                      </div>
                      <div className="flex items-start space-x-2">
                        <RadioGroupItem value="summary" id="summary-report" className="mt-1"/>
                        <div className="grid gap-1.5">
                          <Label htmlFor="summary-report" className="font-normal">{[t('finder.summaryReport')].flat().join(' ')}</Label>
                          <p className="text-xs text-muted-foreground">{[t('finder.summaryReportDesc')].flat().join(' ')}</p>
                           {reportFormat === 'summary' && (
                            <div className="pt-2 pl-6 space-y-4">
                              <div className="space-y-2">
                                <Label htmlFor="summary-key-column" className="flex items-center space-x-2 text-sm font-medium"><BarChartHorizontal className="h-4 w-4" /><span>{[t('finder.summaryKeyCol')].flat().join(' ')}</span></Label>
                                <Input
                                  id="summary-key-column"
                                  value={summaryKeyColumn}
                                  onChange={e => setSummaryKeyColumn(e.target.value)}
                                  disabled={isLoading}
                                  placeholder={[t('finder.summaryKeyColPlaceholder')].flat().join(' ')}
                                />
                                <p className="text-xs text-muted-foreground">{[t('finder.summaryKeyColDesc')].flat().join(' ')}</p>
                              </div>
                               <div className="space-y-2">
                                <Label htmlFor="summary-context-column" className="flex items-center space-x-2 text-sm font-medium"><Columns className="h-4 w-4" /><span>{[t('finder.summaryContextCol')].flat().join(' ')}</span></Label>
                                <Input
                                  id="summary-context-column"
                                  value={summaryContextColumn}
                                  onChange={e => setSummaryContextColumn(e.target.value)}
                                  disabled={isLoading}
                                  placeholder={[t('finder.summaryContextColPlaceholder')].flat().join(' ')}
                                />
                                <p className="text-xs text-muted-foreground">{[t('finder.summaryContextColDesc')].flat().join(' ')}</p>
                              </div>
                              <div className="space-y-2">
                                <Label htmlFor="blank-key-label" className="flex items-center space-x-2 text-sm font-medium">
                                  <Pilcrow className="h-4 w-4" />
                                  <span>{[t('finder.blankKeyLabel')].flat().join(' ')}</span>
                                </Label>
                                <Input
                                  id="blank-key-label"
                                  value={blankKeyLabel}
                                  onChange={e => setBlankKeyLabel(e.target.value)}
                                  disabled={isLoading}
                                  placeholder={[t('finder.blankKeyLabelPlaceholder')].flat().join(' ')}
                                />
                                <p className="text-xs text-muted-foreground">{[t('finder.blankKeyLabelDesc')].flat().join(' ')}</p>
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    </RadioGroup>

                    {reportFormat === 'detailed' && (
                      <div className="pt-2 pl-6 space-y-4">
                        <Label className="flex items-center space-x-2 text-sm font-medium">
                          <Columns className="h-5 w-5" />
                          <span>{[t('finder.detailedColsToInclude')].flat().join(' ')}</span>
                        </Label>
                        <div className="flex items-center space-x-2">
                            <Checkbox id="include-all-data" checked={includeAllData} onCheckedChange={c => setIncludeAllData(c as boolean)} />
                            <Label htmlFor="include-all-data" className="font-normal">{[t('finder.detailedColsToIncludeAll')].flat().join(' ')}</Label>
                        </div>
                        <Input 
                          id="columns-to-include"
                          value={columnsToInclude}
                          onChange={e => setColumnsToInclude(e.target.value)}
                          disabled={isLoading || includeAllData}
                          placeholder={[t('finder.detailedColsToIncludePlaceholder')].flat().join(' ')}
                        />
                        <p className="text-xs text-muted-foreground">{[t('finder.detailedColsToIncludeDesc')].flat().join(' ')}</p>
                      </div>
                    )}
                  </Card>
                )}
              </div>
            </div>
          </CardContent>
        </Card>
        
        <Accordion type="single" collapsible>
            <AccordionItem value="advanced-settings">
                <AccordionTrigger className="text-md font-semibold">{[t('aggregator.advancedSettings.title')].flat().join(' ')}</AccordionTrigger>
                <AccordionContent>
                    <Card className="p-4 border-dashed">
                       <div className="space-y-2">
                            <Label htmlFor="report-chunk-size" className="text-sm font-medium">{[t('aggregator.advancedSettings.maxRows')].flat().join(' ')}</Label>
                            <Input
                                id="report-chunk-size"
                                type="number"
                                min="1000"
                                step="1000"
                                value={reportChunkSize}
                                onChange={(e) => setReportChunkSize(parseInt(e.target.value, 10) || 100000)}
                                disabled={isLoading}
                            />
                            <p className="text-xs text-muted-foreground">
                                {[t('aggregator.advancedSettings.maxRowsDesc')].flat().join(' ')}
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
              {[t('updater.vbsPreviewDesc')].flat().join(' ')}
            </p>
          </div>
        )}

        <Button onClick={handleProcess} disabled={isLoading || !file} className="w-full">
          {isLoading && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
          <FileScan className="mr-2 h-5 w-5" />
          {[t('finder.processBtn')].flat().join(' ')}
        </Button>
      </CardContent>

      {hasResults && processedReport.totalEmpty > 0 && (
        <CardFooter className="flex-col space-y-4 items-stretch">
          <div className="p-4 border rounded-md bg-secondary/30">
            <h3 className="text-lg font-semibold mb-2 font-headline">{[t('finder.resultsTitle')].flat().join(' ')}</h3>
            <p>{[t('finder.resultsFound', { count: processedReport.totalEmpty })].flat().join(' ')}</p>
            <ul className="text-sm mt-2 max-h-32 overflow-y-auto">
              {Object.entries(processedReport.summary).map(([sheetName, count]) => (
                <li key={sheetName} className="flex justify-between">
                  <span>{sheetName}:</span>
                  <span className="font-medium">{count}</span>
                </li>
              ))}
            </ul>
          </div>
          
          {modifiedWorkbook && (
            <>
                <div className="w-full p-4 border rounded-md bg-secondary/30 space-y-4">
                    <Label className="text-md font-semibold font-headline">{[t('common.outputOptions.title')].flat().join(' ')}</Label>
                    <RadioGroup value={outputFormat} onValueChange={(v) => setOutputFormat(v as any)} className="space-y-3">
                        <div>
                            <div className="flex items-center space-x-2">
                                <RadioGroupItem value="xlsx" id="format-xlsx-finder" />
                                <Label htmlFor="format-xlsx-finder" className="font-normal">{[t('common.outputOptions.xlsx')].flat().join(' ')}</Label>
                            </div>
                            <p className="text-xs text-muted-foreground pl-6 pt-1">{[t('common.outputOptions.xlsxDesc')].flat().join(' ')}</p>
                        </div>
                        <div>
                            <div className="flex items-center space-x-2">
                                <RadioGroupItem value="xlsm" id="format-xlsm-finder" />
                                <Label htmlFor="format-xlsm-finder" className="font-normal">{[t('common.outputOptions.xlsm')].flat().join(' ')}</Label>
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
                <Button onClick={handleDownloadUpdatedFile} variant="outline" className="w-full">
                  <Download className="mr-2 h-5 w-5" />
                  {[t('finder.downloadBtn')].flat().join(' ')}
                </Button>
            </>
          )}
        </CardFooter>
      )}
    </Card>
  );
}
