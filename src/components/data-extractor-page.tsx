
"use client";

import React, { useState, useCallback, ChangeEvent, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, FileSearch, Loader2, CheckCircle2, FileSpreadsheet, ListFilter, XCircle, Settings, Columns, Search, FileOutput, ListChecks, ChevronDown, FileText, Scan } from 'lucide-react';
import { useLanguage } from '@/context/language-context';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';
import type { ExtractionReport } from '@/lib/excel-types';
import { findAndExtractData, createExtractionReportWorkbook } from '@/lib/excel-data-extractor';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { RadioGroup, RadioGroupItem } from './ui/radio-group';
import { Alert, AlertDescription } from './ui/alert';
import { Lightbulb } from 'lucide-react';
import { Checkbox } from './ui/checkbox';
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from './ui/dropdown-menu';
import { getColumnIndex } from '@/lib/excel-helpers';
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from './ui/tooltip';

interface DataExtractorPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function DataExtractorPage({ onProcessingChange, onFileStateChange }: DataExtractorPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  
  const [searchScope, setSearchScope] = useState<'single' | 'multiple'>('single');
  const [singleSelectedSheet, setSingleSelectedSheet] = useState<string>('');
  const [multipleSelectedSheets, setMultipleSelectedSheets] = useState<Record<string, boolean>>({});

  const [headerRow, setHeaderRow] = useState<number>(1);
  const [lookupColumn, setLookupColumn] = useState<string>('');
  const [lookupValue, setLookupValue] = useState<string>('');
  const [returnColumns, setReturnColumns] = useState<string>('');

  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const cancellationRequested = useRef(false);

  const [extractionReport, setExtractionReport] = useState<ExtractionReport | null>(null);
  const [outputFormat, setOutputFormat] = useState<'xlsx' | 'xlsm'>('xlsm');
  const [reportChunkSize, setReportChunkSize] = useState<number>(100000);
  
  const [uniqueLookupValues, setUniqueLookupValues] = useState<string[]>([]);
  const [isScanningValues, setIsScanningValues] = useState<boolean>(false);
  const { toast } = useToast();

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
    // Reset all state when the file changes.
    setSheetNames([]);
    setSearchScope('single');
    setSingleSelectedSheet('');
    setMultipleSelectedSheets({});
    setHeaderRow(1);
    setLookupColumn('');
    setLookupValue('');
    setReturnColumns('');
    setExtractionReport(null);
    setReportChunkSize(100000);
    setUniqueLookupValues([]);
    setIsScanningValues(false);

    if (file) {
      const getSheetNamesFromFile = async () => {
        setIsProcessing(true);
        try {
          const arrayBuffer = await file.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
          const names = workbook.SheetNames;
          setSheetNames(names);
          if (names.length > 0) {
            setSingleSelectedSheet(names[0]);
          }
           const initialSelection: Record<string, boolean> = {};
          names.forEach(name => {
            initialSelection[name] = true; // Select all by default for multi-select
          });
          setMultipleSelectedSheets(initialSelection);
        } catch (error) {
          toast({ title: t('toast.errorReadingFile') as string, description: t('toast.errorReadingSheets') as string, variant: "destructive" });
        } finally {
          setIsProcessing(false);
        }
      };
      getSheetNamesFromFile();
    }
  }, [file, toast, t]);

  useEffect(() => {
    setUniqueLookupValues([]);
  }, [lookupColumn, searchScope, singleSelectedSheet, multipleSelectedSheets]);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      if (!selectedFile.name.match(/\.(xlsx|xls|xlsm)$/)) {
        toast({ title: t('toast.invalidFileType') as string, description: t('toast.invalidFileTypeDesc') as string, variant: 'destructive' });
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
  
  const handleCancel = () => {
    cancellationRequested.current = true;
    setProcessingStatus(t('common.cancelling') as string);
  };
  
  const handleSelectAllSheets = (checked: boolean) => {
    const newSelection: Record<string, boolean> = {};
    sheetNames.forEach(name => {
      newSelection[name] = checked;
    });
    setMultipleSelectedSheets(newSelection);
  };

  const handlePartialSelection = (count: number) => {
    const newSelection: Record<string, boolean> = {};
    sheetNames.forEach((name, index) => {
      newSelection[name] = index < count;
    });
    setMultipleSelectedSheets(newSelection);
  };

  const handleSheetSelectionChange = (sheetName: string, checked: boolean) => {
    setMultipleSelectedSheets(prev => ({ ...prev, [sheetName]: checked }));
  };

  const handleScanValues = useCallback(async () => {
    let sheetsToProcess: string[] = [];
    if (searchScope === 'single') {
        if (singleSelectedSheet) sheetsToProcess = [singleSelectedSheet];
    } else {
        sheetsToProcess = Object.keys(multipleSelectedSheets).filter(name => multipleSelectedSheets[name]);
    }
    
    if (!file || sheetsToProcess.length === 0 || !lookupColumn.trim()) {
      toast({ title: t('toast.missingInfo') as string, description: t('extractor.toast.scanError') as string, variant: 'destructive' });
      return;
    }

    setIsScanningValues(true);
    setUniqueLookupValues([]);
    
    try {
        const values = new Set<string>();
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });

        for (const sheetName of sheetsToProcess) {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) continue;
            
            const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
            const headerRowIndex = headerRow - 1;
            if (aoa.length <= headerRowIndex) continue;
            
            const headers = aoa[headerRowIndex].map(h => String(h || ''));
            const colIdx = getColumnIndex(lookupColumn, headers);

            if (colIdx === null) {
                throw new Error(`Column "${lookupColumn}" not found on sheet "${sheetName}".`);
            }

            for (let R = headerRow; R < aoa.length; R++) {
                const row = aoa[R];
                if (row && row[colIdx] !== null && row[colIdx] !== undefined) {
                    const cellValue = String(row[colIdx]).trim();
                    if(cellValue) {
                        values.add(cellValue);
                    }
                }
            }
        }
        
        const sortedValues = Array.from(values).sort((a,b) => a.localeCompare(b));
        setUniqueLookupValues(sortedValues);

        if (sortedValues.length > 0) {
            toast({ title: t('toast.processingComplete') as string, description: t('extractor.toast.scanSuccess', { count: sortedValues.length }) as string });
        } else {
            toast({ title: t('toast.processingComplete') as string, description: t('extractor.toast.scanNoValues') as string });
        }
    } catch (error) {
         const message = error instanceof Error ? error.message : String(error);
         toast({ title: t('toast.errorReadingFile') as string, description: message, variant: 'destructive' });
    } finally {
        setIsScanningValues(false);
    }
  }, [file, searchScope, singleSelectedSheet, multipleSelectedSheets, lookupColumn, headerRow, toast, t]);


  const handleProcess = useCallback(async () => {
    let sheetsToProcess: string[] = [];
    if (searchScope === 'single') {
        if (singleSelectedSheet) sheetsToProcess = [singleSelectedSheet];
    } else {
        sheetsToProcess = Object.keys(multipleSelectedSheets).filter(name => multipleSelectedSheets[name]);
    }
    
    if (!file || sheetsToProcess.length === 0 || !lookupColumn.trim() || !lookupValue.trim() || !returnColumns.trim() || headerRow < 1) {
      toast({ title: t('toast.missingInfo') as string, description: t('extractor.toast.missingInfo') as string, variant: 'destructive' });
      return;
    }

    cancellationRequested.current = false;
    setIsProcessing(true);
    setProcessingStatus('');
    setExtractionReport(null);
    
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, bookVBA: true, bookFiles: true });
      
      const onProgress = (status: { sheetName: string; rowsFound: number; currentSheet: number; totalSheets: number; }) => {
        if (cancellationRequested.current) throw new Error('Cancelled by user.');
        setProcessingStatus(
          t('extractor.toast.processing', {
            current: status.currentSheet,
            total: status.totalSheets,
            sheetName: status.sheetName,
            count: status.rowsFound
          }) as string
        );
      };

      const report = findAndExtractData(
        workbook,
        sheetsToProcess,
        {
            lookupColumn,
            lookupValue,
            returnColumns,
            headerRow,
        },
        onProgress
      );
      
      setExtractionReport(report);
      
      toast({
        title: t('toast.processingComplete') as string,
        description: t('extractor.toast.success', { count: report.summary.totalRowsFound }) as string,
        action: <CheckCircle2 className="text-green-500" />,
      });

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : t('extractor.toast.error') as string;
      if (errorMessage !== 'Cancelled by user.') {
        toast({ title: t('toast.errorReadingFile') as string, description: errorMessage, variant: 'destructive' });
      } else {
        toast({ title: t('toast.cancelledTitle') as string, description: t('toast.cancelledDesc') as string, variant: 'default' });
      }
    } finally {
      setIsProcessing(false);
      cancellationRequested.current = false;
      setProcessingStatus('');
    }
  }, [file, searchScope, singleSelectedSheet, multipleSelectedSheets, headerRow, lookupColumn, lookupValue, returnColumns, toast, t]);

  const handleDownloadReport = useCallback(() => {
    if (!extractionReport || !file) {
      toast({ title: t('toast.noDataToDownload') as string, description: t('extractor.toast.noReport') as string, variant: "destructive" });
      return;
    }
    try {
        const reportWb = createExtractionReportWorkbook(extractionReport, { reportChunkSize });
        const originalFileName = file.name.substring(0, file.name.lastIndexOf('.')) + '_extracted_data';
        XLSX.writeFile(reportWb, `${originalFileName}.${outputFormat}`, { compression: true, bookType: outputFormat, cellStyles: true });
        toast({ title: t('toast.downloadSuccess') as string });
    } catch (error) {
        toast({ title: t('toast.downloadError') as string, description: t('extractor.toast.downloadReportError') as string, variant: 'destructive' });
    }
  }, [extractionReport, file, toast, t, outputFormat, reportChunkSize]);
  
  const handleDownloadPdf = useCallback(async () => {
    if (!extractionReport || !file) {
      toast({ title: t('toast.noDataToDownload') as string, description: t('extractor.toast.noReport') as string, variant: "destructive" });
      return;
    }
    
    const { default: jsPDF } = await import('jspdf');
    const { default: autoTable } = await import('jspdf-autotable');

    const { details, summary } = extractionReport;
    if (details.length === 0) {
        toast({ title: t('toast.noDataToDownload') as string, description: "No matching rows found to generate a PDF.", variant: 'default' });
        return;
    }
    
    // Consistent header logic with Excel report
    const baseHeaders = Object.keys(details[0]).filter(h => h !== "Source Sheet");
    const headers = summary.sheetsSearched.length > 1 
        ? ["Source Sheet", ...baseHeaders] 
        : baseHeaders;
    
    // Manually format the body into an array of arrays of strings
    const body = details.map(rowObject => {
        return headers.map(header => {
            const value = rowObject[header];
            if (value === null || value === undefined) {
                return '';
            }
            return String(value);
        });
    });

    const orientation = headers.length > 7 ? 'landscape' : 'portrait';
    const doc = new jsPDF({ orientation });

    const title = `Data Extraction Report for: ${file.name}`;
    doc.setFontSize(16);
    doc.text(title, 14, 22);

    autoTable(doc, {
        startY: 30,
        head: [headers],
        body: body,
        theme: 'striped',
        headStyles: { fillColor: [75, 85, 99] },
        styles: {
            fontSize: headers.length > 10 ? 7 : (headers.length > 7 ? 8 : 10),
            cellPadding: 2,
            overflow: 'linebreak'
        },
        didDrawPage: (data) => {
            const pageNum = data.pageNumber;
            doc.setFontSize(10);
            doc.text(`Page ${pageNum}`, data.settings.margin.left, doc.internal.pageSize.height - 10);
        }
    });

    const originalFileName = file.name.substring(0, file.name.lastIndexOf('.')) + '_extracted_data.pdf';
    doc.save(originalFileName);
    toast({ title: t('toast.downloadSuccess') as string, description: t('extractor.toast.pdfSuccess') as string });

  }, [extractionReport, file, toast, t]);

  const hasResults = extractionReport !== null;
  const previewData = extractionReport?.details.slice(0, 10);
  const allSheetsSelected = sheetNames.length > 0 && sheetNames.every(name => multipleSelectedSheets[name]);

  return (
    <Card className="w-full max-w-lg md:max-w-4xl xl:max-w-6xl shadow-xl relative">
      {isProcessing && (
        <div className="absolute inset-0 bg-background/80 backdrop-blur-sm flex flex-col items-center justify-center z-10 rounded-lg space-y-4">
            <div className="flex items-center gap-2 text-muted-foreground">
            <Loader2 className="h-6 w-6 animate-spin" />
            <span className="text-lg font-medium">{processingStatus || t('common.processing')}</span>
            </div>
            <Button variant="destructive" onClick={handleCancel}>
                <XCircle className="mr-2 h-4 w-4"/>
                {t('common.cancel')}
            </Button>
        </div>
      )}
      <CardHeader>
        <div className="flex items-center space-x-2 mb-2">
          <FileSearch className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{t('extractor.title')}</CardTitle>
        </div>
        <CardDescription className="font-body">{t('extractor.description')}</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-extractor" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{t('extractor.uploadStep')}</span>
          </Label>
          <Input
            id="file-upload-extractor"
            type="file"
            accept=".xlsx, .xls, .xlsm"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isProcessing}
          />
        </div>

        {sheetNames.length > 0 && (
          <div className="space-y-4">
             <div className="space-y-2">
                <Label className="flex items-center space-x-2 text-sm font-medium">{t('extractor.searchScope')}</Label>
                <RadioGroup value={searchScope} onValueChange={(v) => setSearchScope(v as any)} className="grid grid-cols-2 gap-2">
                    <Label htmlFor="scope-single" className="p-2 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer flex items-center justify-center gap-2">
                        <RadioGroupItem value="single" id="scope-single" />
                        {t('extractor.singleSheet')}
                    </Label>
                    <Label htmlFor="scope-multiple" className="p-2 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer flex items-center justify-center gap-2">
                        <RadioGroupItem value="multiple" id="scope-multiple" />
                        {t('extractor.multipleSheets')}
                    </Label>
                </RadioGroup>
            </div>
            
            {searchScope === 'single' ? (
                <div className="space-y-2">
                    <Label htmlFor="sheet-select-extractor" className="flex items-center space-x-2 text-sm font-medium">
                    <ListFilter className="h-5 w-5" />
                    <span>{t('extractor.selectSheetStep')}</span>
                    </Label>
                    <Select value={singleSelectedSheet} onValueChange={setSingleSelectedSheet} disabled={isProcessing || sheetNames.length === 0}>
                    <SelectTrigger id="sheet-select-extractor">
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
            ) : (
                <div className="space-y-3">
                    <Label className="flex items-center space-x-2 text-sm font-medium mb-2">
                        <ListChecks className="h-5 w-5" />
                        <span>{t('duplicates.selectSheetsStep')}</span>
                    </Label>
                    <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
                        <Checkbox
                            id="select-all-sheets-extractor"
                            checked={allSheetsSelected}
                            onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                            disabled={isProcessing}
                        />
                        <Label htmlFor="select-all-sheets-extractor" className="text-sm font-medium flex-grow">
                            {t('common.selectAll')} ({t('common.selectedCount', {selected: Object.values(multipleSelectedSheets).filter(Boolean).length, total: sheetNames.length})})
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
                            </DropdownMenuContent>
                            </DropdownMenu>
                        )}
                    </div>
                     <Card className="max-h-48 overflow-y-auto p-3 bg-background">
                        <div className="space-y-2">
                            {sheetNames.map(name => (
                            <div key={name} className="flex items-center space-x-2">
                                <Checkbox
                                id={`sheet-extractor-${name}`}
                                checked={multipleSelectedSheets[name] || false}
                                onCheckedChange={(checked) => handleSheetSelectionChange(name, checked as boolean)}
                                disabled={isProcessing}
                                />
                                <Label htmlFor={`sheet-extractor-${name}`} className="text-sm font-normal">{name}</Label>
                            </div>
                            ))}
                        </div>
                    </Card>
                </div>
            )}
          </div>
        )}

        <div className="space-y-2">
            <Label htmlFor="header-row-extractor" className="flex items-center space-x-2 text-sm font-medium">
                <FileSpreadsheet className="h-5 w-5" />
                <span>{t('extractor.headerRowStep')}</span>
            </Label>
            <Input 
                id="header-row-extractor" 
                type="number" 
                min="1" 
                value={headerRow} 
                onChange={(e) => setHeaderRow(parseInt(e.target.value, 10) || 1)} 
                disabled={isProcessing || !file}
            />
            <p className="text-xs text-muted-foreground">{t('extractor.headerRowDesc')}</p>
        </div>
        
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-2">
                <Label htmlFor="lookup-column" className="flex items-center space-x-2 text-sm font-medium">
                    <Search className="h-5 w-5" />
                    <span>{t('extractor.lookupColStep')}</span>
                </Label>
                 <div className="flex items-center gap-2">
                    <Input 
                        id="lookup-column" 
                        value={lookupColumn} 
                        onChange={e => setLookupColumn(e.target.value)} 
                        disabled={isProcessing || !file}
                        placeholder={t('extractor.lookupColPlaceholder') as string}
                        className="flex-grow"
                    />
                    <TooltipProvider>
                        <Tooltip>
                            <TooltipTrigger asChild>
                                <Button onClick={handleScanValues} disabled={isScanningValues || !lookupColumn.trim() || !file} variant="outline" size="icon">
                                    {isScanningValues ? <Loader2 className="h-4 w-4 animate-spin" /> : <Scan className="h-4 w-4" />}
                                    <span className="sr-only">{t('extractor.scanValuesBtn')}</span>
                                </Button>
                            </TooltipTrigger>
                            <TooltipContent>
                                <p>{t('extractor.scanForValues')}</p>
                            </TooltipContent>
                        </Tooltip>
                    </TooltipProvider>
                </div>
                <p className="text-xs text-muted-foreground">{t('extractor.lookupColDesc')}</p>
            </div>
             <div className="space-y-2">
                <Label htmlFor="lookup-value" className="flex items-center space-x-2 text-sm font-medium">
                    <Search className="h-5 w-5" />
                    <span>{t('extractor.lookupValStep')}</span>
                </Label>
                <Input 
                    id="lookup-value" 
                    value={lookupValue} 
                    onChange={e => setLookupValue(e.target.value)} 
                    disabled={isProcessing || !file}
                    placeholder={t('extractor.lookupValPlaceholder') as string}
                />
                <p className="text-xs text-muted-foreground">{t('extractor.lookupValDesc')}</p>
                 {uniqueLookupValues.length > 0 && (
                    <div className="pt-2 space-y-2">
                        <Label htmlFor="discovered-values-select">{t('extractor.discoveredValues')}</Label>
                        <Select onValueChange={value => setLookupValue(value)} value={uniqueLookupValues.includes(lookupValue) ? lookupValue : ''}>
                        <SelectTrigger id="discovered-values-select">
                            <SelectValue placeholder={t('extractor.discoveredValuesPlaceholder') as string} />
                        </SelectTrigger>
                        <SelectContent>
                            {uniqueLookupValues.map(val => (
                            <SelectItem key={val} value={val}>{val.length > 100 ? `${val.substring(0,100)}...` : val}</SelectItem>
                            ))}
                        </SelectContent>
                        </Select>
                    </div>
                )}
            </div>
        </div>
        
        <div className="space-y-2">
            <Label htmlFor="return-columns" className="flex items-center space-x-2 text-sm font-medium">
                <Columns className="h-5 w-5" />
                <span>{t('extractor.returnColsStep')}</span>
            </Label>
            <Input 
                id="return-columns" 
                value={returnColumns} 
                onChange={e => setReturnColumns(e.target.value)} 
                disabled={isProcessing || !file}
                placeholder={t('extractor.returnColsPlaceholder') as string}
            />
            <p className="text-xs text-muted-foreground">{t('extractor.returnColsDesc')}</p>
        </div>

        <Accordion type="single" collapsible>
            <AccordionItem value="advanced-settings">
                <AccordionTrigger className="text-md font-semibold">{t('extractor.advancedSettings.title')}</AccordionTrigger>
                <AccordionContent>
                    <Card className="p-4 border-dashed">
                       <div className="space-y-2">
                            <Label htmlFor="report-chunk-size" className="text-sm font-medium">{t('extractor.advancedSettings.maxRows')}</Label>
                            <Input
                                id="report-chunk-size"
                                type="number"
                                min="1000"
                                step="1000"
                                value={reportChunkSize}
                                onChange={(e) => setReportChunkSize(parseInt(e.target.value, 10) || 100000)}
                                disabled={isProcessing}
                            />
                            <p className="text-xs text-muted-foreground">{t('extractor.advancedSettings.maxRowsDesc')}</p>
                        </div>
                    </Card>
                </AccordionContent>
            </AccordionItem>
        </Accordion>

        <Button onClick={handleProcess} disabled={isProcessing || !file} className="w-full">
          {isProcessing ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <FileSearch className="mr-2 h-5 w-5" />}
          {t('extractor.processBtn')}
        </Button>
      </CardContent>

      {hasResults && (
        <CardFooter className="flex-col space-y-4 items-stretch">
          <div className="p-4 border rounded-md bg-secondary/30">
            <h3 className="text-lg font-semibold mb-2 font-headline">{t('extractor.resultsTitle')}</h3>
            <p>{t('extractor.resultsFound', { count: extractionReport.summary.totalRowsFound })}</p>
            
            {Object.keys(extractionReport.summary.perSheetSummary).length > 1 && (
                <ul className="text-sm mt-2 max-h-32 overflow-y-auto">
                {Object.entries(extractionReport.summary.perSheetSummary).map(([sheetName, count]) => (
                    <li key={sheetName} className="flex justify-between">
                    <span>{sheetName}:</span>
                    <span className="font-medium">{count}</span>
                    </li>
                ))}
                </ul>
            )}

            {previewData && previewData.length > 0 && (
                <div className="mt-4">
                    <Label className="text-sm font-medium">{t('extractor.previewLabel')}</Label>
                    <Table>
                        <TableHeader>
                            <TableRow>
                                {Object.keys(previewData[0]).map(key => <TableHead key={key}>{key}</TableHead>)}
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                            {previewData.map((row, index) => (
                                <TableRow key={index}>
                                    {Object.values(row).map((value, cellIndex) => (
                                        <TableCell key={cellIndex}>{String(value)}</TableCell>
                                    ))}
                                </TableRow>
                            ))}
                        </TableBody>
                    </Table>
                </div>
            )}
          </div>
            <div className="w-full p-4 border rounded-md bg-secondary/30 space-y-4">
                <Label className="text-md font-semibold font-headline">{t('common.outputOptions.title')}</Label>
                <RadioGroup value={outputFormat} onValueChange={(v) => setOutputFormat(v as any)} className="space-y-3">
                    <div>
                        <div className="flex items-center space-x-2">
                            <RadioGroupItem value="xlsx" id="format-xlsx-extractor" />
                            <Label htmlFor="format-xlsx-extractor" className="font-normal">{t('common.outputOptions.xlsx')}</Label>
                        </div>
                        <p className="text-xs text-muted-foreground pl-6 pt-1">{t('common.outputOptions.xlsxDesc')}</p>
                    </div>
                    <div>
                        <div className="flex items-center space-x-2">
                            <RadioGroupItem value="xlsm" id="format-xlsm-extractor" />
                            <Label htmlFor="format-xlsm-extractor" className="font-normal">{t('common.outputOptions.xlsm')}</Label>
                        </div>
                        <p className="text-xs text-muted-foreground pl-6 pt-1">{t('common.outputOptions.xlsmDesc')}</p>
                    </div>
                </RadioGroup>
                <Alert variant="default" className="mt-2">
                    <Lightbulb className="h-4 w-4" />
                    <AlertDescription>{t('common.outputOptions.recommendation')}</AlertDescription>
                </Alert>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Button onClick={handleDownloadReport} variant="outline" className="w-full" disabled={isProcessing || extractionReport.summary.totalRowsFound === 0}>
                    <Download className="mr-2 h-5 w-5" />
                    {t('extractor.downloadBtn')}
                </Button>
                <Button onClick={handleDownloadPdf} variant="outline" className="w-full" disabled={isProcessing || extractionReport.summary.totalRowsFound === 0}>
                    <FileText className="mr-2 h-5 w-5" />
                    {t('extractor.downloadPdfBtn')}
                </Button>
            </div>
        </CardFooter>
      )}
    </Card>
  );
}
