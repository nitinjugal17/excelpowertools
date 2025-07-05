
"use client";

import React, { useState, useCallback, ChangeEvent, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, Paintbrush, ListChecks, CheckCircle2, Loader2, CaseSensitive, WholeWord, Palette, AlignHorizontalJustifyCenter, Type, Pilcrow, ChevronDown, Lightbulb, ScrollText, XCircle, Rows, Columns, HelpCircle } from 'lucide-react';
import { findAndFormatText } from '@/lib/excel-text-formatter';
import { generateTextFormatterVbs } from '@/lib/vbs-generators';
import type { TextFormatConfig, HorizontalAlignment, VerticalAlignment } from '@/lib/excel-types';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { DropdownMenu, DropdownMenuTrigger, DropdownMenuContent, DropdownMenuItem } from '@/components/ui/dropdown-menu';
import { useLanguage } from '@/context/language-context';
import { Textarea } from '@/components/ui/textarea';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';

interface SheetSelection {
  [sheetName: string]: boolean;
}

const isValidHex = (hex: string) => /^([0-9A-F]{6})$/i.test(hex);

interface TextFormatterPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function TextFormatterPage({ onProcessingChange, onFileStateChange }: TextFormatterPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<SheetSelection>({});
  
  const [searchMode, setSearchMode] = useState<'text' | 'regex'>('text');
  const [searchText, setSearchText] = useState<string>('');
  const [matchCase, setMatchCase] = useState<boolean>(false);
  const [matchEntireCell, setMatchEntireCell] = useState<boolean>(true);

  const [enableRange, setEnableRange] = useState<boolean>(false);
  const [rangeConfig, setRangeConfig] = useState({
    startRow: 1,
    endRow: 100,
    startCol: 'A',
    endCol: 'Z',
  });

  const [enableFontFormatting, setEnableFontFormatting] = useState<boolean>(true);
  const [fontConfig, setFontConfig] = useState({
    bold: true,
    italic: false,
    underline: false,
    name: 'Calibri',
    size: 11,
    color: '000000',
  });

  const [enableFillFormatting, setEnableFillFormatting] = useState<boolean>(false);
  const [fillColor, setFillColor] = useState<string>('FFFF00');

  const [enableAlignmentFormatting, setEnableAlignmentFormatting] = useState<boolean>(false);
  const [alignmentConfig, setAlignmentConfig] = useState({
    horizontal: 'general' as HorizontalAlignment,
    vertical: 'center' as VerticalAlignment,
  });

  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const cancellationRequested = useRef(false);

  const [processedWorkbook, setProcessedWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [cellsFormattedCount, setCellsFormattedCount] = useState<number | null>(null);
  const [outputFormat, setOutputFormat] = useState<'xlsx' | 'xlsm'>('xlsm');
  const [vbscriptPreview, setVbscriptPreview] = useState<string>('');
  const { toast } = useToast();

  const regexExamples = t('formatter.regexTips.examples');

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
    if (!file) {
      setSheetNames([]);
      setSelectedSheets({});
      setSearchText('');
      setMatchCase(false);
      setMatchEntireCell(true);
      setSearchMode('text');
      setEnableRange(false);
      setRangeConfig({ startRow: 1, endRow: 100, startCol: 'A', endCol: 'Z' });
      setEnableFontFormatting(true);
      setFontConfig({ bold: true, italic: false, underline: false, name: 'Calibri', size: 11, color: '000000' });
      setEnableFillFormatting(false);
      setFillColor('FFFF00');
      setEnableAlignmentFormatting(false);
      setAlignmentConfig({ horizontal: 'general', vertical: 'center' });
      setProcessedWorkbook(null);
      setCellsFormattedCount(null);
    } else {
      const getSheetNamesFromFile = async () => {
        setIsProcessing(true);
        setProcessedWorkbook(null);
        setCellsFormattedCount(null);
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

    const config: TextFormatConfig = {
      searchText: searchText.split('\n').map(term => term.trim()).filter(Boolean),
      searchMode,
      matchCase,
      matchEntireCell,
      range: enableRange ? rangeConfig : undefined,
      style: {}
    };

    if (enableFontFormatting) config.style.font = fontConfig;
    if (enableFillFormatting) config.style.fill = { color: fillColor };
    if (enableAlignmentFormatting) config.style.alignment = alignmentConfig;
      
    const script = generateTextFormatterVbs(sheetsToUpdate, config);
    setVbscriptPreview(script);

  }, [selectedSheets, searchText, searchMode, matchCase, matchEntireCell, enableFontFormatting, fontConfig, enableFillFormatting, fillColor, enableAlignmentFormatting, alignmentConfig, enableRange, rangeConfig]);


  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      if (!selectedFile.name.match(/\.(xlsx|xls|xlsm)$/)) {
        toast({ title: [t('toast.invalidFileType')].flat().join(' '), description: [t('toast.invalidFileTypeDesc')].flat().join(' '), variant: 'destructive' });
        setFile(null);
        return;
      }
      setFile(selectedFile);
    } else {
      setFile(null);
    }
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

  const handleSheetSelectionChange = (sheetName: string, checked: boolean) => {
    setSelectedSheets(prev => ({ ...prev, [sheetName]: checked }));
  };
  
  const handleCancel = () => {
    cancellationRequested.current = true;
    setProcessingStatus([t('common.cancelling')].flat().join(' '));
  };

  const handleProcess = useCallback(async () => {
    const sheetsToProcess = sheetNames.filter(name => selectedSheets[name]);
    const searchTerms = searchText.split('\n').map(term => term.trim()).filter(Boolean);

    if (!file || sheetsToProcess.length === 0) {
      toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('formatter.toast.missingFileOrSheet')].flat().join(' '), variant: 'destructive' });
      return;
    }
    if (!enableFontFormatting && !enableFillFormatting && !enableAlignmentFormatting) {
        toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('formatter.toast.noFormatting')].flat().join(' '), variant: 'destructive' });
        return;
    }
    if (enableFillFormatting && !isValidHex(fillColor)) {
      toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('formatter.toast.invalidFill')].flat().join(' '), variant: 'destructive'});
      return;
    }
     if (enableFontFormatting && fontConfig.color && !isValidHex(fontConfig.color)) {
      toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('formatter.toast.invalidFontColor')].flat().join(' '), variant: 'destructive'});
      return;
    }

    cancellationRequested.current = false;
    setIsProcessing(true);
    setProcessingStatus('');
    setProcessedWorkbook(null);
    setCellsFormattedCount(null);
    
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, cellDates: true, bookVBA: true, bookFiles: true });
      
      const formatConfig: TextFormatConfig = {
        searchText: searchTerms,
        searchMode,
        matchCase,
        matchEntireCell,
        range: enableRange ? rangeConfig : undefined,
        style: {}
      };

      if (enableFontFormatting) formatConfig.style.font = fontConfig;
      if (enableFillFormatting) formatConfig.style.fill = { color: fillColor };
      if (enableAlignmentFormatting) formatConfig.style.alignment = alignmentConfig;

      const onProgress = (status: { sheetName: string; currentSheet: number; totalSheets: number; cellsFormatted: number }) => {
        if (cancellationRequested.current) throw new Error('Cancelled by user.');
        setProcessingStatus([t('formatter.toast.processing', {current: status.currentSheet, total: status.totalSheets, sheetName: status.sheetName, count: status.cellsFormatted})].flat().join(' '));
      };
      
      const { workbook: newWb, cellsFormatted } = findAndFormatText(
        workbook,
        sheetsToProcess,
        formatConfig,
        onProgress,
        cancellationRequested
      );

      setProcessedWorkbook(newWb);
      setCellsFormattedCount(cellsFormatted);
      
      toast({
        title: [t('toast.processingComplete')].flat().join(' '),
        description: [t('formatter.toast.success', { count: cellsFormatted })].flat().join(' '),
        action: <CheckCircle2 className="text-green-500" />,
      });

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : [t('formatter.toast.error')].flat().join(' ');
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
  }, [file, sheetNames, selectedSheets, searchText, searchMode, matchCase, matchEntireCell, enableFontFormatting, fontConfig, enableFillFormatting, fillColor, enableAlignmentFormatting, alignmentConfig, toast, t, enableRange, rangeConfig]);
  
  const handleDownloadFile = useCallback(() => {
     if (!processedWorkbook || !file) {
       toast({ title: [t('toast.noFileToDownload')].flat().join(' '), description: [t('formatter.toast.noFile')].flat().join(' '), variant: "destructive" });
       return;
     }
     try {
        const originalFileName = file.name.substring(0, file.name.lastIndexOf('.'));
        XLSX.writeFile(processedWorkbook, `${originalFileName}_formatted.${outputFormat}`, { compression: true, bookType: outputFormat, cellStyles: true });
        toast({ title: [t('toast.downloadSuccess')].flat().join(' ') });
     } catch (error) {
        toast({ title: [t('toast.downloadError')].flat().join(' '), description: [t('toast.downloadError')].flat().join(' '), variant: 'destructive' });
     }
  }, [processedWorkbook, file, toast, t, outputFormat]);

  const allSheetsSelected = sheetNames.length > 0 && sheetNames.every(name => selectedSheets[name]);
  const isProcessButtonDisabled = isProcessing || !file || (!enableFontFormatting && !enableFillFormatting && !enableAlignmentFormatting);

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
          <Paintbrush className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{[t('formatter.title')].flat().join(' ')}</CardTitle>
        </div>
        <CardDescription className="font-body">
          {[t('formatter.description')].flat().join(' ')}
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-formatter" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{[t('formatter.uploadStep')].flat().join(' ')}</span>
          </Label>
          <Input
            id="file-upload-formatter"
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
              <span>{[t('formatter.selectSheetsStep')].flat().join(' ')}</span>
            </Label>
            <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
              <Checkbox
                id="select-all-sheets-formatter"
                checked={allSheetsSelected}
                onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                disabled={isProcessing}
              />
              <Label htmlFor="select-all-sheets-formatter" className="text-sm font-medium flex-grow">
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
                      id={`sheet-formatter-${name}`}
                      checked={selectedSheets[name] || false}
                      onCheckedChange={(checked) => handleSheetSelectionChange(name, checked as boolean)}
                      disabled={isProcessing}
                    />
                    <Label htmlFor={`sheet-formatter-${name}`} className="text-sm font-normal">{name}</Label>
                  </div>
                ))}
              </div>
            </Card>
          </div>
        )}

        <div className="space-y-2">
            <Label className="flex items-center space-x-2 text-sm font-medium">{[t('formatter.searchMode')].flat().join(' ')}</Label>
            <RadioGroup value={searchMode} onValueChange={(v) => setSearchMode(v as any)} className="grid grid-cols-2 gap-2">
                <Label htmlFor="mode-text" className="p-2 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer flex items-center justify-center gap-2">
                    <RadioGroupItem value="text" id="mode-text" />
                    {[t('formatter.modeText')].flat().join(' ')}
                </Label>
                <Label htmlFor="mode-regex" className="p-2 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer flex items-center justify-center gap-2">
                    <RadioGroupItem value="regex" id="mode-regex" />
                    {[t('formatter.modeRegex')].flat().join(' ')}
                </Label>
            </RadioGroup>
        </div>

        <div className="space-y-2">
            <Label htmlFor="search-text" className="flex items-center space-x-2 text-sm font-medium">
                <Type className="h-5 w-5" />
                <span>{[t('formatter.findStep')].flat().join(' ')}</span>
            </Label>
            <Textarea
                id="search-text" 
                value={searchText} 
                onChange={e => setSearchText(e.target.value)} 
                disabled={isProcessing || !file}
                placeholder={[t('formatter.findPlaceholder')].flat().join(' ')}
                rows={4}
            />
            <p className="text-xs text-muted-foreground">{[t('formatter.findDesc')].flat().join(' ')}</p>
            <div className="flex items-center gap-4 pt-2">
                 <div className="flex items-center space-x-2">
                    <Checkbox id="match-case" checked={matchCase} onCheckedChange={c => setMatchCase(c as boolean)} disabled={isProcessing || !file} />
                    <Label htmlFor="match-case" className="flex items-center space-x-1 font-normal"><CaseSensitive className="h-4 w-4"/><span>{[t('formatter.matchCase')].flat().join(' ')}</span></Label>
                </div>
                 <div className="flex items-center space-x-2">
                    <Checkbox id="match-cell" checked={matchEntireCell} onCheckedChange={c => setMatchEntireCell(c as boolean)} disabled={isProcessing || !file} />
                    <Label htmlFor="match-cell" className="flex items-center space-x-1 font-normal"><WholeWord className="h-4 w-4"/><span>{[t('formatter.matchEntireCell')].flat().join(' ')}</span></Label>
                </div>
            </div>
        </div>

        {searchMode === 'regex' && Array.isArray(regexExamples) && (
            <Alert>
                <HelpCircle className="h-4 w-4" />
                <AlertTitle>{[t('formatter.regexTips.title')].flat().join(' ')}</AlertTitle>
                <AlertDescription className="pt-2">
                    <p className="pb-4">{[t('formatter.regexTips.description')].flat().join(' ')}</p>
                    <Accordion type="single" collapsible className="w-full">
                        {regexExamples.map((item: any, index: number) => (
                            <AccordionItem value={`item-${index}`} key={index}>
                                <AccordionTrigger>{item.title}</AccordionTrigger>
                                <AccordionContent className="space-y-2">
                                    <p className="text-sm text-muted-foreground">{item.description}</p>
                                    <pre className="text-xs p-2 bg-muted rounded-md font-code"><code>{item.pattern}</code></pre>
                                    <p className="text-xs text-muted-foreground italic">{[t('formatter.regexTips.exampleLabel')].flat().join(' ')} {item.example}</p>
                                </AccordionContent>
                            </AccordionItem>
                        ))}
                    </Accordion>
                </AlertDescription>
            </Alert>
        )}

        <Card className="p-4 bg-secondary/30">
            <CardHeader className="p-0 pb-4">
                <div className="flex items-start space-x-3">
                  <Checkbox id="enable-range" checked={enableRange} onCheckedChange={c => setEnableRange(c as boolean)} className="mt-1" />
                  <div className="grid gap-1.5 leading-none w-full">
                    <Label htmlFor="enable-range" className="text-md font-semibold">{[t('formatter.rangeStep')].flat().join(' ')}</Label>
                    <p className="text-xs text-muted-foreground">{[t('formatter.rangeDesc')].flat().join(' ')}</p>
                  </div>
                </div>
            </CardHeader>
            {enableRange && (
                <CardContent className="p-0 pt-4">
                   <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                        <div className="space-y-2">
                            <Label htmlFor="start-row" className="flex items-center space-x-2 text-sm"><Rows className="h-4 w-4" /><span>{[t('formatter.startRow')].flat().join(' ')}</span></Label>
                            <Input id="start-row" type="number" min="1" value={rangeConfig.startRow} onChange={e => setRangeConfig(p => ({...p, startRow: parseInt(e.target.value, 10) || 1}))} />
                        </div>
                        <div className="space-y-2">
                            <Label htmlFor="end-row" className="flex items-center space-x-2 text-sm"><Rows className="h-4 w-4" /><span>{[t('formatter.endRow')].flat().join(' ')}</span></Label>
                            <Input id="end-row" type="number" min="1" value={rangeConfig.endRow} onChange={e => setRangeConfig(p => ({...p, endRow: parseInt(e.target.value, 10) || 1}))} />
                        </div>
                        <div className="space-y-2">
                            <Label htmlFor="start-col" className="flex items-center space-x-2 text-sm"><Columns className="h-4 w-4" /><span>{[t('formatter.startCol')].flat().join(' ')}</span></Label>
                            <Input id="start-col" value={rangeConfig.startCol} onChange={e => setRangeConfig(p => ({...p, startCol: e.target.value.toUpperCase()}))} placeholder={[t('formatter.startColPlaceholder')].flat().join(' ')}/>
                        </div>
                        <div className="space-y-2">
                            <Label htmlFor="end-col" className="flex items-center space-x-2 text-sm"><Columns className="h-4 w-4" /><span>{[t('formatter.endCol')].flat().join(' ')}</span></Label>
                            <Input id="end-col" value={rangeConfig.endCol} onChange={e => setRangeConfig(p => ({...p, endCol: e.target.value.toUpperCase()}))} placeholder={[t('formatter.endColPlaceholder')].flat().join(' ')}/>
                        </div>
                   </div>
                </CardContent>
            )}
        </Card>

        <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
          <CardHeader className="p-0 pb-4">
            <Label className="flex items-center space-x-2 text-md font-semibold text-primary">
              <Palette className="h-5 w-5" />
              <span>{[t('formatter.formatStep')].flat().join(' ')}</span>
            </Label>
          </CardHeader>
          <CardContent className="p-0 space-y-4">
             <div className="flex items-start space-x-3">
              <Checkbox id="enable-font" checked={enableFontFormatting} onCheckedChange={c => setEnableFontFormatting(c as boolean)} className="mt-1" />
               <div className="grid gap-2 leading-none w-full">
                <Label htmlFor="enable-font">{[t('formatter.fontFormatting')].flat().join(' ')}</Label>
                {enableFontFormatting && <Card className="p-4 mt-2 space-y-4 bg-background">
                  <div className="flex items-center space-x-4">
                    <div className="flex items-center space-x-2"><Checkbox id="font-bold" checked={!!fontConfig.bold} onCheckedChange={c => setFontConfig(p => ({...p, bold: c as boolean}))} /><Label htmlFor="font-bold" className="font-normal">{[t('common.bold')].flat().join(' ')}</Label></div>
                    <div className="flex items-center space-x-2"><Checkbox id="font-italic" checked={!!fontConfig.italic} onCheckedChange={c => setFontConfig(p => ({...p, italic: c as boolean}))} /><Label htmlFor="font-italic" className="font-normal">{[t('common.italic')].flat().join(' ')}</Label></div>
                    <div className="flex items-center space-x-2"><Checkbox id="font-underline" checked={!!fontConfig.underline} onCheckedChange={c => setFontConfig(p => ({...p, underline: c as boolean}))} /><Label htmlFor="font-underline" className="font-normal">{[t('common.underline')].flat().join(' ')}</Label></div>
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div><Label htmlFor="font-name" className="text-sm">{[t('common.fontName')].flat().join(' ')}</Label><Input id="font-name" value={fontConfig.name || ''} onChange={e => setFontConfig(p => ({...p, name: e.target.value}))} placeholder="Calibri" /></div>
                    <div><Label htmlFor="font-size" className="text-sm">{[t('common.fontSize')].flat().join(' ')}</Label><Input id="font-size" type="number" min="1" value={fontConfig.size || 11} onChange={e => setFontConfig(p => ({...p, size: parseInt(e.target.value, 10) || 11}))} /></div>
                  </div>
                  <div><Label htmlFor="font-color" className="text-sm">{[t('updater.fontColorHex')].flat().join(' ')}</Label><Input id="font-color" value={fontConfig.color || ''} onChange={e => setFontConfig(p => ({...p, color: e.target.value.replace('#', '')}))} placeholder="000000" /></div>
                </Card>}
              </div>
            </div>

            <div className="flex items-start space-x-3">
              <Checkbox id="enable-fill" checked={enableFillFormatting} onCheckedChange={c => setEnableFillFormatting(c as boolean)} className="mt-1" />
              <div className="grid gap-2 leading-none w-full">
                <Label htmlFor="enable-fill">{[t('formatter.fillFormatting')].flat().join(' ')}</Label>
                {enableFillFormatting && <Card className="p-4 mt-2 space-y-2 bg-background">
                  <div><Label htmlFor="fill-color" className="text-sm">{[t('formatter.fillColor')].flat().join(' ')}</Label><Input id="fill-color" value={fillColor || ''} onChange={e => setFillColor(e.target.value.replace('#', ''))} placeholder="FFFF00" /></div>
                </Card>}
              </div>
            </div>

            <div className="flex items-start space-x-3">
              <Checkbox id="enable-align" checked={enableAlignmentFormatting} onCheckedChange={c => setEnableAlignmentFormatting(c as boolean)} className="mt-1" />
              <div className="grid gap-2 leading-none w-full">
                <Label htmlFor="enable-align">{[t('formatter.alignmentFormatting')].flat().join(' ')}</Label>
                {enableAlignmentFormatting && <Card className="p-4 mt-2 grid grid-cols-2 gap-4 bg-background">
                   <div>
                    <Label htmlFor="h-align" className="text-sm">{[t('formatter.horizontal')].flat().join(' ')}</Label>
                    <Select value={alignmentConfig.horizontal || 'general'} onValueChange={v => setAlignmentConfig(p => ({...p, horizontal: v as HorizontalAlignment}))}>
                        <SelectTrigger id="h-align"><SelectValue /></SelectTrigger>
                        <SelectContent>
                            <SelectItem value="general">{[t('common.alignments.general')].flat().join(' ')}</SelectItem>
                            <SelectItem value="left">{[t('common.alignments.left')].flat().join(' ')}</SelectItem>
                            <SelectItem value="center">{[t('common.alignments.center')].flat().join(' ')}</SelectItem>
                            <SelectItem value="right">{[t('common.alignments.right')].flat().join(' ')}</SelectItem>
                            <SelectItem value="fill">{[t('common.alignments.fill')].flat().join(' ')}</SelectItem>
                            <SelectItem value="justify">{[t('common.alignments.justify')].flat().join(' ')}</SelectItem>
                        </SelectContent>
                    </Select>
                   </div>
                   <div>
                    <Label htmlFor="v-align" className="text-sm">{[t('formatter.vertical')].flat().join(' ')}</Label>
                     <Select value={alignmentConfig.vertical || 'center'} onValueChange={v => setAlignmentConfig(p => ({...p, vertical: v as VerticalAlignment}))}>
                        <SelectTrigger id="v-align"><SelectValue /></SelectTrigger>
                        <SelectContent>
                            <SelectItem value="top">{[t('common.alignments.top')].flat().join(' ')}</SelectItem>
                            <SelectItem value="center">{[t('common.alignments.center')].flat().join(' ')}</SelectItem>
                            <SelectItem value="bottom">{[t('common.alignments.bottom')].flat().join(' ')}</SelectItem>
                            <SelectItem value="justify">{[t('common.alignments.justify')].flat().join(' ')}</SelectItem>
                            <SelectItem value="distributed">{[t('common.alignments.distributed')].flat().join(' ')}</SelectItem>
                        </SelectContent>
                    </Select>
                   </div>
                </Card>}
              </div>
            </div>
          </CardContent>
        </Card>

        {file && (
          <div className="space-y-2">
            <Label className="flex items-center space-x-2 text-sm font-medium">
              <ScrollText className="h-5 w-5" />
              <span>{[t('formatter.vbsPreviewStep')].flat().join(' ')}</span>
            </Label>
            <Card className="bg-secondary/20">
              <CardContent className="p-0">
                <pre className="text-xs p-4 overflow-x-auto bg-gray-800 text-white rounded-md max-h-60">
                  <code>{vbscriptPreview}</code>
                </pre>
              </CardContent>
            </Card>
             <p className="text-xs text-muted-foreground">
              {[t('formatter.vbsPreviewDesc')].flat().join(' ')}
            </p>
          </div>
        )}

        <Button onClick={handleProcess} disabled={isProcessButtonDisabled} className="w-full">
          {isProcessing ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Paintbrush className="mr-2 h-5 w-5" />}
          {[t('formatter.processBtn')].flat().join(' ')}
        </Button>
      </CardContent>

      {cellsFormattedCount !== null && (
        <CardFooter className="flex-col space-y-4 items-stretch">
          <div className="p-4 border rounded-md bg-secondary/30">
            <h3 className="text-lg font-semibold mb-2 font-headline">{[t('formatter.resultsTitle')].flat().join(' ')}</h3>
            {cellsFormattedCount > 0 ? (
              <p>{[t('formatter.resultsFound', { count: cellsFormattedCount })].flat().join(' ')}</p>
            ) : (
              <p>{[t('formatter.resultsNotFound')].flat().join(' ')}</p>
            )}
          </div>
          
          {processedWorkbook && cellsFormattedCount > 0 && (
            <>
                <div className="w-full p-4 border rounded-md bg-secondary/30 space-y-4">
                    <Label className="text-md font-semibold font-headline">{[t('common.outputOptions.title')].flat().join(' ')}</Label>
                    <RadioGroup value={outputFormat} onValueChange={(v) => setOutputFormat(v as any)} className="space-y-3">
                        <div>
                            <div className="flex items-center space-x-2">
                                <RadioGroupItem value="xlsx" id="format-xlsx-formatter" />
                                <Label htmlFor="format-xlsx-formatter" className="font-normal">{[t('common.outputOptions.xlsx')].flat().join(' ')}</Label>
                            </div>
                            <p className="text-xs text-muted-foreground pl-6 pt-1">{[t('common.outputOptions.xlsxDesc')].flat().join(' ')}</p>
                        </div>
                        <div>
                            <div className="flex items-center space-x-2">
                                <RadioGroupItem value="xlsm" id="format-xlsm-formatter" />
                                <Label htmlFor="format-xlsm-formatter" className="font-normal">{[t('common.outputOptions.xlsm')].flat().join(' ')}</Label>
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
                <Button onClick={handleDownloadFile} variant="outline" className="w-full">
                  <Download className="mr-2 h-5 w-5" />
                  {[t('formatter.downloadBtn')].flat().join(' ')}
                </Button>
            </>
          )}
        </CardFooter>
      )}
    </Card>
  );
}
