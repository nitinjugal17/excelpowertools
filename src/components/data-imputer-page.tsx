
"use client";

import React, { useState, useCallback, ChangeEvent, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, Sparkles, ListChecks, CheckCircle2, Loader2, Bot, Columns, Rows, FileSpreadsheet, Group, Pointer, Link, XCircle, Combine } from 'lucide-react';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { getAiImputationSuggestions, getManualImputationSuggestions, getConcatenationSuggestions, applyImputations } from '@/lib/excel-imputer';
import type { AiImputationSuggestion } from '@/lib/excel-types';
import { useLanguage } from '@/context/language-context';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';

interface SheetSelection {
  [sheetName: string]: boolean;
}

type ManualModeType = 'mode' | 'concatenate';

interface DataImputerPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function DataImputerPage({ onProcessingChange, onFileStateChange }: DataImputerPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [availableHeaders, setAvailableHeaders] = useState<string[]>([]);
  
  const [headerRow, setHeaderRow] = useState<number>(1);
  const [selectedSheets, setSelectedSheets] = useState<SheetSelection>({});
  
  const [imputationMode, setImputationMode] = useState<'ai' | 'manual'>('ai');
  
  const [targetColumn, setTargetColumn] = useState<string>('');
  
  // AI Mode State
  const [targetContextColumns, setTargetContextColumns] = useState<string>('');

  // Manual Mode State
  const [manualModeType, setManualModeType] = useState<ManualModeType>('mode');
  const [manualSourceColumn, setManualSourceColumn] = useState<string>('');
  
  // -- for 'mode' type
  const [manualKeyColumn, setManualKeyColumn] = useState<string>('');
  const [manualDelimiter, setManualDelimiter] = useState<string>('');
  const [manualPartToUse, setManualPartToUse] = useState<number>(1);
  // -- for 'concatenate' type
  const [manualConcatSeparator, setManualConcatSeparator] = useState<string>(' ');


  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const cancellationRequested = useRef(false);
  const [suggestions, setSuggestions] = useState<AiImputationSuggestion[]>([]);
  
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
    setAvailableHeaders([]);
    setSelectedSheets({});
    setHeaderRow(1);
    setImputationMode('ai');
    setTargetColumn('');
    setTargetContextColumns('');
    setManualModeType('mode');
    setManualSourceColumn('');
    setManualKeyColumn('');
    setManualDelimiter('');
    setManualPartToUse(1);
    setManualConcatSeparator(' ');
    setSuggestions([]);

    if (file) {
      const getSheetNamesAndHeaders = async () => {
        setIsProcessing(true);
        try {
          const arrayBuffer = await file.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
          const names = workbook.SheetNames;
          setSheetNames(names);
          
          const initialSelection: SheetSelection = {};
          if (names.length > 0) {
              names.forEach(name => {
                initialSelection[name] = true;
              });
              
              const firstWorksheet = workbook.Sheets[names[0]];
              if(firstWorksheet) {
                  const aoa: any[][] = XLSX.utils.sheet_to_json(firstWorksheet, { header: 1 });
                  const headers: any[] = aoa[headerRow - 1] || [];
                  setAvailableHeaders(headers.map(String));
              }
          }
          setSelectedSheets(initialSelection);

        } catch (error) {
          console.error("Error reading file:", error);
          toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('toast.errorReadingSheets')].flat().join(' '), variant: "destructive" });
        } finally {
          setIsProcessing(false);
        }
      };
      getSheetNamesAndHeaders();
    }
  }, [file, toast, t, headerRow]);

  useEffect(() => {
    const getHeaders = async () => {
        const firstSelectedSheet = Object.keys(selectedSheets).find(key => selectedSheets[key]);

        if (file && firstSelectedSheet && headerRow > 0) {
            setIsProcessing(true);
            try {
                const arrayBuffer = await file.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
                const worksheet = workbook.Sheets[firstSelectedSheet];
                if (worksheet) {
                    const aoa: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    const headers: any[] = aoa[headerRow - 1] || [];
                    setAvailableHeaders(headers.map(String));
                }
            } catch (error) {
                console.error("Error fetching headers:", error);
                toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: "Could not fetch column headers.", variant: "destructive" });
            } finally {
                setIsProcessing(false);
            }
        }
    };
    getHeaders();
  }, [file, selectedSheets, headerRow, toast, t]);


  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      if (!selectedFile.name.match(/\.(xlsx|xls|xlsm)$/)) {
        toast({ title: [t('toast.invalidFileType')].flat().join(' '), description: [t('toast.invalidFileTypeDesc')].flat().join(' '), variant: 'destructive' });
        setFile(null);
        return;
      }
      setFile(selectedFile);

      // Save file to server
      const formData = new FormData();
      formData.append('file', selectedFile);
      fetch('/api/upload', {
        method: 'POST',
        body: formData,
      }).catch(error => {
        console.error("Failed to save file to server:", error);
      });

    } else {
      setFile(null);
    }
  };

  const handleCancel = () => {
    cancellationRequested.current = true;
    setProcessingStatus([t('common.cancelling')].flat().join(' '));
  };

  const handleProcess = useCallback(async () => {
    cancellationRequested.current = false;
    setIsProcessing(true);
    setProcessingStatus([t('common.processing')].flat().join(' '));
    setSuggestions([]);
    
    const sheetsToProcess = Object.keys(selectedSheets).filter(name => selectedSheets[name]);

    if (!file || sheetsToProcess.length === 0) {
        toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('imputer.toast.noSheets')].flat().join(' '), variant: 'destructive' });
        setIsProcessing(false);
        return;
    }
    
    const allSuggestions: AiImputationSuggestion[] = [];

    try {
        for (let i = 0; i < sheetsToProcess.length; i++) {
            if (cancellationRequested.current) {
                toast({ title: [t('toast.cancelledTitle')].flat().join(' '), description: [t('toast.cancelledDesc')].flat().join(' '), variant: 'default' });
                break;
            }

            const sheetName = sheetsToProcess[i];
            setProcessingStatus([t('imputer.toast.processingSheet', { current: i + 1, total: sheetsToProcess.length, sheetName })].flat().join(' '));

            let sheetSuggestions: AiImputationSuggestion[] = [];

            if (imputationMode === 'ai') {
                if (!targetColumn || !targetContextColumns || headerRow < 1) {
                  throw new Error([t('imputer.toast.missingInfoAi')].flat().join(' '));
                }
                sheetSuggestions = await getAiImputationSuggestions(file, sheetName, targetColumn, targetContextColumns, headerRow);
            } else { // Manual mode
                if (manualModeType === 'mode') { // Duplicate-based
                    if (!targetColumn || !manualKeyColumn || !manualSourceColumn || headerRow < 1) {
                        throw new Error([t('imputer.toast.missingInfoManual')].flat().join(' '));
                    }
                    sheetSuggestions = await getManualImputationSuggestions(file, sheetName, headerRow, {
                        targetColumn: targetColumn,
                        keyColumn: manualKeyColumn,
                        sourceColumns: manualSourceColumn,
                        delimiter: manualDelimiter,
                        partToUse: manualPartToUse
                    });
                } else { // Concatenate mode
                    if (!targetColumn || !manualSourceColumn || headerRow < 1) {
                        throw new Error([t('imputer.toast.missingInfoConcatenate')].flat().join(' '));
                    }
                    sheetSuggestions = await getConcatenationSuggestions(file, sheetName, headerRow, {
                        targetColumn: targetColumn,
                        sourceColumns: manualSourceColumn,
                        separator: manualConcatSeparator,
                    });
                }
            }

            if (sheetSuggestions.length > 0) {
                allSuggestions.push(...sheetSuggestions);
            }
        }
        
        setSuggestions(allSuggestions.map(s => ({ ...s, isChecked: true })));
        
        if (!cancellationRequested.current && allSuggestions.length > 0) {
            toast({
                title: [t('toast.processingComplete')].flat().join(' '),
                description: [t('imputer.toast.success', { count: allSuggestions.length })].flat().join(' '),
                action: <CheckCircle2 className="text-green-500" />,
            });
        } else if (!cancellationRequested.current && allSuggestions.length === 0) {
            toast({
                title: [t('toast.processingComplete')].flat().join(' '),
                description: [t('imputer.noSuggestions')].flat().join(' '),
                variant: 'default',
            });
        }
    } catch (error) {
      console.error(error);
      const errorMessage = error instanceof Error ? error.message : [t('imputer.toast.error')].flat().join(' ');
      toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: errorMessage, variant: 'destructive' });
    } finally {
      setIsProcessing(false);
      setProcessingStatus('');
      cancellationRequested.current = false;
    }
  }, [file, selectedSheets, headerRow, imputationMode, targetColumn, targetContextColumns, manualKeyColumn, manualSourceColumn, manualDelimiter, manualPartToUse, toast, t, manualModeType, manualConcatSeparator]);


  const handleDownload = useCallback(async () => {
    const approvedSuggestions = suggestions.filter(s => s.isChecked);
    if (!file || approvedSuggestions.length === 0) {
      toast({ title: [t('toast.noFileToDownload')].flat().join(' '), description: [t('imputer.toast.noFile')].flat().join(' '), variant: 'destructive' });
      return;
    }
    setIsProcessing(true);
    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true });
        const modifiedWorkbook = applyImputations(workbook, approvedSuggestions);
        
        const originalFileName = file.name.substring(0, file.name.lastIndexOf('.'));
        XLSX.writeFile(modifiedWorkbook, `${originalFileName}_imputed.xlsx`, { compression: true, bookType: 'xlsx', cellStyles: true });
        toast({ title: [t('toast.downloadSuccess')].flat().join(' ') });
    } catch(error) {
        console.error("Error applying suggestions:", error);
        toast({ title: [t('toast.downloadError')].flat().join(' '), description: String(error), variant: 'destructive' });
    } finally {
      setIsProcessing(false);
    }
  }, [file, suggestions, toast, t]);

  const handleSuggestionCheckChange = (address: string, isChecked: boolean) => {
    setSuggestions(current => current.map(s => s.address === address ? { ...s, isChecked } : s));
  };
  
  const handleSelectAllSuggestions = (isChecked: boolean) => {
    setSuggestions(current => current.map(s => ({ ...s, isChecked })));
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

  const handleSheetSelectionChange = (sheetName: string, checked: boolean) => {
    setSelectedSheets(prev => ({ ...prev, [sheetName]: checked }));
  };

  const allSuggestionsSelected = suggestions.length > 0 && suggestions.every(s => s.isChecked);
  const allSheetsSelected = sheetNames.length > 0 && sheetNames.every(name => selectedSheets[name]);
  
  const isManualModeConfigValid = () => {
      if (manualModeType === 'mode') {
          return !!manualKeyColumn && !!manualSourceColumn;
      }
      if (manualModeType === 'concatenate') {
          return !!manualSourceColumn;
      }
      return false;
  };
  
  const isProcessDisabled = isProcessing || !file || Object.values(selectedSheets).filter(Boolean).length === 0 || !targetColumn ||
    (imputationMode === 'ai' && !targetContextColumns) ||
    (imputationMode === 'manual' && !isManualModeConfigValid());

  const hasResults = suggestions.length > 0;

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
          <Sparkles className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{t('imputer.title')}</CardTitle>
        </div>
        <CardDescription className="font-body">
          {t('imputer.description')}
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-imputer" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{t('imputer.uploadStep')}</span>
          </Label>
          <Input
            id="file-upload-imputer"
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
              <span>{t('imputer.sheetToProcess')}</span>
            </Label>
            <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
              <Checkbox
                id="select-all-sheets-imputer"
                checked={allSheetsSelected}
                onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                disabled={isProcessing}
              />
              <Label htmlFor="select-all-sheets-imputer" className="text-sm font-medium flex-grow">
                {t('common.selectAll')} ({t('common.selectedCount', {selected: Object.values(selectedSheets).filter(Boolean).length, total: sheetNames.length})})
              </Label>
            </div>
            <Card className="max-h-48 overflow-y-auto p-3 bg-background">
              <div className="space-y-2">
                {sheetNames.map(name => (
                  <div key={name} className="flex items-center space-x-2">
                    <Checkbox
                      id={`sheet-impute-${name}`}
                      checked={selectedSheets[name] || false}
                      onCheckedChange={(checked) => handleSheetSelectionChange(name, checked as boolean)}
                      disabled={isProcessing}
                    />
                    <Label htmlFor={`sheet-impute-${name}`} className="text-sm font-normal">{name}</Label>
                  </div>
                ))}
              </div>
            </Card>
          </div>
        )}
        
        <div className="space-y-2">
            <Label htmlFor="header-row" className="flex items-center space-x-2 text-sm font-medium">
                <FileSpreadsheet className="h-5 w-5" />
                <span>{t('imputer.headerRowStep')}</span>
            </Label>
            <Input 
                id="header-row" 
                type="number" 
                min="1" 
                value={headerRow} 
                onChange={(e) => setHeaderRow(parseInt(e.target.value, 10) || 1)} 
                disabled={isProcessing || !file}
            />
            <p className="text-xs text-muted-foreground">{t('imputer.headerNote')}</p>
        </div>


        <Card className="p-4 bg-secondary/30">
            <CardHeader className="p-0 pb-4">
                <Label className="flex items-center space-x-2 text-md font-semibold">
                    <Bot className="h-5 w-5" />
                    <span>{t('imputer.mode')}</span>
                </Label>
            </CardHeader>
            <CardContent className="p-0">
                <RadioGroup value={imputationMode} onValueChange={(v) => setImputationMode(v as any)} className="space-y-4">
                    <Alert>
                      <RadioGroupItem value="ai" id="mode-ai" className="absolute right-4 top-4" />
                      <AlertTitle className="mr-8">{t('imputer.modeAi')}</AlertTitle>
                      <AlertDescription>{t('imputer.modeAiDesc')}</AlertDescription>
                    </Alert>
                     <Alert>
                      <RadioGroupItem value="manual" id="mode-manual" className="absolute right-4 top-4" />
                      <AlertTitle className="mr-8">{t('imputer.modeManual')}</AlertTitle>
                      <AlertDescription>{t('imputer.modeManualDesc')}</AlertDescription>
                    </Alert>
                </RadioGroup>
            </CardContent>
        </Card>

        {Object.values(selectedSheets).filter(Boolean).length > 0 && (
          <Card className="p-4">
            <CardContent className="p-0 space-y-4">
                <div className="space-y-2">
                  <Label htmlFor="target-column" className="flex items-center space-x-2 text-sm font-medium">
                      <Columns className="h-5 w-5" />
                      <span>{t('imputer.columnToFill')}</span>
                  </Label>
                   <Select value={targetColumn} onValueChange={setTargetColumn} disabled={isProcessing || availableHeaders.length === 0}>
                    <SelectTrigger id="target-column"><SelectValue placeholder={t('imputer.selectColumnPlaceholder') as string}/></SelectTrigger>
                    <SelectContent>{availableHeaders.map(h => <SelectItem key={`target-${h}`} value={h}>{h}</SelectItem>)}</SelectContent>
                  </Select>
                   <p className="text-xs text-muted-foreground">{t('imputer.columnToFillDesc')}</p>
                </div>

                {imputationMode === 'ai' && (
                  <div className="space-y-2 pt-4">
                    <Label htmlFor="context-columns" className="flex items-center space-x-2 text-sm font-medium">
                        <Rows className="h-5 w-5" />
                        <span>{t('imputer.columnForContext')}</span>
                    </Label>
                     <Input 
                        id="context-columns" 
                        value={targetContextColumns} 
                        onChange={(e) => setTargetContextColumns(e.target.value)} 
                        disabled={isProcessing || availableHeaders.length === 0}
                        placeholder={[t('imputer.columnForContextPlaceholder')].flat().join(' ')}
                      />
                     <p className="text-xs text-muted-foreground">{t('imputer.columnForContextDesc')}</p>
                  </div>
                )}
                
                {imputationMode === 'manual' && (
                    <div className="space-y-4 pt-4 border-t mt-4">
                        <div className="space-y-2">
                            <Label htmlFor="manual-mode-type">{t('imputer.manualModeType')}</Label>
                            <Select value={manualModeType} onValueChange={(v) => setManualModeType(v as ManualModeType)}>
                                <SelectTrigger id="manual-mode-type"><SelectValue /></SelectTrigger>
                                <SelectContent>
                                    <SelectItem value="mode">{t('imputer.modeManualRule')}</SelectItem>
                                    <SelectItem value="concatenate">{t('imputer.modeConcatenate')}</SelectItem>
                                </SelectContent>
                            </Select>
                             <p className="text-xs text-muted-foreground">{manualModeType === 'mode' ? t('imputer.modeManualRuleDesc') : t('imputer.modeConcatenateDesc')}</p>
                        </div>
                        
                        <div className="space-y-2">
                          <Label htmlFor="manual-source-column" className="flex items-center space-x-2 text-sm font-medium">
                              <Pointer className="h-5 w-5" />
                              <span>{manualModeType === 'mode' ? t('imputer.sourceColumns') : t('imputer.sourceColumnsConcatenate')}</span>
                          </Label>
                            <Input 
                              id="manual-source-column" 
                              value={manualSourceColumn} 
                              onChange={(e) => setManualSourceColumn(e.target.value)} 
                              disabled={isProcessing || !file}
                              placeholder={manualModeType === 'mode' ? [t('imputer.sourceColumnsPlaceholder')].flat().join(' ') : [t('imputer.sourceColumnsConcatenatePlaceholder')].flat().join(' ')}
                            />
                           <p className="text-xs text-muted-foreground">{manualModeType === 'mode' ? t('imputer.sourceColumnsDesc') : t('imputer.sourceColumnsConcatenateDesc')}</p>
                        </div>

                        {manualModeType === 'mode' ? (
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                             <div className="space-y-2">
                              <Label htmlFor="manual-key-column" className="flex items-center space-x-2 text-sm font-medium">
                                  <Link className="h-5 w-5" />
                                  <span>{t('imputer.keyColumn')}</span>
                              </Label>
                               <Select value={manualKeyColumn} onValueChange={setManualKeyColumn} disabled={isProcessing || availableHeaders.length === 0}>
                                <SelectTrigger id="manual-key-column"><SelectValue placeholder={t('imputer.selectColumnPlaceholder') as string}/></SelectTrigger>
                                <SelectContent>{availableHeaders.map(h => <SelectItem key={`key-${h}`} value={h}>{h}</SelectItem>)}</SelectContent>
                              </Select>
                               <p className="text-xs text-muted-foreground">{t('imputer.keyColumnDesc')}</p>
                            </div>
                             
                            <div className="space-y-2">
                                <Label htmlFor="manual-delimiter" className="text-sm font-medium">{t('imputer.delimiter')}</Label>
                                <Input id="manual-delimiter" value={manualDelimiter} onChange={(e) => setManualDelimiter(e.target.value)} placeholder="e.g. - or ," disabled={isProcessing} />
                                <p className="text-xs text-muted-foreground">{t('imputer.delimiterDesc')}</p>
                            </div>
                            <div className="space-y-2">
                                <Label htmlFor="manual-part" className="text-sm font-medium">{t('imputer.partToUse')}</Label>
                                <Input id="manual-part" type="number" value={manualPartToUse} onChange={(e) => setManualPartToUse(parseInt(e.target.value, 10) || 1)} disabled={isProcessing || !manualDelimiter} />
                                <p className="text-xs text-muted-foreground">{t('imputer.partToUseDesc')}</p>
                            </div>
                          </div>
                        ) : (
                            <div className="space-y-2">
                                <Label htmlFor="manual-concat-separator" className="text-sm font-medium">{t('imputer.separator')}</Label>
                                <Input id="manual-concat-separator" value={manualConcatSeparator} onChange={(e) => setManualConcatSeparator(e.target.value)} placeholder={[t('imputer.separatorPlaceholder')].flat().join(' ')} disabled={isProcessing} />
                                <p className="text-xs text-muted-foreground">{t('imputer.separatorDesc')}</p>
                            </div>
                        )}
                    </div>
                )}
            </CardContent>
          </Card>
        )}

        <Button onClick={handleProcess} disabled={isProcessDisabled} className="w-full">
          {isProcessing ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Bot className="mr-2 h-5 w-5" />}
          {t('imputer.processBtn')}
        </Button>
      </CardContent>

      {hasResults && (
        <CardFooter className="flex-col space-y-4 items-stretch">
          <Card className="p-4 bg-secondary/30">
            <CardHeader className="p-0 pb-4">
              <CardTitle className="text-lg font-headline">{t('imputer.resultsTitle')}</CardTitle>
              <CardDescription>{t('imputer.resultsDesc')}</CardDescription>
            </CardHeader>
            <CardContent className="p-0">
                <div className="border rounded-lg overflow-hidden bg-background">
                    <div className="max-h-96 overflow-y-auto">
                      <Table>
                        <TableHeader className="sticky top-0 bg-background z-10">
                          <TableRow>
                            <TableHead className="w-[120px]">{t('imputer.tableHeaderAddress')}</TableHead>
                            <TableHead>{t('imputer.tableHeaderSuggestion')}</TableHead>
                            <TableHead>{t('imputer.tableHeaderContext')}</TableHead>
                            <TableHead className="w-[100px] text-right">
                                <div className="flex items-center justify-end space-x-2">
                                <Label htmlFor="select-all-suggestions" className="text-xs font-normal">{t('imputer.tableHeaderApply')}</Label>
                                <Checkbox
                                    id="select-all-suggestions"
                                    checked={allSuggestionsSelected}
                                    onCheckedChange={(checked) => handleSelectAllSuggestions(checked as boolean)}
                                />
                                </div>
                            </TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {suggestions.map((s) => (
                            <TableRow key={s.address}>
                              <TableCell className="font-mono text-xs">{s.sheetName}!<br/>{s.address}</TableCell>
                              <TableCell className="font-medium text-primary">{s.suggestion}</TableCell>
                              <TableCell>
                                {s.context?.map((c, index) => (
                                    <div key={index} className="text-xs">
                                        <span className="font-semibold text-muted-foreground">{c.label}:</span>
                                        <span className="ml-1 font-mono">{String(c.value)}</span>
                                    </div>
                                ))}
                              </TableCell>
                              <TableCell className="text-right">
                                <Checkbox checked={s.isChecked} onCheckedChange={(checked) => handleSuggestionCheckChange(s.address, checked as boolean)} />
                              </TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                </div>
            </CardContent>
          </Card>
          
          <Button onClick={handleDownload} variant="outline" className="w-full" disabled={isProcessing}>
              <Download className="mr-2 h-5 w-5" />
              {t('imputer.downloadBtn')}
          </Button>
        </CardFooter>
      )}
       {!hasResults && !isProcessing && (
        <CardFooter>
            <p className="text-sm text-muted-foreground w-full text-center">{t('imputer.noSuggestionsYet')}</p>
        </CardFooter>
      )}
    </Card>
  );
}
