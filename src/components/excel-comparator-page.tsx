
"use client";

import React, { useState, useCallback, ChangeEvent, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, GitCompareArrows, CheckCircle2, Loader2, KeyRound, Info, FileX2, FileSpreadsheet, GitMerge } from 'lucide-react';
import { useLanguage } from '@/context/language-context';
import { ScrollArea } from './ui/scroll-area';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { reconcileWorkbooks, compareWorkbooks, generateComparisonReportWorkbook } from '@/lib/excel-comparator';
import type { ComparisonReport, SheetComparisonResult } from '@/lib/excel-types';
import { RadioGroup, RadioGroupItem } from './ui/radio-group';

interface SheetSelection {
  [sheetName: string]: boolean;
}

const MAX_SHEETS_TO_COMPARE = 15;

interface ExcelComparatorPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function ExcelComparatorPage({ onProcessingChange, onFileStateChange }: ExcelComparatorPageProps) {
  const { t } = useLanguage();
  const [fileA, setFileA] = useState<File | null>(null);
  const [fileB, setFileB] = useState<File | null>(null);
  const [sheetNamesA, setSheetNamesA] = useState<string[]>([]);
  const [sheetNamesB, setSheetNamesB] = useState<string[]>([]);
  
  const [headerRow, setHeaderRow] = useState<number>(1);
  const [primaryKeyColumns, setPrimaryKeyColumns] = useState<string>('');
  const [selectedSheets, setSelectedSheets] = useState<SheetSelection>({});

  const [comparisonReport, setComparisonReport] = useState<ComparisonReport | null>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const { toast } = useToast();
  
  const [reconciledWb, setReconciledWb] = useState<XLSX.WorkBook | null>(null);
  const [mergeTarget, setMergeTarget] = useState<'A' | 'B'>('A');

  useEffect(() => {
    if (onProcessingChange) {
      onProcessingChange(isProcessing);
    }
  }, [isProcessing, onProcessingChange]);
  
  useEffect(() => {
    if (onFileStateChange) {
      onFileStateChange(fileA !== null || fileB !== null);
    }
  }, [fileA, fileB, onFileStateChange]);
  
  const commonSheets = useMemo(() => {
    const setA = new Set(sheetNamesA.map(s => s.toLowerCase()));
    return sheetNamesB.filter(s => setA.has(s.toLowerCase()));
  }, [sheetNamesA, sheetNamesB]);

  const handleFileChange = async (event: ChangeEvent<HTMLInputElement>, fileType: 'A' | 'B') => {
    const file = event.target.files?.[0];
    if (file) {
      if (!file.name.match(/\.(xlsx|xls|xlsm)$/)) {
        toast({ title: [t('toast.invalidFileType')].flat().join(' '), variant: 'destructive' });
        return;
      }
      if (fileType === 'A') {
        setFileA(file);
      } else {
        setFileB(file);
      }
      
      setIsProcessing(true);
      setComparisonReport(null);
      setReconciledWb(null);
      try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
        if (fileType === 'A') {
          setSheetNamesA(workbook.SheetNames);
        } else {
          setSheetNamesB(workbook.SheetNames);
        }
      } catch (error) {
        toast({ title: [t('toast.errorReadingFile')].flat().join(' '), variant: 'destructive' });
      } finally {
        setIsProcessing(false);
      }
    }
  };
  
  const handleSelectAllSheets = (checked: boolean) => {
    const newSelection: SheetSelection = {};
    commonSheets.forEach(name => {
      newSelection[name] = checked;
    });
    setSelectedSheets(newSelection);
  };
  
  const handleSheetSelectionChange = (sheetName: string, checked: boolean) => {
    setSelectedSheets(prev => ({ ...prev, [sheetName]: checked }));
  };

  const handleCompare = async () => {
    const sheetsToCompare = Object.entries(selectedSheets)
      .filter(([, isSelected]) => isSelected)
      .map(([sheetName]) => sheetName);
      
    if (!fileA || !fileB || sheetsToCompare.length === 0 || !primaryKeyColumns.trim()) {
      toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('comparator.toast.missingInfo')].flat().join(' '), variant: 'destructive' });
      return;
    }
    
    if (sheetsToCompare.length > MAX_SHEETS_TO_COMPARE) {
       toast({ title: [t('comparator.toast.limitTitle')].flat().join(' '), description: [t('comparator.toast.limitDesc', { limit: MAX_SHEETS_TO_COMPARE })].flat().join(' '), variant: 'destructive' });
      return;
    }

    setIsProcessing(true);
    setComparisonReport(null);
    setReconciledWb(null);
    try {
      const bufferA = await fileA.arrayBuffer();
      const wbA = XLSX.read(bufferA, { type: 'buffer', cellStyles: true });
      const bufferB = await fileB.arrayBuffer();
      const wbB = XLSX.read(bufferB, { type: 'buffer', cellStyles: true });
      
      const report = compareWorkbooks(wbA, wbB, sheetsToCompare, primaryKeyColumns, headerRow);
      setComparisonReport(report);
      
      toast({
        title: [t('toast.processingComplete')].flat().join(' '),
        description: [t('comparator.toast.success', { count: report.summary.sheetsWithDifferences.length })].flat().join(' '),
        action: <CheckCircle2 className="text-green-500" />
      });
      
    } catch (error) {
       const message = error instanceof Error ? error.message : String(error);
       toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: message, variant: 'destructive' });
    } finally {
      setIsProcessing(false);
    }
  };
  
  const handleReconcile = useCallback(() => {
    if (!fileA || !fileB || !comparisonReport) return;
    
    setIsProcessing(true);
    // Use a timeout to allow the UI to update to the "processing" state
    setTimeout(async () => {
        try {
            const bufferA = await fileA.arrayBuffer();
            const wbA = XLSX.read(bufferA, { type: 'buffer', cellStyles: true });
            const bufferB = await fileB.arrayBuffer();
            const wbB = XLSX.read(bufferB, { type: 'buffer', cellStyles: true });

            const newWb = reconcileWorkbooks(wbA, wbB, comparisonReport, mergeTarget);
            setReconciledWb(newWb);
            toast({
                title: [t('comparator.toast.reconcileSuccess')].flat().join(' '),
            });
        } catch(error) {
            console.error("Reconciliation error:", error);
            const message = error instanceof Error ? error.message : String(error);
            toast({ title: [t('comparator.toast.reconcileError')].flat().join(' '), description: message, variant: 'destructive' });
        } finally {
            setIsProcessing(false);
        }
    }, 50); // Small delay
  }, [fileA, fileB, comparisonReport, mergeTarget, toast, t]);
  
  const handleDownloadReconciled = () => {
    if (!reconciledWb || !fileA || !fileB) return;
    const baseFile = mergeTarget === 'A' ? fileA : fileB;
    const originalFileName = baseFile.name.substring(0, baseFile.name.lastIndexOf('.'));
    const fileName = `${originalFileName}_reconciled.xlsx`;
    XLSX.writeFile(reconciledWb, fileName);
  }
  
  const handleDownloadReport = () => {
    if (!comparisonReport || !fileA || !fileB) return;
    try {
      const reportWb = generateComparisonReportWorkbook(comparisonReport, fileA.name, fileB.name);
      XLSX.writeFile(reportWb, 'Excel_Comparison_Report.xlsx');
    } catch (error) {
       const message = error instanceof Error ? error.message : String(error);
       toast({ title: [t('toast.downloadError')].flat().join(' '), description: message, variant: 'destructive' });
    }
  };

  const renderSheetResult = (sheetName: string, result: SheetComparisonResult) => (
    <AccordionItem value={sheetName} key={sheetName}>
      <AccordionTrigger>{sheetName} - {[t('comparator.results.summary', { new: result.summary.newRows, deleted: result.summary.deletedRows, modified: result.summary.modifiedRows })].flat().join(' ')}</AccordionTrigger>
      <AccordionContent className="space-y-4">
        {result.modified.length > 0 && (
          <div>
            <h4 className="font-semibold text-amber-600">{[t('comparator.results.modifiedRows')].flat().join(' ')} ({result.modified.length})</h4>
             <ScrollArea className="h-60 border rounded-md mt-2">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>{[t('comparator.results.keyHeader')].flat().join(' ')}</TableHead>
                    <TableHead>{[t('comparator.results.columnHeader')].flat().join(' ')}</TableHead>
                    <TableHead>{[t('comparator.results.fileAHeader')].flat().join(' ')}</TableHead>
                    <TableHead>{[t('comparator.results.fileBHeader')].flat().join(' ')}</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {result.modified.map((mod, i) => (
                    mod.diffs.map((diff, j) => (
                       <TableRow key={`mod-${i}-${j}`}>
                          {j === 0 && <TableCell rowSpan={mod.diffs.length} className="align-top font-mono text-xs">{mod.key}</TableCell>}
                          <TableCell className="font-semibold">{diff.colName}</TableCell>
                          <TableCell className="text-muted-foreground">{String(diff.valueA)}</TableCell>
                          <TableCell className="text-primary font-medium">{String(diff.valueB)}</TableCell>
                       </TableRow>
                    ))
                  ))}
                </TableBody>
              </Table>
            </ScrollArea>
          </div>
        )}
         {result.new.length > 0 && (
          <div>
            <h4 className="font-semibold text-green-600">{[t('comparator.results.newRows')].flat().join(' ')} ({result.new.length})</h4>
            <ScrollArea className="h-40 border rounded-md mt-2">
              <Table>
                <TableBody>
                  {result.new.map((row, i) => <TableRow key={`new-${i}`}><TableCell>{row.join(', ')}</TableCell></TableRow>)}
                </TableBody>
              </Table>
            </ScrollArea>
          </div>
        )}
         {result.deleted.length > 0 && (
          <div>
            <h4 className="font-semibold text-red-600">{[t('comparator.results.deletedRows')].flat().join(' ')} ({result.deleted.length})</h4>
            <ScrollArea className="h-40 border rounded-md mt-2">
              <Table>
                <TableBody>
                  {result.deleted.map((row, i) => <TableRow key={`del-${i}`}><TableCell>{row.join(', ')}</TableCell></TableRow>)}
                </TableBody>
              </Table>
            </ScrollArea>
          </div>
        )}
      </AccordionContent>
    </AccordionItem>
  );

  const allSheetsSelected = commonSheets.length > 0 && commonSheets.every(name => selectedSheets[name]);

  return (
    <Card className="w-full max-w-4xl shadow-xl relative">
       {isProcessing && (
        <div className="absolute inset-0 bg-background/80 backdrop-blur-sm flex items-center justify-center z-10 rounded-lg">
          <Loader2 className="h-8 w-8 animate-spin" />
        </div>
      )}
      <CardHeader>
        <div className="flex items-center space-x-2 mb-2">
            <GitCompareArrows className="h-8 w-8 text-primary" />
            <CardTitle className="text-2xl font-headline">{[t('comparator.title')].flat().join(' ')}</CardTitle>
        </div>
        <CardDescription>{[t('comparator.description')].flat().join(' ')}</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-2">
              <Label htmlFor="file-a" className="font-semibold">{[t('comparator.fileA')].flat().join(' ')}</Label>
              <Input id="file-a" type="file" onChange={(e) => handleFileChange(e, 'A')} />
              <p className="text-xs text-muted-foreground">{[t('comparator.fileADesc')].flat().join(' ')}</p>
            </div>
             <div className="space-y-2">
              <Label htmlFor="file-b" className="font-semibold">{[t('comparator.fileB')].flat().join(' ')}</Label>
              <Input id="file-b" type="file" onChange={(e) => handleFileChange(e, 'B')} />
              <p className="text-xs text-muted-foreground">{[t('comparator.fileBDesc')].flat().join(' ')}</p>
            </div>
        </div>
        
        {fileA && fileB && (
            <Card className="p-4 bg-secondary/30">
                <CardHeader className="p-0 pb-4">
                    <CardTitle className="text-lg">{[t('comparator.configTitle')].flat().join(' ')}</CardTitle>
                </CardHeader>
                <CardContent className="p-0 space-y-4">
                    <div className="space-y-2">
                        <Label htmlFor="header-row" className="flex items-center space-x-2 font-medium">
                            <FileSpreadsheet className="h-4 w-4"/>
                            <span>{[t('common.headerRow')].flat().join(' ')}</span>
                        </Label>
                        <Input id="header-row" type="number" min="1" value={headerRow} onChange={e => setHeaderRow(parseInt(e.target.value, 10) || 1)} />
                        <p className="text-xs text-muted-foreground">{[t('comparator.headerRowDesc')].flat().join(' ')}</p>
                    </div>

                     <div className="space-y-2">
                        <Label htmlFor="primary-key" className="flex items-center space-x-2 font-medium">
                            <KeyRound className="h-4 w-4"/>
                            <span>{[t('comparator.primaryKey')].flat().join(' ')}</span>
                        </Label>
                        <Input id="primary-key" value={primaryKeyColumns} onChange={e => setPrimaryKeyColumns(e.target.value)} placeholder={[t('comparator.primaryKeyPlaceholder')].flat().join(' ')} />
                        <p className="text-xs text-muted-foreground">{[t('comparator.primaryKeyDesc')].flat().join(' ')}</p>
                    </div>

                    {commonSheets.length > 0 ? (
                        <div className="space-y-3">
                            <Label className="font-medium">{[t('comparator.selectSheets')].flat().join(' ')}</Label>
                             <Alert variant="default">
                                <Info className="h-4 w-4" />
                                <AlertDescription>{[t('comparator.sheetLimitNote', { limit: MAX_SHEETS_TO_COMPARE })].flat().join(' ')}</AlertDescription>
                             </Alert>
                             <div className="flex items-center space-x-2 p-2 border rounded-md bg-background">
                                <Checkbox
                                    id="select-all-sheets-comparator"
                                    checked={allSheetsSelected}
                                    onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                                />
                                <Label htmlFor="select-all-sheets-comparator" className="flex-grow">{[t('common.selectAll')].flat().join(' ')}</Label>
                                <span>{Object.values(selectedSheets).filter(Boolean).length} / {commonSheets.length}</span>
                             </div>
                             <ScrollArea className="h-40 border rounded-md p-2 bg-background">
                                {commonSheets.map(name => (
                                    <div key={name} className="flex items-center space-x-2 py-1">
                                        <Checkbox
                                            id={`sheet-comp-${name}`}
                                            checked={!!selectedSheets[name]}
                                            onCheckedChange={(checked) => handleSheetSelectionChange(name, !!checked)}
                                        />
                                        <Label htmlFor={`sheet-comp-${name}`} className="font-normal">{name}</Label>
                                    </div>
                                ))}
                             </ScrollArea>
                        </div>
                    ) : (
                        <Alert variant="destructive">
                            <FileX2 className="h-4 w-4" />
                            <AlertTitle>{[t('comparator.noCommonSheetsTitle')].flat().join(' ')}</AlertTitle>
                            <AlertDescription>{[t('comparator.noCommonSheetsDesc')].flat().join(' ')}</AlertDescription>
                        </Alert>
                    )}
                </CardContent>
            </Card>
        )}
        
        <Button onClick={handleCompare} disabled={isProcessing || !fileA || !fileB || Object.values(selectedSheets).filter(Boolean).length === 0 || !primaryKeyColumns.trim()} className="w-full">
            <GitCompareArrows className="mr-2 h-4 w-4" />
            {[t('comparator.compareBtn')].flat().join(' ')}
        </Button>
      </CardContent>

      {comparisonReport && (
        <CardFooter className="flex-col items-stretch space-y-4">
           <Card className="p-4 bg-secondary/30">
            <CardHeader className="p-0 pb-4">
                <CardTitle>{[t('comparator.results.title')].flat().join(' ')}</CardTitle>
            </CardHeader>
             <CardContent className="p-0">
                {comparisonReport.summary.sheetsWithDifferences.length > 0 || comparisonReport.summary.totalRowsFound > 0 ? (
                    <Accordion type="single" collapsible className="w-full">
                        {comparisonReport.summary.sheetsWithDifferences.map(sheetName => 
                            renderSheetResult(sheetName, comparisonReport.details[sheetName])
                        )}
                    </Accordion>
                ) : (
                    <p className="text-green-600 font-medium">{[t('comparator.results.noDiff')].flat().join(' ')}</p>
                )}
             </CardContent>
           </Card>
           
            {comparisonReport.summary.sheetsWithDifferences.length > 0 && (
                <Card className="p-4 bg-secondary/30">
                    <CardHeader className="p-0 pb-4">
                        <CardTitle className="flex items-center gap-2"><GitMerge className="h-5 w-5"/>{[t('comparator.reconciliation.title')].flat().join(' ')}</CardTitle>
                        <CardDescription>{[t('comparator.reconciliation.description')].flat().join(' ')}</CardDescription>
                    </CardHeader>
                    <CardContent className="p-0 space-y-4">
                        <RadioGroup value={mergeTarget} onValueChange={v => setMergeTarget(v as 'A' | 'B')}>
                            <Label className="font-medium">{[t('comparator.reconciliation.targetLabel')].flat().join(' ')}</Label>
                            <div className="flex items-center space-x-2 mt-2">
                                <RadioGroupItem value="A" id="target-a" />
                                <Label htmlFor="target-a" className="font-normal">{[t('comparator.reconciliation.targetA')].flat().join(' ')}</Label>
                            </div>
                            <div className="flex items-center space-x-2">
                                <RadioGroupItem value="B" id="target-b" />
                                <Label htmlFor="target-b" className="font-normal">{[t('comparator.reconciliation.targetB')].flat().join(' ')}</Label>
                            </div>
                        </RadioGroup>
                        <Button onClick={handleReconcile} disabled={isProcessing} className="w-full">
                            <GitMerge className="mr-2 h-4 w-4" />
                            {[t('comparator.reconciliation.reconcileBtn')].flat().join(' ')}
                        </Button>
                    </CardContent>
                </Card>
            )}

            {reconciledWb && (
                <Button onClick={handleDownloadReconciled} variant="secondary" className="w-full">
                    <Download className="mr-2 h-4 w-4"/>
                    {[t('comparator.reconciliation.downloadBtn')].flat().join(' ')}
                </Button>
            )}

           <Button onClick={handleDownloadReport} variant="outline" className="w-full">
              <Download className="mr-2 h-4 w-4"/>
              {[t('comparator.downloadReportBtn')].flat().join(' ')}
           </Button>
        </CardFooter>
      )}
    </Card>
  )
}
