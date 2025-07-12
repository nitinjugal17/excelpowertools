
"use client";

import React, { useState, useCallback, ChangeEvent, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, Combine, CheckCircle2, Loader2, ArrowRightLeft, FileInput, Eye, Lightbulb, FileSpreadsheet } from 'lucide-react';
import { useLanguage } from '@/context/language-context';
import { ScrollArea } from './ui/scroll-area';
import { mergeSheets, combineSheets } from '@/lib/excel-sheet-merger';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Markup } from '@/components/ui/markup';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';

interface SheetSelection {
  [sheetName: string]: boolean;
}

interface SheetMergerPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function SheetMergerPage({ onProcessingChange, onFileStateChange }: SheetMergerPageProps) {
  const { t } = useLanguage();
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const { toast } = useToast();

  // Mode state
  const [operationMode, setOperationMode] = useState<'merge' | 'combine'>('merge');

  // State for 'merge' mode
  const [sourceFile, setSourceFile] = useState<File | null>(null);
  const [destFile, setDestFile] = useState<File | null>(null);
  const [sourceSheetNames, setSourceSheetNames] = useState<string[]>([]);
  const [destSheetNames, setDestSheetNames] = useState<string[]>([]);
  const [selectedSheetsForMerge, setSelectedSheetsForMerge] = useState<SheetSelection>({});
  const [replaceExisting, setReplaceExisting] = useState<boolean>(true);
  const [comparisonData, setComparisonData] = useState<{ sheetName: string, source: any[][], dest: any[][] } | null>(null);
  const [isComparisonModalOpen, setIsComparisonModalOpen] = useState(false);
  
  // State for 'combine' mode
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheetsForCombine, setSelectedSheetsForCombine] = useState<SheetSelection>({});
  const [newSheetName, setNewSheetName] = useState<string>('Combined_Sheet');
  const [headerRow, setHeaderRow] = useState<number>(1);
  const [addSourceColumn, setAddSourceColumn] = useState<boolean>(true);
  const [columnsToIgnore, setColumnsToIgnore] = useState<string>('');

  useEffect(() => {
    if (onProcessingChange) {
      onProcessingChange(isProcessing);
    }
  }, [isProcessing, onProcessingChange]);

  useEffect(() => {
    if (onFileStateChange) {
      onFileStateChange(sourceFile !== null || destFile !== null || file !== null);
    }
  }, [sourceFile, destFile, file, onFileStateChange]);

  const handleSourceFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      setSourceFile(file);
      setIsProcessing(true);
      try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
        setSourceSheetNames(workbook.SheetNames);
        const initialSelection: SheetSelection = {};
        workbook.SheetNames.forEach(name => {
          initialSelection[name] = true;
        });
        setSelectedSheetsForMerge(initialSelection);
      } catch (error) {
        toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('merger.toast.errorReadingSource')].flat().join(' '), variant: 'destructive' });
        setSourceSheetNames([]);
        setSelectedSheetsForMerge({});
      } finally {
        setIsProcessing(false);
      }
    } else {
      setSourceFile(null);
      setSourceSheetNames([]);
      setSelectedSheetsForMerge({});
    }
  };

  const handleDestFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      setDestFile(file);
      setIsProcessing(true);
      try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
        setDestSheetNames(workbook.SheetNames);
      } catch (error) {
        toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('merger.toast.errorReadingDest')].flat().join(' '), variant: 'destructive' });
        setDestSheetNames([]);
      } finally {
        setIsProcessing(false);
      }
    } else {
      setDestFile(null);
      setDestSheetNames([]);
    }
  };

  const handleCombineFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
        setFile(selectedFile);
        setColumnsToIgnore('');
        setIsProcessing(true);
        try {
            const arrayBuffer = await selectedFile.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, {type: 'buffer'});
            setSheetNames(workbook.SheetNames);
            const initialSelection: SheetSelection = {};
            workbook.SheetNames.forEach(name => {
              initialSelection[name] = true;
            });
            setSelectedSheetsForCombine(initialSelection);
        } catch(error) {
            toast({ title: t('toast.errorReadingFile') as string, variant: 'destructive' });
            setSheetNames([]);
            setSelectedSheetsForCombine({});
        } finally {
            setIsProcessing(false);
        }
    } else {
        setFile(null);
        setSheetNames([]);
        setSelectedSheetsForCombine({});
    }
  };

  const handleSelectAllMerge = (checked: boolean) => {
    const newSelection: SheetSelection = {};
    sourceSheetNames.forEach(name => {
      newSelection[name] = checked;
    });
    setSelectedSheetsForMerge(newSelection);
  };
  
  const handleSheetSelectionMerge = (sheetName: string, checked: boolean) => {
    setSelectedSheetsForMerge(prev => ({ ...prev, [sheetName]: checked }));
  };

  const handleSelectAllCombine = (checked: boolean) => {
    const newSelection: SheetSelection = {};
    sheetNames.forEach(name => {
      newSelection[name] = checked;
    });
    setSelectedSheetsForCombine(newSelection);
  };
  
  const handleSheetSelectionCombine = (sheetName: string, checked: boolean) => {
    setSelectedSheetsForCombine(prev => ({ ...prev, [sheetName]: checked }));
  };
  
  const handleCompare = async (sheetName: string) => {
    if (!sourceFile || !destFile) return;

    setIsProcessing(true);
    try {
      const sourceBuffer = await sourceFile.arrayBuffer();
      const sourceWb = XLSX.read(sourceBuffer, { type: 'buffer' });
      const sourceSheet = sourceWb.Sheets[sheetName];
      const sourceData = sourceSheet ? XLSX.utils.sheet_to_json<any[]>(sourceSheet, { header: 1 }).slice(0, 20) : [];

      const destBuffer = await destFile.arrayBuffer();
      const destWb = XLSX.read(destBuffer, { type: 'buffer' });
      const destSheet = destWb.Sheets[sheetName];
      const destData = destSheet ? XLSX.utils.sheet_to_json<any[]>(destSheet, { header: 1 }).slice(0, 20) : [];

      setComparisonData({ sheetName, source: sourceData, dest: destData });
      setIsComparisonModalOpen(true);

    } catch (error) {
      toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('merger.toast.errorReadingSheets')].flat().join(' '), variant: 'destructive' });
    } finally {
      setIsProcessing(false);
    }
  };


  const handleMergeAndDownload = async () => {
    const sheetsToMerge = Object.entries(selectedSheetsForMerge)
      .filter(([, isSelected]) => isSelected)
      .map(([sheetName]) => sheetName);

    if (!sourceFile || !destFile || sheetsToMerge.length === 0) {
      toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('merger.toast.missingInfo')].flat().join(' '), variant: 'destructive' });
      return;
    }

    setIsProcessing(true);
    try {
      const sourceBuffer = await sourceFile.arrayBuffer();
      const destBuffer = await destFile.arrayBuffer();
      const sourceWb = XLSX.read(sourceBuffer, { type: 'buffer', cellStyles: true, bookFiles: true, bookVBA: true });
      let destWb = XLSX.read(destBuffer, { type: 'buffer', cellStyles: true, bookFiles: true, bookVBA: true });
      
      destWb = mergeSheets(sourceWb, destWb, sheetsToMerge, replaceExisting);
      
      const originalDestFileName = destFile.name.substring(0, destFile.name.lastIndexOf('.'));
      XLSX.writeFile(destWb, `${originalDestFileName}_merged.xlsx`, { compression: true, cellStyles: true, bookType: 'xlsx' });
      
      toast({
        title: [t('merger.toast.successTitle')].flat().join(' '),
        description: [t('merger.toast.successDesc', { count: sheetsToMerge.length })].flat().join(' '),
        action: <CheckCircle2 className="text-green-500" />
      });

    } catch (error) {
      console.error(error);
      const errorMessage = error instanceof Error ? error.message : [t('merger.toast.errorGeneral')].flat().join(' ');
      toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: errorMessage, variant: 'destructive' });
    } finally {
      setIsProcessing(false);
    }
  };
  
  const handleCombineAndDownload = async () => {
    const sheetsToCombine = Object.entries(selectedSheetsForCombine)
      .filter(([, isSelected]) => isSelected)
      .map(([sheetName]) => sheetName);
      
    if (!file || sheetsToCombine.length === 0) {
      toast({ title: t('toast.missingInfo') as string, description: [t('merger.toast.missingInfoCombine')].flat().join(' '), variant: 'destructive' });
      return;
    }
    
    setIsProcessing(true);
    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellDates: true });
        
        const newWb = combineSheets(workbook, sheetsToCombine, headerRow, newSheetName, addSourceColumn, columnsToIgnore);
        
        const finalFileName = (newSheetName.trim() || 'Combined') + '.xlsx';
        XLSX.writeFile(newWb, finalFileName, { compression: true, bookType: 'xlsx' });

        toast({
            title: t('merger.toast.combineSuccessTitle') as string,
            description: t('merger.toast.combineSuccessDesc', { count: sheetsToCombine.length }) as string,
            action: <CheckCircle2 className="text-green-500" />
        });

    } catch (error) {
        console.error(error);
        const errorMessage = error instanceof Error ? error.message : t('merger.toast.errorGeneral') as string;
        toast({ title: t('toast.errorReadingFile') as string, description: errorMessage, variant: 'destructive' });
    } finally {
        setIsProcessing(false);
    }
  };

  const renderComparisonTable = (data: any[][], caption: string) => {
    if (!data || data.length === 0) {
      return <p className="text-sm text-muted-foreground">{[t('merger.noContent')].flat().join(' ')}</p>;
    }
    const headers = data[0];
    const rows = data.slice(1);
    return (
      <div>
        <h3 className="font-semibold mb-2">{caption}</h3>
        <ScrollArea className="h-96 border rounded-md">
          <Table>
            <TableHeader>
              <TableRow>
                {headers.map((header, index) => <TableHead key={index}>{header}</TableHead>)}
              </TableRow>
            </TableHeader>
            <TableBody>
              {rows.map((row, rowIndex) => (
                <TableRow key={rowIndex}>
                  {row.map((cell, cellIndex) => <TableCell key={cellIndex}>{cell}</TableCell>)}
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </ScrollArea>
      </div>
    );
  };
  
  const allSheetsSelectedMerge = sourceSheetNames.length > 0 && sourceSheetNames.every(name => selectedSheetsForMerge[name]);
  const allSheetsSelectedCombine = sheetNames.length > 0 && sheetNames.every(name => selectedSheetsForCombine[name]);

  return (
    <>
      <Dialog open={isComparisonModalOpen} onOpenChange={setIsComparisonModalOpen}>
        <DialogContent className="max-w-6xl">
          <DialogHeader>
            <DialogTitle>{[t('merger.comparisonTitle', {sheetName: comparisonData?.sheetName || ''})].flat().join(' ')}</DialogTitle>
          </DialogHeader>
          {comparisonData && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {renderComparisonTable(comparisonData.source, [t('merger.sourceSheet')].flat().join(' '))}
              {renderComparisonTable(comparisonData.dest, [t('merger.destinationSheet')].flat().join(' '))}
            </div>
          )}
        </DialogContent>
      </Dialog>

      <Card className="w-full max-w-lg md:max-w-xl lg:max-w-2xl shadow-xl relative">
        {isProcessing && (
          <div className="absolute inset-0 bg-background/80 backdrop-blur-sm flex flex-col items-center justify-center z-10 rounded-lg space-y-4">
            <Loader2 className="h-8 w-8 animate-spin" />
            <p className="text-lg">{[t('common.processing')].flat().join(' ')}</p>
          </div>
        )}
        <CardHeader>
          <div className="flex items-center space-x-2 mb-2">
            <Combine className="h-8 w-8 text-primary" />
            <CardTitle className="text-2xl font-headline">{[t('merger.title')].flat().join(' ')}</CardTitle>
          </div>
          <CardDescription className="font-body">{[t('merger.description')].flat().join(' ')}</CardDescription>
        </CardHeader>
        <CardContent className="space-y-6">
            <RadioGroup value={operationMode} onValueChange={(v) => setOperationMode(v as 'merge' | 'combine')} className="space-y-2">
                <Label htmlFor="mode-merge" className="flex items-start space-x-3 p-4 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer">
                    <RadioGroupItem value="merge" id="mode-merge" className="mt-1"/>
                    <div className="grid gap-1.5">
                        <span className="font-semibold">{[t('merger.modeMerge')].flat().join(' ')}</span>
                        <p className="text-xs text-muted-foreground">{[t('merger.modeMergeDesc')].flat().join(' ')}</p>
                    </div>
                </Label>
                <Label htmlFor="mode-combine" className="flex items-start space-x-3 p-4 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer">
                    <RadioGroupItem value="combine" id="mode-combine" className="mt-1"/>
                    <div className="grid gap-1.5">
                        <span className="font-semibold">{[t('merger.modeCombine')].flat().join(' ')}</span>
                        <p className="text-xs text-muted-foreground">{[t('merger.modeCombineDesc')].flat().join(' ')}</p>
                    </div>
                </Label>
            </RadioGroup>

          {operationMode === 'merge' && (
            <div className="space-y-6 pt-4 border-t">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-start">
                    <div className="space-y-2">
                    <Label htmlFor="source-file" className="flex items-center space-x-2 text-sm font-medium">
                        <FileInput className="h-5 w-5" />
                        <span>{[t('merger.sourceFile')].flat().join(' ')}</span>
                    </Label>
                    <Input id="source-file" type="file" onChange={handleSourceFileChange} disabled={isProcessing} className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20" />
                    <p className="text-xs text-muted-foreground">{[t('merger.sourceFileDesc')].flat().join(' ')}</p>
                    </div>
                    <div className="space-y-2">
                    <Label htmlFor="dest-file" className="flex items-center space-x-2 text-sm font-medium">
                        <ArrowRightLeft className="h-5 w-5" />
                        <span>{[t('merger.destFile')].flat().join(' ')}</span>
                    </Label>
                    <Input id="dest-file" type="file" onChange={handleDestFileChange} disabled={isProcessing} className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20" />
                    <p className="text-xs text-muted-foreground">{[t('merger.destFileDesc')].flat().join(' ')}</p>
                    </div>
                </div>

                {sourceSheetNames.length > 0 && (
                    <div className="space-y-3">
                    <Label className="flex items-center space-x-2 text-sm font-medium mb-2">
                        <span>{[t('merger.selectSheets')].flat().join(' ')}</span>
                    </Label>
                    <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
                        <Checkbox
                        id="select-all-sheets-merger"
                        checked={allSheetsSelectedMerge}
                        onCheckedChange={(checked) => handleSelectAllMerge(checked as boolean)}
                        disabled={isProcessing}
                        />
                        <Label htmlFor="select-all-sheets-merger" className="text-sm font-medium flex-grow">
                        {[t('common.selectAll')].flat().join(' ')} ({[t('common.selectedCount', {selected: Object.values(selectedSheetsForMerge).filter(Boolean).length, total: sourceSheetNames.length})].flat().join(' ')})
                        </Label>
                    </div>
                    <Card className="max-h-60 overflow-y-auto p-3 bg-background">
                        <div className="space-y-2">
                        {sourceSheetNames.map(name => {
                            const sheetExistsInDest = destSheetNames.some(dn => dn.toLowerCase() === name.toLowerCase());
                            return (
                            <div key={name} className="flex items-center justify-between space-x-2 p-1">
                                <div className="flex items-center space-x-2">
                                <Checkbox
                                    id={`sheet-merger-${name}`}
                                    checked={selectedSheetsForMerge[name] || false}
                                    onCheckedChange={(checked) => handleSheetSelectionMerge(name, checked as boolean)}
                                    disabled={isProcessing}
                                />
                                <Label htmlFor={`sheet-merger-${name}`} className="text-sm font-normal">{name}</Label>
                                </div>
                                {sheetExistsInDest && (
                                <Button variant="outline" size="sm" onClick={() => handleCompare(name)} disabled={isProcessing}>
                                    <Eye className="mr-2 h-4 w-4" />
                                    {[t('merger.compare')].flat().join(' ')}
                                </Button>
                                )}
                            </div>
                            );
                        })}
                        </div>
                    </Card>
                    </div>
                )}
                
                <div className="flex items-start space-x-2">
                    <Checkbox
                    id="replace-existing"
                    checked={replaceExisting}
                    onCheckedChange={(checked) => setReplaceExisting(checked as boolean)}
                    disabled={isProcessing}
                    className="mt-1"
                    />
                    <div className="grid gap-1.5 leading-none">
                        <Label htmlFor="replace-existing" className="font-normal">{[t('merger.replaceLabel')].flat().join(' ')}</Label>
                        <p className="text-xs text-muted-foreground">{[t('merger.replaceDesc')].flat().join(' ')}</p>
                    </div>
                </div>

                <Alert>
                    <Lightbulb className="h-4 w-4" />
                    <AlertTitle>{[t('merger.rowLevelMergeNoteTitle')].flat().join(' ')}</AlertTitle>
                    <AlertDescription>
                        <Markup text={[t('merger.rowLevelMergeNoteDesc')].flat().join(' ')} />
                    </AlertDescription>
                </Alert>
            </div>
          )}

          {operationMode === 'combine' && (
            <div className="space-y-6 pt-4 border-t">
                <div className="space-y-2">
                    <Label htmlFor="combine-file" className="flex items-center space-x-2 text-sm font-medium">
                        <FileInput className="h-5 w-5" />
                        <span>{[t('merger.uploadStepCombine')].flat().join(' ')}</span>
                    </Label>
                    <Input id="combine-file" type="file" onChange={handleCombineFileChange} disabled={isProcessing} className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20" />
                </div>
                {sheetNames.length > 0 && (
                  <>
                    <div className="space-y-3">
                      <Label className="flex items-center space-x-2 text-sm font-medium mb-2">
                        <span>{[t('merger.selectSheetsCombine')].flat().join(' ')}</span>
                      </Label>
                      <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
                          <Checkbox
                          id="select-all-sheets-combine"
                          checked={allSheetsSelectedCombine}
                          onCheckedChange={(checked) => handleSelectAllCombine(checked as boolean)}
                          disabled={isProcessing}
                          />
                          <Label htmlFor="select-all-sheets-combine" className="text-sm font-medium flex-grow">
                          {[t('common.selectAll')].flat().join(' ')} ({[t('common.selectedCount', {selected: Object.values(selectedSheetsForCombine).filter(Boolean).length, total: sheetNames.length})].flat().join(' ')})
                          </Label>
                      </div>
                      <Card className="max-h-60 overflow-y-auto p-3 bg-background">
                          <div className="space-y-2">
                          {sheetNames.map(name => (
                              <div key={`combine-${name}`} className="flex items-center space-x-2 p-1">
                                  <Checkbox
                                      id={`sheet-combine-${name}`}
                                      checked={selectedSheetsForCombine[name] || false}
                                      onCheckedChange={(checked) => handleSheetSelectionCombine(name, checked as boolean)}
                                      disabled={isProcessing}
                                  />
                                  <Label htmlFor={`sheet-combine-${name}`} className="text-sm font-normal">{name}</Label>
                              </div>
                          ))}
                          </div>
                      </Card>
                    </div>
                    <div className="space-y-2">
                        <Label htmlFor="header-row" className="flex items-center space-x-2 text-sm font-medium">
                            <FileSpreadsheet className="h-4 w-4" />
                            <span>{[t('merger.headerRowStep')].flat().join(' ')}</span>
                        </Label>
                        <Input id="header-row" type="number" min="1" value={headerRow} onChange={(e) => setHeaderRow(parseInt(e.target.value, 10) || 1)} disabled={isProcessing}/>
                        <p className="text-xs text-muted-foreground">{[t('merger.headerRowDesc')].flat().join(' ')}</p>
                    </div>
                     <div className="space-y-2">
                        <Label htmlFor="new-sheet-name" className="text-sm font-medium">{[t('merger.newSheetName')].flat().join(' ')}</Label>
                        <Input id="new-sheet-name" value={newSheetName} onChange={(e) => setNewSheetName(e.target.value)} placeholder={[t('merger.newSheetNamePlaceholder')].flat().join(' ')} disabled={isProcessing}/>
                    </div>
                    <div className="flex items-start space-x-2">
                        <Checkbox
                        id="add-source-column"
                        checked={addSourceColumn}
                        onCheckedChange={(checked) => setAddSourceColumn(checked as boolean)}
                        disabled={isProcessing}
                        className="mt-1"
                        />
                        <div className="grid gap-1.5 leading-none">
                            <Label htmlFor="add-source-column" className="font-normal">{[t('merger.addSourceColumn')].flat().join(' ')}</Label>
                            <p className="text-xs text-muted-foreground">{[t('merger.addSourceColumnDesc')].flat().join(' ')}</p>
                        </div>
                    </div>
                    <div className="space-y-2">
                        <Label htmlFor="columns-to-ignore" className="text-sm font-medium">{[t('merger.ignoreColumns')].flat().join(' ')}</Label>
                        <Input 
                            id="columns-to-ignore" 
                            value={columnsToIgnore} 
                            onChange={(e) => setColumnsToIgnore(e.target.value)} 
                            placeholder={[t('merger.ignoreColumnsPlaceholder')].flat().join(' ')} 
                            disabled={isProcessing}
                        />
                        <p className="text-xs text-muted-foreground">{[t('merger.ignoreColumnsDesc')].flat().join(' ')}</p>
                    </div>
                  </>
                )}
            </div>
          )}

        </CardContent>
        <CardFooter>
            {operationMode === 'merge' ? (
                <Button onClick={handleMergeAndDownload} disabled={isProcessing || !sourceFile || !destFile || Object.values(selectedSheetsForMerge).filter(Boolean).length === 0} className="w-full">
                    {isProcessing ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-5 w-5" />}
                    {[t('merger.processBtn')].flat().join(' ')}
                </Button>
            ) : (
                <Button onClick={handleCombineAndDownload} disabled={isProcessing || !file || Object.values(selectedSheetsForCombine).filter(Boolean).length === 0} className="w-full">
                    {isProcessing ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-5 w-5" />}
                    {[t('merger.processBtnCombine')].flat().join(' ')}
                </Button>
            )}
        </CardFooter>
      </Card>
    </>
  );
}
