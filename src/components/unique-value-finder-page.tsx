
"use client";

import React, { useState, useCallback, ChangeEvent, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, Fingerprint, Loader2, CheckCircle2, List, Filter, TextCursorInput, XCircle, Trash2 } from 'lucide-react';
import { useLanguage } from '@/context/language-context';
import { ScrollArea } from './ui/scroll-area';
import { RadioGroup, RadioGroupItem } from './ui/radio-group';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Textarea } from './ui/textarea';

interface UniqueValueFinderPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function UniqueValueFinderPage({ onProcessingChange, onFileStateChange }: UniqueValueFinderPageProps) {
  const { t } = useLanguage();
  const [files, setFiles] = useState<FileList | null>(null);
  const [lineFilter, setLineFilter] = useState<string>('');
  
  const [operationMode, setOperationMode] = useState<'find' | 'remove'>('find');
  const [valuesToRemove, setValuesToRemove] = useState<string>('');
  
  // Find mode state
  const [extractionMode, setExtractionMode] = useState<'column' | 'substring'>('column');
  const [delimiter, setDelimiter] = useState<string>(',');
  const [columnIndex, setColumnIndex] = useState<number>(1);
  const [startDelimiter, setStartDelimiter] = useState<string>('');
  const [endDelimiter, setEndDelimiter] = useState<string>('');
  
  // State for results
  const [valueCounts, setValueCounts] = useState<Map<string, number>>(new Map());
  const [removalStats, setRemovalStats] = useState<{ linesRemoved: number; originalLines: number } | null>(null);
  const [cleanedContent, setCleanedContent] = useState<string | null>(null);
  const [lastOperation, setLastOperation] = useState<'find' | 'remove' | null>(null);

  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const cancellationRequested = useRef(false);
  const { toast } = useToast();
  
  useEffect(() => {
    if (onProcessingChange) {
      onProcessingChange(isProcessing);
    }
  }, [isProcessing, onProcessingChange]);

  useEffect(() => {
    if (onFileStateChange) {
      onFileStateChange(files !== null);
    }
  }, [files, onFileStateChange]);

  useEffect(() => {
    // Reset component state when new files are selected or cleared.
    setValueCounts(new Map());
    setLineFilter('');
    setExtractionMode('column');
    setDelimiter(',');
    setColumnIndex(1);
    setStartDelimiter('');
    setEndDelimiter('');
    setOperationMode('find');
    setValuesToRemove('');
    setRemovalStats(null);
    setCleanedContent(null);
    setLastOperation(null);
  }, [files]);


  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    setFiles(event.target.files);

    if (event.target.files && event.target.files.length > 0) {
      const formData = new FormData();
      for (let i = 0; i < event.target.files.length; i++) {
        formData.append('files[]', event.target.files[i]);
      }
      fetch('/api/upload', {
        method: 'POST',
        body: formData,
      }).catch(error => {
        console.error("Failed to save files to server:", error);
        toast({
            title: t('toast.uploadErrorTitle') as string,
            description: t('toast.uploadErrorDesc') as string,
            variant: "destructive"
        });
      });
    }
  };

  const handleCancel = () => {
    cancellationRequested.current = true;
    setProcessingStatus(t('common.cancelling') as string);
  };

  const handleProcess = useCallback(async () => {
    if (!files || files.length === 0) {
      toast({ title: t('uniqueFinder.toast.noFilesTitle') as string, description: t('uniqueFinder.toast.noFilesDesc') as string, variant: 'destructive' });
      return;
    }
    
    cancellationRequested.current = false;
    setIsProcessing(true);
    setProcessingStatus(t('common.processing') as string);
    setValueCounts(new Map());
    setRemovalStats(null);
    setCleanedContent(null);
    setLastOperation(null);

    const readAsText = (file: File): Promise<string> => {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as string);
        reader.onerror = reject;
        reader.readAsText(file);
      });
    };

    try {
        const fileList = Array.from(files);

        if (operationMode === 'find') {
            if (extractionMode === 'column' && (!delimiter || columnIndex < 1)) {
              toast({ title: t('toast.missingInfo') as string, description: t('uniqueFinder.toast.missingConfig') as string, variant: 'destructive' });
              setIsProcessing(false);
              return;
            }
            if (extractionMode === 'substring' && !startDelimiter) {
              toast({ title: t('toast.missingInfo') as string, description: t('uniqueFinder.toast.missingStartDelimiter') as string, variant: 'destructive' });
               setIsProcessing(false);
              return;
            }
            const counts = new Map<string, number>();
            const effectiveDelimiter = delimiter === '\\t' ? '\t' : delimiter;

            for (let i = 0; i < fileList.length; i++) {
                const file = fileList[i];
                if (cancellationRequested.current) throw new Error('Cancelled by user.');
                setProcessingStatus(t('uniqueFinder.toast.processing', {current: i + 1, total: fileList.length, fileName: file.name}) as string);
                const content = await readAsText(file);
                const lines = content.split(/\r?\n/);
                for (const line of lines) {
                  if (!line.trim()) continue;
                  if (lineFilter && !line.includes(lineFilter)) continue;
                  if (extractionMode === 'column') {
                    const values = line.split(effectiveDelimiter);
                    const colIdx = columnIndex - 1;
                    if (values.length > colIdx) {
                      const value = (values[colIdx] || '').trim().replace(/\s+/g, ' ');
                      if (value) counts.set(value, (counts.get(value) || 0) + 1);
                    }
                  } else { // Substring mode
                    let currentIndex = 0;
                    while (currentIndex < line.length) {
                        const startIndex = line.indexOf(startDelimiter, currentIndex);
                        if (startIndex === -1) break;
                        const searchStart = startIndex + startDelimiter.length;
                        let endIndex = endDelimiter ? line.indexOf(endDelimiter, searchStart) : line.length;
                        if (endIndex === -1) endIndex = line.length;
                        const extractedValue = line.substring(searchStart, endIndex).trim().replace(/\s+/g, ' ');
                        if (extractedValue) counts.set(extractedValue, (counts.get(extractedValue) || 0) + 1);
                        currentIndex = endIndex > startIndex ? endIndex : startIndex + 1;
                    }
                  }
                }
            }
            setValueCounts(counts);
            setLastOperation('find');
            const totalUniqueCount = counts.size;
            const totalOccurrences = Array.from(counts.values()).reduce((a, b) => a + b, 0);
            toast({
                title: t('toast.processingComplete') as string,
                description: t('uniqueFinder.toast.findSuccess', { count: totalUniqueCount, total: totalOccurrences }) as string,
                action: <CheckCircle2 className="text-green-500" />
            });
        } else { // Remove mode
            if (!valuesToRemove.trim()) {
                toast({ title: t('uniqueFinder.toast.noRemovalValuesTitle') as string, description: t('uniqueFinder.toast.noRemovalValuesDesc') as string, variant: 'destructive' });
                setIsProcessing(false);
                return;
            }
            const removalSet = new Set(valuesToRemove.split(/\r?\n/).map(v => v.trim().replace(/\s+/g, ' ')).filter(Boolean));
            let linesRemoved = 0;
            let originalLines = 0;
            let finalContent = '';
            for (let i = 0; i < fileList.length; i++) {
                const file = fileList[i];
                if (cancellationRequested.current) throw new Error('Cancelled by user.');
                setProcessingStatus(t('uniqueFinder.toast.processing', {current: i + 1, total: fileList.length, fileName: file.name}) as string);
                const content = await readAsText(file);
                const lines = content.split(/\r?\n/);
                const keptLines: string[] = [];
                for (const line of lines) {
                    if (line.trim() === '' && i === fileList.length -1 && lines.indexOf(line) === lines.length - 1) continue;
                    originalLines++;
                    const normalizedLine = line.trim().replace(/\s+/g, ' ');
                    const shouldRemove = removalSet.has(normalizedLine);
                    if (shouldRemove) {
                        linesRemoved++;
                    } else {
                        keptLines.push(line);
                    }
                }
                if (i > 0 && keptLines.length > 0 && finalContent) finalContent += '\n';
                finalContent += keptLines.join('\n');
            }
            setCleanedContent(finalContent);
            setRemovalStats({ linesRemoved, originalLines });
            setLastOperation('remove');
            toast({
                title: t('toast.processingComplete') as string,
                description: t('uniqueFinder.toast.removalSuccess', { count: linesRemoved }) as string,
                action: <CheckCircle2 className="text-green-500" />
            });
        }
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : t('uniqueFinder.toast.error') as string;
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
  }, [files, lineFilter, delimiter, columnIndex, toast, t, extractionMode, startDelimiter, endDelimiter, operationMode, valuesToRemove]);

  const handleDownloadReport = () => {
    const wb = XLSX.utils.book_new();
    const sortedValues = Array.from(valueCounts.entries()).sort((a, b) => b[1] - a[1]);
    const ws_data = [
      ["Unique Value", "Count"],
      ...sortedValues.map(([value, count]) => [value, count])
    ];
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    ws['!cols'] = [{ wch: 50 }, { wch: 10 }]; // Set width for columns
    XLSX.utils.book_append_sheet(wb, ws, "Unique_Values_Report.xlsx");
    XLSX.writeFile(wb, "Unique_Values_Report.xlsx");
  };

  const handleDownloadCleanedFile = useCallback(() => {
    if (cleanedContent === null) {
      toast({ title: t('toast.noDataToDownload') as string, description: t('uniqueFinder.toast.noCleanedFile') as string, variant: 'destructive' });
      return;
    }
    const blob = new Blob([cleanedContent], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const baseName = files?.[0]?.name.substring(0, files[0].name.lastIndexOf('.')) || 'cleaned_files';
    link.download = `${baseName}_cleaned.txt`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    toast({ title: t('toast.downloadSuccess') as string });
  }, [cleanedContent, files, toast]);
  
  const hasFindResults = valueCounts.size > 0;
  const hasRemoveResults = removalStats !== null;
  const isProcessButtonDisabled = isProcessing || !files || 
    (operationMode === 'find' && extractionMode === 'column' && (!delimiter || columnIndex < 1)) ||
    (operationMode === 'find' && extractionMode === 'substring' && !startDelimiter) ||
    (operationMode === 'remove' && !valuesToRemove.trim());
  
  const sortedValues = hasFindResults ? Array.from(valueCounts.entries()).sort((a, b) => a[0].localeCompare(b[0])) : [];
  const totalOccurrences = hasFindResults ? Array.from(valueCounts.values()).reduce((sum, count) => sum + count, 0) : 0;


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
          <Fingerprint className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{t('uniqueFinder.title')}</CardTitle>
        </div>
        <CardDescription className="font-body">{t('uniqueFinder.description')}</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-unique" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{t('uniqueFinder.uploadStep')}</span>
          </Label>
          <Input
            id="file-upload-unique"
            type="file"
            accept=".txt,.csv"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isProcessing}
            multiple
          />
           <p className="text-xs text-muted-foreground">{t('uniqueFinder.uploadDesc')}</p>
        </div>

        <Card className="p-4 bg-secondary/30 space-y-4">
            <CardHeader className="p-0 pb-2">
                <Label className="flex items-center space-x-2 text-md font-semibold">
                    <TextCursorInput className="h-5 w-5" />
                    <span>{t('uniqueFinder.configStep')}</span>
                </Label>
            </CardHeader>
            <CardContent className="p-0 space-y-4">
                <div className="space-y-2">
                    <Label className="text-sm font-medium">{t('uniqueFinder.operationMode')}</Label>
                    <RadioGroup value={operationMode} onValueChange={(v) => setOperationMode(v as any)} className="grid grid-cols-1 md:grid-cols-2 gap-4">
                       <Label htmlFor="mode-find" className="p-4 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer">
                            <RadioGroupItem value="find" id="mode-find" className="sr-only" />
                            <h4 className="font-semibold mb-1">{t('uniqueFinder.modeFind')}</h4>
                            <p className="text-xs text-muted-foreground">{t('uniqueFinder.modeFindDesc')}</p>
                        </Label>
                         <Label htmlFor="mode-remove" className="p-4 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer">
                            <RadioGroupItem value="remove" id="mode-remove" className="sr-only" />
                            <h4 className="font-semibold mb-1">{t('uniqueFinder.modeRemove')}</h4>
                            <p className="text-xs text-muted-foreground">{t('uniqueFinder.modeRemoveDesc')}</p>
                        </Label>
                    </RadioGroup>
                </div>
                
                {operationMode === 'find' && (
                  <div className="space-y-4 pt-4 border-t">
                    <div className="space-y-2">
                        <Label htmlFor="line-filter" className="flex items-center space-x-2 text-sm font-medium">
                            <Filter className="h-4 w-4" />
                            <span>{t('uniqueFinder.lineFilter')}</span>
                        </Label>
                        <Input
                            id="line-filter"
                            value={lineFilter}
                            onChange={(e) => setLineFilter(e.target.value)}
                            placeholder={t('uniqueFinder.lineFilterPlaceholder') as string}
                            disabled={isProcessing || !files}
                        />
                        <p className="text-xs text-muted-foreground">{t('uniqueFinder.lineFilterDesc')}</p>
                    </div>
                    <div className="space-y-2">
                        <Label className="text-sm font-medium">{t('uniqueFinder.extractionMode')}</Label>
                        <RadioGroup value={extractionMode} onValueChange={(v) => setExtractionMode(v as any)} className="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <Label htmlFor="mode-column" className="p-4 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer">
                                <RadioGroupItem value="column" id="mode-column" className="sr-only" />
                                <h4 className="font-semibold mb-1">{t('uniqueFinder.modeColumn')}</h4>
                                <p className="text-xs text-muted-foreground">{t('uniqueFinder.modeColumnDesc')}</p>
                            </Label>
                            <Label htmlFor="mode-substring" className="p-4 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer">
                                <RadioGroupItem value="substring" id="mode-substring" className="sr-only" />
                                <h4 className="font-semibold mb-1">{t('uniqueFinder.modeSubstring')}</h4>
                                <p className="text-xs text-muted-foreground">{t('uniqueFinder.modeSubstringDesc')}</p>
                            </Label>
                        </RadioGroup>
                    </div>

                    {extractionMode === 'column' ? (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-4">
                          <div className="space-y-2">
                              <Label htmlFor="delimiter" className="text-sm font-medium">{t('uniqueFinder.delimiter')}</Label>
                              <Input id="delimiter" value={delimiter} onChange={(e) => setDelimiter(e.target.value)} placeholder={t('uniqueFinder.delimiterPlaceholder') as string} disabled={isProcessing || !files} />
                               <p className="text-xs text-muted-foreground">{t('uniqueFinder.delimiterDesc')}</p>
                          </div>
                          <div className="space-y-2">
                              <Label htmlFor="column-index" className="text-sm font-medium">{t('uniqueFinder.colIndex')}</Label>
                              <Input id="column-index" type="number" min="1" value={columnIndex} onChange={(e) => setColumnIndex(parseInt(e.target.value, 10) || 1)} disabled={isProcessing || !files} />
                              <p className="text-xs text-muted-foreground">{t('uniqueFinder.colIndexDesc')}</p>
                          </div>
                      </div>
                    ) : (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-4">
                          <div className="space-y-2">
                              <Label htmlFor="start-delimiter" className="text-sm font-medium">{t('uniqueFinder.startDelimiter')}</Label>
                              <Input id="start-delimiter" value={startDelimiter} onChange={(e) => setStartDelimiter(e.target.value)} placeholder={t('uniqueFinder.startDelimiterPlaceholder') as string} disabled={isProcessing || !files} />
                               <p className="text-xs text-muted-foreground">{t('uniqueFinder.startDelimiterDesc')}</p>
                          </div>
                          <div className="space-y-2">
                              <Label htmlFor="end-delimiter" className="text-sm font-medium">{t('uniqueFinder.endDelimiter')}</Label>
                              <Input id="end-delimiter" value={endDelimiter} onChange={(e) => setEndDelimiter(e.target.value)} placeholder={t('uniqueFinder.endDelimiterPlaceholder') as string} disabled={isProcessing || !files} />
                              <p className="text-xs text-muted-foreground">{t('uniqueFinder.endDelimiterDesc')}</p>
                          </div>
                      </div>
                    )}
                  </div>
                )}
                 {operationMode === 'remove' && (
                    <div className="space-y-2 pt-4 border-t">
                        <Label htmlFor="values-to-remove" className="text-sm font-medium">{t('uniqueFinder.valuesToRemove')}</Label>
                        <Textarea
                            id="values-to-remove"
                            value={valuesToRemove}
                            onChange={(e) => setValuesToRemove(e.target.value)}
                            placeholder={t('uniqueFinder.valuesToRemovePlaceholder') as string}
                            disabled={isProcessing || !files}
                            rows={6}
                        />
                    </div>
                )}
            </CardContent>
        </Card>
        
        <Button onClick={handleProcess} disabled={isProcessButtonDisabled} className="w-full">
          {isProcessing ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : (operationMode === 'find' ? <Fingerprint className="mr-2 h-5 w-5" /> : <Trash2 className="mr-2 h-5 w-5" />)}
          {operationMode === 'find' ? t('uniqueFinder.processBtn') : t('uniqueFinder.processBtnRemove')}
        </Button>
      </CardContent>

      {lastOperation && (
        <CardFooter className="flex-col space-y-4 items-stretch">
            {lastOperation === 'find' && hasFindResults && (
                <Card className="p-4 bg-secondary/30">
                    <CardHeader className="p-0 pb-2">
                      <CardTitle className="text-lg font-headline">{t('uniqueFinder.resultsTitle')}</CardTitle>
                      <CardDescription>{t('uniqueFinder.resultsDesc', { count: valueCounts.size, total: totalOccurrences })}</CardDescription>
                    </CardHeader>
                    <CardContent className="p-0">
                        <ScrollArea className="h-72 mt-4 border rounded-md bg-background">
                           <Table>
                             <TableHeader>
                               <TableRow>
                                 <TableHead>{t('uniqueFinder.tableHeaderValue')}</TableHead>
                                 <TableHead className="text-right">{t('uniqueFinder.tableHeaderCount')}</TableHead>
                               </TableRow>
                             </TableHeader>
                             <TableBody>
                               {sortedValues.map(([value, count]) => (
                                 <TableRow key={value}>
                                   <TableCell className="font-code text-xs">{value}</TableCell>
                                   <TableCell className="text-right">{count}</TableCell>
                                 </TableRow>
                               ))}
                             </TableBody>
                           </Table>
                        </ScrollArea>
                    </CardContent>
                </Card>
            )}
            {lastOperation === 'remove' && hasRemoveResults && (
                 <Card className="p-4 bg-secondary/30">
                    <CardHeader className="p-0 pb-2">
                      <CardTitle className="text-lg font-headline">{t('uniqueFinder.resultsTitleRemove')}</CardTitle>
                      <CardDescription>{t('uniqueFinder.resultsDescRemove', { linesRemoved: removalStats.linesRemoved, originalLines: removalStats.originalLines })}</CardDescription>
                    </CardHeader>
                    {cleanedContent !== null && (
                        <CardContent className="p-0 pt-4">
                            <Label htmlFor="cleaned-preview" className="text-sm font-medium">{t('uniqueFinder.cleanedPreview')}</Label>
                            <ScrollArea className="h-40 mt-2 border rounded-md bg-background">
                                <Textarea
                                    id="cleaned-preview"
                                    readOnly
                                    value={cleanedContent}
                                    className="font-code text-xs w-full h-full p-2 border-0 focus-visible:ring-0"
                                    rows={10}
                                />
                            </ScrollArea>
                        </CardContent>
                    )}
                 </Card>
            )}

            {lastOperation === 'find' && hasFindResults && (
              <Button onClick={handleDownloadReport} variant="outline" className="w-full" disabled={isProcessing}>
                  <Download className="mr-2 h-5 w-5" />
                  {t('uniqueFinder.downloadBtn')}
              </Button>
            )}
            {lastOperation === 'remove' && hasRemoveResults && (
                 <Button onClick={handleDownloadCleanedFile} variant="outline" className="w-full" disabled={isProcessing}>
                    <Download className="mr-2 h-5 w-5" />
                    {t('uniqueFinder.downloadBtnCleaned')}
                </Button>
            )}
        </CardFooter>
      )}
       {!lastOperation && !isProcessing && (
        <CardFooter>
            <p className="text-sm text-muted-foreground w-full text-center">{t('uniqueFinder.noResults')}</p>
        </CardFooter>
      )}
    </Card>
  );
}
