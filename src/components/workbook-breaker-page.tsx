
"use client";

import React, { useState, useCallback, ChangeEvent, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, LibraryBig, Loader2, CheckCircle2, Group, FilterX, Scissors, Lightbulb, XCircle } from 'lucide-react';
import { useLanguage } from '@/context/language-context';
import { ScrollArea } from './ui/scroll-area';
import { RadioGroup, RadioGroupItem } from './ui/radio-group';
import { Alert, AlertDescription } from './ui/alert';
import { Markup } from '@/components/ui/markup';

interface WorkbookBreakerPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function WorkbookBreakerPage({ onProcessingChange, onFileStateChange }: WorkbookBreakerPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [sheetGroups, setSheetGroups] = useState<Record<string, string>>({}); // { [sheetName]: groupName }
  const [processedFiles, setProcessedFiles] = useState<Record<string, XLSX.WorkBook>>({});
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processingStatus, setProcessingStatus] = useState<string>('');
  const cancellationRequested = useRef(false);
  const [filterText, setFilterText] = useState<string>('');
  const [bulkGroupName, setBulkGroupName] = useState<string>('');
  const [desiredFiles, setDesiredFiles] = useState<string>('');
  const [outputFormat, setOutputFormat] = useState<'xlsx' | 'xlsm'>('xlsm');
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
    setSheetGroups({});
    setProcessedFiles({});
    setFilterText('');
    setBulkGroupName('');
    setDesiredFiles('');
    setOutputFormat('xlsm');

    if (file) {
      const getSheetNamesFromFile = async () => {
        setIsProcessing(true);
        try {
          const arrayBuffer = await file.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
          setSheetNames(workbook.SheetNames);
        } catch (error) {
          toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('toast.errorReadingSheets')].flat().join(' '), variant: 'destructive' });
        } finally {
          setIsProcessing(false);
        }
      };
      getSheetNamesFromFile();
    }
  }, [file, toast, t]);


  const handleFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
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

  const handleGroupChange = (sheetName: string, groupName: string) => {
    setSheetGroups(prev => ({ ...prev, [sheetName]: groupName }));
  };
  
  const handleClearGroups = () => {
    setSheetGroups({});
  };

  const filteredSheets = useMemo(() => {
    if (!filterText) return sheetNames;
    return sheetNames.filter(name => name.toLowerCase().includes(filterText.toLowerCase()));
  }, [sheetNames, filterText]);
  
  const handleBulkAssign = () => {
    if (!bulkGroupName) {
        toast({ title: [t('breaker.toast.noBulkNameTitle')].flat().join(' '), description: [t('breaker.toast.noBulkNameDesc')].flat().join(' '), variant: 'destructive' });
        return;
    }
    const newGroups = { ...sheetGroups };
    filteredSheets.forEach(sheetName => {
        newGroups[sheetName] = bulkGroupName;
    });
    setSheetGroups(newGroups);
  };
  
  const handleSuggestBreakup = () => {
    const numFiles = parseInt(desiredFiles, 10);
    if (isNaN(numFiles) || numFiles < 1 || numFiles > sheetNames.length) {
      toast({ title: [t('breaker.toast.noDesiredFilesTitle')].flat().join(' '), description: [t('breaker.toast.noDesiredFilesDesc')].flat().join(' '), variant: 'destructive' });
      return;
    }

    const totalSheets = sheetNames.length;
    const sheetsPerFile = Math.ceil(totalSheets / numFiles);
    const baseFileName = file ? file.name.substring(0, file.name.lastIndexOf('.')) : 'Workbook';

    const newGroups: Record<string, string> = {};
    sheetNames.forEach((sheetName, index) => {
      const groupNumber = Math.floor(index / sheetsPerFile) + 1;
      newGroups[sheetName] = `${baseFileName}_Part_${groupNumber}`;
    });
    setSheetGroups(newGroups);
    toast({ title: [t('breaker.toast.suggestionSuccessTitle')].flat().join(' '), description: [t('breaker.toast.suggestionSuccessDesc', { count: numFiles })].flat().join(' ') });
  };
  
  const handleBreakIntoIndividualFiles = () => {
    const newGroups: Record<string, string> = {};
    sheetNames.forEach(sheetName => {
        newGroups[sheetName] = sheetName;
    });
    setSheetGroups(newGroups);
    toast({ title: [t('breaker.toast.individualBreakupSuccessTitle')].flat().join(' '), description: [t('breaker.toast.individualBreakupSuccessDesc', { count: sheetNames.length })].flat().join(' ') });
  };

  const handleCancel = () => {
    cancellationRequested.current = true;
    setProcessingStatus([t('common.cancelling')].flat().join(' '));
  };

  const handleProcess = useCallback(async () => {
    if (!file) {
      toast({ title: [t('breaker.toast.noFileTitle')].flat().join(' '), description: [t('breaker.toast.noFileDesc')].flat().join(' '), variant: 'destructive' });
      return;
    }
    
    const groups: Record<string, string[]> = {};
    for (const sheetName in sheetGroups) {
      const groupName = sheetGroups[sheetName];
      if (groupName) {
        if (!groups[groupName]) groups[groupName] = [];
        groups[groupName].push(sheetName);
      }
    }
    
    if (Object.keys(groups).length === 0) {
      toast({ title: [t('breaker.toast.noGroupsTitle')].flat().join(' '), description: [t('breaker.toast.noGroupsDesc')].flat().join(' '), variant: 'destructive' });
      return;
    }
    
    cancellationRequested.current = false;
    setIsProcessing(true);
    setProcessingStatus('');
    setProcessedFiles({});

    try {
      const allNeededSheets = Object.values(groups).flat();
      const arrayBuffer = await file.arrayBuffer();
      // Read only the required sheets to save memory
      const originalWorkbook = XLSX.read(arrayBuffer, { type: 'buffer', sheets: allNeededSheets, cellStyles: true, bookVBA: true, bookFiles: true });
      
      const newWorkbooks: Record<string, XLSX.WorkBook> = {};
      const groupNames = Object.keys(groups);

      for (let i = 0; i < groupNames.length; i++) {
        const groupName = groupNames[i];
        if (cancellationRequested.current) throw new Error('Cancelled by user.');
        
        setProcessingStatus([t('breaker.toast.processing', { current: i + 1, total: groupNames.length, groupName })].flat().join(' '));

        const newWb = XLSX.utils.book_new();
        const sheetsInGroup = groups[groupName];
        
        for (const sheetName of sheetsInGroup) {
          const sheet = originalWorkbook.Sheets[sheetName];
          if (sheet) {
            XLSX.utils.book_append_sheet(newWb, sheet, sheetName);
          }
        }
        
        const baseName = groupName.replace(/[^a-z0-9]/gi, '_').substring(0, 50) || 'grouped';
        newWorkbooks[baseName] = newWb;
      }
      
      setProcessedFiles(newWorkbooks);
      toast({
        title: [t('toast.processingComplete')].flat().join(' '),
        description: [t('breaker.toast.success', { count: Object.keys(newWorkbooks).length })].flat().join(' '),
        action: <CheckCircle2 className="text-green-500" />
      });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : [t('breaker.toast.error')].flat().join(' ');
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
  }, [file, sheetGroups, toast, t]);

  const handleDownload = (baseName: string, workbook: XLSX.WorkBook) => {
    try {
        const fileName = `${baseName}.${outputFormat}`;
        XLSX.writeFile(workbook, fileName, { compression: true, bookType: outputFormat, cellStyles: true });
        toast({ title: [t('toast.downloadSuccess')].flat().join(' ') });
    } catch(error) {
        console.error("Download error:", error);
        toast({ title: [t('toast.downloadError')].flat().join(' '), variant: 'destructive' });
    }
  };

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
          <LibraryBig className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{t('breaker.title')}</CardTitle>
        </div>
        <CardDescription className="font-body">{t('breaker.description')}</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-breaker" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{t('breaker.uploadStep')}</span>
          </Label>
          <Input
            id="file-upload-breaker"
            type="file"
            accept=".xlsx, .xls, .xlsm"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isProcessing}
          />
        </div>

        {sheetNames.length > 0 && (
          <div className="space-y-4">
             <Card className="p-4 bg-secondary/30">
                <Label className="flex items-center space-x-2 text-md font-semibold mb-2">
                    <Scissors className="h-5 w-5" />
                    <span>{t('breaker.smartBreakup.title')}</span>
                </Label>
                <div className="space-y-4">
                    <div>
                        <p className="text-xs text-muted-foreground mb-4">{t('breaker.smartBreakup.description')}</p>
                        <div className="flex items-center gap-4">
                            <div className="flex-grow space-y-2">
                                <Label htmlFor="desired-files" className="text-sm font-medium">{t('breaker.smartBreakup.label')}</Label>
                                <Input
                                    id="desired-files"
                                    type="number"
                                    min="1"
                                    value={desiredFiles}
                                    onChange={(e) => setDesiredFiles(e.target.value)}
                                    placeholder={[t('breaker.smartBreakup.placeholder')].flat().join(' ')}
                                    disabled={isProcessing || !file}
                                />
                            </div>
                            <div className="pt-8">
                                <Button onClick={handleSuggestBreakup} disabled={isProcessing || !file || !desiredFiles}>
                                    {t('breaker.smartBreakup.button')}
                                </Button>
                            </div>
                        </div>
                    </div>
                    <div className="border-t pt-4">
                        <p className="text-xs text-muted-foreground mb-2">{t('breaker.smartBreakup.individualDesc')}</p>
                        <Button onClick={handleBreakIntoIndividualFiles} disabled={isProcessing || !file} className="w-full" variant="outline">
                             {t('breaker.smartBreakup.individualBtn')}
                        </Button>
                    </div>
                </div>
            </Card>

            <Card className="p-4 bg-secondary/30">
                <Label className="flex items-center space-x-2 text-md font-semibold mb-2">
                    <Group className="h-5 w-5" />
                    <span>{t('breaker.groupingStep')}</span>
                </Label>
                <div className="space-y-2">
                  <Label htmlFor="filter-sheets" className="text-sm font-medium">{t('breaker.filterSheets')}</Label>
                  <Input 
                    id="filter-sheets"
                    value={filterText}
                    onChange={(e) => setFilterText(e.target.value)}
                    placeholder={[t('breaker.filterPlaceholder')].flat().join(' ')}
                  />
                  <p className="text-xs text-muted-foreground">{t('breaker.filterDesc')}</p>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-4">
                    <div className="space-y-2">
                        <Label htmlFor="bulk-group-name" className="text-sm font-medium">{t('breaker.bulkAssignName')}</Label>
                        <Input
                            id="bulk-group-name"
                            value={bulkGroupName}
                            onChange={(e) => setBulkGroupName(e.target.value)}
                            placeholder={[t('breaker.bulkAssignPlaceholder')].flat().join(' ')}
                            disabled={!filterText}
                        />
                    </div>
                    <div className="flex items-end">
                       <Button onClick={handleBulkAssign} disabled={!filterText || !bulkGroupName} className="w-full">
                         {t('breaker.bulkAssignBtn')} ({filteredSheets.length})
                       </Button>
                    </div>
                </div>
            </Card>

            <div className="border rounded-lg overflow-hidden">
                <div className="p-4 bg-muted/50 grid grid-cols-2 gap-4 items-center">
                    <h4 className="text-sm font-semibold">{t('breaker.sheetName')}</h4>
                    <div className="flex justify-between items-center">
                        <h4 className="text-sm font-semibold">{t('breaker.newWorkbookName')}</h4>
                        <Button variant="outline" size="sm" onClick={handleClearGroups}><FilterX className="mr-2 h-4 w-4" />{t('breaker.clearGroups')}</Button>
                    </div>
                </div>
                <ScrollArea className="h-72">
                    <div className="p-4 space-y-2">
                        {filteredSheets.map(name => (
                            <div key={name} className="grid grid-cols-2 gap-4 items-center">
                                <Label htmlFor={`group-for-${name}`} className="truncate" title={name}>
                                    {name}
                                </Label>
                                <Input
                                    id={`group-for-${name}`}
                                    value={sheetGroups[name] || ''}
                                    onChange={e => handleGroupChange(name, e.target.value)}
                                    placeholder={[t('breaker.groupNamePlaceholder')].flat().join(' ')}
                                />
                            </div>
                        ))}
                    </div>
                </ScrollArea>
            </div>
            <p className="text-xs text-muted-foreground px-1 pt-2">
                <Markup text={[t('breaker.noteFileName')].flat().join(' ')} />
            </p>
            
            <Button onClick={handleProcess} disabled={isProcessing || Object.keys(sheetGroups).filter(key => sheetGroups[key]).length === 0} className="w-full">
              <LibraryBig className="mr-2 h-5 w-5" />
              {t('breaker.processBtn')}
            </Button>
          </div>
        )}
      </CardContent>

      {Object.keys(processedFiles).length > 0 && (
        <CardFooter className="flex-col space-y-4 items-stretch">
            <Card className="p-4 bg-secondary/30">
                <CardHeader className="p-0 pb-4">
                    <CardTitle className="text-lg font-headline">{t('breaker.resultsTitle')}</CardTitle>
                    <CardDescription>{t('breaker.resultsDesc')}</CardDescription>
                </CardHeader>
                <CardContent className="p-0 max-h-60 overflow-y-auto">
                    <div className="w-full p-4 border rounded-md bg-background space-y-4 mb-4">
                        <Label className="text-md font-semibold font-headline">{t('common.outputOptions.title')}</Label>
                        <RadioGroup value={outputFormat} onValueChange={(v) => setOutputFormat(v as any)} className="space-y-3">
                            <div>
                                <div className="flex items-center space-x-2">
                                    <RadioGroupItem value="xlsx" id="format-xlsx-breaker" />
                                    <Label htmlFor="format-xlsx-breaker" className="font-normal">{t('common.outputOptions.xlsx')}</Label>
                                </div>
                                <p className="text-xs text-muted-foreground pl-6 pt-1">{t('common.outputOptions.xlsxDesc')}</p>
                            </div>
                            <div>
                                <div className="flex items-center space-x-2">
                                    <RadioGroupItem value="xlsm" id="format-xlsm-breaker" />
                                    <Label htmlFor="format-xlsm-breaker" className="font-normal">{t('common.outputOptions.xlsm')}</Label>
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

                    <ul className="space-y-2">
                        {Object.entries(processedFiles).map(([baseName, workbook]) => {
                            const fileName = `${baseName}.${outputFormat}`;
                            return (
                                <li key={baseName} className="flex items-center justify-between p-2 bg-background rounded-md">
                                    <span className="font-medium truncate pr-4">{fileName}</span>
                                    <Button size="sm" onClick={() => handleDownload(baseName, workbook)}>
                                        <Download className="mr-2 h-4 w-4" />
                                        {t('common.download')}
                                    </Button>
                                </li>
                            );
                        })}
                    </ul>
                </CardContent>
            </Card>
        </CardFooter>
      )}
    </Card>
  );
}
