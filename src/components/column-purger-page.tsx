
"use client";

import React, { useState, useCallback, ChangeEvent, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, FileMinus, CheckCircle2, Loader2, ListChecks, FileSpreadsheet, Lightbulb, XCircle } from 'lucide-react';
import { purgeColumnsFromSheets } from '@/lib/excel-column-purger';
import { useLanguage } from '@/context/language-context';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Alert, AlertDescription } from '@/components/ui/alert';

interface SheetSelection {
  [sheetName: string]: boolean;
}

interface ColumnPurgerPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function ColumnPurgerPage({ onProcessingChange, onFileStateChange }: ColumnPurgerPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<SheetSelection>({});
  const [columnsToRemove, setColumnsToRemove] = useState<string>('');
  const [headerRow, setHeaderRow] = useState<number>(1);

  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [outputFormat, setOutputFormat] = useState<'xlsx' | 'xlsm'>('xlsm');
  const { toast } = useToast();

  useEffect(() => {
    if (onProcessingChange) {
      onProcessingChange(isLoading);
    }
  }, [isLoading, onProcessingChange]);
  
  useEffect(() => {
    if (onFileStateChange) {
      onFileStateChange(file !== null);
    }
  }, [file, onFileStateChange]);

  useEffect(() => {
    setSheetNames([]);
    setSelectedSheets({});
    setColumnsToRemove('');
    setHeaderRow(1);
    setIsLoading(false);
  }, [file]);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      if (!selectedFile.name.match(/\.(xlsx|xls|xlsm)$/)) {
        toast({ title: [t('toast.invalidFileType')].flat().join(' '), description: [t('toast.invalidFileTypeDesc')].flat().join(' '), variant: 'destructive' });
        setFile(null);
        return;
      }
      setFile(selectedFile);

      setIsLoading(true);
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: 'array' });
          setSheetNames(workbook.SheetNames);
          const initialSelection: SheetSelection = {};
          workbook.SheetNames.forEach(name => {
            initialSelection[name] = true;
          });
          setSelectedSheets(initialSelection);
        } catch (error) {
          toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: [t('toast.errorReadingSheets')].flat().join(' '), variant: 'destructive' });
        } finally {
          setIsLoading(false);
        }
      };
      reader.readAsArrayBuffer(selectedFile);
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

  const handleSheetSelectionChange = (sheetName: string, checked: boolean) => {
    setSelectedSheets(prev => ({ ...prev, [sheetName]: checked }));
  };

  const handleProcessAndDownload = useCallback(async () => {
    const sheetsToProcess = Object.keys(selectedSheets).filter(name => selectedSheets[name]);
    if (!file || sheetsToProcess.length === 0 || !columnsToRemove.trim()) {
      toast({ title: [t('toast.missingInfo')].flat().join(' '), description: [t('purger.toast.missingInfo')].flat().join(' '), variant: 'destructive' });
      return;
    }

    setIsLoading(true);
    try {
      const arrayBuffer = await file.arrayBuffer();
      let workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, bookFiles: true, bookVBA: true, cellDates: true });
      
      workbook = purgeColumnsFromSheets(workbook, sheetsToProcess, columnsToRemove, headerRow);

      const originalFileName = file.name.substring(0, file.name.lastIndexOf('.'));
      XLSX.writeFile(workbook, `${originalFileName}_purged.${outputFormat}`, { compression: true, cellStyles: true, bookType: outputFormat });
      
      toast({
        title: [t('toast.processingComplete')].flat().join(' '),
        description: [t('purger.toast.success')].flat().join(' '),
        action: <CheckCircle2 className="text-green-500" />,
      });
    } catch (error) {
      console.error("Error purging columns:", error);
      const errorMessage = error instanceof Error ? error.message : String(error);
      toast({ title: [t('toast.errorReadingFile')].flat().join(' '), description: errorMessage, variant: "destructive" });
    } finally {
      setIsLoading(false);
    }
  }, [file, selectedSheets, columnsToRemove, headerRow, outputFormat, toast, t]);

  const allSheetsSelected = sheetNames.length > 0 && sheetNames.every(name => selectedSheets[name]);

  return (
    <Card className="w-full max-w-lg md:max-w-xl lg:max-w-2xl shadow-xl relative">
      {isLoading && (
        <div className="absolute inset-0 bg-background/80 backdrop-blur-sm flex items-center justify-center z-10 rounded-lg">
          <div className="flex items-center gap-2 text-muted-foreground">
            <Loader2 className="h-6 w-6 animate-spin" />
            <span className="text-lg font-medium">{[t('common.processing')].flat().join(' ')}</span>
          </div>
        </div>
      )}
      <CardHeader>
        <div className="flex items-center space-x-2 mb-2">
          <FileMinus className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{[t('purger.title')].flat().join(' ')}</CardTitle>
        </div>
        <CardDescription className="font-body">{[t('purger.description')].flat().join(' ')}</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-purger" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{[t('purger.uploadStep')].flat().join(' ')}</span>
          </Label>
          <Input
            id="file-upload-purger"
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
              <span>{[t('purger.selectSheetsStep')].flat().join(' ')}</span>
            </Label>
            <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
              <Checkbox
                id="select-all-sheets-purger"
                checked={allSheetsSelected}
                onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                disabled={isLoading}
              />
              <Label htmlFor="select-all-sheets-purger" className="text-sm font-medium flex-grow">
                {[t('common.selectAll')].flat().join(' ')} ({[t('common.selectedCount', {selected: Object.values(selectedSheets).filter(Boolean).length, total: sheetNames.length})].flat().join(' ')})
              </Label>
            </div>
            <Card className="max-h-48 overflow-y-auto p-3 bg-background">
              <div className="space-y-2">
                {sheetNames.map(name => (
                  <div key={name} className="flex items-center space-x-2">
                    <Checkbox
                      id={`sheet-purger-${name}`}
                      checked={selectedSheets[name] || false}
                      onCheckedChange={(checked) => handleSheetSelectionChange(name, checked as boolean)}
                      disabled={isLoading}
                    />
                    <Label htmlFor={`sheet-purger-${name}`} className="text-sm font-normal">{name}</Label>
                  </div>
                ))}
              </div>
            </Card>
          </div>
        )}

        <div className="space-y-2">
            <Label htmlFor="header-row-purger" className="flex items-center space-x-2 text-sm font-medium">
                <FileSpreadsheet className="h-5 w-5" />
                <span>{[t('common.headerRow')].flat().join(' ')}</span>
            </Label>
            <Input 
                id="header-row-purger" 
                type="number" 
                min="1" 
                value={headerRow} 
                onChange={(e) => setHeaderRow(parseInt(e.target.value, 10) || 1)} 
                disabled={isLoading || !file}
            />
             <p className="text-xs text-muted-foreground">{[t('purger.headerRowDesc')].flat().join(' ')}</p>
        </div>

        <div className="space-y-2">
            <Label htmlFor="columns-to-remove" className="flex items-center space-x-2 text-sm font-medium">
                <FileMinus className="h-5 w-5" />
                <span>{[t('purger.columnStep')].flat().join(' ')}</span>
            </Label>
            <Input 
                id="columns-to-remove" 
                value={columnsToRemove} 
                onChange={e => setColumnsToRemove(e.target.value)} 
                disabled={isLoading || !file}
                placeholder={[t('purger.columnPlaceholder')].flat().join(' ')}
            />
            <p className="text-xs text-muted-foreground">{[t('purger.columnDesc')].flat().join(' ')}</p>
        </div>
      </CardContent>

      <CardFooter className="flex-col items-stretch space-y-4">
        <div className="w-full p-4 border rounded-md bg-secondary/30 space-y-4">
            <Label className="text-md font-semibold font-headline">{[t('common.outputOptions.title')].flat().join(' ')}</Label>
            <RadioGroup value={outputFormat} onValueChange={(v) => setOutputFormat(v as any)} className="space-y-3">
                <div>
                    <div className="flex items-center space-x-2">
                        <RadioGroupItem value="xlsx" id="format-xlsx-purger" />
                        <Label htmlFor="format-xlsx-purger" className="font-normal">{[t('common.outputOptions.xlsx')].flat().join(' ')}</Label>
                    </div>
                    <p className="text-xs text-muted-foreground pl-6 pt-1">{[t('common.outputOptions.xlsxDesc')].flat().join(' ')}</p>
                </div>
                <div>
                    <div className="flex items-center space-x-2">
                        <RadioGroupItem value="xlsm" id="format-xlsm-purger" />
                        <Label htmlFor="format-xlsm-purger" className="font-normal">{[t('common.outputOptions.xlsm')].flat().join(' ')}</Label>
                    </div>
                    <p className="text-xs text-muted-foreground pl-6 pt-1">{[t('common.outputOptions.xlsmDesc')].flat().join(' ')}</p>
                </div>
            </RadioGroup>
            <Alert variant="default" className="mt-2">
                <Lightbulb className="h-4 w-4" />
                <AlertDescription>{[t('common.outputOptions.recommendation')].flat().join(' ')}</AlertDescription>
            </Alert>
        </div>
        <Button 
            onClick={handleProcessAndDownload} 
            disabled={isLoading || !file || Object.values(selectedSheets).filter(Boolean).length === 0 || !columnsToRemove.trim()}
            className="w-full"
        >
          {isLoading ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-5 w-5" />}
          {[t('purger.processBtn')].flat().join(' ')}
        </Button>
      </CardFooter>
    </Card>
  );
}
