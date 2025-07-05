
"use client";

import React, { useState, useCallback, ChangeEvent, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, LayoutGrid, CheckCircle2, Loader2, ListChecks, FileSpreadsheet } from 'lucide-react';
import { useLanguage } from '@/context/language-context';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { createPivotTableFromWorkbook } from '@/lib/excel-pivot-creator';
import type { AggregationType } from '@/lib/excel-types';

interface SheetSelection {
  [sheetName: string]: boolean;
}

interface PivotTableCreatorPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function PivotTableCreatorPage({ onProcessingChange, onFileStateChange }: PivotTableCreatorPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<SheetSelection>({});
  
  const [headerRow, setHeaderRow] = useState<number>(1);
  const [rowFields, setRowFields] = useState<string>('');
  const [columnFields, setColumnFields] = useState<string>('');
  const [valueField, setValueField] = useState<string>('');
  const [aggregationType, setAggregationType] = useState<AggregationType>('SUM');
  const [outputSheetName, setOutputSheetName] = useState<string>('Pivot_Table');

  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processedWb, setProcessedWb] = useState<XLSX.WorkBook | null>(null);
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

  const handleFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    setFile(selectedFile || null);
    if (selectedFile) {
        setIsProcessing(true);
        try {
            const arrayBuffer = await selectedFile.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
            setSheetNames(workbook.SheetNames);
            const initialSelection: SheetSelection = {};
            workbook.SheetNames.forEach(name => {
              initialSelection[name] = true;
            });
            setSelectedSheets(initialSelection);
        } catch (error) {
            toast({ title: t('toast.errorReadingFile') as string, description: t('toast.errorReadingSheets') as string, variant: 'destructive' });
            setSheetNames([]);
        } finally {
            setIsProcessing(false);
        }
    } else {
        setSheetNames([]);
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

  const handleProcess = useCallback(async () => {
    const sheetsToProcess = Object.keys(selectedSheets).filter(name => selectedSheets[name]);
    if (!file || sheetsToProcess.length === 0 || !rowFields.trim() || !valueField.trim()) {
      toast({ title: t('toast.missingInfo') as string, description: t('pivot.toast.missingInfo') as string, variant: 'destructive' });
      return;
    }

    setIsProcessing(true);
    setProcessedWb(null);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, bookVBA: true, bookFiles: true });
      
      const pivotWb = createPivotTableFromWorkbook(workbook, {
          sheetNames: sheetsToProcess,
          rowFields: rowFields.split(',').map(f => f.trim()),
          columnFields: columnFields.split(',').map(f => f.trim()).filter(Boolean),
          valueField: valueField.trim(),
          aggregationType,
          headerRow,
          outputSheetName,
      });

      setProcessedWb(pivotWb);
      toast({
        title: t('toast.processingComplete') as string,
        description: t('pivot.toast.success') as string,
        action: <CheckCircle2 className="text-green-500" />,
      });

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      toast({ title: t('toast.errorReadingFile') as string, description: t('pivot.toast.error', { errorMessage }) as string, variant: "destructive" });
    } finally {
      setIsProcessing(false);
    }
  }, [file, selectedSheets, headerRow, rowFields, columnFields, valueField, aggregationType, outputSheetName, toast, t]);

  const handleDownload = useCallback(() => {
    if (!processedWb) {
      toast({ title: [t('toast.noDataToDownload')].flat().join(' '), variant: "destructive" });
      return;
    }
    const fileName = `${outputSheetName}.xlsx`;
    XLSX.writeFile(processedWb, fileName, { compression: true, bookType: 'xlsx' });
  }, [processedWb, outputSheetName, toast, t]);
  
  const allSheetsSelected = sheetNames.length > 0 && sheetNames.every(name => selectedSheets[name]);

  return (
    <Card className="w-full max-w-2xl shadow-xl relative">
      {isProcessing && (
        <div className="absolute inset-0 bg-background/80 backdrop-blur-sm flex items-center justify-center z-10 rounded-lg">
          <div className="flex items-center gap-2 text-muted-foreground">
            <Loader2 className="h-6 w-6 animate-spin" />
            <span className="text-lg font-medium">{t('common.processing')}</span>
          </div>
        </div>
      )}
      <CardHeader>
        <div className="flex items-center space-x-2 mb-2">
          <LayoutGrid className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{t('pivot.title')}</CardTitle>
        </div>
        <CardDescription className="font-body">{t('pivot.description')}</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-pivot" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{t('pivot.uploadStep')}</span>
          </Label>
          <Input
            id="file-upload-pivot"
            type="file"
            accept=".xlsx, .xls, .xlsm"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isProcessing}
          />
        </div>

        {sheetNames.length > 0 && (
            <Card className="p-4 bg-secondary/30 space-y-4">
                <CardHeader className="p-0 pb-4">
                    <CardTitle className="text-lg flex items-center gap-2"><ListChecks className="h-5 w-5" />{t('pivot.selectSheetsStep')}</CardTitle>
                </CardHeader>
                <CardContent className="p-0">
                    <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-background">
                        <Checkbox
                        id="select-all-sheets-pivot"
                        checked={allSheetsSelected}
                        onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                        disabled={isProcessing}
                        />
                        <Label htmlFor="select-all-sheets-pivot" className="text-sm font-medium flex-grow">
                        {t('common.selectAll')} ({t('common.selectedCount', {selected: Object.values(selectedSheets).filter(Boolean).length, total: sheetNames.length})})
                        </Label>
                    </div>
                    <Card className="max-h-48 overflow-y-auto p-3 bg-background">
                        <div className="space-y-2">
                        {sheetNames.map(name => (
                            <div key={name} className="flex items-center space-x-2">
                            <Checkbox
                                id={`sheet-pivot-${name}`}
                                checked={selectedSheets[name] || false}
                                onCheckedChange={(checked) => handleSheetSelectionChange(name, checked as boolean)}
                                disabled={isProcessing}
                            />
                            <Label htmlFor={`sheet-pivot-${name}`} className="text-sm font-normal">{name}</Label>
                            </div>
                        ))}
                        </div>
                    </Card>
                </CardContent>
            </Card>
        )}

        {file && (
          <Card className="p-4 space-y-4">
            <CardHeader className="p-0 pb-4">
                <CardTitle className="text-lg flex items-center gap-2"><FileSpreadsheet className="h-5 w-5" />{t('pivot.configStep')}</CardTitle>
            </CardHeader>
            <CardContent className="p-0 space-y-6">
                <div className="space-y-2">
                    <Label htmlFor="header-row-pivot">{t('pivot.headerRowStep')}</Label>
                    <Input id="header-row-pivot" type="number" min="1" value={headerRow} onChange={(e) => setHeaderRow(parseInt(e.target.value, 10) || 1)} disabled={isProcessing} />
                    <p className="text-xs text-muted-foreground">{t('pivot.headerRowDesc')}</p>
                </div>
                <div className="space-y-2">
                    <Label htmlFor="row-fields">{t('pivot.rowFields')}</Label>
                    <Input id="row-fields" value={rowFields} onChange={e => setRowFields(e.target.value)} placeholder={t('pivot.rowFieldsPlaceholder') as string} disabled={isProcessing} />
                    <p className="text-xs text-muted-foreground">{t('pivot.rowFieldsDesc')}</p>
                </div>
                <div className="space-y-2">
                    <Label htmlFor="col-fields">{t('pivot.columnFields')}</Label>
                    <Input id="col-fields" value={columnFields} onChange={e => setColumnFields(e.target.value)} placeholder={t('pivot.columnFieldsPlaceholder') as string} disabled={isProcessing} />
                    <p className="text-xs text-muted-foreground">{t('pivot.columnFieldsDesc')}</p>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="space-y-2">
                        <Label htmlFor="value-field">{t('pivot.valueField')}</Label>
                        <Input id="value-field" value={valueField} onChange={e => setValueField(e.target.value)} placeholder={t('pivot.valueFieldPlaceholder') as string} disabled={isProcessing} />
                        <p className="text-xs text-muted-foreground">{t('pivot.valueFieldDesc')}</p>
                    </div>
                    <div className="space-y-2">
                        <Label htmlFor="agg-type">{t('pivot.aggregation')}</Label>
                        <Select value={aggregationType} onValueChange={(v) => setAggregationType(v as AggregationType)}>
                            <SelectTrigger id="agg-type"><SelectValue/></SelectTrigger>
                            <SelectContent>
                                <SelectItem value="SUM">SUM</SelectItem>
                                <SelectItem value="COUNT">COUNT</SelectItem>
                                <SelectItem value="AVERAGE">AVERAGE</SelectItem>
                                <SelectItem value="MIN">MIN</SelectItem>
                                <SelectItem value="MAX">MAX</SelectItem>
                            </SelectContent>
                        </Select>
                    </div>
                </div>
                 <div className="space-y-2">
                    <Label htmlFor="output-sheet-name">{t('pivot.outputSheet')}</Label>
                    <Input id="output-sheet-name" value={outputSheetName} onChange={e => setOutputSheetName(e.target.value)} placeholder={t('pivot.outputSheetPlaceholder') as string} disabled={isProcessing} />
                </div>
            </CardContent>
          </Card>
        )}

        <Button onClick={handleProcess} disabled={isProcessing || !file || !rowFields || !valueField} className="w-full">
          {isProcessing ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <LayoutGrid className="mr-2 h-5 w-5" />}
          {t('pivot.processBtn')}
        </Button>
      </CardContent>

      {processedWb && (
        <CardFooter>
          <Button onClick={handleDownload} variant="outline" className="w-full">
            <Download className="mr-2 h-5 w-5" />
            {t('pivot.downloadBtn')}
          </Button>
        </CardFooter>
      )}
    </Card>
  );
}
