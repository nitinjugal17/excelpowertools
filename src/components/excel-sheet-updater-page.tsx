
"use client";

import React, { useState, useCallback, ChangeEvent, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Checkbox } from '@/components/ui/checkbox';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, Wand2, ListChecks, FileSpreadsheet, Loader2, FileEdit, Settings2, CheckCircle2, PencilLine, ScrollText, SplitSquareHorizontal, Settings, Palette, ChevronDown, Lightbulb, Shield, Ban } from 'lucide-react';
import { formatAndUpdateSheets } from '@/lib/excel-sheet-updater';
import type { FormattingConfig, CustomHeaderConfig, CustomColumnConfig, HorizontalAlignment, VerticalAlignment, RangeFormattingConfig, SheetProtectionConfig, CommandDisablingConfig } from '@/lib/excel-types';
import { generateVbsPreview } from '@/lib/vbs-generators';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { DropdownMenu, DropdownMenuTrigger, DropdownMenuContent, DropdownMenuItem } from '@/components/ui/dropdown-menu';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { useLanguage } from '@/context/language-context';
import { Markup } from '@/components/ui/markup';


interface SheetSelection {
  [sheetName: string]: boolean;
}

interface ExcelSheetUpdaterPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function ExcelSheetUpdaterPage({ onProcessingChange, onFileStateChange }: ExcelSheetUpdaterPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<SheetSelection>({});
  
  const [enableHeaderFormatting, setEnableHeaderFormatting] = useState<boolean>(false);
  const [headerRowNumberForFormatting, setHeaderRowNumberForFormatting] = useState<number>(1);
  const [formatOptions, setFormatOptions] = useState({
    bold: true,
    italic: false,
    underline: false,
    alignment: 'general' as HorizontalAlignment,
    fontName: 'Calibri',
    fontSize: 11,
  });

  const [enableCustomHeaderInsertion, setEnableCustomHeaderInsertion] = useState<boolean>(false);
  const [customHeaderText, setCustomHeaderText] = useState<string>('');
  const [customHeaderInsertBeforeRow, setCustomHeaderInsertBeforeRow] = useState<number>(1);
  const [customHeaderMergeAndCenter, setCustomHeaderMergeAndCenter] = useState<boolean>(true);
  const [customHeaderFormatOptions, setCustomHeaderFormatOptions] = useState({
    bold: true,
    italic: false,
    underline: false,
    fontSize: 12,
    fontName: 'Calibri',
    horizontalAlignment: 'center' as HorizontalAlignment,
    verticalAlignment: 'center' as VerticalAlignment,
    wrapText: false,
    indent: 0,
  });


  const [enableCustomColumnInsertion, setEnableCustomColumnInsertion] = useState<boolean>(false);
  const [newColumnName, setNewColumnName] = useState<string>('New Column');
  const [newColumnHeaderRow, setNewColumnHeaderRow] = useState<number>(1);
  const [insertColumnBefore, setInsertColumnBefore] = useState<string>('A');
  const [sourceDataColumn, setSourceDataColumn] = useState<string>('');
  const [textSplitter, setTextSplitter] = useState<string>('-');
  const [partToUse, setPartToUse] = useState<number>(1);
  const [dataStartRow, setDataStartRow] = useState<number>(2);
  const [customColumnAlignment, setCustomColumnAlignment] = useState<HorizontalAlignment>('general');

  const [enableRangeFormatting, setEnableRangeFormatting] = useState<boolean>(false);
  const [rangeFormatConfig, setRangeFormatConfig] = useState<RangeFormattingConfig>({
    startRow: 1,
    endRow: 1,
    startCol: 'A',
    endCol: 'D',
    merge: true,
    style: {
      font: {
        bold: true,
        italic: false,
        underline: false,
        name: 'Calibri',
        size: 14,
        color: '000000',
      },
      alignment: {
        horizontal: 'center',
        vertical: 'center',
      },
      fill: {
        color: 'FFFF00',
      },
    }
  });

  const [enableSheetProtection, setEnableSheetProtection] = useState<boolean>(false);
  const [protectionPassword, setProtectionPassword] = useState<string>('');
  const [protectionType, setProtectionType] = useState<'full' | 'range'>('full');
  const [protectionRange, setProtectionRange] = useState<string>('A1:D10');
  const [preventSelection, setPreventSelection] = useState<boolean>(false);

  const [enableCommandDisabling, setEnableCommandDisabling] = useState<boolean>(false);
  const [disableCopyPaste, setDisableCopyPaste] = useState<boolean>(true);
  const [disablePrint, setDisablePrint] = useState<boolean>(true);


  const [vbscriptPreview, setVbscriptPreview] = useState<string>('');
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
    if (!file) {
      setSheetNames([]);
      setSelectedSheets({});
      setEnableHeaderFormatting(false);
      setHeaderRowNumberForFormatting(1);
      setFormatOptions({ bold: true, italic: false, underline: false, alignment: 'general', fontName: 'Calibri', fontSize: 11 });
      setEnableCustomHeaderInsertion(false);
      setCustomHeaderText('');
      setCustomHeaderInsertBeforeRow(1);
      setEnableCustomColumnInsertion(false);
      setNewColumnName('New Column');
      setNewColumnHeaderRow(1);
      setInsertColumnBefore('A');
      setSourceDataColumn('');
      setTextSplitter('-');
      setPartToUse(1);
      setDataStartRow(2);
      setCustomColumnAlignment('general');
      setCustomHeaderMergeAndCenter(true);
      setCustomHeaderFormatOptions({
        bold: true,
        italic: false,
        underline: false,
        fontSize: 12,
        fontName: 'Calibri',
        horizontalAlignment: 'center' as HorizontalAlignment,
        verticalAlignment: 'center' as VerticalAlignment,
        wrapText: false,
        indent: 0,
      });
      setEnableRangeFormatting(false);
      setRangeFormatConfig({
        startRow: 1, endRow: 1, startCol: 'A', endCol: 'D', merge: true,
        style: {
          font: { bold: true, italic: false, underline: false, name: 'Calibri', size: 14, color: '000000' },
          alignment: { horizontal: 'center', vertical: 'center' },
          fill: { color: 'FFFF00' }
        }
      });
      setEnableSheetProtection(false);
      setProtectionPassword('');
      setProtectionType('full');
      setProtectionRange('A1:D10');
      setPreventSelection(false);
      setEnableCommandDisabling(false);
      setDisableCopyPaste(true);
      setDisablePrint(true);
    } else {
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
          toast({ title: t('toast.errorReadingFile') as string, description: t('toast.errorReadingSheets') as string, variant: "destructive" });
          setSheetNames([]);
          setSelectedSheets({});
        } finally {
          setIsLoading(false);
        }
      };
      getSheetNamesFromFile();
    }
  }, [file, toast, t]);

  // Effect to update VBScript Preview
  useEffect(() => {
    const commandDisablingConfig: CommandDisablingConfig | undefined = enableCommandDisabling ? {
        disableCopyPaste,
        disablePrint,
        vbaPassword: protectionPassword // Use same password for VBA project lock
    } : undefined;

    const script = generateVbsPreview(commandDisablingConfig);
    setVbscriptPreview(script);

  }, [enableCommandDisabling, disableCopyPaste, disablePrint, protectionPassword]);

  useEffect(() => {
    if (enableCommandDisabling || enableSheetProtection) {
        setOutputFormat('xlsm');
    }
  }, [enableCommandDisabling, enableSheetProtection]);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile) {
      if (!selectedFile.name.match(/\.(xlsx|xls|xlsm)$/)) {
        toast({
          title: t('toast.invalidFileType') as string,
          description: t('toast.invalidFileTypeDesc') as string,
          variant: 'destructive',
        });
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

  const handleUpdateAndDownload = useCallback(async () => {
    if (!file) {
      toast({ title: t('updater.toast.noFile') as string, variant: "destructive" });
      return;
    }
    const sheetsToUpdate = Object.entries(selectedSheets)
      .filter(([,isSelected]) => isSelected)
      .map(([sheetName]) => sheetName);

    if (sheetsToUpdate.length === 0) {
      toast({ title: t('updater.toast.noSheets') as string, variant: "destructive" });
      return;
    }
    if (enableHeaderFormatting && headerRowNumberForFormatting < 1) {
      toast({ title: t('updater.toast.invalidHeaderRow') as string, variant: "destructive" });
      return;
    }
    if (enableCustomHeaderInsertion) {
      if (!customHeaderText.trim()) {
          toast({ title: t('updater.toast.missingHeaderText') as string, variant: "destructive" });
          return;
      }
      if (customHeaderInsertBeforeRow < 1) {
        toast({ title: t('updater.toast.invalidInsertRow') as string, variant: "destructive" });
        return;
      }
    }
    if (enableCustomColumnInsertion) {
        if (!newColumnName.trim()) {
            toast({ title: t('updater.toast.missingColumnName') as string, variant: "destructive" });
            return;
        }
        if (!insertColumnBefore.trim()) {
            toast({ title: t('updater.toast.missingColumnPosition') as string, variant: "destructive" });
            return;
        }
        if (!sourceDataColumn.trim()) {
            toast({ title: t('updater.toast.missingSourceColumn') as string, variant: "destructive" });
            return;
        }
    }
    if (enableRangeFormatting) {
        if (!rangeFormatConfig.startCol || !rangeFormatConfig.endCol || rangeFormatConfig.startRow < 1 || rangeFormatConfig.endRow < 1) {
            toast({ title: t('updater.toast.invalidRange') as string, variant: "destructive" });
            return;
        }
    }
    if ((enableSheetProtection || enableCommandDisabling) && !protectionPassword) {
        toast({ title: t('updater.toast.missingPassword') as string, description: t('updater.toast.missingPasswordDesc') as string, variant: "destructive" });
        return;
    }


    setIsLoading(true);
    try {
      const arrayBuffer = await file.arrayBuffer();
      let workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true, bookFiles: true, bookVBA: true, cellDates: true });
      
      const customHeaderConfig: CustomHeaderConfig | undefined = enableCustomHeaderInsertion ? {
        text: customHeaderText,
        insertBeforeRow: customHeaderInsertBeforeRow,
        mergeAndCenter: customHeaderMergeAndCenter,
        styleOptions: customHeaderFormatOptions,
      } : undefined;

      const customColumnConfig: CustomColumnConfig | undefined = enableCustomColumnInsertion ? {
        newColumnName,
        newColumnHeaderRow,
        insertColumnBefore,
        sourceDataColumn,
        textSplitter,
        partToUse,
        dataStartRow,
        alignment: customColumnAlignment,
      } : undefined;

      const formattingConfig: FormattingConfig | undefined = enableHeaderFormatting ? {
        dataTitlesRowNumber: headerRowNumberForFormatting,
        styleOptions: formatOptions,
      } : undefined;

      const finalRangeFormatConfig = enableRangeFormatting ? rangeFormatConfig : undefined;
      
      const sheetProtectionConfig: SheetProtectionConfig | undefined = enableSheetProtection ? {
        password: protectionPassword,
        type: protectionType,
        range: protectionType === 'range' ? protectionRange : undefined,
        selectLockedCells: !preventSelection,
      } : undefined;
      
      const commandDisablingConfig: CommandDisablingConfig | undefined = enableCommandDisabling ? {
        disableCopyPaste,
        disablePrint,
        vbaPassword: protectionPassword
      } : undefined;

      workbook = formatAndUpdateSheets(workbook, sheetsToUpdate, formattingConfig, customHeaderConfig, customColumnConfig, finalRangeFormatConfig, sheetProtectionConfig, commandDisablingConfig);

      const finalOutputFormat = enableCommandDisabling || enableSheetProtection ? 'xlsm' : outputFormat;

      const originalFileName = file.name.substring(0, file.name.lastIndexOf('.'));
      XLSX.writeFile(workbook, `${originalFileName}_updated.${finalOutputFormat}`, { compression: true, cellStyles: true, bookType: finalOutputFormat });
      toast({
        title: t('toast.processingComplete') as string,
        description: t('updater.toast.updateSuccess') as string,
        action: <CheckCircle2 className="text-green-500" />,
      });

    } catch (error) {
      console.error("Error updating sheets:", error);
      const errorMessage = error instanceof Error ? error.message : t('updater.toast.updateErrorUnknown') as string;
      toast({ title: t('toast.errorReadingFile') as string, description: t('updater.toast.updateError', {errorMessage}) as string, variant: "destructive" });
    } finally {
      setIsLoading(false);
    }
  }, [file, selectedSheets, enableHeaderFormatting, headerRowNumberForFormatting, formatOptions, enableCustomHeaderInsertion, customHeaderText, customHeaderInsertBeforeRow, customHeaderMergeAndCenter, customHeaderFormatOptions, enableCustomColumnInsertion, newColumnName, newColumnHeaderRow, insertColumnBefore, sourceDataColumn, textSplitter, partToUse, dataStartRow, customColumnAlignment, toast, enableRangeFormatting, rangeFormatConfig, outputFormat, t, enableSheetProtection, protectionPassword, protectionType, protectionRange, preventSelection, enableCommandDisabling, disableCopyPaste, disablePrint]);

  const allSheetsSelected = sheetNames.length > 0 && sheetNames.every(name => selectedSheets[name]);

  return (
    <Card className="w-full max-w-lg md:max-w-xl lg:max-w-2xl xl:max-w-6xl shadow-xl relative">
      {isLoading && (
        <div className="absolute inset-0 bg-background/80 backdrop-blur-sm flex items-center justify-center z-10 rounded-lg">
            <div className="flex items-center gap-2 text-muted-foreground">
            <Loader2 className="h-6 w-6 animate-spin" />
            <span className="text-lg font-medium">{t('common.processing')}</span>
            </div>
        </div>
      )}
      <CardHeader>
        <div className="flex items-center space-x-2 mb-2">
          <Wand2 className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{t('updater.title')}</CardTitle>
        </div>
        <CardDescription className="font-body">
          {t('updater.description')}
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-updater" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>{t('updater.uploadStep')}</span>
          </Label>
          <Input
            id="file-upload-updater"
            type="file"
            accept=".xlsx, .xls, .xlsm"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isLoading}
          />
          {file && <p className="text-xs text-muted-foreground font-code">{t('common.selectedFile', {fileName: file.name})}</p>}
        </div>

        {sheetNames.length > 0 && (
          <>
            {/* Sheet Selection */}
            <div className="space-y-3">
                <Label className="flex items-center space-x-2 text-sm font-medium mb-2">
                    <ListChecks className="h-5 w-5" />
                    <span>{t('updater.selectSheetsStep')}</span>
                </Label>
                <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
                    <Checkbox
                        id="select-all-sheets"
                        checked={allSheetsSelected}
                        onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                        aria-label="Select all sheets"
                        disabled={isLoading}
                    />
                    <Label htmlFor="select-all-sheets" className="text-sm font-medium flex-grow">
                        {t('common.selectAll')} ({t('common.selectedCount', {selected: Object.values(selectedSheets).filter(Boolean).length, total: sheetNames.length})})
                    </Label>
                     {sheetNames.length > 50 && (
                        <DropdownMenu>
                            <DropdownMenuTrigger asChild>
                                <Button variant="outline" size="sm" disabled={isLoading}>
                                    {t('common.partial')}
                                    <ChevronDown className="ml-1 h-4 w-4" />
                                </Button>
                            </DropdownMenuTrigger>
                            <DropdownMenuContent>
                                <DropdownMenuItem onSelect={() => handlePartialSelection(50)}>{t('common.first50')}</DropdownMenuItem>
                                {sheetNames.length >= 100 && <DropdownMenuItem onSelect={() => handlePartialSelection(100)}>{t('common.first100')}</DropdownMenuItem>}
                                {sheetNames.length >= 150 && <DropdownMenuItem onSelect={() => handlePartialSelection(150)}>{t('common.first150')}</DropdownMenuItem>}
                            </DropdownMenuContent>
                        </DropdownMenu>
                    )}
                </div>
                <Card className="max-h-48 overflow-y-auto p-3 bg-background">
                    <div className="space-y-2">
                    {sheetNames.map(name => (
                        <div key={name} className="flex items-center space-x-2">
                        <Checkbox
                            id={`sheet-${name}`}
                            checked={selectedSheets[name] || false}
                            onCheckedChange={(checked) => handleSheetSelectionChange(name, checked as boolean)}
                            disabled={isLoading}
                        />
                        <Label htmlFor={`sheet-${name}`} className="text-sm font-normal">
                            {name}
                        </Label>
                        </div>
                    ))}
                    </div>
                </Card>
            </div>
            
            <div className="space-y-4">
              <Label className="flex items-center space-x-2 text-lg font-semibold">
                <Settings className="h-6 w-6" />
                <span>{t('updater.configStep')}</span>
              </Label>

              <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
                <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                  <Checkbox
                    id="enable-header-formatting"
                    checked={enableHeaderFormatting}
                    onCheckedChange={(checked) => setEnableHeaderFormatting(checked as boolean)}
                    disabled={isLoading}
                  />
                  <Label htmlFor="enable-header-formatting" className="flex items-center space-x-2 text-md font-semibold text-primary">
                    <FileEdit className="h-5 w-5" />
                    <span>{t('updater.dataHeaderStep')}</span>
                  </Label>
                </CardHeader>
                {enableHeaderFormatting && (
                  <CardContent className="p-0">
                    <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                        <div className="space-y-2">
                          <Label htmlFor="header-row-updater" className="text-sm font-medium">{t('updater.dataHeaderRow')}</Label>
                          <Input
                            id="header-row-updater"
                            type="number"
                            min="1"
                            value={headerRowNumberForFormatting}
                            onChange={(e) => setHeaderRowNumberForFormatting(Math.max(1, parseInt(e.target.value, 10) || 1))}
                            disabled={isLoading}
                            className="w-full"
                          />
                          <p className="text-xs text-muted-foreground">{t('updater.dataHeaderRowDesc')}</p>
                        </div>
                        <div className="space-y-2">
                            <Label className="text-sm font-medium">{t('updater.dataHeaderFormatting')}</Label>
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 p-3 border rounded-md bg-background">
                                <div className="flex flex-col space-y-2">
                                    <div className="flex items-center space-x-2">
                                        <Checkbox
                                        id="format-bold"
                                        checked={!!formatOptions.bold}
                                        onCheckedChange={(checked) => setFormatOptions(prev => ({ ...prev, bold: checked as boolean }))}
                                        disabled={isLoading}
                                        />
                                        <Label htmlFor="format-bold" className="text-sm font-normal">{t('common.bold')}</Label>
                                    </div>
                                    <div className="flex items-center space-x-2">
                                        <Checkbox
                                        id="format-italic"
                                        checked={!!formatOptions.italic}
                                        onCheckedChange={(checked) => setFormatOptions(prev => ({ ...prev, italic: checked as boolean }))}
                                        disabled={isLoading}
                                        />
                                        <Label htmlFor="format-italic" className="text-sm font-normal">{t('common.italic')}</Label>
                                    </div>
                                    <div className="flex items-center space-x-2">
                                        <Checkbox
                                        id="format-underline"
                                        checked={!!formatOptions.underline}
                                        onCheckedChange={(checked) => setFormatOptions(prev => ({ ...prev, underline: checked as boolean }))}
                                        disabled={isLoading}
                                        />
                                        <Label htmlFor="format-underline" className="text-sm font-normal">{t('common.underline')}</Label>
                                    </div>
                                </div>
                                <div className="space-y-4">
                                    <div className="space-y-1">
                                        <Label htmlFor="header-alignment-select" className="text-sm font-medium">{t('common.alignment')}</Label>
                                        <Select
                                            value={formatOptions.alignment || 'general'}
                                            onValueChange={(value) => setFormatOptions(prev => ({...prev, alignment: value as HorizontalAlignment}))}
                                            disabled={isLoading}
                                        >
                                            <SelectTrigger id="header-alignment-select">
                                                <SelectValue placeholder={t('updater.customColumnAlignmentPlaceholder') as string} />
                                            </SelectTrigger>
                                            <SelectContent>
                                                <SelectItem value="general">{t('common.alignments.general')}</SelectItem>
                                                <SelectItem value="left">{t('common.alignments.left')}</SelectItem>
                                                <SelectItem value="center">{t('common.alignments.center')}</SelectItem>
                                                <SelectItem value="centerContinuous">{t('common.alignments.centerContinuous')}</SelectItem>
                                                <SelectItem value="right">{t('common.alignments.right')}</SelectItem>
                                                <SelectItem value="fill">{t('common.alignments.fill')}</SelectItem>
                                                <SelectItem value="justify">{t('common.alignments.justify')}</SelectItem>
                                            </SelectContent>
                                        </Select>
                                    </div>
                                    <div className="space-y-1">
                                        <Label htmlFor="header-font-name" className="text-sm font-medium">{t('common.fontName')}</Label>
                                        <Input id="header-font-name" type="text" value={formatOptions.fontName || ''} onChange={(e) => setFormatOptions(prev => ({ ...prev, fontName: e.target.value }))} disabled={isLoading} placeholder="e.g., Arial" />
                                    </div>
                                    <div className="space-y-1">
                                        <Label htmlFor="header-font-size" className="text-sm font-medium">{t('common.fontSize')}</Label>
                                        <Input id="header-font-size" type="number" min="1" value={formatOptions.fontSize || 11} onChange={(e) => setFormatOptions(prev => ({ ...prev, fontSize: parseInt(e.target.value, 10) || 11 }))} disabled={isLoading} placeholder="e.g., 12" />
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                  </CardContent>
                )}
              </Card>

              <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
                <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                  <Checkbox
                    id="enable-custom-header"
                    checked={enableCustomHeaderInsertion}
                    onCheckedChange={(checked) => setEnableCustomHeaderInsertion(checked as boolean)}
                    disabled={isLoading}
                  />
                  <Label htmlFor="enable-custom-header" className="flex items-center space-x-2 text-md font-semibold text-primary">
                    <PencilLine className="h-5 w-5" />
                    <span>{t('updater.customHeaderStep')}</span>
                  </Label>
                </CardHeader>
                {enableCustomHeaderInsertion && (
                  <CardContent className="p-0">
                    <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                      <div>
                        <Label htmlFor="custom-header-text" className="text-sm font-medium">{t('updater.customHeaderText')}</Label>
                        <Input id="custom-header-text" type="text" value={customHeaderText} onChange={(e) => setCustomHeaderText(e.target.value)} disabled={isLoading} placeholder={t('updater.customHeaderTextPlaceholder') as string}/>
                        <p className="text-xs text-muted-foreground">{t('updater.customHeaderTextDesc')}</p>
                      </div>
                      <div>
                        <Label htmlFor="custom-header-insert-row" className="text-sm font-medium">{t('updater.customHeaderInsertRow')}</Label>
                        <Input id="custom-header-insert-row" type="number" min="1" value={customHeaderInsertBeforeRow} onChange={(e) => setCustomHeaderInsertBeforeRow(Math.max(1, parseInt(e.target.value, 10) || 1))} disabled={isLoading} placeholder={t('updater.customHeaderInsertRowPlaceholder') as string}/>
                        <p className="text-xs text-muted-foreground">{t('updater.customHeaderInsertRowDesc')}</p>
                      </div>
                      
                      <Card className="p-3 bg-background/50">
                        <h4 className="text-sm font-semibold mb-3">{t('updater.headerFormatting')}</h4>
                        <div className="space-y-4">
                            <div className="grid grid-cols-2 gap-4">
                                <div>
                                    <Label htmlFor="custom-header-font-name" className="text-sm font-medium">{t('common.fontName')}</Label>
                                    <Input id="custom-header-font-name" type="text" value={customHeaderFormatOptions.fontName || ''} onChange={(e) => setCustomHeaderFormatOptions(p => ({...p, fontName: e.target.value}))} disabled={isLoading} placeholder="e.g., Calibri"/>
                                </div>
                                 <div>
                                    <Label htmlFor="custom-header-font-size" className="text-sm font-medium">{t('common.fontSize')}</Label>
                                    <Input id="custom-header-font-size" type="number" min="1" value={customHeaderFormatOptions.fontSize || 12} onChange={(e) => setCustomHeaderFormatOptions(p => ({...p, fontSize: parseInt(e.target.value, 10) || 11}))} disabled={isLoading} />
                                </div>
                            </div>

                            <div className="grid grid-cols-2 gap-4">
                                <div>
                                   <Label htmlFor="custom-header-h-align" className="text-sm font-medium">{t('updater.hAlign')}</Label>
                                    <Select value={customHeaderFormatOptions.horizontalAlignment || 'general'} onValueChange={(v) => setCustomHeaderFormatOptions(p => ({...p, horizontalAlignment: v as HorizontalAlignment}))} disabled={isLoading}>
                                        <SelectTrigger id="custom-header-h-align"><SelectValue /></SelectTrigger>
                                        <SelectContent>
                                            <SelectItem value="general">{t('common.alignments.general')}</SelectItem>
                                            <SelectItem value="left">{t('common.alignments.left')}</SelectItem>
                                            <SelectItem value="center">{t('common.alignments.center')}</SelectItem>
                                            <SelectItem value="centerContinuous">{t('common.alignments.centerContinuous')}</SelectItem>
                                            <SelectItem value="right">{t('common.alignments.right')}</SelectItem>
                                            <SelectItem value="fill">{t('common.alignments.fill')}</SelectItem>
                                            <SelectItem value="justify">{t('common.alignments.justify')}</SelectItem>
                                        </SelectContent>
                                    </Select>
                                </div>
                                <div>
                                    <Label htmlFor="custom-header-v-align" className="text-sm font-medium">{t('updater.vAlign')}</Label>
                                    <Select value={customHeaderFormatOptions.verticalAlignment || 'center'} onValueChange={(v) => setCustomHeaderFormatOptions(p => ({...p, verticalAlignment: v as VerticalAlignment}))} disabled={isLoading}>
                                        <SelectTrigger id="custom-header-v-align"><SelectValue /></SelectTrigger>
                                        <SelectContent>
                                            <SelectItem value="top">{t('common.alignments.top')}</SelectItem>
                                            <SelectItem value="center">{t('common.alignments.center')}</SelectItem>
                                            <SelectItem value="bottom">{t('common.alignments.bottom')}</SelectItem>
                                            <SelectItem value="justify">{t('common.alignments.justify')}</SelectItem>
                                            <SelectItem value="distributed">{t('common.alignments.distributed')}</SelectItem>
                                        </SelectContent>
                                    </Select>
                                </div>
                            </div>
                            <div>
                                <Label htmlFor="custom-header-indent" className="text-sm font-medium">{t('updater.indent')}</Label>
                                <Input id="custom-header-indent" type="number" min="0" value={customHeaderFormatOptions.indent || 0} onChange={(e) => setCustomHeaderFormatOptions(p => ({...p, indent: Math.max(0, parseInt(e.target.value,10) || 0)}))} disabled={isLoading}/>
                                <p className="text-xs text-muted-foreground">{t('updater.indentDesc')}</p>
                            </div>
                            <div className="grid grid-cols-2 gap-4 pt-2">
                                 <div className="flex flex-col space-y-2">
                                   <div className="flex items-center space-x-2"><Checkbox id="h-format-bold-custom" checked={!!customHeaderFormatOptions.bold} onCheckedChange={(checked) => setCustomHeaderFormatOptions(p => ({ ...p, bold: checked as boolean }))} disabled={isLoading} /><Label htmlFor="h-format-bold-custom">{t('common.bold')}</Label></div>
                                   <div className="flex items-center space-x-2"><Checkbox id="h-format-italic-custom" checked={!!customHeaderFormatOptions.italic} onCheckedChange={(checked) => setCustomHeaderFormatOptions(p => ({ ...p, italic: checked as boolean }))} disabled={isLoading} /><Label htmlFor="h-format-italic-custom">{t('common.italic')}</Label></div>
                                   <div className="flex items-center space-x-2"><Checkbox id="h-format-underline-custom" checked={!!customHeaderFormatOptions.underline} onCheckedChange={(checked) => setCustomHeaderFormatOptions(p => ({ ...p, underline: checked as boolean }))} disabled={isLoading} /><Label htmlFor="h-format-underline-custom">{t('common.underline')}</Label></div>
                                </div>
                                 <div className="flex flex-col space-y-2">
                                   <div className="flex items-center space-x-2 pt-2">
                                      <Checkbox id="custom-header-wrap-text" checked={!!customHeaderFormatOptions.wrapText} onCheckedChange={(checked) => setCustomHeaderFormatOptions(p => ({...p, wrapText: checked as boolean}))} disabled={isLoading} />
                                      <Label htmlFor="custom-header-wrap-text" className="text-sm font-normal">{t('updater.wrapText')}</Label>
                                   </div>
                                   <div className="flex items-center space-x-2 pt-2">
                                      <Checkbox id="custom-header-merge-center" checked={!!customHeaderMergeAndCenter} onCheckedChange={(checked) => setCustomHeaderMergeAndCenter(checked as boolean)} disabled={isLoading} />
                                      <Label htmlFor="custom-header-merge-center" className="text-sm font-normal">{t('splitter.mergeAndCenter')}</Label>
                                   </div>
                                </div>
                            </div>
                        </div>
                      </Card>
                    </div>
                  </CardContent>
                )}
              </Card>

              <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
                <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                  <Checkbox
                    id="enable-custom-column"
                    checked={enableCustomColumnInsertion}
                    onCheckedChange={(checked) => setEnableCustomColumnInsertion(checked as boolean)}
                    disabled={isLoading}
                  />
                  <Label htmlFor="enable-custom-column" className="flex items-center space-x-2 text-md font-semibold text-primary">
                    <SplitSquareHorizontal className="h-5 w-5" />
                    <span>{t('updater.customColumnStep')}</span>
                  </Label>
                </CardHeader>
                {enableCustomColumnInsertion && (
                  <CardContent className="p-0">
                    <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                      <div className="space-y-1">
                        <Label htmlFor="new-column-name" className="text-sm font-medium">{t('updater.newColumnHeader')}</Label>
                        <Input id="new-column-name" type="text" value={newColumnName} onChange={(e) => setNewColumnName(e.target.value)} disabled={isLoading} placeholder={t('updater.newColumnHeaderPlaceholder') as string} />
                        <p className="text-xs text-muted-foreground">{t('updater.newColumnHeaderDesc')}</p>
                      </div>
                      <div className="space-y-1">
                        <Label htmlFor="new-column-header-row" className="text-sm font-medium">{t('updater.newColumnHeaderRow')}</Label>
                        <Input id="new-column-header-row" type="number" min="1" value={newColumnHeaderRow} onChange={(e) => setNewColumnHeaderRow(Math.max(1, parseInt(e.target.value, 10) || 1))} disabled={isLoading} placeholder={t('updater.newColumnHeaderRowPlaceholder') as string}/>
                        <p className="text-xs text-muted-foreground">{t('updater.newColumnHeaderRowDesc')}</p>
                      </div>
                      <div className="space-y-1">
                        <Label htmlFor="insert-column-before" className="text-sm font-medium">{t('updater.insertColumnBefore')}</Label>
                        <Input id="insert-column-before" type="text" value={insertColumnBefore} onChange={(e) => setInsertColumnBefore(e.target.value)} disabled={isLoading} placeholder={t('updater.insertColumnBeforePlaceholder') as string} />
                        <p className="text-xs text-muted-foreground">{t('updater.insertColumnBeforeDesc')}</p>
                      </div>
                      <div className="space-y-1">
                        <Label htmlFor="source-data-column" className="text-sm font-medium">{t('updater.sourceDataColumn')}</Label>
                        <Input id="source-data-column" type="text" value={sourceDataColumn} onChange={(e) => setSourceDataColumn(e.target.value)} disabled={isLoading} placeholder={t('updater.sourceDataColumnPlaceholder') as string} />
                         <p className="text-xs text-muted-foreground">{t('updater.sourceDataColumnDesc')}</p>
                      </div>
                      <div className="space-y-1">
                        <Label htmlFor="text-splitter" className="text-sm font-medium">{t('updater.textDelimiter')}</Label>
                        <Input id="text-splitter" type="text" value={textSplitter} onChange={(e) => setTextSplitter(e.target.value)} disabled={isLoading} placeholder={t('updater.textDelimiterPlaceholder') as string} />
                         <p className="text-xs text-muted-foreground">{t('updater.textDelimiterDesc')}</p>
                      </div>
                      <div className="space-y-1">
                        <Label htmlFor="part-to-use" className="text-sm font-medium">{t('updater.partToUse')}</Label>
                        <Input 
                          id="part-to-use" 
                          type="number" 
                          value={partToUse} 
                          onChange={(e) => {
                              const val = parseInt(e.target.value, 10);
                              if (isNaN(val) || val === 0) {
                                  setPartToUse(1);
                              } else {
                                  setPartToUse(val);
                              }
                          }} 
                          disabled={isLoading}
                        />
                         <p className="text-xs text-muted-foreground">{t('updater.partToUseDesc')}</p>
                      </div>
                       <div className="space-y-1">
                        <Label htmlFor="data-start-row" className="text-sm font-medium">{t('updater.dataStartRow')}</Label>
                        <Input id="data-start-row" type="number" min="1" value={dataStartRow} onChange={(e) => setDataStartRow(Math.max(1, parseInt(e.target.value, 10) || 1))} disabled={isLoading} placeholder={t('updater.dataStartRowPlaceholder') as string} />
                         <p className="text-xs text-muted-foreground">{t('updater.dataStartRowDesc')}</p>
                      </div>
                      <div className="space-y-1">
                        <Label htmlFor="custom-column-alignment">{t('updater.customColumnAlignment')}</Label>
                        <Select
                            value={customColumnAlignment}
                            onValueChange={(value) => setCustomColumnAlignment(value as HorizontalAlignment)}
                            disabled={isLoading}
                        >
                            <SelectTrigger id="custom-column-alignment">
                                <SelectValue placeholder={t('updater.customColumnAlignmentPlaceholder') as string} />
                            </SelectTrigger>
                            <SelectContent>
                                <SelectItem value="general">{t('common.alignments.general')}</SelectItem>
                                <SelectItem value="left">{t('common.alignments.left')}</SelectItem>
                                <SelectItem value="center">{t('common.alignments.center')}</SelectItem>
                                <SelectItem value="centerContinuous">{t('common.alignments.centerContinuous')}</SelectItem>
                                <SelectItem value="right">{t('common.alignments.right')}</SelectItem>
                                <SelectItem value="fill">{t('common.alignments.fill')}</SelectItem>
                                <SelectItem value="justify">{t('common.alignments.justify')}</SelectItem>
                            </SelectContent>
                        </Select>
                         <p className="text-xs text-muted-foreground">{t('updater.customColumnAlignmentDesc')}</p>
                      </div>
                    </div>
                  </CardContent>
                )}
              </Card>

              <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
                <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                  <Checkbox
                    id="enable-range-formatting"
                    checked={enableRangeFormatting}
                    onCheckedChange={(checked) => setEnableRangeFormatting(checked as boolean)}
                    disabled={isLoading}
                  />
                  <Label htmlFor="enable-range-formatting" className="flex items-center space-x-2 text-md font-semibold text-primary">
                    <Palette className="h-5 w-5" />
                    <span>{t('updater.rangeFormattingStep')}</span>
                  </Label>
                </CardHeader>
                {enableRangeFormatting && (
                  <CardContent className="p-0">
                     <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                        <p className="text-xs text-muted-foreground pb-2">{t('updater.rangeFormattingDesc')}</p>
                        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                           <div className="space-y-1">
                            <Label htmlFor="rf-start-row" className="text-sm">{t('updater.startRow')}</Label>
                            <Input id="rf-start-row" type="number" min="1" value={rangeFormatConfig.startRow} onChange={e => setRangeFormatConfig(p => ({...p, startRow: parseInt(e.target.value, 10) || 1}))} />
                           </div>
                           <div className="space-y-1">
                            <Label htmlFor="rf-end-row" className="text-sm">{t('updater.endRow')}</Label>
                            <Input id="rf-end-row" type="number" min="1" value={rangeFormatConfig.endRow} onChange={e => setRangeFormatConfig(p => ({...p, endRow: parseInt(e.target.value, 10) || 1}))} />
                           </div>
                           <div className="space-y-1">
                            <Label htmlFor="rf-start-col" className="text-sm">{t('updater.startCol')}</Label>
                            <Input id="rf-start-col" type="text" value={rangeFormatConfig.startCol} onChange={e => setRangeFormatConfig(p => ({...p, startCol: e.target.value}))} />
                           </div>
                           <div className="space-y-1">
                            <Label htmlFor="rf-end-col" className="text-sm">{t('updater.endCol')}</Label>
                            <Input id="rf-end-col" type="text" value={rangeFormatConfig.endCol} onChange={e => setRangeFormatConfig(p => ({...p, endCol: e.target.value}))} />
                           </div>
                        </div>
                        <Card className="p-3 bg-background/50">
                          <h4 className="text-sm font-semibold mb-3">{t('updater.rangeStyleOptions')}</h4>
                          <div className="space-y-4">
                            <div className="grid grid-cols-2 gap-4">
                              <div>
                                <Label htmlFor="rf-font-name" className="text-sm">{t('common.fontName')}</Label>
                                <Input id="rf-font-name" value={rangeFormatConfig.style.font.name} onChange={e => setRangeFormatConfig(p => ({...p, style: {...p.style, font: {...p.style.font, name: e.target.value}}}))} />
                              </div>
                               <div>
                                <Label htmlFor="rf-font-size" className="text-sm">{t('common.fontSize')}</Label>
                                <Input id="rf-font-size" type="number" min="1" value={rangeFormatConfig.style.font.size} onChange={e => setRangeFormatConfig(p => ({...p, style: {...p.style, font: {...p.style.font, size: parseInt(e.target.value, 10) || 11}}}))} />
                              </div>
                            </div>
                             <div className="grid grid-cols-2 gap-4">
                              <div>
                                <Label htmlFor="rf-font-color" className="text-sm">{t('updater.fontColorHex')}</Label>
                                <Input id="rf-font-color" value={rangeFormatConfig.style.font.color} onChange={e => setRangeFormatConfig(p => ({...p, style: {...p.style, font: {...p.style.font, color: e.target.value.replace('#','')}}}))} placeholder="e.g., 000000" />
                              </div>
                              <div>
                                <Label htmlFor="rf-fill-color" className="text-sm">{t('updater.fillColorHex')}</Label>
                                <Input id="rf-fill-color" value={rangeFormatConfig.style.fill.color} onChange={e => setRangeFormatConfig(p => ({...p, style: {...p.style, fill: {...p.style.fill, color: e.target.value.replace('#','')}}}))} placeholder="e.g., FFFF00" />
                              </div>
                            </div>
                             <div className="grid grid-cols-2 gap-4">
                              <div>
                                <Label htmlFor="rf-h-align" className="text-sm">{t('updater.hAlign')}</Label>
                                <Select value={rangeFormatConfig.style.alignment.horizontal} onValueChange={v => setRangeFormatConfig(p => ({...p, style: {...p.style, alignment: {...p.style.alignment, horizontal: v as HorizontalAlignment}}}))}>
                                  <SelectTrigger id="rf-h-align"><SelectValue /></SelectTrigger>
                                  <SelectContent>
                                    <SelectItem value="general">{t('common.alignments.general')}</SelectItem><SelectItem value="left">{t('common.alignments.left')}</SelectItem><SelectItem value="center">{t('common.alignments.center')}</SelectItem><SelectItem value="right">{t('common.alignments.right')}</SelectItem><SelectItem value="fill">{t('common.alignments.fill')}</SelectItem><SelectItem value="justify">{t('common.alignments.justify')}</SelectItem>
                                  </SelectContent>
                                </Select>
                              </div>
                              <div>
                                <Label htmlFor="rf-v-align" className="text-sm">{t('updater.vAlign')}</Label>
                                <Select value={rangeFormatConfig.style.alignment.vertical} onValueChange={v => setRangeFormatConfig(p => ({...p, style: {...p.style, alignment: {...p.style.alignment, vertical: v as VerticalAlignment}}}))}>
                                  <SelectTrigger id="rf-v-align"><SelectValue /></SelectTrigger>
                                  <SelectContent>
                                    <SelectItem value="top">{t('common.alignments.top')}</SelectItem><SelectItem value="center">{t('common.alignments.center')}</SelectItem><SelectItem value="bottom">{t('common.alignments.bottom')}</SelectItem><SelectItem value="justify">{t('common.alignments.justify')}</SelectItem><SelectItem value="distributed">{t('common.alignments.distributed')}</SelectItem>
                                  </SelectContent>
                                </Select>
                              </div>
                            </div>
                             <div className="flex items-center space-x-4 pt-2">
                               <div className="flex items-center space-x-2"><Checkbox id="rf-font-bold" checked={rangeFormatConfig.style.font.bold} onCheckedChange={c => setRangeFormatConfig(p => ({...p, style: {...p.style, font: {...p.style.font, bold: !!c}}}))} /><Label htmlFor="rf-font-bold">{t('common.bold')}</Label></div>
                               <div className="flex items-center space-x-2"><Checkbox id="rf-font-italic" checked={rangeFormatConfig.style.font.italic} onCheckedChange={c => setRangeFormatConfig(p => ({...p, style: {...p.style, font: {...p.style.font, italic: !!c}}}))} /><Label htmlFor="rf-font-italic">{t('common.italic')}</Label></div>
                               <div className="flex items-center space-x-2"><Checkbox id="rf-font-underline" checked={rangeFormatConfig.style.font.underline} onCheckedChange={c => setRangeFormatConfig(p => ({...p, style: {...p.style, font: {...p.style.font, underline: !!c}}}))} /><Label htmlFor="rf-font-underline">{t('common.underline')}</Label></div>
                               <div className="flex items-center space-x-2"><Checkbox id="rf-merge" checked={rangeFormatConfig.merge} onCheckedChange={c => setRangeFormatConfig(p => ({...p, merge: !!c}))} /><Label htmlFor="rf-merge">{t('updater.mergeRange')}</Label></div>
                            </div>
                          </div>
                        </Card>
                     </div>
                  </CardContent>
                )}
              </Card>

              <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
                <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                  <Checkbox
                    id="enable-sheet-protection"
                    checked={enableSheetProtection}
                    onCheckedChange={(checked) => setEnableSheetProtection(checked as boolean)}
                    disabled={isLoading}
                  />
                  <Label htmlFor="enable-sheet-protection" className="flex items-center space-x-2 text-md font-semibold text-primary">
                    <Shield className="h-5 w-5" />
                    <span>{t('updater.protection.title')}</span>
                  </Label>
                </CardHeader>
                {enableSheetProtection && (
                  <CardContent className="p-0">
                     <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                        <div className="space-y-2">
                            <Label htmlFor="protection-password">{t('updater.protection.password')}</Label>
                            <Input id="protection-password" type="password" value={protectionPassword} onChange={e => setProtectionPassword(e.target.value)} />
                            <p className="text-xs text-muted-foreground">{t('updater.protection.passwordDesc')}</p>
                        </div>
                        <div className="space-y-2">
                            <Label>{t('updater.protection.type')}</Label>
                            <RadioGroup value={protectionType} onValueChange={v => setProtectionType(v as any)} className="space-y-2">
                                <Label className="flex items-center space-x-2 font-normal"><RadioGroupItem value="full" id="protect-full" /><span>{t('updater.protection.fullSheet')}</span></Label>
                                <Label className="flex items-center space-x-2 font-normal"><RadioGroupItem value="range" id="protect-range" /><span>{t('updater.protection.specificRange')}</span></Label>
                            </RadioGroup>
                        </div>
                        {protectionType === 'range' && (
                            <div className="space-y-2">
                                <Label htmlFor="protection-range">{t('updater.protection.rangeToLock')}</Label>
                                <Input id="protection-range" value={protectionRange} onChange={e => setProtectionRange(e.target.value)} placeholder="e.g., A1:D10" />
                                <p className="text-xs text-muted-foreground">{t('updater.protection.rangeToLockDesc')}</p>
                            </div>
                        )}
                        <div className="flex items-start space-x-3 pt-4 border-t mt-4">
                            <Checkbox
                                id="protection-prevent-select"
                                checked={preventSelection}
                                onCheckedChange={(checked) => setPreventSelection(checked as boolean)}
                                disabled={isLoading}
                            />
                            <div className="grid gap-1.5 leading-none">
                                <Label htmlFor="protection-prevent-select" className="font-normal">{t('updater.protection.preventSelection')}</Label>
                                <p className="text-xs text-muted-foreground">{t('updater.protection.preventSelectionDesc')}</p>
                            </div>
                        </div>
                     </div>
                  </CardContent>
                )}
              </Card>

                <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
                <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                    <Checkbox
                        id="enable-command-disabling"
                        checked={enableCommandDisabling}
                        onCheckedChange={(checked) => setEnableCommandDisabling(checked as boolean)}
                        disabled={isLoading}
                    />
                    <Label htmlFor="enable-command-disabling" className="flex items-center space-x-2 text-md font-semibold text-primary">
                        <Ban className="h-5 w-5" />
                        <span>{t('updater.commandDisabling.title')}</span>
                    </Label>
                </CardHeader>
                {enableCommandDisabling && (
                    <CardContent className="p-0">
                        <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                             <Alert variant="destructive">
                                <Shield className="h-4 w-4" />
                                <AlertTitle>{t('updater.protection.vbaLockTitle')}</AlertTitle>
                                <AlertDescription>
                                    <Markup text={t('updater.protection.vbaLockDesc') as string} />
                                </AlertDescription>
                            </Alert>
                            <div className="flex items-center space-x-2">
                                <Checkbox
                                    id="disable-copy-paste"
                                    checked={disableCopyPaste}
                                    onCheckedChange={(c) => setDisableCopyPaste(c as boolean)}
                                />
                                <Label htmlFor="disable-copy-paste" className="font-normal">{t('updater.commandDisabling.disableCopyPaste')}</Label>
                            </div>
                             <div className="flex items-center space-x-2">
                                <Checkbox
                                    id="disable-print"
                                    checked={disablePrint}
                                    onCheckedChange={(c) => setDisablePrint(c as boolean)}
                                />
                                <Label htmlFor="disable-print" className="font-normal">{t('updater.commandDisabling.disablePrint')}</Label>
                            </div>
                        </div>
                    </CardContent>
                )}
                </Card>
            </div>

            {/* VBScript Preview Section */}
            {enableCommandDisabling && (
                <div className="space-y-2">
                <Label className="flex items-center space-x-2 text-sm font-medium">
                    <ScrollText className="h-5 w-5" />
                    <span>{t('updater.vbsPreviewStep')}</span>
                </Label>
                <Card className="bg-secondary/20">
                    <CardContent className="p-0">
                    <pre className="text-xs p-4 overflow-x-auto bg-gray-800 text-white rounded-md max-h-60">
                        <code>{vbscriptPreview}</code>
                    </pre>
                    </CardContent>
                </Card>
                <p className="text-xs text-muted-foreground">
                    <Markup text={t('updater.vbsPreviewDesc') as string} />
                </p>
                </div>
            )}
          </>
        )}
      </CardContent>
      {file && sheetNames.length > 0 && (
        <CardFooter className="flex-col items-stretch space-y-4">
            <div className="w-full p-4 border rounded-md bg-secondary/30 space-y-4">
                <Label className="text-md font-semibold font-headline">{t('common.outputOptions.title')}</Label>
                <RadioGroup value={outputFormat} onValueChange={(v) => setOutputFormat(v as any)} className="space-y-3">
                    <div>
                        <div className="flex items-center space-x-2">
                            <RadioGroupItem value="xlsx" id="format-xlsx-updater" disabled={enableCommandDisabling || enableSheetProtection} />
                            <Label htmlFor="format-xlsx-updater" className="font-normal">{t('common.outputOptions.xlsx')}</Label>
                        </div>
                        <p className="text-xs text-muted-foreground pl-6 pt-1">{t('common.outputOptions.xlsxDesc')}</p>
                    </div>
                    <div>
                        <div className="flex items-center space-x-2">
                            <RadioGroupItem value="xlsm" id="format-xlsm-updater" />
                            <Label htmlFor="format-xlsm-updater" className="font-normal">{t('common.outputOptions.xlsm')}</Label>
                        </div>
                        <p className="text-xs text-muted-foreground pl-6 pt-1">{t('common.outputOptions.xlsmDesc')}</p>
                    </div>
                </RadioGroup>
                <Alert variant="default" className="mt-2">
                    <Lightbulb className="h-4 w-4" />
                    <AlertDescription>
                        {enableCommandDisabling || enableSheetProtection ? t('updater.toast.macroRequired') : t('common.outputOptions.recommendation')}
                    </AlertDescription>
                </Alert>
            </div>
          <Button
            onClick={handleUpdateAndDownload}
            disabled={isLoading || Object.values(selectedSheets).filter(Boolean).length === 0}
            className="w-full bg-primary hover:bg-primary/90 text-primary-foreground"
          >
            {isLoading && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
            <FileEdit className="mr-2 h-5 w-5" />
            {t('updater.processBtn')}
          </Button>
        </CardFooter>
      )}
    </Card>
  );
}
