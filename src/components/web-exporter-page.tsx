
"use client";

import React, { useState, useCallback, ChangeEvent, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, Globe, CheckCircle2, Loader2, ListChecks, Shield } from 'lucide-react';
import { useLanguage } from '@/context/language-context';
import { generateCombinedHtmlPage } from '@/lib/excel-web-exporter';
import { Checkbox } from './ui/checkbox';
import { ScrollArea } from './ui/scroll-area';
import { RadioGroup, RadioGroupItem } from './ui/radio-group';

type ExportMode = "multiple" | "single";

interface WebExporterPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

export default function WebExporterPage({ onProcessingChange, onFileStateChange }: WebExporterPageProps) {
  const { t } = useLanguage();
  const [file, setFile] = useState<File | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<Record<string, boolean>>({});
  
  const [enablePassword, setEnablePassword] = useState(false);
  const [fullAccessPassword, setFullAccessPassword] = useState('');
  const [maskedAccessPassword, setMaskedAccessPassword] = useState('');

  const [exportMode, setExportMode] = useState<ExportMode>("multiple");
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const { toast } = useToast();

  useEffect(() => {
    onProcessingChange?.(isProcessing);
  }, [isProcessing, onProcessingChange]);

  useEffect(() => {
    onFileStateChange?.(file !== null);
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
            const allSelected: Record<string, boolean> = {};
            for (const name of workbook.SheetNames) {
                allSelected[name] = true;
            }
            setSelectedSheets(allSelected);
        } catch (error) {
            toast({ title: "Error reading file", description: "Could not read sheet names from the selected file.", variant: 'destructive' });
            setSheetNames([]);
            setSelectedSheets({});
        } finally {
            setIsProcessing(false);
        }
    } else {
        setSheetNames([]);
        setSelectedSheets({});
    }
  };

  const handleSheetSelectionChange = (sheetName: string, checked: boolean) => {
    setSelectedSheets(prev => ({ ...prev, [sheetName]: checked }));
  };

  const handleSelectAllSheets = (checked: boolean) => {
    const newSelection: Record<string, boolean> = {};
    sheetNames.forEach(name => {
      newSelection[name] = checked;
    });
    setSelectedSheets(newSelection);
  };
  
  const generateAndDownloadHtml = async () => {
    const sheetsToExport = Object.keys(selectedSheets).filter(name => selectedSheets[name]);

    if (!file || sheetsToExport.length === 0) {
        toast({ title: "Missing Information", description: "Please upload a file and select at least one sheet to export.", variant: 'destructive' });
        return;
    }
    
    if (enablePassword && !fullAccessPassword && !maskedAccessPassword) {
        toast({ title: "Missing Password", description: "Please enter at least one password to protect the exported files.", variant: 'destructive' });
        return;
    }

    setIsProcessing(true);
    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'buffer', cellStyles: true });

        if (exportMode === "multiple") {
            for (const sheetName of sheetsToExport) {
                // For multiple files, we generate one HTML for each selected sheet
                const htmlContent = await generateCombinedHtmlPage(
                    workbook,
                    [sheetName], // Process one sheet at a time
                    file.name,
                    enablePassword ? fullAccessPassword : undefined,
                    enablePassword ? maskedAccessPassword : undefined
                );

                const blob = new Blob([htmlContent], { type: 'text/html' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `${sheetName}.html`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }
        } else { // Single file mode
            // For a single file, we pass all selected sheets to be included in one HTML
            const finalHtml = await generateCombinedHtmlPage(
                workbook,
                sheetsToExport, // Process all selected sheets
                file.name,
                enablePassword ? fullAccessPassword : undefined,
                enablePassword ? maskedAccessPassword : undefined
            );

            const blob = new Blob([finalHtml], { type: 'text/html' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${file.name.substring(0, file.name.lastIndexOf('.'))}_export.html`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }
        
        toast({ title: "Export Successful", description: `HTML file(s) have been downloaded.`, action: <CheckCircle2 className="text-green-500" /> });

    } catch (error) {
        console.error("Error generating HTML file(s):", error);
        const errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
        toast({ title: "Export Failed", description: `An error occurred while generating the HTML file(s): ${errorMessage}`, variant: "destructive" });
    } finally {
        setIsProcessing(false);
    }
  };

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
          <Globe className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">Web Page Exporter</CardTitle>
        </div>
        <CardDescription className="font-body">Export sheets to secure, self-contained HTML files for read-only viewing.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-web-exporter" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>1. Upload Your Excel File</span>
          </Label>
          <Input
            id="file-upload-web-exporter"
            type="file"
            accept=".xlsx, .xls, .xlsm"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isProcessing}
          />
        </div>

        {sheetNames.length > 0 && (
          <div className="space-y-6">
            <div className="space-y-3">
                <Label className="flex items-center space-x-2 text-sm font-medium mb-2">
                <ListChecks className="h-5 w-5" />
                <span>2. Select Sheets to Export</span>
                </Label>
                <div className="flex items-center space-x-2 mb-2 p-2 border rounded-md bg-secondary/20">
                <Checkbox
                    id="select-all-sheets-exporter"
                    checked={allSheetsSelected}
                    onCheckedChange={(checked) => handleSelectAllSheets(checked as boolean)}
                    disabled={isProcessing}
                />
                <Label htmlFor="select-all-sheets-exporter" className="text-sm font-medium flex-grow">
                    {t('common.selectAll')} ({t('common.selectedCount', {selected: Object.values(selectedSheets).filter(Boolean).length, total: sheetNames.length})})
                </Label>
                </div>
                <Card className="max-h-48 overflow-y-auto p-3 bg-background">
                    <ScrollArea className="h-full">
                        <div className="space-y-2">
                            {sheetNames.map(name => (
                            <div key={name} className="flex items-center space-x-2">
                                <Checkbox
                                id={`sheet-exporter-${name}`}
                                checked={selectedSheets[name] || false}
                                onCheckedChange={(checked) => handleSheetSelectionChange(name, checked as boolean)}
                                disabled={isProcessing}
                                />
                                <Label htmlFor={`sheet-exporter-${name}`} className="text-sm font-normal">{name}</Label>
                            </div>
                            ))}
                        </div>
                    </ScrollArea>
                </Card>
            </div>
            
            <div className="space-y-2">
                <Label className="text-sm font-medium">3. Export Mode</Label>
                <RadioGroup value={exportMode} onValueChange={(v) => setExportMode(v as ExportMode)} className="grid grid-cols-1 md:grid-cols-2 gap-2">
                    <Label htmlFor="mode-multiple" className="p-2 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer flex items-center justify-center gap-2">
                        <RadioGroupItem value="multiple" id="mode-multiple" />
                        Multiple HTML Files
                    </Label>
                    <Label htmlFor="mode-single" className="p-2 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer flex items-center justify-center gap-2">
                        <RadioGroupItem value="single" id="mode-single" />
                        Single HTML File
                    </Label>
                </RadioGroup>
            </div>


            <Card className="p-4 border-dashed border-primary/50 bg-primary/5">
                <CardHeader className="p-0 pb-4 flex-row items-center space-x-3 space-y-0">
                    <Checkbox
                        id="enable-password"
                        checked={enablePassword}
                        onCheckedChange={(checked) => setEnablePassword(checked as boolean)}
                        disabled={isProcessing}
                    />
                    <Label htmlFor="enable-password" className="flex items-center space-x-2 text-md font-semibold text-primary">
                        <Shield className="h-5 w-5" />
                        <span>4. Password Protection (Optional)</span>
                    </Label>
                </CardHeader>
                {enablePassword && (
                    <CardContent className="p-0">
                        <div className="space-y-4 pl-8 border-l-2 border-primary/30 ml-2 pt-4">
                            <div className="space-y-1">
                                <Label htmlFor="full-access-password">Full Access Password</Label>
                                <Input 
                                    id="full-access-password" 
                                    type="password" 
                                    value={fullAccessPassword}
                                    onChange={(e) => setFullAccessPassword(e.target.value)}
                                    placeholder="Password for viewing real data"
                                />
                                <p className="text-xs text-muted-foreground">Unlocks the actual data for viewing and console access.</p>
                            </div>
                            <div className="space-y-1">
                                <Label htmlFor="masked-access-password">Masked Data Password</Label>
                                <Input 
                                    id="masked-access-password" 
                                    type="password" 
                                    value={maskedAccessPassword}
                                    onChange={(e) => setMaskedAccessPassword(e.target.value)}
                                    placeholder="Password for viewing masked data"
                                />
                                <p className="text-xs text-muted-foreground">Unlocks a view where numbers are replaced with random values.</p>
                            </div>
                        </div>
                    </CardContent>
                )}
            </Card>

          </div>
        )}
      </CardContent>
      <CardFooter>
        <Button onClick={generateAndDownloadHtml} disabled={isProcessing || !file || Object.values(selectedSheets).filter(s => s).length === 0} className="w-full">
            <Download className="mr-2 h-5 w-5" />
            Generate and Download HTML File(s)
        </Button>
      </CardFooter>
    </Card>
  );
}
