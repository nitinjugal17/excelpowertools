
"use client";

import React, { useState, useCallback, ChangeEvent } from 'react';
import * as XLSX from 'xlsx-js-style';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { useToast } from '@/hooks/use-toast';
import { UploadCloud, Download, FileCode, CheckCircle2, Loader2, List, Link } from 'lucide-react';
import { useLanguage } from '@/context/language-context';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Alert, AlertDescription } from '@/components/ui/alert';

interface HtmlToExcelPageProps {
  onProcessingChange: (isProcessing: boolean) => void;
  onFileStateChange: (hasFile: boolean) => void;
}

type OperationMode = 'extract' | 'hyperlink';

export default function HtmlToExcelPage({ onProcessingChange, onFileStateChange }: HtmlToExcelPageProps) {
  const { t } = useLanguage();
  const [files, setFiles] = useState<FileList | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processedWorkbook, setProcessedWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [operationMode, setOperationMode] = useState<OperationMode>('extract');
  const [basePath, setBasePath] = useState('');
  const { toast } = useToast();

  React.useEffect(() => {
    onProcessingChange?.(isProcessing);
  }, [isProcessing, onProcessingChange]);

  React.useEffect(() => {
    onFileStateChange?.(files !== null && files.length > 0);
  }, [files, onFileStateChange]);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    setFiles(event.target.files);
    setProcessedWorkbook(null);
  };

  const handleProcess = useCallback(async () => {
    if (!files || files.length === 0) {
      toast({ title: "No Files Selected", description: "Please upload one or more HTML files to process.", variant: 'destructive' });
      return;
    }
    if (operationMode === 'hyperlink' && !basePath.trim()) {
      toast({ title: "Base Path Required", description: "Please provide the base file path for hyperlink mode.", variant: 'destructive' });
      return;
    }

    setIsProcessing(true);
    const newWb = XLSX.utils.book_new();

    if (operationMode === 'extract') {
      let filesProcessed = 0;
      for (const file of Array.from(files)) {
        if (!file.name.toLowerCase().endsWith('.html') && !file.name.toLowerCase().endsWith('.htm')) {
          toast({ title: "Invalid File Type", description: `Skipping non-HTML file: ${file.name}`, variant: 'destructive' });
          continue;
        }

        try {
          const htmlString = await file.text();
          const parser = new DOMParser();
          const doc = parser.parseFromString(htmlString, 'text/html');
          const tables = doc.querySelectorAll('table');
          
          if (tables.length === 0) {
              console.warn(`No tables found in ${file.name}, skipping.`);
              continue;
          }
          
          const mainTable = tables[0];
          const data: any[][] = [];
          
          const thead = mainTable.querySelector('thead');
          if (thead) {
              thead.querySelectorAll('tr').forEach(headerRow => {
                  const rowData: string[] = [];
                  headerRow.querySelectorAll('th').forEach(th => rowData.push(th.innerText));
                  data.push(rowData);
              });
          }
          
          const tbody = mainTable.querySelector('tbody');
          if (tbody) {
              tbody.querySelectorAll('tr').forEach(bodyRow => {
                  const rowData: string[] = [];
                  bodyRow.querySelectorAll('td').forEach(td => rowData.push(td.innerText));
                  data.push(rowData);
              });
          }
          
          const tfoot = mainTable.querySelector('tfoot');
          if (tfoot) {
              tfoot.querySelectorAll('tr').forEach(footerRow => {
                  const rowData: string[] = [];
                  footerRow.querySelectorAll('td, th').forEach(cell => rowData.push(cell.innerText));
                  data.push(rowData);
              });
          }

          const ws = XLSX.utils.aoa_to_sheet(data);
          const sheetName = file.name.replace(/\.html?$/i, '').substring(0, 31);
          XLSX.utils.book_append_sheet(newWb, ws, sheetName);
          filesProcessed++;

        } catch (error) {
          console.error(`Error processing file ${file.name}:`, error);
          toast({ title: "Processing Error", description: `Could not process file: ${file.name}`, variant: 'destructive' });
        }
      }
      
      if (filesProcessed > 0) {
          setProcessedWorkbook(newWb);
          toast({ title: "Processing Complete", description: `${filesProcessed} HTML file(s) have been converted to Excel sheets.`, action: <CheckCircle2 className="text-green-500" /> });
      } else {
          toast({ title: "No Data Processed", description: "No tables could be extracted from the selected files.", variant: "destructive" });
      }

    } else { // Hyperlink Mode
        const aoa = [
            ['File Name', 'Link to File']
        ];
        const finalBasePath = basePath.endsWith('/') || basePath.endsWith('\\') ? basePath : basePath + '/';

        for (const file of Array.from(files)) {
             const fullPath = `${finalBasePath}${file.name}`;
             const linkCell = { 
                v: 'Open File', 
                l: { Target: fullPath, Tooltip: `Click to open ${fullPath}` } ,
                s: { font: { color: { rgb: "0000FF" }, underline: true } }
            };
            aoa.push([file.name, linkCell]);
        }
        
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!cols'] = [{ wch: 40 }, { wch: 20 }];
        XLSX.utils.book_append_sheet(newWb, ws, "File Index");
        setProcessedWorkbook(newWb);
        toast({ title: "Processing Complete", description: `Index file created with ${files.length} hyperlinks.`, action: <CheckCircle2 className="text-green-500" /> });
    }

    setIsProcessing(false);
  }, [files, toast, operationMode, basePath]);

  const handleDownload = useCallback(() => {
    if (!processedWorkbook) {
      toast({ title: "No Data to Download", variant: 'destructive' });
      return;
    }
    const fileName = operationMode === 'extract' ? "Generated_From_HTML.xlsx" : "HTML_File_Index.xlsx";
    XLSX.writeFile(processedWorkbook, fileName);
  }, [processedWorkbook, toast, operationMode]);

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
          <FileCode className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">HTML to Excel Creator</CardTitle>
        </div>
        <CardDescription className="font-body">Upload HTML files to extract their table data into Excel sheets or create an index file with hyperlinks.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-upload-html" className="flex items-center space-x-2 text-sm font-medium">
            <UploadCloud className="h-5 w-5" />
            <span>1. Upload HTML Files</span>
          </Label>
          <Input
            id="file-upload-html"
            type="file"
            accept=".html,.htm"
            onChange={handleFileChange}
            className="file:text-primary file:font-semibold file:bg-primary/10 file:border-0 hover:file:bg-primary/20"
            disabled={isProcessing}
            multiple
          />
        </div>

        <div className="space-y-2">
            <Label className="text-sm font-medium">2. Choose Operation Mode</Label>
            <RadioGroup value={operationMode} onValueChange={(v) => setOperationMode(v as OperationMode)} className="space-y-2">
                <Label htmlFor="mode-extract" className="flex items-start space-x-3 p-4 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer">
                    <RadioGroupItem value="extract" id="mode-extract" className="mt-1"/>
                    <div className="grid gap-1.5">
                        <span className="font-semibold">Extract Table Data</span>
                        <p className="text-xs text-muted-foreground">Reads tables from each HTML file and creates a new sheet for each in the output Excel file.</p>
                    </div>
                </Label>
                <Label htmlFor="mode-hyperlink" className="flex items-start space-x-3 p-4 border rounded-md has-[:checked]:border-primary has-[:checked]:bg-primary/10 cursor-pointer">
                    <RadioGroupItem value="hyperlink" id="mode-hyperlink" className="mt-1"/>
                    <div className="grid gap-1.5">
                        <span className="font-semibold">Insert Hyperlinks</span>
                        <p className="text-xs text-muted-foreground">Creates a single sheet listing all uploaded files, with hyperlinks pointing to their local path.</p>
                    </div>
                </Label>
            </RadioGroup>
        </div>

        {operationMode === 'hyperlink' && (
          <div className="space-y-2">
              <Label htmlFor="base-path" className="flex items-center space-x-2 text-sm font-medium">
                <Link className="h-4 w-4" />
                <span>Base File Path</span>
              </Label>
              <Input
                id="base-path"
                value={basePath}
                onChange={e => setBasePath(e.target.value)}
                placeholder="e.g., C:\Users\YourName\Documents\Reports\"
                disabled={isProcessing}
              />
              <Alert variant="destructive" className="mt-2">
                <AlertDescription>
                  This mode will only work if the end-user has the HTML files stored at this exact path on their computer or a mapped network drive.
                </AlertDescription>
              </Alert>
          </div>
        )}

        {files && files.length > 0 && (
          <div className="space-y-2">
            <Label className="flex items-center space-x-2 text-sm font-medium">
              <List className="h-5 w-5" />
              <span>Selected Files</span>
            </Label>
            <Card className="max-h-40 overflow-y-auto p-3 bg-secondary/20">
              <ul className="space-y-1 text-sm text-muted-foreground">
                {Array.from(files).map((file, index) => (
                  <li key={index} className="truncate">{file.name}</li>
                ))}
              </ul>
            </Card>
          </div>
        )}

        <Button onClick={handleProcess} disabled={isProcessing || !files || files.length === 0} className="w-full">
            <FileCode className="mr-2 h-5 w-5" />
            Process HTML and Create Excel File
        </Button>
      </CardContent>
      {processedWorkbook && (
        <CardFooter>
          <Button onClick={handleDownload} variant="outline" className="w-full">
            <Download className="mr-2 h-5 w-5" />
            Download Generated Excel File
          </Button>
        </CardFooter>
      )}
    </Card>
  );
}
