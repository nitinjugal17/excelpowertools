
"use client";

import React, { useState, useEffect, useCallback } from 'react';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Label } from '@/components/ui/label';
import { Input } from '@/components/ui/input';
import { Textarea } from '@/components/ui/textarea';
import { Button } from '@/components/ui/button';
import { useToast } from '@/hooks/use-toast';
import { FolderSymlink, Copy, AlertTriangle } from 'lucide-react';
import { useLanguage } from '@/context/language-context';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Markup } from './ui/markup';

interface LocalFileEmbedderPageProps {
  onFileStateChange: (hasFile: boolean) => void;
}

export default function LocalFileEmbedderPage({ onFileStateChange }: LocalFileEmbedderPageProps) {
  const { t } = useLanguage();
  const [filePath, setFilePath] = useState('');
  const [width, setWidth] = useState('800');
  const [height, setHeight] = useState('600');
  const [embedCode, setEmbedCode] = useState('');
  const { toast } = useToast();

  useEffect(() => {
    onFileStateChange?.(false);
  }, [onFileStateChange]);

  const generateEmbedCode = useCallback(() => {
    if (!filePath) {
      setEmbedCode('');
      return;
    }

    // Convert backslashes to forward slashes for the file protocol URI
    const formattedPath = filePath.replace(/\\/g, '/');
    const finalUrl = `file:///${formattedPath}`;
    
    const code = `<iframe width="${width}" height="${height}" frameborder="0" src="${finalUrl}"></iframe>`;
    setEmbedCode(code);
  }, [filePath, width, height]);

  useEffect(() => {
    generateEmbedCode();
  }, [generateEmbedCode]);
  
  const handleCopyToClipboard = () => {
    if (embedCode) {
      navigator.clipboard.writeText(embedCode);
      toast({ title: t('localFileEmbedder.toast.copiedTitle') as string, description: t('localFileEmbedder.toast.copiedDesc') as string });
    }
  };

  return (
    <Card className="w-full max-w-2xl shadow-xl">
      <CardHeader>
        <div className="flex items-center space-x-2 mb-2">
          <FolderSymlink className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">{t('localFileEmbedder.title')}</CardTitle>
        </div>
        <CardDescription className="font-body">
          <Markup text={t('localFileEmbedder.description') as string} />
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="file-path">{t('localFileEmbedder.filePathLabel')}</Label>
          <Input 
            id="file-path" 
            placeholder={t('localFileEmbedder.filePathPlaceholder') as string}
            value={filePath}
            onChange={(e) => setFilePath(e.target.value)}
          />
          <p className="text-xs text-muted-foreground">
            {t('localFileEmbedder.filePathDesc')}
          </p>
        </div>
        
        <Card className="p-4 bg-secondary/30 space-y-4">
          <CardHeader className="p-0 pb-2">
            <CardTitle className="text-lg">{t('localFileEmbedder.embedOptions')}</CardTitle>
          </CardHeader>
          <CardContent className="p-0 grid grid-cols-1 md:grid-cols-2 gap-6">
             <div className="space-y-2">
               <Label htmlFor="embed-width">{t('localFileEmbedder.width')}</Label>
               <Input id="embed-width" value={width} onChange={(e) => setWidth(e.target.value)} />
             </div>
              <div className="space-y-2">
               <Label htmlFor="embed-height">{t('localFileEmbedder.height')}</Label>
               <Input id="embed-height" value={height} onChange={(e) => setHeight(e.target.value)} />
             </div>
          </CardContent>
        </Card>

        <div className="space-y-2">
            <Label htmlFor="embed-code">{t('localFileEmbedder.generatedCode')}</Label>
            <Textarea 
                id="embed-code"
                readOnly
                value={embedCode}
                className="font-code h-32"
                placeholder={t('localFileEmbedder.codePlaceholder') as string}
            />
        </div>
        
        <Alert variant="destructive">
          <AlertTriangle className="h-4 w-4" />
          <AlertTitle>{t('localFileEmbedder.securityWarningTitle')}</AlertTitle>
          <AlertDescription>
            {t('localFileEmbedder.securityWarningDesc')}
          </AlertDescription>
        </Alert>

      </CardContent>
      <CardFooter>
        <Button onClick={handleCopyToClipboard} disabled={!embedCode} className="w-full">
            <Copy className="mr-2 h-4 w-4" />
            {t('localFileEmbedder.copyBtn')}
        </Button>
      </CardFooter>
    </Card>
  );
}
