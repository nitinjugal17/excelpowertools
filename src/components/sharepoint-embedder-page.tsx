
"use client";

import React, { useState, useEffect, useCallback } from 'react';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Label } from '@/components/ui/label';
import { Input } from '@/components/ui/input';
import { Checkbox } from '@/components/ui/checkbox';
import { Textarea } from '@/components/ui/textarea';
import { Button } from '@/components/ui/button';
import { useToast } from '@/hooks/use-toast';
import { Share2, Copy } from 'lucide-react';
import { useLanguage } from '@/context/language-context';

interface SharePointEmbedderPageProps {
  onFileStateChange: (hasFile: boolean) => void;
}

export default function SharePointEmbedderPage({ onFileStateChange }: SharePointEmbedderPageProps) {
  const { t } = useLanguage();
  const [shareUrl, setShareUrl] = useState('');
  const [width, setWidth] = useState('500');
  const [height, setHeight] = useState('200');
  const [allowInteractivity, setAllowInteractivity] = useState(true);
  const [hideSheetTabs, setHideSheetTabs] = useState(true);
  const [hideHeaders, setHideHeaders] = useState(true);
  const [namedItem, setNamedItem] = useState('dashboard');
  const [activeCell, setActiveCell] = useState('A1');
  const [embedCode, setEmbedCode] = useState('');
  const { toast } = useToast();

  useEffect(() => {
    onFileStateChange?.(false);
  }, [onFileStateChange]);

  const generateEmbedCode = useCallback(() => {
    if (!shareUrl) {
      setEmbedCode('');
      return;
    }

    let finalUrl = shareUrl.split('?')[0];
    if (!finalUrl) return;
    
    // Ensure the base URL ends with the correct embed action preamble
    const embedAction = "action=embedview";
    if (finalUrl.includes('?')) {
        finalUrl += `&${embedAction}`;
    } else {
        finalUrl += `?${embedAction}`;
    }
    
    finalUrl += '&wdbipreview=True';

    if (allowInteractivity) finalUrl += '&wdAllowInteractivity=True';
    if (hideSheetTabs) finalUrl += '&wdHideSheetTabs=True';
    if (hideHeaders) finalUrl += '&wdHideHeaders=True';
    if (namedItem.trim()) finalUrl += `&item=${encodeURIComponent(namedItem.trim())}`;
    if (activeCell.trim()) finalUrl += `&activeCell=${encodeURIComponent(activeCell.trim())}`;
    
    const code = `<iframe width="${width}" height="${height}" frameborder="0" scrolling="no" src="${finalUrl}"></iframe>`;
    setEmbedCode(code);
  }, [shareUrl, width, height, allowInteractivity, hideSheetTabs, hideHeaders, namedItem, activeCell]);

  useEffect(() => {
    generateEmbedCode();
  }, [generateEmbedCode]);
  
  const handleCopyToClipboard = () => {
    if (embedCode) {
      navigator.clipboard.writeText(embedCode);
      toast({ title: 'Copied!', description: 'The iframe embed code has been copied to your clipboard.' });
    }
  };

  return (
    <Card className="w-full max-w-2xl shadow-xl">
      <CardHeader>
        <div className="flex items-center space-x-2 mb-2">
          <Share2 className="h-8 w-8 text-primary" />
          <CardTitle className="text-2xl font-headline">SharePoint Embedder</CardTitle>
        </div>
        <CardDescription className="font-body">
          Generate an iframe embed code for a SharePoint Excel file with specific viewing options.
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-2">
          <Label htmlFor="share-url">SharePoint Share Link</Label>
          <Input 
            id="share-url" 
            placeholder="https://yourDomain.sharepoint.com/:x:/g/..." 
            value={shareUrl}
            onChange={(e) => setShareUrl(e.target.value)}
          />
          <p className="text-xs text-muted-foreground">
            Paste the full 'Anyone with the link' share URL for your Excel file here.
          </p>
        </div>
        
        <Card className="p-4 bg-secondary/30 space-y-4">
          <CardHeader className="p-0 pb-2">
            <CardTitle className="text-lg">Embed Options</CardTitle>
          </CardHeader>
          <CardContent className="p-0 grid grid-cols-1 md:grid-cols-2 gap-6">
             <div className="space-y-4">
               <div className="space-y-2">
                 <Label htmlFor="embed-width">Width</Label>
                 <Input id="embed-width" value={width} onChange={(e) => setWidth(e.target.value)} />
               </div>
                <div className="space-y-2">
                 <Label htmlFor="embed-height">Height</Label>
                 <Input id="embed-height" value={height} onChange={(e) => setHeight(e.target.value)} />
               </div>
                <div className="space-y-2">
                 <Label htmlFor="named-item">Named Item</Label>
                 <Input id="named-item" value={namedItem} onChange={(e) => setNamedItem(e.target.value)} placeholder="e.g., dashboard, SalesRange" />
               </div>
                <div className="space-y-2">
                 <Label htmlFor="active-cell">Active Cell</Label>
                 <Input id="active-cell" value={activeCell} onChange={(e) => setActiveCell(e.target.value)} placeholder="e.g., A1" />
               </div>
             </div>
             <div className="space-y-4 pt-2">
                 <div className="flex items-center space-x-2">
                    <Checkbox id="allow-interactivity" checked={allowInteractivity} onCheckedChange={(c) => setAllowInteractivity(c as boolean)} />
                    <Label htmlFor="allow-interactivity">Allow Interactivity</Label>
                 </div>
                 <p className="text-xs text-muted-foreground -mt-2 pl-6">Allows users to use slicers and filters.</p>
                 <div className="flex items-center space-x-2">
                    <Checkbox id="hide-tabs" checked={hideSheetTabs} onCheckedChange={(c) => setHideSheetTabs(c as boolean)} />
                    <Label htmlFor="hide-tabs">Hide Sheet Tabs</Label>
                 </div>
                 <div className="flex items-center space-x-2">
                    <Checkbox id="hide-headers" checked={hideHeaders} onCheckedChange={(c) => setHideHeaders(c as boolean)} />
                    <Label htmlFor="hide-headers">Hide Headers</Label>
                 </div>
             </div>
          </CardContent>
        </Card>

        <div className="space-y-2">
            <Label htmlFor="embed-code">Generated Embed Code</Label>
            <Textarea 
                id="embed-code"
                readOnly
                value={embedCode}
                className="font-code h-32"
                placeholder="iframe code will appear here..."
            />
        </div>
      </CardContent>
      <CardFooter>
        <Button onClick={handleCopyToClipboard} disabled={!embedCode} className="w-full">
            <Copy className="mr-2 h-4 w-4" />
            Copy to Clipboard
        </Button>
      </CardFooter>
    </Card>
  );
}
