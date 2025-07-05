
"use client";

import React, { useEffect } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Lightbulb, FileSpreadsheet, Wand2, Sigma, FileScan, CopyCheck, Paintbrush, BookUser, Sparkles, LibraryBig, Fingerprint, Combine, FileSearch, FileMinus, GitCompareArrows, LayoutGrid } from 'lucide-react';
import { useLanguage } from '@/context/language-context';
import { Markup } from '@/components/ui/markup';

interface UserGuidePageProps {
  onFileStateChange: (hasFile: boolean) => void;
}

export default function UserGuidePage({ onFileStateChange }: UserGuidePageProps) {
  const { t } = useLanguage();
  
  useEffect(() => {
    // This page never has a file loaded, so we signal this on mount.
    if (onFileStateChange) {
      onFileStateChange(false);
    }
  }, [onFileStateChange]);

  const toolSections = [
      {
          id: "splitter",
          icon: FileSpreadsheet,
          titleKey: "userGuide.splitter.title",
          descriptionKey: "userGuide.splitter.description",
          whenToUseKey: "userGuide.splitter.whenToUse",
          howToUseKey: "userGuide.splitter.howToUse",
          proTipKey: "userGuide.splitter.proTip"
      },
      {
          id: "updater",
          icon: Wand2,
          titleKey: "userGuide.updater.title",
          descriptionKey: "userGuide.updater.description",
          whenToUseKey: "userGuide.updater.whenToUse",
          howToUseKey: "userGuide.updater.howToUse",
          proTipKey: "userGuide.updater.proTip"
      },
      {
          id: "aggregator",
          icon: Sigma,
          titleKey: "userGuide.aggregator.title",
          descriptionKey: "userGuide.aggregator.description",
          whenToUseKey: "userGuide.aggregator.whenToUse",
          howToUseKey: "userGuide.aggregator.howToUse",
          proTipKey: "userGuide.aggregator.proTip"
      },
       {
          id: "merger",
          icon: Combine,
          titleKey: "userGuide.merger.title",
          descriptionKey: "userGuide.merger.description",
          whenToUseKey: "userGuide.merger.whenToUse",
          howToUseKey: "userGuide.merger.howToUse",
          proTipKey: "userGuide.merger.proTip"
      },
       {
          id: "comparator",
          icon: GitCompareArrows,
          titleKey: "userGuide.comparator.title",
          descriptionKey: "userGuide.comparator.description",
          whenToUseKey: "userGuide.comparator.whenToUse",
          howToUseKey: "userGuide.comparator.howToUse",
          proTipKey: "userGuide.comparator.proTip"
      },
      {
          id: "breaker",
          icon: LibraryBig,
          titleKey: "userGuide.breaker.title",
          descriptionKey: "userGuide.breaker.description",
          whenToUseKey: "userGuide.breaker.whenToUse",
          howToUseKey: "userGuide.breaker.howToUse",
          proTipKey: "userGuide.breaker.proTip"
      },
      {
          id: "finder",
          icon: FileScan,
          titleKey: "userGuide.finder.title",
          descriptionKey: "userGuide.finder.description",
          whenToUseKey: "userGuide.finder.whenToUse",
          howToUseKey: "userGuide.finder.howToUse",
          proTipKey: "userGuide.finder.proTip"
      },
      {
          id: "duplicates",
          icon: CopyCheck,
          titleKey: "userGuide.duplicates.title",
          descriptionKey: "userGuide.duplicates.description",
          whenToUseKey: "userGuide.duplicates.whenToUse",
          howToUseKey: "userGuide.duplicates.howToUse",
          proTipKey: "userGuide.duplicates.proTip"
      },
      {
          id: "columnPurger",
          icon: FileMinus,
          titleKey: "userGuide.columnPurger.title",
          descriptionKey: "userGuide.columnPurger.description",
          whenToUseKey: "userGuide.columnPurger.whenToUse",
          howToUseKey: "userGuide.columnPurger.howToUse",
          proTipKey: "userGuide.columnPurger.proTip"
      },
      {
          id: "formatter",
          icon: Paintbrush,
          titleKey: "userGuide.formatter.title",
          descriptionKey: "userGuide.formatter.description",
          whenToUseKey: "userGuide.formatter.whenToUse",
          howToUseKey: "userGuide.formatter.howToUse",
          proTipKey: "userGuide.formatter.proTip"
      },
      {
          id: "imputer",
          icon: Sparkles,
          titleKey: "userGuide.imputer.title",
          descriptionKey: "userGuide.imputer.description",
          whenToUseKey: "userGuide.imputer.whenToUse",
          howToUseKey: "userGuide.imputer.howToUse",
          proTipKey: "userGuide.imputer.proTip"
      },
      {
          id: "uniqueFinder",
          icon: Fingerprint,
          titleKey: "userGuide.uniqueFinder.title",
          descriptionKey: "userGuide.uniqueFinder.description",
          whenToUseKey: "userGuide.uniqueFinder.whenToUse",
          howToUseKey: "userGuide.uniqueFinder.howToUse",
          proTipKey: "userGuide.uniqueFinder.proTip"
      },
       {
          id: "extractor",
          icon: FileSearch,
          titleKey: "userGuide.extractor.title",
          descriptionKey: "userGuide.extractor.description",
          whenToUseKey: "userGuide.extractor.whenToUse",
          howToUseKey: "userGuide.extractor.howToUse",
          proTipKey: "userGuide.extractor.proTip"
      },
      {
          id: "pivot",
          icon: LayoutGrid,
          titleKey: "userGuide.pivot.title",
          descriptionKey: "userGuide.pivot.description",
          whenToUseKey: "userGuide.pivot.whenToUse",
          howToUseKey: "userGuide.pivot.howToUse",
          proTipKey: "userGuide.pivot.proTip"
      }
  ];

  return (
    <Card className="w-full max-w-4xl shadow-xl">
      <CardHeader>
        <CardTitle className="text-3xl font-headline flex items-center">
            <BookUser className="mr-3 h-8 w-8 text-primary"/>
            {[t('userGuide.title')].flat().join(' ')}
        </CardTitle>
        <CardDescription className="font-body pt-2">
            {[t('userGuide.description')].flat().join(' ')}
        </CardDescription>
      </CardHeader>
      <CardContent>
        <Accordion type="single" collapsible className="w-full">
            {toolSections.sort((a,b) => [t(a.titleKey)].flat().join(' ').localeCompare([t(b.titleKey)].flat().join(' '))).map(tool => {
              const howToUseSteps = t(tool.howToUseKey, {}) as string[];
              return (
                 <AccordionItem value={tool.id} key={tool.id}>
                    <AccordionTrigger className="text-xl hover:no-underline">
                        <div className="flex items-center">
                            <tool.icon className="mr-3 h-6 w-6 text-primary/80"/>
                            {[t(tool.titleKey)].flat().join(' ')}
                        </div>
                    </AccordionTrigger>
                    <AccordionContent className="pt-4 pl-4 border-l-2 ml-4 border-primary/20">
                        <div className="space-y-6">
                            <div className="space-y-2">
                                <h3 className="font-semibold text-lg">{[t('userGuide.whatItDoes')].flat().join(' ')}</h3>
                                <p className="text-muted-foreground">{[t(tool.descriptionKey)].flat().join(' ')}</p>
                            </div>
                             <div className="space-y-2">
                                <h3 className="font-semibold text-lg">{[t('userGuide.whenToUse')].flat().join(' ')}</h3>
                                <p className="text-muted-foreground">{[t(tool.whenToUseKey)].flat().join(' ')}</p>
                            </div>
                             <div className="space-y-2">
                                <h3 className="font-semibold text-lg">{[t('userGuide.howToUse')].flat().join(' ')}</h3>
                                <ol className="list-decimal list-inside space-y-2 text-muted-foreground">
                                    {Array.isArray(howToUseSteps) && howToUseSteps.map((step, index) => <li key={index}><Markup text={step} /></li>)}
                                </ol>
                            </div>
                             <Alert>
                                <Lightbulb className="h-4 w-4" />
                                <AlertTitle>{[t('userGuide.proTip')].flat().join(' ')}</AlertTitle>
                                <AlertDescription>
                                    <Markup text={[t(tool.proTipKey)].flat().join(' ')} />
                                </AlertDescription>
                            </Alert>
                        </div>
                    </AccordionContent>
                </AccordionItem>
            )})}
        </Accordion>
      </CardContent>
    </Card>
  );
}
