
"use client";

import React, { useState, useCallback } from 'react';
import { SidebarProvider, SidebarInset } from "@/components/ui/sidebar";
import { AppSidebar } from "@/components/app-sidebar";
import { AppHeader } from "@/components/app-header";

import ExcelSheetSplitterPage from '@/components/excel-sheet-splitter-page';
import ExcelSheetUpdaterPage from '@/components/excel-sheet-updater-page';
import ExcelDataAggregatorPage from '@/components/excel-data-aggregator-page';
import EmptyCellFinderPage from '@/components/empty-cell-finder-page';
import DuplicateFinderPage from '@/components/duplicate-finder-page';
import TextFormatterPage from '@/components/text-formatter-page';
import UserGuidePage from '@/components/user-guide-page';
import DataImputerPage from '@/components/data-imputer-page';
import WorkbookBreakerPage from '@/components/workbook-breaker-page';
import UniqueValueFinderPage from '@/components/unique-value-finder-page';
import DataExtractorPage from '@/components/data-extractor-page';
import SheetMergerPage from '@/components/sheet-merger-page';
import ColumnPurgerPage from '@/components/column-purger-page';
import ExcelComparatorPage from '@/components/excel-comparator-page';
import PivotTableCreatorPage from '@/components/pivot-table-creator-page';

import { FileSpreadsheet, Wand2, Sigma, FileScan, CopyCheck, Paintbrush, BookUser, Sparkles, LibraryBig, Fingerprint, FileSearch, Combine, FileMinus, GitCompareArrows, LayoutGrid } from "lucide-react";
import { useLanguage } from '@/context/language-context';
import type { Tool, ToolInfo } from '@/types/tools';

const toolComponents: Record<Tool, React.ComponentType<any>> = {
  guide: UserGuidePage,
  splitter: ExcelSheetSplitterPage,
  updater: ExcelSheetUpdaterPage,
  aggregator: ExcelDataAggregatorPage,
  merger: SheetMergerPage,
  breaker: WorkbookBreakerPage,
  finder: EmptyCellFinderPage,
  duplicates: DuplicateFinderPage,
  columnPurger: ColumnPurgerPage,
  formatter: TextFormatterPage,
  imputer: DataImputerPage,
  uniqueFinder: UniqueValueFinderPage,
  extractor: DataExtractorPage,
  comparator: ExcelComparatorPage,
  pivot: PivotTableCreatorPage,
};

const toolInfo: Record<Tool, ToolInfo> = {
    guide: { icon: BookUser, labelKey: "sidebar.userGuide" },
    splitter: { icon: FileSpreadsheet, labelKey: "sidebar.sheetSplitter" },
    updater: { icon: Wand2, labelKey: "sidebar.sheetUpdater" },
    aggregator: { icon: Sigma, labelKey: "sidebar.dataAggregator" },
    merger: { icon: Combine, labelKey: "sidebar.sheetMerger" },
    comparator: { icon: GitCompareArrows, labelKey: "sidebar.comparator" },
    breaker: { icon: LibraryBig, labelKey: "sidebar.workbookBreaker" },
    finder: { icon: FileScan, labelKey: "sidebar.emptyCellFinder" },
    duplicates: { icon: CopyCheck, labelKey: "sidebar.duplicateFinder" },
    columnPurger: { icon: FileMinus, labelKey: "sidebar.columnPurger" },
    formatter: { icon: Paintbrush, labelKey: "sidebar.textFormatter" },
    imputer: { icon: Sparkles, labelKey: "sidebar.aiImputer" },
    uniqueFinder: { icon: Fingerprint, labelKey: "sidebar.uniqueValueFinder" },
    extractor: { icon: FileSearch, labelKey: "sidebar.dataExtractor" },
    pivot: { icon: LayoutGrid, labelKey: "sidebar.pivotCreator" },
};

export default function MainLayout() {
  const [activeTool, setActiveTool] = useState<Tool>('guide');
  const [isToolProcessing, setIsToolProcessing] = useState(false);
  const [isFileLoaded, setIsFileLoaded] = useState(false);
  const { t } = useLanguage();

  const handleProcessingChange = useCallback((processing: boolean) => {
    // A timeout helps prevent a race condition where the state updates
    // just as the user clicks, but before the confirmation dialog can appear.
    setTimeout(() => setIsToolProcessing(processing), 50);
  }, []);

  const handleFileStateChange = useCallback((hasFile: boolean) => {
      setIsFileLoaded(hasFile);
  }, []);

  const handleSetActiveTool = (tool: Tool) => {
    const isDirty = isToolProcessing || (isFileLoaded && activeTool !== 'guide');

    if (isDirty) {
      if (window.confirm(t('common.confirmSwitch') as string)) {
        setActiveTool(tool);
        // Reset processing and file state when the user confirms the switch
        setIsToolProcessing(false);
        setIsFileLoaded(false);
      }
    } else {
      setActiveTool(tool);
    }
  };

  const ActiveComponent = toolComponents[activeTool];
  const activeLabel = [t(toolInfo[activeTool].labelKey)].flat().join(' ');

  const activeComponentProps = {
    onProcessingChange: handleProcessingChange,
    onFileStateChange: handleFileStateChange,
  };

  return (
    <SidebarProvider>
      <AppSidebar
        activeTool={activeTool}
        setActiveTool={handleSetActiveTool}
        toolInfo={toolInfo}
        isProcessing={isToolProcessing || (isFileLoaded && activeTool !== 'guide')}
      />
      <SidebarInset>
        <AppHeader activeLabel={activeLabel} />
        <main className="flex flex-1 flex-col items-center justify-start p-4 md:p-8">
          <ActiveComponent {...activeComponentProps} />
        </main>
      </SidebarInset>
    </SidebarProvider>
  );
}
