
import type React from 'react';
import type { LucideIcon } from 'lucide-react';

export type Tool = 'guide' | 'splitter' | 'updater' | 'aggregator' | 'finder' | 'duplicates' | 'formatter' | 'imputer' | 'breaker' | 'uniqueFinder' | 'extractor' | 'merger' | 'columnPurger' | 'comparator' | 'pivot' | 'webExporter' | 'htmlToExcel' | 'sharepointEmbedder' | 'localFileEmbedder';

export interface ToolInfo {
    icon: LucideIcon;
    labelKey: string;
}
