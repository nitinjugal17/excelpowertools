
"use client";

import {
  Sidebar,
  SidebarHeader,
  SidebarContent,
  SidebarMenu,
  SidebarMenuItem,
  SidebarMenuButton,
} from "@/components/ui/sidebar";
import { CardTitle } from '@/components/ui/card';
import { useLanguage } from '@/context/language-context';
import type { Tool, ToolInfo } from '@/types/tools';

interface AppSidebarProps {
    activeTool: Tool;
    setActiveTool: (tool: Tool) => void;
    toolInfo: Record<Tool, ToolInfo>;
    isProcessing: boolean;
}

export function AppSidebar({ activeTool, setActiveTool, toolInfo, isProcessing }: AppSidebarProps) {
    const { t } = useLanguage();

    return (
        <Sidebar>
            <SidebarHeader>
                <CardTitle className="p-2 text-xl">{[t('sidebar.title')].flat().join(' ')}</CardTitle>
            </SidebarHeader>
            <SidebarContent>
                <SidebarMenu>
                    {(Object.keys(toolInfo) as Tool[]).map((toolKey) => {
                        const { icon: Icon, labelKey } = toolInfo[toolKey];
                        const label = [t(labelKey)].flat().join(' ');
                        return (
                            <SidebarMenuItem key={toolKey}>
                                <SidebarMenuButton
                                    onClick={() => setActiveTool(toolKey)}
                                    isActive={activeTool === toolKey}
                                    disabled={isProcessing && activeTool !== toolKey}
                                    tooltip={{ children: label, side: 'right', align: 'center' }}
                                >
                                    <Icon />
                                    <span>{label}</span>
                                </SidebarMenuButton>
                            </SidebarMenuItem>
                        )
                    })}
                </SidebarMenu>
            </SidebarContent>
        </Sidebar>
    );
}
