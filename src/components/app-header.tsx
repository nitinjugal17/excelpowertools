
"use client";

import { useLanguage } from "@/context/language-context";
import { SidebarTrigger } from "@/components/ui/sidebar";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

interface AppHeaderProps {
    activeLabel: string;
}

export function AppHeader({ activeLabel }: AppHeaderProps) {
    const { t, language, setLanguage } = useLanguage();

    return (
        <header className="flex items-center justify-between p-4 border-b">
             <div className="flex items-center gap-2 md:hidden">
                <SidebarTrigger />
                <h1 className="text-lg font-semibold">{activeLabel}</h1>
             </div>
             <div className="hidden md:flex flex-grow justify-center">
                 <h1 className="text-2xl font-semibold">{activeLabel}</h1>
             </div>
             <div className="flex items-center gap-2">
                <Label htmlFor="language-select">{[t('common.language')].flat().join(' ')}</Label>
                <Select value={language} onValueChange={(value) => setLanguage(value as 'en' | 'hi')}>
                    <SelectTrigger id="language-select" className="w-[120px]">
                        <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                        <SelectItem value="en">English</SelectItem>
                        <SelectItem value="hi">हिन्दी</SelectItem>
                    </SelectContent>
                </Select>
             </div>
        </header>
    );
}
