
"use client";

import React from 'react';

export function Markup({ text }: { text: string }) {
    if (!text) return null;
    
    // Regex to split by [code:...] and capture the content, but not split if it's not found
    const parts = text.split(/(\[code:[^\]]+\])/g);
    
    return (
        <>
            {parts.map((part, index) => {
                const codeMatch = part.match(/\[code:(.+)\]/);
                if (codeMatch && codeMatch[1]) {
                    return <code key={index} className="bg-muted px-1 py-0.5 rounded font-mono text-sm">{codeMatch[1]}</code>;
                }
                return part;
            })}
        </>
    );
}

    