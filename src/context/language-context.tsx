"use client";

import React, { createContext, useState, useContext, ReactNode, useCallback, useEffect } from 'react';
import enTranslations from '@/lib/i18n/en.json';
import hiTranslations from '@/lib/i18n/hi.json';

const translations = {
    en: enTranslations,
    hi: hiTranslations,
};

type Language = 'en' | 'hi';

function getTranslation(
  language: Language,
  key: string,
  replacements?: { [key: string]: string | number }
): string | string[] {
    const keys = key.split('.');
    let translation: any = translations[language];

    for (const k of keys) {
      translation = translation?.[k];
      if (translation === undefined) {
        break;
      }
    }
    
    if (translation === undefined) {
        if (language !== 'en') {
            translation = translations.en;
            for (const k of keys) {
                translation = translation?.[k];
                if (translation === undefined) {
                    return key;
                }
            }
        } else {
             return key;
        }
    }
    
    if (typeof translation === 'string') {
        if (replacements) {
            return Object.keys(replacements).reduce((acc, rKey) => {
                const value = String(replacements[rKey]);
                return acc.replace(new RegExp(`\\{${rKey}\\}`, 'g'), value);
            }, translation);
        }
        return translation;
    }
    
    if (Array.isArray(translation)) {
        if (replacements) {
            return translation.map(item => {
                return Object.keys(replacements!).reduce((acc, rKey) => {
                    const value = String(replacements![rKey]);
                    return acc.replace(new RegExp(`\\{${rKey}\\}`, 'g'), value);
                }, item);
            });
        }
        return translation;
    }

    return key;
}

interface LanguageContextType {
  language: Language;
  setLanguage: (language: Language) => void;
  t: (key: string, replacements?: { [key: string]: string | number }) => string | string[];
}

const LanguageContext = createContext<LanguageContextType | undefined>(undefined);

function getInitialLanguage(): Language {
    if (typeof window !== 'undefined') {
        const savedLanguage = localStorage.getItem('app-language');
        if (savedLanguage === 'en' || savedLanguage === 'hi') {
            return savedLanguage;
        }
    }
    return 'en';
}

export function LanguageProvider({ children }: { children: ReactNode }) {
  const [language, setLanguageState] = useState<Language>('en');

  useEffect(() => {
    setLanguageState(getInitialLanguage());
  }, []);

  const setLanguage = useCallback((lang: Language) => {
    try {
        localStorage.setItem('app-language', lang);
        setLanguageState(lang);
    } catch (error) {
        console.warn('Could not save language preference to localStorage.');
        setLanguageState(lang);
    }
  }, []);

  const t = useCallback((key: string, replacements?: { [key: string]: string | number }) => {
    return getTranslation(language, key, replacements);
  }, [language]);


  return (
    <LanguageContext.Provider value={{ language, setLanguage, t }}>
      {children}
    </LanguageContext.Provider>
  );
}

export function useLanguage() {
  const context = useContext(LanguageContext);
  if (context === undefined) {
    throw new Error('useLanguage must be used within a LanguageProvider');
  }
  return context;
}
