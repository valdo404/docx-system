import { ui, defaultLang, type Lang } from './ui';

export function getLangFromUrl(url: URL): Lang {
  const [, lang] = url.pathname.split('/');
  if (lang in ui) return lang as Lang;
  return defaultLang;
}

const currentYear = new Date().getFullYear().toString();

export function useTranslations(lang: Lang) {
  return function t(key: keyof (typeof ui)[typeof defaultLang]): string | readonly string[] {
    const value = ui[lang][key] ?? ui[defaultLang][key];
    if (typeof value === 'string') {
      return value.replace(/\{\{year\}\}/g, currentYear);
    }
    return value;
  };
}

export function getLocalizedPath(path: string, lang: Lang): string {
  if (lang === defaultLang) return path;
  return `/en${path}`;
}
