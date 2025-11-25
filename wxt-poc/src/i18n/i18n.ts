type Messages = Record<string, string>;

const loaded = import.meta.glob('./locales/*.json', { eager: true, import: 'default' }) as Record<string, Messages>;
const catalogs: Record<string, Messages> = {};
for (const [path, msgs] of Object.entries(loaded)) {
  const lang = path.match(/\/([^/]+)\.json$/)?.[1];
  if (lang) catalogs[lang] = msgs;
}

let currentLang = 'en';
let currentMessages: Messages = catalogs.en || {};

export async function loadLocale(lang: string) {
  if (catalogs[lang]) {
    currentMessages = catalogs[lang];
    currentLang = lang;
    applyDir(lang);
    return;
  }
  currentMessages = catalogs.en || {};
  currentLang = 'en';
  applyDir('en');
}

export function t(key: string, params: Record<string, string | number> = {}) {
  const raw = currentMessages[key] ?? catalogs.en?.[key] ?? key;
  return Object.entries(params).reduce((acc, [k, v]) => acc.replace(new RegExp(`{${k}}`, 'g'), String(v)), raw);
}

export function setLanguage(lang: string) {
  return loadLocale(lang);
}

export function getLanguage() {
  return currentLang;
}

function applyDir(lang: string) {
  const rtl = ['ar', 'he'].includes(lang.toLowerCase());
  const dir = rtl ? 'rtl' : 'ltr';
  if (typeof document !== 'undefined') {
    document.documentElement.lang = lang;
    document.documentElement.dir = dir;
  }
}
