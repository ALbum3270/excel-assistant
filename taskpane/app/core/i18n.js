// Pane language. One i18next instance serves the plain-JS core, the React view
// (through react-i18next) and the vendored AI Elements components (patched to
// call `t` on their few fixed English strings, which are their own keys).
import i18next from "i18next";
import en from "./locales/en.js";
import zh from "./locales/zh.js";

export const LANGUAGES = ["zh", "en"];

export const i18n = i18next.createInstance();
i18n.init({
  lng: "en",
  fallbackLng: "en",
  // Resources are inline, so t() works as soon as this module loads.
  initAsync: false,
  resources: { en: { translation: en }, zh: { translation: zh } },
  // Keys are ids or English sentences, never paths.
  keySeparator: false,
  nsSeparator: false,
  // React escapes what it renders; model and daemon text never goes through t().
  interpolation: { escapeValue: false },
  returnNull: false,
});

export const t = (key, options) => i18n.t(key, options);

// "auto" follows Office's display language (e.g. "zh-CN", "en-US").
export function resolveLanguage(setting, hostLanguage) {
  if (LANGUAGES.includes(setting)) return setting;
  return /^zh\b/i.test(hostLanguage || "") ? "zh" : "en";
}

export function setLanguage(language) {
  if (i18n.language !== language) i18n.changeLanguage(language);
  if (typeof document !== "undefined") {
    document.documentElement.lang = language === "zh" ? "zh-CN" : "en";
  }
}
