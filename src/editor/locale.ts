import zhCN from "./locales/zh-CN";
import zhTW from "./locales/zh-TW";
import enUS from "./locales/en-US";
import jaJP from "./locales/ja-JP";
import koKR from "./locales/ko-KR";

export const supportedLocales = [
  "zh-CN",
  "zh-TW",
  "en-US",
  "ja-JP",
  "ko-KR",
] as const;
export type Locale = (typeof supportedLocales)[number];
export type Label = keyof typeof zhCN.labels;
export type LocaleMessages = {
  [Section in keyof typeof zhCN]: Record<keyof (typeof zhCN)[Section], string>;
};
const messages: Record<Locale, LocaleMessages> = {
  "zh-CN": zhCN,
  "zh-TW": zhTW,
  "en-US": enUS,
  "ja-JP": jaJP,
  "ko-KR": koKR,
};

/** UI language is per editor; unsupported JavaScript values fall back to zh-CN. */
export function resolveLocale(locale?: string): Locale {
  return supportedLocales.includes(locale as Locale)
    ? (locale as Locale)
    : "zh-CN";
}
export function localeMessages(locale?: string): LocaleMessages {
  return messages[resolveLocale(locale)];
}
export function labels(locale?: string): Record<Label, string> {
  return localeMessages(locale).labels;
}
export function errorMessage(error: Error, locale: Locale): string {
  const messages = localeMessages(locale);
  if (error.name === "AbortError") return messages.labels.cancelled;
  const prefix = /^(Validation failed|Protected cell): (.+)$/.exec(
    error.message,
  );
  if (prefix)
    return (
      messages.errors[prefix[1] as keyof typeof messages.errors] +
      ": " +
      prefix[2]
    );
  return (
    messages.errors[error.message as keyof typeof messages.errors] ??
    error.message
  );
}
