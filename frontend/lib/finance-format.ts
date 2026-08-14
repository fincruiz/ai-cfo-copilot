export function toNumber(value: string | number | null | undefined): number {
  const numeric = typeof value === "number" ? value : Number(value ?? 0);
  return Number.isFinite(numeric) ? numeric : 0;
}

export function formatMoney(
  value: string | number | null | undefined,
  currency = "AUD",
  decimals = 0,
): string {
  return new Intl.NumberFormat("en-AU", {
    style: "currency",
    currency,
    maximumFractionDigits: Math.min(2, Math.max(0, decimals)),
    minimumFractionDigits: Math.min(2, Math.max(0, decimals)),
  }).format(toNumber(value));
}

export function formatCompactMoney(
  value: string | number | null | undefined,
  currency = "AUD",
): string {
  return new Intl.NumberFormat("en-AU", {
    style: "currency",
    currency,
    notation: "compact",
    maximumFractionDigits: 2,
  }).format(toNumber(value));
}

export function formatNumber(value: string | number | null | undefined, decimals = 2): string {
  return new Intl.NumberFormat("en-AU", {
    maximumFractionDigits: Math.min(2, Math.max(0, decimals)),
    minimumFractionDigits: Math.min(2, Math.max(0, decimals)),
  }).format(toNumber(value));
}

export function formatPercent(value: string | number | null | undefined, decimals = 2): string {
  return `${formatNumber(value, decimals)}%`;
}

export function formatDays(value: string | number | null | undefined): string {
  return `${formatNumber(value, 2)} days`;
}

export function formatRatio(value: string | number | null | undefined): string {
  return `${formatNumber(value, 2)}x`;
}

export function humanize(value: string): string {
  return value
    .replaceAll("_", " ")
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}
