export function toNumber(value: string | number | null | undefined): number {
  const numeric = typeof value === "number" ? value : Number(value ?? 0);
  return Number.isFinite(numeric) ? numeric : 0;
}

export function formatMoney(
  value: string | number | null | undefined,
  currency = "AUD",
): string {
  return new Intl.NumberFormat("en-AU", {
    style: "currency",
    currency,
    maximumFractionDigits: 0,
  }).format(toNumber(value));
}

export function formatNumber(value: string | number | null | undefined, decimals = 2): string {
  return new Intl.NumberFormat("en-AU", {
    maximumFractionDigits: decimals,
    minimumFractionDigits: decimals,
  }).format(toNumber(value));
}

export function humanize(value: string): string {
  return value
    .replaceAll("_", " ")
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}
