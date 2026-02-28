import dayjs from "dayjs";

export function formatDate(d: string | null | undefined): string {
  if (!d) return "-";
  return dayjs(d).format("YYYY/MM/DD");
}

export function formatDateTime(d: string | null | undefined): string {
  if (!d) return "-";
  return dayjs(d).format("YYYY/MM/DD HH:mm");
}

export function formatWeight(w: number, unit = "g"): string {
  if (w >= 1000 && unit === "g") {
    return `${(w / 1000).toFixed(2)} kg`;
  }
  return `${w.toFixed(1)} ${unit}`;
}

export function formatNumber(n: number): string {
  return n.toLocaleString("ja-JP");
}
