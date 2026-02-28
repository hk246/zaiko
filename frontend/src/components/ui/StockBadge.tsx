import { clsx } from "clsx";

interface StockBadgeProps {
  currentStock: number;
  minWeight: number;
  isLowStock: boolean;
  unit?: string;
}

export function StockBadge({ currentStock, minWeight, isLowStock, unit = "g" }: StockBadgeProps) {
  const display = currentStock >= 1000 && unit === "g"
    ? `${(currentStock / 1000).toFixed(1)}kg`
    : `${currentStock.toFixed(1)}${unit}`;

  return (
    <span
      className={clsx(
        "badge",
        isLowStock ? "bg-red-100 text-red-700" : "bg-green-100 text-green-700",
      )}
    >
      {display}
      {isLowStock && (
        <svg className="ml-1 h-3 w-3" fill="currentColor" viewBox="0 0 20 20">
          <path
            fillRule="evenodd"
            d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z"
            clipRule="evenodd"
          />
        </svg>
      )}
    </span>
  );
}
