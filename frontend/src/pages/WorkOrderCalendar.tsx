import { useMemo, useState, useRef, useEffect } from "react";
import { Link } from "react-router-dom";
import { useWorkOrders } from "@/api/hooks/useWorkOrders";
import { PageHeader } from "@/components/ui/PageHeader";
import { PRIORITY_COLORS, STATUS_COLORS, STATUS_LABELS } from "@/lib/constants";
import { clsx } from "clsx";
import dayjs from "dayjs";
import type { WorkOrder } from "@/api/types";

/** セル内に表示する最大件数（これ以上は「+N件」で展開） */
const MAX_VISIBLE = 3;

export default function WorkOrderCalendar() {
  const [currentMonth, setCurrentMonth] = useState(dayjs().startOf("month"));
  const { data: workOrders } = useWorkOrders({ archived: false });
  const [popover, setPopover] = useState<{ key: string; items: WorkOrder[]; x: number; y: number } | null>(null);
  const popRef = useRef<HTMLDivElement>(null);

  // Close popover on outside click
  useEffect(() => {
    if (!popover) return;
    const handler = (e: MouseEvent) => {
      if (popRef.current && !popRef.current.contains(e.target as Node)) {
        setPopover(null);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [popover]);

  const days = useMemo(() => {
    const start = currentMonth.startOf("week");
    const end = currentMonth.endOf("month").endOf("week");
    const result: dayjs.Dayjs[] = [];
    let d = start;
    while (d.isBefore(end) || d.isSame(end, "day")) {
      result.push(d);
      d = d.add(1, "day");
    }
    return result;
  }, [currentMonth]);

  const woByDay = useMemo(() => {
    const map: Record<string, WorkOrder[]> = {};
    workOrders?.forEach((wo) => {
      if (wo.planned_start) {
        const key = dayjs(wo.planned_start).format("YYYY-MM-DD");
        (map[key] ??= []).push(wo);
      }
    });
    return map;
  }, [workOrders]);

  const today = dayjs().format("YYYY-MM-DD");

  const openPopover = (e: React.MouseEvent, key: string, items: WorkOrder[]) => {
    e.stopPropagation();
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
    setPopover({ key, items, x: rect.left + rect.width / 2, y: rect.bottom + 4 });
  };

  return (
    <div>
      <PageHeader
        title="カレンダー"
        actions={<Link to="/work-orders" className="btn-secondary">&larr; 一覧に戻る</Link>}
      />

      <div className="card p-0 relative">
        {/* Month Nav */}
        <div className="flex items-center justify-between border-b border-gray-100 px-4 py-3">
          <button className="btn-secondary btn-sm" onClick={() => { setCurrentMonth(currentMonth.subtract(1, "month")); setPopover(null); }}>
            &larr;
          </button>
          <h2 className="text-sm font-bold text-gray-800">{currentMonth.format("YYYY年 M月")}</h2>
          <button className="btn-secondary btn-sm" onClick={() => { setCurrentMonth(currentMonth.add(1, "month")); setPopover(null); }}>
            &rarr;
          </button>
        </div>

        {/* Weekday Header */}
        <div className="grid grid-cols-7 border-b border-gray-100">
          {["日", "月", "火", "水", "木", "金", "土"].map((d) => (
            <div key={d} className="px-2 py-2 text-center text-xs font-medium text-gray-500">{d}</div>
          ))}
        </div>

        {/* Days Grid */}
        <div className="grid grid-cols-7">
          {days.map((d) => {
            const key = d.format("YYYY-MM-DD");
            const isToday = key === today;
            const isCurrentMonth = d.month() === currentMonth.month();
            const items = woByDay[key] ?? [];
            const overflow = items.length > MAX_VISIBLE;
            const visible = overflow ? items.slice(0, MAX_VISIBLE) : items;
            const hiddenCount = items.length - MAX_VISIBLE;

            return (
              <div
                key={key}
                className={clsx(
                  "min-h-[80px] border-b border-r border-gray-50 p-1",
                  !isCurrentMonth && "bg-gray-50/50",
                )}
              >
                <p className={clsx(
                  "mb-0.5 text-xs leading-tight",
                  isToday ? "font-bold text-primary-600" : isCurrentMonth ? "text-gray-600" : "text-gray-300",
                )}>
                  {d.date()}
                </p>
                <div className="space-y-px">
                  {visible.map((wo) => (
                    <div
                      key={wo.id}
                      className={clsx(
                        "truncate rounded px-1 py-px text-[10px] leading-tight cursor-default",
                        PRIORITY_COLORS[wo.priority],
                      )}
                      title={`${wo.process_name} (${STATUS_LABELS[wo.status]}${wo.worker_name ? ` / ${wo.worker_name}` : ""})`}
                    >
                      {wo.process_name}
                    </div>
                  ))}
                  {overflow && (
                    <button
                      className="w-full text-left rounded px-1 py-px text-[10px] leading-tight text-blue-600 hover:bg-blue-50 cursor-pointer"
                      onClick={(e) => openPopover(e, key, items)}
                    >
                      +{hiddenCount}件を表示
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>

        {/* Popover for overflow items */}
        {popover && (
          <div
            ref={popRef}
            className="fixed z-50 w-64 rounded-lg border border-gray-200 bg-white shadow-xl"
            style={{
              left: Math.min(popover.x - 128, window.innerWidth - 280),
              top: Math.min(popover.y, window.innerHeight - 300),
            }}
          >
            <div className="flex items-center justify-between border-b border-gray-100 px-3 py-2">
              <h3 className="text-xs font-bold text-gray-700">
                {dayjs(popover.key).format("M月D日")}の作業 ({popover.items.length}件)
              </h3>
              <button
                className="text-gray-400 hover:text-gray-600 text-sm leading-none"
                onClick={() => setPopover(null)}
              >
                &times;
              </button>
            </div>
            <div className="max-h-60 overflow-y-auto p-2 space-y-1">
              {popover.items.map((wo) => (
                <div
                  key={wo.id}
                  className={clsx(
                    "rounded px-2 py-1.5 text-xs",
                    PRIORITY_COLORS[wo.priority],
                  )}
                >
                  <p className="font-medium truncate">{wo.process_name}</p>
                  <p className="text-[10px] opacity-75 mt-0.5">
                    <span className={clsx("inline-block rounded px-1 mr-1", STATUS_COLORS[wo.status])}>
                      {STATUS_LABELS[wo.status]}
                    </span>
                    {wo.worker_name && <span>{wo.worker_name}</span>}
                  </p>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
