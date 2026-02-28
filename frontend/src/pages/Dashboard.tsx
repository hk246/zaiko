import { Link } from "react-router-dom";
import { useDashboardSummary, useAlerts } from "@/api/hooks/useDashboard";
import { useEmailDraft } from "@/api/hooks/useEmail";
import { PageHeader } from "@/components/ui/PageHeader";
import { clsx } from "clsx";
import { formatDate } from "@/lib/format";
import toast from "react-hot-toast";

export default function Dashboard() {
  const { data: summary, isLoading } = useDashboardSummary();
  const { data: alerts } = useAlerts();
  const emailDraft = useEmailDraft();

  if (isLoading) return <div className="py-20 text-center text-gray-400">読み込み中...</div>;

  const cards = [
    { label: "材料数", value: summary?.total_materials ?? 0, to: "/materials", color: "text-primary-600" },
    { label: "ロット数", value: summary?.total_lots ?? 0, to: "/materials", color: "text-gray-700" },
    { label: "低在庫", value: summary?.low_stock_count ?? 0, to: "/materials?low_stock=1", color: "text-red-600", alert: true },
    { label: "未実行予約", value: summary?.pending_reservations ?? 0, to: "/reservations", color: "text-yellow-600" },
    { label: "期限切れ予約", value: summary?.overdue_reservations ?? 0, to: "/reservations", color: "text-red-600", alert: true },
    { label: "進行中作業", value: summary?.active_work_orders ?? 0, to: "/work-orders", color: "text-blue-600" },
  ];

  return (
    <div>
      <PageHeader title="ダッシュボード" description="在庫状況の概要" />

      {/* Summary Cards */}
      <div className="mb-8 grid grid-cols-6 gap-4">
        {cards.map((c) => (
          <Link key={c.label} to={c.to} className="card hover:shadow-md transition-shadow">
            <p className="text-xs font-medium text-gray-500">{c.label}</p>
            <p className={clsx("mt-1 text-2xl font-bold", c.alert && c.value > 0 ? "text-red-600" : c.color)}>
              {c.value}
            </p>
          </Link>
        ))}
      </div>

      {/* Low Stock Alerts */}
      {alerts && alerts.length > 0 && (
        <div className="card mb-6 border-red-200">
          <div className="mb-3 flex items-center gap-2">
            <svg className="h-5 w-5 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4.5c-.77-.833-2.694-.833-3.464 0L3.34 16.5c-.77.833.192 2.5 1.732 2.5z" />
            </svg>
            <h2 className="text-sm font-bold text-red-800">在庫アラート ({alerts.length}件)</h2>
          </div>
          <div className="space-y-3">
            {alerts.map((a, i) => {
              const isNow = new Date(a.start_date) <= new Date();
              return (
                <div key={i} className={clsx(
                  "rounded-lg border px-4 py-3",
                  isNow ? "border-red-300 bg-red-50" : "border-yellow-300 bg-yellow-50",
                )}>
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0 flex-1">
                      <div className="flex items-center gap-2 mb-1">
                        <span className={clsx(
                          "badge text-[10px]",
                          isNow ? "bg-red-600 text-white" : "bg-yellow-500 text-white",
                        )}>
                          {isNow ? "不足中" : "不足予測"}
                        </span>
                        <Link to={`/materials/${a.material_id}`} className={clsx(
                          "text-sm font-bold hover:underline",
                          isNow ? "text-red-800" : "text-yellow-800",
                        )}>
                          {a.material_name}
                        </Link>
                      </div>
                      <div className={clsx("flex flex-wrap gap-x-4 gap-y-1 text-xs", isNow ? "text-red-700" : "text-yellow-700")}>
                        <span>現在在庫: <strong>{a.current_stock.toFixed(1)}g</strong></span>
                        <span>不足量: <strong>{Math.abs(a.shortage_amount).toFixed(1)}g</strong></span>
                        <span>
                          期間: {formatDate(a.start_date)} ～ {a.end_date ? formatDate(a.end_date) : "継続中"}
                        </span>
                      </div>
                      {a.contact_name && (
                        <p className="mt-1 text-xs text-gray-500">
                          担当: {a.contact_name} ({a.contact_email})
                        </p>
                      )}
                    </div>
                    {a.action_type === "email" && a.contact_id && (
                      <button
                        className="shrink-0 flex items-center gap-1.5 rounded-lg bg-blue-600 px-3 py-2 text-xs font-medium text-white hover:bg-blue-700 transition-colors shadow-sm"
                        onClick={() => emailDraft.mutate(
                          { material_id: a.material_id, contact_id: a.contact_id! },
                          { onSuccess: () => toast.success("Outlookで下書きを開きました"), onError: (e) => toast.error(e.message) }
                        )}
                        disabled={emailDraft.isPending}
                      >
                        <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                          <path strokeLinecap="round" strokeLinejoin="round" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
                        </svg>
                        {a.contact_name}に連絡
                      </button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Quick Links */}
      <div className="grid grid-cols-3 gap-4">
        <Link to="/materials" className="card flex items-center gap-3 hover:shadow-md transition-shadow">
          <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-primary-100 text-primary-600">
            <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4" />
            </svg>
          </div>
          <div>
            <p className="text-sm font-medium text-gray-800">材料管理</p>
            <p className="text-xs text-gray-500">在庫の追加・編集</p>
          </div>
        </Link>
        <Link to="/work-orders" className="card flex items-center gap-3 hover:shadow-md transition-shadow">
          <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-blue-100 text-blue-600">
            <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" />
            </svg>
          </div>
          <div>
            <p className="text-sm font-medium text-gray-800">作業工程</p>
            <p className="text-xs text-gray-500">作業の管理・追跡</p>
          </div>
        </Link>
        <Link to="/backup" className="card flex items-center gap-3 hover:shadow-md transition-shadow">
          <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-green-100 text-green-600">
            <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
            </svg>
          </div>
          <div>
            <p className="text-sm font-medium text-gray-800">バックアップ</p>
            <p className="text-xs text-gray-500">データの保護</p>
          </div>
        </Link>
      </div>
    </div>
  );
}
