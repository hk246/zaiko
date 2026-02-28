import { useState } from "react";
import { Link } from "react-router-dom";
import { useWorkOrders, useCreateWorkOrder, useDeleteWorkOrder, useArchiveWorkOrder } from "@/api/hooks/useWorkOrders";
import { PageHeader } from "@/components/ui/PageHeader";
import { EmptyState } from "@/components/ui/EmptyState";
import { Modal } from "@/components/ui/Modal";
import { ImportModal } from "@/components/ui/ImportModal";
import { formatDate } from "@/lib/format";
import { STATUS_LABELS, STATUS_COLORS, PRIORITY_LABELS, PRIORITY_COLORS } from "@/lib/constants";
import { ContactPicker } from "@/components/ui/ContactPicker";
import { useConfirm } from "@/components/ui/ConfirmDialog";
import { clsx } from "clsx";
import toast from "react-hot-toast";

export default function WorkOrders() {
  const [statusFilter, setStatusFilter] = useState<string | undefined>();
  const [showArchived, setShowArchived] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const { data: workOrders, isLoading } = useWorkOrders({ status: statusFilter, archived: showArchived });
  const createMutation = useCreateWorkOrder();
  const deleteMutation = useDeleteWorkOrder();
  const archiveMutation = useArchiveWorkOrder();
  const confirm = useConfirm();

  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({
    process_name: "", priority: "none", worker_name: "", experiment_type: "standard",
    planned_start: "", planned_end: "", notes: "",
  });

  const handleCreate = () => {
    createMutation.mutate(
      {
        process_name: form.process_name,
        priority: form.priority,
        worker_name: form.worker_name || null,
        experiment_type: form.experiment_type,
        planned_start: form.planned_start || null,
        planned_end: form.planned_end || null,
        notes: form.notes || null,
      },
      {
        onSuccess: () => {
          toast.success("作業工程を作成しました");
          setShowAdd(false);
          setForm({ process_name: "", priority: "none", worker_name: "", experiment_type: "standard", planned_start: "", planned_end: "", notes: "" });
        },
        onError: (e) => toast.error(e.message),
      },
    );
  };

  return (
    <div>
      <PageHeader
        title="作業工程"
        description="作業の管理・追跡"
        actions={
          <div className="flex gap-2">
            <button className="btn-secondary" onClick={() => setShowImport(true)}>Excelインポート</button>
            <Link to="/work-orders/calendar" className="btn-secondary">カレンダー</Link>
            <button className="btn-primary" onClick={() => setShowAdd(true)}>+ 作業追加</button>
          </div>
        }
      />

      {/* Filters */}
      <div className="mb-4 flex items-center gap-3">
        <div className="flex gap-2">
          {[undefined, "pending", "in_progress", "completed"].map((s) => (
            <button
              key={s ?? "all"}
              onClick={() => setStatusFilter(s)}
              className={clsx("btn-sm rounded-full", statusFilter === s ? "bg-primary-600 text-white" : "btn-secondary")}
            >
              {s ? STATUS_LABELS[s] : "全て"}
            </button>
          ))}
        </div>
        <label className="ml-auto flex items-center gap-1.5 text-sm text-gray-500">
          <input type="checkbox" checked={showArchived} onChange={(e) => setShowArchived(e.target.checked)} className="rounded" />
          アーカイブ表示
        </label>
      </div>

      {isLoading ? (
        <div className="py-20 text-center text-gray-400">読み込み中...</div>
      ) : !workOrders?.length ? (
        <EmptyState title="作業工程がありません" action={{ label: "作業を追加", onClick: () => setShowAdd(true) }} />
      ) : (
        <div className="space-y-2">
          {workOrders.map((wo) => (
            <div key={wo.id} className={clsx("card flex items-center justify-between", wo.is_overdue && "border-red-200 bg-red-50/50")}>
              <div className="flex items-center gap-4 min-w-0">
                <span className={clsx("badge", PRIORITY_COLORS[wo.priority])}>{PRIORITY_LABELS[wo.priority]}</span>
                <div className="min-w-0">
                  <p className="text-sm font-medium text-gray-800 truncate">{wo.process_name}</p>
                  <p className="text-xs text-gray-500">
                    {wo.worker_name && `${wo.worker_name} · `}
                    {formatDate(wo.planned_start)} ～ {formatDate(wo.planned_end)}
                    {wo.experiment_type === "experiment" && " · 実験"}
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-3 shrink-0">
                {/* Progress bar */}
                {wo.status === "in_progress" && (
                  <div className="w-20">
                    <div className="h-1.5 rounded-full bg-gray-200">
                      <div className="h-1.5 rounded-full bg-blue-500" style={{ width: `${wo.progress_pct}%` }} />
                    </div>
                    <p className="mt-0.5 text-[10px] text-gray-400 text-right">{wo.progress_pct}%</p>
                  </div>
                )}
                <span className={clsx("badge", STATUS_COLORS[wo.status])}>{STATUS_LABELS[wo.status]}</span>
                {wo.is_overdue && <span className="badge bg-red-100 text-red-700">遅延</span>}
                {wo.invoice_issued && <span className="badge bg-purple-100 text-purple-700">請求済</span>}
                {!wo.archived && wo.status === "completed" && (
                  <button className="text-xs text-gray-400 hover:text-gray-600" onClick={() => archiveMutation.mutate(wo.id)}>
                    アーカイブ
                  </button>
                )}
                <button className="text-xs text-red-400 hover:text-red-600" onClick={async () => { if (await confirm("削除しますか？")) deleteMutation.mutate(wo.id); }}>
                  削除
                </button>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Add Modal */}
      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="作業工程追加" wide>
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">工程名</label>
            <input className="input" value={form.process_name} onChange={(e) => setForm({ ...form, process_name: e.target.value })} />
          </div>
          <div className="grid grid-cols-3 gap-4">
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">優先度</label>
              <select className="input" value={form.priority} onChange={(e) => setForm({ ...form, priority: e.target.value })}>
                <option value="none">なし</option>
                <option value="low">低</option>
                <option value="high">高</option>
                <option value="critical">緊急</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">担当者</label>
              <ContactPicker value={form.worker_name} onChange={(name) => setForm({ ...form, worker_name: name })} placeholder="担当者を選択" />
            </div>
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">種別</label>
              <select className="input" value={form.experiment_type} onChange={(e) => setForm({ ...form, experiment_type: e.target.value })}>
                <option value="standard">通常</option>
                <option value="experiment">実験</option>
              </select>
            </div>
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">予定開始日</label>
              <input className="input" type="date" value={form.planned_start} onChange={(e) => setForm({ ...form, planned_start: e.target.value })} />
            </div>
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">予定終了日</label>
              <input className="input" type="date" value={form.planned_end} onChange={(e) => setForm({ ...form, planned_end: e.target.value })} />
            </div>
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">メモ</label>
            <textarea className="input" rows={2} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} />
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => setShowAdd(false)}>キャンセル</button>
            <button className="btn-primary" onClick={handleCreate} disabled={!form.process_name || createMutation.isPending}>追加</button>
          </div>
        </div>
      </Modal>

      <ImportModal
        open={showImport}
        onClose={() => setShowImport(false)}
        dataType="work_orders"
        label="作業工程"
        invalidateKeys={[["work-orders"]]}
      />
    </div>
  );
}
