import { useState } from "react";
import { useReservations, useCreateReservation, useExecuteReservation, useCancelReservation, useDeleteReservation } from "@/api/hooks/useReservations";
import { useMaterials } from "@/api/hooks/useMaterials";
import { PageHeader } from "@/components/ui/PageHeader";
import { EmptyState } from "@/components/ui/EmptyState";
import { Modal } from "@/components/ui/Modal";
import { ImportModal } from "@/components/ui/ImportModal";
import { formatDate } from "@/lib/format";
import { STATUS_COLORS, STATUS_LABELS } from "@/lib/constants";
import { ContactPicker } from "@/components/ui/ContactPicker";
import { useConfirm } from "@/components/ui/ConfirmDialog";
import { clsx } from "clsx";
import toast from "react-hot-toast";

export default function Reservations() {
  const [statusFilter, setStatusFilter] = useState("pending");
  const [showImport, setShowImport] = useState(false);
  const { data: reservations, isLoading } = useReservations({ status: statusFilter });
  const { data: materials } = useMaterials();
  const createMutation = useCreateReservation();
  const executeMutation = useExecuteReservation();
  const cancelMutation = useCancelReservation();
  const deleteMutation = useDeleteReservation();
  const confirm = useConfirm();

  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({ material_id: "", type: "use", quantity: "", scheduled_date: "", user_name: "", purpose: "" });

  const handleCreate = () => {
    createMutation.mutate(
      {
        material_id: parseInt(form.material_id),
        type: form.type,
        quantity: parseFloat(form.quantity),
        scheduled_date: form.scheduled_date,
        user_name: form.user_name || null,
        purpose: form.purpose || null,
      },
      {
        onSuccess: () => {
          toast.success("予約を作成しました");
          setShowAdd(false);
          setForm({ material_id: "", type: "use", quantity: "", scheduled_date: "", user_name: "", purpose: "" });
        },
        onError: (e) => toast.error(e.message),
      },
    );
  };

  const handleExecute = async (id: number, qty: number) => {
    if (!(await confirm("この予約を実行しますか？"))) return;
    executeMutation.mutate(
      { id, body: { actual_quantity: qty } },
      { onSuccess: () => toast.success("実行しました"), onError: (e) => toast.error(e.message) },
    );
  };

  return (
    <div>
      <PageHeader
        title="予約管理"
        description="使用・補充の予約"
        actions={
          <div className="flex gap-2">
            <button className="btn-secondary" onClick={() => setShowImport(true)}>Excelインポート</button>
            <button className="btn-primary" onClick={() => setShowAdd(true)}>+ 予約作成</button>
          </div>
        }
      />

      {/* Status Filter */}
      <div className="mb-4 flex gap-2">
        {["pending", "executed", "cancelled"].map((s) => (
          <button
            key={s}
            onClick={() => setStatusFilter(s)}
            className={clsx("btn-sm rounded-full", statusFilter === s ? "bg-primary-600 text-white" : "btn-secondary")}
          >
            {STATUS_LABELS[s]}
          </button>
        ))}
      </div>

      {isLoading ? (
        <div className="py-20 text-center text-gray-400">読み込み中...</div>
      ) : !reservations?.length ? (
        <EmptyState title="予約がありません" action={{ label: "予約を作成", onClick: () => setShowAdd(true) }} />
      ) : (
        <div className="space-y-2">
          {reservations.map((r) => (
            <div key={r.id} className={clsx("card flex items-center justify-between", r.is_overdue && "border-red-200 bg-red-50/50")}>
              <div className="flex items-center gap-4">
                <span className={clsx("badge", r.type === "use" ? "bg-orange-100 text-orange-700" : "bg-blue-100 text-blue-700")}>
                  {r.type === "use" ? "使用" : "補充"}
                </span>
                <div>
                  <p className="text-sm font-medium text-gray-800">{r.material_name}</p>
                  <p className="text-xs text-gray-500">
                    {r.quantity.toFixed(1)}g &middot; {formatDate(r.scheduled_date)}
                    {r.user_name && ` &middot; ${r.user_name}`}
                    {r.purpose && ` &middot; ${r.purpose}`}
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <span className={clsx("badge", STATUS_COLORS[r.status])}>{STATUS_LABELS[r.status]}</span>
                {r.is_overdue && <span className="badge bg-red-100 text-red-700">期限切れ</span>}
                {r.status === "pending" && (
                  <>
                    <button
                      className="btn-primary btn-sm"
                      onClick={() => handleExecute(r.id, r.quantity)}
                    >
                      実行
                    </button>
                    <button
                      className="btn-sm text-gray-500 hover:text-red-600"
                      onClick={() => cancelMutation.mutate(r.id, { onSuccess: () => toast.success("キャンセルしました") })}
                    >
                      取消
                    </button>
                  </>
                )}
                {r.status !== "executed" && (
                  <button
                    className="text-xs text-red-400 hover:text-red-600"
                    onClick={async () => { if (await confirm("削除しますか？")) deleteMutation.mutate(r.id); }}
                  >
                    削除
                  </button>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Add Modal */}
      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="予約作成">
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">材料</label>
            <select className="input" value={form.material_id} onChange={(e) => setForm({ ...form, material_id: e.target.value })}>
              <option value="">選択してください</option>
              {materials?.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
            </select>
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">種類</label>
              <select className="input" value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                <option value="use">使用</option>
                <option value="replenish">補充</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">数量 (g)</label>
              <input className="input" type="number" step="0.1" value={form.quantity} onChange={(e) => setForm({ ...form, quantity: e.target.value })} />
            </div>
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">予定日</label>
            <input className="input" type="date" value={form.scheduled_date} onChange={(e) => setForm({ ...form, scheduled_date: e.target.value })} />
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">担当者</label>
              <ContactPicker value={form.user_name} onChange={(name) => setForm({ ...form, user_name: name })} placeholder="担当者を選択" />
            </div>
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">目的</label>
              <input className="input" value={form.purpose} onChange={(e) => setForm({ ...form, purpose: e.target.value })} />
            </div>
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => setShowAdd(false)}>キャンセル</button>
            <button
              className="btn-primary"
              onClick={handleCreate}
              disabled={!form.material_id || !form.quantity || !form.scheduled_date || createMutation.isPending}
            >
              作成
            </button>
          </div>
        </div>
      </Modal>

      <ImportModal
        open={showImport}
        onClose={() => setShowImport(false)}
        dataType="reservations"
        label="予約"
        invalidateKeys={[["reservations"]]}
      />
    </div>
  );
}
