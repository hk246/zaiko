import { useState, useMemo } from "react";
import { useQuery } from "@tanstack/react-query";
import { useReservations, useCreateReservation, useExecuteReservation, useCancelReservation, useDeleteReservation } from "@/api/hooks/useReservations";
import { useMaterials, useLots } from "@/api/hooks/useMaterials";
import { api } from "@/api/client";
import type { Reservation, Tare } from "@/api/types";
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

  // ── Create Form ──
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({ material_id: "", type: "use", quantity: "", scheduled_date: "", user_name: "", purpose: "" });

  // ── Execute Modal State ──
  const [execTarget, setExecTarget] = useState<Reservation | null>(null);
  const [execQty, setExecQty] = useState("");
  const [lotAllocations, setLotAllocations] = useState<Array<{ lot_id: number; quantity: string }>>([]);
  const [replenishMode, setReplenishMode] = useState<"existing" | "new">("new");
  const [replenishLotId, setReplenishLotId] = useState("");
  const [replenishLotName, setReplenishLotName] = useState("");
  const [replenishTareId, setReplenishTareId] = useState("");
  const [execBy, setExecBy] = useState("");
  const [execNote, setExecNote] = useState("");

  // ── Data for Execute Modal ──
  const { data: execLots } = useLots(execTarget?.material_id ?? 0);
  const { data: tares } = useQuery({
    queryKey: ["tares"],
    queryFn: () => api.get<Tare[]>("/tares"),
    enabled: !!execTarget && execTarget.type === "replenish",
  });

  // ── Create Handlers ──
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

  // ── Execute Modal Open / Close ──
  const openExecModal = (r: Reservation) => {
    setExecTarget(r);
    setExecQty(String(r.quantity));
    setLotAllocations([]);
    setReplenishMode(r.lot_id ? "existing" : "new");
    setReplenishLotId(r.lot_id ? String(r.lot_id) : "");
    setReplenishLotName(r.lot_name || "");
    setReplenishTareId("");
    setExecBy("");
    setExecNote("");
  };
  const closeExecModal = () => setExecTarget(null);

  // ── Lot allocation helpers (use type) ──
  const toggleLot = (lotId: number) => {
    setLotAllocations((prev) => {
      const exists = prev.find((a) => a.lot_id === lotId);
      if (exists) return prev.filter((a) => a.lot_id !== lotId);
      return [...prev, { lot_id: lotId, quantity: "" }];
    });
  };
  const setLotQty = (lotId: number, qty: string) => {
    setLotAllocations((prev) => prev.map((a) => (a.lot_id === lotId ? { ...a, quantity: qty } : a)));
  };

  // ── Validation (use type) ──
  const useValidation = useMemo(() => {
    if (!execTarget || execTarget.type !== "use") return { valid: true, errors: [] as string[] };
    const qty = parseFloat(execQty) || 0;
    const errors: string[] = [];
    if (qty <= 0) errors.push("数量は0より大きい必要があります");
    if (lotAllocations.length > 0) {
      const total = lotAllocations.reduce((s, a) => s + (parseFloat(a.quantity) || 0), 0);
      if (Math.abs(total - qty) > 0.01) {
        errors.push(`ロット合計 ${total.toFixed(1)}g ≠ 実行数量 ${qty.toFixed(1)}g`);
      }
      for (const alloc of lotAllocations) {
        const lot = execLots?.find((l) => l.id === alloc.lot_id);
        const allocQty = parseFloat(alloc.quantity) || 0;
        if (lot && allocQty > lot.weight) {
          errors.push(`${lot.lot_name}: ${allocQty.toFixed(1)}g > 在庫 ${lot.weight.toFixed(1)}g`);
        }
      }
    }
    return { valid: errors.length === 0, errors };
  }, [execTarget, execQty, lotAllocations, execLots]);

  // ── Validation (replenish type) ──
  const replenishValidation = useMemo(() => {
    if (!execTarget || execTarget.type !== "replenish") return { valid: true, errors: [] as string[] };
    const qty = parseFloat(execQty) || 0;
    const errors: string[] = [];
    if (qty <= 0) errors.push("数量は0より大きい必要があります");
    if (replenishMode === "existing" && !replenishLotId) errors.push("ロットを選択してください");
    if (replenishMode === "new" && !replenishLotName.trim()) errors.push("ロット名を入力してください");
    return { valid: errors.length === 0, errors };
  }, [execTarget, execQty, replenishMode, replenishLotId, replenishLotName]);

  const currentValidation = execTarget?.type === "use" ? useValidation : replenishValidation;

  // ── Execute Submit ──
  const handleSubmitExec = () => {
    if (!execTarget) return;
    const body: Record<string, unknown> = { actual_quantity: parseFloat(execQty) };

    if (execTarget.type === "use") {
      if (lotAllocations.length > 0) {
        body.lot_usages = lotAllocations.map((a) => ({
          lot_id: a.lot_id,
          quantity: parseFloat(a.quantity),
        }));
      }
    } else {
      // replenish
      if (replenishMode === "existing" && replenishLotId) {
        body.lot_id = parseInt(replenishLotId);
      } else if (replenishMode === "new") {
        if (replenishLotName) body.lot_name = replenishLotName;
        if (replenishTareId) body.tare_id = parseInt(replenishTareId);
      }
    }

    if (execBy) body.executed_by = execBy;
    if (execNote) body.note = execNote;

    executeMutation.mutate(
      { id: execTarget.id, body },
      {
        onSuccess: () => {
          toast.success("実行しました");
          closeExecModal();
        },
        onError: (e) => toast.error(e.message),
      },
    );
  };

  // ── Tare calculation for replenish ──
  const selectedTare = tares?.find((t) => String(t.id) === replenishTareId);
  const grossWeight = parseFloat(execQty) || 0;
  const netWeight = selectedTare ? Math.max(0, grossWeight - selectedTare.weight) : grossWeight;

  // ── Available lots (non-fraction, with stock) for use type ──
  const availableLots = useMemo(() => {
    if (!execLots) return [];
    return execLots.filter((l) => !l.is_fraction && l.weight > 0);
  }, [execLots]);

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
                    {r.user_name && ` · ${r.user_name}`}
                    {r.purpose && ` · ${r.purpose}`}
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <span className={clsx("badge", STATUS_COLORS[r.status])}>{STATUS_LABELS[r.status]}</span>
                {r.is_overdue && <span className="badge bg-red-100 text-red-700">期限切れ</span>}
                {r.status === "pending" && (
                  <>
                    <button className="btn-primary btn-sm" onClick={() => openExecModal(r)}>
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

      {/* ── Add Modal ── */}
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

      {/* ── Execute Modal ── */}
      <Modal
        open={!!execTarget}
        onClose={closeExecModal}
        title={execTarget?.type === "use" ? "予約実行 – 使用" : "予約実行 – 補充"}
        wide
      >
        {execTarget && (
          <div className="space-y-4">
            {/* Material & reservation info */}
            <div className="rounded-lg bg-gray-50 px-4 py-2.5 text-sm">
              <div className="flex items-center gap-3">
                <span className={clsx("badge text-[10px]", execTarget.type === "use" ? "bg-orange-100 text-orange-700" : "bg-blue-100 text-blue-700")}>
                  {execTarget.type === "use" ? "使用" : "補充"}
                </span>
                <span className="font-bold text-gray-800">{execTarget.material_name}</span>
                <span className="text-gray-500">予約数量: {execTarget.quantity.toFixed(1)}g</span>
              </div>
            </div>

            {/* Actual quantity */}
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">実行数量 (g)</label>
              <input
                className="input"
                type="number"
                step="0.1"
                value={execQty}
                onChange={(e) => setExecQty(e.target.value)}
              />
            </div>

            {/* ── Use type: Lot selection ── */}
            {execTarget.type === "use" && (
              <div>
                <label className="mb-2 block text-sm font-medium text-gray-700">使用ロット</label>
                {availableLots.length > 0 ? (
                  <div className="rounded-lg border border-gray-200 overflow-hidden">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="bg-gray-50 text-left text-xs text-gray-500">
                          <th className="px-3 py-2 w-8"></th>
                          <th className="px-3 py-2">ロット名</th>
                          <th className="px-3 py-2 text-right">在庫 (g)</th>
                          <th className="px-3 py-2 text-right w-36">使用量 (g)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {availableLots.map((lot) => {
                          const alloc = lotAllocations.find((a) => a.lot_id === lot.id);
                          const isChecked = !!alloc;
                          return (
                            <tr
                              key={lot.id}
                              className={clsx(
                                "border-t border-gray-100 transition-colors",
                                isChecked ? "bg-blue-50/50" : "hover:bg-gray-50",
                              )}
                            >
                              <td className="px-3 py-2">
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => toggleLot(lot.id)}
                                  className="rounded border-gray-300 text-primary-600 focus:ring-primary-500"
                                />
                              </td>
                              <td className="px-3 py-2 font-medium text-gray-800">{lot.lot_name}</td>
                              <td className="px-3 py-2 text-right text-gray-600">{lot.weight.toFixed(1)}</td>
                              <td className="px-3 py-2 text-right">
                                {isChecked ? (
                                  <input
                                    type="number"
                                    step="0.1"
                                    className="input !py-1 !px-2 !text-sm w-28 text-right"
                                    value={alloc!.quantity}
                                    onChange={(e) => setLotQty(lot.id, e.target.value)}
                                    placeholder="0.0"
                                  />
                                ) : (
                                  <span className="text-gray-300">—</span>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                    {lotAllocations.length > 0 && (
                      <div className="border-t border-gray-200 bg-gray-50 px-3 py-2 text-xs text-gray-600 flex justify-between">
                        <span>選択ロット合計</span>
                        <span className="font-bold">
                          {lotAllocations.reduce((s, a) => s + (parseFloat(a.quantity) || 0), 0).toFixed(1)}g
                          {" / "}
                          {(parseFloat(execQty) || 0).toFixed(1)}g
                        </span>
                      </div>
                    )}
                  </div>
                ) : (
                  <p className="text-sm text-gray-400">利用可能なロットがありません</p>
                )}
                {lotAllocations.length === 0 && availableLots.length > 0 && (
                  <p className="mt-2 text-xs text-blue-600 bg-blue-50 rounded-lg px-3 py-2">
                    ロット未選択の場合、古いロットから順に自動で割り当てられます (FIFO)
                  </p>
                )}
              </div>
            )}

            {/* ── Replenish type: Lot target ── */}
            {execTarget.type === "replenish" && (
              <div className="space-y-3">
                <label className="mb-1 block text-sm font-medium text-gray-700">補充先</label>
                {/* Mode toggle */}
                <div className="flex rounded-lg border border-gray-200 overflow-hidden">
                  <button
                    className={clsx("flex-1 px-3 py-2 text-xs font-medium transition-colors",
                      replenishMode === "existing" ? "bg-primary-600 text-white" : "bg-white text-gray-600 hover:bg-gray-50")}
                    onClick={() => setReplenishMode("existing")}
                  >
                    既存ロットに追加
                  </button>
                  <button
                    className={clsx("flex-1 px-3 py-2 text-xs font-medium transition-colors",
                      replenishMode === "new" ? "bg-primary-600 text-white" : "bg-white text-gray-600 hover:bg-gray-50")}
                    onClick={() => setReplenishMode("new")}
                  >
                    新規ロット作成
                  </button>
                </div>

                {replenishMode === "existing" ? (
                  <div>
                    <select
                      className="input"
                      value={replenishLotId}
                      onChange={(e) => setReplenishLotId(e.target.value)}
                    >
                      <option value="">ロットを選択</option>
                      {execLots?.filter((l) => l.weight >= 0).map((l) => (
                        <option key={l.id} value={l.id}>
                          {l.lot_name} ({l.weight.toFixed(1)}g){l.is_fraction ? " [端数]" : ""}
                        </option>
                      ))}
                    </select>
                    {replenishLotId && execLots && (() => {
                      const lot = execLots.find((l) => String(l.id) === replenishLotId);
                      if (!lot) return null;
                      return (
                        <div className="mt-2 rounded-lg bg-blue-50 px-3 py-2 text-xs text-blue-700">
                          現在: {lot.weight.toFixed(1)}g → 実行後: {(lot.weight + (parseFloat(execQty) || 0)).toFixed(1)}g
                        </div>
                      );
                    })()}
                  </div>
                ) : (
                  <div className="space-y-3">
                    <div>
                      <label className="mb-1 block text-xs font-medium text-gray-600">ロット名</label>
                      <input
                        className="input"
                        value={replenishLotName}
                        onChange={(e) => setReplenishLotName(e.target.value)}
                        placeholder="例: Lot-20260301"
                      />
                    </div>
                    <div>
                      <label className="mb-1 block text-xs font-medium text-gray-600">風袋 (任意)</label>
                      <select
                        className="input"
                        value={replenishTareId}
                        onChange={(e) => setReplenishTareId(e.target.value)}
                      >
                        <option value="">なし</option>
                        {tares?.map((t) => (
                          <option key={t.id} value={t.id}>
                            {t.name} ({t.weight.toFixed(1)}g)
                          </option>
                        ))}
                      </select>
                    </div>
                    {selectedTare && (
                      <div className="rounded-lg bg-blue-50 px-3 py-2 text-xs text-blue-700">
                        入力値 {grossWeight.toFixed(1)}g − 風袋 {selectedTare.weight.toFixed(1)}g = 正味 <strong>{netWeight.toFixed(1)}g</strong>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            {/* Common fields */}
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="mb-1 block text-sm font-medium text-gray-700">実行者</label>
                <ContactPicker value={execBy} onChange={setExecBy} placeholder="実行者を選択" />
              </div>
              <div>
                <label className="mb-1 block text-sm font-medium text-gray-700">メモ</label>
                <input className="input" value={execNote} onChange={(e) => setExecNote(e.target.value)} placeholder="メモ (任意)" />
              </div>
            </div>

            {/* Validation errors */}
            {currentValidation.errors.length > 0 && (
              <div className="rounded-lg bg-red-50 px-3 py-2 space-y-0.5">
                {currentValidation.errors.map((err, i) => (
                  <p key={i} className="text-xs text-red-600">⚠ {err}</p>
                ))}
              </div>
            )}

            {/* Actions */}
            <div className="flex justify-end gap-2 pt-2">
              <button className="btn-secondary" onClick={closeExecModal}>キャンセル</button>
              <button
                className="btn-primary"
                onClick={handleSubmitExec}
                disabled={!currentValidation.valid || executeMutation.isPending}
              >
                {executeMutation.isPending ? "実行中..." : "実行"}
              </button>
            </div>
          </div>
        )}
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
