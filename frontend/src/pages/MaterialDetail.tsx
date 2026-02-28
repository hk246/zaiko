import { useState } from "react";
import { useParams, Link } from "react-router-dom";
import { useMaterial, useLots, useCreateLot, useDeleteLot, useToggleFraction, useUpdateMaterial } from "@/api/hooks/useMaterials";
import { useMaterialStats, useMaterialTimeline } from "@/api/hooks/useDashboard";
import { PageHeader } from "@/components/ui/PageHeader";
import { StockBadge } from "@/components/ui/StockBadge";
import { Modal } from "@/components/ui/Modal";
import { ImportModal } from "@/components/ui/ImportModal";
import { useContacts } from "@/api/hooks/useContacts";
import { useEmailDraft } from "@/api/hooks/useEmail";
import { useConfirm } from "@/components/ui/ConfirmDialog";
import { formatDate, formatWeight } from "@/lib/format";
import { LineChart, Line, XAxis, YAxis, Tooltip, ReferenceLine, ResponsiveContainer } from "recharts";
import toast from "react-hot-toast";

export default function MaterialDetail() {
  const { id } = useParams<{ id: string }>();
  const { data: material, isLoading } = useMaterial(id!);
  const { data: lots } = useLots(id!);
  const { data: stats } = useMaterialStats(id!);
  const { data: timeline } = useMaterialTimeline(id!);
  const createLot = useCreateLot(id!);
  const deleteLot = useDeleteLot(id!);
  const toggleFraction = useToggleFraction(id!);
  const updateMaterial = useUpdateMaterial(id!);
  const { data: contacts } = useContacts();
  const emailDraft = useEmailDraft();
  const confirm = useConfirm();

  const [showAddLot, setShowAddLot] = useState(false);
  const [showImportLot, setShowImportLot] = useState(false);
  const [showImportProps, setShowImportProps] = useState(false);
  const [showEdit, setShowEdit] = useState(false);
  const [lotForm, setLotForm] = useState({ lot_name: "", weight: "" });
  const [editForm, setEditForm] = useState({ name: "", min_weight: "", action_type: "none", contact_id: "" });

  if (isLoading) return <div className="py-20 text-center text-gray-400">読み込み中...</div>;
  if (!material) return <div className="py-20 text-center text-gray-400">材料が見つかりません</div>;

  const openEdit = () => {
    setEditForm({
      name: material.name,
      min_weight: String(material.min_weight),
      action_type: material.action_type,
      contact_id: material.contact_id ? String(material.contact_id) : "",
    });
    setShowEdit(true);
  };

  const handleSaveEdit = () => {
    updateMaterial.mutate(
      { name: editForm.name, min_weight: parseFloat(editForm.min_weight) || 0, action_type: editForm.action_type, contact_id: editForm.contact_id ? parseInt(editForm.contact_id) : null },
      { onSuccess: () => { toast.success("更新しました"); setShowEdit(false); }, onError: (e) => toast.error(e.message) },
    );
  };

  const handleAddLot = () => {
    createLot.mutate(
      { lot_name: lotForm.lot_name, weight: parseFloat(lotForm.weight) || 0 },
      { onSuccess: () => { toast.success("ロットを追加しました"); setShowAddLot(false); setLotForm({ lot_name: "", weight: "" }); }, onError: (e) => toast.error(e.message) },
    );
  };

  return (
    <div>
      <div className="mb-2">
        <Link to="/materials" className="text-xs text-gray-400 hover:text-gray-600">&larr; 材料一覧</Link>
      </div>
      <PageHeader
        title={material.name}
        description={material.label?.name ?? undefined}
        actions={
          <div className="flex gap-2">
            <button onClick={() => setShowImportProps(true)} className="btn-secondary">物性値インポート</button>
            <button onClick={openEdit} className="btn-secondary">編集</button>
          </div>
        }
      />

      {/* Overview Cards */}
      <div className="mb-6 grid grid-cols-4 gap-4">
        <div className="card">
          <p className="text-xs text-gray-500">現在在庫</p>
          <p className="mt-1"><StockBadge currentStock={material.current_stock} minWeight={material.min_weight} isLowStock={material.is_low_stock} /></p>
        </div>
        <div className="card">
          <p className="text-xs text-gray-500">予測在庫</p>
          <p className="mt-1 text-lg font-bold text-gray-800">{formatWeight(material.predicted_stock)}</p>
        </div>
        <div className="card">
          <p className="text-xs text-gray-500">端数在庫</p>
          <p className="mt-1 text-lg font-bold text-gray-600">{formatWeight(material.fraction_stock)}</p>
        </div>
        <div className="card">
          <p className="text-xs text-gray-500">最低在庫量</p>
          <p className="mt-1 text-lg font-bold text-gray-600">{formatWeight(material.min_weight)}</p>
        </div>
      </div>

      {/* Low Stock Alert + Email */}
      {material.is_low_stock && (
        <div className="mb-6 rounded-lg border border-red-300 bg-red-50 px-4 py-3">
          <div className="flex items-center gap-2 mb-2">
            <svg className="h-5 w-5 text-red-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4.5c-.77-.833-2.694-.833-3.464 0L3.34 16.5c-.77.833.192 2.5 1.732 2.5z" />
            </svg>
            <span className="text-sm font-bold text-red-800">
              {material.current_stock < material.min_weight
                ? "在庫が最低量を下回っています"
                : "在庫不足の見込みがあります"}
            </span>
          </div>
          <p className="text-xs text-red-700 mb-2">
            現在在庫 <strong>{formatWeight(material.current_stock)}</strong> / 最低量 <strong>{formatWeight(material.min_weight)}</strong>
            {material.predicted_stock < material.min_weight && (
              <> &middot; 予測在庫 <strong>{formatWeight(material.predicted_stock)}</strong></>
            )}
          </p>
          {material.action_type === "email" && material.contact_id ? (
            <button
              className="flex items-center gap-1.5 rounded-lg bg-blue-600 px-3 py-2 text-xs font-medium text-white hover:bg-blue-700 transition-colors shadow-sm"
              onClick={() => emailDraft.mutate(
                { material_id: material.id, contact_id: material.contact_id! },
                { onSuccess: () => toast.success("Outlookで下書きを開きました"), onError: (e) => toast.error(e.message) }
              )}
              disabled={emailDraft.isPending}
            >
              <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
              </svg>
              {material.contact_name}にOutlookで補充依頼
            </button>
          ) : material.action_type !== "email" ? (
            <p className="text-xs text-gray-500">
              メール通知を有効にするには「編集」から アラートアクション を「メール」に設定してください。
            </p>
          ) : (
            <p className="text-xs text-gray-500">
              メール送信先を設定するには「編集」から 通知先連絡先 を選択してください。
            </p>
          )}
        </div>
      )}

      {/* Timeline Chart */}
      {timeline && timeline.length > 1 && (
        <div className="card mb-6">
          <h2 className="mb-3 text-sm font-bold text-gray-800">在庫推移予測</h2>
          <ResponsiveContainer width="100%" height={200}>
            <LineChart data={timeline}>
              <XAxis dataKey="date" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} />
              <Tooltip />
              {material.min_weight > 0 && (
                <ReferenceLine y={material.min_weight} stroke="#ef4444" strokeDasharray="4 4" label="最低量" />
              )}
              <Line type="stepAfter" dataKey="stock" stroke="#3b82f6" strokeWidth={2} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      )}

      {/* Stats */}
      {stats && stats.length > 0 && (
        <div className="card mb-6">
          <h2 className="mb-3 text-sm font-bold text-gray-800">使用統計</h2>
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs">
              <thead>
                <tr className="border-b border-gray-100">
                  <th className="px-3 py-2 font-medium text-gray-500">期間</th>
                  <th className="px-3 py-2 font-medium text-gray-500 text-right">使用量</th>
                  <th className="px-3 py-2 font-medium text-gray-500 text-right">補充量</th>
                  <th className="px-3 py-2 font-medium text-gray-500 text-right">増減</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-50">
                {stats.map((s) => (
                  <tr key={s.period_label}>
                    <td className="px-3 py-2 text-gray-700">{s.period_label}</td>
                    <td className="px-3 py-2 text-right text-red-600">-{s.total_used.toFixed(1)}</td>
                    <td className="px-3 py-2 text-right text-green-600">+{s.total_replenished.toFixed(1)}</td>
                    <td className={`px-3 py-2 text-right font-medium ${s.net_change >= 0 ? "text-green-600" : "text-red-600"}`}>
                      {s.net_change >= 0 ? "+" : ""}{s.net_change.toFixed(1)}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Lots Table */}
      <div className="card p-0">
        <div className="flex items-center justify-between border-b border-gray-100 px-4 py-3">
          <h2 className="text-sm font-bold text-gray-800">ロット一覧 ({lots?.length ?? 0})</h2>
          <div className="flex gap-2">
            <button onClick={() => setShowImportLot(true)} className="btn-secondary btn-sm">Excelインポート</button>
            <button onClick={() => setShowAddLot(true)} className="btn-primary btn-sm">+ ロット追加</button>
          </div>
        </div>
        {lots?.length ? (
          <table className="w-full text-left text-sm">
            <thead className="border-b border-gray-100 bg-gray-50">
              <tr>
                <th className="px-4 py-2 font-medium text-gray-500">ロット名</th>
                <th className="px-4 py-2 font-medium text-gray-500 text-right">重量</th>
                <th className="px-4 py-2 font-medium text-gray-500 text-center">端数</th>
                <th className="px-4 py-2 font-medium text-gray-500">作成日</th>
                <th className="px-4 py-2"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {lots.map((l) => (
                <tr key={l.id} className={`hover:bg-gray-50 ${l.is_fraction ? "opacity-60" : ""}`}>
                  <td className="px-4 py-2 font-medium text-gray-800">{l.lot_name}</td>
                  <td className="px-4 py-2 text-right">{formatWeight(l.weight)}</td>
                  <td className="px-4 py-2 text-center">
                    <button
                      onClick={() => toggleFraction.mutate(l.id)}
                      className={`badge cursor-pointer ${l.is_fraction ? "bg-yellow-100 text-yellow-700" : "bg-gray-100 text-gray-500"}`}
                    >
                      {l.is_fraction ? "端数" : "-"}
                    </button>
                  </td>
                  <td className="px-4 py-2 text-gray-500 text-xs">{formatDate(l.created_at)}</td>
                  <td className="px-4 py-2 text-right">
                    <button
                      onClick={async () => { if (await confirm(`「${l.lot_name}」を削除しますか？`)) deleteLot.mutate(l.id); }}
                      className="text-xs text-red-500 hover:text-red-700"
                    >
                      削除
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <div className="py-10 text-center text-sm text-gray-400">ロットがありません</div>
        )}
      </div>

      {/* Add Lot Modal */}
      <Modal open={showAddLot} onClose={() => setShowAddLot(false)} title="ロット追加">
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">ロット名</label>
            <input className="input" value={lotForm.lot_name} onChange={(e) => setLotForm({ ...lotForm, lot_name: e.target.value })} />
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">重量 (g)</label>
            <input className="input" type="number" step="0.1" value={lotForm.weight} onChange={(e) => setLotForm({ ...lotForm, weight: e.target.value })} />
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => setShowAddLot(false)}>キャンセル</button>
            <button className="btn-primary" onClick={handleAddLot} disabled={!lotForm.lot_name || createLot.isPending}>追加</button>
          </div>
        </div>
      </Modal>

      <ImportModal
        open={showImportLot}
        onClose={() => setShowImportLot(false)}
        dataType="lots"
        label="ロット"
        materialId={parseInt(id!)}
        invalidateKeys={[["lots", id!], ["material", id!], ["materials"]]}
      />
      <ImportModal
        open={showImportProps}
        onClose={() => setShowImportProps(false)}
        dataType="lot_properties"
        label="物性値"
        materialId={parseInt(id!)}
        invalidateKeys={[["lots", id!], ["material", id!]]}
      />

      {/* Edit Material Modal */}
      <Modal open={showEdit} onClose={() => setShowEdit(false)} title="材料編集">
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">名前</label>
            <input className="input" value={editForm.name} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} />
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">最低在庫量 (g)</label>
            <input className="input" type="number" value={editForm.min_weight} onChange={(e) => setEditForm({ ...editForm, min_weight: e.target.value })} />
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">アラートアクション</label>
            <select className="input" value={editForm.action_type} onChange={(e) => setEditForm({ ...editForm, action_type: e.target.value })}>
              <option value="none">なし</option>
              <option value="email">メール</option>
              <option value="excel">Excel</option>
            </select>
          </div>
          {editForm.action_type === "email" && (
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">通知先連絡先</label>
              <select
                className="input"
                value={editForm.contact_id}
                onChange={(e) => setEditForm({ ...editForm, contact_id: e.target.value })}
              >
                <option value="">選択してください</option>
                {contacts?.map((c) => (
                  <option key={c.id} value={c.id}>{c.name} ({c.email})</option>
                ))}
              </select>
            </div>
          )}
          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => setShowEdit(false)}>キャンセル</button>
            <button className="btn-primary" onClick={handleSaveEdit} disabled={updateMaterial.isPending}>保存</button>
          </div>
        </div>
      </Modal>
    </div>
  );
}
