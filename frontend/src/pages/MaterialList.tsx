import { useState } from "react";
import { Link, useSearchParams } from "react-router-dom";
import { useMaterials, useCreateMaterial, useDeleteMaterial } from "@/api/hooks/useMaterials";
import { useLabels } from "@/api/hooks/useLabels";
import { PageHeader } from "@/components/ui/PageHeader";
import { StockBadge } from "@/components/ui/StockBadge";
import { EmptyState } from "@/components/ui/EmptyState";
import { Modal } from "@/components/ui/Modal";
import { ImportModal } from "@/components/ui/ImportModal";
import { useConfirm } from "@/components/ui/ConfirmDialog";
import toast from "react-hot-toast";

export default function MaterialList() {
  const [searchParams] = useSearchParams();
  const [search, setSearch] = useState("");
  const [labelFilter, setLabelFilter] = useState<number | undefined>();
  const lowStockOnly = searchParams.get("low_stock") === "1";

  const { data: materials, isLoading } = useMaterials({ q: search, label_id: labelFilter, low_stock_only: lowStockOnly });
  const { data: labels } = useLabels();
  const createMutation = useCreateMaterial();
  const deleteMutation = useDeleteMaterial();
  const confirm = useConfirm();

  const [showAdd, setShowAdd] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [form, setForm] = useState({ name: "", unit: "g", min_weight: "0", label_id: "" });

  const handleCreate = () => {
    createMutation.mutate(
      {
        name: form.name,
        unit: form.unit,
        min_weight: parseFloat(form.min_weight) || 0,
        label_id: form.label_id ? parseInt(form.label_id) : null,
      },
      {
        onSuccess: () => {
          toast.success("材料を追加しました");
          setShowAdd(false);
          setForm({ name: "", unit: "g", min_weight: "0", label_id: "" });
        },
        onError: (e) => toast.error(e.message),
      },
    );
  };

  const handleDelete = async (id: number, name: string) => {
    if (!(await confirm(`「${name}」を削除しますか？`))) return;
    deleteMutation.mutate(id, {
      onSuccess: () => toast.success("削除しました"),
      onError: (e) => toast.error(e.message),
    });
  };

  return (
    <div>
      <PageHeader
        title="材料一覧"
        description="原料の在庫管理"
        actions={
          <div className="flex gap-2">
            <button onClick={() => setShowImport(true)} className="btn-secondary">Excelインポート</button>
            <button onClick={() => setShowAdd(true)} className="btn-primary">+ 材料追加</button>
          </div>
        }
      />

      {/* Filters */}
      <div className="mb-4 flex items-center gap-3">
        <input
          type="text"
          placeholder="検索..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          className="input w-64"
        />
        <select
          value={labelFilter ?? ""}
          onChange={(e) => setLabelFilter(e.target.value ? parseInt(e.target.value) : undefined)}
          className="input w-40"
        >
          <option value="">全ラベル</option>
          {labels?.map((l) => (
            <option key={l.id} value={l.id}>{l.name}</option>
          ))}
        </select>
      </div>

      {/* Table */}
      {isLoading ? (
        <div className="py-20 text-center text-gray-400">読み込み中...</div>
      ) : !materials?.length ? (
        <EmptyState title="材料がありません" action={{ label: "材料を追加", onClick: () => setShowAdd(true) }} />
      ) : (
        <div className="card overflow-hidden p-0">
          <table className="w-full text-left text-sm">
            <thead className="border-b border-gray-100 bg-gray-50">
              <tr>
                <th className="px-4 py-3 font-medium text-gray-500">名前</th>
                <th className="px-4 py-3 font-medium text-gray-500">ラベル</th>
                <th className="px-4 py-3 font-medium text-gray-500 text-right">現在在庫</th>
                <th className="px-4 py-3 font-medium text-gray-500 text-right">予測在庫</th>
                <th className="px-4 py-3 font-medium text-gray-500 text-right">最低量</th>
                <th className="px-4 py-3 font-medium text-gray-500 text-center">ロット</th>
                <th className="px-4 py-3 font-medium text-gray-500"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {materials.map((m) => (
                <tr key={m.id} className="hover:bg-gray-50 transition-colors">
                  <td className="px-4 py-3">
                    <Link to={`/materials/${m.id}`} className="font-medium text-gray-800 hover:text-primary-600">
                      {m.name}
                    </Link>
                  </td>
                  <td className="px-4 py-3">
                    {m.label && (
                      <span className="badge" style={{ backgroundColor: m.label.color + "20", color: m.label.color }}>
                        {m.label.name}
                      </span>
                    )}
                  </td>
                  <td className="px-4 py-3 text-right">
                    <StockBadge currentStock={m.current_stock} minWeight={m.min_weight} isLowStock={m.is_low_stock} />
                  </td>
                  <td className="px-4 py-3 text-right text-gray-600">{m.predicted_stock.toFixed(1)}g</td>
                  <td className="px-4 py-3 text-right text-gray-600">{m.min_weight.toFixed(1)}g</td>
                  <td className="px-4 py-3 text-center text-gray-500">{m.lot_count}</td>
                  <td className="px-4 py-3 text-right">
                    <button
                      onClick={() => handleDelete(m.id, m.name)}
                      className="text-xs text-red-500 hover:text-red-700"
                    >
                      削除
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Add Modal */}
      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="材料追加">
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">名前</label>
            <input className="input" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">単位</label>
              <input className="input" value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })} />
            </div>
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">最低在庫量</label>
              <input className="input" type="number" value={form.min_weight} onChange={(e) => setForm({ ...form, min_weight: e.target.value })} />
            </div>
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">ラベル</label>
            <select className="input" value={form.label_id} onChange={(e) => setForm({ ...form, label_id: e.target.value })}>
              <option value="">なし</option>
              {labels?.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
            </select>
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => setShowAdd(false)}>キャンセル</button>
            <button className="btn-primary" onClick={handleCreate} disabled={!form.name || createMutation.isPending}>
              追加
            </button>
          </div>
        </div>
      </Modal>

      <ImportModal
        open={showImport}
        onClose={() => setShowImport(false)}
        dataType="materials"
        label="材料"
        invalidateKeys={[["materials"]]}
      />
    </div>
  );
}
