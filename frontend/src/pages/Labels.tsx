import { useState } from "react";
import { useLabels, useCreateLabel, useDeleteLabel } from "@/api/hooks/useLabels";
import { PageHeader } from "@/components/ui/PageHeader";
import { EmptyState } from "@/components/ui/EmptyState";
import { Modal } from "@/components/ui/Modal";
import { ImportModal } from "@/components/ui/ImportModal";
import { useConfirm } from "@/components/ui/ConfirmDialog";
import toast from "react-hot-toast";

export default function Labels() {
  const { data: labels, isLoading } = useLabels();
  const createMutation = useCreateLabel();
  const deleteMutation = useDeleteLabel();
  const confirm = useConfirm();

  const [showAdd, setShowAdd] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [form, setForm] = useState({ name: "", color: "#6c757d" });

  const handleCreate = () => {
    createMutation.mutate(form, {
      onSuccess: () => { toast.success("ラベルを作成しました"); setShowAdd(false); setForm({ name: "", color: "#6c757d" }); },
      onError: (e) => toast.error(e.message),
    });
  };

  return (
    <div>
      <PageHeader title="ラベル管理" description="材料のカテゴリ分類" actions={
        <div className="flex gap-2">
          <button className="btn-secondary" onClick={() => setShowImport(true)}>Excelインポート</button>
          <button className="btn-primary" onClick={() => setShowAdd(true)}>+ ラベル追加</button>
        </div>
      } />

      {isLoading ? (
        <div className="py-20 text-center text-gray-400">読み込み中...</div>
      ) : !labels?.length ? (
        <EmptyState title="ラベルがありません" action={{ label: "ラベルを作成", onClick: () => setShowAdd(true) }} />
      ) : (
        <div className="grid gap-3 grid-cols-3">
          {labels.map((l) => (
            <div key={l.id} className="card flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="h-8 w-8 rounded-lg" style={{ backgroundColor: l.color }} />
                <div>
                  <p className="text-sm font-medium text-gray-800">{l.name}</p>
                  <p className="text-xs text-gray-500">{l.material_count} 材料</p>
                </div>
              </div>
              <button
                onClick={async () => { if (await confirm(`「${l.name}」を削除しますか？`)) deleteMutation.mutate(l.id, { onSuccess: () => toast.success("削除しました") }); }}
                className="text-xs text-red-400 hover:text-red-600"
              >
                削除
              </button>
            </div>
          ))}
        </div>
      )}

      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="ラベル追加">
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">名前</label>
            <input className="input" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">色</label>
            <div className="flex items-center gap-3">
              <input type="color" value={form.color} onChange={(e) => setForm({ ...form, color: e.target.value })} className="h-10 w-14 cursor-pointer rounded border" />
              <input className="input w-28" value={form.color} onChange={(e) => setForm({ ...form, color: e.target.value })} />
            </div>
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => setShowAdd(false)}>キャンセル</button>
            <button className="btn-primary" onClick={handleCreate} disabled={!form.name || createMutation.isPending}>追加</button>
          </div>
        </div>
      </Modal>

      <ImportModal
        open={showImport}
        onClose={() => setShowImport(false)}
        dataType="labels"
        label="ラベル"
        invalidateKeys={[["labels"]]}
      />
    </div>
  );
}
