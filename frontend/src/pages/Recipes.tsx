import { useState } from "react";
import { useRecipes, useCreateRecipe, useDeleteRecipe, useExecuteRecipe } from "@/api/hooks/useRecipes";
import { useMaterials } from "@/api/hooks/useMaterials";
import { PageHeader } from "@/components/ui/PageHeader";
import { EmptyState } from "@/components/ui/EmptyState";
import { Modal } from "@/components/ui/Modal";
import { useConfirm } from "@/components/ui/ConfirmDialog";
import toast from "react-hot-toast";

export default function Recipes() {
  const { data: recipes, isLoading } = useRecipes();
  const { data: materials } = useMaterials();
  const createMutation = useCreateRecipe();
  const deleteMutation = useDeleteRecipe();
  const executeMutation = useExecuteRecipe();
  const confirm = useConfirm();

  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({ name: "", type: "use", description: "" });
  const [items, setItems] = useState<{ material_id: string; quantity: string }[]>([{ material_id: "", quantity: "" }]);

  const [showExec, setShowExec] = useState<number | null>(null);
  const [execDate, setExecDate] = useState("");

  const addItem = () => setItems([...items, { material_id: "", quantity: "" }]);
  const removeItem = (i: number) => setItems(items.filter((_, idx) => idx !== i));
  const updateItem = (i: number, field: string, value: string) =>
    setItems(items.map((item, idx) => (idx === i ? { ...item, [field]: value } : item)));

  const handleCreate = () => {
    createMutation.mutate(
      {
        name: form.name,
        type: form.type,
        description: form.description || null,
        items: items
          .filter((it) => it.material_id && it.quantity)
          .map((it) => ({ material_id: parseInt(it.material_id), quantity: parseFloat(it.quantity) })),
      },
      {
        onSuccess: () => {
          toast.success("レシピを作成しました");
          setShowAdd(false);
          setForm({ name: "", type: "use", description: "" });
          setItems([{ material_id: "", quantity: "" }]);
        },
        onError: (e) => toast.error(e.message),
      },
    );
  };

  const handleExecute = (id: number) => {
    executeMutation.mutate(
      { id, body: { scheduled_date: execDate } },
      {
        onSuccess: () => { toast.success("レシピから予約を作成しました"); setShowExec(null); },
        onError: (e) => toast.error(e.message),
      },
    );
  };

  return (
    <div>
      <PageHeader title="レシピ" description="材料の組み合わせテンプレート" actions={<button className="btn-primary" onClick={() => setShowAdd(true)}>+ レシピ作成</button>} />

      {isLoading ? (
        <div className="py-20 text-center text-gray-400">読み込み中...</div>
      ) : !recipes?.length ? (
        <EmptyState title="レシピがありません" action={{ label: "レシピを作成", onClick: () => setShowAdd(true) }} />
      ) : (
        <div className="grid gap-4 grid-cols-3">
          {recipes.map((r) => (
            <div key={r.id} className="card">
              <div className="mb-3 flex items-start justify-between">
                <div>
                  <h3 className="text-sm font-bold text-gray-800">{r.name}</h3>
                  <span className="badge mt-1 bg-gray-100 text-gray-600">{r.type === "use" ? "使用" : "補充"}</span>
                </div>
                <button onClick={async () => { if (await confirm("削除しますか？")) deleteMutation.mutate(r.id); }} className="text-xs text-red-400 hover:text-red-600">削除</button>
              </div>
              {r.description && <p className="mb-3 text-xs text-gray-500">{r.description}</p>}
              <ul className="mb-3 space-y-1">
                {r.items.map((item) => (
                  <li key={item.id} className="flex justify-between text-xs text-gray-600">
                    <span>{item.material_name}</span>
                    <span>{item.quantity.toFixed(1)}g</span>
                  </li>
                ))}
              </ul>
              <button className="btn-primary btn-sm w-full" onClick={() => { setShowExec(r.id); setExecDate(""); }}>
                予約作成
              </button>
            </div>
          ))}
        </div>
      )}

      {/* Add Recipe Modal */}
      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="レシピ作成" wide>
        <div className="space-y-4">
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">名前</label>
              <input className="input" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            </div>
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">種類</label>
              <select className="input" value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                <option value="use">使用</option>
                <option value="replenish">補充</option>
              </select>
            </div>
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">説明</label>
            <input className="input" value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} />
          </div>
          <div>
            <div className="mb-2 flex items-center justify-between">
              <label className="text-sm font-medium text-gray-700">材料</label>
              <button className="btn-sm text-primary-600 hover:text-primary-700" onClick={addItem}>+ 追加</button>
            </div>
            {items.map((item, i) => (
              <div key={i} className="mb-2 flex items-center gap-2">
                <select className="input flex-1" value={item.material_id} onChange={(e) => updateItem(i, "material_id", e.target.value)}>
                  <option value="">材料選択</option>
                  {materials?.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
                </select>
                <input className="input w-24" type="number" step="0.1" placeholder="数量" value={item.quantity} onChange={(e) => updateItem(i, "quantity", e.target.value)} />
                {items.length > 1 && <button onClick={() => removeItem(i)} className="text-red-400 hover:text-red-600 text-sm">x</button>}
              </div>
            ))}
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => setShowAdd(false)}>キャンセル</button>
            <button className="btn-primary" onClick={handleCreate} disabled={!form.name || createMutation.isPending}>作成</button>
          </div>
        </div>
      </Modal>

      {/* Execute Modal */}
      <Modal open={showExec !== null} onClose={() => setShowExec(null)} title="レシピから予約作成">
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">予定日</label>
            <input className="input" type="date" value={execDate} onChange={(e) => setExecDate(e.target.value)} />
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => setShowExec(null)}>キャンセル</button>
            <button className="btn-primary" onClick={() => showExec && handleExecute(showExec)} disabled={!execDate || executeMutation.isPending}>作成</button>
          </div>
        </div>
      </Modal>
    </div>
  );
}
