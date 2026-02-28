import { useState } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { api } from "@/api/client";
import { useLabels } from "@/api/hooks/useLabels";
import { PageHeader } from "@/components/ui/PageHeader";
import { EmptyState } from "@/components/ui/EmptyState";
import { Modal } from "@/components/ui/Modal";
import { ImportModal } from "@/components/ui/ImportModal";
import { useConfirm } from "@/components/ui/ConfirmDialog";
import type { PropertyField } from "@/api/types";
import toast from "react-hot-toast";

export default function PropertyFields() {
  const [labelFilter, setLabelFilter] = useState<string>("");
  const { data: labels } = useLabels();
  const qc = useQueryClient();
  const confirm = useConfirm();

  const qs = labelFilter ? `?label_id=${labelFilter}` : "";
  const { data: fields, isLoading } = useQuery({
    queryKey: ["property-fields", qs],
    queryFn: () => api.get<PropertyField[]>(`/property-fields${qs}`),
  });

  const createMutation = useMutation({
    mutationFn: (body: Record<string, unknown>) => api.post("/property-fields", body),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["property-fields"] }),
  });
  const deleteMutation = useMutation({
    mutationFn: (id: number) => api.delete(`/property-fields/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["property-fields"] }),
  });

  const [showAdd, setShowAdd] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [form, setForm] = useState({ name: "", field_type: "string", unit: "", label_id: "" });

  const handleCreate = () => {
    createMutation.mutate(
      { name: form.name, field_type: form.field_type, unit: form.unit || null, label_id: form.label_id ? parseInt(form.label_id) : null },
      { onSuccess: () => { toast.success("フィールドを追加しました"); setShowAdd(false); setForm({ name: "", field_type: "string", unit: "", label_id: "" }); } },
    );
  };

  return (
    <div>
      <PageHeader title="物性値フィールド" description="ロットに追加するカスタム属性" actions={
        <div className="flex gap-2">
          <button className="btn-secondary" onClick={() => setShowImport(true)}>Excelインポート</button>
          <button className="btn-primary" onClick={() => setShowAdd(true)}>+ フィールド追加</button>
        </div>
      } />

      <div className="mb-4">
        <select className="input w-48" value={labelFilter} onChange={(e) => setLabelFilter(e.target.value)}>
          <option value="">全ラベル</option>
          {labels?.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
        </select>
      </div>

      {isLoading ? (
        <div className="py-20 text-center text-gray-400">読み込み中...</div>
      ) : !fields?.length ? (
        <EmptyState title="フィールドがありません" action={{ label: "フィールドを追加", onClick: () => setShowAdd(true) }} />
      ) : (
        <div className="card overflow-hidden p-0">
          <table className="w-full text-left text-sm">
            <thead className="border-b border-gray-100 bg-gray-50">
              <tr>
                <th className="px-4 py-3 font-medium text-gray-500">名前</th>
                <th className="px-4 py-3 font-medium text-gray-500">型</th>
                <th className="px-4 py-3 font-medium text-gray-500">単位</th>
                <th className="px-4 py-3 font-medium text-gray-500">ラベル</th>
                <th className="px-4 py-3"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {fields.map((f) => (
                <tr key={f.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 font-medium text-gray-800">{f.name}</td>
                  <td className="px-4 py-3"><span className="badge bg-gray-100 text-gray-600">{f.field_type === "number" ? "数値" : "文字列"}</span></td>
                  <td className="px-4 py-3 text-gray-500">{f.unit || "-"}</td>
                  <td className="px-4 py-3 text-gray-500">{labels?.find((l) => l.id === f.label_id)?.name || "共通"}</td>
                  <td className="px-4 py-3 text-right">
                    <button onClick={async () => { if (await confirm("削除しますか？")) deleteMutation.mutate(f.id); }} className="text-xs text-red-400 hover:text-red-600">削除</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <Modal open={showAdd} onClose={() => setShowAdd(false)} title="フィールド追加">
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">名前</label>
            <input className="input" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">型</label>
              <select className="input" value={form.field_type} onChange={(e) => setForm({ ...form, field_type: e.target.value })}>
                <option value="string">文字列</option>
                <option value="number">数値</option>
              </select>
            </div>
            <div>
              <label className="mb-1 block text-sm font-medium text-gray-700">単位</label>
              <input className="input" placeholder="例: mPa·s" value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })} />
            </div>
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">ラベル</label>
            <select className="input" value={form.label_id} onChange={(e) => setForm({ ...form, label_id: e.target.value })}>
              <option value="">共通（全材料）</option>
              {labels?.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
            </select>
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
        dataType="property_fields"
        label="物性値フィールド"
        invalidateKeys={[["property-fields"]]}
      />
    </div>
  );
}
