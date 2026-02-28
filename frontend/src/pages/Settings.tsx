import { useState } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { api } from "@/api/client";
import { API_BASE } from "@/lib/constants";
import { PageHeader } from "@/components/ui/PageHeader";
import { Modal } from "@/components/ui/Modal";
import { ImportModal } from "@/components/ui/ImportModal";
import { useMyName } from "@/api/hooks/useMyName";
import type { Tare, Contact } from "@/api/types";
import toast from "react-hot-toast";

export default function Settings() {
  const qc = useQueryClient();

  const { data: settings } = useQuery({
    queryKey: ["settings"],
    queryFn: () => api.get<{ database_folder: string; backup_folder: string; app_version: string }>("/settings"),
  });

  // Tares
  const { data: tares } = useQuery({ queryKey: ["tares"], queryFn: () => api.get<Tare[]>("/tares") });
  const createTare = useMutation({
    mutationFn: (body: Record<string, unknown>) => api.post("/tares", body),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["tares"] }),
  });
  const deleteTare = useMutation({
    mutationFn: (id: number) => api.delete(`/tares/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["tares"] }),
  });

  // Contacts
  const { data: contacts } = useQuery({ queryKey: ["contacts"], queryFn: () => api.get<Contact[]>("/contacts") });
  const createContact = useMutation({
    mutationFn: (body: Record<string, unknown>) => api.post("/contacts", body),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["contacts"] }),
  });
  const deleteContact = useMutation({
    mutationFn: (id: number) => api.delete(`/contacts/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["contacts"] }),
  });

  const [showTare, setShowTare] = useState(false);
  const [tareForm, setTareForm] = useState({ name: "", weight: "" });
  const [showContact, setShowContact] = useState(false);
  const [contactForm, setContactForm] = useState({ name: "", email: "" });
  const [showImportTare, setShowImportTare] = useState(false);
  const [showImportContact, setShowImportContact] = useState(false);
  const { myContactId, setMyContactId } = useMyName();
  const [exporting, setExporting] = useState(false);
  const [dbFolder, setDbFolder] = useState("");
  const [backupFolder, setBackupFolder] = useState("");
  const [dbFolderEditing, setDbFolderEditing] = useState(false);
  const [backupFolderEditing, setBackupFolderEditing] = useState(false);

  // Reimport state
  const [reimportFile, setReimportFile] = useState<File | null>(null);
  const [reimportStep, setReimportStep] = useState<"select" | "preview" | "done">("select");
  const [reimportPreview, setReimportPreview] = useState<any>(null);
  const [reimportLoading, setReimportLoading] = useState(false);
  const [reimportResult, setReimportResult] = useState<any>(null);

  const saveDbFolder = useMutation({
    mutationFn: (folder: string) => api.put("/settings/database-folder", { folder }),
    onSuccess: () => {
      toast.success("データベースフォルダを変更しました（再起動後に反映）");
      setDbFolderEditing(false);
      qc.invalidateQueries({ queryKey: ["settings"] });
    },
    onError: (e: any) => toast.error(e.message),
  });

  const saveBackupFolder = useMutation({
    mutationFn: (folder: string) => api.put("/settings/backup-folder", { folder }),
    onSuccess: () => {
      toast.success("バックアップフォルダを変更しました（再起動後に反映）");
      setBackupFolderEditing(false);
      qc.invalidateQueries({ queryKey: ["settings"] });
    },
    onError: (e: any) => toast.error(e.message),
  });

  const handleReimportPreview = async () => {
    if (!reimportFile) return;
    setReimportLoading(true);
    try {
      const form = new FormData();
      form.append("file", reimportFile);
      const res = await fetch(`${API_BASE}/export/import?confirm=false`, { method: "POST", body: form });
      if (!res.ok) { const b = await res.json().catch(() => ({})); throw new Error(b.detail || "プレビューに失敗しました"); }
      const data = await res.json();
      setReimportPreview(data);
      setReimportStep("preview");
    } catch (e: any) {
      toast.error(e.message);
    } finally {
      setReimportLoading(false);
    }
  };

  const handleReimportConfirm = async () => {
    if (!reimportFile) return;
    setReimportLoading(true);
    try {
      const form = new FormData();
      form.append("file", reimportFile);
      const res = await fetch(`${API_BASE}/export/import?confirm=true`, { method: "POST", body: form });
      if (!res.ok) { const b = await res.json().catch(() => ({})); throw new Error(b.detail || "インポートに失敗しました"); }
      const data = await res.json();
      setReimportResult(data);
      setReimportStep("done");
      qc.invalidateQueries();
      toast.success("インポートが完了しました");
    } catch (e: any) {
      toast.error(e.message);
    } finally {
      setReimportLoading(false);
    }
  };

  const resetReimport = () => {
    setReimportFile(null);
    setReimportStep("select");
    setReimportPreview(null);
    setReimportResult(null);
  };

  const handleExport = async () => {
    setExporting(true);
    try {
      const res = await fetch(`${API_BASE}/export`);
      if (!res.ok) throw new Error("エクスポートに失敗しました");
      const blob = await res.blob();
      const cd = res.headers.get("content-disposition") || "";
      // Prefer RFC 5987 filename*=UTF-8''... format, fallback to filename=...
      const match = cd.match(/filename\*=UTF-8''([^;\s]+)/i) || cd.match(/filename=([^;\s]+)/i);
      const filename = match ? decodeURIComponent(match[1]) : "在庫データ.xlsx";
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
      toast.success("エクスポートが完了しました");
    } catch (e: any) {
      toast.error(e.message || "エクスポートに失敗しました");
    } finally {
      setExporting(false);
    }
  };

  return (
    <div>
      <PageHeader title="設定" />

      {/* App Info */}
      <div className="card mb-6">
        <h2 className="mb-3 text-sm font-bold text-gray-800">アプリケーション情報</h2>
        <div className="space-y-3">
          {/* Database folder */}
          <div>
            <label className="block text-xs font-medium text-gray-500 mb-1">データベースフォルダ</label>
            {dbFolderEditing ? (
              <div className="flex gap-2">
                <input className="input flex-1 text-sm" value={dbFolder} onChange={(e) => setDbFolder(e.target.value)} placeholder="フォルダパスを入力" />
                <button className="btn-primary btn-sm" onClick={() => saveDbFolder.mutate(dbFolder)} disabled={!dbFolder || saveDbFolder.isPending}>保存</button>
                <button className="btn-secondary btn-sm" onClick={() => setDbFolderEditing(false)}>取消</button>
              </div>
            ) : (
              <div className="flex items-center gap-2">
                <span className="flex-1 text-sm font-medium text-gray-800 truncate">{settings?.database_folder ?? "-"}</span>
                <button className="text-xs text-primary-600 hover:text-primary-800 font-medium" onClick={() => { setDbFolder(settings?.database_folder ?? ""); setDbFolderEditing(true); }}>変更</button>
              </div>
            )}
          </div>
          {/* Backup folder */}
          <div>
            <label className="block text-xs font-medium text-gray-500 mb-1">バックアップフォルダ</label>
            {backupFolderEditing ? (
              <div className="flex gap-2">
                <input className="input flex-1 text-sm" value={backupFolder} onChange={(e) => setBackupFolder(e.target.value)} placeholder="フォルダパスを入力" />
                <button className="btn-primary btn-sm" onClick={() => saveBackupFolder.mutate(backupFolder)} disabled={!backupFolder || saveBackupFolder.isPending}>保存</button>
                <button className="btn-secondary btn-sm" onClick={() => setBackupFolderEditing(false)}>取消</button>
              </div>
            ) : (
              <div className="flex items-center gap-2">
                <span className="flex-1 text-sm font-medium text-gray-800 truncate">{settings?.backup_folder ?? "-"}</span>
                <button className="text-xs text-primary-600 hover:text-primary-800 font-medium" onClick={() => { setBackupFolder(settings?.backup_folder ?? ""); setBackupFolderEditing(true); }}>変更</button>
              </div>
            )}
          </div>
          <p className="text-[11px] text-gray-400">※ フォルダ変更後はアプリケーションの再起動が必要です</p>
        </div>
      </div>

      {/* Data Export & Import */}
      <div className="card mb-6">
        <h2 className="mb-3 text-sm font-bold text-gray-800">データエクスポート / インポート</h2>

        {/* Export */}
        <div className="mb-4 flex items-center justify-between">
          <p className="text-xs text-gray-500">全データをExcelファイル（シート別）で一括ダウンロード</p>
          <button
            className="flex items-center gap-2 rounded-lg bg-green-600 px-4 py-2 text-sm font-medium text-white hover:bg-green-700 transition-colors shadow-sm disabled:opacity-50"
            onClick={handleExport}
            disabled={exporting}
          >
            <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
            </svg>
            {exporting ? "エクスポート中..." : "Excelエクスポート"}
          </button>
        </div>

        {/* Reimport */}
        <div className="border-t border-gray-100 pt-4">
          <p className="text-xs text-gray-500 mb-3">
            エクスポートしたExcelを編集して再インポート（IDあり→更新、IDなし→新規作成）
          </p>

          {reimportStep === "select" && (
            <div className="flex items-center gap-3">
              <input
                type="file"
                accept=".xlsx,.xls"
                className="text-sm text-gray-600 file:mr-3 file:rounded-lg file:border-0 file:bg-blue-50 file:px-3 file:py-1.5 file:text-xs file:font-medium file:text-blue-700 hover:file:bg-blue-100"
                onChange={(e) => setReimportFile(e.target.files?.[0] ?? null)}
              />
              <button
                className="btn-primary btn-sm"
                onClick={handleReimportPreview}
                disabled={!reimportFile || reimportLoading}
              >
                {reimportLoading ? "解析中..." : "プレビュー"}
              </button>
            </div>
          )}

          {reimportStep === "preview" && reimportPreview && (
            <div>
              <div className="mb-3 grid grid-cols-4 gap-3">
                {reimportPreview.sheets?.map((s: any) => (
                  <div key={s.name} className="rounded-lg border border-gray-200 px-3 py-2">
                    <p className="text-xs font-bold text-gray-700">{s.name}</p>
                    <div className="mt-1 flex gap-3 text-[11px]">
                      {s.updates > 0 && <span className="text-blue-600">更新: {s.updates}</span>}
                      {s.creates > 0 && <span className="text-green-600">新規: {s.creates}</span>}
                      {s.errors?.length > 0 && <span className="text-red-600">エラー: {s.errors.length}</span>}
                      {s.updates === 0 && s.creates === 0 && s.errors?.length === 0 && <span className="text-gray-400">変更なし</span>}
                    </div>
                  </div>
                ))}
              </div>

              {/* Errors */}
              {reimportPreview.sheets?.some((s: any) => s.errors?.length > 0) && (
                <div className="mb-3 max-h-40 overflow-y-auto rounded-lg border border-red-200 bg-red-50 p-3">
                  <p className="mb-1 text-xs font-bold text-red-700">エラー一覧</p>
                  {reimportPreview.sheets.filter((s: any) => s.errors?.length > 0).map((s: any) =>
                    s.errors.map((e: any, i: number) => (
                      <p key={`${s.name}-${i}`} className="text-[11px] text-red-600">
                        [{s.name}] 行{e.row}: {e.field} - {e.message}
                      </p>
                    ))
                  )}
                </div>
              )}

              <div className="flex gap-2">
                <button className="btn-secondary btn-sm" onClick={resetReimport}>戻る</button>
                <button
                  className="btn-primary btn-sm"
                  onClick={handleReimportConfirm}
                  disabled={!reimportPreview.valid || reimportLoading}
                >
                  {reimportLoading ? "インポート中..." : "インポート実行"}
                </button>
              </div>
            </div>
          )}

          {reimportStep === "done" && reimportResult && (
            <div>
              <div className="mb-3 rounded-lg border border-green-200 bg-green-50 px-4 py-3">
                <p className="text-sm font-medium text-green-800">✓ インポート完了</p>
                <div className="mt-1 flex gap-4 text-xs text-green-700">
                  <span>更新: {reimportResult.total_updates}件</span>
                  <span>新規: {reimportResult.total_creates}件</span>
                </div>
              </div>
              <button className="btn-secondary btn-sm" onClick={resetReimport}>閉じる</button>
            </div>
          )}
        </div>
      </div>

      {/* My Name */}
      <div className="card mb-6">
        <h2 className="mb-3 text-sm font-bold text-gray-800">自分の名前</h2>
        <p className="mb-2 text-xs text-gray-500">担当者の入力時に「自分」として選択できる名前を設定します。</p>
        <select
          className="input max-w-xs"
          value={myContactId ?? ""}
          onChange={(e) => {
            const val = e.target.value;
            setMyContactId(val ? parseInt(val, 10) : null);
            toast.success(val ? "自分の名前を設定しました" : "自分の名前をクリアしました");
          }}
        >
          <option value="">未設定</option>
          {contacts?.map((c) => (
            <option key={c.id} value={c.id}>{c.name} ({c.email})</option>
          ))}
        </select>
      </div>

      {/* Tares */}
      <div className="card mb-6">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-bold text-gray-800">風袋管理</h2>
          <div className="flex gap-2">
            <button className="btn-secondary btn-sm" onClick={() => setShowImportTare(true)}>Excelインポート</button>
            <button className="btn-primary btn-sm" onClick={() => setShowTare(true)}>+ 追加</button>
          </div>
        </div>
        {tares?.length ? (
          <div className="space-y-2">
            {tares.map((t) => (
              <div key={t.id} className="flex items-center justify-between rounded-lg bg-gray-50 px-3 py-2">
                <div>
                  <span className="text-sm font-medium text-gray-800">{t.name}</span>
                  <span className="ml-2 text-xs text-gray-500">{t.weight}g</span>
                </div>
                <button onClick={() => deleteTare.mutate(t.id)} className="text-xs text-red-400 hover:text-red-600">削除</button>
              </div>
            ))}
          </div>
        ) : (
          <p className="text-sm text-gray-400">風袋が登録されていません</p>
        )}
      </div>

      {/* Contacts */}
      <div className="card mb-6">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-bold text-gray-800">連絡先</h2>
          <div className="flex gap-2">
            <button className="btn-secondary btn-sm" onClick={() => setShowImportContact(true)}>Excelインポート</button>
            <button className="btn-primary btn-sm" onClick={() => setShowContact(true)}>+ 追加</button>
          </div>
        </div>
        {contacts?.length ? (
          <div className="space-y-2">
            {contacts.map((c) => (
              <div key={c.id} className="flex items-center justify-between rounded-lg bg-gray-50 px-3 py-2">
                <div>
                  <span className="text-sm font-medium text-gray-800">{c.name}</span>
                  <span className="ml-2 text-xs text-gray-500">{c.email}</span>
                </div>
                <button onClick={() => deleteContact.mutate(c.id)} className="text-xs text-red-400 hover:text-red-600">削除</button>
              </div>
            ))}
          </div>
        ) : (
          <p className="text-sm text-gray-400">連絡先が登録されていません</p>
        )}
      </div>

      {/* Tare Modal */}
      <Modal open={showTare} onClose={() => setShowTare(false)} title="風袋追加">
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">名前</label>
            <input className="input" value={tareForm.name} onChange={(e) => setTareForm({ ...tareForm, name: e.target.value })} />
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">重量 (g)</label>
            <input className="input" type="number" step="0.01" value={tareForm.weight} onChange={(e) => setTareForm({ ...tareForm, weight: e.target.value })} />
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => setShowTare(false)}>キャンセル</button>
            <button className="btn-primary" onClick={() => { createTare.mutate({ name: tareForm.name, weight: parseFloat(tareForm.weight) }, { onSuccess: () => { toast.success("追加しました"); setShowTare(false); setTareForm({ name: "", weight: "" }); } }); }} disabled={!tareForm.name}>追加</button>
          </div>
        </div>
      </Modal>

      {/* Contact Modal */}
      <Modal open={showContact} onClose={() => setShowContact(false)} title="連絡先追加">
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">名前</label>
            <input className="input" value={contactForm.name} onChange={(e) => setContactForm({ ...contactForm, name: e.target.value })} />
          </div>
          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">メール</label>
            <input className="input" type="email" value={contactForm.email} onChange={(e) => setContactForm({ ...contactForm, email: e.target.value })} />
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => setShowContact(false)}>キャンセル</button>
            <button className="btn-primary" onClick={() => { createContact.mutate({ name: contactForm.name, email: contactForm.email }, { onSuccess: () => { toast.success("追加しました"); setShowContact(false); setContactForm({ name: "", email: "" }); } }); }} disabled={!contactForm.name || !contactForm.email}>追加</button>
          </div>
        </div>
      </Modal>

      <ImportModal
        open={showImportTare}
        onClose={() => setShowImportTare(false)}
        dataType="tares"
        label="風袋"
        invalidateKeys={[["tares"]]}
      />
      <ImportModal
        open={showImportContact}
        onClose={() => setShowImportContact(false)}
        dataType="contacts"
        label="連絡先"
        invalidateKeys={[["contacts"]]}
      />
    </div>
  );
}
