import { useState, useRef } from "react";
import { useQueryClient } from "@tanstack/react-query";
import { Modal } from "@/components/ui/Modal";
import { useImportPreview, useImportConfirm } from "@/api/hooks/useImport";
import type { ImportPreview } from "@/api/hooks/useImport";
import { API_BASE } from "@/lib/constants";
import toast from "react-hot-toast";

interface ImportModalProps {
  open: boolean;
  onClose: () => void;
  dataType: string;
  label: string;
  materialId?: number;
  invalidateKeys?: string[][];
}

type Step = "select" | "preview" | "done";

export function ImportModal({ open, onClose, dataType, label, materialId, invalidateKeys }: ImportModalProps) {
  const qc = useQueryClient();
  const fileRef = useRef<HTMLInputElement>(null);
  const [step, setStep] = useState<Step>("select");
  const [file, setFile] = useState<File | null>(null);
  const [preview, setPreview] = useState<ImportPreview | null>(null);
  const [importedCount, setImportedCount] = useState(0);

  const previewMutation = useImportPreview();
  const confirmMutation = useImportConfirm();

  const reset = () => {
    setStep("select");
    setFile(null);
    setPreview(null);
    setImportedCount(0);
    if (fileRef.current) fileRef.current.value = "";
  };

  const handleClose = () => {
    reset();
    onClose();
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0] ?? null;
    setFile(f);
  };

  const handlePreview = () => {
    if (!file) return;
    previewMutation.mutate(
      { dataType, file, materialId },
      {
        onSuccess: (data) => {
          setPreview(data);
          setStep("preview");
        },
        onError: (e) => toast.error(e.message),
      },
    );
  };

  const handleConfirm = () => {
    if (!file) return;
    confirmMutation.mutate(
      { dataType, file, materialId },
      {
        onSuccess: (data) => {
          setImportedCount(data.imported);
          setStep("done");
          toast.success(data.message);
          if (invalidateKeys) {
            invalidateKeys.forEach((key) => qc.invalidateQueries({ queryKey: key }));
          }
        },
        onError: (e) => toast.error(e.message),
      },
    );
  };

  const templateUrl = `${API_BASE}/import/template/${dataType}`;

  return (
    <Modal open={open} onClose={handleClose} title={`${label}インポート`} wide>
      {step === "select" && (
        <div className="space-y-5">
          <div>
            <p className="mb-2 text-sm text-gray-600">
              Excelファイル (.xlsx) を選択してインポートします。
            </p>
            <a
              href={templateUrl}
              download
              className="inline-flex items-center gap-1.5 text-sm text-primary-600 hover:text-primary-700 hover:underline"
            >
              <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2M7 10l5 5 5-5M12 15V3" />
              </svg>
              テンプレートをダウンロード
            </a>
          </div>

          <div>
            <label className="mb-1 block text-sm font-medium text-gray-700">ファイル選択</label>
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileChange}
              className="block w-full text-sm text-gray-500 file:mr-3 file:rounded-lg file:border-0 file:bg-primary-50 file:px-4 file:py-2 file:text-sm file:font-medium file:text-primary-700 hover:file:bg-primary-100"
            />
          </div>

          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={handleClose}>キャンセル</button>
            <button
              className="btn-primary"
              onClick={handlePreview}
              disabled={!file || previewMutation.isPending}
            >
              {previewMutation.isPending ? "解析中..." : "プレビュー"}
            </button>
          </div>
        </div>
      )}

      {step === "preview" && preview && (
        <div className="space-y-4">
          {/* Summary */}
          <div className="flex items-center gap-4">
            <div className="rounded-lg bg-blue-50 px-4 py-2">
              <p className="text-xs text-blue-500">データ行数</p>
              <p className="text-lg font-bold text-blue-700">{preview.total_rows}</p>
            </div>
            <div className={`rounded-lg px-4 py-2 ${preview.valid ? "bg-green-50" : "bg-red-50"}`}>
              <p className={`text-xs ${preview.valid ? "text-green-500" : "text-red-500"}`}>ステータス</p>
              <p className={`text-lg font-bold ${preview.valid ? "text-green-700" : "text-red-700"}`}>
                {preview.valid ? "OK" : `${preview.errors.length}件のエラー`}
              </p>
            </div>
          </div>

          {/* Errors */}
          {preview.errors.length > 0 && (
            <div className="max-h-48 overflow-y-auto rounded-lg border border-red-200 bg-red-50">
              <table className="w-full text-left text-xs">
                <thead className="border-b border-red-200 bg-red-100/50">
                  <tr>
                    <th className="px-3 py-2 font-medium text-red-700">行</th>
                    <th className="px-3 py-2 font-medium text-red-700">フィールド</th>
                    <th className="px-3 py-2 font-medium text-red-700">エラー</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-red-100">
                  {preview.errors.map((err, i) => (
                    <tr key={i}>
                      <td className="px-3 py-1.5 text-red-600">{err.row}</td>
                      <td className="px-3 py-1.5 text-red-600">{err.field}</td>
                      <td className="px-3 py-1.5 text-red-600">{err.message}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          <div className="flex justify-end gap-2 pt-2">
            <button className="btn-secondary" onClick={() => { setStep("select"); setPreview(null); }}>
              戻る
            </button>
            <button
              className="btn-primary"
              onClick={handleConfirm}
              disabled={!preview.valid || confirmMutation.isPending}
            >
              {confirmMutation.isPending ? "インポート中..." : "インポート実行"}
            </button>
          </div>
        </div>
      )}

      {step === "done" && (
        <div className="space-y-4 text-center py-4">
          <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-full bg-green-100">
            <svg className="h-7 w-7 text-green-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
            </svg>
          </div>
          <p className="text-sm text-gray-700">{importedCount}件の{label}をインポートしました</p>
          <button className="btn-primary" onClick={handleClose}>閉じる</button>
        </div>
      )}
    </Modal>
  );
}
