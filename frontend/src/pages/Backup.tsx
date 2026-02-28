import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { api } from "@/api/client";
import { PageHeader } from "@/components/ui/PageHeader";
import { EmptyState } from "@/components/ui/EmptyState";
import { useConfirm } from "@/components/ui/ConfirmDialog";
import { formatDateTime } from "@/lib/format";
import toast from "react-hot-toast";

interface BackupItem {
  filename: string;
  size_bytes: number;
  created_at: string;
}

export default function Backup() {
  const qc = useQueryClient();
  const confirm = useConfirm();

  const { data: backups, isLoading } = useQuery({
    queryKey: ["backups"],
    queryFn: () => api.get<BackupItem[]>("/backup"),
  });

  const createMutation = useMutation({
    mutationFn: () => api.post("/backup"),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["backups"] });
      toast.success("バックアップを作成しました");
    },
    onError: (e) => toast.error(e.message),
  });

  const restoreMutation = useMutation({
    mutationFn: (filename: string) => api.post(`/backup/restore/${filename}`),
    onSuccess: () => toast.success("復元しました。再起動が必要です。"),
    onError: (e) => toast.error(e.message),
  });

  const deleteMutation = useMutation({
    mutationFn: (filename: string) => api.delete(`/backup/${filename}`),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["backups"] });
      toast.success("削除しました");
    },
  });

  const formatSize = (bytes: number) => {
    if (bytes >= 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
    return `${(bytes / 1024).toFixed(1)} KB`;
  };

  return (
    <div>
      <PageHeader
        title="バックアップ"
        description="データベースの保護と復元"
        actions={
          <button className="btn-primary" onClick={() => createMutation.mutate()} disabled={createMutation.isPending}>
            + バックアップ作成
          </button>
        }
      />

      {isLoading ? (
        <div className="py-20 text-center text-gray-400">読み込み中...</div>
      ) : !backups?.length ? (
        <EmptyState title="バックアップがありません" description="データを保護するためにバックアップを作成しましょう" action={{ label: "バックアップ作成", onClick: () => createMutation.mutate() }} />
      ) : (
        <div className="space-y-2">
          {backups.map((b) => (
            <div key={b.filename} className="card flex items-center justify-between">
              <div>
                <p className="text-sm font-medium text-gray-800">{b.filename}</p>
                <p className="text-xs text-gray-500">
                  {formatDateTime(b.created_at)} &middot; {formatSize(b.size_bytes)}
                </p>
              </div>
              <div className="flex items-center gap-2">
                <a
                  href={`/api/backup/download/${b.filename}`}
                  className="btn-secondary btn-sm"
                  download
                >
                  DL
                </a>
                <button
                  className="btn-secondary btn-sm"
                  onClick={async () => { if (await confirm("このバックアップから復元しますか？")) restoreMutation.mutate(b.filename); }}
                >
                  復元
                </button>
                <button
                  className="text-xs text-red-400 hover:text-red-600"
                  onClick={async () => { if (await confirm("削除しますか？")) deleteMutation.mutate(b.filename); }}
                >
                  削除
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
