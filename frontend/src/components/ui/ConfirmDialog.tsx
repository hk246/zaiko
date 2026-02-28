import { createContext, useContext, useState, useCallback, useRef, type ReactNode } from "react";

interface ConfirmState {
  message: string;
  resolve: (result: boolean) => void;
}

const ConfirmContext = createContext<((message: string) => Promise<boolean>) | null>(null);

export function useConfirm() {
  const fn = useContext(ConfirmContext);
  if (!fn) throw new Error("useConfirm must be used within ConfirmProvider");
  return fn;
}

export function ConfirmProvider({ children }: { children: ReactNode }) {
  const [state, setState] = useState<ConfirmState | null>(null);
  const dialogRef = useRef<HTMLDialogElement>(null);

  const confirm = useCallback((message: string): Promise<boolean> => {
    return new Promise<boolean>((resolve) => {
      setState({ message, resolve });
      // open on next tick so ref is rendered
      setTimeout(() => dialogRef.current?.showModal(), 0);
    });
  }, []);

  const handleResult = (result: boolean) => {
    state?.resolve(result);
    dialogRef.current?.close();
    setState(null);
  };

  return (
    <ConfirmContext.Provider value={confirm}>
      {children}

      {state && (
        <dialog
          ref={dialogRef}
          onClose={() => handleResult(false)}
          className="rounded-xl border-0 p-0 shadow-2xl backdrop:bg-black/40 w-[380px]"
        >
          <div className="px-5 py-4">
            <div className="flex items-start gap-3">
              <div className="mt-0.5 flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-red-100">
                <svg className="h-5 w-5 text-red-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4.5c-.77-.833-2.694-.833-3.464 0L3.34 16.5c-.77.833.192 2.5 1.732 2.5z" />
                </svg>
              </div>
              <div>
                <h3 className="text-sm font-bold text-gray-800">確認</h3>
                <p className="mt-1 text-sm text-gray-600">{state.message}</p>
              </div>
            </div>
          </div>
          <div className="flex justify-end gap-2 border-t border-gray-100 px-5 py-3">
            <button
              className="btn-secondary btn-sm"
              onClick={() => handleResult(false)}
            >
              キャンセル
            </button>
            <button
              className="btn-sm rounded-lg bg-red-600 px-3 py-1.5 text-sm font-medium text-white hover:bg-red-700 transition-colors"
              onClick={() => handleResult(true)}
              autoFocus
            >
              実行
            </button>
          </div>
        </dialog>
      )}
    </ConfirmContext.Provider>
  );
}
