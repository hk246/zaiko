import { useState, useCallback, useEffect } from "react";

const STORAGE_KEY = "my_contact_id";

export function useMyName() {
  const [myContactId, setMyContactIdState] = useState<number | null>(() => {
    const stored = localStorage.getItem(STORAGE_KEY);
    return stored ? parseInt(stored, 10) : null;
  });

  const setMyContactId = useCallback((id: number | null) => {
    if (id === null) {
      localStorage.removeItem(STORAGE_KEY);
    } else {
      localStorage.setItem(STORAGE_KEY, String(id));
    }
    setMyContactIdState(id);
  }, []);

  // Listen for changes from other tabs / components
  useEffect(() => {
    const handler = (e: StorageEvent) => {
      if (e.key === STORAGE_KEY) {
        setMyContactIdState(e.newValue ? parseInt(e.newValue, 10) : null);
      }
    };
    window.addEventListener("storage", handler);
    return () => window.removeEventListener("storage", handler);
  }, []);

  return { myContactId, setMyContactId };
}
