import { useContacts } from "@/api/hooks/useContacts";
import { useMyName } from "@/api/hooks/useMyName";

interface ContactPickerProps {
  value: string;
  onChange: (name: string) => void;
  className?: string;
  placeholder?: string;
}

/**
 * アドレス帳から担当者を選択するプルダウン。
 * 「自分」設定済みの場合は先頭に「自分 (名前)」オプションを表示。
 * 値は名前文字列（worker_name / user_name との互換性維持）。
 */
export function ContactPicker({ value, onChange, className = "input", placeholder }: ContactPickerProps) {
  const { data: contacts } = useContacts();
  const { myContactId } = useMyName();

  const myContact = contacts?.find((c) => c.id === myContactId);

  const handleChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const v = e.target.value;
    if (v === "__my__" && myContact) {
      onChange(myContact.name);
    } else {
      onChange(v);
    }
  };

  // Determine current select value: if value matches "my" contact, show __my__
  const selectValue = myContact && value === myContact.name ? "__my__" : value;

  return (
    <select className={className} value={selectValue} onChange={handleChange}>
      <option value="">{placeholder || "選択してください"}</option>
      {myContact && (
        <option value="__my__">自分 ({myContact.name})</option>
      )}
      {contacts
        ?.filter((c) => !myContact || c.id !== myContact.id)
        .map((c) => (
          <option key={c.id} value={c.name}>
            {c.name}
          </option>
        ))}
    </select>
  );
}
