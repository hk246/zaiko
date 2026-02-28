import { NavLink } from "react-router-dom";
import { clsx } from "clsx";

interface NavItem {
  to: string;
  label: string;
  icon: string;
}

const sections: { title: string; items: NavItem[] }[] = [
  {
    title: "概要",
    items: [
      { to: "/dashboard", label: "ダッシュボード", icon: "M3 12l2-2m0 0l7-7 7 7M5 10v10a1 1 0 001 1h3m10-11l2 2m-2-2v10a1 1 0 01-1 1h-3m-4 0h4" },
    ],
  },
  {
    title: "在庫",
    items: [
      { to: "/materials", label: "材料一覧", icon: "M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4" },
      { to: "/reservations", label: "予約", icon: "M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" },
      { to: "/recipes", label: "レシピ", icon: "M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" },
    ],
  },
  {
    title: "作業",
    items: [
      { to: "/work-orders", label: "作業工程", icon: "M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" },
      { to: "/work-orders/calendar", label: "カレンダー", icon: "M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" },
    ],
  },
  {
    title: "設定",
    items: [
      { to: "/labels", label: "ラベル", icon: "M7 7h.01M7 3h5c.512 0 1.024.195 1.414.586l7 7a2 2 0 010 2.828l-7 7a2 2 0 01-2.828 0l-7-7A1.994 1.994 0 013 12V7a4 4 0 014-4z" },
      { to: "/property-fields", label: "物性値", icon: "M4 6h16M4 10h16M4 14h16M4 18h16" },
      { to: "/settings", label: "設定", icon: "M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.066 2.573c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.573 1.066c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.066-2.573c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z M15 12a3 3 0 11-6 0 3 3 0 016 0z" },
      { to: "/backup", label: "バックアップ", icon: "M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" },
    ],
  },
];

function SvgIcon({ d }: { d: string }) {
  return (
    <svg className="h-5 w-5 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
      <path strokeLinecap="round" strokeLinejoin="round" d={d} />
    </svg>
  );
}

export function Sidebar() {
  return (
    <aside className="fixed inset-y-0 left-0 z-30 flex w-56 flex-col border-r border-gray-200 bg-white">
      {/* Header */}
      <div className="flex h-14 items-center gap-2 border-b border-gray-100 px-4">
        <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-primary-600">
          <svg viewBox="0 0 16 16" className="h-5 w-5" shapeRendering="crispEdges">
            {/* Pixel-art warehouse icon */}
            <rect x="7" y="1" width="2" height="1" fill="white" />
            <rect x="5" y="2" width="6" height="1" fill="white" />
            <rect x="3" y="3" width="10" height="1" fill="white" />
            <rect x="2" y="4" width="12" height="1" fill="white" />
            <rect x="2" y="5" width="2" height="8" fill="white" />
            <rect x="12" y="5" width="2" height="8" fill="white" />
            <rect x="2" y="13" width="12" height="1" fill="white" />
            <rect x="6" y="9" width="4" height="4" fill="white" opacity="0.5" />
            <rect x="4" y="6" width="8" height="1" fill="white" opacity="0.4" />
            <rect x="4" y="8" width="8" height="1" fill="white" opacity="0.4" />
          </svg>
        </div>
        <span className="text-sm font-bold text-gray-800">在庫管理</span>
      </div>

      {/* Nav */}
      <nav className="flex-1 overflow-y-auto px-3 py-4 space-y-5">
        {sections.map((section) => (
          <div key={section.title}>
            <p className="mb-1.5 px-2 text-[11px] font-semibold uppercase tracking-wider text-gray-400">
              {section.title}
            </p>
            <ul className="space-y-0.5">
              {section.items.map((item) => (
                <li key={item.to}>
                  <NavLink
                    to={item.to}
                    className={({ isActive }) =>
                      clsx(
                        "flex items-center gap-2.5 rounded-lg px-2.5 py-2 text-sm font-medium transition-colors",
                        isActive
                          ? "bg-primary-50 text-primary-700"
                          : "text-gray-600 hover:bg-gray-50 hover:text-gray-900",
                      )
                    }
                  >
                    <SvgIcon d={item.icon} />
                    {item.label}
                  </NavLink>
                </li>
              ))}
            </ul>
          </div>
        ))}
      </nav>

      {/* Footer */}
      <div className="border-t border-gray-100 px-4 py-3" />
    </aside>
  );
}
