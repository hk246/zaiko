// ── Label ──
export interface Label {
  id: number;
  name: string;
  color: string;
  created_at: string;
  material_count: number;
}

// ── Material ──
export interface Material {
  id: number;
  name: string;
  unit: string;
  min_weight: number;
  email: string | null;
  excel_path: string | null;
  action_type: string;
  contact_id: number | null;
  label_id: number | null;
  label: Label | null;
  created_at: string;
  current_stock: number;
  fraction_stock: number;
  predicted_stock: number;
  is_low_stock: boolean;
  lot_count: number;
  contact_name: string | null;
  contact_email: string | null;
}

export interface MaterialListItem {
  id: number;
  name: string;
  unit: string;
  min_weight: number;
  label: Label | null;
  current_stock: number;
  predicted_stock: number;
  is_low_stock: boolean;
  lot_count: number;
}

// ── Lot ──
export interface LotProperty {
  field_id: number;
  field_name: string;
  field_type: string;
  unit: string | null;
  value_string: string | null;
  value_number: number | null;
}

export interface Lot {
  id: number;
  material_id: number;
  lot_name: string;
  weight: number;
  is_fraction: boolean;
  created_at: string;
  properties: LotProperty[];
}

// ── Reservation ──
export interface Reservation {
  id: number;
  material_id: number;
  material_name: string;
  lot_id: number | null;
  lot_name: string | null;
  recipe_id: number | null;
  type: "use" | "replenish";
  quantity: number;
  actual_quantity: number | null;
  user_name: string | null;
  purpose: string | null;
  scheduled_date: string;
  status: string;
  executed_at: string | null;
  created_at: string;
  is_overdue: boolean;
}

// ── Recipe ──
export interface RecipeItem {
  id: number;
  material_id: number;
  material_name: string;
  quantity: number;
  lot_name: string | null;
}

export interface Recipe {
  id: number;
  name: string;
  description: string | null;
  type: string;
  items: RecipeItem[];
  created_at: string;
}

// ── WorkOrder ──
export interface WorkOrder {
  id: number;
  process_name: string;
  lot_name: string | null;
  material_id: number | null;
  material_name: string | null;
  status: string;
  priority: string;
  invoice_issued: boolean;
  worker_name: string | null;
  experiment_type: string;
  planned_start: string | null;
  planned_end: string | null;
  actual_start: string | null;
  actual_end: string | null;
  notes: string | null;
  archived: boolean;
  created_at: string;
  is_overdue: boolean;
  progress_pct: number;
}

// ── PropertyField ──
export interface PropertyField {
  id: number;
  label_id: number | null;
  name: string;
  field_type: string;
  unit: string | null;
  order_index: number;
  created_at: string;
}

// ── Master ──
export interface Tare {
  id: number;
  name: string;
  weight: number;
  description: string | null;
  created_at: string;
}

export interface Contact {
  id: number;
  name: string;
  email: string;
  description: string | null;
  created_at: string;
}

// ── Dashboard ──
export interface DashboardSummary {
  total_materials: number;
  total_lots: number;
  low_stock_count: number;
  pending_reservations: number;
  overdue_reservations: number;
  active_work_orders: number;
  overdue_work_orders: number;
}

export interface CriticalPeriod {
  material_id: number;
  material_name: string;
  start_date: string;
  end_date: string | null;
  shortage_amount: number;
  contact_id: number | null;
  contact_name: string | null;
  contact_email: string | null;
  action_type: string;
  current_stock: number;
}

export interface MaterialStats {
  period_label: string;
  total_used: number;
  total_replenished: number;
  net_change: number;
  use_count: number;
  replenish_count: number;
}

export interface TimelineEvent {
  date: string;
  stock: number;
  event_type: string | null;
  description: string;
}
