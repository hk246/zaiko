import { Routes, Route, Navigate } from "react-router-dom";
import { AppShell } from "./components/layout/AppShell";
import Dashboard from "./pages/Dashboard";
import MaterialList from "./pages/MaterialList";
import MaterialDetail from "./pages/MaterialDetail";
import Reservations from "./pages/Reservations";
import Recipes from "./pages/Recipes";
import WorkOrders from "./pages/WorkOrders";
import WorkOrderCalendar from "./pages/WorkOrderCalendar";
import Labels from "./pages/Labels";
import PropertyFields from "./pages/PropertyFields";
import Settings from "./pages/Settings";
import Backup from "./pages/Backup";

export default function App() {
  return (
    <AppShell>
      <Routes>
        <Route path="/" element={<Navigate to="/dashboard" replace />} />
        <Route path="/dashboard" element={<Dashboard />} />
        <Route path="/materials" element={<MaterialList />} />
        <Route path="/materials/:id" element={<MaterialDetail />} />
        <Route path="/reservations" element={<Reservations />} />
        <Route path="/recipes" element={<Recipes />} />
        <Route path="/work-orders" element={<WorkOrders />} />
        <Route path="/work-orders/calendar" element={<WorkOrderCalendar />} />
        <Route path="/labels" element={<Labels />} />
        <Route path="/property-fields" element={<PropertyFields />} />
        <Route path="/settings" element={<Settings />} />
        <Route path="/backup" element={<Backup />} />
      </Routes>
    </AppShell>
  );
}
