import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

function Button({ children, onClick, disabled = false, variant = "default", size = "md", className = "", type = "button" }) {
  return (
    <button type={type} onClick={onClick} disabled={disabled} className={`btn ${variant === "outline" ? "btn-outline" : "btn-default"} ${size === "sm" ? "btn-sm" : ""} ${className}`}>
      {children}
    </button>
  );
}
function Card({ children, className = "" }) { return <div className={`card ${className}`}>{children}</div>; }
function CardContent({ children, className = "" }) { return <div className={`card-content ${className}`}>{children}</div>; }

const STORAGE_KEY = "cleaning-two-workspaces-final-v1";
const todayDate = new Date().toISOString().slice(0, 10);
const monthName = new Date().toLocaleDateString("el-GR", { month: "long", year: "numeric" });

const workspaces = {
  units: {
    id: "units",
    title: "4 Μονάδες Πελάτη",
    subtitle: "4 μονάδες, 4 καθαρίστριες, σεντόνια κάθε 3 μέρες, stock και βαριές μέρες",
    hasCleaners: true,
    hasLinen: true,
    heavyDayLimit: 7,
    properties: [
      { id: "vrachos", name: "Vrachos Lofts", aliases: ["Vrachos Lofts, Άφυτος", "Vrachos Lofts", "vrachos"], area: "Μονάδες", rooms: ["V01", "V02", "V03", "V04", "V05", "V06", "V07"], user: "vrachos", pass: "1234" },
      { id: "inlofts", name: "In Lofts", aliases: ["InLofts, Άφυτος", "InLofts", "In Lofts", "inlofts"], area: "Μονάδες", rooms: ["I01", "I02", "I03", "I04", "I05", "I06", "I07"], user: "inlofts", pass: "1234" },
      { id: "theros", name: "Theros Apartments", aliases: ["THEROS APARTMENTS", "Theros Apartments", "theros"], area: "Μονάδες", rooms: ["S1", "S2", "S3", "S4", "S5", "S6", "A1", "A2", "A3", "A4", "A5", "A6"], user: "theros", pass: "1234" },
      { id: "okkio", name: "Okkio Sea View", aliases: ["Okkio Sea View, Γαλήνη", "Okkio Sea View", "Okkio Hotel", "okkio"], area: "Μονάδες", rooms: ["No1", "No2", "No3", "No4"], user: "okkio", pass: "1234" },
    ],
    linenStock: {
      vrachos: { cleanSets: 18, usedSets: 3, minimum: 10 },
      inlofts: { cleanSets: 16, usedSets: 4, minimum: 10 },
      theros: { cleanSets: 24, usedSets: 7, minimum: 16 },
      okkio: { cleanSets: 6, usedSets: 3, minimum: 6 },
    },
    reservations: [
      { propertyId: "vrachos", room: "V02", checkIn: "2026-05-01", checkOut: "2026-05-07" },
      { propertyId: "vrachos", room: "V02", checkIn: "2026-05-07", checkOut: "2026-05-12" },
      { propertyId: "inlofts", room: "I03", checkIn: "2026-05-01", checkOut: "2026-05-05" },
      { propertyId: "theros", room: "S3", checkIn: "2026-05-01", checkOut: "2026-05-06" },
      { propertyId: "okkio", room: "No2", checkIn: "2026-05-01", checkOut: "2026-05-04" },
    ],
  },
  sithon: {
    id: "sithon",
    title: "Τα δικά μου καταλύματα",
    subtitle: "Μία ενιαία προβολή για όλα τα σπίτια που καθαρίζεις εσύ",
    hasCleaners: false,
    hasLinen: false,
    heavyDayLimit: 4,
    properties: [
      { id: "vrachos_lofts", name: "Vrachos Lofts", aliases: ["Vrachos Lofts"], area: "Vrachos" },
      { id: "valti_lux_studio_1", name: "Valti Lux Studio No1", aliases: ["Valti Lux Studio No1"], area: "Sikia Beach" },
      { id: "valti_lux_studio_2", name: "Valti Lux Studio No2", aliases: ["Valti Lux Studio No2"], area: "Sikia Beach" },
      { id: "sai_luxury_sea_view", name: "Sai Luxury Sea View", aliases: ["Sai Luxury Sea View"], area: "Sithonia" },
      { id: "luxhr_m", name: "LuxHR M Maisonette", aliases: ["LuxHR M Maisonette"], area: "Nikiti" },
      { id: "farm_apartment", name: "Farm Apartment", aliases: ["Farm Apartment"], area: "Sykia" },
    ],
    linenStock: {},
    reservations: [
      { propertyId: "vrachos_lofts", checkIn: "2026-05-01", checkOut: "2026-05-04" },
      { propertyId: "valti_lux_studio_1", checkIn: "2026-05-01", checkOut: "2026-05-04" },
      { propertyId: "valti_lux_studio_2", checkIn: "2026-05-02", checkOut: "2026-05-04" },
      { propertyId: "sai_luxury_sea_view", checkIn: "2026-05-01", checkOut: "2026-05-04" },
    ],
  },
};

function getWorkspace(id) { return workspaces[id] || workspaces.units; }
function propertyById(workspace, id) { return workspace.properties.find((p) => p.id === id); }
function propertyName(workspace, id) { return propertyById(workspace, id)?.name || id; }
function propertyArea(workspace, id) { return propertyById(workspace, id)?.area || ""; }
function addDays(dateString, days) { const date = new Date(`${dateString}T00:00:00`); date.setDate(date.getDate() + days); return date.toISOString().slice(0, 10); }
function isBefore(a, b) { return new Date(`${a}T00:00:00`) < new Date(`${b}T00:00:00`); }

const users = [
  { username: "admin", password: "admin2026", role: "admin", name: "Admin", defaultWorkspaceId: "units" },
  { username: "client", password: "client2026", role: "client", name: "Πελάτης 4 Μονάδων", defaultWorkspaceId: "units" },
  { username: "sithon", password: "sithon2026", role: "owner", name: "Τα δικά μου καταλύματα", defaultWorkspaceId: "sithon" },
  { username: "vrachos", password: "vrachos2026", role: "cleaner", name: "Vrachos Lofts", defaultWorkspaceId: "units", propertyId: "vrachos" },
  { username: "inlofts", password: "inlofts2026", role: "cleaner", name: "In Lofts", defaultWorkspaceId: "units", propertyId: "inlofts" },
  { username: "theros", password: "theros2026", role: "cleaner", name: "Theros Apartments", defaultWorkspaceId: "units", propertyId: "theros" },
  { username: "okkio", password: "okkio2026", role: "cleaner", name: "Okkio Sea View", defaultWorkspaceId: "units", propertyId: "okkio" },
];
function authenticate(username, password) {
  const found = users.find((user) => user.username === username && user.password === password);
  if (!found) return null;
  return { role: found.role, name: found.name, workspaceId: found.defaultWorkspaceId, propertyId: found.propertyId || null };
}
function hasSameDayArrival(reservations, reservation, workspace) {
  return reservations.some((item) => item.propertyId === reservation.propertyId && (!workspace.hasCleaners || item.room === reservation.room) && item.checkIn === reservation.checkOut && item !== reservation);
}
function generateTasks(workspace) {
  let id = 1;
  const tasks = [];
  workspace.reservations.forEach((reservation) => {
    const sameDayArrival = hasSameDayArrival(workspace.reservations, reservation, workspace);
    tasks.push({ id: id++, workspaceId: workspace.id, source: "excel", date: reservation.checkOut, propertyId: reservation.propertyId, propertyName: propertyName(workspace, reservation.propertyId), area: propertyArea(workspace, reservation.propertyId), room: reservation.room || "", guestName: reservation.guestName || "", type: sameDayArrival ? "Check-out + άφιξη ίδια μέρα" : "Check-out καθαριότητα", urgent: sameDayArrival, linen: workspace.hasLinen, status: "pending", note: "" });
    if (workspace.hasLinen) {
      let linenDate = addDays(reservation.checkIn, 3);
      while (isBefore(linenDate, reservation.checkOut)) {
        tasks.push({ id: id++, workspaceId: workspace.id, source: "auto-linen", date: linenDate, propertyId: reservation.propertyId, propertyName: propertyName(workspace, reservation.propertyId), area: propertyArea(workspace, reservation.propertyId), room: reservation.room || "", guestName: reservation.guestName || "", type: "Αλλαγή σεντονιών / πετσετών", urgent: false, linen: true, status: "pending", note: "Αυτόματα κάθε 3 μέρες" });
        linenDate = addDays(linenDate, 3);
      }
    }
  });
  return tasks.sort((a, b) => a.date.localeCompare(b.date) || a.propertyName.localeCompare(b.propertyName) || a.room.localeCompare(b.room));
}
function initialState() { return { activeWorkspaceId: "units", tasksByWorkspace: { units: generateTasks(workspaces.units), sithon: generateTasks(workspaces.sithon) }, linenStock: workspaces.units.linenStock }; }
function loadSavedState() { try { const raw = localStorage.getItem(STORAGE_KEY); if (!raw) return null; const parsed = JSON.parse(raw); return parsed.tasksByWorkspace ? parsed : null; } catch { return null; } }
function saveState(state) { try { localStorage.setItem(STORAGE_KEY, JSON.stringify({ ...state, savedAt: new Date().toISOString() })); return true; } catch { return false; } }
function clearSavedState() { try { localStorage.removeItem(STORAGE_KEY); } catch {} }
function filterTasks(tasks, selectedProperty, selectedArea, user) { let list = tasks; if (user?.role === "cleaner") list = list.filter((task) => task.propertyId === user.propertyId); if (selectedProperty !== "all") list = list.filter((task) => task.propertyId === selectedProperty); if (selectedArea !== "all") list = list.filter((task) => task.area === selectedArea); return list; }
function filterTasksByView(tasks, viewMode) { return viewMode === "today" ? tasks.filter((task) => task.date === todayDate) : tasks; }
function getStats(tasks) { const done = tasks.filter((t) => t.status === "done").length; return { total: tasks.length, done, pending: tasks.length - done, urgent: tasks.filter((t) => t.urgent && t.status !== "done").length, linen: tasks.filter((t) => t.source === "auto-linen").length }; }
function formatDate(date) { const [, month, day] = date.split("-"); return `${day}/${month}`; }
function groupTasksByDate(tasks) { return tasks.reduce((g, t) => { if (!g[t.date]) g[t.date] = []; g[t.date].push(t); return g; }, {}); }
function groupBy(items, key) { return items.reduce((g, item) => { const value = item[key] || "Χωρίς περιοχή"; if (!g[value]) g[value] = []; g[value].push(item); return g; }, {}); }
function getHeavyDays(tasks, workspace) { return Object.entries(groupTasksByDate(tasks.filter((t) => t.status !== "done"))).map(([date, dayTasks]) => ({ date, total: dayTasks.length, heavy: dayTasks.length >= workspace.heavyDayLimit, areaGroups: groupBy(dayTasks, "area") })).filter((d) => d.heavy).sort((a, b) => a.date.localeCompare(b.date)); }
function sortTodayRoute(tasks) { return [...tasks].sort((a, b) => Number(b.urgent) - Number(a.urgent) || a.area.localeCompare(b.area) || a.propertyName.localeCompare(b.propertyName)); }
function getTodayTasks(tasks) { return tasks.filter((task) => task.date === todayDate && task.status !== "done"); }
function isHeavyToday(tasks, workspace) { return getTodayTasks(tasks).length >= workspace.heavyDayLimit; }
function stockStatus(stock) { if (!stock) return "ok"; if (stock.cleanSets <= stock.minimum) return "low"; if (stock.cleanSets <= stock.minimum + 3) return "warning"; return "ok"; }
function getReports(tasks, workspace) { return workspace.properties.map((p) => { const list = tasks.filter((t) => t.propertyId === p.id); const done = list.filter((t) => t.status === "done").length; return { propertyId: p.id, propertyName: p.name, area: p.area, total: list.length, done, urgent: list.filter((t) => t.urgent).length, linen: list.filter((t) => t.source === "auto-linen").length, completion: list.length ? Math.round((done / list.length) * 100) : 0 }; }); }
function useOnlineStatus() { const [isOnline, setIsOnline] = useState(typeof navigator === "undefined" ? true : navigator.onLine); useEffect(() => { const on = () => setIsOnline(true); const off = () => setIsOnline(false); window.addEventListener("online", on); window.addEventListener("offline", off); return () => { window.removeEventListener("online", on); window.removeEventListener("offline", off); }; }, []); return isOnline; }
function canSwitchWorkspaces(user) { return user?.role === "admin"; }
function canUploadExcel(user) { return user?.role === "admin" || user?.role === "client" || user?.role === "owner"; }
function excelDateToIso(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") { const parsed = XLSX.SSF.parse_date_code(value); if (!parsed) return ""; return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`; }
  const text = String(value).trim(); const separator = text.includes("/") ? "/" : text.includes("-") ? "-" : text.includes(".") ? "." : "";
  if (separator) { const parts = text.split(separator); if (parts.length === 3) { const first = parts[0].padStart(2, "0"); const second = parts[1].padStart(2, "0"); const year = parts[2].length === 2 ? "20" + parts[2] : parts[2]; return `${year}-${second}-${first}`; } }
  const parsed = new Date(text); if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10); return "";
}
function normalizeKey(value) { return String(value || "").toLowerCase().replaceAll(" ", "").replaceAll("-", "").replaceAll("_", "").replaceAll(",", ""); }
function pickValue(row, names) { const normalized = {}; Object.keys(row).forEach((key) => { normalized[normalizeKey(key)] = row[key]; }); for (const name of names) { const key = normalizeKey(name); if (normalized[key] !== undefined) return normalized[key]; } return ""; }
function findPropertyFromText(workspace, text) { const value = normalizeKey(text); if (!value) return null; return workspace.properties.find((property) => { const candidates = [property.id, property.name, ...(property.aliases || [])].map(normalizeKey); return candidates.some((candidate) => value === candidate || value.includes(candidate) || candidate.includes(value)); }) || null; }
function parseExcelRowsToReservations(rows, workspace, selectedUploadProperty) {
  const reservations = []; const errors = [];
  rows.forEach((row, index) => {
    const roomOrProperty = pickValue(row, ["Room", "ROOM", "Δωμάτιο", "Κατάλυμα", "Property", "Accommodation", "Unit"]);
    const guestName = String(pickValue(row, ["Όνομα", "Ονομα", "Name", "Guest", "Guest Name", "Πελάτης", "Πελατης"] ) || "").trim();
    const checkIn = excelDateToIso(pickValue(row, ["Check In", "Check-in", "CheckIn", "Arrival", "Άφιξη", "Αφιξη"]));
    const checkOut = excelDateToIso(pickValue(row, ["Check Out", "Check-out", "CheckOut", "Departure", "Αναχώρηση", "Αναχωρηση"]));
    if (!checkIn || !checkOut) { errors.push("Γραμμή " + (index + 2) + ": λείπει Check-in ή Check-out"); return; }
    if (workspace.id === "units") {
      const propertyId = selectedUploadProperty !== "all" ? selectedUploadProperty : pickValue(row, ["Μονάδα", "Unit", "Property", "Κατάλυμα"]);
      const selectedProperty = workspace.properties.find((property) => property.id === propertyId) || findPropertyFromText(workspace, propertyId);
      if (!selectedProperty) { errors.push("Γραμμή " + (index + 2) + ": διάλεξε μονάδα ή βάλε στήλη Μονάδα"); return; }
      reservations.push({ propertyId: selectedProperty.id, room: String(roomOrProperty || "").trim(), guestName, checkIn, checkOut }); return;
    }
    const property = findPropertyFromText(workspace, roomOrProperty) || findPropertyFromText(workspace, pickValue(row, ["Κατάλυμα", "Property", "Accommodation"]));
    if (!property) { errors.push("Γραμμή " + (index + 2) + ": δεν βρέθηκε κατάλυμα για " + roomOrProperty); return; }
    reservations.push({ propertyId: property.id, guestName, checkIn, checkOut });
  });
  return { reservations, errors };
}
function generateTasksFromReservations(workspace, reservations) { return generateTasks({ ...workspace, reservations }); }

export default function CleaningTwoWorkspaces() {
  const saved = useMemo(() => loadSavedState(), []);
  const base = useMemo(() => saved || initialState(), [saved]);
  const [user, setUser] = useState(null);
  const [username, setUsername] = useState("admin");
  const [password, setPassword] = useState("admin2026");
  const [loginError, setLoginError] = useState("");
  const [activeWorkspaceId, setActiveWorkspaceId] = useState(base.activeWorkspaceId || "units");
  const [tasksByWorkspace, setTasksByWorkspace] = useState(base.tasksByWorkspace);
  const [linenStock, setLinenStock] = useState(base.linenStock || workspaces.units.linenStock);
  const [selectedProperty, setSelectedProperty] = useState("all");
  const [selectedUploadProperty, setSelectedUploadProperty] = useState("all");
  const [selectedArea, setSelectedArea] = useState("all");
  const [viewMode, setViewMode] = useState("today");
  const [tab, setTab] = useState("schedule");
  const [message, setMessage] = useState(saved ? "Φορτώθηκαν αποθηκευμένα δεδομένα." : "");
  const [lastSaved, setLastSaved] = useState(saved?.savedAt || null);
  const isOnline = useOnlineStatus();
  const workspace = getWorkspace(user?.role === "cleaner" || user?.role === "client" ? "units" : user?.role === "owner" ? "sithon" : activeWorkspaceId);
  const tasks = tasksByWorkspace[workspace.id] || [];
  const areas = useMemo(() => [...new Set(workspace.properties.map((p) => p.area))].sort(), [workspace]);
  useEffect(() => { const ok = saveState({ activeWorkspaceId, tasksByWorkspace, linenStock }); if (ok) setLastSaved(new Date().toISOString()); }, [activeWorkspaceId, tasksByWorkspace, linenStock]);
  const login = () => { const nextUser = authenticate(username.trim(), password.trim()); if (!nextUser) { setLoginError("Λάθος username ή password."); return; } setLoginError(""); setUser(nextUser); setActiveWorkspaceId(nextUser.workspaceId || "units"); };
  const filteredTasks = useMemo(() => filterTasks(tasks, selectedProperty, selectedArea, user), [tasks, selectedProperty, selectedArea, user]);
  const displayedTasks = useMemo(() => filterTasksByView(filteredTasks, viewMode), [filteredTasks, viewMode]);
  const groupedTasks = useMemo(() => groupTasksByDate(displayedTasks), [displayedTasks]);
  const stats = useMemo(() => getStats(displayedTasks), [displayedTasks]);
  const heavyDays = useMemo(() => getHeavyDays(filteredTasks, workspace), [filteredTasks, workspace]);
  const reports = useMemo(() => getReports(tasks, workspace), [tasks, workspace]);
  const switchWorkspace = (id) => { setActiveWorkspaceId(id); setSelectedProperty("all"); setSelectedArea("all"); setViewMode("today"); setTab("schedule"); };
  const updateTask = (id, patch) => setTasksByWorkspace((prev) => ({ ...prev, [workspace.id]: prev[workspace.id].map((t) => (t.id === id ? { ...t, ...patch } : t)) }));
  const markDone = (id) => { const task = tasks.find((t) => t.id === id); if (workspace.hasLinen && task?.linen && task.status !== "done") setLinenStock((prev) => ({ ...prev, [task.propertyId]: { ...prev[task.propertyId], cleanSets: Math.max(0, prev[task.propertyId].cleanSets - 1), usedSets: prev[task.propertyId].usedSets + 1 } })); updateTask(id, { status: "done", note: task?.note || "Ολοκληρώθηκε" }); };
  const addNote = (id, note) => updateTask(id, { note });
  const washLinen = (propertyId) => setLinenStock((prev) => ({ ...prev, [propertyId]: { ...prev[propertyId], cleanSets: prev[propertyId].cleanSets + prev[propertyId].usedSets, usedSets: 0 } }));
  const simulateExcelUpload = () => { const freshTasks = generateTasks(workspace); setTasksByWorkspace((prev) => ({ ...prev, [workspace.id]: freshTasks })); setMessage(workspace.id === "units" ? "Φορτώθηκε demo Excel για όλες τις 4 μονάδες." : "Φορτώθηκε demo Excel για τα δικά σου."); setViewMode("month"); setTab("schedule"); };
  const handleExcelFile = async (file) => {
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer(); const workbook = XLSX.read(buffer, { type: "array", cellDates: true }); let allReservations = []; let allErrors = [];
      workbook.SheetNames.forEach((sheetName) => { const sheet = workbook.Sheets[sheetName]; const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" }); let detectedProperty = null; if (workspace.id === "units") detectedProperty = findPropertyFromText(workspace, sheetName); const result = parseExcelRowsToReservations(rows, workspace, detectedProperty ? detectedProperty.id : selectedUploadProperty); allReservations = allReservations.concat(result.reservations); allErrors = allErrors.concat(result.errors); });
      if (!allReservations.length) { setMessage("Δεν δημιουργήθηκαν κρατήσεις από το Excel. " + allErrors.slice(0, 3).join(" • ")); return; }
      const freshTasks = generateTasksFromReservations(workspace, allReservations); setTasksByWorkspace((prev) => ({ ...prev, [workspace.id]: freshTasks })); setMessage("Το Excel " + file.name + " φορτώθηκε από " + workbook.SheetNames.length + " sheets: " + freshTasks.length + " εργασίες." + (allErrors.length ? " Προσοχή: " + allErrors.length + " γραμμές δεν διαβάστηκαν." : "")); setViewMode("month"); setTab("schedule");
    } catch (error) { setMessage("Σφάλμα στο Excel: " + error.message); }
  };
  const resetDemo = () => { clearSavedState(); const fresh = initialState(); setActiveWorkspaceId(fresh.activeWorkspaceId); setTasksByWorkspace(fresh.tasksByWorkspace); setLinenStock(fresh.linenStock); setSelectedProperty("all"); setSelectedArea("all"); setViewMode("today"); setTab("schedule"); setMessage("Έγινε reset και στα 2 συστήματα."); };
  const logout = () => { setUser(null); setSelectedProperty("all"); setSelectedArea("all"); setViewMode("today"); setTab("schedule"); };
  if (!user) return <LoginScreen username={username} setUsername={setUsername} password={password} setPassword={setPassword} login={login} loginError={loginError} isOnline={isOnline} />;
  return <div className="app"><div className="container"><div className="topbar"><div><h1>SithonBreak Cleaning System</h1><div className="subtitle-strong">{user.role === "cleaner" ? user.name : workspace.title}</div><p>{workspace.subtitle}</p></div><div className="top-actions"><StatusBadge isOnline={isOnline} lastSaved={lastSaved} /><Button variant="outline" onClick={logout}>🚪 Έξοδος</Button></div></div>
    {canSwitchWorkspaces(user) && <div className="workspace-grid"><WorkspaceCard active={workspace.id === "units"} title="4 Μονάδες Πελάτη" text="Με καθαρίστριες, σεντόνια, stock" onClick={() => switchWorkspace("units")} /><WorkspaceCard active={workspace.id === "sithon"} title="Τα δικά μου" text="Μία προβολή για όλα τα σπίτια που καθαρίζεις εσύ" onClick={() => switchWorkspace("sithon")} /></div>}
    <div className="tabs">{workspace.id === "sithon" && (user.role === "admin" || user.role === "owner") && <Button variant={tab === "today-work" ? "default" : "outline"} onClick={() => setTab("today-work")}>Σήμερα δουλεύω</Button>}<Button variant={tab === "schedule" ? "default" : "outline"} onClick={() => setTab("schedule")}>Πρόγραμμα</Button><Button variant={tab === "heavy" ? "default" : "outline"} onClick={() => setTab("heavy")}>Βαριές μέρες</Button>{workspace.hasLinen && (user.role === "admin" || user.role === "client") && <Button variant={tab === "linen" ? "default" : "outline"} onClick={() => setTab("linen")}>Stock σεντονιών</Button>}<Button variant={tab === "reports" ? "default" : "outline"} onClick={() => setTab("reports")}>Reports</Button></div>
    {canUploadExcel(user) && <UploadPanel workspace={workspace} user={user} simulateExcelUpload={simulateExcelUpload} handleExcelFile={handleExcelFile} resetDemo={resetDemo} selectedUploadProperty={selectedUploadProperty} setSelectedUploadProperty={setSelectedUploadProperty} />}
    {message && <div className="message">✅ {message}</div>}
    {tab === "today-work" && workspace.id === "sithon" ? <TodayWorkView tasks={tasks} workspace={workspace} markDone={markDone} addNote={addNote} /> : tab === "linen" && workspace.hasLinen ? <LinenStockView workspace={workspace} linenStock={linenStock} washLinen={washLinen} /> : tab === "heavy" ? <HeavyDaysView heavyDays={heavyDays} workspace={workspace} /> : tab === "reports" ? <ReportsView reports={reports} tasks={tasks} workspace={workspace} /> : <ScheduleView workspace={workspace} user={user} areas={areas} selectedArea={selectedArea} setSelectedArea={setSelectedArea} selectedProperty={selectedProperty} setSelectedProperty={setSelectedProperty} viewMode={viewMode} setViewMode={setViewMode} stats={stats} displayedTasks={displayedTasks} groupedTasks={groupedTasks} markDone={markDone} addNote={addNote} />}
  </div></div>;
}
function LoginScreen({ username, setUsername, password, setPassword, login, loginError, isOnline }) { return <div className="login-wrap"><Card className="login-card"><CardContent><div className="login-icon">🧹</div><h1>SithonBreak Cleaning System</h1><p>4 μονάδες πελάτη + δικά μου καταλύματα</p><input value={username} onChange={(e) => setUsername(e.target.value)} placeholder="Username" /><input value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Password" type="password" onKeyDown={(e) => e.key === "Enter" && login()} />{loginError && <div className="error">{loginError}</div>}<Button onClick={login} className="full">Είσοδος</Button><div className="login-info"><b>Admin:</b> admin / admin2026<br/><b>Πελάτης:</b> client / client2026<br/><b>Δικά σου:</b> sithon / sithon2026<br/><b>Καθαρίστριες:</b> vrachos, inlofts, theros, okkio / αντίστοιχο 2026<br/><b>Κατάσταση:</b> {isOnline ? "Online" : "Offline"}</div></CardContent></Card></div>; }
function UploadPanel({ workspace, user, simulateExcelUpload, handleExcelFile, resetDemo, selectedUploadProperty, setSelectedUploadProperty }) { const canChooseUnitExcel = workspace.id === "units" && (user.role === "admin" || user.role === "client"); const inputId = "excel-upload-" + workspace.id; return <Card><CardContent><div className="upload-row"><div><b>📤 Ανέβασμα Excel: {workspace.title}</b><p>{workspace.id === "units" ? "Για τον πελάτη ανεβάζεις Excel με tabs για κάθε μονάδα." : "Για τα δικά σου περνάς ένα Excel και βλέπεις όλα τα σπίτια μαζί."}</p></div><div className="actions"><input id={inputId} type="file" accept=".xlsx,.xls" className="hidden" onChange={(event) => handleExcelFile(event.target.files?.[0])} /><Button onClick={() => document.getElementById(inputId)?.click()}>Ανέβασμα Excel</Button><Button variant="outline" onClick={simulateExcelUpload}>Demo</Button>{user.role === "admin" && <Button variant="outline" onClick={resetDemo}>Reset</Button>}</div></div>{canChooseUnitExcel && <div><div className="label">Ποιο Excel ανεβάζεις;</div><div className="tabs"><Button variant={selectedUploadProperty === "all" ? "default" : "outline"} onClick={() => setSelectedUploadProperty("all")}>Όλες</Button>{workspace.properties.map((p) => <Button key={p.id} variant={selectedUploadProperty === p.id ? "default" : "outline"} onClick={() => setSelectedUploadProperty(p.id)}>{p.name}</Button>)}</div></div>}<div className="hint"><b>Στήλες:</b> Room/Κατάλυμα, Check-in, Check-out, Όνομα.</div></CardContent></Card>; }
function WorkspaceCard({ active, title, text, onClick }) { return <button onClick={onClick} className={`workspace-card ${active ? "active" : ""}`}><b>{title}</b><span>{text}</span></button>; }
function StatusBadge({ isOnline, lastSaved }) { return <div className={`status ${isOnline ? "ok" : "warn"}`}><b>{isOnline ? "Online" : "Offline"}</b><small>{lastSaved ? `Αποθήκευση: ${new Date(lastSaved).toLocaleTimeString("el-GR", { hour: "2-digit", minute: "2-digit" })}` : "Δεν έχει αποθηκευτεί"}</small></div>; }
function ScheduleView({ workspace, user, areas, selectedArea, setSelectedArea, selectedProperty, setSelectedProperty, viewMode, setViewMode, stats, displayedTasks, groupedTasks, markDone, addNote }) { const canFilter = user.role === "admin" || user.role === "client" || user.role === "owner"; return <><div className="tabs"><Button variant={viewMode === "today" ? "default" : "outline"} onClick={() => setViewMode("today")}>Σήμερα</Button><Button variant={viewMode === "month" ? "default" : "outline"} onClick={() => setViewMode("month")}>Όλος ο μήνας</Button></div><Stats stats={stats} workspace={workspace}/>{canFilter && <><div className="tabs"><Button variant={selectedArea === "all" ? "default" : "outline"} onClick={() => setSelectedArea("all")}>Όλες οι περιοχές</Button>{areas.map((area) => <Button key={area} variant={selectedArea === area ? "default" : "outline"} onClick={() => setSelectedArea(area)}>{area}</Button>)}</div><div className="tabs"><Button variant={selectedProperty === "all" ? "default" : "outline"} onClick={() => setSelectedProperty("all")}>Όλα</Button>{workspace.properties.map((p) => <Button key={p.id} variant={selectedProperty === p.id ? "default" : "outline"} onClick={() => setSelectedProperty(p.id)}>{p.name}</Button>)}</div></>}{displayedTasks.length === 0 ? <Empty /> : viewMode === "month" ? <MonthView groupedTasks={groupedTasks} markDone={markDone} addNote={addNote} /> : <div className="grid">{displayedTasks.map((task) => <TaskCard key={task.id} task={task} markDone={markDone} addNote={addNote} />)}</div>}</>; }
function Stats({ stats, workspace }) { return <div className="stats"><Stat icon="📅" title="Εργασίες" value={stats.total}/><Stat icon="⚠️" title="Ίδια μέρα" value={stats.urgent}/>{workspace.hasLinen && <Stat icon="🧺" title="Σεντόνια" value={stats.linen}/>}<Stat icon="✅" title="Ολοκληρωμένα" value={stats.done}/><Stat icon="⏳" title="Εκκρεμούν" value={stats.pending}/></div>; }
function TodayWorkView({ tasks, workspace, markDone, addNote }) { const todayTasks = sortTodayRoute(getTodayTasks(tasks)); const heavy = isHeavyToday(tasks, workspace); const areaGroups = groupBy(todayTasks, "area"); if (!todayTasks.length) return <Empty title="Σήμερα δεν έχεις καθαριότητες"/>; return <div><Card className={heavy ? "heavy-card" : ""}><CardContent><h2>Σήμερα δουλεύω</h2><p>Έχεις {todayTasks.length} σπίτια για καθαριότητα.</p>{heavy && <div className="heavy-alert">⚠️ Βαριά μέρα – πάρε άτομο</div>}</CardContent></Card>{Object.entries(areaGroups).map(([area, areaTasks]) => <Card key={area}><CardContent><h2>{area}</h2>{areaTasks.map((task, index) => <TaskLine key={task.id} task={task} index={index + 1} markDone={markDone} addNote={addNote}/>)}</CardContent></Card>)}</div>; }
function MonthView({ groupedTasks, markDone, addNote }) { const dates = Object.keys(groupedTasks).sort(); return <div>{dates.map((date) => { const dayTasks = groupedTasks[date]; return <Card key={date}><CardContent><h2>{formatDate(date)} <small>({dayTasks.length})</small></h2>{dayTasks.map((task) => <TaskLine key={task.id} task={task} markDone={markDone} addNote={addNote}/>)}</CardContent></Card>; })}</div>; }
function TaskLine({ task, index, markDone, addNote }) { return <div className="task-line"><div>{index && <span className="num">{index}</span>}<b>{task.propertyName} {task.room && `• ${task.room}`}</b>{task.guestName && <div className="muted">👤 {task.guestName}</div>}<div className="muted">{task.area} • {task.type}</div>{task.note && <div className="note">💬 {task.note}</div>}</div><TaskActions task={task} markDone={markDone} addNote={addNote}/></div>; }
function TaskCard({ task, markDone, addNote }) { return <Card><CardContent><div className="task-head"><div><div className="muted">{task.area} • {formatDate(task.date)}</div><h2>{task.propertyName}</h2>{task.room && <b>{task.room}</b>}{task.guestName && <div className="muted">👤 {task.guestName}</div>}</div>{task.urgent && task.status !== "done" && <span className="badge danger">Ίδια μέρα</span>}{task.status === "done" && <span className="badge success">Έγινε</span>}</div><div className="task-type">{task.source === "auto-linen" ? "🧺" : "🧹"} {task.type}</div>{task.note && <div className="note">💬 {task.note}</div>}<TaskActions task={task} markDone={markDone} addNote={addNote}/></CardContent></Card>; }
function TaskActions({ task, markDone, addNote }) { const writeNote = () => { const text = window.prompt("Γράψε σημείωση:", task.note || ""); if (text !== null) addNote(task.id, text.trim()); }; return <div className="actions"><Button size="sm" disabled={task.status === "done"} onClick={() => markDone(task.id)}>Ολοκληρώθηκε</Button><Button size="sm" variant="outline" onClick={writeNote}>Σημείωση</Button></div>; }
function HeavyDaysView({ heavyDays, workspace }) { if (!heavyDays.length) return <Empty title="Δεν υπάρχουν βαριές μέρες"/>; return <div>{heavyDays.map((day) => <Card key={day.date} className="heavy-card"><CardContent><h2>{formatDate(day.date)} - Βαριά μέρα</h2><p>Έχει {day.total} εργασίες. {workspace.id === "sithon" ? "Πιθανόν να χρειαστείς άτομο." : "Πιθανόν χρειάζεται ενίσχυση."}</p>{Object.entries(day.areaGroups).map(([area, tasks]) => <div key={area} className="hint"><b>{area}:</b> {tasks.map((t) => `${t.propertyName}${t.room ? ` ${t.room}` : ""}`).join(", ")}</div>)}</CardContent></Card>)}</div>; }
function LinenStockView({ workspace, linenStock, washLinen }) { return <div className="grid">{workspace.properties.map((p) => { const stock = linenStock[p.id]; const status = stockStatus(stock); return <Card key={p.id}><CardContent><h2>{p.name}</h2><span className={`badge ${status === "low" ? "danger" : status === "warning" ? "warn" : "success"}`}>{status === "low" ? "Χαμηλό" : status === "warning" ? "Προσοχή" : "ΟΚ"}</span><div className="stats two"><Stat icon="🧺" title="Καθαρά" value={stock?.cleanSets}/><Stat icon="🧼" title="Για πλύσιμο" value={stock?.usedSets}/></div><Button variant="outline" onClick={() => washLinen(p.id)}>Καταχώρηση πλυσίματος</Button></CardContent></Card>; })}</div>; }
function ReportsView({ reports, tasks, workspace }) { const totals = getStats(tasks); const heavyDays = getHeavyDays(tasks, workspace); return <div><div className="stats"><Stat icon="📊" title="Σύνολο" value={totals.total}/><Stat icon="✅" title="Έγιναν" value={totals.done}/><Stat icon="⚠️" title="Ίδια μέρα" value={totals.urgent}/><Stat icon="🔥" title="Βαριές" value={heavyDays.length}/></div><Card><CardContent><h2>Αναφορά</h2><div className="table-wrap"><table><thead><tr><th>Όνομα</th><th>Περιοχή</th><th>Εργασίες</th>{workspace.hasLinen && <th>Σεντόνια</th>}<th>Ίδια μέρα</th><th>%</th></tr></thead><tbody>{reports.map((r) => <tr key={r.propertyId}><td>{r.propertyName}</td><td>{r.area}</td><td>{r.total}</td>{workspace.hasLinen && <td>{r.linen}</td>}<td>{r.urgent}</td><td>{r.completion}%</td></tr>)}</tbody></table></div></CardContent></Card></div>; }
function Stat({ icon, title, value }) { return <Card className="stat"><CardContent><div className="stat-icon">{icon}</div><div><small>{title}</small><strong>{value}</strong></div></CardContent></Card>; }
function Empty({ title = "Δεν υπάρχουν εργασίες" }) { return <Card><CardContent><div className="empty">{title}</div></CardContent></Card>; }
