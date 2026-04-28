import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";

const STORAGE_KEY = "cleaning-two-workspaces-final-v1";
const todayDate = "2026-05-04";
const monthName = "Μάιος 2026";

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
      { propertyId: "vrachos", room: "V05", checkIn: "2026-05-01", checkOut: "2026-05-06" },
      { propertyId: "inlofts", room: "I03", checkIn: "2026-05-01", checkOut: "2026-05-05" },
      { propertyId: "inlofts", room: "I07", checkIn: "2026-05-01", checkOut: "2026-05-08" },
      { propertyId: "theros", room: "S3", checkIn: "2026-05-01", checkOut: "2026-05-06" },
      { propertyId: "theros", room: "S3", checkIn: "2026-05-06", checkOut: "2026-05-11" },
      { propertyId: "theros", room: "A2", checkIn: "2026-05-01", checkOut: "2026-05-10" },
      { propertyId: "okkio", room: "No2", checkIn: "2026-05-01", checkOut: "2026-05-04" },
      { propertyId: "okkio", room: "No2", checkIn: "2026-05-04", checkOut: "2026-05-09" },
      { propertyId: "okkio", room: "No4", checkIn: "2026-05-01", checkOut: "2026-05-06" },
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
      { id: "vrachos_lofts", name: "Vrachos Lofts", area: "Vrachos" },
      { id: "valti_lux_studio_1", name: "Valti Lux Studio No1", area: "Sikia Beach" },
      { id: "valti_lux_studio_2", name: "Valti Lux Studio No2", area: "Sikia Beach" },
      { id: "sai_luxury_sea_view", name: "Sai Luxury Sea View", area: "Sithonia" },
      { id: "luxhr_m", name: "LuxHR M Maisonette", area: "Nikiti" },
      { id: "farm_apartment", name: "Farm Apartment", area: "Sykia" },
    ],
    linenStock: {},
    reservations: [
      { propertyId: "vrachos_lofts", checkIn: "2026-05-01", checkOut: "2026-05-04" },
      { propertyId: "valti_lux_studio_1", checkIn: "2026-05-01", checkOut: "2026-05-04" },
      { propertyId: "valti_lux_studio_2", checkIn: "2026-05-02", checkOut: "2026-05-04" },
      { propertyId: "sai_luxury_sea_view", checkIn: "2026-05-01", checkOut: "2026-05-04" },
      { propertyId: "luxhr_m", checkIn: "2026-05-01", checkOut: "2026-05-06" },
      { propertyId: "farm_apartment", checkIn: "2026-05-03", checkOut: "2026-05-06" },
      { propertyId: "vrachos_lofts", checkIn: "2026-05-04", checkOut: "2026-05-10" },
      { propertyId: "valti_lux_studio_1", checkIn: "2026-05-04", checkOut: "2026-05-09" },
      { propertyId: "sai_luxury_sea_view", checkIn: "2026-05-06", checkOut: "2026-05-12" },
      { propertyId: "valti_lux_studio_2", checkIn: "2026-05-08", checkOut: "2026-05-13" },
      { propertyId: "luxhr_m", checkIn: "2026-05-10", checkOut: "2026-05-15" },
      { propertyId: "farm_apartment", checkIn: "2026-05-10", checkOut: "2026-05-15" },
      { propertyId: "vrachos_lofts", checkIn: "2026-05-15", checkOut: "2026-05-20" },
      { propertyId: "valti_lux_studio_1", checkIn: "2026-05-15", checkOut: "2026-05-20" },
      { propertyId: "valti_lux_studio_2", checkIn: "2026-05-16", checkOut: "2026-05-20" },
      { propertyId: "sai_luxury_sea_view", checkIn: "2026-05-17", checkOut: "2026-05-20" },
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
  { username: "okkio", password: "okkio2026", role: "cleaner", name: "Okkio Hotel", defaultWorkspaceId: "units", propertyId: "okkio" },
];

function authenticate(username, password) {
  const found = users.find((user) => user.username === username && user.password === password);
  if (!found) return null;
  return {
    role: found.role,
    name: found.name,
    workspaceId: found.defaultWorkspaceId,
    propertyId: found.propertyId || null,
  };
}

function hasSameDayArrival(reservations, reservation, workspace) {
  return reservations.some((item) => item.propertyId === reservation.propertyId && (!workspace.hasCleaners || item.room === reservation.room) && item.checkIn === reservation.checkOut && item !== reservation);
}

function generateTasks(workspace) {
  let id = 1;
  const tasks = [];
  workspace.reservations.forEach((reservation) => {
    const sameDayArrival = hasSameDayArrival(workspace.reservations, reservation, workspace);
    tasks.push({
      id: id++, workspaceId: workspace.id, source: "excel", date: reservation.checkOut,
      propertyId: reservation.propertyId, propertyName: propertyName(workspace, reservation.propertyId), area: propertyArea(workspace, reservation.propertyId),
      room: reservation.room || "", guestName: reservation.guestName || "", type: sameDayArrival ? "Check-out + άφιξη ίδια μέρα" : "Check-out καθαριότητα",
      urgent: sameDayArrival, linen: workspace.hasLinen, status: "pending", note: "",
    });
    if (workspace.hasLinen) {
      let linenDate = addDays(reservation.checkIn, 3);
      while (isBefore(linenDate, reservation.checkOut)) {
        tasks.push({
          id: id++, workspaceId: workspace.id, source: "auto-linen", date: linenDate,
          propertyId: reservation.propertyId, propertyName: propertyName(workspace, reservation.propertyId), area: propertyArea(workspace, reservation.propertyId),
          room: reservation.room || "", guestName: reservation.guestName || "", type: "Αλλαγή σεντονιών / πετσετών", urgent: false, linen: true, status: "pending", note: "Αυτόματα κάθε 3 μέρες",
        });
        linenDate = addDays(linenDate, 3);
      }
    }
  });
  return tasks.sort((a, b) => a.date.localeCompare(b.date) || a.propertyName.localeCompare(b.propertyName) || a.room.localeCompare(b.room));
}

function initialState() {
  return { activeWorkspaceId: "units", tasksByWorkspace: { units: generateTasks(workspaces.units), sithon: generateTasks(workspaces.sithon) }, linenStock: workspaces.units.linenStock };
}
function loadSavedState() {
  // Δεν φορτώνουμε πια μεγάλα Excel από localStorage γιατί μπορεί να κολλήσει ο browser.
  return null;
}
function saveState() {
  // Disabled για σταθερότητα σε μεγάλα Excel / πολλά uploads.
  return false;
}
function clearSavedState() {
  try { localStorage.removeItem(STORAGE_KEY); } catch {}
}

function filterTasks(tasks, selectedProperty, selectedArea, user) {
  let list = tasks;
  if (user?.role === "cleaner") list = list.filter((task) => task.propertyId === user.propertyId);
  if (selectedProperty !== "all") list = list.filter((task) => task.propertyId === selectedProperty);
  if (selectedArea !== "all") list = list.filter((task) => task.area === selectedArea);
  return list;
}
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

function useOnlineStatus() {
  const [isOnline, setIsOnline] = useState(typeof navigator === "undefined" ? true : navigator.onLine);
  useEffect(() => { const on = () => setIsOnline(true); const off = () => setIsOnline(false); window.addEventListener("online", on); window.addEventListener("offline", off); return () => { window.removeEventListener("online", on); window.removeEventListener("offline", off); }; }, []);
  return isOnline;
}

function canSwitchWorkspaces(user) {
  return user?.role === "admin";
}

function canUploadExcel(user) {
  return user?.role === "admin" || user?.role === "client" || user?.role === "owner";
}

function canSeeSithon(user) {
  return user?.role === "admin" || user?.role === "owner";
}

function canSeeUnits(user) {
  return user?.role === "admin" || user?.role === "client" || user?.role === "cleaner";
}

function excelDateToIso(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return "";
    return String(parsed.y) + "-" + String(parsed.m).padStart(2, "0") + "-" + String(parsed.d).padStart(2, "0");
  }
  const text = String(value).trim();
  const separator = text.includes("/") ? "/" : text.includes("-") ? "-" : text.includes(".") ? "." : "";
  if (separator) {
    const parts = text.split(separator);
    if (parts.length === 3) {
      const first = parts[0].padStart(2, "0");
      const second = parts[1].padStart(2, "0");
      const year = parts[2].length === 2 ? "20" + parts[2] : parts[2];
      return year + "-" + second + "-" + first;
    }
  }
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return "";
}

function normalizeKey(value) {
  return String(value || "").toLowerCase().replaceAll(" ", "").replaceAll("-", "").replaceAll("_", "");
}

function pickValue(row, names) {
  const normalized = {};
  Object.keys(row).forEach((key) => { normalized[normalizeKey(key)] = row[key]; });
  for (const name of names) {
    const key = normalizeKey(name);
    if (normalized[key] !== undefined) return normalized[key];
  }
  return "";
}

function findPropertyFromText(workspace, text) {
  const value = normalizeKey(text);
  if (!value) return null;
  return workspace.properties.find((property) => {
    const candidates = [property.id, property.name, ...(property.aliases || [])].map(normalizeKey);
    return candidates.some((candidate) => value === candidate || value.includes(candidate) || candidate.includes(value));
  }) || null;
}

function shouldSkipSheet(sheetName) {
  const value = normalizeKey(sheetName);
  const blockedWords = ["ημερ", "ημερολογιο", "calendar", "ντουβας", "γιωργος", "prices", "τιμες"];
  return blockedWords.some((word) => value.includes(normalizeKey(word)));
}

function parseExcelRowsToReservations(rows, workspace, selectedUploadProperty) {
  const reservations = [];
  const errors = [];

  rows.forEach((row, index) => {
    const roomOrProperty = pickValue(row, ["Room", "ROOM", "Δωμάτιο", "Κατάλυμα", "Property", "Accommodation", "Unit"]);
    const guestName = String(pickValue(row, ["Όνομα", "Ονομα", "Name", "Guest", "Guest Name", "Πελάτης", "Πελατης"] ) || "").trim();
    const rawCheckIn = pickValue(row, ["Check In", "Check-in", "CheckIn", "Arrival", "Άφιξη", "Αφιξη"]);
    const rawCheckOut = pickValue(row, ["Check Out", "Check-out", "CheckOut", "Departure", "Αναχώρηση", "Αναχωρηση"]);

    // Αγνοεί τελείως κενές ή φορμαρισμένες γραμμές Excel
    if (!roomOrProperty && !rawCheckIn && !rawCheckOut && !guestName) return;

    const checkIn = excelDateToIso(rawCheckIn);
    const checkOut = excelDateToIso(rawCheckOut);

    if (!checkIn || !checkOut) {
      errors.push("Γραμμή " + (index + 2) + ": λείπει Check-in ή Check-out");
      return;
    }

    if (workspace.id === "units") {
      const propertyId = selectedUploadProperty !== "all" ? selectedUploadProperty : pickValue(row, ["Μονάδα", "Unit", "Property", "Κατάλυμα"]);
      const selectedProperty = workspace.properties.find((property) => property.id === propertyId) || findPropertyFromText(workspace, propertyId);
      if (!selectedProperty) {
        errors.push("Γραμμή " + (index + 2) + ": διάλεξε μονάδα ή βάλε στήλη Μονάδα");
        return;
      }
      reservations.push({ propertyId: selectedProperty.id, room: String(roomOrProperty || "").trim(), guestName, checkIn, checkOut });
      return;
    }

    const property = findPropertyFromText(workspace, roomOrProperty) || findPropertyFromText(workspace, pickValue(row, ["Κατάλυμα", "Property", "Accommodation"]));
    if (!property) {
      errors.push("Γραμμή " + (index + 2) + ": δεν βρέθηκε κατάλυμα για " + roomOrProperty);
      return;
    }
    reservations.push({ propertyId: property.id, guestName, checkIn, checkOut });
  });

  return { reservations, errors };
}

function generateTasksFromReservations(workspace, reservations) {
  return generateTasks({ ...workspace, reservations });
}

function taskKey(task) {
  return [task.workspaceId, task.propertyId, task.room || "", task.date, task.type, task.guestName || ""].join("|");
}

function mergeTasks(existingTasks, newTasks, replacePropertyId = null) {
  const baseTasks = replacePropertyId
    ? existingTasks.filter((task) => task.propertyId !== replacePropertyId)
    : existingTasks;

  const merged = [];
  const seen = new Set();

  [...baseTasks, ...newTasks].forEach((task) => {
    const key = taskKey(task);
    if (seen.has(key)) return;
    seen.add(key);
    merged.push(task);
  });

  return merged
    .map((task, index) => ({ ...task, id: index + 1 }))
    .sort((a, b) => a.date.localeCompare(b.date) || a.propertyName.localeCompare(b.propertyName) || a.room.localeCompare(b.room));
}

function runDemoTests() {
  console.assert(generateTasks(workspaces.units).some((t) => t.source === "auto-linen"), "Οι 4 μονάδες πρέπει να έχουν σεντόνια κάθε 3 μέρες.");
  console.assert(!generateTasks(workspaces.sithon).some((t) => t.source === "auto-linen"), "Τα δικά μου πρέπει να έχουν μόνο check-out.");
  console.assert(authenticate("vrachos", "vrachos2026")?.role === "cleaner", "Οι καθαρίστριες πρέπει να μπαίνουν στις 4 μονάδες.");
  console.assert(authenticate("client", "client2026")?.role === "client", "Ο πελάτης πρέπει να έχει δικό του login.");
  console.assert(authenticate("admin", "admin2026")?.role === "admin", "Το admin πρέπει να δουλεύει.");
  console.assert(authenticate("sithon", "sithon2026")?.role === "owner", "Ο κωδικός για τα δικά μου πρέπει να ανοίγει μόνο τα δικά μου.");
  console.assert(authenticate("admin", "admin") === null, "Ο παλιός admin κωδικός δεν πρέπει να δουλεύει.");
  console.assert(canSeeUnits(authenticate("client", "client2026")) === true, "Ο πελάτης πρέπει να βλέπει τις 4 μονάδες.");
  console.assert(canSeeSithon(authenticate("client", "client2026")) === false, "Ο πελάτης δεν πρέπει να βλέπει τα δικά μου.");
}
runDemoTests();

function Icon({ children, className = "" }) { return <span className={`inline-flex items-center justify-center ${className}`}>{children}</span>; }

export default function CleaningTwoWorkspaces() {
  const saved = useMemo(() => loadSavedState(), []);
  const base = useMemo(() => saved || initialState(), [saved]);
  const [user, setUser] = useState(null);
  const [username, setUsername] = useState("admin");
  const [password, setPassword] = useState("admin");
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
  const [isUploading, setIsUploading] = useState(false);
  const [lastSaved, setLastSaved] = useState(null);
  const isOnline = useOnlineStatus();

  const workspace = getWorkspace(user?.role === "cleaner" || user?.role === "client" ? "units" : user?.role === "owner" ? "sithon" : activeWorkspaceId);
  const tasks = tasksByWorkspace[workspace.id] || [];
  const areas = useMemo(() => [...new Set(workspace.properties.map((p) => p.area))].sort(), [workspace]);

  useEffect(() => {
    // Δεν κάνουμε auto-save μεγάλων δεδομένων για να μη γίνεται RESULT_CODE_HUNG.
  }, [activeWorkspaceId, tasksByWorkspace, linenStock]);

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
  const simulateExcelUpload = () => {
    const freshTasks = generateTasks(workspace);
    if (workspace.id === "units" && selectedUploadProperty !== "all") {
      const propertyTasks = freshTasks.filter((task) => task.propertyId === selectedUploadProperty);
      setTasksByWorkspace((prev) => ({
        ...prev,
        units: [...prev.units.filter((task) => task.propertyId !== selectedUploadProperty), ...propertyTasks].sort((a, b) => a.date.localeCompare(b.date) || a.propertyName.localeCompare(b.propertyName) || a.room.localeCompare(b.room)),
      }));
      setMessage("Φορτώθηκε ξεχωριστό demo Excel για " + propertyName(workspace, selectedUploadProperty) + ".");
    } else {
      setTasksByWorkspace((prev) => ({ ...prev, [workspace.id]: freshTasks }));
      setMessage(workspace.id === "units" ? "Φορτώθηκε demo Excel για όλες τις 4 μονάδες." : "Φορτώθηκε demo Excel για τα δικά σου: ενιαία προβολή όλων των σπιτιών ανά ημέρα.");
    }
    setViewMode("month");
    setTab("schedule");
  };

  const handleExcelFile = async (file) => {
    if (!file || isUploading) return;
    setIsUploading(true);
    setMessage("Διαβάζω το Excel, περίμενε λίγο...");
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });

      let allReservations = [];
      let allErrors = [];

      workbook.SheetNames.forEach((sheetName) => {
        if (shouldSkipSheet(sheetName)) return;

        let detectedProperty = null;
        if (workspace.id === "units") {
          detectedProperty = findPropertyFromText(workspace, sheetName);
          // Αν το Excel έχει πολλά sheets, διαβάζουμε μόνο αυτά που είναι γνωστές μονάδες.
          // Έτσι αγνοούμε ημερολόγια/βοηθητικά φύλλα όπως "ΗΜΕΡ ΝΤΟΥΒΑΣ ΓΙΩΡΓΟΣ".
          if (!detectedProperty && selectedUploadProperty === "all") return;
        }

        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", blankrows: false });

        const result = parseExcelRowsToReservations(
          rows,
          workspace,
          detectedProperty ? detectedProperty.id : selectedUploadProperty
        );

        allReservations = allReservations.concat(result.reservations);
        allErrors = allErrors.concat(result.errors);
      });

      if (!allReservations.length) {
        setMessage("Δεν δημιουργήθηκαν κρατήσεις από το Excel. " + allErrors.slice(0, 3).join(" • "));
        return;
      }

      const freshTasks = generateTasksFromReservations(workspace, allReservations);

      setTasksByWorkspace((prev) => {
        const currentTasks = prev[workspace.id] || [];
        const replacePropertyId = workspace.id === "units" && selectedUploadProperty !== "all" ? selectedUploadProperty : null;
        return {
          ...prev,
          [workspace.id]: mergeTasks(currentTasks, freshTasks, replacePropertyId),
        };
      });
      setSelectedProperty("all");
      setSelectedArea("all");

      setMessage(
        "Το Excel " + file.name + " προστέθηκε από " + workbook.SheetNames.length + " sheets: " +
        freshTasks.length + " νέες εργασίες." +
        (allErrors.length ? " Προσοχή: " + allErrors.length + " γραμμές δεν διαβάστηκαν." : "")
      );

      setViewMode("month");
      setTab("schedule");
    } catch (error) {
      setMessage("Σφάλμα στο Excel: " + error.message);
    } finally {
      setIsUploading(false);
    }
  };
  const resetDemo = () => { clearSavedState(); const fresh = initialState(); setActiveWorkspaceId(fresh.activeWorkspaceId); setTasksByWorkspace(fresh.tasksByWorkspace); setLinenStock(fresh.linenStock); setSelectedProperty("all"); setSelectedArea("all"); setViewMode("today"); setTab("schedule"); setMessage("Έγινε reset και στα 2 συστήματα."); };
  const logout = () => { setUser(null); setSelectedProperty("all"); setSelectedArea("all"); setViewMode("today"); setTab("schedule"); };

  if (!user) return <LoginScreen username={username} setUsername={setUsername} password={password} setPassword={setPassword} login={login} loginError={loginError} isOnline={isOnline} />;

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-6"><div className="max-w-6xl mx-auto space-y-5">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3"><div><h1 className="text-2xl md:text-3xl font-bold">SithonBreak Cleaning System</h1><div className="text-lg font-semibold text-slate-700">{user.role === "cleaner" ? user.name : workspace.title}</div><p className="text-slate-500">{workspace.subtitle}</p></div><div className="flex flex-wrap gap-2"><StatusBadge isOnline={isOnline} lastSaved={lastSaved} /><Button variant="outline" className="rounded-xl" onClick={logout}>🚪 Έξοδος</Button></div></div>

      {canSwitchWorkspaces(user) && <div className="grid grid-cols-1 md:grid-cols-2 gap-3"><WorkspaceCard active={workspace.id === "units"} title="4 Μονάδες Πελάτη" text="Με καθαρίστριες, σεντόνια, stock" onClick={() => switchWorkspace("units")} /><WorkspaceCard active={workspace.id === "sithon"} title="Τα δικά μου" text="Μία προβολή για όλα τα σπίτια που καθαρίζεις εσύ" onClick={() => switchWorkspace("sithon")} /></div>}

      <div className="flex gap-2 overflow-x-auto pb-1">{workspace.id === "sithon" && (user.role === "admin" || user.role === "owner") && <Button variant={tab === "today-work" ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setTab("today-work")}>Σήμερα δουλεύω</Button>}<Button variant={tab === "schedule" ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setTab("schedule")}>Πρόγραμμα</Button><Button variant={tab === "heavy" ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setTab("heavy")}>Βαριές μέρες</Button>{workspace.hasLinen && (user.role === "admin" || user.role === "client") && <Button variant={tab === "linen" ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setTab("linen")}>Stock σεντονιών</Button>}<Button variant={tab === "reports" ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setTab("reports")}>Reports</Button></div>

      {canUploadExcel(user) && <UploadPanel workspace={workspace} user={user} simulateExcelUpload={simulateExcelUpload} handleExcelFile={handleExcelFile} isUploading={isUploading} resetDemo={resetDemo} selectedUploadProperty={selectedUploadProperty} setSelectedUploadProperty={setSelectedUploadProperty} />}
      {message && <div className="rounded-2xl bg-green-50 p-4 text-sm font-medium text-green-800">✅ {message}</div>}

      {tab === "today-work" && workspace.id === "sithon" ? <TodayWorkView tasks={tasks} workspace={workspace} markDone={markDone} addNote={addNote} /> : tab === "linen" && workspace.hasLinen ? <LinenStockView workspace={workspace} linenStock={linenStock} washLinen={washLinen} /> : tab === "heavy" ? <HeavyDaysView heavyDays={heavyDays} workspace={workspace} /> : tab === "reports" ? <ReportsView reports={reports} tasks={tasks} workspace={workspace} /> : <ScheduleView workspace={workspace} user={user} areas={areas} selectedArea={selectedArea} setSelectedArea={setSelectedArea} selectedProperty={selectedProperty} setSelectedProperty={setSelectedProperty} viewMode={viewMode} setViewMode={setViewMode} stats={stats} displayedTasks={displayedTasks} groupedTasks={groupedTasks} markDone={markDone} addNote={addNote} />}
    </div></div>
  );
}

function LoginScreen({ username, setUsername, password, setPassword, login, loginError, isOnline }) { return <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4"><div className="w-full max-w-md"><Card className="rounded-2xl shadow-lg"><CardContent className="p-6 space-y-5"><div className="text-center space-y-2"><div className="mx-auto h-14 w-14 rounded-2xl bg-slate-900 text-white flex items-center justify-center text-2xl">🧹</div><h1 className="text-2xl font-bold">SithonBreak Cleaning System</h1><p className="text-sm text-slate-500">4 μονάδες πελάτη + δικά μου καταλύματα</p></div><div className="space-y-3"><input className="w-full rounded-xl border p-3" value={username} onChange={(e) => setUsername(e.target.value)} placeholder="Username" /><input className="w-full rounded-xl border p-3" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Password" type="password" onKeyDown={(e) => e.key === "Enter" && login()} />{loginError && <div className="rounded-xl bg-red-50 p-3 text-sm font-medium text-red-700">{loginError}</div>}<Button onClick={login} className="w-full rounded-xl py-6">Είσοδος</Button></div><div className="rounded-xl bg-slate-100 p-3 text-sm text-slate-600">
<div><b>Σύνδεση μόνο με προσωπικούς κωδικούς.</b></div>
<div className="text-xs mt-1">(Οι κωδικοί δεν εμφανίζονται για λόγους ασφάλειας)</div>
</div><div><b>Πελάτης 4 μονάδων:</b> client / client2026</div><div><b>Δικά σου:</b> sithon / sithon2026</div><div><b>Καθαρίστριες:</b> vrachos, inlofts, theros, okkio / αντίστοιχο 2026</div><div><b>Κατάσταση:</b> {isOnline ? "Online" : "Offline mode"}</div></div></CardContent></Card></div></div>; }

function UploadPanel({ workspace, user, simulateExcelUpload, handleExcelFile, isUploading, resetDemo, selectedUploadProperty, setSelectedUploadProperty }) {
  const canChooseUnitExcel = workspace.id === "units" && (user.role === "admin" || user.role === "client");
  const inputId = "excel-upload-" + workspace.id;
  return <Card className="rounded-2xl shadow-sm"><CardContent className="p-4 space-y-4"><div className="flex flex-col md:flex-row gap-3 md:items-center justify-between"><div className="flex items-center gap-3"><div className="h-11 w-11 rounded-xl bg-slate-100 flex items-center justify-center text-xl">📤</div><div><div className="font-semibold">Ανέβασμα Excel: {workspace.title}</div><div className="text-sm text-slate-500">{workspace.id === "units" ? "Για τον πελάτη μπορείς να περνάς ξεχωριστό Excel για κάθε μονάδα." : "Για τα δικά σου περνάς ένα Excel και βλέπεις όλα τα σπίτια μαζί."}</div></div></div><div className="flex flex-col sm:flex-row gap-2"><input id={inputId} type="file" accept=".xlsx,.xls" className="hidden" onChange={(event) => { handleExcelFile(event.target.files?.[0]); event.target.value = ""; }} /><Button className="rounded-xl" disabled={isUploading} onClick={() => document.getElementById(inputId)?.click()}>{isUploading ? "Διαβάζω Excel..." : "Ανέβασμα πραγματικού Excel"}</Button><Button variant="outline" className="rounded-xl" onClick={simulateExcelUpload}>Demo Excel</Button>{user.role === "admin" && <Button variant="outline" className="rounded-xl" onClick={resetDemo}>Reset</Button>}</div></div>{canChooseUnitExcel && <div className="space-y-2"><div className="text-sm font-medium text-slate-700">Ποιο Excel ανεβάζεις;</div><div className="flex gap-2 overflow-x-auto pb-1"><Button variant={selectedUploadProperty === "all" ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setSelectedUploadProperty("all")}>Όλες οι μονάδες</Button>{workspace.properties.map((p) => <Button key={p.id} variant={selectedUploadProperty === p.id ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setSelectedUploadProperty(p.id)}>{p.name}</Button>)}</div></div>}<div className="rounded-xl bg-slate-100 p-3 text-sm text-slate-600"><b>Στήλες που διαβάζει:</b> Room/Κατάλυμα, Check-in, Check-out. Για τις 4 μονάδες διάλεξε πρώτα τη μονάδα ή βάλε στήλη Μονάδα στο Excel.</div></CardContent></Card>;
}
function WorkspaceCard({ active, title, text, onClick }) { return <button onClick={onClick} className={`text-left rounded-2xl border p-4 shadow-sm ${active ? "bg-slate-900 text-white" : "bg-white"}`}><div className="text-xl font-bold">{title}</div><div className={`text-sm ${active ? "text-slate-200" : "text-slate-500"}`}>{text}</div></button>; }
function StatusBadge({ isOnline, lastSaved }) { return <div className={`rounded-xl border px-3 py-2 text-sm ${isOnline ? "bg-green-50 text-green-800 border-green-100" : "bg-orange-50 text-orange-800 border-orange-100"}`}><div className="font-semibold">{isOnline ? "Online" : "Offline"}</div><div className="text-xs opacity-80">Σταθερή λειτουργία χωρίς βαρύ auto-save</div></div>; }

function ScheduleView({ workspace, user, areas, selectedArea, setSelectedArea, selectedProperty, setSelectedProperty, viewMode, setViewMode, stats, displayedTasks, groupedTasks, markDone, addNote }) { const canFilter = user.role === "admin" || user.role === "client" || user.role === "owner"; return <><div className="flex gap-2 overflow-x-auto pb-1"><Button variant={viewMode === "today" ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setViewMode("today")}>Σήμερα</Button><Button variant={viewMode === "month" ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setViewMode("month")}>Όλος ο μήνας</Button></div><div className="grid grid-cols-1 md:grid-cols-5 gap-3"><Stat icon="📅" title="Εργασίες" value={stats.total} /><Stat icon="⚠️" title="Ίδια μέρα" value={stats.urgent} />{workspace.hasLinen && <Stat icon="🧺" title="Σεντόνια" value={stats.linen} />}<Stat icon="✅" title="Ολοκληρωμένα" value={stats.done} /><Stat icon="⏳" title="Εκκρεμούν" value={stats.pending} /></div>{(user.role === "admin" || user.role === "owner") && workspace.id === "sithon" && <div className="rounded-2xl bg-blue-50 p-4 text-sm text-blue-800">📌 Ενιαία προβολή: εδώ βλέπεις όλα τα σπίτια μαζί, από όλες τις περιοχές, για να ξέρεις κάθε μέρα ποια καθαρίζεις.</div>}{canFilter && <><div className="flex gap-2 overflow-x-auto pb-1"><Button variant={selectedArea === "all" ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setSelectedArea("all")}>Όλες οι περιοχές</Button>{areas.map((area) => <Button key={area} variant={selectedArea === area ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setSelectedArea(area)}>{area}</Button>)}</div><div className="flex gap-2 overflow-x-auto pb-1"><Button variant={selectedProperty === "all" ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setSelectedProperty("all")}>Όλα</Button>{workspace.properties.map((p) => <Button key={p.id} variant={selectedProperty === p.id ? "default" : "outline"} className="rounded-xl whitespace-nowrap" onClick={() => setSelectedProperty(p.id)}>{p.name}</Button>)}</div></>}{displayedTasks.length === 0 ? <Card className="rounded-2xl shadow-sm"><CardContent className="p-6 text-center text-slate-500">Δεν υπάρχουν εργασίες.</CardContent></Card> : viewMode === "month" ? <MonthView groupedTasks={groupedTasks} markDone={markDone} addNote={addNote} /> : <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">{displayedTasks.map((task) => <TaskCard key={task.id} task={task} markDone={markDone} addNote={addNote} />)}</div>}</>; }

function TodayWorkView({ tasks, workspace, markDone, addNote }) {
  const todayTasks = sortTodayRoute(getTodayTasks(tasks));
  const heavy = isHeavyToday(tasks, workspace);
  const areaGroups = groupBy(todayTasks, "area");

  if (!todayTasks.length) {
    return <Card className="rounded-2xl shadow-sm"><CardContent className="p-6 text-center text-slate-500"><div className="text-3xl mb-2">🎉</div><div className="font-semibold">Σήμερα δεν έχεις καθαριότητες</div><div className="text-sm">Δεν υπάρχει check-out για σήμερα.</div></CardContent></Card>;
  }

  return <div className="space-y-4">
    <Card className={`rounded-2xl shadow-sm ${heavy ? "border-orange-200 bg-orange-50" : "bg-white"}`}><CardContent className="p-5 space-y-2"><div className="flex flex-col md:flex-row md:items-center justify-between gap-3"><div><div className="text-2xl font-bold">Σήμερα δουλεύω</div><div className="text-slate-600">Έχεις {todayTasks.length} σπίτια για καθαριότητα.</div></div>{heavy && <div className="rounded-full bg-orange-100 px-4 py-2 text-sm font-bold text-orange-800">⚠️ Βαριά μέρα – πάρε άτομο</div>}</div><div className="text-sm text-slate-500">Αυτόματη σειρά: πρώτα ίδια μέρα άφιξη, μετά ανά περιοχή.</div></CardContent></Card>

    {Object.entries(areaGroups).map(([area, areaTasks]) => <Card key={area} className="rounded-2xl shadow-sm overflow-hidden"><CardContent className="p-0"><div className="p-4 bg-white border-b"><div className="text-xl font-bold">{area}</div><div className="text-sm text-slate-500">{areaTasks.length} καθαριότητες</div></div><div className="divide-y">{areaTasks.map((task, index) => <div key={task.id} className="p-4 flex flex-col md:flex-row md:items-center justify-between gap-3"><div className="flex items-start gap-3"><div className="h-10 w-10 rounded-xl bg-slate-100 flex items-center justify-center font-bold">{index + 1}</div><div><div className="font-bold text-lg">{task.propertyName}</div>{task.guestName && <div className="text-sm text-slate-500">👤 {task.guestName}</div>}<div className="text-sm text-slate-600">{task.type}</div>{task.urgent && <div className="mt-1 inline-block rounded-full bg-red-100 px-3 py-1 text-xs font-semibold text-red-700">Ίδια μέρα άφιξη</div>}{task.note && <div className="text-sm text-slate-500 mt-1">💬 {task.note}</div>}</div></div><TaskActions task={task} markDone={markDone} addNote={addNote} /></div>)}</div></CardContent></Card>)}
  </div>;
}

function MonthView({ groupedTasks, markDone, addNote }) { const dates = Object.keys(groupedTasks).sort(); return <div className="space-y-4">{dates.map((date) => { const dayTasks = groupedTasks[date]; const urgentCount = dayTasks.filter((t) => t.urgent && t.status !== "done").length; const pendingCount = dayTasks.filter((t) => t.status !== "done").length; return <Card key={date} className="rounded-2xl shadow-sm overflow-hidden"><CardContent className="p-0"><div className="p-4 bg-white border-b flex flex-col md:flex-row md:items-center justify-between gap-2"><div><div className="text-xl font-bold">{formatDate(date)}</div><div className="text-sm text-slate-500">{dayTasks.length} εργασίες • {pendingCount} εκκρεμούν</div></div>{urgentCount > 0 && <div className="rounded-full bg-red-100 px-3 py-1 text-sm font-semibold text-red-700">⚠️ {urgentCount} ίδια μέρα</div>}</div><div className="divide-y">{dayTasks.map((task) => <div key={task.id} className="p-4 flex flex-col md:flex-row md:items-center justify-between gap-3"><div><div className="font-semibold">{task.propertyName} {task.room && `• ${task.room}`}</div>{task.guestName && <div className="text-sm text-slate-500">👤 {task.guestName}</div>}<div className="text-sm text-slate-600">{task.area} • {task.type}</div>{task.note && <div className="text-sm text-slate-500 mt-1">💬 {task.note}</div>}</div><TaskActions task={task} markDone={markDone} addNote={addNote} /></div>)}</div></CardContent></Card>; })}</div>; }
function TaskCard({ task, markDone, addNote }) { return <Card className={`rounded-2xl shadow-sm ${task.status === "done" ? "opacity-70" : ""}`}><CardContent className="p-4 space-y-4"><div className="flex justify-between gap-3"><div><div className="text-sm text-slate-500">{task.area} • {formatDate(task.date)}</div><div className="text-2xl font-bold">{task.propertyName}</div>{task.room && <div className="font-semibold text-slate-700">{task.room}</div>}{task.guestName && <div className="text-sm text-slate-500 mt-1">👤 {task.guestName}</div>}</div><div>{task.urgent && task.status !== "done" && <span className="rounded-full bg-red-100 px-3 py-1 text-sm font-semibold text-red-700">Ίδια μέρα</span>}{task.status === "done" && <span className="rounded-full bg-green-100 px-3 py-1 text-sm font-semibold text-green-700">Έγινε</span>}</div></div><div className="rounded-xl bg-slate-100 p-3">{task.source === "auto-linen" ? "🧺" : "🧹"} {task.type}</div>{task.note && <div className="rounded-xl border p-3 text-sm">💬 {task.note}</div>}<TaskActions task={task} markDone={markDone} addNote={addNote} /></CardContent></Card>; }
function TaskActions({ task, markDone, addNote }) { const writeNote = () => { const text = window.prompt("Γράψε σημείωση για αυτή την καθαριότητα:", task.note || ""); if (text !== null) addNote(task.id, text.trim()); }; return <div className="flex flex-wrap gap-2 md:justify-end"><Button size="sm" className="rounded-xl" disabled={task.status === "done"} onClick={() => markDone(task.id)}>Ολοκληρώθηκε</Button><Button size="sm" variant="outline" className="rounded-xl" onClick={writeNote}>Γράψε σημείωση</Button></div>; }

function HeavyDaysView({ heavyDays, workspace }) { if (!heavyDays.length) return <Card className="rounded-2xl shadow-sm"><CardContent className="p-6 text-center text-slate-500">Δεν υπάρχουν βαριές μέρες.</CardContent></Card>; return <div className="space-y-4">{heavyDays.map((day) => <Card key={day.date} className="rounded-2xl shadow-sm"><CardContent className="p-5 space-y-3"><div className="flex items-center justify-between gap-3"><div><div className="text-xl font-bold">{formatDate(day.date)}</div><div className="text-sm text-slate-500">Έχει {day.total} εργασίες. {workspace.id === "sithon" ? "Πιθανόν να χρειαστείς άτομο για βοήθεια." : "Πιθανόν χρειάζεται ενίσχυση."}</div></div><div className="rounded-full bg-orange-100 px-3 py-1 text-sm font-semibold text-orange-700">Βαριά μέρα</div></div>{Object.entries(day.areaGroups).map(([area, tasks]) => <div key={area} className="rounded-xl bg-slate-100 p-3"><div className="font-semibold">{area} • {tasks.length}</div><div className="text-sm text-slate-600 mt-1">{tasks.map((t) => `${t.propertyName}${t.room ? ` ${t.room}` : ""}`).join(", ")}</div></div>)}</CardContent></Card>)}</div>; }
function LinenStockView({ workspace, linenStock, washLinen }) { return <div className="grid grid-cols-1 md:grid-cols-2 gap-4">{workspace.properties.map((p) => { const stock = linenStock[p.id]; const status = stockStatus(stock); return <Card key={p.id} className="rounded-2xl shadow-sm"><CardContent className="p-5 space-y-4"><div className="flex justify-between"><div><div className="text-xl font-bold">{p.name}</div><div className="text-sm text-slate-500">Όριο: {stock?.minimum}</div></div><div className={`rounded-full px-3 py-1 text-sm font-semibold ${status === "low" ? "bg-red-100 text-red-700" : status === "warning" ? "bg-yellow-100 text-yellow-700" : "bg-green-100 text-green-700"}`}>{status === "low" ? "Χαμηλό" : status === "warning" ? "Προσοχή" : "ΟΚ"}</div></div><div className="grid grid-cols-2 gap-3"><div className="rounded-xl bg-slate-100 p-4"><div className="text-sm text-slate-500">Καθαρά</div><div className="text-3xl font-bold">{stock?.cleanSets}</div></div><div className="rounded-xl bg-slate-100 p-4"><div className="text-sm text-slate-500">Για πλύσιμο</div><div className="text-3xl font-bold">{stock?.usedSets}</div></div></div><Button variant="outline" className="rounded-xl w-full" onClick={() => washLinen(p.id)}>Καταχώρηση πλυσίματος</Button></CardContent></Card>; })}</div>; }
function ReportsView({ reports, tasks, workspace }) { const totals = getStats(tasks); const heavyDays = getHeavyDays(tasks, workspace); return <div className="space-y-4"><div className="grid grid-cols-1 md:grid-cols-4 gap-3"><Stat icon="📊" title="Σύνολο" value={totals.total} /><Stat icon="✅" title="Έγιναν" value={totals.done} /><Stat icon="⚠️" title="Ίδια μέρα" value={totals.urgent} /><Stat icon="🔥" title="Βαριές" value={heavyDays.length} /></div><Card className="rounded-2xl shadow-sm overflow-hidden"><CardContent className="p-0"><div className="p-4 border-b bg-white"><div className="text-xl font-bold">Αναφορά</div></div><div className="overflow-x-auto"><table className="w-full text-sm"><thead className="bg-slate-100 text-slate-600"><tr><th className="text-left p-3">Όνομα</th><th className="text-left p-3">Περιοχή</th><th className="text-center p-3">Εργασίες</th>{workspace.hasLinen && <th className="text-center p-3">Σεντόνια</th>}<th className="text-center p-3">Ίδια μέρα</th><th className="text-center p-3">%</th></tr></thead><tbody className="divide-y">{reports.map((r) => <tr key={r.propertyId}><td className="p-3 font-semibold">{r.propertyName}</td><td className="p-3">{r.area}</td><td className="p-3 text-center">{r.total}</td>{workspace.hasLinen && <td className="p-3 text-center">{r.linen}</td>}<td className="p-3 text-center">{r.urgent}</td><td className="p-3 text-center font-semibold">{r.completion}%</td></tr>)}</tbody></table></div></CardContent></Card></div>; }
function Stat({ icon, title, value }) { return <Card className="rounded-2xl shadow-sm"><CardContent className="p-4 flex items-center gap-3"><div className="h-11 w-11 rounded-xl bg-slate-100 flex items-center justify-center text-xl">{icon}</div><div><div className="text-sm text-slate-500">{title}</div><div className="text-2xl font-bold">{value}</div></div></CardContent></Card>; }

