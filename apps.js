import { useState, useMemo, useEffect, useCallback } from "react";

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
// Replace these values with your own before deploying
const CONFIG = {
  clientId: "46cea36b-1063-4567-9796-3c0f3c5a24df", // Azure App Registration client ID
  tenantId: "bf266c07-4cc7-4ba5-97f1-99a4a77b5acd", // Your Azure AD tenant ID
  sharepointSite: "tsortho.sharepoint.com", // e.g. contoso.sharepoint.com
  siteRelPath: "/sites/allcompany", // SharePoint site path
  listName: "InOutBoard", // Exact name of your SharePoint list
  managersGroupId: "18dd5e6f-fab9-4a57-a703-7eeca76aad79", // Azure AD security group object ID
};
// ─────────────────────────────────────────────────────────────────────────────

const STATUSES = [
  { label: "In", color: "#27500A", bg: "#EAF3DE", dot: "#639922" },
  { label: "Out", color: "#791F1F", bg: "#FCEBEB", dot: "#E24B4A" },
  { label: "Lunch", color: "#633806", bg: "#FAEEDA", dot: "#EF9F27" },
  { label: "In a Meeting", color: "#0C447C", bg: "#E6F1FB", dot: "#378ADD" },
  {
    label: "Working Remotely",
    color: "#3C3489",
    bg: "#EEEDFE",
    dot: "#7F77DD",
  },
  { label: "On Vacation", color: "#085041", bg: "#E1F5EE", dot: "#1D9E75" },
  { label: "Nurses Station", color: "#72243E", bg: "#FBEAF0", dot: "#D4537E" },
  { label: "Custom Remark", color: "#444441", bg: "#F1EFE8", dot: "#888780" },
];

const AVATAR_COLORS = [
  ["#E6F1FB", "#0C447C"],
  ["#E1F5EE", "#085041"],
  ["#FAEEDA", "#633806"],
  ["#EEEDFE", "#3C3489"],
  ["#FBEAF0", "#72243E"],
  ["#EAF3DE", "#27500A"],
  ["#FCEBEB", "#791F1F"],
  ["#FAE8DC", "#712B13"],
];

// ─── DEMO DATA (used when not connected to SharePoint) ───────────────────────
const NOW = Date.now();
const DEMO_EMPLOYEES = [
  {
    id: "1",
    name: "Sarah Mitchell",
    dept: "Nursing",
    ext: "2210",
    email: "s.mitchell@org.local",
    status: "In",
    remark: "",
    isManager: true,
    modifiedAt: NOW - 1000 * 60 * 8,
    spItemId: null,
  },
  {
    id: "2",
    name: "James Okafor",
    dept: "Nursing",
    ext: "2211",
    email: "j.okafor@org.local",
    status: "Nurses Station",
    remark: "",
    isManager: false,
    modifiedAt: NOW - 1000 * 60 * 32,
    spItemId: null,
  },
  {
    id: "3",
    name: "Linda Vance",
    dept: "Administration",
    ext: "2100",
    email: "l.vance@org.local",
    status: "Out",
    remark: "Back Thursday",
    isManager: true,
    modifiedAt: NOW - 1000 * 60 * 60 * 2,
    spItemId: null,
  },
  {
    id: "4",
    name: "Tom Bergstrom",
    dept: "Administration",
    ext: "2101",
    email: "t.bergstrom@org.local",
    status: "In",
    remark: "",
    isManager: false,
    modifiedAt: NOW - 1000 * 60 * 60 * 5,
    spItemId: null,
  },
  {
    id: "5",
    name: "Dr. Priya Sharma",
    dept: "Physicians",
    ext: "2300",
    email: "p.sharma@org.local",
    status: "In a Meeting",
    remark: "",
    isManager: false,
    modifiedAt: NOW - 1000 * 60 * 15,
    spItemId: null,
  },
  {
    id: "6",
    name: "Carlos Reyes",
    dept: "Physicians",
    ext: "2301",
    email: "c.reyes@org.local",
    status: "Working Remotely",
    remark: "Available by phone",
    isManager: false,
    modifiedAt: NOW - 1000 * 60 * 60 * 24,
    spItemId: null,
  },
  {
    id: "7",
    name: "Angela Wu",
    dept: "Radiology",
    ext: "2400",
    email: "a.wu@org.local",
    status: "In",
    remark: "",
    isManager: true,
    modifiedAt: NOW - 1000 * 60 * 3,
    spItemId: null,
  },
  {
    id: "8",
    name: "Derek Hall",
    dept: "Radiology",
    ext: "2401",
    email: "d.hall@org.local",
    status: "Lunch",
    remark: "",
    isManager: false,
    modifiedAt: NOW - 1000 * 60 * 45,
    spItemId: null,
  },
  {
    id: "9",
    name: "Fiona Marsh",
    dept: "HR",
    ext: "2050",
    email: "f.marsh@org.local",
    status: "On Vacation",
    remark: "Returns Mon",
    isManager: false,
    modifiedAt: NOW - 1000 * 60 * 60 * 48,
    spItemId: null,
  },
  {
    id: "10",
    name: "Kevin Park",
    dept: "IT",
    ext: "2001",
    email: "k.park@org.local",
    status: "In",
    remark: "",
    isManager: true,
    modifiedAt: NOW - 1000 * 60 * 20,
    spItemId: null,
  },
  {
    id: "11",
    name: "Tanya Ellis",
    dept: "IT",
    ext: "2002",
    email: "t.ellis@org.local",
    status: "In",
    remark: "",
    isManager: false,
    modifiedAt: NOW - 1000 * 60 * 60,
    spItemId: null,
  },
  {
    id: "12",
    name: "Samuel Grant",
    dept: "Lab",
    ext: "2500",
    email: "s.grant@org.local",
    status: "Out",
    remark: "",
    isManager: false,
    modifiedAt: NOW - 1000 * 60 * 60 * 3,
    spItemId: null,
  },
];

// ─── MSAL / GRAPH HELPERS ────────────────────────────────────────────────────
// These functions only run when MSAL is loaded (real deployment).
// In demo mode the app skips all of this.

async function getMsalToken(msal) {
  const accounts = msal.getAllAccounts();
  if (!accounts.length)
    await msal.loginPopup({
      scopes: ["User.Read", "Sites.ReadWrite.All", "GroupMember.Read.All"],
    });
  const result = await msal.acquireTokenSilent({
    scopes: ["User.Read", "Sites.ReadWrite.All", "GroupMember.Read.All"],
    account: msal.getAllAccounts()[0],
  });
  return result.accessToken;
}

async function graphGet(token, url) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  return r.json();
}

async function getSignedInUser(token) {
  const me = await graphGet(
    token,
    "https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,department,businessPhones"
  );
  const groups = await graphGet(
    token,
    `https://graph.microsoft.com/v1.0/me/memberOf?$select=id`
  );
  const isManager = (groups.value || []).some(
    (g) => g.id === CONFIG.managersGroupId
  );
  return {
    azureId: me.id,
    name: me.displayName,
    email: me.mail,
    dept: me.department || "",
    ext: (me.businessPhones || [])[0] || "",
    isManager,
  };
}

async function getSharePointEmployees(token) {
  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${CONFIG.sharepointSite}:${CONFIG.siteRelPath}`;
  const siteData = await graphGet(token, siteUrl);
  const siteId = siteData.id;

  const listsUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$filter=displayName eq '${CONFIG.listName}'`;
  const listsData = await graphGet(token, listsUrl);
  const listId = listsData.value[0].id;

  const itemsUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=999`;
  const itemsData = await graphGet(token, itemsUrl);

  return {
    siteId,
    listId,
    employees: itemsData.value.map((item) => ({
      id: item.fields.AzureObjectId || item.id,
      spItemId: item.id,
      name: item.fields.Title || "",
      dept: item.fields.Department || "",
      ext: item.fields.Extension || "",
      email: item.fields.Email || "",
      status: item.fields.Status || "In",
      remark: item.fields.Remark || "",
      isManager: item.fields.IsManager === true,
      modifiedAt: new Date(item.lastModifiedDateTime).getTime(),
    })),
  };
}

async function updateSharePointItem(token, siteId, listId, spItemId, fields) {
  await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${spItemId}/fields`,
    {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        Status: fields.status,
        Remark: fields.remark,
      }),
    }
  );
}

// ─── UTILITIES ───────────────────────────────────────────────────────────────
function formatDateTime(ts) {
  if (!ts) return "—";
  const d = new Date(ts);
  return (
    d.toLocaleDateString(undefined, {
      month: "short",
      day: "numeric",
      year: "numeric",
    }) +
    " " +
    d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" })
  );
}
function initials(name) {
  return name
    .split(" ")
    .map((p) => p[0])
    .join("")
    .slice(0, 2)
    .toUpperCase();
}
function statusInfo(label) {
  return (
    STATUSES.find((s) => s.label === label) || STATUSES[STATUSES.length - 1]
  );
}

// ─── EDIT MODAL ──────────────────────────────────────────────────────────────
function EditModal({ employee, onSave, onClose, saving }) {
  const [status, setStatus] = useState(employee.status);
  const [remark, setRemark] = useState(employee.remark);
  return (
    <div
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(0,0,0,0.45)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 200,
      }}
    >
      <div
        style={{
          background: "#fff",
          border: "0.5px solid #d0d0d0",
          borderRadius: 14,
          padding: "1.5rem",
          width: "100%",
          maxWidth: 400,
          margin: "0 1rem",
          boxShadow: "0 8px 32px rgba(0,0,0,0.12)",
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            marginBottom: "1.25rem",
          }}
        >
          <span style={{ fontSize: 20 }}>✏️</span>
          <h3 style={{ fontWeight: 500, fontSize: 16, margin: 0 }}>
            Update status — {employee.name}
          </h3>
        </div>
        <label
          style={{
            fontSize: 13,
            color: "#666",
            display: "block",
            marginBottom: 4,
          }}
        >
          Status
        </label>
        <select
          value={status}
          onChange={(e) => setStatus(e.target.value)}
          style={{
            width: "100%",
            padding: "8px 10px",
            borderRadius: 8,
            border: "0.5px solid #ccc",
            fontSize: 14,
            marginBottom: 12,
            background: "#fff",
          }}
        >
          {STATUSES.map((s) => (
            <option key={s.label} value={s.label}>
              {s.label}
            </option>
          ))}
        </select>
        <label
          style={{
            fontSize: 13,
            color: "#666",
            display: "block",
            marginBottom: 4,
          }}
        >
          {status === "Custom Remark"
            ? "Custom remark"
            : "Note / remark (optional)"}
        </label>
        <input
          type="text"
          value={remark}
          onChange={(e) => setRemark(e.target.value)}
          maxLength={80}
          placeholder={
            status === "Custom Remark"
              ? "Enter your custom note…"
              : "e.g. Back at 3pm…"
          }
          style={{
            width: "100%",
            padding: "8px 10px",
            borderRadius: 8,
            border: "0.5px solid #ccc",
            fontSize: 14,
            background: "#fff",
          }}
        />
        <div
          style={{
            display: "flex",
            gap: 10,
            marginTop: "1.25rem",
            justifyContent: "flex-end",
          }}
        >
          <button
            onClick={onClose}
            style={{
              padding: "8px 16px",
              borderRadius: 8,
              border: "0.5px solid #ccc",
              background: "none",
              fontSize: 14,
              cursor: "pointer",
              color: "#666",
            }}
          >
            Cancel
          </button>
          <button
            onClick={() =>
              onSave({ ...employee, status, remark, modifiedAt: Date.now() })
            }
            disabled={saving}
            style={{
              padding: "8px 20px",
              borderRadius: 8,
              border: "none",
              background: saving ? "#aac" : "#185FA5",
              color: "#fff",
              fontSize: 14,
              fontWeight: 500,
              cursor: saving ? "not-allowed" : "pointer",
            }}
          >
            {saving ? "Saving…" : "Save"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── MAIN COMPONENT ──────────────────────────────────────────────────────────
export default function InOutBoard() {
  const [mode, setMode] = useState("loading"); // "loading" | "demo" | "live"
  const [employees, setEmployees] = useState([]);
  const [viewer, setViewer] = useState(null);
  const [demoViewerId, setDemoViewerId] = useState("1");
  const [editingId, setEditingId] = useState(null);
  const [saving, setSaving] = useState(false);
  const [search, setSearch] = useState("");
  const [deptFilter, setDeptFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [sortField, setSortField] = useState("name");
  const [sortDir, setSortDir] = useState("asc");
  const [error, setError] = useState(null);
  const [spContext, setSpContext] = useState(null); // { token, siteId, listId }

  // On mount: try MSAL, fall back to demo
  useEffect(() => {
    if (CONFIG.clientId === "YOUR_APP_CLIENT_ID") {
      // No config set — run in demo mode
      setEmployees(DEMO_EMPLOYEES);
      setViewer(DEMO_EMPLOYEES[0]);
      setMode("demo");
      return;
    }
    // Try to load MSAL and sign in
    const script = document.createElement("script");
    script.src =
      "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.10.0/lib/msal-browser.min.js";
    script.onload = async () => {
      try {
        const msal = new window.msal.PublicClientApplication({
          auth: {
            clientId: CONFIG.clientId,
            tenantId: CONFIG.tenantId,
            redirectUri: window.location.href,
          },
        });
        await msal.initialize();
        const token = await getMsalToken(msal);
        const me = await getSignedInUser(token);
        const {
          siteId,
          listId,
          employees: spEmployees,
        } = await getSharePointEmployees(token);
        setSpContext({ token, siteId, listId, msal });
        setEmployees(spEmployees);
        const myRow =
          spEmployees.find(
            (e) => e.email?.toLowerCase() === me.email?.toLowerCase()
          ) || spEmployees[0];
        setViewer({ ...myRow, isManager: me.isManager });
        setMode("live");
      } catch (err) {
        setError(
          "Could not connect to Microsoft 365. Running in demo mode. (" +
            err.message +
            ")"
        );
        setEmployees(DEMO_EMPLOYEES);
        setViewer(DEMO_EMPLOYEES[0]);
        setMode("demo");
      }
    };
    script.onerror = () => {
      setEmployees(DEMO_EMPLOYEES);
      setViewer(DEMO_EMPLOYEES[0]);
      setMode("demo");
    };
    document.head.appendChild(script);
  }, []);

  const demoViewer = useMemo(
    () =>
      mode === "demo"
        ? employees.find((e) => e.id === demoViewerId) || employees[0]
        : null,
    [mode, employees, demoViewerId]
  );

  const activeViewer = mode === "live" ? viewer : demoViewer;
  const canEdit = useCallback(
    (empId) =>
      activeViewer && (activeViewer.isManager || activeViewer.id === empId),
    [activeViewer]
  );
  const depts = useMemo(
    () => [...new Set(employees.map((e) => e.dept))].sort(),
    [employees]
  );

  function toggleSort(field) {
    if (sortField === field) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else {
      setSortField(field);
      setSortDir("asc");
    }
  }

  const filtered = useMemo(() => {
    const q = search.toLowerCase();
    let rows = employees.filter((e) => {
      if (deptFilter && e.dept !== deptFilter) return false;
      if (statusFilter && e.status !== statusFilter) return false;
      if (
        q &&
        ![e.name, e.dept, e.email, e.status, e.remark].some((x) =>
          (x || "").toLowerCase().includes(q)
        )
      )
        return false;
      return true;
    });
    return [...rows].sort((a, b) => {
      let av = a[sortField],
        bv = b[sortField];
      if (sortField === "modifiedAt")
        return sortDir === "asc"
          ? (av || 0) - (bv || 0)
          : (bv || 0) - (av || 0);
      return sortDir === "asc"
        ? String(av || "").localeCompare(String(bv || ""))
        : String(bv || "").localeCompare(String(av || ""));
    });
  }, [employees, search, deptFilter, statusFilter, sortField, sortDir]);

  const stats = useMemo(() => {
    const present = employees.filter((e) =>
      ["In", "Nurses Station", "In a Meeting", "Working Remotely"].includes(
        e.status
      )
    ).length;
    const away = employees.filter((e) =>
      ["Out", "Lunch", "On Vacation"].includes(e.status)
    ).length;
    return { total: employees.length, present, away, depts: depts.length };
  }, [employees, depts]);

  async function handleSave(updated) {
    setSaving(true);
    try {
      if (mode === "live" && spContext && updated.spItemId) {
        await updateSharePointItem(
          spContext.token,
          spContext.siteId,
          spContext.listId,
          updated.spItemId,
          updated
        );
      }
      setEmployees((prev) =>
        prev.map((e) => (e.id === updated.id ? updated : e))
      );
      if (mode === "live" && viewer?.id === updated.id) setViewer(updated);
    } catch (err) {
      setError("Save failed: " + err.message);
    }
    setSaving(false);
    setEditingId(null);
  }

  const editingEmployee = editingId
    ? employees.find((e) => e.id === editingId)
    : null;
  const inputStyle = {
    padding: "7px 12px",
    border: "0.5px solid #ccc",
    borderRadius: 8,
    fontSize: 14,
    background: "#fff",
    color: "inherit",
  };

  function SortTh({ field, label, style: extraStyle }) {
    const active = sortField === field;
    return (
      <th
        onClick={() => toggleSort(field)}
        style={{
          fontSize: 11,
          fontWeight: 500,
          color: active ? "#185FA5" : "#999",
          textAlign: "left",
          padding: "10px 12px",
          textTransform: "uppercase",
          letterSpacing: ".05em",
          cursor: "pointer",
          userSelect: "none",
          whiteSpace: "nowrap",
          ...extraStyle,
        }}
      >
        {label}{" "}
        {active ? (
          sortDir === "asc" ? (
            "↑"
          ) : (
            "↓"
          )
        ) : (
          <span style={{ opacity: 0.3 }}>↕</span>
        )}
      </th>
    );
  }

  if (mode === "loading") {
    return (
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          height: "60vh",
          color: "#999",
          fontSize: 14,
        }}
      >
        Connecting to Microsoft 365…
      </div>
    );
  }

  return (
    <div
      style={{
        fontFamily: "system-ui,sans-serif",
        padding: "1.5rem",
        maxWidth: "100%",
        minHeight: "100vh",
      }}
    >
      {/* Header */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 12,
          marginBottom: "1.5rem",
        }}
      >
        <div
          style={{
            width: 40,
            height: 40,
            borderRadius: 10,
            background: "#185FA5",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            color: "#fff",
            fontSize: 20,
          }}
        >
          📋
        </div>
        <div>
          <h1 style={{ fontSize: 20, fontWeight: 500, margin: 0 }}>
            Staff In/Out Board
          </h1>
          <p style={{ fontSize: 13, color: "#888", margin: 0 }}>
            Live presence for your organization
          </p>
        </div>
        <div
          style={{
            marginLeft: "auto",
            display: "flex",
            alignItems: "center",
            gap: 8,
          }}
        >
          {mode === "live" ? (
            <span
              style={{
                background: "#EAF3DE",
                color: "#27500A",
                fontSize: 11,
                padding: "3px 10px",
                borderRadius: 20,
                fontWeight: 500,
              }}
            >
              🟢 Connected to SharePoint
            </span>
          ) : (
            <span
              style={{
                background: "#FAEEDA",
                color: "#633806",
                fontSize: 11,
                padding: "3px 10px",
                borderRadius: 20,
                fontWeight: 500,
              }}
            >
              🟡 Demo mode
            </span>
          )}
        </div>
      </div>

      {/* Error banner */}
      {error && (
        <div
          style={{
            background: "#FCEBEB",
            color: "#791F1F",
            padding: "10px 14px",
            borderRadius: 8,
            fontSize: 13,
            marginBottom: "1rem",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <span>⚠️ {error}</span>
          <button
            onClick={() => setError(null)}
            style={{
              background: "none",
              border: "none",
              cursor: "pointer",
              color: "#791F1F",
              fontSize: 16,
            }}
          >
            ×
          </button>
        </div>
      )}

      {/* Viewer bar */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          padding: "10px 14px",
          background: "#f5f5f3",
          border: "0.5px solid #e0e0e0",
          borderRadius: 10,
          marginBottom: "1.25rem",
          flexWrap: "wrap",
        }}
      >
        <span style={{ fontSize: 13, color: "#666" }}>
          👤 {mode === "live" ? "Signed in as:" : "Viewing as (demo):"}
        </span>
        {mode === "live" ? (
          <strong style={{ fontSize: 14 }}>{viewer?.name}</strong>
        ) : (
          <select
            value={demoViewerId}
            onChange={(e) => setDemoViewerId(e.target.value)}
            style={{ ...inputStyle, fontSize: 13, padding: "4px 10px" }}
          >
            {employees.map((e) => (
              <option key={e.id} value={e.id}>
                {e.name}
                {e.isManager ? " (Manager)" : ""}
              </option>
            ))}
          </select>
        )}
        {activeViewer?.isManager ? (
          <span
            style={{
              background: "#E6F1FB",
              color: "#185FA5",
              fontSize: 11,
              padding: "3px 10px",
              borderRadius: 20,
              fontWeight: 500,
            }}
          >
            🛡 Manager — can edit all
          </span>
        ) : (
          <span style={{ fontSize: 12, color: "#999" }}>
            Can edit own row only
          </span>
        )}
        {mode === "demo" && (
          <span style={{ marginLeft: "auto", fontSize: 11, color: "#bbb" }}>
            Configure CONFIG block to connect to your Microsoft 365
          </span>
        )}
      </div>

      {/* Stats */}
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit,minmax(100px,1fr))",
          gap: 10,
          marginBottom: "1.5rem",
        }}
      >
        {[
          { v: stats.total, l: "Total staff", c: "inherit" },
          { v: stats.present, l: "Present/available", c: "#27500A" },
          { v: stats.away, l: "Out / away", c: "#791F1F" },
          { v: stats.depts, l: "Departments", c: "inherit" },
        ].map((s) => (
          <div
            key={s.l}
            style={{
              background: "#f7f7f5",
              borderRadius: 10,
              padding: "12px 14px",
            }}
          >
            <div style={{ fontSize: 24, fontWeight: 500, color: s.c }}>
              {s.v}
            </div>
            <div style={{ fontSize: 12, color: "#888", marginTop: 2 }}>
              {s.l}
            </div>
          </div>
        ))}
      </div>

      {/* Filters */}
      <div
        style={{
          display: "flex",
          gap: 10,
          marginBottom: "1.25rem",
          flexWrap: "wrap",
        }}
      >
        <input
          type="text"
          placeholder="Search name, department, email…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{ ...inputStyle, flex: 1, minWidth: 180, maxWidth: 320 }}
          aria-label="Search"
        />
        <select
          value={deptFilter}
          onChange={(e) => setDeptFilter(e.target.value)}
          style={inputStyle}
          aria-label="Filter by department"
        >
          <option value="">All departments</option>
          {depts.map((d) => (
            <option key={d} value={d}>
              {d}
            </option>
          ))}
        </select>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
          style={inputStyle}
          aria-label="Filter by status"
        >
          <option value="">All statuses</option>
          {STATUSES.map((s) => (
            <option key={s.label} value={s.label}>
              {s.label}
            </option>
          ))}
        </select>
      </div>

      {/* Table */}
      <div
        style={{
          overflowX: "auto",
          border: "0.5px solid #e0e0e0",
          borderRadius: 12,
        }}
      >
        <table
          style={{
            width: "100%",
            borderCollapse: "collapse",
            tableLayout: "fixed",
            minWidth: 900,
          }}
        >
          <colgroup>
            <col style={{ width: "17%" }} />
            <col style={{ width: "13%" }} />
            <col style={{ width: "6%" }} />
            <col style={{ width: "17%" }} />
            <col style={{ width: "13%" }} />
            <col style={{ width: "10%" }} />
            <col style={{ width: "17%" }} />
            <col style={{ width: "7%" }} />
          </colgroup>
          <thead>
            <tr style={{ borderBottom: "0.5px solid #e0e0e0" }}>
              <SortTh field="name" label="Name" />
              <SortTh field="dept" label="Department" />
              <th
                style={{
                  fontSize: 11,
                  fontWeight: 500,
                  color: "#999",
                  textAlign: "left",
                  padding: "10px 12px",
                  textTransform: "uppercase",
                  letterSpacing: ".05em",
                }}
              >
                Ext.
              </th>
              <SortTh field="email" label="Email" />
              <SortTh field="status" label="Status" />
              <th
                style={{
                  fontSize: 11,
                  fontWeight: 500,
                  color: "#999",
                  textAlign: "left",
                  padding: "10px 12px",
                  textTransform: "uppercase",
                  letterSpacing: ".05em",
                }}
              >
                Remark
              </th>
              <SortTh field="modifiedAt" label="Last Updated" />
              <th></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((emp) => {
              const si = statusInfo(emp.status);
              const [abg, afg] =
                AVATAR_COLORS[employees.indexOf(emp) % AVATAR_COLORS.length];
              const editable = canEdit(emp.id);
              return (
                <tr
                  key={emp.id}
                  style={{
                    borderBottom: "0.5px solid #eee",
                    transition: "background .1s",
                  }}
                  onMouseEnter={(e) =>
                    (e.currentTarget.style.background = "#fafaf8")
                  }
                  onMouseLeave={(e) => (e.currentTarget.style.background = "")}
                >
                  <td style={{ padding: "11px 12px" }}>
                    <div
                      style={{ display: "flex", alignItems: "center", gap: 10 }}
                    >
                      <div
                        style={{
                          width: 32,
                          height: 32,
                          borderRadius: "50%",
                          background: abg,
                          color: afg,
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          fontSize: 12,
                          fontWeight: 500,
                          flexShrink: 0,
                        }}
                        aria-hidden="true"
                      >
                        {initials(emp.name)}
                      </div>
                      <span
                        style={{
                          fontWeight: 500,
                          fontSize: 14,
                          whiteSpace: "nowrap",
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                        }}
                      >
                        {emp.name}
                      </span>
                    </div>
                  </td>
                  <td
                    style={{
                      fontSize: 13,
                      color: "#888",
                      padding: "11px 12px",
                      whiteSpace: "nowrap",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                    }}
                  >
                    {emp.dept}
                  </td>
                  <td style={{ fontSize: 14, padding: "11px 12px" }}>
                    {emp.ext}
                  </td>
                  <td
                    style={{
                      fontSize: 12,
                      color: "#185FA5",
                      padding: "11px 12px",
                      whiteSpace: "nowrap",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                    }}
                  >
                    {emp.email}
                  </td>
                  <td style={{ padding: "11px 12px" }}>
                    <span
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 5,
                        padding: "3px 10px",
                        borderRadius: 20,
                        fontSize: 12,
                        fontWeight: 500,
                        background: si.bg,
                        color: si.color,
                        whiteSpace: "nowrap",
                      }}
                    >
                      <span
                        style={{
                          width: 7,
                          height: 7,
                          borderRadius: "50%",
                          background: si.dot,
                          flexShrink: 0,
                        }}
                        aria-hidden="true"
                      />
                      {emp.status}
                    </span>
                  </td>
                  <td
                    style={{
                      fontSize: 12,
                      color: "#999",
                      padding: "11px 12px",
                      whiteSpace: "nowrap",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                    }}
                    title={emp.remark}
                  >
                    {emp.remark}
                  </td>
                  <td
                    style={{
                      fontSize: 11,
                      color: "#aaa",
                      padding: "11px 12px",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {formatDateTime(emp.modifiedAt)}
                  </td>
                  <td style={{ padding: "11px 12px", textAlign: "center" }}>
                    <button
                      onClick={() => editable && setEditingId(emp.id)}
                      aria-label={`Edit ${emp.name}`}
                      style={{
                        background: "none",
                        border: "0.5px solid #ddd",
                        borderRadius: 7,
                        padding: "4px 8px",
                        fontSize: 14,
                        cursor: editable ? "pointer" : "not-allowed",
                        opacity: editable ? 1 : 0.3,
                        transition: "background .15s",
                        lineHeight: 1,
                      }}
                      onMouseEnter={(e) =>
                        editable &&
                        (e.currentTarget.style.background = "#f0f0ee")
                      }
                      onMouseLeave={(e) =>
                        (e.currentTarget.style.background = "none")
                      }
                    >
                      ✏️
                    </button>
                  </td>
                </tr>
              );
            })}
            {filtered.length === 0 && (
              <tr>
                <td
                  colSpan={8}
                  style={{
                    textAlign: "center",
                    padding: "2rem",
                    color: "#aaa",
                    fontSize: 14,
                  }}
                >
                  No employees match your filters.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <p
        style={{
          fontSize: 12,
          color: "#bbb",
          marginTop: "1rem",
          textAlign: "right",
        }}
      >
        {filtered.length} of {employees.length} employees shown ·{" "}
        {mode === "live" ? "Live data" : "Demo data"} ·{" "}
        {new Date().toLocaleTimeString()}
      </p>

      {editingEmployee && (
        <EditModal
          employee={editingEmployee}
          onSave={handleSave}
          onClose={() => setEditingId(null)}
          saving={saving}
        />
      )}
    </div>
  );
}
