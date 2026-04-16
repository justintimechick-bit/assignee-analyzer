import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
} from "recharts";

/* ── Simple SVG icon components (avoids lucide-react version conflicts) ── */
const Icon = ({ d, size = 16, color = "currentColor", style }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={style}>
    <path d={d} />
  </svg>
);

const Icons = {
  upload: (p) => <Icon {...p} d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12" />,
  download: (p) => <Icon {...p} d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3" />,
  users: (p) => <Icon {...p} d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2M9 7a4 4 0 100-8 4 4 0 000 8M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75" />,
  user: (p) => <Icon {...p} d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2M12 7a4 4 0 100-8 4 4 0 000 8" />,
  chevronDown: (p) => <Icon {...p} d="M6 9l6 6 6-6" />,
  chevronLeft: (p) => <Icon {...p} d="M15 18l-6-6 6-6" />,
  alert: (p) => <Icon {...p} d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0zM12 9v4M12 17h.01" />,
  zap: (p) => <Icon {...p} d="M13 2L3 14h9l-1 10 10-12h-9l1-10z" />,
  file: (p) => <Icon {...p} d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8zM14 2v6h6M16 13H8M16 17H8M10 9H8" />,
  trophy: (p) => <Icon {...p} d="M6 9H4.5a2.5 2.5 0 010-5H6M18 9h1.5a2.5 2.5 0 000-5H18M4 22h16M10 14.66V17c0 .55-.47.98-.97 1.21C7.85 18.75 7 20 7 22M14 14.66V17c0 .55.47.98.97 1.21C16.15 18.75 17 20 17 22M18 2H6v7a6 6 0 1012 0V2z" />,
  target: (p) => <Icon {...p} d="M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10zM12 18a6 6 0 100-12 6 6 0 000 12zM12 14a2 2 0 100-4 2 2 0 000 4z" />,
  brain: (p) => <Icon {...p} d="M12 2a7 7 0 017 7c0 2.38-1.19 4.47-3 5.74V17a2 2 0 01-2 2h-4a2 2 0 01-2-2v-2.26C6.19 13.47 5 11.38 5 9a7 7 0 017-7zM9 22h6M10 2v1M14 2v1M12 17v5" />,
  refresh: (p) => <Icon {...p} d="M23 4v6h-6M1 20v-6h6M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15" />,
  search: (p) => <Icon {...p} d="M11 19a8 8 0 100-16 8 8 0 000 16zM21 21l-4.35-4.35" />,
  bar: (p) => <Icon {...p} d="M18 20V10M12 20V4M6 20v-6" />,
  check: (p) => <Icon {...p} d="M22 11.08V12a10 10 0 11-5.93-9.14M22 4L12 14.01l-3-3" />,
};

/* ── palette ── */
const P = ["#f4a15d","#6bc5a0","#e07a5f","#81b0ff","#c89bdb","#5ec4d4","#e8c468","#f07167","#a0d995","#b8a9e8","#ff9b85","#7ec8e3"];
const FONT = `'Outfit', 'Segoe UI', sans-serif`;

const T = {
  "--bg": "#0c0e14", "--fg": "#e4e6eb", "--card": "#14161e",
  "--card2": "#1a1d28", "--border": "#252835", "--muted": "#7d819a",
  "--accent": "#f4a15d", "--green": "#6bc5a0", "--red": "#e07a5f",
  "--blue": "#81b0ff", "--purple": "#c89bdb",
};

/* ══════════════════════════════════════════════════════════════════
   PMS CASE CATEGORIZATION ENGINE
   Keyword-based classification tuned for hospitality property
   management system support (HMS, Epitome, IHG, DoD lodging)
   ══════════════════════════════════════════════════════════════════ */

const CATEGORY_RULES = [
  {
    name: "Night Audit",
    keywords: ["night audit", "nite audit", "end of day", "eod", "rollover", "roll over",
      "close day", "day close", "audit process", "auto post", "autopost",
      "bucket check", "no-show post", "no show post", "end of day process",
      "night audit fail", "audit report", "audit error"],
    weight: 10,
  },
  {
    name: "Check-In / Check-Out",
    keywords: ["check in", "checkin", "check-in", "check out", "checkout", "check-out",
      "walk-in", "walkin", "walk in", "early check", "late check",
      "express checkout", "front desk", "arrival", "departure",
      "key card", "keycard", "room key", "key packet", "guest arrival",
      "guest departure", "checked in", "checked out"],
    weight: 8,
  },
  {
    name: "Reservation Management",
    keywords: ["reservation", "booking", "availability", "avail", "cancel reservation",
      "modify reservation", "confirmation number", "overbooking", "overbook",
      "waitlist", "wait list", "allotment", "group booking",
      "travel order", "tdy", "pcs", "reservation lookup", "reservation search",
      "res not found", "duplicate reservation", "future reservation"],
    weight: 8,
  },
  {
    name: "Billing / Folio",
    keywords: ["folio", "billing", "invoice", "charge", "posting", "post charge",
      "payment", "refund", "credit card", "cc auth", "adjustment",
      "transfer charge", "split folio", "routing", "direct bill",
      "direct billing", "account receivable", "ar ", "a/r", "tax exempt",
      "tax post", "city ledger", "advance deposit", "folio balance",
      "zero balance", "balance due", "guest ledger", "incidental",
      "auth code", "settlement", "cashier"],
    weight: 9,
  },
  {
    name: "Rate / Revenue Management",
    keywords: ["rate code", "rate plan", "rate change", "baq rate", "per diem",
      "lodging rate", "pricing", "rate discrepancy", "rate setup", "yield",
      "revenue manage", "rate schedule", "govt rate", "government rate",
      "tla rate", "tdy rate", "seasonal rate", "rack rate", "negotiated rate",
      "rate override", "best available", "rate mismatch"],
    weight: 9,
  },
  {
    name: "Reporting",
    keywords: ["report", "crystal report", "export report", "print report",
      "occupancy report", "revenue report", "daily report", "manager report",
      "flash report", "statistical", "analytics", "forecast report",
      "history report", "housekeeping report", "guest count",
      "arrival list", "departure list", "in-house list", "inhouse list",
      "trial balance", "daily revenue", "monthly report", "generate report"],
    weight: 7,
  },
  {
    name: "User Access / Permissions",
    keywords: ["login", "log in", "password", "reset password", "access denied",
      "permission", "user account", "locked out", "lockout", "unlock",
      "credential", "authentication", "user setup", "new user",
      "deactivate user", "disable user", "role", "security access",
      "cac ", "common access card", "user rights", "access level",
      "cannot log", "forgot password", "expired password"],
    weight: 9,
  },
  {
    name: "Room / Housekeeping",
    keywords: ["housekeeping", "housekeeper", "room status", "dirty", "clean room",
      "inspected", "out of order", "ooo ", "out of service", "oos ",
      "room type", "room assignment", "room move", "room change",
      "maintenance", "work order", "room inventory", "vacant",
      "occupied", "room block", "floor plan", "room number",
      "room not ready", "hskp", "dnd", "do not disturb"],
    weight: 8,
  },
  {
    name: "Interface / Integration",
    keywords: ["interface", "integration", "pos ", "point of sale",
      "credit card interface", "door lock", "key system", "kaba",
      "saflok", "ving", "onity", "pbx", "phone system",
      "call accounting", "ota ", "channel manager", "gds ",
      "crs ", "central reservation", "pmsi", "dfas", "olvims",
      "shift4", "freedompay", "payment interface", "key encoder"],
    weight: 9,
  },
  {
    name: "System Connectivity",
    keywords: ["cannot connect", "connection", "timeout", "time out",
      "server down", "system down", "vpn", "securelink", "remote access",
      "network", "slow system", "performance", "latency", "outage",
      "not responding", "error connecting", "database connection",
      "sql server", "citrix", "rdp", "remote desktop", "dns",
      "cannot access", "site unreachable", "offline"],
    weight: 8,
  },
  {
    name: "Guest Profile",
    keywords: ["guest profile", "guest history", "profile merge",
      "duplicate profile", "loyalty", "vip ", "guest record",
      "guest info", "guest name", "guest data", "profile search",
      "profile update", "frequent guest", "guest preference",
      "guest notes", "profile duplicate"],
    weight: 7,
  },
  {
    name: "Correspondence / Letters",
    keywords: ["confirmation letter", "correspondence", "email template",
      "letter template", "folio print", "receipt print",
      "registration card", "reg card", "mail merge", "template",
      "print confirmation", "email confirmation", "auto email"],
    weight: 7,
  },
  {
    name: "Data Correction",
    keywords: ["correction", "fix data", "wrong date", "wrong amount",
      "wrong room", "wrong name", "incorrect", "data error",
      "update record", "modify record", "manual correction",
      "override", "back date", "backdate", "wrong guest",
      "wrong charge", "mispost", "mis-post", "reverse"],
    weight: 6,
  },
  {
    name: "System Configuration",
    keywords: ["configuration", "config", "setup", "property setup",
      "system setup", "parameter", "settings", "preferences",
      "room type setup", "rate code setup", "market code",
      "source code", "transaction code", "package setup",
      "tax setup", "currency", "property config", "system parameter"],
    weight: 7,
  },
  {
    name: "System Error / Bug",
    keywords: ["error", "bug", "crash", "frozen", "freeze", "not working",
      "broken", "glitch", "fail", "failure", "exception",
      "unexpected", "cannot save", "won't open", "does not open",
      "blank screen", "missing data", "data loss", "error message",
      "stack trace", "unhandled", "corrupt"],
    weight: 3,
  },
  {
    name: "Training / How-To",
    keywords: ["how to", "how do i", "training", "procedure",
      "documentation", "instruction", "help with", "show me",
      "walk me through", "step by step", "best practice",
      "guidance", "tutorial", "new employee", "refresher",
      "knowledge base", "sop "],
    weight: 5,
  },
  {
    name: "Upgrade / Patch",
    keywords: ["upgrade", "update", "patch", "version", "install",
      "migration", "release", "hotfix", "hot fix", "service pack",
      "new version", "deploy", "software update", "system update"],
    weight: 8,
  },
  {
    name: "Group / Event Management",
    keywords: ["group", "block", "allotment", "rooming list", "event",
      "conference", "meeting room", "function", "banquet",
      "group pickup", "group resume", "cutoff", "group rate",
      "group block", "group folio", "master folio"],
    weight: 7,
  },
];

function categorizeCase(description, resolution) {
  const text = `${description || ""} ${resolution || ""}`.toLowerCase();
  if (!text.trim()) return "Uncategorized";

  let bestCategory = "General Inquiry";
  let bestScore = 0;

  for (const rule of CATEGORY_RULES) {
    let matchCount = 0;
    let matchedKeywords = 0;
    for (const kw of rule.keywords) {
      const idx = text.indexOf(kw);
      if (idx !== -1) {
        matchedKeywords++;
        matchCount += kw.length * rule.weight;
        if (idx < 150) matchCount += kw.length * 2;
      }
    }
    if (matchedKeywords > 0) {
      const score = matchCount * (1 + matchedKeywords * 0.3);
      if (score > bestScore) {
        bestScore = score;
        bestCategory = rule.name;
      }
    }
  }

  return bestCategory;
}

/* ── helpers ── */
function detectColumns(headers) {
  const l = headers.map(h => h.toLowerCase().replace(/[^a-z0-9]/g, ""));
  const find = (kws) => { for (const k of kws) { const i = l.findIndex(h => h.includes(k)); if (i !== -1) return headers[i]; } return null; };
  return {
    assignee: find(["assignedto","assignee","owner","agent","assignmentgroup","group"]),
    state: find(["state","status","stage"]),
    description: find(["shortdescription","description","summary","subject","title"]),
    resolution: find(["resolutionnotes","resolution","closenotes","closingnotes","worknotes","comments","resolvenotes"]),
    created: find(["createdon","created","opened","openedat","sys_created_on"]),
    priority: find(["priority","urgency","severity"]),
    id: find(["number","incidentid","ticketid","caseid","sysid","id"]),
  };
}

function isClosed(state) {
  const s = (state || "").toString().toLowerCase();
  return s.includes("close") || s.includes("resolve") || s === "7" || s === "6";
}

/* ── Excel export ── */
function generateExcel(assigneeData, allCategories) {
  const wb = XLSX.utils.book_new();
  const ranked = [...assigneeData].sort((a, b) => b.closed - a.closed);
  const totalClosed = ranked.reduce((s, a) => s + a.closed, 0);

  const rankRows = [
    ["Assignee Performance Report"],
    ["Generated", new Date().toLocaleString()],
    ["Total Cases Analyzed", totalClosed],
    [],
    ["Rank", "Assignee", "Cases Closed", "% of Total", "Top Specialty", "Specialty %", "# Case Types"],
    ...ranked.map((a, i) => {
      const topCat = a.catData[0];
      return [i + 1, a.name, a.closed, `${((a.closed / totalClosed) * 100).toFixed(1)}%`,
        topCat?.name || "—", topCat ? `${((topCat.value / a.closed) * 100).toFixed(1)}%` : "—", a.catData.length];
    }),
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(rankRows);
  ws1["!cols"] = [{ wch: 6 }, { wch: 32 }, { wch: 14 }, { wch: 10 }, { wch: 26 }, { wch: 12 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Rankings");

  for (const a of ranked.slice(0, 20)) {
    const name = a.name.slice(0, 28).replace(/[\\\/\?\*\[\]]/g, "");
    const rows = [[a.name], ["Total Closed", a.closed], [],
      ["Case Type", "Count", "% of Their Cases"],
      ...a.catData.map(c => [c.name, c.value, `${((c.value / a.closed) * 100).toFixed(1)}%`])];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 30 }, { wch: 12 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, ws, name);
  }

  const catCounts = {};
  for (const a of assigneeData) for (const c of a.catData) catCounts[c.name] = (catCounts[c.name] || 0) + c.value;
  const catSorted = Object.entries(catCounts).sort((a, b) => b[1] - a[1]);
  const catRows = [["Case Type Distribution"], [],
    ["Category", "Total Cases", "% of All", "Top Assignee", "Top Assignee Count"],
    ...catSorted.map(([name, count]) => {
      const topA = ranked.reduce((best, a) => {
        const m = a.catData.find(c => c.name === name);
        return m && m.value > (best?.value || 0) ? { name: a.name, value: m.value } : best;
      }, null);
      return [name, count, `${((count / totalClosed) * 100).toFixed(1)}%`, topA?.name || "—", topA?.value || 0];
    })];
  const ws2 = XLSX.utils.aoa_to_sheet(catRows);
  ws2["!cols"] = [{ wch: 28 }, { wch: 14 }, { wch: 10 }, { wch: 30 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Case Types");

  XLSX.writeFile(wb, `Assignee_Analysis_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* ── Components ── */

function StatCard({ icon, label, value, sub, color }) {
  return (
    <div style={{
      background: "var(--card)", borderRadius: 14, padding: "18px 20px",
      border: "1px solid var(--border)", display: "flex", flexDirection: "column", gap: 4,
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span style={{ fontSize: 11, fontWeight: 600, letterSpacing: "0.07em", textTransform: "uppercase", color: "var(--muted)" }}>{label}</span>
        {icon}
      </div>
      <span style={{ fontSize: 26, fontWeight: 700, color: color || "var(--fg)", lineHeight: 1.1 }}>{value}</span>
      {sub && <span style={{ fontSize: 12, color: "var(--muted)" }}>{sub}</span>}
    </div>
  );
}

function ColumnMapper({ headers, colMap, setColMap }) {
  const fields = [
    { key: "assignee", label: "Assignee / Group", required: true },
    { key: "state", label: "State / Status", required: true },
    { key: "description", label: "Description", required: true },
    { key: "resolution", label: "Resolution Notes" },
    { key: "priority", label: "Priority" },
    { key: "id", label: "Ticket ID" },
    { key: "created", label: "Created Date" },
  ];
  return (
    <div style={{ background: "var(--card)", borderRadius: 14, padding: 22, border: "1px solid var(--border)", marginBottom: 20 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
        {Icons.target({ size: 16, color: "var(--accent)" })}
        <h3 style={{ fontSize: 14, fontWeight: 600, margin: 0, color: "var(--fg)" }}>Map Your Columns</h3>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 10 }}>
        {fields.map(f => (
          <div key={f.key}>
            <label style={{
              fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em",
              color: f.required ? "var(--accent)" : "var(--muted)", display: "block", marginBottom: 3,
            }}>{f.label} {f.required && "*"}</label>
            <div style={{ position: "relative" }}>
              <select value={colMap[f.key] || ""}
                onChange={e => setColMap(p => ({ ...p, [f.key]: e.target.value || null }))}
                style={{
                  width: "100%", padding: "7px 26px 7px 9px", borderRadius: 8,
                  border: "1px solid var(--border)", background: "var(--bg)",
                  color: "var(--fg)", fontSize: 12, appearance: "none", cursor: "pointer",
                }}>
                <option value="">— none —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
              {Icons.chevronDown({ size: 12, style: { position: "absolute", right: 7, top: "50%", transform: "translateY(-50%)", pointerEvents: "none", color: "var(--muted)" } })}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function AssigneeCard({ data, rank, onClick, maxClosed }) {
  const pct = maxClosed ? (data.closed / maxClosed) * 100 : 0;
  return (
    <div onClick={onClick} style={{
      background: "var(--card)", borderRadius: 14, padding: "16px 18px",
      border: "1px solid var(--border)", cursor: "pointer",
      transition: "all 0.2s ease", position: "relative", overflow: "hidden",
    }}
    onMouseEnter={e => { e.currentTarget.style.borderColor = "var(--accent)"; e.currentTarget.style.transform = "translateY(-2px)"; }}
    onMouseLeave={e => { e.currentTarget.style.borderColor = "var(--border)"; e.currentTarget.style.transform = "translateY(0)"; }}>
      <div style={{ position: "absolute", bottom: 0, left: 0, width: `${pct}%`, height: 3, background: `linear-gradient(90deg, var(--accent), transparent)`, opacity: 0.5 }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{
            width: 36, height: 36, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center",
            background: rank <= 3 ? "rgba(244,161,93,0.15)" : "rgba(125,129,154,0.1)",
            fontSize: 14, fontWeight: 700,
          }}>
            {rank <= 3 ? Icons.trophy({ size: 18, color: "var(--accent)" }) : <span style={{ color: "var(--muted)" }}>#{rank}</span>}
          </div>
          <div>
            <div style={{ fontSize: 14, fontWeight: 600, color: "var(--fg)" }}>{data.name}</div>
            <div style={{ fontSize: 11, color: "var(--muted)" }}>{data.catData.length} case type{data.catData.length !== 1 ? "s" : ""}</div>
          </div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 22, fontWeight: 700, color: "var(--accent)", lineHeight: 1 }}>{data.closed}</div>
          <div style={{ fontSize: 10, color: "var(--muted)", textTransform: "uppercase", letterSpacing: "0.05em" }}>closed</div>
        </div>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
        {data.catData.slice(0, 3).map((c, i) => (
          <span key={i} style={{
            padding: "3px 8px", borderRadius: 6, fontSize: 10, fontWeight: 600,
            background: `${P[i % P.length]}18`, color: P[i % P.length],
          }}>{c.name} ({c.value})</span>
        ))}
        {data.catData.length > 3 && (
          <span style={{ padding: "3px 8px", borderRadius: 6, fontSize: 10, color: "var(--muted)" }}>+{data.catData.length - 3} more</span>
        )}
      </div>
    </div>
  );
}

function AssigneeDetail({ data, onBack }) {
  const radarData = data.catData.slice(0, 8).map(c => ({
    subject: c.name.length > 18 ? c.name.slice(0, 16) + "…" : c.name, count: c.value,
  }));

  return (
    <div>
      <button onClick={onBack} style={{
        display: "flex", alignItems: "center", gap: 6, padding: "8px 14px",
        borderRadius: 10, border: "1px solid var(--border)", background: "var(--card)",
        color: "var(--fg)", fontSize: 13, fontWeight: 600, cursor: "pointer", marginBottom: 20,
      }}>{Icons.chevronLeft({ size: 16 })} All Assignees</button>

      <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 24 }}>
        <div style={{
          width: 52, height: 52, borderRadius: 14, display: "flex", alignItems: "center", justifyContent: "center",
          background: "rgba(244,161,93,0.12)",
        }}>{Icons.user({ size: 26, color: "var(--accent)" })}</div>
        <div>
          <h2 style={{ fontSize: 22, fontWeight: 700, margin: 0, color: "var(--fg)" }}>{data.name}</h2>
          <p style={{ fontSize: 13, color: "var(--muted)", margin: 0 }}>
            {data.closed} cases closed · {data.catData.length} case type{data.catData.length !== 1 ? "s" : ""} · top specialty: <span style={{ color: "var(--accent)" }}>{data.catData[0]?.name || "—"}</span>
          </p>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: radarData.length >= 3 ? "1fr 1fr" : "1fr", gap: 16, marginBottom: 16 }}>
        {radarData.length >= 3 && (
          <div style={{ background: "var(--card)", borderRadius: 14, padding: 20, border: "1px solid var(--border)" }}>
            <h3 style={{ fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--muted)", margin: "0 0 10px" }}>Specialty Profile</h3>
            <ResponsiveContainer width="100%" height={280}>
              <RadarChart data={radarData}>
                <PolarGrid stroke="#252835" />
                <PolarAngleAxis dataKey="subject" tick={{ fill: "#7d819a", fontSize: 10 }} />
                <PolarRadiusAxis tick={{ fill: "#7d819a", fontSize: 9 }} />
                <Radar dataKey="count" stroke="#f4a15d" fill="#f4a15d" fillOpacity={0.2} strokeWidth={2} />
              </RadarChart>
            </ResponsiveContainer>
          </div>
        )}
        <div style={{ background: "var(--card)", borderRadius: 14, padding: 20, border: "1px solid var(--border)" }}>
          <h3 style={{ fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--muted)", margin: "0 0 10px" }}>Case Type Breakdown</h3>
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={data.catData.slice(0, 10)} layout="vertical" margin={{ left: 0, right: 10 }}>
              <XAxis type="number" tick={{ fill: "#7d819a", fontSize: 10 }} />
              <YAxis type="category" dataKey="name" width={130} tick={{ fill: "#e4e6eb", fontSize: 10 }} />
              <Tooltip contentStyle={{ background: "#14161e", border: "1px solid #252835", borderRadius: 8, fontSize: 12 }} />
              <Bar dataKey="value" radius={[0, 6, 6, 0]}>
                {data.catData.slice(0, 10).map((_, i) => <Cell key={i} fill={P[i % P.length]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div style={{ background: "var(--card)", borderRadius: 14, padding: 20, border: "1px solid var(--border)", marginBottom: 16 }}>
        <h3 style={{ fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--muted)", margin: "0 0 10px" }}>Case Distribution</h3>
        <ResponsiveContainer width="100%" height={250}>
          <PieChart>
            <Pie data={data.catData} dataKey="value" nameKey="name" cx="50%" cy="50%"
              outerRadius={90} innerRadius={50} paddingAngle={2} strokeWidth={0}
              label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
              labelLine={{ stroke: "#7d819a" }}>
              {data.catData.map((_, i) => <Cell key={i} fill={P[i % P.length]} />)}
            </Pie>
            <Tooltip contentStyle={{ background: "#14161e", border: "1px solid #252835", borderRadius: 8, fontSize: 12 }} />
          </PieChart>
        </ResponsiveContainer>
      </div>

      <div style={{ background: "var(--card)", borderRadius: 14, padding: 20, border: "1px solid var(--border)" }}>
        <h3 style={{ fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--muted)", margin: "0 0 14px" }}>All Case Types</h3>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              {["Case Type", "Count", "% of Total"].map(h => (
                <th key={h} style={{ textAlign: "left", padding: "8px 12px", fontSize: 11, fontWeight: 600, color: "var(--muted)", borderBottom: "1px solid var(--border)", textTransform: "uppercase", letterSpacing: "0.04em" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.catData.map((c, i) => (
              <tr key={i} style={{ borderBottom: i < data.catData.length - 1 ? "1px solid var(--border)" : "none" }}>
                <td style={{ padding: "10px 12px", fontSize: 13, display: "flex", alignItems: "center", gap: 8 }}>
                  <div style={{ width: 8, height: 8, borderRadius: 3, background: P[i % P.length], flexShrink: 0 }} />
                  {c.name}
                </td>
                <td style={{ padding: "10px 12px", fontSize: 13, fontWeight: 600 }}>{c.value}</td>
                <td style={{ padding: "10px 12px", fontSize: 13, color: "var(--muted)" }}>{((c.value / data.closed) * 100).toFixed(1)}%</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ── Main App ── */
export default function AssigneeAnalyzer() {
  const [rawData, setRawData] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [colMap, setColMap] = useState({});
  const [fileName, setFileName] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [error, setError] = useState(null);
  const [assigneeResults, setAssigneeResults] = useState(null);
  const [selectedAssignee, setSelectedAssignee] = useState(null);
  const [searchFilter, setSearchFilter] = useState("");

  const processFile = useCallback((file) => {
    setError(null); setFileName(file.name); setAssigneeResults(null); setSelectedAssignee(null);
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv" || ext === "tsv") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: (r) => {
          if (!r.data.length) return setError("No data rows found.");
          setHeaders(r.meta.fields || []); setColMap(detectColumns(r.meta.fields || [])); setRawData(r.data);
        },
        error: () => setError("Failed to parse CSV."),
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
          const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
          if (!json.length) return setError("No data rows found.");
          const h = Object.keys(json[0]);
          setHeaders(h); setColMap(detectColumns(h)); setRawData(json);
        } catch { setError("Failed to parse Excel file."); }
      };
      reader.readAsArrayBuffer(file);
    } else setError("Upload a .csv, .tsv, or .xlsx file.");
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault(); setDragOver(false);
    if (e.dataTransfer?.files?.[0]) processFile(e.dataTransfer.files[0]);
  }, [processFile]);

  const runAnalysis = useCallback(() => {
    if (!rawData || !colMap.assignee || !colMap.state || !colMap.description) return;
    setError(null);
    const closedCases = rawData.filter(r => isClosed(r[colMap.state]));
    if (!closedCases.length) { setError("No closed/resolved cases found. Check that your State column is mapped correctly."); return; }

    const byAssignee = {};
    for (const row of closedCases) {
      const name = (row[colMap.assignee] || "Unassigned").toString().trim();
      if (!byAssignee[name]) byAssignee[name] = [];
      const desc = (row[colMap.description] || "").toString();
      const res = colMap.resolution ? (row[colMap.resolution] || "").toString() : "";
      byAssignee[name].push({ category: categorizeCase(desc, res), row });
    }

    const results = Object.entries(byAssignee).map(([name, cases]) => {
      const catCounts = {};
      for (const c of cases) catCounts[c.category] = (catCounts[c.category] || 0) + 1;
      const catData = Object.entries(catCounts).sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value }));
      return { name, closed: cases.length, catData };
    });

    results.sort((a, b) => b.closed - a.closed);
    setAssigneeResults(results);
  }, [rawData, colMap]);

  const reset = () => {
    setRawData(null); setHeaders([]); setColMap({}); setFileName("");
    setError(null); setAssigneeResults(null); setSelectedAssignee(null);
  };

  const filteredResults = useMemo(() => {
    if (!assigneeResults) return [];
    if (!searchFilter) return assigneeResults;
    const q = searchFilter.toLowerCase();
    return assigneeResults.filter(a => a.name.toLowerCase().includes(q) || a.catData.some(c => c.name.toLowerCase().includes(q)));
  }, [assigneeResults, searchFilter]);

  const totalClosed = assigneeResults?.reduce((s, a) => s + a.closed, 0) || 0;
  const allCategories = useMemo(() => {
    if (!assigneeResults) return [];
    const counts = {};
    for (const a of assigneeResults) for (const c of a.catData) counts[c.name] = (counts[c.name] || 0) + c.value;
    return Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([name, value]) => ({ name, value }));
  }, [assigneeResults]);

  const maxClosed = assigneeResults?.[0]?.closed || 1;
  const rootStyle = { ...T, fontFamily: FONT, color: "var(--fg)", background: "var(--bg)", minHeight: "100vh", padding: "28px 24px" };

  if (!rawData) {
    return (
      <div style={rootStyle}>
        <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&family=JetBrains+Mono:wght@700&display=swap" rel="stylesheet" />
        <div style={{ maxWidth: 600, margin: "0 auto", textAlign: "center" }}>
          <div style={{ marginBottom: 36 }}>
            <div style={{ display: "inline-flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
              {Icons.brain({ size: 28, color: "var(--accent)" })}
              <h1 style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 24, fontWeight: 700, margin: 0, color: "var(--accent)" }}>Assignee Analyzer</h1>
            </div>
            <p style={{ fontSize: 14, color: "var(--muted)", margin: 0 }}>PMS case-type analysis per assignee · instant categorization · no API needed</p>
          </div>
          <div onDragOver={e => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onDrop={handleDrop}
            onClick={() => document.getElementById("fu").click()}
            style={{
              border: `2px dashed ${dragOver ? "var(--accent)" : "var(--border)"}`,
              borderRadius: 18, padding: "52px 28px", cursor: "pointer",
              background: dragOver ? "rgba(244,161,93,0.04)" : "var(--card)", transition: "all 0.2s",
            }}>
            {Icons.upload({ size: 36, color: dragOver ? "#f4a15d" : "#7d819a", style: { marginBottom: 14 } })}
            <p style={{ fontSize: 15, fontWeight: 600, margin: "0 0 4px" }}>Drop your ServiceNow export</p>
            <p style={{ fontSize: 13, color: "var(--muted)", margin: 0 }}>.csv, .tsv, or .xlsx — needs description + assignee columns</p>
            <input id="fu" type="file" accept=".csv,.tsv,.xlsx,.xls" style={{ display: "none" }} onChange={e => e.target.files?.[0] && processFile(e.target.files[0])} />
          </div>
          {error && (
            <div style={{ marginTop: 14, padding: "10px 14px", borderRadius: 10, background: "rgba(224,122,95,0.12)", color: "var(--red)", fontSize: 13, display: "flex", alignItems: "center", gap: 8 }}>
              {Icons.alert({ size: 16, color: "var(--red)" })} {error}
            </div>
          )}
        </div>
      </div>
    );
  }

  if (selectedAssignee) {
    return (
      <div style={rootStyle}>
        <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&family=JetBrains+Mono:wght@700&display=swap" rel="stylesheet" />
        <div style={{ maxWidth: 900, margin: "0 auto" }}>
          <AssigneeDetail data={selectedAssignee} onBack={() => setSelectedAssignee(null)} />
        </div>
      </div>
    );
  }

  return (
    <div style={rootStyle}>
      <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&family=JetBrains+Mono:wght@700&display=swap" rel="stylesheet" />
      <div style={{ maxWidth: 960, margin: "0 auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, marginBottom: 20 }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              {Icons.brain({ size: 20, color: "var(--accent)" })}
              <h1 style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 18, fontWeight: 700, margin: 0, color: "var(--accent)" }}>Assignee Analyzer</h1>
            </div>
            <p style={{ fontSize: 12, color: "var(--muted)", margin: "2px 0 0", display: "flex", alignItems: "center", gap: 5 }}>
              {Icons.file({ size: 12, color: "var(--muted)" })} {fileName} — {rawData.length.toLocaleString()} rows
            </p>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={reset} style={{
              padding: "8px 14px", borderRadius: 10, border: "1px solid var(--border)",
              background: "var(--card)", color: "var(--fg)", fontSize: 12, fontWeight: 600, cursor: "pointer",
              display: "flex", alignItems: "center", gap: 5,
            }}>{Icons.refresh({ size: 13 })} New File</button>
            {assigneeResults && (
              <button onClick={() => generateExcel(assigneeResults, allCategories)} style={{
                padding: "8px 16px", borderRadius: 10, border: "none",
                background: "var(--accent)", color: "#0c0e14", fontSize: 12, fontWeight: 700, cursor: "pointer",
                display: "flex", alignItems: "center", gap: 5,
              }}>{Icons.download({ size: 13, color: "#0c0e14" })} Export .xlsx</button>
            )}
          </div>
        </div>

        <ColumnMapper headers={headers} colMap={colMap} setColMap={setColMap} />

        {!assigneeResults && (
          <div style={{ textAlign: "center", padding: "16px 0 24px" }}>
            {(!colMap.assignee || !colMap.state || !colMap.description) ? (
              <p style={{ color: "var(--muted)", fontSize: 13 }}>
                Map <strong>Assignee</strong>, <strong>State</strong>, and <strong>Description</strong> columns to continue.
              </p>
            ) : (
              <button onClick={runAnalysis} style={{
                padding: "12px 28px", borderRadius: 12, border: "none",
                background: "linear-gradient(135deg, var(--accent), #e8c468)",
                color: "#0c0e14", fontSize: 15, fontWeight: 700, cursor: "pointer",
                display: "inline-flex", alignItems: "center", gap: 8,
                boxShadow: "0 4px 24px rgba(244,161,93,0.3)",
              }}>{Icons.zap({ size: 18, color: "#0c0e14" })} Analyze Assignees</button>
            )}
            {error && (
              <div style={{ marginTop: 12, padding: "10px 14px", borderRadius: 10, background: "rgba(224,122,95,0.12)", color: "var(--red)", fontSize: 13, display: "inline-flex", alignItems: "center", gap: 6 }}>
                {Icons.alert({ size: 14, color: "var(--red)" })} {error}
              </div>
            )}
          </div>
        )}

        {assigneeResults && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 20 }}>
              <StatCard icon={Icons.users({ size: 16, color: "var(--accent)", style: { opacity: 0.5 } })} label="Assignees" value={assigneeResults.length} color="var(--accent)" />
              <StatCard icon={Icons.check({ size: 16, color: "var(--green)", style: { opacity: 0.5 } })} label="Cases Closed" value={totalClosed.toLocaleString()} color="var(--green)" />
              <StatCard icon={Icons.brain({ size: 16, color: "var(--purple)", style: { opacity: 0.5 } })} label="Case Types" value={allCategories.length} color="var(--purple)" />
              <StatCard icon={Icons.trophy({ size: 16, color: "var(--blue)", style: { opacity: 0.5 } })} label="Top Closer" value={assigneeResults[0]?.name?.split(/[\s,]/)[0] || "—"} sub={`${assigneeResults[0]?.closed || 0} cases`} color="var(--blue)" />
            </div>

            {allCategories.length > 0 && (
              <div style={{ background: "var(--card)", borderRadius: 14, padding: 20, border: "1px solid var(--border)", marginBottom: 20 }}>
                <h3 style={{ fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--muted)", margin: "0 0 14px" }}>Overall Case Type Distribution</h3>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={allCategories.slice(0, 12)} margin={{ left: 0, right: 10 }}>
                    <XAxis dataKey="name" tick={{ fill: "#7d819a", fontSize: 10, angle: -25, textAnchor: "end" }} height={70} />
                    <YAxis tick={{ fill: "#7d819a", fontSize: 10 }} />
                    <Tooltip contentStyle={{ background: "#14161e", border: "1px solid #252835", borderRadius: 8, fontSize: 12, maxWidth: 240 }} />
                    <Bar dataKey="value" radius={[6, 6, 0, 0]}>
                      {allCategories.slice(0, 12).map((_, i) => <Cell key={i} fill={P[i % P.length]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}

            <div style={{ position: "relative", marginBottom: 16 }}>
              <div style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)" }}>
                {Icons.search({ size: 15, color: "var(--muted)" })}
              </div>
              <input value={searchFilter} onChange={e => setSearchFilter(e.target.value)}
                placeholder="Search by assignee name or case type..."
                style={{
                  width: "100%", padding: "10px 12px 10px 36px", borderRadius: 10,
                  border: "1px solid var(--border)", background: "var(--card)",
                  color: "var(--fg)", fontSize: 13, boxSizing: "border-box",
                }} />
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 12 }}>
              {filteredResults.map((a) => (
                <AssigneeCard key={a.name} data={a} rank={assigneeResults.indexOf(a) + 1}
                  maxClosed={maxClosed} onClick={() => setSelectedAssignee(a)} />
              ))}
            </div>

            {filteredResults.length === 0 && (
              <p style={{ textAlign: "center", color: "var(--muted)", padding: 24, fontSize: 13 }}>No assignees match your search.</p>
            )}
          </>
        )}
      </div>
    </div>
  );
}
