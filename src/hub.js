import * as SDK from "azure-devops-extension-sdk";

const API = "api-version=7.1";
const STALE_DAYS = 7;
const PALETTE = ["#0078d4", "#8764b8", "#00a3a3", "#c47f00", "#d13438", "#107c10", "#5c2d91", "#e3008c", "#4f6bed", "#498205"];

let ctx, base, token, project;
let allItems = [];
let currentRows = [];

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

async function api(path, options = {}) {
  const res = await fetch(`${base}${path}${path.includes("?") ? "&" : "?"}${API}`, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(options.headers || {}) }
  });
  if (!res.ok) throw new Error(`Request failed (${res.status}) for ${path}`);
  return res.json();
}

async function init() {
  await SDK.init({ loaded: false, applyTheme: true });
  await SDK.ready();
  ctx = SDK.getWebContext();
  token = await SDK.getAccessToken();
  base = (ctx.collection && ctx.collection.uri) || (ctx.host && ctx.host.uri) || "";
  if (!base.endsWith("/")) base += "/";
  project = ctx.project.name;

  $("refreshBtn").addEventListener("click", () => load());
  $("csvBtn").addEventListener("click", exportCsv);
  $("teamSelect").addEventListener("change", () => load());

  try {
    const teams = await api(`_apis/projects/${encodeURIComponent(ctx.project.id)}/teams?$top=200`);
    const sel = $("teamSelect");
    sel.innerHTML = teams.value
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((t) => `<option value="${esc(t.id)}">${esc(t.name)}</option>`)
      .join("");
    if (ctx.team && ctx.team.id) sel.value = ctx.team.id;
  } catch (e) {
    showError(e);
  }
  await load();
  SDK.notifyLoadSucceeded();
}

function showError(e) {
  $("content").innerHTML = `<div class="error">Could not load sprint data.<br><small>${esc(e.message || e)}</small></div>`;
  $("sprintInfo").textContent = "";
}

async function load() {
  $("content").innerHTML = `<div class="spinner">Loading sprint data…</div>`;
  const team = $("teamSelect").value;
  if (!team) return showError(new Error("No team found in this project."));
  try {
    const scope = `${encodeURIComponent(project)}/${encodeURIComponent(team)}`;

    const iters = await api(`${scope}/_apis/work/teamsettings/iterations?$timeframe=current`);
    const iter = iters.value && iters.value[0];
    if (!iter) {
      $("sprintInfo").textContent = "No current sprint";
      $("content").innerHTML = `<div class="empty">This team has no active sprint. Set a current iteration in Team Settings to see insights.</div>`;
      return;
    }
    const start = iter.attributes && iter.attributes.startDate ? new Date(iter.attributes.startDate) : null;
    const end = iter.attributes && iter.attributes.finishDate ? new Date(iter.attributes.finishDate) : null;
    const daysLeft = end ? Math.max(0, Math.ceil((end - new Date()) / 86400000)) : null;
    $("sprintInfo").textContent =
      `${iter.name}` + (start && end ? ` · ${fmt(start)} – ${fmt(end)}` : "") + (daysLeft != null ? ` · ${daysLeft} day(s) left` : "");

    const typesRes = await api(`${encodeURIComponent(project)}/_apis/wit/workitemtypes`);
    const catMap = {};
    typesRes.value.forEach((t) => (t.states || []).forEach((s) => (catMap[`${t.name}|${s.name}`] = s.category)));

    const wiql = await api(`${scope}/_apis/wit/wiql`, {
      method: "POST",
      body: JSON.stringify({
        query: "SELECT [System.Id] FROM WorkItems WHERE [System.IterationPath] = @CurrentIteration AND [System.WorkItemType] NOT IN ('Epic','Feature','Test Plan','Test Suite') ORDER BY [System.Id]"
      })
    });
    const ids = (wiql.workItems || []).map((w) => w.id);
    allItems = [];
    const fields = [
      "System.Id", "System.Title", "System.State", "System.WorkItemType", "System.AssignedTo", "System.Tags",
      "Microsoft.VSTS.Common.StateChangeDate", "Microsoft.VSTS.Scheduling.StoryPoints",
      "Microsoft.VSTS.Scheduling.RemainingWork", "Microsoft.VSTS.Scheduling.CompletedWork"
    ];
    for (let i = 0; i < ids.length; i += 200) {
      const batch = await api(`${encodeURIComponent(project)}/_apis/wit/workitemsbatch`, {
        method: "POST",
        body: JSON.stringify({ ids: ids.slice(i, i + 200), fields })
      });
      allItems.push(...batch.value);
    }
    render(allItems.map((w) => normalize(w, catMap)), daysLeft);
  } catch (e) {
    showError(e);
  } finally {
    SDK.resize();
  }
}

function fmt(d) {
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

function normalize(w, catMap) {
  const f = w.fields;
  const type = f["System.WorkItemType"];
  const state = f["System.State"];
  let cat = catMap[`${type}|${state}`];
  if (!cat) {
    const s = String(state).toLowerCase();
    cat = ["done", "closed", "completed"].includes(s) ? "Completed" : ["new", "to do", "proposed"].includes(s) ? "Proposed" : "InProgress";
  }
  const a = f["System.AssignedTo"];
  const tags = String(f["System.Tags"] || "").split(";").map((t) => t.trim().toLowerCase());
  const changed = f["Microsoft.VSTS.Common.StateChangeDate"] ? new Date(f["Microsoft.VSTS.Common.StateChangeDate"]) : null;
  return {
    id: w.id,
    url: `${base}${encodeURIComponent(project)}/_workitems/edit/${w.id}`,
    title: f["System.Title"],
    type, state, cat,
    assignee: a ? a.displayName || a : "Unassigned",
    blocked: tags.includes("blocked"),
    ageDays: changed ? Math.floor((Date.now() - changed) / 86400000) : 0,
    points: Number(f["Microsoft.VSTS.Scheduling.StoryPoints"] || 0),
    remaining: Number(f["Microsoft.VSTS.Scheduling.RemainingWork"] || 0)
  };
}

function render(rows, daysLeft) {
  currentRows = rows;
  if (!rows.length) {
    $("content").innerHTML = `<div class="empty">No work items in the current sprint.</div>`;
    return;
  }
  const active = rows.filter((r) => r.cat !== "Removed");
  const total = active.length;
  const done = active.filter((r) => r.cat === "Completed").length;
  const inProg = active.filter((r) => r.cat === "InProgress" || r.cat === "Resolved").length;
  const notStarted = active.filter((r) => r.cat === "Proposed").length;
  const blocked = active.filter((r) => r.blocked && r.cat !== "Completed");
  const unassigned = active.filter((r) => r.assignee === "Unassigned" && r.cat !== "Completed");
  const stale = active.filter((r) => r.cat === "InProgress" && r.ageDays >= STALE_DAYS);
  const pts = active.reduce((s, r) => s + r.points, 0);
  const ptsDone = active.filter((r) => r.cat === "Completed").reduce((s, r) => s + r.points, 0);
  const remaining = active.reduce((s, r) => s + r.remaining, 0);
  const pct = Math.round((done / total) * 100);

  // Health score: completion, minus penalties for risk
  const risk = (blocked.length * 6 + stale.length * 3 + unassigned.length * 2) / total * 100;
  let score = Math.max(0, Math.min(100, Math.round(100 - risk * 0.8 - (notStarted / total) * 15 * (daysLeft != null && daysLeft <= 3 ? 2 : 1))));
  const scoreColor = score >= 75 ? "var(--good)" : score >= 50 ? "var(--warn)" : "var(--bad)";

  // Workload by assignee, stacked by category
  const byPerson = {};
  active.forEach((r) => {
    byPerson[r.assignee] = byPerson[r.assignee] || { Completed: 0, InProgress: 0, Proposed: 0 };
    const k = r.cat === "Completed" ? "Completed" : r.cat === "Proposed" ? "Proposed" : "InProgress";
    byPerson[r.assignee][k]++;
  });
  const people = Object.entries(byPerson).sort((a, b) => sum(b[1]) - sum(a[1]));
  const maxP = Math.max(...people.map((p) => sum(p[1])));

  // States
  const byState = {};
  active.forEach((r) => (byState[r.state] = (byState[r.state] || 0) + 1));
  const states = Object.entries(byState).sort((a, b) => b[1] - a[1]);

  const attention = [
    ...blocked.map((r) => ({ ...r, flag: "Blocked", color: "var(--bad)" })),
    ...stale.filter((r) => !r.blocked).map((r) => ({ ...r, flag: `No change ${r.ageDays}d`, color: "var(--warn)" })),
    ...unassigned.filter((r) => !r.blocked && !(r.cat === "InProgress" && r.ageDays >= STALE_DAYS)).map((r) => ({ ...r, flag: "Unassigned", color: "var(--accent)" }))
  ].slice(0, 25);

  $("content").innerHTML = `
    <div class="grid">
      <div class="card score">
        <div class="ring" style="background:conic-gradient(${scoreColor} ${score * 3.6}deg, var(--line) 0)"><span>${score}</span></div>
        <div><strong>Sprint health</strong><div class="sub">${score >= 75 ? "On track" : score >= 50 ? "Needs attention" : "At risk"}</div></div>
      </div>
      <div class="card kpi"><div class="v">${pct}%</div><div class="l">Completed (${done}/${total})</div>
        <div class="progress"><div style="width:${pct}%"></div></div></div>
      <div class="card kpi"><div class="v">${inProg}</div><div class="l">In progress</div></div>
      <div class="card kpi"><div class="v">${notStarted}</div><div class="l">Not started</div></div>
      <div class="card kpi"><div class="v" style="color:${blocked.length ? "var(--bad)" : "inherit"}">${blocked.length}</div><div class="l">Blocked (tag “Blocked”)</div></div>
      <div class="card kpi"><div class="v" style="color:${stale.length ? "var(--warn)" : "inherit"}">${stale.length}</div><div class="l">Stale &gt; ${STALE_DAYS} days</div></div>
      <div class="card kpi"><div class="v">${unassigned.length}</div><div class="l">Unassigned open</div></div>
      <div class="card kpi"><div class="v">${pts ? `${ptsDone}/${pts}` : "–"}</div><div class="l">Story points done</div></div>
      <div class="card kpi"><div class="v">${remaining ? Math.round(remaining) + "h" : "–"}</div><div class="l">Remaining work</div></div>
    </div>
    <div class="cols">
      <div class="card">
        <h2>Workload by person</h2>
        ${people.map(([name, c]) => `
          <div class="bar-row"><div class="name" title="${esc(name)}">${esc(name)}</div>
            <div class="track" style="width:${(sum(c) / maxP) * 100}%">
              <div class="seg" style="width:${(c.Completed / sum(c)) * 100}%;background:var(--good)"></div>
              <div class="seg" style="width:${(c.InProgress / sum(c)) * 100}%;background:var(--accent)"></div>
              <div class="seg" style="width:${(c.Proposed / sum(c)) * 100}%;background:#8a8886"></div>
            </div><div>${sum(c)}</div></div>`).join("")}
        <div class="legend"><span><i class="dot" style="background:var(--good)"></i>Done</span><span><i class="dot" style="background:var(--accent)"></i>In progress</span><span><i class="dot" style="background:#8a8886"></i>Not started</span></div>
      </div>
      <div class="card">
        <h2>Items by state</h2>
        ${states.map(([s, n], i) => `
          <div class="bar-row"><div class="name">${esc(s)}</div>
            <div class="track"><div class="seg" style="width:${(n / states[0][1]) * 100}%;background:${PALETTE[i % PALETTE.length]}"></div></div><div>${n}</div></div>`).join("")}
      </div>
    </div>
    <div class="card" style="margin-top:12px">
      <h2>Needs attention</h2>
      ${attention.length ? `<table><thead><tr><th>ID</th><th>Title</th><th>Assignee</th><th>State</th><th>Flag</th></tr></thead><tbody>
        ${attention.map((r) => `<tr><td><a href="${r.url}" target="_blank" rel="noopener">#${r.id}</a></td><td>${esc(r.title)}</td><td>${esc(r.assignee)}</td><td>${esc(r.state)}</td><td><span class="pill" style="background:${r.color}">${esc(r.flag)}</span></td></tr>`).join("")}
      </tbody></table>` : `<div class="empty">Nothing flagged. Great job!</div>`}
    </div>`;
  SDK.resize();
}

function sum(o) {
  return o.Completed + o.InProgress + o.Proposed;
}

function exportCsv() {
  if (!currentRows.length) return;
  const head = ["ID", "Type", "Title", "State", "Assignee", "Blocked", "Days since state change", "Story points", "Remaining work"];
  const q = (v) => `"${String(v).replace(/"/g, '""')}"`;
  const lines = [head.map(q).join(",")].concat(
    currentRows.map((r) => [r.id, r.type, r.title, r.state, r.assignee, r.blocked ? "Yes" : "No", r.ageDays, r.points, r.remaining].map(q).join(","))
  );
  const blob = new Blob(["\ufeff" + lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "sprint-pulse.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
}

init().catch(showError);
