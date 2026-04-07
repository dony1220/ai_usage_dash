/**
 * CSV 기반 Gemini 이용현황 대시보드.
 * 같은 경로의 data.csv 로드 또는 파일 선택. 헤더 자동 매핑.
 */

const STATUS_PRIORITY = { 사용: 0, 미등록: 1, 퇴사: 2, 휴직: 3 };
const CHART_PALETTE = [
  "#2563eb", "#f59e0b", "#ef4444", "#8b5cf6", "#10b981", "#6366f1", "#14b8a6", "#f97316",
];

/** @type {any[]} */
let chartInstances = [];

function $(id) {
  return document.getElementById(id);
}

/** @param {string[]} headers */
function resolveColumnMapping(headers) {
  const raw = headers.map((h) => (h ?? "").trim());
  const norm = raw.map((h) => h.toLowerCase());

  const exact = (name) => raw.indexOf(name);
  const includes = (sub) => raw.findIndex((h) => h.includes(sub));

  let nameIdx = exact("이름");
  if (nameIdx < 0) nameIdx = includes("이름");

  let deptIdx = exact("소속1");
  if (deptIdx < 0) deptIdx = exact("소속");
  if (deptIdx < 0) deptIdx = includes("소속");

  let rankIdx = exact("직급");
  if (rankIdx < 0) rankIdx = includes("직급");

  let emailIdx = exact("이메일");
  if (emailIdx < 0) {
    emailIdx = raw.findIndex((h) => /^e-?mail/i.test(h) || h.toLowerCase() === "email");
  }

  let statusIdx = exact("사용/미등록/퇴사/휴직");
  if (statusIdx < 0) {
    statusIdx = raw.findIndex(
      (h) => h.includes("미등록") && h.includes("퇴사") && h.includes("휴직")
    );
  }
  if (statusIdx < 0) statusIdx = exact("상태");

  let usageIdx = raw.findIndex((h) => h.includes("전체 사용량"));
  if (usageIdx < 0) usageIdx = raw.findIndex((h) => h.includes("사용량"));
  if (usageIdx < 0) {
    usageIdx = norm.findIndex((h) => h.includes("usage") || h.includes("count"));
  }

  const map = { nameIdx, deptIdx, rankIdx, emailIdx, statusIdx, usageIdx };
  const missing = [];
  for (const k of ["nameIdx", "deptIdx", "rankIdx", "emailIdx", "statusIdx", "usageIdx"]) {
    if (map[k] < 0) missing.push(k.replace("Idx", ""));
  }
  return { map, rawHeaders: raw, missing };
}

function mapFromObjects(rows) {
  if (!rows.length) return { map: null, rawHeaders: [], missing: ["(빈 파일)"] };
  const rawHeaders = Object.keys(rows[0]).map((h) => (h ?? "").trim());
  return resolveColumnMapping(rawHeaders);
}

function parseUsage(val) {
  const s = String(val ?? "")
    .trim()
    .replace(/,/g, "");
  if (!s) return 0;
  const n = parseInt(String(parseFloat(s)), 10);
  return Number.isFinite(n) ? n : 0;
}

function objectRowsToNormalized(rows, col) {
  const { map, rawHeaders } = col;
  const kName = rawHeaders[map.nameIdx];
  const kDept = rawHeaders[map.deptIdx];
  const kRank = rawHeaders[map.rankIdx];
  const kEmail = rawHeaders[map.emailIdx];
  const kStatus = rawHeaders[map.statusIdx];
  const kUsage = rawHeaders[map.usageIdx];
  const out = [];
  for (const r of rows) {
    if (!r || typeof r !== "object") continue;
    let status = (r[kStatus] || "").trim() || "미분류";
    let rank = (r[kRank] || "").trim() || "미기재";
    let dept1 = (r[kDept] || "").trim();
    const name = (r[kName] || "").trim();
    const email = (r[kEmail] || "").trim();
    if (status === "휴직") {
      rank = "휴직";
      dept1 = "휴직";
    }
    const usage = parseUsage(r[kUsage]);
    if (!email && !name) continue;
    out.push({ name, dept1, rank, email, status, usage });
  }
  return out;
}

function buildPayload(normalized) {
  const status_count = {};
  for (const r of normalized) {
    status_count[r.status] = (status_count[r.status] || 0) + 1;
  }
  const statuses = Object.keys(status_count).sort(
    (a, b) =>
      (STATUS_PRIORITY[a] !== undefined ? STATUS_PRIORITY[a] : 99) -
        (STATUS_PRIORITY[b] !== undefined ? STATUS_PRIORITY[b] : 99) || a.localeCompare(b, "ko")
  );

  const rank_status = {};
  for (const r of normalized) {
    if (!rank_status[r.rank]) rank_status[r.rank] = {};
    rank_status[r.rank][r.status] = (rank_status[r.rank][r.status] || 0) + 1;
  }
  const rank_labels = Object.keys(rank_status).sort(
    (a, b) =>
      Object.values(rank_status[b]).reduce((s, n) => s + n, 0) -
        Object.values(rank_status[a]).reduce((s, n) => s + n, 0) || a.localeCompare(b, "ko")
  );
  const stacked_series = statuses.map((s) => ({
    name: s,
    data: rank_labels.map((rank) => rank_status[rank][s] || 0),
  }));

  let total_usage = 0;
  const dept_usage = {};
  for (const r of normalized) {
    total_usage += r.usage;
    dept_usage[r.dept1] = (dept_usage[r.dept1] || 0) + r.usage;
  }
  const dept1_labels = Object.keys(dept_usage).sort(
    (a, b) => dept_usage[b] - dept_usage[a] || a.localeCompare(b, "ko")
  );

  const all_users = [...normalized].sort((a, b) => b.usage - a.usage);

  return {
    kpi: {
      totalUsers: normalized.length,
      totalUsage: total_usage,
      activeUsers: status_count["사용"] || 0,
      inactiveUsers: normalized.length - (status_count["사용"] || 0),
    },
    statusLabels: statuses,
    statusCounts: statuses.map((s) => status_count[s]),
    dept1UsageLabels: dept1_labels,
    dept1UsageValues: dept1_labels.map((d) => dept_usage[d]),
    rankLabels: rank_labels,
    stackedSeries: stacked_series,
    allUsers: all_users,
  };
}

function destroyCharts() {
  chartInstances.forEach((c) => {
    try {
      c.destroy();
    } catch (_) {}
  });
  chartInstances = [];
}

function renderCharts(payload) {
  destroyCharts();
  const tickFont = { size: 13 };
  const legendBottom = {
    legend: {
      position: "bottom",
      labels: { font: { size: 13 }, boxWidth: 14, padding: 12 },
    },
  };

  chartInstances.push(
    new Chart($("chartRankStatus"), {
      type: "bar",
      data: {
        labels: payload.rankLabels,
        datasets: payload.stackedSeries.map((s, i) => ({
          label: s.name,
          data: s.data,
          backgroundColor: CHART_PALETTE[i % CHART_PALETTE.length],
        })),
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: legendBottom,
        scales: {
          x: {
            stacked: true,
            ticks: { font: tickFont, maxRotation: 45, minRotation: 0, autoSkip: true },
          },
          y: { stacked: true, beginAtZero: true, ticks: { font: tickFont } },
        },
      },
    })
  );

  chartInstances.push(
    new Chart($("chartStatusCount"), {
      type: "doughnut",
      data: {
        labels: payload.statusLabels,
        datasets: [
          {
            data: payload.statusCounts,
            backgroundColor: CHART_PALETTE.slice(0, payload.statusLabels.length),
          },
        ],
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: legendBottom },
    })
  );

  chartInstances.push(
    new Chart($("chartDeptUsage"), {
      type: "bar",
      data: {
        labels: payload.dept1UsageLabels,
        datasets: [{ label: "사용량", data: payload.dept1UsageValues, backgroundColor: "#0ea5e9" }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: {
            ticks: { font: tickFont, maxRotation: 50, minRotation: 0, autoSkip: true },
          },
          y: { beginAtZero: true, ticks: { font: tickFont } },
        },
      },
    })
  );
}

function formatInt(n) {
  return new Intl.NumberFormat("ko-KR").format(n);
}

function renderKpis(payload) {
  $("kpiTotalUsers").textContent = formatInt(payload.kpi.totalUsers);
  $("kpiTotalUsage").textContent = formatInt(payload.kpi.totalUsage);
  $("kpiActiveUsers").textContent = formatInt(payload.kpi.activeUsers);
  $("kpiInactiveUsers").textContent = formatInt(payload.kpi.inactiveUsers);
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function setupTableFilters(payload) {
  const statusSel = $("userStatusFilter");
  const searchInp = $("userSearchInput");
  const tbody = $("allUsersBody");
  const metaEl = $("userTableMeta");

  statusSel.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = "";
  optAll.textContent = "전체";
  statusSel.appendChild(optAll);

  const fromChart = new Set(payload.statusLabels);
  const fromRows = new Set(payload.allUsers.map((u) => u.status));
  const allSt = [...fromChart, ...[...fromRows].filter((s) => !fromChart.has(s))];
  allSt.sort(
    (a, b) =>
      (STATUS_PRIORITY[a] !== undefined ? STATUS_PRIORITY[a] : 99) -
        (STATUS_PRIORITY[b] !== undefined ? STATUS_PRIORITY[b] : 99) || a.localeCompare(b, "ko")
  );
  allSt.forEach((s) => {
    const o = document.createElement("option");
    o.value = s;
    o.textContent = s;
    statusSel.appendChild(o);
  });

  function filteredUsers() {
    const st = statusSel.value;
    const q = (searchInp.value || "").trim().toLowerCase();
    let list = payload.allUsers;
    if (st) list = list.filter((u) => u.status === st);
    if (q) {
      list = list.filter((u) => {
        const hay = `${u.name} ${u.dept1} ${u.rank} ${u.email} ${u.status}`.toLowerCase();
        return hay.includes(q);
      });
    }
    return list;
  }

  function renderUserTable() {
    const list = filteredUsers();
    tbody.innerHTML = "";
    const fmt = new Intl.NumberFormat("ko-KR");
    list.forEach((u, i) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${i + 1}</td><td>${escapeHtml(u.name)}</td><td>${escapeHtml(u.dept1)}</td><td>${escapeHtml(
        u.rank
      )}</td><td>${escapeHtml(u.email)}</td><td>${fmt.format(u.usage)}</td><td>${escapeHtml(u.status)}</td>`;
      tbody.appendChild(tr);
    });
    metaEl.textContent = `표시 ${list.length}명 / 전체 ${payload.allUsers.length}명`;
  }

  statusSel.onchange = renderUserTable;
  searchInp.oninput = renderUserTable;
  renderUserTable();
}

function setLoading(on) {
  $("loading").classList.toggle("visible", on);
}

function showError(msg) {
  const el = $("errorBanner");
  el.textContent = msg;
  el.classList.add("visible");
}

function hideError() {
  $("errorBanner").classList.remove("visible");
}

function showInfo(html) {
  const el = $("infoBanner");
  el.innerHTML = html;
  el.classList.add("visible");
}

function hideInfo() {
  $("infoBanner").classList.remove("visible");
}

function processCsvText(text, sourceLabel) {
  const results = Papa.parse(text, {
    header: true,
    skipEmptyLines: "greedy",
  });
  if (results.errors && results.errors.length) {
    console.warn("CSV parse warnings:", results.errors);
  }
  const rows = (results.data || []).filter((r) => r && typeof r === "object");
  if (!rows.length) throw new Error("데이터 행이 없습니다.");

  const col = mapFromObjects(rows);
  if (col.missing.length) {
    throw new Error(`필수 컬럼을 찾을 수 없습니다: ${col.missing.join(", ")}`);
  }
  const normalized = objectRowsToNormalized(rows, col);
  if (!normalized.length) throw new Error("유효한 행(이름 또는 이메일)이 없습니다.");

  const payload = buildPayload(normalized);

  hideError();
  showInfo(
    `<strong>${sourceLabel}</strong> · ${normalized.length}명 · 매핑: 이름=${escapeHtml(
      col.rawHeaders[col.map.nameIdx]
    )}, 소속=${escapeHtml(col.rawHeaders[col.map.deptIdx])}, 직급=${escapeHtml(
      col.rawHeaders[col.map.rankIdx]
    )}, 이메일=${escapeHtml(col.rawHeaders[col.map.emailIdx])}, 상태=${escapeHtml(
      col.rawHeaders[col.map.statusIdx]
    )}, 사용량=${escapeHtml(col.rawHeaders[col.map.usageIdx])}`
  );

  $("mainContent").hidden = false;
  renderKpis(payload);
  renderCharts(payload);
  setupTableFilters(payload);
  $("sourceMeta").textContent = sourceLabel;
}

async function loadDataCsv() {
  hideError();
  hideInfo();
  setLoading(true);
  $("sourceMeta").textContent = "data.csv 불러오는 중…";
  try {
    const res = await fetch("./data.csv", { cache: "no-store" });
    if (!res.ok) throw new Error(`data.csv를 찾을 수 없습니다 (${res.status}). 파일을 선택하거나 배포 폴더에 data.csv를 넣어 주세요.`);
    const text = await res.text();
    processCsvText(text, "data.csv (서버/로컬)");
  } catch (e) {
    $("mainContent").hidden = true;
    showError(e.message || String(e));
    $("sourceMeta").textContent = "data.csv 없음 → CSV 선택";
  } finally {
    setLoading(false);
  }
}

function onFileSelected(file) {
  if (!file) return;
  setLoading(true);
  hideError();
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const text = String(reader.result || "");
      processCsvText(text, `업로드: ${file.name}`);
    } catch (e) {
      $("mainContent").hidden = true;
      showError(e.message || String(e));
      $("sourceMeta").textContent = "오류";
    } finally {
      setLoading(false);
    }
  };
  reader.onerror = () => {
    setLoading(false);
    showError("파일을 읽을 수 없습니다.");
  };
  reader.readAsText(file, "UTF-8");
}

window.addEventListener("DOMContentLoaded", () => {
  $("csvFile").addEventListener("change", (e) => {
    const f = e.target.files && e.target.files[0];
    onFileSelected(f);
    e.target.value = "";
  });
  $("btnReloadData").addEventListener("click", () => loadDataCsv());
  loadDataCsv();
});
