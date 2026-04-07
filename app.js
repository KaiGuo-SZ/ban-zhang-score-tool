const GROUP_MAP = new Map([
  ["两广总督", "社群一部"],
  ["双面鬼龙", "社群一部"],
  ["量子启明", "社群一部"],
  ["金戈铁马", "社群一部"],
  ["超星不凡", "社群一部"],
  ["北斗七星", "社群一部"],
  ["梦之队", "社群二部"],
  ["纵横四海", "社群二部"],
  ["铃兰泽门", "社群三部"],
  ["铃兰邓门", "社群三部"],
  ["铃兰中学", "社群三部"],
  ["暂无小组", "未分组"],
]);

const REQUIRED_COLUMNS = ["营期", "班级", "小组", "获客", "流水", "添加人数"];
const PERIODS = [-1, -2, -3];

function qs(sel) {
  const el = document.querySelector(sel);
  if (!el) throw new Error(`缺少元素: ${sel}`);
  return el;
}

function tokenizeQuery(input) {
  const s = String(input ?? "").trim().toLowerCase();
  if (!s) return [];
  return s
    .split(/[,\s，、;；]+/)
    .map((t) => t.trim())
    .filter(Boolean);
}

function toNumber(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (!s) return null;
  const cleaned = s.replaceAll(",", "").replaceAll("%", "").trim();
  if (!cleaned) return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function round2(n) {
  if (!Number.isFinite(n)) return null;
  return Math.round(n * 100) / 100;
}

function roundHalfUp(x) {
  if (!Number.isFinite(x)) return null;
  return Math.floor(x + 0.5);
}

function cleanLeaderName(name) {
  const s = String(name ?? "").trim();
  if (!s) return "";
  return s.replace(/[（(].*$/, "").trim();
}

function decodeFileToText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error("读取文件失败"));
    reader.onload = () => {
      const buf = reader.result;
      try {
        const text = new TextDecoder("utf-16le").decode(buf);
        resolve(text);
      } catch (e) {
        reject(e);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function parseTSV(text) {
  const lines = text.split(/\r\n|\n|\r/).filter((l) => l !== "");
  if (lines.length === 0) return { rows: [], headers: [] };
  const headers = lines[0].split("\t").map((h) => h.trim());
  const rows = [];
  for (let i = 1; i < lines.length; i += 1) {
    const cols = lines[i].split("\t");
    const obj = {};
    for (let j = 0; j < headers.length; j += 1) obj[headers[j]] = cols[j] ?? "";
    rows.push(obj);
  }
  return { rows, headers };
}

function modeOrFirst(values, fallback) {
  const map = new Map();
  for (const v of values) {
    const s = String(v ?? "").trim();
    if (!s) continue;
    map.set(s, (map.get(s) ?? 0) + 1);
  }
  if (map.size === 0) return fallback;
  let bestKey = null;
  let bestCount = -1;
  for (const [k, c] of map.entries()) {
    if (c > bestCount) {
      bestCount = c;
      bestKey = k;
    }
  }
  return bestKey ?? fallback;
}

function groupBy(rows, keyFn) {
  const map = new Map();
  for (const r of rows) {
    const k = keyFn(r);
    const bucket = map.get(k) ?? [];
    bucket.push(r);
    map.set(k, bucket);
  }
  return map;
}

function assignLatestPeriodByLeader(rows) {
  const byLeader = groupBy(rows, (r) => r["班长"]);
  const out = [];
  for (const [leader, items] of byLeader.entries()) {
    const terms = Array.from(new Set(items.map((x) => x["营期"]))).filter((x) => x !== null);
    terms.sort((a, b) => b - a);
    const termToRank = new Map();
    for (let i = 0; i < terms.length; i += 1) termToRank.set(terms[i], -(i + 1));
    for (const it of items) out.push({ ...it, 最新营期: termToRank.get(it["营期"]) ?? null });
  }
  return out;
}

function assignRankAndScore(rows) {
  const byTerm = groupBy(rows, (r) => r["营期"]);
  const enriched = [];

  for (const [term, items] of byTerm.entries()) {
    const participants = items
      .map((r, idx) => ({ idx, v: r["添加产值"] }))
      .filter((x) => x.v !== null && Number.isFinite(x.v));

    participants.sort((a, b) => b.v - a.v);

    const rankByIdx = new Map();
    let currentRank = 0;
    let lastValue = null;
    for (let i = 0; i < participants.length; i += 1) {
      const { idx, v } = participants[i];
      if (lastValue === null || v !== lastValue) {
        currentRank = i + 1;
        lastValue = v;
      }
      rankByIdx.set(idx, currentRank);
    }

    const count = participants.length;
    for (let i = 0; i < items.length; i += 1) {
      const base = items[i];
      const rank = rankByIdx.get(i) ?? null;
      const pct = rank !== null && count > 0 ? rank / count : null;
      const rankPct = pct !== null ? round2(pct * 100) : null;
      const score =
        pct === null
          ? 0
          : pct <= 0.2
            ? 40
            : pct <= 0.5
              ? 30
              : pct <= 0.6
                ? 20
                : pct <= 0.8
                  ? 10
                  : 0;
      enriched.push({
        ...base,
        产值排名: rank,
        "产值排名%": rankPct,
        产值得分: score,
      });
    }
  }

  return enriched;
}

function buildWideTable(rows) {
  const byLeader = groupBy(rows.filter((r) => PERIODS.includes(r["最新营期"])), (r) => r["班长"]);
  const leaders = [];

  for (const [leader, items] of byLeader.entries()) {
    const row = {
      班长: leader,
      大组: "",
      小组: "",
    };

    const pNeg1 = items.find((x) => x["最新营期"] === -1);
    row["小组"] = String(pNeg1?.["小组"] ?? "暂无小组");
    row["大组"] = String(pNeg1?.["大组"] ?? "未分组");

    let total = 0;

    for (const p of PERIODS) {
      const it = items.find((x) => x["最新营期"] === p);
      const leads = it?.["获客"] ?? null;
      const classType = it?.["班型"] ?? null;
      const classScore = it?.["班型得分"] ?? null;
      const addValue = it?.["添加产值"] ?? null;
      const rank = it?.["产值排名"] ?? null;
      const rankPct = it?.["产值排名%"] ?? null;
      const valueScore = it?.["产值得分"] ?? 0;

      row[`获客_${p}`] = leads;
      row[`班型_${p}`] = classType;
      row[`班型得分_${p}`] = classScore;
      row[`添加产值_${p}`] = addValue === null ? null : round2(Number(addValue));
      row[`产值排名_${p}`] = rank;
      row[`产值排名%_${p}`] = rankPct === null ? null : round2(Number(rankPct));
      row[`产值得分_${p}`] = valueScore;

      total += (Number(classScore ?? 0) || 0) + (Number(valueScore ?? 0) || 0);
    }

    row["班长总分"] = total;
    row["带班数"] = total < 120 ? 1 : total < 150 ? 2 : total < 180 ? 3 : 4;
    leaders.push(row);
  }

  leaders.sort((a, b) => {
    const band = (Number(b["带班数"] ?? 0) || 0) - (Number(a["带班数"] ?? 0) || 0);
    if (band !== 0) return band;
    const score = (Number(b["班长总分"] ?? 0) || 0) - (Number(a["班长总分"] ?? 0) || 0);
    if (score !== 0) return score;
    const g = String(a["大组"]).localeCompare(String(b["大组"]), "zh");
    if (g !== 0) return g;
    const s = String(a["小组"]).localeCompare(String(b["小组"]), "zh");
    if (s !== 0) return s;
    return String(a["班长"]).localeCompare(String(b["班长"]), "zh");
  });

  return leaders;
}

function toCSV(rows, columns) {
  const lines = [];
  lines.push(columns.join(","));
  for (const r of rows) {
    const line = columns
      .map((c) => {
        const v = r[c];
        const s = v === null || v === undefined ? "" : String(v);
        const escaped = s.includes('"') || s.includes(",") || s.includes("\n") ? `"${s.replaceAll('"', '""')}"` : s;
        return escaped;
      })
      .join(",");
    lines.push(line);
  }
  return "\ufeff" + lines.join("\n");
}

function downloadText(filename, text) {
  const blob = new Blob([text], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function renderTable(tableEl, rows, columns, { searchInput, metaEl, hintEl }) {
  let sortKey = "带班数";
  let sortDir = -1;

  function normalizeForSort(v) {
    if (v === null || v === undefined) return { t: "null", v: "" };
    if (typeof v === "number" && Number.isFinite(v)) return { t: "num", v };
    const n = toNumber(v);
    if (n !== null) return { t: "num", v: n };
    return { t: "str", v: String(v) };
  }

  function headerText(c) {
    if (c === "产值排名%") return "产值排名百分位";
    return c;
  }

  function cellText(c, v) {
    if (v === null || v === undefined) return "";
    if (c === "产值排名%") {
      const n = typeof v === "number" && Number.isFinite(v) ? v : toNumber(v);
      if (n === null) return "";
      const out = round2(n);
      if (out === null) return "";
      return `${out.toFixed(2)}%`;
    }
    return String(v);
  }

  function apply(rows0) {
    const tokens = tokenizeQuery(searchInput?.value ?? "");
    let filtered = rows0;
    if (tokens.length > 0) {
      filtered = rows0.filter((r) => {
        for (const t of tokens) {
          for (const c of columns) {
            if (String(r[c] ?? "").toLowerCase().includes(t)) return true;
          }
        }
        return false;
      });
    }

    if (sortKey) {
      const key = sortKey;
      const dir = sortDir;
      filtered = [...filtered].sort((a, b) => {
        const av = normalizeForSort(a[key]);
        const bv = normalizeForSort(b[key]);
        let primary = 0;
        if (av.t !== bv.t) primary = av.t.localeCompare(bv.t) * dir;
        else if (av.t === "num") primary = (av.v - bv.v) * dir;
        else primary = String(av.v).localeCompare(String(bv.v), "zh") * dir;
        if (primary !== 0) return primary;
        const score = (Number(b["班长总分"] ?? 0) || 0) - (Number(a["班长总分"] ?? 0) || 0);
        if (score !== 0) return score;
        const g = String(a["大组"]).localeCompare(String(b["大组"]), "zh");
        if (g !== 0) return g;
        const s = String(a["小组"]).localeCompare(String(b["小组"]), "zh");
        if (s !== 0) return s;
        return String(a["班长"]).localeCompare(String(b["班长"]), "zh");
      });
    }

    tableEl.__filteredRows = filtered;
    metaEl.textContent = `显示 ${filtered.length} / ${rows0.length}`;
    hintEl.textContent = filtered.length === 0 ? "未匹配到数据，请调整搜索条件或重新上传文件。" : "";

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    for (const c of columns) {
      const th = document.createElement("th");
      th.textContent = headerText(c);
      th.addEventListener("click", () => {
        if (sortKey === c) sortDir = -sortDir;
        else {
          sortKey = c;
          sortDir = c === "带班数" || c === "班长总分" ? -1 : 1;
        }
        apply(rows0);
      });
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);

    const tbody = document.createElement("tbody");
    for (const r of filtered) {
      const tr = document.createElement("tr");
      for (const c of columns) {
        const td = document.createElement("td");
        const v = r[c];
        td.textContent = cellText(c, v);
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }

    tableEl.innerHTML = "";
    tableEl.appendChild(thead);
    tableEl.appendChild(tbody);
  }

  if (searchInput) searchInput.addEventListener("input", () => apply(rows));
  apply(rows);
}

function renderOverviewTable(tableEl, rows, { searchInput, metaEl, hintEl }) {
  const columns = [
    "大组",
    "小组",
    "班长",
    "产值得分_-1",
    "产值得分_-2",
    "产值得分_-3",
    "班型得分_-1",
    "班型得分_-2",
    "班型得分_-3",
    "班长总分",
    "带班数",
  ];
  const detailRows = [
    { label: "获客", key: "获客" },
    { label: "班型", key: "班型" },
    { label: "添加产值", key: "添加产值" },
    { label: "产值排名", key: "产值排名" },
    { label: "产值排名百分位", key: "产值排名%" },
  ];

  let sortKey = null;
  let sortDir = 1;

  const expanded = tableEl.__expandedKeys instanceof Set ? tableEl.__expandedKeys : new Set();
  tableEl.__expandedKeys = expanded;

  function normalizeForSort(v) {
    if (v === null || v === undefined) return { t: "null", v: "" };
    if (typeof v === "number" && Number.isFinite(v)) return { t: "num", v };
    const n = toNumber(v);
    if (n !== null) return { t: "num", v: n };
    return { t: "str", v: String(v) };
  }

  function valueText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "number" && Number.isFinite(v)) {
      if (Number.isInteger(v)) return String(v);
      return round2(v).toFixed(2);
    }
    const s = String(v);
    const n = toNumber(s);
    if (n !== null && s.includes(".")) return round2(n).toFixed(2);
    return s;
  }

  function valueTextByKey(key, v) {
    if (key.includes("产值排名%")) {
      const base = valueText(v);
      if (!base) return "";
      const n = toNumber(base);
      if (n === null) return "";
      return `${round2(n).toFixed(2)}%`;
    }
    return valueText(v);
  }

  function apply(rows0) {
    const tokens = tokenizeQuery(searchInput?.value ?? "");
    let filtered = rows0;
    if (tokens.length > 0) {
      filtered = rows0.filter((r) => {
        for (const t of tokens) {
          for (const c of ["大组", "小组", "班长"]) {
            if (String(r[c] ?? "").toLowerCase().includes(t)) return true;
          }
        }
        return false;
      });
    }

    if (sortKey) {
      const key = sortKey;
      const dir = sortDir;
      filtered = [...filtered].sort((a, b) => {
        const av = normalizeForSort(a[key]);
        const bv = normalizeForSort(b[key]);
        if (av.t !== bv.t) return av.t.localeCompare(bv.t) * dir;
        if (av.t === "num") return (av.v - bv.v) * dir;
        return String(av.v).localeCompare(String(bv.v), "zh") * dir;
      });
    }

    tableEl.__filteredRows = filtered;
    metaEl.textContent = `显示 ${filtered.length} / ${rows0.length}`;
    hintEl.textContent = filtered.length === 0 ? "未匹配到数据，请调整搜索条件或重新上传文件。" : "";

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    for (const c of columns) {
      const th = document.createElement("th");
      th.textContent = c;
      th.addEventListener("click", () => {
        if (sortKey === c) sortDir = -sortDir;
        else {
          sortKey = c;
          sortDir = 1;
        }
        apply(rows0);
      });
      headRow.appendChild(th);
    }
    const thMore = document.createElement("th");
    thMore.textContent = "详情";
    thMore.className = "no-sort";
    headRow.appendChild(thMore);
    thead.appendChild(headRow);

    const tbody = document.createElement("tbody");
    for (const r of filtered) {
      const key = String(r["班长"] ?? "");
      const isOpen = expanded.has(key);

      const tr = document.createElement("tr");
      for (const c of columns) {
        const td = document.createElement("td");
        const v = r[c];
        if (c === "带班数") {
          const span = document.createElement("span");
          const n = Number(v);
          span.className = `badge b${Number.isFinite(n) ? n : 0}`.trim();
          span.textContent = valueTextByKey(c, v);
          td.appendChild(span);
        } else {
          td.textContent = valueTextByKey(c, v);
          if (c.includes("得分") || c === "班长总分") td.className = "num";
        }
        tr.appendChild(td);
      }

      const tdMore = document.createElement("td");
      tdMore.className = "detail-toggle";
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = isOpen ? "收起" : "展开";
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        const cur = expanded.has(key);
        if (cur) expanded.delete(key);
        else expanded.add(key);
        apply(rows0);
      });
      tdMore.appendChild(btn);
      tr.appendChild(tdMore);
      tbody.appendChild(tr);

      if (isOpen) {
        const trDetail = document.createElement("tr");
        trDetail.className = "detail-row";
        const td = document.createElement("td");
        td.colSpan = columns.length + 1;

        const box = document.createElement("div");
        box.className = "detail-box";

        const sub = document.createElement("table");
        sub.className = "subtable";
        const subHead = document.createElement("thead");
        const subHeadTr = document.createElement("tr");
        ["指标", "-1", "-2", "-3"].forEach((h) => {
          const th = document.createElement("th");
          th.textContent = h;
          subHeadTr.appendChild(th);
        });
        subHead.appendChild(subHeadTr);
        sub.appendChild(subHead);

        const subBody = document.createElement("tbody");
        for (const item of detailRows) {
          const subTr = document.createElement("tr");
          const tdLabel = document.createElement("td");
          tdLabel.textContent = item.label;
          subTr.appendChild(tdLabel);
          for (const p of PERIODS) {
            const tdV = document.createElement("td");
            tdV.textContent = valueTextByKey(item.key, r[`${item.key}_${p}`]);
            subTr.appendChild(tdV);
          }
          subBody.appendChild(subTr);
        }
        sub.appendChild(subBody);

        box.appendChild(sub);
        td.appendChild(box);
        trDetail.appendChild(td);
        tbody.appendChild(trDetail);
      }
    }

    tableEl.innerHTML = "";
    tableEl.appendChild(thead);
    tableEl.appendChild(tbody);
  }

  if (searchInput) searchInput.addEventListener("input", () => apply(rows));
  apply(rows);
}

function setActiveTab(tabName) {
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === tabName));
  document.querySelectorAll(".tab-content").forEach((s) => s.classList.toggle("active", s.id === tabName));
}

function computeFromRawRows(rawRows) {
  for (const col of REQUIRED_COLUMNS) {
    if (!(col in rawRows[0])) throw new Error(`缺少字段：${col}`);
  }

  const normalized = rawRows.map((r) => {
    const term = toNumber(r["营期"]);
    const leaderRaw = r["班级"];
    const leader = cleanLeaderName(leaderRaw);
    const group = String(r["小组"] ?? "").trim() || "暂无小组";
    const data = {
      营期: term,
      班长: leader,
      小组: group,
      获客: toNumber(r["获客"]) ?? 0,
      流水: toNumber(r["流水"]) ?? 0,
      添加人数: toNumber(r["添加人数"]) ?? 0,
    };
    return data;
  });

  const grouped = groupBy(normalized, (r) => `${r["营期"]}||${r["班长"]}`);
  const aggregated = [];
  for (const items of grouped.values()) {
    const base = items[0];
    const sum = (k) => items.reduce((acc, it) => acc + (Number(it[k] ?? 0) || 0), 0);
    aggregated.push({
      营期: base["营期"],
      班长: base["班长"],
      小组: modeOrFirst(items.map((x) => x["小组"]), "暂无小组"),
      获客: sum("获客"),
      流水: sum("流水"),
      添加人数: sum("添加人数"),
    });
  }

  for (const r of aggregated) r["大组"] = GROUP_MAP.get(r["小组"]) ?? (r["小组"] === "暂无小组" ? "未分组" : "未分组");

  const withPeriods = assignLatestPeriodByLeader(aggregated);

  for (const r of withPeriods) {
    const add = Number(r["添加人数"]);
    const flow = Number(r["流水"]);
    r["添加产值"] = add > 0 && Number.isFinite(flow) ? flow / add : null;
    r["班型"] = roundHalfUp(Number(r["获客"]) / 250);
    r["班型得分"] = r["班型"] === null ? null : r["班型"] * 10;
  }

  const withRank = assignRankAndScore(withPeriods);
  const evidence = withRank
    .map((r) => ({
      营期: r["营期"],
      班长: r["班长"],
      大组: r["大组"],
      小组: r["小组"],
      获客: r["获客"],
      流水: r["流水"],
      添加人数: r["添加人数"],
      添加产值: r["添加产值"] === null ? null : round2(r["添加产值"]),
      最新营期: r["最新营期"],
      产值排名: r["产值排名"],
      "产值排名%": r["产值排名%"],
      产值得分: r["产值得分"],
      班型: r["班型"],
      班型得分: r["班型得分"],
    }))
    .sort((a, b) => {
      const t = (a["营期"] ?? 0) - (b["营期"] ?? 0);
      if (t !== 0) return t;
      return String(a["班长"]).localeCompare(String(b["班长"]), "zh");
    });

  const overview = buildWideTable(withRank);

  return { overview, evidence };
}

function initTabs() {
  document.querySelectorAll(".tab").forEach((b) =>
    b.addEventListener("click", () => setActiveTab(b.dataset.tab)),
  );
}

function initApp() {
  initTabs();

  const fileInput = qs("#fileInput");
  const dropzone = qs("#dropzone");
  const fileMeta = qs("#fileMeta");
  const confirmClosed = qs("#confirmClosed");
  const statusPill = qs("#statusPill");
  const summaryChips = qs("#summaryChips");
  const bandFilter = qs("#bandFilter");
  const clearData = qs("#clearData");

  const overviewTable = qs("#overviewTable");
  const evidenceTable = qs("#evidenceTable");

  const overviewSearch = qs("#overviewSearch");
  const evidenceSearch = qs("#evidenceSearch");

  const overviewMeta = qs("#overviewMeta");
  const evidenceMeta = qs("#evidenceMeta");

  const overviewHint = qs("#overviewHint");
  const evidenceHint = qs("#evidenceHint");

  const downloadResult = qs("#downloadResult");
  const downloadEvidence = qs("#downloadEvidence");

  let currentOverview = [];
  let currentEvidence = [];
  let activeBand = null;
  let uploadEnabled = false;

  function setStatus(kind, text) {
    statusPill.className = `pill ${kind ?? ""}`.trim();
    statusPill.textContent = text;
  }

  function setDownloadsEnabled(enabled) {
    downloadResult.disabled = !enabled;
    downloadEvidence.disabled = !enabled;
    clearData.disabled = !enabled;
  }

  function setUploadEnabled(enabled) {
    uploadEnabled = enabled;
    fileInput.disabled = !enabled;
    dropzone.classList.toggle("disabled", !enabled);
    dropzone.setAttribute("aria-disabled", String(!enabled));
    dropzone.tabIndex = enabled ? 0 : -1;
  }

  function renderAll() {
    const filteredOverview =
      activeBand === null ? currentOverview : currentOverview.filter((r) => Number(r["带班数"]) === activeBand);

    renderOverviewTable(overviewTable, filteredOverview, {
      searchInput: overviewSearch,
      metaEl: overviewMeta,
      hintEl: overviewHint,
    });

    const evidenceColumns = [
      "营期",
      "班长",
      "大组",
      "小组",
      "获客",
      "流水",
      "添加人数",
      "添加产值",
      "最新营期",
      "产值排名",
      "产值排名%",
      "产值得分",
      "班型",
      "班型得分",
    ];

    renderTable(evidenceTable, currentEvidence, evidenceColumns, {
      searchInput: evidenceSearch,
      metaEl: evidenceMeta,
      hintEl: evidenceHint,
    });
  }

  function renderBandFilter() {
    const counts = new Map();
    for (const r of currentOverview) {
      const b = Number(r["带班数"]);
      if (!Number.isFinite(b)) continue;
      counts.set(b, (counts.get(b) ?? 0) + 1);
    }

    const allChip = { label: `全部 (${currentOverview.length})`, value: null };
    const options = [4, 3, 2, 1].map((b) => ({ label: `${b} (${counts.get(b) ?? 0})`, value: b }));
    const chips = [allChip, ...options];

    bandFilter.innerHTML = "";
    for (const c of chips) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `chip ${activeBand === c.value ? "active" : ""}`.trim();
      btn.textContent = c.label;
      btn.addEventListener("click", () => {
        activeBand = c.value;
        renderAll();
        renderBandFilter();
      });
      bandFilter.appendChild(btn);
    }

    summaryChips.innerHTML = "";
    for (const b of [4, 3, 2, 1]) {
      const el = document.createElement("div");
      el.className = "pill";
      el.textContent = `带班数 ${b}: ${counts.get(b) ?? 0}`;
      summaryChips.appendChild(el);
    }
  }

  function resetUI() {
    currentOverview = [];
    currentEvidence = [];
    activeBand = null;
    overviewSearch.value = "";
    evidenceSearch.value = "";
    overviewTable.innerHTML = "";
    evidenceTable.innerHTML = "";
    overviewMeta.textContent = "";
    evidenceMeta.textContent = "";
    summaryChips.innerHTML = "";
    bandFilter.innerHTML = "";
    fileMeta.textContent = "未上传（请先勾选确认）";
    confirmClosed.checked = false;
    setUploadEnabled(false);
    setStatus("warn", "请先确认已结营");
    setDownloadsEnabled(false);
    overviewHint.textContent = "上传后将展示：组-班长 + 三期两维度得分 + 总分与带班数；其余过程数据可点击“展开”或切换到“过程数据”。";
    evidenceHint.textContent = "";
  }

  downloadResult.addEventListener("click", () => {
    const rows = Array.isArray(overviewTable.__filteredRows) ? overviewTable.__filteredRows : currentOverview;
    if (!rows.length) return;
    const columns = Object.keys(rows[0]);
    downloadText("班长排班得分-结果.csv", toCSV(rows, columns));
  });

  downloadEvidence.addEventListener("click", () => {
    const rows = Array.isArray(evidenceTable.__filteredRows) ? evidenceTable.__filteredRows : currentEvidence;
    if (!rows.length) return;
    const columns = Object.keys(rows[0]);
    downloadText("班长排班得分-过程数据.csv", toCSV(rows, columns));
  });

  async function handleFile(file) {
    if (!uploadEnabled) return;
    if (!file) return;
    setDownloadsEnabled(false);
    setStatus("warn", "计算中");
    fileMeta.textContent = `${file.name}`;
    overviewHint.textContent = "正在读取并计算，请稍候…";
    evidenceHint.textContent = "";
    try {
      const text = await decodeFileToText(file);
      const { rows, headers } = parseTSV(text);
      for (const c of REQUIRED_COLUMNS) {
        if (!headers.includes(c)) throw new Error(`上传文件缺少字段：${c}`);
      }
      if (rows.length === 0) throw new Error("文件无数据行");
      const { overview, evidence } = computeFromRawRows(rows);
      currentOverview = overview;
      currentEvidence = evidence;
      activeBand = null;
      renderBandFilter();
      renderAll();
      setDownloadsEnabled(true);
      setStatus("ok", "已完成");
      overviewHint.textContent = "已完成计算。总览默认按带班数与总分从高到低排序；点击“展开”可查看班型/产值等佐证数据。";
      evidenceHint.textContent = "可通过搜索框按营期/班长快速定位并导出。";
      setActiveTab("overview");
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatus("bad", "计算失败");
      overviewHint.textContent = `计算失败：${msg}`;
      evidenceHint.textContent = "";
      currentOverview = [];
      currentEvidence = [];
      summaryChips.innerHTML = "";
      bandFilter.innerHTML = "";
      overviewTable.innerHTML = "";
      evidenceTable.innerHTML = "";
      overviewMeta.textContent = "";
      evidenceMeta.textContent = "";
      setDownloadsEnabled(false);
    } finally {
      fileInput.value = "";
    }
  }

  clearData.addEventListener("click", () => resetUI());

  confirmClosed.addEventListener("change", () => {
    if (confirmClosed.checked) {
      setUploadEnabled(true);
      fileMeta.textContent = "未上传";
      setStatus("", "等待上传");
    } else {
      setUploadEnabled(false);
      fileMeta.textContent = "未上传（请先勾选确认）";
      setStatus("warn", "请先确认已结营");
    }
  });

  fileInput.addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    await handleFile(file);
  });

  dropzone.addEventListener("click", () => {
    if (!uploadEnabled) return;
    fileInput.click();
  });
  dropzone.addEventListener("keydown", (e) => {
    if (!uploadEnabled) return;
    if (e.key === "Enter" || e.key === " ") fileInput.click();
  });
  dropzone.addEventListener("dragover", (e) => {
    if (!uploadEnabled) return;
    e.preventDefault();
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
  dropzone.addEventListener("drop", async (e) => {
    if (!uploadEnabled) return;
    e.preventDefault();
    dropzone.classList.remove("dragover");
    const file = e.dataTransfer?.files?.[0];
    await handleFile(file);
  });

  resetUI();
}

initApp();
