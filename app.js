const REQUIRED_COLUMNS = ["项目", "营期", "开营时间", "班级", "小组", "大组", "获客", "流水", "添加人数"];
const PERIODS = [-1, -2, -3];
const PIP_PERIODS = [-1, -2, -3, -4, -5, -6, -7, -8];

function qs(sel) {
  const el = document.querySelector(sel);
  if (!el) throw new Error(`缺少元素: ${sel}`);
  return el;
}

function tokenizeQuery(input) {
  const s = String(input ?? "").trim().toLowerCase();
  if (!s) return [];
  const parts = s.match(/[^\s,，、;；]+/g) ?? [];
  const seen = new Set();
  const out = [];
  for (const p of parts) {
    const t = String(p ?? "").trim();
    if (!t) continue;
    if (seen.has(t)) continue;
    seen.add(t);
    out.push(t);
  }
  return out;
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

function debounce(fn, waitMs) {
  let t = null;
  return (...args) => {
    if (t) clearTimeout(t);
    t = setTimeout(() => fn(...args), waitMs);
  };
}

function cleanLeaderName(name, projectType) {
  const s = String(name ?? "").trim();
  if (!s) return "";
  if (projectType === "ziyang") {
    // 梓洋项目：删除括号及连字符（-、_ 等）及其后的内容
    return s.replace(/[（(\-_].*$/, "").trim();
  }
  // 其他项目：仅删除括号及其后的内容
  return s.replace(/[（(].*$/, "").trim();
}

function parseDate(s) {
  if (!s) return null;
  let cleaned = String(s).trim();
  if (!cleaned) return null;
  
  // 处理常见的日期格式：2026/4/6, 2026-4-6, 2026年4月6日
  cleaned = cleaned.replace(/年|月/g, "/").replace(/日/g, "");
  cleaned = cleaned.replace(/-/g, "/");
  
  // 如果格式是 "4/6" 或 "04/06"，自动补全当前年份
  const parts = cleaned.split("/");
  if (parts.length === 2) {
    const currentYear = new Date().getFullYear();
    cleaned = `${currentYear}/${cleaned}`;
  }
  
  const d = new Date(cleaned);
  return isNaN(d.getTime()) ? null : d;
}

function decodeFileToText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error("读取文件失败"));
    reader.onload = () => {
      const buf = reader.result;
      const view = new Uint8Array(buf);
      
      let encoding = "utf-8";
      // 简单的 UTF-16LE 检测 (BOM: FF FE)
      if (view.length >= 2 && view[0] === 0xff && view[1] === 0xfe) {
        encoding = "utf-16le";
      } else if (view.length >= 2 && view[0] === 0xfe && view[1] === 0xff) {
        encoding = "utf-16be";
      }
      
      try {
        const text = new TextDecoder(encoding).decode(buf);
        resolve(text);
      } catch (e) {
        reject(e);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function parseTSV(text0) {
  // 去除 BOM
  const text = text0.startsWith("\ufeff") ? text0.slice(1) : text0;
  const lines = text.split(/\r\n|\n|\r/).filter((l) => l.trim() !== "");
  if (lines.length === 0) return { rows: [], headers: [] };
  
  // 自动检测分隔符：取第一行看逗号多还是制表符多
  const firstLine = lines[0];
  const tabCount = (firstLine.match(/\t/g) || []).length;
  const commaCount = (firstLine.match(/,/g) || []).length;
  const sep = tabCount >= commaCount ? "\t" : ",";
  
  const headers = firstLine.split(sep).map((h) => h.trim().replace(/^["']|["']$/g, ""));
  const rows = [];
  for (let i = 1; i < lines.length; i += 1) {
    const cols = lines[i].split(sep);
    const obj = {};
    for (let j = 0; j < headers.length; j += 1) {
      let val = (cols[j] ?? "").trim();
      // 去除引号
      val = val.replace(/^["']|["']$/g, "");
      obj[headers[j]] = val;
    }
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

function yesNo(v) {
  if (v === true) return "是";
  if (v === false) return "否";
  return "";
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

function buildWideTable(rows, projectType, probationSet) {
  const byLeader = groupBy(rows.filter((r) => PIP_PERIODS.includes(r["最新营期"])), (r) => r["班长"]);
  const leaders = [];

  for (const [leader, items] of byLeader.entries()) {
    const row = {
      班长: leader,
      大组: "",
      小组: "",
    };
    
    if (projectType === "ziyang") {
      row["试用期"] = probationSet.has(leader) ? "是" : "";
    }

    const pNeg1 = items.find((x) => x["最新营期"] === -1);
    row["小组"] = String(pNeg1?.["小组"] ?? "暂无小组");
    row["大组"] = String(pNeg1?.["大组"] ?? "未分组");

    for (const p of PERIODS) {
      const it = items.find((x) => x["最新营期"] === p);
      const leads = it?.["获客"] ?? null;
      const classType = it?.["班型"] ?? null;
      const classScore = it?.["班型得分"] ?? null;
      const addValue = it?.["添加产值"] ?? null;
      const market = it?.["大盘产值"] ?? null;
      const below = it?.["低于大盘"] ?? null;
      const rank = it?.["产值排名"] ?? null;
      const rankPct = it?.["产值排名%"] ?? null;
      const valueScore = it?.["产值得分"] ?? 0;

      row[`营期_${p}`] = it?.["营期"] ?? null;
      row[`获客_${p}`] = leads;
      row[`班型_${p}`] = classType;
      row[`添加产值_${p}`] = addValue === null ? null : round2(Number(addValue));
      row[`大盘产值_${p}`] = market === null ? null : round2(Number(market));
      row[`低于大盘_${p}`] = yesNo(below);
      row[`大盘差值_${p}`] =
        addValue === null || market === null ? null : round2(Number(addValue) - Number(market));
      row[`产值排名_${p}`] = rank;
      row[`产值排名%_${p}`] = rankPct === null ? null : round2(Number(rankPct));
      row[`产值_${p}`] = valueScore;
      row[`班型_${p}`] = classScore;
    }

    const belowByPeriod = new Map();
    for (const p of PIP_PERIODS) {
      const it = items.find((x) => x["最新营期"] === p);
      const below = it?.["低于大盘"];
      belowByPeriod.set(p, below === true);
    }
    
    let pip = false;
    let eliminated = false;

    if (projectType === "ziyang") {
      const isProbation = probationSet.has(leader);
      if (isProbation) {
        // 试用期员工：近5个营期，2期产值低大盘则标记PIP，3期产值低大盘则标记淘汰；
        const belowCount5 = [-1, -2, -3, -4, -5].map(p => belowByPeriod.get(p)).filter(Boolean).length;
        if (belowCount5 >= 3) eliminated = true;
        else if (belowCount5 >= 2) pip = true;
      } else {
        // 转正员工：近8个营期，3期产值低大盘则标记PIP，5期产值低大盘则标记淘汰；
        const belowCount8 = PIP_PERIODS.map(p => belowByPeriod.get(p)).filter(Boolean).length;
        if (belowCount8 >= 5) eliminated = true;
        else if (belowCount8 >= 3) pip = true;
      }
    } else {
      const b1 = belowByPeriod.get(-1) === true;
      const b2 = belowByPeriod.get(-2) === true;
      const b3 = belowByPeriod.get(-3) === true;
      const b4 = belowByPeriod.get(-4) === true;
      const count3 = [b1, b2, b3].filter(Boolean).length;
      const count4 = [b1, b2, b3, b4].filter(Boolean).length;
      pip = (b1 && b2) || count3 >= 2;
      eliminated = b1 && count4 >= 3;
    }

    row["PIP"] = pip ? "PIP" : "";
    row["淘汰"] = eliminated ? "淘汰" : "";

    const availablePeriods = new Set(items.map((x) => x["最新营期"]).filter((x) => x !== null));
    if (availablePeriods.size === 2 && availablePeriods.has(-1) && availablePeriods.has(-2) && !availablePeriods.has(-3)) {
      const a1 = Number(row["产值_-1"] ?? 0) || 0;
      const a2 = Number(row["产值_-2"] ?? 0) || 0;
      const b1 = Number(row["班型_-1"] ?? 0) || 0;
      const b2 = Number(row["班型_-2"] ?? 0) || 0;
      row["产值_-3"] = round2((a1 + a2) / 2);
      row["班型_-3"] = round2((b1 + b2) / 2);
    }

    let total = 0;
    for (const p of PERIODS) {
      total += (Number(row[`班型_${p}`] ?? 0) || 0) + (Number(row[`产值_${p}`] ?? 0) || 0);
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

function buildPipTable(rows, projectType, probationSet) {
  const byLeader = groupBy(rows.filter((r) => PIP_PERIODS.includes(r["最新营期"])), (r) => r["班长"]);
  const leaders = [];

  for (const [leader, items] of byLeader.entries()) {
    const pNeg1 = items.find((x) => x["最新营期"] === -1);
    const row = {
      班长: leader,
      大组: String(pNeg1?.["大组"] ?? "未分组"),
      小组: String(pNeg1?.["小组"] ?? "暂无小组"),
    };
    
    if (projectType === "ziyang") {
      row["试用期"] = probationSet.has(leader) ? "是" : "";
    }

    const belowFlags = [];
    for (const p of PIP_PERIODS) {
      const it = items.find((x) => x["最新营期"] === p);
      const term = it?.["营期"] ?? null;
      const addValue = it?.["添加产值"] ?? null;
      const market = it?.["大盘产值"] ?? null;
      const below = it?.["低于大盘"] ?? null;
      row[`营期_${p}`] = term;
      row[`添加产值_${p}`] = addValue === null ? null : round2(Number(addValue));
      row[`大盘产值_${p}`] = market === null ? null : round2(Number(market));
      row[`低于大盘_${p}`] = yesNo(below);
      belowFlags.push(below === true);
    }
    
    let pip = false;
    let eliminated = false;

    if (projectType === "ziyang") {
      const isProbation = probationSet.has(leader);
      if (isProbation) {
        // 试用期员工：近5个营期，2期产值低大盘则标记PIP，3期产值低大盘则标记淘汰；
        const belowCount5 = belowFlags.slice(0, 5).filter(Boolean).length;
        if (belowCount5 >= 3) eliminated = true;
        else if (belowCount5 >= 2) pip = true;
      } else {
        // 转正员工：近8个营期，3期产值低大盘则标记PIP，5期产值低大盘则标记淘汰；
        const belowCount8 = belowFlags.filter(Boolean).length;
        if (belowCount8 >= 5) eliminated = true;
        else if (belowCount8 >= 3) pip = true;
      }
    } else {
      const b1 = belowFlags[0] === true;
      const b2 = belowFlags[1] === true;
      const b3 = belowFlags[2] === true;
      const count3 = [b1, b2, b3].filter(Boolean).length;
      pip = (b1 && b2) || count3 >= 2;
      eliminated = b1 && belowFlags.slice(0, 4).filter(Boolean).length >= 3;
    }

    row["PIP"] = pip ? "PIP" : "";
    row["淘汰"] = eliminated ? "淘汰" : "";
    leaders.push(row);
  }

  leaders.sort((a, b) => {
    const elim = String(b["淘汰"] ?? "").localeCompare(String(a["淘汰"] ?? ""), "zh");
    if (elim !== 0) return elim;
    const pip = String(b["PIP"] ?? "").localeCompare(String(a["PIP"] ?? ""), "zh");
    if (pip !== 0) return pip;
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

function renderTable(tableEl, rows, columns, { searchInput, metaEl, hintEl, searchColumns }) {
  let sortKey = columns.includes("带班数") ? "带班数" : null;
  let sortDir = sortKey ? -1 : 1;

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
      const cols = Array.isArray(searchColumns) && searchColumns.length > 0 ? searchColumns : columns;
      const cacheKey = `__searchText__${cols.join("|")}`;
      filtered = rows0.filter((r) => {
        const hay =
          typeof r[cacheKey] === "string"
            ? r[cacheKey]
            : (r[cacheKey] = cols.map((c) => String(r[c] ?? "")).join("\u0001").toLowerCase());
        for (const t of tokens) {
          if (hay.includes(t)) return true;
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
        
        if (c === "试用期" && String(v ?? "") === "是") {
          const span = document.createElement("span");
          span.className = "badge probation";
          span.textContent = "试用期";
          td.appendChild(span);
        } else {
          td.textContent = cellText(c, v);
        }
        
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }

    tableEl.innerHTML = "";
    tableEl.appendChild(thead);
    tableEl.appendChild(tbody);
  }

  if (searchInput) {
    const prev = searchInput.__boundApply;
    if (typeof prev === "function") searchInput.removeEventListener("input", prev);
    const next = debounce(() => apply(rows), 120);
    searchInput.__boundApply = next;
    searchInput.addEventListener("input", next);
  }
  apply(rows);
}

function renderOverviewTable(tableEl, rows, { searchInput, metaEl, hintEl, projectType }) {
  const columns = [
    "大组",
    "小组",
    "班长",
    "产值_-1",
    "产值_-2",
    "产值_-3",
    "班型_-1",
    "班型_-2",
    "班型_-3",
    "班长总分",
    "带班数",
    ...(projectType === "ziyang" ? ["试用期"] : []),
    "PIP",
    "淘汰",
  ];
  const detailRows = [
    { label: "营期", key: "营期" },
    { label: "获客", key: "获客" },
    { label: "班型", key: "班型" },
    { label: "添加产值", key: "添加产值" },
    { label: "大盘产值", key: "大盘产值" },
    { label: "低于大盘", key: "低于大盘" },
    { label: "与大盘差值", key: "大盘差值" },
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
    if (key === "大盘差值") {
      const n = typeof v === "number" && Number.isFinite(v) ? v : toNumber(v);
      if (n === null) return "";
      const out = round2(n);
      if (out === null) return "";
      const sign = out > 0 ? "+" : "";
      return `${sign}${out.toFixed(2)}`;
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
        } else if (c === "试用期" && String(v ?? "") === "是") {
          const span = document.createElement("span");
          span.className = "badge probation";
          span.textContent = "试用期";
          td.appendChild(span);
        } else if (c === "PIP" && String(v ?? "") === "PIP") {
          const span = document.createElement("span");
          span.className = "badge pip";
          span.textContent = "PIP";
          td.appendChild(span);
        } else if (c === "淘汰" && String(v ?? "") === "淘汰") {
          const span = document.createElement("span");
          span.className = "badge elim";
          span.textContent = "淘汰";
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

        const cardsContainer = document.createElement("div");
        cardsContainer.className = "period-cards";

        for (const p of PERIODS) {
          const card = document.createElement("div");
          card.className = "period-card";
          
          const header = document.createElement("div");
          header.className = "period-card-header";
          
          const termVal = valueTextByKey("营期", r[`营期_${p}`]);
          header.innerHTML = `<span class="period-badge">近${Math.abs(p)}期</span><span class="period-term">${termVal ? `营期 ${termVal}` : "暂无营期"}</span>`;
          card.appendChild(header);

          const body = document.createElement("div");
          body.className = "period-card-body";

          for (const item of detailRows) {
            if (item.key === "营期") continue;
            const rowDiv = document.createElement("div");
            rowDiv.className = "period-stat";
            
            const lbl = document.createElement("span");
            lbl.className = "stat-label";
            lbl.textContent = item.label;

            const val = document.createElement("span");
            val.className = "stat-value";
            const valText = valueTextByKey(item.key, r[`${item.key}_${p}`]);
            val.textContent = valText || "-";

            if (item.key === "低于大盘" && valText === "是") {
              val.style.color = "var(--danger-text)";
              val.style.fontWeight = "600";
            } else if (item.key === "低于大盘" && valText === "否") {
              val.style.color = "var(--success-text)";
              val.style.fontWeight = "600";
            } else if (item.key === "大盘差值") {
              const raw = r[`${item.key}_${p}`];
              const n = typeof raw === "number" && Number.isFinite(raw) ? raw : toNumber(raw);
              if (n !== null) {
                if (n > 0) {
                  val.style.color = "var(--success-text)";
                  val.style.fontWeight = "600";
                } else if (n < 0) {
                  val.style.color = "var(--danger-text)";
                  val.style.fontWeight = "600";
                } else {
                  val.style.color = "var(--text-muted)";
                }
              }
            }

            rowDiv.appendChild(lbl);
            rowDiv.appendChild(val);
            body.appendChild(rowDiv);
          }
          card.appendChild(body);
          cardsContainer.appendChild(card);
        }

        box.appendChild(cardsContainer);
        td.appendChild(box);
        trDetail.appendChild(td);
        tbody.appendChild(trDetail);
      }
    }

    tableEl.innerHTML = "";
    tableEl.appendChild(thead);
    tableEl.appendChild(tbody);
  }

  if (searchInput) {
    const prev = searchInput.__boundApply;
    if (typeof prev === "function") searchInput.removeEventListener("input", prev);
    const next = debounce(() => apply(rows), 120);
    searchInput.__boundApply = next;
    searchInput.addEventListener("input", next);
  }
  apply(rows);
}

function renderPipTable(tableEl, rows, { searchInput, metaEl, hintEl, projectType }) {
  const dynamicBelowCols = PIP_PERIODS.map(p => `低于大盘_${p}`);
  const columns = [
    "大组", 
    "小组", 
    ...(projectType === "ziyang" ? ["试用期"] : []),
    "班长", 
    "PIP", 
    "淘汰", 
    ...(projectType === "ziyang" ? dynamicBelowCols : ["低于大盘_-1", "低于大盘_-2", "低于大盘_-3", "低于大盘_-4"])
  ];
  const detailRows = [
    { label: "营期", key: "营期" },
    { label: "添加产值", key: "添加产值" },
    { label: "大盘产值", key: "大盘产值" },
    { label: "低于大盘", key: "低于大盘" },
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

  function apply(rows0) {
    const tokens = tokenizeQuery(searchInput?.value ?? "");
    let filtered = rows0;
    if (tokens.length > 0) {
      filtered = rows0.filter((r) => {
        const hay = [r["大组"], r["小组"], r["班长"], r["PIP"], r["淘汰"]].map((x) => String(x ?? "")).join("\u0001").toLowerCase();
        for (const t of tokens) if (hay.includes(t)) return true;
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
    } else {
      filtered = [...filtered].sort((a, b) => {
        const elim = String(b["淘汰"] ?? "").localeCompare(String(a["淘汰"] ?? ""), "zh");
        if (elim !== 0) return elim;
        const pip = String(b["PIP"] ?? "").localeCompare(String(a["PIP"] ?? ""), "zh");
        if (pip !== 0) return pip;
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
        if (c === "PIP" && String(v ?? "") === "PIP") {
          const span = document.createElement("span");
          span.className = "badge pip";
          span.textContent = "PIP";
          td.appendChild(span);
        } else if (c === "试用期" && String(v ?? "") === "是") {
          const span = document.createElement("span");
          span.className = "badge probation";
          span.textContent = "试用期";
          td.appendChild(span);
        } else if (c === "淘汰" && String(v ?? "") === "淘汰") {
          const span = document.createElement("span");
          span.className = "badge elim";
          span.textContent = "淘汰";
          td.appendChild(span);
        } else if (c.startsWith("低于大盘_")) {
          const s = String(v ?? "");
          if (s === "是" || s === "否") {
            const span = document.createElement("span");
            span.className = `badge ${s === "是" ? "low-yes" : "low-no"}`.trim();
            span.textContent = s;
            td.appendChild(span);
          } else td.textContent = "";
        } else td.textContent = valueText(v);
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

        const cardsContainer = document.createElement("div");
        cardsContainer.className = "period-cards";
        const pList = projectType === "ziyang" ? PIP_PERIODS : [-1, -2, -3, -4];

        for (const p of pList) {
          const card = document.createElement("div");
          card.className = "period-card";
          
          const header = document.createElement("div");
          header.className = "period-card-header";
          
          const termVal = valueText(r[`营期_${p}`]);
          header.innerHTML = `<span class="period-badge">近${Math.abs(p)}期</span><span class="period-term">${termVal ? `营期 ${termVal}` : "暂无营期"}</span>`;
          card.appendChild(header);

          const body = document.createElement("div");
          body.className = "period-card-body";

          for (const item of detailRows) {
            if (item.key === "营期") continue;
            const rowDiv = document.createElement("div");
            rowDiv.className = "period-stat";
            
            const lbl = document.createElement("span");
            lbl.className = "stat-label";
            lbl.textContent = item.label;

            const val = document.createElement("span");
            val.className = "stat-value";
            const valText = valueText(r[`${item.key}_${p}`]);
            val.textContent = valText || "-";

            if (item.key === "低于大盘" && valText === "是") {
              val.style.color = "var(--danger-text)";
              val.style.fontWeight = "600";
            } else if (item.key === "低于大盘" && valText === "否") {
              val.style.color = "var(--success-text)";
              val.style.fontWeight = "600";
            }

            rowDiv.appendChild(lbl);
            rowDiv.appendChild(val);
            body.appendChild(rowDiv);
          }
          card.appendChild(body);
          cardsContainer.appendChild(card);
        }

        box.appendChild(cardsContainer);
        td.appendChild(box);
        trDetail.appendChild(td);
        tbody.appendChild(trDetail);
      }
    }

    tableEl.innerHTML = "";
    tableEl.appendChild(thead);
    tableEl.appendChild(tbody);
  }

  if (searchInput) {
    const prev = searchInput.__boundApply;
    if (typeof prev === "function") searchInput.removeEventListener("input", prev);
    const next = debounce(() => apply(rows), 120);
    searchInput.__boundApply = next;
    searchInput.addEventListener("input", next);
  }
  apply(rows);
}

function setActiveTab(tabName) {
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === tabName));
  document.querySelectorAll(".tab-content").forEach((s) => s.classList.toggle("active", s.id === tabName));
}

function computeFromRawRows(rawRows, analysisDateStr, projectType, probationSet) {
  for (const col of REQUIRED_COLUMNS) {
    if (!(col in (rawRows[0] || {}))) throw new Error(`缺少字段：${col}`);
  }

  const analysisDate = parseDate(analysisDateStr) || new Date();
  const sevenDaysMs = 7 * 24 * 60 * 60 * 1000;

  const filteredRawRows = rawRows.filter((r) => {
    const project = String(r["项目"] ?? "").trim();
    const term = toNumber(r["营期"]);
    if (term === null) return false;

    if (projectType === "ziyang") {
      const klass = String(r["班级"] ?? "");
      if (klass.includes("非标-类目一")) return false;
    }
    
    // 兼容 V2.0 逻辑：过滤掉英语项目的第 10 期
    if (project === "英语" && term === 10) return false;
    
    // 营期过滤逻辑：仅保留分析当日已结营的营期
    const startDateRaw = r["开营时间"];
    const startDate = parseDate(startDateRaw);
    if (!startDate) {
      console.warn(`无法解析日期: ${startDateRaw}, 营期: ${term}`);
      return false;
    }
    const endDate = new Date(startDate.getTime() + sevenDaysMs);
    if (endDate > analysisDate) return false;

    return true;
  });

  if (filteredRawRows.length === 0) {
    const dates = Array.from(new Set(rawRows.map(r => r["开营时间"]).filter(Boolean))).slice(0, 3).join(", ");
    throw new Error(`过滤后无有效数据。请检查“分析当日日期”设置。
      (检测到的开营时间示例: ${dates || "无"}，当前分析日期: ${analysisDate.toLocaleDateString()})`);
  }

  const normalized = filteredRawRows.map((r) => {
    const term = toNumber(r["营期"]);
    const leaderRaw = r["班级"];
    const leader = cleanLeaderName(leaderRaw, projectType);
    const group = String(r["小组"] ?? "").trim() || "暂无小组";
    const bigGroup = String(r["大组"] ?? "").trim() || "未分组";
    const data = {
      营期: term,
      班长: leader,
      小组: group,
      大组: bigGroup,
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
      大组: modeOrFirst(items.map((x) => x["大组"]), "未分组"),
      获客: sum("获客"),
      流水: sum("流水"),
      添加人数: sum("添加人数"),
    });
  }

  const marketByTerm = new Map();
  for (const [term, items] of groupBy(aggregated, (r) => r["营期"]).entries()) {
    const sumFlow = items.reduce((acc, it) => acc + (Number(it["流水"] ?? 0) || 0), 0);
    const sumAdd = items.reduce((acc, it) => acc + (Number(it["添加人数"] ?? 0) || 0), 0);
    marketByTerm.set(term, sumAdd > 0 ? sumFlow / sumAdd : null);
  }

  const withPeriods = assignLatestPeriodByLeader(aggregated);

  for (const r of withPeriods) {
    const add = Number(r["添加人数"]);
    const flow = Number(r["流水"]);
    r["添加产值"] = add > 0 && Number.isFinite(flow) ? flow / add : null;
    r["班型"] = roundHalfUp(Number(r["获客"]) / 250);
    r["班型得分"] = r["班型"] === null ? null : r["班型"] * 10;
    r["大盘产值"] = marketByTerm.get(r["营期"]) ?? null;
    r["低于大盘"] =
      r["添加产值"] !== null && r["大盘产值"] !== null ? Number(r["添加产值"]) < Number(r["大盘产值"]) : null;
  }

  const withRank = assignRankAndScore(withPeriods);
  const evidence = withRank
    .map((r) => ({
      营期: r["营期"],
      班长: r["班长"],
      大组: r["大组"],
      小组: r["小组"],
      ...(projectType === "ziyang" ? { 试用期: probationSet.has(r["班长"]) ? "是" : "" } : {}),
      获客: r["获客"],
      流水: r["流水"],
      添加人数: r["添加人数"],
      添加产值: r["添加产值"] === null ? null : round2(r["添加产值"]),
      大盘产值: r["大盘产值"] === null ? null : round2(r["大盘产值"]),
      低于大盘: yesNo(r["低于大盘"]),
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

  const overview = buildWideTable(withRank, projectType, probationSet);
  const pip = buildPipTable(withRank, projectType, probationSet);

  return { overview, evidence, pip };
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
  const statusPill = qs("#statusPill");
  const summaryChips = qs("#summaryChips");
  const bandFilter = qs("#bandFilter");
  const clearData = qs("#clearData");
  const analysisDateInput = qs("#analysisDate");

  // 项目类型选择
  document.querySelectorAll('input[name="projectType"]').forEach((radio) => {
    radio.addEventListener("change", (e) => {
      currentProjectType = e.target.value;
      loadFromCache();
    });
  });

  // 设置默认分析日期为今天
  const today = new Date().toISOString().split("T")[0];
  analysisDateInput.value = today;

  const overviewTable = qs("#overviewTable");
  const evidenceTable = qs("#evidenceTable");
  const pipTable = qs("#pipTable");

  const overviewSearch = qs("#overviewSearch");
  const evidenceSearch = qs("#evidenceSearch");
  const pipSearch = qs("#pipSearch");

  const overviewMeta = qs("#overviewMeta");
  const evidenceMeta = qs("#evidenceMeta");
  const pipMeta = qs("#pipMeta");

  const overviewHint = qs("#overviewHint");
  const evidenceHint = qs("#evidenceHint");
  const pipHint = qs("#pipHint");

  const downloadResult = qs("#downloadResult");
  const downloadEvidence = qs("#downloadEvidence");
  const downloadPip = qs("#downloadPip");

  let currentOverview = [];
  let currentEvidence = [];
  let currentPip = [];
  let activeBand = null;
  let uploadEnabled = true; // Always enabled now
  let currentProjectType = "other";
  
  // 增加多项目数据记忆
  const projectCache = {
    other: {
      overview: [],
      evidence: [],
      pip: [],
      probationSet: new Set(),
      fileName: null,
      rawRows: null
    },
    ziyang: {
      overview: [],
      evidence: [],
      pip: [],
      probationSet: new Set(),
      fileName: null,
      rawRows: null
    }
  };

  let currentProbationSet = projectCache[currentProjectType].probationSet;

  function setStatus(kind, text) {
    statusPill.className = `pill ${kind || "default"}`.trim();
    statusPill.textContent = text;
  }

  function setDownloadsEnabled(enabled) {
    downloadResult.disabled = !enabled;
    downloadEvidence.disabled = !enabled;
    downloadPip.disabled = !enabled;
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
      projectType: currentProjectType,
    });

    const evidenceColumns = [
      "营期",
      "大组",
      "小组",
      ...(currentProjectType === "ziyang" ? ["试用期"] : []),
      "班长",
      "获客",
      "流水",
      "添加人数",
      "添加产值",
      "大盘产值",
      "低于大盘",
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
      searchColumns: ["营期", "班长", "小组", "大组"],
    });

    renderPipTable(pipTable, currentPip, {
      searchInput: pipSearch,
      metaEl: pipMeta,
      hintEl: pipHint,
      projectType: currentProjectType,
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
    currentPip = [];
    activeBand = null;
    currentProbationSet = projectCache[currentProjectType].probationSet;
    overviewSearch.value = "";
    evidenceSearch.value = "";
    pipSearch.value = "";
    overviewTable.innerHTML = "";
    evidenceTable.innerHTML = "";
    pipTable.innerHTML = "";
    overviewMeta.textContent = "";
    evidenceMeta.textContent = "";
    pipMeta.textContent = "";
    summaryChips.innerHTML = "";
    bandFilter.innerHTML = "";
    fileMeta.textContent = "未上传";
    setUploadEnabled(true);
    setStatus("", "等待上传");
    setDownloadsEnabled(false);
    overviewHint.textContent = "上传后将展示：组-班长 + 三期两维度得分 + 总分与带班数；其余过程数据可查看下方明细佐证。";
    evidenceHint.textContent = "";
    pipHint.textContent = "用于判断：是否低于当期大盘产值、是否触发 PIP/淘汰；默认展示相关过程数据。";
  }

  function loadFromCache() {
    const cache = projectCache[currentProjectType];
    currentOverview = cache.overview || [];
    currentEvidence = cache.evidence || [];
    currentPip = cache.pip || [];
    currentProbationSet = cache.probationSet;
    activeBand = null;

    if (currentOverview.length > 0) {
      fileMeta.textContent = cache.fileName || "已恢复缓存数据";
      renderBandFilter();
      renderAll();
      setDownloadsEnabled(true);
      setStatus("ok", "已恢复");
      overviewHint.textContent = "已恢复该项目上次计算的数据。";
      evidenceHint.textContent = "可通过搜索框按营期/班长快速定位并导出。";
      pipHint.textContent = "可通过搜索框按班长/组别定位；导出用于策略运营做 PIP/淘汰辅助决策。";
    } else {
      resetUI();
    }
  }

  downloadResult.addEventListener("click", () => {
    const rows = Array.isArray(overviewTable.__filteredRows) ? overviewTable.__filteredRows : currentOverview;
    if (!rows.length) return;
    const columns = Object.keys(rows[0]);
    downloadText("升班数据.csv", toCSV(rows, columns));
  });

  downloadEvidence.addEventListener("click", () => {
    const rows = Array.isArray(evidenceTable.__filteredRows) ? evidenceTable.__filteredRows : currentEvidence;
    if (!rows.length) return;
    const columns = Object.keys(rows[0]);
    downloadText("过程数据.csv", toCSV(rows, columns));
  });

  downloadPip.addEventListener("click", () => {
    const rows = Array.isArray(pipTable.__filteredRows) ? pipTable.__filteredRows : currentPip;
    if (!rows.length) return;
    const columns = Object.keys(rows[0]);
    downloadText("PIP淘汰.csv", toCSV(rows, columns));
  });

  const probationModal = qs("#probationModal");
  const probationList = qs("#probationList");
  const probationSearch = qs("#probationSearch");
  const probationSelectAll = qs("#probationSelectAll");
  const probationCancel = qs("#probationCancel");
  const probationConfirm = qs("#probationConfirm");

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

      if (currentProjectType === "ziyang") {
        // 展示批量标记试用期 Modal
        showProbationModal(rows, file);
      } else {
        processAndRender(rows);
      }
    } catch (err) {
      handleError(err);
    } finally {
      fileInput.value = "";
    }
  }

  function showProbationModal(rawRows, file) {
    // 提取去重的班长及其最新营期
    const leaderMap = new Map(); // name -> latest term
    for (const r of rawRows) {
      if (currentProjectType === "ziyang") {
        const klass = String(r["班级"] ?? "");
        if (klass.includes("非标-类目一")) continue;
      }
      const name = cleanLeaderName(r["班级"], currentProjectType);
      if (!name) continue;
      const term = toNumber(r["营期"]);
      if (term === null) continue;
      const existing = leaderMap.get(name);
      if (existing === undefined || term > existing) {
        leaderMap.set(name, term);
      }
    }

    const leaderArray = Array.from(leaderMap.entries()).map(([name, term]) => ({ name, term }));
    // 按营期倒序，同营期按姓名排序
    leaderArray.sort((a, b) => {
      if (b.term !== a.term) return b.term - a.term;
      return a.name.localeCompare(b.name, "zh");
    });

    probationSearch.value = "";
    
    let currentFiltered = [];

    function renderProbationList() {
      const tokens = tokenizeQuery(probationSearch.value);
      currentFiltered = leaderArray;
      
      if (tokens.length > 0) {
        currentFiltered = leaderArray.filter(l => {
          const hay = l.name.toLowerCase();
          for (const t of tokens) {
            if (hay.includes(t)) return true;
          }
          return false;
        });
      }

      probationList.innerHTML = "";
      
      if (currentFiltered.length === 0) {
        const emptyState = document.createElement("div");
        emptyState.className = "probation-empty";
        emptyState.textContent = "未找到匹配的班长";
        probationList.appendChild(emptyState);
        probationSelectAll.checked = false;
        probationSelectAll.disabled = true;
        return;
      }
      
      probationSelectAll.disabled = false;
      let allChecked = true;

      for (const l of currentFiltered) {
        const item = document.createElement("label");
        item.className = "probation-item";
        
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = l.name;
        // Keep selection state when searching
        if (currentProbationSet.has(l.name)) {
          checkbox.checked = true;
        } else {
          allChecked = false;
        }
        
        checkbox.addEventListener("change", (e) => {
          if (e.target.checked) {
            currentProbationSet.add(l.name);
          } else {
            currentProbationSet.delete(l.name);
          }
          // Update select all checkbox state
          const allCurrentlyChecked = currentFiltered.every(f => currentProbationSet.has(f.name));
          probationSelectAll.checked = allCurrentlyChecked;
        });
        
        const nameSpan = document.createElement("span");
        nameSpan.className = "leader-name";
        nameSpan.textContent = l.name;
        
        const termSpan = document.createElement("span");
        termSpan.className = "term-tag";
        termSpan.textContent = `最新营期: ${l.term}`;

        item.appendChild(checkbox);
        item.appendChild(nameSpan);
        item.appendChild(termSpan);
        
        probationList.appendChild(item);
      }

      probationSelectAll.checked = allChecked;
    }

    probationSelectAll.onchange = (e) => {
      const isChecked = e.target.checked;
      for (const l of currentFiltered) {
        if (isChecked) {
          currentProbationSet.add(l.name);
        } else {
          currentProbationSet.delete(l.name);
        }
      }
      renderProbationList();
    };

    const onSearch = debounce(() => {
      renderProbationList();
    }, 150);

    probationSearch.addEventListener("input", onSearch);
    
    renderProbationList();

    probationModal.classList.add("active");

    probationCancel.onclick = () => {
      probationSearch.removeEventListener("input", onSearch);
      probationModal.classList.remove("active");
      handleError(new Error("已取消操作"));
    };

    probationConfirm.onclick = () => {
      probationSearch.removeEventListener("input", onSearch);
      probationModal.classList.remove("active");
      processAndRender(rawRows);
    };
  }

  function processAndRender(rows) {
    try {
      const { overview, evidence, pip } = computeFromRawRows(rows, analysisDateInput.value, currentProjectType, currentProbationSet);
      currentOverview = overview;
      currentEvidence = evidence;
      currentPip = pip;
      activeBand = null;

      // 保存到当前项目缓存中
      const cache = projectCache[currentProjectType];
      cache.overview = overview;
      cache.evidence = evidence;
      cache.pip = pip;
      cache.fileName = fileMeta.textContent;
      cache.rawRows = rows;

      renderBandFilter();
      renderAll();
      setDownloadsEnabled(true);
      setStatus("ok", "已完成");
      overviewHint.textContent = "已完成计算。默认按带班数与总分从高到低排序；点击“展开”可查看班型/产值等数据。";
      evidenceHint.textContent = "可通过搜索框按营期/班长快速定位并导出。";
      pipHint.textContent = "可通过搜索框按班长/组别定位；导出用于策略运营做 PIP/淘汰辅助决策。";
      setActiveTab("overview");
    } catch (err) {
      handleError(err);
    }
  }

  function handleError(err) {
    const msg = err instanceof Error ? err.message : String(err);
    setStatus("bad", "计算失败");
    overviewHint.textContent = `计算失败：${msg}`;
    evidenceHint.textContent = "";
    currentOverview = [];
    currentEvidence = [];
    currentPip = [];
    summaryChips.innerHTML = "";
    bandFilter.innerHTML = "";
    overviewTable.innerHTML = "";
    evidenceTable.innerHTML = "";
    pipTable.innerHTML = "";
    overviewMeta.textContent = "";
    evidenceMeta.textContent = "";
    pipMeta.textContent = "";
    setDownloadsEnabled(false);
  }

  clearData.addEventListener("click", () => resetUI());

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
