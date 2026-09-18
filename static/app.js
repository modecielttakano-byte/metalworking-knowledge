const $ = (sel) => document.querySelector(sel);

const state = {
  offset: 0,
  limit: 20,
  total: 0,
  q: "",
};

function selectedSources() {
  return [...document.querySelectorAll('input[name="source"]:checked')].map((el) => el.value);
}

function params(extraOffset) {
  const p = new URLSearchParams();
  p.set("q", $("#q").value.trim());
  p.set("limit", String(state.limit));
  p.set("offset", String(extraOffset ?? state.offset));
  const sources = selectedSources();
  if (sources.length && sources.length < 3) p.set("source", sources.join(","));
  const process = $("#process").value;
  if (process) p.set("process", process);
  if ($("#answered-only").checked) p.set("answered_only", "true");
  if ($("#result-only").checked) p.set("result_only", "true");
  return p;
}

function badgeClass(type) {
  if (type === "internal_log") return "log";
  if (type === "internal_ts_tool") return "ts";
  return "qa";
}

function highlight(text, q) {
  const esc = (s) => s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  let html = esc(text || "");
  const terms = (q || "").trim().split(/\s+/).filter((t) => t.length >= 1);
  for (const term of terms) {
    const re = new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
    html = html.replace(re, (m) => `<mark>${m}</mark>`);
  }
  return html;
}

function fieldChips(hit) {
  const chips = [];
  if (hit.process) chips.push(`工程 ${hit.process}`);
  if (hit.model) chips.push(`機種 ${hit.model}`);
  if (hit.machine) chips.push(`機械 ${hit.machine}`);
  if (hit.factory) chips.push(`工場 ${hit.factory}`);
  if (hit.tool) chips.push(`工具 ${hit.tool}`);
  if (hit.has_result) chips.push("結果あり");
  if (hit.source_type === "external_qa") chips.push(hit.has_answer ? "回答あり" : "未回答");
  return chips.map((c) => `<span class="field">${escText(c)}</span>`).join("");
}

function escText(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function renderHits(hits, append) {
  const root = $("#results");
  if (!append) root.innerHTML = "";
  for (const hit of hits) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "card";
    btn.innerHTML = `
      <div class="meta">
        <span class="badge ${badgeClass(hit.source_type)}">${escText(hit.source_label)}</span>
        <span class="chip ${hit.source_type === "external_qa" ? "trust-forum" : "trust-internal"}">${escText(hit.trust_label)}</span>
        ${fieldChips(hit)}
      </div>
      <h2>${highlight(hit.title, state.q)}</h2>
      <p class="snippet">${highlight(hit.snippet, state.q)}</p>
      ${hit.url ? `<p class="snippet"><a class="ext-link" href="${escText(hit.url)}" target="_blank" rel="noopener">元のQ&amp;Aを開く</a></p>` : ""}
    `;
    btn.addEventListener("click", (ev) => {
      if (ev.target.closest("a")) return;
      openDetail(hit.id);
    });
    root.appendChild(btn);
  }
}

async function runSearch(reset) {
  if (reset) {
    state.offset = 0;
    state.q = $("#q").value.trim();
  }
  $("#status").textContent = "検索中…";
  const res = await fetch("/api/search?" + params(state.offset).toString());
  const data = await res.json();
  if (!res.ok) {
    $("#status").textContent = data.detail || "検索に失敗しました";
    return;
  }
  state.total = data.total;
  renderHits(data.hits, !reset && state.offset > 0);
  const shown = Math.min(state.offset + data.hits.length, data.total);
  if (!data.total) {
    $("#status").textContent = state.q
      ? `「${state.q}」に一致する件はありません。フィルタを緩めるか、別のキーワードを試してください。`
      : "ヒットがありません。取り込み済みか、フィルタを確認してください。";
  } else {
    $("#status").textContent = `「${state.q || "（条件のみ）"}」 ${data.total} 件中 ${shown} 件を表示。社内の対策・結果があるものを先に並べています。`;
  }
  $("#more-wrap").classList.toggle("hidden", shown >= data.total);
}

async function openDetail(id) {
  const res = await fetch("/api/docs/" + id);
  const doc = await res.json();
  $("#detail-title").textContent = doc.title;
  const meta = [];
  meta.push(`<span class="badge ${badgeClass(doc.source_type)}">${escText(doc.source_label)}</span>`);
  meta.push(`<span class="chip ${doc.source_type === "external_qa" ? "trust-forum" : "trust-internal"}">${escText(doc.trust_label)}</span>`);
  if (doc.process) meta.push(`<span class="field">工程 ${escText(doc.process)}</span>`);
  if (doc.model) meta.push(`<span class="field">機種 ${escText(doc.model)}</span>`);
  if (doc.machine) meta.push(`<span class="field">機械 ${escText(doc.machine)}</span>`);
  if (doc.factory) meta.push(`<span class="field">工場 ${escText(doc.factory)}</span>`);
  if (doc.tool) meta.push(`<span class="field">工具 ${escText(doc.tool)}</span>`);
  if (doc.sheet_name) meta.push(`<span class="field">シート ${escText(doc.sheet_name)}</span>`);
  if (doc.url) meta.push(`<a class="ext-link" href="${escText(doc.url)}" target="_blank" rel="noopener">外部URL</a>`);
  $("#detail-meta").innerHTML = meta.join("");
  $("#detail-body").textContent = doc.body || "";
  $("#detail").showModal();
}

async function loadFacets() {
  const health = await fetch("/api/health").then((r) => r.json());
  if (!health.ok) {
    $("#status").textContent = health.hint || "インデックスがまだありません。python ingest.py を実行してください。";
    return;
  }
  const facets = await fetch("/api/facets").then((r) => r.json());
  const sel = $("#process");
  for (const p of facets.processes) {
    const opt = document.createElement("option");
    opt.value = p.process;
    opt.textContent = `${p.process}（${p.n}）`;
    sel.appendChild(opt);
  }
  const stats = health.stats || {};
  $("#status").textContent = `索引 ${stats.total || "?"} 件（社内ログ ${stats.internal_log || 0} / TS・工具 ${stats.internal_ts_tool || 0} / 外部Q&A ${stats.external_qa || 0}）`;
}

$("#search-form").addEventListener("submit", (e) => {
  e.preventDefault();
  runSearch(true);
});
for (const el of document.querySelectorAll(".filters input, .filters select")) {
  el.addEventListener("change", () => runSearch(true));
}
$("#more").addEventListener("click", () => {
  state.offset += state.limit;
  runSearch(false);
});
$("#detail-close").addEventListener("click", () => $("#detail").close());
$("#detail").addEventListener("click", (e) => {
  if (e.target.id === "detail") $("#detail").close();
});

loadFacets();
