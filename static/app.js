"use strict";

/* ================================================================ 共通 */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
const view = $("#view");
const state = { meta: null, equipment: [], recordQuery: { q: "", sort: "done_date", dir: "desc" }, wrap: false, area: "" };
/* 画面右上で選んだエリア（前処理・後処理など）を API に渡す */
const areaQS = (sep = "?") => (state.area ? `${sep}area=${encodeURIComponent(state.area)}` : "");
const inArea = e => !state.area || e.area === state.area;

function h(v) {
  return String(v ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
const fmtDate = d => (d ? d.replaceAll("-", "/") : "");
const fmtNum = n => (n == null ? "" : Number(n).toLocaleString("ja-JP"));
const yen = n => (n ? "¥" + fmtNum(n) : "");

async function api(path, opts = {}) {
  const init = { headers: {}, ...opts };
  if (opts.body && typeof opts.body !== "string") {
    init.body = JSON.stringify(opts.body);
    init.headers["Content-Type"] = "application/json";
  }
  const res = await fetch("/api" + path, init);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `エラー (${res.status})`);
  return data;
}

function toast(msg, ms = 2600) {
  const t = $("#toast");
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => (t.hidden = true), ms);
}

function statusBadge(s) {
  return `<span class="status" style="background:var(--st-${h(s || "完了")})">${h(s || "完了")}</span>`;
}
function kindBadge(k) {
  if (k === "定期") return `<span class="kind teiki">定期</span>`;
  if (k === "不定期") return `<span class="kind futeiki">不定期</span>`;
  return `<span class="muted">—</span>`;
}
const LV_LABEL = { over: "⚠ 期限超過", soon: "● まもなく", ok: "✓ 余裕あり", unknown: "？ 基準なし", none: "—" };
function lvBadge(lv, text) {
  return `<span class="lv ${lv}">${h(text || LV_LABEL[lv])}</span>`;
}
function eqLabel(r) {
  return `${r.equipment_code ? `<span class="code">${h(r.equipment_code)}</span> ` : ""}${h(r.equipment_name)}`;
}
function dateCell(r) {
  const shown = r.done_date ? fmtDate(r.done_date) : r.date_raw || "日付なし";
  const chip = r.date_uncertain ? `<span class="chip-uncertain" title="元の表記: ${h(r.date_raw)}">要確認</span>` : "";
  return `${h(shown)}${chip}`;
}
function currentUser() {
  return $("#current-user").value.trim();
}

/* ツールチップ */
const tip = $("#tip");
document.addEventListener("mousemove", e => {
  const el = e.target.closest("[data-tip]");
  if (!el) { tip.hidden = true; return; }
  tip.innerHTML = el.dataset.tip;
  tip.hidden = false;
  const x = Math.min(e.clientX + 14, window.innerWidth - tip.offsetWidth - 8);
  const y = Math.min(e.clientY + 14, window.innerHeight - tip.offsetHeight - 8);
  tip.style.left = x + "px";
  tip.style.top = y + "px";
});

/* ================================================================ ルーティング */
const routes = {
  dashboard: ["ダッシュボード", renderDashboard],
  records: ["作業計画/履歴", renderRecords],
  calendar: ["点検計画", renderCalendar],
  plans: ["交換アラート", renderPlans],
  equipment: ["設備台帳", renderEquipment],
  parts: ["部品・研磨管理", renderParts],
  analytics: ["分析", renderAnalytics],
  vendors: ["取引業者", renderVendors],
  notify: ["通知設定", renderNotify],
};

async function router() {
  const [name, arg] = (location.hash.replace(/^#\//, "") || "dashboard").split("/");
  const route = routes[name] || routes.dashboard;
  $$("#sidebar nav a").forEach(a => a.classList.toggle("active", a.dataset.nav === name));
  $("#page-title").textContent = route[0];
  document.title = `${route[0]} | 設備保全`;
  document.body.classList.remove("nav-open");
  $("#scrim").hidden = true;
  view.innerHTML = `<p class="muted">読み込み中…</p>`;
  try {
    await route[1](arg);
  } catch (e) {
    view.innerHTML = `<div class="card"><p class="err">${h(e.message)}</p></div>`;
  }
  refreshBadges();
}

async function refreshBadges() {
  try {
    const d = await api("/dashboard" + areaQS());
    setBadge("#badge-plans", d.rules_over.length + d.rules_soon.length);
    setBadge("#badge-parts", d.parts_alert.length);
  } catch { /* バッジは補助表示なので無視 */ }
}
function setBadge(sel, n) {
  const b = $(sel);
  b.hidden = !n;
  b.textContent = n > 99 ? "99+" : n;
}

/* ================================================================ ダッシュボード */
async function renderDashboard() {
  const d = await api("/dashboard" + areaQS());
  const ruleRow = r => `
    <tr data-rule="${r.id}">
      <td>${lvBadge(r.level)}</td>
      <td class="eqcol">${eqLabel(r)}</td>
      <td>${h(r.title)}</td>
      <td class="nowrap">${r.due_date ? fmtDate(r.due_date) : "—"}</td>
      <td class="nowrap num">${r.days_left != null ? (r.days_left < 0 ? `<span class="warn-date">${-r.days_left}日超過</span>` : `あと${r.days_left}日`) : "—"}</td>
    </tr>`;
  const alertRules = [...d.rules_over, ...d.rules_soon];
  view.innerHTML = `
    <div class="grid kpi">
      <a class="kpi-tile" href="#/plans"><div class="kpi-label">${lvBadge("over", "⚠ 期限超過")}</div><div class="kpi-value">${d.rules_over.length}<small>項目</small></div></a>
      <a class="kpi-tile" href="#/plans"><div class="kpi-label">${lvBadge("soon", "● 14日以内")}</div><div class="kpi-value">${d.rules_soon.length}<small>項目</small></div></a>
      <a class="kpi-tile" href="#/records" id="kpi-open"><div class="kpi-label">未完了の作業</div><div class="kpi-value">${d.open_records.length}<small>件</small></div></a>
      <a class="kpi-tile" href="#/parts"><div class="kpi-label">研磨上限が近い部品</div><div class="kpi-value">${d.parts_alert.length}<small>個</small></div></a>
      <a class="kpi-tile" href="#/records" id="kpi-uncertain"><div class="kpi-label">日付が要確認の履歴</div><div class="kpi-value">${d.uncertain_count}<small>件</small></div></a>
    </div>
    <div class="grid cols-2">
      <div class="card">
        <h3>交換・点検アラート <span class="sub">期限超過と14日以内</span></h3>
        ${alertRules.length ? `<div class="table-wrap"><table class="list"><thead><tr><th>状態</th><th>設備</th><th>項目</th><th>期限</th><th class="num">残り</th></tr></thead><tbody>${alertRules.slice(0, 10).map(ruleRow).join("")}</tbody></table></div>
        ${alertRules.length > 10 ? `<p><a href="#/plans">ほか ${alertRules.length - 10} 項目を見る →</a></p>` : ""}` : `<p class="empty">期限が近い項目はありません</p>`}
      </div>
      <div class="card">
        <h3>未完了の作業</h3>
        ${d.open_records.length ? `<div class="table-wrap"><table class="list"><tbody>${d.open_records.map(r => `
          <tr data-record="${r.id}"><td class="nowrap">${dateCell(r)}</td><td class="eqcol">${eqLabel(r)}</td><td class="clip">${h(r.content)}</td><td class="st-cell">${statusBadge(r.status)}</td></tr>`).join("")}
        </tbody></table></div>` : `<p class="empty">未完了の作業はありません</p>`}
      </div>
      <div class="card">
        <h3>最近の作業</h3>
        <ul class="timeline">${d.recent.map(timelineItem).join("")}</ul>
      </div>
      <div class="card">
        <h3>研磨上限が近い部品</h3>
        ${d.parts_alert.length ? `<div class="table-wrap"><table class="list"><tbody>${d.parts_alert.map(p => `
          <tr data-part="${p.id}"><td>${lvBadge(p.level, p.level === "over" ? "⚠ 上限到達" : "● 残りわずか")}</td><td class="code">${h(p.serial)}</td><td>${h(p.equipment_name || "")}</td><td>${grindMeter(p)}</td></tr>`).join("")}
        </tbody></table></div>` : `<p class="empty">該当なし</p>`}
        ${d.rules_unknown.length ? `<p class="muted" style="font-size:13px;margin-top:12px">※ ${d.rules_unknown.length} 項目は前回実施日またはアワーメーターが未入力のため判定できません（<a href="#/plans">交換アラート</a>で確認）</p>` : ""}
      </div>
    </div>`;
  $("#kpi-uncertain").addEventListener("click", () => {
    state.recordQuery = { q: "", sort: "done_date", dir: "desc", uncertain: "1" };
  });
  $("#kpi-open").addEventListener("click", () => {
    state.recordQuery = { q: "", sort: "done_date", dir: "desc", status: state.meta.record_statuses.filter(s => s !== "完了").join(",") };
  });
  bindRowClicks();
}

function timelineItem(r) {
  return `<li class="${r.kind === "定期" ? "teiki" : r.kind === "不定期" ? "futeiki" : ""}" data-record="${r.id}">
    <div class="tl-date">${dateCell(r)} ${kindBadge(r.kind)} ${r.equipment_name ? eqLabel(r) : ""}</div>
    <div class="tl-title">${h(r.content)}</div>
    ${r.reason ? `<div class="tl-body">${h(r.reason)}</div>` : ""}
  </li>`;
}

function bindRowClicks(root = view) {
  $$("[data-record]", root).forEach(el => el.addEventListener("click", () => openRecordForm(+el.dataset.record)));
  $$("[data-part]", root).forEach(el => el.addEventListener("click", () => openPartDetail(+el.dataset.part)));
  $$("[data-rule]", root).forEach(el => el.addEventListener("click", () => openRuleActions(+el.dataset.rule)));
  $$("[data-eq]", root).forEach(el => el.addEventListener("click", () => (location.hash = `#/equipment/${el.dataset.eq}`)));
}

/* ================================================================ 作業計画/履歴 */
async function renderRecords() {
  await loadEquipment();
  const q = state.recordQuery;
  const statuses = state.meta.record_statuses;
  view.innerHTML = `
    <div class="toolbar">
      <div class="search"><span>🔍</span><input id="q" type="search" placeholder="キーワード検索（例：ベアリング、異音、大家）" value="${h(q.q)}"></div>
      <button class="btn primary lg" id="new-record">新規作成</button>
      <a class="btn" id="csv" href="#">CSV出力</a>
    </div>
    <div class="filters">
      <label>設備 <select id="f-eq"><option value="">すべて</option>${equipmentOptions(inArea, q.equipment_id)}</select></label>
      <label>区分 <select id="f-kind"><option value="">すべて</option><option ${q.kind === "定期" ? "selected" : ""}>定期</option><option ${q.kind === "不定期" ? "selected" : ""}>不定期</option></select></label>
      <label>作業状況 <select id="f-status"><option value="">すべて</option><option value="${statuses.filter(s => s !== "完了").join(",")}" ${q.status && q.status !== "完了" ? "selected" : ""}>未完了のみ</option>${statuses.map(s => `<option ${q.status === s ? "selected" : ""}>${s}</option>`).join("")}</select></label>
      <label>期間 <input type="date" id="f-from" value="${h(q.from || "")}">〜<input type="date" id="f-to" value="${h(q.to || "")}"></label>
      <label><input type="checkbox" id="f-uncertain" ${q.uncertain === "1" ? "checked" : ""}>日付要確認のみ</label>
      <button class="btn small" id="f-reset">条件をリセット</button>
    </div>
    <div class="count-line"><label><input type="checkbox" id="wrap" ${state.wrap ? "checked" : ""}> 改行して表示</label><strong id="rec-count"></strong></div>
    <div class="table-wrap ${state.wrap ? "wrap-on" : ""}" id="rec-wrap"><table class="list responsive" id="rec-table">
      <thead><tr>
        <th class="sortable" data-sort="done_date">実施日</th>
        <th class="sortable" data-sort="kind">区分</th>
        <th class="sortable" data-sort="equipment_code">設備</th>
        <th>実施内容</th><th>理由・状況</th><th>今後の対策</th>
        <th class="sortable num" data-sort="downtime_min">停止・修理(分)</th>
        <th class="sortable" data-sort="author">記入者</th>
        <th class="sortable" data-sort="status">作業状況</th>
      </tr></thead><tbody></tbody></table></div>`;

  const reload = async () => {
    const params = new URLSearchParams(Object.entries({ ...q, area: state.area }).filter(([, v]) => v));
    const recs = await api("/records?" + params);
    $("#rec-count").textContent = `${recs.length}件が検索されました`;
    $("#csv").href = "/api/records.csv?" + params;
    $$("#rec-table th[data-sort]").forEach(th => {
      th.classList.toggle("sort-asc", th.dataset.sort === q.sort && q.dir === "asc");
      th.classList.toggle("sort-desc", th.dataset.sort === q.sort && q.dir !== "asc");
    });
    $("#rec-table tbody").innerHTML = recs.length ? recs.map(r => `
      <tr data-record="${r.id}">
        <td class="nowrap" data-label="実施日">${dateCell(r)}</td>
        <td class="nowrap">${kindBadge(r.kind)}</td>
        <td class="eqcol" data-label="設備">${eqLabel(r)}</td>
        <td class="clip" data-label="内容">${h(r.content)}</td>
        <td class="clip" data-label="理由">${h(r.reason)}${r.result ? `\n→ ${h(r.result)}` : ""}</td>
        <td class="clip" data-label="対策">${h(r.countermeasure)}</td>
        <td class="num" data-label="停止・修理(分)">${r.downtime_min != null ? fmtNum(r.downtime_min) : r.hour_meter != null ? `<span class="muted">${r.hour_meter}h</span>` : ""}</td>
        <td class="nowrap" data-label="記入者">${h(r.author)}</td>
        <td class="st-cell">${statusBadge(r.status)}</td>
      </tr>`).join("") : `<tr><td colspan="9" class="empty">該当する履歴はありません</td></tr>`;
    bindRowClicks($("#rec-table"));
  };

  let timer;
  $("#q").addEventListener("input", e => { clearTimeout(timer); timer = setTimeout(() => { q.q = e.target.value; reload(); }, 250); });
  const bindF = (id, key, ev = "change") => $(id).addEventListener(ev, e => { q[key] = e.target.type === "checkbox" ? (e.target.checked ? "1" : "") : e.target.value; reload(); });
  bindF("#f-eq", "equipment_id"); bindF("#f-kind", "kind"); bindF("#f-status", "status");
  bindF("#f-from", "from"); bindF("#f-to", "to"); bindF("#f-uncertain", "uncertain");
  $("#f-reset").addEventListener("click", () => { state.recordQuery = { q: "", sort: "done_date", dir: "desc" }; renderRecords(); });
  $("#wrap").addEventListener("change", e => { state.wrap = e.target.checked; $("#rec-wrap").classList.toggle("wrap-on", state.wrap); });
  $$("#rec-table th[data-sort]").forEach(th => th.addEventListener("click", () => {
    q.dir = q.sort === th.dataset.sort && q.dir !== "asc" ? "asc" : "desc";
    q.sort = th.dataset.sort;
    reload();
  }));
  $("#new-record").addEventListener("click", () => openRecordForm(null, { equipment_id: q.equipment_id }));
  await reload();
}

/* ================================================================ 作業記録フォーム */
async function openRecordForm(id, preset = {}) {
  await loadEquipment();
  const rec = id ? await api(`/records/${id}`) : {
    kind: "不定期", status: "完了", done_date: state.meta.today, author: currentUser(), ...preset,
  };
  const eqOptions = equipmentOptions(e => (e.status !== "廃止" && inArea(e)) || e.id == rec.equipment_id, rec.equipment_id);
  const body = `
    <div class="form-grid">
      <label class="field full"><span>設備<em class="req">*</em></span><select name="equipment_id" required><option value="">選択してください</option>${eqOptions}</select></label>
      <div class="field"><span>区分</span><div class="seg">
        <label><input type="radio" name="kind" value="定期" ${rec.kind === "定期" ? "checked" : ""}>定期</label>
        <label><input type="radio" name="kind" value="不定期" ${rec.kind === "不定期" ? "checked" : ""}>不定期（故障）</label>
      </div></div>
      <label class="field"><span>作業状況</span><select name="status">${state.meta.record_statuses.map(s => `<option ${rec.status === s ? "selected" : ""}>${s}</option>`).join("")}</select></label>
      <label class="field"><span>実施日<em class="hint" style="font-weight:400">（未完了なら予定日）</em></span><input type="date" name="done_date" value="${h(rec.done_date || "")}">
        ${rec.date_uncertain ? `<span class="hint">⚠ 元の表記「${h(rec.date_raw)}」。正しい日付を入れると要確認が外れます</span>` : ""}</label>
      <label class="field"><span>記入者</span><input name="author" list="author-list" value="${h(rec.author || "")}"></label>
      <label class="field full"><span>実施内容<em class="req">*</em></span><textarea name="content" rows="2" required>${h(rec.content || "")}</textarea></label>
      <label class="field full"><span>実施の理由・状況</span><textarea name="reason" rows="3">${h(rec.reason || "")}</textarea></label>
      <label class="field full"><span>実施の結果</span><textarea name="result" rows="2">${h(rec.result || "")}</textarea></label>
      <label class="field full"><span>備考（今後の対策）</span><textarea name="countermeasure" rows="2">${h(rec.countermeasure || "")}</textarea></label>
      <label class="field"><span>停止・修理時間（分）</span><input type="number" min="0" name="downtime_min" value="${h(rec.downtime_min ?? "")}" inputmode="numeric">
        ${rec.downtime_raw && rec.downtime_min == null && rec.hour_meter == null ? `<span class="hint">元の表記: ${h(rec.downtime_raw)}</span>` : ""}</label>
      <label class="field"><span>部品の稼働時間（h）</span><input type="number" min="0" step="0.1" name="hour_meter" value="${h(rec.hour_meter ?? "")}" inputmode="decimal"><span class="hint">フィルタースクリュー交換時のアワーメーター等</span></label>
      <label class="field"><span>関連する定期項目</span><select name="rule_id"><option value="">なし</option></select><span class="hint">完了で保存すると前回実施日が更新されます</span></label>
      <label class="field"><span>部品シリアル</span><input name="part_serial" list="serial-list" value="${h(rec.part_serial || "")}" placeholder="例: M1005.1"><datalist id="serial-list"></datalist></label>
      <label class="field"><span>費用（円）</span><input type="number" min="0" name="cost" value="${h(rec.cost ?? "")}" inputmode="numeric"></label>
    </div>
    <p class="err" id="form-err" hidden></p>`;
  const foot = `
    ${id ? `<button type="button" class="btn danger left" id="del">削除</button>` : ""}
    <button type="button" class="btn" data-close>キャンセル</button>
    <button type="submit" class="btn primary" value="save">保存</button>`;
  openModal(id ? "作業記録の編集" : "作業を記録", body, foot);

  const form = $("#modal-form");
  const fillEqLinked = async () => {
    const eqId = form.equipment_id.value;
    const ruleSel = form.rule_id;
    ruleSel.innerHTML = `<option value="">なし</option>`;
    $("#serial-list").innerHTML = "";
    if (!eqId) return;
    const eq = await api(`/equipment/${eqId}`);
    ruleSel.innerHTML += eq.rules.map(r => `<option value="${r.id}" ${rec.rule_id == r.id ? "selected" : ""}>${h(r.title)}</option>`).join("");
    $("#serial-list").innerHTML = eq.parts.map(p => `<option value="${h(p.serial)}">${h(p.status)}</option>`).join("");
  };
  form.equipment_id.addEventListener("change", fillEqLinked);
  await fillEqLinked();

  form.onsubmit = async e => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(form));
    data.kind = data.kind || null;
    const err = $("#form-err");
    if (!data.equipment_id || !data.content.trim()) {
      err.textContent = "設備と実施内容は必須です";
      err.hidden = false;
      return;
    }
    try {
      if (id) await api(`/records/${id}`, { method: "PUT", body: data });
      else await api("/records", { method: "POST", body: data });
      if (data.author) rememberUser(data.author);
      closeModal();
      toast("保存しました");
      router();
    } catch (ex) {
      err.textContent = ex.message;
      err.hidden = false;
    }
  };
  if (id) $("#del").addEventListener("click", async () => {
    if (!confirm("この作業記録を削除しますか？")) return;
    await api(`/records/${id}`, { method: "DELETE" });
    closeModal();
    toast("削除しました");
    router();
  });
}

/* ================================================================ 交換アラート（定期項目） */
async function renderPlans() {
  const rules = await api("/rules" + areaQS());
  const counts = { over: 0, soon: 0, unknown: 0, ok: 0 };
  rules.filter(r => r.active).forEach(r => counts[r.level]++);
  view.innerHTML = `
    <div class="toolbar">
      <div class="filters" style="margin:0;flex:1">
        ${["over", "soon", "unknown", "ok"].map(l => `<label><input type="checkbox" class="lv-filter" value="${l}" checked> ${lvBadge(l)} ${counts[l]}</label>`).join("")}
      </div>
      <button class="btn primary lg" id="new-rule">項目を追加</button>
    </div>
    <p class="muted" style="font-size:13px;margin:0 0 10px">日数基準は前回実施日から、稼働時間基準は設備のアワーメーター（設備台帳で入力）から判定します。「14日以内」または「残り2割以下」で「まもなく」になります。</p>
    <div class="table-wrap"><table class="list responsive"><thead><tr>
      <th>状態</th><th>設備</th><th>項目</th><th>間隔</th><th>前回実施</th><th>次回期限（日数）</th><th>稼働時間</th><th></th>
    </tr></thead><tbody id="rules-body"></tbody></table></div>`;

  const draw = () => {
    const show = new Set($$(".lv-filter").filter(c => c.checked).map(c => c.value));
    $("#rules-body").innerHTML = rules.filter(r => show.has(r.level)).map(r => `
      <tr data-rule="${r.id}" style="${r.active ? "" : "opacity:.5"}">
        <td>${lvBadge(r.level)}${r.active ? "" : ' <span class="muted">停止中</span>'}</td>
        <td class="eqcol" data-label="設備">${eqLabel(r)}</td>
        <td data-label="項目"><strong>${h(r.title)}</strong>${r.note ? `<div class="muted" style="font-size:12px">${h(r.note)}</div>` : ""}</td>
        <td class="nowrap" data-label="間隔">${[r.interval_days ? `${r.interval_days}日` : "", r.interval_hours ? `${r.interval_hours}h` : ""].filter(Boolean).join(" / ")}</td>
        <td class="nowrap" data-label="前回">${r.last_done_date ? fmtDate(r.last_done_date) : '<span class="muted">未登録</span>'}</td>
        <td class="nowrap" data-label="期限">${r.due_date ? `${fmtDate(r.due_date)}<br>${r.days_left < 0 ? `<span class="warn-date">${-r.days_left}日超過</span>` : `<span class="muted">あと${r.days_left}日</span>`}` : "—"}</td>
        <td data-label="稼働">${r.interval_hours ? hoursMeter(r) : '<span class="muted">—</span>'}</td>
        <td class="nowrap"><button class="btn small do-rule" data-id="${r.id}">実施を記録</button></td>
      </tr>`).join("") || `<tr><td colspan="8" class="empty">該当なし</td></tr>`;
    bindRowClicks($("#rules-body"));
    $$(".do-rule").forEach(b => b.addEventListener("click", e => { e.stopPropagation(); recordForRule(rules.find(r => r.id == b.dataset.id)); }));
  };
  $$(".lv-filter").forEach(c => c.addEventListener("change", draw));
  $("#new-rule").addEventListener("click", () => openRuleForm(null));
  draw();
}

function hoursMeter(r) {
  if (r.current_hours == null) return `<span class="muted">メーター未入力</span>`;
  const pct = Math.max(0, Math.min(100, (r.hours_used / r.interval_hours) * 100));
  const cls = r.hours_left < 0 ? "over" : r.hours_left <= r.interval_hours * 0.2 ? "soon" : "";
  return `<div class="meter" data-tip="${h(`<div class='tip-title'>${h(r.title)}</div>前回から ${r.hours_used}h / 基準 ${r.interval_hours}h<br>メーター更新: ${h(fmtDate(r.hours_updated) || "—")}`)}">
    <div class="meter-bar"><div class="meter-fill ${cls}" style="width:${pct}%"></div></div>
    <span class="meter-text">${r.hours_used}/${r.interval_hours}h</span></div>`;
}

function recordForRule(r) {
  openRecordForm(null, { equipment_id: r.equipment_id, rule_id: r.id, kind: "定期", content: r.title });
}

async function openRuleActions(id) {
  const rules = await api("/rules");
  const r = rules.find(x => x.id === id);
  if (!r) return;
  openModal(r.title, `
    <dl class="info">
      <dt>設備</dt><dd>${eqLabel(r)}</dd>
      <dt>状態</dt><dd>${lvBadge(r.level)}</dd>
      <dt>間隔</dt><dd>${[r.interval_days ? `${r.interval_days}日ごと` : "", r.interval_hours ? `${r.interval_hours}時間ごと` : ""].filter(Boolean).join("、")}</dd>
      <dt>前回実施</dt><dd>${fmtDate(r.last_done_date) || "未登録"}</dd>
      <dt>次回期限</dt><dd>${fmtDate(r.due_date) || "—"}</dd>
      ${r.interval_hours ? `<dt>稼働時間</dt><dd>${hoursMeter(r)}</dd>` : ""}
      <dt>根拠・メモ</dt><dd>${h(r.note || "")}</dd>
    </dl>`, `
    <button type="button" class="btn left" id="edit-rule">設定を変更</button>
    <a class="btn" href="#/equipment/${r.equipment_id}" data-close>設備を見る</a>
    <button type="button" class="btn primary" id="do-rule">実施を記録</button>`);
  $("#edit-rule").addEventListener("click", () => openRuleForm(r));
  $("#do-rule").addEventListener("click", () => recordForRule(r));
}

async function openRuleForm(rule, preset = {}) {
  await loadEquipment();
  const r = rule || { active: 1, ...preset };
  openModal(rule ? "定期項目の設定" : "定期項目を追加", `
    <div class="form-grid">
      <label class="field full"><span>設備<em class="req">*</em></span><select name="equipment_id">${equipmentOptions(e => (e.status !== "廃止" && inArea(e)) || e.id == r.equipment_id, r.equipment_id)}</select></label>
      <label class="field full"><span>項目名<em class="req">*</em></span><input name="title" value="${h(r.title || "")}" placeholder="例: フィルタースクリュー交換"></label>
      <label class="field"><span>間隔（日数）</span><input type="number" min="1" name="interval_days" value="${h(r.interval_days ?? "")}"></label>
      <label class="field"><span>間隔（稼働時間 h）</span><input type="number" min="1" step="0.1" name="interval_hours" value="${h(r.interval_hours ?? "")}"></label>
      <label class="field"><span>前回実施日</span><input type="date" name="last_done_date" value="${h(r.last_done_date || "")}"></label>
      <label class="field"><span>前回実施時のメーター値（h）</span><input type="number" min="0" step="0.1" name="last_hours" value="${h(r.last_hours ?? "")}"></label>
      <label class="field full"><span>根拠・メモ</span><textarea name="note">${h(r.note || "")}</textarea></label>
      <label class="field"><span>有効</span><select name="active"><option value="1" ${r.active ? "selected" : ""}>有効</option><option value="0" ${r.active ? "" : "selected"}>停止</option></select></label>
    </div><p class="err" id="form-err" hidden></p>`, `
    ${rule ? `<button type="button" class="btn danger left" id="del">削除</button>` : ""}
    <button type="button" class="btn" data-close>キャンセル</button>
    <button type="submit" class="btn primary">保存</button>`);
  const form = $("#modal-form");
  form.onsubmit = async e => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(form));
    try {
      if (rule) await api(`/rules/${rule.id}`, { method: "PUT", body: data });
      else await api("/rules", { method: "POST", body: data });
      closeModal(); toast("保存しました"); router();
    } catch (ex) { $("#form-err").textContent = ex.message; $("#form-err").hidden = false; }
  };
  if (rule) $("#del").addEventListener("click", async () => {
    if (!confirm("この定期項目を削除しますか？（作業履歴は残ります）")) return;
    await api(`/rules/${rule.id}`, { method: "DELETE" });
    closeModal(); toast("削除しました"); router();
  });
}

/* ================================================================ 点検計画（カレンダー） */
const pad2 = n => String(n).padStart(2, "0");
const isoDate = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const parseISO = s => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];
/* 色だけに頼らないよう、種類はラベルでも示す */
const EV_TYPES = {
  planned: { label: "予定", desc: "日付を決めた作業（未完了）" },
  due: { label: "期限", desc: "定期項目の次回期限" },
  projected: { label: "見込み", desc: "期限どおりに実施した場合のその次" },
  done: { label: "実績", desc: "完了した定期・その他の作業" },
  failure: { label: "故障", desc: "完了した不定期（故障）の作業" },
};
const evKey = ev => (ev.type === "done" && ev.kind === "不定期" ? "failure" : ev.type);
const calState = {
  mode: "month", ym: null, year: null, selected: null,
  show: { planned: true, due: true, projected: true, done: true, failure: false },
};

async function renderCalendar(arg) {
  const today = state.meta.today;
  calState.mode = arg === "year" ? "year" : "month";
  calState.ym ||= today.slice(0, 7);
  calState.year ||= +today.slice(0, 4);
  view.innerHTML = `
    <div class="toolbar cal-toolbar">
      <div class="seg small" id="cal-mode">
        <label><input type="radio" name="cal-mode" value="month" ${calState.mode === "month" ? "checked" : ""}>月</label>
        <label><input type="radio" name="cal-mode" value="year" ${calState.mode === "year" ? "checked" : ""}>年間計画表</label>
      </div>
      <div class="cal-nav">
        <button class="btn small" id="cal-prev" aria-label="前へ">‹</button>
        <strong id="cal-label"></strong>
        <button class="btn small" id="cal-next" aria-label="次へ">›</button>
        <button class="btn small" id="cal-today">今日</button>
      </div>
      <span class="spacer"></span>
      <button class="btn" id="cal-print">印刷</button>
      <button class="btn primary" id="cal-add">予定を追加</button>
    </div>
    <div id="cal-body"></div>`;
  $$("#cal-mode input").forEach(r => r.addEventListener("change", () => { setCalMode(r.value); drawCalendar(); }));
  const move = step => {
    if (calState.mode === "year") calState.year += step;
    else {
      const d = parseISO(calState.ym + "-01");
      d.setMonth(d.getMonth() + step);
      calState.ym = isoDate(d).slice(0, 7);
    }
    drawCalendar();
  };
  $("#cal-prev").addEventListener("click", () => move(-1));
  $("#cal-next").addEventListener("click", () => move(1));
  $("#cal-today").addEventListener("click", () => {
    calState.ym = today.slice(0, 7); calState.year = +today.slice(0, 4); calState.selected = today; drawCalendar();
  });
  $("#cal-print").addEventListener("click", () => window.print());
  $("#cal-add").addEventListener("click", () => openPlanForm(calState.selected || today));
  await drawCalendar();
}

/* 表示切替は URL にも残す（#/calendar/year で年間計画表を直接開ける） */
function setCalMode(mode) {
  calState.mode = mode;
  history.replaceState(null, "", mode === "year" ? "#/calendar/year" : "#/calendar");
}

async function drawCalendar() {
  return calState.mode === "year" ? drawYearPlan() : drawMonth();
}

/* 未完了（未着手）の作業として予定を入れる。完了にすると定期項目の前回実施日が更新される */
function openPlanForm(date, preset = {}) {
  openRecordForm(null, { status: "未着手", kind: "定期", done_date: date, ...preset });
}
async function scheduleRule(ruleInfo, date) {
  await api("/records", {
    method: "POST",
    body: { equipment_id: ruleInfo.equipment_id, rule_id: ruleInfo.rule_id, kind: "定期", status: "未着手",
            done_date: date, content: ruleInfo.title, author: currentUser() },
  });
  toast(`${fmtDate(date)} に予定を入れました`);
}
async function completePlanned(ev) {
  // 予定日より前に終わった場合は今日の日付で完了にする
  const date = ev.date > state.meta.today ? state.meta.today : ev.date;
  await api(`/records/${ev.id}`, { method: "PUT", body: { status: "完了", done_date: date } });
  toast("完了にしました");
}

function evChip(ev, i) {
  const key = evKey(ev);
  const late = ev.type === "planned" && ev.date < state.meta.today;
  const cls = ["ev", `ev-${key}`, ev.type === "due" ? ev.level : "", late ? "late" : ""].join(" ");
  const eq = ev.equipment_nickname || ev.equipment_name || "";
  const drag = ev.type === "planned" || ev.type === "due";
  return `<div class="${cls}" data-ev="${i}" ${drag ? 'draggable="true"' : ""} title="${h(`${EV_TYPES[key].label}${late ? "（遅れ）" : ""}：${eq} ${ev.title || ""}`)}">
    <span class="ev-tag">${late ? "遅れ" : EV_TYPES[key].label}</span>${h(eq)} <span class="ev-title">${h(ev.title || "")}</span></div>`;
}

async function drawMonth() {
  const today = state.meta.today;
  const first = parseISO(calState.ym + "-01");
  const month = first.getMonth();
  const gridStart = new Date(first);
  gridStart.setDate(1 - first.getDay());
  const last = new Date(first.getFullYear(), month + 1, 0);
  const gridEnd = new Date(last);
  gridEnd.setDate(last.getDate() + (6 - last.getDay()));
  $("#cal-label").textContent = `${first.getFullYear()}年${month + 1}月`;

  const d = await api(`/schedule?from=${isoDate(gridStart)}&to=${isoDate(gridEnd)}${areaQS("&")}`);
  if (!calState.selected || calState.selected.slice(0, 7) !== calState.ym) {
    calState.selected = today.slice(0, 7) === calState.ym ? today : calState.ym + "-01";
  }
  const byDate = {};
  d.events.forEach((ev, i) => { ev.i = i; (byDate[ev.date] ||= []).push(ev); });
  const visible = ev => calState.show[evKey(ev)];
  const counts = {};
  d.events.filter(ev => ev.date.slice(0, 7) === calState.ym).forEach(ev => (counts[evKey(ev)] = (counts[evKey(ev)] || 0) + 1));

  const cells = [];
  for (let c = new Date(gridStart); c <= gridEnd; c.setDate(c.getDate() + 1)) {
    const iso = isoDate(c);
    const evs = (byDate[iso] || []).filter(visible);
    const MAX = 3;
    cells.push(`<div class="cal-day ${c.getMonth() !== month ? "out" : ""} ${iso === today ? "today" : ""} ${iso === calState.selected ? "sel" : ""} wd${c.getDay()}" data-date="${iso}" tabindex="0">
      <div class="cal-dnum">${c.getDate()}${evs.length ? `<span class="cal-dots" aria-label="${evs.length}件">${evs.length}</span>` : ""}</div>
      <div class="cal-evs">${evs.slice(0, MAX).map(ev => evChip(ev, ev.i)).join("")}
        ${evs.length > MAX ? `<div class="ev-more">他 ${evs.length - MAX} 件</div>` : ""}</div>
    </div>`);
  }

  $("#cal-body").innerHTML = `
    ${overduePanel(d.overdue)}
    <div class="filters cal-legend">
      ${Object.entries(EV_TYPES).map(([k, t]) => `<label title="${h(t.desc)}"><input type="checkbox" data-show="${k}" ${calState.show[k] ? "checked" : ""}>
        <span class="ev ev-${k} ${k === "due" ? "soon" : ""} legend-chip"><span class="ev-tag">${t.label}</span></span>${counts[k] || 0}</label>`).join("")}
      <span class="muted cal-hint">予定・期限はドラッグで別の日へ動かせます</span>
    </div>
    <div class="cal-layout">
      <div class="cal-grid" role="grid">
        ${WEEKDAYS.map((w, i) => `<div class="cal-wd wd${i}">${w}</div>`).join("")}
        ${cells.join("")}
      </div>
      <aside class="card cal-side" id="cal-side"></aside>
    </div>
    ${d.hours_only.length ? `<p class="muted" style="font-size:13px">※ 稼働時間だけで管理している ${d.hours_only.length} 項目は日付が決まらないためカレンダーには出ません（<a href="#/plans">交換アラート</a>で確認）。</p>` : ""}`;

  const drawSide = () => {
    const iso = calState.selected;
    const evs = byDate[iso] || [];
    const dd = parseISO(iso);
    $("#cal-side").innerHTML = `
      <h3>${dd.getMonth() + 1}月${dd.getDate()}日（${WEEKDAYS[dd.getDay()]}）${iso === today ? '<span class="sub">今日</span>' : ""}</h3>
      ${evs.length ? `<ul class="day-list">${evs.map(ev => dayItem(ev)).join("")}</ul>` : `<p class="empty">予定はありません</p>`}
      <button class="btn small" id="day-add">＋ この日に予定を追加</button>`;
    $("#day-add").addEventListener("click", () => openPlanForm(iso));
    $$("#cal-side [data-act]").forEach(b => b.addEventListener("click", async e => {
      e.stopPropagation();
      const ev = d.events[+b.dataset.i];
      try {
        if (b.dataset.act === "open") openRecordForm(ev.id);
        else if (b.dataset.act === "complete") { await completePlanned(ev); router(); }
        else if (b.dataset.act === "schedule") { await scheduleRule(ev, iso); router(); }
        else if (b.dataset.act === "record") recordForRule({ id: ev.rule_id, equipment_id: ev.equipment_id, title: ev.title });
        else if (b.dataset.act === "rule") openRuleActions(ev.rule_id);
      } catch (ex) { toast(ex.message); }
    }));
  };
  const select = iso => {
    calState.selected = iso;
    $$(".cal-day.sel").forEach(x => x.classList.remove("sel"));
    $(`.cal-day[data-date="${iso}"]`)?.classList.add("sel");
    drawSide();
  };

  $$(".cal-day").forEach(cell => {
    cell.addEventListener("click", () => select(cell.dataset.date));
    cell.addEventListener("keydown", e => { if (e.key === "Enter") select(cell.dataset.date); });
    cell.addEventListener("dblclick", () => openPlanForm(cell.dataset.date));
    cell.addEventListener("dragover", e => { e.preventDefault(); cell.classList.add("drop"); });
    cell.addEventListener("dragleave", () => cell.classList.remove("drop"));
    cell.addEventListener("drop", async e => {
      e.preventDefault();
      cell.classList.remove("drop");
      const date = cell.dataset.date;
      try {
        const src = JSON.parse(e.dataTransfer.getData("text/plain") || "{}");
        if (src.record) {
          await api(`/records/${src.record}`, { method: "PUT", body: { done_date: date } });
          toast(`${fmtDate(date)} に移しました`);
        } else if (src.rule) {
          await scheduleRule(src.rule, date);
        } else return;
        calState.selected = date;
        router();
      } catch (ex) { toast(ex.message); }
    });
  });
  $$(".cal-grid .ev[draggable]").forEach(chip => chip.addEventListener("dragstart", e => {
    const ev = d.events[+chip.dataset.ev];
    e.dataTransfer.setData("text/plain", JSON.stringify(ev.type === "planned" ? { record: ev.id } : { rule: pickRule(ev) }));
  }));
  $$(".cal-grid .ev").forEach(chip => chip.addEventListener("click", e => {
    // チップを押したらその日を選ぶ（詳細は右の一覧から操作する）
    e.stopPropagation();
    select(chip.closest(".cal-day").dataset.date);
  }));
  $$("[data-show]").forEach(c => c.addEventListener("change", () => { calState.show[c.dataset.show] = c.checked; drawMonth(); }));
  bindOverdue(d.overdue);
  drawSide();
}
const pickRule = r => ({ rule_id: r.rule_id, equipment_id: r.equipment_id, title: r.title });

function dayItem(ev) {
  const key = evKey(ev);
  const late = ev.type === "planned" && ev.date < state.meta.today;
  const btn = (act, label, primary) => `<button class="btn small ${primary ? "primary" : ""}" data-act="${act}" data-i="${ev.i}">${label}</button>`;
  const actions = {
    planned: btn("complete", "完了にする", true) + btn("open", "開く"),
    due: btn("schedule", "この日に予定を入れる", true) + btn("record", "実施を記録") + btn("rule", "設定"),
    projected: btn("rule", "設定"),
    done: btn("open", "開く"),
    failure: btn("open", "開く"),
  }[key];
  return `<li class="day-item">
    <div class="day-head"><span class="ev ev-${key} ${ev.type === "due" ? ev.level : ""} ${late ? "late" : ""} legend-chip"><span class="ev-tag">${late ? "遅れ" : EV_TYPES[key].label}</span></span>
      ${ev.type === "planned" && ev.status !== "未着手" ? statusBadge(ev.status) : ""}</div>
    <div class="day-eq">${eqLabel({ equipment_code: ev.equipment_code, equipment_name: ev.equipment_nickname || ev.equipment_name })}</div>
    <div class="day-title">${h(ev.title || "")}</div>
    ${ev.type === "projected" ? `<div class="muted" style="font-size:12px">${ev.interval_days}日ごと。前回の実施日によって前後します</div>` : ""}
    ${ev.author ? `<div class="muted" style="font-size:12px">記入者 ${h(ev.author)}</div>` : ""}
    <div class="day-actions">${actions}</div>
  </li>`;
}

function overduePanel(list) {
  if (!list.length) return "";
  return `<details class="card overdue-panel" ${calState.overdueOpen ? "open" : ""}>
    <summary>${lvBadge("over", `⚠ 期限超過・予定未定 ${list.length}項目`)}<span class="muted">日付を決めて予定に入れてください（カレンダーへドラッグしても入れられます）</span></summary>
    <div class="table-wrap"><table class="list responsive plain"><thead><tr><th>設備</th><th>項目</th><th>前回実施</th><th>期限</th><th>予定日</th></tr></thead><tbody>
      ${list.map((r, i) => `<tr>
        <td class="eqcol" data-label="設備"><span class="ev-grip" draggable="true" data-od="${i}" title="カレンダーの日付へドラッグ">⠿</span> ${eqLabel({ equipment_code: r.equipment_code, equipment_name: r.equipment_nickname || r.equipment_name })}</td>
        <td data-label="項目"><strong>${h(r.title)}</strong> <span class="muted">${r.interval_days}日ごと</span></td>
        <td class="nowrap" data-label="前回">${fmtDate(r.last_done_date) || '<span class="muted">未登録</span>'}</td>
        <td class="nowrap" data-label="期限">${r.due_date ? `${fmtDate(r.due_date)} <span class="warn-date">${-r.days_left}日超過</span>` : '<span class="muted">前回不明</span>'}</td>
        <td class="nowrap"><input type="date" class="od-date" data-od="${i}" value="${state.meta.today}"> <button class="btn small primary od-go" data-od="${i}">予定に入れる</button></td>
      </tr>`).join("")}
    </tbody></table></div>
  </details>`;
}
function bindOverdue(list) {
  const panel = $(".overdue-panel");
  if (!panel) return;
  panel.addEventListener("toggle", () => (calState.overdueOpen = panel.open));
  $$(".od-go").forEach(b => b.addEventListener("click", async () => {
    const date = $(`.od-date[data-od="${b.dataset.od}"]`).value;
    if (!date) return toast("予定日を入れてください");
    try { await scheduleRule(list[+b.dataset.od], date); router(); } catch (ex) { toast(ex.message); }
  }));
  $$(".ev-grip").forEach(g => g.addEventListener("dragstart", e => {
    e.dataTransfer.setData("text/plain", JSON.stringify({ rule: pickRule(list[+g.dataset.od]) }));
  }));
}

/* 年間計画表：定期項目 × 12か月。工場でよく使う「年間保全計画表」の形 */
async function drawYearPlan() {
  const y = calState.year;
  $("#cal-label").textContent = `${y}年`;
  const d = await api(`/schedule?from=${y}-01-01&to=${y}-12-31${areaQS("&")}`);
  const marks = {}; // rule_id → month(1-12) → Set(種類)
  const put = (rid, date, kind) => {
    if (!rid || !date || +date.slice(0, 4) !== y) return;
    ((marks[rid] ||= {})[+date.slice(5, 7)] ||= new Set()).add(kind);
  };
  d.events.forEach(ev => put(ev.rule_id, ev.date, ev.type === "planned" && ev.date < d.today ? "late" : ev.type));
  d.rules.forEach(r => put(r.id, r.last_done_date, "done"));
  const overdue = new Set(d.overdue.map(r => r.rule_id));
  const curMonth = +d.today.slice(0, 4) === y ? +d.today.slice(5, 7) : 0;
  const SYM = { done: ["●", "実施済み"], planned: ["◎", "予定"], late: ["◎", "予定（遅れ）"], due: ["○", "期限"], projected: ["◌", "見込み"] };
  const cell = (r, m) => {
    const set = marks[r.id]?.[m];
    let inner = set ? ["done", "late", "planned", "due", "projected"].filter(k => set.has(k)).map(k => `<span class="ym ym-${k}" title="${SYM[k][1]}">${SYM[k][0]}</span>`).join("") : "";
    if (!inner && overdue.has(r.id) && m === curMonth) inner = `<span class="ym ym-over" title="期限超過">⚠</span>`;
    return `<td class="ym-cell ${m === curMonth ? "cur" : ""}">${inner}</td>`;
  };
  const byCode = [...d.rules].sort((x, z) => (x.equipment_code || "ZZZ").localeCompare(z.equipment_code || "ZZZ") || x.equipment_id - z.equipment_id);
  const groups = groupBy(byCode, r => String(r.equipment_id));
  const eqName = r => `${r.equipment_code || ""} ${r.equipment_nickname || r.equipment_name}`.trim();
  $("#cal-body").innerHTML = `
    <div class="filters ym-legend">${Object.entries(SYM).filter(([k]) => k !== "late").map(([k, [s, l]]) => `<span><span class="ym ym-${k}">${s}</span> ${l}</span>`).join("")}
      <span><span class="ym ym-over">⚠</span> 期限超過（今月に表示）</span>
      <span class="muted">月の見出しを押すとその月のカレンダーを開きます</span></div>
    ${d.rules.length ? `<div class="table-wrap"><table class="list plain year-plan"><thead><tr>
      <th>設備</th><th>項目</th><th>間隔</th>${Array.from({ length: 12 }, (_, i) => `<th class="ym-head ${i + 1 === curMonth ? "cur" : ""}" data-month="${i + 1}">${i + 1}月</th>`).join("")}
    </tr></thead><tbody>
      ${groups.map(([, rs]) => rs.map((r, i) => `<tr>
        ${i === 0 ? `<td rowspan="${rs.length}" class="eqcol"><a href="#/equipment/${r.equipment_id}">${h(eqName(r))}</a></td>` : ""}
        <td><a href="#" data-rule-open="${r.id}">${h(r.title)}</a>${overdue.has(r.id) ? ` ${lvBadge("over", "超過")}` : ""}</td>
        <td class="nowrap muted">${[r.interval_days ? `${r.interval_days}日` : "", r.interval_hours ? `${r.interval_hours}h` : ""].filter(Boolean).join(" / ")}</td>
        ${Array.from({ length: 12 }, (_, m) => cell(r, m + 1)).join("")}
      </tr>`).join("")).join("")}
    </tbody></table></div>` : `<p class="empty">定期項目がありません。<a href="#/plans">交換アラート</a>から追加してください</p>`}`;
  $$(".ym-head").forEach(th => th.addEventListener("click", () => {
    setCalMode("month");
    calState.ym = `${y}-${pad2(th.dataset.month)}`;
    $(`#cal-mode input[value="month"]`).checked = true;
    drawMonth();
  }));
  $$("[data-rule-open]").forEach(a => a.addEventListener("click", e => { e.preventDefault(); openRuleActions(+a.dataset.ruleOpen); }));
}

/* ================================================================ 設備台帳 */
function equipmentOptions(filter, selected) {
  const groups = {};
  state.equipment.filter(filter).forEach(e => (groups[e.area || "未分類"] ||= []).push(e));
  const opt = e => `<option value="${e.id}" ${selected == e.id ? "selected" : ""}>${h((e.code ? e.code + " " : "") + eqTitle(e))}${e.status === "廃止" ? "（廃止）" : ""}</option>`;
  const names = Object.keys(groups);
  if (names.length === 1) return groups[names[0]].map(opt).join("");
  return names.map(a => `<optgroup label="${h(a)}">${groups[a].map(opt).join("")}</optgroup>`).join("");
}

async function loadEquipment(force) {
  if (!state.meta) state.meta = await api("/meta");
  if (force || !state.equipment.length) state.equipment = await api("/equipment");
  return state.equipment;
}

const eqTitle = e => e.nickname || e.name;
const eqPrefs = (() => {
  const def = { view: "card", group: "area", retired: false };
  try { return { ...def, ...JSON.parse(localStorage.getItem("maint-eq-prefs") || "{}") }; } catch { return def; }
})();
function saveEqPrefs() {
  try { localStorage.setItem("maint-eq-prefs", JSON.stringify(eqPrefs)); } catch { /* 保存できなくても動作には影響しない */ }
}
const eqCollapsed = new Set();

/* 配列を key ごとにまとめる（並び順は order の順、なければ出てきた順） */
function groupBy(list, key, order = []) {
  const map = new Map(order.map(k => [k, []]));
  list.forEach(x => { const k = key(x) || "その他"; if (!map.has(k)) map.set(k, []); map.get(k).push(x); });
  return [...map].filter(([, v]) => v.length);
}

function alertChips(list) {
  const over = list.reduce((s, e) => s + e.alerts.over, 0);
  const soon = list.reduce((s, e) => s + e.alerts.soon, 0);
  return `${over ? lvBadge("over", `⚠ 期限超過 ${over}`) : ""}${soon ? lvBadge("soon", `● まもなく ${soon}`) : ""}`;
}

function eqCard(e) {
  return `<div class="eq-card ${e.status === "廃止" ? "retired" : ""}" data-eq="${e.id}">
    <div class="eq-name">${h(eqTitle(e))}${e.status === "廃止" ? ' <span class="lv unknown">廃止</span>' : ""}</div>
    ${e.nickname ? `<div class="eq-formal" title="${h(e.name)}">${h(e.name)}</div>` : ""}
    <div class="eq-meta">${e.code ? `<span class="code">${h(e.code)}</span>` : ""}<span title="全${e.record_count}件のうち故障${e.failure_count}件">故障 ${e.failure_count}/${e.record_count}件</span><span>最終 ${h(fmtDate(e.last_date) || "—")}</span></div>
    ${e.alerts.over || e.alerts.soon || e.current_hours != null ? `<div class="eq-badges">${alertChips([e])}${e.current_hours != null ? `<span class="lv unknown">メーター ${e.current_hours}h</span>` : ""}</div>` : ""}
  </div>`;
}

function eqTable(list) {
  return `<div class="table-wrap"><table class="list responsive eq-table"><thead><tr>
    <th>設備No</th><th>通称</th><th>正式名</th><th>メーカー</th><th class="num">故障</th><th class="num">履歴</th><th>最終作業</th><th>アラート</th>
  </tr></thead><tbody>${list.map(e => `<tr data-eq="${e.id}" class="${e.status === "廃止" ? "retired" : ""}">
    <td class="code nowrap" data-label="設備No">${h(e.code || "—")}</td>
    <td data-label="通称"><strong>${h(eqTitle(e))}</strong>${e.status === "廃止" ? ' <span class="lv unknown">廃止</span>' : ""}</td>
    <td class="clip muted" data-label="正式名">${e.nickname ? h(e.name) : ""}</td>
    <td class="clip" data-label="メーカー">${h(e.maker)}</td>
    <td class="num" data-label="故障">${e.failure_count}</td>
    <td class="num" data-label="履歴">${e.record_count}</td>
    <td class="nowrap" data-label="最終作業">${h(fmtDate(e.last_date) || "—")}</td>
    <td data-label="アラート">${alertChips([e])}</td>
  </tr>`).join("")}</tbody></table></div>`;
}

async function renderEquipment(id) {
  if (id) return renderEquipmentDetail(+id);
  const all = await loadEquipment(true);
  let category = "";
  view.innerHTML = `
    <div class="toolbar">
      <div class="search"><span>🔍</span><input id="eq-q" type="search" placeholder="通称・正式名・設備No・メーカーで絞り込み"></div>
      <button class="btn primary lg" id="new-eq">設備を追加</button>
    </div>
    <div class="filters eq-controls">
      <span class="ctl-label">まとめ方</span>
      <div class="seg small" id="eq-group">
        <label><input type="radio" name="eq-group" value="area" ${eqPrefs.group === "area" ? "checked" : ""}>エリア → 種類</label>
        <label><input type="radio" name="eq-group" value="category" ${eqPrefs.group === "category" ? "checked" : ""}>種類だけ</label>
      </div>
      <span class="ctl-label">表示</span>
      <div class="seg small" id="eq-view">
        <label><input type="radio" name="eq-view" value="card" ${eqPrefs.view === "card" ? "checked" : ""}>カード</label>
        <label><input type="radio" name="eq-view" value="list" ${eqPrefs.view === "list" ? "checked" : ""}>一覧</label>
      </div>
      <label><input type="checkbox" id="eq-retired" ${eqPrefs.retired ? "checked" : ""}> 廃止設備も表示</label>
      <span class="spacer"></span>
      <button class="btn small" id="eq-open">すべて開く</button>
      <button class="btn small" id="eq-close">すべて閉じる</button>
    </div>
    <div class="cat-chips" id="cat-chips"></div>
    <div id="eq-body"></div>`;

  const groupHtml = (key, label, list, level) => {
    const open = !eqCollapsed.has(key);
    return `<details class="eq-group lv${level}" data-key="${h(key)}" ${open ? "open" : ""}>
      <summary><span class="g-title">${h(label)}</span><span class="g-count">${list.length}台</span>${alertChips(list)}</summary>
      ${eqPrefs.view === "list" ? eqTable(list) : `<div class="eq-grid">${list.map(eqCard).join("")}</div>`}
    </details>`;
  };

  const draw = () => {
    const q = norm($("#eq-q").value).toLowerCase();
    const base = all.filter(e => inArea(e) && (eqPrefs.retired || e.status !== "廃止") &&
      (!q || norm([e.code, e.name, e.nickname, e.maker, e.vendor_name, e.category].join(" ")).toLowerCase().includes(q)));
    // 種類チップ（件数つき）
    const cats = groupBy(base, e => e.category, state.meta.categories);
    $("#cat-chips").innerHTML = `<button class="chip ${category ? "" : "on"}" data-cat="">すべて <b>${base.length}</b></button>` +
      cats.map(([c, l]) => `<button class="chip ${category === c ? "on" : ""}" data-cat="${h(c)}">${h(c)} <b>${l.length}</b></button>`).join("");
    $$("#cat-chips .chip").forEach(b => b.addEventListener("click", () => { category = b.dataset.cat; draw(); }));

    const list = base.filter(e => !category || e.category === category);
    if (!list.length) { $("#eq-body").innerHTML = `<p class="empty">該当する設備はありません</p>`; return; }
    const byCat = items => groupBy(items, e => e.category, state.meta.categories);
    let html;
    if (eqPrefs.group === "area" && !state.area) {
      html = groupBy(list, e => e.area, state.meta.areas).map(([area, items]) => `
        <section class="eq-area">
          <h2 class="area-title">${h(area)} <span class="g-count">${items.length}台</span>${alertChips(items)}</h2>
          ${byCat(items).map(([c, l]) => groupHtml(`${area}|${c}`, c, l, 2)).join("")}
        </section>`).join("");
    } else {
      html = byCat(list).map(([c, l]) => groupHtml(`*|${c}`, c, l, 1)).join("");
    }
    // 検索中・種類で絞り込み中は全部開いて見せる
    $("#eq-body").innerHTML = html;
    if (q || category) $$("#eq-body details").forEach(d => (d.open = true));
    $$("#eq-body details").forEach(d => d.addEventListener("toggle", () => {
      if (q || category) return;
      d.open ? eqCollapsed.delete(d.dataset.key) : eqCollapsed.add(d.dataset.key);
    }));
    bindRowClicks($("#eq-body"));
  };

  $("#eq-q").addEventListener("input", draw);
  $$("#eq-group input").forEach(r => r.addEventListener("change", () => { eqPrefs.group = r.value; saveEqPrefs(); draw(); }));
  $$("#eq-view input").forEach(r => r.addEventListener("change", () => { eqPrefs.view = r.value; saveEqPrefs(); draw(); }));
  $("#eq-retired").addEventListener("change", e => { eqPrefs.retired = e.target.checked; saveEqPrefs(); draw(); });
  $("#eq-open").addEventListener("click", () => { $$("#eq-body details").forEach(d => (d.open = true)); });
  $("#eq-close").addEventListener("click", () => { $$("#eq-body details").forEach(d => (d.open = false)); });
  $("#new-eq").addEventListener("click", () => openEquipmentForm(null));
  draw();
}

/* 半角カナ・全角英数をそろえる（サーバー側の norm と同じ考え方） */
function norm(s) {
  return String(s ?? "").normalize("NFKC").trim();
}

async function renderEquipmentDetail(id) {
  const [eq, recs] = await Promise.all([api(`/equipment/${id}`), api(`/records?equipment_id=${id}`)]);
  $("#page-title").textContent = eqTitle(eq);
  view.innerHTML = `
    <a class="back" href="#/equipment">← 設備台帳へ戻る</a> <span class="muted crumb">${h(eq.area)} ＞ ${h(eq.category)}</span>
    <div class="grid cols-2">
      <div class="card">
        <h3>${h(eq.code || "設備No未採番")} ${h(eqTitle(eq))} ${eq.status === "廃止" ? '<span class="lv unknown">廃止</span>' : ""}</h3>
        <dl class="info">
          <dt>正式名</dt><dd>${h(eq.name)}</dd>
          <dt>エリア・種類</dt><dd>${h(eq.area)} ／ ${h(eq.category)}</dd>
          <dt>型式</dt><dd>${h(eq.model)}</dd>
          <dt>メーカー</dt><dd>${h(eq.maker)}</dd>
          <dt>納入業者</dt><dd>${h(eq.supplier)}</dd>
          <dt>資産No</dt><dd class="code">${h(eq.asset_no)}</dd>
          <dt>製造番号</dt><dd>${h(eq.serial_no)}</dd>
          <dt>製造年</dt><dd>${h(eq.manufactured)}</dd>
          <dt>納入年月日</dt><dd>${h(eq.delivered)}</dd>
          <dt>メモ</dt><dd>${h(eq.note)}</dd>
        </dl>
        <div style="margin-top:12px"><button class="btn small" id="edit-eq">設備情報を編集</button></div>
      </div>
      <div class="card">
        <h3>取引業者</h3>
        ${eq.vendor_name ? `<dl class="info">
          <dt>会社名</dt><dd>${h(eq.vendor_name)}</dd>
          <dt>担当者</dt><dd>${h(eq.vendor_contact)}</dd>
          <dt>TEL</dt><dd>${eq.vendor_tel ? `<a href="tel:${h(eq.vendor_tel)}">${h(eq.vendor_tel)}</a>` : ""}</dd>
          <dt>携帯</dt><dd>${eq.vendor_mobile ? `<a href="tel:${h(eq.vendor_mobile)}">${h(eq.vendor_mobile)}</a>` : ""}</dd>
          <dt>FAX</dt><dd>${h(eq.vendor_fax)}</dd>
          <dt>住所</dt><dd>${h(eq.vendor_address)}</dd>
        </dl>` : `<p class="empty">未登録</p>`}
        <h3 style="margin-top:18px">アワーメーター</h3>
        <form id="hours-form" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
          <input type="number" step="0.1" min="0" name="current_hours" value="${h(eq.current_hours ?? "")}" placeholder="現在値 (h)" inputmode="decimal" style="width:140px;padding:8px;border:1px solid #cfcec9;border-radius:6px">
          <button class="btn primary">更新</button>
          <span class="muted" style="font-size:13px">${eq.hours_updated ? `最終更新 ${fmtDate(eq.hours_updated)}` : "未入力"}</span>
        </form>
      </div>
    </div>

    <div class="section-title"><h2>定期交換・点検項目</h2><button class="btn small" id="add-rule">項目を追加</button></div>
    ${eq.rules.length ? `<div class="table-wrap"><table class="list responsive"><thead><tr><th>状態</th><th>項目</th><th>間隔</th><th>前回</th><th>次回期限</th><th>稼働時間</th></tr></thead><tbody>
      ${eq.rules.map(r => `<tr data-rule="${r.id}"><td>${lvBadge(r.level)}</td><td><strong>${h(r.title)}</strong></td>
        <td class="nowrap">${[r.interval_days ? `${r.interval_days}日` : "", r.interval_hours ? `${r.interval_hours}h` : ""].filter(Boolean).join(" / ")}</td>
        <td class="nowrap">${fmtDate(r.last_done_date) || "—"}</td>
        <td class="nowrap">${fmtDate(r.due_date) || "—"}</td>
        <td>${r.interval_hours ? hoursMeter(r) : "—"}</td></tr>`).join("")}
    </tbody></table></div>` : `<p class="muted">登録なし</p>`}

    ${eq.parts.length ? `<div class="section-title"><h2>専用部品</h2></div>
    <div class="table-wrap"><table class="list responsive"><thead><tr><th>シリアル</th><th>種別</th><th>状態</th><th>研磨回数</th></tr></thead><tbody>
      ${eq.parts.map(p => `<tr data-part="${p.id}"><td class="code">${h(p.serial)}</td><td>${h(p.part_type)}</td><td>${h(p.status)}</td><td>${grindMeter(p)}</td></tr>`).join("")}
    </tbody></table></div>` : ""}

    <div class="section-title"><h2>作業履歴（${recs.length}件）</h2><button class="btn primary" id="add-rec">作業を記録</button></div>
    <div class="card"><ul class="timeline">${recs.map(r => timelineItem({ ...r, equipment_name: "" })).join("") || '<li class="muted">履歴なし</li>'}</ul></div>`;

  $("#edit-eq").addEventListener("click", () => openEquipmentForm(eq));
  $("#add-rule").addEventListener("click", () => openRuleForm(null, { equipment_id: id }));
  $("#add-rec").addEventListener("click", () => openRecordForm(null, { equipment_id: id }));
  $("#hours-form").addEventListener("submit", async e => {
    e.preventDefault();
    try {
      await api(`/equipment/${id}/hours`, { method: "PUT", body: { current_hours: e.target.current_hours.value } });
      toast("アワーメーターを更新しました");
      router();
    } catch (ex) { toast(ex.message); }
  });
  bindRowClicks();
}

async function openEquipmentForm(eq) {
  const vendors = await api("/vendors");
  const e = eq || { status: "稼働", area: "前処理" };
  const f = (name, label, cls = "") => `<label class="field ${cls}"><span>${label}</span><input name="${name}" value="${h(e[name] ?? "")}"></label>`;
  openModal(eq ? "設備情報の編集" : "設備を追加", `
    <div class="form-grid">
      ${f("code", "設備No")}${f("nickname", "通称（現場での呼び名）")}
      ${f("name", "正式名 *", "full")}
      <label class="field"><span>種類</span><select name="category">${state.meta.categories.map(c => `<option ${e.category === c ? "selected" : ""}>${h(c)}</option>`).join("")}</select></label>
      ${f("model", "型式")}${f("maker", "メーカー")}
      ${f("supplier", "納入業者名")}
      <label class="field"><span>取引業者</span><select name="vendor_id"><option value="">なし</option>${vendors.map(v => `<option value="${v.id}" ${e.vendor_id == v.id ? "selected" : ""}>${h(v.name)}</option>`).join("")}</select></label>
      ${f("asset_no", "資産No")}${f("serial_no", "製造番号")}
      ${f("manufactured", "製造年")}${f("delivered", "納入年月日")}
      <label class="field"><span>状態</span><select name="status"><option ${e.status === "稼働" ? "selected" : ""}>稼働</option><option ${e.status === "廃止" ? "selected" : ""}>廃止</option></select></label>
      ${f("area", "エリア")}
      <label class="field full"><span>メモ</span><textarea name="note">${h(e.note || "")}</textarea></label>
    </div><p class="err" id="form-err" hidden></p>`, `
    <button type="button" class="btn" data-close>キャンセル</button><button type="submit" class="btn primary">保存</button>`);
  const form = $("#modal-form");
  form.onsubmit = async ev => {
    ev.preventDefault();
    const data = Object.fromEntries(new FormData(form));
    try {
      if (eq) await api(`/equipment/${eq.id}`, { method: "PUT", body: data });
      else { const r = await api("/equipment", { method: "POST", body: data }); location.hash = `#/equipment/${r.id}`; }
      state.equipment = [];
      closeModal(); toast("保存しました"); router();
    } catch (ex) { $("#form-err").textContent = ex.message; $("#form-err").hidden = false; }
  };
}

/* ================================================================ 部品・研磨管理 */
function grindMeter(p) {
  if (p.grind_count == null) return `<span class="muted">研磨回数不明</span>`;
  const limit = p.grind_limit || 10;
  const pct = Math.min(100, (p.grind_count / limit) * 100);
  const cls = p.grind_count >= limit ? "over" : p.grind_count >= limit - 2 ? "soon" : "";
  return `<div class="meter" data-tip="研磨 ${p.grind_count} 回 / 上限 ${limit} 回">
    <div class="meter-bar"><div class="meter-fill ${cls}" style="width:${pct}%"></div></div>
    <span class="meter-text">${p.grind_count}/${limit}回</span></div>`;
}

async function renderParts() {
  const parts = await api("/parts");
  view.innerHTML = `
    <div class="toolbar">
      <div class="filters" style="margin:0;flex:1">
        <label>状態 <select id="p-status"><option value="">使用可能のみ（使用中・予備・研磨中）</option><option value="all">すべて</option>${state.meta.part_statuses.map(s => `<option>${s}</option>`).join("")}</select></label>
      </div>
      <button class="btn primary lg" id="new-part">部品を登録</button>
    </div>
    <p class="muted" style="font-size:13px;margin:0 0 10px">肉盛り研磨は一般に10回まで（フィルターは内径φ110が限界、メーカーで現物判断）。「研磨出し」を記録すると回数が増えます。</p>
    <div class="table-wrap"><table class="list responsive"><thead><tr>
      <th>シリアル</th><th>種別</th><th>専用機</th><th>状態</th><th>研磨回数</th><th>最終記録</th><th>メモ</th>
    </tr></thead><tbody id="parts-body"></tbody></table></div>`;
  const draw = () => {
    const f = $("#p-status").value;
    const list = parts.filter(p => f === "all" ? true : f ? p.status === f : ["使用中", "予備", "研磨中"].includes(p.status));
    $("#parts-body").innerHTML = list.map(p => `
      <tr data-part="${p.id}">
        <td class="code" data-label="シリアル"><strong>${h(p.serial)}</strong></td>
        <td data-label="種別">${h(p.part_type)}</td>
        <td data-label="専用機">${p.equipment_name ? eqLabel(p) : '<span class="muted">共通/不明</span>'}</td>
        <td data-label="状態">${h(p.status)}</td>
        <td>${grindMeter(p)}</td>
        <td class="nowrap" data-label="最終記録">${fmtDate(p.last_event_date)}</td>
        <td class="clip" data-label="メモ">${h(p.note)}</td>
      </tr>`).join("") || `<tr><td colspan="7" class="empty">該当なし</td></tr>`;
    bindRowClicks($("#parts-body"));
  };
  $("#p-status").addEventListener("change", draw);
  $("#new-part").addEventListener("click", () => openPartForm(null));
  draw();
}

async function openPartDetail(id) {
  const p = await api(`/parts/${id}`);
  await loadEquipment();
  const eq = state.equipment.find(e => e.id === p.equipment_id);
  openModal(`部品 ${p.serial}`, `
    <dl class="info">
      <dt>種別</dt><dd>${h(p.part_type)}</dd>
      <dt>専用機</dt><dd>${eq ? h(eq.name) : "共通/不明"}</dd>
      <dt>状態</dt><dd>${h(p.status)}</dd>
      <dt>研磨回数</dt><dd>${grindMeter(p)}</dd>
      <dt>メモ</dt><dd>${h(p.note)}</dd>
    </dl>
    <h3 style="margin:18px 0 8px;font-size:15px">記録を追加</h3>
    <div class="form-grid" id="ev-form">
      <label class="field"><span>内容</span><select name="event">${state.meta.part_events.map(e => `<option>${e}</option>`).join("")}</select></label>
      <label class="field"><span>日付</span><input type="date" name="event_date" value="${state.meta.today}"></label>
      <label class="field"><span>アワーメーター（h）</span><input type="number" step="0.1" name="hour_meter" inputmode="decimal"></label>
      <label class="field"><span>メモ</span><input name="note"></label>
    </div>
    <h3 style="margin:18px 0 8px;font-size:15px">履歴</h3>
    <ul class="timeline">${p.events.map(ev => `<li style="cursor:default"><div class="tl-date">${fmtDate(ev.event_date) || "日付不明"}${ev.hour_meter != null ? ` ・ ${ev.hour_meter}h` : ""}${ev.equipment_name ? ` ・ ${h(ev.equipment_name)}` : ""}</div><div class="tl-title">${h(ev.event)}</div>${ev.note ? `<div class="tl-body">${h(ev.note)}</div>` : ""}</li>`).join("") || '<li class="muted">記録なし</li>'}</ul>`, `
    <button type="button" class="btn left" id="edit-part">部品情報を編集</button>
    <button type="button" class="btn" data-close>閉じる</button>
    <button type="submit" class="btn primary">記録する</button>`);
  $("#edit-part").addEventListener("click", () => openPartForm(p));
  $("#modal-form").onsubmit = async e => {
    e.preventDefault();
    const data = Object.fromEntries($$("#ev-form [name]").map(el => [el.name, el.value]));
    try {
      const r = await api(`/parts/${id}/events`, { method: "POST", body: data });
      closeModal();
      toast(r.message || "記録しました", r.message ? 6000 : 2600);
      router();
    } catch (ex) { toast(ex.message); }
  };
}

async function openPartForm(p) {
  await loadEquipment();
  const v = p || { part_type: "フィルタースクリュー", grind_count: 0, grind_limit: 10, status: "予備" };
  openModal(p ? "部品情報の編集" : "部品を登録", `
    <div class="form-grid">
      <label class="field"><span>種別 *</span><input name="part_type" value="${h(v.part_type)}"></label>
      <label class="field"><span>シリアル *</span><input name="serial" value="${h(v.serial || "")}" ${p ? "readonly" : ""}></label>
      <label class="field"><span>専用機</span><select name="equipment_id"><option value="">共通/不明</option>${equipmentOptions(e => inArea(e) || e.id == v.equipment_id, v.equipment_id)}</select></label>
      <label class="field"><span>状態</span><select name="status">${state.meta.part_statuses.map(s => `<option ${v.status === s ? "selected" : ""}>${s}</option>`).join("")}</select></label>
      <label class="field"><span>研磨回数</span><input type="number" min="0" name="grind_count" value="${h(v.grind_count ?? "")}"></label>
      <label class="field"><span>研磨上限</span><input type="number" min="1" name="grind_limit" value="${h(v.grind_limit ?? 10)}"></label>
      <label class="field full"><span>メモ</span><textarea name="note">${h(v.note || "")}</textarea></label>
    </div><p class="err" id="form-err" hidden></p>`, `
    <button type="button" class="btn" data-close>キャンセル</button><button type="submit" class="btn primary">保存</button>`);
  const form = $("#modal-form");
  form.onsubmit = async e => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(form));
    try {
      if (p) await api(`/parts/${p.id}`, { method: "PUT", body: data });
      else await api("/parts", { method: "POST", body: data });
      closeModal(); toast("保存しました"); router();
    } catch (ex) { $("#form-err").textContent = ex.message; $("#form-err").hidden = false; }
  };
}

/* ================================================================ 分析 */
async function renderAnalytics() {
  const years = state.analyticsYears || 5;
  const a = await api(`/analytics?years=${years}${areaQS("&")}`);
  const byYear = a.by_year.filter(y => y.year >= a.since.slice(0, 4));
  const eqs = a.by_equipment.filter(e => e.downtime > 0 || e.futeiki > 0);
  const kw = a.keywords.filter(k => k.count > 0);
  const totalDown = eqs.reduce((s, e) => s + e.downtime, 0);
  const totalCost = byYear.reduce((s, y) => s + y.cost, 0);
  const totalFail = byYear.reduce((s, y) => s + (y.futeiki || 0), 0);

  view.innerHTML = `
    <div class="toolbar">
      <div class="filters" style="margin:0">
        <label>集計期間 <select id="a-years">${[3, 5, 10].map(n => `<option value="${n}" ${n === years ? "selected" : ""}>直近${n}年（${new Date().getFullYear() - n + 1}年〜）</option>`).join("")}</select></label>
      </div>
    </div>
    <div class="grid kpi">
      <div class="kpi-tile"><div class="kpi-label">故障（不定期）件数</div><div class="kpi-value">${totalFail}<small>件</small></div></div>
      <div class="kpi-tile"><div class="kpi-label">停止・修理時間（記録分）</div><div class="kpi-value">${fmtNum(Math.round(totalDown / 60 * 10) / 10)}<small>時間</small></div></div>
      <div class="kpi-tile"><div class="kpi-label">修繕費用（記録分）</div><div class="kpi-value">${totalCost ? fmtNum(Math.round(totalCost / 10000)) : 0}<small>万円</small></div></div>
    </div>
    <div class="grid cols-2">
      <div class="card chart-card">
        <h3>年別の作業件数</h3>
        <div class="legend"><span><i class="sw" style="background:var(--series-1)"></i>定期</span><span><i class="sw" style="background:var(--series-2)"></i>不定期（故障）</span></div>
        ${stackedYearChart(byYear)}
        ${tableView(["年", "定期", "不定期", "区分なし", "停止・修理(分)"], byYear.map(y => [y.year, y.teiki, y.futeiki, y.other, y.downtime]))}
      </div>
      <div class="card chart-card">
        <h3>設備別の停止・修理時間 <span class="sub">分・時間の記載がある履歴のみ</span></h3>
        ${hbarChart(eqs.filter(e => e.downtime > 0).slice(0, 10).map(e => ({ label: e.name, value: e.downtime, tip: `<div class='tip-title'>${h(e.name)}</div>停止・修理 ${fmtNum(e.downtime)}分 / 故障 ${e.futeiki}件` })), v => `${fmtNum(v)}分`)}
        ${tableView(["設備", "停止・修理(分)", "故障件数", "費用"], eqs.map(e => [e.name, e.downtime, e.futeiki, yen(e.cost)]))}
      </div>
      <div class="card chart-card">
        <h3>故障内容のキーワード <span class="sub">不定期の履歴に出てくる回数</span></h3>
        ${hbarChart(kw.map(k => ({ label: k.keyword, kw: k.keyword, value: k.count, tip: `<div class='tip-title'>${h(k.keyword)}</div>${k.count}件` })), v => `${v}件`)}
        <p class="muted" style="font-size:12px">クリックすると、そのキーワードで履歴を検索します。</p>
      </div>
      <div class="card chart-card">
        <h3>設備別の故障件数</h3>
        ${hbarChart(eqs.filter(e => e.futeiki > 0).sort((x, y) => y.futeiki - x.futeiki).slice(0, 10).map(e => ({ label: e.name, value: e.futeiki, tip: `<div class='tip-title'>${h(e.name)}</div>故障 ${e.futeiki}件 / 全${e.count}件` })), v => `${v}件`)}
      </div>
    </div>`;
  $("#a-years").addEventListener("change", e => { state.analyticsYears = +e.target.value; renderAnalytics(); });
  $$(".hrow[data-kw]").forEach(r => r.addEventListener("click", () => {
    state.recordQuery = { q: r.dataset.kw.split("・")[0], sort: "done_date", dir: "desc", kind: "不定期" };
    location.hash = "#/records";
  }));
}

function niceMax(v) {
  if (v <= 5) return 5;
  const p = Math.pow(10, Math.floor(Math.log10(v)));
  return Math.ceil(v / p) * p;
}

function stackedYearChart(rowsIn) {
  if (!rowsIn.length) return `<p class="empty">データなし</p>`;
  const max = niceMax(Math.max(...rowsIn.map(y => (y.teiki || 0) + (y.futeiki || 0) + (y.other || 0))));
  const ticks = [0.5, 1].map(t => Math.round(max * t));
  const cols = rowsIn.map(y => {
    const segs = [["teiki", "var(--series-1)", "定期"], ["futeiki", "var(--series-2)", "不定期"]].filter(([k]) => y[k] > 0);
    const total = (y.teiki || 0) + (y.futeiki || 0) + (y.other || 0);
    const tipHtml = `<div class='tip-title'>${y.year}年</div>
      <div class='tip-row'><i class='sw' style='background:var(--series-1)'></i>定期 ${y.teiki || 0}件</div>
      <div class='tip-row'><i class='sw' style='background:var(--series-2)'></i>不定期 ${y.futeiki || 0}件</div>
      ${y.other ? `<div class='tip-row'>区分なし ${y.other}件</div>` : ""}`;
    return `<div class="vcol" data-tip="${h(tipHtml)}">
      <div class="total">${total}</div>
      ${segs.map(([k, c], i) => `<div class="seg-fill ${i === 0 ? "top" : ""}" style="height:${(y[k] / max) * 100}%;background:${c}"></div>`).join("")}
      ${y.other ? `<div class="seg-fill ${segs.length ? "" : "top"}" style="height:${(y.other / max) * 100}%;background:#b9b8b3"></div>` : ""}
    </div>`;
  }).join("");
  return `<div class="vbars">${ticks.map(t => `<div class="gridline" style="bottom:${(t / max) * 100}%"><span>${t}</span></div>`).join("")}${cols}</div>
    <div class="vlabels">${rowsIn.map(y => `<span>${y.year}</span>`).join("")}</div>`;
}

function hbarChart(items, fmt) {
  if (!items.length) return `<p class="empty">データなし</p>`;
  const max = Math.max(...items.map(i => i.value));
  return `<div class="hbars">${items.map(i => `
    <div class="hrow" data-tip="${h(i.tip)}" ${i.kw ? `data-kw="${h(i.kw)}" style="cursor:pointer"` : ""}>
      <span class="hname" title="${h(i.label)}">${h(i.label)}</span>
      <span class="htrack"><span class="hfill" style="width:${(i.value / max) * 85}%"></span><span class="hval">${fmt(i.value)}</span></span>
    </div>`).join("")}</div>`;
}

function tableView(head, body) {
  return `<details class="table-view"><summary>表で見る</summary>
    <table class="list plain"><thead><tr>${head.map(x => `<th>${h(x)}</th>`).join("")}</tr></thead>
    <tbody>${body.map(r => `<tr>${r.map((c, i) => `<td class="${i ? "num" : ""}">${h(c ?? "")}</td>`).join("")}</tr>`).join("")}</tbody></table></details>`;
}

/* ================================================================ 取引業者 */
/* ================================================================ 通知設定 */
async function renderNotify() {
  const d = await api("/notify");
  const st = d.status;
  const areas = ["*", ...state.meta.areas];
  const areaLabel = a => (a === "*" ? "全エリア共通" : a);
  view.innerHTML = `
    <div class="grid cols-2">
      <div class="card">
        <h3>送信設定 ${st.configured ? lvBadge("ok", "✓ 設定済み") : lvBadge("over", "未設定")}</h3>
        <dl class="info">
          <dt>メールサーバー</dt><dd>${st.host ? `${h(st.host)}:${st.port}${st.starttls ? "（STARTTLS）" : ""}${st.auth ? "（認証あり）" : ""}` : '<span class="muted">未設定</span>'}</dd>
          <dt>差出人</dt><dd>${h(st.from_addr) || '<span class="muted">未設定</span>'}</dd>
          <dt>本文のリンク</dt><dd>${h(st.app_url)}</dd>
          <dt>送るタイミング</dt><dd>平日の朝（タスクスケジューラで設定）。期限超過のままの項目は毎週${h(st.reminder_weekday)}曜にまとめて送る</dd>
          <dt>通知済み</dt><dd>${d.sent_count}件${d.last_sent ? `（最終 ${h(d.last_sent)}）` : ""}</dd>
        </dl>
        <p class="muted" style="font-size:13px">サーバーやパスワードは画面からは変更できません。サーバー上の <span class="code">${h(st.config_path)}</span> を編集してください${st.config_exists ? "" : "（まだありません。notify.ini.example をコピーして作成）"}。</p>
        <form id="test-form" style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px">
          <input type="email" name="email" placeholder="テスト送信先のアドレス" required style="flex:1;min-width:200px;padding:8px;border:1px solid #cfcec9;border-radius:6px">
          <button class="btn" ${st.configured ? "" : "disabled"}>テスト送信</button>
        </form>
      </div>
      <div class="card">
        <h3>宛先 <span class="sub">エリアごと。複数はカンマか改行で区切る</span></h3>
        <form id="recip-form" class="form-grid">
          ${areas.map(a => `<label class="field ${a === "*" ? "full" : ""}"><span>${h(areaLabel(a))}${a === "*" ? '<em class="hint" style="font-weight:400"> ・全エリアの通知が届く</em>' : ""}</span>
            <textarea name="${h(a)}" rows="2" placeholder="name@example.co.jp">${h((d.recipients[a] || []).join("\n"))}</textarea></label>`).join("")}
          <div class="full"><button class="btn primary">宛先を保存</button> <span class="err" id="recip-err" hidden></span></div>
        </form>
      </div>
    </div>
    <div class="section-title"><h2>送信内容のプレビュー</h2>
      <div style="display:flex;gap:8px;align-items:center">
        <label class="muted" style="font-size:13px"><input type="checkbox" id="pv-all"> 通知済みも含める</label>
        <button class="btn" id="pv-btn">今送るとどうなるか見る</button>
      </div>
    </div>
    <div id="preview"><p class="muted">ボタンを押すと、次回の送信で届くメールを表示します（実際には送りません）。</p></div>`;

  $("#recip-form").addEventListener("submit", async e => {
    e.preventDefault();
    try {
      const r = await api("/notify/recipients", { method: "PUT", body: Object.fromEntries(new FormData(e.target)) });
      $("#recip-err").hidden = true;
      toast(`宛先を保存しました（${r.count}件）`);
    } catch (ex) { $("#recip-err").textContent = ex.message; $("#recip-err").hidden = false; }
  });
  $("#test-form").addEventListener("submit", async e => {
    e.preventDefault();
    try {
      await api("/notify/test", { method: "POST", body: { email: e.target.email.value } });
      toast("テストメールを送信しました");
    } catch (ex) { toast(ex.message, 6000); }
  });
  $("#pv-btn").addEventListener("click", async () => {
    const r = await api("/notify/preview" + ($("#pv-all").checked ? "?all=1" : ""));
    $("#preview").innerHTML =
      r.warnings.map(w => `<p class="err">⚠ ${h(w)}</p>`).join("") +
      (r.messages.length ? r.messages.map(m => `
        <div class="card mail-preview">
          <div class="mail-head"><div><span class="muted">宛先</span> ${h(m.to.join(", "))}</div><div><span class="muted">件名</span> <strong>${h(m.subject)}</strong></div></div>
          <pre>${h(m.body)}</pre>
        </div>`).join("") : `<p class="empty">今送る項目はありません${r.warnings.length ? "（宛先未設定のエリアを除く）" : ""}</p>`);
  });
}

async function renderVendors() {
  const vendors = await api("/vendors");
  view.innerHTML = `<div class="table-wrap"><table class="list responsive plain"><thead><tr>
    <th>会社名</th><th>担当者</th><th>TEL</th><th>携帯</th><th>FAX</th><th>住所</th><th>担当設備</th></tr></thead><tbody>
    ${vendors.map(v => `<tr>
      <td><strong>${h(v.name)}</strong></td>
      <td data-label="担当者">${h(v.contact)}</td>
      <td class="nowrap" data-label="TEL">${v.tel ? `<a href="tel:${h(v.tel)}">${h(v.tel)}</a>` : ""}</td>
      <td class="nowrap" data-label="携帯">${v.mobile ? `<a href="tel:${h(v.mobile)}">${h(v.mobile)}</a>` : ""}</td>
      <td class="nowrap" data-label="FAX">${h(v.fax)}</td>
      <td data-label="住所">${h(v.address)}</td>
      <td data-label="設備" style="font-size:13px">${h(v.equipment_names)}</td></tr>`).join("")}
    </tbody></table></div>`;
}

/* ================================================================ モーダル */
const modal = $("#modal");
function openModal(title, body, foot) {
  $("#modal-title").textContent = title;
  $("#modal-body").innerHTML = body;
  $("#modal-foot").innerHTML = foot;
  $("#modal-form").onsubmit = e => e.preventDefault();
  if (!modal.open) modal.showModal();
  $("#modal-body").scrollTop = 0;
}
function closeModal() { modal.close(); }
modal.addEventListener("click", e => {
  if (e.target.closest("[data-close]")) closeModal();
  else if (e.target === modal) closeModal();
});

/* ================================================================ 起動 */
function rememberUser(name) {
  try { localStorage.setItem("maint-user", name); } catch { /* 保存できなくても動作には影響しない */ }
}

async function init() {
  try { $("#current-user").value = localStorage.getItem("maint-user") || ""; } catch { /* noop */ }
  $("#current-user").addEventListener("change", e => rememberUser(e.target.value.trim()));
  $("#menu-btn").addEventListener("click", () => { document.body.classList.add("nav-open"); $("#scrim").hidden = false; });
  $("#scrim").addEventListener("click", () => { document.body.classList.remove("nav-open"); $("#scrim").hidden = true; });
  $("#fab").addEventListener("click", () => openRecordForm(null));
  state.meta = await api("/meta");
  try { state.area = localStorage.getItem("maint-area") || ""; } catch { /* noop */ }
  if (state.area && !state.meta.areas.includes(state.area)) state.area = "";
  const areaSel = $("#area");
  areaSel.innerHTML = `<option value="">全エリア</option>` + state.meta.areas.map(a => `<option ${a === state.area ? "selected" : ""}>${h(a)}</option>`).join("");
  areaSel.addEventListener("change", () => {
    state.area = areaSel.value;
    try { localStorage.setItem("maint-area", state.area); } catch { /* noop */ }
    state.recordQuery.equipment_id = "";
    router();
  });
  $("#author-list").innerHTML = state.meta.authors.map(a => `<option value="${h(a)}">`).join("");
  window.addEventListener("hashchange", router);
  router();
}
init();
