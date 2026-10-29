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
  if (opts.body && typeof opts.body !== "string" && !(opts.body instanceof FormData)) {
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
  if (k === "工事") return `<span class="kind koji">工事</span>`;
  if (k === "臨時") return `<span class="kind rinji">臨時</span>`;
  return `<span class="muted">—</span>`;
}
const LV_LABEL = { over: "⚠ 期限超過", soon: "● まもなく", ok: "✓ 余裕あり", unknown: "？ 基準なし", none: "—" };
function lvBadge(lv, text) {
  return `<span class="lv ${lv}">${h(text || LV_LABEL[lv])}</span>`;
}
function eqLabel(r) {
  const loc = r.location ? `<span class="loc">場所 ${h(r.location)}</span>` : "";
  return `${r.equipment_code ? `<span class="code">${h(r.equipment_code)}</span> ` : ""}${h(r.equipment_name)}${r.equipment_name && loc ? " " : ""}${loc}`;
}
function dateCell(r) {
  const shown = r.done_date ? fmtDate(r.done_date) : r.date_raw || "日付なし";
  const chip = r.date_uncertain ? `<span class="chip-uncertain" title="元の表記: ${h(r.date_raw)}">要確認</span>` : "";
  const time = timeRange(r);
  return `${h(shown)}${time ? ` <span class="time">${h(time)}</span>` : ""}${chip}`;
}
/* 工事の時間（9:00〜12:00 / 9:00〜 / 〜12:00） */
function timeRange(r) {
  if (!r.start_time && !r.end_time) return "";
  const t = s => (s ? s.replace(/^0(\d)/, "$1") : "");
  return `${t(r.start_time)}〜${t(r.end_time)}`;
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
  calendar: ["工事点検計画", renderCalendar],
  plans: ["交換アラート", renderPlans],
  equipment: ["設備台帳", renderEquipment],
  parts: ["部品・研磨管理", renderParts],
  analytics: ["分析", renderAnalytics],
  vendors: ["取引業者", renderVendors],
  notify: ["通知設定", renderNotify],
  pla: ["プラ点検", renderPla],
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

/* 添付ファイル 1 件の表示（url があればリンク、btnAttr のボタンで外す） */
function fileChip(name, size, url, btnAttr, isNew = false) {
  const icon = /\.pdf$/i.test(name) ? "PDF" : "写真";
  const kb = size ? (size >= 1024 * 1024 ? `${(size / 1024 / 1024).toFixed(1)}MB` : `${Math.ceil(size / 1024)}KB`) : "";
  const label = `<span class="rf-name"><span class="rf-type">${icon}</span>${h(name)} <span class="hint">${kb}${isNew ? "・未保存" : ""}</span></span>`;
  return `<span class="rf-item">${url ? `<a href="${h(url)}" target="_blank" rel="noopener">${label}</a>` : label}
    <button type="button" class="rf-del" ${btnAttr} aria-label="外す">✕</button></span>`;
}

function timelineItem(r) {
  return `<li class="${r.kind === "定期" ? "teiki" : r.kind === "不定期" ? "futeiki" : ""}" data-record="${r.id}">
    <div class="tl-date">${dateCell(r)} ${kindBadge(r.kind)} ${r.equipment_name || r.location ? eqLabel(r) : ""}${r.file_count ? ` <span class="rf-count" title="添付ファイル">📎${r.file_count}</span>` : ""}</div>
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
      <label>区分 <select id="f-kind"><option value="">すべて</option><option ${q.kind === "定期" ? "selected" : ""}>定期</option><option ${q.kind === "不定期" ? "selected" : ""}>不定期</option><option ${q.kind === "臨時" ? "selected" : ""}>臨時</option><option ${q.kind === "工事" ? "selected" : ""}>工事</option></select></label>
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
  /* 工事は設備に限らない（建屋・配管など）。設備は任意にして場所を入れられるようにする */
  const koji = rec.kind === "工事";
  const body = `
    <div class="form-grid">
      <div class="field full"><span>設備<em class="req" id="eq-req">*</em></span>
        <div class="eq-row"><select name="equipment_id"><option value="">選択してください</option>${eqOptions}</select>
          <button type="button" class="btn koji-toggle" id="koji-toggle" aria-pressed="${koji}">工事</button></div></div>
      <label class="field full" id="loc-field"><span>場所<em class="req">*</em><em class="hint" style="font-weight:400">（設備を選ばない場合）</em></span><input name="location" list="location-list" value="${h(rec.location || "")}" placeholder="例: 2階 前処理室 天井配管"><datalist id="location-list"></datalist></label>
      <input type="hidden" name="koji" value="${koji ? "1" : ""}">
      <div class="field" id="kind-field"><span>区分</span><div class="seg">
        <label><input type="radio" name="kind" value="定期" ${rec.kind === "定期" ? "checked" : ""}>定期</label>
        <label><input type="radio" name="kind" value="不定期" ${rec.kind === "不定期" ? "checked" : ""}>不定期（故障）</label>
        <label><input type="radio" name="kind" value="臨時" ${rec.kind === "臨時" ? "checked" : ""}>臨時</label>
      </div></div>
      <label class="field"><span>作業状況</span><select name="status">${state.meta.record_statuses.map(s => `<option ${rec.status === s ? "selected" : ""}>${s}</option>`).join("")}</select></label>
      <label class="field"><span>実施日<em class="hint" style="font-weight:400">（未完了なら予定日）</em></span><input type="date" name="done_date" value="${h(rec.done_date || "")}">
        ${rec.date_uncertain ? `<span class="hint">⚠ 元の表記「${h(rec.date_raw)}」。正しい日付を入れると要確認が外れます</span>` : ""}</label>
      <div class="field" id="time-field"><span>時間</span>
        <div class="time-row"><input type="time" name="start_time" value="${h(rec.start_time || "")}" aria-label="開始時刻">〜<input type="time" name="end_time" value="${h(rec.end_time || "")}" aria-label="終了時刻"></div></div>
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
      <div class="field full"><span>添付ファイル<em class="hint" style="font-weight:400">（PDF・写真。報告書・見積書・現場写真など）</em></span>
        <div class="rf-list" id="rf-list"></div>
        <div><button type="button" class="btn small" id="rf-add">ファイルを追加</button> <span class="hint">1ファイル30MBまで。保存すると添付されます</span></div>
      </div>
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

  const setKoji = on => {
    form.koji.value = on ? "1" : "";
    $("#koji-toggle").setAttribute("aria-pressed", on);
    $("#eq-req").hidden = on;
    $("#loc-field").hidden = !on;
    $("#time-field").hidden = !on;
    $("#kind-field").hidden = on;
  };
  $("#koji-toggle").addEventListener("click", () => {
    setKoji(!form.koji.value);
    if (form.koji.value) form.location.focus();
  });
  setKoji(koji);
  api("/records/locations").then(locs => {
    $("#location-list").innerHTML = locs.map(l => `<option value="${h(l)}">`).join("");
  }).catch(() => { /* 候補は補助なので無視 */ });

  /* 添付ファイル: 登録済みはすぐ外せる。新しく選んだものは保存のときに送る */
  const savedFiles = rec.files || [];
  let pending = [];
  const drawFiles = () => {
    $("#rf-list").innerHTML = [
      ...savedFiles.map(f => fileChip(f.name, f.size, `/record-files/${f.file}`, `data-rf="${f.id}"`)),
      ...pending.map((f, i) => fileChip(f.name, f.size, "", `data-pend="${i}"`, true)),
    ].join("") || `<span class="hint">なし</span>`;
    $$("#rf-list [data-rf]").forEach(b => b.addEventListener("click", async () => {
      if (!confirm("この添付ファイルを外しますか？")) return;
      try {
        await api(`/record-files/${b.dataset.rf}`, { method: "DELETE" });
        savedFiles.splice(savedFiles.findIndex(f => f.id == b.dataset.rf), 1);
        drawFiles();
      } catch (ex) { toast(ex.message); }
    }));
    $$("#rf-list [data-pend]").forEach(b => b.addEventListener("click", () => { pending.splice(+b.dataset.pend, 1); drawFiles(); }));
  };
  $("#rf-add").addEventListener("click", () => {
    const inp = document.createElement("input");
    inp.type = "file"; inp.multiple = true;
    inp.accept = "application/pdf,.pdf,image/jpeg,.jpg,.jpeg,image/png,.png";
    inp.onchange = () => {
      const big = [...inp.files].filter(f => f.size > 30 * 1024 * 1024);
      if (big.length) toast(`30MB を超えるファイルは添付できません: ${big.map(f => f.name).join("、")}`, 4000);
      pending.push(...[...inp.files].filter(f => f.size <= 30 * 1024 * 1024));
      drawFiles();
    };
    inp.click();
  });
  drawFiles();

  form.onsubmit = async e => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(form));
    const isKoji = !!data.koji;
    delete data.koji;
    data.kind = isKoji ? "工事" : data.kind || null;
    if (!isKoji) data.location = data.start_time = data.end_time = "";
    const err = $("#form-err");
    const msg = !data.content.trim() ? "実施内容は必須です"
      : isKoji ? (!data.equipment_id && !data.location.trim() ? "工事は設備か場所のどちらかを入力してください" : "")
      : !data.equipment_id ? "設備と実施内容は必須です" : "";
    if (msg) {
      err.textContent = msg;
      err.hidden = false;
      return;
    }
    let recId = id;
    try {
      if (id) await api(`/records/${id}`, { method: "PUT", body: data });
      else recId = (await api("/records", { method: "POST", body: data })).id;
      if (data.author) rememberUser(data.author);
      if (pending.length) {
        const fd = new FormData();
        pending.forEach(f => fd.append("file", f));
        try { await api(`/records/${recId}/files`, { method: "POST", body: fd }); }
        catch (ex) {  // 記録は保存できているので、開き直してファイルだけ付け直してもらう
          closeModal(); router();
          toast(`記録は保存しましたが、ファイルを添付できませんでした: ${ex.message}`, 6000);
          return;
        }
      }
      closeModal();
      toast(pending.length ? `保存しました（ファイル ${pending.length} 件を添付）` : "保存しました");
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

/* ================================================================ 工事点検計画（カレンダー） */
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
  const eq = ev.equipment_nickname || ev.equipment_name || ev.location || "";
  const drag = ev.type === "planned" || ev.type === "due";
  const start = ev.start_time ? ev.start_time.replace(/^0(\d)/, "$1") : "";
  /* 1 段目: 種類・時刻・設備（場所）、2 段目: 作業内容（2 行まで） */
  return `<div class="${cls}" data-ev="${i}" ${drag ? 'draggable="true"' : ""} title="${h(`${EV_TYPES[key].label}${late ? "（遅れ）" : ""}：${timeRange(ev)} ${eq}\n${ev.title || ""}`)}">
    <div class="ev-head"><span class="ev-tag">${late ? "遅れ" : EV_TYPES[key].label}</span>${ev.kind === "工事" || ev.kind === "臨時" ? `<span class="ev-kind">${ev.kind}</span>` : ""}${start ? `<span class="ev-time">${start}</span>` : ""}<span class="ev-eq">${h(eq)}</span></div>
    ${ev.title ? `<div class="ev-title">${h(ev.title)}</div>` : ""}</div>`;
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
    <div class="day-eq">${ev.kind === "工事" || ev.kind === "臨時" ? kindBadge(ev.kind) + " " : ""}${eqLabel({ equipment_code: ev.equipment_code, equipment_name: ev.equipment_nickname || ev.equipment_name, location: ev.location })}</div>
    <div class="day-title">${timeRange(ev) ? `<span class="time">${h(timeRange(ev))}</span> ` : ""}${h(ev.title || "")}</div>
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
      <a class="btn" href="/api/equipment/export.xlsx" download>Excel出力</a>
      <button class="btn" id="eq-import">Excel取込</button>
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
  $("#eq-import").addEventListener("click", () => openXlsxImport(EQUIPMENT_IMPORT));
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
        <div style="margin-top:12px"><button class="btn small" id="edit-eq">設備情報を編集</button> <button class="btn small" id="copy-eq">この設備をコピー</button></div>
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
    <div class="table-wrap"><table class="list responsive"><thead><tr><th>シリアル</th><th>種別</th><th>状態</th><th>研磨回数</th><th>保管場所</th></tr></thead><tbody>
      ${eq.parts.map(p => `<tr data-part="${p.id}"><td class="code">${h(p.serial)}</td><td>${h(p.part_type)}</td><td>${h(p.status)}</td><td>${grindMeter(p)}</td><td>${h(p.location || "")}</td></tr>`).join("")}
    </tbody></table></div>` : ""}

    ${eq.catalog.length ? `<div class="section-title"><h2>交換部品・単価（${eq.catalog.length}件）</h2><a class="btn small" href="#/parts/catalog">部品一覧へ</a></div>
    <div class="table-wrap"><table class="list responsive"><thead><tr><th>部品名</th><th class="num">単価</th><th class="num">数量</th><th>保管場所</th><th>備考</th></tr></thead><tbody>
      ${eq.catalog.map(c => `<tr style="cursor:default"><td data-label="部品名"><strong>${h(c.name)}</strong></td><td class="num nowrap" data-label="単価">${c.unit_price != null ? yen(c.unit_price) : '<span class="muted">未登録</span>'}</td><td class="num" data-label="数量">${c.qty != null ? fmtNum(c.qty) : ""}</td><td data-label="保管場所">${h(c.location || "")}</td><td class="clip" data-label="備考">${h([c.note, c.memo].filter(Boolean).join(" ／ "))}</td></tr>`).join("")}
    </tbody></table></div>` : ""}

    <div class="section-title"><h2 id="rec-count">作業履歴（${recs.length}件）</h2><button class="btn primary" id="add-rec">作業を記録</button></div>
    ${recs.length ? `<div class="search" style="margin-bottom:8px"><span>🔍</span><input id="rec-q" type="search" placeholder="作業内容・原因・結果・対策・作業者・日付で絞り込み"></div>` : ""}
    <div class="card"><ul class="timeline" id="rec-list"></ul></div>`;

  /* 作業履歴の絞り込み（スペース区切りで AND 検索） */
  const recText = r => norm([r.done_date, r.date_raw, r.kind, r.status, r.content, r.reason, r.result,
    r.countermeasure, r.author, r.part_serial, r.rule_title, r.file_names].filter(Boolean).join(" ")).toLowerCase();
  const drawRecs = () => {
    const words = norm($("#rec-q")?.value).toLowerCase().split(/\s+/).filter(Boolean);
    const hit = words.length ? recs.filter(r => { const t = recText(r); return words.every(w => t.includes(w)); }) : recs;
    $("#rec-count").textContent = words.length ? `作業履歴（${hit.length} / ${recs.length}件）` : `作業履歴（${recs.length}件）`;
    $("#rec-list").innerHTML = hit.map(r => timelineItem({ ...r, equipment_name: "" })).join("")
      || `<li class="muted">${recs.length ? "該当する履歴なし" : "履歴なし"}</li>`;
    bindRowClicks($("#rec-list"));
  };
  $("#edit-eq").addEventListener("click", () => openEquipmentForm(eq));
  $("#copy-eq").addEventListener("click", () => openEquipmentForm(null, eq));
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
  drawRecs();  // 履歴のクリックは drawRecs 側で付ける（上の bindRowClicks と二重にしない）
  $("#rec-q")?.addEventListener("input", drawRecs);
}

/* src を渡すとその設備をコピーした内容で「追加」の画面を開く（1台ごとに違う設備No・資産No・製造番号などは空にする） */
async function openEquipmentForm(eq, src) {
  const vendors = await api("/vendors");
  const e = eq || (src ? { ...src, code: "", asset_no: "", serial_no: "", manufactured: "", delivered: "", status: "稼働" }
    : { status: "稼働", area: "前処理" });
  const f = (name, label, cls = "") => `<label class="field ${cls}"><span>${label}</span><input name="${name}" value="${h(e[name] ?? "")}"></label>`;
  const copyRules = src?.rules?.length
    ? `<label class="field full"><span><input type="checkbox" name="copy_rules_from" value="${src.id}" checked> 定期交換・点検項目もコピーする（${src.rules.length}件。前回実施日は空になります）</span></label>` : "";
  openModal(eq ? "設備情報の編集" : src ? `「${eqTitle(src)}」をコピーして追加` : "設備を追加", `
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
      ${copyRules}
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
/* 保管場所の入力候補（今までに入れた場所） */
const locList = () => `<datalist id="loc-list">${(state.meta.locations || []).map(l => `<option value="${h(l)}">`).join("")}</datalist>`;

function grindMeter(p) {
  if (p.grind_count == null) return `<span class="muted">研磨回数不明</span>`;
  const limit = p.grind_limit || 10;
  const pct = Math.min(100, (p.grind_count / limit) * 100);
  const cls = p.grind_count >= limit ? "over" : p.grind_count >= limit - 2 ? "soon" : "";
  return `<div class="meter" data-tip="研磨 ${p.grind_count} 回 / 上限 ${limit} 回">
    <div class="meter-bar"><div class="meter-fill ${cls}" style="width:${pct}%"></div></div>
    <span class="meter-text">${p.grind_count}/${limit}回</span></div>`;
}

/* 部品・研磨管理の上のタブと Excel ボタン（両タブ共通。1つの Excel に両方のシートが入る） */
function partsHeader(tab) {
  return `<div class="tabs-row">
    <nav class="tabs">
      <a href="#/parts" class="${tab === "serial" ? "on" : ""}">研磨管理（シリアル）</a>
      <a href="#/parts/catalog" class="${tab === "catalog" ? "on" : ""}">部品一覧（価格表）</a>
    </nav>
    <div class="tabs-actions">
      <a class="btn" href="/api/parts/export.xlsx" download>Excel出力</a>
      <button class="btn" id="xlsx-import">Excel取込</button>
    </div>
  </div>`;
}

async function renderParts(tab) {
  if (tab === "catalog") return renderCatalog();
  const parts = await api("/parts");
  view.innerHTML = `
    ${partsHeader("serial")}
    <div class="toolbar">
      <div class="filters" style="margin:0;flex:1">
        <label>状態 <select id="p-status"><option value="">使用可能のみ（使用中・予備・研磨中）</option><option value="all">すべて</option>${state.meta.part_statuses.map(s => `<option>${s}</option>`).join("")}</select></label>
      </div>
      <button class="btn primary lg" id="new-part">部品を登録</button>
    </div>
    <p class="muted" style="font-size:13px;margin:0 0 10px">肉盛り研磨は一般に10回まで（フィルターは内径φ110が限界、メーカーで現物判断）。「研磨出し」を記録すると回数が増えます。</p>
    <div class="table-wrap"><table class="list responsive"><thead><tr>
      <th>シリアル</th><th>種別</th><th>専用機</th><th>状態</th><th>研磨回数</th><th>保管場所</th><th>最終記録</th><th>メモ</th>
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
        <td data-label="保管場所">${h(p.location || "")}</td>
        <td class="nowrap" data-label="最終記録">${fmtDate(p.last_event_date)}</td>
        <td class="clip" data-label="メモ">${h(p.note)}</td>
      </tr>`).join("") || `<tr><td colspan="8" class="empty">該当なし</td></tr>`;
    bindRowClicks($("#parts-body"));
  };
  $("#p-status").addEventListener("change", draw);
  $("#new-part").addEventListener("click", () => openPartForm(null));
  $("#xlsx-import").addEventListener("click", () => openXlsxImport(PARTS_IMPORT));
  draw();
}

/* ---------------------------------------------------------------- 部品一覧（価格表） */
const catState = { q: "", group: "" };

async function renderCatalog() {
  const [items] = await Promise.all([api("/catalog" + areaQS()), loadEquipment()]);
  view.innerHTML = `
    ${partsHeader("catalog")}
    <div class="toolbar">
      <div class="search"><span>🔍</span><input id="c-q" type="search" placeholder="部品名・型番・機械・保管場所・備考で絞り込み" value="${h(catState.q)}"></div>
      <button class="btn primary lg" id="new-cat">部品を追加</button>
    </div>
    <div class="cat-chips" id="c-groups"></div>
    <p class="muted" style="font-size:13px;margin:0 0 10px">交換部品・消耗品の単価の一覧です。まとめて直すときは「Excel出力」で書き出し、直したファイルを「Excel取込」で読み込みます。</p>
    <div id="c-body"></div>`;

  const draw = () => {
    const q = norm(catState.q).toLowerCase();
    const hit = items.filter(c => !q || norm([c.group_name, c.machine, c.name, c.note, c.memo, c.location, c.equipment_nickname, c.equipment_name].join(" ")).toLowerCase().includes(q));
    const groups = groupBy(hit, c => c.group_name || "未分類");
    if (catState.group && !groups.some(([g]) => g === catState.group)) catState.group = "";
    $("#c-groups").innerHTML = `<button class="chip ${catState.group ? "" : "on"}" data-g="">すべて <b>${hit.length}</b></button>` +
      groups.map(([g, l]) => `<button class="chip ${catState.group === g ? "on" : ""}" data-g="${h(g)}">${h(g)} <b>${l.length}</b></button>`).join("");
    $$("#c-groups .chip").forEach(b => b.addEventListener("click", () => { catState.group = b.dataset.g; draw(); }));
    const shown = groups.filter(([g]) => !catState.group || g === catState.group);
    $("#c-body").innerHTML = shown.length ? shown.map(([g, list]) => `
      <section class="cat-group">
        <h2 class="area-title">${h(g)} <span class="g-count">${list.length}件</span></h2>
        <div class="table-wrap"><table class="list responsive cat-table"><thead><tr>
          <th>機械・場所</th><th>部品名</th><th class="num">単価</th><th class="num">数量</th><th>保管場所</th><th>備考</th><th>設備</th>
        </tr></thead><tbody>${list.map(c => `
          <tr data-cat="${c.id}">
            <td data-label="機械・場所">${h(c.machine || "")}</td>
            <td data-label="部品名"><strong>${h(c.name)}</strong></td>
            <td class="num nowrap" data-label="単価">${c.unit_price != null ? yen(c.unit_price) : '<span class="muted">未登録</span>'}</td>
            <td class="num" data-label="数量">${c.qty != null ? fmtNum(c.qty) : ""}</td>
            <td data-label="保管場所">${c.location ? h(c.location) : '<span class="muted">—</span>'}</td>
            <td class="clip" data-label="備考">${h(c.note || "")}${c.memo ? `<div class="muted" style="font-size:12px">${h(c.memo)}</div>` : ""}</td>
            <td data-label="設備">${c.equipment_id ? `<a href="#/equipment/${c.equipment_id}" class="eq-link">${eqLabel({ equipment_code: c.equipment_code, equipment_name: c.equipment_nickname || c.equipment_name })}</a>` : '<span class="muted">—</span>'}</td>
          </tr>`).join("")}</tbody></table></div>
      </section>`).join("") : `<p class="empty">該当する部品はありません</p>`;
    $$("#c-body [data-cat]").forEach(tr => tr.addEventListener("click", e => {
      if (e.target.closest(".eq-link")) return;
      openCatalogForm(items.find(c => c.id == tr.dataset.cat));
    }));
  };
  let timer;
  $("#c-q").addEventListener("input", e => { clearTimeout(timer); timer = setTimeout(() => { catState.q = e.target.value; draw(); }, 200); });
  $("#new-cat").addEventListener("click", () => openCatalogForm(null, { group_name: catState.group || "" }));
  $("#xlsx-import").addEventListener("click", () => openXlsxImport(PARTS_IMPORT));
  draw();
}

async function openCatalogForm(c, preset = {}) {
  await loadEquipment();
  const v = c || { area: state.area || "前処理", ...preset };
  const groups = [...new Set((await api("/catalog")).map(x => x.group_name).filter(Boolean))];
  openModal(c ? "部品の編集" : "部品を追加", `
    <div class="form-grid">
      <label class="field"><span>区分</span><input name="group_name" list="cat-groups" value="${h(v.group_name || "")}" placeholder="例: Aライン"><datalist id="cat-groups">${groups.map(g => `<option value="${h(g)}">`).join("")}</datalist></label>
      <label class="field"><span>機械・場所</span><input name="machine" value="${h(v.machine || "")}" placeholder="例: DeJong成型機"></label>
      <label class="field full"><span>部品名<em class="req">*</em></span><input name="name" value="${h(v.name || "")}" placeholder="例: 打ち出しベルト"></label>
      <label class="field"><span>単価（円）</span><input type="number" min="0" name="unit_price" value="${h(v.unit_price ?? "")}" inputmode="numeric"></label>
      <label class="field"><span>数量</span><input type="number" min="0" step="any" name="qty" value="${h(v.qty ?? "")}" inputmode="decimal"><span class="hint">一度の交換で使う数</span></label>
      <label class="field full"><span>保管場所</span><input name="location" list="loc-list" value="${h(v.location || "")}" placeholder="例: 部品庫 棚A-2">${locList()}</label>
      <label class="field full"><span>備考</span><textarea name="note" rows="2">${h(v.note || "")}</textarea></label>
      <label class="field full"><span>補足</span><textarea name="memo" rows="2">${h(v.memo || "")}</textarea></label>
      <label class="field"><span>紐づける設備</span><select name="equipment_id"><option value="">なし</option>${equipmentOptions(e => e.status !== "廃止" || e.id == v.equipment_id, v.equipment_id)}</select><span class="hint">設備台帳の詳しい画面にも表示されます</span></label>
      <label class="field"><span>エリア</span><select name="area">${state.meta.areas.map(a => `<option ${a === v.area ? "selected" : ""}>${h(a)}</option>`).join("")}</select></label>
    </div><p class="err" id="form-err" hidden></p>`, `
    ${c ? `<button type="button" class="btn danger left" id="del">削除</button>` : ""}
    <button type="button" class="btn" data-close>キャンセル</button><button type="submit" class="btn primary">保存</button>`);
  const form = $("#modal-form");
  form.onsubmit = async ev => {
    ev.preventDefault();
    const data = Object.fromEntries(new FormData(form));
    if (!data.name.trim()) { $("#form-err").textContent = "部品名は必須です"; $("#form-err").hidden = false; return; }
    try {
      if (c) await api(`/catalog/${c.id}`, { method: "PUT", body: data });
      else await api("/catalog", { method: "POST", body: data });
      state.meta.locations = (await api("/meta")).locations;  // 保管場所の候補を最新にする
      closeModal(); toast("保存しました"); router();
    } catch (ex) { $("#form-err").textContent = ex.message; $("#form-err").hidden = false; }
  };
  if (c) $("#del").addEventListener("click", async () => {
    if (!confirm(`「${c.name}」を部品一覧から削除しますか？`)) return;
    await api(`/catalog/${c.id}`, { method: "DELETE" });
    closeModal(); toast("削除しました"); router();
  });
}

/* ---------------------------------------------------------------- Excel 取込（確認してから反映） */
const IMPORT_LABELS = {
  group_name: "区分", machine: "機械・場所", name: "部品名", unit_price: "単価", qty: "数量", note: "備考", memo: "補足",
  location: "保管場所", equipment_id: "紐づけ設備", area: "エリア",
  category: "種類", code: "設備No", nickname: "通称", model: "型式", maker: "メーカー", supplier: "納入業者", vendor_id: "取引業者",
  asset_no: "資産No", serial_no: "製造番号", manufactured: "製造年", delivered: "納入年月日", part_type: "種別", status: "状態", grind_count: "研磨回数", grind_limit: "研磨上限",
};

/* Excel 取込の画面（部品・設備台帳で共通）。cfg で送り先と確認内容の出し方を変える */
async function openXlsxImport(cfg) {
  await loadEquipment();
  openModal(cfg.title, `
    <p style="margin-top:0">「Excel出力」で書き出して直したファイル（.xlsx）を選んでください。<br>取り込む前に、${cfg.what}される内容を表示します。</p>
    <input type="file" id="imp-file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
    <div id="imp-result" style="margin-top:14px"></div>`, `
    <button type="button" class="btn" data-close>閉じる</button>
    <button type="submit" class="btn primary" id="imp-apply" disabled>この内容で取り込む</button>`);
  let file = null;
  const send = async apply => {
    const fd = new FormData();
    fd.append("file", file);
    return api(`${cfg.endpoint}${apply ? "?apply=1" : ""}`, { method: "POST", body: fd });
  };
  $("#imp-file").addEventListener("change", async e => {
    file = e.target.files[0];
    $("#imp-apply").disabled = true;
    if (!file) return;
    $("#imp-result").innerHTML = `<p class="muted">読み込み中…</p>`;
    try {
      const plan = await send(false);
      const n = cfg.count(plan);
      $("#imp-result").innerHTML = `
        ${plan.errors.length ? `<div class="imp-errors"><strong>直してから取り込んでください（${plan.errors.length}件）</strong><ul>${plan.errors.map(e => `<li>${h(e)}</li>`).join("")}</ul></div>` : ""}
        ${!plan.errors.length && !n ? `<p class="empty">変更はありません（Excel の内容は今のデータと同じです）</p>` : ""}
        ${cfg.preview(plan)}`;
      $("#imp-apply").disabled = !!plan.errors.length || !n;
    } catch (ex) {
      $("#imp-result").innerHTML = `<p class="err">${h(ex.message)}</p>`;
    }
  });
  $("#modal-form").onsubmit = async ev => {
    ev.preventDefault();
    if (!file) return;
    try {
      const r = (await send(true)).applied;
      closeModal();
      toast(cfg.done(r), 6000);
      if (cfg.after) await cfg.after();
      router();
    } catch (ex) {
      $("#imp-result").insertAdjacentHTML("afterbegin", `<p class="err">${h(ex.message)}</p>`);
    }
  };
}
const impSection = (title, list, fmt) => list.length
  ? `<details class="imp-sec" ${list.length <= 15 ? "open" : ""}><summary>${title} <b>${list.length}件</b></summary><ul>${list.map(fmt).join("")}</ul></details>` : "";
const impChanges = (ch, val, labels = IMPORT_LABELS) => Object.entries(ch).map(([k, [a, b]]) =>
  `<div class="chg"><b>${h(labels[k] || k)}</b> ${h(val(k, a))} → <strong>${h(val(k, b))}</strong></div>`).join("");

const PARTS_IMPORT = {
  title: "部品の Excel取込", what: "追加・変更・削除", endpoint: "/parts/import",
  count: p => ["add", "update", "delete"].reduce((s, k) => s + p.catalog[k].length, 0) + p.serial.add.length + p.serial.update.length,
  preview: p => {
    const c = p.catalog, s = p.serial;
    const eqName = id => { const e = state.equipment.find(x => x.id == id); return e ? (e.code ? e.code + " " : "") + eqTitle(e) : "なし"; };
    const val = (k, v) => (v == null || v === "" ? "（空欄）" : k === "equipment_id" ? eqName(v) : k === "unit_price" ? yen(v) : String(v));
    return `
      ${impSection("部品一覧：追加", c.add, a => `<li><span class="muted">${h(a.row)}行目</span> ${h([a.group_name, a.machine].filter(Boolean).join(" / "))} ／ <strong>${h(a.name)}</strong> ${a.unit_price != null ? yen(a.unit_price) : ""}</li>`)}
      ${impSection("部品一覧：変更", c.update, u => `<li><span class="muted">${h(u.row)}行目</span> <strong>${h(u.name)}</strong>${impChanges(u.changes, val)}</li>`)}
      ${impSection("部品一覧：削除", c.delete, d => `<li><span class="muted">${h(d.row)}行目</span> <strong>${h(d.name)}</strong></li>`)}
      ${impSection("研磨管理部品：追加", s.add, a => `<li><span class="muted">${h(a.row)}行目</span> <strong>${h(a.serial)}</strong> ${h(a.part_type || "")} ${h(a.status || "")}</li>`)}
      ${impSection("研磨管理部品：変更", s.update, u => `<li><span class="muted">${h(u.row)}行目</span> <strong>${h(u.serial)}</strong>${impChanges(u.changes, val)}</li>`)}`;
  },
  done: r => `取り込みました（部品一覧 追加${r.catalog.add}・変更${r.catalog.update}・削除${r.catalog.delete}／研磨管理 追加${r.serial.add}・変更${r.serial.update}）`,
};

const EQUIPMENT_IMPORT = {
  title: "設備台帳の Excel取込", what: "追加・変更", endpoint: "/equipment/import",
  count: p => p.add.length + p.update.length,
  preview: p => {
    const val = (k, v) => (v == null || v === "" ? "（空欄）" : k === "vendor_id" ? p.vendors[v] || v : String(v));
    return `
      ${impSection("設備：追加", p.add, a => `<li><span class="muted">${h(a.row)}行目</span> ${a.code ? `<span class="code">${h(a.code)}</span> ` : ""}<strong>${h(a.label)}</strong> <span class="muted">${h(a.area)} ／ ${h(a.category)}</span></li>`)}
      ${impSection("設備：変更", p.update, u => `<li><span class="muted">${h(u.row)}行目</span> <strong>${h(u.label)}</strong>${impChanges(u.changes, val, { ...IMPORT_LABELS, name: "正式名", note: "メモ" })}</li>`)}`;
  },
  done: r => `取り込みました（設備 追加${r.add}・変更${r.update}）`,
  after: async () => { state.equipment = []; state.meta = await api("/meta"); },  // 設備・エリアの一覧を読み直す
};

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
      <dt>保管場所</dt><dd>${h(p.location || "")}</dd>
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
      <label class="field full"><span>保管場所</span><input name="location" list="loc-list" value="${h(v.location || "")}" placeholder="例: 部品庫 棚A-2">${locList()}</label>
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
      state.meta.locations = (await api("/meta")).locations;  // 保管場所の候補を最新にする
      closeModal(); toast("保存しました"); router();
    } catch (ex) { $("#form-err").textContent = ex.message; $("#form-err").hidden = false; }
  };
}

/* ================================================================ プラ点検 */
/* チェック表（紙の「プラスチック部品チェック表」）の 1 行ごとに ○×－ をつける。行はプラリストの行と紐づく */
const PLA_DEFECTS = ["破損", "欠け", "割れ", "変形", "紛失", "ゆるみ", "摩耗"];
const plaState = {
  meta: null, process: null, date: null, timing: null, sheet: null, onlyTodo: false,
  month: null, q: "", showInactive: false, onlyUnlinked: false, editSheet: null, showOldSheets: false,
};
const plaStamp = () => `${state.meta.today}T${new Date().toTimeString().slice(0, 5)}`;
const hazardBadge = hz => (hz ? `<span class="hz hz-${h(hz)}" title="危害度 ${h(hz)}">${h(hz)}</span>` : "");
const photoUrl = f => `/pla-photos/${encodeURIComponent(f)}`;

async function plaInit() {
  plaState.meta ||= await api("/pla/meta");
  if (!plaState.process) {
    // 右上のエリア（前処理・後処理・タレ工程）に合わせて工程を選んでおく
    const byArea = Object.entries(plaState.meta.process_area).find(([, a]) => a === state.area);
    plaState.process = byArea ? byArea[0] : plaState.meta.processes[0];
  }
  plaState.date ||= state.meta.today;
  plaState.timing ||= new Date().getHours() < 12 ? "始業時" : "終業時";
  plaState.month ||= state.meta.today.slice(0, 7);
}

function plaHeader(tab) {
  const p = plaState.meta.processes;
  const t = (key, label) => `<a href="#/pla${key ? "/" + key : ""}" class="${tab === (key || "check") ? "on" : ""}">${label}</a>`;
  return `<div class="tabs-row">
    <nav class="tabs">${t("", "点検する")}${t("review", "確認・記録")}${t("sheets", "チェック表")}${t("items", "プラリスト")}</nav>
    <div class="seg small" id="pla-proc">${p.map(x => `<label><input type="radio" name="pla-proc" value="${h(x)}" ${x === plaState.process ? "checked" : ""}>${h(x)}</label>`).join("")}</div>
  </div>`;
}
function bindPlaHeader() {
  $$("#pla-proc input").forEach(r => r.addEventListener("change", () => {
    plaState.process = r.value; plaState.sheet = null; plaState.editSheet = null; router();
  }));
}

async function renderPla(tab) {
  await plaInit();
  if (tab === "review") return renderPlaReview();
  if (tab === "items") return renderPlaItems();
  if (tab === "sheets") return renderPlaSheets();
  return renderPlaCheck();
}

/* 写真を大きく見る（← → で切り替え） */
function openLightbox(photos, start = 0, caption = "") {
  let i = start;
  const box = document.createElement("div");
  box.className = "lightbox";
  const draw = () => {
    box.innerHTML = `<button class="lb-close" aria-label="閉じる">✕</button>
      ${photos.length > 1 ? `<button class="lb-prev" aria-label="前の写真">‹</button><button class="lb-next" aria-label="次の写真">›</button>` : ""}
      <img src="${photoUrl(photos[i].file)}" alt="">
      <div class="lb-cap">${h(caption)}${photos.length > 1 ? `　${i + 1} / ${photos.length}` : ""}</div>`;
    box.querySelector(".lb-close").onclick = close;
    box.querySelector(".lb-prev")?.addEventListener("click", e => { e.stopPropagation(); i = (i + photos.length - 1) % photos.length; draw(); });
    box.querySelector(".lb-next")?.addEventListener("click", e => { e.stopPropagation(); i = (i + 1) % photos.length; draw(); });
  };
  const key = e => {
    if (e.key === "Escape") close();
    if (e.key === "ArrowLeft" && photos.length > 1) { i = (i + photos.length - 1) % photos.length; draw(); }
    if (e.key === "ArrowRight" && photos.length > 1) { i = (i + 1) % photos.length; draw(); }
  };
  const close = () => { box.remove(); document.removeEventListener("keydown", key); };
  box.addEventListener("click", e => { if (e.target === box) close(); });
  document.addEventListener("keydown", key);
  draw();
  document.body.append(box);
}

const pointPhotos = p => (p.photos.length
  ? `<button type="button" class="pla-thumb" data-lb="${p.id}"><img src="${photoUrl(p.photos[0].file)}" alt="写真" loading="lazy">${p.photos.length > 1 ? `<span>${p.photos.length}</span>` : ""}</button>`
  : `<div class="pla-thumb none">写真なし</div>`);
const pointLinks = p => (p.items.length
  ? `<details class="pla-links"><summary>プラリスト ${p.items.length}件 ${[...new Set(p.items.map(i => i.hazard).filter(Boolean))].sort().map(hazardBadge).join("")}</summary>
      <ul>${p.items.map(i => `<li>${hazardBadge(i.hazard)}${h(i.item)} <span class="muted">${h([i.color, i.material_maker || i.material].filter(Boolean).join(" ・ "))}</span></li>`).join("")}</ul></details>`
  : "");

/* ---------------------------------------------------------------- 点検する */
async function renderPlaCheck() {
  const qs = `process=${encodeURIComponent(plaState.process)}&date=${plaState.date}&timing=${encodeURIComponent(plaState.timing)}`;
  const d = await api(`/pla/check?${qs}`);
  const res = d.results;  // point_id → 結果
  const sheets = d.sheets;
  const allPoints = sheets.flatMap(s => s.points);
  if (!plaState.sheet || !sheets.some(s => s.id === plaState.sheet)) {
    plaState.sheet = (sheets.find(s => s.points.some(p => !res[p.id])) || sheets[0] || {}).id;
  }
  view.innerHTML = `
    ${plaHeader("check")}
    <div class="pla-bar">
      <label class="pla-date">日付 <input type="date" id="pla-date" value="${h(plaState.date)}"></label>
      <div class="seg" id="pla-timing">${plaState.meta.timings.map(t => `<label><input type="radio" name="pla-timing" value="${t}" ${t === plaState.timing ? "checked" : ""}>${t}</label>`).join("")}</div>
      <div class="pla-progress" id="pla-progress"></div>
    </div>
    ${d.review ? `<p class="pla-reviewed">✓ ${h(d.review.reviewer)} さんが確認済みです（${h(fmtDate(d.review.reviewed_at.slice(0, 10)))}）。直したときは確認をやり直してもらってください。</p>` : ""}
    ${!currentUser() ? `<p class="err" style="font-size:15px">右上の「記入者」に名前を入れてから点検してください。</p>` : ""}
    <div class="pla-layout">
      <aside class="pla-eqs card">
        <select id="pla-eq-select" class="pla-eq-select">${sheets.map(s => `<option value="${s.id}" ${s.id === plaState.sheet ? "selected" : ""}>${h(s.title)}</option>`).join("")}</select>
        <ul id="pla-eq-list"></ul>
      </aside>
      <section class="pla-main" id="pla-main"></section>
    </div>`;
  bindPlaHeader();
  $("#pla-date").addEventListener("change", e => { if (e.target.value) { plaState.date = e.target.value; router(); } });
  $$("#pla-timing input").forEach(r => r.addEventListener("change", () => { plaState.timing = r.value; router(); }));
  if (!sheets.length) { $("#pla-main").innerHTML = `<p class="empty">チェック表がありません。「チェック表」タブで作ってください。</p>`; return; }

  const stat = pts => ({ done: pts.filter(p => res[p.id]).length, ng: pts.filter(p => res[p.id]?.result === "ng").length, total: pts.length });
  const drawProgress = () => {
    const s = stat(allPoints);
    const sh = sheets.filter(x => stat(x.points).done === x.points.length).length;
    $("#pla-progress").innerHTML = `
      <div class="meter-bar"><div class="meter-fill ${s.done === s.total ? "ok" : ""}" style="width:${Math.round((s.done / s.total) * 100)}%"></div></div>
      <span>表 <b>${sh}</b>/${sheets.length}・箇所 <b>${s.done}</b>/${s.total}</span>${s.ng ? `<span class="lv over">× 異常 ${s.ng}</span>` : ""}`;
  };
  const drawList = () => {
    $("#pla-eq-list").innerHTML = sheets.map(s => {
      const st = stat(s.points);
      const cls = st.ng ? "ng" : st.done === st.total ? "done" : st.done ? "part" : "";
      return `<li class="${cls} ${s.id === plaState.sheet ? "on" : ""}" data-s="${s.id}"><span class="pla-eq-name">${h(s.title)}</span><span class="pla-eq-count">${st.ng ? "× " : st.done === st.total ? "✓ " : ""}${st.done}/${st.total}</span></li>`;
    }).join("");
    $$("#pla-eq-list li").forEach(li => li.addEventListener("click", () => select(+li.dataset.s)));
    $$("#pla-eq-select option").forEach(o => {
      const st = stat(sheets.find(s => s.id == o.value).points);
      o.textContent = `${st.ng ? "× " : st.done === st.total ? "✓ " : ""}${sheets.find(s => s.id == o.value).title}（${st.done}/${st.total}）`;
    });
  };
  const save = async list => {
    try {
      await api("/pla/results", { method: "PUT", body: { date: plaState.date, timing: plaState.timing, checked_by: currentUser(), results: list } });
      return true;
    } catch (ex) { toast(ex.message, 4000); return false; }
  };
  const drawMain = () => {
    const sheet = sheets.find(s => s.id === plaState.sheet);
    const shown = plaState.onlyTodo ? sheet.points.filter(p => !res[p.id]) : sheet.points;
    const idx = sheets.indexOf(sheet);
    const next = sheets.slice(idx + 1).find(s => s.points.some(p => !res[p.id])) || sheets[idx + 1];
    $("#pla-main").innerHTML = `
      <div class="pla-eq-head">
        <h2>${h(sheet.title)}</h2>
        <div class="pla-eq-actions">
          <label class="muted" style="font-size:13px"><input type="checkbox" id="pla-todo" ${plaState.onlyTodo ? "checked" : ""}> 未点検だけ表示</label>
          <button class="btn" id="pla-all-ok">残りをすべて ○</button>
        </div>
      </div>
      ${sheet.photos.length ? `<div class="pla-place"><span class="muted">場所の写真</span>${sheet.photos.map((ph, i) => `<button type="button" class="pla-place-ph" data-i="${i}"><img src="${photoUrl(ph.file)}" alt="場所の写真" loading="lazy"></button>`).join("")}</div>` : ""}
      <ul class="pla-items">${shown.map((p, n) => {
        const r = res[p.id];
        return `<li class="pla-item ${r ? "r-" + r.result : ""}" data-id="${p.id}">
          ${pointPhotos(p)}
          <div class="pla-item-info">
            ${p.location ? `<div class="pla-loc">📍 ${h(p.location)}</div>` : ""}
            <div class="pla-item-name">${h(p.item_text)}</div>
            ${pointLinks(p)}
            ${r?.result === "ng" ? `<div class="pla-ng-note">× ${h(r.note || "")}</div>` : ""}
            ${r ? `<div class="pla-who">${h(r.checked_by || "")} ${h((r.checked_at || "").slice(11, 16))}</div>` : ""}
          </div>
          <div class="pla-btns">
            <button class="pla-btn ok ${r?.result === "ok" ? "on" : ""}" data-r="ok" title="異常なし">○</button>
            <button class="pla-btn ng ${r?.result === "ng" ? "on" : ""}" data-r="ng" title="異常あり">×</button>
            <button class="pla-btn na ${r?.result === "na" ? "on" : ""}" data-r="na" title="稼働なし・使っていない">－</button>
          </div>
        </li>`;
      }).join("") || `<li class="empty">このチェック表はすべて点検済みです</li>`}</ul>
      ${next && next !== sheet ? `<div class="pla-next"><button class="btn primary lg" id="pla-next">次のチェック表へ：${h(next.title)} →</button></div>` : ""}`;

    $("#pla-todo").addEventListener("change", e => { plaState.onlyTodo = e.target.checked; drawMain(); });
    $$(".pla-place-ph").forEach(b => b.addEventListener("click", () => openLightbox(sheet.photos, +b.dataset.i, `${sheet.title}　場所の写真`)));
    $$("#pla-main [data-lb]").forEach(b => b.addEventListener("click", () => {
      const p = sheet.points.find(x => x.id == b.dataset.lb);
      openLightbox(p.photos, 0, [p.location, p.item_text.replace(/\n/g, " ")].filter(Boolean).join("　"));
    }));
    $("#pla-all-ok").addEventListener("click", async () => {
      const todo = sheet.points.filter(p => !res[p.id]);
      if (!todo.length) return toast("このチェック表はすべて点検済みです");
      if (!(await save(todo.map(p => ({ point_id: p.id, result: "ok" }))))) return;
      todo.forEach(p => (res[p.id] = { result: "ok", checked_by: currentUser(), checked_at: plaStamp() }));
      toast(`${todo.length} 箇所を ○ にしました`);
      redraw();
    });
    $("#pla-next")?.addEventListener("click", () => { select(next.id); window.scrollTo({ top: 0, behavior: "smooth" }); });
    $$("#pla-main .pla-item").forEach(li => {
      const id = +li.dataset.id;
      li.querySelectorAll(".pla-btn").forEach(b => b.addEventListener("click", async () => {
        const cur = res[id];
        const r = b.dataset.r;
        if (r === "ng") return openPlaNg(sheet, sheet.points.find(p => p.id === id), cur, async note => {
          if (!(await save([{ point_id: id, result: "ng", note }]))) return false;
          res[id] = { result: "ng", note, checked_by: currentUser(), checked_at: plaStamp() };
          redraw();
          return true;
        });
        const clear = cur?.result === r;  // 同じボタンをもう一度押したら取り消し
        if (!(await save([{ point_id: id, result: clear ? "" : r }]))) return;
        if (clear) delete res[id];
        else res[id] = { result: r, checked_by: currentUser(), checked_at: plaStamp() };
        redraw();
      }));
    });
  };
  const redraw = () => { drawProgress(); drawList(); drawMain(); };
  const select = id => { plaState.sheet = id; $("#pla-eq-select").value = id; drawList(); drawMain(); };
  $("#pla-eq-select").addEventListener("change", e => select(+e.target.value));
  redraw();
}

function openPlaNg(sheet, p, cur, onSave) {
  openModal("異常あり", `
    <div class="pla-ng-head">
      ${p.photos.length ? `<img src="${photoUrl(p.photos[0].file)}" alt="">` : ""}
      <p><b>${h(sheet.title)}</b><br>${p.location ? h(p.location) + "／" : ""}${h(p.item_text.replace(/\n/g, " "))}</p>
    </div>
    <div class="pla-defects">${PLA_DEFECTS.map(x => `<button type="button" class="chip" data-x="${x}">${x}</button>`).join("")}</div>
    <label class="field full" style="margin-top:12px"><span>異常の内容<em class="req">*</em></span>
      <textarea name="ng_note" rows="3" placeholder="例: 角が欠けている。欠片は見つからず">${h(cur?.result === "ng" ? cur.note || "" : "")}</textarea></label>
    <p class="muted" style="font-size:13px">欠片が見つからないときは、製品への混入がないか責任者にすぐ報告してください。</p>
    <p class="err" id="form-err" hidden></p>`, `
    <button type="button" class="btn" data-close>キャンセル</button>
    <button type="submit" class="btn danger-fill">× 異常ありで記録</button>`);
  const ta = $("#modal-form [name=ng_note]");
  $$(".pla-defects .chip").forEach(c => c.addEventListener("click", () => {
    ta.value = ta.value.includes(c.dataset.x) ? ta.value : (ta.value ? ta.value + "、" : "") + c.dataset.x;
    ta.focus();
  }));
  $("#modal-form").onsubmit = async e => {
    e.preventDefault();
    if (!ta.value.trim()) { $("#form-err").textContent = "異常の内容を書いてください"; $("#form-err").hidden = false; return; }
    if (await onSave(ta.value.trim())) { closeModal(); toast("異常ありで記録しました"); }
  };
  setTimeout(() => ta.focus(), 50);
}

/* ---------------------------------------------------------------- 確認・記録 */
async function renderPlaReview() {
  const d = await api(`/pla/summary?process=${encodeURIComponent(plaState.process)}&month=${plaState.month}`);
  const [y, m] = plaState.month.split("-").map(Number);
  const last = new Date(y, m, 0).getDate();
  const dates = [];
  for (let i = last; i >= 1; i--) {
    const iso = `${plaState.month}-${String(i).padStart(2, "0")}`;
    if (iso <= state.meta.today) dates.push(iso);
  }
  const cell = (iso, t) => {
    const c = d.days[iso]?.[t];
    if (!c || (!c.done && !c.review)) return `<td class="pla-cell" data-d="${iso}" data-t="${t}"><span class="muted">未実施</span></td>`;
    return `<td class="pla-cell ${c.ng ? "has-ng" : ""}" data-d="${iso}" data-t="${t}">
      <div class="pla-cell-top"><b>${c.done}</b>/${d.total}${c.done >= d.total ? ' <span class="lv ok">完了</span>' : ' <span class="lv soon">途中</span>'}${c.ng ? ` <span class="lv over">× ${c.ng}</span>` : ""}</div>
      <div class="pla-cell-sub">${c.checkers ? "点検 " + h(c.checkers.replaceAll(",", "、")) : ""}</div>
      <div class="pla-cell-sub">${c.review ? `<span class="pla-ok">✓ 確認 ${h(c.review.reviewer)}</span>` : '<span class="muted">未確認</span>'}</div>
    </td>`;
  };
  const wd = iso => "日月火水木金土"[new Date(iso + "T00:00").getDay()];
  view.innerHTML = `
    ${plaHeader("review")}
    <div class="toolbar">
      <div class="cal-nav">
        <button class="btn small" id="pm-prev" aria-label="前の月">‹</button>
        <strong>${y}年${m}月</strong>
        <button class="btn small" id="pm-next" aria-label="次の月">›</button>
      </div>
      <span style="flex:1"></span>
      <a class="btn" href="/api/pla/export.xlsx?process=${encodeURIComponent(plaState.process)}&month=${plaState.month}" download>この月の点検記録をExcel出力</a>
    </div>
    ${d.ng.length ? `<div class="card pla-ng-card"><h3>この月の異常 <span class="sub">${d.ng.length}件</span></h3>
      <div class="table-wrap"><table class="list responsive"><thead><tr><th>日付</th><th>チェック表・箇所</th><th>内容</th><th>点検者</th><th></th></tr></thead><tbody>
      ${d.ng.map(r => `<tr><td class="nowrap" data-label="日付">${fmtDate(r.check_date)} ${h(r.timing)}</td>
        <td data-label="チェック表・箇所"><b>${h(r.sheet_title)}</b><br>${h([r.location, r.item_text.replace(/\n/g, " ")].filter(Boolean).join("／"))}</td>
        <td data-label="内容">${h(r.note || "")}</td><td class="nowrap" data-label="点検者">${h(r.checked_by || "")}</td>
        <td class="nowrap"><button class="btn small pla-fix" data-i="${r.id}">修理を記録</button></td></tr>`).join("")}
      </tbody></table></div></div>` : ""}
    <p class="muted" style="font-size:13px;margin:8px 0">使っているチェック表の箇所は ${d.total} 箇所です。マスを押すと中身を見て「確認」できます。</p>
    <div class="table-wrap"><table class="list pla-month"><thead><tr><th>日付</th>${plaState.meta.timings.map(t => `<th>${t}</th>`).join("")}</tr></thead><tbody>
      ${dates.map(iso => `<tr class="${wd(iso) === "日" ? "sun" : wd(iso) === "土" ? "sat" : ""}"><td class="nowrap"><b>${Number(iso.slice(8))}日</b>（${wd(iso)}）</td>${plaState.meta.timings.map(t => cell(iso, t)).join("")}</tr>`).join("") || `<tr><td colspan="3" class="empty">まだ日付がありません</td></tr>`}
    </tbody></table></div>`;
  const move = step => { const dd = new Date(y, m - 1 + step, 1); plaState.month = `${dd.getFullYear()}-${String(dd.getMonth() + 1).padStart(2, "0")}`; router(); };
  $("#pm-prev").addEventListener("click", () => move(-1));
  $("#pm-next").addEventListener("click", () => move(1));
  $$(".pla-cell").forEach(td => td.addEventListener("click", () => openPlaDay(td.dataset.d, td.dataset.t)));
  $$(".pla-fix").forEach(b => b.addEventListener("click", () => {
    const r = d.ng.find(x => x.id == b.dataset.i);
    openRecordForm(null, { kind: "不定期", status: "未着手", content: `プラ点検で異常：${r.sheet_title}`, reason: `${fmtDate(r.check_date)} ${r.timing}　${r.sheet_title}／${[r.location, r.item_text.replace(/\n/g, " ")].filter(Boolean).join(" ")}：${r.note || ""}（点検者 ${r.checked_by || ""}）` });
  }));
  bindPlaHeader();
}

async function openPlaDay(date, timing) {
  const d = await api(`/pla/check?process=${encodeURIComponent(plaState.process)}&date=${date}&timing=${encodeURIComponent(timing)}`);
  const res = d.results;
  const pts = d.sheets.flatMap(s => s.points.map(p => ({ ...p, sheet: s.title })));
  const ng = pts.filter(p => res[p.id]?.result === "ng");
  const todo = d.sheets.map(s => [s.title, s.points.filter(p => !res[p.id]).length]).filter(([, n]) => n);
  const done = pts.filter(p => res[p.id]).length;
  const rv = d.review;
  openModal(`${fmtDate(date)} ${timing}　${plaState.process}`, `
    <p style="margin-top:0">点検 <b>${done}</b> / ${pts.length} 箇所${ng.length ? `　<span class="lv over">× 異常 ${ng.length}</span>` : ""}</p>
    ${ng.length ? `<h3 class="pla-h">異常</h3><ul class="pla-ng-list">${ng.map(p => `<li><b>${h(p.sheet)}</b> ／ ${h([p.location, p.item_text.replace(/\n/g, " ")].filter(Boolean).join(" "))}<div class="pla-ng-note">× ${h(res[p.id].note || "")}（${h(res[p.id].checked_by || "")}）</div></li>`).join("")}</ul>` : ""}
    ${todo.length && done ? `<h3 class="pla-h">終わっていないチェック表（${todo.length}）</h3><p class="pla-todo-list">${todo.map(([t, n]) => `${h(t)} <span class="muted">残り${n}</span>`).join("　/　")}</p>` : ""}
    ${!done ? `<p class="empty">まだ点検されていません</p>` : ""}
    <h3 class="pla-h">確認</h3>
    ${rv ? `<p class="pla-ok">✓ ${h(rv.reviewer)} さんが確認しました（${h(rv.reviewed_at.replace("T", " ").slice(0, 16))}）${rv.comment ? `<br>コメント: ${h(rv.comment)}` : ""}</p>` : ""}
    <label class="field full"><span>コメント（任意）</span><input name="comment" value="${h(rv?.comment || "")}" placeholder="例: 異常箇所は交換済みを確認"></label>
    <p class="muted" style="font-size:13px">確認者は右上の「記入者」（${h(currentUser() || "未入力")}）で記録されます。</p>
    <p class="err" id="form-err" hidden></p>`, `
    ${rv ? `<button type="button" class="btn left" id="rv-cancel">確認を取り消す</button>` : ""}
    <a class="btn" href="#/pla" id="rv-open">点検画面を開く</a>
    ${done ? `<button type="submit" class="btn primary">${rv ? "確認し直す" : "確認しました"}</button>` : ""}`);
  $("#rv-open").addEventListener("click", () => { plaState.date = date; plaState.timing = timing; closeModal(); });
  const send = async cancel => {
    try {
      await api("/pla/reviews", { method: "PUT", body: { process: plaState.process, date, timing, reviewer: currentUser(), comment: $("#modal-form [name=comment]").value, cancel } });
      closeModal(); toast(cancel ? "確認を取り消しました" : "確認しました"); router();
    } catch (ex) { $("#form-err").textContent = ex.message; $("#form-err").hidden = false; }
  };
  $("#modal-form").onsubmit = e => { e.preventDefault(); send(false); };
  $("#rv-cancel")?.addEventListener("click", () => send(true));
}

/* ---------------------------------------------------------------- チェック表（行・写真・プラリストとの紐づけを直す） */
/* スマホの写真は大きいので、長い辺 1600px の JPEG にしてから送る */
async function shrinkPhoto(file) {
  if (!file.type.startsWith("image/") || file.type === "image/gif") return file;
  try {
    const img = await createImageBitmap(file);
    const scale = Math.min(1, 1600 / Math.max(img.width, img.height));
    const c = document.createElement("canvas");
    c.width = Math.round(img.width * scale); c.height = Math.round(img.height * scale);
    c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
    const blob = await new Promise(r => c.toBlob(r, "image/jpeg", 0.85));
    return blob ? new File([blob], "photo.jpg", { type: "image/jpeg" }) : file;
  } catch { return file; }
}
function pickPhotos(onFiles) {
  const inp = document.createElement("input");
  inp.type = "file"; inp.accept = "image/*"; inp.multiple = true;
  inp.onchange = () => onFiles([...inp.files]);
  inp.click();
}
async function uploadPhotos(files, sheetId, pointId) {
  for (const f of files) {
    const fd = new FormData();
    fd.append("file", await shrinkPhoto(f));
    fd.append("sheet_id", sheetId);
    if (pointId) fd.append("point_id", pointId);
    await api("/pla/photos", { method: "POST", body: fd });
  }
}

async function renderPlaSheets() {
  const [sheets, items] = await Promise.all([
    api(`/pla/sheets?process=${encodeURIComponent(plaState.process)}&all=1`),
    api(`/pla/items?process=${encodeURIComponent(plaState.process)}`),
  ]);
  const shown = sheets.filter(s => plaState.showOldSheets || s.active || s.id === plaState.editSheet);
  if (!plaState.editSheet || !sheets.some(s => s.id === plaState.editSheet)) plaState.editSheet = (shown[0] || {}).id;
  const sheet = sheets.find(s => s.id === plaState.editSheet);
  view.innerHTML = `
    ${plaHeader("sheets")}
    <div class="pla-layout">
      <aside class="pla-eqs card">
        <select id="ps-select" class="pla-eq-select">${shown.map(s => `<option value="${s.id}" ${s.id === plaState.editSheet ? "selected" : ""}>${h(s.title)}${s.active ? "" : "（使わない）"}</option>`).join("")}</select>
        <label class="muted ps-old"><input type="checkbox" id="ps-old" ${plaState.showOldSheets ? "checked" : ""}> 使わない表（旧版・撤去）も表示</label>
        <ul id="pla-eq-list">${shown.map(s => {
          const unl = s.points.filter(p => !p.items.length).length;
          return `<li class="${s.id === plaState.editSheet ? "on" : ""} ${s.active ? "" : "old"}" data-s="${s.id}"><span class="pla-eq-name">${h(s.title)}</span><span class="pla-eq-count">${s.points.length}行</span></li>`;
        }).join("")}</ul>
        <button class="btn small" id="ps-new" style="margin:8px">＋ チェック表を追加</button>
      </aside>
      <section class="pla-main" id="ps-main"></section>
    </div>`;
  bindPlaHeader();
  $("#ps-old").addEventListener("change", e => { plaState.showOldSheets = e.target.checked; router(); });
  $$("#pla-eq-list li").forEach(li => li.addEventListener("click", () => { plaState.editSheet = +li.dataset.s; router(); }));
  $("#ps-select").addEventListener("change", e => { plaState.editSheet = +e.target.value; router(); });
  $("#ps-new").addEventListener("click", async () => {
    const title = prompt("新しいチェック表の名前");
    if (!title) return;
    const r = await api("/pla/sheets", { method: "POST", body: { process: plaState.process, title } });
    plaState.editSheet = r.id; router();
  });
  if (!sheet) { $("#ps-main").innerHTML = `<p class="empty">チェック表がありません</p>`; return; }

  $("#ps-main").innerHTML = `
    <div class="card ps-head">
      <div class="form-grid">
        <label class="field full"><span>チェック表の名前</span><input id="ps-title" value="${h(sheet.title)}"></label>
        <label class="field"><span>状態</span><select id="ps-active"><option value="1" ${sheet.active ? "selected" : ""}>使う（毎日点検する）</option><option value="0" ${sheet.active ? "" : "selected"}>使わない（旧版・撤去）</option></select></label>
        <div class="field"><span>元のファイル</span><div class="muted" style="font-size:13px;padding-top:6px">${h(sheet.file || "アプリで作成")}${sheet.sheet_name ? "／" + h(sheet.sheet_name) : ""}</div></div>
        <label class="field full"><span>改定履歴・メモ</span><textarea id="ps-note" rows="3">${h(sheet.note || "")}</textarea></label>
      </div>
      <div style="text-align:right"><button class="btn primary" id="ps-save">保存</button></div>
    </div>
    <div class="section-title"><h2>場所の写真</h2><button class="btn small" id="ps-add-place">📷 写真を追加</button></div>
    <div class="ps-photos">${sheet.photos.map((ph, i) => `<div class="ps-ph"><img src="${photoUrl(ph.file)}" data-i="${i}" alt=""><button class="ps-del" data-ph="${ph.id}" title="この写真を外す">✕</button></div>`).join("") || '<p class="muted">なし（設備全体や場所が分かる写真があると、初めての人でも探しやすくなります）</p>'}</div>
    <div class="section-title"><h2>点検する行（${sheet.points.length}）</h2></div>
    <ul class="ps-points">${sheet.points.map((p, n) => `
      <li class="ps-point ${p.active ? "" : "old"}" data-p="${p.id}">
        <div class="ps-point-photos">${p.photos.map((ph, i) => `<div class="ps-ph small"><img src="${photoUrl(ph.file)}" data-i="${i}" alt=""><button class="ps-del" data-ph="${ph.id}" title="この写真を外す">✕</button></div>`).join("")}
          <button class="ps-add-ph" title="写真を追加">📷<br>追加</button></div>
        <div class="ps-point-body">
          <div class="form-grid">
            <label class="field"><span>場所</span><input class="ps-loc" value="${h(p.location || "")}" placeholder="例: 操作BOX"></label>
            <label class="field"><span>項目<em class="req">*</em></span><textarea class="ps-text" rows="${Math.min(6, (p.item_text.match(/\n/g) || []).length + 1)}">${h(p.item_text)}</textarea></label>
          </div>
          <div class="ps-links">${p.items.map(i => `<span class="ps-link">${hazardBadge(i.hazard)}${h(i.item)}</span>`).join("") || '<span class="lv soon">プラリストと未紐づけ</span>'}
            <button class="btn small ps-link-btn">紐づけを変更</button></div>
          <div class="ps-actions">
            <button class="btn small ps-up" ${n === 0 ? "disabled" : ""}>↑</button><button class="btn small ps-down" ${n === sheet.points.length - 1 ? "disabled" : ""}>↓</button>
            <label class="muted"><input type="checkbox" class="ps-act" ${p.active ? "checked" : ""}> 点検する</label>
            <button class="btn small danger ps-del-point">行を削除</button>
          </div>
        </div>
      </li>`).join("")}</ul>
    <div class="card ps-add">
      <h3>行を追加</h3>
      <div class="form-grid">
        <label class="field"><span>場所</span><input id="ps-new-loc"></label>
        <label class="field"><span>項目<em class="req">*</em></span><textarea id="ps-new-text" rows="2" placeholder="例: 非常停止ボタン(赤)×1"></textarea></label>
      </div>
      <div style="text-align:right"><button class="btn primary" id="ps-new-point">追加</button></div>
    </div>`;

  const reload = () => router();
  const put = async (url, body, msg) => { try { await api(url, { method: "PUT", body }); if (msg) toast(msg); return true; } catch (ex) { toast(ex.message, 4000); return false; } };
  $("#ps-save").addEventListener("click", async () => {
    if (await put(`/pla/sheets/${sheet.id}`, { title: $("#ps-title").value, active: $("#ps-active").value, note: $("#ps-note").value }, "保存しました")) reload();
  });
  $("#ps-add-place").addEventListener("click", () => pickPhotos(async fs => { await uploadPhotos(fs, sheet.id, null); toast("写真を追加しました"); reload(); }));
  $$(".ps-photos img").forEach(img => img.addEventListener("click", () => openLightbox(sheet.photos, +img.dataset.i, sheet.title)));
  $$("#ps-main .ps-del").forEach(b => b.addEventListener("click", async () => {
    if (!confirm("この写真を外しますか？")) return;
    await api(`/pla/photos/${b.dataset.ph}`, { method: "DELETE" }); reload();
  }));
  $$(".ps-point").forEach(li => {
    const p = sheet.points.find(x => x.id == li.dataset.p);
    li.querySelectorAll(".ps-point-photos img").forEach(img => img.addEventListener("click", () => openLightbox(p.photos, +img.dataset.i, p.item_text)));
    li.querySelector(".ps-add-ph").addEventListener("click", () => pickPhotos(async fs => { await uploadPhotos(fs, sheet.id, p.id); toast("写真を追加しました"); reload(); }));
    li.querySelector(".ps-loc").addEventListener("change", e => put(`/pla/points/${p.id}`, { location: e.target.value }, "保存しました"));
    li.querySelector(".ps-text").addEventListener("change", e => put(`/pla/points/${p.id}`, { item_text: e.target.value }, "保存しました"));
    li.querySelector(".ps-act").addEventListener("change", async e => { if (await put(`/pla/points/${p.id}`, { active: e.target.checked ? 1 : 0 }, "保存しました")) li.classList.toggle("old", !e.target.checked); });
    li.querySelector(".ps-up").addEventListener("click", async () => { if (await put(`/pla/points/${p.id}`, { move: "up" })) reload(); });
    li.querySelector(".ps-down").addEventListener("click", async () => { if (await put(`/pla/points/${p.id}`, { move: "down" })) reload(); });
    li.querySelector(".ps-del-point").addEventListener("click", async () => {
      if (!confirm("この行を削除しますか？")) return;
      try { await api(`/pla/points/${p.id}`, { method: "DELETE" }); reload(); } catch (ex) { toast(ex.message, 4000); }
    });
    li.querySelector(".ps-link-btn").addEventListener("click", () => openPlaLink(p, items, reload));
  });
  $("#ps-new-point").addEventListener("click", async () => {
    const text = $("#ps-new-text").value.trim();
    if (!text) return toast("項目を入れてください");
    await api("/pla/points", { method: "POST", body: { sheet_id: sheet.id, location: $("#ps-new-loc").value, item_text: text } });
    toast("行を追加しました"); reload();
  });
}

/* チェック表の 1 行に、プラリストのどの行が入っているかを選ぶ */
function openPlaLink(p, items, done) {
  const chosen = new Set(p.items.map(i => i.id));
  // 最初は、すでに紐づいている行の設備（なければ全部）を出す
  const eqs = [...new Set(p.items.map(i => i.equipment))];
  openModal("プラリストと紐づけ", `
    <p style="margin-top:0"><b>${h(p.location || "")}</b> ${h(p.item_text.replace(/\n/g, " "))}</p>
    <div class="search" style="margin-bottom:8px"><span>🔍</span><input id="lk-q" type="search" placeholder="設備名・箇所で探す" value="${h(eqs[0] || "")}"></div>
    <div id="lk-list" class="lk-list"></div>
    <p class="muted" style="font-size:13px" id="lk-count"></p>`, `
    <button type="button" class="btn" data-close>キャンセル</button><button type="submit" class="btn primary">保存</button>`);
  const draw = () => {
    const q = norm($("#lk-q").value).toLowerCase();
    const list = items.filter(i => i.active && (chosen.has(i.id) || !q || norm(i.equipment + " " + i.item).toLowerCase().includes(q))).slice(0, 300);
    $("#lk-list").innerHTML = groupBy(list, i => i.equipment).map(([g, l]) => `
      <div class="lk-group"><b>${h(g)}</b>${l.map(i => `<label><input type="checkbox" value="${i.id}" ${chosen.has(i.id) ? "checked" : ""}> ${hazardBadge(i.hazard)}${h(i.item)}</label>`).join("")}</div>`).join("") || `<p class="empty">見つかりません</p>`;
    $("#lk-count").textContent = `選んでいる行: ${chosen.size}`;
    $$("#lk-list input").forEach(c => c.addEventListener("change", () => { c.checked ? chosen.add(+c.value) : chosen.delete(+c.value); $("#lk-count").textContent = `選んでいる行: ${chosen.size}`; }));
  };
  let timer;
  $("#lk-q").addEventListener("input", () => { clearTimeout(timer); timer = setTimeout(draw, 200); });
  $("#modal-form").onsubmit = async e => {
    e.preventDefault();
    await api(`/pla/points/${p.id}`, { method: "PUT", body: { item_ids: [...chosen] } });
    closeModal(); toast("紐づけを保存しました"); done();
  };
  draw();
}

/* ---------------------------------------------------------------- プラリスト（危害度分析の一覧） */
async function renderPlaItems() {
  const items = await api(`/pla/items?process=${encodeURIComponent(plaState.process)}`);
  view.innerHTML = `
    ${plaHeader("items")}
    <div class="toolbar">
      <div class="search"><span>🔍</span><input id="pi-q" type="search" placeholder="設備名・箇所・色・材質・備考で絞り込み" value="${h(plaState.q)}"></div>
      <button class="btn primary lg" id="pi-new">行を追加</button>
    </div>
    <div class="filters">
      <label><input type="checkbox" id="pi-unlinked" ${plaState.onlyUnlinked ? "checked" : ""}> チェック表にない行だけ（日々の点検なのに、どのチェック表にも入っていない）</label>
      <label><input type="checkbox" id="pi-inactive" ${plaState.showInactive ? "checked" : ""}> 点検しない行（撤去・項目から外した）も表示</label>
      <span class="muted" id="pi-count"></span>
    </div>
    <div id="pi-body"></div>`;
  const draw = () => {
    const q = norm(plaState.q).toLowerCase();
    const list = items.filter(i => (plaState.showInactive || i.active) && (!plaState.onlyUnlinked || (i.daily && !i.sheets.length)) &&
      (!q || norm([i.equipment, i.item, i.color, i.material, i.material_maker, i.note].join(" ")).toLowerCase().includes(q)));
    const unlinked = items.filter(i => i.daily && !i.sheets.length).length;
    $("#pi-count").textContent = `全 ${items.length} 行（日々の点検 ${items.filter(i => i.daily).length}・うちチェック表にない ${unlinked}）／ 表示 ${list.length}`;
    const groups = groupBy(list, i => i.equipment);
    $("#pi-body").innerHTML = groups.length ? groups.map(([g, l]) => `
      <section class="cat-group">
        <h2 class="area-title">${h(g)} <span class="g-count">${l.length}行</span><button class="btn small pi-add" data-g="${h(g)}" style="margin-left:auto">＋ 行</button></h2>
        <div class="table-wrap"><table class="list responsive"><thead><tr>
          <th>危害度</th><th>機械・器具</th><th>色</th><th>材質</th><th>頻度</th><th>チェック表</th><th>備考</th>
        </tr></thead><tbody>${l.map(i => `
          <tr data-pi="${i.id}" class="${i.active ? "" : "retired"}">
            <td data-label="危害度">${hazardBadge(i.hazard) || '<span class="muted">—</span>'}</td>
            <td data-label="機械・器具"><strong>${h(i.item)}</strong>${i.active ? "" : ' <span class="lv unknown">点検しない</span>'}</td>
            <td data-label="色">${h(i.color || "")}</td>
            <td data-label="材質">${h([i.material, i.material_pct, i.material_maker].filter(Boolean).join(" ／ "))}</td>
            <td class="nowrap" data-label="頻度">${h(i.frequency || "")}</td>
            <td data-label="チェック表">${i.sheets.length ? [...new Map(i.sheets.map(s => [s.sheet_id, s])).values()].map(s => `<a href="#/pla/sheets" class="pi-sheet" data-s="${s.sheet_id}">${h(s.title)}</a>`).join("<br>") : i.daily ? '<span class="lv soon">なし</span>' : ""}</td>
            <td class="clip" data-label="備考">${h(i.note || "")}</td>
          </tr>`).join("")}</tbody></table></div>
      </section>`).join("") : `<p class="empty">該当する行はありません</p>`;
    $$("#pi-body [data-pi]").forEach(tr => tr.addEventListener("click", e => {
      if (e.target.closest(".pi-sheet")) return;
      openPlaItemForm(items.find(i => i.id == tr.dataset.pi), items);
    }));
    $$("#pi-body .pi-sheet").forEach(a => a.addEventListener("click", () => { plaState.editSheet = +a.dataset.s; plaState.showOldSheets = false; }));
    $$("#pi-body .pi-add").forEach(b => b.addEventListener("click", e => { e.stopPropagation(); openPlaItemForm(null, items, { equipment: b.dataset.g }); }));
  };
  let timer;
  $("#pi-q").addEventListener("input", e => { clearTimeout(timer); timer = setTimeout(() => { plaState.q = e.target.value; draw(); }, 200); });
  $("#pi-inactive").addEventListener("change", e => { plaState.showInactive = e.target.checked; draw(); });
  $("#pi-unlinked").addEventListener("change", e => { plaState.onlyUnlinked = e.target.checked; draw(); });
  $("#pi-new").addEventListener("click", () => openPlaItemForm(null, items));
  bindPlaHeader();
  draw();
}

function openPlaItemForm(it, items, preset = {}) {
  const v = it || { process: plaState.process, active: 1, frequency: "2回/日", methods: "プラリスク点検表", hazard: "C", ...preset };
  const methods = (v.methods || "").split(",").filter(Boolean);
  const eqs = [...new Set(items.map(i => i.equipment))];
  openModal(it ? "プラリストの行を編集" : "プラリストに行を追加", `
    <div class="form-grid">
      <label class="field full"><span>設備名<em class="req">*</em></span><input name="equipment" list="pla-eqs" value="${h(v.equipment || "")}"><datalist id="pla-eqs">${eqs.map(e => `<option value="${h(e)}">`).join("")}</datalist></label>
      <label class="field full"><span>機械・器具（点検箇所）<em class="req">*</em></span><input name="item" value="${h(v.item || "")}" placeholder="例: 非常停止ボタン(赤)×1"></label>
      <label class="field"><span>色</span><input name="color" value="${h(v.color || "")}"></label>
      <div class="field"><span>危害度</span><div class="seg">${["", ...plaState.meta.hazards].map(x => `<label><input type="radio" name="hazard" value="${x}" ${(v.hazard || "") === x ? "checked" : ""}>${x || "なし"}</label>`).join("")}</div></div>
      <label class="field"><span>材質（ぷらしる）</span><input name="material" value="${h(v.material || "")}" placeholder="例: PA"></label>
      <label class="field"><span>％（ぷらしる）</span><input name="material_pct" value="${h(v.material_pct || "")}"></label>
      <label class="field full"><span>材質（メーカー提示）</span><input name="material_maker" value="${h(v.material_maker || "")}"></label>
      <div class="field full"><span>点検方法</span><div class="pla-methods">${plaState.meta.methods.map(mt => `<label><input type="checkbox" name="m" value="${h(mt)}" ${methods.includes(mt) ? "checked" : ""}> ${h(mt)}</label>`).join("")}</div></div>
      <label class="field"><span>頻度</span><input name="frequency" list="pla-freq" value="${h(v.frequency || "")}"><datalist id="pla-freq"><option value="2回/日"><option value="配合毎"><option value="1時間毎"></datalist></label>
      <label class="field"><span>状態</span><select name="active"><option value="1" ${v.active ? "selected" : ""}>点検する</option><option value="0" ${v.active ? "" : "selected"}>点検しない（撤去・項目から外した）</option></select></label>
      <label class="field full"><span>備考（更新記録など）</span><textarea name="note" rows="2">${h(v.note || "")}</textarea></label>
      <label class="field full"><input type="checkbox" name="is_new" ${v.is_new ? "checked" : ""}> 新規</label>
      ${it?.sheets?.length ? `<div class="field full"><span>入っているチェック表</span><div style="font-size:14px">${[...new Set(it.sheets.map(s => s.title))].map(h).join("、")}</div></div>` : ""}
    </div><p class="err" id="form-err" hidden></p>`, `
    ${it ? `<button type="button" class="btn danger left" id="pi-del">削除</button>` : ""}
    <button type="button" class="btn" data-close>キャンセル</button><button type="submit" class="btn primary">保存</button>`);
  const form = $("#modal-form");
  form.onsubmit = async e => {
    e.preventDefault();
    const fd = new FormData(form);
    const data = Object.fromEntries([...fd].filter(([k]) => k !== "m"));
    data.methods = fd.getAll("m").join(",");
    data.is_new = fd.get("is_new") ? 1 : 0;
    data.process = v.process;
    if (!data.equipment.trim() || !data.item.trim()) { $("#form-err").textContent = "設備名と機械・器具は必須です"; $("#form-err").hidden = false; return; }
    try {
      if (it) await api(`/pla/items/${it.id}`, { method: "PUT", body: data });
      else await api("/pla/items", { method: "POST", body: data });
      closeModal(); toast("保存しました"); router();
    } catch (ex) { $("#form-err").textContent = ex.message; $("#form-err").hidden = false; }
  };
  $("#pi-del")?.addEventListener("click", async () => {
    if (!confirm(`「${it.item}」をプラリストから削除しますか？（チェック表との紐づけも外れます）`)) return;
    try { await api(`/pla/items/${it.id}`, { method: "DELETE" }); closeModal(); toast("削除しました"); router(); }
    catch (ex) { $("#form-err").textContent = ex.message; $("#form-err").hidden = false; }
  });
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
