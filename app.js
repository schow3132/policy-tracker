/* Policy Tracker: renders data/tracker.json (generated from the spreadsheet by update-data.ps1).
   Sheets are found by name and tables by their header labels, so the workbook layout can grow
   (more rows, more candidates, more policies) without code changes. */
(function () {
  "use strict";

  const $ = (s, el = document) => el.querySelector(s);
  const esc = (v) => String(v == null ? "" : v).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const clean = (v) => (v == null ? "" : String(v).replace(/\s+/g, " ").trim());
  const num = (v) => { const n = Number(v); return v !== null && v !== "" && isFinite(n) ? n : null; };
  const pct = (v) => (num(v) == null ? "–" : Math.round(num(v) * 100) + "%");
  const money = (v) => {
    const n = num(v); if (n == null) return "–";
    const a = Math.abs(n), s = n < 0 ? "−$" : "$";
    if (a >= 1e9) return s + (a / 1e9).toFixed(2) + "B";
    if (a >= 1e6) return s + (a / 1e6).toFixed(a >= 1e8 ? 0 : 1) + "M";
    if (a >= 1e3) return s + Math.round(a / 1e3) + "K";
    return s + Math.round(a);
  };
  const safeUrl = (u) => (/^https?:\/\//i.test(clean(u)) ? clean(u) : "");
  const ST = { AL: "", Michigan: "MI", Texas: "TX", Ohio: "OH", Wisconsin: "WI", Kansas: "KS" };
  const abbr = (s) => ST[s] || s;
  const GLYPH = { Support: "✔", Middle: "◐", Oppose: "✖" };
  const DETAIL_ORDER = ["Detailed plan", "Specific proposal", "Claim only"];

  let BOOK = null;

  /* ---------- sheet helpers ---------- */
  function rows(sheet) {
    const s = BOOK.sheets[sheet];
    if (!s) return [];
    return (s.rows || []).map((r) => (Array.isArray(r) ? r : r == null ? [] : [r]));
  }
  const isEmpty = (r) => !r || !r.some((c) => clean(c) !== "");
  function subtitle(sheet) {
    const r = rows(sheet)[1] || [];
    return clean(r.find((c) => clean(c)) || "");
  }
  /** Table whose header row contains `first`; reads until the first blank row. */
  function table(sheet, first, after) {
    const R = rows(sheet);
    let start = 0;
    if (after) start = Math.max(0, R.findIndex((r) => r.some((c) => clean(c).startsWith(after))));
    let hi = -1, col = 0;
    for (let i = start; i < R.length && hi < 0; i++) {
      const j = R[i].findIndex((c) => clean(c) === first);
      if (j >= 0) { hi = i; col = j; }
    }
    if (hi < 0) return { rows: [], after: [] };
    const head = R[hi].slice(col).map(clean);
    const out = [];
    let i = hi + 1;
    for (; i < R.length; i++) {
      const r = R[i];
      if (isEmpty(r) || clean(r[col]) === "") break;
      // a lone text cell under a multi-column table is a footnote, not a data row
      if (head.length > 2 && r.filter((c) => clean(c) !== "").length === 1) break;
      const o = {};
      head.forEach((h, k) => { if (h) o[h] = r[col + k] == null ? "" : r[col + k]; });
      out.push(o);
    }
    const tail = R.slice(i).filter((r) => !isEmpty(r)).map((r) => r.map(clean).filter(Boolean));
    return { rows: out, after: tail };
  }
  function findText(sheet, prefix) {
    for (const r of rows(sheet)) for (const c of r) if (clean(c).startsWith(prefix)) return clean(c);
    return "";
  }
  /** Value in the row below the cell labelled `label` (dashboard tiles). */
  function below(sheet, label) {
    const R = rows(sheet);
    for (let i = 0; i < R.length; i++) {
      const j = R[i].findIndex((c) => clean(c) === label);
      if (j >= 0) return (R[i + 1] || [])[j];
    }
    return null;
  }

  /* ---------- tooltip ---------- */
  const tips = [];
  const tip = (html) => { tips.push(html); return tips.length - 1; };
  const tipEl = () => $("#tip");
  const noHover = window.matchMedia("(hover: none)");
  let pinned = null;
  function showTip(target, x, y) {
    const t = tipEl();
    t.innerHTML = tips[+target.dataset.tip];
    t.hidden = false;
    const r = target.getBoundingClientRect();
    const w = t.offsetWidth, h = t.offsetHeight;
    let left = x != null ? x + 14 : r.right + 8;
    let top = y != null ? y + 14 : r.top;
    if (left + w > innerWidth - 8) left = Math.max(8, (x != null ? x : r.left) - w - 14);
    if (top + h > innerHeight - 8) top = Math.max(8, (y != null ? y : r.bottom) - h - 14);
    t.style.left = left + "px"; t.style.top = top + "px";
  }
  function hideTip() { pinned = null; const t = tipEl(); t.hidden = true; t.style.pointerEvents = "none"; }
  document.addEventListener("mousemove", (e) => {
    if (pinned) return;
    const el = e.target.closest && e.target.closest("[data-tip]");
    if (el) showTip(el, e.clientX, e.clientY); else hideTip();
  });
  document.addEventListener("focusin", (e) => {
    const el = e.target.closest && e.target.closest("[data-tip]");
    if (el) showTip(el); else if (!pinned) hideTip();
  });
  document.addEventListener("click", (e) => {
    const el = e.target.closest && e.target.closest("[data-tip]");
    if (e.target.closest && e.target.closest("#tip")) return;
    if (el && noHover.matches) {
      // touch: first tap shows the details (with the source link inside), not the page behind it
      e.preventDefault();
      pinned = el; showTip(el); tipEl().style.pointerEvents = "auto";
      return;
    }
    if (pinned) hideTip();
  });
  window.addEventListener("scroll", () => { if (!pinned) hideTip(); }, { passive: true });

  /* ---------- shared bits ---------- */
  const party = (p) => `<span class="pty pty--${esc(p)}" aria-label="${p === "D" ? "Democrat" : p === "R" ? "Republican" : esc(p)}">${esc(p)}</span>`;
  const glyph = (st) => GLYPH[st] ? `<span class="glyph glyph--${st}" aria-hidden="true">${GLYPH[st]}</span>` : `<span class="glyph glyph--silent" aria-hidden="true"></span>`;
  const isPrimary = (c) => clean(c["Source class"]).toLowerCase().startsWith("primary");
  const srcLabel = (c) => {
    const s = clean(c["Source class"]);
    const [kind, rest] = s.split(/\s+[—-]\s+/);
    return `<span class="src ${isPrimary(c) ? "src--primary" : ""}">${esc(kind)}</span>${rest ? " · " + esc(rest) : ""}`;
  };
  function claimCard(c, opts = {}) {
    const url = safeUrl(c["Source URL"]);
    return `<article class="claim">
      <div class="claim__top">
        ${opts.noWho ? "" : `<span class="claim__who">${party(c.Party)}${esc(c.Candidate)}</span><span class="claim__office">${esc(abbr(c.State))} ${esc(c.Office)}</span>`}
        <span class="st">${glyph(c.Stance)}${esc(c.Stance)}</span>
        <span class="chip">${esc(c["Plan detail"])}</span>
      </div>
      ${opts.noPolicy ? "" : `<div class="claim__policy">${esc(c["Issue area"])} · ${esc(c.Policy)}</div>`}
      <p class="claim__text">${esc(c["What they said (paraphrased)"])}</p>
      <div class="claim__foot">
        <span class="claim__id">${esc(c["Claim #"])}</span>
        <span>${srcLabel(c)}</span>
        ${url ? `<a href="${esc(url)}" target="_blank" rel="noopener">Open source ↗</a>` : ""}
        ${c.Accessed ? `<span>Accessed ${esc(c.Accessed)}</span>` : ""}
      </div>
    </article>`;
  }
  function select(id, label, values, all) {
    return `<label class="visually-hidden" for="${id}">${esc(label)}</label>
      <select id="${id}" aria-label="${esc(label)}"><option value="">${esc(all)}</option>${values.map((v) => `<option>${esc(v)}</option>`).join("")}</select>`;
  }
  const uniq = (arr) => [...new Set(arr.filter((v) => clean(v) !== ""))];

  /* ---------- data model ---------- */
  let M = {};
  function model() {
    const claims = table("Claims Log", "Claim #").rows;
    const policies = table("Policy Index", "ID");
    const score = table("Candidate Scorecard", "Candidate");
    const statewide = table("Statewide Races", "State");
    const detail = table("Policy Detail", "Candidate").rows;
    const house = table("House Races", "State");
    const donors = table("Donor Profile", "Candidate");
    const houseByState = table("Dashboard", "State", "U.S. HOUSE BY STATE");
    const byKey = new Map(claims.map((c) => [clean(c.Candidate) + "|" + clean(c.Policy), c]));
    const cands = score.rows.map((s) => {
      const office = clean(s.Office);
      const short = abbr(clean(s.State)) + " " + (office.startsWith("Senate") ? "Sen" : office.startsWith("Gov") ? "Gov" : office);
      return { ...s, name: clean(s.Candidate), party: clean(s.Party), short };
    });
    M = { claims, policies, score, statewide, detail, house, donors, houseByState, byKey, cands };
  }

  /* ---------- views ---------- */
  function renderOverview() {
    const asOf = (subtitle("Dashboard").match(/Data as of[^|]*/) || [""])[0].trim();
    $("#hero-eyebrow").textContent = clean(subtitle("Dashboard").split("|")[0]) || "2026 midterms";
    const prim = M.claims.filter(isPrimary).length;
    const urls = uniq(M.claims.map((c) => safeUrl(c["Source URL"]))).length;
    const k = [
      [M.claims.length, "claims logged, each linked to a source"],
      [M.policies.rows.length, "contested policies indexed"],
      [M.cands.length, "statewide nominees coded"],
      [M.claims.length ? Math.round((prim / M.claims.length) * 100) + "%" : "–", "of claims from the candidate's own site or plan"],
      [urls, "distinct source pages"],
      [num(below("Dashboard", "NOMINEES TRACKED")) ?? "–", "nominees tracked, including House"],
    ];
    $("#kpis").innerHTML = k.map(([v, l]) => `<div class="kpi"><div class="kpi__v">${esc(v)}</div><div class="kpi__l">${esc(l)}</div></div>`).join("");

    // race cards
    const groups = [];
    M.statewide.rows.forEach((r) => {
      const key = clean(r.State) + "|" + clean(r.Office);
      let g = groups.find((x) => x.key === key);
      if (!g) groups.push((g = { key, state: clean(r.State), office: clean(r.Office), rows: [] }));
      g.rows.push(r);
    });
    const maxRec = Math.max(1, ...M.statewide.rows.map((r) => num(r["Total receipts ($)"]) || 0));
    $("#race-cards").innerHTML = groups.map((g) => {
      const res = g.rows.map((r) => clean(r["Result (media projections)"]));
      const close = res.some((x) => /too close/i.test(x));
      const chip = close ? `<span class="chip chip--close">Too close to call</span>` : res.some((x) => /winner/i.test(x)) ? `<span class="chip chip--win">Projected</span>` : `<span class="chip">Uncalled</span>`;
      return `<article class="race">
        <div class="race__head"><span class="race__title">${esc(g.state)} · ${esc(g.office)}</span>${chip}</div>
        ${g.rows.map((r) => {
          const rec = num(r["Total receipts ($)"]);
          const win = /winner/i.test(r["Result (media projections)"]);
          const t = tip(`<b>${esc(r.Candidate)}</b>${esc(r["Seat status"])} · ${esc(r["Result (media projections)"])}<div class="tip__m">${rec != null ? `Raised ${money(rec)} · spent ${money(r["Total disbursements ($)"])} (FEC)` : esc(r["Funding source"])}</div>`);
          return `<div class="race__row" data-tip="${t}">
            ${party(clean(r.Party))}
            <div><div class="race__name ${win ? "is-win" : ""}">${esc(r.Candidate)}</div>
              ${rec != null ? `<div class="bar" style="width:${(rec / maxRec) * 100}%;background:var(--${clean(r.Party) === "D" ? "dem" : "rep"})"></div>` : ""}</div>
            <div class="race__money">${rec != null ? money(rec) + " raised" : "State filer"}</div>
          </div>`;
        }).join("")}
      </article>`;
    }).join("");

    // consensus bars (single series)
    const counts = {};
    M.policies.rows.forEach((p) => { const c = clean(p.Consensus); if (c) counts[c] = (counts[c] || 0) + 1; });
    const cmax = Math.max(1, ...Object.values(counts));
    $("#consensus-chart").innerHTML = `<div class="hbars">${Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([c, n]) => {
      const names = M.policies.rows.filter((p) => clean(p.Consensus) === c).map((p) => esc(p.Policy)).join("<br>");
      return `<div class="hbar"><div class="hbar__label"><span class="t">${esc(c)}</span></div>
        <div class="hbar__track"><div class="hbar__seg" data-tip="${tip(`<b>${esc(c)}</b>${names}`)}" style="width:${(n / cmax) * 85}%;background:var(--ink-2);border-radius:0 4px 4px 0"></div><span class="hbar__val">${n}</span></div></div>`;
    }).join("")}</div>`;

    // house by state, stacked
    const hs = M.houseByState.rows;
    const hmax = Math.max(1, ...hs.map((r) => num(r.Seats) || 0));
    $("#house-note").textContent = M.houseByState.after.map((r) => r.join(" ")).join(" ") || "Seats by projected party.";
    $("#house-chart").innerHTML = `<div class="legend"><span><i class="dot" style="background:var(--dem)"></i>D projected</span><span><i class="dot" style="background:var(--rep)"></i>R projected</span><span><i class="dot" style="background:var(--axis)"></i>Uncalled or not yet pulled</span></div>
      <div class="hbars">${hs.map((r) => {
        const d = num(r["D projected"]) || 0, rp = num(r["R projected"]) || 0, u = num(r["Uncalled / not pulled"]) || 0, seats = num(r.Seats) || d + rp + u;
        const seg = (n, col, label) => n ? `<div class="hbar__seg" data-tip="${tip(`<b>${esc(r.State)}: ${n} ${label}</b>of ${seats} seats`)}" style="width:${(n / hmax) * 85}%;background:${col}"></div>` : "";
        return `<div class="hbar"><div class="hbar__label"><span class="t">${esc(r.State)}</span></div>
          <div class="hbar__track">${seg(d, "var(--dem)", "D projected")}${seg(rp, "var(--rep)", "R projected")}${seg(u, "var(--axis)", "uncalled or not pulled")}<span class="hbar__val">${seats} seats${d || rp ? ` · ${d} D, ${rp} R` : ""}</span></div></div>`;
      }).join("")}</div>`;

    $("#notice").innerHTML = `<div class="wrap">${asOf ? "<b>" + esc(asOf.replace(/\s*\(.*\)/, "")) + ".</b> " : ""}Election results shown are media projections, not certified. Positions are coded only from what candidates stated.</div>`;
  }

  function renderPolicies() {
    const P = M.policies;
    const D = M.cands.filter((c) => c.party === "D"), R = M.cands.filter((c) => c.party === "R");
    $("#policies-sub").textContent = `${P.rows.length} contested policies, and where each of the ${D.length} Democratic and ${R.length} Republican statewide nominees stands. Open a policy to read every claim behind it.`;
    $("#policy-method").textContent = findText("Policy Index", "Method:");
    $("#policy-legend").innerHTML = `<span><i class="unit unit--Support"></i>Support</span><span><i class="unit unit--Middle"></i>Middle (partial or conditional)</span><span><i class="unit unit--Oppose"></i>Oppose</span><span><i class="unit unit--silent"></i>No stated position</span><span>One dot per nominee</span>`;
    $("#policy-filters").innerHTML = `<input type="search" id="pf-q" placeholder="Search policies" aria-label="Search policies">
      ${select("pf-issue", "Issue area", uniq(P.rows.map((p) => clean(p["Issue area"]))), "All issue areas")}
      ${select("pf-cons", "Consensus", uniq(P.rows.map((p) => clean(p.Consensus))), "Any consensus")}`;
    const order = { Support: 0, Middle: 1, Oppose: 2 };
    const side = (list, pol, label) => {
      const u = list.map((c) => ({ c, cl: M.byKey.get(c.name + "|" + clean(pol)) }))
        .sort((a, b) => (a.cl ? order[a.cl.Stance] ?? 3 : 4) - (b.cl ? order[b.cl.Stance] ?? 3 : 4));
      return `<div class="units__side"><small>${label}</small><div class="units__row">${u.map(({ c, cl }) =>
        `<i class="unit unit--${cl ? esc(cl.Stance) : "silent"}" data-tip="${tip(`<b>${esc(c.name)}</b>${esc(c.short)} · ${cl ? esc(cl.Stance) + " · " + esc(cl["Plan detail"]) : "No stated position"}`)}"></i>`).join("")}</div></div>`;
    };
    const draw = () => {
      const q = $("#pf-q").value.toLowerCase(), is = $("#pf-issue").value, co = $("#pf-cons").value;
      const list = P.rows.filter((p) => (!is || clean(p["Issue area"]) === is) && (!co || clean(p.Consensus) === co) && (!q || (clean(p.Policy) + " " + clean(p["Issue area"])).toLowerCase().includes(q)));
      $("#policy-list").innerHTML = list.map((p) => {
        const cl = M.claims.filter((c) => clean(c.Policy) === clean(p.Policy));
        const prim = cl.filter(isPrimary).length;
        return `<details class="pol">
          <summary>
            <div><div class="pol__name">${esc(p.Policy)}</div>
              <div class="pol__meta"><span>${esc(p.ID)} · ${esc(p["Issue area"])}</span><span class="chip">${esc(p.Consensus)}</span><span>${cl.length} claim${cl.length === 1 ? "" : "s"}, ${prim} primary-sourced</span></div></div>
            <div class="units">${side(D, p.Policy, "Democrats")}${side(R, p.Policy, "Republicans")}</div>
          </summary>
          <div class="pol__body">${cl.length ? cl.map((c) => claimCard(c, { noPolicy: true })).join("") : `<p class="empty">No claims logged yet.</p>`}</div>
        </details>`;
      }).join("") || `<p class="empty">No policies match.</p>`;
    };
    ["pf-q", "pf-issue", "pf-cons"].forEach((id) => $("#" + id).addEventListener("input", draw));
    draw();
  }

  function renderGrid() {
    const C = M.cands;
    $("#grid-legend").innerHTML = `<span>${glyph("Support")}Support</span><span>${glyph("Middle")}Middle</span><span>${glyph("Oppose")}Oppose</span><span>Blank: no stated position found</span>`;
    const firstR = C.findIndex((c) => c.party === "R");
    const head = `<thead><tr><th class="corner">Policy</th>${C.map((c, i) => `<th class="colh ${i === firstR ? "party-gap" : ""}" scope="col"><div>${party(c.party)}${esc(c.name)} <small>${esc(c.short)}</small></div></th>`).join("")}</tr></thead>`;
    const body = M.policies.rows.map((p) => `<tr><th scope="row">${esc(p.Policy)}<small>${esc(p["Issue area"])}</small></th>${C.map((c, i) => {
      const cl = M.byKey.get(c.name + "|" + clean(p.Policy));
      const gap = i === firstR ? ' class="party-gap"' : "";
      if (!cl) return `<td${gap}></td>`;
      const url = safeUrl(cl["Source URL"]);
      const t = tip(`<b>${esc(c.name)}: ${esc(cl.Stance)}</b>${esc(cl["What they said (paraphrased)"])}<div class="tip__m">${esc(cl["Plan detail"])} · ${srcLabel(cl)}</div>${url ? `<div class="tip__m"><a href="${esc(url)}" target="_blank" rel="noopener">Open source ↗</a></div>` : ""}`);
      const inner = glyph(cl.Stance);
      const label = `${c.name}, ${p.Policy}: ${cl.Stance}. Open source`;
      return `<td${gap}>${url ? `<a class="cell" href="${esc(url)}" target="_blank" rel="noopener" data-tip="${t}" aria-label="${esc(label)}">${inner}</a>` : `<span class="cell" data-tip="${t}" tabindex="0">${inner}</span>`}</td>`;
    }).join("")}</tr>`).join("");
    $("#stance-grid").innerHTML = `<table class="sg">${head}<tbody>${body}</tbody></table>`;
  }

  let claimState = {};
  function renderClaims() {
    const C = M.claims;
    $("#claim-filters").innerHTML = `<input type="search" id="cf-q" placeholder="Search what they said" aria-label="Search claims">
      <span class="seg" role="group" aria-label="Party"><button type="button" data-p="" aria-pressed="true">All</button><button type="button" data-p="D" aria-pressed="false">Dem</button><button type="button" data-p="R" aria-pressed="false">Rep</button></span>
      ${select("cf-cand", "Candidate", uniq(M.cands.map((c) => c.name)), "All candidates")}
      ${select("cf-issue", "Issue area", uniq(C.map((c) => clean(c["Issue area"]))), "All issue areas")}
      ${select("cf-stance", "Stance", ["Support", "Middle", "Oppose"], "Any stance")}
      ${select("cf-detail", "Plan detail", DETAIL_ORDER, "Any detail level")}
      ${select("cf-src", "Source", ["Primary", "Secondary"], "Any source")}
      <button type="button" id="cf-reset">Reset</button>`;
    let partyF = "";
    const draw = () => {
      const q = $("#cf-q").value.toLowerCase();
      const f = { cand: $("#cf-cand").value, issue: $("#cf-issue").value, st: $("#cf-stance").value, det: $("#cf-detail").value, src: $("#cf-src").value };
      const list = C.filter((c) => (!partyF || clean(c.Party) === partyF) && (!f.cand || clean(c.Candidate) === f.cand) && (!f.issue || clean(c["Issue area"]) === f.issue) && (!f.st || clean(c.Stance) === f.st) && (!f.det || clean(c["Plan detail"]) === f.det) && (!f.src || clean(c["Source class"]).startsWith(f.src)) && (!q || Object.values(c).join(" ").toLowerCase().includes(q)));
      $("#claim-count").textContent = `Showing ${list.length} of ${C.length} claims`;
      $("#claim-list").innerHTML = list.map((c) => claimCard(c)).join("") || `<p class="empty">No claims match these filters.</p>`;
    };
    $("#claim-filters").addEventListener("input", draw);
    $("#claim-filters").addEventListener("click", (e) => {
      const b = e.target.closest("button"); if (!b) return;
      if (b.id === "cf-reset") { $("#claim-filters").querySelectorAll("select,input").forEach((x) => (x.value = "")); partyF = ""; }
      else partyF = b.dataset.p;
      $("#claim-filters").querySelectorAll(".seg button").forEach((x) => x.setAttribute("aria-pressed", String(x.dataset.p === partyF)));
      draw();
    });
    claimState.apply = (params) => {
      $("#claim-filters").querySelectorAll("select,input").forEach((x) => (x.value = ""));
      partyF = "";
      if (params.cand) $("#cf-cand").value = params.cand;
      if (params.issue) $("#cf-issue").value = params.issue;
      $("#claim-filters").querySelectorAll(".seg button").forEach((x) => x.setAttribute("aria-pressed", String(x.dataset.p === partyF)));
      draw();
    };
    draw();
  }

  function renderCandidates() {
    $("#cand-sub").textContent = subtitle("Candidate Scorecard");
    $("#cand-note").textContent = (M.score.after[0] || []).join(" ");
    $("#cand-legend").innerHTML = `<span><i class="dot" style="background:var(--d-plan)"></i>Detailed plans</span><span><i class="dot" style="background:var(--d-spec)"></i>Specific proposals</span><span><i class="dot" style="background:var(--d-claim)"></i>Claims only</span>`;
    $("#cand-list").innerHTML = M.cands.map((c) => {
      const n = num(c["Policies with a stated position"]) || 0;
      const parts = [["Detailed plans", "--d-plan"], ["Specific proposals", "--d-spec"], ["Claims only", "--d-claim"]];
      const mix = n ? parts.map(([k, v]) => { const x = num(c[k]) || 0; return x ? `<span data-tip="${tip(`<b>${esc(c.name)}</b>${x} ${k.toLowerCase()} of ${n} positions`)}" style="flex:${x};background:var(${v})"></span>` : ""; }).join("") : `<span style="flex:1;background:var(--grid)"></span>`;
      const pos = M.detail.filter((d) => clean(d.Candidate) === c.name);
      const nClaims = M.claims.filter((x) => clean(x.Candidate) === c.name).length;
      const result = clean(c.Result);
      const rchip = /winner/i.test(result) ? "chip--win" : /close/i.test(result) ? "chip--close" : "";
      const score = num(c["Detail score (0–100)"]);
      const main = safeUrl(c["Main campaign source"]);
      return `<article class="cand">
        <div class="cand__head">
          <div><div class="cand__name">${party(c.party)}${esc(c.name)}</div>
            <div class="cand__office">${esc(c.State)} · ${esc(c.Office)}</div>
            <div style="margin-top:6px"><span class="chip ${rchip}">${esc(result)}</span></div></div>
          <div class="cand__score"><b>${score == null ? "–" : score}</b><small>detail score</small></div>
        </div>
        <div><div class="mix" role="img" aria-label="${esc(`${c["Detailed plans"] || 0} detailed plans, ${c["Specific proposals"] || 0} specific proposals, ${c["Claims only"] || 0} claims only`)}">${mix}</div></div>
        <div class="cand__stats">
          <span><b>${n}</b> indexed positions</span>
          <span>${glyph("Support")} ${esc(c.Supports || 0)}</span><span>${glyph("Middle")} ${esc(c.Middle || 0)}</span><span>${glyph("Oppose")} ${esc(c.Opposes || 0)}</span>
          <span>${pct(c["% of claims from primary sources"])} primary-sourced</span>
        </div>
        <details><summary>Positions by issue (${pos.length})</summary>
          ${pos.length ? `<dl>${pos.map((d) => { const u = safeUrl(d["Source URL"]); return `<div><dt>${esc(d.Issue)}</dt><dd>${esc(d["Position (paraphrased)"])} ${u ? `<a href="${esc(u)}" target="_blank" rel="noopener">${esc(d["Source type"] || "source")} ↗</a>` : ""}</dd></div>`; }).join("")}</dl>` : `<p class="empty">No positions found on the sourced pages yet.</p>`}
        </details>
        <div class="cand__stats">${nClaims ? `<a href="#claims?cand=${encodeURIComponent(c.name)}">See ${nClaims} coded claim${nClaims === 1 ? "" : "s"} →</a>` : `<span class="empty">No coded claims</span>`}${main ? `<a href="${esc(main)}" target="_blank" rel="noopener">Main source ↗</a>` : ""}</div>
      </article>`;
    }).join("");
  }

  function renderMoney() {
    const fed = M.statewide.rows.filter((r) => num(r["Total receipts ($)"]) != null);
    const max = Math.max(1, ...fed.map((r) => num(r["Total receipts ($)"])));
    $("#senate-money").innerHTML = `<div class="legend"><span><i class="dot" style="background:var(--dem)"></i>Democrat</span><span><i class="dot" style="background:var(--rep)"></i>Republican</span><span>Total receipts reported to the FEC</span></div>
      <div class="hbars">${fed.map((r) => {
        const v = num(r["Total receipts ($)"]);
        const t = tip(`<b>${esc(r.Candidate)} (${esc(r.Party)}), ${esc(abbr(r.State))} ${esc(r.Office)}</b>Raised ${money(v)} · spent ${money(r["Total disbursements ($)"])}<div class="tip__m">Cash difference ${money(r["Receipts − disbursements ($)"])} · ${esc(r["Result (media projections)"])}</div>`);
        return `<div class="hbar"><div class="hbar__label">${party(clean(r.Party))}<span class="t">${esc(r.Candidate)} <span style="color:var(--muted)">${esc(abbr(r.State))}</span></span></div>
          <div class="hbar__track"><div class="hbar__seg" data-tip="${t}" style="width:${(v / max) * 80}%;background:var(--${clean(r.Party) === "D" ? "dem" : "rep"})"></div><span class="hbar__val">${money(v)}</span></div></div>`;
      }).join("")}</div>`;
    $("#senate-note").textContent = (M.statewide.after[0] || []).join(" ");

    const B = [["≤$200 ($)", "Under $200", "--b1"], ["$200–499 ($)", "$200–499", "--b2"], ["$500–999 ($)", "$500–999", "--b3"], ["$1,000–1,999 ($)", "$1,000–1,999", "--b4"], ["$2,000+ ($)", "$2,000+", "--b5"]];
    $("#donor-legend").innerHTML = B.map(([, l, v]) => `<span><i class="dot" style="background:var(${v})"></i>${l}</span>`).join("");
    $("#donor-note").textContent = subtitle("Donor Profile");
    $("#donor-filters").innerHTML = `<span class="seg" role="group" aria-label="Race type"><button type="button" data-r="Senate" aria-pressed="true">Senate</button><button type="button" data-r="House" aria-pressed="false">House</button></span>`;
    let race = "Senate";
    const draw = () => {
      const list = M.donors.rows.filter((d) => (race === "Senate") === /senate/i.test(d.Race));
      $("#donor-chart").innerHTML = `<div class="hbars">${list.map((d) => {
        const tot = num(d["Total individual ($)"]) || B.reduce((s, [k]) => s + (num(d[k]) || 0), 0);
        const segs = B.map(([k, l, v]) => { const x = num(d[k]) || 0; if (!x || !tot) return ""; return `<div class="hbar__seg" data-tip="${tip(`<b>${esc(d.Candidate)}: ${l}</b>${money(x)} · ${Math.round((x / tot) * 100)}% of individual giving`)}" style="width:${(x / tot) * 78}%;background:var(${v})"></div>`; }).join("");
        return `<div class="hbar"><div class="hbar__label">${party(clean(d.Party))}<span class="t">${esc(d.Candidate)} <span style="color:var(--muted)">${esc(d.Race)}</span></span></div>
          <div class="hbar__track">${segs}<span class="hbar__val">${money(tot)} · ${pct(d["Small-dollar share"])} small</span></div></div>`;
      }).join("") || `<p class="empty">No donor data yet.</p>`}</div>`;
    };
    $("#donor-filters").addEventListener("click", (e) => {
      const b = e.target.closest("button"); if (!b) return;
      race = b.dataset.r;
      $("#donor-filters").querySelectorAll("button").forEach((x) => x.setAttribute("aria-pressed", String(x.dataset.r === race)));
      draw();
    });
    draw();
  }

  function renderHouse() {
    const H = M.house.rows;
    $("#house-sub").textContent = subtitle("House Races");
    $("#house-fine").textContent = (M.house.after[0] || []).join(" ");
    $("#house-filters").innerHTML = `<input type="search" id="hf-q" placeholder="Search name or district" aria-label="Search House nominees">
      ${select("hf-state", "State", uniq(H.map((r) => clean(r.State))), "All states")}
      ${select("hf-res", "Result", uniq(H.map((r) => clean(r["Result (media projections)"]))), "Any result")}`;
    const draw = () => {
      const q = $("#hf-q").value.toLowerCase(), s = $("#hf-state").value, res = $("#hf-res").value;
      const list = H.filter((r) => (!s || clean(r.State) === s) && (!res || clean(r["Result (media projections)"]) === res) && (!q || (clean(r.Candidate) + " " + clean(r.District)).toLowerCase().includes(q)));
      $("#house-count").textContent = `Showing ${list.length} of ${H.length} nominees`;
      let prev = "";
      $("#house-table").innerHTML = `<thead><tr><th>District</th><th>Candidate</th><th>Seat</th><th>Result</th><th class="num">Raised</th><th class="num">Spent</th><th>FEC</th></tr></thead><tbody>${list.map((r) => {
        const start = clean(r.District) !== prev; prev = clean(r.District);
        const u = safeUrl(r["FEC link"]);
        const win = /winner/i.test(r["Result (media projections)"]);
        return `<tr class="${start ? "district-start" : ""}"><td>${start ? esc(r.District) : ""}</td><td><span class="who">${party(clean(r.Party))}${esc(r.Candidate)}</span></td><td>${esc(r["Seat status"])}</td>
          <td>${win ? `<span class="chip chip--win">${esc(r["Result (media projections)"])}</span>` : esc(r["Result (media projections)"])}</td>
          <td class="num">${num(r["Total receipts ($)"]) != null ? money(r["Total receipts ($)"]) : "–"}</td><td class="num">${num(r["Total disbursements ($)"]) != null ? money(r["Total disbursements ($)"]) : "–"}</td>
          <td>${u ? `<a href="${esc(u)}" target="_blank" rel="noopener">${esc(r["FEC candidate ID"] || "FEC")} ↗</a>` : ""}</td></tr>`;
      }).join("")}</tbody>`;
    };
    $("#house-filters").addEventListener("input", draw);
    draw();
  }

  function renderMethod() {
    const R = rows("Sources & Method").slice(2).filter((r) => !isEmpty(r));
    let html = "", open = false;
    R.forEach((r) => {
      const a = clean(r[0]), b = clean(r[1]), u = safeUrl(r[2]);
      if (a && !b && a === a.toUpperCase()) { if (open) html += "</div>"; html += `<h2 class="h2">${esc(a.charAt(0) + a.slice(1).toLowerCase())}</h2><div class="method">`; open = true; return; }
      if (!open) { html += `<div class="method">`; open = true; }
      html += `<div class="method__row"><b>${esc(a)}</b><p>${esc(b)}${u ? ` <a href="${esc(u)}" target="_blank" rel="noopener">Source ↗</a>` : ""}</p></div>`;
    });
    if (open) html += "</div>";
    const legend = findText("Stance Grid", "Legend:");
    const scoreDef = subtitle("Candidate Scorecard");
    html += `<h2 class="h2">Definitions</h2><div class="method">
      <div class="method__row"><b>Stances</b><p>${esc(legend.replace(/^Legend:\s*/, ""))}</p></div>
      <div class="method__row"><b>Plan detail</b><p>${esc((findText("Policy Index", "Method:").match(/Plan detail:.*/) || [""])[0].replace(/^Plan detail:\s*/, ""))}</p></div>
      <div class="method__row"><b>Detail score</b><p>${esc(scoreDef)}</p></div>
      <div class="method__row"><b>Source classes</b><p>${esc(subtitle("Claims Log").replace(/^Click[^.]*\.\s*/, ""))}</p></div>
    </div>`;
    $("#method-body").innerHTML = html;
  }

  /* ---------- routing ---------- */
  const VIEWS = ["overview", "policies", "grid", "claims", "candidates", "money", "races", "method"];
  function route() {
    const [name, qs] = (location.hash.slice(1) || "overview").split("?");
    const v = VIEWS.includes(name) ? name : "overview";
    VIEWS.forEach((x) => ($("#view-" + x).hidden = x !== v));
    document.querySelectorAll(".tabs a").forEach((a) => { if (a.getAttribute("href") === "#" + v) a.setAttribute("aria-current", "page"); else a.removeAttribute("aria-current"); });
    const active = $(`.tabs a[href="#${v}"]`); if (active) active.scrollIntoView({ block: "nearest", inline: "nearest" });
    if (v === "claims" && claimState.apply) claimState.apply(Object.fromEntries(new URLSearchParams(qs || "")));
    hideTip();
    window.scrollTo(0, 0);
  }

  fetch("data/tracker.json", { cache: "no-cache" })
    .then((r) => { if (!r.ok) throw new Error(r.status); return r.json(); })
    .then((json) => {
      BOOK = json;
      model();
      renderOverview(); renderPolicies(); renderGrid(); renderClaims(); renderCandidates(); renderMoney(); renderHouse(); renderMethod();
      $("#loading").remove();
      $("#foot-meta").textContent = `Built from “${clean(json.source)}” · data file generated ${clean(json.generated)}.`;
      window.addEventListener("hashchange", route);
      route();
    })
    .catch((e) => { $("#loading").textContent = "Could not load data/tracker.json (" + e.message + "). If you opened index.html directly from disk, use a local server or the published site."; });
})();
