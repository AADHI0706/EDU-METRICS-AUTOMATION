/* Registrar — Result Ledger: frontend logic */
(() => {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = v => v == null ? "" : String(v).replace(/[&<>"']/g,
    c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const f2 = v => (v == null || v === "" || isNaN(v)) ? "–" : Number(v).toFixed(2);

  let currentId = null;
  const charts = {};

  async function api(url, opts) {
    const res = await fetch(url, opts);
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || res.statusText);
    return data;
  }

  function drawChart(id, config) {
    if (!window.Chart) return;               // CDN unavailable (offline)
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart($("#" + id), config);
  }

  // ---------- Tabs ----------
  function switchTab(name) {
    $$(".tab").forEach(t => t.classList.toggle("active", t.dataset.tab === name));
    $$(".pane").forEach(p => p.classList.toggle("active", p.id === "pane-" + name));
    window.scrollTo({ top: 0 });
  }
  $$(".tab").forEach(t => t.addEventListener("click", () => !t.disabled && switchTab(t.dataset.tab)));
  const enableTabs = () => $$(".tab").forEach(t => t.disabled = false);

  // ---------- Previous intakes ----------
  async function refreshReports() {
    const box = $("#reports-list");
    try {
      const list = await api("/api/reports");
      if (!list.length) { box.innerHTML = '<p class="hint">None yet.</p>'; return; }
      box.innerHTML = list.map(r => {
        const when = r.uploaded_at ? new Date(r.uploaded_at + "Z").toLocaleString() : "";
        return `<button type="button" class="report-item ${r.id === currentId ? "current" : ""}" data-id="${r.id}">
          <span><span class="name">#${r.id} · ${esc(r.result_filename)}</span><br>
          <span class="meta">${r.student_count} students · ${r.semester_count} semesters · ${esc(when)}</span></span>
          <span class="open">Open →</span></button>`;
      }).join("");
      $$(".report-item", box).forEach(b => b.addEventListener("click", () => loadReport(+b.dataset.id)));
    } catch (e) { box.innerHTML = `<p class="status-line err">${esc(e.message)}</p>`; }
  }

  // ---------- Upload ----------
  $("#upload-form").addEventListener("submit", async e => {
    e.preventDefault();
    const btn = $("#upload-btn"), st = $("#upload-status");
    btn.disabled = true; st.className = "status-line"; st.textContent = "Processing…";
    try {
      const d = await api("/api/upload", { method: "POST", body: new FormData(e.target) });
      st.className = "status-line ok";
      st.textContent = `Processed ${d.student_count} students across ${d.semester_count} semesters.`;
      await loadReport(d.report_id);
    } catch (err) {
      st.className = "status-line err"; st.textContent = err.message;
    }
    btn.disabled = false;
  });

  // ---------- Load a report ----------
  async function loadReport(id) {
    currentId = id;
    const tag = $("#active-report-tag");
    tag.textContent = `Ledger #${id}`; tag.classList.add("on");
    $("#dl-excel").href = `/api/reports/${id}/download/excel`;
    $("#dl-html").href = `/api/reports/${id}/download/html`;
    try {
      const [sg, ar, nil, st] = await Promise.all([
        api(`/api/reports/${id}/sgpa`), api(`/api/reports/${id}/arrears`),
        api(`/api/reports/${id}/nil-arrear`), api(`/api/reports/${id}/staff`)]);
      renderRegister(sg, ar, nil, st);
      renderArrears(ar, nil);
      renderStaff(st);
      enableTabs();
      switchTab("dashboard");
      refreshReports();
    } catch (e) {
      const s = $("#upload-status"); s.className = "status-line err"; s.textContent = e.message;
    }
  }

  const regLink = r => `<a class="reg" data-reg="${esc(r)}">${esc(r)}</a>`;
  document.addEventListener("click", e => {
    const a = e.target.closest("a.reg");
    if (a) { e.preventDefault(); lookup(a.dataset.reg); }
  });

  // ---------- Register / dashboard ----------
  function renderRegister(sg, ar, nil, staff) {
    const students = sg.students, sems = sg.semesters;
    const cg = students.map(s => s.cgpa).filter(v => v != null);
    const avg = cg.length ? cg.reduce((a, b) => a + b, 0) / cg.length : null;
    const totalPass = staff.reduce((a, r) => a + (r.pass || 0), 0);
    const totalAll = staff.reduce((a, r) => a + (r.total_students || 0), 0);
    const overall = totalAll ? (totalPass / totalAll * 100) : null;

    const stats = [
      ["Students", students.length, "on the register"],
      ["Semesters", sems.length, sems.join(" · ")],
      ["Average CGPA", f2(avg), "all students"],
      ["Highest CGPA", f2(cg.length ? Math.max(...cg) : null), students[0] ? esc(students[0].name) : ""],
      ["Nil arrear", nil.length, "clean record"],
      ["Arrear entries", ar.length, "open papers"],
      ["Overall pass %", overall == null ? "–" : overall.toFixed(1) + "%", "all subjects"],
    ];
    $("#stat-row").innerHTML = stats.map(([l, v, n]) =>
      `<div class="stat"><div class="label">${esc(l)}</div><div class="value">${v}</div><div class="note">${n}</div></div>`).join("");

    // SGPA distribution (every student-semester SGPA)
    const bins = ["< 5", "5 – 6", "6 – 7", "7 – 8", "8 – 9", "9 – 10"], counts = [0, 0, 0, 0, 0, 0];
    students.forEach(s => Object.values(s.sgpa).forEach(v => {
      if (v == null) return;
      counts[v < 5 ? 0 : v < 6 ? 1 : v < 7 ? 2 : v < 8 ? 3 : v < 9 ? 4 : 5]++;
    }));
    drawChart("sgpa-chart", {
      type: "bar",
      data: { labels: bins, datasets: [{ label: "Semester results", data: counts, backgroundColor: "#7a1f2b" }] },
      options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } },
    });

    // Top of the register
    $("#rank-list").innerHTML = students.slice(0, 10).map(s =>
      `<li><span class="nm">${esc(s.name)}<br><span class="rg">${esc(s.reg_no)}</span></span><span class="cg">${f2(s.cgpa)}</span></li>`).join("")
      || '<li class="empty">No data</li>';

    // Full register
    const head = $("#sgpa-table thead tr");
    head.innerHTML = ["Rank", "Reg No", "Name", ...sems, "CGPA"].map((h, i) =>
      `<th class="${i === 0 || i > 2 ? "num" : ""}">${esc(h)}</th>`).join("");
    $("#sgpa-table tbody").innerHTML = students.map(s =>
      `<tr><td class="num">${esc(s.rank ?? "–")}</td><td class="mono">${regLink(s.reg_no)}</td><td>${esc(s.name)}</td>` +
      sems.map(m => `<td class="num">${f2(s.sgpa[m])}</td>`).join("") +
      `<td class="num"><b>${f2(s.cgpa)}</b></td></tr>`).join("");
  }

  // ---------- Arrears ----------
  function renderArrears(ar, nil) {
    $("#arrear-table tbody").innerHTML = ar.length ? ar.map(r =>
      `<tr><td class="mono">${regLink(r.reg_no)}</td><td>${esc(r.name)}</td><td>${esc(r.semester)}</td>
       <td class="mono">${esc(r.subject)}</td><td><span class="badge fail">${esc(r.grade)}</span></td></tr>`).join("")
      : '<tr><td colspan="5" class="empty">No arrears on record.</td></tr>';
    $("#nil-table tbody").innerHTML = nil.length ? nil.map(r =>
      `<tr><td class="mono">${regLink(r.reg_no)}</td><td>${esc(r.name)}</td></tr>`).join("")
      : '<tr><td colspan="2" class="empty">None.</td></tr>';
  }

  // ---------- Faculty ----------
  function renderStaff(rows) {
    const color = p => p >= 90 ? "#2f6b4f" : p >= 70 ? "#a86b12" : "#7a1f2b";
    drawChart("staff-chart", {
      type: "bar",
      data: {
        labels: rows.map(r => `${r.subject} (${r.semester})`),
        datasets: [{ label: "Pass %", data: rows.map(r => r.pass_percent), backgroundColor: rows.map(r => color(r.pass_percent)) }],
      },
      options: {
        plugins: { legend: { display: false } },
        scales: { y: { min: 0, max: 100 }, x: { ticks: { maxRotation: 70, minRotation: 45, font: { size: 10 } } } },
      },
    });
    $("#staff-table tbody").innerHTML = rows.length ? rows.map(r => {
      const cls = r.pass_percent >= 90 ? "pass" : r.pass_percent >= 70 ? "mid" : "fail";
      return `<tr><td>${esc(r.semester)}</td><td class="mono">${esc(r.subject)}</td><td class="mono">${esc(r.faculty_id)}</td>
        <td>${esc(r.faculty_name)}</td><td class="num">${r.total_students}</td><td class="num">${r.pass}</td>
        <td class="num">${r.fail}</td><td><span class="badge ${cls}">${Number(r.pass_percent).toFixed(1)}%</span></td></tr>`;
    }).join("") : '<tr><td colspan="8" class="empty">No faculty data.</td></tr>';
  }

  // ---------- Student lookup ----------
  async function lookup(reg) {
    reg = (reg || "").trim();
    if (!reg || currentId == null) return;
    switchTab("student");
    $("#lookup-input").value = reg;
    const out = $("#lookup-result");
    out.innerHTML = '<p class="hint">Searching…</p>';
    try {
      const d = await api(`/api/reports/${currentId}/student/${encodeURIComponent(reg)}`);
      let h = `<div class="student-head"><h3>${esc(d.name)}</h3><span class="mono">${esc(d.reg_no)}</span></div>
        <div class="student-stats"><div>CGPA<b>${f2(d.cgpa)}</b></div><div>Rank<b>${esc(d.rank ?? "–")}</b></div></div>
        <div class="sem-grid">`;
      for (const [sem, subs] of Object.entries(d.subjects)) {
        h += `<div class="sem-block"><h4><span>${esc(sem)}</span><span class="mono">SGPA ${f2(d.sgpa[sem])}</span></h4>
          <div class="table-wrap"><table><thead><tr><th>Subject</th><th>Grade</th><th class="num">Credit</th><th>Status</th></tr></thead><tbody>` +
          subs.map(s => `<tr><td class="mono">${esc(s.subject)}</td><td>${esc(s.grade)}</td><td class="num">${esc(s.credit)}</td>
            <td><span class="badge ${s.is_arrear ? "fail" : "pass"}">${s.is_arrear ? "FAIL" : "PASS"}</span></td></tr>`).join("") +
          "</tbody></table></div></div>";
      }
      out.innerHTML = h + "</div>";
    } catch (e) { out.innerHTML = `<p class="status-line err">${esc(e.message)}</p>`; }
  }
  $("#lookup-btn").addEventListener("click", () => lookup($("#lookup-input").value));
  $("#lookup-input").addEventListener("keydown", e => e.key === "Enter" && lookup(e.target.value));


  // ---------- Opened without the server? ----------
  if (location.protocol === "file:") {
    const b = document.createElement("div");
    b.style.cssText = "background:#7a1f2b;color:#fff;padding:12px 28px;font-weight:600;text-align:center";
    b.textContent = "This page was opened as a file. Start the program (python app.py or the .exe) and open http://127.0.0.1:5000 instead.";
    document.body.prepend(b);
  }

  // ---------- Init ----------
  (async () => {
    await refreshReports();
    try {
      const list = await api("/api/reports");
      if (list.length) await loadReport(list[0].id);   // reopen the latest ledger automatically
    } catch (_) { /* stay on intake */ }
  })();
})();