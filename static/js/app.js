/* ==========================================================================
 * SheetForge — frontend logic
 * ========================================================================== */
"use strict";

(function () {
  const $ = (id) => document.getElementById(id);

  const ALLOWED_EXT = [".xlsx", ".xlsm", ".xls", ".csv"];
  const COMPRESS_EXT = [".xlsx", ".xlsm"];
  const MAX_FILES = 30;

  const TABS = {
    merge: { endpoint: "merge", label: "Merge files" },
    compress: { endpoint: "compress", label: "Compress files" },
    zip: { endpoint: "zip", label: "Create ZIP bundle" },
  };

  const state = { tab: "merge", files: [] };

  /* ------------------------------ helpers ------------------------------ */
  function fmtBytes(n) {
    if (n === null || n === undefined) return "—";
    if (n < 1024) return n + " B";
    const units = ["KB", "MB", "GB"];
    let i = -1;
    let v = n;
    do { v /= 1024; i++; } while (v >= 1024 && i < units.length - 1);
    return v.toFixed(v >= 100 ? 0 : 1) + " " + units[i];
  }

  function extOf(name) {
    const i = name.lastIndexOf(".");
    return i < 0 ? "" : name.slice(i).toLowerCase();
  }

  function fileIcon(name) {
    const ext = extOf(name);
    if (ext === ".csv") return "📄";
    if (ext === ".xls") return "📗";
    return "📊";
  }

  function addFiles(fileList) {
    let rejected = [];
    for (const f of fileList) {
      if (state.files.length >= MAX_FILES) {
        rejected.push({ name: f.name, why: "file limit reached (30)" });
        break;
      }
      if (!ALLOWED_EXT.includes(extOf(f.name))) {
        rejected.push({ name: f.name, why: "unsupported type" });
        continue;
      }
      if (f.size === 0) {
        rejected.push({ name: f.name, why: "file is empty" });
        continue;
      }
      state.files.push(f);
    }
    renderFileList();
    if (rejected.length) {
      showError(
        "Some files were skipped:<br>" +
        rejected.map((r) => "• <b>" + escapeHtml(r.name) + "</b> — " + r.why).join("<br>")
      );
    }
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    }[c]));
  }

  function renderFileList() {
    const ul = $("filelist");
    ul.innerHTML = "";
    state.files.forEach((f, idx) => {
      const li = document.createElement("li");
      li.className = "filechip";
      li.innerHTML =
        '<span class="f-icon">' + fileIcon(f.name) + "</span>" +
        '<span class="f-meta">' +
        '  <span class="f-name" title="' + escapeHtml(f.name) + '">' + escapeHtml(f.name) + "</span>" +
        '  <span class="f-size">' + fmtBytes(f.size) + "</span>" +
        "</span>" +
        '<button class="f-remove" type="button" aria-label="Remove ' +
        escapeHtml(f.name) + '">×</button>';
      li.querySelector(".f-remove").addEventListener("click", () => {
        state.files.splice(idx, 1);
        renderFileList();
      });
      ul.appendChild(li);
    });
  }

  /* ------------------------------ tabs -------------------------------- */
  function switchTab(tab) {
    state.tab = tab;
    document.querySelectorAll(".tab").forEach((b) => {
      const active = b.dataset.tab === tab;
      b.classList.toggle("active", active);
      b.setAttribute("aria-selected", active ? "true" : "false");
    });
    ["merge", "compress", "zip"].forEach((t) => {
      $("options-" + t).classList.toggle("hidden", t !== tab);
    });
    $("runbtn-label").textContent = TABS[tab].label;
    hideResults();
  }

  /* --------------------------- validation ------------------------------ */
  function validateFiles() {
    if (state.files.length === 0) {
      showError("Please add at least one file first.");
      return false;
    }
    if (state.tab === "compress") {
      const bad = state.files.filter((f) => !COMPRESS_EXT.includes(extOf(f.name)));
      if (bad.length) {
        showError(
          "The compressor works on <b>.xlsx / .xlsm</b> files only. " +
          "“" + escapeHtml(bad[0].name) + "” can't be compressed — " +
          "use the ZIP bundle tab for other formats."
        );
        return false;
      }
    }
    return true;
  }

  /* ---------------------------- processing ----------------------------- */
  function buildFormData() {
    const fd = new FormData();
    state.files.forEach((f) => fd.append("files", f));
    if (state.tab === "merge") {
      const mode = document.querySelector('input[name="mode"]:checked').value;
      fd.append("mode", mode);
      fd.append("header", $("opt-header").checked ? "1" : "0");
      fd.append("add_source", $("opt-source").checked ? "1" : "0");
      fd.append("dedupe", $("opt-dedupe").checked ? "1" : "0");
      fd.append("include_all", $("opt-include-all").checked ? "1" : "0");
      if (mode === "stack") fd.append("strategy", $("strategy").value);
    }
    if (state.tab === "compress") {
      const preset = document.querySelector('input[name="preset"]:checked').value;
      fd.append("preset", preset);
      fd.append("max_dim", $("max_dim").value || "1600");
      fd.append("quality", $("quality").value || "72");
    }
    return fd;
  }

  function setProgress(pct, text) {
    const fill = $("fill");
    fill.classList.remove("pulse");
    fill.style.width = Math.max(2, Math.min(100, pct)) + "%";
    $("status").textContent = text;
  }

  function setProcessing(text) {
    const fill = $("fill");
    fill.classList.add("pulse");
    fill.style.width = "100%";
    $("status").textContent = text;
  }

  async function run() {
    hideResults();
    hideError();

    if (!validateFiles()) return;

    const btn = $("runbtn");

    btn.disabled = true;
    $("progress").classList.remove("hidden");

    const fd = buildFormData();

    try {
        setProgress(5, "Uploading files...");

        const response = await fetch("/api/merge", {
            method: "POST",
            body: fd
        });

        const data = await response.json();

        if (!response.ok || !data.ok) {
            throw new Error(
                data.error || `Server error (${response.status})`
            );
        }

        // The server has accepted the job.
        // It is now processing in the background.
        setProgress(10, "Files uploaded. Starting merge...");

        pollJob(data.job_id);

    } catch (error) {
        $("progress").classList.add("hidden");
        btn.disabled = false;

        showError(
            escapeHtml(
                error.message || "Could not start the merge."
            )
        );
    }
}
  async function pollJob(jobId) {

    const btn = $("runbtn");

    const timer = setInterval(async () => {

        try {

            const response = await fetch(`/api/jobs/${jobId}`);

            const job = await response.json();

            if (!response.ok || !job.ok) {
                clearInterval(timer);

                $("progress").classList.add("hidden");
                btn.disabled = false;

                showError(
                    escapeHtml(
                        job.error || "Could not retrieve merge status."
                    )
                );

                return;
            }

            setProgress(
                job.progress || 0,
                job.message || "Processing..."
            );

            if (job.status === "completed") {

                clearInterval(timer);

                $("progress").classList.add("hidden");
                btn.disabled = false;

                const result = {
                    ok: true,
                    stats: job.result?.stats || {},
                    download: job.download,
                    name: job.result?.file || "merged.xlsx"
                };

                renderResults(result);

                return;
            }

            if (job.status === "failed") {

                clearInterval(timer);

                $("progress").classList.add("hidden");
                btn.disabled = false;

                showError(
                    escapeHtml(
                        job.error || "The merge failed."
                    )
                );

                return;
            }

        } catch (error) {

            clearInterval(timer);

            $("progress").classList.add("hidden");
            btn.disabled = false;

            showError(
                "Network error while checking merge status."
            );
        }

    }, 1500);
}
  /* ------------------------------ results ------------------------------ */
  function statTile(label, value) {
    return '<div class="stat-tile"><b>' + escapeHtml(String(value)) + "</b><span>" +
           escapeHtml(label) + "</span></div>";
  }

  function mergeResults(resp) {
    const s = resp.stats || {};
    return (
      statTile("Files", s.files || 0) +
      statTile("Input rows", (s.input_rows || 0).toLocaleString()) +
      statTile("Output rows", (s.output_rows || 0).toLocaleString()) +
      statTile("Columns", s.columns || "—") +
      (s.duplicates_removed
        ? statTile("Duplicates removed", s.duplicates_removed.toLocaleString())
        : "") +
      statTile("Sheets", s.sheets || "1")
    );
  }

  function compressResults(resp) {
    const s = resp.stats || {};
    return (
      statTile("Original", fmtBytes(s.original_bytes)) +
      statTile("Compressed", fmtBytes(s.final_bytes)) +
      statTile("Saved", fmtBytes(s.saved_bytes)) +
      statTile("Reduction", (s.percent_saved || 0) + "%") +
      statTile("Images optimized", (s.images_compressed || 0) + " / " + (s.images || 0)) +
      statTile("Removed parts", (s.removed_parts || []).length)
    );
  }

  function zipResults(resp) {
    const s = resp.stats || {};
    return (
      statTile("Files", s.files || 0) +
      statTile("Total size", fmtBytes(s.total_bytes)) +
      statTile("ZIP size", fmtBytes(s.zip_bytes)) +
      statTile("Saved", fmtBytes((s.total_bytes || 0) - (s.zip_bytes || 0)))
    );
  }

  function renderResults(resp) {
    const el = $("results");
    el.classList.remove("hidden");

    const titles = {
      merge: "✅ Merged successfully",
      compress: "✅ Compression complete",
      zip: "✅ ZIP bundle created",
    };
    const subs = {
      merge: (r) => r.stats.mode === "sheets"
        ? "Every file is now a separate sheet in one workbook."
        : "All rows combined into a single sheet with aligned columns.",
      compress: (r) => "Your optimized file" + (r.results ? "s" : "") + " are ready to download.",
      zip: () => "All files packed into one archive.",
    };
    const title = titles[state.tab];
    const sub = subs[state.tab](resp);

    let statsHtml = "";
    let downloads = "";

    if (state.tab === "merge") {
      statsHtml = mergeResults(resp);
      downloads = downloadLink(resp.download, "Download " + resp.name);
    } else if (state.tab === "compress") {
      const results = resp.results || [{ stats: resp.stats, download: resp.download, name: resp.name }];
      statsHtml = results.map((r) => compressResults(r)).join("");
      downloads = results
        .map((r) => downloadLink(r.download, "Download " + r.name))
        .join("");
      if (results.length > 1) {
        downloads +=
          '<p class="hint" style="color:#047857;margin:6px 0 0">' +
          results.map((r) => escapeHtml(r.file)).join(", ") + "</p>";
      }
    } else {
      statsHtml = zipResults(resp);
      downloads = downloadLink(resp.download, "Download " + resp.name);
    }

    el.innerHTML =
      "<h3>" + title + "</h3>" +
      '<p class="r-sub">' + sub + "</p>" +
      '<div class="stats-grid">' + statsHtml + "</div>" +
      '<div class="downloads">' + downloads + "</div>" +
      '<button type="button" class="again" id="again-btn">Start over with new files</button>';

    $("again-btn").addEventListener("click", () => {
      state.files = [];
      renderFileList();
      hideResults();
    });
  }

  function downloadLink(url, text) {
    return '<a class="dlbtn" href="' + url + '">' +
           '<svg viewBox="0 0 24 24" width="17" height="17"><path d="M12 4v12m0 0 5-5m-5 5-5-5M5 20h14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>' +
           escapeHtml(text) + "</a>";
  }

  function hideResults() { $("results").classList.add("hidden"); $("results").innerHTML = ""; }
  function hideError() { $("error").classList.add("hidden"); $("error").innerHTML = ""; }
  function showError(msg) {
    const el = $("error");
    el.innerHTML = "⚠️ " + msg;
    el.classList.remove("hidden");
  }

  /* ------------------------------ events ------------------------------- */
  const dz = $("dropzone");
  const input = $("file-input");

  dz.addEventListener("click", () => input.click());
  dz.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
  input.addEventListener("change", () => { addFiles(input.files); input.value = ""; });

  ["dragenter", "dragover"].forEach((ev) =>
    dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("dragover"); })
  );
  ["dragleave", "drop"].forEach((ev) =>
    dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("dragover"); })
  );
  dz.addEventListener("drop", (e) => addFiles(e.dataTransfer.files));

  document.querySelectorAll(".tab").forEach((b) =>
    b.addEventListener("click", () => switchTab(b.dataset.tab))
  );
  $("runbtn").addEventListener("click", run);

  // merge mode radio cards
  document.querySelectorAll('input[name="mode"]').forEach((r) =>
    r.addEventListener("change", () => {
      document.querySelectorAll('input[name="mode"]').forEach((x) =>
        x.closest(".radio-card").classList.toggle("active", x.checked)
      );
      $("stack-only").classList.toggle("hidden", document.querySelector('input[name="mode"]:checked').value !== "stack");
    })
  );
  // compress preset radio cards
  document.querySelectorAll('input[name="preset"]').forEach((r) =>
    r.addEventListener("change", () => {
      document.querySelectorAll('input[name="preset"]').forEach((x) =>
        x.closest(".radio-card").classList.toggle("active", x.checked)
      );
    })
  );

  $("quality").addEventListener("input", () => {
    $("quality-val").textContent = $("quality").value;
  });

  // prevent accidental file-list leaks
  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => e.preventDefault());
})();
