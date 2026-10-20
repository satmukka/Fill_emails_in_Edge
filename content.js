(function () {
  // Toggle: if the panel already exists, remove it.
  const existing = document.getElementById("aef-panel");
  if (existing) {
    existing.remove();
    return;
  }

  // Track the last focused editable element that is NOT part of our panel.
  let target = null;
  const isEditable = (el) =>
    el && (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable);
  document.addEventListener(
    "focusin",
    (e) => {
      if (e.target.closest && e.target.closest("#aef-panel")) return;
      if (isEditable(e.target)) target = e.target;
    },
    true
  );
  if (isEditable(document.activeElement)) target = document.activeElement;

  let emails = [];
  let stop = false;
  let running = false;

  const panel = document.createElement("div");
  panel.id = "aef-panel";
  panel.style.cssText =
    "position:fixed;top:12px;right:12px;z-index:2147483647;width:270px;background:#fff;" +
    "border:1px solid #888;border-radius:6px;padding:10px;font:13px Segoe UI,Arial,sans-serif;" +
    "color:#111;box-shadow:0 2px 10px rgba(0,0,0,.3)";
  panel.innerHTML =
    '<div style="font-weight:600;margin-bottom:6px">Email Filler</div>' +
    '<input type="file" id="aef-file" accept=".xlsx,.xls,.csv" style="width:100%;margin-bottom:6px">' +
    '<div style="margin-bottom:6px">Wait before Enter (ms): ' +
    '<input id="aef-d1" type="number" value="1000" style="width:70px"></div>' +
    '<div style="margin-bottom:6px">Wait after Enter (ms): ' +
    '<input id="aef-d2" type="number" value="1000" style="width:70px"></div>' +
    '<button id="aef-start" disabled>Start</button> ' +
    '<button id="aef-stop" disabled>Stop</button> ' +
    '<button id="aef-close">Close</button>' +
    '<div id="aef-status" style="margin-top:8px;color:#444">Pick the Excel file.</div>';
  document.body.appendChild(panel);

  const $ = (id) => panel.querySelector("#" + id);
  const status = (t) => ($("aef-status").textContent = t);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // Keep focus in the recipient box when clicking panel buttons.
  panel.addEventListener("mousedown", (e) => {
    if (e.target.tagName === "BUTTON") e.preventDefault();
  });

  $("aef-close").onclick = () => {
    stop = true;
    panel.remove();
  };
  $("aef-stop").onclick = () => {
    stop = true;
  };

  $("aef-file").onchange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
      const seen = new Set();
      emails = [];
      // Scan every cell, keep anything that looks like an email (skips headers automatically).
      for (const row of rows) {
        for (const cell of row) {
          const v = String(cell == null ? "" : cell).trim();
          if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) && !seen.has(v.toLowerCase())) {
            seen.add(v.toLowerCase());
            emails.push(v);
          }
        }
      }
      status(emails.length + " emails loaded. Click in the recipient box, then Start.");
      $("aef-start").disabled = emails.length === 0;
    } catch (err) {
      status("Could not read file: " + err.message);
    }
  };

  function setText(el, text) {
    el.focus();
    if (el.isContentEditable) {
      document.execCommand("selectAll", false, null);
      document.execCommand("insertText", false, text);
      return;
    }
    const proto = el.tagName === "TEXTAREA" ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, "value").set.call(el, text);
    el.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertFromPaste", data: text }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function pressEnter(el) {
    const opts = { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true, cancelable: true };
    el.dispatchEvent(new KeyboardEvent("keydown", opts));
    el.dispatchEvent(new KeyboardEvent("keypress", opts));
    el.dispatchEvent(new KeyboardEvent("keyup", opts));
  }

  $("aef-start").onclick = async () => {
    if (running) return;
    if (!target || !document.contains(target)) {
      status("Click inside the recipient box first, then press Start.");
      return;
    }
    running = true;
    stop = false;
    $("aef-start").disabled = true;
    $("aef-stop").disabled = false;
    const d1 = parseInt($("aef-d1").value, 10) || 1000;
    const d2 = parseInt($("aef-d2").value, 10) || 1000;
    let i = 0;
    for (; i < emails.length && !stop; i++) {
      status("Entering " + (i + 1) + " / " + emails.length + ": " + emails[i]);
      setText(target, emails[i]);
      await sleep(d1);
      pressEnter(target);
      await sleep(d2);
    }
    status(stop ? "Stopped after " + i + " / " + emails.length + "." : "Done: " + i + " emails entered.");
    running = false;
    $("aef-start").disabled = false;
    $("aef-stop").disabled = true;
  };
})();
