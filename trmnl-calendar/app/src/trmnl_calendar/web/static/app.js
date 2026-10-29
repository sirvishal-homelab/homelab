// Tiny enhancements. Every page works without this file.
(function () {
  document.documentElement.classList.add("js");

  var btn = document.getElementById("zoom-toggle");
  var box = document.getElementById("preview");
  if (btn && box) {
    btn.addEventListener("click", function () {
      var zoomed = box.classList.toggle("zoomed");
      btn.textContent = zoomed ? "Normal size" : "Enlarge";
    });
  }

  // Event form: typing in an end date or a count picks that "Ends" choice.
  [["until", "until"], ["count", "count"]].forEach(function (pair) {
    var field = document.getElementById(pair[0]);
    if (!field) return;
    field.addEventListener("input", function () {
      var radio = document.querySelector("input[name=ends][value=" + pair[1] + "]");
      if (radio && field.value) radio.checked = true;
    });
  });

  // Event form drafts: what you type is kept in this browser, so a failed save, a dropped
  // connection or an accidental close does not lose it. Stored only on this device.
  (function drafts() {
    var DAY = 7 * 24 * 3600 * 1000;
    function store() {
      try { return window.localStorage; } catch (e) { return null; }
    }
    var ls = store();
    var evForm = document.querySelector("form.event[action^='/events/']");
    // Any other page: forget drafts that were submitted (we got somewhere, so the save worked).
    function sweep() {
      if (!ls) return;
      try {
        Object.keys(ls).forEach(function (k) {
          if (k.indexOf("draft:") !== 0) return;
          var d = JSON.parse(ls.getItem(k));
          if (!d || Date.now() - d.t > DAY || (d.sent && !evForm)) ls.removeItem(k);
        });
      } catch (e) { /* storage unavailable: nothing to do */ }
    }
    sweep();
    if (!ls || !evForm || evForm.getAttribute("action").indexOf("/delete") > -1) return;
    var key = "draft:" + evForm.getAttribute("action");
    function read() {
      try { return JSON.parse(ls.getItem(key)); } catch (e) { return null; }
    }
    function snapshot() {
      var out = [];
      Array.prototype.forEach.call(evForm.elements, function (el) {
        if (!el.name || el.type === "hidden" || el.type === "submit" || el.type === "button") return;
        if ((el.type === "checkbox" || el.type === "radio") && !el.checked) return;
        out.push([el.name, el.value]);
      });
      return out;
    }
    var initial = JSON.stringify(snapshot());
    function save(sent) {
      try {
        var now = JSON.stringify(snapshot());
        if (now === initial) { ls.removeItem(key); return; }
        ls.setItem(key, JSON.stringify({ t: Date.now(), sent: !!sent, f: snapshot() }));
      } catch (e) { /* full or blocked: carry on without a draft */ }
    }
    evForm.addEventListener("input", function () { save(false); });
    evForm.addEventListener("change", function () { save(false); });
    evForm.addEventListener("submit", function () { save(true); });
    var d = read();
    if (!d || !d.f || document.querySelector(".notice.error")) return;
    var box = document.createElement("div");
    box.className = "notice";
    box.setAttribute("role", "status");
    var when = new Date(d.t).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
    box.appendChild(document.createTextNode("You have unsaved changes from " + when + ". "));
    function button(label, fn) {
      var b = document.createElement("button");
      b.type = "button";
      b.textContent = label;
      b.addEventListener("click", fn);
      box.appendChild(b);
      box.appendChild(document.createTextNode(" "));
    }
    button("Restore", function () {
      Array.prototype.forEach.call(evForm.elements, function (el) {
        if (!el.name || el.type === "hidden" || el.type === "submit" || el.type === "button") return;
        if (el.type === "checkbox" || el.type === "radio") el.checked = false;
      });
      d.f.forEach(function (p) {
        var els = evForm.querySelectorAll("[name='" + p[0] + "']");
        Array.prototype.forEach.call(els, function (el) {
          if (el.type === "checkbox" || el.type === "radio") { if (el.value === p[1]) el.checked = true; }
          else el.value = p[1];
        });
      });
      Array.prototype.forEach.call(evForm.querySelectorAll("details"), function (x) { x.open = true; });
      box.remove();
      save(false);
    });
    button("Discard", function () { ls.removeItem(key); box.remove(); });
    evForm.parentNode.insertBefore(box, evForm);
  })();

  var form = document.getElementById("view-form");
  if (!form) return;

  // Phone: Settings and Preview tabs.
  var editor = document.getElementById("editor");
  document.querySelectorAll("[data-tab-btn]").forEach(function (b) {
    b.addEventListener("click", function () {
      editor.setAttribute("data-tab", b.getAttribute("data-tab-btn"));
      document.querySelectorAll("[data-tab-btn]").forEach(function (o) {
        o.setAttribute("aria-selected", o === b ? "true" : "false");
      });
    });
  });

  // Day presets.
  var presets = form.querySelector(".presets");
  if (presets) presets.hidden = false;
  form.querySelectorAll("[data-days]").forEach(function (b) {
    b.addEventListener("click", function () {
      var on = b.getAttribute("data-days").split(" ");
      form.querySelectorAll("input[name=weekday]").forEach(function (c) {
        c.checked = on.indexOf(c.value) >= 0;
      });
      changed();
    });
  });

  // Show only the options the chosen layout uses.
  var layout = document.getElementById("layout");
  function syncOptions() {
    var allowed = layout.options[layout.selectedIndex].getAttribute("data-options").split(" ");
    form.querySelectorAll("[data-option]").forEach(function (l) {
      l.hidden = allowed.indexOf(l.getAttribute("data-option")) < 0;
    });
  }

  // Live preview: debounce, send the unsaved settings, swap the image.
  var dirty = false, submitting = false, timer = null, seq = 0;
  var updateBtn = document.getElementById("preview-btn");
  if (updateBtn) updateBtn.hidden = true;

  function collect() {
    var f = form.elements;
    function checked(name) {
      return Array.prototype.filter.call(form.querySelectorAll("input[name=" + name + "]"), function (c) { return c.checked; })
        .map(function (c) { return c.value; });
    }
    function mode(name) {
      var r = form.querySelector("input[name=" + name + "]:checked");
      return r ? r.value : "all";
    }
    var opts = {};
    form.querySelectorAll("[data-option] input").forEach(function (c) { opts[c.name] = c.checked; });
    return {
      view: {
        name: f.name.value || "Preview",
        layout: layout.value,
        weekdays: checked("weekday"),
        categories: { mode: mode("categories_mode"), ids: checked("category").map(Number) },
        people: { mode: mode("people_mode"), ids: checked("person").map(Number) },
        options: opts
      }
    };
  }

  function refresh() {
    var mine = ++seq;
    fetch("/api/v1/views/preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(collect())
    }).then(function (r) { return r.json().then(function (j) { return { ok: r.ok, body: j }; }); })
      .then(function (res) {
        if (mine !== seq) return;
        var fit = document.getElementById("fit");
        var img = document.getElementById("preview-img");
        if (!res.ok) {
          var msgs = res.body.error && res.body.error.fields ? Object.values(res.body.error.fields) : [res.body.error.message];
          fit.innerHTML = "";
          var p = document.createElement("p"); p.className = "err"; p.textContent = msgs.join(" "); fit.appendChild(p);
          return;
        }
        if (img) img.src = res.body.image.url;
        fit.innerHTML = "";
        if (res.body.fit.length) {
          var ul = document.createElement("ul");
          res.body.fit.forEach(function (f) { var li = document.createElement("li"); li.textContent = f.message; ul.appendChild(li); });
          fit.appendChild(ul);
        } else {
          var ok = document.createElement("p"); ok.className = "ok"; ok.textContent = "Everything fits."; fit.appendChild(ok);
        }
      }).catch(function () {});
  }

  function changed() {
    dirty = true;
    clearTimeout(timer);
    timer = setTimeout(refresh, 300);
  }
  form.addEventListener("input", changed);
  form.addEventListener("change", function () { syncOptions(); changed(); });
  form.addEventListener("submit", function () { submitting = true; });
  window.addEventListener("beforeunload", function (e) {
    if (dirty && !submitting) { e.preventDefault(); e.returnValue = ""; }
  });
  syncOptions();
})();
