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
