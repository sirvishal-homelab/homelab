// Tiny enhancements. Every page works without this file.
(function () {
  var btn = document.getElementById("zoom-toggle");
  var box = document.getElementById("preview");
  if (btn && box) {
    btn.addEventListener("click", function () {
      var zoomed = box.classList.toggle("zoomed");
      btn.textContent = zoomed ? "Fit to screen" : "Show actual size";
    });
  }
})();
