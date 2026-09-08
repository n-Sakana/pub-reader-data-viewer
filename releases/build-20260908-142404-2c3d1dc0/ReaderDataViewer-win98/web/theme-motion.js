/* Presentation-only WebView hook. No business messages, timers or mutation observers.
   The native host sends surfaceShown AFTER its size/position is stable. */
(function () {
  'use strict';
  var root = document.documentElement;
  var reduced = window.matchMedia('(prefers-reduced-motion: reduce)');
  var contrast = window.matchMedia('(forced-colors: active)');
  var active = null;
  function cancel() { if (active) { active.cancel(); active = null; } }
  function allowed() {
    return root.getAttribute('data-rdv-modern') === 'true' &&
      root.getAttribute('data-rdv-motion') === 'auto' && !reduced.matches && !contrast.matches;
  }
  function present(event) {
    if (!event.data || event.data.type !== 'surfaceShown' || !allowed()) { return; }
    var body = document.querySelector('.veil.show>.dlg>.body');
    if (!body || !body.animate) { return; }
    cancel();
    var style = window.getComputedStyle(root);
    var duration = Math.min(160, Math.max(80, parseFloat(style.getPropertyValue('--ui-enter')) || 140));
    active = body.animate([{opacity:0.86}, {opacity:1}], {
      duration:duration, easing:style.getPropertyValue('--ui-ease').trim() || 'ease-out', iterations:1
    });
    active.onfinish = function () { active = null; };
  }
  function preferenceChanged() { if (!allowed()) { cancel(); } }
  [reduced, contrast].forEach(function (query) {
    if (query.addEventListener) { query.addEventListener('change', preferenceChanged); }
    else if (query.addListener) { query.addListener(preferenceChanged); }
  });
  if (window.chrome && window.chrome.webview) { window.chrome.webview.addEventListener('message', present); }
})();
