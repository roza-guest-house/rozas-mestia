/* Basic Cloudflare statistics. No booking events, visitor IDs, cookies or popup. */
(function (root) {
  "use strict";
  if (root.RozaBasicStatistics) return;
  const config = root.ROZA_BASIC_STATISTICS_CONFIG || {};
  const preferenceKey = "roza-basic-statistics-opt-out-v1";
  const token = typeof config.token === "string" ? config.token.trim() : "";
  const correctOrigin = root.location.protocol === "https:" && root.location.origin === config.origin;
  const validPath = path => typeof path === "string" && /^\/[a-z0-9_./-]*$/.test(path) && !path.includes("..");
  const validConfiguration = validPath(config.privacyPath) && Array.isArray(config.paths) && config.paths.length > 0 && config.paths.every(validPath);
  const privacyPage = validPath(config.privacyPath) && root.location.pathname === config.privacyPath;
  const pageEligible = !privacyPage && Array.isArray(config.paths) && config.paths.some(path => validPath(path) && path === root.location.pathname);
  const configured = config.enabled === true && /^[a-f0-9]{32}$/i.test(token) && correctOrigin && validConfiguration;
  const inScope = configured && (pageEligible || privacyPage);
  const state = { optedOut: null, storageAvailable: false, preferenceSaveFailed: false, browserOptOut: false, beaconRequested: false, scriptLoaded: false, loadFailed: false };
  let controlsBound = false;

  function readPreference() {
    try {
      // Any existing value is treated as an opt-out, including an unexpected old value.
      state.optedOut = root.localStorage.getItem(preferenceKey) !== null;
      state.storageAvailable = true;
    } catch (_) {
      state.optedOut = null;
      state.storageAvailable = false;
    }
  }

  function browserPrefersOptOut() {
    try {
      const navigator = root.navigator || {};
      return navigator.globalPrivacyControl === true || [navigator.doNotTrack, navigator.msDoNotTrack, root.doNotTrack]
        .some(value => value === "1" || value === "yes");
    } catch (_) { return true; }
  }

  function getStatus() {
    const reason = config.enabled !== true ? "disabled" : !/^[a-f0-9]{32}$/i.test(token) ? "missing_token" : !correctOrigin ? "wrong_origin" : !validConfiguration ? "invalid_config" :
      !pageEligible && !privacyPage ? "excluded_page" : state.preferenceSaveFailed ? "preference_not_saved" : state.browserOptOut ? "browser_opt_out" : !state.storageAvailable ? "storage_unavailable" :
      state.optedOut ? "opted_out" : privacyPage ? "privacy_page" : state.loadFailed ? "load_failed" : state.scriptLoaded ? "script_loaded" : state.beaconRequested ? "loading" : "not_loaded";
    return { configured, pageEligible: correctOrigin && pageEligible, privacyPage: correctOrigin && privacyPage, optedOut: state.optedOut,
      storageAvailable: state.storageAvailable, preferenceSaveFailed: state.preferenceSaveFailed, browserOptOut: state.browserOptOut, beaconRequested: state.beaconRequested,
      scriptLoaded: state.scriptLoaded, reason };
  }

  function renderControls() {
    if (!privacyPage || !correctOrigin) return;
    const status = getStatus();
    const output = root.document.getElementById("basic-statistics-status");
    const off = root.document.getElementById("basic-statistics-opt-out");
    const on = root.document.getElementById("basic-statistics-opt-in");
    let message = "Basic statistics are not enabled for this website.";
    if (configured) {
      message = status.preferenceSaveFailed ? "Your choice could not be saved. Statistics may still run on other pages or future visits. Try again or use your browser's Do Not Track or Global Privacy Control setting." :
        !status.storageAvailable ? "Basic statistics are off because this browser's preference storage is unavailable." :
        status.browserOptOut ? "Basic statistics are off because this browser sends a privacy preference." :
        status.optedOut ? "Basic statistics are off for future page loads in this browser." :
        "Basic statistics are allowed on the home and booking pages. This privacy page is not measured.";
      message += " A changed choice applies to future page loads; pages already open may finish sending statistics. No page is reloaded.";
    }
    if (output) output.textContent = message;
    if (off) off.disabled = !configured || !status.storageAvailable || status.optedOut === true;
    if (on) on.disabled = !configured || !status.storageAvailable || status.browserOptOut || status.optedOut !== true;
  }

  function setOptOut(optedOut) {
    if (typeof optedOut !== "boolean") return { ok: false, saved: false, reason: "invalid_choice" };
    if (!inScope) return { ok: false, saved: false, reason: getStatus().reason };
    try {
      if (optedOut) root.localStorage.setItem(preferenceKey, "1");
      else root.localStorage.removeItem(preferenceKey);
      readPreference();
      if (!state.storageAvailable || state.optedOut !== optedOut) throw new Error("Preference was not saved.");
      state.preferenceSaveFailed = false;
      renderControls();
      return { ok: true, saved: true, optedOut, appliesToFuturePageLoads: true };
    } catch (_) {
      readPreference();
      state.preferenceSaveFailed = true;
      renderControls();
      return { ok: false, saved: false, reason: "preference_not_saved", appliesToFuturePageLoads: false };
    }
  }

  function bindControls() {
    if (controlsBound || !privacyPage || !correctOrigin) return;
    controlsBound = true;
    const off = root.document.getElementById("basic-statistics-opt-out");
    const on = root.document.getElementById("basic-statistics-opt-in");
    if (off) off.addEventListener("click", () => setOptOut(true));
    if (on) on.addEventListener("click", () => setOptOut(false));
    renderControls();
  }

  function loadBeacon() {
    if (!inScope || !pageEligible || !state.storageAvailable || state.optedOut || state.browserOptOut || state.beaconRequested) return;
    state.beaconRequested = true;
    try {
      const script = root.document.createElement("script");
      script.type = "module";
      script.async = true;
      script.src = "https://static.cloudflareinsights.com/beacon.min.js";
      script.setAttribute("data-cf-beacon", JSON.stringify({ token, spa: false }));
      script.onload = () => { state.scriptLoaded = true; };
      script.onerror = () => { state.loadFailed = true; };
      root.document.head.appendChild(script);
    } catch (_) { state.loadFailed = true; }
  }

  root.RozaBasicStatistics = Object.freeze({ getStatus, setOptOut });
  if (inScope) {
    state.browserOptOut = browserPrefersOptOut();
    readPreference();
    root.addEventListener("storage", event => {
      if (event.key !== preferenceKey && event.key !== null) return;
      readPreference();
      state.preferenceSaveFailed = false;
      renderControls();
      // An existing third-party beacon cannot be unloaded reliably. Never reload a booking.
    });
    loadBeacon();
  }
  if (privacyPage && correctOrigin) {
    if (root.document.readyState === "loading") root.document.addEventListener("DOMContentLoaded", bindControls, { once: true });
    else bindControls();
  }
})(window);
