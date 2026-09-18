/* Passive JSON diagnostics: never log bodies, query strings, tokens or customer data. */
(function () {
  'use strict';
  if (window.vtdJsonDiagnostics) return;
  const responses = new WeakMap();
  let active = 0, sequence = 0;
  function fingerprint(value) {
    let h = 2166136261;
    for (const c of String(value)) h = Math.imul(h ^ c.charCodeAt(0), 16777619);
    return (h >>> 0).toString(16);
  }
  function endpoint(value) {
    try { const u = new URL(value); return {requestHost:u.hostname, endpointId:fingerprint(u.origin + u.pathname)}; }
    catch (_) { return {}; }
  }
  function classify(text) {
    const s = String(text).slice(0, 100000).toLowerCase();
    if (/too many requests|quota exceeded|service invoked too many|rate limit/.test(s)) return 'quota_or_rate_limit';
    if (/accounts.google.com|servicelogin|sign in to continue/.test(s)) return 'google_login';
    if (/access denied|permission denied|you need access/.test(s)) return 'access_denied';
    if (/page not found|file you have requested does not exist|404 not found/.test(s)) return 'not_found';
    if (/temporarily unavailable|internal server error/.test(s)) return 'server_error';
    return /^\s*</.test(s) ? 'unclassified_html' : 'non_json';
  }
  function emit(detail) { try { window.vtdDiagnostics?.record('json_diagnostic', detail); } catch (_) {} }
  const fetchOriginal = window.fetch;
  window.fetch = function (url, options) {
    let host = ''; try { host = new URL(typeof url === 'string' ? url : url.url).hostname; } catch (_) {}
    if (host !== 'script.google.com') return fetchOriginal.apply(this, arguments);
    // Diagnostics uploads use their own saved fetch and do not pass through this wrapper.
    const start = Date.now(), requestId = 'json_' + start.toString(36) + '_' + (++sequence);
    const info = Object.assign(endpoint(typeof url === 'string' ? url : url.url), {requestId, transport:'web_fetch', concurrentRequests:++active});
    let pending;
    try { pending = fetchOriginal.apply(this, arguments); } catch (e) { active--; throw e; }
    return Promise.resolve(pending).then(response => {
      responses.set(response, Object.assign(info, {durationMs:Date.now()-start, redirected:!!response.redirected, responseType:response.type}));
      return response;
    }).finally(() => { active--; });
  };
  const readOriginal = window.readApiJson;
  if (readOriginal) window.readApiJson = async function (response, label, action) {
    const metadata = Object.assign({}, responses.get(response), {label, action, status:response.status, contentType:response.headers?.get('content-type') || ''});
    try { metadata.responseHost = new URL(response.url).hostname; } catch (_) {}
    return window.parseApiResponse(await response.text(), metadata);
  };
  const parseOriginal = window.parseApiResponse;
  if (parseOriginal) window.parseApiResponse = function (text, metadata) {
    const result = parseOriginal.apply(this, arguments);
    try {
      const m = metadata || {};
      const invalid = result?.code === 'INVALID_API_RESPONSE';
      const mismatch = m.action === 'skuSync' && result?.ok && result.protocol !== 1;
      if (invalid || mismatch) emit(Object.assign({}, m, {
        transport:m.transport || 'native_bridge',
        responseLength:String(text).length,
        pageCategory:invalid ? classify(text) : 'json_protocol_mismatch',
        protocolPresent:!!result && Object.prototype.hasOwnProperty.call(result, 'protocol'),
        servicePresent:!!result && Object.prototype.hasOwnProperty.call(result, 'service'),
        reason:invalid ? result.responseCategory : 'unexpected_json_shape',
        jsonDiagnosticsVersion:'1'
      }));
    } catch (_) {}
    return result;
  };
  window.vtdJsonDiagnostics = {version:'1'};
  try { window.vtdDiagnostics?.record('json_diagnostics_loaded', {jsonDiagnosticsVersion:'1'}); } catch (_) {}
})();
