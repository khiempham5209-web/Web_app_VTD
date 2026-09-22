
// Bounded evidence only; never store complete HTML or successful business JSON.
function vtdDiagRedact_(value, limit) {
  return String(value || '').slice(0,100000)
    .replace(/<script\b[^>]*>[\s\S]*?<\/script\s*>/gi,' ')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style\s*>/gi,' ')
    .replace(/<!--[\s\S]*?-->/g,' ').replace(/<[^>]*>/g,' ')
    .replace(/&(?:nbsp|lt|gt|quot|amp|#\d+);/gi,' ')
    .replace(/https?:\/\/[^\s<>"']+/gi,'[URL]')
    .replace(/[\w.+-]+@[\w.-]+\.[a-z]{2,}/gi,'[EMAIL]')
    .replace(/((?:sessionToken|access_token|id_token|password|passwd|authorization|cookie|token)\s*[=:]\s*)[^\s,;]+/gi,'$1[REDACTED]')
    .replace(/Bearer\s+\S+/gi,'Bearer [REDACTED]')
    .replace(/[A-Za-z0-9_-]{28,}/g,'[ID]').replace(/\b\d{7,}\b/g,'[NUMBER]')
    .replace(/\s+/g,' ').trim().slice(0,limit||600);
}
function vtdDiagEvidence_(text, status) {
  const raw=String(text||'').slice(0,100000);
  const title=(raw.match(/<title\b[^>]*>([\s\S]*?)<\/title>/i)||[])[1]||'';
  const plain=vtdDiagRedact_(raw,10000);
  const rules=[
    ['quota_or_rate_limit',/service invoked too many times|quota exceeded|too many requests|rate limit|dịch vụ bị gọi quá nhiều/i],
    ['google_login',/sign in to continue|sign in with your google account|đăng nhập để tiếp tục/i],
    ['access_denied',/access denied|permission denied|you need access|authorization is required/i],
    ['not_found',/the file you have requested does not exist|page not found|404[ .:-]*not found/i],
    ['bad_request',/malformed|400[ .:-]*bad request|invalid request/i],
    ['javascript_reference',/ReferenceError:\s*[A-Za-z_$][\w$]* is not defined/i],
    ['javascript_type',/TypeError:\s*Cannot read properties of (?:undefined|null)|TypeError:\s*[A-Za-z_$][\w.$]* is not a function/i],
    ['json_parse',/SyntaxError:\s*Unexpected token|Unexpected end of JSON input/i],
    ['timeout',/exceeded maximum execution time|execution timed out/i],
    ['server_error',/internal server error|temporarily unavailable/i]
  ];
  let match,category='unknown';for(const rule of rules){const m=plain.match(rule[1]);if(m){category=rule[0];match=m[0];break;}}
  return {htmlTitle:vtdDiagRedact_(title,180),errorExcerpt:match||'',evidenceCategory:category,
    conclusionBasis:match?'response_text':'insufficient_evidence',
    diagnosticConclusion:match?'Phản hồi có thông báo: '+category+'; cần đối chiếu server để xác định nguồn lỗi.':'Chưa xác định nguyên nhân; HTTP '+(status||'không rõ')+' không đủ bằng chứng.',
    evidenceTruncated:String(text||'').length>100000};
}

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
      if (invalid || mismatch) emit(Object.assign({}, m, invalid ? vtdDiagEvidence_(text,m.status) : {diagnosticConclusion:'JSON trả về thiếu protocol=1 của skuSync; cần đối chiếu action và server.',conclusionBasis:'json_shape'}, {
        transport:m.transport || 'native_bridge',
        responseLength:String(text).length,
        pageCategory:invalid ? classify(text) : 'json_protocol_mismatch',
        protocolPresent:!!result && Object.prototype.hasOwnProperty.call(result, 'protocol'),
        servicePresent:!!result && Object.prototype.hasOwnProperty.call(result, 'service'),
        reason:invalid ? result.responseCategory : 'unexpected_json_shape',
        jsonDiagnosticsVersion:'3'
      }));
    } catch (_) {}
    return result;
  };
  window.vtdJsonDiagnostics = {version:'3'};
  try { window.vtdDiagnostics?.record('json_diagnostics_loaded', {jsonDiagnosticsVersion:'3'}); } catch (_) {}
})();
