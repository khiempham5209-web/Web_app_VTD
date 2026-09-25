// Kiem thu hang doi mang 5.9.14: gioi han song song, uu tien thao tac nguoi dung,
// dung chung request trung, tu thu lai lenh doc khi mat phan hoi, gop bao cao tien do cache.
const fs = require('fs'), vm = require('vm'), assert = require('assert/strict');
const html = fs.readFileSync(require('path').join(process.argv[2] || '.', 'index.html'), 'utf8');
const start = html.indexOf('    // Hàng đợi mạng dùng chung');
const end = html.indexOf('    function timeoutResult(');
assert.ok(start > 0 && end > start, 'network queue block not found');
const source = html.slice(start, end);

function makeContext(respond) {
  const calls = [];
  let active = 0, maxActive = 0;
  const ctx = {
    console, Promise, Map, Set, JSON, Object, String, Number, Date, Math, setTimeout, clearTimeout,
    window: {}, live: {token: 't'}, APP_VERSION: 'test', nativeCallbacks: {},
    activeApiUrl: () => 'https://script.google.com/macros/s/x/exec',
    setLoading: () => {}, handleApiResponse: (action, res) => res,
    readApiJson: r => r, AbortController: undefined,
    scriptApiFetch: (url, payload) => {
      const call = {action: payload.action, params: payload.params, done: null};
      calls.push(call); active++; maxActive = Math.max(maxActive, active);
      return new Promise(resolve => { call.done = res => { active--; resolve(res); }; if (respond) respond(call); });
    }
  };
  vm.createContext(ctx);
  vm.runInContext(source, ctx);
  return {ctx, calls, max: () => maxActive, active: () => active};
}
const tick = (ms = 0) => new Promise(r => setTimeout(r, ms));
async function test(name, fn) { await fn(); console.log('PASS', name); }

(async () => {
  await test('Background requests capped at 3, low-priority reports at 2, total at 4', async () => {
    const t = makeContext();
    const actions = ['init','dashboard','masterSkuKeySync','cacheDataManifest','lookupOrderManifest','queueErrors','forgotPassRequests','loginReport','storeInfo','queueRetryRequests'];
    const all = actions.map((a, i) => t.ctx.apiTransport(a, {_silent: true, n: i}));
    await tick();
    assert.equal(t.calls.length, 3, 'only 3 background requests start');
    let lowActive = t.calls.filter(c => ['queueErrors','forgotPassRequests','loginReport','storeInfo','queueRetryRequests'].includes(c.action)).length;
    assert.equal(lowActive, 0, 'core data goes before admin/report requests');
    while (t.calls.some(c => !c.finished)) {
      const next = t.calls.find(c => !c.finished);
      next.finished = true; next.done({ok: true, action: next.action});
      await tick();
      const running = t.calls.filter(c => !c.finished);
      assert.ok(running.length <= 3);
      assert.ok(running.filter(c => ['queueErrors','forgotPassRequests','loginReport','storeInfo','queueRetryRequests'].includes(c.action)).length <= 2);
    }
    const results = await Promise.all(all);
    assert.equal(results.length, actions.length);
    assert.equal(t.max(), 3);
  });

  await test('Saving an order starts immediately even when background slots are full', async () => {
    const t = makeContext();
    ['init','dashboard','masterSkuKeySync','lookupOrderManifest'].forEach(a => t.ctx.apiTransport(a, {_silent: true}));
    await tick();
    assert.equal(t.calls.length, 3);
    t.ctx.apiTransport('saveRaw', {_silent: true, id: 'R1'});
    t.ctx.apiTransport('search', {q: 'x'});
    await tick();
    assert.equal(t.calls[3].action, 'saveRaw', 'reserved slot used by the order save');
    assert.equal(t.calls.length, 4, 'total cap 4 respected');
  });

  await test('Identical background reads share one network call; user-forced reads are not shared', async () => {
    const t = makeContext(c => setTimeout(() => c.done({ok: true, rows: [1]}), 5));
    const r = await Promise.all([1,2,3,4,5,6].map(() => t.ctx.apiTransport('forgotPassRequests', {_silent: true})));
    assert.equal(t.calls.filter(c => c.action === 'forgotPassRequests').length, 1);
    assert.ok(r.every(x => x.ok));
    await Promise.all([t.ctx.apiTransport('queueErrors', {_silent: false}), t.ctx.apiTransport('queueErrors', {_silent: false})]);
    assert.equal(t.calls.filter(c => c.action === 'queueErrors').length, 2);
  });

  await test('Lost response: reads retried once, writes never retried', async () => {
    const lost = {ok: false, code: 'INVALID_API_RESPONSE', httpStatus: 404, message: 'API hệ thống trả về HTML (HTTP 404, not_found).'};
    let n = 0;
    const t = makeContext(c => setTimeout(() => c.done(c.action === 'dashboard' && n++ === 0 ? lost : c.action === 'saveRaw' ? lost : {ok: true}), 1));
    const read = await t.ctx.apiTransport('dashboard', {_silent: true});
    assert.equal(read.ok, true);
    assert.equal(read.retriedAfterLostResponse, true);
    assert.equal(t.calls.filter(c => c.action === 'dashboard').length, 2);
    const write = await t.ctx.apiTransport('saveRaw', {_silent: true});
    assert.equal(write.ok, false);
    assert.equal(t.calls.filter(c => c.action === 'saveRaw').length, 1);
    const typeErr = makeContext(c => setTimeout(() => c.done({ok: false, message: 'TypeError: Load failed'}), 1));
    await typeErr.ctx.apiTransport('uploadImage', {_silent: true});
    assert.equal(typeErr.calls.length, 1, 'image upload not retried automatically');
    const timeout = makeContext(c => setTimeout(() => c.done({ok: false, timeout: true, message: 'API phản hồi quá lâu'}), 1));
    await timeout.ctx.apiTransport('dashboard', {_silent: true});
    assert.equal(timeout.calls.length, 1, 'client timeout not retried');
  });

  await test('Cache progress reports coalesced; done/error sent immediately and cancel pending progress', async () => {
    const t = makeContext(c => c.done({ok: true}));
    const r1 = await t.ctx.apiTransport('reportLookupCacheStatus', {_silent: true, cacheType: 'masterSku', status: 'checking'});
    const r2 = await t.ctx.apiTransport('reportLookupCacheStatus', {_silent: true, cacheType: 'masterSku', status: 'downloading', downloaded: 100});
    await t.ctx.apiTransport('reportLookupCacheStatus', {_silent: true, cacheType: 'masterSku', status: 'downloading', downloaded: 300});
    assert.equal(r1.deferred, true); assert.equal(r2.deferred, true);
    assert.equal(t.calls.length, 0, 'progress steps not sent immediately');
    await tick(3200);
    assert.equal(t.calls.length, 1, 'one combined progress report');
    assert.equal(t.calls[0].params.downloaded, 300, 'latest progress sent');
    await t.ctx.apiTransport('reportLookupCacheStatus', {_silent: true, cacheType: 'masterSku', status: 'downloading', downloaded: 350});
    await t.ctx.apiTransport('reportLookupCacheStatus', {_silent: true, cacheType: 'masterSku', status: 'done', downloaded: 357});
    assert.equal(t.calls.length, 2);
    assert.equal(t.calls[1].params.status, 'done');
    await tick(3200);
    assert.equal(t.calls.length, 2, 'pending progress dropped after done');
    await t.ctx.apiTransport('reportLookupCacheStatus', {_silent: true, cacheType: 'localQueue', status: 'error', lastError: 'x'});
    assert.equal(t.calls.length, 3, 'errors always sent immediately');
  });
  console.log('TOTAL network queue tests passed');
})().catch(err => { console.error(err); process.exit(1); });
