# Web 5.9.3

- PN SKU cache uses material keys in a new pnSku store; legacy cache remains available during migration. API skuSync compares fingerprints and returns only new/changed SKU records plus inventory. Identical duplicates coalesce; missing materials and conflicting duplicates report errors.
- Validate the entire response before one atomic IndexedDB transaction. Timeout, invalid payload, or failed commit retains the previous cache. Persist and display SKU error status.
- Daily attempt claims persist in IndexedDB for PN SKU, master SKU, PN order reconciliation and PN append. Reopening/reconnecting does not repeat these attempts the same day, even after failure. Manual cache action bypasses the limit. A reconciliation may contain multiple requests when needed; operational lookups/uploads are unaffected.
- Remove email deduplication: each received error report requests an immediate email. Offline clients cannot deliver reports; mail transport/quota failures are still possible and are logged.

Deploy BOTH ChungTuFF_WebApi.gs (new skuSync action) and Code_chinh_VTD.gs (email behavior), then load web 5.9.3 and run cache manually once if recovering from an error. APK unchanged.
Validation: 23 tests, including PN SKU deltas/duplicates, cache retention on timeout/commit failure, daily attempt persistence and email per error report. No real email sent during testing.
