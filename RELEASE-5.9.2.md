# Web 5.9.2

- Compact PN notes: six aligned 36px options, no internal scrolling; outside click closes without losing selection. Done button retained. Photo behavior unchanged.
- Successful keyed reconciliation schedules the next daily pass instead of a 10-minute loop. Startup, reconnect, lookup misses and manual requests can also initiate reconciliation; failed passes retain retry backoff. New source changes are discovered at the next pass, not instantaneously.
- API VTD emails cache errors from its executing account fulfillment.wms.3pl@gmail.com to khiempham5209@gmail.com. Identical cache/error signatures across devices are limited to one email per six hours. Requires the deployment to execute as that account and MailApp permission. Devices must reach the API to report an error. Mail failures are logged and returned as alertError; cache reports still save.

Deploy index.html and redeploy Code_chinh_VTD.gs. PN API and APK unchanged.
Validation: 20 tests pass; browser verified all six options at 36px, outside dismissal, retained selection, and six IndexedDB fixture checks. Email tested with mocked transport; actual delivery requires deployment verification.
