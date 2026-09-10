# Web 5.9.1

- Restore compact PN notes dropdown with multiple selections and exclusive Normal option.
- Web photo capture retains previous photos, snapshots each new file, and prevents submitting before image reading finishes. All five image groups use the shared handler.
- Show cache error details on the current device and update committed progress after each batch. Source changes trigger pending reconciliation instead of a generic failure.
- MASTER SKU API coalesces identical records by material code; conflicting contents remain errors. No source spreadsheet rows are deleted.

Deployment: publish index.html and redeploy Code_chinh_VTD.gs to the existing Apps Script deployment. ChungTuFF_WebApi.gs is unchanged. APK 5.9 remains unchanged; these client fixes are web 5.9.1.

Validation: 19 automated tests pass. Dropdown checked in a mobile-sized browser; physical iPhone camera verification still required. Initial PN transient error was not captured; later device screenshot shows 1352/1352.
