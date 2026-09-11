# 5.9.6 Master SKU startup and staff recovery

- Master reconciliation runs before PN; unrelated background traffic no longer blocks idle indefinitely. Foreground busy waits have a 90-second error limit.
- Earlier scheduled work cannot be postponed by later scheduling calls. Daily attempt is claimed after idle, with visible checking/error reports. New marker namespace replaces stale marks from the earlier implementation.
- First real master data replaces demo products even while input is open. Existing real selections defer refresh as before.
- Every signed-in user sees own cache status in Settings / App update. Manual download requires downloadCache (Tải lại cache máy này) or admin config permission; it affects this device only. The API advertises the new action for permission management.
- Reporting/email failures remain visible locally. Email delivery still requires the device to reach the API and Google's mail service to accept it; there is no end-to-end delivery proof from unit tests.

Deploy Code_chinh_VTD.gs, grant downloadCache to staff, and install VTD-ChungTu-v5.9.6.apk over the existing app. PN API unchanged. Update latestVersion/APK link in shared settings after deployment. Do not uninstall to update.

Validation: 30 automated tests, including scheduler ordering, idle timeout, daily claims, same-named photo accumulation, staff permission visibility, and first master replacing demo products. APK signature verified against existing signing certificate. Actual letuyet device download and email receipt require deployment validation.
