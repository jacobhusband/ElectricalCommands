# Microsoft Outlook Consent Troubleshooting

This app uses Microsoft Entra desktop OAuth with these effective settings:

- App registration: `Desktop Application`
- Tenant ID: `417e5aea-af22-46eb-9b47-8667938cdcaf`
- Client ID: `ad3bdf05-66e2-4a1a-9660-4ca1ba237007`
- Redirect URI family: `http://localhost:{ephemeral_port}/`
- Delegated Microsoft Graph scope: `Mail.Read`

## Tenant policy checks

1. Open Microsoft Entra admin center.
2. Go to `App registrations > Desktop Application > Authentication`.
3. Verify `Mobile and desktop applications` includes `http://localhost`.
4. Verify `Allow public client flows` is enabled.
5. Go to `App registrations > Desktop Application > API permissions`.
6. Confirm Microsoft Graph permissions are delegated only.
7. Keep `Mail.Read` and the standard sign-in scopes.
8. Remove unexpected application permissions or extra delegated permissions if they were added accidentally.
9. Go to `Enterprise applications > Consent and permissions > User consent settings`.
10. Check the current tenant setting.
11. Preferred target: `Allow limited user consent only for apps from verified publishers and apps registered in your tenant, for selected permissions`.
12. Go to `Enterprise applications > Consent and permissions > Permission classifications`.
13. Classify the delegated permissions the app requests as low impact where appropriate, especially `Mail.Read`, `offline_access`, and `User.Read` if Microsoft added it to the app registration.
14. Go to `Enterprise applications > Desktop Application > Properties`.
15. Check `User assignment required?`.
16. Preferred default: set it to `No`.
17. If policy requires assignment, assign `jacobh@acies.net` or the appropriate user/group.

## Retest

1. Start the desktop app.
2. Open the Outlook mail dialog.
3. Click `Connect Outlook`.
4. If the app reports a tenant consent-policy issue, use the in-app `Retry minimal consent` action.
5. If reduced-scope consent succeeds, treat that as evidence that `Mail.Read` or `offline_access` is what tenant policy is blocking.
6. Sign in with the target Microsoft 365 account again after the tenant settings are updated.
7. Confirm the `Need admin approval` page no longer appears.
8. Confirm the app returns to the local callback and shows the account as connected.
9. Run `Scan inbox` and verify Outlook messages load successfully.

## Notes

- The app currently requests delegated `Mail.Read`, not application-wide mailbox access.
- Delegated `Mail.Read` is normally user-consentable, so the approval screen usually indicates tenant policy or enterprise app access restrictions.
- The app can now run a reduced-scope diagnostic with `openid profile email User.Read` to isolate whether `Mail.Read` or `offline_access` is the trigger.
- If the issue remains after tenant policy and assignment checks, fall back to tenant-wide admin consent or a custom app consent policy for this internal app.
