# Connecting Sasha's Third Eye to OneDrive

"Push to OneDrive" and "Read from OneDrive" sign each team lead in with their own Microsoft work
account and save uploads to **their** OneDrive, under:

```
Sashas Third Eye / Internal Reporting / <name> / <date time> /
Sashas Third Eye / Client Reporting  / <client> / <date time> /
```

Each saved folder holds the uploaded files and a `manifest.json` that says which upload box each file came from.

## One-time setup (Microsoft 365 admin, about 5 minutes)

1. Sign in to https://entra.microsoft.com and go to **Identity > Applications > App registrations > New registration**.
2. **Name:** `Sasha's Third Eye`
3. **Supported account types:** *Accounts in this organizational directory only* (single tenant).
4. **Redirect URI:** platform **Single-page application (SPA)**, address:
   `https://purview-healthcare.github.io/Sashas-third-eye/auth-redirect.html`
   (the site's published address followed by `auth-redirect.html`).
5. Click **Register**.
6. Open **API permissions > Add a permission > Microsoft Graph > Delegated permissions** and add
   `User.Read` and `Files.ReadWrite`. Then click **Grant admin consent** so team leads are not asked
   to approve it themselves.
7. On the app's **Overview** page, copy the **Application (client) ID** and the **Directory (tenant) ID**.

No client secret or certificate is needed. Neither ID is a secret.

## Turning it on

Put the two IDs into `CONFIG` at the top of `onedrive.js`:

```js
clientId: '<Application (client) ID>',
tenantId: '<Directory (tenant) ID>',
```

Until both are filled in, the buttons explain that OneDrive isn't connected yet.

## Notes

- Files go straight from the browser to Microsoft. This site has no server and never sees the data.
- Each team lead sees only their own OneDrive folder unless they share it.
- Pop-ups must be allowed for the site, because Microsoft sign-in opens in a small window.
