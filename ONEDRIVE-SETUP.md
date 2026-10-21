# Sasha's Third Eye and OneDrive

Every Excel upload box, on Internal Reporting and Client Reporting, has an **Upload from OneDrive** button.
It opens a small file browser on the OneDrive of **Purviewautomate@purview-healthcare.com** only. The
picked file loads into that box exactly as if it had been chosen from the computer.

## Where the files live

Supervisors keep their own folders in the Purviewautomate OneDrive (for example `Brahmastra / <client> /`)
and pick from there. Third Eye only reads files; it never writes anything to OneDrive.

## Using it

1. Click **Upload from OneDrive** on the box you want to fill.
2. The first time on a computer, click **Sign in to Purview Automate's OneDrive** and sign in as
   Purviewautomate@purview-healthcare.com. Any other account is turned away and forgotten again, so a
   person's own OneDrive never opens. After that the sign-in is remembered on that computer.
3. Open the folders, click the file and click **Upload** (or double-click the file). Only file types the
   box accepts are listed. Each box reopens in the folder it last used.
4. **Sign out of OneDrive** at the bottom of the browser forgets the Purviewautomate sign-in on that computer.

Clicking the box itself still uploads from the computer, as before.

## Microsoft setup (one time, done 19 Oct 2026)

Third Eye has its own app registration in Entra, **Purview Sasha's Third Eye** (client ID
`2e3408ab-40a3-4e0d-8985-8a1474e867e1`, tenant `94be55e5-2194-40e4-9d98-1dcc7e91c268`, single tenant):

- Single-page application redirect URI: `https://purview-healthcare.github.io/Sashas-third-eye/auth-redirect.html`
- Delegated Microsoft Graph permissions, admin-consented: User.Read, Files.ReadWrite.All, offline_access, openid, profile

No client secret is needed. The buttons only show on the team site, because Microsoft only signs people in
at the registered address.
