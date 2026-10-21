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
3. Open the folders, click the file and click **Upload** (or double-click the file). It always opens at Purview
   Automate's **My files**, laid out like OneDrive (Name, Modified, Modified By, File size, Sharing), and only
   file types the box accepts are listed.
4. **Sign out of OneDrive** at the bottom of the browser forgets the Purviewautomate sign-in on that computer.

Clicking the box itself still uploads from the computer, as before.

## Microsoft setup (one time, done 19 Oct 2026)

Third Eye has its own app registration in Entra, **Purview Sasha's Third Eye** (client ID
`2e3408ab-40a3-4e0d-8985-8a1474e867e1`, tenant `94be55e5-2194-40e4-9d98-1dcc7e91c268`, single tenant):

- Single-page application redirect URI: `https://purview-healthcare.github.io/Sashas-third-eye/auth-redirect.html`
- Delegated Microsoft Graph permissions, admin-consented: User.Read, Files.ReadWrite.All, offline_access, openid, profile

No client secret is needed. The buttons only show on the team site, because Microsoft only signs people in
at the registered address.

## Saved work (Internal Reporting)

Below the upload boxes, a **Saved work** line keeps what supervisors set up for a date range, so it is not
done twice:

- **Save to OneDrive** saves, for the scope and production dates in view: the targets as adjusted, the RCA &
  mitigation plans of the clients in view, and a period summary (one row per client and biller: claims worked,
  target, attainment, claims audited, sampling, QA score, QA errors, top error). It opens the same OneDrive
  window, starting in the folder the production file came from; the supervisor can pick any folder, types
  their name once (shown as Saved by) and clicks **Save here**. The work goes into
  **Third Eye saved work.xlsx** in that folder (created the first time), with the tabs RCA & Mitigation,
  Targets and Period summary for people to read, and a hidden "Third Eye data" tab that Third Eye reads back.
  Each client is saved as its own record, with only the targets that reach that client (its client and
  practice figures, and those of the billers, team leads and weeks that worked it). Saving the same scope,
  dates and client again replaces only that client's save, in that file.
- When files are loaded or the dates change, Third Eye searches Purview Automate's OneDrive for every saved
  work file, wherever it is, and matches one client at a time. Work saved for client A is never loaded when
  only client B is in view, even for the same dates. For each client with a save for the same scope and
  dates, it loads that client's targets and RCA and says who saved them, when and where; other clients'
  targets and RCA are left as they are. If two supervisors saved the same client and dates, both are listed
  and the supervisor picks one. For new dates it offers **Copy from** an earlier save of that client.
  The default target covers every client, so when several clients are loaded a saved default comes back as
  that client's own target instead.
- Edits typed into the readable tabs in Excel are not read back; change things in Third Eye and save again.
