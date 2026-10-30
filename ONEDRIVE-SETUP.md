# Sasha's Third Eye and OneDrive

Every Excel upload box, on Internal Reporting and Client Reporting, opens OneDrive: click the box or its **Upload from OneDrive** button. Files are never picked from the computer or dragged in; that is turned away with a message.
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

## Quality insights (Internal Reporting)

Third Eye saves what it works out from the QA tracker and the production sheet (not the sheets themselves) by
itself, a few seconds after anything changes, in the same folders Brahmastra uses for each team lead:

```
Brahmastra Reports / <Team lead> / <Client> /
   Production        Brahmastra's own production file
   QA tracker        the QA file the lead keeps; the QA Sheet box opens here when Third Eye is opened from Brahmastra
   Quality insights  Third Eye: "Quality - <Client> - AR - Oct 2026.xlsx", one file per client, scope and month
   SLA & KPI         Drishti
```

- The team lead comes from Brahmastra when Third Eye is opened inside it (My day & reports, Team quality &
  production). Opened on its own, it is the lead most of the client's billers report to on the roster. A client
  with two team leads gets the file under each. Folder names are cleaned the same way Brahmastra cleans them.
- Saving needs Purview Automate's OneDrive signed in on that computer; it never opens a sign-in window by
  itself. The **Quality insights** line under the upload boxes says when it last saved, has **Save now**, and
  shows **Sign in to OneDrive** when needed. If the file is open in the Excel desktop app, it tries again a
  minute later. Nothing saves before the client's saved targets and RCA for that month have loaded.
- The file has these tabs, rebuilt on every save (edits typed into them are not read back):
  - **Overview**: the month's claims, attainment and QA against last month, sampling, open RCA items, a
    **Needs attention** box for the latest week next to **Doing well**, and the month's weeks.
  - **Daily**: one line per biller and work date, with a day total. A day is never added twice; a day whose
    figures changed (late audits, a new target) is updated and **Updated on** says what changed.
  - **Weekly**: one line per biller with the month's weeks across (claims, attainment, QA) and the month at the end.
  - **Monthly**: the month per biller with the change from last month (read from last month's file) and an
    Outlier flag.
  - **Sampling**: each biller against the sampling band for their status (PIP 100%, OJT, Production,
    Observation, a detected Outlier), with the verdict and the audits still needed, and the weeks across.
  - **RCA & Mitigation**: a line when an RCA is added or changes, with what changed.
  - **Targets**: a line when a target is set, changed or cleared, with the old and new Target / Day.
  - **Third Eye data** (hidden): the records Third Eye reads back. Please do not edit it.
- When a client and month are opened again, that month's targets and RCA load back from its file; a new month
  starts from last month's. If another supervisor saves the same file at the same moment, Third Eye reads
  their save and adds to it.
- Older "Third Eye saved work" files are left where they are and no longer written.
