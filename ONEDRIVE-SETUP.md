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

Below the upload boxes, a **Saved work** line keeps a daily log of what supervisors upload and set up, with
weekly and monthly summaries, for every day of the year:

- **Save to OneDrive** opens the same OneDrive window. The first time, the supervisor opens their client
  folder (for example My files › Sashas Third Eye › Shreya › JH), types their name once (shown as Saved by)
  and clicks **Save here**; after that the window opens in that folder for that scope. The work goes into
  one file per scope and year, **Third Eye saved work – AR – 2026.xlsx** (or Payment Posting / Submission),
  created the first time. On 1 January the next save starts the 2027 file in the same folder; a date range
  that runs across the new year adds each day to its own year's file.
- The file has these tabs, all rebuilt by Third Eye on every save (edits typed into them are not read back):
  - **Overview**: the latest month for each client (claims, attainment and QA against last month, sampling,
    open RCA items, last save) and that month's weeks.
  - **Daily**: one line per client, biller and work date (claims worked, share-of-day target, attainment,
    claims audited, sampling, QA score, QA errors, top error, saved by, saved on), oldest first, with a day
    total under each client's day. A save adds the work days in view that are not in the file yet. A day
    that is already there is never added twice: if its figures changed (late audits, a new target) the line
    is updated, **Updated on** says when, by whom and what changed, and the old figures are kept in the
    hidden data. Lines from today's save are shaded yellow.
  - **Weekly**: the Daily lines added up per month, week, client, team lead and biller, with team and client
    totals. Weeks run Monday to Sunday and stop at the month's edges (Week 1 of October 2026 is 1–4 Oct), so
    a month's weeks add up to the month. QA score is weighted by claims audited.
  - **Monthly**: one line per client, team lead, biller and month, client totals per month, and the change in
    attainment since the month before. The current month says "(so far)".
  - **RCA log**: a line when an RCA is added or any of its fields change, with what changed (Status: In
    progress → Implemented). Saving it again unchanged adds nothing. An RCA counts as open until Implemented.
  - **Targets log**: a line when a target is set, changed or cleared, with the old and new Target / Day.
  - **Third Eye data** (hidden): the records Third Eye reads back. Please do not edit it.
- Each client is saved on its own, with only the targets that reach that client (its client and practice
  figures, and those of the billers, team leads and weeks that worked it).
- When files are loaded or the dates change, Third Eye searches Purview Automate's OneDrive for every saved
  work file, wherever it is, and matches one client at a time; work saved for client A is never loaded when
  only client B is in view. For each client it takes each supervisor's save for the same dates, or else
  their latest save for these dates or earlier, loads that client's targets and RCA, and says who saved
  them, when and where. If more than one supervisor saved that client, they are listed and the supervisor
  picks one. A save already on screen is not loaded again, so changes made since stay. Saves in the older
  single **Third Eye saved work.xlsx** still load back. The default target covers every client, so when
  several clients are loaded a saved default comes back as that client's own target instead.
- If another supervisor saves to the same file at the same moment, Third Eye reads their save and adds to
  it. If the file is open in the Excel desktop app, OneDrive refuses the change; close it there and save
  again.
- Edits typed into the readable tabs in Excel are not read back; change things in Third Eye and save again.
