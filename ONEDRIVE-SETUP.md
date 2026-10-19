# Sasha's Third Eye and OneDrive

Internal Reporting can read its files straight from OneDrive and write its results back into the
production file. This follows the same pattern as Purview Brahmastra.

## Where the files live

All files sit in the OneDrive of **Purviewautomate@purview-healthcare.com**:

- Production: Brahmastra's master production workbook for each client (`Brahmastra / <client> / ...xlsx`).
- QA tracker and team roster: anywhere in that OneDrive, for example `Sashas Third Eye / AR /`.
- `Sashas Third Eye / Third Eye links.json`: which file each client, scope and box uses. Third Eye writes
  this file itself when someone links a file. It holds file IDs and client names only, never claim data.

Share files with specific people only (never "Anyone with the link"), and give Can edit only to people who
click Update.

## Using it

1. Sign in to Microsoft on the OneDrive files bar with your Purview work account.
2. Link each file once: click **Link file**, paste the file's OneDrive link (right-click the file in
   OneDrive, Copy link) and click **Save link**. Everyone then uses the same file. **Change link** replaces it.
3. Pick the dates worked and click **Fetch**. The files load into the upload boxes as if picked by hand.
   Every month sheet with Date Worked and Worked By columns is read; columns and tabs whose names start
   with "Brahmastra" or "Third Eye" are left out. With "Previous month too" ticked, the month before the
   From date goes into Previous Production for the Comparison tab.
4. Click **Update** to write the results back into each production file:
   - columns at the right of each month sheet: Third Eye Audited, Third Eye QA Score, Third Eye QA Errors,
     Third Eye Note Check, Third Eye Updated By, Third Eye Updated On;
   - a **Third Eye summary** tab with one row per biller per date: claims worked, target, attainment,
     claims audited, sampling, QA score and QA errors.

   Only those columns and that tab are written. The team's columns and Brahmastra's are never touched.
   Rows are matched on Date Worked, Worked By and Claim No; if rows moved since Fetch, nothing is written
   to that file and you are asked to Fetch again. Only dates whose results changed are rewritten.
   OneDrive version history is the undo.

## Microsoft setup (one time, done 19 Oct 2026)

Third Eye has its own app registration in Entra, **Purview Sasha's Third Eye** (client ID
`2e3408ab-40a3-4e0d-8985-8a1474e867e1`, tenant `94be55e5-2194-40e4-9d98-1dcc7e91c268`, single tenant):

- Single-page application redirect URI: `https://purview-healthcare.github.io/Sashas-third-eye/auth-redirect.html`
- Delegated Microsoft Graph permissions, admin-consented: User.Read, Files.ReadWrite.All, offline_access, openid, profile

No client secret is needed. People on purviewservices.com need a guest invite to purview-healthcare.com first.
The bar only shows on the team site, because Microsoft only signs people in at the registered address.
