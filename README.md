# Sashas-third-eye

Two reporting apps, published together on GitHub Pages:

- `index.html`: Internal reporting (Sasha's Third Eye: QA, production, insights and team reports).
  A Scope choice at the top picks AR, Payment Posting or Submission (`?scope=pp`, `?scope=sub`). Each scope runs every report (Production, Agent Quality, QA Performance, Insights, Team Status, Comparison) on its own uploads and keeps its own saved settings. Payment Posting reads the Payment Posting QA sheet the same way the Audit Dashboard does. Payment Posting and Submission also get a details tab built from their production sheets (`scopes.js`). Uploads stay for the visit in the browser's IndexedDB, so switching scope and back keeps them; a new visit, a new or duplicated tab, a reopened tab or a refresh starts empty. If the browser will not keep a sheet, the page says which boxes to upload again after a scope switch. An uploaded team roster with a Department column is narrowed to each scope's department (Posting, Submission, AR), plus anyone with work in that scope's sheets; people with no Department (blank, "-" or "N/A") stay in every scope.
- `client.html`: Client reporting .

Internal Reporting and Client Reporting tabs sit under the heading on both pages. Client Reporting opens on one upload slot per report (AR inventory current and previous month, AR production, submission, payment posting, rejections, EV).

OneDrive: every Excel upload box on both pages has an **Upload from OneDrive** button that opens a file browser on Purviewautomate@purview-healthcare.com's OneDrive only (no other account is accepted) and loads the picked file into that box. On Internal Reporting, the **Quality insights** (QA, sampling, attainment, RCA, targets) save by themselves, one file per client and month, to Brahmastra Reports / <team lead> / <client> / Quality insights, written with xlsx-js-style (`vendor/`); that month's targets and RCA load back (`onedrive.js`). See ONEDRIVE-SETUP.md.
