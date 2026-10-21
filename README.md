# Sashas-third-eye

Two reporting apps, published together on GitHub Pages:

- `index.html`: Internal reporting (Sasha's Third Eye: QA, production, insights and team reports).
  A Scope choice at the top picks AR, Payment Posting or Submission (`?scope=pp`, `?scope=sub`). Each scope runs every report (Production, Agent Quality, QA Performance, Insights, Team Status, Comparison) on its own uploads and keeps its own saved settings. Payment Posting reads the Payment Posting QA sheet the same way the Audit Dashboard does. Payment Posting and Submission also get a details tab built from their production sheets (`scopes.js`).
- `client.html`: Client reporting .

Internal Reporting and Client Reporting tabs sit under the heading on both pages. Client Reporting opens on one upload slot per report (AR inventory current and previous month, AR production, submission, payment posting, rejections, EV).

OneDrive: every Excel upload box on both pages has an **Upload from OneDrive** button that opens a file browser on Purviewautomate@purview-healthcare.com's OneDrive only (no other account is accepted) and loads the picked file into that box. On Internal Reporting, **Save to OneDrive** keeps the targets, RCA and a period summary for the dates in view in "Third Eye saved work.xlsx" in a folder the supervisor picks, and the same dates later load them back (`onedrive.js`). See ONEDRIVE-SETUP.md.
