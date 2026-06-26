## Summary
- Adds parser.py, which reads the most recent XLS/XLSX maintenance report in each aircraft's OneDrive folder and generates site/data.json + site/data.js for the dashboard (no longer hardcodes filenames, since reports get a new date suffix on every update).
- Adds site/ with the static dashboard (index.html / dashboard.html) plus the generated data files — this is the folder meant to be deployed to Cloudflare Pages.
- Adds .github/workflows/update-data.yml, a scheduled GitHub Actions workflow that syncs the OneDrive folders via rclone, reruns parser.py, and deploys the refreshed data to Cloudflare Pages automatically (daily, plus manual trigger).
- The login screen's password is now compared as a SHA-256 hash instead of being stored in plaintext in the page source.

## Test plan
- [x] Ran python parser.py locally against the 4 real OneDrive spreadsheets — confirms it locates the latest file per folder and produces correct site/data.json / site/data.js.
- [ ] Add the required GitHub Actions secrets (RCLONE_CONF, CF_API_TOKEN, CF_ACCOUNT_ID, CF_PROJECT_NAME) so the scheduled workflow can run end-to-end.
- [ ] Trigger the workflow manually (workflow_dispatch) once secrets are set, and confirm the Cloudflare Pages deployment updates with fresh data.
