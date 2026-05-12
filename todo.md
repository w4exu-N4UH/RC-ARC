### TODO List

#### Migration: Personal → Club Accounts

##### Code Changes (in this repo)

- [x] Update SheetMonkey endpoint in all 3 form pages (old: `tsXjKtuurQAQJFvwaGGvuX` → new: `cCtbQ5HZdYrQS7fVKN4dZV`):
  - `content/pages/contact.md`
  - `content/pages/membership.md`
  - `content/pages/elmer.md`
- [x] Update GitHub repo URL references from `mrbell-dev/RC-ARC` → `w4exu-N4UH/RC-ARC` in:
  - `README.md`
  - `CLAUDE.md`
  - `docs/wiki/Getting-Started-with-GitHub.md`
  - `docs/wiki/Previewing-Changes.md`
  - `docs/wiki/Editing-Content.md`
  - `docs/wiki/Creating-New-Posts.md`
  - `docs/wiki/README.md`

##### Manual Tasks (outside the repo)

- [x] **Google Drive setup** — Sheets, PDF template, and PDF folder copied to club account (`rowansociety2@gmail.com`); Email tabs and columns confirmed good
- [x] **Google Apps Script** — Installed on master sheet (`rowanars.net-form-submit`); `installTrigger` and `installDailyTrigger` both run; Discord health check tested and working
- [x] **Google Calendar** — Updated embed to club calendar (`rowansociety2@gmail.com`)
- [x] **GitHub Pages** — Confirmed deploying from `gh-pages` branch, custom domain and HTTPS active
- [x] **GitHub Actions** — Both workflows green in `w4exu-N4UH`: `deploy.yaml` and `update-theme.yaml` (daily cron)
- [x] **GitHub Wiki** — Confirmed transferred to `github.com/w4exu-N4UH/RC-ARC/wiki`
- [x] **Old SheetMonkey** — Keeping in place; only `rowanars.net` can submit to it so no action needed
- [x] **Old Google Apps Script** — Keeping as archive; trigger not active so no action needed

##### Smoke Tests (after push to GitHub)

- [x] Visit [www.rowanars.net](https://www.rowanars.net) — site loads, no broken layout
- [x] Activities page — Google Calendar embed displays correctly and the "Add to Google Calendar" link works
- [x] Contact form — submit a test message and confirm it lands in the club SheetMonkey sheet
- [x] Membership form — submit a test entry and confirm it lands in the club SheetMonkey sheet
- [x] Elmer form — submit a test request and confirm it lands in the club SheetMonkey sheet
- [x] Check [GitHub Actions](https://github.com/w4exu-N4UH/RC-ARC/actions) — deploy workflow shows green checkmark

#### Content Updates

- [ ] **2026 Field Day** — Update `content/events/field-day.md` with 2026 event details; Discord message sent to channel requesting info from chairman (date, location, etc)

#### Future Enhancements

- Set up Obsidian Sync to manage all files in this repository. This would allow editing content on the fly and adding new posts directly from Obsidian.
