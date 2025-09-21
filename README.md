# Google Sheets → Email Rejections
A friendly, point-and-click tool for Recruiters, HR, and Talent Acquisition to send personalized, templated rejection emails from a Google Sheet.

This README walks you through setup and day-to-day use. No coding required.

---

## What you can do

- Read candidates from a Google Sheet (your pipeline/export).
- Personalize subject & body using placeholders (name, role, company…).
- Choose plain-text and/or HTML templates.
- Add attachments (e.g., “Interview Tips” PDF).
- Run in **Dry Run** (no emails) or **Send** mode.
- **Test Send To Me** (sends one example to yourself).
- Progress bar + live log; safely **resumes** without re-sending people already marked as sent.
- Saves your last settings next to the app for easy reuse.

---

## Quick start (10–15 minutes)

### 1) Get the files
Place these two files together in a normal, writable folder on your computer:

```
your-folder/
├─ gui_rejections_app.py        # the app you run
└─ rejections_core.py           # the engine (don’t edit)
```

The app will create a `rejections_gui_settings.json` alongside these to remember your settings.

### 2) Install Python + packages
- Install **Python 3.10+** (3.11/3.12 recommended).
- Open a terminal (Command Prompt/PowerShell on Windows; Terminal on macOS).
- (Optional but recommended) create a virtual environment.
- Install dependencies:

```bash
python -m pip install --upgrade pip
python -m pip install customtkinter jinja2 google-api-python-client google-auth google-auth-oauthlib
```

> Tip (macOS): use `python3` instead of `python` if needed.

### 3) Create a Google OAuth client (one-time)
This gives the app permission to access your Sheets & Gmail.

1. Go to Google Cloud Console → **APIs & Services** → **Credentials**.  
2. Create **OAuth client ID** → **Desktop app**.  
3. Download the JSON → save it next to the app as **`credentials.json`**.

> If your org restricts OAuth, your Google Workspace admin may need to approve the app scopes:
> - `https://www.googleapis.com/auth/spreadsheets`
> - `https://www.googleapis.com/auth/gmail.send`

### 4) Prepare your Google Sheet
Have (at minimum) these columns in the first row (header):

- `email` (candidate’s email)
- `name`
- `role`
- `company`

Optional but useful columns (add if you want to use them in templates):

- `stage`, `reason`, `application_date`
- **Gates:**  
  - `send` → only rows with `yes` are processed (if this column exists)  
  - `skip` → rows with `yes` are always skipped
- **Status (auto-managed by the app):**  
  - `sent_status` (app writes “sent”)  
  - `sent_at` (app writes timestamp)

> The app will auto-add `sent_status` and `sent_at` if missing.  
> Re-running the app will **not** re-send to rows where `sent_status` already says `sent`.

### 5) Make simple templates
Put these next to the app, or anywhere on disk you can browse to.

**`template.txt`** (required; plain text)
```
Hi {{ name }},

Thank you for your interest in {{ company }} and the {{ role }} role.
After careful review, we’re moving forward with other candidates at this time.

We appreciate the time you spent applying{{ " and interviewing" if stage|lower == "onsite" else "" }} and wish you the best in your search.

Best,
{{ sender_name }}
{{ sender_title }}
```

**`template.html`** (optional; nicer formatting)
```html
<p>Hi {{ name }},</p>
<p>Thank you for your interest in <strong>{{ company }}</strong> and the <strong>{{ role }}</strong> role.
After careful review, we’re moving forward with other candidates at this time.</p>
<p>We appreciate the time you spent applying{% if stage|lower == "onsite" %} and interviewing{% endif %} and wish you the best in your search.</p>
<p>Best,<br>
{{ sender_name }}<br>
{{ sender_title }}</p>
```

Subject example (set in the app):  
`Regarding your application for {{ role }} at {{ company }}`

Available placeholders: `name`, `email`, `role`, `company`, `stage`, `reason`, `application_date`, `sender_name`, `sender_title`.

---

## Running the app

1. In your terminal, cd into the folder and run:
   ```bash
   python gui_rejections_app.py
   ```
2. The **Config** tab:
   - **Sender (email address):** Use your Gmail/Workspace address or approved alias.
   - **Reply-To:** Where replies should go (often a shared inbox).
   - **Cc/Bcc:** Optional.
   - **Throttle:** Seconds to wait between sends (default 2.0).
   - **Domain throttle:** Extra delay per domain (optional; helps avoid bursts to the same company).
   - **Preview N:** Limit to the first N candidates (leave 0 to send all eligible).
   - **Spreadsheet ID:** Copy the long ID from your Google Sheet URL.
   - **Preferred Tab:** The tab name (e.g., `Applicants`). Case-insensitive.
   - **Read Range:** Usually `A:Z` (or narrow it if you want).
   - **Credentials JSON:** Browse to your downloaded `credentials.json`.
   - **Token JSON:** A file the app creates after your first Google sign-in (defaults to `token.json` next to the app).
   - **Sender Signature:** Shown in templates if not overridden by environment variables.

3. **Templates** tab:
   - Choose your `template.txt` (required) and optional `template.html`.

4. **Attachments** tab:
   - Add one or more files.  
   - Select any item and click **Remove selected** to remove it.  
   - **Remove missing** drops files that no longer exist on disk.

5. **Run** tab:
   - **Dry Run:** Parses the sheet and renders emails **without sending**.
   - **Test Send To Me:** Sends the **first eligible** email to you (From: your Sender; To: you; CC/BCC suppressed; subject prefixed with `[TEST]`).
   - **Send:** Sends to all eligible rows (respecting gates and throttles).
   - **Cancel:** Politely stops the run after the current item.

> Settings are saved to `rejections_gui_settings.json` right next to the app, so you don’t have to re-enter them next time.

---

## How eligibility works (who gets an email)

A row is eligible if:

- Required fields are present: `email`, `name`, `role`, `company`.
- If a `send` column exists, it must be `yes`.
- If a `skip` column exists, it must **not** be `yes`.
- `sent_status` is not already `sent`.

When an email is successfully sent (not dry run, not test mode), the app writes:

- `sent_status = sent`
- `sent_at = <timestamp>`

You can safely re-run the app; previously sent rows are skipped.

---

## Tips for good sending hygiene

- Start with **Preview N = 5** and **Dry Run** to verify subjects and content.
- Use **Test Send To Me** to see exactly what lands in an inbox.
- Keep **Throttle ≥ 2 seconds** and consider a **Domain throttle** (e.g., `5–10 s`) for large sends.
- Use **Cc/Bcc** sparingly.
- If sending from an alias, ensure your Gmail / Workspace is configured to **Send mail as** that alias.

---

## First-time Google sign-in & the token file

On your first run, Google will open a browser window to ask for permission. After you approve:

- The app tries to write the token to the path in **Token JSON** (default `token.json` next to the app).
- If that folder is read-only, the app **falls back** to saving `token.json` under your home folder at:
  ```
  ~/.rejections_gui/token.json
  ```
  and will show a clear message in the Run log like:
  > “Cannot write token to ‘…/token.json’. Wrote it to ‘~/.rejections_gui/token.json’ instead. Update your Token JSON path in the GUI to this location.”

Just copy that fallback path into the **Token JSON** field, click **Save Settings**, and you’re set.

---

## Spreadsheet setup example

Header row (row 1):

```
email | name | role | company | stage | reason | application_date | send | skip | sent_status | sent_at
```

Example rows:

```
jane@domain.com | Jane Doe | Product Manager | Acme | onsite | skills gap | 2025-08-01 | yes |   |
joe@domain.com  | Joe Lin  | Backend Engineer | Acme | screen | scheduling | 2025-08-03 | yes |   |
```

- If you **don’t** have a `send` column, the app will process **all** rows that have the required fields and are not marked as sent.
- Add `skip = yes` for any one-off rows you want to ignore.

---

## Personalization (templates)

You can use these placeholders in subject/text/HTML:

- `{{ name }}`, `{{ email }}`, `{{ role }}`, `{{ company }}`
- `{{ stage }}`, `{{ reason }}`, `{{ application_date }}`
- `{{ sender_name }}`, `{{ sender_title }}`

Conditionals are supported (Jinja2):

```jinja
{% if stage|lower == "onsite" %}
Thank you again for the time you spent with our team.
{% endif %}
```

---

## Common workflows

### Bulk reject with careful checks
1) Set **Preview N = 10** → **Dry Run**  
2) **Test Send To Me** → read the actual message in your inbox  
3) Uncheck **Dry run** → **Send**

### Send only some candidates
- Add `send` column and set `yes` only for rows you want.
- Or add `skip` for rows to exclude.

### Add a helpful PDF
- Go to **Attachments** → **Add attachment(s)** → pick your PDF.
- Reuse the same attachments next time (they’re saved in your settings file).

---

## Troubleshooting

**“Permission denied” when creating token**  
- You’re writing `token.json` to a folder you can’t write to.  
- Check the Run log—if it says it saved to `~/.rejections_gui/token.json`, copy that path into the **Token JSON** field and click **Save Settings**.

**“Missing required columns”**  
- The app didn’t find one of: `email`, `name`, `role`, `company`.  
- Fix your header row (row 1) so those names exist (spelling matters).

**“No rows found” or wrong tab**  
- Confirm the **Spreadsheet ID**, **Preferred Tab** (tab name), and **Read Range** (`A:Z` is a safe default).  
- Make sure you (and the Google account you’re signed in with) have at least **Viewer** access to the Sheet; **Editor** access is needed to write `sent_status/sent_at`.

**Gmail “send as” / alias isn’t honored**  
- Gmail must be configured to **Send mail as** that alias in your account settings. Ask IT if needed.

**Messages appear blank**  
- `template.txt` is required. If you only use HTML, the app will auto-create a plain-text fallback by stripping HTML tags.

**Re-sending to the same candidates**  
- The app won’t resend rows with `sent_status == sent`.  
- If you want to re-send, clear that cell for those rows.

**Org blocks the app**  
- Your Workspace admin may need to approve the OAuth scopes or add the OAuth client to an allowlist.

---

## Safety, privacy & compliance

- **Least access**: The app only requests access to send email and read/write the specific spreadsheet.
- **Token storage**: The token file is local to your machine (either next to the app or under `~/.rejections_gui/token.json`). Treat it like a password.
- **Audit trail**: Each sent row is stamped with `sent_status` + `sent_at`.
- **Review content**: Use **Dry Run** and **Test Send To Me** to double-check tone and accuracy.
- **Respect candidate privacy**: Keep attachments appropriate and avoid sensitive info unless required and approved.

---

## FAQ

**Q: Can multiple recruiters use this?**  
Yes. Each recruiter can keep a copy of the folder on their machine with their own `credentials.json` and `token.json`, or share the same Sheet with appropriate access.

**Q: Does it support rich HTML?**  
Yes—pick an HTML template. The app sends a multipart email containing both plain-text and HTML, so all email clients render something clean.

**Q: Will the progress bar show total items?**  
Yes. It reflects the position within the eligible list (after filtering by `send/skip/sent_status`).

**Q: Can I schedule or automate?**  
This app is designed for manual runs. If you need a scheduled job, ask your tech partner to wrap `rejections_core.run_sender` in a script or CI job.

---

## Uninstall / cleanup

- Delete the folder with the two `.py` files and `rejections_gui_settings.json`.
- Optionally remove `~/.rejections_gui/token.json` if it was created.

---

## Need a hand?

If you run into anything confusing, share:
- A screenshot of your **Config** tab,
- The last 20 lines from the **Run** tab log,
- The header row of your Google Sheet.

We’ll help you get it sorted quickly.
