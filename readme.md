# Invoice Processor

Automated time extraction from Microsoft PPM and Google Calendar.

---

# Export Module Setup

## Microsoft Project Portfolio Management (PPM)
* Will be prompted for user, password, and url. Asked if you'd like to save (writes to credentials on ppm.db plaintext)
* URL example `https://ORGNAME.sharepoint.com/sites/pwa/Timesheet.aspx`
* Uses Chrome driver in Selenium

## Google Calendar (GCal)
* Follow steps at [developers.google.com](https://developers.google.com/workspace/guides/get-started) to get your API credentials setup on the Google Cloud console.
* Name your Google OAuth 2.0 secret.json file `gauth-credentials.json`
* Special feature - use a second calendar in Google to track different invoicing projects as all-day events when PPM prompts for invoicing you can put `*GCAL` which will evenly divide the PPM projects total hours among the all-day Google calendar events.

---

# Usage

* (once) `pip install -r requirements.txt`
* Run `pyinv.py`
* Follow prompts

# Building

* (once) `pip install -r requirements.txt`
* (once) `pip install pylint, pipreqs`
* In vscode `Python: Run Linting` or `pylint '.\Invoice Processor'` (pylint 10.00/10)
* `pipreqs --force` to create new requirements.txt

---

# Todo

* Remove code repetition in `pyinv.py`
* Dynamically load modules in `/export_modules/*`
* Better CLI/GUI than `print` and `input`
* `.envs` for things like Debug flags. Probably move credentials into here.