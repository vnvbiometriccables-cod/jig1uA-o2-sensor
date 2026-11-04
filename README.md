# O₂ SENSOR IN ENVIRONMENTAL CHAMBER
### Live Sample Maximum / Minimum Monitoring

## 1. Project overview

A live Streamlit dashboard for five O₂ sensor samples in the environmental chamber:
Sample 36, Sample 28, Sample 27, Sample 30 and Sample 32. Readings are in **µA**
and are never converted.

The dashboard reads the **O2_Data** tab of the Google Sheet every 60 seconds and:

* calculates each sample's **actual maximum, minimum, difference and average** from
  every valid numeric reading, including readings outside the limits;
* finds the **overall highest and lowest** reading, with the sample and timestamp;
* checks every reading against the **threshold band**:

  | Reading | Status | Colour |
  |---|---|---|
  | below 50 µA | BELOW LIMIT | red |
  | 50–146 µA (both limits included) | WITHIN RANGE | green |
  | above 146 µA | ABOVE LIMIT | orange |
  | missing or invalid | NO DATA | grey |

* draws live trend graphs with 50 µA and 146 µA reference lines;
* produces a **five-sheet Excel report** using the same calculations.

| Item | Value |
|---|---|
| Google Sheet | <https://docs.google.com/spreadsheets/d/1fJjKezE5Dujuc_QTmtwefe6z2h8ff-E3QuOO49Yjd1g/edit?gid=1371427582> |
| Sheet ID / gid | `1fJjKezE5Dujuc_QTmtwefe6z2h8ff-E3QuOO49Yjd1g` / `1371427582` (tab O2_Data) |
| Data method | Public CSV export: `…/export?format=csv&gid=1371427582` |
| Unit | µA |
| Default limits | 50 µA (min) – 146 µA (max), inclusive, adjustable in the app |
| Refresh | Every 60 s (can be switched off), plus **REFRESH NOW** |

This is a separate app from the Jig 1 and Jig 2 dashboards. It has its own repo,
its own Streamlit app and its own indigo colour theme.

### Project files

```
O2_Sensor_Dashboard/
├── app.py                        # the whole application
├── requirements.txt
├── README.md
├── .gitignore
└── .streamlit/
    ├── config.toml               # theme / server settings
    └── secrets.toml.example      # OPTIONAL – only for the private-sheet method
```

## 2. Install Python

Install **Python 3.11 or later** from <https://www.python.org/downloads/>.
On Windows, tick **"Add python.exe to PATH"** in the installer.

To check it worked, open a terminal (Command Prompt on Windows, Terminal on Mac)
and run `python --version`.

## 3. Project setup

Unzip `O2_Sensor_Dashboard.zip`. Then open a terminal **inside** the
`O2_Sensor_Dashboard` folder:

* **Windows:** open the folder in File Explorer, click the address bar, type
  `cmd`, and press Enter.
* **Mac:** in Terminal, type `cd ` (with a space), drag the folder into the
  window, and press Enter.

Optionally, create a virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows
source .venv/bin/activate       # Mac / Linux
```

## 4. Install the packages

```bash
pip install -r requirements.txt
```

## 5. Google Sheet sharing

The simple CSV method needs the sheet to be shared as **"Anyone with the link –
Viewer"**. On 4 Nov 2026 the sheet already had this setting, so **no login or
secrets are needed**.

To check or set it: open the sheet → **Share** → **General access** →
**Anyone with the link** → **Viewer**.

> Anyone who has the link can view the sheet. If the data must stay private,
> use the service-account method below instead.

**Private sheet option (service account).** If public access isn't allowed:

1. Share the sheet with your service account's e-mail as **Viewer**. You can use
   the one already set up for the Jig dashboards,
   `jig2-dashboard@jig2-dashboard.iam.gserviceaccount.com`.
2. Fill in `.streamlit/secrets.toml.example` with the real key and set
   `data_source = "api"`.
3. Paste it into **Streamlit Cloud → App → Settings → Secrets**. For local runs,
   save it as `.streamlit/secrets.toml` instead.

Credentials are never written in the source code.

## 6. Run locally

```bash
streamlit run app.py
```

Your browser opens at <http://localhost:8501>. Leave the terminal open while
you use the dashboard.

## 7. Selecting Google Sheet columns

Open **Column mapping** in the sidebar. The app detects the columns automatically:

* **Timestamp:** `DATE AND TIME`
* **Sample 36 / 28 / 27 / 30 / 32:** the `SAMPLE 36 (uA)` … `SAMPLE 32 (uA)`
  columns

If a header changes, pick the right column from its dropdown.
**Reset to detected columns** restores the automatic choice. The other columns in
the sheet (S6, S7, S8) are ignored. If a sample isn't mapped, it shows **NO DATA**
and a warning appears.

## 8. Changing the threshold values

Open **Threshold configuration** in the sidebar and change **Minimum limit** or
**Maximum limit**. **Restore 50 – 146 µA** puts the defaults back.

* Statuses, compliance counts, graph reference lines and the Excel report all
  update at once.
* The readings and the maximum/minimum values never change, because thresholds
  don't filter them.
* Changes apply to your browser session. To change the permanent defaults, edit
  `DEFAULT_MIN` and `DEFAULT_MAX` near the top of `app.py`.

## 9. How automatic refresh works

* While **Auto-refresh** is on, the page reloads its data every 60 seconds as
  long as the browser tab is open. Turn it off to pause; **REFRESH NOW** always
  fetches straight away.
* The sheet is downloaded at most once every 50 seconds and shared by everyone
  viewing the app. This keeps Google requests low.
* The status bar shows several things:
  * **LIVE**, **CONNECTION ERROR** or **NO DATA**
  * the last successful refresh and the latest data timestamp
  * the number of new rows since the last refresh

**If a refresh fails**, nothing is cleared. The last successfully loaded data and
its maximum/minimum values stay on screen with a red **CONNECTION ERROR**
message, and the app tries again on the next cycle. LIVE is never shown after a
failed refresh. No readings are ever invented.

## 10. Downloading the Excel report

Click **⬇️ DOWNLOAD EXCEL REPORT** in the sidebar. You get
`O2_Sensor_Environmental_Chamber_Monitoring_Report.xlsx` with five sheets:

| Sheet | Contents |
|---|---|
| Maximum Minimum Summary | Per sample: maximum and its timestamp, minimum and its timestamp, difference, average, total readings |
| Overall Summary | Title, generation time, unit, limits, overall highest and lowest (with sample and timestamp), sample count, valid readings |
| Threshold Validation Summary | Per sample and in total: valid readings, within range, below limit, above limit, compliance % |
| Raw Readings | Every row: timestamp plus the five original sample values. Nothing removed, out-of-range values kept. |
| Sample-wise Analysis | One row per reading: timestamp, sample, output (µA), threshold status (colour-coded) |

The report has headers, borders on the summary sheets, frozen header rows,
filters, number formatting, and conditional formatting on the status column.

**Size note.** The sheet holds about 180,000 rows (about 900,000 sample
readings), so the report takes about a minute to build after you click and is
about 30 MB. Excel allows up to 1,048,576 rows per sheet. When **Sample-wise
Analysis** reaches that limit (in about two days at the current logging rate),
the rows continue on **Sample-wise Analysis (2)**, **(3)** and so on. No readings
are dropped.

## 11. Troubleshooting Google Sheets connections

| Message | Cause and fix |
|---|---|
| *Google returned a sign-in page instead of CSV* | The sheet is no longer public. Share it as "Anyone with the link – Viewer", or use the service-account method (section 5). |
| *Google Sheets refused access* (HTTP 401/403) | Same as above. |
| *Sheet or tab not found* (HTTP 404) | Check the sheet ID and the gid (`1371427582` = O2_Data tab). |
| *did not respond in time* | Google was slow or the network dropped. The app keeps the last data and retries automatically. |
| *No connection to Google Sheets* | Internet is down where the app runs. |
| *Not mapped to a sheet column* | A header changed. Fix it under **Column mapping**. |
| Dates look a month off | Add `[app_settings]` with `timestamp_order = "monthfirst"` (or `"dayfirst"`) to the secrets. |
| Max of 255.9922 µA or negative minimum | These are real values in the sheet. 255.9922 is the logger's top of scale, and negative values appear during sensor or connection faults. They're included because the specification says actual max/min must include every reading. Find them with the filters in the Sample-wise Analysis sheet. |

---

## 12. Free deployment (Streamlit Community Cloud)

1. **GitHub account.** Sign up at <https://github.com> if you don't have one.
   You can reuse `vnvbiometriccables-cod`.
2. **Create a repository.** Go to <https://github.com/new>. Name it, for example,
   `o2-sensor-max-min-dashboard`, choose **Private**, and click **Create
   repository**. Don't add a README. Use a new repo so the Jig dashboards are
   never touched.
3. **Upload the files.** Click **uploading an existing file** and drag in
   `app.py`, `requirements.txt`, `README.md`, `.gitignore` and the `.streamlit`
   folder. Then click **Commit changes**.
   * Check that `.streamlit/config.toml` appears inside a folder.
   * On Windows, turn on View → Show → Hidden items to see the dot-files.
4. **Connect GitHub to Streamlit.** Sign in at <https://share.streamlit.io> with
   GitHub. If Streamlit can't see the new repo, go to GitHub → **Settings →
   Applications → Installed GitHub Apps → Streamlit → Configure** and give it
   access to that repo.
5. **Deploy.** Click **Create app** → **Deploy a public app from GitHub**. Click
   **Paste GitHub URL** and paste
   `https://github.com/<your-user>/o2-sensor-max-min-dashboard/blob/main/app.py`.
   In **Advanced settings**, choose Python **3.11**, then click **Deploy**.
6. **Secrets.** Not needed for the public-link method. Use them only for the
   private-sheet method (section 5).
7. **Open the dashboard** at the `https://….streamlit.app` address Streamlit
   shows. Under the app's **Settings → Sharing**, you can limit who can open it.

### Free-tier limits that affect continuous monitoring

* **Sleeping.** An app with no visitors for a while (about 12 hours or more) is
  put to sleep. The next visitor clicks **"Yes, get this app back up"** and it
  starts in about 30 seconds. Nothing is lost, because all data comes from the
  sheet.
* **Refresh needs an open tab.** Auto-refresh runs in the browser. With no
  dashboard open, nothing is fetched. That's fine here, because the logger keeps
  writing to the sheet regardless.
* **Memory (about 1 GB).** Plenty for the dashboard. The full Excel report is
  the heaviest operation, so generate it when needed rather than leaving many
  report downloads running at once.
* **One app per repo or entry point.** Keep this dashboard in its own repo, as
  described above.
