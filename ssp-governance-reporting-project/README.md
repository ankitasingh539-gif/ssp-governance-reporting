# SSP Governance report builder prototype

## Purpose

This is a local report-making prototype. Its main view is a Board report preview, not a consolidated source workbook. It reads Area Manager submissions, filters by reporting month and Area, calculates initial headline totals from the named template fields, and downloads a Word report draft. The prototype is designed to be refined with service owners before formal circulation.

## Current workflow

1. In the app, use **Select a folder** to choose the local OneDrive-synced Area Governance folder. Excel reports inside it and its subfolders are uploaded to the app session.
2. Review submission cover details and validation items on **Data sources**.
3. Select **All areas** or an individual Area and choose a reporting month.
4. Review the headline measures, Area comparison, and report sections.
5. Add optional report focus text and download the Word report draft.
6. Optionally map an HR workbook or CSV by Area, employee identifier, absence hours, and reporting month. The report uses only aggregates from mapped HR columns; raw HR rows are not displayed.

## Current calculation boundaries

- People supported, admissions, discharges, and voids are summed from the `Service figures` section.
- Referrals listed count non-empty referrer records; they are not treated as admissions.
- Absence hours (all recorded absence types) and agency hours are summed from their respective Area Manager sections.
- Incidents and accidents are summed from their two source columns.
- Complaints/compliments and audit records are counts of populated source records.
- Risks marked `Open` are counted from the Risk section's Status field.
- Occupancy rate is not calculated because the template does not provide an agreed capacity denominator.
- Quarter comparison wording remains a placeholder until the required historical period and comparison method are agreed.
- Compliance rates and conclusions are not inferred from raw check fields.

## Generated Board report

The Word report follows the supplied Architecture structure: cover details, executive summary, occupancy and referrals, quarter comparison, improvements and risks, compliance and monitoring, complaints and compliments, package changes, staffing, and absence. It uses compact tables with one row per Area wherever possible. Absence totals include every recorded type in the Area Manager absence register. If HR absence hours are mapped, they appear separately to avoid accidental double counting.

The current Area Manager workbook does not provide service comments or issues from the previous report, so those fields are marked as unavailable. The quarter comparison remains uncalculated until prior-quarter files and a comparison method are confirmed. Service-level source details remain available in the app and original workbooks.

The input template has multi-section sheets and no formal Excel tables. The app uses observed row and column positions. Template changes may require updating `TEMPLATE` in `app.py`. `Report_Data` and `Lists` are not treated as source records.

## HR data

The HR file schema has not yet been provided. On the **Data sources** page, map the HR Area field, an employee identifier for distinct headcount, an optional absence-hours field, and (recommended) a reporting-month field. If the HR file is a single-month snapshot, confirm that in the app before using it. Confirm that HR and Area Manager absence measures do not overlap before combining them in an approved report.

The current app processes data in the local Streamlit session. It does not retain an HR database or connect to Microsoft Graph. Team-wide production use needs organisation sign-in, role-based access, an approved hosting location, retention rules, and a privacy review for staffing and absence information.

## Run locally

Requires Python 3.11 or newer and package access for `requirements.txt`.

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

The theme settings in `.streamlit/config.toml` use a light background and blue accent. The user selects the folder through the browser. The prototype does not connect directly to SharePoint or OneDrive online; it uploads selected workbooks into the local app session for processing.

## Run with Docker on Windows

Docker packages the Python runtime and app dependencies in a container, so Python does not need to be installed separately. Docker Desktop must be installed and running; on Windows it uses WSL 2 for Linux containers. Check with IT whether Docker Desktop and WSL 2 are permitted on your work device. See the [official Docker Desktop installation guide](https://docs.docker.com/desktop/setup/install/windows-install/).

1. Open PowerShell in the project folder and run:

   ```powershell
   docker compose up --build
   ```

2. Open `http://localhost:8501` in a browser, choose the synced Area Governance folder, and select **Create dashboard**.
3. Stop the app with `Ctrl+C` in PowerShell. To start it again later, run `docker compose up` from the same folder.

This Docker setup is for running on one computer with Docker Desktop. The selected workbook files are sent from the browser to the local container for the app session; it does not read SharePoint online, and SharePoint does not run Docker containers. For team access, run the container on an approved server/cloud host, add Microsoft Entra sign-in, and use an approved SharePoint/Graph connection.

## GitHub contents

Commit the app code, dependency list, README, `.gitignore`, `.dockerignore`, `Dockerfile`, `compose.yaml`, and `.streamlit/config.toml`. Do not commit Area Manager or HR workbooks, contact lists, generated reports, local logs, `.env`, or credentials. The ignore files exclude common data and local runtime files by default.

## Production direction

For a team tool, retain the same report workflow but deploy to an organisation-approved internal service, authenticate through Microsoft Entra ID, read SharePoint files through Microsoft Graph with least privilege, and restrict views to approved SSP Governance users. Store outputs in a separate permission-controlled library. Do not embed app secrets in code. A tenant administrator must approve the app registration, authentication model, SharePoint scope, and hosting method.
