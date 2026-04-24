# Weekly Report Automation

The app currently supports three main workflows:

- upload an `.xlsx` workbook or an `Intervention Log .csv` file through the interface
- open the dashboard without a workbook and use the `Flat Time` tab with `.csv` files only
- use automatic `Flat Time` activity translation on hover through `Aramco Activity Codes.csv` stored in the app root

## How It Works

- the user starts the local server
- uploads an `.xlsx` file or an `Intervention Log .csv` file through the interface, or opens the dashboard without a workbook
- the `Interactive Dashboard` and `Weekly Report` are generated automatically
- uploaded files are processed in memory and are not stored on the server
- the `Flat Time` tab can work entirely from `.csv` files, even with no workbook loaded

## Run The App

From the project folder:

```bash
python3 server.py
```

Then open:

```text
http://127.0.0.1:8000
```

## Coolify Deployment

The deployment files are already included:

- `Dockerfile`
- `requirements.txt`
- `.dockerignore`

Recommended Coolify configuration:

1. Select the `Weekly-Report` repository
2. Use the root `Dockerfile`
3. Keep port `8000`
4. Use this start command:

```text
python server.py
```

The server already supports `HOST` and `PORT` environment variables, so it works well in containers.

## Security Hardening

For more restrictive corporate environments, the app supports:

- optional Basic Authentication
- optional IP or network allowlisting
- HTTP security headers
- in-memory `.xlsx` processing only

Useful environment variables for Coolify:

```text
HOST=0.0.0.0
PORT=8000
BASIC_AUTH_USER=your_user
BASIC_AUTH_PASSWORD=your_strong_password
ALLOWED_IPS=187.77.250.164,10.0.0.0/8,192.168.0.0/16
```

Notes:

- `BASIC_AUTH_USER` and `BASIC_AUTH_PASSWORD` protect the entire site
- `ALLOWED_IPS` accepts individual IPs and CIDR ranges separated by commas
- `/health` remains open for deployment health checks

## Usage Flow

1. Click `Upload And Open Report` and choose an `.xlsx` file or `Intervention Log .csv`
2. Or use `Open Dashboard Without Workbook` to enter the app directly
3. Without a workbook, the `Flat Time` tab still works with `.csv` uploads
4. Use the `Interactive Dashboard`, `Weekly Report`, and `Flat Time` tabs

## Flat Time Activity Translation

- if `Aramco Activity Codes.csv` is present in the app root, the dashboard uses it automatically
- descriptions appear on hover over activity names in the `Flat Time` tab
- in deployment, keep this file in the project so the app does not depend on manual upload

## Main Files

- `server.py`: starts the server, receives uploads, and opens the dashboard
- `generate_report.py`: generates the HTML dashboard
- `Dockerfile`: deployment image for Coolify
- `requirements.txt`: Python dependencies
- `report_output/report.html`: latest generated HTML output

## Important Notes

- the app no longer requires the spreadsheet to be stored inside the project
- uploaded files are processed in memory only
- no `.xlsx` file is bundled on the server
- in corporate environments, `.csv` usually passes more easily because it is plain text, while `.xlsx` is a compressed Office package and is often inspected by antivirus, DLP, upload filters, and anti-malware policies
- in `CSV` mode, the system assumes the file represents the `Intervention Log` sheet
- in `CSV` mode, the system builds the dashboard from the available intervention data only
