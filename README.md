# Project Monitoring Dashboard

Java-powered data pipeline with a GitHub Pages dashboard for monitoring project progress.

## What this includes

- `src/main/java/.../ExcelToJsonExporter.java` reads `Projects.xlsx` and exports JSON.
- `docs/` contains the static dashboard site for GitHub Pages.
- `docs/assets/logo.png` is used as the dashboard cover/logo.

## Expected Excel format

Use the first row as headers. Include at least:

- a project name column (e.g., `Project Name`)
- a progress column (e.g., `Progress` or `Completion`) with values from `0` to `100`

Any additional columns are shown in the table.

## Generate dashboard data from Excel

1. Put `Projects.xlsx` in the repository root.
2. Run:

```bash
mvn compile exec:java
```

This writes to `docs/data/projects.json`.

## Preview locally

You can use any static server from the repo root:

```bash
python -m http.server 8080
```

Then open `http://localhost:8080/docs/`.

## Publish on GitHub Pages

1. Push this repository to `master`.
2. In GitHub: **Settings -> Pages**
3. Set:
   - **Source**: Deploy from a branch
   - **Branch**: `master`
   - **Folder**: `/docs`
4. Save.

Published URL will be:

`https://vbos-dashboards.github.io/Project_monitoring_dashboard/`
