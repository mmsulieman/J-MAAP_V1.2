# J-MAAP
## Jijiga Monitoring, Analytics & Action Platform

J-MAAP is a Streamlit application for monthly monitoring data intake, data-quality review, cross-file reconciliation, raw-data exploration, findings analysis, action tracking, and generation of management outputs.

### Local run

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

### Streamlit Cloud

Upload/commit **all extracted files and folders** to the repository root. The main file path is:

```text
app.py
```

See `DEPLOYMENT_CHECKLIST.md` for the exact required repository structure and troubleshooting steps.

### Important v1.2 fix

Earlier builds imported local modules through `src.*`. If the `src` folder was not committed to the deployed repository, Streamlit raised `ModuleNotFoundError` before the app could load. v1.2 uses flat root-level local modules to make deployment more robust.

## Version 1.5 cohort-safety update
- Detects actual monitoring months available in uploaded MoDa files by sub-office before analysis.
- Blocks zero-record analysis runs and shows which reporting months are actually available.
- Shows eligible record count and first/last monitoring dates for the selected cohort.
- Warns when a completed month appears only partially covered by the uploaded export(s).
- Supports combining multiple MoDa form versions/exports; UUID deduplication prevents double counting.
- Sub-office filtering is case/whitespace tolerant.


## v1.5 update — Sub-office filter

- Sub-office is now a dropdown detected directly from uploaded MoDa data.
- Jijiga, Gode and Dollo Addo are prioritised at the top when present.
- The selected Sub-office controls cohort validation, monthly analysis, MoDa raw preview and generated report/output filenames.
- Uploading a new MoDa source refreshes the dropdown automatically.

## v1.6 update — Visited Sites Map

J-MAAP now includes a dedicated **Visited Sites Map** page driven by the accepted monthly MoDa cohort and its GPS fields.

Map views:
- **Street** — OpenStreetMap basemap.
- **Satellite imagery** — Esri World Imagery basemap.
- **Administrative hollow** — locally bundled Somali Region 2024 zone/woreda boundaries with no polygon fill.

Map functions:
- Uses the global Reporting Month and Sub-office filters.
- Additional Activity, Zone, Woreda and Provider filters.
- Point-display modes: Activity, Visit frequency, Provider and Finding severity.
- Zone labels on by default; woreda labels can be toggled on.
- Popups show site, geography, activity, provider, monitoring events, submissions, latest visit and finding severity.
- Downloadable mapped-site register.
- Median site coordinate is used where several MoDa submissions at the same site contain slightly different handheld-GPS readings.

### Somali Region administrative source
The supplied 2024 shapefile was validated before conversion:
- 11 zones and 101 woredas.
- Zone and woreda polygon layers are EPSG:4326 and align with MoDa GPS coordinates.
- The app bundles converted GeoJSON boundary files under `assets/maps/`; no shapefile runtime dependency is required.
- Woreda label points use the geometry of the supplied `Centroid.shp`, because the polygon attribute columns labelled Latitude/Longitude are not reliable coordinate labels in the source.
- `SRS_Settlements_2024.shp` is **not** used by v1.6; its projected coordinate metadata/extent requires separate validation before it should be used for automated site geocoding.

Street and satellite basemaps require internet access in the user's browser. Administrative boundary data are packaged locally with J-MAAP.

## v1.7 update — MoDa photo evidence on the map

J-MAAP now detects MoDa attachment URLs and links them to the accepted monthly cohort at **site → monitoring event → submission** level.

Photo functions:
- A camera symbol (`📷`) appears in the site tooltip when photo evidence exists.
- Clicking a site can show up to four recent photo thumbnails directly in the Leaflet popup.
- The map page includes a site-level photo gallery with up to 12 recent images and a complete photo-evidence register.
- Each photo record retains monitoring date, activity, provider, monitoring event key, source question, source worksheet/file and submission UUID.
- The map register includes photo count and photo types for each monitored site.
- Main-sheet photo fields and the MoDa operational-picture repeat sheet are supported.
- MoDa attachment links exported as `http://api.moda.wfp.org/...` are upgraded to HTTPS for display inside the HTTPS Streamlit application.
- If an image cannot be embedded because the browser is not authenticated to MoDa, J-MAAP keeps an **Open full image** link rather than failing the map.

J-MAAP does **not** download, cache or republish the underlying images. Photos remain hosted by MoDa and normal MoDa/WFP access controls continue to apply.

August 2026 regression using the two Jijiga MoDa exports:
- 1,817 accepted submissions
- 247 detected photo links
- 146 sites with photo evidence
- 414 detailed findings
- 11 aggregated actions

## v1.8 update — Dashboard UI rebuild

Version 1.8 rebuilds the real Streamlit interface around the approved J-MAAP dashboard design rather than relying on default Streamlit page styling.

### Findings & Actions
- Horizontal management filters for reporting month, sub-office, activity, provider, zone, woreda and monitoring type.
- KPI cards for detailed findings, aggregated issues, high-priority findings, open/overdue actions and closures.
- Dashboard tabs: Overview, Findings Analysis, Action Tracker, Issues by Woreda and Trends.
- Findings-by-severity, findings-by-theme, action-status and woreda × theme visuals.
- Styled action register with severity/status/validation badges.
- Right-hand priority-action and recurring-issue panels.
- All numbers remain dynamic from the selected uploaded MoDa cohort; the reference screenshots are layout specifications only.

### Visited Sites Map
- Dashboard KPI strip and horizontal operational filters.
- The map and Site Details rail are rendered inside one custom HTML/JavaScript Leaflet component.
- Clicking a site updates the right-hand panel without re-running Streamlit.
- Site panel shows activity, provider, monitoring events, submissions, latest visit, finding signals and photo evidence.
- Street, Esri satellite/imagery and Somali Region hollow administrative views are switchable inside the map.
- Latest image and thumbnail strip use the MoDa attachment links and fall back gracefully when MoDa authentication prevents embedding.

### Other UI changes
- Dark WFP-style navigation sidebar and white application header.
- Monitoring Overview rebuilt as a management dashboard.
- Dedicated Photo Evidence gallery.
- Reports & Exports rebuilt as an export center with data-readiness status.
- New root module: `ui_components.py`.

No additional Python package is required for the v1.8 UI layer; the enhanced map uses the same browser-loaded Leaflet library already used by J-MAAP.

### August 2026 regression benchmark
The core analytical regression remains unchanged:
- 1,817 accepted Jijiga submissions
- 22 configured indicators
- 414 detailed findings
- 11 aggregated actions
- 109 manual-month/date inconsistencies
- 247 photo links across 146 sites
