# J-MAAP Streamlit Cloud Deployment Checklist

This v1.2 package fixes the `ModuleNotFoundError` seen at `from src.io_utils ...` by removing the fragile `src` package dependency. All local Python modules now sit beside `app.py` at the repository root.

## GitHub repository root must contain

- `app.py`
- `io_utils.py`
- `dq_matching.py`
- `pipeline.py`
- `analytics.py`
- `outputs.py`
- `monitoring_automation_v1.py`
- `requirements.txt`
- `config/`
- `.streamlit/`

Do **not** upload only `app.py`, and do **not** keep the ZIP as a single file in GitHub. Extract the ZIP and upload/commit all contents.

## Streamlit Community Cloud

1. Push all files above to the root of the GitHub repository (for example `j-maap`).
2. In Streamlit Cloud, choose the repository and set **Main file path** to `app.py`.
3. Use Python 3.11 or 3.12 if the deployment settings allow selecting a version.
4. Reboot/redeploy the app after replacing the old files.
5. If an old build remains cached, use **Manage app → Reboot app**.

## Expected startup

The app should load before any MoDa file is uploaded. If it fails before the upload screen, inspect the Streamlit logs; the first unresolved import/package name identifies the deployment problem.

## v1.6 map files

In addition to the existing root files, make sure GitHub contains:

```text
mapping.py
assets/
  maps/
    somali_zones_2024.geojson
    somali_woredas_2024.geojson
    zone_labels.json
    woreda_labels.json
    somali_bounds.json
```

If `assets/maps/` is missing, the app will run but the **Visited Sites Map** page will fail when it tries to load the administrative layer.

No additional Python GIS package is required in Streamlit Cloud: the shapefile has already been converted to GeoJSON. Street and imagery layers are browser-served web tiles.

## v1.7 photo-evidence update

Replace/commit these updated root files:

```text
app.py
pipeline.py
mapping.py
io_utils.py
```

No additional Python package is required for photo previews. Image files are **not** stored in GitHub or Streamlit; J-MAAP reads only the attachment URLs contained in the uploaded MoDa export.

After deployment, upload MoDa data and run the monthly analysis before opening **Visited Sites Map**. A site with photo evidence should show `📷` in its tooltip and thumbnails in the popup when the current browser can access the MoDa image URL.

## v1.8 dashboard UI update

In addition to all v1.7 files, the repository root must now contain:

```text
ui_components.py
```

For an existing v1.7 deployment, replace/commit at minimum:

```text
app.py
mapping.py
ui_components.py
README.md
DEPLOYMENT_CHECKLIST.md
```

Keep the existing `assets/maps/`, `config/`, `.streamlit/` and other root Python modules. Do not upload `__pycache__/` folders.

After the commit:
1. Open Streamlit Community Cloud.
2. Use **Manage app → Reboot app**.
3. Upload the two August reference exports if you want to run the regression acceptance check.
4. Run August 2026 / Jijiga analysis.
5. Confirm 1,817 accepted submissions, 414 detailed findings and 11 aggregated actions.
6. Open **Visited Sites Map** and click a site marker. The right-hand Site Details panel should update inside the map component.
7. Open **Findings & Actions** and verify the dashboard cards, charts, action table and right-side priority panels.

The v1.8 UI relies on external browser access for OpenStreetMap, Esri imagery, Leaflet CDN resources and any MoDa images. The locally packaged Somali Region administrative GeoJSON continues to work without a Python GIS dependency.
