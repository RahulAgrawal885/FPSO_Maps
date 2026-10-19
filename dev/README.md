# Developer helpers

Nothing in this folder is used by the live map; it's here for testing and working on `index.html` locally.

| File | What it's for |
|---|---|
| `serve.py` | Runs the map on your computer with caching turned off, so edits show on reload. From the repo root: `python3 dev/serve.py` then open http://127.0.0.1:8766/ (pass another port as an argument if 8766 is busy). |
| `FPSO_test_upload_no_TAM_SAM_SOM.xlsx` | Test file to upload: 15 projects from `dataset2.xlsx` (same OED IDs) plus 5 made-up "TEST …" projects, with no TAM/SAM/SOM columns. |
| `make_test_file.py` | Rebuilds the test file: `python3 dev/make_test_file.py dataset2.xlsx dev/FPSO_test_upload_no_TAM_SAM_SOM.xlsx` (needs `pip install openpyxl`). |

## What the test file exercises

- **Upload check:** it has the two required columns (Platform Name, Location), so it loads without the popup.
- **No funnel columns:** all 20 projects start as Out; Export adds `Updated: TAM / SAM / SOM`.
- **Compare with Map 2 (dataset2.xlsx):** 5 new, 8 dropped, plus these edits on matched projects:

| Project | Change from dataset2 |
|---|---|
| Nam Du FPSO | Hull award 2026 → 2027 |
| Bonga Southwest - Aparo FPSO | Installation 2031 → 2033 |
| Fangtooth FPSO (FPSO-9) | Hull award 2027 → 2026 |
| Bay du Nord FPSO | Status FEED → Bidding |
| Abadi FPSO | EPCI capex 2550 → 2700 |

Until you set stages, every matched project also shows "Stage moved" (Map 2 has them in SAM/SOM, the test file starts them at Out).
