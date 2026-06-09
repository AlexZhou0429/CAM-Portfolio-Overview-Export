# CAM Portfolio Overview Export

Exports CAM portfolio positions to Excel using the CAM Open API.

The program:

- selects portfolios tagged `SP Core`
- downloads Asset & Position data
- sorts positions by USD Exposure
- creates one sheet per portfolio with the top 10 and bottom 10 exposures

## Setup

Install Python 3, open a terminal in this folder, then run:

```bash
python -m pip install -r requirements.txt
```

## Run

```bash
python export_portfolio_overview.py
```

Enter the CAM API Key and Secret when prompted. They are hidden and are not
saved to disk.

Two files are created in `outputs/`:

- `portfolio_overview_*.xlsx`: complete API data
- `portfolio_managers_top_bottom_*.xlsx`: one formatted sheet per portfolio

Optional date and hour:

```bash
python export_portfolio_overview.py --date 2026-06-09 --hour 17
```
