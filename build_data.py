"""
V1.2 data pipeline for the Snitch Control Room dashboard.

THREE ways to run this:

  A) Fully automatic, authenticated, IMPORTRANGE-safe (recommended -- set up once):
       python build_data.py
     Requires a Google Service Account (one-time ~10 min setup, see README).
     Save its JSON key as ./service_account.json next to this script, and share
     the Google Sheet with the service account's email (Viewer access).
     This goes through Google's live Sheets API, which correctly resolves
     IMPORTRANGE (and any other cross-sheet formula) because it's a real
     authenticated read, unlike a plain anonymous URL fetch.

  B) Automatic but UNSAFE for IMPORTRANGE-driven tabs:
       python build_data.py --export-url
     Downloads via the plain https://.../export?format=xlsx URL. Fine for
     Sales Data and Pipeline (plain formulas), but Inv Data 2 (IMPORTRANGE)
     will come back BLANK -- Google's export endpoint can't resolve
     IMPORTRANGE without an authenticated session. Kept only as a fallback.

  C) Offline / manual:
     python build_data.py --local
     Reads ./Automation_Data.xlsx if you've placed one there yourself.

Either way it writes data/data.json, which is the only thing index.html reads.

Sales now comes from Snowflake (SALES_FOR_AUTO_2, daily grain), not the "Sales
Data" Google Sheet tab. Requires these env vars (or a local
snowflake_credentials.json -- see snowflake_credentials.example.json):
  SNOWFLAKE_ACCOUNT, SNOWFLAKE_USER, SNOWFLAKE_PASSWORD,
  SNOWFLAKE_WAREHOUSE, SNOWFLAKE_DATABASE, SNOWFLAKE_SCHEMA
Inventory and Pipeline still come from Google Sheets as before.

Setup for mode A (one-time):
  pip install google-api-python-client google-auth --break-system-packages
  1. console.cloud.google.com -> new/existing project -> enable "Google Sheets API"
  2. IAM & Admin -> Service Accounts -> Create Service Account (any name, no roles needed)
  3. That service account -> Keys -> Add Key -> Create new key -> JSON -> downloads a file
  4. Save that file as service_account.json in this same folder
  5. Open the file, copy the "client_email" value
  6. Share the actual Google Sheet with that email, Viewer access
  Done -- python build_data.py now works with zero manual downloads, forever.
"""
import json, datetime, sys, os, csv
import openpyxl

# Rolling window for Snowflake-sourced data (Sales, Returns actual, Sales-by-SKU).
# Computed fresh every run from *today*, not a fixed date -- this is what keeps
# sales.json/returns_actual.json/sales_by_sku.json from growing bigger every
# single day forever. 400 days covers: current month, full prior-year MoM
# comparison, and any reasonable Custom Range pick in the Analytics tab, while
# old days automatically age out of the window as time moves forward.
ROLLING_WINDOW_DAYS = 400
ROLLING_CUTOFF_DATE = (datetime.date.today() - datetime.timedelta(days=ROLLING_WINDOW_DAYS)).isoformat()

SHEET_ID = "1PqtpL9w2Tneon_-6zz7BGiW5YUz4fhDCrgn687r_CZw"
HERE = os.path.dirname(os.path.abspath(__file__))
LOCAL_XLSX = os.path.join(HERE, "Automation_Data.xlsx")
SERVICE_ACCOUNT_FILE = os.path.join(HERE, "service_account.json")
TARGETS_DIR = os.path.join(HERE, "data", "targets")
OUT = os.path.join(HERE, "data", "data.json")
TAB_NAMES = ["Inv Data 2", "Pipeline"]  # Sales Data removed -- now sourced from Snowflake

# Exact filenames expected in data/targets/ -- one per category family, each with
# a "Sheet1" tab at store/branch grain with monthly qty columns (AUG_2026.. JUN_2027).
TARGET_FILES = [
    "Jeans Planned Qty only.xlsx",
    "Shirts Planned Qty Only.xlsx",
    "Trousers Planned Qty only.xlsx",
    "tshirts Planned Qty Only.xlsx",
]
MONTH_COL_MAP = {
    "AUG_2026": "2026-08", "SEP_2026": "2026-09", "OCT_2026": "2026-10",
    "NOV_2026": "2026-11", "DEC_2026": "2026-12", "JAN_2027": "2027-01",
    "FEB_2027": "2027-02", "MAR_2027": "2027-03", "APR_2027": "2027-04",
    "MAY_2027": "2027-05", "JUN_2027": "2027-06",
}
# normalize target's channel labels to match Sales' channel naming where they're
# clearly the same thing; Warehouse kept as its own distinct channel since Sales
# doesn't have an equivalent yet. MP-SOR merged into Marketplace per instruction.
TARGET_CHANNEL_MAP = {"Online - Shopify": "Shopify"}
# NOTE: MP-SOR deliberately kept distinct from Marketplace here (not merged),
# per instruction -- the channel-mix split treats them as separate percentages.
# The live dashboard's "Marketplace" filter still sums MP + MP-SOR together at
# query time (same as it already does for Sales), so nothing visible changes.

NEW_TOTAL_FILE = os.path.join(HERE, "data", "targets_new_total.xlsx")


# ---------------------------------------------------------------------------
# Mode A: Google Sheets API via service account (authenticated, IMPORTRANGE-safe)
# ---------------------------------------------------------------------------
GSHEET_EPOCH = datetime.datetime(1899, 12, 30)

def _gsheet_serial_to_date(value):
    """Sheets API returns dates as serial-day floats (same epoch as Excel).
    Converts back to a python datetime so downstream code (month_key etc.)
    doesn't need to know the difference between this and an openpyxl cell."""
    if isinstance(value, (int, float)):
        return GSHEET_EPOCH + datetime.timedelta(days=value)
    return value


class SimpleSheet:
    """Minimal stand-in for an openpyxl worksheet, backed by raw API rows,
    so the same parsing code in build() works for either data source."""
    def __init__(self, rows, width, date_cols=()):
        # pad every row out to `width` columns (Sheets API omits trailing blanks)
        self.rows = []
        for row in rows:
            padded = list(row) + [None] * (width - len(row))
            for c in date_cols:
                if isinstance(padded[c], str) and padded[c] == "":
                    padded[c] = None
                elif isinstance(padded[c], (int, float)):
                    padded[c] = _gsheet_serial_to_date(padded[c])
            self.rows.append(tuple(padded))

    def iter_rows(self, min_row=1, max_row=None, values_only=True):
        start = min_row - 1
        end = max_row if max_row else len(self.rows)
        for r in self.rows[start:end]:
            yield r

    @property
    def max_row(self):
        return len(self.rows)


class SimpleWorkbook:
    def __init__(self, sheets_dict):
        self._sheets = sheets_dict
        self.sheetnames = list(sheets_dict.keys())

    def __getitem__(self, name):
        return self._sheets[name]


def fetch_via_sheets_api(sheet_id):
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        print(f"ERROR: {SERVICE_ACCOUNT_FILE} not found.")
        print("See the module docstring (top of this file) for one-time setup steps.")
        sys.exit(1)

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    service = build("sheets", "v4", credentials=creds)
    values_api = service.spreadsheets().values()

    # date columns (0-indexed) per tab, so serials get converted correctly
    date_cols = {"Inv Data 2": (0,), "Pipeline": (0, 1, 20, 21)}
    widths = {"Inv Data 2": 17, "Pipeline": 28}

    sheets = {}
    for tab in TAB_NAMES:
        print(f"Fetching '{tab}' via Sheets API...")
        result = values_api.get(
            spreadsheetId=sheet_id, range=tab,
            valueRenderOption="UNFORMATTED_VALUE"
        ).execute()
        rows = result.get("values", [])
        sheets[tab] = SimpleSheet(rows, widths[tab], date_cols.get(tab, ()))
        print(f"  -> {len(rows)} rows (incl. header)")

    return SimpleWorkbook(sheets)


def download_from_gsheet(sheet_id, dest):
    import requests
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    print("Downloading workbook from Google Sheets (unauthenticated export)...")
    print("WARNING: this will NOT resolve IMPORTRANGE-driven tabs (e.g. Inv Data 2).")
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    with open(dest, "wb") as f:
        f.write(r.content)
    print(f"Saved -> {dest}")


# ---------------------------------------------------------------------------
# Snowflake: daily-grain Sales (SALES_FOR_AUTO_2), replaces the "Sales Data"
# Google Sheet tab. Credentials come from env vars so the same code works
# locally (via snowflake_credentials.json, loaded into env vars below) and
# in GitHub Actions (via repo secrets injected as env vars in the workflow).
# ---------------------------------------------------------------------------
SNOWFLAKE_CREDS_FILE = os.path.join(HERE, "snowflake_credentials.json")

SNOWFLAKE_QUERY = f"""
    SELECT DATE, CHANNEL, L1_CATEGORY, CATEGORY, META1, META2, META3,
           SUM(GROSS_SALES_VALUE) AS GROSS_SALES_VALUE, SUM(MRP_VALUE) AS MRP_VALUE,
           SUM(QTY) AS QTY, SUM(COGS_SOLD) AS COGS_SOLD
    FROM SNITCH_DB.MAPLEMONK.SALES_FOR_AUTO_3
    WHERE DATE >= '{ROLLING_CUTOFF_DATE}'
    GROUP BY DATE, CHANNEL, L1_CATEGORY, CATEGORY, META1, META2, META3
"""

# Separate, much smaller SKU-level dataset for the Pareto/Quartile tabs only --
# aggregated by MONTH (not by day), since those tabs only need period totals
# per sku_group, not daily granularity. Keeping this out of the main SNOWFLAKE_QUERY
# above is what keeps data.json from ballooning to 1GB+ (SALES_FOR_AUTO_3 is raw
# SKU-day grain; exploding that into the main sales array multiplies file size
# by roughly the average SKU count per leaf combo).
SALES_BY_SKU_QUERY = f"""
    SELECT
        TO_CHAR(DATE, 'YYYY-MM') AS MONTH, CHANNEL, SKU_GROUP,
        L1_CATEGORY, CATEGORY, META1, META2, META3,
        SUM(GROSS_SALES_VALUE) AS GROSS_SALES_VALUE, SUM(QTY) AS QTY,
        SUM(COGS_SOLD) AS COGS_SOLD, SUM(MRP_VALUE) AS MRP_VALUE
    FROM SNITCH_DB.MAPLEMONK.SALES_FOR_AUTO_3
    WHERE DATE >= '{ROLLING_CUTOFF_DATE}'
    GROUP BY TO_CHAR(DATE, 'YYYY-MM'), CHANNEL, SKU_GROUP, L1_CATEGORY, CATEGORY, META1, META2, META3
"""
# NOTE: this now uses the same rolling ROLLING_CUTOFF_DATE as the main sales
# query (computed fresh from today's date every run -- see top of file), not a
# fixed date. That's what keeps this file's size roughly constant over time
# instead of growing forever as more months of history accumulate.

# Returns at the same monthly SKU_GROUP grain as SALES_BY_SKU_QUERY, for the
# Pareto/Quartile "Return" and "Net Value" columns. Joined directly on raw
# SKU_GROUP (NOT the parent-cleaned version used in the main RETURNS_QUERY) --
# RETURNS_DATA and SALES_FOR_AUTO_3 both come from the same order-line SKU
# identifiers, so they should match exactly without needing the meta_mapping
# product-master join at all here. Only OVERALL_RETURNS_QTY exists (no return
# *value* field in the source table) -- return value in Rupee-metric terms is
# therefore an ESTIMATE (return_qty x that SKU-month's average per-unit rate),
# clearly labeled "Est." in the UI. Qty-based return counts are exact.
RETURNS_BY_SKU_QUERY = f"""
    SELECT
        TO_CHAR(DATE, 'YYYY-MM') AS MONTH, SKU_GROUP,
        SUM(OVERALL_RETURNS_QTY) AS RETURN_QTY
    FROM SNITCH_DB.MAPLEMONK.RETURNS_DATA
    WHERE DATE >= '{ROLLING_CUTOFF_DATE}'
    GROUP BY TO_CHAR(DATE, 'YYYY-MM'), SKU_GROUP
"""

# ---------------------------------------------------------------------------
# Snowflake: one representative image URL per sku_group, for the Pareto /
# Quartile tabs (thumbnail next to each ranked SKU group). Picks the first
# IMAGE-type media (by media index) per sku_group across all its variants.
# ---------------------------------------------------------------------------
IMAGES_QUERY = r"""
WITH images AS (
    SELECT
        UPPER(
    TRIM(
        CASE
            WHEN variant.value:sku::STRING ILIKE '4C%'
              OR variant.value:sku::STRING ILIKE 'MP%'
            THEN REGEXP_SUBSTR(variant.value:sku::STRING, '^[^-]+-[^-]+-[^-]+')
            ELSE REGEXP_SUBSTR(variant.value:sku::STRING, '^[^-]+-[^-]+')
        END
    )
) AS sku_group,
        image.value:preview:image:url::STRING AS image_url,
        image.index AS image_index
    FROM snitch_db.maplemonk.new_meafields_product_products_graph_ql t,
         LATERAL FLATTEN(input => PARSE_JSON(t.media)) AS image,
         LATERAL FLATTEN(input => PARSE_JSON(t.variants)) AS variant
    WHERE image.value:mediaContentType::STRING = 'IMAGE'
)
SELECT
    sku_group,
    image_url
FROM images
QUALIFY ROW_NUMBER() OVER (
    PARTITION BY sku_group
    ORDER BY image_index
) = 1
"""


def _load_snowflake_creds():
    """Local dev convenience: if snowflake_credentials.json exists, load its
    keys into os.environ (only if not already set). In GitHub Actions, the
    env vars are set directly by the workflow from repo secrets instead, so
    this file is never needed/present there."""
    if os.path.exists(SNOWFLAKE_CREDS_FILE):
        with open(SNOWFLAKE_CREDS_FILE) as f:
            creds = json.load(f)
        for k, v in creds.items():
            os.environ.setdefault(k, v)


def _snowflake_connect():
    import platform
    platform.libc_ver = lambda *a, **k: ('', '')  # Windows Store python.exe workaround
    import snowflake.connector

    _load_snowflake_creds()
    required = ["SNOWFLAKE_ACCOUNT", "SNOWFLAKE_USER", "SNOWFLAKE_PASSWORD",
                "SNOWFLAKE_WAREHOUSE", "SNOWFLAKE_DATABASE", "SNOWFLAKE_SCHEMA"]
    missing = [k for k in required if not os.environ.get(k)]
    if missing:
        print(f"ERROR: missing Snowflake env vars: {', '.join(missing)}")
        print("Locally: create snowflake_credentials.json (see README/chat). In GitHub")
        print("Actions: set these as repo secrets and pass them into the workflow env.")
        sys.exit(1)

    print("Connecting to Snowflake...")
    return snowflake.connector.connect(
        account=os.environ["SNOWFLAKE_ACCOUNT"],
        user=os.environ["SNOWFLAKE_USER"],
        password=os.environ["SNOWFLAKE_PASSWORD"],
        warehouse=os.environ["SNOWFLAKE_WAREHOUSE"],
        database=os.environ["SNOWFLAKE_DATABASE"],
        schema=os.environ["SNOWFLAKE_SCHEMA"],
    )


def fetch_sales_from_snowflake():
    conn = _snowflake_connect()
    try:
        cur = conn.cursor()
        print("Fetching sales from SALES_FOR_AUTO_3 (aggregated to leaf grain)...")
        cur.execute(SNOWFLAKE_QUERY)
        rows = cur.fetchall()
        cur.close()
    finally:
        conn.close()

    sales = []
    for row in rows:
        (dt, channel, l1, cat, m1, m2, m3, gross_sales, mrp_value, qty, cogs_sold) = row
        d = dt.strftime("%Y-%m-%d") if hasattr(dt, "strftime") else str(dt)[:10]
        sales.append({
            "date": d, "month": month_key(dt), "channel": channel, "l1": l1, "cat": cat,
            "m1": m1, "m2": m2, "m3": m3,
            "gross_sales": float(gross_sales or 0), "mrp_value": float(mrp_value or 0),
            "qty": float(qty or 0), "cogs_sold": float(cogs_sold or 0),
        })
    print(f"  -> {len(sales)} sales rows from Snowflake")
    return sales


def fetch_sales_by_sku_from_snowflake():
    """Separate, monthly-grain SKU-level dataset for Pareto/Quartile tabs only.
    Kept out of the main sales array (see SALES_BY_SKU_QUERY comment) to avoid
    the data.json size blowup."""
    conn = _snowflake_connect()
    try:
        cur = conn.cursor()
        print("Fetching sales_by_sku (monthly grain) from SALES_FOR_AUTO_3...")
        cur.execute(SALES_BY_SKU_QUERY)
        rows = cur.fetchall()
        cur.close()
    finally:
        conn.close()

    sales_by_sku = []
    for row in rows:
        (month, channel, sku_group, l1, cat, m1, m2, m3, gross_sales, qty, cogs_sold, mrp_value) = row
        sales_by_sku.append({
            "month": month, "channel": channel, "sku_group": sku_group,
            "l1": l1, "cat": cat, "m1": m1, "m2": m2, "m3": m3,
            "gross_sales": float(gross_sales or 0), "qty": float(qty or 0),
            "cogs_sold": float(cogs_sold or 0), "mrp_value": float(mrp_value or 0),
        })
    print(f"  -> {len(sales_by_sku)} sales_by_sku rows from Snowflake")
    return sales_by_sku


def fetch_returns_by_sku_from_snowflake():
    """Monthly-grain per-SKU returns, for the Pareto/Quartile 'Return' and 'Net
    Value' columns. See RETURNS_BY_SKU_QUERY comment for the value-estimation
    caveat (qty is exact, Rupee value is derived client-side as an estimate)."""
    conn = _snowflake_connect()
    try:
        cur = conn.cursor()
        print("Fetching returns_by_sku (monthly grain) from RETURNS_DATA...")
        cur.execute(RETURNS_BY_SKU_QUERY)
        rows = cur.fetchall()
        cur.close()
    finally:
        conn.close()

    returns_by_sku = []
    for row in rows:
        month, sku_group, return_qty = row
        if not sku_group:
            continue
        returns_by_sku.append({"month": month, "sku_group": sku_group, "return_qty": float(return_qty or 0)})
    print(f"  -> {len(returns_by_sku)} returns_by_sku rows from Snowflake")
    return returns_by_sku


def fetch_images_from_snowflake():
    """One image URL per sku_group -> dict, for O(1) lookup client-side
    (Pareto/Quartile tab thumbnails). Empty dict (not an error) if the query
    returns nothing, so the tabs still render without images."""
    conn = _snowflake_connect()
    try:
        cur = conn.cursor()
        print("Fetching sku_group -> image_url map...")
        cur.execute(IMAGES_QUERY)
        rows = cur.fetchall()
        cur.close()
    finally:
        conn.close()

    images = {}
    for sku_group, image_url in rows:
        if sku_group and image_url:
            images[sku_group] = image_url
    print(f"  -> {len(images)} sku_group images from Snowflake")
    return images


# ---------------------------------------------------------------------------
# Snowflake: daily-grain actual Returns (RETURNS_DATA), joined to the product
# master (meta_mapping_cogs_sku) to get L1/Category/Meta1/2/3 per SKU_GROUP,
# for the new Analytics tab's "Returns" metric (actual, not the learned/frozen
# % used by the existing Returns tab).
#
# Channel split: RETURNS_DATA only has OVERALL_RETURNS_QTY and
# SHOPIFY_RETURNS_QTY (no separate Marketplace/Offline columns). Per Aditya:
# Offline returns are always 0, so:
#   shopify_returns_qty    = SHOPIFY_RETURNS_QTY (exact)
#   offline_returns_qty    = 0 (fixed)
#   marketplace_returns_qty = OVERALL_RETURNS_QTY - SHOPIFY_RETURNS_QTY (exact,
#                             since Offline contributes nothing to the gap)
# One row per (date, sku_group) -- exploded into per-channel qty here, at the
# SAME leaf grain (l1/cat/m1/m2/m3) as Sales, so it can be sliced identically.
# ---------------------------------------------------------------------------
RETURNS_QUERY = rf"""
WITH meta_map AS (
    SELECT
        UPPER(
            IFF(
                UPPER(REPLACE(a.SKU_GROUP, ' ', '')) LIKE 'MP%'
                OR UPPER(REPLACE(a.SKU_GROUP, ' ', '')) LIKE '4C-%',
                REGEXP_REPLACE(REPLACE(a.SKU_GROUP, ' ', ''), '^([^-]+-[^-]+).*$', '\1'),
                REGEXP_REPLACE(REPLACE(a.SKU_GROUP, ' ', ''), '-.*$', '')
            )
        ) AS sku_group_clean,
        COALESCE(
            MIN(CASE WHEN UPPER(TRIM(a.l1_category)) = 'LONG TAIL' THEN a.l1_category END),
            MIN(CASE WHEN UPPER(TRIM(a.l1_category)) = 'PLUS' THEN a.l1_category END),
            MIN(CASE WHEN UPPER(TRIM(a.l1_category)) = 'LUXE' THEN a.l1_category END),
            MIN(CASE WHEN UPPER(TRIM(a.l1_category)) = 'SNITCH' THEN a.l1_category END),
            MIN(a.l1_category)
        ) AS l1_category,
        COALESCE(
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'SHIRTS' THEN a.category END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'TSHIRTS' THEN a.category END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'JEANS' THEN a.category END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'TROUSERS' THEN a.category END),
            MIN(a.category)
        ) AS category,
        COALESCE(
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'SHIRTS' THEN a.meta1 END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'TSHIRTS' THEN a.meta1 END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'JEANS' THEN a.meta1 END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'TROUSERS' THEN a.meta1 END),
            MIN(a.meta1)
        ) AS meta1,
        COALESCE(
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'SHIRTS' THEN a.meta2 END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'TSHIRTS' THEN a.meta2 END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'JEANS' THEN a.meta2 END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'TROUSERS' THEN a.meta2 END),
            MIN(a.meta2)
        ) AS meta2,
        COALESCE(
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'SHIRTS' THEN a.meta3 END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'TSHIRTS' THEN a.meta3 END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'JEANS' THEN a.meta3 END),
            MIN(CASE WHEN UPPER(TRIM(a.category)) = 'TROUSERS' THEN a.meta3 END),
            MIN(a.meta3)
        ) AS meta3,
        MAX(a.cogs) AS cogs
    FROM snitch_db.maplemonk.meta_mapping_cogs_sku a
    GROUP BY 1
),
returns_with_parent AS (
    SELECT
        r.*,
        UPPER(
            IFF(
                UPPER(REPLACE(r.SKU_GROUP, ' ', '')) LIKE 'MP%'
                OR UPPER(REPLACE(r.SKU_GROUP, ' ', '')) LIKE '4C-%',
                REGEXP_REPLACE(REPLACE(r.SKU_GROUP, ' ', ''), '^([^-]+-[^-]+).*$', '\1'),
                REGEXP_REPLACE(REPLACE(r.SKU_GROUP, ' ', ''), '-.*$', '')
            )
        ) AS sku_group_clean
    FROM SNITCH_DB.MAPLEMONK.RETURNS_DATA r
    WHERE r.DATE >= '{ROLLING_CUTOFF_DATE}'
)
SELECT
    r.DATE,
    m.l1_category, m.category, m.meta1, m.meta2, m.meta3,
    SUM(r.SHOPIFY_REVENUE) AS SHOPIFY_REVENUE, SUM(r.SHOPIFY_QTY) AS SHOPIFY_QTY,
    SUM(r.MP_REVENUE) AS MP_REVENUE, SUM(r.MP_QTY) AS MP_QTY,
    SUM(r.OFFLINE_REVENUE) AS OFFLINE_REVENUE, SUM(r.OFFLINE_QTY) AS OFFLINE_QTY,
    SUM(r.OVERALL_RETURNS_QTY) AS OVERALL_RETURNS_QTY, SUM(r.SHOPIFY_RETURNS_QTY) AS SHOPIFY_RETURNS_QTY
FROM returns_with_parent r
LEFT JOIN meta_map m
    ON r.sku_group_clean = m.sku_group_clean
GROUP BY r.DATE, m.l1_category, m.category, m.meta1, m.meta2, m.meta3
"""


def fetch_returns_from_snowflake():
    conn = _snowflake_connect()
    try:
        cur = conn.cursor()
        print("Fetching returns from RETURNS_DATA (joined to meta_mapping_cogs_sku)...")
        cur.execute(RETURNS_QUERY)
        rows = cur.fetchall()
        cur.close()
    finally:
        conn.close()

    returns_actual = []
    unmapped = 0
    for row in rows:
        (dt, l1, cat, m1, m2, m3,
         shopify_rev, shopify_qty, mp_rev, mp_qty, offline_rev, offline_qty,
         overall_returns_qty, shopify_returns_qty) = row
        if l1 is None or cat is None:
            unmapped += 1
            continue
        d = dt.strftime("%Y-%m-%d") if hasattr(dt, "strftime") else str(dt)[:10]
        overall_ret = float(overall_returns_qty or 0)
        shopify_ret = float(shopify_returns_qty or 0)
        mp_ret = max(0.0, overall_ret - shopify_ret)  # offline always 0, per Aditya
        base = {"date": d, "month": month_key(dt), "l1": l1, "cat": cat, "m1": m1, "m2": m2, "m3": m3}
        if float(shopify_qty or 0) or shopify_ret:
            returns_actual.append({**base, "channel": "Shopify",
                                    "sales_qty": float(shopify_qty or 0), "sales_value": float(shopify_rev or 0),
                                    "return_qty": shopify_ret})
        if float(mp_qty or 0) or mp_ret:
            returns_actual.append({**base, "channel": "Marketplace",
                                    "sales_qty": float(mp_qty or 0), "sales_value": float(mp_rev or 0),
                                    "return_qty": mp_ret})
        if float(offline_qty or 0):
            returns_actual.append({**base, "channel": "Offline",
                                    "sales_qty": float(offline_qty or 0), "sales_value": float(offline_rev or 0),
                                    "return_qty": 0.0})
    if unmapped:
        print(f"  NOTE: {unmapped} rows had no L1/Category match in meta_mapping_cogs_sku, skipped.")
    print(f"  -> {len(returns_actual)} returns rows (per-channel) from Snowflake")
    return returns_actual


def month_key(dt):
    if dt is None:
        return None
    if isinstance(dt, str):
        return None  # e.g. "No Date mentioned" -- treat as undated, not a real month
    if hasattr(dt, "year") and dt.year < 2020:
        return None  # placeholder/epoch-error dates (e.g. 1970-01-01) -- not real
    return dt.strftime("%Y-%m")


def build_returns():
    """Loads the frozen, one-time 'learning' dataset (data/returns_learned.json,
    monthly Sales Qty / Return Qty at Channel x L1 x Category grain, 2026 only,
    MP-SOR already merged into Marketplace). This is NOT re-derived from a live
    source -- it's a static snapshot used as the starting point for the editable
    Returns tab in the dashboard, which is where the actually-applied rates
    (with caps, and any manual overrides) live from here on."""
    path = os.path.join(HERE, "data", "returns_learned.json")
    if not os.path.exists(path):
        print(f"NOTE: {path} not found -- Returns tab will start empty (0% assumed everywhere).")
        return []
    with open(path) as f:
        learned = json.load(f)
    print(f"Returns: loaded {len(learned)} frozen monthly learning rows.")
    return learned


def build_channel_mix_pct():
    """Computes OLD channel-mix % per (L1, Category, Month) from the 4
    Planned_Qty_only files -- summed across Meta/ASP-Bin/store, kept distinct
    per Channel (Warehouse excluded, not a sales channel)."""
    old_qty = {}  # (l1, cat, channel, month) -> qty
    any_found = False
    for fname in TARGET_FILES:
        path = os.path.join(TARGETS_DIR, fname)
        if not os.path.exists(path):
            print(f"NOTE: target file not found (skipping): data/targets/{fname}")
            continue
        any_found = True
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb["Sheet1"]
        col_idx = None
        for row in ws.iter_rows(values_only=True):
            if col_idx is None:
                if row and "L1_CATEGORY" in row:
                    col_idx = {name: i for i, name in enumerate(row)}
                continue
            l1 = row[col_idx["L1_CATEGORY"]]
            if l1 is None or l1 == "L1_CATEGORY":
                continue
            cat = (row[col_idx["CATEGORY"]] or "").strip().lower()
            channel_raw = row[col_idx["type"]]
            if channel_raw is None or channel_raw == "type" or channel_raw == "Warehouse":
                continue
            channel = TARGET_CHANNEL_MAP.get(channel_raw, channel_raw)
            for col_name, month_iso in MONTH_COL_MAP.items():
                idx = col_idx.get(col_name)
                if idx is None:
                    continue
                val = row[idx] or 0
                key = (l1, cat, channel, month_iso)
                old_qty[key] = old_qty.get(key, 0) + val
        wb.close()
    if not any_found:
        print("NOTE: no target files found in data/targets/ -- cannot compute channel mix.")
        return {}

    # totals per (l1, cat, month) across all channels
    totals = {}
    for (l1, cat, channel, month), qty in old_qty.items():
        tkey = (l1, cat, month)
        totals[tkey] = totals.get(tkey, 0) + qty

    pct = {}
    for (l1, cat, channel, month), qty in old_qty.items():
        total = totals.get((l1, cat, month), 0)
        if total > 0:
            pct[(l1, cat, channel, month)] = qty / total
    return pct


def build_new_totals():
    """Reads data/targets_new_total.xlsx (the new all-channel-combined Planned
    Qty numbers), aggregating away Meta1/2/3 and ASP Bin (ASP is parked for now)
    down to L1 x Category x Month."""
    if not os.path.exists(NEW_TOTAL_FILE):
        print("NOTE: data/targets_new_total.xlsx not found -- no new target totals to apply.")
        return {}
    wb = openpyxl.load_workbook(NEW_TOTAL_FILE, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    col_idx = None
    NEW_MONTH_COLS = {
        "Planned Qty AUG'26": "2026-08", "Planned Qty SEP'26": "2026-09",
        "Planned Qty OCT'26": "2026-10", "Planned Qty NOV'26": "2026-11",
        "Planned Qty DEC'26": "2026-12",
    }
    totals = {}
    for row in ws.iter_rows(values_only=True):
        if col_idx is None:
            if row and "Category" in row and "L1" in row:
                col_idx = {name: i for i, name in enumerate(row)}
            continue
        cat_raw = row[col_idx["Category"]]
        if cat_raw is None:
            continue
        cat = str(cat_raw).strip().lower()
        l1 = row[col_idx["L1"]]
        for col_name, month_iso in NEW_MONTH_COLS.items():
            idx = col_idx.get(col_name)
            if idx is None:
                continue
            val = row[idx] or 0
            key = (l1, cat, month_iso)
            totals[key] = totals.get(key, 0) + val
    return totals


def build_targets():
    """New target logic: preserves the OLD channel-mix % (from the 4
    Planned_Qty_only files) and applies it to the NEW total (from
    data/targets_new_total.xlsx), per L1 x Category x Month. The old files
    are now used only to derive the split ratio, not as target values
    themselves -- the new file's totals are authoritative."""
    pct = build_channel_mix_pct()
    new_totals = build_new_totals()

    targets = []
    for (l1, cat, month), new_total in new_totals.items():
        matching_channels = [k for k in pct if k[0] == l1 and k[1] == cat and k[3] == month]
        if not matching_channels:
            continue  # no old channel-mix data for this L1+Cat+Month -- can't split honestly
        for key in matching_channels:
            _, _, channel, _ = key
            share = pct[key]
            targets.append({"channel": channel, "l1": l1, "cat": cat, "month": month, "qty": new_total * share})
    return targets


def build(xlsx_path_or_workbook):
    if isinstance(xlsx_path_or_workbook, str):
        wb = openpyxl.load_workbook(xlsx_path_or_workbook, data_only=True)
    else:
        wb = xlsx_path_or_workbook  # already a loaded workbook (or SimpleWorkbook from the API path)

    # ---- Sales (now from Snowflake, daily grain -- see fetch_sales_from_snowflake) ----
    sales = fetch_sales_from_snowflake()
    sales_by_sku = fetch_sales_by_sku_from_snowflake()
    returns_by_sku = fetch_returns_by_sku_from_snowflake()
    returns_actual = fetch_returns_from_snowflake()
    images = fetch_images_from_snowflake()

    # ---- Inventory ----
    # Kept at DAILY grain (not summed to month) because Closing(month) = the
    # latest dated snapshot within that month, per leaf combo -- summing days
    # together would wildly overstate stock. The frontend does the "latest
    # snapshot in month" pick, exactly matching the Overall tab's
    # SUMIFS(...,'CI v2'!date, MAXIFS(...)) formula.
    ws = wb["Inv Data 2"]
    inventory = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        inv_date, l1, cat, m1, m2, m3, mp_qty, online_qty, offline_qty, total_qty, mp_val, online_val, offline_val, total_val = row[:14]
        if l1 is None or cat is None:
            # stray/blank rows found in the source sheet (no L1 or Category) -- not real
            # inventory, skip so they don't pollute L1/Category rollups.
            continue
        d = inv_date.strftime("%Y-%m-%d") if hasattr(inv_date, "strftime") else str(inv_date)[:10]
        inventory.append({
            "date": d, "l1": l1, "cat": cat, "m1": m1, "m2": m2, "m3": m3,
            "mp_qty": mp_qty or 0, "online_qty": online_qty or 0, "offline_qty": offline_qty or 0,
            "total_qty": total_qty or 0,
            "mp_value": mp_val or 0, "online_value": online_val or 0, "offline_value": offline_val or 0,
            "total_value": total_val or 0
        })

    # ---- Pipeline (for Inwards) ----
    # Matched at the same leaf grain as inventory (L1+Cat+Meta1+Meta2+Meta3), keyed
    # by FDD_MONTH -- the month a batch is expected to land. No STATUS filter,
    # matching the original Overall-tab formula exactly.
    ws = wb["Pipeline"]
    pipeline = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        l1, cat, m1, m2, m3 = row[9], row[10], row[11], row[12], row[13]
        fdd_month, qty, cogs_unit = row[21], row[22], row[23]
        if l1 is None or cat is None:
            continue
        qty = qty or 0
        pipeline.append({
            "l1": l1, "cat": cat, "m1": m1, "m2": m2, "m3": m3,
            "fdd_month": month_key(fdd_month),  # None if undated
            "qty": qty,
            "cogs_value": qty * (cogs_unit or 0)
        })

    targets = build_targets()
    returns = build_returns()

    og_path = os.path.join(HERE, "data", "targets_og_fallback.json")
    targets_og = []
    if os.path.exists(og_path):
        with open(og_path) as f:
            targets_og = json.load(f)
    else:
        print("NOTE: data/targets_og_fallback.json not found -- no fallback for uncovered categories.")

    data_dir = os.path.dirname(OUT)
    os.makedirs(data_dir, exist_ok=True)

    # Split across multiple files instead of one giant data.json:
    #   1. GitHub hard-blocks any single file over 100MB -- one big file was
    #      already past that (225MB+) and would fail to push outright.
    #   2. The browser can fetch these in parallel and start rendering the
    #      Dashboard as soon as sales/inventory/meta arrive, without waiting
    #      on the (much larger) returns_actual/sales_by_sku to finish parsing.
    parts = {
        "meta.json": {
            "generated_at": datetime.datetime.now().isoformat(),
            "images": images,
        },
        "sales.json": {"sales": sales},
        "sales_by_sku.json": {"sales_by_sku": sales_by_sku, "returns_by_sku": returns_by_sku},
        "inventory.json": {"inventory": inventory},
        "pipeline.json": {"pipeline": pipeline},
        "returns_actual.json": {"returns_actual": returns_actual},
        "targets.json": {
            "targets": targets,        # primary source: the 4 category files, multi-channel
            "targets_og": targets_og,  # fallback: old Targets tab, Shopify-only, all categories
            "returns": returns,        # frozen monthly learning rows: Channel x L1 x Cat x Month (2026)
        },
    }
    written = []
    for filename, payload in parts.items():
        path = os.path.join(data_dir, filename)
        with open(path, "w") as f:
            json.dump(payload, f)
        size_mb = os.path.getsize(path) / (1024*1024)
        written.append((filename, size_mb))

    print(f"sales rows: {len(sales)}, sales_by_sku rows: {len(sales_by_sku)}, returns_by_sku rows: {len(returns_by_sku)}, "
          f"inventory rows: {len(inventory)}, pipeline rows: {len(pipeline)}, "
          f"target rows (new files): {len(targets)}, target rows (OG fallback): {len(targets_og)}, "
          f"returns_actual rows: {len(returns_actual)}, images: {len(images)}")
    if len(sales) > 0 and len(inventory) == 0:
        print("\n⚠️  WARNING: Sales has rows but Inventory is empty. This is the exact symptom")
        print("   of an unauthenticated download failing to resolve IMPORTRANGE on Inv Data 2.")
        print("   Fix: re-download Automation_Data.xlsx manually via your browser (File > Download")
        print("   > Microsoft Excel), then re-run with --local. See the module docstring for detail.\n")

    print("\nWrote data files:")
    total_mb = 0
    for filename, size_mb in written:
        flag = "  ⚠️ still >90MB, close to GitHub's 100MB limit!" if size_mb > 90 else ""
        print(f"  data/{filename}: {size_mb:.1f} MB{flag}")
        total_mb += size_mb
    print(f"  TOTAL: {total_mb:.1f} MB across {len(written)} files -> {data_dir}")


if __name__ == "__main__":
    if "--local" in sys.argv:
        if not os.path.exists(LOCAL_XLSX):
            print(f"ERROR: {LOCAL_XLSX} not found.")
            sys.exit(1)
        build(LOCAL_XLSX)
    elif "--export-url" in sys.argv:
        download_from_gsheet(SHEET_ID, LOCAL_XLSX)
        build(LOCAL_XLSX)
    else:
        wb = fetch_via_sheets_api(SHEET_ID)
        build(wb)