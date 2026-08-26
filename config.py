FISCAL_YEAR_FIXED = "2026"
CURRENCY_FIXED = "INR"
ASSET_MANAGED_HISTORICALLY_FIXED = "X"

MODE_LANDING = "landing"
MODE_MASTER = "master"
MODE_SUB = "sub"

TARGET_COLUMNS = [
    "Company Code",
    "External/Legacy Asset Number",
    "Asset Subnumber",
    "Depreciation Area",
    "Current Fiscal Year",
    "Cumulated Acquisition Value",
    "Cum. Revaluation on the Replacem. Value",
    "Cumulative Investment Grants",
    "Accumulated Ordinary Depreciation",
    "Cumulative Special Depreciation",
    "Cumulative Unplanned Depreciation",
    "Cumulative Transfer of Reserves",
    "Cumulative Interest",
    "Cumulative Revaluation of Ordinary Depr.",
    "Cumulative Down Payments",
    "Currency Key",
]

MASTER_DETAILS_FIELDS = [
    ("A", "Company code"),
    ("B", "External/Legacy Asset Number*"),
    ("E", "Asset Subnumber*"),
    ("F", "Asset Class*"),
    ("L", "Asset Description"),
    ("M", "Asset Description 2"),
    ("O", "Quantity"),
    ("P", "Base Unit of Measure"),
    ("L", "Asset Main Number Text"),
]

INVENTORY_FIELDS = [
    ("A", "Company code"),
    ("B", "External/Legacy Asset Number*"),
    ("E", "Asset Subnumber*"),
]

ORIGIN_FIELDS = [
    ("A", "Company code"),
    ("B", "External/Legacy Asset Number*"),
    ("E", "Asset Subnumber*"),
]