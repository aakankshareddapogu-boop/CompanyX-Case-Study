from __future__ import annotations

from io import BytesIO
import re
import unicodedata

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Company X | Slow Stock Action Tool",
    page_icon="📦",
    layout="wide",
)

REQUIRED_COLUMNS = {
    "items": [
        "SKU", "Item Name", "Category", "Master Pack Qty",
        "Unit Cost (Rs)", "Alternate Item (name)", "Alternate Flag", "Status"
    ],
    "stock": [
        "SKU", "Qty 0-30 days", "Qty 31-60 days", "Qty 61-90 days",
        "Qty 91-180 days", "Qty 180+ days"
    ],
    "sales": [
        "Month", "CustomerID", "CustomerName", "Industry", "SKU", "Qty", "Value (Rs)"
    ],
    "customers": ["CustomerID", "CustomerName", "Industry", "Area"],
}

STOCK_COLUMNS = [
    "Qty 0-30 days", "Qty 31-60 days", "Qty 61-90 days",
    "Qty 91-180 days", "Qty 180+ days"
]


def clean_header(value):
    """Normalize column headings so harmless spaces/case differences do not break upload."""
    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text).casefold()


def identify_dataset(df):
    """Identify a data table by its column headings, not by its filename."""
    normalized = {clean_header(c): c for c in df.columns}
    for dataset_name, required in REQUIRED_COLUMNS.items():
        required_norm = {clean_header(c) for c in required}
        if required_norm.issubset(set(normalized)):
            result = df.rename(
                columns={normalized[clean_header(c)]: c for c in required}
            ).copy()
            return dataset_name, result
    return None, None


def read_uploaded_files(uploaded_files):
    """Read one workbook with four sheets or four separate Excel files."""
    found = {}
    ignored = []
    problems = []

    for uploaded in uploaded_files:
        try:
            uploaded.seek(0)
            excel = pd.ExcelFile(uploaded)
            for sheet_name in excel.sheet_names:
                df = pd.read_excel(excel, sheet_name=sheet_name)
                dataset_name, recognized = identify_dataset(df)
                if dataset_name is None:
                    if not df.dropna(how="all").empty:
                        ignored.append(
                            f"{uploaded.name} — sheet '{sheet_name}' "
                            "(not a required data table)"
                        )
                    continue

                if dataset_name in found:
                    problems.append(
                        f"More than one table was detected for **{dataset_name}**. "
                        f"Using the first detected table and ignoring the duplicate."
                    )
                    continue

                found[dataset_name] = recognized
        except Exception as exc:
            problems.append(f"Could not read **{uploaded.name}**: {exc}")

    missing = [name for name in REQUIRED_COLUMNS if name not in found]
    if missing:
        friendly = {
            "items": "Item Master",
            "stock": "Stock Aging",
            "sales": "Sales Lines",
            "customers": "Customers",
        }
        missing_names = ", ".join(friendly[name] for name in missing)
        problems.append(
            f"Missing required data table(s): {missing_names}. Upload either one workbook "
            "with all four tables on separate sheets, or the four separate Excel files."
        )
    return found, problems, ignored


def normalize_text(series):
    return series.fillna("").astype(str).str.strip()


def normalize_month(value):
    """Normalize common month labels such as Apr-2026 and 2026-06."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        parsed = pd.to_datetime(text, format="%Y-%m", errors="coerce")
    return text if pd.isna(parsed) else parsed.strftime("%Y-%m")


def prepare_data(raw):
    """Clean the data, combine switched SKUs, classify stock and create target lists."""
    issues = []
    audit = []

    items = raw["items"].copy()
    stock = raw["stock"].copy()
    sales = raw["sales"].copy()
    customers = raw["customers"].copy()

    # Standardize SKU and customer IDs before matching tables.
    for df in [items, stock, sales]:
        df["SKU"] = normalize_text(df["SKU"]).str.upper()
    for df in [customers, sales]:
        df["CustomerID"] = normalize_text(df["CustomerID"]).str.upper()

    for df, columns in [
        (items, ["Item Name", "Category", "Alternate Item (name)", "Alternate Flag", "Status"]),
        (sales, ["Month", "CustomerName", "Industry"]),
        (customers, ["CustomerName", "Industry", "Area"]),
    ]:
        for col in columns:
            df[col] = normalize_text(df[col])

    # Convert numeric columns; report bad values instead of silently hiding them.
    numeric_columns = {
        "items": ["Master Pack Qty", "Unit Cost (Rs)"],
        "stock": STOCK_COLUMNS,
        "sales": ["Qty", "Value (Rs)"],
    }
    for name, cols in numeric_columns.items():
        df = {"items": items, "stock": stock, "sales": sales}[name]
        for col in cols:
            old_missing = int(df[col].isna().sum())
            df[col] = pd.to_numeric(df[col], errors="coerce")
            invalid_count = int(df[col].isna().sum()) - old_missing
            if invalid_count:
                issues.append(
                    f"{name}: {invalid_count} value(s) in '{col}' were not numeric."
                )

    # Standardize month labels before identifying exact duplicate sales lines.
    raw_duplicate_sales = int(sales.duplicated().sum())
    sales["Month"] = sales["Month"].map(normalize_month)
    normalized_duplicate_sales = int(sales.duplicated().sum())
    sales = sales.drop_duplicates().copy()
    audit.append(
        f"Excluded {normalized_duplicate_sales} exact duplicate sales row(s) after "
        f"normalizing text/month formats. Exact duplicates before month normalization: "
        f"{raw_duplicate_sales}."
    )
    if normalized_duplicate_sales:
        issues.append(
            f"{normalized_duplicate_sales} exact duplicate sales row(s) were excluded "
            "from calculations. Confirm with the source owner that these are accidental."
        )

    # Duplicate keys can make joins ambiguous.
    duplicate_item_skus = (
        items.loc[items["SKU"].duplicated(keep=False), "SKU"].dropna().unique().tolist()
    )
    duplicate_stock_skus = (
        stock.loc[stock["SKU"].duplicated(keep=False), "SKU"].dropna().unique().tolist()
    )
    if duplicate_item_skus:
        issues.append(
            "Duplicate Item Master SKU(s): " + ", ".join(duplicate_item_skus)
            + ". The first row per SKU is used."
        )
        items = items.drop_duplicates("SKU", keep="first")
    if duplicate_stock_skus:
        issues.append(
            "Duplicate Stock Aging SKU(s): " + ", ".join(duplicate_stock_skus)
            + ". Duplicate stock rows are summed."
        )

    items["Alternate Flag"] = items["Alternate Flag"].str.upper()
    items["Status"] = items["Status"].str.casefold()
    items["Category"] = items["Category"].replace("", "Unknown category")

    if items["Master Pack Qty"].isna().any() or (items["Master Pack Qty"] <= 0).any():
        issues.append(
            "Item Master has missing/non-positive Master Pack Qty values. "
            "Affected items need source correction."
        )
    if items["Unit Cost (Rs)"].isna().any() or (items["Unit Cost (Rs)"] < 0).any():
        issues.append(
            "Item Master has missing/negative Unit Cost values; aged values may be incomplete."
        )
    if sales[["Qty", "Value (Rs)"]].isna().any().any():
        issues.append(
            "Sales Lines has missing/non-numeric Qty or Value (Rs); affected cells are treated as zero."
        )
    sales[["Qty", "Value (Rs)"]] = sales[["Qty", "Value (Rs)"]].fillna(0)

    # Flag negative stock and use zero for those bucket cells in calculations only.
    stock[STOCK_COLUMNS] = stock[STOCK_COLUMNS].fillna(0)
    negative_mask = stock[STOCK_COLUMNS] < 0
    negative_cell_count = int(negative_mask.sum().sum())
    if negative_cell_count:
        bad_skus = stock.loc[negative_mask.any(axis=1), "SKU"].tolist()
        issues.append(
            f"Negative stock found in {negative_cell_count} bucket cell(s), SKU(s): "
            f"{', '.join(bad_skus)}. Negative bucket values are set to zero in calculations "
            "only; the uploaded source file is not changed."
        )
        audit.append(
            f"Set {negative_cell_count} negative stock bucket value(s) to zero in the calculation copy only."
        )
    stock[STOCK_COLUMNS] = stock[STOCK_COLUMNS].clip(lower=0)

    # Check unmatched SKU codes before applying the switch mapping.
    master_skus = set(items["SKU"].dropna())
    unmatched_stock = sorted(set(stock["SKU"].dropna()) - master_skus)
    unmatched_sales = sorted(set(sales["SKU"].dropna()) - master_skus)
    if unmatched_stock:
        issues.append(
            "Stock SKU(s) not found in Item Master: " + ", ".join(unmatched_stock)
            + ". They are excluded from item analysis."
        )
    if unmatched_sales:
        issues.append(
            "Sales SKU(s) not found in Item Master: " + ", ".join(unmatched_sales)
            + ". They are excluded from item analysis."
        )

    # Resolve only explicit SWITCH items, using alternate item name to find new SKU.
    name_to_sku = {
        str(row["Item Name"]).strip().casefold(): row["SKU"]
        for _, row in items.iterrows()
        if str(row["Item Name"]).strip()
    }
    switch_map = {}
    unresolved_switches = []
    for _, row in items.iterrows():
        if row["Alternate Flag"] == "SWITCH":
            old_sku = row["SKU"]
            alternate_name = str(row["Alternate Item (name)"]).strip().casefold()
            new_sku = name_to_sku.get(alternate_name)
            if not alternate_name or not new_sku:
                unresolved_switches.append(old_sku)
            elif old_sku != new_sku:
                switch_map[old_sku] = new_sku

    if unresolved_switches:
        issues.append(
            "SWITCH item(s) with a missing/unmatched alternate item name: "
            + ", ".join(unresolved_switches)
            + ". These could not be combined automatically."
        )
    if switch_map:
        audit.append(
            "Combined switched SKU pairs under replacement SKU: "
            + ", ".join(f"{old} → {new}" for old, new in switch_map.items()) + "."
        )

    # Flag inactive items that still have stock before excluding them.
    inactive_skus = set(items.loc[items["Status"].eq("inactive"), "SKU"])
    inactive_stock_check = stock.copy()
    inactive_stock_check["Total stock in source buckets"] = inactive_stock_check[STOCK_COLUMNS].sum(axis=1)
    inactive_with_stock = inactive_stock_check[
        inactive_stock_check["SKU"].isin(inactive_skus)
        & (inactive_stock_check["Total stock in source buckets"] > 0)
    ][["SKU", "Total stock in source buckets"]].copy()
    if not inactive_with_stock.empty:
        issues.append(
            "Inactive item(s) still have stock: "
            + ", ".join(inactive_with_stock["SKU"].astype(str).tolist())
            + ". Excluded from active classification; inventory owner should confirm disposition."
        )

    # Map old SKU stock/sales to replacement SKU and sum the stock buckets.
    stock["SKU"] = stock["SKU"].replace(switch_map)
    sales["SKU"] = sales["SKU"].replace(switch_map)
    stock = stock.groupby("SKU", as_index=False)[STOCK_COLUMNS].sum()

    active_items = items[items["Status"].eq("active")].copy()
    active_items = active_items[~active_items["SKU"].isin(switch_map.keys())].copy()

    stock["Total Stock"] = stock[STOCK_COLUMNS].sum(axis=1)
    stock["Aged Qty"] = stock["Qty 91-180 days"] + stock["Qty 180+ days"]

    sales_summary = sales.groupby("SKU", as_index=False).agg(
        **{
            "Six-Month Sales Qty": ("Qty", "sum"),
            "Six-Month Sales Value (Rs)": ("Value (Rs)", "sum"),
        }
    )

    analysis = active_items.merge(stock, on="SKU", how="left", validate="one_to_one")
    analysis = analysis.merge(sales_summary, on="SKU", how="left", validate="one_to_one")

    fill_cols = STOCK_COLUMNS + [
        "Total Stock", "Aged Qty", "Six-Month Sales Qty", "Six-Month Sales Value (Rs)"
    ]
    analysis[fill_cols] = analysis[fill_cols].fillna(0)

    # Aged value uses the replacement SKU's cost after pack-size switches.
    analysis["Aged Value (Rs)"] = analysis["Aged Qty"] * analysis["Unit Cost (Rs)"]
    analysis["Average Monthly Sales Qty"] = analysis["Six-Month Sales Qty"] / 6

    analysis["Months to Clear"] = float("inf")
    has_sales = analysis["Average Monthly Sales Qty"] > 0
    analysis.loc[has_sales, "Months to Clear"] = (
        analysis.loc[has_sales, "Total Stock"]
        / analysis.loc[has_sales, "Average Monthly Sales Qty"]
    )

    # Classification order follows the brief exactly.
    def classify(row):
        pack = row["Master Pack Qty"]
        if pd.isna(pack) or pack <= 0:
            return "Review Data"
        if row["Six-Month Sales Qty"] < pack:
            return "Red List"
        if row["Aged Qty"] <= 0 or row["Months to Clear"] < 3:
            return "No Action"
        return "Push"

    analysis["Class"] = analysis.apply(classify, axis=1)
    analysis["Months to Clear Display"] = analysis["Months to Clear"].map(
        lambda value: "No sales" if pd.isna(value) or value == float("inf")
        else round(float(value), 2)
    )

    # Customer value eligibility is based on total purchase value over six months.
    customer_sales = sales.groupby("CustomerID", as_index=False).agg(
        **{"Six-Month Purchase Value (Rs)": ("Value (Rs)", "sum")}
    )
    customer_master = customers.drop_duplicates("CustomerID", keep="first").copy()
    customer_master["Industry"] = customer_master["Industry"].replace("", pd.NA)
    customer_master["CustomerName"] = customer_master["CustomerName"].replace("", "Unknown customer")
    customer_summary = customer_master.merge(customer_sales, on="CustomerID", how="left")
    customer_summary["Six-Month Purchase Value (Rs)"] = (
        customer_summary["Six-Month Purchase Value (Rs)"].fillna(0)
    )

    if customers["CustomerID"].duplicated().any():
        duplicate_ids = customers.loc[
            customers["CustomerID"].duplicated(keep=False), "CustomerID"
        ].unique().tolist()
        issues.append(
            "Duplicate CustomerID(s) found; first record is used: " + ", ".join(duplicate_ids) + "."
        )

    if customers["Industry"].eq("").any():
        blank_ids = customers.loc[customers["Industry"].eq(""), "CustomerID"].tolist()
        issues.append(
            "Customer industry is missing for: " + ", ".join(blank_ids)
            + ". These customers cannot enter Similar Business bucket."
        )

    # Customer history is calculated after switched SKUs are consolidated.
    item_customer_sales = sales.groupby(["SKU", "CustomerID"], as_index=False).agg(
        **{"Item Purchase Qty": ("Qty", "sum")}
    )
    prior_buyers = {
        sku: set(group.loc[group["Item Purchase Qty"] > 0, "CustomerID"])
        for sku, group in item_customer_sales.groupby("SKU")
    }

    sku_to_category = dict(zip(items["SKU"], items["Category"]))
    for old_sku, new_sku in switch_map.items():
        sku_to_category[old_sku] = sku_to_category.get(
            new_sku, sku_to_category.get(old_sku, "Unknown category")
        )
    sales["Category for target logic"] = sales["SKU"].map(sku_to_category)
    category_buyers = {
        category: set(group.loc[group["Qty"] > 0, "CustomerID"])
        for category, group in sales.groupby("Category for target logic", dropna=False)
    }

    target_rows = []
    eligible_customers = customer_summary[
        customer_summary["Six-Month Purchase Value (Rs)"] > 100000
    ].copy()
    customer_lookup = eligible_customers.set_index("CustomerID", drop=False).to_dict("index")

    for _, item in analysis[analysis["Class"].eq("Push")].iterrows():
        sku = item["SKU"]
        buyers = prior_buyers.get(sku, set())

        buyer_industries = set(
            customers.loc[
                customers["CustomerID"].isin(buyers)
                & customers["Industry"].notna()
                & customers["Industry"].astype(str).str.strip().ne(""),
                "Industry"
            ].astype(str).str.strip()
        )
        same_category_buyers = category_buyers.get(item["Category"], set())
        bucket_lists = {1: [], 2: [], 3: []}

        for _, customer in eligible_customers.iterrows():
            cid = customer["CustomerID"]
            industry = customer["Industry"]
            if cid in buyers:
                bucket_lists[1].append(cid)
            elif (
                pd.notna(industry) and str(industry).strip()
                and str(industry).strip() in buyer_industries
            ):
                bucket_lists[2].append(cid)
            elif cid in same_category_buyers:
                # Bucket 3 is reached only if bucket 1 and bucket 2 did not apply.
                bucket_lists[3].append(cid)

        for bucket, customer_ids in bucket_lists.items():
            for cid in customer_ids:
                info = customer_lookup[cid]
                target_rows.append({
                    "SKU": sku,
                    "Item Name": item["Item Name"],
                    "Category": item["Category"],
                    "CustomerID": cid,
                    "CustomerName": info.get("CustomerName", "Unknown customer"),
                    "Industry": info.get("Industry", ""),
                    "Area": info.get("Area", ""),
                    "Bucket": bucket,
                    "Bucket Label": {
                        1: "Bought this item",
                        2: "Similar business",
                        3: "Same category",
                    }[bucket],
                    "Six-Month Purchase Value (Rs)": info.get(
                        "Six-Month Purchase Value (Rs)", 0
                    ),
                })

    targets = pd.DataFrame(target_rows, columns=[
        "SKU", "Item Name", "Category", "CustomerID", "CustomerName", "Industry",
        "Area", "Bucket", "Bucket Label", "Six-Month Purchase Value (Rs)"
    ])
    if not targets.empty:
        targets = targets.sort_values(
            ["SKU", "Bucket", "Six-Month Purchase Value (Rs)", "CustomerName"],
            ascending=[True, True, False, True]
        ).reset_index(drop=True)

    if sales["Month"].nunique() != 6:
        issues.append(
            f"Sales data contains {sales['Month'].nunique()} distinct month label(s) after "
            "normalization; the brief expects six months. Verify the date range."
        )
    if sales["CustomerID"].eq("").any():
        issues.append("Sales Lines contains blank CustomerID values.")
    if customers["CustomerID"].eq("").any():
        issues.append("Customers contains blank CustomerID values.")

    final_sales_skus = set(sales["SKU"].dropna())
    no_sales_skus = sorted(set(analysis["SKU"]) - final_sales_skus)
    if no_sales_skus:
        issues.append(
            "Active final SKU(s) with no sales in the period: " + ", ".join(no_sales_skus)
            + ". Classification follows the specified Red List rule where applicable."
        )

    return {
        "analysis": analysis,
        "targets": targets,
        "customer_summary": customer_summary,
        "issues": pd.DataFrame({"Data quality finding": issues}),
        "audit": pd.DataFrame({"Handling applied": audit}),
        "inactive_stock": inactive_with_stock,
        "switch_map": switch_map,
        "raw_duplicate_sales": raw_duplicate_sales,
        "normalized_duplicate_sales": normalized_duplicate_sales,
    }


def export_results(result):
    """Export summary, item analysis, customer targets and checks to an Excel workbook."""
    analysis = result["analysis"].copy()
    targets = result["targets"].copy()

    summary = (
        analysis["Class"].value_counts()
        .rename_axis("Class")
        .reset_index(name="Item Count")
    )
    push_value = analysis.loc[
        analysis["Class"] == "Push", "Aged Value (Rs)"
    ].sum()
    summary.loc[len(summary)] = ["Total Push aged value (Rs)", push_value]

    top_push = analysis[analysis["Class"] == "Push"].sort_values(
        "Aged Value (Rs)", ascending=False
    ).head(10)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        analysis.to_excel(writer, sheet_name="Item Analysis", index=False)
        top_push.to_excel(writer, sheet_name="Top 10 Push", index=False)
        targets.to_excel(writer, sheet_name="Customer Targets", index=False)
        result["issues"].to_excel(writer, sheet_name="Data Quality", index=False)
        result["audit"].to_excel(writer, sheet_name="Handling Log", index=False)
        result["customer_summary"].to_excel(writer, sheet_name="Customer Summary", index=False)
    output.seek(0)
    return output.getvalue()


st.title("📦 Company X — Slow Stock Action Tool")
st.caption("Prioritize aged inventory and give sales a ranked customer call list.")

with st.expander("How to use this tool", expanded=True):
    st.markdown(
        """
        1. Upload either **one Excel workbook with four data sheets** or the **four separate data files**.
        2. Review data-quality findings and the handling log.
        3. Filter the item list by class or category.
        4. Select a Push item to view its ranked customer targets.
        5. Download the Excel output for sharing.

        The README file is optional. It will be ignored if it does not contain a required data table.
        """
    )

uploaded_files = st.file_uploader(
    "Upload the combined workbook or your Excel files",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    help="Upload one workbook with four sheets, or Item Master + Stock Aging + Sales Lines + Customers.",
)

if not uploaded_files:
    st.info("Upload your Excel file(s) to start. The app identifies tables from their column headings.")
    st.stop()

raw_data, load_problems, ignored_files = read_uploaded_files(uploaded_files)

if load_problems:
    st.warning("Please review the upload checks below.")
    for problem in load_problems:
        st.markdown(f"- {problem}")

if ignored_files:
    with st.expander("Files/sheets ignored"):
        for item in ignored_files:
            st.write(f"- {item}")

if any(name not in raw_data for name in REQUIRED_COLUMNS):
    st.stop()

try:
    result = prepare_data(raw_data)
except Exception as exc:
    st.error(f"The data could not be processed: {exc}")
    st.exception(exc)
    st.stop()

analysis = result["analysis"]
targets = result["targets"]

st.subheader("Overview")
counts = analysis["Class"].value_counts()
push_count = int(counts.get("Push", 0))
red_count = int(counts.get("Red List", 0))
no_action_count = int(counts.get("No Action", 0))
review_count = int(counts.get("Review Data", 0))
push_aged_value = float(
    analysis.loc[analysis["Class"] == "Push", "Aged Value (Rs)"].sum()
)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Push items", push_count)
c2.metric("Red List items", red_count)
c3.metric("No Action items", no_action_count)
c4.metric("Review Data", review_count)
c5.metric("Push aged value", f"₹{push_aged_value:,.0f}")

st.subheader("Data quality checks")
if result["issues"].empty:
    st.success("No data quality issues were detected by the current checks.")
else:
    st.warning(f"{len(result['issues'])} data-quality finding(s) need attention.")
    st.dataframe(result["issues"], use_container_width=True, hide_index=True)

with st.expander("Handling applied by the tool"):
    if result["audit"].empty:
        st.write("No special handling was required.")
    else:
        st.dataframe(result["audit"], use_container_width=True, hide_index=True)

if not result["inactive_stock"].empty:
    st.markdown("**Inactive products with stock**")
    st.dataframe(result["inactive_stock"], use_container_width=True, hide_index=True)

st.subheader("Top 10 Push items by aged stock value")
top_push = analysis[analysis["Class"] == "Push"].sort_values(
    "Aged Value (Rs)", ascending=False
).head(10)
if top_push.empty:
    st.info("No items are currently classified as Push.")
else:
    st.dataframe(
        top_push[[
            "SKU", "Item Name", "Category", "Total Stock", "Aged Qty",
            "Aged Value (Rs)", "Six-Month Sales Qty", "Average Monthly Sales Qty",
            "Months to Clear", "Class"
        ]],
        use_container_width=True,
        hide_index=True,
    )

st.subheader("Filter and inspect items")
filter1, filter2 = st.columns(2)
class_options = ["All"] + sorted(analysis["Class"].dropna().unique().tolist())
category_options = ["All"] + sorted(
    analysis["Category"].fillna("Unknown category").unique().tolist()
)
selected_class = filter1.selectbox("Filter by class", class_options)
selected_category = filter2.selectbox("Filter by category", category_options)

filtered = analysis.copy()
if selected_class != "All":
    filtered = filtered[filtered["Class"] == selected_class]
if selected_category != "All":
    filtered = filtered[
        filtered["Category"].fillna("Unknown category") == selected_category
    ]

item_columns = [
    "SKU", "Item Name", "Category", "Master Pack Qty", "Unit Cost (Rs)",
    "Total Stock", "Aged Qty", "Aged Value (Rs)", "Six-Month Sales Qty",
    "Average Monthly Sales Qty", "Months to Clear Display", "Class"
]
st.dataframe(filtered[item_columns], use_container_width=True, hide_index=True)

st.subheader("Customer call list for a Push item")
push_skus = sorted(
    analysis.loc[analysis["Class"] == "Push", "SKU"].dropna().unique().tolist()
)
if not push_skus:
    st.info("There are no Push items to show customer targets for.")
else:
    selected_sku = st.selectbox("Choose a Push SKU", push_skus)
    selected_targets = targets[targets["SKU"] == selected_sku].copy()
    if selected_targets.empty:
        st.info("No eligible customer targets matched the rules for this item.")
    else:
        st.caption(
            "Eligible customers must have more than ₹1,00,000 total purchase value. "
            "Bucket 1 ranks before bucket 2, then bucket 3; within each bucket, "
            "higher purchase value ranks first."
        )
        st.dataframe(
            selected_targets[[
                "Bucket", "Bucket Label", "CustomerName", "CustomerID",
                "Industry", "Area", "Six-Month Purchase Value (Rs)"
            ]],
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("**Quick view for a sales representative**")
        for rank, (_, row) in enumerate(selected_targets.iterrows(), start=1):
            with st.container(border=True):
                st.markdown(f"**{rank}. {row['CustomerName']}** · {row['CustomerID']}")
                st.write(f"Bucket {row['Bucket']}: {row['Bucket Label']}")
                industry = (
                    row["Industry"]
                    if pd.notna(row["Industry"]) and str(row["Industry"]).strip()
                    else "Not provided"
                )
                area = (
                    row["Area"]
                    if pd.notna(row["Area"]) and str(row["Area"]).strip()
                    else "Not provided"
                )
                st.write(f"Industry: {industry} · Area: {area}")
                st.write(
                    f"Six-month purchase value: ₹{row['Six-Month Purchase Value (Rs)']:,.0f}"
                )

st.subheader("Download results")
excel_bytes = export_results(result)
st.download_button(
    "Download Excel results",
    data=excel_bytes,
    file_name="CompanyX_Slow_Stock_Results.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
st.caption(
    "The export includes summary counts, item analysis, top 10 Push items, customer targets, "
    "data-quality findings, handling log and customer purchase totals."
)
