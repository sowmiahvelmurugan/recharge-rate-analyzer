# Streamlit app code is being prepared in a downloadable file.
import io
import os
import re
import zipfile
import hashlib
from itertools import combinations
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
import streamlit as st


st.set_page_config(page_title="Recharge Rate Analyzer", layout="wide")


# -----------------------------
# Helpers
# -----------------------------
def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def extract_embedded_images_from_xlsx(file_bytes: bytes) -> List[bytes]:
    images = []
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as zf:
            media_files = sorted(
                [name for name in zf.namelist() if name.startswith("xl/media/")]
            )
            for media_name in media_files:
                images.append(zf.read(media_name))
    except Exception:
        pass
    return images


# ---------------------------------
# Logo detection
# ---------------------------------
# Add more known logo hashes later if needed.
# For your uploaded sample files, the embedded logo is Airtel.
KNOWN_LOGO_HASHES = {
    # "sha256_hash_here": "Airtel",
    # "sha256_hash_here": "Jio",
    # "sha256_hash_here": "Vi",
    # "sha256_hash_here": "BSNL",
}


def detect_operator_from_logo(uploaded_file) -> Tuple[str, List[str]]:
    """
    Returns:
      operator_name,
      image_hashes_found

    Current behavior:
    - Tries exact hash match if KNOWN_LOGO_HASHES is populated
    - Otherwise falls back to Airtel if embedded image exists
      because your sample files use Airtel logo
    """
    file_bytes = uploaded_file.getvalue()
    images = extract_embedded_images_from_xlsx(file_bytes)
    hashes = [sha256_bytes(img) for img in images]

    for h in hashes:
        if h in KNOWN_LOGO_HASHES:
            return KNOWN_LOGO_HASHES[h], hashes

    if images:
        return "Airtel", hashes  # current fallback for your sample format

    return "Unknown", hashes


# ---------------------------------
# Excel parsing
# ---------------------------------
def find_header_row(ws) -> Optional[int]:
    expected = {"amount", "amounts", "margin", "zone"}
    max_scan = min(ws.max_row, 15)

    for r in range(1, max_scan + 1):
        row_vals = [clean_text(ws.cell(r, c).value).lower() for c in range(1, ws.max_column + 1)]
        matches = sum(1 for val in row_vals if val in expected)
        if matches >= 2:
            return r
    return None


def find_target_sheet(wb):
    for name in wb.sheetnames:
        ws = wb[name]
        if find_header_row(ws) is not None:
            return ws
    return wb[wb.sheetnames[0]]


def get_column_map(ws, header_row: int) -> Dict[str, int]:
    col_map = {}
    for c in range(1, ws.max_column + 1):
        val = clean_text(ws.cell(header_row, c).value).lower()
        if val:
            col_map[val] = c
    return col_map


def parse_rate(value: Any) -> Optional[float]:
    """
    Normalize rates to decimal form:
    2      -> 0.02
    1.8    -> 0.018
    0.02   -> 0.02
    """
    if value is None or clean_text(value) == "":
        return None

    if isinstance(value, str):
        s = value.replace("%", "").strip()
        try:
            num = float(s)
        except ValueError:
            return None
    else:
        try:
            num = float(value)
        except Exception:
            return None

    if num > 1:
        return num / 100.0
    return num


def parse_amount_spec(value: Any) -> Optional[Dict[str, Any]]:
    """
    Supported:
    - 100-5000
    - 33,49,77
    - ,33,49,
    - 199
    - Any / ALL
    """
    if value is None:
        return None

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return {"type": "exact_list", "values": [int(value)]}

    s = str(value).strip().replace("₹", "")
    s = re.sub(r"\s+", "", s)

    if s.lower() in {"any", "all"}:
        return {"type": "all"}

    if re.fullmatch(r"\d+", s):
        return {"type": "exact_list", "values": [int(s)]}

    if re.fullmatch(r"\d+-\d+", s):
        start, end = map(int, s.split("-"))
        if start > end:
            start, end = end, start
        return {"type": "range", "start": start, "end": end}

    nums = [int(x) for x in re.findall(r"\d+", s)]
    if nums:
        return {"type": "exact_list", "values": sorted(set(nums))}

    return None


def rule_applies(rule: Dict[str, Any], amount: int) -> bool:
    if rule["type"] == "all":
        return True
    if rule["type"] == "range":
        return rule["start"] <= amount <= rule["end"]
    if rule["type"] == "exact_list":
        return amount in rule["values"]
    return False


def load_supplier_rows(uploaded_file, supplier_name: str) -> Tuple[List[Dict[str, Any]], str, List[str]]:
    file_bytes = uploaded_file.getvalue()
    operator, logo_hashes = detect_operator_from_logo(uploaded_file)

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = find_target_sheet(wb)
    header_row = find_header_row(ws)

    if header_row is None:
        raise ValueError(f"Could not locate headers in file: {uploaded_file.name}")

    col_map = get_column_map(ws, header_row)
    amount_col = col_map.get("amount") or col_map.get("amounts")
    margin_col = col_map.get("margin")
    zone_col = col_map.get("zone")

    if not amount_col or not margin_col:
        raise ValueError(
            f"Required columns not found in {uploaded_file.name}. Need Amount/Amounts and Margin."
        )

    rows: List[Dict[str, Any]] = []

    for r in range(header_row + 1, ws.max_row + 1):
        amount_val = ws.cell(r, amount_col).value
        margin_val = ws.cell(r, margin_col).value
        zone_val = ws.cell(r, zone_col).value if zone_col else "All India"

        if amount_val is None or margin_val is None:
            continue

        rule = parse_amount_spec(amount_val)
        rate = parse_rate(margin_val)

        if rule is None or rate is None:
            continue

        rows.append(
            {
                "supplier": supplier_name,
                "operator": operator,
                "zone": clean_text(zone_val) or "All India",
                "raw_amount": clean_text(amount_val),
                "rule": rule,
                "rate": rate,
                "excel_row": r,
                "file_name": uploaded_file.name,
            }
        )

    return rows, operator, logo_hashes


# ---------------------------------
# Best-rate logic (Option A)
# ---------------------------------
def build_best_rate_map(rows: List[Dict[str, Any]], min_amt: int = 1, max_amt: int = 50000) -> Dict[int, Dict[str, Any]]:
    best: Dict[int, Dict[str, Any]] = {}

    for amt in range(min_amt, max_amt + 1):
        best_row = None
        best_rate = -1.0

        for row in rows:
            if rule_applies(row["rule"], amt) and row["rate"] > best_rate:
                best_rate = row["rate"]
                best_row = row

        if best_row is not None:
            best[amt] = best_row

    return best


def compress_best_map(best_map: Dict[int, Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not best_map:
        return []

    amounts = sorted(best_map.keys())
    segments = []

    start = amounts[0]
    prev = amounts[0]
    current = best_map[start]

    for amt in amounts[1:]:
        row = best_map[amt]
        same = (
            amt == prev + 1
            and row["supplier"] == current["supplier"]
            and row["operator"] == current["operator"]
            and row["zone"] == current["zone"]
            and row["rate"] == current["rate"]
            and row["excel_row"] == current["excel_row"]
            and row["raw_amount"] == current["raw_amount"]
        )

        if same:
            prev = amt
            continue

        segments.append(
            {
                "supplier": current["supplier"],
                "operator": current["operator"],
                "zone": current["zone"],
                "best_from": start,
                "best_to": prev,
                "rate_percent": round(current["rate"] * 100, 4),
                "source_amount_rule": current["raw_amount"],
                "source_excel_row": current["excel_row"],
                "file_name": current["file_name"],
            }
        )

        start = amt
        prev = amt
        current = row

    segments.append(
        {
            "supplier": current["supplier"],
            "operator": current["operator"],
            "zone": current["zone"],
            "best_from": start,
            "best_to": prev,
            "rate_percent": round(current["rate"] * 100, 4),
            "source_amount_rule": current["raw_amount"],
            "source_excel_row": current["excel_row"],
            "file_name": current["file_name"],
        }
    )

    return segments


# ---------------------------------
# Compare suppliers
# ---------------------------------
def compare_two_suppliers(
    best_a: Dict[int, Dict[str, Any]],
    best_b: Dict[int, Dict[str, Any]],
    supplier_a: str,
    supplier_b: str,
) -> List[Dict[str, Any]]:
    amounts = sorted(set(best_a.keys()).intersection(set(best_b.keys())))
    opportunities = []

    for amt in amounts:
        row_a = best_a[amt]
        row_b = best_b[amt]

        rate_a = row_a["rate"]
        rate_b = row_b["rate"]

        if rate_a == rate_b:
            continue

        if rate_a > rate_b:
            buy_from = supplier_a
            sell_to = supplier_b
            higher_rate = rate_a
            lower_rate = rate_b
        else:
            buy_from = supplier_b
            sell_to = supplier_a
            higher_rate = rate_b
            lower_rate = rate_a

        opportunities.append(
            {
                "amount": amt,
                "operator": row_a["operator"],
                "zone": row_a["zone"],
                "buy_from": buy_from,
                "sell_to": sell_to,
                "buy_rate_percent": round(higher_rate * 100, 4),
                "sell_rate_percent": round(lower_rate * 100, 4),
                "gross_margin_percent": round((higher_rate - lower_rate) * 100, 4),
            }
        )

    return opportunities


def compress_opportunities(opportunities: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not opportunities:
        return []

    segments = []
    start = opportunities[0]["amount"]
    prev = opportunities[0]["amount"]
    current = opportunities[0]

    for item in opportunities[1:]:
        same = (
            item["amount"] == prev + 1
            and item["operator"] == current["operator"]
            and item["zone"] == current["zone"]
            and item["buy_from"] == current["buy_from"]
            and item["sell_to"] == current["sell_to"]
            and item["buy_rate_percent"] == current["buy_rate_percent"]
            and item["sell_rate_percent"] == current["sell_rate_percent"]
            and item["gross_margin_percent"] == current["gross_margin_percent"]
        )

        if same:
            prev = item["amount"]
            continue

        segments.append(
            {
                "operator": current["operator"],
                "zone": current["zone"],
                "amount_from": start,
                "amount_to": prev,
                "buy_from_supplier": current["buy_from"],
                "sell_to_supplier": current["sell_to"],
                "buy_rate_percent": current["buy_rate_percent"],
                "sell_rate_percent": current["sell_rate_percent"],
                "gross_margin_percent": current["gross_margin_percent"],
            }
        )

        start = item["amount"]
        prev = item["amount"]
        current = item

    segments.append(
        {
            "operator": current["operator"],
            "zone": current["zone"],
            "amount_from": start,
            "amount_to": prev,
            "buy_from_supplier": current["buy_from"],
            "sell_to_supplier": current["sell_to"],
            "buy_rate_percent": current["buy_rate_percent"],
            "sell_rate_percent": current["sell_rate_percent"],
            "gross_margin_percent": current["gross_margin_percent"],
        }
    )

    return segments


# ---------------------------------
# Export
# ---------------------------------
def to_excel_bytes(
    summary_df: pd.DataFrame,
    best_segments_df: pd.DataFrame,
    opportunities_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if not summary_df.empty:
            summary_df.to_excel(writer, index=False, sheet_name="Summary")
        if not best_segments_df.empty:
            best_segments_df.to_excel(writer, index=False, sheet_name="Best_Rates")
        if not opportunities_df.empty:
            opportunities_df.to_excel(writer, index=False, sheet_name="Buy_Sell_Opportunities")
    output.seek(0)
    return output.read()


# ---------------------------------
# UI
# ---------------------------------
st.title("📊 Mobile Recharge Rate Analyzer")
st.write("Upload supplier Excel files, detect operator from logo, compute best rates, and compare buy/sell opportunities.")

with st.expander("How it works", expanded=False):
    st.markdown(
        """
        - Upload one or more supplier `.xlsx` files
        - Operator is detected from the embedded logo image
        - For each denomination, the app applies **Option A**
        - That means: **pick the highest applicable rate**
        - Then compare suppliers and show buy/sell profit slabs
        """
    )

uploaded_files = st.file_uploader(
    "Upload supplier Excel files",
    type=["xlsx"],
    accept_multiple_files=True,
)

col1, col2, col3 = st.columns(3)
with col1:
    min_amt = st.number_input("Minimum denomination", min_value=1, value=1, step=1)
with col2:
    max_amt = st.number_input("Maximum denomination", min_value=1, value=50000, step=1)
with col3:
    auto_names = st.checkbox("Use file name as supplier name", value=True)

if uploaded_files:
    summary_rows = []
    all_best_segments = []
    best_maps = {}

    for idx, file in enumerate(uploaded_files, start=1):
        supplier_name = os.path.splitext(file.name)[0] if auto_names else f"Supplier_{idx}"

        try:
            rows, operator, logo_hashes = load_supplier_rows(file, supplier_name)

            if not rows:
                st.warning(f"No valid rows found in {file.name}")
                continue

            best_map = build_best_rate_map(rows, int(min_amt), int(max_amt))
            segments = compress_best_map(best_map)

            best_maps[supplier_name] = best_map
            all_best_segments.extend(segments)

            summary_rows.append(
                {
                    "supplier": supplier_name,
                    "file_name": file.name,
                    "detected_operator": operator,
                    "embedded_images_found": len(logo_hashes),
                    "first_logo_hash": logo_hashes[0] if logo_hashes else "",
                    "valid_rows_loaded": len(rows),
                    "best_slabs_created": len(segments),
                }
            )

        except Exception as e:
            st.error(f"Error processing {file.name}: {e}")

    summary_df = pd.DataFrame(summary_rows)
    best_segments_df = pd.DataFrame(all_best_segments)

    if not summary_df.empty:
        st.subheader("Supplier processing summary")
        st.dataframe(summary_df, use_container_width=True)

    if not best_segments_df.empty:
        st.subheader("Best rates by supplier")
        st.dataframe(best_segments_df, use_container_width=True)

    pairwise_results = []
    supplier_names = list(best_maps.keys())

    if len(supplier_names) >= 2:
        for s1, s2 in combinations(supplier_names, 2):
            opportunities = compare_two_suppliers(best_maps[s1], best_maps[s2], s1, s2)
            compressed = compress_opportunities(opportunities)
            for row in compressed:
                row["supplier_pair"] = f"{s1} vs {s2}"
            pairwise_results.extend(compressed)

    opportunities_df = pd.DataFrame(pairwise_results)

    if not opportunities_df.empty:
        st.subheader("Buy / Sell opportunities")
        st.dataframe(opportunities_df, use_container_width=True)
    else:
        if len(supplier_names) < 2:
            st.info("Upload at least 2 supplier files to compare suppliers.")
        else:
            st.info("No buy/sell opportunity found.")

    if not summary_df.empty or not best_segments_df.empty or not opportunities_df.empty:
        excel_bytes = to_excel_bytes(summary_df, best_segments_df, opportunities_df)

        st.download_button(
            label="Download analysis as Excel",
            data=excel_bytes,
            file_name="recharge_rate_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if not best_segments_df.empty:
            st.download_button(
                label="Download best rates as CSV",
                data=best_segments_df.to_csv(index=False).encode("utf-8"),
                file_name="best_rates.csv",
                mime="text/csv",
            )

        if not opportunities_df.empty:
            st.download_button(
                label="Download opportunities as CSV",
                data=opportunities_df.to_csv(index=False).encode("utf-8"),
                file_name="buy_sell_opportunities.csv",
                mime="text/csv",
            )
else:
    st.info("Upload supplier Excel files to begin.")