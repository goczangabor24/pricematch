import io
import math
import re
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Table Matcher", page_icon="✅", layout="wide")


# ---------------------------
# Helpers
# ---------------------------

def parse_eu_number(value) -> Optional[float]:
    if value is None:
        return None

    s = str(value).strip()
    if not s:
        return None

    s = s.replace("\u00a0", " ")
    s = s.replace("€", "")
    s = re.sub(r"\bEUR\b", "", s, flags=re.IGNORECASE)
    s = s.strip().replace(" ", "")

    # Keep only digits, comma, dot, minus
    s = re.sub(r"[^0-9,.\-]", "", s)
    if not s:
        return None

    if "," in s and "." in s:
        # 1.234,56
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        # 1,234.56
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2:
            decimal_part = parts[-1]
            int_part = "".join(parts[:-1])
            if len(decimal_part) in (1, 2, 3):
                s = f"{int_part}.{decimal_part}"
            else:
                s = "".join(parts)

    try:
        return float(s)
    except Exception:
        return None


def format_eu_number(value: Optional[float], decimals: int = 2) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return f"{value:.{decimals}f}".replace(".", ",")


def normalize_code(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def read_main_table(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()

    if name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(uploaded_file)
        return df

    raw = uploaded_file.read()

    # Try common encodings / separators
    attempts = [
        {"encoding": "utf-16", "sep": "\t"},
        {"encoding": "utf-8", "sep": "\t"},
        {"encoding": "utf-8-sig", "sep": "\t"},
        {"encoding": "latin1", "sep": "\t"},
        {"encoding": "utf-16", "sep": ","},
        {"encoding": "utf-8", "sep": ","},
        {"encoding": "utf-8-sig", "sep": ","},
        {"encoding": "latin1", "sep": ","},
        {"encoding": "utf-16", "sep": ";"},
        {"encoding": "utf-8", "sep": ";"},
        {"encoding": "utf-8-sig", "sep": ";"},
        {"encoding": "latin1", "sep": ";"},
    ]

    last_error = None
    for attempt in attempts:
        try:
            return pd.read_csv(
                io.BytesIO(raw),
                encoding=attempt["encoding"],
                sep=attempt["sep"]
            )
        except Exception as e:
            last_error = e

    raise ValueError(f"Could not read main file. Last error: {last_error}")


def read_reference_table(uploaded_file, pasted_text: str) -> pd.DataFrame:
    if uploaded_file is not None:
        name = uploaded_file.name.lower()

        if name.endswith(".xlsx") or name.endswith(".xls"):
            df = pd.read_excel(uploaded_file, header=None)
        else:
            raw = uploaded_file.read()
            attempts = [
                {"encoding": "utf-16", "sep": "\t"},
                {"encoding": "utf-8", "sep": "\t"},
                {"encoding": "utf-8-sig", "sep": "\t"},
                {"encoding": "latin1", "sep": "\t"},
                {"encoding": "utf-16", "sep": ","},
                {"encoding": "utf-8", "sep": ","},
                {"encoding": "utf-8-sig", "sep": ","},
                {"encoding": "latin1", "sep": ","},
                {"encoding": "utf-16", "sep": ";"},
                {"encoding": "utf-8", "sep": ";"},
                {"encoding": "utf-8-sig", "sep": ";"},
                {"encoding": "latin1", "sep": ";"},
            ]

            df = None
            last_error = None
            for attempt in attempts:
                try:
                    df = pd.read_csv(
                        io.BytesIO(raw),
                        encoding=attempt["encoding"],
                        sep=attempt["sep"],
                        header=None
                    )
                    break
                except Exception as e:
                    last_error = e

            if df is None:
                raise ValueError(f"Could not read reference file. Last error: {last_error}")
    else:
        lines = [line.strip() for line in pasted_text.splitlines() if line.strip()]
        rows = []
        for line in lines:
            parts = re.split(r"\t|;|,", line, maxsplit=1)
            if len(parts) >= 2:
                rows.append(parts[:2])

        df = pd.DataFrame(rows)

    if df.shape[1] < 2:
        raise ValueError("Reference input must contain at least 2 columns.")

    df = df.iloc[:, :2].copy()
    df.columns = ["ref_code", "ref_value"]
    return df


def find_best_match(target: float, d: Optional[float], f: Optional[float], g: Optional[float], tolerance: float):
    candidates = []

    if f is not None:
        candidates.append(("F", f))
    if g is not None:
        candidates.append(("G", g))
    if d is not None and f is not None:
        candidates.append(("D*F", d * f))
    if d is not None and g is not None:
        candidates.append(("D*G", d * g))

    if not candidates:
        return {
            "exact": False,
            "exact_formula": "",
            "exact_value": None,
            "closest_formula": "",
            "closest_value": None,
            "difference": None,
        }

    for formula, value in candidates:
        if abs(value - target) <= tolerance:
            return {
                "exact": True,
                "exact_formula": formula,
                "exact_value": value,
                "closest_formula": formula,
                "closest_value": value,
                "difference": 0.0,
            }

    closest_formula, closest_value = min(candidates, key=lambda x: abs(x[1] - target))
    diff = abs(closest_value - target)

    return {
        "exact": False,
        "exact_formula": "",
        "exact_value": None,
        "closest_formula": closest_formula,
        "closest_value": closest_value,
        "difference": diff,
    }


def build_results(main_df: pd.DataFrame, ref_df: pd.DataFrame, tolerance: float) -> pd.DataFrame:
    df = main_df.copy()

    if df.shape[1] < 7:
        raise ValueError("Main table must contain at least 7 columns so A, B, D, F, G exist.")

    # Fixed positional columns
    col_a = df.columns[0]
    col_b = df.columns[1]
    col_d = df.columns[3]
    col_f = df.columns[5]
    col_g = df.columns[6]

    df["_A_code"] = df[col_a].apply(normalize_code)
    df["_B_code"] = df[col_b].apply(normalize_code)
    df["_D_num"] = df[col_d].apply(parse_eu_number)
    df["_F_num"] = df[col_f].apply(parse_eu_number)
    df["_G_num"] = df[col_g].apply(parse_eu_number)

    results = []

    for _, ref_row in ref_df.iterrows():
        ref_code = normalize_code(ref_row["ref_code"])
        ref_value_raw = ref_row["ref_value"]
        ref_value_num = parse_eu_number(ref_value_raw)

        matches_a = df[df["_A_code"] == ref_code]
        matches_b = df[df["_B_code"] == ref_code]
        matches = pd.concat([matches_a, matches_b]).drop_duplicates()

        if matches.empty:
            results.append({
                "reference_code": ref_code,
                "reference_value": str(ref_value_raw),
                "found": "No",
                "found_in": "",
                "exact_match": "",
                "matched_formula": "",
                "closest_formula": "",
                "closest_value": "",
                "difference": "",
                "A_value": "",
                "B_value": "",
                "D_value": "",
                "F_value": "",
                "G_value": "",
            })
            continue

        best_row_result = None
        best_main_row = None
        best_found_in = ""

        for main_idx, main_row in matches.iterrows():
            found_in_list = []
            if normalize_code(main_row["_A_code"]) == ref_code:
                found_in_list.append("A")
            if normalize_code(main_row["_B_code"]) == ref_code:
                found_in_list.append("B")
            found_in = "/".join(found_in_list)

            if ref_value_num is None:
                comparison = {
                    "exact": False,
                    "exact_formula": "",
                    "exact_value": None,
                    "closest_formula": "",
                    "closest_value": None,
                    "difference": None,
                }
            else:
                comparison = find_best_match(
                    target=ref_value_num,
                    d=main_row["_D_num"],
                    f=main_row["_F_num"],
                    g=main_row["_G_num"],
                    tolerance=tolerance
                )

            if best_row_result is None:
                best_row_result = comparison
                best_main_row = main_row
                best_found_in = found_in
            else:
                current_diff = comparison["difference"]
                best_diff = best_row_result["difference"]

                if comparison["exact"] and not best_row_result["exact"]:
                    best_row_result = comparison
                    best_main_row = main_row
                    best_found_in = found_in
                elif comparison["exact"] == best_row_result["exact"]:
                    if current_diff is not None and best_diff is not None and current_diff < best_diff:
                        best_row_result = comparison
                        best_main_row = main_row
                        best_found_in = found_in

        results.append({
            "reference_code": ref_code,
            "reference_value": str(ref_value_raw),
            "found": "Yes",
            "found_in": best_found_in,
            "exact_match": "✓" if best_row_result["exact"] else "",
            "matched_formula": best_row_result["exact_formula"],
            "closest_formula": best_row_result["closest_formula"],
            "closest_value": format_eu_number(best_row_result["closest_value"]),
            "difference": format_eu_number(best_row_result["difference"]) if best_row_result["difference"] is not None else "",
            "A_value": normalize_code(best_main_row[col_a]),
            "B_value": normalize_code(best_main_row[col_b]),
            "D_value": format_eu_number(best_main_row["_D_num"]) if best_main_row["_D_num"] is not None else "",
            "F_value": format_eu_number(best_main_row["_F_num"]) if best_main_row["_F_num"] is not None else "",
            "G_value": format_eu_number(best_main_row["_G_num"]) if best_main_row["_G_num"] is not None else "",
        })

    return pd.DataFrame(results)


def to_tsv(df: pd.DataFrame) -> str:
    return df.to_csv(sep="\t", index=False)


# ---------------------------
# UI
# ---------------------------

st.title("Code + Value Matcher")
st.write(
    "Upload the main table and a 2-column reference table. "
    "The app checks whether the reference code exists in column A or B of the main table, "
    "then compares the reference value against F, G, D*F, and D*G."
)

with st.sidebar:
    st.header("Settings")
    tolerance = st.number_input(
        "Matching tolerance",
        min_value=0.0,
        value=0.01,
        step=0.01,
        help="Two values are treated as equal if their difference is within this tolerance."
    )

st.markdown("### 1. Upload main table")
main_file = st.file_uploader(
    "Main table file",
    type=["csv", "tsv", "txt", "xlsx", "xls"],
    key="main_file"
)

st.markdown("### 2. Upload or paste reference table")
reference_file = st.file_uploader(
    "Reference table file (2 columns)",
    type=["csv", "tsv", "txt", "xlsx", "xls"],
    key="ref_file"
)

reference_text = st.text_area(
    "Or paste the 2-column reference table here",
    height=180,
    placeholder="61118\t60,00\n61132\t60,00\n61146\t60,00"
)

run = st.button("Run comparison", type="primary")

if run:
    if main_file is None:
        st.error("Please upload the main table.")
        st.stop()

    if reference_file is None and not reference_text.strip():
        st.error("Please upload or paste the reference table.")
        st.stop()

    try:
        main_df = read_main_table(main_file)
        ref_df = read_reference_table(reference_file, reference_text)

        result_df = build_results(main_df, ref_df, tolerance=tolerance)

        st.markdown("### 3. Result")
        st.dataframe(result_df, use_container_width=True)

        tsv_output = to_tsv(result_df)

        st.markdown("### 4. TSV output")
        st.text_area("Copy TSV", value=tsv_output, height=250)

        st.download_button(
            "Download result TSV",
            data=tsv_output.encode("utf-8"),
            file_name="comparison_result.tsv",
            mime="text/tab-separated-values"
        )

    except Exception as e:
        st.error(str(e))