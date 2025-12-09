import io
import os
import re
from datetime import date, datetime

import pandas as pd
import requests
import streamlit as st

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ============================================================
# GOOGLE MAPS API KEY
# ============================================================
# 🔴 IMPORTANT: put your REAL working key here
API_KEY = "AIzaSyAQ6JK-XEvoZfvNu7rAeFSVMTnLmBxmK7E"

# ============================================================
# CONFIG: static files for downloads (update paths if needed)
# ============================================================

PYTHON_FILES = {
    "Address Validation Notebook": "/Users/sreenidhikommineni/Downloads/Address_Validation.ipynb",
    "Coupon Generation Notebook": "/Users/sreenidhikommineni/Downloads/CouponGeneration.ipynb",
}

WORD_TEMPLATES = {
    # used on Refunds page
    "Refund Check Template": "/Users/sreenidhikommineni/Library/CloudStorage/OneDrive-FlexDay/Chase Temp Bulk.docx",
    # used on Letters & Labels page
    "Letter Template": "/Users/sreenidhikommineni/Library/CloudStorage/OneDrive-FlexDay/DM-Letters Temp.docx",
    "Label Template": "/Users/sreenidhikommineni/Library/CloudStorage/OneDrive-FlexDay/Labels-Temp.docx",
}

# include both possible ZIP spellings; we’ll keep whichever exists
REQUIRED_COUPON_COLS = [
    "Ticket ID", "Created Time", "Enclosures",
    "Consumer Full Name", "Consumer Street Address",
    "Consumer City", "Consumer State",
    "Consumer ZIP Code", "Consumer Zip Code",
    "No of coupons", "Coupons"
]

REQUIRED_REFUND_COLS = [
    "Ticket ID", "Created Time", "Enclosures",
    "Consumer Full Name", "Consumer Street Address",
    "Consumer City", "Consumer State",
    "Consumer ZIP Code", "Consumer Zip Code",
    "Refund Amount"
]

# ============================================================
# Helpers: files & Excel
# ============================================================

def get_file_bytes(path: str):
    """Read a file in binary mode, return bytes or None if not found."""
    if not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return f.read()


def to_excel_bytes(sheets_dict: dict):
    """sheets_dict = {sheet_name: dataframe} -> in-memory Excel file."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output


# ============================================================
# ADDRESS VALIDATION (Google, flexible + ZIP fix)
# ============================================================

def normalize_street(street: str) -> str:
    """Normalize street address for comparison."""
    street = str(street).lower().strip()
    street = re.sub(r'\s+', ' ', street)           # Remove extra spaces
    street = re.sub(r'[.,#]', '', street)          # Remove punctuation

    abbrev_map = {
        r'\bstreet\b': 'st',
        r'\bavenue\b': 'ave',
        r'\bdrive\b': 'dr',
        r'\broad\b': 'rd',
        r'\bboulevard\b': 'blvd',
        r'\blane\b': 'ln',
        r'\bcourt\b': 'ct',
        r'\bcircle\b': 'cir',
        r'\bparkway\b': 'pkwy',
        r'\bplace\b': 'pl',
        r'\bnorth\b': 'n',
        r'\bsouth\b': 's',
        r'\beast\b': 'e',
        r'\bwest\b': 'w',
        r'\bnortheast\b': 'ne',
        r'\bnorthwest\b': 'nw',
        r'\bsoutheast\b': 'se',
        r'\bsouthwest\b': 'sw',
    }
    for full, abbrev in abbrev_map.items():
        street = re.sub(full, abbrev, street)
    return street


STATE_ABBREV = {
    'alabama': 'al', 'alaska': 'ak', 'arizona': 'az', 'arkansas': 'ar',
    'california': 'ca', 'colorado': 'co', 'connecticut': 'ct', 'delaware': 'de',
    'florida': 'fl', 'georgia': 'ga', 'hawaii': 'hi', 'idaho': 'id',
    'illinois': 'il', 'indiana': 'in', 'iowa': 'ia', 'kansas': 'ks',
    'kentucky': 'ky', 'louisiana': 'la', 'maine': 'me', 'maryland': 'md',
    'massachusetts': 'ma', 'michigan': 'mi', 'minnesota': 'mn', 'mississippi': 'ms',
    'missouri': 'mo', 'montana': 'mt', 'nebraska': 'ne', 'nevada': 'nv',
    'new hampshire': 'nh', 'new jersey': 'nj', 'new mexico': 'nm', 'new york': 'ny',
    'north carolina': 'nc', 'north dakota': 'nd', 'ohio': 'oh', 'oklahoma': 'ok',
    'oregon': 'or', 'pennsylvania': 'pa', 'rhode island': 'ri', 'south carolina': 'sc',
    'south dakota': 'sd', 'tennessee': 'tn', 'texas': 'tx', 'utah': 'ut',
    'vermont': 'vt', 'virginia': 'va', 'washington': 'wa', 'west virginia': 'wv',
    'wisconsin': 'wi', 'wyoming': 'wy', 'district of columbia': 'dc'
}


def normalize_state(state: str) -> str:
    """Convert state to 2-letter abbreviation if possible."""
    state = str(state).lower().strip()
    if len(state) == 2:
        return state
    return STATE_ABBREV.get(state, state)


def validate_address(street, city, state, zipcode, api_key: str):
    """Call Google Geocoding API to validate one address."""
    full_address = f"{street}, {city}, {state} {zipcode}"
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {"address": full_address, "key": api_key}

    try:
        response = requests.get(url, params=params, timeout=10).json()
    except Exception as e:
        return {
            "valid": "error",
            "google_address": "",
            "mistake": f"API Error: {e}"
        }

    status = response.get("status")
    results = response.get("results", [])

    if status != "OK" or not results:
        msg = response.get("error_message", "Address not found by Google")

        # treat true API problems as errors
        if status in ("REQUEST_DENIED", "OVER_QUERY_LIMIT", "UNKNOWN_ERROR"):
            return {
                "valid": "error",
                "google_address": "",
                "mistake": f"{status}: {msg}"
            }

        # ZERO_RESULTS etc. are real invalid addresses
        return {
            "valid": "no",
            "google_address": "",
            "mistake": f"{status}: {msg}"
        }

    result = results[0]
    formatted_address = result.get("formatted_address", "")

    # Extract components (keep both long & short for state)
    comps_long = {}
    comps_short = {}
    for comp in result["address_components"]:
        for t in comp["types"]:
            comps_long[t] = comp.get("long_name", "")
            comps_short[t] = comp.get("short_name", "")

    google_street_number = comps_long.get("street_number", "").lower()
    google_route = comps_long.get("route", "").lower()
    google_city = comps_long.get("locality", "").lower()
    google_state_long = comps_long.get("administrative_area_level_1", "").lower()
    google_state_short = comps_short.get("administrative_area_level_1", "").lower()
    google_zip_raw = comps_long.get("postal_code", "").lower()

    # Normalize ZIP: keep only digits; use 5-digit form
    google_zip_digits = re.sub(r"[^0-9]", "", google_zip_raw)
    google_zip_5 = google_zip_digits[:5] if google_zip_digits else ""

    zipcode_str = str(zipcode)
    input_zip_digits = re.sub(r"[^0-9]", "", zipcode_str)
    input_zip_5 = input_zip_digits[:5] if input_zip_digits else ""

    # Build & normalize street
    google_full_street = f"{google_street_number} {google_route}".strip()
    google_normalized = normalize_street(google_full_street)
    input_normalized = normalize_street(street)

    # Normalize city/state
    input_city = str(city).lower().strip()
    input_state_norm = normalize_state(state)
    google_state_norm = normalize_state(google_state_long or google_state_short)

    mistakes = []

    # 1. Street match
    if google_normalized != input_normalized:
        mistakes.append(f"Street mismatch: '{street}' vs '{google_full_street}'")

    # 2. City
    if google_city != input_city:
        mistakes.append(f"City mismatch: '{city}' vs '{google_city}'")

    # 3. State
    if google_state_norm != input_state_norm:
        mistakes.append(f"State mismatch: '{state}' vs '{google_state_long or google_state_short}'")

    # 4. ZIP (allow ZIP+4; use 5-digit comparison)
    if input_zip_5 and google_zip_5 and google_zip_5 != input_zip_5:
        mistakes.append(f"ZIP mismatch: '{zipcode}' vs '{google_zip_raw}'")

    if mistakes:
        return {
            "valid": "no",
            "google_address": formatted_address,
            "mistake": "; ".join(mistakes)
        }
    else:
        return {
            "valid": "yes",
            "google_address": formatted_address,
            "mistake": ""
        }


def validate_addresses_dataframe(df: pd.DataFrame, api_key: str):
    """
    Run Google-based address validation for every row in df.
    Returns:
      df_with_cols, results_df
    """
    df = df.copy()

    # choose zip column (support both names)
    if "Consumer ZIP Code" in df.columns:
        zip_col = "Consumer ZIP Code"
    elif "Consumer Zip Code" in df.columns:
        zip_col = "Consumer Zip Code"
    else:
        zip_col = None

    results = []
    for idx, row in df.iterrows():
        street = row.get("Consumer Street Address", "")
        city = row.get("Consumer City", "")
        state = row.get("Consumer State", "")
        zipcode = row.get(zip_col, "") if zip_col else ""

        ticket_id = row.get("Ticket ID", f"Row_{idx}")
        input_address = f"{street}, {city}, {state} {zipcode}"

        result = validate_address(street, city, state, zipcode, api_key)

        results.append({
            "Ticket ID": ticket_id,
            "Input Address": input_address,
            "Valid": result["valid"],
            "Google Address": result["google_address"],
            "Mistake": result["mistake"],
        })

    results_df = pd.DataFrame(results)

    df["Address Valid"] = results_df["Valid"]
    df["Google Address"] = results_df["Google Address"]
    df["Validation Issue"] = results_df["Mistake"]

    return df, results_df


# ============================================================
# COUPON PDF GENERATION
# ============================================================

PAGE_WIDTH, PAGE_HEIGHT = A4
COUPONS_PER_PAGE = 4
BARCODE_WIDTH = 1.46 * inch
BARCODE_HEIGHT = 1.11 * inch

COUPON_LAYOUTS = {
    0: {
        "expiry_x": 7.7 * inch, "expiry_y": 11.58 * inch,
        "code_x": 2.50 * inch,  "code_y": 10.90 * inch,
        "name_x": 2.30 * inch,  "name_y": 10.55 * inch,
        "address_x": 2.50 * inch, "address_y": 10.25 * inch,
        "max_retail_x": 7.85 * inch, "max_retail_y": 10.47 * inch,
        "barcode_x": 6.0 * inch, "barcode_y": 9.10 * inch,
    },
    1: {
        "expiry_x": 7.7 * inch, "expiry_y": 8.54 * inch,
        "code_x": 2.50 * inch,  "code_y": 7.90 * inch,
        "name_x": 2.30 * inch,  "name_y": 7.55 * inch,
        "address_x": 2.50 * inch, "address_y": 7.25 * inch,
        "max_retail_x": 7.85 * inch, "max_retail_y": 7.50 * inch,
        "barcode_x": 6.0 * inch, "barcode_y": 6.30 * inch,
    },
    2: {
        "expiry_x": 7.7 * inch, "expiry_y": 5.55 * inch,
        "code_x": 2.50 * inch,  "code_y": 4.75 * inch,
        "name_x": 2.30 * inch,  "name_y": 4.40 * inch,
        "address_x": 2.51 * inch, "address_y": 4.10 * inch,
        "max_retail_x": 7.85 * inch, "max_retail_y": 4.50 * inch,
        "barcode_x": 6.0 * inch, "barcode_y": 3.20 * inch,
    },
    3: {
        "expiry_x": 7.7 * inch, "expiry_y": 2.55 * inch,
        "code_x": 2.50 * inch,  "code_y": 1.85 * inch,
        "name_x": 2.30 * inch,  "name_y": 1.50 * inch,
        "address_x": 2.51 * inch, "address_y": 1.20 * inch,
        "max_retail_x": 7.85 * inch, "max_retail_y": 1.47 * inch,
        "barcode_x": 6.0 * inch, "barcode_y": 0.25 * inch,
    },
}

# register font
try:
    pdfmetrics.registerFont(
        TTFont("Aptos-Bold", "/System/Library/Fonts/Supplemental/Arial Bold.ttf")
    )
    FONT_NAME = "Aptos-Bold"
except Exception:
    FONT_NAME = "Helvetica-Bold"


def build_barcodes_lookup(barcodes_df: pd.DataFrame):
    lookup = {}
    for _, row in barcodes_df.iterrows():
        name = str(row.get("Coupons", row.get("NAME", ""))).strip()
        if not name:
            continue
        lookup[name] = {
            "max_retail": row.get("MAX RETAIL", ""),
            "expiry": row.get("Expiry", ""),
            "barcode_path": str(row.get("Barcode Path", "")).strip().strip("'"),
        }
    return lookup


def draw_coupon(c, cust_row, coupon_name, idx_on_page, barcodes_lookup):
    layout = COUPON_LAYOUTS.get(idx_on_page, COUPON_LAYOUTS[0])
    info = barcodes_lookup.get(coupon_name, {})

    expiry = info.get("expiry", "")
    max_retail = info.get("max_retail", "")
    barcode_path = info.get("barcode_path", "")

    if isinstance(expiry, (datetime, pd.Timestamp)):
        expiry_str = expiry.strftime("%m/%d/%Y")
    else:
        expiry_str = str(expiry).split(" ")[0]
        try:
            expiry_str = datetime.strptime(expiry_str, "%Y-%m-%d").strftime("%m/%d/%Y")
        except Exception:
            pass

    coupon_code = str(cust_row.get("Ticket ID", "")).strip()

    # pick ZIP column for the address line
    if "Consumer ZIP Code" in cust_row.index:
        zip_val = cust_row.get("Consumer ZIP Code", "")
    else:
        zip_val = cust_row.get("Consumer Zip Code", "")

    address_lines = [
        cust_row.get("Consumer Full Name", ""),
        cust_row.get("Consumer Street Address", ""),
        f"{cust_row.get('Consumer City', '')}, {cust_row.get('Consumer State', '')} {zip_val}",
    ]

    c.setFont(FONT_NAME, 7.5)

    c.drawRightString(layout["expiry_x"], layout["expiry_y"], f"Expires: {expiry_str}")
    # Ticket ID only (no Tags)
    c.drawString(layout["code_x"], layout["code_y"], coupon_code)
    c.drawString(layout["name_x"], layout["name_y"], coupon_name)

    text_obj = c.beginText(layout["address_x"], layout["address_y"])
    text_obj.setFont(FONT_NAME, 7.5)
    for line in address_lines:
        text_obj.textLine(str(line))
    c.drawText(text_obj)

    c.drawRightString(
        layout["max_retail_x"],
        layout["max_retail_y"],
        f"Max Retail Value: ${max_retail}",
    )

    if barcode_path and os.path.exists(barcode_path):
        c.drawImage(
            barcode_path,
            layout["barcode_x"],
            layout["barcode_y"],
            width=BARCODE_WIDTH,
            height=BARCODE_HEIGHT,
            preserveAspectRatio=True,
        )
    else:
        c.drawString(layout["barcode_x"], layout["barcode_y"], "[Barcode not found]")


def generate_coupon_pdf(coupons_df: pd.DataFrame, barcodes_df: pd.DataFrame) -> bytes:
    """Create coupons PDF and return as bytes."""
    barcodes_lookup = build_barcodes_lookup(barcodes_df)
    df = coupons_df.copy()

    # unify ZIP + number-of-coupons column
    if "Consumer ZIP Code" not in df.columns and "Consumer Zip Code" in df.columns:
        df["Consumer ZIP Code"] = df["Consumer Zip Code"]
    if "Number of Coupons" not in df.columns and "No of coupons" in df.columns:
        df["Number of Coupons"] = df["No of coupons"]

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    coupon_counter = 0

    for _, cust in df.iterrows():
        raw_names = str(cust.get("Coupons", "")).strip()
        if not raw_names or raw_names.lower() == "nan":
            continue

        no_of_coupons = cust.get("Number of Coupons", 1)
        try:
            if pd.isna(no_of_coupons):
                num_coupons = 1
            else:
                num_coupons = int(float(no_of_coupons))
        except (ValueError, TypeError):
            num_coupons = 1

        coupon_names = [n.strip() for n in raw_names.split(",") if n.strip()]

        for cname in coupon_names:
            for _ in range(num_coupons):
                idx_on_page = coupon_counter % COUPONS_PER_PAGE
                draw_coupon(c, cust, cname, idx_on_page, barcodes_lookup)
                coupon_counter += 1

                if coupon_counter % COUPONS_PER_PAGE == 0:
                    c.showPage()

    c.save()
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# PAGES
# ============================================================

def show_overview():
    st.title("Del Monte Fulfillment – Overview")

    st.markdown("### What this app is for")
    st.write("""
    This app documents and supports the **Coupons**, **Refunds**, **Letters & Labels**,
    and **Envelope & Mailing** processes for Del Monte.
    """)

    st.markdown("### Main pages")
    st.write("""
    - **Coupons Process** – work only with coupon cases  
    - **Refunds Process** – work only with refund cases  
    - **Letters & Labels** – how to generate letters & address labels  
    - **Envelope & Mailing** – how to package and send mail  
    """)

    st.markdown("### High-level workflow")
    st.write("""
    1. Export tickets from **Freshdesk** (CSV/Excel).
    2. Go to **Coupons Process** or **Refunds Process**:
       - Upload the export.
       - Filter & clean data.
       - Run address validation.
       - Download ready-to-use Excel/PDF files.
    3. Use:
       - Coupons Excel → coupon PDF generator & printing.
       - Refunds Excel → Word mail merge for checks & letters.
    4. Use **Letters & Labels** page instructions to prepare letters and labels.
    5. Use **Envelope & Mailing** page for physical mailing steps.
    """)

    st.markdown("---")
    st.markdown("## ⬇️ Download Python Notebooks")

    for label, path in PYTHON_FILES.items():
        data = get_file_bytes(path)
        if data is None:
            st.warning(f"File not found on this machine: `{path}`.")
        else:
            st.download_button(
                f"Download {label}",
                data=data,
                file_name=os.path.basename(path),
                mime="application/octet-stream",
            )


def show_coupons_process():
    st.title("🎫 Coupons Fulfillment Process")

    st.markdown("## 1. Upload Freshdesk Export (Coupons)")
    uploaded = st.file_uploader(
        "Upload Freshdesk CSV/Excel for Coupons",
        type=["csv", "xlsx", "xls"],
        key="coupons_upload"
    )

    st.markdown("### Coupon Barcodes Excel")
    barcodes_file = st.file_uploader(
        "Upload Coupon Barcodes Excel (with Coupons / MAX RETAIL / Expiry / Barcode Path)",
        type=["xlsx", "xls"],
        key="barcodes_upload"
    )

    if uploaded is not None:
        if not API_KEY or "YOUR_REAL_GOOGLE_MAPS_API_KEY_HERE" in API_KEY:
            st.error("API_KEY is not set correctly. Please add your real Google Maps API key at the top of app.py.")
            return

        try:
            if uploaded.name.lower().endswith(".csv"):
                df_raw = pd.read_csv(uploaded)
            else:
                df_raw = pd.read_excel(uploaded)
        except Exception as e:
            st.error(f"Could not read file: {e}")
            return

        st.subheader("Preview of uploaded data")
        st.dataframe(df_raw.head())

        st.markdown("### Choose the **Enclosures** column & Filter for Coupons")
        enclosure_col = st.selectbox(
            "Column",
            options=df_raw.columns.tolist(),
            index=(df_raw.columns.tolist().index("Enclosures")
                   if "Enclosures" in df_raw.columns else 0),
            key="coupons_enclosures_col"
        )

        col_values = df_raw[enclosure_col].dropna().astype(str).unique().tolist()
        col_values_sorted = sorted(col_values)

        coupon_value = st.selectbox(
            "Filter - **Coupons**",
            options=col_values_sorted,
            index=0,
            key="coupon_value_select"
        )

        if st.button("Process & Validate Coupons", key="process_coupons_btn"):
            with st.spinner("Filtering coupons and validating addresses..."):
                df_coupons = df_raw[df_raw[enclosure_col].astype(str) == str(coupon_value)].copy()
                coupon_cols_present = [c for c in REQUIRED_COUPON_COLS if c in df_coupons.columns]
                df_coupons = df_coupons[coupon_cols_present]

                # unify ZIP name
                if "Consumer ZIP Code" not in df_coupons.columns and "Consumer Zip Code" in df_coupons.columns:
                    df_coupons["Consumer ZIP Code"] = df_coupons["Consumer Zip Code"]

                df_coupons, results_df = validate_addresses_dataframe(df_coupons, API_KEY)

                st.session_state["coupons_df"] = df_coupons  # for PDF step

                st.success("✅ Coupons processed!")

                num_rows = len(results_df)
                valid_count = (results_df["Valid"] == "yes").sum()
                invalid_count = (results_df["Valid"] == "no").sum()
                error_count = (results_df["Valid"] == "error").sum()

                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("Total coupon rows", num_rows)
                with c2:
                    st.metric("Valid addresses", int(valid_count))
                with c3:
                    st.metric("Invalid addresses", int(invalid_count))
                with c4:
                    st.metric("Errors", int(error_count))

                st.markdown("### Validation Results (first 30)")
                st.dataframe(results_df.head(30))

                invalid_df = results_df[results_df["Valid"] == "no"]
                if not invalid_df.empty:
                    st.markdown("### ❗ Invalid Addresses (need fixing)")
                    st.dataframe(invalid_df)
                else:
                    st.success("All coupon addresses look valid based on Google responses.")

                today_str = date.today().strftime("%Y%m%d")
                coupons_bytes = to_excel_bytes({"Coupons": df_coupons})
                st.download_button(
                    "⬇️ Download Coupons Excel (with validation)",
                    data=coupons_bytes,
                    file_name=f"Coupons_prepared_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    st.markdown("---")
    st.markdown("## 2. Generate Coupon PDFs (from validated data)")

    if st.button("Run Coupon PDF Generator", key="run_coupon_generator_btn"):
        df_coupons = st.session_state.get("coupons_df")
        if df_coupons is None:
            st.error("Please process coupons first (above) before generating PDFs.")
        elif barcodes_file is None:
            st.error("Please upload the Coupon Barcodes Excel file first.")
        else:
            try:
                barcodes_df = pd.read_excel(barcodes_file, sheet_name="Barcodes")
            except Exception as e:
                st.error(f"Could not read Barcodes sheet: {e}")
            else:
                with st.spinner("Generating coupon PDF..."):
                    pdf_bytes = generate_coupon_pdf(df_coupons, barcodes_df)
                today_str = date.today().strftime("%Y%m%d")
                st.success("✅ Coupon PDF generated.")
                st.download_button(
                    "⬇️ Download Coupons PDF",
                    data=pdf_bytes,
                    file_name=f"Coupons_{today_str}.pdf",
                    mime="application/pdf",
                )

    st.markdown("---")
    st.markdown("## 3. Process Explanation (Coupons)")
    st.write("""
    1. Export Freshdesk tickets for desired date range.  
    2. Check the Names and Address and alsd for any misplacemnt of Lower and Upper Case
    3. Upload the exported file here.  
    4. Select the Enclosures column and the value for **Coupons**.  
    5. Click **Process & Validate Coupons**   
    6. Check the invalid addresses and fix them if needed and re-run.  
    7. Upload the **Barcodes** Excel.  
    8. Click **Run Coupon PDF Generator** to create the final coupon PDF.  
    9. Print coupons on coupon paper.
    """)


def show_refunds_process():
    st.title("💸 Refunds Fulfillment Process")

    st.markdown("## 1. Upload Freshdesk Export (Refunds)")
    uploaded = st.file_uploader(
        "Upload Freshdesk CSV/Excel for Refunds",
        type=["csv", "xlsx", "xls"],
        key="refunds_upload"
    )

    if uploaded is not None:
        if not API_KEY or "YOUR_REAL_GOOGLE_MAPS_API_KEY_HERE" in API_KEY:
            st.error("API_KEY is not set correctly. Please add your real Google Maps API key at the top of app.py.")
            return

        try:
            if uploaded.name.lower().endswith(".csv"):
                df_raw = pd.read_csv(uploaded)
            else:
                df_raw = pd.read_excel(uploaded)
        except Exception as e:
            st.error(f"Could not read file: {e}")
            return

        st.subheader("Preview of uploaded data")
        st.dataframe(df_raw.head())

        st.markdown("### Choose the **Enclosures** column & Filter for Refunds")
        enclosure_col = st.selectbox(
            "Column",
            options=df_raw.columns.tolist(),
            index=(df_raw.columns.tolist().index("Enclosures")
                   if "Enclosures" in df_raw.columns else 0),
            key="refunds_enclosures_col"
        )

        col_values = df_raw[enclosure_col].dropna().astype(str).unique().tolist()
        col_values_sorted = sorted(col_values)

        refund_value = st.selectbox(
            "Filter - **Refunds**",
            options=col_values_sorted,
            index=0,
            key="refund_value_select"
        )

        if st.button("Process & Validate Refunds", key="process_refunds_btn"):
            with st.spinner("Filtering refunds and validating addresses..."):
                df_refunds = df_raw[df_raw[enclosure_col].astype(str) == str(refund_value)].copy()
                refund_cols_present = [c for c in REQUIRED_REFUND_COLS if c in df_refunds.columns]
                df_refunds = df_refunds[refund_cols_present]

                if "Consumer ZIP Code" not in df_refunds.columns and "Consumer Zip Code" in df_refunds.columns:
                    df_refunds["Consumer ZIP Code"] = df_refunds["Consumer Zip Code"]

                df_refunds, results_df = validate_addresses_dataframe(df_refunds, API_KEY)

                st.success("✅ Refunds processed!")

                num_rows = len(results_df)
                valid_count = (results_df["Valid"] == "yes").sum()
                invalid_count = (results_df["Valid"] == "no").sum()
                error_count = (results_df["Valid"] == "error").sum()

                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("Total refund rows", num_rows)
                with c2:
                    st.metric("Valid addresses", int(valid_count))
                with c3:
                    st.metric("Invalid addresses", int(invalid_count))
                with c4:
                    st.metric("Errors", int(error_count))

                st.markdown("### Validation Results (first 30)")
                st.dataframe(results_df.head(30))

                invalid_df = results_df[results_df["Valid"] == "no"]
                if not invalid_df.empty:
                    st.markdown("### ❗ Invalid Addresses (need fixing)")
                    st.dataframe(invalid_df)
                else:
                    st.success("All refund addresses look valid based on Google responses.")

                today_str = date.today().strftime("%Y%m%d")
                refunds_bytes = to_excel_bytes({"Refunds": df_refunds})
                st.download_button(
                    "⬇️ Download Refunds Excel (with validation)",
                    data=refunds_bytes,
                    file_name=f"Refunds_prepared_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    st.markdown("---")
    st.markdown("## 2. Download Refund Word Template")

    for label, path in WORD_TEMPLATES.items():
        if "Refund" not in label:
            continue
        data = get_file_bytes(path)
        if data is None:
            st.warning(f"Template not found: `{path}`.")
        else:
            st.download_button(
                f"Download {label}",
                data=data,
                file_name=os.path.basename(path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    st.markdown("## 3. Process Explanation (Refunds)")
    st.write("""
    1. Export Freshdesk tickets for your date range.
    2. Check the Names and Address and alsd for any misplacemnt of Lower and Upper Case
    3. Upload export here and process.  
    4. Download **Refunds_prepared** Excel.  
    5. Fix invalid addresses (Maps / Zillow) and re-run if needed.  
    6. In Word:
       - Open Refund Check Template.  
       - Add today's date.  
       - Mailings → Select Recipients → Use Existing List → pick Refunds Excel.  
       - Finish & Merge → Edit Individual Documents.  
    7. Manually type in the **Name**, **Refund Amount**, and **Amount in words**.  
    8. Load check stock and print.
    """)


def show_letters_labels_process():
    st.title("✉️ Letters & Labels Process")

    st.markdown("## 1. Download Letter & Label Templates")

    for label, path in WORD_TEMPLATES.items():
        if "Refund" in label:
            continue  # refund template is on refunds page
        data = get_file_bytes(path)
        if data is None:
            st.warning(f"Template not found: `{path}`.")
        else:
            st.download_button(
                f"Download {label}",
                data=data,
                file_name=os.path.basename(path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    st.markdown("## 2. How to use in Word (Mail Merge)")
    st.write("""
    **Letters:**
    1. Open **Letter Template** in Word.  
    2. Change the date to today.  
    3. For coupon cases → keep text about coupons.  
       For refund cases → change body to refer to refund check.  
    4. Mailings → Select Recipients → Use Existing List → choose:
       - `Coupons_prepared` for coupon letters, or  
       - `Refunds_prepared` for refund letters.  
    5. Finish & Merge → Edit Individual Documents → Print.  

    **Labels:**
    1. Open **Label Template** in Word.  
    2. Mailings → Select Recipients → Use Existing List → choose the same Excel
       (Coupons_prepared or Refunds_prepared).  
    3. Finish & Merge → Edit Individual Documents.  
    4. Print to the **thermal label printer**.
    """)


def show_envelope_filling_process():
    st.title("📦 Envelope Filling & Mailing Instructions")

    st.markdown("## 1. Outside the Envelope")
    st.write("""
    - **Top left corner:** Del Monte label sticker  
    - **Top right corner:** Postage stamp  
    - **Center:** Name & address label sticker (from thermal printer)
    """)

    st.markdown("## 2. Inside the Envelope")
    st.write("""
    - Fold the letter into **2/3rds**.  
    - Insert:
      - Letter  
      - Coupons (for coupon cases) **or**  
      - Refund check (for refund cases).  
    """)

    st.markdown("## 3. Final Steps")
    st.write("""
    - Seal envelopes.  
    - Drop them into the outgoing mail box.  
    - Go to Freshdesk and close the corresponding tickets.    
    """)


# ============================================================
# MAIN APP
# ============================================================

def main():
    st.set_page_config(page_title="Del Monte Fulfillment", layout="wide")

    st.sidebar.title("Navigation")
    page = st.sidebar.radio(
        "Go to:",
        (
            "Overview",
            "Coupons Process",
            "Refunds Process",
            "Letters & Labels",
            "Envelope Filling & Mailing",
        )
    )

    if page == "Overview":
        show_overview()
    elif page == "Coupons Process":
        show_coupons_process()
    elif page == "Refunds Process":
        show_refunds_process()
    elif page == "Letters & Labels":
        show_letters_labels_process()
    elif page == "Envelope Filling & Mailing":
        show_envelope_filling_process()


if __name__ == "__main__":
    main()
