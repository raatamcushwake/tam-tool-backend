from fastapi import APIRouter, UploadFile, File, HTTPException
import pandas as pd
import io
import logging
import difflib

logger = logging.getLogger(__name__)
router = APIRouter()

def get_similarity(a, b):
    return difflib.SequenceMatcher(None, str(a).lower(), str(b).lower()).ratio()

def clean_num(val):
    if pd.isna(val) or val == "" or str(val).strip().lower() in ["nan", "n/a", "-", "empty"]:
        return 0.0
    try:
        cleaned = "".join(c for c in str(val) if c.isdigit() or c in [".", "-"])
        return float(cleaned) if cleaned else 0.0
    except:
        return 0.0

def extract_mis_data(file_content):
    try:
        df_raw = pd.read_excel(io.BytesIO(file_content), header=None, engine='openpyxl')
        header_row_idx = 0
        for i in range(min(10, len(df_raw))):
            row_vals = [str(v).strip() for v in df_raw.iloc[i] if pd.notna(v)]
            if any("Unit No" in v or "Customer Name" in v for v in row_vals):
                header_row_idx = i
                break
        df = pd.read_excel(io.BytesIO(file_content), header=header_row_idx, engine='openpyxl')
        if df.empty:
            return pd.DataFrame()
        df = df.iloc[1:].reset_index(drop=True)

        standard_sequence = [
            "Unit No.", "Tower", "Booking Date", "Registration Date",
            "Unit Type", "Customer Name", "Saleable area in sft",
            "Carpet area in sft", "Agreement value",
            "Amount Received excl. Tax",
            "Demand Raised as on Current Month excl. tax"
        ]
        mapping = {}
        for std_name in standard_sequence:
            for col in df.columns:
                if str(col).strip().lower().startswith(std_name.lower()):
                    mapping[col] = std_name
                    break
        df = df[list(mapping.keys())].rename(columns=mapping)
        existing = [c for c in standard_sequence if c in df.columns]
        df = df[existing]
        if "Unit No." in df.columns:
            df = df.dropna(subset=["Unit No."])
            df["Unit No."] = df["Unit No."].astype(str).str.strip()
            df = df[df["Unit No."] != ""]
        date_cols = ["Booking Date", "Registration Date"]
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].apply(lambda x: x.strftime('%d %b %Y') if pd.notna(x) else '-')
        calc_cols = [
            "Agreement value", "Amount Received excl. Tax",
            "Demand Raised as on Current Month excl. tax",
            "Saleable area in sft", "Carpet area in sft"
        ]
        for col in calc_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(r'[^\d.]', '', regex=True),
                    errors='coerce'
                ).fillna(0)
        return df.fillna("").infer_objects(copy=False)
    except Exception as e:
        logger.error(f"Error extracting MIS data: {e}")
        return pd.DataFrame()

def process_comparison(prev_content, curr_content):
    try:
        df_prev = extract_mis_data(prev_content)
        df_curr = extract_mis_data(curr_content)

        if df_curr.empty:
            return {"status": "error", "message": "Current file is empty"}

        prev_lookup = {}
        if not df_prev.empty and "Unit No." in df_prev.columns:
            prev_lookup = df_prev.drop_duplicates("Unit No.").set_index("Unit No.").to_dict(orient='index')

        curr_lookup = {}
        if "Unit No." in df_curr.columns:
            curr_lookup = df_curr.drop_duplicates("Unit No.").set_index("Unit No.").to_dict(orient='index')

        final_rows = []
        sold_count = 0

        for _, row in df_curr.iterrows():
            unit_no = str(row.get("Unit No.", "")).strip()
            if not unit_no:
                continue

            unit_data = row.to_dict()
            curr_name_raw = str(unit_data.get("Customer Name", "")).strip()
            curr_name_lower = curr_name_raw.lower()
            is_sold = curr_name_lower not in ["", "nan", "n/a", "-", "unsold"]
            if is_sold:
                sold_count += 1

            # Compute financial fields
            agreement_val = clean_num(unit_data.get("Agreement value", 0))
            demand_raised = clean_num(unit_data.get("Demand Raised as on Current Month excl. tax", 0))
            amount_received = clean_num(unit_data.get("Amount Received excl. Tax", 0))
            saleable_area = clean_num(unit_data.get("Saleable area in sft", 0))

            rate_per_sft = (agreement_val / saleable_area) if saleable_area > 0 else 0
            os_against_demand = demand_raised - amount_received
            os_against_sale = agreement_val - amount_received
            os_pct_demand = (os_against_demand / demand_raised * 100) if demand_raised > 0 else (-1 if amount_received > 0 else 0)
            os_pct_sale = (os_against_sale / agreement_val * 100) if agreement_val > 0 else (-1 if amount_received > 0 else 0)

            # Disbursement
            try:
                booking_date = pd.to_datetime(unit_data.get("Booking Date", ""), errors='coerce')
                disbursement = "Pre" if pd.notna(booking_date) and booking_date < pd.Timestamp("2024-07-17") else "Post"
            except:
                disbursement = "N/A"

            status = "EXISTING"
            change_log = ""
            demand_inc = 0.0
            received_inc = 0.0
            agreement_inc = 0.0

            is_new = unit_no not in prev_lookup

            if not is_new:
                prev_unit = prev_lookup[unit_no]
                prev_name = str(prev_unit.get("Customer Name", "")).strip()
                prev_was_unsold = prev_name.lower() in ["", "nan", "n/a", "-", "unsold"]

                if curr_name_lower in ["unsold", "-", ""]:
                    status = "CANCELLATION"
                    change_log = f"[CANCELLATION] {prev_name}"
                elif prev_was_unsold and is_sold:
                    status = "NEW"
                    change_log = "[NEW BOOKING] Unit was previously UNSOLD"
                elif prev_name.lower() != curr_name_lower and is_sold:
                    similarity = get_similarity(prev_name, curr_name_raw)
                    if similarity >= 0.80:
                        status = "NAME_CORRECTION"
                        change_log = f"[NAME CORRECTION] {prev_name} → {curr_name_raw}"
                    else:
                        prev_exists_elsewhere = any(
                            str(r.get("Customer Name", "")).strip().lower() == prev_name.lower()
                            for u, r in curr_lookup.items() if u != unit_no
                        )
                        curr_existed_in_prev = any(
                            str(r.get("Customer Name", "")).strip().lower() == curr_name_lower
                            for u, r in prev_lookup.items() if u != unit_no
                        )
                        if prev_exists_elsewhere or curr_existed_in_prev:
                            status = "TRANSFER"
                            change_log = f"[TRANSFER] {prev_name} → {curr_name_raw}"
                        else:
                            status = "ANOMALY"
                            change_log = f"[ANOMALY] {prev_name} → {curr_name_raw}"

                # Value changes
                params = [
                    ("Agreement value", "agreement", "AGREEMENT_VALUE"),
                    ("Amount Received excl. Tax", "amount_received", "AMOUNT_RECEIVED_CHANGE"),
                    ("Demand Raised as on Current Month excl. tax", "demand", "DEMAND_RAISED_CHANGE"),
                ]
                for excel_key, prefix, change_status in params:
                    p_val = clean_num(prev_unit.get(excel_key, 0))
                    c_val = clean_num(unit_data.get(excel_key, 0))
                    delta = round(c_val - p_val, 2)
                    unit_data[f"prev_{prefix}"] = p_val
                    unit_data[f"{prefix}_delta"] = delta
                    if abs(delta) > 0.01:
                        action = "raised" if delta > 0 else "decreased"
                        log_entry = f"[{change_status}] {p_val:,.2f} → {c_val:,.2f} ({action} by {abs(delta):,.2f})"
                        if status == "EXISTING":
                            status = change_status
                        change_log = change_log + f" | {log_entry}" if change_log else log_entry
                        if prefix == "demand":
                            demand_inc = delta
                        elif prefix == "amount_received":
                            received_inc = delta
                        elif prefix == "agreement":
                            agreement_inc = delta

                # Saleable / Carpet deltas
                for excel_key, prefix in [("Saleable area in sft", "saleable"), ("Carpet area in sft", "carpet")]:
                    p_val = clean_num(prev_unit.get(excel_key, 0))
                    c_val = clean_num(unit_data.get(excel_key, 0))
                    delta = round(c_val - p_val, 2)
                    unit_data[f"prev_{prefix}"] = p_val
                    unit_data[f"{prefix}_delta"] = delta

            else:
                if is_sold:
                    status = "NEW"
                    change_log = "[NEW BOOKING]"
                demand_inc = demand_raised
                received_inc = amount_received
                for prefix in ["agreement", "saleable", "carpet"]:
                    unit_data[f"prev_{prefix}"] = 0
                    unit_data[f"{prefix}_delta"] = 0
                unit_data["prev_amount_received"] = 0
                unit_data["amount_received_delta"] = amount_received
                unit_data["prev_demand"] = 0
                unit_data["demand_delta"] = demand_raised

            # Aging computation
            variation_in_demand = demand_inc if demand_inc > 0 else 0
            aging_3060 = 0.0
            aging_60plus = 0.0
            if os_against_demand > 0:
                if os_against_demand > variation_in_demand:
                    aging_3060 = variation_in_demand
                    aging_60plus = os_against_demand - variation_in_demand
                else:
                    aging_3060 = os_against_demand
                    aging_60plus = 0.0

            unit_data.update({
                "Status": status,
                "Change Details": change_log,
                "DEMAND_INCREMENT_VAL": demand_inc,
                "RECEIVED_INCREMENT_VAL": received_inc,
                "AGREEMENT_INCREMENT_VAL": agreement_inc,
                "Rate per sft": round(rate_per_sft, 2),
                "Disbursement": disbursement,
                "Amount Received excl. Tax Current Month": amount_received,
                "Outstanding against demand": round(os_against_demand, 2),
                "O/S % Demand": round(os_pct_demand, 2),
                "Outstanding against sale value": round(os_against_sale, 2),
                "O/S against Sale Value": round(os_pct_sale, 2),
                "Upto 30 days": "Not applicable",
                "30 - 60 days": round(aging_3060, 2),
                "Greater than 60 days": round(aging_60plus, 2),
                "Total aging": round(aging_3060 + aging_60plus, 2),
                "REFERENCE_MSP": 0,
            })

            final_rows.append(unit_data)

        # Cancelled units missing from current
        for unit_no, prev_data in prev_lookup.items():
            if unit_no not in curr_lookup:
                prev_name = str(prev_data.get("Customer Name", "")).strip()
                if prev_name.lower() not in ["", "nan", "n/a", "-", "unsold"]:
                    p_demand = clean_num(prev_data.get("Demand Raised as on Current Month excl. tax", 0))
                    p_received = clean_num(prev_data.get("Amount Received excl. Tax", 0))
                    cancelled = prev_data.copy()
                    cancelled.update({
                        "Unit No.": unit_no,
                        "Status": "CANCELLATION",
                        "Change Details": f"[CANCELLATION] {prev_name}",
                        "Customer Name": "MISSING IN CURRENT",
                        "DEMAND_INCREMENT_VAL": 0.0,
                        "RECEIVED_INCREMENT_VAL": 0.0,
                        "AGREEMENT_INCREMENT_VAL": 0.0,
                        "Rate per sft": 0,
                        "Disbursement": "N/A",
                        "Amount Received excl. Tax Current Month": p_received,
                        "Outstanding against demand": p_demand - p_received,
                        "O/S % Demand": ((p_demand - p_received) / p_demand * 100) if p_demand > 0 else 0,
                        "Outstanding against sale value": 0,
                        "O/S against Sale Value": 0,
                        "Upto 30 days": "Not applicable",
                        "30 - 60 days": 0,
                        "Greater than 60 days": 0,
                        "Total aging": 0,
                        "REFERENCE_MSP": 0,
                        "prev_agreement": 0, "agreement_delta": 0,
                        "prev_amount_received": 0, "amount_received_delta": 0,
                        "prev_demand": 0, "demand_delta": 0,
                        "prev_saleable": 0, "saleable_delta": 0,
                        "prev_carpet": 0, "carpet_delta": 0,
                    })
                    final_rows.append(cancelled)

        # Transfer detection pass
        prev_customer_to_unit = {}
        for u, d in prev_lookup.items():
            n = str(d.get("Customer Name", "")).strip().upper()
            if n and n not in ["", "NAN", "N/A", "-", "UNSOLD"]:
                prev_customer_to_unit[n] = u

        cancelled_customers = {}
        for row in final_rows:
            if row.get("Status") == "CANCELLATION":
                cust = str(row.get("Customer Name", "")).strip().upper()
                unit = str(row.get("Unit No.", "")).strip()
                if cust and cust not in ["", "NAN", "N/A", "-", "UNSOLD", "MISSING IN CURRENT"]:
                    cancelled_customers[cust] = unit

        for row in final_rows:
            if row.get("Status") == "ANOMALY":
                curr_cust = str(row.get("Customer Name", "")).strip().upper()
                curr_unit = str(row.get("Unit No.", "")).strip()
                if curr_cust in prev_customer_to_unit:
                    old_unit = prev_customer_to_unit[curr_cust]
                    if old_unit != curr_unit:
                        row["Status"] = "TRANSFER"
                        row["Change Details"] = f"[TRANSFER] Unit {old_unit} → Unit {curr_unit} | Customer: {row.get('Customer Name')}"

            if row.get("Status") == "NEW":
                curr_cust = str(row.get("Customer Name", "")).strip().upper()
                curr_unit = str(row.get("Unit No.", "")).strip()
                from_unit = cancelled_customers.get(curr_cust)
                if from_unit and from_unit != curr_unit:
                    row["Status"] = "TRANSFER"
                    row["Change Details"] = f"[TRANSFER] Unit {from_unit} → Unit {curr_unit} | Customer: {row.get('Customer Name')}"

        # Clean NaN
        import math
        def clean_val(v):
            if isinstance(v, float) and math.isnan(v):
                return None
            return v

        result = [{k: clean_val(v) for k, v in row.items()} for row in final_rows]

        total_units = len([r for r in result if r.get("Status") != "CANCELLATION"])

        return {
            "status": "success",
            "extracted_data": result,
            "total_unit_count": total_units,
            "sold_units": sold_count,
            "unsold_units": total_units - sold_count,
        }

    except Exception as e:
        logger.error(f"Comparison error: {e}", exc_info=True)
        return {"status": "error", "message": str(e)}


@router.post("/compare")
async def compare_mis(
    prev_month: UploadFile = File(...),
    curr_month: UploadFile = File(...)
):
    try:
        prev_content = await prev_month.read()
        curr_content = await curr_month.read()
        result = process_comparison(prev_content, curr_content)
        if result["status"] == "error":
            raise HTTPException(status_code=400, detail=result["message"])
        return result
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))