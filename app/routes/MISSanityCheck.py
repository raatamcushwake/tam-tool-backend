from fastapi import APIRouter, UploadFile, File, HTTPException
import pandas as pd
import io
import logging
import re
import difflib

logger = logging.getLogger(__name__)
try:
    pd.set_option('future.no_silent_downcasting', True)
except Exception:
    pass

router = APIRouter()

def process_mis_sanity(prev_content, curr_content):
    try:
        standard_sequence = [
            "Unit No.", "Tower", "Booking Date", "Registration Date", 
            "Unit Type", "Customer Name", "Saleable area in sft", 
            "Carpet area in sft", "Agreement value", 
            "Amount Received excl. Tax", 
            "Demand Raised as on Current Month excl. tax"
        ]

        def clean_sheet(content):
            df = pd.read_excel(io.BytesIO(content), header=0)
            df = df.iloc[1:].reset_index(drop=True)
            actual_cols = df.columns.tolist()
            mapping = {}
            for std_name in standard_sequence:
                for col in actual_cols:
                    if str(col).strip().lower().startswith(std_name.lower()):
                        mapping[col] = std_name
                        break
            df = df[list(mapping.keys())].rename(columns=mapping)
            existing_std_cols = [c for c in standard_sequence if c in df.columns]
            df = df[existing_std_cols]
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
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0)
            return df.fillna("").infer_objects(copy=False)

        prev_df = clean_sheet(prev_content)
        curr_df = clean_sheet(curr_content)

        duplicate_units = curr_df[curr_df.duplicated(subset=['Unit No.'], keep=False)]
        has_duplicates = len(duplicate_units) > 0
        duplicate_units_list = duplicate_units.to_dict(orient='records') if has_duplicates else []

        prev_df_clean = prev_df.drop_duplicates(subset=['Unit No.'], keep='first')
        curr_df_clean = curr_df.drop_duplicates(subset=['Unit No.'], keep='first')
        prev_lookup = prev_df_clean.set_index("Unit No.").to_dict(orient='index')
        curr_lookup = curr_df_clean.set_index("Unit No.").to_dict(orient='index')

        all_units = []
        new_bookings = []
        already_cancelled_units = set()
        transferred_units = []
        name_corrections = []
        cancelled_units = []
        anomaly_units = []

        agreement_decreased, agreement_increased = [], []
        amount_decreased, amount_increased = [], []
        demand_decreased, demand_increased = [], []
        saleable_decreased, saleable_increased = [], []
        carpet_decreased, carpet_increased = [], []

        for _, row in curr_df.iterrows():
            unit_no = str(row["Unit No."])
            unit_data = row.to_dict()
            curr_name_raw = str(unit_data.get('Customer Name', '')).strip()
            curr_name_lower = curr_name_raw.lower()
            is_new = unit_no not in prev_lookup

            if not is_new:
                prev_unit = prev_lookup[unit_no]
                prev_name = str(prev_unit.get('Customer Name', '')).strip()
                unit_data['is_new'] = False
                prev_was_unsold = prev_name.lower() in ["", "nan", "n/a", "-", "unsold"]

                if curr_name_lower in ["unsold", "-", ""]:
                    unit_data['is_cancelled'] = True
                    unit_data['prev_customer'] = prev_name
                    cancelled_units.append(unit_data)
                elif prev_was_unsold:
                    unit_data['is_new'] = True
                    new_bookings.append(unit_data)
                elif prev_name.lower() != curr_name_lower:
                    similarity = difflib.SequenceMatcher(None, prev_name.lower(), curr_name_lower).ratio()
                    unit_data['prev_customer'] = prev_name
                    unit_data['curr_customer'] = curr_name_raw
                    if similarity >= 0.80:
                        unit_data['name_correction_detected'] = True
                        name_corrections.append(unit_data)
                    else:
                        prev_name_lower = prev_name.lower()
                        curr_name_lower_check = curr_name_raw.lower()
                        prev_customer_exists_elsewhere = any(
                            str(curr_row.get('Customer Name', '')).strip().lower() == prev_name_lower
                            for curr_unit_no, curr_row in curr_lookup.items()
                            if curr_unit_no != unit_no
                        )
                        curr_customer_existed_in_prev = any(
                            str(prev_row.get('Customer Name', '')).strip().lower() == curr_name_lower_check
                            for prev_unit_no, prev_row in prev_lookup.items()
                            if prev_unit_no != unit_no
                        )
                        if prev_customer_exists_elsewhere or curr_customer_existed_in_prev:
                            unit_data['anomaly_detected'] = True
                            unit_data['is_resale'] = True
                            anomaly_units.append(unit_data)
                        else:
                            # Same unit, different customer name, not a fuzzy match, and neither
                            # name shows up anywhere else this/last month.
                            # Same booking date on both months -> Resale (Anomaly).
                            # Different/missing booking date -> old customer cancelled this unit
                            # and a new customer booked it fresh: split into Cancellation + New Booking.
                            prev_bd = str(prev_unit.get('Booking Date', '')).strip()
                            curr_bd = str(unit_data.get('Booking Date', '')).strip()
                            same_booking_date = (
                                prev_bd not in ('', '-', 'nan', 'NaT') and
                                curr_bd not in ('', '-', 'nan', 'NaT') and
                                prev_bd == curr_bd
                            )
                            if same_booking_date:
                                unit_data['transfer_detected'] = True
                                unit_data['anomaly_detected'] = True
                                unit_data['is_resale'] = True
                                unit_data['prev_customer'] = prev_name
                                unit_data['curr_customer'] = curr_name_raw
                                transferred_units.append(unit_data)
                                anomaly_units.append(unit_data)
                            else:
                                cancelled_prev = prev_unit.copy()
                                cancelled_prev.update({
                                    'Unit No.': unit_no,
                                    'is_cancelled': True,
                                    'prev_customer': prev_name,
                                    'Customer Name': prev_name,
                                })
                                cancelled_units.append(cancelled_prev)
                                already_cancelled_units.add(unit_no)
                                unit_data['is_new'] = True
                                unit_data['is_split_new_booking'] = True
                                new_bookings.append(unit_data)

                if unit_data.get('is_split_new_booking'):
                    unit_data['prev_agreement'] = 0
                    unit_data['agreement_delta'] = 0
                    unit_data['prev_amount_received'] = 0
                    unit_data['amount_received_delta'] = unit_data.get('Amount Received excl. Tax', 0)
                    if unit_data['amount_received_delta'] > 0:
                        amount_increased.append(unit_data)
                    unit_data['prev_demand'] = 0
                    unit_data['demand_delta'] = unit_data.get('Demand Raised as on Current Month excl. tax', 0)
                    if unit_data['demand_delta'] > 0:
                        demand_increased.append(unit_data)
                    unit_data['prev_saleable'] = 0
                    unit_data['saleable_delta'] = 0
                    unit_data['prev_carpet'] = 0
                    unit_data['carpet_delta'] = 0
                else:
                    params = [
                        ('Agreement value', 'agreement', agreement_decreased, agreement_increased),
                        ('Amount Received excl. Tax', 'amount_received', amount_decreased, amount_increased),
                        ('Demand Raised as on Current Month excl. tax', 'demand', demand_decreased, demand_increased),
                        ('Saleable area in sft', 'saleable', saleable_decreased, saleable_increased),
                        ('Carpet area in sft', 'carpet', carpet_decreased, carpet_increased)
                    ]
                    for excel_key, prefix, dec_list, inc_list in params:
                        p_val = prev_unit.get(excel_key, 0)
                        c_val = unit_data.get(excel_key, 0)
                        delta = round(c_val - p_val, 2)
                        unit_data[f'prev_{prefix}'] = p_val
                        unit_data[f'{prefix}_delta'] = delta
                        threshold = 100 if prefix in ('amount_received', 'demand') else 0
                        if delta < -threshold:
                            dec_list.append(unit_data)
                        elif delta > 0:
                            inc_list.append(unit_data)
            else:
                is_sold_new = curr_name_lower not in ["", "nan", "n/a", "-", "unsold"]
                unit_data['is_new'] = True
                if is_sold_new:
                    new_bookings.append(unit_data)
                financial_params = [
                    ('Amount Received excl. Tax', 'amount_received', amount_increased),
                    ('Demand Raised as on Current Month excl. tax', 'demand', demand_increased),
                ]
                for excel_key, prefix, inc_list in financial_params:
                    curr_val = unit_data.get(excel_key, 0)
                    unit_data[f'prev_{prefix}'] = 0
                    unit_data[f'{prefix}_delta'] = curr_val
                    if curr_val > 0:
                        inc_list.append(unit_data)
                for prefix in ['agreement', 'saleable', 'carpet']:
                    unit_data[f'prev_{prefix}'] = 0
                    unit_data[f'{prefix}_delta'] = 0

            all_units.append(unit_data)

        for unit_no, prev_data in prev_lookup.items():
            if unit_no not in curr_lookup and unit_no not in already_cancelled_units:
                cancelled_unit = prev_data.copy()
                cancelled_unit.update({
                    'Unit No.': unit_no,
                    'is_cancelled': True,
                    'prev_customer': prev_data.get('Customer Name', ''),
                    'Customer Name': "MISSING IN CURRENT"
                })
                cancelled_units.append(cancelled_unit)

        # Cross-unit reassignment pass — mirrors MIS Analysis logic.
        # If a customer who got cancelled from one unit reappears as a
        # NEW booking (or was already flagged anomaly) on a different
        # unit, treat it as a single Anomaly/Resale event and drop the
        # now-redundant standalone Cancellation row.
        prev_customer_to_unit = {}
        for u, d in prev_lookup.items():
            n = str(d.get('Customer Name', '')).strip().upper()
            if n and n not in ["", "NAN", "N/A", "-", "UNSOLD"]:
                prev_customer_to_unit[n] = u

        cancelled_customers = {}
        for row in cancelled_units:
            cust = str(row.get('prev_customer', '') or row.get('Customer Name', '')).strip().upper()
            unit = str(row.get('Unit No.', '')).strip()
            if cust and cust not in ["", "NAN", "N/A", "-", "UNSOLD", "MISSING IN CURRENT"]:
                cancelled_customers[cust] = unit

        units_to_drop_from_cancellations = set()

        for row in anomaly_units:
            curr_cust = str(row.get('Customer Name', '')).strip().upper()
            curr_unit = str(row.get('Unit No.', '')).strip()
            old_unit = prev_customer_to_unit.get(curr_cust)
            if old_unit and old_unit != curr_unit:
                row['is_resale'] = True
                row['from_unit'] = old_unit
                row['to_unit'] = curr_unit
                units_to_drop_from_cancellations.add(old_unit)

        units_to_drop_from_new_bookings = set()

        for row in new_bookings:
            curr_cust = str(row.get('Customer Name', '')).strip().upper()
            curr_unit = str(row.get('Unit No.', '')).strip()
            from_unit = cancelled_customers.get(curr_cust)
            if from_unit and from_unit != curr_unit:
                from_unit_prev_name = str(prev_lookup.get(from_unit, {}).get('Customer Name', '')).strip()
                curr_unit_prev_name = str(prev_lookup.get(curr_unit, {}).get('Customer Name', '')).strip()
                curr_name_for_unit = str(row.get('Customer Name', '')).strip()
                row['anomaly_detected'] = True
                row['is_resale'] = True
                row['from_unit'] = from_unit
                row['to_unit'] = curr_unit
                row['prev_customer'] = curr_unit_prev_name or from_unit_prev_name or 'Unsold'
                row['curr_customer'] = curr_name_for_unit
                if row not in anomaly_units:
                    anomaly_units.append(row)
                units_to_drop_from_cancellations.add(from_unit)
                # This row is being reclassified as Anomaly — mirror Analysis's
                # `row["Status"] = "ANOMALY"` behavior by removing it from New Bookings
                # so it isn't double-counted in both buckets.
                units_to_drop_from_new_bookings.add(curr_unit)

        if units_to_drop_from_new_bookings:
            new_bookings = [
                row for row in new_bookings
                if str(row.get('Unit No.', '')).strip() not in units_to_drop_from_new_bookings
            ]

        if units_to_drop_from_cancellations:
            cancelled_units = [
                row for row in cancelled_units
                if str(row.get('Unit No.', '')).strip() not in units_to_drop_from_cancellations
            ]

        def get_sums(list_data, delta_key):
            return sum(abs(u.get(delta_key, 0)) for u in list_data)

        agreement_dec = get_sums(agreement_decreased, 'agreement_delta')
        amount_dec = get_sums(amount_decreased, 'amount_received_delta')
        demand_dec = get_sums(demand_decreased, 'demand_delta')
        saleable_dec = get_sums(saleable_decreased, 'saleable_delta')
        carpet_dec = get_sums(carpet_decreased, 'carpet_delta')
        agreement_inc = get_sums(agreement_increased, 'agreement_delta')
        saleable_inc = get_sums(saleable_increased, 'saleable_delta')
        carpet_inc = get_sums(carpet_increased, 'carpet_delta')

        sanity_passed = (
            amount_dec == 0 and demand_dec == 0 and
            agreement_dec == 0 and agreement_inc == 0 and
            saleable_dec == 0 and saleable_inc == 0 and
            carpet_dec == 0 and carpet_inc == 0 and
            not has_duplicates
        )

        issues = []
        if has_duplicates:
            issues.append(f"Duplicate Unit Numbers found ({len(duplicate_units)} entries)")
        if amount_dec > 0:
            issues.append("Amount Received decreased (not allowed)")
        if demand_dec > 0:
            issues.append("Demand Raised decreased (not allowed)")
        if agreement_dec > 0 or agreement_inc > 0:
            issues.append("Agreement Value changed (must remain constant)")
        if saleable_dec > 0 or saleable_inc > 0:
            issues.append("Saleable Area changed (must remain constant)")
        if carpet_dec > 0 or carpet_inc > 0:
            issues.append("Carpet Area changed (must remain constant)")

        return {
            "status": "success",
            "sanity_check_passed": sanity_passed,
            "issues": issues,
            "summary": {
                "total_units_scanned": len(all_units),
                "new_bookings_count": len(new_bookings),
                "transferred_count": len(transferred_units),
                "name_correction_count": len(name_corrections),
                "cancelled_count": len(cancelled_units),
                "duplicate_count": len(duplicate_units_list),
                "anomaly_count": len(anomaly_units),
                "agreement_dec": agreement_dec, "agreement_inc": agreement_inc,
                "agreement_dec_count": len(agreement_decreased), "agreement_inc_count": len(agreement_increased),
                "amount_dec": amount_dec, "amount_inc": get_sums(amount_increased, 'amount_received_delta'),
                "amount_dec_count": len(amount_decreased), "amount_inc_count": len(amount_increased),
                "demand_dec": demand_dec, "demand_inc": get_sums(demand_increased, 'demand_delta'),
                "demand_dec_count": len(demand_decreased), "demand_inc_count": len(demand_increased),
                "saleable_dec": saleable_dec, "saleable_inc": saleable_inc,
                "saleable_dec_count": len(saleable_decreased), "saleable_inc_count": len(saleable_increased),
                "carpet_dec": carpet_dec, "carpet_inc": carpet_inc,
                "carpet_dec_count": len(carpet_decreased), "carpet_inc_count": len(carpet_increased),
            },
            "units": all_units,
            "new_bookings": new_bookings,
            "transferred_units": transferred_units,
            "name_corrections": name_corrections,
            "cancelled_units": cancelled_units,
            "duplicate_units": duplicate_units_list,
            "anomaly_units": anomaly_units,
            "decreases": {
                "agreement": agreement_decreased, "amount": amount_decreased,
                "demand": demand_decreased, "saleable": saleable_decreased, "carpet": carpet_decreased
            },
            "increases": {
                "agreement": agreement_increased, "amount": amount_increased,
                "demand": demand_increased, "saleable": saleable_increased, "carpet": carpet_increased
            }
        }
    except Exception as e:
        logger.error(f"Error: {str(e)}", exc_info=True)
        return {"status": "error", "message": str(e)}


@router.post("/run")
async def run_mis_sanity(
    prev_month: UploadFile = File(...),
    curr_month: UploadFile = File(...)
):
    try:
        prev_content = await prev_month.read()
        curr_content = await curr_month.read()
        result = process_mis_sanity(prev_content, curr_content)
        if result["status"] == "error":
            raise HTTPException(status_code=400, detail=result["message"])
        return result
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
