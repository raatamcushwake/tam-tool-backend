from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse
import pandas as pd
import io
import re
import logging
from app.core.firebase import get_storage_bucket, get_firestore

logger = logging.getLogger(__name__)
router = APIRouter()


# ─── Helper ───────────────────────────────────────────────────

def clean_num(val):
    if pd.isna(val) or val == "" or str(val).strip().lower() in ["nan", "n/a", "-", "empty"]:
        return 0.0
    try:
        cleaned = "".join(c for c in str(val) if c.isdigit() or c in [".", "-"])
        return float(cleaned) if cleaned else 0.0
    except:
        return 0.0


# ─── Extract Business Plan from uploaded file ─────────────────

def extract_business_plan_data(bp_content):
    try:
        xl = pd.ExcelFile(io.BytesIO(bp_content), engine='openpyxl')
        sheet_name = next((s for s in xl.sheet_names if "Business Plan" in s), xl.sheet_names[0])
        df_bp = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        
        print("DEBUG BP rows:", df_bp.shape)

        # Find Projected Expenses row
        outflow_row_idx = -1
        for i in range(len(df_bp)):
            cell_val = str(df_bp.iloc[i, 0]).strip().lower()
            if "projected expenses" in cell_val:
                outflow_row_idx = i
                break
            
        print("OUTFLOW TOTAL row index:", outflow_row_idx)

        # Find quarter row (row with Q1, Q2, Q3, Q4)
        quarter_row_idx = -1
        for i in range(min(10, len(df_bp))):
            row_vals = [str(v).strip().upper() for v in df_bp.iloc[i]]
            if any(v in ["Q1","Q2","Q3","Q4"] for v in row_vals):
                quarter_row_idx = i
                break
            
        # Find period row (row with month ranges like "Jul-Sep 2024")
        period_row_idx = -1
        for i in range(min(10, len(df_bp))):
            row_str = " ".join(str(v) for v in df_bp.iloc[i]).lower()
            if any(m in row_str for m in ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"]):
                if any(str(y) in row_str for y in range(2020, 2035)):
                    period_row_idx = i
                    break
                
        print("Quarter row:", quarter_row_idx, "Period row:", period_row_idx)

        budget_by_period = {}

        if outflow_row_idx >= 0 and quarter_row_idx >= 0:
            quarter_row = df_bp.iloc[quarter_row_idx]
            period_row = df_bp.iloc[period_row_idx] if period_row_idx >= 0 else [None]*len(quarter_row)
            outflow_row = df_bp.iloc[outflow_row_idx]

            for col_idx in range(len(quarter_row)):
                q_val = str(quarter_row.iloc[col_idx]).strip().upper()
                if q_val not in ["Q1","Q2","Q3","Q4"]:
                    continue
                period_label = str(period_row.iloc[col_idx]).strip() if period_row_idx >= 0 else ""
                outflow_val = clean_num(outflow_row.iloc[col_idx])
                # Convert to Crores (divide by 10000000)
                outflow_cr = outflow_val / 10000000
                if period_label and period_label != "nan":
                    budget_by_period[period_label] = outflow_cr
                    print(f"  Period: {period_label} | Quarter: {q_val} | Outflow: {outflow_cr} Cr")

        return {"available_periods": [], "period_data": {}, "budget_by_period": budget_by_period}
    except Exception as e:
        logger.error(f"BP Extraction Error: {e}", exc_info=True)
        return {"available_periods": [], "period_data": {}}


# ─── Process Cost Analysis ─────────────────────────────────────

def process_cost_analysis(business_plan_content, cleared_bills_content, reference_bp_content=None):
    try:
        bp_stats = extract_business_plan_data(reference_bp_content if reference_bp_content else business_plan_content)

        bills_df = pd.read_excel(io.BytesIO(cleared_bills_content))

        def simplify(text):
            return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()

        required_cols_order = [
            "Tranche", "Payment Clearance date", "Month",
            "Costing Particular", "Supplier/ Vendor/Customer/Salaries/Others",
            "Name", "Payment cleared"
        ]
        excel_col_map = {simplify(col): col for col in bills_df.columns}
        final_src_mapping = {}
        for col in required_cols_order:
            simplified = simplify(col)
            if simplified in excel_col_map:
                final_src_mapping[col] = excel_col_map[simplified]

        extracted_df = bills_df[list(final_src_mapping.values())].copy()
        extracted_df = extracted_df.fillna("").replace([float('inf'), float('-inf')], "")
        if "Payment Clearance date" in extracted_df.columns:
            extracted_df["Payment Clearance date"] = extracted_df["Payment Clearance date"].astype(str)

        if "Month" in final_src_mapping:
            month_col = final_src_mapping["Month"]
            def format_month_label(m):
                m = str(m).strip()
                parts = m.split("-")
                if len(parts) == 2:
                    mon = parts[0].capitalize()
                    yr = parts[1]
                    if len(yr) == 2:
                        yr = "20" + yr
                    return f"{mon}-{yr}"
                return m
            extracted_df[month_col] = extracted_df[month_col].astype(str).str.strip().apply(format_month_label)

        month_col_name = final_src_mapping.get("Month")
        payment_col_name = final_src_mapping.get("Payment cleared")
        particular_col_name = final_src_mapping.get("Costing Particular")

        all_bill_months = []
        if month_col_name:
            raw_months = extracted_df[month_col_name].dropna().unique().tolist()
            raw_months = [m for m in raw_months if m and m != 'nan' and m != '']
            month_order = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"]
            def parse_month_sort(m):
                m = str(m).strip().lower()
                parts = m.split("-")
                if len(parts) == 2:
                    mon = parts[0][:3].lower()
                    yr = parts[1]
                    if len(yr) == 2: yr = "20" + yr
                    try:
                        return (int(yr), month_order.index(mon))
                    except:
                        return (9999, 99)
                return (9999, 99)
            all_bill_months = sorted(raw_months, key=parse_month_sort)

        # ── Build bp_stats periods from bill months (NOT from BP Excel columns) ──
        for month in all_bill_months:
            bp_stats["period_data"][month] = {
                "planned_budget": 0,
                "period": month,
                "is_quarterly": False
            }
            bp_stats["available_periods"].append({
                "label": month,
                "display_label": month,
                "col_index": -1,
                "period": month
            })

        # ── Read BP Excel for static columns only (Particular, BP CTC, Revised CTC, Pre) ──
        bp_df_raw = pd.read_excel(io.BytesIO(business_plan_content), header=None, engine='openpyxl')

        header_row_idx = 0
        bp_ctc_col = None
        revised_ctc_cols = []  # list of (col_index, header_label)
        pre_col = None

        for i in range(min(5, len(bp_df_raw))):
            row_vals = [str(v).strip() for v in bp_df_raw.iloc[i]]
            row_vals_lower = [v.lower() for v in row_vals]
            for j, val in enumerate(row_vals_lower):
                if "ctc" in val and ("bp" in val or "business" in val or "plan" in val):
                    bp_ctc_col = j
                if "revised" in val and "ctc" in val:
                    revised_ctc_cols.append((j, row_vals[j]))  # store index + original label
                if "pre" in val or ("total" in val and "expense" in val):
                    pre_col = j
            if bp_ctc_col is not None:
                header_row_idx = i
                break

        # Remove duplicates keeping order
        seen = set()
        revised_ctc_cols = [(j, lbl) for j, lbl in revised_ctc_cols if not (j in seen or seen.add(j))]

        # Most recent revised CTC = last one found
        latest_revised_ctc_col = revised_ctc_cols[-1][0] if revised_ctc_cols else None

        print(f"Found columns — BP CTC: {bp_ctc_col}, Revised CTCs: {revised_ctc_cols}, Pre: {pre_col}")

        raw_rows = []
        for row_idx in range(header_row_idx + 1, len(bp_df_raw)):
            particular_name = str(bp_df_raw.iloc[row_idx, 0]).strip()
            if not particular_name or particular_name.lower() in ["nan", "none", ""]:
                particular_name = str(bp_df_raw.iloc[row_idx, 1]).strip() if bp_df_raw.shape[1] > 1 else ""

            if not particular_name or simplify(particular_name) in ["total", "nan", "", "none"]:
                continue

            bp_ctc_raw = clean_num(bp_df_raw.iloc[row_idx, bp_ctc_col]) if bp_ctc_col is not None else 0.0
            pre_val_raw = clean_num(bp_df_raw.iloc[row_idx, pre_col]) if pre_col is not None else 0.0

            # All revised CTCs as list of {label, value}
            all_revised_ctcs = []
            for (col_idx, col_label) in revised_ctc_cols:
                val = clean_num(bp_df_raw.iloc[row_idx, col_idx])
                all_revised_ctcs.append({"label": col_label, "value": val})

            # Latest revised CTC used for calculations
            revised_ctc_raw = clean_num(bp_df_raw.iloc[row_idx, latest_revised_ctc_col]) if latest_revised_ctc_col is not None else 0.0

            print(f"ROW {row_idx}: particular='{particular_name}' bp={bp_ctc_raw} revised_ctcs={all_revised_ctcs} pre={pre_val_raw}")

            per_month_bills = {}
            if particular_col_name and month_col_name and payment_col_name:
                p_clean = simplify(particular_name)
                for month in all_bill_months:
                    mask = (
                        extracted_df[particular_col_name].apply(simplify) == p_clean
                    ) & (
                        extracted_df[month_col_name].str.strip().str.lower() == month.lower()
                    )
                    subset = extracted_df[mask]
                    amount = float(subset[payment_col_name].sum()) if len(subset) > 0 else 0.0
                    transactions = subset.to_dict('records') if len(subset) > 0 else []
                    per_month_bills[month] = {"amount": amount, "transactions": transactions}

            raw_rows.append({
                "particular": particular_name,
                "bp_ctc": bp_ctc_raw,
                "all_revised_ctcs": all_revised_ctcs,       # all versions
                "revised_ctc": revised_ctc_raw,              # latest one for calculations
                "pre_val": pre_val_raw,
                "per_month_bills": per_month_bills
            })

        return {
            "status": "success",
            "raw_rows": raw_rows,
            "all_bill_months": all_bill_months,
            "revised_ctc_headers": [lbl for (_, lbl) in revised_ctc_cols],
            "extracted_bills": {
                "columns": required_cols_order,
                "data": extracted_df.to_dict('records')
            },
            "bp_stats": bp_stats
        }

    except Exception as e:
        logger.error(f"Cost Analysis Error: {e}", exc_info=True)
        return {"status": "error", "message": str(e)}


# ─── ROUTES ───────────────────────────────────────────────────

@router.post("/upload-bp")
async def upload_cost_bp(
    project_id: str = Form(...),
    file: UploadFile = File(...)
):
    """Manager uploads Cost Budget (BP) Excel — stored in Firebase Storage"""
    try:
        content = await file.read()
        bucket = get_storage_bucket()
        blob = bucket.blob(f"projects/{project_id}/costReference/businessPlan_latest.xlsx")
        blob.upload_from_string(content,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        blob.make_public()
        file_url = blob.public_url

        db = get_firestore()
        db.collection("projects").document(project_id)\
          .collection("costReferenceData").document("businessPlan")\
          .set({
              "fileUrl": file_url,
              "fileName": file.filename,
              "uploadedAt": pd.Timestamp.now().isoformat(),
          })

        return {"status": "success", "message": "BP uploaded successfully", "fileUrl": file_url}
    except Exception as e:
        logger.error(f"BP upload error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/run")
async def run_cost_analysis(
    project_id: str = Form(...),
    cleared_bills: UploadFile = File(...)
):
    """Maker runs cost analysis — fetches BP from Firebase Storage, processes with cleared bills"""
    try:
        # Get BP file URL from Firestore
        db = get_firestore()
        bp_doc = db.collection("projects").document(project_id)\
                   .collection("costReferenceData").document("businessPlan").get()

        if not bp_doc.exists:
            raise HTTPException(status_code=400,
                detail="No Cost Budget uploaded yet. Ask Manager to upload BP first.")

        bp_data = bp_doc.to_dict()
        file_url = bp_data.get("fileUrl")
        if not file_url:
            raise HTTPException(status_code=400, detail="BP file URL missing.")

        # Download BP from Firebase Storage
        import urllib.request
        with urllib.request.urlopen(file_url) as response:
            bp_content = response.read()

        # Read cleared bills
        # Read cleared bills
        cleared_bills_content = await cleared_bills.read()

        # Fetch Reference Business Plan for OUTFLOW budget figures
        reference_bp_content = None
        try:
            ref_doc = db.collection("projects").document(project_id)\
                        .collection("referenceData").document("businessPlan").get()
            if ref_doc.exists:
                ref_url = ref_doc.to_dict().get("fileUrl")
                if ref_url:
                    with urllib.request.urlopen(ref_url) as resp:
                        reference_bp_content = resp.read()
                    print("Reference BP fetched for budget extraction")
        except Exception as e:
            print(f"Reference BP fetch failed: {e}")

        # Run analysis
        # Fetch last approved bills from Firestore for incremental tracking
        last_approved_months = []
        try:
            approved_doc = db.collection("projects").document(project_id)\
                             .collection("costApprovedData").document("latestApproved").get()
            if approved_doc.exists:
                approved_data = approved_doc.to_dict()
                last_approved_months = approved_data.get("allBillMonths", [])
                print(f"Last approved months: {last_approved_months}")
        except Exception as e:
            print(f"Could not fetch last approved data: {e}")

        # Run analysis
        result = process_cost_analysis(bp_content, cleared_bills_content, reference_bp_content)

        # Tag which months are new vs previously approved
        if result.get("status") == "success" and last_approved_months:
            all_months = result.get("all_bill_months", [])
            result["new_bill_months"] = [m for m in all_months if m not in last_approved_months]
            result["previously_approved_months"] = last_approved_months
        else:
            result["new_bill_months"] = result.get("all_bill_months", [])
            result["previously_approved_months"] = []

        if result["status"] == "error":
            raise HTTPException(status_code=400, detail=result["message"])

        return result

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Cost analysis run error: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))
