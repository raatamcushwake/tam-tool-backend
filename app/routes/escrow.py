from fastapi import APIRouter, UploadFile, File, HTTPException, Form
from fastapi.responses import JSONResponse
import pdfplumber
import io
import re
import json
from typing import Optional, List

router = APIRouter()

def parse_amount(text):
    if not text:
        return 0.0
    cleaned = re.sub(r'[^\d.]', '', str(text).replace(',', ''))
    try:
        return float(cleaned)
    except:
        return 0.0

def matches_account(description, acc_no, acc_desc):
    """
    Returns True if the transaction description matches the given account —
    either by account number or by the account's description text
    (e.g. '70% Account').
    """
    desc_lower = (description or "").lower()
    acc_no = (acc_no or "").strip()
    acc_desc = (acc_desc or "").strip().lower()
    if acc_no and acc_no in description:
        return True
    if acc_desc and acc_desc in desc_lower:
        return True
    return False

def extract_statement_account_number(full_text):
    """
    Looks for a line like 'Account Number: 57500001513888' near the top of
    the PDF and returns the digits found, or None if not found.
    """
    match = re.search(r'Account Number\s*:?\s*([0-9]+)', full_text)
    if match:
        return match.group(1).strip()
    return None

def categorize(description, allowed_accounts, all_enabled):
    """
    allowed_accounts: list of {accNo, description} ticked as allowed for
    this account/direction.
    all_enabled: if True, every transaction is categorized as matched
    (green for inflow / red-as-matched for outflow), overriding ticks.
    Returns "matched" or "unmatched".
    """
    if all_enabled:
        return "matched"
    for acc in allowed_accounts:
        if matches_account(description, acc.get("accNo"), acc.get("description")):
            return "matched"
    return "unmatched"

from fastapi import APIRouter, UploadFile, File, HTTPException, Form
from fastapi.responses import JSONResponse
import pdfplumber
import io
import re
import json
from typing import Optional, List

router = APIRouter()

def parse_amount(text):
    if not text:
        return 0.0
    cleaned = re.sub(r'[^\d.]', '', str(text).replace(',', ''))
    try:
        return float(cleaned)
    except:
        return 0.0

def matches_account(description, acc_no, acc_desc):
    desc_lower = (description or "").lower()
    acc_no = (acc_no or "").strip()
    acc_desc = (acc_desc or "").strip().lower()
    if acc_no and acc_no in description:
        return True
    if acc_desc and acc_desc in desc_lower:
        return True
    return False

def extract_statement_account_number(full_text):
    match = re.search(r'Account Number\s*:?\s*([0-9]+)', full_text)
    if match:
        return match.group(1).strip()
    return None

def categorize(description, allowed_accounts, all_enabled):
    if all_enabled:
        return "matched"
    for acc in allowed_accounts:
        if matches_account(description, acc.get("accNo"), acc.get("description")):
            return "matched"
    return "unmatched"

@router.post("/parse-pdf")
async def parse_escrow_pdf(
    files: List[UploadFile] = File(...),
    accounts_config: str = Form(...),
):
    try:
        accounts = json.loads(accounts_config)
        accounts_by_no = {a["accNo"]: a for a in accounts}

        results = {}

        for file in files:
            contents = await file.read()
            inflow = []
            outflow = []
            full_text = ""
            all_rows = []  # to track first and last transaction

            with pdfplumber.open(io.BytesIO(contents)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text() or ""
                    full_text += page_text + "\n"

                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            if not row or len(row) < 5:
                                continue

                            row_text = ' '.join([str(c) for c in row if c])
                            if any(h in row_text for h in ['Transaction', 'Opening', 'Closing', 'Debit Amount', 'Credit Amount', 'Balance', 'Count', 'End Of', 'Summary']):
                                continue

                            description = str(row[1] or '').strip().replace('\n', ' ')
                            transaction_date = str(row[0] or '').strip().replace('\n', ' ')
                            reference_no = str(row[2] or '').strip().replace('\n', ' ')

                            debit = parse_amount(row[4]) if len(row) > 4 else 0
                            credit = parse_amount(row[5]) if len(row) > 5 else 0
                            closing_bal = parse_amount(row[6]) if len(row) > 6 else 0

                            if not description or description == 'None':
                                continue
                            if description.strip().isdigit():
                                continue

                            if transaction_date == 'None':
                                transaction_date = ''
                            if reference_no == 'None':
                                reference_no = ''

                            all_rows.append({
                                "debit": debit,
                                "credit": credit,
                                "closing_bal": closing_bal,
                                "transaction_date": transaction_date
                            })

                            if credit > 0:
                                inflow.append({
                                    "description": description,
                                    "amount": credit,
                                    "remark": "",
                                    "transactionDate": transaction_date,
                                    "referenceNo": reference_no
                                })
                            elif debit > 0:
                                outflow.append({
                                    "description": description,
                                    "amount": debit,
                                    "remark": "",
                                    "transactionDate": transaction_date,
                                    "referenceNo": reference_no
                                })

            # Calculate opening and closing balance
            opening_balance = 0.0
            closing_balance = 0.0

            if all_rows:
                first_row = all_rows[0]
                last_row = all_rows[-1]

                # Opening balance formula
                if first_row["credit"] > 0:
                    # First transaction is credit
                    opening_balance = first_row["closing_bal"] - first_row["credit"]
                elif first_row["debit"] > 0:
                    # First transaction is debit
                    opening_balance = first_row["closing_bal"] + first_row["debit"]

                # Closing balance from last transaction
                closing_balance = last_row["closing_bal"]

            # Identify which account this statement belongs to
            stmt_acc_no = extract_statement_account_number(full_text)
            matched_config = accounts_by_no.get(stmt_acc_no) if stmt_acc_no else None

            if not matched_config:
                continue

            acc_no = matched_config["accNo"]
            allowed_inflow = matched_config.get("allowedInflow", [])
            allowed_outflow = matched_config.get("allowedOutflow", [])
            all_inflow_enabled = matched_config.get("allInflowEnabled", False)
            all_outflow_enabled = matched_config.get("allOutflowEnabled", False)

            for row in inflow:
                cat = categorize(row["description"], allowed_inflow, all_inflow_enabled)
                row["category"] = "green" if cat == "matched" else "red"
            for row in outflow:
                cat = categorize(row["description"], allowed_outflow, all_outflow_enabled)
                row["category"] = "green" if cat == "matched" else "red"

            # Reversal detection
            for in_row in inflow:
                in_ref = (in_row.get("referenceNo") or "").strip()
                if not in_ref:
                    continue
                for out_row in outflow:
                    out_ref = (out_row.get("referenceNo") or "").strip()
                    if out_ref and out_ref == in_ref and out_row["amount"] == in_row["amount"]:
                        in_row["category"] = "reversal"
                        out_row["category"] = "reversal"

            results[acc_no] = {
                "inflow": inflow,
                "outflow": outflow,
                "matchedFile": file.filename,
                "openingBalance": opening_balance,
                "closingBalance": closing_balance,
                "balanceAsOf": all_rows[-1]["transaction_date"] if all_rows else "",
            }

        return JSONResponse({
            "success": True,
            "results": results
        })

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))