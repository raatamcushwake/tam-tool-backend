from fastapi import APIRouter, HTTPException
from app.core.firebase import get_firestore

router = APIRouter()


def _find_units_array(doc_data: dict):
    """
    Scans a Firestore document's top-level fields and returns the first
    list of dict records that looks like unit-wise MIS data (i.e. each
    item has a 'Customer Name' key). Works regardless of what the field
    is actually named.
    """
    for key, value in doc_data.items():
        if isinstance(value, list) and len(value) > 0:
            if isinstance(value[0], dict) and "Customer Name" in value[0]:
                return value
    return []


@router.get("/mis-months/{project_id}")
async def get_mis_months(project_id: str):
    """Returns the list of available MIS month document IDs for a project."""
    try:
        db = get_firestore()
        docs = db.collection("projects").document(project_id).collection("misSubmissions").stream()
        months = [d.id for d in docs]
        return {"months": months}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.get("/mis-data/{project_id}/{month}")
async def get_mis_data(project_id: str, month: str):
    """
    Returns only Unit No., Customer Name, and RECEIVED_INCREMENT_VAL
    for every unit in the selected month's MIS document.
    """
    try:
        db = get_firestore()
        doc = db.collection("projects").document(project_id)\
                 .collection("misSubmissions").document(month).get()

        if not doc.exists:
            raise HTTPException(status_code=404, detail="MIS month not found")

        doc_data = doc.to_dict()
        units = _find_units_array(doc_data)

        result = []
        for u in units:
            result.append({
                "unitNo": u.get("Unit No.", ""),
                "customerName": u.get("Customer Name", ""),
                "receivedIncrementVal": u.get("RECEIVED_INCREMENT_VAL", 0),
            })

        return {"month": month, "units": result}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


from typing import List, Optional
from pydantic import BaseModel
from app.services.entity_matcher import get_entity_for_unit, match_transaction_to_entity_detailed, get_all_entities


class BankTransaction(BaseModel):
    date: str
    valueDate: str = ""
    narration: str
    referenceNo: str = ""
    depositAmount: float
    accountNumber: Optional[str] = ""
    monthKey: str = "UNKNOWN"


def _group_matched_rows(matched_rows: list) -> list:
    """
    Groups matched transaction rows by (unitNo, monthKey), so multiple
    transactions in the same month for the same unit appear as ONE row —
    narrations stacked together, deposit amounts summed, and the MIS
    collection value shown once per group.
    """
    groups = {}
    for row in matched_rows:
        key = (row["unitNo"], row["monthKey"])
        if key not in groups:
            groups[key] = {
                "unitNo": row["unitNo"],
                "unitNoBankStmt": row["unitNoBankStmt"],
                "customerName": row["customerName"],
                "collectionRaised": row["collectionRaised"],
                "monthKey": row["monthKey"],
                "transactions": [],
                "totalDepositAmount": 0.0,
            }
        groups[key]["transactions"].append({
            "date": row["date"],
            "valueDate": row["valueDate"],
            "narration": row["narration"],
            "depositAmount": row["depositAmount"],
        })
        groups[key]["totalDepositAmount"] += row["depositAmount"]

    return list(groups.values())


class ManualUnit(BaseModel):
    unitNo: str
    customerName: str = ""
    receivedIncrementVal: float = 0


class CompareRequest(BaseModel):
    projectId: str
    month: str
    transactions: List[BankTransaction]
    manualUnits: Optional[List[ManualUnit]] = None


def _get_units_for_month(project_id: str, month: str) -> list:
    db = get_firestore()
    doc = db.collection("projects").document(project_id)\
             .collection("misSubmissions").document(month).get()

    if not doc.exists:
        raise HTTPException(status_code=404, detail="MIS month not found")

    doc_data = doc.to_dict()
    units = _find_units_array(doc_data)

    result = []
    for u in units:
        result.append({
            "unitNo": u.get("Unit No.", ""),
            "customerName": u.get("Customer Name", ""),
            "receivedIncrementVal": u.get("RECEIVED_INCREMENT_VAL", 0),
        })
    return result


@router.post("/compare")
async def compare_collection(request: CompareRequest):
    """
    Matches each MIS unit's entity fingerprint (from entity_reference.xlsx)
    against the uploaded bank statement's transactions. A unit can match
    multiple transactions (e.g. partial payments).
    """
    try:
        if request.manualUnits:
            mis_units = [
                {"unitNo": u.unitNo, "customerName": u.customerName, "receivedIncrementVal": u.receivedIncrementVal}
                for u in request.manualUnits
            ]
        else:
            mis_units = _get_units_for_month(request.projectId, request.month)

        matched_rows = []
        unmatched_units = []
        matched_txn_indices = set()

        match_method_counts = {"unit_number": 0, "account": 0, "account_last4": 0, "upi": 0, "mobile": 0, "name": 0, "ifsc": 0}
        unique_matched_units = set()

        # Transactions that matched a real unit, but that unit has no
        # recorded increment this month — these should be excluded
        # entirely (not Matched, not Unmatched) since there's nothing to
        # reconcile them against.
        zero_increment_txn_indices = set()

        # Signal strength: lower number = stronger/more trustworthy match.
        # Used to pick the single best unit for each transaction when more
        # than one entity's fingerprint happens to match the same narration.
        # Unit number is checked FIRST/highest priority — it's the entity's
        # own unique identifier and should win over any coincidental
        # name/account/UPI overlap from an unrelated unit.
        METHOD_RANK = {"unit_number": 0, "account": 1, "account_last4": 2, "upi": 3, "mobile": 4, "name": 5, "ifsc": 6}

        # Build (unit, entity) pairs once
        unit_entities = []
        for unit in mis_units:
            entity = get_entity_for_unit(unit["unitNo"])
            if not entity:
                unmatched_units.append({**unit, "reason": "Unit not found in entity reference"})
                continue
            unit_entities.append((unit, entity))

        units_with_a_match = set()

        # For EACH transaction, find the single best-matching unit across
        # ALL entities — instead of letting every entity independently claim
        # the same transaction. Only the strongest match wins.
        debug_collisions = []  # transactions matched by more than one unit

        for idx, txn in enumerate(request.transactions):
            best = None  # (rank, -specificity, unit, method)
            all_candidates = []  # every unit that matched this transaction, for debugging

            for unit, entity in unit_entities:
                method = match_transaction_to_entity_detailed(txn.narration, txn.accountNumber or "", entity)
                if method is None:
                    continue
                all_candidates.append({"unitNo": unit["unitNo"], "customerName": unit["customerName"], "method": method})
                rank = METHOD_RANK[method]

                # When two candidates tie on the same method (e.g. both
                # matched via "name"), prefer whichever one's identifying
                # text is LONGER/more specific — a 3-word exact name match
                # ("Ashok Kumar Koul") is far more trustworthy than a
                # shorter 2-word match ("Ashok Kumar") that happens to be
                # a subset of the same narration.
                specificity = len(unit["customerName"] or "")
                candidate = (rank, -specificity, unit, method)

                if best is None or candidate[:2] < best[:2]:
                    best = candidate

            if len(all_candidates) > 1:
                debug_collisions.append({
                    "narration": txn.narration,
                    "depositAmount": txn.depositAmount,
                    "candidates": all_candidates,
                    "winner": best[2]["unitNo"] if best else None,
                })

            if best is None:
                continue

            _, _, unit, method = best

            # Unit found, but no increment recorded this month — exclude
            # this transaction entirely (not Matched, not Unmatched).
            if not unit["receivedIncrementVal"]:
                zero_increment_txn_indices.add(idx)
                units_with_a_match.add(unit["unitNo"])
                continue

            matched_txn_indices.add(idx)
            match_method_counts[method] += 1
            units_with_a_match.add(unit["unitNo"])
            unique_matched_units.add(unit["unitNo"])

            matched_rows.append({
                "unitNo": unit["unitNo"],
                "unitNoBankStmt": unit["unitNo"],
                "customerName": unit["customerName"],
                "date": txn.date,
                "valueDate": txn.valueDate,
                "narration": txn.narration,
                "collectionRaised": unit["receivedIncrementVal"],
                "depositAmount": txn.depositAmount,
                "matchMethod": method,
                "monthKey": txn.monthKey,
            })

        for unit, entity in unit_entities:
            if unit["unitNo"] not in units_with_a_match:
                unmatched_units.append({**unit, "reason": "No matching transaction found in bank statement"})

        grouped_matched = _group_matched_rows(matched_rows)

        # Transactions that never matched ANY unit — these need manual unit mapping
        unmatched_transactions = [
            {
                "date": txn.date,
                "valueDate": txn.valueDate,
                "narration": txn.narration,
                "depositAmount": txn.depositAmount,
                "monthKey": txn.monthKey,
            }
            for idx, txn in enumerate(request.transactions)
            if idx not in matched_txn_indices and idx not in zero_increment_txn_indices
        ]

        all_entities = get_all_entities()
        for txn_record in unmatched_transactions:
            best = None
            for entity in all_entities:
                method = match_transaction_to_entity_detailed(txn_record["narration"], "", entity)
                if method is None:
                    continue
                rank = METHOD_RANK[method]
                if best is None or rank < best[0]:
                    best = (rank, entity["unitNo"], entity.get("customerName", ""), method)
            if best:
                txn_record["suggestedUnitNo"] = best[1]
                txn_record["suggestedCustomerName"] = best[2]
                txn_record["suggestedMatchMethod"] = best[3]

        return {
            "month": request.month,
            "matched": grouped_matched,
            "unmatched": unmatched_units,
            "unmatchedTransactions": unmatched_transactions,
            "uniqueMatchedUnitsCount": len(unique_matched_units),
            "totalUnitsCount": len(mis_units),
            "matchMethodCounts": match_method_counts,
            "debugCollisions": debug_collisions,
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    