from datetime import datetime
from firebase_admin import firestore

# Extend this as more services get added.
SERVICE_CODES = {
    "tdd": {"code": "TDD", "reset": "monthly"},
    "continuous-monitoring": {"code": "COM", "reset": "never"},
    "periodic-monitoring": {"code": "POM", "reset": "never"},
    "lender-independent-engineering": {"code": "LIE", "reset": "never"},  # confirm this key matches Valuation Tool's actual serviceKey
}


def _region_code(region: str) -> str:
    return (region or "").strip().upper()[:3]


@firestore.transactional
def _next_sequence(transaction, counter_ref):
    snapshot = counter_ref.get(transaction=transaction)
    current = snapshot.get("seq") if snapshot.exists else 0
    next_seq = (current or 0) + 1
    transaction.set(counter_ref, {"seq": next_seq}, merge=True)
    return next_seq


def generate_service_code(db, region: str, service_key: str, created_at=None):
    """
    Generates codes like WES-TDD-2026-08-001.
    TDD resets to 001 every month per region. COM/POM/LIE count forever, never resetting.
    Returns None if region or serviceKey is missing/unrecognized (project just stays without a code).
    """
    service_info = SERVICE_CODES.get(service_key)
    if not service_info or not region:
        return None

    reg_code = _region_code(region)
    svc_code = service_info["code"]
    created_at = created_at or datetime.utcnow()
    year = created_at.year
    month = f"{created_at.month:02d}"

    if service_info["reset"] == "monthly":
        counter_id = f"{reg_code}-{svc_code}-{year}-{month}"
    else:
        counter_id = f"{reg_code}-{svc_code}"

    counter_ref = db.collection("code_counters").document(counter_id)
    transaction = db.transaction()
    seq = _next_sequence(transaction, counter_ref)

    return f"{reg_code}-{svc_code}-{year}-{month}-{seq:03d}"
