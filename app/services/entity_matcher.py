import re


def normalize_unit(unit: str) -> str:
    """
    Normalizes a unit number so slash and dash variants match.
    e.g. 'T1/51' and 'T1-51' both become 'T1_51'
    Also handles multi-part units like 'C1/GF/9' -> 'C1_GF_9'
    """
    if not unit:
        return ""
    return re.sub(r'[/\-]', '_', unit.strip().upper())

import os
import pandas as pd
from functools import lru_cache

ENTITY_REFERENCE_PATH = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "matched_output.xlsx"
)


def _split_multi(value) -> list:
    """
    Splits a comma-separated cell (e.g. '434015317168, 434016458465') into
    a clean list of individual values. Returns [] for blank/NaN cells.
    """
    if pd.isna(value) or str(value).strip() == "":
        return []
    return [v.strip() for v in str(value).split(",") if v.strip()]


# If a "name" appears under more than this many DIFFERENT units, it's
# almost certainly boilerplate text (e.g. the collection account's own
# name) rather than a real customer identifier. Kept high on purpose —
# real customers can legitimately co-own several units (joint owners,
# family members listed as "Additional Names"), and a low threshold here
# was wrongly stripping their real names too. This should only catch
# something appearing under DOZENS of units, not a handful.
MAX_UNITS_PER_NAME = 15

# Known boilerplate text that shows up in almost every bank narration
# (the collection account's own name) — explicitly excluded regardless of
# the unit-count threshold above, as a more precise safeguard.
KNOWN_BOILERPLATE_NAMES = {
    "KV DEVELOPERS",
    "K V DEVELOPERS",
    "K V DEVELOPERS PRIVATE LIMITED",
    "KV DEVELOPERS PRIVATE LIMITED",
    "KV DEVELOPERS PVT LTD",
    "K V DEVELOPERS PVT LTD",
}


@lru_cache(maxsize=1)
def load_entity_reference() -> dict:
    """
    Loads matched_output.xlsx ONCE (cached) and builds a lookup dict
    keyed by normalized unit number, so repeated calls don't re-read the
    file from disk every time.
    """
    df = pd.read_excel(ENTITY_REFERENCE_PATH)
    df = df.dropna(subset=["Unit No."])

    raw_by_unit = {}

    for _, row in df.iterrows():
        unit_key = normalize_unit(str(row["Unit No."]))

        names = set()
        for col in [
            "Customer Name (MIS)", "Name Inferred From Email",
            "Person Name (Bank Stmt)", "Additional Names",
            "Customer Name (Unit Mapping)",
        ]:
            val = row.get(col)
            if pd.notna(val) and str(val).strip():
                for name_part in _split_multi(val):
                    names.add(name_part)

        mobiles = set()
        for col in ["Mobile (MIS)", "Mobile (Bank Stmt)", "Mobile (Unit Mapping)"]:
            mobiles.update(_split_multi(row.get(col)))

        upi_ids = set()
        for col in ["UPI ID", "UPI ID (Unit Mapping)"]:
            upi_ids.update(_split_multi(row.get(col)))

        account_numbers = set()
        for col in ["Account Number", "Account Number (Unit Mapping)"]:
            account_numbers.update(_split_multi(row.get(col)))

        account_last4s = set()
        for col in ["Account Last4 (masked)", "Account Last4 (Unit Mapping)"]:
            account_last4s.update(_split_multi(row.get(col)))

        ifsc_codes = set()
        for col in ["IFSC (Unit Mapping)", "IFSC (Matched)"]:
            ifsc_codes.update(_split_multi(row.get(col)))

        bank_names = set()
        for col in ["Bank Name (Unit Mapping)", "Bank Name (Matched)"]:
            bank_names.update(_split_multi(row.get(col)))

        bank_stmt_unit_no = row.get("Unit No. (Bank Stmt)")
        bank_stmt_unit_no = str(bank_stmt_unit_no).strip() if pd.notna(bank_stmt_unit_no) else ""

        customer_name = ""
        for col in ["Customer Name (MIS)", "Customer Name (Unit Mapping)"]:
            val = row.get(col)
            if pd.notna(val) and str(val).strip():
                customer_name = str(val).strip()
                break

        entry = {
            "unitKey": unit_key,
            "unitNo": str(row["Unit No."]).strip(),
            "unitNoBankStmt": bank_stmt_unit_no,
            "customerName": customer_name,
            "names": names,
            "mobiles": mobiles,
            "upiIds": upi_ids,
            "accountNumbers": account_numbers,
            "accountLast4s": account_last4s,
            "ifscCodes": ifsc_codes,
            "bankNames": bank_names,
        }

        if unit_key not in raw_by_unit:
            raw_by_unit[unit_key] = entry
        else:
            existing = raw_by_unit[unit_key]
            for field in ["names", "mobiles", "upiIds", "accountNumbers", "accountLast4s", "ifscCodes", "bankNames"]:
                existing[field] |= entry[field]
            if not existing["unitNoBankStmt"] and entry["unitNoBankStmt"]:
                existing["unitNoBankStmt"] = entry["unitNoBankStmt"]
            if not existing["customerName"] and entry["customerName"]:
                existing["customerName"] = entry["customerName"]

    raw_entities = list(raw_by_unit.values())

    def _build_generic_set(field_name: str, extra_denylist: set = None) -> set:
        counts = {}
        for e in raw_entities:
            for val in e[field_name]:
                key = val.strip().upper()
                counts.setdefault(key, set()).add(e["unitKey"])
        generic = {val for val, units in counts.items() if len(units) > MAX_UNITS_PER_NAME}
        if extra_denylist:
            generic |= extra_denylist
        return generic

    generic_names = _build_generic_set("names", extra_denylist=KNOWN_BOILERPLATE_NAMES)
    generic_mobiles = _build_generic_set("mobiles")
    generic_upi_ids = _build_generic_set("upiIds")
    generic_account_numbers = _build_generic_set("accountNumbers")
    generic_account_last4s = _build_generic_set("accountLast4s")
    generic_ifsc_codes = _build_generic_set("ifscCodes")

    lookup = {}
    for e in raw_entities:
        entity = {
            "unitNo": e["unitNo"],
            "unitNoBankStmt": e["unitNoBankStmt"],
            "customerName": e["customerName"],
            "names": {v for v in e["names"] if v.strip().upper() not in generic_names},
            "mobiles": {v for v in e["mobiles"] if v.strip().upper() not in generic_mobiles},
            "upiIds": {v for v in e["upiIds"] if v.strip().upper() not in generic_upi_ids},
            "accountNumbers": {v for v in e["accountNumbers"] if v.strip().upper() not in generic_account_numbers},
            "accountLast4s": {v for v in e["accountLast4s"] if v.strip().upper() not in generic_account_last4s},
            "ifscCodes": {v for v in e["ifscCodes"] if v.strip().upper() not in generic_ifsc_codes},
            "bankNames": e["bankNames"],
        }
        lookup[e["unitKey"]] = entity

    return lookup


def get_all_entities() -> list:
    """
    Returns every unit's entity fingerprint from matched_output.xlsx,
    regardless of whether that unit has a collection increment in the
    current month's MIS. Used to suggest a likely unit for transactions
    that don't match any unit in this month's active comparison set.
    """
    return list(load_entity_reference().values())


def get_entity_for_unit(unit_no: str) -> dict | None:
    """
    Returns the entity fingerprint (names, mobiles, UPI IDs, account
    numbers, account last-4s) for a given unit number, or None if that
    unit isn't found in entity_reference.xlsx.
    """
    lookup = load_entity_reference()
    return lookup.get(normalize_unit(unit_no))


MIN_NAME_LENGTH = 4  # names/inferred-names shorter than this cause false positives
MIN_NAME_WORDS = 2   # single-word names (e.g. just "Seema") are too generic to trust

# Generic legal/business suffix words. A name fragment made up ENTIRELY of
# these (e.g. "Pvt Ltd", "Private Limited") can accidentally match almost
# any company's transaction narration and must never be trusted alone —
# even if it technically has 2+ words and isn't caught by the unit-count
# generic filter (e.g. it came from a comma inside one company's own name
# being wrongly split into fragments).
print("=== entity_matcher.py loaded — LEGAL_SUFFIX_WORDS fix v2 ===")

LEGAL_SUFFIX_WORDS = {
    "PVT", "PVT.", "LTD", "LTD.", "LIMITED", "PRIVATE", "LLP", "INC",
    "INC.", "CO", "CO.", "COMPANY", "CORP", "CORPORATION", "GROUP",
    "HOLDINGS", "PROPERTIES", "DEVELOPERS", "INFRA", "PROJECTS",
    "REALTORS", "REALTY", "ESTATE", "ESTATES", "VENTURES", "ENTERPRISES",
    "COLLECTION", "ACCOUNT",
}


def _is_generic_suffix_only(name: str) -> bool:
    """Returns True if every word in the name is a generic legal/business
    suffix word — meaning the fragment carries no real identifying info
    on its own (e.g. 'Pvt Ltd', 'Private Limited')."""
    words = [w.strip(".,").upper() for w in name.split() if w.strip(".,")]
    if not words:
        return True
    return all(w in LEGAL_SUFFIX_WORDS for w in words)


def _contains_known_boilerplate(name: str) -> bool:
    """Returns True if the name CONTAINS the collection account's own
    name/brand anywhere within it — e.g. 'Kv Developers', 'K V Developers
    Pvt Ltd', 'KV Development', with or without a space between K and V,
    and regardless of what other words surround it. This catches stray
    data-entry fragments (like 'Kv Developers' mixed into a customer's
    Additional Names cell) that an exact-phrase list would miss."""
    upper = name.upper()
    if re.search(r'K\s*V\s*DEVELOP', upper):
        return True
    for phrase in KNOWN_BOILERPLATE_NAMES:
        if phrase in upper:
            return True
    return False


VALID_MOBILE_LENGTH = 10  # real Indian mobile numbers are 10 digits


def _is_valid_mobile(mobile: str) -> bool:
    digits = "".join(c for c in (mobile or "") if c.isdigit())
    return len(digits) == VALID_MOBILE_LENGTH


def _unit_number_in_narration(narration_upper: str, unit_no: str) -> bool:
    """
    Checks if the entity's OWN unit number appears in the narration text,
    trying common variants (slash, dash, space, no separator) since bank
    narrations write unit numbers inconsistently (e.g. 'T1/111',
    'T1-111', 'T1 111', 'T1111').
    """
    if not unit_no:
        return False
    base = unit_no.strip().upper()
    variants = {
        base,
        base.replace("/", "-"),
        base.replace("-", "/"),
        base.replace("/", " ").replace("-", " "),
        re.sub(r'[/\-\s]', '', base),
    }
    return any(v and v in narration_upper for v in variants)


def match_transaction_to_entity(narration: str, account_number: str, entity: dict) -> bool:
    """
    Checks if a single bank transaction (narration text + account number,
    if available) matches the given entity, using this priority:
    1. Unit Number (own unit number literally appears in narration)
    2. Account Number (exact)
    3. UPI ID (substring, only if reasonably specific)
    4. Mobile number (exact 10-digit match only — rejects junk like '0')
    5. Name variants (substring, only names >= MIN_NAME_LENGTH chars)
    """
    narration = _normalize_whitespace(narration).upper()
    account_number = (account_number or "").strip()

    # 1. Unit Number match — strongest possible signal
    if _unit_number_in_narration(narration, entity.get("unitNo", "")):
        return True

    # 2. Account Number match
    if account_number and account_number in entity["accountNumbers"]:
        return True
    # account_last4 matching intentionally removed — see note in the
    # detailed version of this function.

    # 2. UPI ID match — only trust UPI IDs with a reasonable length to avoid junk
    for upi in entity["upiIds"]:
        if upi and len(upi) >= 8 and upi.upper() in narration:
            return True

    # 3. Mobile number match — only valid 10-digit numbers, exact substring
    for mobile in entity["mobiles"]:
        if _is_valid_mobile(mobile) and mobile in narration:
            return True

    # 4. Name variant match — reject short/generic names that cause false positives
    for name in entity["names"]:
        clean_name = _normalize_whitespace(name)
        if len(clean_name) < MIN_NAME_LENGTH:
            continue
        if len(clean_name.split()) < MIN_NAME_WORDS:
            continue  # single-word names are too generic (e.g. "Seema")
        if _is_generic_suffix_only(clean_name):
            continue  # e.g. "Pvt Ltd" alone — carries no real identity
        if _contains_known_boilerplate(clean_name):
            continue  # e.g. "Kv Developers" — the collection account's own name
        if clean_name.upper() in narration:
            return True

    return False


def _normalize_whitespace(text: str) -> str:
    """Collapses multiple spaces/tabs into one, so 'KUMAR  SHARMA' matches 'KUMAR SHARMA'."""
    return re.sub(r'\s+', ' ', (text or "").strip())


IFSC_PATTERN = re.compile(r'\b([A-Z]{4}0[A-Z0-9]{6})\b')
TRAILING_DIGITS_PATTERN = re.compile(r'(\d{4})\b(?!.*\d)')


def _extract_ifsc(narration_upper: str) -> str | None:
    match = IFSC_PATTERN.search(narration_upper)
    return match.group(1) if match else None


def _extract_account_last4(narration_upper: str) -> str | None:
    match = TRAILING_DIGITS_PATTERN.search(narration_upper)
    return match.group(1) if match else None


def match_transaction_to_entity_detailed(narration: str, account_number: str, entity: dict) -> str | None:
    """
    Same matching logic as match_transaction_to_entity, but returns WHICH
    signal caused the match ('unit_number', 'account', 'account_last4',
    'upi', 'mobile', 'name', 'ifsc') instead of just True/False — or None
    if no match. Used for debugging match quality.
    """
    narration = _normalize_whitespace(narration).upper()
    account_number = (account_number or "").strip()

    if _unit_number_in_narration(narration, entity.get("unitNo", "")):
        return "unit_number"

    if account_number and account_number in entity["accountNumbers"]:
        return "account"

    last4 = _extract_account_last4(narration)
    if last4 and last4 in entity["accountLast4s"]:
        return "account_last4"

    for upi in entity["upiIds"]:
        if upi and len(upi) >= 8 and upi.upper() in narration:
            return "upi"

    for mobile in entity["mobiles"]:
        if _is_valid_mobile(mobile) and mobile in narration:
            return "mobile"

    for name in entity["names"]:
        clean_name = _normalize_whitespace(name)
        if len(clean_name) < MIN_NAME_LENGTH:
            continue
        if len(clean_name.split()) < MIN_NAME_WORDS:
            continue
        if _is_generic_suffix_only(clean_name):
            continue
        if _contains_known_boilerplate(clean_name):
            continue
        if clean_name.upper() in narration:
            return "name"

    ifsc = _extract_ifsc(narration)
    if ifsc and ifsc in entity["ifscCodes"]:
        return "ifsc"

    return None