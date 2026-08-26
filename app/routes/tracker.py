import os
from fastapi import APIRouter, HTTPException
from app.core.firebase import get_firestore

router = APIRouter()

MOCK_MODE = os.environ.get("MOCK_FIRESTORE", "false").lower() == "true"


@router.get("/assignments")
async def get_assignments():
    """
    Read-only list of all tam_assignments docs.

    tam_assignments itself does NOT store region (only projects does),
    so we join it in per-project, same pattern as enrich_project_roles()
    in auth.py. tam_assignments also does NOT currently have a `status`
    field on any doc written by assignRole() in Valuation Services --
    absence of status is treated as "OPEN" here, display-only.
    """
    if MOCK_MODE:
        print("DEBUG: MOCK_FIRESTORE active — returning fake assignment list")
        return [
            {
                "id": "proj123_uid456",
                "projectId": "proj123",
                "projectName": "Sample Project",
                "userId": "uid456",
                "userEmail": "maker@cushwake.com",
                "userName": "Test Maker",
                "role": "MAKER",
                "serviceKey": "CM",
                "serviceLabel": "Cost Management",
                "region": "WES",
                "assignedBy": "admin@cushwake.com",
                "assignedAt": "2026-01-01T00:00:00Z",
                "status": "OPEN",
            }
        ]
    try:
        db = get_firestore()
        docs = db.collection("tam_assignments").stream()
        result = []
        project_region_cache = {}

        for d in docs:
            data = d.to_dict()
            project_id = data.get("projectId")
            region = None

            if project_id:
                if project_id not in project_region_cache:
                    try:
                        proj_doc = db.collection("projects").document(project_id).get()
                        project_region_cache[project_id] = (
                            proj_doc.to_dict().get("region") if proj_doc.exists else None
                        )
                    except Exception as e:
                        print(f"WARNING: could not fetch project {project_id}: {e}")
                        project_region_cache[project_id] = None
                region = project_region_cache[project_id]

            result.append({
                "id": d.id,
                "projectId": data.get("projectId"),
                "projectName": data.get("projectName"),
                "userId": data.get("userId"),
                "userEmail": data.get("userEmail"),
                "userName": data.get("userName"),
                "role": data.get("role"),
                "serviceKey": data.get("serviceKey"),
                "serviceLabel": data.get("serviceLabel"),
                "region": region,
                "assignedBy": data.get("assignedBy"),
                "assignedAt": data.get("assignedAt"),
                "status": data.get("status", "OPEN"),
            })

        return result
    except Exception as e:
        print(f"ERROR in get_assignments: {type(e).__name__}: {e}")
        raise HTTPException(status_code=400, detail=str(e))


@router.patch("/assignments/{assignment_id}/close")
async def close_assignment(assignment_id: str):
    """
    The ONLY write this codebase makes to tam_assignments — flips
    status to CLOSED. assignment_id is the composite `{projectId}_{userId}`
    doc id used by Valuation Services' assignRole(). No other field on
    this doc should ever be touched from TAM Tool.
    """
    if MOCK_MODE:
        print(f"DEBUG: MOCK_FIRESTORE active — pretending to close {assignment_id}")
        return {"message": "Assignment closed (mock)"}
    try:
        db = get_firestore()
        ref = db.collection("tam_assignments").document(assignment_id)
        doc = ref.get()
        if not doc.exists:
            raise HTTPException(status_code=404, detail="Assignment not found")
        ref.update({"status": "CLOSED"})
        return {"message": "Assignment closed"}
    except HTTPException:
        raise
    except Exception as e:
        print(f"ERROR in close_assignment: {type(e).__name__}: {e}")
        raise HTTPException(status_code=400, detail=str(e))
