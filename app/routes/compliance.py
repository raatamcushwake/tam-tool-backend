from fastapi import APIRouter, HTTPException
from app.core.firebase import get_firestore
from pydantic import BaseModel
from firebase_admin import firestore

router = APIRouter()

class ComplianceEntry(BaseModel):
    projectId: str
    checkpointName: str
    status: str   # PENDING, COMPLETED, OVERDUE
    dueDate: str
    addedBy: str

@router.post("/add")
async def add_compliance(request: ComplianceEntry):
    try:
        db = get_firestore()
        ref = db.collection("compliance").document()
        ref.set({
            "id": ref.id,
            **request.dict(),
            "createdAt": firestore.SERVER_TIMESTAMP
        })
        return {"message": "Compliance entry added", "id": ref.id}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/project/{project_id}")
async def get_project_compliance(project_id: str):
    try:
        db = get_firestore()
        docs = db.collection("compliance")\
            .where("projectId", "==", project_id)\
            .stream()
        return [d.to_dict() for d in docs]
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))