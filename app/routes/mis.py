from fastapi import APIRouter, HTTPException
from app.core.firebase import get_firestore
from pydantic import BaseModel
from firebase_admin import firestore

router = APIRouter()

class MISSubmitRequest(BaseModel):
    projectId: str
    submittedBy: str
    data: dict

@router.post("/submit")
async def submit_mis(request: MISSubmitRequest):
    try:
        db = get_firestore()
        ref = db.collection("mis_submissions").document()
        ref.set({
            "id": ref.id,
            "projectId": request.projectId,
            "submittedBy": request.submittedBy,
            "data": request.data,
            "status": "PENDING_REVIEW",
            "submittedAt": firestore.SERVER_TIMESTAMP
        })
        return {"message": "MIS submitted", "id": ref.id}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/project/{project_id}")
async def get_project_mis(project_id: str):
    try:
        db = get_firestore()
        docs = db.collection("mis_submissions")\
            .where("projectId", "==", project_id)\
            .stream()
        return [d.to_dict() for d in docs]
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/all")
async def get_all_mis():
    try:
        db = get_firestore()
        docs = db.collection("mis_submissions").stream()
        return [d.to_dict() for d in docs]
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))