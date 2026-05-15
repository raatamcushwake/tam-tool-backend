from fastapi import APIRouter, HTTPException
from app.core.firebase import get_firestore
from pydantic import BaseModel
from firebase_admin import firestore

router = APIRouter()

class ApprovalRequest(BaseModel):
    submissionId: str
    projectId: str
    actionBy: str
    stage: str       # REVIEWER or MANAGER
    action: str      # APPROVED or REJECTED
    comment: str = ""

@router.post("/action")
async def take_action(request: ApprovalRequest):
    try:
        db = get_firestore()

        # Save approval record
        ref = db.collection("approvals").document()
        ref.set({
            "id": ref.id,
            "submissionId": request.submissionId,
            "projectId": request.projectId,
            "stage": request.stage,
            "action": request.action,
            "actionBy": request.actionBy,
            "comment": request.comment,
            "actionAt": firestore.SERVER_TIMESTAMP
        })

        # Update submission status
        new_status = ""
        if request.action == "REJECTED":
            new_status = "REJECTED"
        elif request.stage == "REVIEWER" and request.action == "APPROVED":
            new_status = "PENDING_MANAGER"
        elif request.stage == "MANAGER" and request.action == "APPROVED":
            new_status = "APPROVED"

        db.collection("mis_submissions").document(request.submissionId).update({
            "status": new_status
        })

        return {"message": f"Action taken: {request.action}", "newStatus": new_status}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/submission/{submission_id}")
async def get_approvals(submission_id: str):
    try:
        db = get_firestore()
        docs = db.collection("approvals")\
            .where("submissionId", "==", submission_id)\
            .stream()
        return [d.to_dict() for d in docs]
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))