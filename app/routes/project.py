from fastapi import APIRouter, HTTPException
from app.core.firebase import get_firestore
from pydantic import BaseModel
from typing import Optional
from firebase_admin import firestore

router = APIRouter()


class SimpleProjectRequest(BaseModel):
    name: str
    description: str = ""

@router.post("")
async def create_project_simple(request: SimpleProjectRequest):
    try:
        db = get_firestore()
        project_ref = db.collection("projects").document()
        project_data = {
            "id": project_ref.id,
            "name": request.name,
            "projectName": request.name,
            "description": request.description,
            "status": "ACTIVE",
            "members": {},
            "createdAt": firestore.SERVER_TIMESTAMP
        }
        project_ref.set(project_data)
        return {"message": "Project created", "id": project_ref.id}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
class ProjectRequest(BaseModel):
    name: str
    description: str = ""
    makerId: str
    reviewerId: str
    managerId: str

class AssignRoleRequest(BaseModel):
    userId: str
    projectId: str
    role: str  # MAKER, REVIEWER, MANAGER

@router.post("/create")
async def create_project(request: ProjectRequest):
    try:
        db = get_firestore()
        project_ref = db.collection("projects").document()
        project_data = {
            "id": project_ref.id,
            "name": request.name,
            "projectName": request.name,
            "description": request.description,
            "status": "ACTIVE",
            "members": {
                "makerId": request.makerId,
                "reviewerId": request.reviewerId,
                "managerId": request.managerId
            },
            "createdAt": firestore.SERVER_TIMESTAMP
        }
        project_ref.set(project_data)

        # Assign roles to users
        for uid, role in [
            (request.makerId, "MAKER"),
            (request.reviewerId, "REVIEWER"),
            (request.managerId, "MANAGER")
        ]:
            user_ref = db.collection("users").document(uid)
            user_ref.update({
                "projectRoles": firestore.ArrayUnion([{
                    "projectId": project_ref.id,
                    "role": role
                }])
            })

        return {"message": "Project created", "projectId": project_ref.id}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("")
async def get_all_projects():
    try:
        db = get_firestore()
        projects = db.collection("projects").stream()
        return [p.to_dict() for p in projects]
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/user/{uid}")
async def get_user_projects(uid: str):
    try:
        db = get_firestore()
        user_doc = db.collection("users").document(uid).get()
        if not user_doc.exists:
            raise HTTPException(status_code=404, detail="User not found")
        user_data = user_doc.to_dict()
        project_roles = user_data.get("projectRoles", [])
        projects = []
        for pr in project_roles:
            proj = db.collection("projects").document(pr["projectId"]).get()
            if proj.exists:
                data = proj.to_dict()
                data["myRole"] = pr["role"]
                projects.append(data)
        return projects
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/{project_id}")
async def get_project(project_id: str):
    try:
        db = get_firestore()
        doc = db.collection("projects").document(project_id).get()
        if not doc.exists:
            raise HTTPException(status_code=404, detail="Project not found")
        return doc.to_dict()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))