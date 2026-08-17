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

class TowerConfig(BaseModel):
    name: str
    constructionArea: Optional[float] = 0
    basements: Optional[float] = 0
    ground: Optional[float] = 0
    stilt: Optional[float] = 0
    podiums: Optional[float] = 0
    serviceFloor: Optional[float] = 0
    upperFloors: Optional[float] = 0

class TowerConfigRequest(BaseModel):
    totalTowerArea: Optional[float] = 0
    nonTowerArea: Optional[float] = 0
    towers: list[TowerConfig] = []


@router.get("/{project_id}/tower-config")
async def get_tower_config(project_id: str):
    try:
        db = get_firestore()
        doc = db.collection("projects").document(project_id).get()
        if not doc.exists:
            raise HTTPException(status_code=404, detail="Project not found")
        data = doc.to_dict()
        tower_config = data.get("towerConfig")
        if not tower_config:
            raise HTTPException(status_code=404, detail="Tower config not set yet")
        return tower_config
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/{project_id}/tower-config")
async def save_tower_config(project_id: str, request: TowerConfigRequest):
    try:
        db = get_firestore()
        project_ref = db.collection("projects").document(project_id)
        doc = project_ref.get()
        if not doc.exists:
            raise HTTPException(status_code=404, detail="Project not found")

        tower_config = {
            "totalTowerArea": request.totalTowerArea,
            "nonTowerArea": request.nonTowerArea,
            "towers": [t.dict() for t in request.towers],
            "status": "locked",
        }
        project_ref.update({"towerConfig": tower_config})
        return tower_config
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.patch("/{project_id}/tower-config/unlock")
async def unlock_tower_config(project_id: str):
    try:
        db = get_firestore()
        project_ref = db.collection("projects").document(project_id)
        doc = project_ref.get()
        if not doc.exists:
            raise HTTPException(status_code=404, detail="Project not found")
        project_ref.update({"towerConfig.status": "draft"})
        return {"status": "draft"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

class WeightageActivity(BaseModel):
    name: str
    values: dict = {}
    costWeightage: Optional[float] = 0
    remarks: Optional[str] = ""

class WeightagePackage(BaseModel):
    name: str
    activities: list[WeightageActivity] = []

class WeightageConfigRequest(BaseModel):
    packages: list[WeightagePackage] = []


@router.get("/{project_id}/weightage-config")
async def get_weightage_config(project_id: str):
    try:
        db = get_firestore()
        doc = db.collection("projects").document(project_id).get()
        if not doc.exists:
            raise HTTPException(status_code=404, detail="Project not found")
        config = doc.to_dict().get("weightageConfig")
        if not config:
            raise HTTPException(status_code=404, detail="Weightage config not set yet")
        return config
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/{project_id}/weightage-config")
async def save_weightage_config(project_id: str, request: WeightageConfigRequest):
    try:
        db = get_firestore()
        project_ref = db.collection("projects").document(project_id)
        if not project_ref.get().exists:
            raise HTTPException(status_code=404, detail="Project not found")
        config = {
            "packages": [p.dict() for p in request.packages],
            "status": "locked",
        }
        project_ref.update({"weightageConfig": config})
        return config
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.patch("/{project_id}/weightage-config/unlock")
async def unlock_weightage_config(project_id: str):
    try:
        db = get_firestore()
        project_ref = db.collection("projects").document(project_id)
        if not project_ref.get().exists:
            raise HTTPException(status_code=404, detail="Project not found")
        project_ref.update({"weightageConfig.status": "draft"})
        return {"status": "draft"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

class ActivityMatrixStructureRequest(BaseModel):
    floors: list[dict] = []
    applicability: dict = {}

class ActivityMatrixValuesRequest(BaseModel):
    values: dict = {}


@router.get("/{project_id}/activity-matrix/{tower_name}")
async def get_activity_matrix(project_id: str, tower_name: str):
    try:
        db = get_firestore()
        doc = db.collection("projects").document(project_id).get()
        if not doc.exists:
            raise HTTPException(status_code=404, detail="Project not found")
        matrix = doc.to_dict().get("activityMatrix", {}).get(tower_name)
        if not matrix:
            raise HTTPException(status_code=404, detail="Activity matrix not set yet")
        return matrix
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/{project_id}/activity-matrix/{tower_name}")
@router.post("/{project_id}/activity-matrix/{tower_name}")
async def save_activity_matrix_structure(project_id: str, tower_name: str, request: ActivityMatrixStructureRequest):
    try:
        db = get_firestore()
        project_ref = db.collection("projects").document(project_id)
        doc = project_ref.get()
        if not doc.exists:
            raise HTTPException(status_code=404, detail="Project not found")

        existing = doc.to_dict().get("activityMatrix", {}).get(tower_name, {})
        update_data = {
            f"activityMatrix.{tower_name}.floors": request.floors,
            f"activityMatrix.{tower_name}.applicability": request.applicability,
            f"activityMatrix.{tower_name}.status": "locked",
        }
        # First-ever save for this tower — make sure `values` exists so the frontend never reads undefined
        if "values" not in existing:
            update_data[f"activityMatrix.{tower_name}.values"] = {}

        project_ref.update(update_data)
        doc2 = project_ref.get().to_dict()
        return doc2.get("activityMatrix", {}).get(tower_name)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.patch("/{project_id}/activity-matrix/{tower_name}/values")
async def save_activity_matrix_values(project_id: str, tower_name: str, request: ActivityMatrixValuesRequest):
    try:
        db = get_firestore()
        project_ref = db.collection("projects").document(project_id)
        if not project_ref.get().exists:
            raise HTTPException(status_code=404, detail="Project not found")
        project_ref.update({f"activityMatrix.{tower_name}.values": request.values})
        return {"values": request.values}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.patch("/{project_id}/activity-matrix/{tower_name}/unlock")
async def unlock_activity_matrix(project_id: str, tower_name: str):
    try:
        db = get_firestore()
        project_ref = db.collection("projects").document(project_id)
        if not project_ref.get().exists:
            raise HTTPException(status_code=404, detail="Project not found")
        project_ref.update({f"activityMatrix.{tower_name}.status": "draft"})
        return {"status": "draft"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))