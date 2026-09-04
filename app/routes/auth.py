import os
from fastapi import APIRouter, HTTPException
from firebase_admin import auth, firestore
from app.core.firebase import get_firestore
from pydantic import BaseModel

router = APIRouter()

SERVICE_LABELS = {
    "continuous-monitoring": "Continuous Monitoring",
    "periodic-monitoring": "Periodic Monitoring",
    "tdd": "TDD",
    "lie": "Lender Independent Engineering",
}

def enrich_project_roles(db, project_roles):
    """For each role, expand into one entry per service enabled on that project."""
    enriched = []
    for pr in project_roles:
        project_id = pr.get("projectId")
        enabled_services = {}
        if project_id:
            try:
                proj_doc = db.collection("projects").document(project_id).get()
                if proj_doc.exists:
                    proj_data = proj_doc.to_dict()
                    enabled_services = proj_data.get("enabledServices", {})
            except Exception as e:
                print(f"WARNING: could not fetch project {project_id}: {e}")

        if enabled_services:
            for service_key, modules in enabled_services.items():
                enriched.append({
                    **pr,
                    "serviceKey": service_key,
                    "serviceLabel": SERVICE_LABELS.get(service_key, service_key),
                    "enabledModules": modules or [],
                })
        else:
            # No services enabled on the project yet — keep the role visible as-is.
            enriched.append({
                **pr,
                "serviceLabel": pr.get("serviceLabel"),
                "serviceKey": pr.get("serviceKey"),
                "enabledModules": [],
            })
    return enriched

MOCK_MODE = os.environ.get("MOCK_FIRESTORE", "false").lower() == "true"

class CreateUserRequest(BaseModel):
    email: str
    password: str
    name: str
    phone: str = ""

class RegisterRequest(BaseModel):
    email: str
    password: str
    name: str
    phone: str = ""

class ForgotPasswordRequest(BaseModel):
    email: str

class ResetPasswordRequest(BaseModel):
    uid: str
    new_password: str

@router.post("/create-user")
async def create_user(request: CreateUserRequest):
    try:
        user = auth.create_user(
            email=request.email,
            password=request.password,
            display_name=request.name
        )
        db = get_firestore()
        db.collection("users").document(user.uid).set({
            "uid": user.uid,
            "name": request.name,
            "email": request.email,
            "phone": request.phone,
            "isAdmin": False,
            "status": "PENDING",
            "projectRoles": [],
            "createdAt": firestore.SERVER_TIMESTAMP
        })
        return {"message": "User created successfully", "uid": user.uid}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.post("/register")
async def register_user(request: RegisterRequest):
    try:
        user = auth.create_user(
            email=request.email,
            password=request.password,
            display_name=request.name
        )
        db = get_firestore()
        db.collection("users").document(user.uid).set({
            "uid": user.uid,
            "name": request.name,
            "email": request.email,
            "phone": request.phone,
            "isAdmin": False,
            "status": "PENDING",
            "projectRoles": [],
            "createdAt": firestore.SERVER_TIMESTAMP
        })
        return {"message": "Registration successful", "uid": user.uid}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/reset-password")
async def reset_password(request: ResetPasswordRequest):
    try:
        auth.update_user(request.uid, password=request.new_password)
        return {"message": "Password reset successfully"}
    except Exception as e:
        print(f"ERROR in reset_password: {e}")
        raise HTTPException(status_code=500, detail="Failed to reset password")
    
@router.post("/forgot-password")
async def forgot_password(request: ForgotPasswordRequest):
    try:
        auth.get_user_by_email(request.email)
        
        import httpx
        async with httpx.AsyncClient() as client:
            response = await client.post(
                f"https://identitytoolkit.googleapis.com/v1/accounts:sendOobCode?key={os.getenv('FIREBASE_WEB_API_KEY')}",
                json={
                    "requestType": "PASSWORD_RESET",
                    "email": request.email
                }
            )
        
        print(f"✅ Password reset email sent to {request.email}")
        return {"message": "Password reset email sent successfully"}

    except auth.UserNotFoundError:
        return {"message": "If this email exists, a reset link has been sent"}
    except Exception as e:
        print(f"ERROR in forgot_password: {e}")
        raise HTTPException(status_code=500, detail="Failed to send reset email")

@router.get("/user/{uid}")
async def get_user(uid: str):
    if MOCK_MODE:
        print(f"DEBUG: MOCK_FIRESTORE active — returning fake profile for {uid}")
        return {
            "uid": uid,
            "name": "Test User",
            "email": "testuser@cushwake.com",
            "phone": "",
            "isAdmin": True,
            "status": "ACTIVE",
            "projectRoles": []
        }
    try:
        db = get_firestore()
        print(f"DEBUG: Fetching user with UID: {uid}")
        user_doc = db.collection("users").document(uid).get()
        print(f"DEBUG: User document exists: {user_doc.exists}")
        print(f"PROJECT BEING USED: {db.project}")
        if user_doc.exists:
            user_data = user_doc.to_dict()
            print(f"DEBUG: User data: {user_data}")
            firebase_user = auth.get_user(uid)
            is_admin_email = firebase_user.email.lower().startswith("admin@")
            if is_admin_email and not user_data.get("isAdmin"):
                db.collection("users").document(uid).update({
                    "isAdmin": True,
                    "status": "ACTIVE"
                })
                user_data["isAdmin"] = True
                user_data["status"] = "ACTIVE"

            user_data["projectRoles"] = enrich_project_roles(db, user_data.get("projectRoles", []))

            return user_data
        else:
            print(f"DEBUG: User not found in Firestore, attempting to create from Firebase Auth")
            try:
                firebase_user = auth.get_user(uid)
                print(f"DEBUG: Found Firebase Auth user: {firebase_user.email}")
                is_admin = firebase_user.email.lower().startswith("admin@")
                initial_status = "ACTIVE" if is_admin else "PENDING"
                db.collection("users").document(uid).set({
                    "uid": uid,
                    "name": firebase_user.display_name or firebase_user.email.split("@")[0],
                    "email": firebase_user.email,
                    "phone": "",
                    "isAdmin": is_admin,
                    "status": initial_status,
                    "projectRoles": [],
                    "createdAt": firestore.SERVER_TIMESTAMP
                })
                return {
                    "uid": uid,
                    "name": firebase_user.display_name or firebase_user.email.split("@")[0],
                    "email": firebase_user.email,
                    "phone": "",
                    "isAdmin": is_admin,
                    "status": initial_status,
                    "projectRoles": []
                }
            except Exception as auth_err:
                print(f"ERROR: Could not get Firebase Auth user: {auth_err}")
                raise HTTPException(status_code=404, detail="User not found")
    except HTTPException:
        raise
    except Exception as e:
        print(f"ERROR in get_user: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/users")
async def get_all_users():
    if MOCK_MODE:
        print("DEBUG: MOCK_FIRESTORE active — returning fake user list")
        return [
            {
                "uid": "GIKFBWhGbtQYWtY7rwPXZxzYf0m1",
                "name": "Test User",
                "email": "testuser@cushwake.com",
                "phone": "",
                "isAdmin": True,
                "status": "ACTIVE",
                "projectRoles": []
            },
            {
                "uid": "fakeuid002",
                "name": "Jane Pending",
                "email": "jane.pending@cushwake.com",
                "phone": "",
                "isAdmin": False,
                "status": "PENDING",
                "projectRoles": []
            },
            {
                "uid": "fakeuid003",
                "name": "John Active",
                "email": "john.active@cushwake.com",
                "phone": "",
                "isAdmin": False,
                "status": "ACTIVE",
                "projectRoles": []
            }
        ]
    try:
        db = get_firestore()
        users = db.collection("users").stream()
        return [u.to_dict() for u in users]
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
@router.patch("/user/{uid}/status")
async def update_user_status(uid: str, body: dict):
    try:
        db = get_firestore()
        db.collection("users").document(uid).update({
            "status": body.get("status")
        })
        return {"message": "User status updated"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

class AssignProjectRequest(BaseModel):
    projectId: str
    projectName: str
    role: str
    serviceKey: str
    serviceLabel: str = ""

@router.post("/user/{uid}/assign-project")
async def assign_project(uid: str, body: AssignProjectRequest):
    try:
        db = get_firestore()
        user_ref = db.collection("users").document(uid)
        user_ref.update({
            "projectRoles": firestore.ArrayUnion([{
                "projectId": body.projectId,
                "projectName": body.projectName,
                "role": body.role,
                "serviceKey": body.serviceKey,
                "serviceLabel": body.serviceLabel,
            }])
        })
        return {"message": "Project assigned successfully"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@router.delete("/user/{uid}/remove-project")
async def remove_project(uid: str, body: dict):
    try:
        db = get_firestore()
        user_ref = db.collection("users").document(uid)
        user_doc = user_ref.get()
        if not user_doc.exists:
            raise HTTPException(status_code=404, detail="User not found")
        current_roles = user_doc.to_dict().get("projectRoles", [])
        updated_roles = [
            r for r in current_roles
            if not (r["projectId"] == body.get("projectId") and r["role"] == body.get("role"))
        ]
        user_ref.update({"projectRoles": updated_roles})
        return {"message": "Project removed successfully"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
