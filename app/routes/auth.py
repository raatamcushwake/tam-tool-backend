from fastapi import APIRouter, HTTPException
from firebase_admin import auth, firestore
from app.core.firebase import get_firestore
from pydantic import BaseModel

router = APIRouter()

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

@router.get("/user/{uid}")
async def get_user(uid: str):
    try:
        db = get_firestore()
        print(f"DEBUG: Fetching user with UID: {uid}")
        user_doc = db.collection("users").document(uid).get()
        print(f"DEBUG: User document exists: {user_doc.exists}")
        

        print(f"PROJECT BEING USED: {db.project}")
        if user_doc.exists:
            user_data = user_doc.to_dict()
            print(f"DEBUG: User data: {user_data}")
            
            # Check if this is an admin user and update if needed
            firebase_user = auth.get_user(uid)
            is_admin_email = firebase_user.email.lower().startswith("admin@")
            
            if is_admin_email and not user_data.get("isAdmin"):
                print(f"DEBUG: Admin user detected, updating document to set isAdmin=true and status=ACTIVE")
                db.collection("users").document(uid).update({
                    "isAdmin": True,
                    "status": "ACTIVE"
                })
                user_data["isAdmin"] = True
                user_data["status"] = "ACTIVE"
            
            return user_data
        else:
            # User doesn't exist in Firestore, try to get from Firebase Auth and create
            print(f"DEBUG: User not found in Firestore, attempting to create from Firebase Auth")
            try:
                firebase_user = auth.get_user(uid)
                print(f"DEBUG: Found Firebase Auth user: {firebase_user.email}")
                
                # Check if user is admin
                is_admin = firebase_user.email.lower().startswith("admin@")
                initial_status = "ACTIVE" if is_admin else "PENDING"
                
                print(f"DEBUG: Setting isAdmin={is_admin}, status={initial_status}")
                
                # Create user document in Firestore
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
                print(f"DEBUG: Created new user document for {uid}")
                
                # Return the created user data (without SERVER_TIMESTAMP)
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

@router.post("/user/{uid}/assign-project")
async def assign_project(uid: str, body: AssignProjectRequest):
    try:
        db = get_firestore()
        user_ref = db.collection("users").document(uid)
        user_ref.update({
            "projectRoles": firestore.ArrayUnion([{
                "projectId": body.projectId,
                "projectName": body.projectName,
                "role": body.role
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