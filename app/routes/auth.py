import os
import resend
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

class ForgotPasswordRequest(BaseModel):
    email: str

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

@router.post("/forgot-password")
async def forgot_password(request: ForgotPasswordRequest):
    try:
        user = auth.get_user_by_email(request.email)
        reset_link = auth.generate_password_reset_link(request.email)

        html = f"""
        <html><body style="font-family:Arial,sans-serif;padding:20px;">
        <div style="max-width:500px;margin:auto;border:1px solid #e5e7eb;border-radius:12px;padding:32px;">
          <div style="text-align:center;margin-bottom:24px;">
            <div style="background:#2563eb;width:48px;height:48px;border-radius:12px;display:inline-flex;align-items:center;justify-content:center;">
              <span style="color:white;font-size:24px;font-weight:bold;">T</span>
            </div>
            <h2 style="color:#1f2937;margin-top:12px;">TAM Tool</h2>
          </div>
          <p style="color:#374151;">Hi {user.display_name or "there"},</p>
          <p style="color:#374151;">You requested a password reset for your TAM Tool account.</p>
          <div style="text-align:center;margin:32px 0;">
            <a href="{reset_link}" 
               style="background:#2563eb;color:white;padding:12px 32px;border-radius:8px;text-decoration:none;font-weight:600;">
              Reset Password
            </a>
          </div>
          <p style="color:#6b7280;font-size:13px;">This link expires in 1 hour. If you didn't request this, you can safely ignore this email.</p>
          <hr style="border:none;border-top:1px solid #e5e7eb;margin:24px 0;">
          <p style="color:#9ca3af;font-size:12px;text-align:center;">© 2026 TAM Tool. All rights reserved.</p>
        </div>
        </body></html>
        """

        resend.api_key = os.getenv("RESEND_API_KEY")
        resend.Emails.send({
            "from": "TAM Tool <onboarding@resend.dev>",
            "to": request.email,
            "subject": "Reset your TAM Tool password",
            "html": html
        })

        print(f"✅ Password reset email sent to {request.email}")
        return {"message": "Password reset email sent successfully"}

    except auth.UserNotFoundError:
        return {"message": "If this email exists, a reset link has been sent"}
    except Exception as e:
        print(f"ERROR in forgot_password: {e}")
        raise HTTPException(status_code=500, detail="Failed to send reset email")

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
            firebase_user = auth.get_user(uid)
            is_admin_email = firebase_user.email.lower().startswith("admin@")
            if is_admin_email and not user_data.get("isAdmin"):
                db.collection("users").document(uid).update({
                    "isAdmin": True,
                    "status": "ACTIVE"
                })
                user_data["isAdmin"] = True
                user_data["status"] = "ACTIVE"
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
