from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.core.firebase import init_firebase
from app.routes import auth, project, mis, approval, cost, compliance
from app.routes.MISSanityCheck import router as mis_sanity_router
from app.routes.MISAnalysis import router as mis_analysis_router

print("DEBUG: Starting app initialization...")

# Initialize Firebase
try:
    init_firebase()
    print("DEBUG: Firebase initialized in main.py")
except Exception as e:
    print(f"ERROR: Failed to initialize Firebase in main.py: {e}")

app = FastAPI(
    title="TAM Tool API",
    description="Backend API for TAM Tool Dashboard",
    version="1.0.0"
)

# CORS Middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Register Routes
app.include_router(auth.router, prefix="/api/auth", tags=["Auth"])
app.include_router(project.router, prefix="/api/projects", tags=["Projects"])
app.include_router(mis.router, prefix="/api/mis", tags=["MIS"])
app.include_router(approval.router, prefix="/api/approvals", tags=["Approvals"])
app.include_router(cost.router, prefix="/api/cost", tags=["Cost"])
app.include_router(compliance.router, prefix="/api/compliance", tags=["Compliance"])
app.include_router(mis_sanity_router, prefix="/api/mis-sanity", tags=["MIS Sanity"])
app.include_router(mis_analysis_router, prefix="/api/mis-analysis", tags=["MIS Analysis"])


@app.get("/")
def root():
    return {"message": "TAM Tool API is running!"}

@app.get("/health")
def health():
    return {"status": "healthy", "version": "1.0.0"}