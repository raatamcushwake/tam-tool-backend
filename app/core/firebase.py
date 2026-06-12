import firebase_admin
from firebase_admin import credentials, firestore, auth, storage
import json
import os

_app = None

def init_firebase():
    global _app
    if not firebase_admin._apps:
        key_json = os.environ.get("FIREBASE_SERVICE_ACCOUNT_KEY")
        if key_json:
            key_dict = json.loads(key_json)
            cred = credentials.Certificate(key_dict)
        else:
            cred = credentials.Certificate("serviceAccountKey.json")
        
        # Read project ID from service account key to set storage bucket
        if key_json:
            project_id = json.loads(key_json).get("project_id", "")
        else:
            with open("serviceAccountKey.json") as f:
                project_id = json.load(f).get("project_id", "")
        
        _app = firebase_admin.initialize_app(cred, {
            'storageBucket': f'{project_id}.appspot.com'
        })
    return _app

def get_firestore():
    return firestore.client()

def get_auth():
    return auth

def get_storage_bucket():
    return storage.bucket()