import firebase_admin
from firebase_admin import credentials, firestore, auth
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
        _app = firebase_admin.initialize_app(cred)
    return _app

def get_firestore():
    return firestore.client()

def get_auth():
    return auth
