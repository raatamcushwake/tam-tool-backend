import firebase_admin
from firebase_admin import credentials, firestore, auth

_app = None

def init_firebase():
    global _app
    if not firebase_admin._apps:
        cred = credentials.Certificate("serviceAccountKey.json")
        _app = firebase_admin.initialize_app(cred)
    return _app

def get_firestore():
    return firestore.client()

def get_auth():
    return auth