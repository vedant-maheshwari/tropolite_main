from fastapi import HTTPException, FastAPI, Depends, status
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from db import get_db_admin
from sqlalchemy.orm import Session
from datetime import datetime, timedelta
from jose import JWTError, jwt
from passlib.context import CryptContext
from dotenv import load_dotenv
import os
import models

load_dotenv('.env')

SECRET_KEY = os.getenv('SECRET_KEY')
ALGORITHM = os.getenv('ALGORITHM')
ACCESS_TOKEN_EXPIRE = int(os.getenv('ACCESS_TOKEN_EXPIRE'))

oauth2_schema = OAuth2PasswordBearer(tokenUrl='token')
pwd_context = CryptContext(schemes=['bcrypt'], deprecated='auto')

user_db = {
    "user@example.com": {
        "username": "user@example.com",
        "hashed_password": pwd_context.hash("password"),
        "role": "user"
    },
    "admin": {
        "username": "admin",
        "hashed_password": pwd_context.hash("admin123"),
        "role": "admin"
    }
}



def hash_password(password : str):
    return pwd_context.hash(password)

def create_access_token(data : dict):
    to_encode = data.copy()
    expire = datetime.now() + timedelta(minutes=ACCESS_TOKEN_EXPIRE)
    to_encode.update({'exp' : expire})
    return jwt.encode(to_encode, SECRET_KEY, ALGORITHM)

def get_current_user(token : str = Depends(oauth2_schema)):
    try:
        payload = jwt.decode(token, SECRET_KEY, ALGORITHM)
        username = payload.get('sub')
        if username is None:
            raise HTTPException(401, 'invalid token')
    except Exception:
        raise HTTPException(401, 'token is invalid or expired')
    return username

def get_current_admin(token: str = Depends(oauth2_schema)):
    """Dependency that only passes for users with role='admin'."""
    try:
        payload = jwt.decode(token, SECRET_KEY, ALGORITHM)
        username = payload.get('sub')
        if username is None:
            raise HTTPException(401, 'invalid token')
    except Exception:
        raise HTTPException(401, 'token is invalid or expired')
    user = user_db.get(username)
    if not user or user.get('role') != 'admin':
        raise HTTPException(status_code=403, detail='Admin access required.')
    return username

def get_current_admin_for_working(db: Session = Depends(get_db_admin), token: str = Depends(oauth2_schema)):
    try : 
        payload = jwt.decode(token, SECRET_KEY, ALGORITHM)
        email = payload.get('sub')
        if email is None:
            raise HTTPException(401, 'invaild token')
    except Exception as e:
        raise HTTPException(401, 'token is invalid or expired')
    user = db.query(models.User).filter(models.User.email == email).first()
    if user.role == 1 :
        return user
    else:
        raise HTTPException(401, 'Admin only access')
