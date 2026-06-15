import os
from sqlalchemy.orm import Session
from db import SessionLocal_admin
from models import User
from auth import hash_password

def create_admin_user(email: str, password: str, role: int = 1):
    db = SessionLocal_admin()
    try:
        # Check if user already exists
        existing_user = db.query(User).filter(User.email == email).first()
        if existing_user:
            print(f"User with email {email} already exists.")
            return

        # Hash the password
        hashed_pw = hash_password(password)

        # Create new user
        new_user = User(
            email=email,
            password=hashed_pw,
            role=role
        )
        
        db.add(new_user)
        db.commit()
        db.refresh(new_user)
        
        print(f"Successfully created user: {email} with role: {role}")
        
    except Exception as e:
        db.rollback()
        print(f"Error creating user: {e}")
    finally:
        db.close()

if __name__ == "__main__":
    # You can change these details to whatever you'd like
    NEW_USER_EMAIL = "admin@example.com"
    NEW_USER_PASSWORD = "securepassword123"
    
    # role=1 usually denotes an admin in your app, change if necessary
    create_admin_user(NEW_USER_EMAIL, NEW_USER_PASSWORD, role=1)
