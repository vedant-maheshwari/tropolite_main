from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base, Session
from dotenv import load_dotenv
from sqlalchemy import text
import os

load_dotenv('.env')
DATABASE_URL = os.getenv("DATABASE_URL")
DATABASE_URL_MD = os.getenv("DATABASE_URL_MD")

DATABASE_ADMIN_URL = os.getenv("DATABASE_ADMIN_DATABASE_URL") or os.getenv("DATABASE_ADMIN_POSTGRES_URL") or os.getenv("DATABASE_ADMIN_PRISMA_DATABASE_URL")

if DATABASE_ADMIN_URL and DATABASE_ADMIN_URL.startswith("postgres://"):
    DATABASE_ADMIN_URL = DATABASE_ADMIN_URL.replace("postgres://", "postgresql://", 1)

if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

# DATABASE_ADMIN_URL = os.getenv("DATABASE_ADMIN_URL")

print(DATABASE_URL)
print(DATABASE_URL_MD)
print(DATABASE_ADMIN_URL)

if not DATABASE_ADMIN_URL:
    raise RuntimeError("DATABASE_ADMIN_URL is not set. Check your environment variables.")

engine_admin = create_engine(DATABASE_ADMIN_URL)
SessionLocal_admin = sessionmaker(autoflush=False, autocommit=False, bind=engine_admin)

# try:
#     with engine_admin.connect() as conn:
#         print("Connected")

#         db_name = conn.execute(
#             text("SELECT DB_NAME()")
#         ).scalar()

#         print("Database:", db_name)

# except Exception as e:
#     print(e)

def get_db_admin():
    db_admin = SessionLocal_admin()
    try:
        yield db_admin
    finally:
        db_admin.close()


# engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autoflush=False, autocommit=False, bind=engine)
Base = declarative_base()

try:
    with engine.connect() as conn:
        print("Connected")
        print("Database:", engine.url.database)

except Exception as e:
    print(e)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


engine_md = create_engine(DATABASE_URL_MD)
SessionLocalMD = sessionmaker(autoflush=False, autocommit=False,bind=engine_md)

try:
    with engine_md.connect() as conn_md:
        print('Connected to MD')
        db_name = conn_md.execute(
            text('SELECT DB_NAME()')
        ).scalar()

        print("MD Database :", db_name)
except Exception as e:
    print(e)

def get_db_md():
    db_md = SessionLocalMD()
    try:
        yield db_md
    finally:
        db_md.close()
    