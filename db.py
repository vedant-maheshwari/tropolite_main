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

try:
    with engine_admin.connect() as conn:
        print("Connected to Admin DB")
        # Ensure permissions column exists for tab-based access
        try:
            conn.execute(text("ALTER TABLE users ADD COLUMN permissions VARCHAR(500) DEFAULT 'final_order,per_fg_bom,per_fg_cost,px_item_cost,final_fg_price'"))
            conn.commit()
            print("Added permissions column to users table.")
        except Exception as e:
            # Column likely already exists
            conn.rollback()
except Exception as e:
    print(f"Admin DB Connection Error: {e}")

def get_db_admin():
    db_admin = SessionLocal_admin()
    try:
        yield db_admin
    finally:
        db_admin.close()


# engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})

# SAP MSSQL engine — strictly limited pool to avoid overloading SAP
# pool_size: max persistent connections kept alive
# max_overflow: extra connections allowed beyond pool_size (then blocked)
# pool_timeout: seconds to wait for a free connection before raising an error
# pool_recycle: recycle connections after 30 min to avoid stale/broken sockets
engine = create_engine(
    DATABASE_URL,
    pool_size=3,
    max_overflow=2,
    pool_timeout=30,
    pool_recycle=1800,
    pool_pre_ping=True,   # validates connection health before use
)
SessionLocal = sessionmaker(autoflush=False, autocommit=False, bind=engine)
Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# SAP MD MSSQL engine — same conservative pool settings
engine_md = create_engine(
    DATABASE_URL_MD,
    pool_size=2,
    max_overflow=1,
    pool_timeout=30,
    pool_recycle=1800,
    pool_pre_ping=True,
)
SessionLocalMD = sessionmaker(autoflush=False, autocommit=False, bind=engine_md)

def get_db_md():
    db_md = SessionLocalMD()
    try:
        yield db_md
    finally:
        db_md.close()
    