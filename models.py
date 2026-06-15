from datetime import time, date
from sqlalchemy import String, Date, Time
from db import Base
from sqlalchemy import String
from sqlalchemy.orm import Mapped, mapped_column

class User(Base):
    __tablename__ = 'users'
    id : Mapped[int] = mapped_column(primary_key=True)
    role : Mapped[int] = mapped_column(nullable=False)
    email : Mapped[str] = mapped_column(String(300), nullable=False)
    password : Mapped[str] = mapped_column(String(300), nullable=False)

class FileUploadMetaData(Base):
    __tablename__ = 'file_upload_meta_data'
    id : Mapped[int] = mapped_column(primary_key=True)
    user : Mapped[str] = mapped_column(String(100))
    date_uploaded: Mapped[date] = mapped_column(Date) 
    time_uploaded: Mapped[time] = mapped_column(Time)
    file_uploaded : Mapped[str] = mapped_column(String(300))
    final_procurement_file : Mapped[str] = mapped_column(String(300))

# class formula(Base):
#     __tablename__ = 'formula'

#     id : Mapped[int] = mapped_column(primary_key=True)
#     product_id : Mapped[str] = mapped_column(String(30),nullable=False)
#     material_id : Mapped[str] = mapped_column(String(30),nullable=False)
#     material_qty : Mapped[int] = mapped_column(nullable=False)
#     material_req : Mapped[int] = mapped_column(nullable=False)
#     material_metric : Mapped[str] = mapped_column(String(10), nullable=False)
#     material_cost : Mapped[int] = mapped_column(nullable=False) # per material_req



# [CBS_BOM] '622'