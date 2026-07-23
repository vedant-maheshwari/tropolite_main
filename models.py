from datetime import time, date, datetime
from sqlalchemy import String, Date, Time, DateTime, Float
from db import Base
from sqlalchemy import String
from sqlalchemy.orm import Mapped, mapped_column

class User(Base):
    __tablename__ = 'users'
    id : Mapped[int] = mapped_column(primary_key=True)
    role : Mapped[int] = mapped_column(nullable=False)
    email : Mapped[str] = mapped_column(String(300), nullable=False)
    password : Mapped[str] = mapped_column(String(300), nullable=False)
    permissions : Mapped[str] = mapped_column(String(500), default='final_order,per_fg_bom,per_fg_cost,px_item_cost,final_fg_price')

class FileUploadMetaData(Base):
    __tablename__ = 'file_upload_meta_data'
    id : Mapped[int] = mapped_column(primary_key=True)
    user : Mapped[str] = mapped_column(String(100))
    date_uploaded: Mapped[date] = mapped_column(Date) 
    time_uploaded: Mapped[time] = mapped_column(Time)
    file_uploaded : Mapped[str] = mapped_column(String(300))
    final_procurement_file : Mapped[str] = mapped_column(String(300))


class PriceList(Base):
    """
    Stores the most recent price list uploaded by an admin.
    Every upload truncates the table and inserts fresh rows so the active
    price list is always the latest upload. uploaded_at / uploaded_by
    provide an audit trail.
    """
    __tablename__ = 'price_list'

    id          : Mapped[int]   = mapped_column(primary_key=True)
    item_code   : Mapped[str]   = mapped_column(String(100), nullable=False, index=True)
    price       : Mapped[float] = mapped_column(Float,       nullable=False)
    uploaded_by : Mapped[str]   = mapped_column(String(100), nullable=True)
    uploaded_at : Mapped[datetime] = mapped_column(DateTime, nullable=True)

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