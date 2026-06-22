# pyrefly: ignore [missing-import]
from fastapi import FastAPI, HTTPException, Depends, status, Request, File, UploadFile
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from db import Base, SessionLocal, get_db, engine, get_db_admin, engine_admin
from db import get_db_md
from sqlalchemy.orm import Session
from sqlalchemy import text
import models, schemas, services
from typing import List
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse
from fastapi.security import OAuth2PasswordRequestForm
import auth
import os
import shutil
from fastapi.middleware.cors import CORSMiddleware
import time
from datetime import date, datetime

app = FastAPI()
templates = Jinja2Templates(directory='templates')
app.mount('/static', StaticFiles(directory='static'), name='static')


class ApproveRequest(BaseModel):
    excel_password: str


Base.metadata.create_all(bind=engine_admin)

app.add_middleware(
    CORSMiddleware,
    allow_origins=['*'],
    allow_credentials=True,
    allow_methods=['*'],
    allow_headers=['*']
)


@app.get('/', response_class=HTMLResponse)
async def home(request : Request):
    return templates.TemplateResponse(request, 'index.html')

@app.post('/register_user')
def create_user(user_info : schemas.register_user, current_user : str = Depends(auth.get_current_admin_for_working),
                db : Session = Depends(get_db_admin)):
    
    if db.query(models.User).filter(models.User.email == user_info.email).first():
        raise HTTPException(400, 'user already exists')
    
    hashed_password = auth.hash_password(user_info.password)
    user_info.password = hashed_password

    response = services.register_user(user_info, db)
    if response == True:
        return {'user created successfully'}
    else:
        raise HTTPException(401, 'issue creating user')

@app.get('/users')
def list_users(current_user: str = Depends(auth.get_current_admin_for_working), db: Session = Depends(get_db_admin)):
    users = db.query(models.User).order_by(models.User.id.asc()).all()
    return [{"id": u.id, "email": u.email, "role": u.role} for u in users]

@app.put('/users/{user_id}')
def update_user(user_id: int, user_info: schemas.user_update, current_user: str = Depends(auth.get_current_admin_for_working), db: Session = Depends(get_db_admin)):
    user = db.query(models.User).filter(models.User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="User not found")
    
    # Check if we are changing email and it already exists on another user
    existing = db.query(models.User).filter(models.User.email == user_info.email, models.User.id != user_id).first()
    if existing:
        raise HTTPException(status_code=400, detail="Email already in use")
        
    user.email = user_info.email
    user.role = user_info.role
    if user_info.password:
        user.password = auth.hash_password(user_info.password)
        
    db.commit()
    return {"message": "User updated successfully"}

@app.delete('/users/{user_id}')
def delete_user(user_id: int, current_user: str = Depends(auth.get_current_admin_for_working), db: Session = Depends(get_db_admin)):
    user = db.query(models.User).filter(models.User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="User not found")
    
    db.delete(user)
    db.commit()
    return {"message": "User deleted successfully"}

@app.get('/metadata')
def get_metadata(current_user: str = Depends(auth.get_current_admin_for_working), db: Session = Depends(get_db_admin)):
    meta = db.query(models.FileUploadMetaData).order_by(models.FileUploadMetaData.id.desc()).all()
    return [
        {
            "id": m.id,
            "user": m.user,
            "date_uploaded": m.date_uploaded.isoformat() if m.date_uploaded else None,
            "time_uploaded": m.time_uploaded.isoformat() if m.time_uploaded else None,
            "file_uploaded": m.file_uploaded,
            "final_procurement_file": m.final_procurement_file
        } for m in meta
    ]

# @app.post('/token')
# def login(form_data: OAuth2PasswordRequestForm = Depends()):
#     user = auth.user_db.get(form_data.username)
#     if not user:
#         raise HTTPException(400, 'user not found')
#     if not auth.pwd_context.verify(form_data.password, user['hashed_password']):
#         raise HTTPException(400, 'invalid email or password')
#     access_token = auth.create_access_token(data={
#         'sub': user['username']
#     })
#     return {
#         'access_token': access_token,
#         'token_type': 'bearer',
#         'role': user.get('role', 'user')
#     }

@app.post('/token')
def login(form_data: OAuth2PasswordRequestForm = Depends(), db : Session = Depends(get_db_admin)):
    user = db.query(models.User).filter(models.User.email == form_data.username).first()
    if not user:
        raise HTTPException(400, 'user not found')
    if not auth.pwd_context.verify(form_data.password, user.password):
        raise HTTPException(400, 'invalid email or password')
    access_token = auth.create_access_token(data={
        'sub': user.email, 'id' : user.id, 'role' : user.role
    })
    return {
        'access_token': access_token,
        'token_type': 'bearer',
        'role': user.role
    }

@app.get('/me')
def get_me(current_user: str = Depends(auth.get_current_user)):
    """Returns the logged-in user's username and role."""
    user = auth.user_db.get(current_user)
    return {
        'username': current_user,
        'role': user.get('role', 'user') if user else 'user'
    }

@app.post('/upload_excel')
async def upload_excel(
    file: UploadFile = File(...),
    current_user=Depends(auth.get_current_user),
    db: Session = Depends(get_db),
    db_md : Session = Depends(get_db_md),
    db_admin : Session = Depends(get_db_admin)
):
    if not file.filename.endswith(('.xlsx', '.xls', '.csv')):
        raise HTTPException(400, detail='Invalid file type.')

    save_dir = '/tmp/uploaded_files'
    os.makedirs(save_dir, exist_ok=True)
    ext = os.path.splitext(file.filename)[1]
    file_location = f'{save_dir}/{datetime.now()}{ext}'

    with open(file_location, "wb+") as f:
        shutil.copyfileobj(file.file, f)

    check = services.check_file_structure_and_conditions(file_location)
    if check['status'] != 'True':
        raise HTTPException(status_code=400, detail=check['details'])

    try:
        json_payload = services.process_file(file_location)
        results = services.run_query(json_payload, db)

        px_items = services.extract_px_items(results)
        print('Extracted PX items')
        # print(px_items)

        print('making PX to MD conversion dict')
        conversion_dict = services.px_to_md_conversion_lookup(db_md)
        # print(conversion_dict)

        print('converting PX to MD')
        md_items = services.convert_px_to_md(px_items, conversion_dict)

        # print(md_items)
        
        print('sending to firebase')
        px_payload = services.Payload(batch_id = services.generate_production_batch(),items = md_items)
        services.conversion(px_payload)

        print('px_code sent')
        print(current_user)

        metadata = models.FileUploadMetaData(
            user = current_user,
            date_uploaded = date.today(),
            time_uploaded = datetime.now().time().isoformat(timespec='seconds'),
            file_uploaded = file_location,
            final_procurement_file = file_location
        )

        db_admin.add(metadata)
        db_admin.commit()
        db_admin.refresh(metadata)
        
        return results
    
    except Exception as e:
        raise HTTPException(400, detail=f'Error occurred: {e}')


@app.post('/get_bom_only')
async def get_bom_only(
    file: UploadFile = File(...),
    current_user=Depends(auth.get_current_user),
    db: Session = Depends(get_db),
    db_admin: Session = Depends(get_db_admin)
):
    """Runs only the BOM query (no PX conversion). Returns the per-FG BOM table directly."""
    if not file.filename.endswith(('.xlsx', '.xls', '.csv')):
        raise HTTPException(400, detail='Invalid file type.')

    save_dir = '/tmp/uploaded_files'
    os.makedirs(save_dir, exist_ok=True)
    ext = os.path.splitext(file.filename)[1]
    file_location = f'{save_dir}/{datetime.now()}{ext}'

    with open(file_location, "wb+") as f:
        shutil.copyfileobj(file.file, f)

    check = services.check_file_structure_and_conditions(file_location)
    if check['status'] != 'True':
        raise HTTPException(status_code=400, detail=check['details'])

    try:
        json_payload = services.process_file(file_location)
        results = services.run_query(json_payload, db)
        return results
    except Exception as e:
        raise HTTPException(400, detail=f'Error occurred: {e}')


class RFItem(BaseModel):
    code: str
    qty: float

@app.post('/explode_rf')
async def explode_rf(
    rf_items: List[RFItem],
    current_user=Depends(auth.get_current_user),
    db: Session = Depends(get_db)
):
    """
    Accepts a list of RF codes with quantities (from PX conversion results).
    Runs BOM explosion on each and returns the broken-down raw materials.
    """
    try:
        items = [{'code': item.code, 'qty': item.qty} for item in rf_items]
        results = services.run_rf_explosion(items, db)
        return results
    except Exception as e:
        raise HTTPException(400, detail=f'Error occurred during RF explosion: {e}')


@app.get('/admin', response_class=HTMLResponse)
async def serve_home(request : Request):
    return templates.TemplateResponse(request, 'admin.html')

# @app.post('/api/conversion')
# async def conversion(payload : TablePayload):

#     ref = db.reference('vault_system')

#     ref.update({
#         'status' : 'pending_approval',
#         'request_data' : payload.model_dump(),
#         'final_result' : None
#     })

@app.post('/approve')
async def approve_request(
    body: ApproveRequest,
    current_admin: str = Depends(auth.get_current_admin)
):
    """Admin-only: sets Firebase status to approved and stores the Excel password for the watchdog."""
    services.approve(body.excel_password)
    return {'status': 'approved'}

@app.post('/cleanup')
async def cleanup_firebase(current_user: str = Depends(auth.get_current_user)):
    """Wipes all vault_system data from Firebase after the frontend has read the result.
    Requires any valid logged-in user (not just admin)."""
    services.cleanup_firebase()
    return {'status': 'cleared'}

@app.get('/health')
def health():
    return {'healthy'}

@app.get('/health/db')
def check_db_conn():
    try :
        with engine.connect() as conn:
            conn.execute(text('SELECT 1'))
            return {'status' : 'connected'}
    except Exception as e:
        return {'error encountered' : f'{e}'}


@app.post('/get_item_by_id')
def get_item_by_id(db: Session = Depends(get_db)):
    return services.get_closing_stock(db)

@app.get('/closing_stock')
def closing_stock(
    current_user: str = Depends(auth.get_current_user),
    db: Session = Depends(get_db)
):
    """Returns a dict of {ItemCode: {qty, last_pur_prc}} using today's date across all configured warehouses."""
    return services.get_closing_stock(db)


@app.post('/upload_price_excel')
async def upload_price_excel(
    file: UploadFile = File(...),
    current_user: str = Depends(auth.get_current_admin_for_working)
):
    """
    Admin-only. Upload a monthly price Excel with columns [ItemCode, Price].
    The prices are stored in-memory and used in the Stock View pricing columns.
    """
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, detail='Invalid file type. Only .xlsx / .xls accepted.')

    save_dir = '/tmp/uploaded_files'
    os.makedirs(save_dir, exist_ok=True)
    ext = os.path.splitext(file.filename)[1]
    file_location = f'{save_dir}/price_list_{datetime.now().strftime("%Y%m%d_%H%M%S")}{ext}'

    with open(file_location, 'wb+') as f:
        shutil.copyfileobj(file.file, f)

    try:
        price_dict = services.load_new_prices_from_excel(file_location)
        return {'status': 'ok', 'items_loaded': len(price_dict)}
    except Exception as e:
        raise HTTPException(400, detail=f'Error loading price Excel: {e}')


@app.get('/new_prices')
def get_new_prices(
    current_user: str = Depends(auth.get_current_user)
):
    """Returns the currently loaded monthly new-price dict: {ItemCode: price}."""
    return services.get_new_prices()


# @app.post('/run_MD_query')
# def px_to_md_conversion_lookup(db : Session = Depends(get_db_md)):
#     query = text(""" 
#     ;WITH FormulaExplosion AS
# (
#     -----------------------------------------------------------------
#     -- Direct Formula Components
#     -----------------------------------------------------------------
#     SELECT
#         BH.U_ITNO              AS PX_Code,
#         F1.U_ITEMCODE          AS MD_Code,
#         F1.U_QTY               AS Formula_Qty,
#         1                      AS Level
#     FROM [@C_BOMH] BH
#     INNER JOIN [@OFES] FH
#         ON BH.U_FORMULA = FH.U_FCODE
#     INNER JOIN [@FES1] F1
#         ON FH.DocEntry = F1.DocEntry
#     WHERE BH.U_STATUS = 'Active'
#       AND FH.U_STATUS = 'Active'

#     UNION ALL

#     -----------------------------------------------------------------
#     -- Recursive Formula Expansion
#     -----------------------------------------------------------------
#     SELECT
#         FE.PX_Code,
#         F1.U_ITEMCODE,
#         F1.U_QTY,
#         FE.Level + 1
#     FROM FormulaExplosion FE
#     INNER JOIN [@C_BOMH] BH
#         ON FE.MD_Code = BH.U_ITNO
#     INNER JOIN [@OFES] FH
#         ON BH.U_FORMULA = FH.U_FCODE
#     INNER JOIN [@FES1] F1
#         ON FH.DocEntry = F1.DocEntry
#     WHERE BH.U_STATUS = 'Active'
#       AND FH.U_STATUS = 'Active'
# )
# SELECT
#     PX_Code,
#     MD_Code,
#     Formula_Qty,
#     ROUND(
#         Formula_Qty * 100.0 /
#         SUM(Formula_Qty) OVER (PARTITION BY PX_Code),
#         4
#     ) AS Formula_Percent,
#     Level
# FROM FormulaExplosion
# ORDER BY
#     PX_Code,
#     MD_Code
# OPTION (MAXRECURSION 1000);
#     """)

#     results = db.execute(query)
#     records = results.mappings().all()
#     return records

@app.post('/run_query')
def px_to_md_conversion_lookup(db : Session = Depends(get_db_md)):
    query = text(""" 
    select top 2 * from OITM;
    """)

    results = db.execute(query)
    records = results.mappings().all()
    return records