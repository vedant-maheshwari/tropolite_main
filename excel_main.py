from fastapi import FastAPI, HTTPException, Depends, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
import firebase_admin
from firebase_admin import credentials, db
from typing import List

app = FastAPI()
templates = Jinja2Templates(directory="templates")

cred = credentials.Certificate('excel-vault-331bd-firebase-adminsdk-fbsvc-41dd8e1c77.json')
firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://excel-vault-331bd-default-rtdb.asia-southeast1.firebasedatabase.app'
})

class ConversionRow(BaseModel):
    fg_code: str
    px_code: str
    quantity: int  # Changed from 'qty' to 'quantity'

class TablePayload(BaseModel):
    batch_id: str  # Changed from 'int' to 'str'
    items: List[ConversionRow]


@app.get('/', response_class=HTMLResponse)
async def serve_home(request : Request):
    return templates.TemplateResponse(request, 'phone_index.html')

@app.get('/admin', response_class=HTMLResponse)
async def serve_home(request : Request):
    return templates.TemplateResponse(request, 'admin.html')

@app.post('/api/conversion')
async def conversion(payload : TablePayload):

    ref = db.reference('vault_system')

    ref.update({
        'status' : 'pending_approval',
        'request_data' : payload.model_dump(),
        'final_result' : None
    })

@app.post('/approve')
async def approve_request():
    ref = db.reference('vault_system')

    ref.update({
        'status' : 'approved'
    })