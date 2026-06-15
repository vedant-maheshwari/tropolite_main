from collections import defaultdict
from datetime import datetime
import json
import os
import random
import sqlite3
import string
from typing import List

from pydantic import BaseModel
from sqlalchemy import text
from sqlalchemy.orm import Session
import schemas, models
import pandas as pd

import firebase_admin
from firebase_admin import credentials, db

_firebase_creds_env = os.getenv("FIREBASE_CREDENTIALS")
if _firebase_creds_env:
    # Vercel/production: credentials stored as a JSON string in env var
    _cred_dict = json.loads(_firebase_creds_env)
    cred = credentials.Certificate(_cred_dict)
else:
    # Local dev: load from file
    cred = credentials.Certificate('excel-vault-331bd-firebase-adminsdk-fbsvc-41dd8e1c77.json')

firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://excel-vault-331bd-default-rtdb.asia-southeast1.firebasedatabase.app'
})


class ConversionRow(BaseModel):
    MD_code: str
    Qty: float  

class Payload(BaseModel):
    batch_id: str  
    items: List[ConversionRow]

def select_by_id(id, db : Session):
    output = db.query(models.formula).filter(models.formula.product_id == id).all()
    result = {}
    for item in output:
        result[item.material_id] = item.material_req
    return result

def get_multiple_item(item_ids : schemas.multi_id_input, db : Session):
    final_result = {}
    for item_id in item_ids.ids:
        output = select_by_id(item_id, db)
        final_result[item_id] = output
    return final_result

def get_total_materials(item_ids : schemas.multi_id_input, db : Session):
    material_dict = {}

    for item_id in item_ids.ids:
        product = db.query(models.formula).filter(models.formula.product_id == item_id).all()
        for item in product:
            if item.material_id not in material_dict:
                material_dict[item.material_id] = item.material_req
            else:
                material_dict[item.material_id] += item.material_req
        
    return material_dict

def get_total_materials_with_op_balance_forecast_qty(payload : schemas.MultiIdInputWithOpBalanceForecastQty, db : Session):
    # print(payload.root)
    material_dict = {}

    for item in payload.root:
        for item_code, item_qty in item.items():
            # print(item_code, item_qty[0]-item_qty[1])
            qty = item_qty[0]-item_qty[1]
            product = db.query(models.formula).filter(models.formula.product_id == item_code).all()
            for item in product:
                if item.material_id not in material_dict:
                    material_dict[item.material_id] = (item.material_req/item.material_qty) * qty
                else:
                    material_dict[item.material_id] += (item.material_req/item.material_qty) * qty
    return material_dict

# ---------------------------------- main functions --------------------------------------------

def register_user(user_info : schemas.register_user, db : Session):
    user = models.User(
        email = user_info.email,
        role = user_info.role,
        password = user_info.password
    )
    
    db.add(user)
    db.commit()
    db.refresh(user)
    return True

def fg_conditions(fg_list : list, fg_qty : list):
    length = len(fg_list)
    if length < 1:
        return False
    # if length != len(set(fg_list)):
    #     return False
    # if 0 in fg_qty:
    #     return False
    return True

def check_file_structure_and_conditions(filepath):
    df = pd.read_excel(filepath, engine='openpyxl', skiprows=1)
    column_list = df.columns
    print(column_list)
    if not (len(column_list) == 2 and 'FG_CODE' in column_list and 'QTY' in column_list):
        return {'status' : 'error', 'details' : 'File structure not correct'}
    if not fg_conditions(df['FG_CODE'], df['QTY']):
        return {'status' : 'error', 'details' : 'not adhering to FG conditions'}
    return {'status' : 'True', 'details' : 'everything is alright'}
    
# def run_query(json_payload: str, db: Session):
#     import json as json_lib

#     # Parse the JSON back to build VALUES rows directly
#     items = json_lib.loads(json_payload)

#     # Build: ('FG001', 100), ('FG002', 200), ...
#     values_rows = ",\n".join(
#         f"(N'{row['FG_Code'].replace(chr(39), chr(39)*2)}', {row['FG_Qty']})"
#         for row in items
#     )

#     raw_conn = db.connection().connection
#     cursor = raw_conn.cursor()

#     sql = f"""
#         SET NOCOUNT ON;

#         DECLARE @FGQty TABLE
#         (
#             FG_Code NVARCHAR(50),
#             FG_Qty  NUMERIC(19,4)
#         )

#         INSERT INTO @FGQty (FG_Code, FG_Qty)
#         VALUES
#         {values_rows}

#         ;WITH FG_FORMULA AS
#         (
#             SELECT
#                 I.FG_Code,
#                 FG.ItemName AS FG_Name,
#                 I.FG_Qty,
#                 BH.U_ITNO AS Parent_Code,
#                 F1.U_ITEMCODE AS Item_Code,
#                 IT.ItemName AS Item_Name,
#                 GRP.ItmsGrpNam AS Item_Group,
#                 'Formula Item' AS Item_Type,
#                 CAST(
#                     (F1.U_QTY / NULLIF(OH.U_TLT, 0)) * I.FG_Qty
#                     AS NUMERIC(19,8)
#                 ) AS Item_Projection
#             FROM @FGQty I
#             INNER JOIN [@C_BOMH] BH ON I.FG_Code = BH.U_ITNO AND BH.U_STATUS = 'Active'
#             INNER JOIN [@OFES]   OH ON BH.U_FORMULA = OH.U_FCODE AND OH.U_STATUS = 'Active'
#             INNER JOIN [@FES1]   F1 ON OH.DocEntry = F1.DocEntry
#             INNER JOIN OITM      FG ON BH.U_ITNO = FG.ItemCode
#             INNER JOIN OITM      IT ON F1.U_ITEMCODE = IT.ItemCode
#             LEFT  JOIN OITB     GRP ON IT.ItmsGrpCod = GRP.ItmsGrpCod
#         ),

#         RF_EXPLODE AS
#         (
#             SELECT
#                 P.FG_Code,
#                 P.FG_Name,
#                 P.FG_Qty,
#                 BH.U_ITNO AS Parent_Code,
#                 F1.U_ITEMCODE AS Item_Code,
#                 IT.ItemName AS Item_Name,
#                 GRP.ItmsGrpNam AS Item_Group,
#                 'Formula Item' AS Item_Type,
#                 CAST(
#                     (F1.U_QTY / NULLIF(OH.U_TLT, 0)) * P.Item_Projection
#                     AS NUMERIC(19,8)
#                 ) AS Item_Projection
#             FROM FG_FORMULA P
#             INNER JOIN [@C_BOMH] BH ON P.Item_Code = BH.U_ITNO AND BH.U_STATUS = 'Active'
#             INNER JOIN [@OFES]   OH ON BH.U_FORMULA = OH.U_FCODE AND OH.U_STATUS = 'Active'
#             INNER JOIN [@FES1]   F1 ON OH.DocEntry = F1.DocEntry
#             INNER JOIN OITM      IT ON F1.U_ITEMCODE = IT.ItemCode
#             LEFT  JOIN OITB     GRP ON IT.ItmsGrpCod = GRP.ItmsGrpCod
#             WHERE P.Item_Code LIKE 'RF%'
#         ),

#         PK AS
#         (
#             SELECT
#                 I.FG_Code,
#                 FG.ItemName AS FG_Name,
#                 I.FG_Qty,
#                 BH.U_ITNO AS Parent_Code,
#                 BD.U_ITNO AS Item_Code,
#                 IT.ItemName AS Item_Name,
#                 GRP.ItmsGrpNam AS Item_Group,
#                 'BOM Item' AS Item_Type,
#                 CAST(BD.U_QTY * I.FG_Qty AS NUMERIC(19,8)) AS Item_Projection
#             FROM @FGQty I
#             INNER JOIN [@C_BOMH] BH ON I.FG_Code = BH.U_ITNO AND BH.U_STATUS = 'Active'
#             INNER JOIN [@C_BOMD] BD ON BH.DocEntry = BD.DocEntry
#             INNER JOIN OITM      FG ON BH.U_ITNO = FG.ItemCode
#             INNER JOIN OITM      IT ON BD.U_ITNO = IT.ItemCode
#             LEFT  JOIN OITB     GRP ON IT.ItmsGrpCod = GRP.ItmsGrpCod
#             WHERE BD.U_ITNO LIKE 'PK%'
#         )

#         SELECT
#             A.Item_Group      AS [Item Group],
#             A.FG_Code         AS [Finish Good],
#             A.FG_Qty          AS [FG Qty],
#             A.Parent_Code,
#             A.FG_Name         AS [Finish Good Name],
#             A.Item_Code       AS [Item Code],
#             A.Item_Type,
#             CAST(A.Item_Projection AS NUMERIC(19,4)) AS [Required Qty]
#         FROM
#         (
#             SELECT * FROM FG_FORMULA WHERE Item_Code NOT LIKE 'RF%'
#             UNION ALL
#             SELECT * FROM RF_EXPLODE
#             UNION ALL
#             SELECT * FROM PK
#         ) A
#         WHERE
#             A.Item_Code LIKE 'RM%'
#             OR A.Item_Code LIKE 'PX%'
#             OR A.Item_Code LIKE 'PK%'
#             OR A.Item_Code LIKE 'R&d%'
#             OR A.Item_Code LIKE 'R&D%'
#             OR A.Item_Code LIKE 'LB%'
#         ORDER BY
#             A.FG_Code,
#             A.Parent_Code,
#             A.Item_Code
#     """

#     cursor.execute(sql)
#     columns = [col[0] for col in cursor.description]
#     rows = cursor.fetchall()
#     cursor.close()

#     return [dict(zip(columns, row)) for row in rows]

## ----------------------------------------------- vikas query -------------------------------------------- 

# def run_query(json_payload: str, db: Session):
#     import json as json_lib

#     # Parse the JSON back to build VALUES rows directly
#     items = json_lib.loads(json_payload)

#     # Build: ('FG001', 100), ('FG002', 200), ...
#     values_rows = ",\n".join(
#         f"(N'{row['FG_Code'].replace(chr(39), chr(39)*2)}', {row['FG_Qty']})"
#         for row in items
#     )

#     raw_conn = db.connection().connection
#     cursor = raw_conn.cursor()

#     sql = f"""
#         SET NOCOUNT ON;

#         IF OBJECT_ID('tempdb..#InputData') IS NOT NULL DROP TABLE #InputData;
#         IF OBJECT_ID('tempdb..#final') IS NOT NULL DROP TABLE #final;
#         IF OBJECT_ID('tempdb..#dd') IS NOT NULL DROP TABLE #dd;

#         CREATE TABLE #InputData (
#             U_ItemCode NVARCHAR(50) COLLATE DATABASE_DEFAULT,
#             U_Quantity NUMERIC(19,6)
#         );

#         INSERT INTO #InputData (U_ItemCode, U_Quantity)
#         VALUES
#         {values_rows};

#         With RPL (
#         U_FGCODE, U_FORMULAID, U_ITEMCODE, U_QTYINDISPLAYUOM, U_FGWHSE,
#         Level, Formula_Weight, Parent_Code, Base_Requirement,
#         Finish_Goods_Help, Component_Item_Help, Item_Group,
#         Inventory_UOM, Parent_Inventory_UOM, Item_Type, U_RevisionNO
#         ) As
#         (
#             SELECT
#             T0.U_ITNO AS 'Finish Good', T0.U_FORMULA As 'Formula ID', T1.U_ITNO As 'Item Code',
#             T1.U_QTY As 'Base Quantity', T0.U_WHS As 'FG Warehouse', 1 as 'Level', 0 as 'Total Formula Weight',
#             T0.U_ITNO As 'Parent_Code', T1.U_QTY*1 as 'Base_Req',
#             Cast(Concat(T0.U_ITNO,'-',T0.U_ITNO) as nvarchar(MAX)) as 'Finish_Goods_Help',
#             Cast(Concat(T0.U_ITNO,'-',T1.U_ITNO) as nvarchar(MAX)) as 'Component_Item_Help',
#             T3.ItmsGrpNam as 'Item Group', T2.InvntryUom as 'Inventory UOM',
#             T4.InvntryUom as 'Parent Inventory UOM', 'BOM Item' As 'Item_Type', T0.U_REVISION

#             FROM [@C_BOMH] T0
#             inner JOIN [@C_BOMD] T1 ON T0.DocEntry = T1.DocEntry
#             LEFT OUTER JOIN OITM T2 ON T1.U_ITNO = T2.ItemCode
#             LEFT OUTER JOIN OITB T3 ON T2.ItmsGrpCod = T3.ItmsGrpCod
#             LEFT OUTER JOIN OITM T4 ON T0.U_ITNO = T4.ItemCode
#             where T0.U_STATUS ='Active'

#             Union All

#             SELECT
#             T0.U_ITNO AS 'Finish Good', T0.U_FORMULA As 'Formula ID', T3.U_ITEMCODE As 'Item Code',
#             T3.U_QTY As 'Base Quantity', T0.U_WHS As 'FG Warehouse', 1 as 'Level',
#             T2.U_TLT as 'Total Formula Weight', T0.U_ITNO As 'Parent_Code',
#             T3.U_QTY/nullif(T2.U_TLT,0) as 'Base_Req',
#             Cast(Concat(T0.U_ITNO,'-',T0.U_ITNO) as nvarchar(MAX)) as 'Finish_Goods_Help',
#             Cast(Concat(T0.U_ITNO,'-',T3.U_ITEMCODE) as nvarchar(MAX)) as 'Component_Item_Help',
#             T5.ItmsGrpNam as 'Item Group', T4.InvntryUom as 'Inventory UOM',
#             T6.InvntryUom as 'Parent Inventory UOM', 'Formula Item' As 'Item_Type', T0.U_REVISION

#             FROM dbo.[@C_BOMH] T0
#             LEFT OUTER JOIN dbo.[@OFES] T2 ON T0.U_FORMULA = T2.U_FCODE
#             LEFT OUTER JOIN dbo.[@FES1] T3 ON T3.DocEntry = T2.DocEntry
#             LEFT OUTER JOIN OITM T4 ON T3.U_ITEMCODE = T4.ItemCode
#             LEFT OUTER JOIN OITB T5 ON T4.ItmsGrpCod = T5.ItmsGrpCod
#             LEFT OUTER JOIN OITM T6 ON T0.U_ITNO = T6.ItemCode
#             Where T2.U_STATUS ='Active' and T0.U_STATUS='Active'

#             Union All

#             SELECT
#             T0.U_FGCODE AS 'Finish Good', T1.U_FORMULA As 'Formula ID', T2.U_ITNO As 'Item Code',
#             T2.U_QTY As 'Base Quantity', T1.U_WHS As 'FG Warehouse', LEVEL + 1,
#             0 as 'Total Formula Weight', T1.U_ITNO as 'Parent_Code',
#             T2.U_QTY*1 as 'Base_Req',
#             Cast(Concat(T0.U_FGCODE,'-',T1.U_ITNO) as nvarchar(MAX)) as 'Finish_Goods_Help',
#             Cast(Concat(T1.U_ITNO,'-',T2.U_ITNO) as nvarchar(MAX)) as 'Component_Item_Help',
#             T5.ItmsGrpNam as 'Item Group', T4.InvntryUom as 'Inventory UOM',
#             T6.InvntryUom as 'Parent Inventory UOM', 'BOM Item' As 'Item_Type', T1.U_REVISION

#             FROM RPL T0, dbo.[@C_BOMH] T1
#             INNER JOIN dbo.[@C_BOMD] T2 ON T1.DocEntry = T2.DocEntry
#             INNER JOIN OITM T4 ON T2.U_ITNO = T4.ItemCode
#             INNER JOIN OITB T5 ON T4.ItmsGrpCod = T5.ItmsGrpCod
#             INNER JOIN OITM T6 ON T1.U_ITNO = T6.ItemCode
#             Where T0.U_ITEMCODE = T1.U_ITNO and Level < 4

#             Union All

#             Select
#             T0.U_FGCODE AS 'Finish Good', T1.U_FORMULA As 'Formula ID', T3.U_ITEMCODE As 'Item Code',
#             T3.U_QTY As 'Base Quantity', T1.U_WHS As 'FG Warehouse', LEVEL + 1,
#             T2.U_TLT as 'Total Formula Weight', T1.U_ITNO as 'Parent_Code',
#             T3.U_QTY/nullif(T2.U_TLT,0) as 'Base_Req',
#             Cast(Concat(T0.U_FGCODE,'-',T1.U_ITNO) as nvarchar(MAX)) as 'Finish_Goods_Help',
#             Cast(Concat(T0.U_FGCODE,'-',T3.U_ITEMCODE) as nvarchar(MAX)) as 'Component_Item_Help',
#             T5.ItmsGrpNam as 'Item Group', T4.InvntryUom as 'Inventory UOM',
#             T6.InvntryUom as 'Parent Inventory UOM', 'Formula Item' As 'Item_Type', T1.U_REVISION

#             FROM RPL T0
#             INNER JOIN [@C_BOMH] T1 ON T0.U_ITEMCODE = T1.U_ITNO
#             INNER JOIN [@OFES] T2 ON T1.U_FORMULA = T2.U_FCODE
#             INNER JOIN [@FES1] T3 ON T2.DocEntry = T3.DocEntry
#             INNER JOIN OITM T4 ON T3.U_ITEMCODE = T4.ItemCode
#             INNER JOIN OITB T5 ON T4.ItmsGrpCod = T5.ItmsGrpCod
#             INNER JOIN OITM T6 ON T1.U_ITNO = T6.ItemCode
#             WHERE T2.U_STATUS='Active'
#             AND T0.Level < 4
#             AND ISNULL(T0.U_ITEMCODE,'') <> 'RM2159'
#             AND T3.U_ITEMCODE <> T0.U_ITEMCODE
#             AND T1.U_STATUS = T2.U_STATUS
#         )

#         Select distinct U_FGCODE As 'Finish Good', U_FORMULAID As 'Formula ID',
#         C.U_ITEMCODE As 'Item Code',
#         U_QTYINDISPLAYUOM As 'Base Quantity', U_FGWHSE As 'FG Warehouse',
#         Level, Formula_Weight as 'Total Formula Weight',
#         Parent_Code as 'Parent_Code',
#         Base_Requirement as 'Base_Requirement',
#         Row_Number() Over (Partition by U_FGCODE Order By U_FGCODE) As 'Count',
#         Finish_Goods_Help As 'Finish_Goods_Help',
#         Component_Item_Help AS 'Component_Item_Help',
#         Item_Group As 'Item Group',
#         Inventory_UOM as 'Inventory UOM',
#         Parent_Inventory_UOM as 'Parent Inventory UOM',
#         Item_Type,
#         convert(numeric(19,2), E.U_Quantity) 'Projection Qty',
#         convert(numeric(19,2), 0.00) 'Item Projection',
#         convert(numeric(19,2), 0.00) 'Convertion Factor',
#         convert(numeric(19,2), 0.00) 'Item Projection in KG',
#         convert(nvarchar(100),'') 'Projection_UOM'
#         into #final
#         From RPL C, #InputData E
#         WHERE C.U_FGCODE = E.U_ItemCode COLLATE DATABASE_DEFAULT
#         AND E.U_Quantity > 0.00
#         Group By U_FGCODE, U_FORMULAID, C.U_ITEMCODE, U_QTYINDISPLAYUOM,
#         U_FGWHSE, Level, Formula_Weight, Parent_Code, Base_Requirement,
#         Finish_Goods_Help, Component_Item_Help, Item_Group,
#         Inventory_UOM, Parent_Inventory_UOM, Item_Type, E.U_Quantity
#         option (maxrecursion 0);

#         update t
#         set [Convertion Factor] = isnull(cd1.U_CONFACTR,1)
#         from #final t,
#         (select distinct cd.U_CONFACTR, ch.U_ITNO
#          from [@C_ITMSTR] ch, [@C_ITMUOM] cd
#          where cd.DocEntry = ch.DocEntry and cd.U_TOUOM <> '') cd1
#         where t.Parent_Code = cd1.U_ITNO COLLATE DATABASE_DEFAULT;

#         update t
#         set [Convertion Factor] = 1
#         from #final t where [Convertion Factor] = 0.00;

#         update t
#         set [Projection Qty] = c.U_Quantity,
#             Projection_UOM = 'KG'
#         from #final t, #InputData c
#         where t.[Finish Good] = c.U_ItemCode COLLATE DATABASE_DEFAULT
#         and c.U_Quantity <> 0.00;

#         select [Finish Good], [Item Code], Parent_Code,
#         (case when [Projection Qty] = 0.00 then [Projection Qty]
#               else Base_Requirement end) Base_Requirement,
#         (Base_Requirement * [Projection Qty] * [Convertion Factor]) ff
#         into #dd
#         from #final
#         where [Item Code] in (select [Item Code] from #final where Item_Type <> 'BOM Item');

#         update t
#         set [Item Projection] =
#             case when Item_Type = 'BOM Item'
#                  then (case when [Inventory UOM] = 'NO'
#                             then Base_Requirement * [Projection Qty] * [Convertion Factor]
#                             else Base_Requirement * [Projection Qty]
#                        end)
#                  else Base_Requirement * [Projection Qty] * [Convertion Factor]
#             end
#         from #final t;

#         update t
#         set [Projection Qty] = T1.[Item Projection]
#         from #final t, #final t1
#         where t.Parent_Code = T1.[Item Code] COLLATE DATABASE_DEFAULT
#         AND T.Level <> 1 AND T1.Level = 1
#         and t.[Finish Good] = t1.Parent_Code COLLATE DATABASE_DEFAULT
#         AND T1.[Item Projection] <> 0;

#         update t
#         set [Item Projection] = Base_Requirement * [Projection Qty] * [Convertion Factor]
#         from #final t
#         where Item_Type = 'Formula Item' and Level > 1;

#         update t
#         set [Item Projection] = Base_Requirement * [Projection Qty] * [Convertion Factor]
#         from #final t
#         where Item_Type = 'BOM Item' and Level > 1;

#         update t
#         set [Item Projection in KG] =
#             case when item_type = 'BOM'
#                  then (case when Projection_UOM = 'NO'
#                             then Base_Requirement * [Convertion Factor] * [Projection Qty]
#                             else Base_Requirement * [Projection Qty]
#                        end)
#                  else Base_Requirement * [Projection Qty] * [Convertion Factor]
#             end
#         from #final t;

#         select [Item Group], [Finish Good], Parent_Code,
#         (select ItemName from oitm where ItemCode = [Finish Good]) 'Finish Good Name',
#         [Item Code], Item_Type, [Item Projection],
#         [Inventory UOM] as Inventory_UOM
#         from #final
#         Order By [Finish Good], Level, [Item Code], [Base Quantity];

#         IF OBJECT_ID('tempdb..#dd') IS NOT NULL DROP TABLE #dd;
#         IF OBJECT_ID('tempdb..#final') IS NOT NULL DROP TABLE #final;
#         IF OBJECT_ID('tempdb..#InputData') IS NOT NULL DROP TABLE #InputData;
#     """

#     cursor.execute(sql)
#     columns = [col[0] for col in cursor.description]
#     rows = cursor.fetchall()
#     cursor.close()

#     return [dict(zip(columns, row)) for row in rows]

# ---------------------------------------------------------------------------------------

# ----------------------------------------------- vedant query -----------------------------------------------

def run_query(json_payload: str, db: Session) -> list[dict]:
    import json as json_lib

    # Parse the JSON back to build VALUES rows directly
    items = json_lib.loads(json_payload)

    # Build: ('FG001', 100), ('FG002', 200), ...
    values_rows = ",\n".join(
        f"(N'{row['FG_Code'].replace(chr(39), chr(39)*2)}', {row['FG_Qty']})"
        for row in items
    )

    raw_conn = db.connection().connection
    cursor = raw_conn.cursor()

    sql = f"""
        SET NOCOUNT ON;

        IF OBJECT_ID('tempdb..#BomResult') IS NOT NULL DROP TABLE #BomResult;

        DECLARE @FGQty TABLE
        (
            FG_Code NVARCHAR(50),
            FG_Qty  NUMERIC(19,4)
        );

        INSERT INTO @FGQty (FG_Code, FG_Qty)
        VALUES
        {values_rows};

        ;WITH BOM_Explosion AS (

            -- ================================================================
            -- ANCHOR LEVEL 0: The FG itself, seeded with requested qty
            -- ================================================================
            SELECT
                CAST(h.U_ITNO     AS NVARCHAR(100))  AS fg_item,
                CAST(h.U_ITNO     AS NVARCHAR(100))  AS parent_item,
                CAST(h.U_ITNO     AS NVARCHAR(100))  AS child_item,
                CAST(itm.ItemName AS NVARCHAR(200))  AS child_name,
                CAST(q.FG_Qty     AS DECIMAL(28,10)) AS qty_per_fg,
                CAST(0            AS INT)            AS bom_level,
                CAST(h.U_UOM      AS NVARCHAR(50))   AS uom,
                CAST('BOM'        AS NVARCHAR(10))   AS source,
                CAST(h.U_ITNO     AS NVARCHAR(4000)) AS explosion_path
            FROM [@C_BOMH] h
            INNER JOIN OITM itm ON itm.ItemCode = h.U_ITNO
            INNER JOIN @FGQty q ON q.FG_Code = h.U_ITNO
            WHERE h.U_STATUS = 'Active'

            UNION ALL

            -- ================================================================
            -- RECURSE BRANCH 1: BOM detail children (PK, LB, RF, etc.)
            -- ================================================================
            SELECT
                CAST(bx.fg_item                             AS NVARCHAR(100)),
                CAST(bx.child_item                          AS NVARCHAR(100)),
                CAST(d.U_ITNO                               AS NVARCHAR(100)),
                CAST(itm.ItemName                           AS NVARCHAR(200)),
                CAST(bx.qty_per_fg * d.U_QTY              AS DECIMAL(28,10)),
                CAST(bx.bom_level + 1                      AS INT),
                CAST(d.U_UOM                                AS NVARCHAR(50)),
                CAST('BOM'                                  AS NVARCHAR(10)),
                CAST(bx.explosion_path + ' > ' + d.U_ITNO AS NVARCHAR(4000))
            FROM BOM_Explosion bx
            INNER JOIN [@C_BOMH] h
                ON  h.U_ITNO   = bx.child_item
                AND h.U_STATUS = 'Active'
            INNER JOIN [@C_BOMD] d
                ON  d.DocEntry = h.DocEntry
            INNER JOIN OITM itm
                ON  itm.ItemCode = d.U_ITNO
            WHERE bx.bom_level < 4
              AND (
                    bx.bom_level = 0
                    OR bx.child_item LIKE 'RF%'
                    OR bx.child_item LIKE 'FG%'
                    OR bx.child_item LIKE 'FGE%'
                    OR bx.child_item LIKE 'RFG%'
                    OR bx.child_item LIKE 'FGH%'
                    OR bx.child_item LIKE 'FB%'
                    OR bx.child_item LIKE 'FBS%'
                    OR bx.child_item LIKE 'PM%'
                    OR bx.child_item LIKE 'PX%'
                    OR bx.child_item LIKE 'RM%'
                  )
              -- When the top-level FG itself starts with FB, stop after level 1
              AND NOT (bx.fg_item LIKE 'FB%' AND bx.bom_level >= 1)

            UNION ALL

            -- ================================================================
            -- RECURSE BRANCH 2: Formula ingredients
            -- ================================================================
            SELECT
                CAST(bx.fg_item                                  AS NVARCHAR(100)),
                CAST(bx.child_item                               AS NVARCHAR(100)),
                CAST(f1.U_ITEMCODE                               AS NVARCHAR(100)),
                CAST(itm.ItemName                                AS NVARCHAR(200)),
                CAST(
                    bx.qty_per_fg * (f1.U_QTY / NULLIF(fh.U_TLT, 0))
                                                                 AS DECIMAL(28,10)),
                CAST(bx.bom_level + 1                           AS INT),
                CAST(f1.U_UOM                                    AS NVARCHAR(50)),
                CAST('FORMULA'                                   AS NVARCHAR(10)),
                CAST(bx.explosion_path + ' > ' + f1.U_ITEMCODE AS NVARCHAR(4000))
            FROM BOM_Explosion bx
            INNER JOIN [@C_BOMH] h
                ON  h.U_ITNO   = bx.child_item
                AND h.U_STATUS = 'Active'
            INNER JOIN [@OFES] fh
                ON  fh.U_FCODE  = h.U_FORMULA
                AND fh.U_STATUS = 'Active'
            INNER JOIN [@FES1] f1
                ON  f1.DocEntry = fh.DocEntry
            INNER JOIN OITM itm
                ON  itm.ItemCode = f1.U_ITEMCODE
            WHERE bx.bom_level < 4
              AND (
                    bx.bom_level = 0
                    OR bx.child_item LIKE 'RF%'
                    OR bx.child_item LIKE 'FG%'
                    OR bx.child_item LIKE 'FGE%'
                    OR bx.child_item LIKE 'RFG%'
                    OR bx.child_item LIKE 'FGH%'
                    OR bx.child_item LIKE 'FB%'
                    OR bx.child_item LIKE 'FBS%'
                    OR bx.child_item LIKE 'PM%'
                    OR bx.child_item LIKE 'PX%'
                    OR bx.child_item LIKE 'RM%'
                  )
              -- When the top-level FG itself starts with FB, stop after level 1
              AND NOT (bx.fg_item LIKE 'FB%' AND bx.bom_level >= 1)

        )

        -- ================================================================
        -- Pipe CTE results into temp table, add conversion factor column
        -- ================================================================
        SELECT
            bx.bom_level                                                       AS [Level],
            bx.fg_item                                                         AS [Finish Good],
            (SELECT ItemName FROM OITM WHERE ItemCode = bx.fg_item)            AS [Finish Good Name],
            bx.parent_item                                                     AS [Parent_Code],
            bx.child_item                                                      AS [Item Code],
            CASE bx.source
                WHEN 'BOM'     THEN 'BOM Item'
                WHEN 'FORMULA' THEN 'Formula Item'
                ELSE bx.source
            END                                                                AS [Item_Type],
            CAST(bx.qty_per_fg AS DECIMAL(18,6))                              AS [Item Projection],
            itm.InvntryUom                                                     AS [Inventory_UOM],
            itb.ItmsGrpNam                                                     AS [Item Group],
            CONVERT(NUMERIC(19,2), 1.00)                                       AS [Convertion Factor]
        INTO #BomResult
        FROM BOM_Explosion bx
        INNER JOIN OITM itm ON itm.ItemCode = bx.child_item
        LEFT  JOIN OITB itb ON itb.ItmsGrpCod = itm.ItmsGrpCod
        WHERE bx.bom_level > 0
        OPTION (MAXRECURSION 100);

        -- ================================================================
        -- Conversion factor logic from Vikas query
        -- ================================================================

        -- Step 1: Lookup conversion factor per Parent_Code
        UPDATE t
        SET [Convertion Factor] = ISNULL(cd1.U_CONFACTR, 1)
        FROM #BomResult t,
        (
            SELECT DISTINCT cd.U_CONFACTR, ch.U_ITNO
            FROM [@C_ITMSTR] ch, [@C_ITMUOM] cd
            WHERE cd.DocEntry = ch.DocEntry
              AND cd.U_TOUOM <> ''
        ) cd1
        WHERE t.[Parent_Code] = cd1.U_ITNO;

        -- Step 2: Default conversion factor to 1 where still 0
        UPDATE t
        SET [Convertion Factor] = 1
        FROM #BomResult t
        WHERE [Convertion Factor] = 0.00;

        -- Step 3: Recalculate Item Projection applying conversion factor
        --   BOM Items:     UOM='NO'  → qty * factor  |  otherwise → qty (no factor)
        --   Formula Items: always    → qty * factor
        UPDATE t
        SET [Item Projection] =
            CASE
                WHEN [Item_Type] = 'BOM Item'
                THEN (
                    CASE WHEN [Inventory_UOM] = 'NO'
                         THEN [Item Projection] * [Convertion Factor]
                         ELSE [Item Projection]
                    END
                )
                ELSE [Item Projection] * [Convertion Factor]
            END
        FROM #BomResult t;

        -- ================================================================
        -- Final SELECT — 8 columns the frontend expects
        -- ================================================================
        SELECT
            [Item Group],
            [Finish Good],
            [Parent_Code],
            [Finish Good Name],
            [Item Code],
            [Item_Type],
            [Item Projection],
            [Inventory_UOM]
        FROM #BomResult
        ORDER BY [Finish Good], [Level], [Item Code];

        IF OBJECT_ID('tempdb..#BomResult') IS NOT NULL DROP TABLE #BomResult;
    """

    cursor.execute(sql)
    columns = [col[0] for col in cursor.description]
    rows = cursor.fetchall()
    cursor.close()

    return [dict(zip(columns, row)) for row in rows]

# ----------------------------------------------------------------------------------------------

def extract_px_items(results: list[dict]) -> list[dict]:
    """
    Aggregates all PX items from run_query results.
    Returns list of {Item_Code, Total_Qty}
    Uses new BOM_Explosion column names: Item Code and Item Projection.
    """
    px_agg = {}

    for row in results:
        code = row.get('Item Code')
        if not code or not code.startswith('PX'):
            continue

        if code not in px_agg:
            px_agg[code] = {
                'Item_Code': code,
                'Total_Qty': 0
            }

        px_agg[code]['Total_Qty'] += float(row.get('Item Projection') or 0)

    return list(px_agg.values())


def process_file(filepath):
    df = pd.read_excel(filepath, engine='openpyxl', skiprows=1)
    # Drop rows where FG_CODE or QTY is blank/NaN (e.g. trailing empty rows in Excel)
    df = df.dropna(subset=['FG_CODE', 'QTY'])
    # Also drop rows where FG_CODE is an empty string after stripping
    df = df[df['FG_CODE'].astype(str).str.strip() != '']
    json_payload_list = []
    for i in range(len(df)):
        json_payload_list.append({
            'FG_Code': str(df.iloc[i]['FG_CODE']).strip(),
            'FG_Qty': float(df.iloc[i]['QTY'])
        })
    payload = json.dumps(json_payload_list)
    return payload.replace("'", "''")

def generate_production_batch(prefix="BAT"):
   
    date_str = datetime.now().strftime("%Y%m%d-%H%M")
    # Generate 4 random uppercase letters and digits
    chars = string.ascii_uppercase + string.digits
    random_suffix = ''.join(random.choices(chars, k=4))
    
    return f"{prefix}-{date_str}-{random_suffix}"

def conversion(payload : Payload):

    ref = db.reference('vault_system')

    ref.update({
        'status' : 'pending_approval',
        'request_data' : payload.model_dump(),
        'final_result' : None
    })

def approve(excel_password: str):
    ref = db.reference('vault_system')

    ref.update({
        'status' : 'approved',
        'excel_password' : excel_password
    })


def px_to_md_conversion_lookup(db : Session):
    query = text(""" 
    ;WITH FormulaExplosion AS
(
    -----------------------------------------------------------------
    -- Direct Formula Components
    -----------------------------------------------------------------
    SELECT
        BH.U_ITNO              AS PX_Code,
        F1.U_ITEMCODE          AS MD_Code,
        F1.U_QTY               AS Formula_Qty,
        1                      AS Level
    FROM [@C_BOMH] BH
    INNER JOIN [@OFES] FH
        ON BH.U_FORMULA = FH.U_FCODE
    INNER JOIN [@FES1] F1
        ON FH.DocEntry = F1.DocEntry
    WHERE BH.U_STATUS = 'Active'
      AND FH.U_STATUS = 'Active'

    UNION ALL

    -----------------------------------------------------------------
    -- Recursive Formula Expansion
    -----------------------------------------------------------------
    SELECT
        FE.PX_Code,
        F1.U_ITEMCODE,
        F1.U_QTY,
        FE.Level + 1
    FROM FormulaExplosion FE
    INNER JOIN [@C_BOMH] BH
        ON FE.MD_Code = BH.U_ITNO
    INNER JOIN [@OFES] FH
        ON BH.U_FORMULA = FH.U_FCODE
    INNER JOIN [@FES1] F1
        ON FH.DocEntry = F1.DocEntry
    WHERE BH.U_STATUS = 'Active'
      AND FH.U_STATUS = 'Active'
)
SELECT
    PX_Code,
    MD_Code,
    Formula_Qty,
    ROUND(
        Formula_Qty * 100.0 /
        SUM(Formula_Qty) OVER (PARTITION BY PX_Code),
        4
    ) AS Formula_Percent,
    Level
FROM FormulaExplosion
ORDER BY
    PX_Code,
    MD_Code
OPTION (MAXRECURSION 1000);
    """)

    results = db.execute(query)
    records = results.mappings().all()

    conversion_dict = defaultdict(list)
    for item in records:
        px_val = item.get('PX_Code')
        px_val = px_val.lower()
        print(px_val)
        md_val = item.get('MD_Code')
        percentage = item.get('Formula_Percent')
        
        # Appends to a list instead of overwriting
        conversion_dict[px_val].append({
            'MD_code' : md_val,
            'percentage' : percentage
        })

    return conversion_dict

def convert_px_to_md(px_codes : list[dict], conversion_dict):
    final_list = []
    for px_code in px_codes:
        # print(px_code)
        # final_list.extend(conversion_dict[px_code.get('Item_Code')])
        px_codes_lower = px_code.get('Item_Code')
        px_codes_lower = px_codes_lower.lower()
        print(f'px code to match : {px_codes_lower}')
        for item in conversion_dict[px_codes_lower]:
            final_list.append({'MD_code' : item.get('MD_code'),
            'Qty' : (float(item.get('percentage')) / 100) * float(px_code.get('Total_Qty'))})
    # print(final_list)

    return final_list

def get_closing_stock(db: Session):
    
    
    today_date = datetime.today().strftime('%Y%m%d')

    
    sql_query = text("""
        SELECT 
            OITM.ItemCode, 
            SUM(m.InQty - m.OutQty) AS Closing_Qty
        FROM OITM (NOLOCK)
        JOIN OINM (NOLOCK) m ON OITM.ItemCode = m.ItemCode
        WHERE OITM.InvntItem = 'Y'
          AND m.DocDate <= :date_to
          AND (:whs_code IS NULL OR m.Warehouse = :whs_code)
          AND m.Warehouse IN (
              'DFGN01', 'DFGN02', 'DFGN03', 'DFGN04', 'DFGN05', 'DFGN06', 'DFGN09',
              'DFGN10', 'DFGN11', 'DFGN14', 'DFGN15', 'DFGN16', 'DFGN17', 'DFGN18', 
              'RMSTOR10', 'RMSTOR11', 'RMSTOR13', 'RMSTORE1', 'RMSTORE2', 'RMSTORE3', 
              'RMSTORE4', 'RMSTORE5', 'RMSTORE6'
          )
        GROUP BY 
            OITM.ItemCode;
    """)

    # 3. Execute the query, passing the dynamic variables as a dictionary
    result = db.execute(sql_query, {
        "date_to": today_date, 
        "whs_code": None
    })
    
    records = result.mappings().all()
    
    # 2. Transform the list of rows into a single Key: Value dictionary
    stock_dict = {row['ItemCode']: row['Closing_Qty'] for row in records}
    
    # Return the clean dictionary
    return stock_dict

