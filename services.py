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

def read_fg_excel(filepath):
    if filepath.lower().endswith('.csv'):
        df = pd.read_csv(filepath)
    else:
        try:
            df = pd.read_excel(filepath, engine='openpyxl', skiprows=1)
            cols = [str(c).strip() for c in df.columns]
            if 'FG_CODE' not in cols or 'QTY' not in cols:
                df0 = pd.read_excel(filepath, engine='openpyxl', skiprows=0)
                cols0 = [str(c).strip() for c in df0.columns]
                if 'FG_CODE' in cols0 and 'QTY' in cols0:
                    df = df0
        except Exception:
            try:
                df = pd.read_excel(filepath, skiprows=1)
                cols = [str(c).strip() for c in df.columns]
                if 'FG_CODE' not in cols or 'QTY' not in cols:
                    df0 = pd.read_excel(filepath, skiprows=0)
                    cols0 = [str(c).strip() for c in df0.columns]
                    if 'FG_CODE' in cols0 and 'QTY' in cols0:
                        df = df0
            except Exception:
                df = pd.read_excel(filepath, engine='openpyxl', skiprows=0)
    return df

def check_file_structure_and_conditions(filepath):
    df = read_fg_excel(filepath)
    column_list = [str(c).strip() for c in df.columns]
    print(column_list)
    if 'FG_CODE' not in column_list or 'QTY' not in column_list:
        return {'status' : 'error', 'details' : 'File structure not correct (missing FG_CODE and QTY headers)'}
    df.columns = column_list
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

def run_query(json_payload: str, db: Session):
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
        FG_Qty NUMERIC(19,4)
        );

        INSERT INTO @FGQty (FG_Code, FG_Qty)
        VALUES
        {values_rows};

        ;WITH BOM_Explosion AS (

        -- ================================================================
        -- ANCHOR LEVEL 0: The FG itself, seeded with requested qty
        -- ================================================================
        SELECT
        CAST(h.U_ITNO AS NVARCHAR(100)) AS fg_item,
        CAST(h.U_ITNO AS NVARCHAR(100)) AS parent_item,
        CAST(h.U_ITNO AS NVARCHAR(100)) AS child_item,
        CAST(itm.ItemName AS NVARCHAR(200)) AS child_name,
        CAST(q.FG_Qty AS DECIMAL(28,10)) AS qty_per_fg,
        CAST(0 AS INT) AS bom_level,
        CAST(h.U_UOM AS NVARCHAR(50)) AS uom,
        CAST('BOM' AS NVARCHAR(10)) AS source,
        CAST(h.U_ITNO AS NVARCHAR(4000)) AS explosion_path
        FROM [@C_BOMH] h
        INNER JOIN OITM itm ON itm.ItemCode = h.U_ITNO
        INNER JOIN @FGQty q ON q.FG_Code = h.U_ITNO
        WHERE h.U_STATUS = 'Active'

        UNION ALL

        -- ================================================================
        -- RECURSE BRANCH 1: BOM detail children
        -- ================================================================
        SELECT
        CAST(bx.fg_item AS NVARCHAR(100)),
        CAST(bx.child_item AS NVARCHAR(100)),
        CAST(d.U_ITNO AS NVARCHAR(100)),
        CAST(itm.ItemName AS NVARCHAR(200)),
        CAST(bx.qty_per_fg * d.U_QTY AS DECIMAL(28,10)),
        CAST(bx.bom_level + 1 AS INT),
        CAST(d.U_UOM AS NVARCHAR(50)),
        CAST('BOM' AS NVARCHAR(10)),
        CAST(bx.explosion_path + ' > ' + d.U_ITNO AS NVARCHAR(4000))
        FROM BOM_Explosion bx
        INNER JOIN [@C_BOMH] h
            ON h.U_ITNO = bx.child_item
            AND h.U_STATUS = 'Active'
        INNER JOIN [@C_BOMD] d
            ON d.DocEntry = h.DocEntry
        INNER JOIN OITM itm
            ON itm.ItemCode = d.U_ITNO
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
        AND NOT (bx.fg_item LIKE 'FB%' AND bx.bom_level >= 1)
        AND NOT (bx.parent_item LIKE 'RM%')

        UNION ALL

        -- ================================================================
        -- RECURSE BRANCH 2: Formula ingredients
        -- ================================================================
        SELECT
        CAST(bx.fg_item AS NVARCHAR(100)),
        CAST(bx.child_item AS NVARCHAR(100)),
        CAST(f1.U_ITEMCODE AS NVARCHAR(100)),
        CAST(itm.ItemName AS NVARCHAR(200)),
        CAST(
            bx.qty_per_fg * (f1.U_QTY / NULLIF(fh.U_TLT, 0))
            AS DECIMAL(28,10)),
        CAST(bx.bom_level + 1 AS INT),
        CAST(f1.U_UOM AS NVARCHAR(50)),
        CAST('FORMULA' AS NVARCHAR(10)),
        CAST(bx.explosion_path + ' > ' + f1.U_ITEMCODE AS NVARCHAR(4000))
        FROM BOM_Explosion bx
        INNER JOIN [@C_BOMH] h
            ON h.U_ITNO = bx.child_item
            AND h.U_STATUS = 'Active'
        INNER JOIN [@OFES] fh
            ON fh.U_FCODE = h.U_FORMULA
            AND fh.U_STATUS = 'Active'
        INNER JOIN [@FES1] f1
            ON f1.DocEntry = fh.DocEntry
        INNER JOIN OITM itm
            ON itm.ItemCode = f1.U_ITEMCODE
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
        AND NOT (bx.fg_item LIKE 'FB%' AND bx.bom_level >= 1)
        AND NOT (bx.parent_item LIKE 'RM%')

        )

        -- ================================================================
        -- Pipe CTE results into temp table
        -- ================================================================
        SELECT
        bx.bom_level AS [Level],
        bx.fg_item AS [Finish Good],
        (SELECT ItemName FROM OITM WHERE ItemCode = bx.fg_item) AS [Finish Good Name],
        bx.parent_item AS [Parent_Code],
        bx.child_item AS [Item Code],
        CASE bx.source
            WHEN 'BOM' THEN 'BOM Item'
            WHEN 'FORMULA' THEN 'Formula Item'
            ELSE bx.source
        END AS [Item_Type],
        CAST(bx.qty_per_fg AS DECIMAL(18,6)) AS [Item Projection],
        itm.InvntryUom AS [Inventory_UOM],
        itb.ItmsGrpNam AS [Item Group]
        INTO #BomResult
        FROM BOM_Explosion bx
        INNER JOIN OITM itm ON itm.ItemCode = bx.child_item
        LEFT JOIN OITB itb ON itb.ItmsGrpCod = itm.ItmsGrpCod
        WHERE bx.bom_level > 0
        OPTION (MAXRECURSION 100);


        -- ================================================================
        -- Final SELECT
        -- Items with Item Projection = 0 are excluded (BOM U_QTY is 0 in
        -- SAP for those items — showing them as 0 in the portal is misleading).
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
        WHERE [Item Projection] > 0
        ORDER BY [Finish Good], [Level], [Item Code];


        IF OBJECT_ID('tempdb..#BomResult') IS NOT NULL DROP TABLE #BomResult;
    """

    try:
        cursor.execute(sql)
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
    finally:
        cursor.close()  # Always release the cursor to free SAP locks

    results = [dict(zip(columns, row)) for row in rows]
    # Normalize R&D and FG item codes to uppercase so mixed-case entries
    # in SAP (e.g. 'R&d111' vs 'R&D111') are treated as the same item.
    for r in results:
        code = r.get('Item Code') or ''
        upper = code.upper()
        if upper.startswith('R&D') or upper.startswith('FG'):
            r['Item Code'] = upper
    return results

# ---------------------------------------------------------------------------------------

# ----------------------------------------------- vedant query -----------------------------------------------

# def run_query(json_payload: str, db: Session) -> list[dict]:
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

#         IF OBJECT_ID('tempdb..#BomResult') IS NOT NULL DROP TABLE #BomResult;

#         DECLARE @FGQty TABLE
#         (
#             FG_Code NVARCHAR(50),
#             FG_Qty  NUMERIC(19,4)
#         );

#         INSERT INTO @FGQty (FG_Code, FG_Qty)
#         VALUES
#         {values_rows};

#         ;WITH BOM_Explosion AS (

#             -- ================================================================
#             -- ANCHOR LEVEL 0: The FG itself, seeded with requested qty.
#             -- We carry is_fb_root = 1 if the root FG starts with 'FB',
#             -- so every recursive row knows the root type without re-checking fg_item.
#             -- ================================================================
#             SELECT
#                 CAST(h.U_ITNO     AS NVARCHAR(100))  AS fg_item,
#                 CAST(h.U_ITNO     AS NVARCHAR(100))  AS parent_item,
#                 CAST(h.U_ITNO     AS NVARCHAR(100))  AS child_item,
#                 CAST(itm.ItemName AS NVARCHAR(200))  AS child_name,
#                 CAST(q.FG_Qty     AS DECIMAL(28,10)) AS qty_per_fg,
#                 CAST(0            AS INT)             AS bom_level,
#                 CAST(h.U_UOM      AS NVARCHAR(50))    AS uom,
#                 CAST('BOM'        AS NVARCHAR(10))    AS source,
#                 CAST(h.U_ITNO     AS NVARCHAR(4000))  AS explosion_path,
#                 -- Carry a flag: 1 if the root FG is an FB item, 0 otherwise.
#                 -- This propagates unchanged through every recursive step.
#                 CAST(CASE WHEN h.U_ITNO LIKE 'FB%' THEN 1 ELSE 0 END AS BIT) AS is_fb_root
#             FROM [@C_BOMH] h
#             INNER JOIN OITM itm ON itm.ItemCode = h.U_ITNO
#             INNER JOIN @FGQty q  ON q.FG_Code   = h.U_ITNO
#             WHERE h.U_STATUS = 'Active'

#             UNION ALL

#             -- ================================================================
#             -- RECURSE BRANCH 1: BOM detail children
#             -- Depth rule: if the root FG is FB, only expand level 0 → level 1.
#             --             For all other roots, expand up to level 4.
#             -- "Only expand from bx" means: bx.bom_level must be 0 for FB roots.
#             -- ================================================================
#             SELECT
#                 CAST(bx.fg_item                              AS NVARCHAR(100)),
#                 CAST(bx.child_item                           AS NVARCHAR(100)),
#                 CAST(d.U_ITNO                                AS NVARCHAR(100)),
#                 CAST(itm.ItemName                            AS NVARCHAR(200)),
#                 CAST(bx.qty_per_fg * d.U_QTY                AS DECIMAL(28,10)),
#                 CAST(bx.bom_level + 1                        AS INT),
#                 CAST(d.U_UOM                                 AS NVARCHAR(50)),
#                 CAST('BOM'                                   AS NVARCHAR(10)),
#                 CAST(bx.explosion_path + ' > ' + d.U_ITNO   AS NVARCHAR(4000)),
#                 bx.is_fb_root
#             FROM BOM_Explosion bx
#             INNER JOIN [@C_BOMH] h
#                 ON  h.U_ITNO   = bx.child_item
#                 AND h.U_STATUS = 'Active'
#             INNER JOIN [@C_BOMD] d
#                 ON  d.DocEntry = h.DocEntry
#             INNER JOIN OITM itm
#                 ON  itm.ItemCode = d.U_ITNO
#             WHERE
#                 -- For FB roots: only expand from level 0 (producing level 1 only)
#                 -- For all other roots: expand up to level 4
#                 (bx.is_fb_root = 1 AND bx.bom_level = 0)
#                 OR
#                 (bx.is_fb_root = 0 AND bx.bom_level < 4)

#             UNION ALL

#             -- ================================================================
#             -- RECURSE BRANCH 2: Formula ingredients
#             -- Same depth rule applied identically.
#             -- ================================================================
#             SELECT
#                 CAST(bx.fg_item                                  AS NVARCHAR(100)),
#                 CAST(bx.child_item                               AS NVARCHAR(100)),
#                 CAST(f1.U_ITEMCODE                               AS NVARCHAR(100)),
#                 CAST(itm.ItemName                                AS NVARCHAR(200)),
#                 CAST(bx.qty_per_fg * (f1.U_QTY / NULLIF(fh.U_TLT, 0))
#                                                                  AS DECIMAL(28,10)),
#                 CAST(bx.bom_level + 1                            AS INT),
#                 CAST(f1.U_UOM                                    AS NVARCHAR(50)),
#                 CAST('FORMULA'                                   AS NVARCHAR(10)),
#                 CAST(bx.explosion_path + ' > ' + f1.U_ITEMCODE  AS NVARCHAR(4000)),
#                 bx.is_fb_root
#             FROM BOM_Explosion bx
#             INNER JOIN [@C_BOMH] h
#                 ON  h.U_ITNO   = bx.child_item
#                 AND h.U_STATUS = 'Active'
#             INNER JOIN [@OFES] fh
#                 ON  fh.U_FCODE  = h.U_FORMULA
#                 AND fh.U_STATUS = 'Active'
#             INNER JOIN [@FES1] f1
#                 ON  f1.DocEntry = fh.DocEntry
#             INNER JOIN OITM itm
#                 ON  itm.ItemCode = f1.U_ITEMCODE
#             WHERE
#                 -- Same rule as Branch 1
#                 (bx.is_fb_root = 1 AND bx.bom_level = 0)
#                 OR
#                 (bx.is_fb_root = 0 AND bx.bom_level < 4)

#         )

#         -- ================================================================
#         -- Pipe CTE results into temp table, add conversion factor column
#         -- ================================================================
#         SELECT
#             bx.bom_level                                                       AS [Level],
#             bx.fg_item                                                         AS [Finish Good],
#             (SELECT ItemName FROM OITM WHERE ItemCode = bx.fg_item)            AS [Finish Good Name],
#             bx.parent_item                                                     AS [Parent_Code],
#             bx.child_item                                                      AS [Item Code],
#             CASE bx.source
#                 WHEN 'BOM'     THEN 'BOM Item'
#                 WHEN 'FORMULA' THEN 'Formula Item'
#                 ELSE bx.source
#             END                                                                AS [Item_Type],
#             CAST(bx.qty_per_fg AS DECIMAL(18,6))                              AS [Item Projection],
#             itm.InvntryUom                                                     AS [Inventory_UOM],
#             itb.ItmsGrpNam                                                     AS [Item Group],
#             CONVERT(NUMERIC(19,2), 1.00)                                       AS [Convertion Factor]
#         INTO #BomResult
#         FROM BOM_Explosion bx
#         INNER JOIN OITM itm ON itm.ItemCode = bx.child_item
#         LEFT  JOIN OITB itb ON itb.ItmsGrpCod = itm.ItmsGrpCod
#         WHERE bx.bom_level > 0
#         OPTION (MAXRECURSION 100);

#         -- ================================================================
#         -- Conversion factor logic
#         -- ================================================================

#         -- Step 1: Lookup conversion factor per Parent_Code
#         UPDATE t
#         SET [Convertion Factor] = ISNULL(cd1.U_CONFACTR, 1)
#         FROM #BomResult t,
#         (
#             SELECT DISTINCT cd.U_CONFACTR, ch.U_ITNO
#             FROM [@C_ITMSTR] ch, [@C_ITMUOM] cd
#             WHERE cd.DocEntry = ch.DocEntry
#               AND cd.U_TOUOM <> ''
#         ) cd1
#         WHERE t.[Parent_Code] = cd1.U_ITNO;

#         -- Step 2: Default conversion factor to 1 where still 0
#         UPDATE t
#         SET [Convertion Factor] = 1
#         FROM #BomResult t
#         WHERE [Convertion Factor] = 0.00;

#         -- Step 3: Recalculate Item Projection applying conversion factor
#         --   BOM Items:     UOM='NO'  → qty * factor  |  otherwise → qty (no factor)
#         --   Formula Items: always    → qty * factor
#         UPDATE t
#         SET [Item Projection] =
#             CASE
#                 WHEN [Item_Type] = 'BOM Item'
#                 THEN (
#                     CASE WHEN [Inventory_UOM] = 'NO'
#                          THEN [Item Projection] * [Convertion Factor]
#                          ELSE [Item Projection]
#                     END
#                 )
#                 ELSE [Item Projection] * [Convertion Factor]
#             END
#         FROM #BomResult t;

#         -- ================================================================
#         -- Final SELECT
#         -- ================================================================
#         SELECT
#             [Item Group],
#             [Finish Good],
#             [Parent_Code],
#             [Finish Good Name],
#             [Item Code],
#             [Item_Type],
#             [Item Projection],
#             [Inventory_UOM]
#         FROM #BomResult
#         ORDER BY [Finish Good], [Level], [Item Code];

#         IF OBJECT_ID('tempdb..#BomResult') IS NOT NULL DROP TABLE #BomResult;
#     """

#     cursor.execute(sql)
#     columns = [col[0] for col in cursor.description]
#     rows = cursor.fetchall()
#     cursor.close()

#     return [dict(zip(columns, row)) for row in rows]

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
    df = read_fg_excel(filepath)
    df.columns = [str(c).strip() for c in df.columns]
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


def fetch_fg_conversion_factors(fg_codes: list, db: Session) -> dict:
    """
    For each FG code, look up the conversion factor from [@C_ITMSTR] / [@C_ITMUOM].
    Returns {fg_code: conversion_factor} — defaults to 1.0 if no entry is found.

    The factor is pulled as MAX(U_CONFACTR) per FG so that multiple UOM rows
    (e.g. a weight entry AND a piece-count entry) collapse to one deterministic value.
    """
    if not fg_codes:
        return {}

    raw_conn = db.connection().connection
    cursor = raw_conn.cursor()

    placeholders = ', '.join(
        f"N'{code.replace(chr(39), chr(39)*2)}'" for code in fg_codes
    )

    sql = f"""
        SELECT ch.U_ITNO,
               MAX(ISNULL(cd.U_CONFACTR, 1)) AS U_CONFACTR
        FROM   [@C_ITMSTR] ch
        INNER JOIN [@C_ITMUOM] cd ON cd.DocEntry = ch.DocEntry
        WHERE  ch.U_ITNO IN ({placeholders})
          AND  cd.U_TOUOM <> ''
        GROUP BY ch.U_ITNO
    """

    try:
        cursor.execute(sql)
        rows = cursor.fetchall()
    finally:
        cursor.close()

    result = {}
    for row in rows:
        fg_code, cf = row[0], row[1]
        result[fg_code] = float(cf) if cf and float(cf) != 0 else 1.0

    return result


def apply_fg_conversions(json_payload: str, db: Session):
    """
    Looks up the conversion factor for every FG item in the payload,
    multiplies FG_Qty by that factor, and returns:

      converted_payload (str) : JSON string with the converted qty — pass
                                this straight to run_query().
      conversion_info   (dict): keyed by FG_Code, each value is a dict:
                                  {
                                    'Uploaded_Qty'      : float,  # qty from Excel
                                    'Conversion_Factor' : float,  # CF from SAP (1.0 if none)
                                    'Converted_Qty'     : float   # Uploaded_Qty * CF
                                  }

    Use conversion_info to annotate the BOM result rows returned by run_query().
    """
    import json as _json

    items = _json.loads(json_payload)          # safe — single-quotes are fine in JSON strings
    fg_codes = [item['FG_Code'] for item in items]
    cf_map = fetch_fg_conversion_factors(fg_codes, db)

    conversion_info: dict = {}
    converted_items: list = []

    for item in items:
        fg_code = item['FG_Code']
        uploaded_qty = float(item['FG_Qty'])
        cf = cf_map.get(fg_code, 1.0)
        converted_qty = round(uploaded_qty * cf, 6)

        conversion_info[fg_code] = {
            'Uploaded_Qty':       uploaded_qty,
            'Conversion_Factor':  cf,
            'Converted_Qty':      converted_qty,
        }
        converted_items.append({'FG_Code': fg_code, 'FG_Qty': converted_qty})

    converted_payload = _json.dumps(converted_items)
    return converted_payload, conversion_info


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
      AND FE.Level < 10
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
OPTION (MAXRECURSION 100);
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
        px_codes_lower = px_codes_lower.lower() if px_codes_lower else ''
        print(f'px code to match : {px_codes_lower}')
        if px_codes_lower in conversion_dict and conversion_dict[px_codes_lower]:
            for item in conversion_dict[px_codes_lower]:
                final_list.append({'MD_code' : item.get('MD_code'),
                'Qty' : (float(item.get('percentage') or 0) / 100) * float(px_code.get('Total_Qty') or 0)})
        else:
            final_list.append({'MD_code' : px_code.get('Item_Code'),
            'Qty' : float(px_code.get('Total_Qty') or 0)})
    # print(final_list)

    return final_list

def get_closing_stock(db: Session):
    
    today_date = datetime.today().strftime('%Y%m%d')

    sql_query = text("""
        SELECT 
            OITM.ItemCode,
            OITM.LastPurPrc,
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
            OITM.ItemCode,
            OITM.LastPurPrc;
    """)

    result = db.execute(sql_query, {
        "date_to": today_date, 
        "whs_code": None
    })
    
    records = result.mappings().all()
    
    # Return a dict: { ItemCode: { 'qty': <Closing_Qty>, 'last_pur_prc': <LastPurPrc> } }
    stock_dict = {
        row['ItemCode']: {
            'qty': float(row['Closing_Qty']) if row['Closing_Qty'] is not None else 0.0,
            'last_pur_prc': float(row['LastPurPrc']) if row['LastPurPrc'] is not None else 0.0
        }
        for row in records
    }
    
    return stock_dict


# ── PRICE LIST — POSTGRES-BACKED ──────────────────────────────────────────────

def save_price_list_to_db(filepath: str, uploaded_by: str, db_admin) -> int:
    """
    Parse the price-list Excel (required columns: ItemCode, Price).
    Truncates the `price_list` table and inserts a fresh row per item so the
    table always reflects the latest upload. Returns the number of rows saved.

    Parameters
    ----------
    filepath    : path to the uploaded .xlsx file
    uploaded_by : username of the admin who uploaded (for audit)
    db_admin    : SQLAlchemy Session bound to the Postgres admin DB
    """
    import models as _models

    df = pd.read_excel(filepath, engine='openpyxl')
    df.columns = [c.strip() for c in df.columns]

    if 'ItemCode' not in df.columns or 'Price' not in df.columns:
        raise ValueError("Price Excel must have columns: ItemCode, Price")

    df = df.dropna(subset=['ItemCode', 'Price'])
    df['ItemCode'] = df['ItemCode'].astype(str).str.strip()
    df['Price']    = pd.to_numeric(df['Price'], errors='coerce').fillna(0.0)

    now = datetime.utcnow()

    # Full replace: delete all then bulk insert
    db_admin.query(_models.PriceList).delete()
    db_admin.bulk_save_objects([
        _models.PriceList(
            item_code   = row['ItemCode'],
            price       = float(row['Price']),
            uploaded_by = uploaded_by,
            uploaded_at = now,
        )
        for _, row in df.iterrows()
    ])
    db_admin.commit()
    return len(df)


def get_price_list_from_db(db_admin) -> dict:
    """
    Returns the active price list as {item_code: price} from Postgres.
    Returns an empty dict if no price list has been uploaded yet.
    """
    import models as _models
    rows = db_admin.query(_models.PriceList.item_code, _models.PriceList.price).all()
    return {row.item_code: float(row.price) for row in rows}




def run_rf_explosion(rf_items: list[dict], db: Session) -> list[dict]:
    """
    Given a list of {'code': <RF_code>, 'qty': <float>} dicts that came in
    from the PX → MD conversion, explodes each RF code through the BOM and
    returns only the leaf raw-material rows (RM, PK, LB, R&D, PX prefixes).

    Parameters
    ----------
    rf_items : list[dict]
        e.g. [{'code': 'RF001', 'qty': 123.45}, ...]
    db : SQLAlchemy Session  (same production DB used by run_query)

    Returns
    -------
    list[dict] with keys: 'RF_Root', 'Parent_Code', 'Item Code',
                          'Item Projection', 'Inventory_UOM'
    """
    if not rf_items:
        return []

    values_rows = ",\n".join(
        f"(N'{item['code'].replace(chr(39), chr(39)*2)}', {item['qty']})"
        for item in rf_items
    )

    raw_conn = db.connection().connection
    cursor = raw_conn.cursor()

    sql = f"""
        SET NOCOUNT ON;

        IF OBJECT_ID('tempdb..#RfResult') IS NOT NULL DROP TABLE #RfResult;

        DECLARE @RFQty TABLE
        (
            FG_Code NVARCHAR(50),
            FG_Qty  NUMERIC(19,4)
        );

        INSERT INTO @RFQty (FG_Code, FG_Qty)
        VALUES
        {values_rows};

        ;WITH BOM_Explosion AS (

            -- Anchor: the RF item itself
            SELECT
                CAST(h.U_ITNO AS NVARCHAR(100)) AS fg_item,
                CAST(h.U_ITNO AS NVARCHAR(100)) AS parent_item,
                CAST(h.U_ITNO AS NVARCHAR(100)) AS child_item,
                CAST(q.FG_Qty AS DECIMAL(28,10)) AS qty_per_fg,
                CAST(0 AS INT) AS bom_level,
                CAST(h.U_UOM AS NVARCHAR(50)) AS uom,
                CAST('BOM' AS NVARCHAR(10)) AS source,
                CAST(h.U_ITNO AS NVARCHAR(4000)) AS explosion_path
            FROM [@C_BOMH] h
            INNER JOIN @RFQty q ON q.FG_Code = h.U_ITNO
            WHERE h.U_STATUS = 'Active'

            UNION ALL

            -- BOM children
            SELECT
                CAST(bx.fg_item AS NVARCHAR(100)),
                CAST(bx.child_item AS NVARCHAR(100)),
                CAST(d.U_ITNO AS NVARCHAR(100)),
                CAST(bx.qty_per_fg * d.U_QTY AS DECIMAL(28,10)),
                CAST(bx.bom_level + 1 AS INT),
                CAST(d.U_UOM AS NVARCHAR(50)),
                CAST('BOM' AS NVARCHAR(10)),
                CAST(bx.explosion_path + ' > ' + d.U_ITNO AS NVARCHAR(4000))
            FROM BOM_Explosion bx
            INNER JOIN [@C_BOMH] h
                ON h.U_ITNO = bx.child_item AND h.U_STATUS = 'Active'
            INNER JOIN [@C_BOMD] d
                ON d.DocEntry = h.DocEntry
            INNER JOIN OITM itm ON itm.ItemCode = d.U_ITNO
            WHERE bx.bom_level < 4
            AND NOT (bx.parent_item LIKE 'RM%')

            UNION ALL

            -- Formula ingredients
            SELECT
                CAST(bx.fg_item AS NVARCHAR(100)),
                CAST(bx.child_item AS NVARCHAR(100)),
                CAST(f1.U_ITEMCODE AS NVARCHAR(100)),
                CAST(bx.qty_per_fg * (f1.U_QTY / NULLIF(fh.U_TLT, 0)) AS DECIMAL(28,10)),
                CAST(bx.bom_level + 1 AS INT),
                CAST(f1.U_UOM AS NVARCHAR(50)),
                CAST('FORMULA' AS NVARCHAR(10)),
                CAST(bx.explosion_path + ' > ' + f1.U_ITEMCODE AS NVARCHAR(4000))
            FROM BOM_Explosion bx
            INNER JOIN [@C_BOMH] h
                ON h.U_ITNO = bx.child_item AND h.U_STATUS = 'Active'
            INNER JOIN [@OFES] fh
                ON fh.U_FCODE = h.U_FORMULA AND fh.U_STATUS = 'Active'
            INNER JOIN [@FES1] f1
                ON f1.DocEntry = fh.DocEntry
            INNER JOIN OITM itm ON itm.ItemCode = f1.U_ITEMCODE
            WHERE bx.bom_level < 4
            AND NOT (bx.parent_item LIKE 'RM%')
        )

        SELECT
            bx.fg_item       AS [RF_Root],
            bx.parent_item   AS [Parent_Code],
            bx.child_item    AS [Item Code],
            CAST(bx.qty_per_fg AS DECIMAL(18,6)) AS [Item Projection],
            itm.InvntryUom   AS [Inventory_UOM]
        INTO #RfResult
        FROM BOM_Explosion bx
        INNER JOIN OITM itm ON itm.ItemCode = bx.child_item
        WHERE bx.bom_level > 0
        OPTION (MAXRECURSION 100);

        -- UOM Normalization to KG
        UPDATE t
        SET [Item Projection] =
            CASE [Inventory_UOM]
                WHEN 'G'  THEN [Item Projection] / 1000.0
                WHEN 'MG' THEN [Item Projection] / 1000000.0
                WHEN 'ML' THEN [Item Projection] / 1000.0
                WHEN 'L'  THEN [Item Projection] * 1.0
                ELSE [Item Projection]
            END,
        [Inventory_UOM] =
            CASE [Inventory_UOM]
                WHEN 'G'  THEN 'KG'
                WHEN 'MG' THEN 'KG'
                WHEN 'ML' THEN 'KG'
                WHEN 'L'  THEN 'KG'
                ELSE [Inventory_UOM]
            END
        FROM #RfResult t;

        -- Return only leaf-level procurement items
        SELECT [RF_Root], [Parent_Code], [Item Code], [Item Projection], [Inventory_UOM]
        FROM #RfResult
        WHERE
            [Item Code] LIKE 'RM%'
            OR [Item Code] LIKE 'PK%'
            OR [Item Code] LIKE 'LB%'
            OR [Item Code] LIKE 'R&d%'
            OR [Item Code] LIKE 'R&D%'
            OR [Item Code] LIKE 'PX%'
        ORDER BY [RF_Root], [Item Code];

        IF OBJECT_ID('tempdb..#RfResult') IS NOT NULL DROP TABLE #RfResult;
    """

    cursor.execute(sql)
    columns = [col[0] for col in cursor.description]
    rows = cursor.fetchall()
    cursor.close()

    results = [dict(zip(columns, row)) for row in rows]
    # Normalize R&D and FG item codes to uppercase
    for r in results:
        code = r.get('Item Code') or ''
        upper = code.upper()
        if upper.startswith('R&D') or upper.startswith('FG'):
            r['Item Code'] = upper
    return results


def cleanup_firebase():
    ref = db.reference('vault_system')
    ref.delete()


def get_px_formula_costs(db: Session) -> dict:
    """
    Returns the PX → MD formula breakdown needed for per-FG cost pricing.

    For each active PX code that has a formula in the MD database, returns:
      {
        "PX0001": [
          { "md_code": "MD001", "formula_qty": 0.45, "formula_pct": 45.0 },
          ...
        ],
        ...
      }

    formula_qty is the raw ingredient qty per batch.
    formula_pct is the percentage contribution of each MD ingredient (sums to 100 per PX code).
    The caller should weight by formula_pct to compute the PX item's cost from constituent MD costs.
    """
    query = text("""
    ;WITH FormulaExplosion AS (
        SELECT
            BH.U_ITNO       AS PX_Code,
            F1.U_ITEMCODE   AS MD_Code,
            CAST(F1.U_QTY AS FLOAT)  AS Formula_Qty,
            1               AS Level
        FROM [@C_BOMH] BH
        INNER JOIN [@OFES] FH ON BH.U_FORMULA = FH.U_FCODE
        INNER JOIN [@FES1] F1 ON FH.DocEntry = F1.DocEntry
        WHERE BH.U_STATUS = 'Active'
          AND FH.U_STATUS = 'Active'
          AND BH.U_ITNO LIKE 'PX%'
    )
    SELECT
        PX_Code,
        MD_Code,
        Formula_Qty,
        ROUND(
            Formula_Qty * 100.0 / NULLIF(SUM(Formula_Qty) OVER (PARTITION BY PX_Code), 0),
            6
        ) AS Formula_Percent
    FROM FormulaExplosion
    ORDER BY PX_Code, MD_Code
    OPTION (MAXRECURSION 1000);
    """)

    results = db.execute(query)
    records = results.mappings().all()

    px_formula_dict: dict = {}
    for row in records:
        px = str(row['PX_Code'])
        if px not in px_formula_dict:
            px_formula_dict[px] = []
        px_formula_dict[px].append({
            'md_code': str(row['MD_Code']),
            'formula_qty': float(row['Formula_Qty']),
            'formula_pct': float(row['Formula_Percent']) if row['Formula_Percent'] is not None else 0.0
        })

    return px_formula_dict


def diagnose_bom_zeros(fg_codes: list[str], db: Session) -> dict:
    """
    Diagnostic: for each FG code, check why PK/LB items might show 0
    in the perFG BOM table.

    Returns three sections:
      - bom_headers   : BOM header row + status for each FG
      - bom_detail_pk_lb : U_QTY of every PK/LB child in [@C_BOMD]
                           (0 here = root cause of 0 in BOM projection)
      - formula_check : active formula, U_TLT (total weight), and
                        all ingredients (U_TLT=0 causes NULL projection)
    """
    if not fg_codes:
        return {"bom_headers": [], "bom_detail_pk_lb": [], "formula_check": []}

    # Build safe IN list  e.g.  N'FB0001', N'FB0002'
    in_list = ", ".join(f"N'{c.replace(chr(39), chr(39)*2)}'" for c in fg_codes)

    raw_conn = db.connection().connection
    cursor = raw_conn.cursor()

    # ── CHECK 1: BOM header status ───────────────────────────────────────
    sql_headers = f"""
        SELECT
            h.U_ITNO    AS FG_Code,
            h.U_STATUS  AS BOM_Status,
            h.U_FORMULA AS Formula_Code
        FROM [@C_BOMH] h
        WHERE h.U_ITNO IN ({in_list})
        ORDER BY h.U_ITNO;
    """

    cursor.execute(sql_headers)
    cols = [c[0] for c in cursor.description]
    bom_headers = [dict(zip(cols, row)) for row in cursor.fetchall()]

    # ── CHECK 2: BOM detail U_QTY for PK / LB items ─────────────────────
    sql_detail = f"""
        SELECT
            h.U_ITNO   AS FG_Code,
            d.U_ITNO   AS Item_Code,
            d.U_QTY    AS BOM_Qty,
            d.U_UOM    AS UOM
        FROM [@C_BOMH] h
        INNER JOIN [@C_BOMD] d ON d.DocEntry = h.DocEntry
        WHERE h.U_ITNO IN ({in_list})
          AND (d.U_ITNO LIKE 'PK%' OR d.U_ITNO LIKE 'LB%')
          AND h.U_STATUS = 'Active'
        ORDER BY h.U_ITNO, d.U_ITNO;
    """

    cursor.execute(sql_detail)
    cols = [c[0] for c in cursor.description]
    bom_detail = [dict(zip(cols, row)) for row in cursor.fetchall()]

    # ── CHECK 3: Formula status and total weight ──────────────────────────
    sql_formula = f"""
        SELECT
            h.U_ITNO       AS FG_Code,
            fh.U_FCODE     AS Formula_Code,
            fh.U_STATUS    AS Formula_Status,
            fh.U_TLT       AS Formula_Total_Weight,
            f1.U_ITEMCODE  AS Ingredient_Code,
            f1.U_QTY       AS Ingredient_Qty,
            f1.U_UOM       AS Ingredient_UOM
        FROM [@C_BOMH] h
        INNER JOIN [@OFES] fh
            ON fh.U_FCODE = h.U_FORMULA
        INNER JOIN [@FES1] f1
            ON f1.DocEntry = fh.DocEntry
        WHERE h.U_ITNO IN ({in_list})
          AND h.U_STATUS = 'Active'
        ORDER BY h.U_ITNO, fh.U_FCODE, f1.U_ITEMCODE;
    """

    cursor.execute(sql_formula)
    cols = [c[0] for c in cursor.description]
    formula_check = [dict(zip(cols, row)) for row in cursor.fetchall()]

    cursor.close()

    return {
        "bom_headers": bom_headers,
        "bom_detail_pk_lb": bom_detail,
        "formula_check": formula_check,
    }
