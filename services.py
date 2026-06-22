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
        itb.ItmsGrpNam AS [Item Group],
        CONVERT(NUMERIC(19,2), 1.00) AS [Convertion Factor]
        INTO #BomResult
        FROM BOM_Explosion bx
        INNER JOIN OITM itm ON itm.ItemCode = bx.child_item
        LEFT JOIN OITB itb ON itb.ItmsGrpCod = itm.ItmsGrpCod
        WHERE bx.bom_level > 0
        OPTION (MAXRECURSION 100);

        -- ================================================================
        -- Conversion factor logic
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
        -- Step 4: UOM Normalization to KG
        -- ================================================================
        UPDATE t
        SET [Item Projection] =
            CASE [Inventory_UOM]
                WHEN 'G'  THEN [Item Projection] / 1000.0
                WHEN 'MG' THEN [Item Projection] / 1000000.0
                WHEN 'ML' THEN [Item Projection] / 1000.0
                WHEN 'L'  THEN [Item Projection] * 1.0
                WHEN 'NO' THEN [Item Projection]
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
        FROM #BomResult t;

        -- ================================================================
        -- Final SELECT
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

    try:
        cursor.execute(sql)
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
    finally:
        cursor.close()  # Always release the cursor to free SAP locks

    return [dict(zip(columns, row)) for row in rows]

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


# ── IN-MEMORY MONTHLY PRICE STORE ─────────────────────────────────────────────
# Holds the latest price Excel uploaded by admin: { ItemCode (str): price (float) }
_monthly_price_store: dict = {}


def load_new_prices_from_excel(filepath: str) -> dict:
    """
    Parse the monthly price Excel (columns: ItemCode, Price).
    Updates the in-memory store and returns the new price dict.
    """
    global _monthly_price_store
    df = pd.read_excel(filepath, engine='openpyxl')
    df.columns = [c.strip() for c in df.columns]
    print(df.columns)

    if 'ItemCode' not in df.columns or 'Price' not in df.columns:
        raise ValueError("Price Excel must have columns: ItemCode, Price")

    df = df.dropna(subset=['ItemCode', 'Price'])
    df['ItemCode'] = df['ItemCode'].astype(str).str.strip()
    df['Price'] = pd.to_numeric(df['Price'], errors='coerce').fillna(0.0)

    _monthly_price_store = dict(zip(df['ItemCode'], df['Price'].astype(float)))
    return _monthly_price_store


def get_new_prices() -> dict:
    """Returns the current in-memory monthly price store."""
    return _monthly_price_store


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

    return [dict(zip(columns, row)) for row in rows]


def cleanup_firebase():
    ref = db.reference('vault_system')
    ref.delete()
