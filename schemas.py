from pydantic import BaseModel, Field, RootModel, EmailStr

class register_user(BaseModel):
    email : EmailStr
    password : str
    role : int

class user_update(BaseModel):
    email : EmailStr
    role : int
    password : str | None = None

class multi_id_input(BaseModel):
    ids : list


MultiIdInputWithOpBalanceForecastQty = RootModel[list[dict[str, tuple[int, int]]]]
# [
#     {'FG01' : [100, 10] } # [forecast_quantity, opening_balance],
#     {'FG02' : [110, 40] } # [forecast_quantity, opening_balance]
# ]


get_multiple_item = RootModel[dict[str, dict[str, int]]]
