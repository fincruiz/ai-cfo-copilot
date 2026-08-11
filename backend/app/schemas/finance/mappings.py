from datetime import datetime
from uuid import UUID
from pydantic import BaseModel,ConfigDict,Field
class AccountMappingUpsert(BaseModel):
    source_account_code:str=Field(min_length=1); source_account_name:str|None=None; statement:str; reporting_group:str; reporting_subgroup:str|None=None; sign_convention:str="positive"; display_order:int|None=None; is_confirmed:bool=True
class AccountMappingResponse(AccountMappingUpsert):
    model_config=ConfigDict(from_attributes=True)
    id:UUID; company_id:UUID; created_at:datetime; updated_at:datetime
class MappingSuggestionResponse(BaseModel):
    source_account_code:str; source_account_name:str|None; statement:str; reporting_group:str; reporting_subgroup:str|None; sign_convention:str; confidence:float; reason:str
class MappingBulkRequest(BaseModel):items:list[AccountMappingUpsert]
