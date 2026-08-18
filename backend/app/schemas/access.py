from pydantic import BaseModel, Field

class InvitationCreate(BaseModel):
    email:str=Field(min_length=3,max_length=320)
    role:str="viewer"
    expires_in_days:int=Field(default=7,ge=1,le=30)

class InvitationAccept(BaseModel):
    token:str=Field(min_length=32,max_length=512)

class PersonalProfileUpdate(BaseModel):
    full_name:str=Field(min_length=2,max_length=255)
    job_title:str|None=Field(default=None,max_length=255)

class MemberRoleUpdate(BaseModel):
    role:str
