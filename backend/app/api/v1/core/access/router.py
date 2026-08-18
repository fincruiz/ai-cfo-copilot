from datetime import datetime,timedelta,timezone
import hashlib,secrets
from typing import Annotated
from uuid import UUID
from fastapi import APIRouter,Depends,status
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.exceptions import ApplicationError
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.dependencies.company import get_current_company,require_company_admin
from app.schemas.auth import CurrentUser
from app.schemas.access import InvitationCreate,InvitationAccept,PersonalProfileUpdate,MemberRoleUpdate
from app.schemas.responses import APIResponse
from app.services.audit_service import AuditService

router=APIRouter(prefix="/access",tags=["Access & invitations"])
ALLOWED={"admin","cfo","finance_manager","accountant","board_member","viewer"}

def digest(token:str)->str:return hashlib.sha256(token.encode()).hexdigest()

@router.get("/me",response_model=APIResponse[dict])
async def access_me(company:Annotated[Company,Depends(get_current_company)],user:Annotated[CurrentUser,Depends(get_current_user)],session:Annotated[AsyncSession,Depends(get_db_session)]):
 row=(await session.execute(text("SELECT role::text role FROM public.company_members WHERE company_id=:c AND user_id=:u AND is_active=true"),{"c":company.id,"u":user.id})).mappings().one()
 role=row["role"]
 return APIResponse(message="Access retrieved.",data={"role":role,"can_write_finance":role in {"owner","admin","cfo","finance_manager","accountant"},"can_manage_members":role in {"owner","admin"},"can_reset_all":role in {"owner","admin"}})

@router.get("/members",response_model=APIResponse[list[dict]])
async def members(company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
 rows=(await session.execute(text("""SELECT cm.id,cm.user_id,cm.role::text role,cm.joined_at,COALESCE(p.full_name,'Workspace user') full_name,p.job_title FROM public.company_members cm LEFT JOIN public.profiles p ON p.id=cm.user_id WHERE cm.company_id=:c AND cm.is_active=true ORDER BY CASE WHEN cm.role::text='owner' THEN 0 ELSE 1 END,cm.joined_at"""),{"c":company.id})).mappings().all()
 return APIResponse(message="Members retrieved.",data=[dict(x) for x in rows])

@router.post("/invitations",response_model=APIResponse[dict],status_code=status.HTTP_201_CREATED)
async def invite(req:InvitationCreate,company:Annotated[Company,Depends(get_current_company)],user:Annotated[CurrentUser,Depends(get_current_user)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
 role=req.role.lower().strip();email=req.email.lower().strip()
 if role not in ALLOWED:raise ApplicationError(message="Unsupported role.",error_code="INVALID_COMPANY_ROLE",status_code=422)
 await session.execute(text("UPDATE public.company_invitations SET status='expired',updated_at=now() WHERE company_id=:c AND lower(email)=:e AND status='pending' AND expires_at<=now()"),{"c":company.id,"e":email})
 if (await session.execute(text("SELECT 1 FROM public.company_invitations WHERE company_id=:c AND lower(email)=:e AND status='pending' AND expires_at>now()"),{"c":company.id,"e":email})).first():
  raise ApplicationError(message="A pending invitation already exists.",error_code="INVITATION_ALREADY_PENDING",status_code=409)
 token=secrets.token_urlsafe(48);expires=datetime.now(timezone.utc)+timedelta(days=req.expires_in_days)
 row=(await session.execute(text("""INSERT INTO public.company_invitations(company_id,email,role,token_hash,invited_by,expires_at) VALUES(:c,:e,CAST(:r AS public.company_role),:h,:u,:x) RETURNING id,email,role::text role,status,expires_at"""),{"c":company.id,"e":email,"r":role,"h":digest(token),"u":user.id,"x":expires})).mappings().one()
 await session.commit()
 await AuditService(session).record(company_id=company.id,user_id=user.id,action="member_invitation_created",module="access",summary=f"Invitation created for {email}.",metadata={"role":role},commit=True)
 data=dict(row);data["invite_token"]=token
 return APIResponse(message="Invitation created. The secure token is shown once.",data=data)

@router.get("/invitations",response_model=APIResponse[list[dict]])
async def invitations(company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
 await session.execute(text("UPDATE public.company_invitations SET status='expired',updated_at=now() WHERE company_id=:c AND status='pending' AND expires_at<=now()"),{"c":company.id})
 rows=(await session.execute(text("SELECT id,email,role::text role,status,expires_at,created_at FROM public.company_invitations WHERE company_id=:c ORDER BY created_at DESC LIMIT 100"),{"c":company.id})).mappings().all();await session.commit()
 return APIResponse(message="Invitations retrieved.",data=[dict(x) for x in rows])

@router.post("/invitations/{iid}/revoke",response_model=APIResponse[dict])
async def revoke(iid:UUID,company:Annotated[Company,Depends(get_current_company)],user:Annotated[CurrentUser,Depends(get_current_user)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
 row=(await session.execute(text("UPDATE public.company_invitations SET status='revoked',revoked_at=now(),updated_at=now() WHERE id=:i AND company_id=:c AND status='pending' RETURNING id,email"),{"i":iid,"c":company.id})).mappings().first()
 if not row:raise ApplicationError(message="Pending invitation not found.",error_code="INVITATION_NOT_FOUND",status_code=404)
 await session.commit();await AuditService(session).record(company_id=company.id,user_id=user.id,action="member_invitation_revoked",module="access",summary=f"Invitation revoked for {row['email']}.",commit=True)
 return APIResponse(message="Invitation revoked.",data={"id":str(iid)})

@router.post("/invitations/accept",response_model=APIResponse[dict])
async def accept(req:InvitationAccept,user:Annotated[CurrentUser,Depends(get_current_user)],session:Annotated[AsyncSession,Depends(get_db_session)]):
 row=(await session.execute(text("SELECT * FROM public.company_invitations WHERE token_hash=:h FOR UPDATE"),{"h":digest(req.token)})).mappings().first()
 if not row:raise ApplicationError(message="Invitation is invalid.",error_code="INVITATION_INVALID",status_code=404)
 if row["status"]!="pending":raise ApplicationError(message="Invitation is no longer available.",error_code="INVITATION_NOT_PENDING",status_code=409)
 if row["expires_at"]<=datetime.now(timezone.utc):
  await session.execute(text("UPDATE public.company_invitations SET status='expired' WHERE id=:i"),{"i":row["id"]});await session.commit();raise ApplicationError(message="Invitation expired.",error_code="INVITATION_EXPIRED",status_code=410)
 if not user.email or user.email.lower()!=row["email"].lower():raise ApplicationError(message="Sign in with the invited email address.",error_code="INVITATION_EMAIL_MISMATCH",status_code=403)
 await session.execute(text("""INSERT INTO public.company_members(company_id,user_id,role,is_active,invited_by) VALUES(:c,:u,CAST(:r AS public.company_role),false,:b) ON CONFLICT(company_id,user_id) DO UPDATE SET role=EXCLUDED.role,is_active=false,invited_by=EXCLUDED.invited_by,updated_at=now()"""),{"c":row["company_id"],"u":user.id,"r":str(row["role"]),"b":row["invited_by"]})
 await session.execute(text("UPDATE public.company_invitations SET status='accepted',accepted_by=:u,accepted_at=now(),updated_at=now() WHERE id=:i"),{"u":user.id,"i":row["id"]});await session.commit()
 return APIResponse(message="Invitation accepted. Complete your profile to activate access.",data={"profile_required":True,"company_id":str(row["company_id"])})

@router.get("/profile",response_model=APIResponse[dict])
async def profile(user:Annotated[CurrentUser,Depends(get_current_user)],session:Annotated[AsyncSession,Depends(get_db_session)]):
 row=(await session.execute(text("SELECT id,full_name,job_title FROM public.profiles WHERE id=:u"),{"u":user.id})).mappings().first()
 return APIResponse(message="Profile retrieved.",data=dict(row) if row else {"id":str(user.id),"full_name":"","job_title":None})

@router.put("/profile",response_model=APIResponse[dict])
async def save_profile(req:PersonalProfileUpdate,user:Annotated[CurrentUser,Depends(get_current_user)],session:Annotated[AsyncSession,Depends(get_db_session)]):
 await session.execute(text("""INSERT INTO public.profiles(id,full_name,job_title,updated_at) VALUES(:u,:n,:j,now()) ON CONFLICT(id) DO UPDATE SET full_name=EXCLUDED.full_name,job_title=EXCLUDED.job_title,updated_at=now()"""),{"u":user.id,"n":req.full_name.strip(),"j":req.job_title})
 rows=(await session.execute(text("SELECT id,company_id FROM public.company_invitations WHERE accepted_by=:u AND status='accepted' FOR UPDATE"),{"u":user.id})).mappings().all()
 for x in rows:
  await session.execute(text("UPDATE public.company_members SET is_active=true,joined_at=now(),updated_at=now() WHERE company_id=:c AND user_id=:u AND is_active=false"),{"c":x["company_id"],"u":user.id})
  await session.execute(text("UPDATE public.company_invitations SET status='completed',completed_at=now(),updated_at=now() WHERE id=:i"),{"i":x["id"]})
 await session.commit()
 return APIResponse(message="Profile updated; accepted memberships activated.",data={"id":str(user.id),"full_name":req.full_name.strip(),"job_title":req.job_title,"activated_workspaces":len(rows)})

@router.patch("/members/{mid}/role",response_model=APIResponse[dict])
async def role(mid:UUID,req:MemberRoleUpdate,company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
 if req.role not in ALLOWED:raise ApplicationError(message="Unsupported role.",error_code="INVALID_COMPANY_ROLE",status_code=422)
 row=(await session.execute(text("SELECT role::text role FROM public.company_members WHERE id=:m AND company_id=:c AND is_active=true"),{"m":mid,"c":company.id})).mappings().first()
 if not row:raise ApplicationError(message="Member not found.",error_code="MEMBER_NOT_FOUND",status_code=404)
 if row["role"]=="owner":raise ApplicationError(message="Owner role is protected.",error_code="OWNER_ROLE_PROTECTED",status_code=409)
 await session.execute(text("UPDATE public.company_members SET role=CAST(:r AS public.company_role),updated_at=now() WHERE id=:m AND company_id=:c"),{"r":req.role,"m":mid,"c":company.id});await session.commit()
 return APIResponse(message="Role updated.",data={"id":str(mid),"role":req.role})

@router.post("/members/{mid}/deactivate",response_model=APIResponse[dict])
async def deactivate(mid:UUID,company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
 row=(await session.execute(text("SELECT role::text role FROM public.company_members WHERE id=:m AND company_id=:c AND is_active=true"),{"m":mid,"c":company.id})).mappings().first()
 if not row:raise ApplicationError(message="Member not found.",error_code="MEMBER_NOT_FOUND",status_code=404)
 if row["role"]=="owner":raise ApplicationError(message="Owner access is protected.",error_code="OWNER_ACCESS_PROTECTED",status_code=409)
 await session.execute(text("UPDATE public.company_members SET is_active=false,updated_at=now() WHERE id=:m AND company_id=:c"),{"m":mid,"c":company.id});await session.commit()
 return APIResponse(message="Access removed.",data={"id":str(mid)})
