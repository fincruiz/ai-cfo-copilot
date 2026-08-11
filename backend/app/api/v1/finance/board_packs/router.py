from pathlib import Path
from typing import Annotated
from uuid import UUID
from fastapi import APIRouter,Depends,HTTPException
from fastapi.responses import FileResponse
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.schemas.responses import APIResponse
from app.schemas.finance.advanced_forecasting import BoardPackGenerateRequest,ArtifactResponse
from app.services.finance.board_pack_service import BoardPackService
router=APIRouter(prefix='/board-packs',tags=['Board Packs'])
def svc(session:Annotated[AsyncSession,Depends(get_db_session)]):return BoardPackService(session)
@router.post('/generate',response_model=APIResponse[list[ArtifactResponse]])
async def generate(request:BoardPackGenerateRequest,current_company:Annotated[Company,Depends(get_current_company)],service:Annotated[BoardPackService,Depends(svc)]):return APIResponse(message='Board pack generated.',data=[ArtifactResponse(**x) for x in await service.generate(current_company,request)])
@router.get('/artifacts')
async def list_artifacts(current_company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)]):
 rows=(await session.execute(text('SELECT id,artifact_type,file_name,file_size_bytes,created_at FROM public.generated_artifacts WHERE company_id=:c ORDER BY created_at DESC'),{'c':current_company.id})).mappings().all();return APIResponse(message='Artifacts retrieved.',data=[{**dict(r),'download_url':f"/api/v1/board-packs/artifacts/{r['id']}/download"} for r in rows])
@router.get('/artifacts/{artifact_id}/download')
async def download(artifact_id:UUID,current_company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)]):
 row=(await session.execute(text('SELECT artifact_type,file_name,storage_path FROM public.generated_artifacts WHERE company_id=:c AND id=:i'),{'c':current_company.id,'i':artifact_id})).mappings().first()
 if not row or not Path(row['storage_path']).exists():raise HTTPException(404,'Artifact not found')
 media={'pptx':'application/vnd.openxmlformats-officedocument.presentationml.presentation','pdf':'application/pdf','xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}[row['artifact_type']]
 return FileResponse(row['storage_path'],filename=row['file_name'],media_type=media)
