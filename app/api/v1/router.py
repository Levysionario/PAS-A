from fastapi import APIRouter
from app.api.v1.endpoints import uploads, pdf

api_router = APIRouter()
api_router.include_router(uploads.router, prefix="/uploads", tags=["1. Upload & Parsing (.docx)"])
api_router.include_router(pdf.router, prefix="/pdf", tags=["2. Motor de Diagramação (PDF)"])