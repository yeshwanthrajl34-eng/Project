from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from backend.ai_core.gemini_generator import GeminiDocumentGenerator


router = APIRouter()

generator = GeminiDocumentGenerator()


class DocumentRequest(BaseModel):
    document_type: str
    parties: str
    dates: str
    terms: str


@router.post("/generate")
def generate_document(request: DocumentRequest):

    try:
        document = generator.generate_document(
            document_type=request.document_type,
            parties=request.parties,
            dates=request.dates,
            terms=request.terms
        )

        return {
            "success": True,
            "document": document
        }

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )