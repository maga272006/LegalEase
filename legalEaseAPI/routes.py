from fastapi import APIRouter
from pydantic import BaseModel

from ai_core.gemini_generator import generate_legal_document


router = APIRouter()


class DocumentRequest(BaseModel):

    document_type: str
    parties: str
    terms: str
    dates: str


@router.post("/generate")
def generate_document(request: DocumentRequest):

    try:

        print("====================================")
        print("GENERATE DOCUMENT REQUEST")
        print("Document Type:", request.document_type)
        print("Parties:", request.parties)
        print("Terms:", request.terms)
        print("Dates:", request.dates)
        print("====================================")

        generated_document = generate_legal_document(

            document_type=request.document_type,

            parties=request.parties,

            terms=request.terms,

            dates=request.dates
        )

        # -----------------------------------------------
        # Check Gemini response
        # -----------------------------------------------

        if generated_document is None:

            return {
                "success": False,
                "error": "Gemini returned no document."
            }

        if not str(generated_document).strip():

            return {
                "success": False,
                "error": "Gemini returned an empty document."
            }

        print("Document generated successfully.")
        print("Document length:", len(generated_document))

        return {

            "success": True,

            "document_type": request.document_type,

            "document": generated_document
        }

    except Exception as e:

        print("====================================")
        print("GENERATE DOCUMENT ERROR")
        print("Error Type:", type(e).__name__)
        print("Error:", str(e))
        print("====================================")

        return {

            "success": False,

            "error": str(e)
        }