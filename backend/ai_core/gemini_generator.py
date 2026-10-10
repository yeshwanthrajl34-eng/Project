import os
import time

from dotenv import load_dotenv
from google import genai


load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    raise ValueError("GEMINI_API_KEY is missing from the .env file")

client = genai.Client(api_key=API_KEY)


class GeminiDocumentGenerator:

    def generate_document(
        self,
        document_type: str,
        parties: str,
        dates: str,
        terms: str
    ):

        prompt = f"""
You are LegalEase, an AI-powered legal document drafting assistant.

Create a professional draft of the following legal document.

DOCUMENT TYPE:
{document_type}

PARTIES:
{parties}

DATES:
{dates}

TERMS AND CONDITIONS:
{terms}

INSTRUCTIONS:

1. Create a clear professional legal document.
2. Use appropriate headings and sections.
3. Include the information provided by the user.
4. Do not invent important personal information.
5. Use formal legal language.
6. Make the document easy to read and edit.
7. Include relevant clauses for the selected document type.
8. Return ONLY the document itself.
9. Do not explain what you are doing.
10. Do not include markdown code fences.

This is an AI-generated draft and should be reviewed by a qualified legal professional before use.
"""

        # Try the Gemini request up to 4 times
        # if Google's service temporarily returns 503.

        for attempt in range(4):

            try:

                response = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=prompt
                )

                return response.text

            except Exception as e:

                error_message = str(e)

                if "503" in error_message or "UNAVAILABLE" in error_message:

                    if attempt < 3:

                        wait_time = 2 ** attempt

                        print(
                            f"Gemini temporarily unavailable. "
                            f"Retrying in {wait_time} seconds..."
                        )

                        time.sleep(wait_time)

                        continue

                raise e