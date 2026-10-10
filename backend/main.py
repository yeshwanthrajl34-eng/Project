from fastapi import FastAPI
from backend.routes import router

app = FastAPI(
    title="LegalEase API",
    description="AI-powered legal document generation API",
    version="1.0.0"
)

app.include_router(router)


@app.get("/")
def home():
    return {
        "message": "Welcome to LegalEase API!",
        "status": "running"
    }