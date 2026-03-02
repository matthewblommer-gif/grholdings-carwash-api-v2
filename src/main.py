from typing import Any, Coroutine

from fastapi import FastAPI, APIRouter
from fastapi.responses import JSONResponse, HTMLResponse
from pathlib import Path
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request

from src.core.request_tracker import init_request_counts
from src.routers import address_pipeline


class RequestTrackingMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        counts = init_request_counts()
        response = await call_next(request)
        if request.url.path.startswith("/api/pipeline"):
            counts.log_summary(endpoint=request.url.path)
        return response


app = FastAPI(title="Washville API", description="API for carwash management", version="1.0.0")
app.add_middleware(RequestTrackingMiddleware)

router = APIRouter(prefix="/api", tags=["health"])


@router.get("/healthcheck")
async def healthcheck() -> JSONResponse:
    return JSONResponse(status_code=200, content={"status": "healthy", "service": "grholdings-carwash-api"})


@router.get("/debug-keys")
async def debug_keys() -> JSONResponse:
    """Temporary endpoint to verify API keys are loaded. Remove after debugging."""
    from src.core.settings import get_settings
    s = get_settings()
    return JSONResponse(status_code=200, content={
        "placer_key_length": len(s.placer_api_key) if s.placer_api_key else 0,
        "placer_key_prefix": s.placer_api_key[:6] + "..." if s.placer_api_key and len(s.placer_api_key) > 6 else "MISSING",
        "google_key_length": len(s.google_api_key) if s.google_api_key else 0,
        "google_key_prefix": s.google_api_key[:6] + "..." if s.google_api_key and len(s.google_api_key) > 6 else "MISSING",
    })


@router.get("/debug-placer")
async def debug_placer() -> JSONResponse:
    """Temporary endpoint to test Placer API connectivity. Remove after debugging."""
    import requests as req
    from src.core.settings import get_settings
    s = get_settings()
    headers = {"accept": "application/json", "x-api-key": s.placer_api_key}
    try:
        resp = req.get("https://papi.placer.ai/v1/poi", params={
            "lat": 40.7128, "lng": -74.0060, "radius": 1.0,
            "entityType": "venue", "limit": 1, "category": "Car Wash Services",
        }, headers=headers)
        return JSONResponse(status_code=200, content={
            "placer_status_code": resp.status_code,
            "placer_response_body": resp.text[:500],
            "request_headers_sent": {"x-api-key": s.placer_api_key[:6] + "...(redacted)"},
        })
    except Exception as e:
        return JSONResponse(status_code=200, content={"error": str(e)})


app.include_router(router)
app.include_router(address_pipeline.router)


@app.get("/", response_class=HTMLResponse)
async def root() -> str:
    dashboard_path = Path(__file__).parent.parent / "dashboard" / "index.html"
    return dashboard_path.read_text()


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
