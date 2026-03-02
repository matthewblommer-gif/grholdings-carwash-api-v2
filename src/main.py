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


@router.get("/debug-post")
async def debug_post() -> JSONResponse:
    """Test POST to the exact endpoint that's failing. Remove after debugging."""
    import requests as req
    from src.core.settings import get_settings
    s = get_settings()
    headers = {"accept": "application/json", "content-type": "application/json", "x-api-key": s.placer_api_key}
    payload = {
        "method": "driveTime",
        "benchmarkScope": "nationwide",
        "allocationType": "weightedCentroid",
        "trafficVolPct": 70,
        "withinRadius": 10,
        "ringRadius": 3,
        "dataset": "sti_popstats",
        "startDate": "2025-01-01",
        "endDate": "2025-12-31",
        "apiId": "store_id_not_real",
        "driveTime": 10,
        "template": "default",
    }
    try:
        resp = req.post("https://papi.placer.ai/v1/reports/trade-area-demographics", json=payload, headers=headers)
        return JSONResponse(status_code=200, content={
            "status_code": resp.status_code,
            "response_headers": dict(resp.headers),
            "response_body": resp.text[:1000],
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
