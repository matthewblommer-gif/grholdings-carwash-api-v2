from fastapi import FastAPI, APIRouter
from fastapi.responses import JSONResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
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


app.include_router(router)
app.include_router(address_pipeline.router)

# Serve static assets (logo, etc.)
resources_path = Path(__file__).parent.parent / "resources"
app.mount("/static", StaticFiles(directory=str(resources_path)), name="static")


@app.get("/", response_class=HTMLResponse)
async def root() -> str:
    dashboard_path = Path(__file__).parent.parent / "dashboard" / "index.html"
    return dashboard_path.read_text()


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
