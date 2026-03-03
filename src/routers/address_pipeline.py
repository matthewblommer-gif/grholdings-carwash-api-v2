import re

from fastapi import APIRouter, status, HTTPException
from fastapi.responses import Response

from src.clients.placer_client import PlacerClient
from src.core.logging import logger
from src.core.settings import get_settings
from src.models.address import AddressRequest, AnalyzeRequest, POISearchResponse, POIWithDistance
from src.services.analysis_orchestrator_service import AnalysisOrchestratorService
from src.services.car_parc_service import CarParcService
from src.services.competitor_service import CompetitorService
from src.services.excel_export_service import ExcelExportService
from src.services.google_location_service import GoogleLocationService
from src.services.key_stats_service import KeyStatsService
from src.services.retail_performance_service import RetailPerformanceService

router = APIRouter(prefix="/api/pipeline", tags=["pipeline"])

settings = get_settings()
location_service = GoogleLocationService(api_key=settings.google_api_key)

placer_client = PlacerClient(api_key=settings.placer_api_key)

car_parc_service = CarParcService(placer_client=placer_client)
competitor_service = CompetitorService(placer_client=placer_client, google_location_service=location_service)
key_stats_service = KeyStatsService(placer_client=placer_client)
retail_performance_service = RetailPerformanceService(placer_client=placer_client, google_location_service=location_service)

orchestrator = AnalysisOrchestratorService(
    car_parc_service=car_parc_service,
    competitor_service=competitor_service,
    key_stats_service=key_stats_service,
    retail_performance_service=retail_performance_service,
)

excel_service = ExcelExportService(location_service=location_service)


@router.post("/search-pois", status_code=status.HTTP_200_OK)
async def search_pois(request: AddressRequest) -> POISearchResponse:
    logger.info(f"Searching POIs for address: {request.address}")
    is_valid_address = location_service.verify_address_exists(request.address)

    if not is_valid_address:
        logger.warning(f"Invalid address submitted: {request.address}")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="The address provided could not be verified. Please check the address and try again.",
        )

    logger.info(f"Address verified successfully: {request.address}")

    coordinates = location_service.lookup_address(request.address)

    if not coordinates:
        logger.error(f"Could not get coordinates for address: {request.address}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Could not geocode the provided address.",
        )

    logger.info(f"Address geocoded: lat={coordinates.latitude}, lng={coordinates.longitude}")

    try:
        pois_with_distance = car_parc_service.search_pois_with_distance(coordinates.latitude, coordinates.longitude, location_service, 0.5)
        logger.info(f"Found {len(pois_with_distance)} POIs")

        if not pois_with_distance:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="No businesses found near the target location. Please try a different address.",
            )

        pois = [POIWithDistance(venue=venue, distance_miles=round(distance, 2)) for venue, distance in pois_with_distance]

        logger.info(f"Found {len(pois)} POIs for address: {request.address}")

        return POISearchResponse(
            latitude=coordinates.latitude,
            longitude=coordinates.longitude,
            formatted_address=coordinates.formatted_address,
            pois=pois,
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"POI search failed: {e}", exc_info=True)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"POI search failed: {str(e)}",
        )


@router.post("/analyze", status_code=status.HTTP_200_OK)
async def analyze_market(request: AnalyzeRequest) -> Response:
    logger.info(f"Starting market analysis for address: {request.address}, POI: {request.poi_name}")

    try:
        market_analysis = orchestrator.analyze_market(
            address=request.address,
            latitude=request.latitude,
            longitude=request.longitude,
            poi_id=request.poi_id,
            poi_name=request.poi_name,
            land_cost=request.land_cost,
            traffic_counts=request.traffic_counts,
        )

        logger.info(f"Market analysis completed successfully for {request.address}")

    except Exception as e:
        logger.error(f"Market analysis failed: {e}", exc_info=True)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Market analysis failed: {str(e)}",
        )

    try:
        excel_bytes = excel_service.export_market_analysis(market_analysis)
        logger.info("Excel export completed")

        clean_address = re.sub(r"[^a-zA-Z0-9\s-]", "", market_analysis.address)
        clean_address = re.sub(r"\s+", "_", clean_address.strip())
        excel_filename = f"{clean_address}.xlsx"

        logger.info(f"Excel file created: {excel_filename}")

        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename={excel_filename}",
                "X-Content-Type-Options": "nosniff",
                "Cache-Control": "no-cache, no-store, must-revalidate",
                "Pragma": "no-cache",
                "Expires": "0",
            },
        )

    except Exception as e:
        logger.error(f"Export failed: {e}", exc_info=True)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Export failed: {str(e)}",
        )
