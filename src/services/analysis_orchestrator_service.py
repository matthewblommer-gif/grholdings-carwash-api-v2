from typing import List, Optional

from src.core.logging import logger
from src.models.market_analysis import MarketAnalysis
from src.services.car_parc_service import CarParcService
from src.services.competitor_service import CompetitorService
from src.services.key_stats_service import KeyStatsService
from src.services.retail_performance_service import RetailPerformanceService


class AnalysisOrchestratorService:

    # Business rule: Default drive times to analyze (5-25 minutes in 1-minute increments)
    DEFAULT_DRIVE_TIMES = list(range(5, 26))

    def __init__(
        self,
        car_parc_service: CarParcService,
        competitor_service: CompetitorService,
        key_stats_service: KeyStatsService,
        retail_performance_service: RetailPerformanceService,
    ) -> None:
        self.car_parc_service = car_parc_service
        self.competitor_service = competitor_service
        self.key_stats_service = key_stats_service
        self.retail_performance_service = retail_performance_service

    def analyze_market(
        self,
        address: str,
        latitude: float,
        longitude: float,
        poi_id: str,
        poi_name: str,
        drive_times: Optional[List[int]] = None,
        land_cost: Optional[int] = None,
        traffic_counts: Optional[int] = None,
    ) -> MarketAnalysis:
        logger.info(f"Starting full market analysis for address: {address}, POI: {poi_name}")

        if drive_times is None:
            drive_times = self.DEFAULT_DRIVE_TIMES

        logger.info("Step 1: Analyzing car parc data")
        car_parc_analysis = self.car_parc_service.analyze_car_parc(poi_id, poi_name, drive_times)

        logger.info("Step 2: Analyzing competitors")
        competitors = self.competitor_service.analyze_competitors(latitude, longitude, poi_id)

        logger.info("Step 3: Fetching key stats")
        key_stats = self.key_stats_service.get_key_stats(car_parc_analysis.reference_poi_id)

        logger.info("Step 4: Analyzing retail performance")
        retailers, reference_poi_retail = self.retail_performance_service.analyze_retail_performance(latitude, longitude, poi_id, poi_name)

        logger.info("Step 5: Calculating market metrics")

        # Business formula: Total Market Members = sum of all competitors' members in market
        total_market_members = sum(c.members_in_market for c in competitors)

        logger.info(f"Market analysis complete: {len(competitors)} competitors, " f"{total_market_members:,} total market members")

        return MarketAnalysis(
            address=address,
            latitude=latitude,
            longitude=longitude,
            reference_poi_id=car_parc_analysis.reference_poi_id,
            reference_poi_name=car_parc_analysis.reference_poi_name,
            car_parc_results=car_parc_analysis.results,
            competitors=competitors,
            key_stats=key_stats,
            retailers=retailers,
            reference_poi_retail=reference_poi_retail,
            total_market_members=total_market_members,
            land_cost=land_cost,
            traffic_counts=traffic_counts,
        )
