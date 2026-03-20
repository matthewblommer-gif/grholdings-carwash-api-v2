import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Optional

from src.core.logging import logger
from src.models.car_parc_result import CarParcResult
from src.models.competitor import Competitor
from src.models.key_stats import KeyStats
from src.models.market_analysis import MarketAnalysis
from src.models.retailer import Retailer
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
        start_time = time.time()
        logger.info(f"Starting parallel market analysis for address: {address}, POI: {poi_name}")

        if drive_times is None:
            drive_times = self.DEFAULT_DRIVE_TIMES

        warnings: List[str] = []

        # Defaults in case a step fails
        car_parc_results: List[CarParcResult] = []
        competitors: List[Competitor] = []
        key_stats = KeyStats(car_parc=0, median_income=0, median_age=0.0, population=0, households=0)
        retailers: List[Retailer] = []
        reference_poi_retail: Optional[Retailer] = None

        def run_car_parc():
            logger.info("Step 1: Analyzing car parc data")
            return self.car_parc_service.analyze_car_parc(poi_id, poi_name, drive_times)

        def run_competitors():
            logger.info("Step 2: Analyzing competitors")
            return self.competitor_service.analyze_competitors(latitude, longitude, poi_id)

        def run_key_stats():
            logger.info("Step 3: Fetching key stats")
            return self.key_stats_service.get_key_stats(poi_id)

        def run_retail():
            logger.info("Step 4: Analyzing retail performance")
            return self.retail_performance_service.analyze_retail_performance(latitude, longitude, poi_id, poi_name)

        with ThreadPoolExecutor(max_workers=4) as executor:
            future_car_parc = executor.submit(run_car_parc)
            future_competitors = executor.submit(run_competitors)
            future_key_stats = executor.submit(run_key_stats)
            future_retail = executor.submit(run_retail)

            # Collect car parc results
            try:
                car_parc_analysis = future_car_parc.result()
                car_parc_results = car_parc_analysis.results
            except Exception as e:
                logger.error(f"Car parc analysis failed: {e}", exc_info=True)
                warnings.append(f"CAR PARC DATA UNAVAILABLE: {e}")

            # Collect competitor results
            try:
                competitors, comp_warnings = future_competitors.result()
                warnings.extend(comp_warnings)
            except Exception as e:
                logger.error(f"Competitor analysis failed: {e}", exc_info=True)
                warnings.append(f"COMPETITOR DATA UNAVAILABLE: {e}")

            # Collect key stats results
            try:
                key_stats = future_key_stats.result()
            except Exception as e:
                logger.error(f"Key stats failed: {e}", exc_info=True)
                warnings.append(f"KEY STATS DATA UNAVAILABLE: {e}")

            # Collect retail results
            try:
                retailers, reference_poi_retail = future_retail.result()
            except Exception as e:
                logger.error(f"Retail performance analysis failed: {e}", exc_info=True)
                warnings.append(f"RETAIL PERFORMANCE DATA UNAVAILABLE: {e}")

        # Business formula: Total Market Members = sum of all competitors' members in market
        total_market_members = sum(c.members_in_market for c in competitors)

        elapsed = time.time() - start_time
        logger.info(
            f"Market analysis complete in {elapsed:.1f}s: {len(competitors)} competitors, "
            f"{total_market_members:,} total market members, {len(warnings)} warnings"
        )

        return MarketAnalysis(
            address=address,
            latitude=latitude,
            longitude=longitude,
            reference_poi_id=poi_id,
            reference_poi_name=poi_name,
            car_parc_results=car_parc_results,
            competitors=competitors,
            key_stats=key_stats,
            retailers=retailers,
            reference_poi_retail=reference_poi_retail,
            total_market_members=total_market_members,
            land_cost=land_cost,
            traffic_counts=traffic_counts,
            warnings=warnings,
        )
