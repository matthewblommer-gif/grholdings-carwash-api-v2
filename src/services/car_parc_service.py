from typing import List, Optional, Tuple

import numpy as np

from src.clients.placer_client import PlacerClient
from src.core.date_utils import get_last_12_months_date_range
from src.core.logging import logger
from src.models.car_parc_analysis import CarParcAnalysis
from src.models.car_parc_result import CarParcResult
from src.models.placer.poi import PlacerPOIResponse, Venue


class CarParcService:

    # Business rules: TAM and market share percentages by drive time
    KNOWN_DRIVE_TIMES = np.array([10, 12, 15])
    KNOWN_TAM_PERCENTAGES = np.array([30, 25, 20])
    KNOWN_MARKET_SHARE_PERCENTAGES = np.array([95, 85, 65])

    POI_SEARCH_RADIUS_MILES = 2.5
    POI_SEARCH_LIMIT = 20

    BENCHMARK_SCOPE = "nationwide"
    ALLOCATION_TYPE = "weightedCentroid"
    TRAFFIC_VOL_PCT = 70
    RING_RADIUS = 3
    DATASET = "sti_popstats"
    TEMPLATE = "default"

    MIN_TAM_PERCENTAGE = 10
    MAX_TAM_PERCENTAGE = 40

    MIN_MARKET_SHARE_PERCENTAGE = 40
    MAX_MARKET_SHARE_PERCENTAGE = 100

    def __init__(self, placer_client: PlacerClient) -> None:
        self.placer_client = placer_client

    def search_pois(self, latitude: float, longitude: float, search_radius: Optional[float]) -> List[Venue]:
        if search_radius is None:
            search_radius = self.POI_SEARCH_RADIUS_MILES

        logger.info(f"Searching POIs near lat={latitude}, lng={longitude} search radius={search_radius}")

        response_data = self.placer_client.search_poi(lat=latitude, lng=longitude, radius=search_radius, limit=self.POI_SEARCH_LIMIT)

        parsed_response = PlacerPOIResponse.model_validate(response_data)

        if not parsed_response.data:
            logger.warning("No POIs found near target location")
            return []

        logger.info(f"Found {len(parsed_response.data)} POIs near target location")
        return parsed_response.data

    def search_pois_with_distance(
        self,
        latitude: float,
        longitude: float,
        google_location_service,
        search_radius: Optional[float] = None,
    ) -> List[Tuple[Venue, float]]:
        DISTANCE_FALLBACK = 999.0

        venues = self.search_pois(latitude, longitude, search_radius)

        venues_with_distance = []
        for venue in venues:
            distance_miles = google_location_service.calculate_distance(latitude, longitude, venue.address.formattedAddress)

            if distance_miles is None:
                distance_miles = DISTANCE_FALLBACK

            venues_with_distance.append((venue, distance_miles))

        venues_with_distance = [(v, d) for v, d in venues_with_distance if d < DISTANCE_FALLBACK]
        venues_with_distance.sort(key=lambda x: x[1])

        logger.info(f"Sorted {len(venues_with_distance)} POIs by distance")
        return venues_with_distance

    def calculate_tam_percentage(self, drive_time_minutes: int) -> float:
        # Business rule: Interpolate TAM percentage based on drive time
        tam_pct = np.interp(drive_time_minutes, self.KNOWN_DRIVE_TIMES, self.KNOWN_TAM_PERCENTAGES)
        return float(np.clip(tam_pct, self.MIN_TAM_PERCENTAGE, self.MAX_TAM_PERCENTAGE) / 100)

    def calculate_market_share_percentage(self, drive_time_minutes: int) -> float:
        # Business rule: Interpolate market share percentage based on drive time
        market_share_pct = np.interp(
            drive_time_minutes,
            self.KNOWN_DRIVE_TIMES,
            self.KNOWN_MARKET_SHARE_PERCENTAGES,
        )
        return float(np.clip(market_share_pct, self.MIN_MARKET_SHARE_PERCENTAGE, self.MAX_MARKET_SHARE_PERCENTAGE) / 100)

    def get_car_parc_for_drive_time(self, reference_poi_id: str, drive_time_minutes: int) -> CarParcResult:
        logger.info(f"Fetching car parc data for {drive_time_minutes} min drive time")

        start_date, end_date = get_last_12_months_date_range()

        payload = {
            "method": "driveTime",
            "benchmarkScope": self.BENCHMARK_SCOPE,
            "allocationType": self.ALLOCATION_TYPE,
            "trafficVolPct": self.TRAFFIC_VOL_PCT,
            "withinRadius": drive_time_minutes,
            "ringRadius": self.RING_RADIUS,
            "dataset": self.DATASET,
            "startDate": start_date,
            "endDate": end_date,
            "apiId": reference_poi_id,
            "driveTime": drive_time_minutes,
            "template": self.TEMPLATE,
        }

        demographics_data = self.placer_client.get_demographics(payload)

        if demographics_data is None:
            logger.info(f"No demographics data available for {drive_time_minutes} min drive time - using zeros")
            return CarParcResult(
                drive_time_minutes=drive_time_minutes,
                car_parc=0,
                population=0,
                households=0,
                tam_percentage=self.calculate_tam_percentage(drive_time_minutes),
                total_addressable_market=0,
                market_share_percentage=self.calculate_market_share_percentage(drive_time_minutes),
            )

        data = demographics_data.get("data", {})
        vehicles_data = data.get("Vehicles per Household", {})
        overview_data = data.get("Overview", {})

        car_parc = vehicles_data.get("Total Number of Vehicles", {}).get("value", 0)
        population = overview_data.get("Population", {}).get("value", 0)
        households = overview_data.get("Households", {}).get("value", 0)

        tam_percentage = self.calculate_tam_percentage(drive_time_minutes)
        total_addressable_market = int(car_parc * tam_percentage)

        market_share_percentage = self.calculate_market_share_percentage(drive_time_minutes)

        logger.info(f"Drive time {drive_time_minutes} min: Car Parc={car_parc:,}, TAM={total_addressable_market:,}")

        return CarParcResult(
            drive_time_minutes=drive_time_minutes,
            car_parc=int(car_parc),
            population=int(population),
            households=int(households),
            tam_percentage=tam_percentage,
            total_addressable_market=total_addressable_market,
            market_share_percentage=market_share_percentage,
        )

    def analyze_car_parc(self, reference_poi_id: str, reference_poi_name: str, drive_times: List[int]) -> CarParcAnalysis:
        logger.info(f"Analyzing car parc for POI: {reference_poi_name} ({reference_poi_id})")

        results = []
        for drive_time in drive_times:
            result = self.get_car_parc_for_drive_time(reference_poi_id, drive_time)
            results.append(result)

        return CarParcAnalysis(
            results=results,
            reference_poi_id=reference_poi_id,
            reference_poi_name=reference_poi_name,
        )
