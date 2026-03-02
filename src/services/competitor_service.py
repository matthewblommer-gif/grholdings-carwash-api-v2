from typing import Dict, List, Tuple

from shapely.geometry import shape

from src.clients.placer_client import PlacerClient
from src.core.date_utils import get_last_12_months_as_two_halves, get_last_12_months_date_range
from src.core.logging import logger
from src.models.competitor import Competitor
from src.models.placer.loyalty import LoyaltyFrequencyResponse
from src.models.placer.poi import PlacerPOIResponse, Venue
from src.models.placer.trade_area import TradeAreaResponse
from src.models.placer.visit_trends import VisitTrendsData, VisitTrendsResponse
from src.services.google_location_service import GoogleLocationService


class CompetitorService:

    # Business rule: Minimum yearly visits to qualify as competitor
    MIN_YEARLY_VISITS = 50000

    # Business rule: Loyalty threshold (minimum 6 visits in last 6 months)
    LOYALTY_MIN_VISITS = 6

    DEFAULT_SEARCH_RADIUS_MILES = 5.0
    MAX_DRIVE_TIME_MINUTES = 15.0

    DRIVE_TIME_FALLBACK = 999.0
    DISTANCE_FALLBACK = 999.0

    def __init__(self, placer_client: PlacerClient, google_location_service: GoogleLocationService) -> None:
        self.placer_client = placer_client
        self.google_location_service = google_location_service

    def search_car_wash_competitors(self, latitude: float, longitude: float, search_radius_miles: float = None) -> List[Venue]:
        if search_radius_miles is None:
            search_radius_miles = self.DEFAULT_SEARCH_RADIUS_MILES
        logger.info(f"Searching for car wash competitors within {search_radius_miles} miles")

        response_data = self.placer_client.search_poi(
            lat=latitude,
            lng=longitude,
            radius=search_radius_miles,
            category="Car Wash Services",
            sub_category="Car Wash",
            limit=50,
        )

        parsed_response = PlacerPOIResponse.model_validate(response_data)
        logger.info(f"Found {len(parsed_response.data)} car wash venues")

        return parsed_response.data

    def calculate_drive_time(self, origin_latitude: float, origin_longitude: float, destination_address: str) -> Tuple[float, float]:
        duration_minutes, distance_miles = self.google_location_service.calculate_drive_time_and_distance(origin_latitude, origin_longitude, destination_address)

        if duration_minutes is None or distance_miles is None:
            return self.DRIVE_TIME_FALLBACK, self.DISTANCE_FALLBACK

        return duration_minutes, distance_miles

    def determine_drive_time_band(self, drive_time_minutes: float) -> str:
        # Business rule: Drive time band classification
        if drive_time_minutes <= 10:
            return "0-10 min"
        elif drive_time_minutes <= 12:
            return "10-12 min"
        elif drive_time_minutes <= 15:
            return "12-15 min"
        else:
            return "15+ min"

    def filter_by_visit_threshold(self, venues: List[Venue], start_date: str, end_date: str) -> Tuple[List[Venue], Dict[str, int]]:
        logger.info(f"Filtering {len(venues)} venues by visit threshold (>= {self.MIN_YEARLY_VISITS})")

        api_ids = [venue.apiId for venue in venues]

        payload = {
            "apiIds": api_ids,
            "granularity": "month",
            "visitDurationSegmentation": "default",
            "startDate": start_date,
            "endDate": end_date,
        }

        response_data = self.placer_client.get_visit_trends(payload)
        parsed_response = VisitTrendsResponse.model_validate(response_data)

        qualified_venues = []
        visit_data = {}

        for item in parsed_response.data:
            if isinstance(item, VisitTrendsData):
                yearly_total = sum(item.visits)

                if yearly_total >= self.MIN_YEARLY_VISITS:
                    venue = next((v for v in venues if v.apiId == item.apiId), None)
                    if venue:
                        qualified_venues.append(venue)
                        visit_data[item.apiId] = yearly_total
                        logger.debug(f"{venue.name}: {yearly_total:,} visits/year - QUALIFIED")

        logger.info(f"Qualified {len(qualified_venues)} venues with >= {self.MIN_YEARLY_VISITS} visits/year")
        return qualified_venues, visit_data

    def _count_loyal_visitors(self, api_id: str, start_date: str, end_date: str) -> int:
        payload = {"apiId": api_id, "startDate": start_date, "endDate": end_date}

        response_data = self.placer_client.get_loyalty_frequency(payload)
        parsed_response = LoyaltyFrequencyResponse.model_validate(response_data)

        loyal_visitors = 0
        for i, bin_value in enumerate(parsed_response.data.bins):
            if bin_value >= self.LOYALTY_MIN_VISITS and i < len(parsed_response.data.visitors):
                loyal_visitors += parsed_response.data.visitors[i]

        return loyal_visitors

    def calculate_total_members(self, api_id: str, start_date: str, end_date: str) -> int:
        return self._count_loyal_visitors(api_id, start_date, end_date)

    def calculate_trade_area_overlap(
        self,
        reference_polygon,
        competitor_api_id: str,
        start_date: str,
        end_date: str,
    ) -> float:
        payload = {
            "apiId": competitor_api_id,
            "startDate": start_date,
            "endDate": end_date,
            "tradeAreaType": "70percentTrueTradeArea",
        }

        response_data = self.placer_client.get_trade_area(payload)
        parsed_response = TradeAreaResponse.model_validate(response_data)

        geojson = {
            "type": parsed_response.data.type,
            "coordinates": parsed_response.data.coordinates,
        }
        competitor_polygon = shape(geojson)

        if reference_polygon and competitor_polygon.is_valid and reference_polygon.is_valid:
            intersection = reference_polygon.intersection(competitor_polygon)
            competitor_area = competitor_polygon.area
            intersection_area = intersection.area
            overlap_pct = (intersection_area / competitor_area * 100) if competitor_area > 0 else 0
            return overlap_pct

        return 0.0

    def get_competitor_tta_demographics(
        self,
        api_id: str,
        start_date: str,
        end_date: str,
    ) -> Tuple[int, int]:
        payload = {
            "method": "tta",
            "benchmarkScope": "nationwide",
            "allocationType": "weightedCentroid",
            "trafficVolPct": 70,
            "withinRadius": 15,
            "ringRadius": 3,
            "dataset": "sti_popstats",
            "startDate": start_date,
            "endDate": end_date,
            "apiId": api_id,
        }

        demographics_data = self.placer_client.get_demographics(payload)

        if demographics_data is None:
            return 0, 0

        data = demographics_data.get("data", {})
        vehicles_data = data.get("Vehicles per Household", {})
        car_parc = vehicles_data.get("Total Number of Vehicles", {}).get("value", 0)

        overview_data = data.get("Overview", {})
        visits = overview_data.get("Visits", {}).get("value", 0)

        return int(car_parc), int(visits)

    def analyze_competitors(
        self,
        latitude: float,
        longitude: float,
        reference_poi_id: str,
        max_drive_time_minutes: float = None,
    ) -> List[Competitor]:
        if max_drive_time_minutes is None:
            max_drive_time_minutes = self.MAX_DRIVE_TIME_MINUTES

        logger.info("Starting competitor analysis")

        start_date, end_date = get_last_12_months_date_range()
        _, second_half_dates = get_last_12_months_as_two_halves()

        all_venues = self.search_car_wash_competitors(latitude, longitude)

        venues_with_drive_times = []
        drive_time_data = {}

        logger.info("Calculating drive times for competitors")
        for venue in all_venues:
            duration_minutes, distance_miles = self.calculate_drive_time(latitude, longitude, venue.address.formattedAddress)

            if duration_minutes <= max_drive_time_minutes:
                venues_with_drive_times.append(venue)
                drive_time_data[venue.apiId] = {
                    "duration_minutes": duration_minutes,
                    "distance_miles": distance_miles,
                }
                logger.debug(f"{venue.name}: {duration_minutes:.1f} min, {distance_miles:.2f} mi")

        logger.info(f"Found {len(venues_with_drive_times)} venues within {max_drive_time_minutes} min")

        qualified_venues, visit_data = self.filter_by_visit_threshold(venues_with_drive_times, start_date, end_date)

        logger.info("Fetching reference POI trade area")
        ref_payload = {
            "apiId": reference_poi_id,
            "startDate": start_date,
            "endDate": end_date,
            "tradeAreaType": "70percentTrueTradeArea",
        }
        ref_response_data = self.placer_client.get_trade_area(ref_payload)
        ref_parsed = TradeAreaResponse.model_validate(ref_response_data)
        ref_geojson = {"type": ref_parsed.data.type, "coordinates": ref_parsed.data.coordinates}
        reference_polygon = shape(ref_geojson)

        competitors = []
        logger.info(f"Building competitor models for {len(qualified_venues)} venues")

        for venue in qualified_venues:
            drive_info = drive_time_data[venue.apiId]
            yearly_visits = visit_data[venue.apiId]

            total_members = self.calculate_total_members(venue.apiId, second_half_dates[0], second_half_dates[1])
            overlap_pct = self.calculate_trade_area_overlap(reference_polygon, venue.apiId, start_date, end_date)
            car_parc, tta_visits = self.get_competitor_tta_demographics(venue.apiId, start_date, end_date)

            members_in_market = int((overlap_pct / 100) * total_members)

            drive_time_band = self.determine_drive_time_band(drive_info["duration_minutes"])

            competitor = Competitor(
                name=venue.name,
                api_id=venue.apiId,
                address=venue.address.formattedAddress,
                drive_time_minutes=drive_info["duration_minutes"],
                drive_time_band=drive_time_band,
                distance_miles=drive_info["distance_miles"],
                total_members=total_members,
                overlap_percentage=overlap_pct,
                members_in_market=members_in_market,
                visits_per_year=yearly_visits,
                visits_per_month=yearly_visits // 12,
                visits_per_day=yearly_visits // 365,
                car_parc=car_parc,
                tta_visits=tta_visits,
            )

            competitors.append(competitor)
            logger.debug(
                f"{competitor.name}: {competitor.total_members:,} members, " f"{competitor.overlap_percentage:.1f}% overlap, " f"{competitor.members_in_market:,} in market"
            )

        logger.info(f"Competitor analysis complete: {len(competitors)} competitors")
        return competitors
