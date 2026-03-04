from typing import Dict, List, Optional, Tuple

from src.clients.placer_client import PlacerClient
from src.core.date_utils import get_last_12_months_date_range
from src.core.logging import logger
from src.models.placer.poi import PlacerPOIResponse, Venue
from src.models.placer.ranking import RankingData, RankingResponse
from src.models.placer.visit_trends import VisitTrendsData, VisitTrendsResponse
from src.models.retailer import Retailer
from src.services.google_location_service import GoogleLocationService


class RetailPerformanceService:

    RETAIL_CATEGORIES = [
        {"category": "Breakfast, Coffee, Bakeries & Dessert Shops", "subCategory": None},
        {"category": "Fast Food & QSR", "subCategory": None},
        {"category": "Electronics Stores", "subCategory": None},
        {"category": "Groceries", "subCategory": None},
        {"category": "Home Improvement", "subCategory": None},
        {"category": "Office Supplies", "subCategory": None},
        {"category": "Drugstores & Pharmacies", "subCategory": None},
        {"category": "Superstores", "subCategory": None},
        {"category": "Furniture and Home Furnishings", "subCategory": None},
    ]

    DEFAULT_SEARCH_RADIUS_MILES = 2.0
    POI_SEARCH_LIMIT = 50

    DISTANCE_FALLBACK = 999.0

    def __init__(self, placer_client: PlacerClient, google_location_service: GoogleLocationService) -> None:
        self.placer_client = placer_client
        self.google_location_service = google_location_service

    def search_retail_venues(self, latitude: float, longitude: float, search_radius_miles: float = None) -> List[Venue]:
        if search_radius_miles is None:
            search_radius_miles = self.DEFAULT_SEARCH_RADIUS_MILES

        logger.info(f"Searching for retail venues within {search_radius_miles} miles")

        all_venues = []
        venue_ids = set()

        for criteria in self.RETAIL_CATEGORIES:
            category = criteria["category"]
            sub_category = criteria["subCategory"]

            logger.debug(f"Searching category: {category}")

            response_data = self.placer_client.search_poi(
                lat=latitude,
                lng=longitude,
                radius=search_radius_miles,
                category=category,
                sub_category=sub_category,
                limit=self.POI_SEARCH_LIMIT,
            )

            parsed_response = PlacerPOIResponse.model_validate(response_data)

            new_venues = 0
            for venue in parsed_response.data:
                if venue.apiId not in venue_ids:
                    all_venues.append(venue)
                    venue_ids.add(venue.apiId)
                    new_venues += 1

            logger.debug(f"  {category}: Found {new_venues} new venues")

        logger.info(f"Found {len(all_venues)} total unique retail venues")
        return all_venues

    def calculate_distance(self, origin_latitude: float, origin_longitude: float, destination_address: str) -> float:
        distance_miles = self.google_location_service.calculate_distance(origin_latitude, origin_longitude, destination_address)

        if distance_miles is None:
            return self.DISTANCE_FALLBACK

        return distance_miles

    def get_rankings(self, venues: List[Venue]) -> Dict[str, RankingData]:
        logger.info(f"Fetching rankings for {len(venues)} retail venues")

        start_date, end_date = get_last_12_months_date_range()
        rankings_dict = {}

        for venue in venues:
            payload = {
                "apiId": venue.apiId,
                "startDate": start_date,
                "endDate": end_date,
                "distanceMiles": 15,
                "scope": "chain",
                "metric": "visits",
            }

            try:
                response_data = self.placer_client.get_ranking_single(payload)
                parsed_response = RankingResponse.model_validate(response_data)

                if parsed_response.data:
                    rankings_dict[venue.apiId] = parsed_response.data[0]
            except Exception as e:
                logger.warning(f"Failed to get ranking for venue {venue.apiId}: {e}")

        logger.info(f"Retrieved rankings for {len(rankings_dict)} venues")
        return rankings_dict

    def get_visit_trends(self, venues: List[Venue]) -> Dict[str, int]:
        logger.info(f"Fetching visit trends for {len(venues)} retail venues")

        if not venues:
            logger.info("No venues to fetch visit trends for, returning empty results")
            return {}

        api_ids = [venue.apiId for venue in venues]

        start_date, end_date = get_last_12_months_date_range()

        payload = {
            "apiIds": api_ids,
            "granularity": "month",
            "visitDurationSegmentation": "allVisits",
            "startDate": start_date,
            "endDate": end_date,
        }

        response_data = self.placer_client.get_visit_trends(payload)
        parsed_response = VisitTrendsResponse.model_validate(response_data)

        visits_dict = {}
        for item in parsed_response.data:
            if isinstance(item, VisitTrendsData):
                yearly_total = sum(item.visits)
                visits_dict[item.apiId] = yearly_total

        logger.info(f"Retrieved visit trends for {len(visits_dict)} venues")
        return visits_dict

    def get_poi_retail_stats(self, poi_id: str, poi_name: str) -> Optional[Retailer]:
        logger.info(f"{poi_id}: Fetching retail stats for reference POI: {poi_name}")

        start_date, end_date = get_last_12_months_date_range()

        ranking_payload = {
            "apiId": poi_id,
            "startDate": start_date,
            "endDate": end_date,
            "distanceMiles": 15,
            "scope": "chain",
            "metric": "visits",
        }

        response_data = self.placer_client.get_ranking_single(ranking_payload)
        parsed_response = RankingResponse.model_validate(response_data)

        if not parsed_response.data:
            logger.warning(f"{poi_id}: No ranking data returned for POI")
            return None

        ranking_data = parsed_response.data[0]

        visits_payload = {
            "apiIds": [poi_id],
            "granularity": "month",
            "visitDurationSegmentation": "allVisits",
            "startDate": start_date,
            "endDate": end_date,
        }

        visits_response = self.placer_client.get_visit_trends(visits_payload)
        parsed_visits = VisitTrendsResponse.model_validate(visits_response)

        visits = None
        for item in parsed_visits.data:
            if isinstance(item, VisitTrendsData) and item.apiId == poi_id:
                visits = sum(item.visits)
                break

        national_percentile = None
        state_percentile = None

        if not ranking_data.ranking.rankError:
            if ranking_data.ranking.nationwide:
                national_percentile = ranking_data.ranking.nationwide.percentile / 100
            if ranking_data.ranking.state:
                state_percentile = ranking_data.ranking.state.percentile / 100

        poi_retailer = Retailer(
            name=poi_name,
            api_id=poi_id,
            category=ranking_data.info.categoryInfo.category,
            sub_category=ranking_data.info.categoryInfo.subCategory,
            address="",
            distance_miles=0.0,
            national_percentile=national_percentile,
            state_percentile=state_percentile,
            visits=visits,
        )

        logger.info(f"{poi_id}: POI retail stats - National: {national_percentile}, State: {state_percentile}, Visits: {visits}")
        return poi_retailer

    def analyze_retail_performance(self, latitude: float, longitude: float, poi_id: str, poi_name: str) -> Tuple[List[Retailer], Optional[Retailer]]:
        logger.info("Starting retail performance analysis")

        poi_retailer = self.get_poi_retail_stats(poi_id, poi_name)

        all_venues = self.search_retail_venues(latitude, longitude)

        logger.info("Calculating distances for retail venues")
        venues_with_distances = []
        distance_data = {}

        for venue in all_venues:
            if venue.apiId == poi_id:
                logger.debug(f"{venue.apiId}: Excluding reference POI from retail list")
                continue

            distance_miles = self.calculate_distance(latitude, longitude, venue.address.formattedAddress)

            if distance_miles < self.DISTANCE_FALLBACK:
                venues_with_distances.append(venue)
                distance_data[venue.apiId] = distance_miles
                logger.debug(f"{venue.name}: {distance_miles:.2f} mi")

        logger.info(f"Found {len(venues_with_distances)} venues with valid distances (excluding POI)")

        if len(venues_with_distances) > 100:
            venues_with_distances = venues_with_distances[:100]
            logger.info(f"Limited to first 100 venues")

        rankings_dict = self.get_rankings(venues_with_distances)
        visits_dict = self.get_visit_trends(venues_with_distances)

        retailers = []
        for venue in venues_with_distances:
            distance_miles = distance_data[venue.apiId]
            ranking_data = rankings_dict.get(venue.apiId)
            visits = visits_dict.get(venue.apiId)

            national_percentile = None
            state_percentile = None

            if ranking_data and not ranking_data.ranking.rankError:
                if ranking_data.ranking.nationwide:
                    national_percentile = ranking_data.ranking.nationwide.percentile / 100
                if ranking_data.ranking.state:
                    state_percentile = ranking_data.ranking.state.percentile / 100

            retailer = Retailer(
                name=venue.name,
                api_id=venue.apiId,
                category=venue.categoryInfo.category,
                sub_category=venue.categoryInfo.subCategory,
                address=venue.address.formattedAddress,
                distance_miles=distance_miles,
                national_percentile=national_percentile,
                state_percentile=state_percentile,
                visits=visits,
            )

            retailers.append(retailer)
            logger.debug(f"{retailer.name}: {retailer.distance_miles:.2f} mi, National: {national_percentile}, State: {state_percentile}, Visits: {visits}")

        filtered_retailers = [r for r in retailers if r.distance_miles <= 2.0 and r.national_percentile is not None and r.state_percentile is not None]

        filtered_retailers.sort(key=lambda r: r.visits or 0, reverse=True)
        top_retailers = filtered_retailers[:50]

        top_retailers.sort(key=lambda r: r.distance_miles)

        logger.info(f"Retail performance analysis complete: {len(top_retailers)} retailers (filtered from {len(retailers)})")
        return top_retailers, poi_retailer
