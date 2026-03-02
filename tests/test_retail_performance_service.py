from unittest.mock import Mock, MagicMock
import pytest

from src.services.retail_performance_service import RetailPerformanceService
from src.models.placer.poi import Venue, CategoryInfo, Address, Regions, RegionDetail


def create_mock_venue(api_id: str, name: str, category: str = "Groceries") -> Venue:
    return Venue(
        entityId=f"entity_{api_id}",
        entityType="venue",
        name=name,
        categoryInfo=CategoryInfo(category=category, group="Retail", subCategory=""),
        address=Address(
            city="Portsmouth",
            state="NH",
            countryCode="US",
            streetName="Main St",
            formattedAddress=f"{name}, Portsmouth, NH",
            shortFormattedAddress=f"{name}, Portsmouth",
            zipCode="03801",
        ),
        isFlagged=False,
        regions=Regions(dma=RegionDetail(code="500", name="Boston"), state=RegionDetail(code="NH", name="New Hampshire"), cbsa=RegionDetail(code="12345", name="Boston Metro")),
        apiId=api_id,
        placerUrl=f"https://placer.ai/venue/{api_id}",
        isPermitted=True,
    )


def create_mock_ranking_response(api_id: str, national_percentile: int, state_percentile: int):
    return {
        "data": [
            {
                "visitDurationSegmentation": "allVisits",
                "info": {
                    "entityId": f"entity_{api_id}",
                    "entityType": "venue",
                    "name": "Test Venue",
                    "flagged": False,
                    "rankedBy": "visits",
                    "categoryInfo": {"category": "Groceries", "group": "Retail", "subCategory": ""},
                },
                "apiId": api_id,
                "metricType": "visits",
                "ranking": {
                    "nationwide": {"rank": 100, "percentile": national_percentile, "rankedOutOf": 1000, "regionCode": "US"},
                    "state": {"rank": 10, "percentile": state_percentile, "rankedOutOf": 100, "regionCode": "NH"},
                },
            }
        ],
        "requestId": "test-request-id",
    }


def create_mock_visit_trends_response(api_id: str, total_visits: int):
    return {
        "data": [
            {
                "visitDurationSegmentation": "allVisits",
                "apiId": api_id,
                "dates": ["2024-01-01", "2024-02-01", "2024-03-01"],
                "visits": [total_visits // 3, total_visits // 3, total_visits // 3 + total_visits % 3],
                "panelVisits": [1000, 1000, 1000],
            }
        ],
        "requestId": "test-request-id",
    }


class TestRetailPerformanceService:

    # REQ: POI should be excluded from retail search results
    def test_poi_excluded_from_retail_results(self):
        mock_placer_client = Mock()
        mock_google_service = Mock()

        poi_id = "poi_market_basket_123"
        poi_name = "Market Basket"

        nearby_venue_1 = create_mock_venue("nearby_1", "Dunkin")
        nearby_venue_2 = create_mock_venue("nearby_2", "McDonald's")
        poi_venue = create_mock_venue(poi_id, poi_name)

        mock_placer_client.search_poi.return_value = {"data": [nearby_venue_1.model_dump(), nearby_venue_2.model_dump(), poi_venue.model_dump()], "requestId": "test"}

        def mock_get_ranking_single(payload, cache=None):
            api_id = payload.get("apiId")
            if api_id == poi_id:
                return create_mock_ranking_response(poi_id, 88, 84)
            elif api_id == "nearby_1":
                return create_mock_ranking_response("nearby_1", 70, 65)
            elif api_id == "nearby_2":
                return create_mock_ranking_response("nearby_2", 60, 55)
            return {"data": [], "requestId": "test"}

        mock_placer_client.get_ranking_single.side_effect = mock_get_ranking_single

        mock_placer_client.get_visit_trends.return_value = {
            "data": [
                create_mock_visit_trends_response(poi_id, 100000)["data"][0],
                create_mock_visit_trends_response("nearby_1", 50000)["data"][0],
                create_mock_visit_trends_response("nearby_2", 30000)["data"][0],
            ],
            "requestId": "test",
        }

        mock_google_service.calculate_distance.return_value = 0.5

        service = RetailPerformanceService(mock_placer_client, mock_google_service)

        retailers, poi_retail = service.analyze_retail_performance(latitude=43.0859778, longitude=-70.78677, poi_id=poi_id, poi_name=poi_name)

        retailer_api_ids = [r.api_id for r in retailers]
        assert poi_id not in retailer_api_ids, "POI should not appear in retailer list"

    # REQ: POI's own retail stats should be returned separately
    def test_poi_retail_stats_returned_separately(self):
        mock_placer_client = Mock()
        mock_google_service = Mock()

        poi_id = "poi_market_basket_123"
        poi_name = "Market Basket"

        mock_placer_client.search_poi.return_value = {"data": [], "requestId": "test"}
        mock_placer_client.get_ranking_single.return_value = create_mock_ranking_response(poi_id, 88, 84)
        mock_placer_client.get_visit_trends.return_value = create_mock_visit_trends_response(poi_id, 782247)

        service = RetailPerformanceService(mock_placer_client, mock_google_service)

        retailers, poi_retail = service.analyze_retail_performance(latitude=43.0859778, longitude=-70.78677, poi_id=poi_id, poi_name=poi_name)

        assert poi_retail is not None, "POI retail stats should be returned"
        assert poi_retail.api_id == poi_id
        assert poi_retail.name == poi_name
        assert poi_retail.national_percentile == 0.88
        assert poi_retail.state_percentile == 0.84
        assert poi_retail.visits == 782247

    # REQ: POI retail stats should have correct percentile format (0-1 scale)
    def test_poi_retail_stats_percentile_format(self):
        mock_placer_client = Mock()
        mock_google_service = Mock()

        poi_id = "test_poi"
        poi_name = "Test POI"

        mock_placer_client.search_poi.return_value = {"data": [], "requestId": "test"}
        mock_placer_client.get_ranking_single.return_value = create_mock_ranking_response(poi_id, 45, 30)
        mock_placer_client.get_visit_trends.return_value = create_mock_visit_trends_response(poi_id, 100000)

        service = RetailPerformanceService(mock_placer_client, mock_google_service)

        _, poi_retail = service.analyze_retail_performance(latitude=43.0, longitude=-70.0, poi_id=poi_id, poi_name=poi_name)

        assert poi_retail.national_percentile == 0.45
        assert poi_retail.state_percentile == 0.30
