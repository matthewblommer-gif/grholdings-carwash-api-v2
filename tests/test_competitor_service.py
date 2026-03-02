from unittest.mock import Mock

from src.services.competitor_service import CompetitorService


def create_mock_loyalty_frequency_response(bins: list[int], visitors: list[int]):
    return {
        "data": {
            "apiId": "venue:7d55054520e387813d764b03",
            "startDate": "2024-01-01",
            "endDate": "2024-06-30",
            "avgVisitsPerCustomer": 2.5,
            "medianVisitsPerCustomer": 2,
            "bins": bins,
            "visitors": visitors,
            "visitorsPercentage": [0.5] * len(visitors),
            "visits": [v * 2 for v in visitors],
            "visitsPercentage": [0.5] * len(visitors),
            "visitDurationSegmentation": "allVisits",
        },
        "requestId": "test-request-id",
    }


def test_count_loyal_visitors_returns_visitors_at_or_above_threshold():
    mock_placer_client = Mock()
    mock_google_service = Mock()

    mock_placer_client.get_loyalty_frequency.return_value = create_mock_loyalty_frequency_response(
        bins=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, 20, 30],
        visitors=[319142, 106216, 51907, 37322, 22606, 13947, 9933, 7663, 6064, 14217, 5645, 2698, 2522],
    )

    service = CompetitorService(mock_placer_client, mock_google_service)

    result = service._count_loyal_visitors("venue:7d55054520e387813d764b03", "2024-01-01", "2024-06-30")

    # Sum of visitors with bins >= 6: 13947 + 9933 + 7663 + 6064 + 14217 + 5645 + 2698 + 2522 = 62689
    assert result == 62689


def test_count_loyal_visitors_returns_zero_when_all_below_threshold():
    mock_placer_client = Mock()
    mock_google_service = Mock()

    mock_placer_client.get_loyalty_frequency.return_value = create_mock_loyalty_frequency_response(
        bins=[1, 2, 3, 4, 5],
        visitors=[100, 200, 300, 400, 500],
    )

    service = CompetitorService(mock_placer_client, mock_google_service)

    result = service._count_loyal_visitors("venue:test", "2024-01-01", "2024-06-30")

    assert result == 0


def test_count_loyal_visitors_handles_empty_response():
    mock_placer_client = Mock()
    mock_google_service = Mock()

    mock_placer_client.get_loyalty_frequency.return_value = create_mock_loyalty_frequency_response(
        bins=[],
        visitors=[],
    )

    service = CompetitorService(mock_placer_client, mock_google_service)

    result = service._count_loyal_visitors("venue:test", "2024-01-01", "2024-06-30")

    assert result == 0


def test_calculate_total_members_uses_single_period():
    mock_placer_client = Mock()
    mock_google_service = Mock()

    mock_placer_client.get_loyalty_frequency.return_value = create_mock_loyalty_frequency_response(
        bins=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        visitors=[1000, 500, 300, 200, 150, 100, 80, 60, 40, 20],
    )

    service = CompetitorService(mock_placer_client, mock_google_service)

    result = service.calculate_total_members(
        api_id="venue:test",
        start_date="2024-07-01",
        end_date="2024-12-31",
    )

    # Sum of visitors with bins >= 6: 100 + 80 + 60 + 40 + 20 = 300
    assert result == 300
