from src.clients.placer_client import PlacerClient
from src.core.date_utils import get_last_12_months_date_range
from src.core.logging import logger
from src.models.key_stats import KeyStats


class KeyStatsService:

    # Business rule: Use 10 minute drive time for key stats
    DRIVE_TIME_MINUTES = 10

    def __init__(self, placer_client: PlacerClient) -> None:
        self.placer_client = placer_client

    def get_key_stats(self, reference_poi_id: str) -> KeyStats:
        logger.info(f"Fetching key stats for {self.DRIVE_TIME_MINUTES} min drive time")

        start_date, end_date = get_last_12_months_date_range()

        payload = {
            "method": "driveTime",
            "benchmarkScope": "nationwide",
            "allocationType": "weightedCentroid",
            "trafficVolPct": 70,
            "withinRadius": self.DRIVE_TIME_MINUTES,
            "ringRadius": 3,
            "dataset": "sti_popstats",
            "startDate": start_date,
            "endDate": end_date,
            "apiId": reference_poi_id,
            "driveTime": self.DRIVE_TIME_MINUTES,
            "template": "default",
        }

        demographics_data = self.placer_client.get_demographics(payload)

        data = demographics_data.get("data", {})
        overview = data.get("Overview", {})
        vehicles = data.get("Vehicles per Household", {})

        car_parc = vehicles.get("Total Number of Vehicles", {}).get("value", 0)
        median_income = overview.get("Household Median Income", {}).get("value", 0)
        median_age = overview.get("Median Age", {}).get("value", 0)
        population = overview.get("Population", {}).get("value", 0)
        households = overview.get("Households", {}).get("value", 0)

        logger.info(f"Key stats: Car Parc={car_parc:,.0f}, " f"Median Income=${median_income:,.0f}, " f"Median Age={median_age:.1f}")

        return KeyStats(
            car_counts=None,
            car_parc=int(car_parc),
            median_income=int(median_income),
            median_age=float(median_age),
            population=int(population),
            households=int(households),
        )
