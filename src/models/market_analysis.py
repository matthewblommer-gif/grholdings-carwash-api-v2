from typing import List, Optional

from pydantic import BaseModel

from src.models.car_parc_result import CarParcResult
from src.models.competitor import Competitor
from src.models.key_stats import KeyStats
from src.models.retailer import Retailer


class MarketAnalysis(BaseModel):
    address: str
    latitude: float
    longitude: float
    reference_poi_id: str
    reference_poi_name: str
    car_parc_results: List[CarParcResult]
    competitors: List[Competitor]
    key_stats: KeyStats
    retailers: List[Retailer]
    reference_poi_retail: Optional[Retailer] = None
    total_market_members: int
    land_cost: Optional[int] = None
    traffic_counts: Optional[int] = None
    warnings: List[str] = []
