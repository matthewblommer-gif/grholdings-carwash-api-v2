from typing import List, Optional

from pydantic import BaseModel

from src.models.placer.poi import Venue


class AddressRequest(BaseModel):
    address: str


class Coordinates(BaseModel):
    latitude: float
    longitude: float
    formatted_address: str


class POIWithDistance(BaseModel):
    venue: Venue
    distance_miles: float


class POISearchResponse(BaseModel):
    latitude: float
    longitude: float
    formatted_address: str
    pois: List[POIWithDistance]


class AnalyzeRequest(BaseModel):
    address: str
    latitude: float
    longitude: float
    poi_id: str
    poi_name: str
    land_cost: Optional[int] = None
    traffic_counts: Optional[int] = None
