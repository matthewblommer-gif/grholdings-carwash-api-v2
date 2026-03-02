from pydantic import BaseModel
from typing import Optional


class Competitor(BaseModel):
    name: str
    api_id: str
    address: str
    drive_time_minutes: float
    drive_time_band: str
    distance_miles: float
    competitor_type: Optional[str] = None
    quality: Optional[str] = None
    total_members: int
    overlap_percentage: float
    members_in_market: int
    visits_per_year: int
    visits_per_month: int
    visits_per_day: int
    car_parc: Optional[int] = None
    tta_visits: Optional[int] = None
