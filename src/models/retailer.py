from pydantic import BaseModel
from typing import Optional


class Retailer(BaseModel):
    name: str
    api_id: str
    category: str
    sub_category: Optional[str] = None
    address: str
    distance_miles: float
    national_percentile: Optional[float] = None
    state_percentile: Optional[float] = None
    visits: Optional[int] = None
