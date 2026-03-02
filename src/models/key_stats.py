from pydantic import BaseModel
from typing import Optional


class KeyStats(BaseModel):
    car_counts: Optional[int] = None
    car_parc: int
    median_income: int
    median_age: float
    population: int
    households: int
