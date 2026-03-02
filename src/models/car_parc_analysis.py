from pydantic import BaseModel
from typing import List
from src.models.car_parc_result import CarParcResult


class CarParcAnalysis(BaseModel):
    results: List[CarParcResult]
    reference_poi_id: str
    reference_poi_name: str
