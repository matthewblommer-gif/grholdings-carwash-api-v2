from pydantic import BaseModel
from typing import List, Union


class TradeAreaData(BaseModel):
    type: str  # "Polygon" or "MultiPolygon"
    coordinates: Union[List[List[List[float]]], List[List[List[List[float]]]]]  # Polygon (3 levels)  # MultiPolygon (4 levels)
    visitDurationSegmentation: str


class TradeAreaResponse(BaseModel):
    apiId: str
    data: TradeAreaData
