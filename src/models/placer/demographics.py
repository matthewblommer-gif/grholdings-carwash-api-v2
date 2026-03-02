from pydantic import BaseModel
from typing import Dict, Any


class DemographicsDataValue(BaseModel):
    value: float


class DemographicsResponse(BaseModel):
    data: Dict[str, Any]
    requestId: str
