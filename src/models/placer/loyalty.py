from pydantic import BaseModel
from typing import List


class LoyaltyFrequencyData(BaseModel):
    apiId: str
    startDate: str
    endDate: str
    avgVisitsPerCustomer: float
    medianVisitsPerCustomer: int
    bins: List[int]
    visitors: List[int]
    visitorsPercentage: List[float]
    visits: List[int]
    visitsPercentage: List[float]
    visitDurationSegmentation: str


class LoyaltyFrequencyResponse(BaseModel):
    data: LoyaltyFrequencyData
    requestId: str
