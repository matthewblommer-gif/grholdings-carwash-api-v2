from pydantic import BaseModel
from typing import Union


class VisitTrendsData(BaseModel):
    visitDurationSegmentation: str
    apiId: str
    dates: list[str]
    visits: list[int]
    panelVisits: list[int]


class NoContentError(BaseModel):
    apiId: str
    details: str
    error: str
    code: int


class VisitTrendsResponse(BaseModel):
    data: list[Union[VisitTrendsData, NoContentError]]
    requestId: str
