from pydantic import BaseModel
from typing import Optional, List
from src.models.placer.poi import CategoryInfo


class RankingDetail(BaseModel):
    rank: int
    percentile: int
    rankedOutOf: int
    regionCode: Optional[str] = None


class VenueRanking(BaseModel):
    nationwide: Optional[RankingDetail] = None
    state: Optional[RankingDetail] = None
    dma: Optional[RankingDetail] = None
    cbsa: Optional[RankingDetail] = None
    rankError: Optional[str] = None


class VenueInfo(BaseModel):
    entityId: str
    entityType: str
    name: str
    flagged: bool
    rankedBy: str
    parentChain: Optional[str] = None
    categoryInfo: CategoryInfo


class RankingData(BaseModel):
    visitDurationSegmentation: str
    info: VenueInfo
    apiId: str
    metricType: str
    ranking: VenueRanking


class RankingResponse(BaseModel):
    data: List[RankingData]
    requestId: str
