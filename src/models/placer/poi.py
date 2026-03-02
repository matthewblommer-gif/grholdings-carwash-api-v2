from pydantic import BaseModel
from typing import Optional, List


class CategoryInfo(BaseModel):
    category: str
    group: str
    subCategory: str


class Address(BaseModel):
    city: Optional[str] = ""
    state: Optional[str] = ""
    countryCode: Optional[str] = ""
    streetName: Optional[str] = ""
    formattedAddress: Optional[str] = ""
    shortFormattedAddress: Optional[str] = ""
    zipCode: Optional[str] = ""


class RegionDetail(BaseModel):
    code: str = ""
    name: str = ""


class Regions(BaseModel):
    dma: Optional[RegionDetail] = None
    state: Optional[RegionDetail] = None
    cbsa: Optional[RegionDetail] = None


class Venue(BaseModel):
    entityId: str
    entityType: str
    name: str
    categoryInfo: CategoryInfo
    address: Address
    isFlagged: bool
    regions: Regions
    apiId: str
    placerUrl: str
    storeId: Optional[str] = None
    isPermitted: bool


class PlacerPOIResponse(BaseModel):
    data: List[Venue]
    requestId: str
