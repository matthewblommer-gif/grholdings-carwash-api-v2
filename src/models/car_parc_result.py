from pydantic import BaseModel


class CarParcResult(BaseModel):
    drive_time_minutes: int
    car_parc: int
    population: int
    households: int
    tam_percentage: float
    total_addressable_market: int
    market_share_percentage: float
