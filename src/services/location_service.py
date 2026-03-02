from abc import ABC, abstractmethod
from typing import Optional

from src.models.address import Coordinates


class ILocationService(ABC):

    @abstractmethod
    def lookup_address(self, address: str) -> Optional[Coordinates]:
        pass

    @abstractmethod
    def verify_address_exists(self, address: str) -> bool:
        pass

    @abstractmethod
    def get_street_view_url(self, coordinates: Coordinates) -> str:
        pass

    @abstractmethod
    def get_satellite_url(self, coordinates: Coordinates, satellite_zoom_level: int) -> str:
        pass

    @abstractmethod
    def download_street_view_image(self, coordinates: Coordinates) -> Optional[bytes]:
        pass

    @abstractmethod
    def download_satellite_image(self, coordinates: Coordinates, satellite_zoom_level: int) -> Optional[bytes]:
        pass
