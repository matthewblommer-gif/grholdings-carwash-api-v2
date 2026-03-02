from typing import Optional

import googlemaps
import requests

from src.core.cache import calculate_cache_ttl, get_cache, normalize_address
from src.core.logging import logger
from src.core.request_tracker import increment_google_count
from src.models.address import Coordinates
from src.services.location_service import ILocationService


class GoogleLocationService(ILocationService):

    METERS_TO_MILES = 0.000621371

    def __init__(self, api_key: str, image_size: str = "600x400"):
        self._client = googlemaps.Client(key=api_key)
        self._api_key = api_key
        self._image_size = image_size
        self._street_view_base_url = "https://maps.googleapis.com/maps/api/streetview"
        self._static_map_base_url = "https://maps.googleapis.com/maps/api/staticmap"
        self._routes_api_base_url = "https://routes.googleapis.com/directions/v2:computeRoutes"
        self._cache = get_cache()

    def _make_cache_key(self, prefix: str, *args: str) -> str:
        return f"google:{prefix}:{':'.join(str(a) for a in args)}"

    def lookup_address(self, address: str) -> Optional[Coordinates]:
        normalized = normalize_address(address)
        cache_key = self._make_cache_key("geocode", normalized)

        cached = self._cache.get(cache_key)
        if cached is not None:
            logger.debug(f"Cache hit for geocode: {address}")
            if cached == "__NONE__":
                return None
            return Coordinates(**cached)

        try:
            increment_google_count()
            result = self._client.geocode(address)

            if not result:
                self._cache.set(cache_key, "__NONE__", expire=calculate_cache_ttl())
                return None

            location = result[0]["geometry"]["location"]
            formatted_address = result[0]["formatted_address"]

            coords = Coordinates(latitude=location["lat"], longitude=location["lng"], formatted_address=formatted_address)
            self._cache.set(cache_key, coords.model_dump(), expire=calculate_cache_ttl())
            return coords

        except Exception as e:
            logger.error(f"Error looking up address '{address}': {e}")
            return None

    def verify_address_exists(self, address: str) -> bool:
        normalized = normalize_address(address)
        cache_key = self._make_cache_key("validate", normalized)

        cached = self._cache.get(cache_key)
        if cached is not None:
            logger.debug(f"Cache hit for address validation: {address}")
            return cached

        try:
            increment_google_count()
            response = self._client.addressvalidation([address], regionCode="US")

            if not response:
                self._cache.set(cache_key, False, expire=calculate_cache_ttl())
                return False

            result = response.get("result", {})
            verdict = result.get("verdict", {})
            has_unconfirmed_components = verdict.get("hasUnconfirmedComponents", False)
            logger.debug(f"has_unconfirmed_components: {has_unconfirmed_components}")

            is_valid = not has_unconfirmed_components
            self._cache.set(cache_key, is_valid, expire=calculate_cache_ttl())
            return is_valid

        except Exception as e:
            logger.error(f"Error verifying address: {e}")
            return False

    def get_street_view_url(self, coordinates: Coordinates) -> str:
        params = {
            "size": self._image_size,
            "location": f"{self._format_address_for_streetview(coordinates.formatted_address)}",
            "key": self._api_key,
        }

        param_string = "&".join([f"{k}={v}" for k, v in params.items()])
        return f"{self._street_view_base_url}?{param_string}"

    def _format_address_for_streetview(self, address: str) -> str:
        return address.replace(",", "").replace(" ", "+")

    def get_satellite_url(self, coordinates: Coordinates, satellite_zoom_level: int) -> str:
        params = {
            "center": f"{coordinates.latitude},{coordinates.longitude}",
            "zoom": satellite_zoom_level,
            "size": self._image_size,
            "maptype": "hybrid",
            "markers": f"color:red|{coordinates.latitude},{coordinates.longitude}",
            "key": self._api_key,
        }

        param_string = "&".join([f"{k}={v}" for k, v in params.items()])
        return f"{self._static_map_base_url}?{param_string}"

    def download_street_view_image(self, coordinates: Coordinates) -> Optional[bytes]:
        try:
            url = self.get_street_view_url(coordinates)
            increment_google_count()
            response = requests.get(url)

            if response.status_code == 200:
                return response.content

            logger.warning(f"Street view image request returned status {response.status_code} for coordinates {coordinates.latitude},{coordinates.longitude}")
            return None

        except Exception as e:
            logger.error(f"Error downloading street view image for coordinates {coordinates.latitude},{coordinates.longitude}: {e}")
            return None

    def download_satellite_image(self, coordinates: Coordinates, satellite_zoom_level: int) -> Optional[bytes]:
        try:
            url = self.get_satellite_url(coordinates, satellite_zoom_level)
            increment_google_count()
            response = requests.get(url)

            if response.status_code == 200:
                return response.content

            logger.warning(f"Satellite image request returned status {response.status_code} for coordinates {coordinates.latitude},{coordinates.longitude}")
            return None

        except Exception as e:
            logger.error(f"Error downloading satellite image for coordinates {coordinates.latitude},{coordinates.longitude}: {e}")
            return None

    def calculate_drive_time_and_distance(self, origin_latitude: float, origin_longitude: float, destination_address: str) -> tuple[Optional[float], Optional[float]]:
        normalized_dest = normalize_address(destination_address)
        cache_key = self._make_cache_key("routes", f"{origin_latitude},{origin_longitude}", normalized_dest)

        cached = self._cache.get(cache_key)
        if cached is not None:
            logger.debug(f"Cache hit for drive time: {destination_address}")
            return cached

        try:
            increment_google_count()
            geocode_result = self._client.geocode(destination_address)

            if not geocode_result:
                logger.warning(f"Could not geocode destination address: {destination_address}")
                self._cache.set(cache_key, (None, None), expire=calculate_cache_ttl())
                return None, None

            venue_location = geocode_result[0]["geometry"]["location"]
            venue_latitude = venue_location["lat"]
            venue_longitude = venue_location["lng"]

            payload = {
                "origin": {"location": {"latLng": {"latitude": origin_latitude, "longitude": origin_longitude}}},
                "destination": {"location": {"latLng": {"latitude": venue_latitude, "longitude": venue_longitude}}},
                "travelMode": "DRIVE",
                "routingPreference": "TRAFFIC_AWARE",
                "computeAlternativeRoutes": False,
                "routeModifiers": {
                    "avoidTolls": False,
                    "avoidHighways": False,
                    "avoidFerries": False,
                },
                "languageCode": "en-US",
                "units": "IMPERIAL",
            }

            headers = {
                "Content-Type": "application/json",
                "X-Goog-Api-Key": self._api_key,
                "X-Goog-FieldMask": "routes.duration,routes.distanceMeters",
            }

            increment_google_count()
            response = requests.post(self._routes_api_base_url, json=payload, headers=headers)

            if response.status_code == 200:
                result = response.json()

                if "routes" in result and len(result["routes"]) > 0:
                    route = result["routes"][0]
                    distance_meters = route["distanceMeters"]
                    distance_miles = distance_meters * self.METERS_TO_MILES
                    duration_seconds = int(route["duration"].rstrip("s"))
                    duration_minutes = duration_seconds / 60

                    self._cache.set(cache_key, (duration_minutes, distance_miles), expire=calculate_cache_ttl())
                    return duration_minutes, distance_miles

            logger.warning(f"No route found for destination: {destination_address}")
            self._cache.set(cache_key, (None, None), expire=calculate_cache_ttl())
            return None, None

        except requests.RequestException as e:
            logger.error(f"HTTP error calculating drive time to {destination_address}: {e}")
            return None, None
        except (KeyError, ValueError, IndexError) as e:
            logger.error(f"Error parsing route response for {destination_address}: {e}")
            return None, None

    def calculate_distance(self, origin_latitude: float, origin_longitude: float, destination_address: str) -> Optional[float]:
        _, distance_miles = self.calculate_drive_time_and_distance(origin_latitude, origin_longitude, destination_address)
        return distance_miles
