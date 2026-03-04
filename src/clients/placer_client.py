import json
import time
from typing import Any, Optional

import requests

from src.core.cache import calculate_cache_ttl, get_cache
from src.core.date_utils import shift_dates_back
from src.core.logging import logger
from src.core.request_tracker import increment_placer_count


class PlacerClient:

    def __init__(self, api_key: str, base_url: str = "https://papi.placer.ai/v1") -> None:
        self.api_key = api_key
        self.base_url = base_url
        self._headers = {"accept": "application/json", "x-api-key": api_key}
        self._cache = get_cache()

    def _make_cache_key(self, endpoint: str, payload: Any) -> str:
        if isinstance(payload, dict):
            sorted_payload = json.dumps(payload, sort_keys=True)
            return f"placer:{endpoint}:{sorted_payload}"
        return f"placer:{endpoint}:{str(payload)}"

    def _get_cached(self, key: str) -> Optional[Any]:
        return self._cache.get(key)

    def _set_cached(self, key: str, value: Any) -> None:
        self._cache.set(key, value, expire=calculate_cache_ttl())

    def _shift_payload_dates(self, payload: dict, months: int = 1) -> dict:
        """Create a copy of the payload with dates shifted back by the given number of months for Placer data lag."""
        shifted = payload.copy()
        if "startDate" in shifted and "endDate" in shifted:
            shifted["startDate"], shifted["endDate"] = shift_dates_back(
                shifted["startDate"], shifted["endDate"], months=months
            )
        return shifted

    def search_poi(
        self,
        lat: float,
        lng: float,
        radius: float,
        category: Optional[str] = None,
        sub_category: Optional[str] = None,
        entity_type: str = "venue",
        limit: int = 50,
    ) -> dict:
        params = {
            "lat": lat,
            "lng": lng,
            "radius": radius,
            "entityType": entity_type,
            "limit": limit,
            "category": category,
            "subCategory": sub_category,
        }
        cache_key = self._make_cache_key("poi", params)

        cached = self._get_cached(cache_key)
        if cached is not None:
            logger.debug(f"Cache hit for POI search: {category}")
            return cached

        url = f"{self.base_url}/poi"
        request_params = {
            "lat": lat,
            "lng": lng,
            "radius": radius,
            "entityType": entity_type,
            "limit": limit,
        }

        if category:
            request_params["category"] = category
        if sub_category:
            request_params["subCategory"] = sub_category

        logger.debug(f"Searching POI: lat={lat}, lng={lng}, radius={radius}, category={category}")

        increment_placer_count()
        response = requests.get(url, params=request_params, headers=self._headers)
        response.raise_for_status()

        result = response.json()
        logger.info(f"POI search response for category {category}: {type(result)} keys={list(result.keys()) if result else 'empty'}")

        if result == {} or "data" not in result or not result.get("data"):
            logger.info(f"POI search returned no data for category: {category}")
            result = {"data": [], "requestId": ""}

        self._set_cached(cache_key, result)
        return result

    def get_demographics(self, payload: dict, max_retries: int = 15, retry_delay_seconds: int = 2, _shift_count: int = 0) -> Optional[dict]:
        cache_key = self._make_cache_key("demographics", payload)

        cached = self._get_cached(cache_key)
        if cached is not None:
            logger.debug(f"Cache hit for demographics API ID: {payload.get('apiId')}, driveTime: {payload.get('driveTime')}")
            return cached

        url = f"{self.base_url}/reports/trade-area-demographics"
        headers = {**self._headers, "content-type": "application/json"}

        logger.debug(f"Fetching demographics for API ID: {payload.get('apiId')}, driveTime: {payload.get('driveTime')}")

        for attempt in range(max_retries):
            increment_placer_count()
            response = requests.post(url, json=payload, headers=headers)

            if response.status_code == 204:
                logger.info(f"Demographics API returned no data (204) for API ID: {payload.get('apiId')}, driveTime: {payload.get('driveTime')}")
                return None

            if response.status_code not in [200, 202]:
                if response.status_code == 400 and _shift_count < 3 and "startDate" in payload:
                    shifted = self._shift_payload_dates(payload, months=_shift_count + 1)
                    logger.warning(
                        f"Demographics API returned 400 for dates {payload['startDate']} to {payload['endDate']}. "
                        f"Retrying with {shifted['startDate']} to {shifted['endDate']}"
                    )
                    return self.get_demographics(shifted, max_retries, retry_delay_seconds, _shift_count=_shift_count + 1)
                logger.error(f"Demographics API error: {response.status_code} - {response.text}")
                response.raise_for_status()

            if "IN_PROGRESS" in response.text or response.status_code == 202:
                if attempt < max_retries - 1:
                    logger.debug(f"Demographics processing, retry {attempt + 1}/{max_retries}")
                    time.sleep(retry_delay_seconds)
                    continue
                else:
                    logger.error(f"Demographics request timed out after {max_retries} attempts")
                    raise TimeoutError(f"Demographics request timed out after {max_retries} attempts")

            result = response.json()
            logger.debug("Demographics data received successfully")

            self._set_cached(cache_key, result)
            return result

        raise TimeoutError(f"Demographics request timed out after {max_retries} attempts")

    def get_visit_trends(self, payload: dict, max_retries: int = 10, retry_delay_seconds: int = 3, _shift_count: int = 0) -> dict:
        api_ids = payload.get("apiIds", [])

        # Batch large requests to avoid 400 errors from too many apiIds
        if len(api_ids) > 20:
            logger.info(f"Batching visit trends request: {len(api_ids)} venues in chunks of 20")
            all_data = []
            request_id = ""
            for i in range(0, len(api_ids), 20):
                chunk = api_ids[i:i + 20]
                chunk_payload = {**payload, "apiIds": chunk}
                chunk_result = self.get_visit_trends(chunk_payload, max_retries, retry_delay_seconds, _shift_count)
                all_data.extend(chunk_result.get("data", []))
                request_id = chunk_result.get("requestId", request_id)
            combined = {"data": all_data, "requestId": request_id}
            self._set_cached(self._make_cache_key("visit_trends", payload), combined)
            return combined

        cache_key = self._make_cache_key("visit_trends", payload)

        cached = self._get_cached(cache_key)
        if cached is not None:
            logger.debug(f"Cache hit for visit trends: {len(api_ids)} venues")
            return cached

        url = f"{self.base_url}/reports/visit-trends"
        headers = {**self._headers, "content-type": "application/json"}

        logger.debug(f"Fetching visit trends for {len(api_ids)} venues")

        for attempt in range(max_retries):
            increment_placer_count()
            response = requests.post(url, json=payload, headers=headers)

            if response.status_code not in [200, 207]:
                if response.status_code == 400 and _shift_count < 3 and "startDate" in payload:
                    shifted = self._shift_payload_dates(payload, months=_shift_count + 1)
                    logger.warning(
                        f"Visit trends API returned 400 for dates {payload['startDate']} to {payload['endDate']}. "
                        f"Retrying with {shifted['startDate']} to {shifted['endDate']}"
                    )
                    return self.get_visit_trends(shifted, max_retries, retry_delay_seconds, _shift_count=_shift_count + 1)
                logger.error(f"Visit trends API error: {response.status_code} - {response.text}")
                response.raise_for_status()

            if "IN_PROGRESS" in response.text:
                if attempt < max_retries - 1:
                    logger.debug(f"Visit trends processing, retry {attempt + 1}/{max_retries}")
                    time.sleep(retry_delay_seconds)
                    continue
                else:
                    logger.error(f"Visit trends request timed out after {max_retries} attempts")
                    raise TimeoutError(f"Visit trends request timed out after {max_retries} attempts")

            result = response.json()
            logger.debug("Visit trends data received successfully")

            self._set_cached(cache_key, result)
            return result

        raise TimeoutError(f"Visit trends request timed out after {max_retries} attempts")

    def get_loyalty_frequency(self, payload: dict, max_retries: int = 10, retry_delay_seconds: int = 3, _shift_count: int = 0) -> dict:
        cache_key = self._make_cache_key("loyalty", payload)

        cached = self._get_cached(cache_key)
        if cached is not None:
            logger.debug(f"Cache hit for loyalty API ID: {payload.get('apiId')}")
            return cached

        url = f"{self.base_url}/reports/loyalty/visits-frequency"
        headers = {**self._headers, "content-type": "application/json"}

        logger.debug(f"Fetching loyalty data for API ID: {payload.get('apiId')}")

        for attempt in range(max_retries):
            increment_placer_count()
            response = requests.post(url, json=payload, headers=headers)

            if "IN_PROGRESS" in response.text or response.status_code == 202:
                if attempt < max_retries - 1:
                    logger.debug(f"Loyalty data processing, retry {attempt + 1}/{max_retries}")
                    time.sleep(retry_delay_seconds)
                    continue
                else:
                    logger.error(f"Loyalty request timed out after {max_retries} attempts")
                    raise TimeoutError(f"Loyalty request timed out after {max_retries} attempts")

            if response.status_code not in [200, 207]:
                if response.status_code == 400 and _shift_count < 3 and "startDate" in payload:
                    shifted = self._shift_payload_dates(payload, months=_shift_count + 1)
                    logger.warning(
                        f"Loyalty API returned 400 for dates {payload['startDate']} to {payload['endDate']}. "
                        f"Retrying with {shifted['startDate']} to {shifted['endDate']}"
                    )
                    return self.get_loyalty_frequency(shifted, max_retries, retry_delay_seconds, _shift_count=_shift_count + 1)
                logger.error(f"Loyalty API error: {response.status_code}")
                response.raise_for_status()

            result = response.json()
            logger.debug("Loyalty data received successfully")

            self._set_cached(cache_key, result)
            return result

        raise TimeoutError(f"Loyalty request timed out after {max_retries} attempts")

    def get_trade_area(self, payload: dict, max_retries: int = 10, retry_delay_seconds: int = 3, _shift_count: int = 0) -> dict:
        cache_key = self._make_cache_key("trade_area", payload)

        cached = self._get_cached(cache_key)
        if cached is not None:
            logger.debug(f"Cache hit for trade area API ID: {payload.get('apiId')}")
            return cached

        url = f"{self.base_url}/reports/true-trade-area"
        headers = {**self._headers, "content-type": "application/json"}

        logger.debug(f"Fetching trade area for API ID: {payload.get('apiId')}")

        for attempt in range(max_retries):
            increment_placer_count()
            response = requests.post(url, json=payload, headers=headers)

            if "IN_PROGRESS" in response.text or response.status_code == 202:
                if attempt < max_retries - 1:
                    logger.debug(f"Trade area processing, retry {attempt + 1}/{max_retries}")
                    time.sleep(retry_delay_seconds)
                    continue
                else:
                    logger.error(f"Trade area request timed out after {max_retries} attempts")
                    raise TimeoutError(f"Trade area request timed out after {max_retries} attempts")

            if response.status_code != 200:
                if response.status_code == 400 and _shift_count < 3 and "startDate" in payload:
                    shifted = self._shift_payload_dates(payload, months=_shift_count + 1)
                    logger.warning(
                        f"Trade area API returned 400 for dates {payload['startDate']} to {payload['endDate']}. "
                        f"Retrying with {shifted['startDate']} to {shifted['endDate']}"
                    )
                    return self.get_trade_area(shifted, max_retries, retry_delay_seconds, _shift_count=_shift_count + 1)
                logger.error(f"Trade area API error: {response.status_code}")
                response.raise_for_status()

            result = response.json()
            logger.debug("Trade area data received successfully")

            self._set_cached(cache_key, result)
            return result

        raise TimeoutError(f"Trade area request timed out after {max_retries} attempts")

    def get_ranking_single(self, payload: dict, _shift_count: int = 0) -> dict:
        cache_key = self._make_cache_key("ranking_single", payload)

        cached = self._get_cached(cache_key)
        if cached is not None:
            logger.debug(f"Cache hit for ranking single: {payload.get('apiId')}")
            return cached

        url = f"{self.base_url}/reports/ranking-overview"
        headers = {**self._headers, "content-type": "application/json"}

        logger.debug(f"Fetching ranking for venue: {payload.get('apiId')}")

        increment_placer_count()
        response = requests.post(url, json=payload, headers=headers)

        if response.status_code not in [200, 207]:
            if response.status_code == 400 and _shift_count < 3 and "startDate" in payload:
                shifted = self._shift_payload_dates(payload, months=_shift_count + 1)
                logger.warning(
                    f"Ranking API returned 400 for dates {payload['startDate']} to {payload['endDate']}. "
                    f"Retrying with {shifted['startDate']} to {shifted['endDate']}"
                )
                return self.get_ranking_single(shifted, _shift_count=_shift_count + 1)
            logger.error(f"Ranking single API error: {response.status_code} - {response.text}")
            response.raise_for_status()

        result = response.json()
        logger.debug("Ranking single data received successfully")

        self._set_cached(cache_key, result)
        return result
