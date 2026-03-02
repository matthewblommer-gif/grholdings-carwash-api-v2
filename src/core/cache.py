from datetime import datetime
from functools import lru_cache
from pathlib import Path

from dateutil.relativedelta import relativedelta
from diskcache import Cache

from src.core.logging import logger

CACHE_DIR = Path(__file__).parent.parent.parent / ".cache"
CACHE_SIZE_LIMIT = 1 * 1024 * 1024 * 1024


@lru_cache(maxsize=1)
def get_cache() -> Cache:
    cache = Cache(directory=str(CACHE_DIR), size_limit=CACHE_SIZE_LIMIT)
    logger.info(f"Cache initialized at {CACHE_DIR}")
    return cache


def calculate_cache_ttl() -> int:
    now = datetime.now()
    next_month = now.replace(day=1) + relativedelta(months=1)
    return int((next_month - now).total_seconds())


def normalize_address(address: str) -> str:
    return " ".join(address.lower().split())
