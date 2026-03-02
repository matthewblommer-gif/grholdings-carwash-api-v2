from contextvars import ContextVar
from dataclasses import dataclass
from typing import Optional

from src.core.logging import logger


@dataclass
class RequestCounts:
    placer: int = 0
    google: int = 0

    def increment_placer(self) -> None:
        self.placer += 1

    def increment_google(self) -> None:
        self.google += 1

    def total(self) -> int:
        return self.placer + self.google

    def log_summary(self, endpoint: str) -> None:
        logger.info(f"API calls for {endpoint}: Placer={self.placer}, Google={self.google}, Total={self.total()}")


_request_counts: ContextVar[Optional[RequestCounts]] = ContextVar("request_counts", default=None)


def init_request_counts() -> RequestCounts:
    counts = RequestCounts()
    _request_counts.set(counts)
    return counts


def get_request_counts() -> Optional[RequestCounts]:
    return _request_counts.get()


def increment_placer_count() -> None:
    counts = _request_counts.get()
    if counts is not None:
        counts.increment_placer()


def increment_google_count() -> None:
    counts = _request_counts.get()
    if counts is not None:
        counts.increment_google()
