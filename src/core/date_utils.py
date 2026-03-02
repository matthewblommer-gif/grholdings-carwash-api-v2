from datetime import datetime
from calendar import monthrange
from dateutil.relativedelta import relativedelta


def get_last_12_months_date_range() -> tuple[str, str]:
    now = datetime.now()
    end_month = now.replace(day=1) - relativedelta(months=1)
    last_day = monthrange(end_month.year, end_month.month)[1]
    end_date = end_month.replace(day=last_day)
    start_date = end_month - relativedelta(months=11)
    return start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")


def get_last_12_months_as_two_halves() -> tuple[tuple[str, str], tuple[str, str]]:
    start_date_str, end_date_str = get_last_12_months_date_range()
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")

    # First half: months 1-6
    first_half_end_month = start_date + relativedelta(months=5)
    first_half_end_day = monthrange(first_half_end_month.year, first_half_end_month.month)[1]
    first_half_end = first_half_end_month.replace(day=first_half_end_day)

    # Second half: months 7-12
    second_half_start = first_half_end_month + relativedelta(months=1)

    return (
        (start_date_str, first_half_end.strftime("%Y-%m-%d")),
        (second_half_start.strftime("%Y-%m-%d"), end_date_str),
    )


def shift_dates_back(start_date: str, end_date: str, months: int = 1) -> tuple[str, str]:
    """Shift a date range back by the specified number of months.

    Used as a fallback when Placer.ai returns 400 because the most recent
    month's data is not yet available (e.g., during their first-Monday-of-month refresh).
    """
    start = datetime.strptime(start_date, "%Y-%m-%d")
    end = datetime.strptime(end_date, "%Y-%m-%d")

    new_start = start - relativedelta(months=months)

    end_first_of_month = end.replace(day=1)
    new_end_month = end_first_of_month - relativedelta(months=months)
    new_end_day = monthrange(new_end_month.year, new_end_month.month)[1]
    new_end = new_end_month.replace(day=new_end_day)

    return new_start.strftime("%Y-%m-%d"), new_end.strftime("%Y-%m-%d")
