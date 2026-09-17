"""Compatibility wrapper for the former investpy-based EcoCalendar API.

The implementation now uses the local HTTP client. This module is kept so
existing scripts importing ``EcoCalendar`` do not break immediately. New code
should import ``InvestingAPIClient`` from ``investing_calendar``.
"""

from datetime import datetime

import pandas as pd

from investing_calendar import InvestingAPIClient


class EcoCalendar:
    """Backward-compatible facade over :class:`InvestingAPIClient`."""

    available_categories = [
        "credit",
        "inflation",
        "employment",
        "activity",
        "economic_activity",
        "central_banks",
        "balance",
        "bonds",
    ]

    def __init__(self, dt_format="%d/%m/%Y", client=None):
        self.dt_format = dt_format
        self.client = client or InvestingAPIClient()

    def get_economic_calendar(
        self,
        from_date: datetime,
        to_date: datetime,
        categories: list | None = None,
        countries: list | None = None,
    ) -> pd.DataFrame:
        normalized_categories = [
            "economic_activity" if category == "activity" else category
            for category in (categories or [])
        ] or None
        return self.client.get_economic_calendar(
            from_date=from_date,
            to_date=to_date,
            categories=normalized_categories,
            countries=countries,
        )
