import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from macrocalendar.config import AppConfig
from macrocalendar.outlook import OutlookClient
from investing_calendar import InvestingAPIClient, _iso_range_from_ddmmyyyy, _map_countries_to_ids


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def raise_for_status(self):
        return None

    def json(self):
        return self.payload


class FakeSession:
    def __init__(self):
        self.url = None
        self.params = None

    def get(self, url, **kwargs):
        self.url = url
        self.params = kwargs["params"]
        return FakeResponse(
            {
                "events": [
                    {
                        "event_id": 10,
                        "short_name": "CPI",
                        "event_cycle_suffix": "YoY",
                        "reference_period": "Aug",
                        "country_id": 32,
                        "currency": "BRL",
                        "importance": "high",
                    }
                ],
                "occurrences": [
                    {
                        "event_id": 10,
                        "occurrence_time": "2026-09-17T15:00:00Z",
                        "actual": None,
                        "forecast": "4.0",
                        "previous": "3.9",
                        "unit": "%",
                    }
                ],
            }
        )


class MacroCalendarTests(unittest.TestCase):
    def test_date_and_country_helpers(self):
        self.assertEqual(_map_countries_to_ids(["Brazil", "United States"]), [32, 5])
        self.assertEqual(
            _iso_range_from_ddmmyyyy("17/09/2026", "18/09/2026", "GMT -3:00"),
            ("2026-09-17T00:00:00.000-03:00", "2026-09-18T23:59:59.999-03:00"),
        )

    def test_http_client_normalizes_api_response(self):
        session = FakeSession()
        client = InvestingAPIClient(session=session)
        result = client.get_economic_calendar(
            countries=["brazil"],
            categories=["inflation"],
            from_date="17/09/2026",
            to_date="18/09/2026",
        )

        self.assertEqual(result.iloc[0]["zone"], "brazil")
        self.assertEqual(result.iloc[0]["event"], "CPI YoY (Aug)")
        self.assertEqual(result.iloc[0]["time"], "12:00")
        self.assertEqual(session.params["country_ids"], "32")
        self.assertEqual(session.params["categories"], "inflation")

    def test_config_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "config.json"
            original = AppConfig(calendar_name="Macro", days=5)
            original.save(path)
            loaded = AppConfig.load(path)
            self.assertEqual(loaded.calendar_name, "Macro")
            self.assertEqual(loaded.days, 5)
            self.assertFalse(loaded.verify_tls)
            self.assertEqual(json.loads(path.read_text(encoding="utf-8"))["days"], 5)

    def test_event_pipeline_honors_category_and_country_configuration(self):
        import calendar_events

        calls = []

        class CalendarStub:
            def get_economic_calendar(self, **kwargs):
                calls.append(kwargs)
                return FakeEmptyDataFrame()

        with patch.object(calendar_events, "calendar", CalendarStub()):
            result = calendar_events.get_and_filter_events(
                1,
                include_holidays=False,
                countries=["brazil"],
                categories=["inflation"],
            )

        self.assertTrue(result.empty)
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0]["categories"], ["inflation"])
        self.assertEqual(calls[0]["countries"], ["brazil"])

    def test_outlook_calendar_lookup_accepts_full_folder_path(self):
        class Folder:
            def __init__(self, name, folders=None):
                self.Name = name
                self.Folders = folders or []

        target = Folder("Macro")
        root = Folder("Calendar", [Folder("Trading", [target])])
        client = OutlookClient("Calendar / Trading / Macro")
        self.assertIs(client._find_calendar(root, "Calendar / Trading / Macro"), target)


class FakeEmptyDataFrame:
    """Minimal empty event frame for the pipeline test."""

    def __init__(self):
        import pandas as pd

        self._frame = pd.DataFrame(
            columns=[
                "id", "date", "time", "zone", "currency", "importance",
                "event", "actual", "forecast", "previous",
            ]
        )

    def __getattr__(self, name):
        return getattr(self._frame, name)


if __name__ == "__main__":
    unittest.main()
