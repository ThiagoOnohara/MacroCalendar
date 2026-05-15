# -*- coding: utf-8 -*-
"""
Created on Mon Jan 19 09:54:42 2026

@author: thiago.onohara
"""

from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple, Union

import pandas as pd
import requests
from requests.adapters import HTTPAdapter

try:
    from urllib3.util.retry import Retry
except Exception:  # pragma: no cover
    Retry = None


# =============================================================================
# Constants (from your investpy-like snippet)
# =============================================================================
from investing_constants import (
    COUNTRY_ID_FILTERS,
    CATEGORY_FILTERS,
    IMPORTANCE_RATINGS,
    TIME_FILTERS,
)


# =============================================================================
# Helpers
# =============================================================================

def _norm(s: str) -> str:
    return " ".join(s.strip().lower().split())


def _as_csv(values: Optional[Iterable[Union[str, int]]]) -> Optional[str]:
    if not values:
        return None
    return ",".join(str(v) for v in values)


def _tz_to_offset(time_zone: Optional[str]) -> str:
    """
    Aceita:
      - None -> "-03:00" (padrão BR; ajuste se quiser)
      - "-03:00" / "+00:00"
      - "GMT -3:00" (estilo investpy)
    Retorna sempre "±HH:MM".
    """
    if time_zone is None:
        return "-03:00"

    tz = time_zone.strip()
    if tz.startswith(("+", "-")) and len(tz) == 6 and tz[3] == ":":
        return tz

    if tz.upper().startswith("GMT"):
        rest = tz[3:].strip()  # "-3:00" / "+1:00" etc.
        if not rest:
            return "+00:00"
        sign = "+" if rest[0] == "+" else "-"
        hh, mm = rest[1:].strip().split(":")
        return f"{sign}{int(hh):02d}:{int(mm):02d}"

    return tz


def _iso_range_from_ddmmyyyy(
    from_date: Optional[str],
    to_date: Optional[str],
    tz_offset: str,
) -> Tuple[str, str]:
    """
    Converte dd/mm/yyyy em intervalo ISO com offset:
      start_date = YYYY-mm-ddT00:00:00.000±HH:MM
      end_date   = YYYY-mm-ddT23:59:59.999±HH:MM
    Se None/None: usa 'hoje'.
    """
    off = _tz_to_offset(tz_offset)

    def parse_d(s: str) -> date:
        return datetime.strptime(s, "%d/%m/%Y").date()

    if from_date is None and to_date is None:
        d0 = date.today()
        d1 = d0
    else:
        if from_date is None or to_date is None:
            raise ValueError("Se passar 'from_date', passe também 'to_date' (dd/mm/yyyy).")
        
        if isinstance(from_date, datetime):
            from_date = datetime.strftime(from_date, '%d/%m/%Y')
        if isinstance(to_date, datetime):
            to_date = datetime.strftime(to_date, '%d/%m/%Y')
            
        d0 = parse_d(from_date)
        d1 = parse_d(to_date)
        if d0 > d1:
            raise ValueError("to_date deve ser >= from_date.")

    start = f"{d0.isoformat()}T00:00:00.000{off}"
    end = f"{d1.isoformat()}T23:59:59.999{off}"
    return start, end


def _map_countries_to_ids(countries: Optional[Sequence[str]]) -> Optional[List[int]]:
    if not countries:
        return None
    ids: List[int] = []
    for c in countries:
        key = _norm(c)
        v = COUNTRY_ID_FILTERS.get(key)
        if v is not None:
            ids.append(int(v))
    return ids or None


def _map_categories_to_api_slugs(categories: Optional[Sequence[str]]) -> Optional[List[str]]:
    """
    Mantive seu comportamento atual (passa slug direto),
    mas deixei o CATEGORY_FILTERS disponível caso você queira voltar a usar.
    """
    if not categories:
        return None

    out: List[str] = []
    for cat in categories:
        key = _norm(cat)

        # Você tinha forçado "mapped = key" (mantive)
        mapped = key  # ou: CATEGORY_FILTERS.get(key, key)
        slug = str(mapped).lstrip("_")
        out.append(slug)

    return out or None


def _build_requests_session(
    *,
    total_retries: int = 5,
    backoff_factor: float = 0.4,
    status_forcelist: Sequence[int] = (429, 500, 502, 503, 504),
) -> requests.Session:
    s = requests.Session()
    if Retry is None:
        return s

    retry = Retry(
        total=total_retries,
        connect=total_retries,
        read=total_retries,
        status=total_retries,
        backoff_factor=backoff_factor,
        status_forcelist=set(status_forcelist),
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=50)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


# =============================================================================
# Client
# =============================================================================

@dataclass(frozen=True)
class InvestingApiConfig:
    domain_id: int = 1
    timeout_s: int = 30
    verify_tls: bool = False  # mantém o default que você já vinha usando no wrapper
    user_agent: str = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )


class DataFrameParser:
    DEFAULT_TZ: str = "America/Sao_Paulo"

    @staticmethod
    def _parse_occurrence_time_sp_naive(
        s: pd.Series,
        tz: str = "America/Sao_Paulo",
    ) -> pd.Series:
        print("--- Parsing Occurrence Time -> Datetime ---")
        # utc=True cria tz-aware em UTC; depois tz_convert e remove tz (naive) :contentReference[oaicite:1]{index=1}
        return (
            pd.to_datetime(s, utc=True, errors="coerce")
            .dt.tz_convert(tz)
            .dt.tz_localize(None)
        )

    @staticmethod
    def _parse_country_id(s: pd.Series) -> pd.Series:
        print("--- Parsing Country ID -> Name ---")
        reversed_map = dict(zip(COUNTRY_ID_FILTERS.values(), COUNTRY_ID_FILTERS.keys()))
        return s.map(reversed_map)

    @staticmethod
    def _parse_event_name(df_: pd.DataFrame) -> pd.Series:
        print("--- Parsing Event Name ---")
        
        if 'event_cycle_suffix' in df_.columns:
            s = (
                df_["short_name"]
                + " "
                + df_["event_cycle_suffix"].fillna("")
                + " "
                + df_["reference_period"].apply(lambda r: f"({r})" if isinstance(r, str) else "")
            ).str.strip()
            return s
        
        else:
            print('"event_cycle_suffix" not in df_.columns')
            print('dataframe columns: ', df_.columns)
            s = (
                df_["short_name"]
                + " "
                + df_["reference_period"].apply(lambda r: f"({r})" if isinstance(r, str) else "")
            ).str.strip()
            return s
            
class InvestingAPIClient:
    BASE_URL = "https://endpoints.investing.com/pd-instruments/v1"

    def __init__(self, config: Optional[InvestingApiConfig] = None, session: Optional[requests.Session] = None):
        self.config = config or InvestingApiConfig()
        self.session = session or _build_requests_session()

    def _headers(self, extra: Optional[Mapping[str, str]] = None) -> Dict[str, str]:
        h = {
            "Accept": "*/*",
            "Origin": "https://www.investing.com",
            "Referer": "https://www.investing.com/",
            "User-Agent": self.config.user_agent,
        }
        if extra:
            h.update(dict(extra))
        return h

    def get_economic_occurrences(
        self,
        *,
        start_date_iso: str,
        end_date_iso: str,
        country_ids: Optional[Sequence[int]] = None,
        categories: Optional[Sequence[str]] = None,
        limit: int = 200,
        extra_params: Optional[Mapping[str, Any]] = None,
        headers: Optional[Mapping[str, str]] = None,
    ) -> Dict[str, Any]:
        url = f"{self.BASE_URL}/calendars/economic/events/occurrences"

        params: Dict[str, Any] = {
            "domain_id": str(self.config.domain_id),
            "limit": str(int(limit)),
            "start_date": start_date_iso,
            "end_date": end_date_iso,
        }
        if country_ids:
            params["country_ids"] = _as_csv(country_ids)
        if categories:
            params["categories"] = _as_csv(categories)
        if extra_params:
            params.update(dict(extra_params))

        resp = self.session.get(
            url,
            headers=self._headers(headers),
            params=params,
            timeout=self.config.timeout_s,
            verify=self.config.verify_tls,
        )
        resp.raise_for_status()
        return resp.json()

    # =============================================================================
    # NEW USAGE: investing = InvestingAPIClient(); investing.economic_calendar(...)
    # =============================================================================
    def get_economic_calendar(
        self,
        time_zone: Optional[str] = "GMT -3:00",
        time_filter: str = "time_only",  # compat
        countries: Optional[List[str]] = None,
        importances: Optional[List[str]] = None,  # (ainda não usado; mantido p/ compat)
        categories: Optional[List[str]] = None,
        from_date: Optional[str] = None,  # dd/mm/yyyy
        to_date: Optional[str] = None,    # dd/mm/yyyy
        *,
        limit: int = 300,
        extra_params: Optional[Mapping[str, Any]] = None,
    ) -> pd.DataFrame:
        if time_filter not in TIME_FILTERS:
            raise ValueError(f"time_filter inválido: {time_filter}. Use {list(TIME_FILTERS.keys())}")

        start_iso, end_iso = _iso_range_from_ddmmyyyy(
            from_date, to_date, tz_offset=_tz_to_offset(time_zone)
        )

        country_ids = _map_countries_to_ids(countries)
        category_slugs = _map_categories_to_api_slugs(categories)

        raw = self.get_economic_occurrences(
            start_date_iso=start_iso,
            end_date_iso=end_iso,
            country_ids=country_ids,
            categories=category_slugs,
            limit=limit,
            extra_params=extra_params,
        )

        print("TOTAL EVENTS (dict keys): ", len(raw), sep='\n'*3)
        print('EVENTS: ', pd.json_normalize(raw['events']), sep='\n'*3)
        print('OCCURENCES: ', pd.json_normalize(raw['occurrences']), sep='\n'*3)

        raw_event_df = pd.json_normalize(raw["events"])
        raw_occ_df = pd.json_normalize(raw["occurrences"])

        # Merge Events and Occurrences
        df = raw_event_df.merge(raw_occ_df, on=["event_id"])

        # Parsing
        df["id"] = df["event_id"]
        df["datetime"] = DataFrameParser._parse_occurrence_time_sp_naive(df["occurrence_time"])
        df["date"] = df["datetime"].dt.strftime("%d/%m/%Y")
        df["time"] = df["datetime"].dt.strftime("%H:%M")
        df["zone"] = DataFrameParser._parse_country_id(df["country_id"])
        df["event"] = DataFrameParser._parse_event_name(df)
        if 'actual' in df.columns:
            df['actual'] = df['actual'].fillna('').astype(str) + df['unit'].fillna('').astype(str)
        base_cols = pd.Index(["id", "date", "time", "zone", "currency", "importance", "event", "actual", "forecast", "previous"])
        matched_cols = df.columns.intersection(base_cols)
        unmatched_cols = df.columns.difference(base_cols)

        print('MATCHED COLUMNS: ', matched_cols)
        print('UNMATCHED COLUMNS: ', unmatched_cols)
        return df[matched_cols]


# =============================================================================
# Example
# =============================================================================
if __name__ == "__main__":
    investing = InvestingAPIClient()  # <- novo modo de uso

    df = investing.get_economic_calendar(
        countries=[
            "brazil", "united states", "euro zone", "canada", "united kingdom", "canada",
            "china", "sweden", "norway", "south korea", "mexico", "chile", "colombia",
            "south africa", "czech republic", "hungary", "poland", "singapore", "japan",
            "switzerland", "france", "spain", "italy"
        ],
        categories=["inflation", "employment", "central_banks"],
        from_date="19/01/2026",
        to_date="20/01/2026",
    )

    print(df.head(30))
