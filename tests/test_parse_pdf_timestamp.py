from datetime import datetime
from pathlib import Path
import sys
import types

from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

try:  # pragma: no cover - exercised indirectly in tests
    import pandas  # type: ignore  # noqa: F401
except ImportError:  # pragma: no cover - fallback used when pandas no disponible
    pd_stub = types.ModuleType("pandas")
    pd_stub.DataFrame = lambda *args, **kwargs: None  # type: ignore[assignment]
    pd_stub.to_excel = lambda *args, **kwargs: None  # type: ignore[assignment]
    sys.modules.setdefault("pandas", pd_stub)

from tickets_parser import parse_pdf_timestamp


def test_prefers_mmdd_when_closer_to_now():
    tz = ZoneInfo("UTC")
    now = datetime(2024, 3, 6, 12, 0, tzinfo=tz)

    parsed = parse_pdf_timestamp("03/05/2024 10:00 AM", now, tz)

    assert parsed == datetime(2024, 3, 5, 10, 0, tzinfo=tz)


def test_prefers_ddmm_when_future_option_is_reasonable():
    tz = ZoneInfo("UTC")
    now = datetime(2024, 5, 4, 12, 0, tzinfo=tz)

    parsed = parse_pdf_timestamp("03/05/2024 10:00 AM", now, tz)

    assert parsed == datetime(2024, 5, 3, 10, 0, tzinfo=tz)


def test_returns_none_when_both_candidates_are_far_future():
    tz = ZoneInfo("UTC")
    now = datetime(2024, 1, 1, 0, 0, tzinfo=tz)

    parsed = parse_pdf_timestamp("12/11/2099 10:00 AM", now, tz)

    assert parsed is None
