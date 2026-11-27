import sys
from pathlib import Path

import pytest
import requests

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))


@pytest.fixture(autouse=True)
def block_network(monkeypatch):
    """Fail any test that tries to reach the internet through requests.get."""
    def blocked_get(*args, **kwargs):
        raise RuntimeError("Network access is disabled in tests")

    monkeypatch.setattr(requests, "get", blocked_get)
