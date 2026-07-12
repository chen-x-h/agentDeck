import json
import os
from pathlib import Path
import pytest
from ppt_render_engine.core.template import get_template_manager
from ppt_render_engine.core.color_scheme import get_color_scheme_manager
from ppt_render_engine.temp_util import get_temp_path

FIXTURE_DIR = Path(__file__).parent / "fixtures"


def pytest_configure(config):
    tm = get_template_manager()
    tm.preload_all()
    get_color_scheme_manager()


def load_fixture(name: str) -> dict:
    path = FIXTURE_DIR / name
    if not path.exists():
        raise FileNotFoundError(f"Fixture not found: {path}")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def list_fixtures() -> list[tuple[str, str]]:
    items = []
    for f in sorted(FIXTURE_DIR.glob("*.json")):
        items.append((f.stem, str(f)))
    return items


@pytest.fixture
def temp_output():
    path = get_temp_path(".pptx")
    yield path
    if os.path.isfile(path):
        os.unlink(path)
