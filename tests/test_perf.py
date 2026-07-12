"""Performance benchmarks: measure build + parse time per fixture."""

import json
import time
import pytest
from ppt_render_engine.models.schema import Presentation
from ppt_render_engine.core.pptx_builder import build_pptx
from ppt_render_engine.core.pptx_parser import parse_pptx
from conftest import load_fixture, list_fixtures, temp_output

# Fixtures excluded from perf test (network-dependent or template-dependent)
SKIP_PERF = {"image_url", "template_placeholder"}


def _build(pres, path):
    build_pptx(pres, path)


def _parse(path):
    return parse_pptx(path)


def _multiply_fixture(data, n: int) -> dict:
    """Duplicate slides to create an n-page presentation."""
    out = dict(data)
    slides = data.get("slides", [])
    if not slides:
        slides = [{"id": 0, "shapes": []}]
    multiplied = []
    for i in range(n):
        s = dict(slides[i % len(slides)])
        s["id"] = i
        multiplied.append(s)
    out["slides"] = multiplied
    return out


@pytest.mark.parametrize("fixture_name,fixture_path", list_fixtures(), ids=lambda x: x[0])
def test_build_single(fixture_name, fixture_path, temp_output):
    if fixture_name in SKIP_PERF:
        pytest.skip("Skipped in perf test")
    data = load_fixture(f"{fixture_name}.json")
    pres = Presentation(**data)
    slide_count = len(data.get("slides", []))

    # warmup: 1 run
    build_pptx(pres, temp_output)
    parse_pptx(temp_output)

    # timed runs
    N = 5
    build_times = []
    parse_times = []
    for _ in range(N):
        t0 = time.perf_counter()
        build_pptx(pres, temp_output)
        build_times.append(time.perf_counter() - t0)

        t0 = time.perf_counter()
        parse_pptx(temp_output)
        parse_times.append(time.perf_counter() - t0)

    avg_build = sum(build_times) / N
    avg_parse = sum(parse_times) / N
    avg_total = avg_build + avg_parse
    per_slide_build = avg_build / slide_count
    per_slide_parse = avg_parse / slide_count
    per_slide_total = avg_total / slide_count

    print(f"\n  [{fixture_name}] {slide_count} slide(s)")
    print(f"    Build  : {avg_build*1000:.1f}ms  ({per_slide_build*1000:.1f}ms/slide)")
    print(f"    Parse  : {avg_parse*1000:.1f}ms  ({per_slide_parse*1000:.1f}ms/slide)")
    print(f"    Roundtrip: {avg_total*1000:.1f}ms  ({per_slide_total*1000:.1f}ms/slide)")

    # Accumulate results on the module
    if not hasattr(test_build_single, "results"):
        test_build_single.results = []
    test_build_single.results.append({
        "fixture": fixture_name,
        "slides": slide_count,
        "build_ms": avg_build * 1000,
        "parse_ms": avg_parse * 1000,
        "total_ms": avg_total * 1000,
        "per_slide_ms": per_slide_total * 1000,
    })


def test_synthetic_multipage(temp_output, request):
    """Test a synthetic multi-page presentation (1, 5, 10, 20 pages)."""
    data = load_fixture("simple_text.json")
    sizes = [1, 5, 10, 20]
    # collect results
    results = []

    for n in sizes:
        multiplied = _multiply_fixture(data, n)
        pres = Presentation(**multiplied)

        # warmup
        build_pptx(pres, temp_output)
        parse_pptx(temp_output)

        N = 5
        b_times = []
        p_times = []
        for _ in range(N):
            t0 = time.perf_counter()
            build_pptx(pres, temp_output)
            b_times.append(time.perf_counter() - t0)

            t0 = time.perf_counter()
            parse_pptx(temp_output)
            p_times.append(time.perf_counter() - t0)

        avg_b = sum(b_times) / N
        avg_p = sum(p_times) / N
        avg_t = avg_b + avg_p

        results.append({
            "pages": n,
            "build_ms": avg_b * 1000,
            "parse_ms": avg_p * 1000,
            "total_ms": avg_t * 1000,
            "per_slide_ms": avg_t / n * 1000,
        })

    print("\n\n  --- Synthetic Multi-Page Benchmark (simple_text fixture) ---")
    for r in results:
        print(f"  {r['pages']:2d} pages | Build {r['build_ms']:.1f}ms Parse {r['parse_ms']:.1f}ms "
              f"Roundtrip {r['total_ms']:.1f}ms ({r['per_slide_ms']:.1f}ms/page)")

    # Verify O(n) scaling: 20-page should not be >3x 10-page total time
    r10 = next(r["total_ms"] for r in results if r["pages"] == 10)
    r20 = next(r["total_ms"] for r in results if r["pages"] == 20)
    ratio = r20 / r10 if r10 > 0 else 0
    assert ratio < 3.0, f"Scaling anomaly: 20-page time / 10-page = {ratio:.2f} (expected < 3.0)"

    if not hasattr(test_build_single, "results"):
        test_build_single.results = []


def test_perf_summary(request):
    if not hasattr(test_build_single, "results"):
        pytest.skip("No results collected")

    results = test_build_single.results
    total_slides = sum(r["slides"] for r in results)
    total_build = sum(r["build_ms"] for r in results)
    total_parse = sum(r["parse_ms"] for r in results)
    avg_build_per_slide = total_build / max(total_slides, 1)
    avg_parse_per_slide = total_parse / max(total_slides, 1)
    avg_roundtrip_per_slide = (total_build + total_parse) / max(total_slides, 1)

    print("\n" + "=" * 65)
    print("  PERF SUMMARY (13 fixtures × 5 runs each)")
    print("=" * 65)
    for r in sorted(results, key=lambda x: x["fixture"]):
        print(f"  {r['fixture']:25s} {r['slides']:2d} slides  "
              f"build {r['build_ms']:7.1f}ms  parse {r['parse_ms']:7.1f}ms  "
              f"total {r['total_ms']:7.1f}ms")
    print(f"  {'-' * 62}")
    print(f"  Average per slide:")
    print(f"    Build    : {avg_build_per_slide:.1f}ms")
    print(f"    Parse    : {avg_parse_per_slide:.1f}ms")
    print(f"    Roundtrip: {avg_roundtrip_per_slide:.1f}ms")
    print("=" * 65)
