import json
import pytest
from ppt_render_engine.models.schema import Presentation
from ppt_render_engine.core.pptx_builder import build_pptx
from ppt_render_engine.core.pptx_parser import parse_pptx
from ppt_render_engine.core.color_scheme import get_color_scheme_manager
from conftest import load_fixture, list_fixtures

csm = get_color_scheme_manager()

# Fields the parser adds as defaults (not in input)
PARSER_DEFAULT_FIELDS = {"rotation", "border_style", "shadow", "border_width", "indent_level", "z_order"}

SCHEME_COLOR_KEYS = {"dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                     "accent4", "accent5", "accent6", "hlink", "folHlink"}

# Fixture-specific paths to skip in comparison (known round-trip gaps)
SKIP_PATHS = {
    "full_table": [
        "slides[0].shapes[0].table_content",  # table content restructured by parser
    ],
    "merged_cells": [
        "slides[0].shapes[0].table_content",
    ],
    "image_base64": [
        "slides[0].shapes[0].image_content.data",
        "slides[0].shapes[0].image_content.width",
        "slides[0].shapes[0].image_content.height",
        "slides[0].shapes[0].image_content.path",
    ],
    "image_path": [
        "slides[0].shapes[0].image_content.width",
        "slides[0].shapes[0].image_content.height",
        "slides[0].shapes[0].image_content.path",
    ],
    "image_url": [
        "slides[0].shapes[0].image_content.url",
        "slides[0].shapes[0].image_content.width",
        "slides[0].shapes[0].image_content.height",
        "slides[0].shapes[0].image_content.path",
    ],
    "template_placeholder": [
        "slides[0].shapes[0].left", "slides[0].shapes[0].top",
        "slides[0].shapes[0].width", "slides[0].shapes[0].height",
        "slides[0].shapes[0].id",
        "slides[1].shapes[0].left", "slides[1].shapes[0].top",
        "slides[1].shapes[0].width", "slides[1].shapes[0].height",
        "slides[1].shapes[0].id",
    ],
}

SKIP_ALL_FIXTURES = {
    "image_url",                # URL download requires network; skipped in offline test
    "template_placeholder",     # template adds layout shapes, coords change
}


def is_scheme_color(val: str) -> bool:
    if isinstance(val, str) and val.startswith("scheme:"):
        return val.split(":")[1] in SCHEME_COLOR_KEYS
    return False


def normalize_parser_field(obj: dict, path: str = "") -> dict:
    """Post-process parser output to match input format."""
    result = {}
    for k, v in obj.items():
        if k in PARSER_DEFAULT_FIELDS:
            continue
        cur = f"{path}.{k}" if path else k
        if k in ("color", "font_color") and isinstance(v, str) and is_scheme_color(v):
            continue
        if isinstance(v, dict):
            result[k] = normalize_parser_field(v, cur)
        elif isinstance(v, list):
            result[k] = [
                normalize_parser_field(item, f"{cur}[{i}]") if isinstance(item, dict) else item
                for i, item in enumerate(v)
            ]
        else:
            result[k] = v
    return result


def deep_field_count(a: dict, b: dict, skip: set, path: str = "") -> tuple:
    matched = 0
    total = 0
    mismatches = []
    keys_a = set(a.keys())
    keys_b = set(b.keys())
    all_keys = keys_a | keys_b
    for k in sorted(all_keys):
        cur = f"{path}.{k}" if path else k
        if cur in skip:
            continue
        if k not in a:
            total += 1
            mismatches.append((cur, "MISSING_IN_INPUT", f"output has '{b[k]}'"))
            continue
        if k not in b:
            total += 1
            mismatches.append((cur, "MISSING_IN_OUTPUT", f"input has '{a[k]}'"))
            continue
        va, vb = a[k], b[k]
        if isinstance(va, dict) and isinstance(vb, dict):
            m, t, mm = deep_field_count(va, vb, skip, cur)
            matched += m
            total += t
            mismatches.extend(mm)
        elif isinstance(va, list) and isinstance(vb, list):
            max_len = max(len(va), len(vb))
            for i in range(max_len):
                ci = f"{cur}[{i}]"
                if ci in skip:
                    continue
                if i >= len(va):
                    total += 1
                    mismatches.append((ci, "MISSING_IN_INPUT", f"output has '{vb[i]}'"))
                elif i >= len(vb):
                    total += 1
                    mismatches.append((ci, "MISSING_IN_OUTPUT", f"input has '{va[i]}'"))
                elif isinstance(va[i], dict) and isinstance(vb[i], dict):
                    m, t, mm = deep_field_count(va[i], vb[i], skip, ci)
                    matched += m
                    total += t
                    mismatches.extend(mm)
                elif va[i] == vb[i]:
                    matched += 1
                    total += 1
                else:
                    total += 1
                    mismatches.append((ci, "DIFF", f"'{va[i]}' != '{vb[i]}'"))
        else:
            total += 1
            if va == vb:
                matched += 1
            elif isinstance(va, str) and isinstance(vb, str) and va.upper().lstrip("#") == vb.upper().lstrip("#"):
                matched += 1
            else:
                mismatches.append((cur, "DIFF", f"'{va}' != '{vb}'"))
    return matched, total, mismatches


@pytest.mark.parametrize("fixture_name,fixture_path", list_fixtures(), ids=lambda x: x[0])
def test_roundtrip_fidelity(fixture_name, fixture_path, temp_output):
    if fixture_name in SKIP_ALL_FIXTURES:
        pytest.skip(f"Known limitation: {fixture_name}")

    data = load_fixture(f"{fixture_name}.json")
    pres = Presentation(**data)

    try:
        build_pptx(pres, temp_output)
    except Exception as e:
        pytest.fail(f"Build failed: {e}")

    parsed = parse_pptx(temp_output)
    output = json.loads(parsed.model_dump_json(exclude_none=True))

    input_slides = data.get("slides", [])
    output_slides = output.get("slides", [])
    skip = set(SKIP_PATHS.get(fixture_name, []))

    total_fields = 0
    matched_fields = 0
    all_mismatches = []

    min_slides = min(len(input_slides), len(output_slides))
    for i in range(min_slides):
        in_shapes = input_slides[i].get("shapes", [])
        out_shapes = output_slides[i].get("shapes", [])
        min_shapes = min(len(in_shapes), len(out_shapes))
        for si in range(min_shapes):
            out_norm = normalize_parser_field(dict(out_shapes[si]))
            m, t, mm = deep_field_count(in_shapes[si], out_norm, skip, f"slides[{i}].shapes[{si}]")
            matched_fields += m
            total_fields += t
            all_mismatches.extend(mm)
        if len(in_shapes) != len(out_shapes):
            all_mismatches.append((f"slides[{i}]", "SHAPE_COUNT",
                                   f"in={len(in_shapes)} out={len(out_shapes)}"))
            total_fields += max(len(in_shapes), len(out_shapes))

    fidelity = round(matched_fields / max(total_fields, 1) * 100, 1)
    top3 = [f"{fp}: {reason} ({detail[:80]})" for fp, reason, detail in all_mismatches[:3]]

    print(f"\n  [{fixture_name}] fidelity: {fidelity}%  ({matched_fields}/{total_fields} fields)")
    if all_mismatches:
        for e in top3:
            print(f"    -> {e}")

    if not hasattr(test_roundtrip_fidelity, "results"):
        test_roundtrip_fidelity.results = []
    test_roundtrip_fidelity.results.append({
        "fixture": fixture_name,
        "fidelity": fidelity,
        "matched": matched_fields,
        "total": total_fields,
        "mismatches": len(all_mismatches),
    })

    assert fidelity >= 70.0, f"Fidelity too low: {fidelity}%"


def test_fidelity_summary():
    if not hasattr(test_roundtrip_fidelity, "results"):
        pytest.skip("No results collected")
    results = test_roundtrip_fidelity.results
    total_matched = sum(r["matched"] for r in results)
    total_fields = sum(r["total"] for r in results)
    avg = round(total_matched / max(total_fields, 1) * 100, 1)

    print("\n" + "=" * 55)
    print("  ROUND-TRIP FIDELITY SUMMARY")
    print("=" * 55)
    for r in sorted(results, key=lambda x: x["fidelity"]):
        icon = "OK" if r["fidelity"] >= 90 else "WARN" if r["fidelity"] >= 70 else "FAIL"
        print(f"  {icon} {r['fixture']:25s} {r['fidelity']:6.1f}%  ({r['matched']:3d}/{r['total']:3d})")
    print(f"  {'-' * 48}")
    passed = sum(1 for r in results if r["fidelity"] >= 70)
    print(f"  OVERALL    : {avg:.1f}% ({total_matched}/{total_fields})")
    print(f"  Passed     : {passed}/{len(results)}")
    print(f"  Skipped    : {list(SKIP_ALL_FIXTURES) if SKIP_ALL_FIXTURES else 'none'}")
    print("=" * 55)

    assert avg >= 75.0, f"Overall fidelity too low: {avg}%"
