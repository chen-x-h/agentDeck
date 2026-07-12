import os
import json
import yaml
from pathlib import Path
from pptx import Presentation as PptxPresentation
from lxml import etree
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from ppt_render_engine.log_config import app_logger

logger = app_logger("core.color_scheme")
_cfg_path = Path(__file__).parents[3] / "config.yaml"


class ColorSchemeNotFoundError(KeyError):
    pass


class ColorSchemeManager:
    def __init__(self):
        self._schemes: dict[str, dict[str, str]] = {}
        self._default_name: str = "经典黑"

    def get_storage_path(self) -> str:
        try:
            with open(_cfg_path, encoding="utf-8") as f:
                cfg = yaml.safe_load(f) or {}
            path = cfg.get("color_scheme", {}).get("storage_path", "./storage/color_schemes/color_schemes.json")
        except Exception:
            path = "./storage/color_schemes/color_schemes.json"
        return os.path.abspath(path)

    def load(self, filepath: str | None = None) -> None:
        path = filepath or self.get_storage_path()
        if not os.path.isfile(path):
            logger.warning("Color scheme file not found", path=path)
            return
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        self._default_name = data.get("default", "经典黑")
        self._schemes.clear()
        for s in data.get("schemes", []):
            self._schemes[s["name"]] = s["colors"]
        logger.info("Color schemes loaded", count=len(self._schemes), default=self._default_name)

    def save(self, filepath: str | None = None) -> None:
        path = filepath or self.get_storage_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        data = {
            "default": self._default_name,
            "schemes": [{"name": name, "colors": colors} for name, colors in self._schemes.items()],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        logger.info("Color schemes saved", count=len(self._schemes))

    def list_schemes(self) -> list[str]:
        return list(self._schemes.keys())

    def get_colors(self, name: str) -> dict[str, str]:
        if name not in self._schemes:
            raise ColorSchemeNotFoundError(name)
        return dict(self._schemes[name])

    def get_default_name(self) -> str:
        return self._default_name

    def get_default_colors(self) -> dict[str, str]:
        return self.get_colors(self._default_name)

    def add_scheme(self, name: str, colors: dict[str, str]) -> None:
        self._schemes[name] = dict(colors)
        logger.info("Color scheme added", name=name)

    def remove_scheme(self, name: str) -> None:
        if name not in self._schemes:
            raise ColorSchemeNotFoundError(name)
        del self._schemes[name]
        if self._default_name == name:
            self._default_name = next(iter(self._schemes)) if self._schemes else "经典黑"
        logger.info("Color scheme removed", name=name)

    def set_default(self, name: str) -> None:
        if name not in self._schemes:
            raise ColorSchemeNotFoundError(name)
        self._default_name = name
        logger.info("Default color scheme set", name=name)

    def get_color_name(self, hex_value: str, scheme_name: str | None = None) -> str | None:
        """Find the scheme color name (dk1, accent1, etc.) matching a hex value, or None."""
        if scheme_name is None:
            scheme_name = self._default_name
        scheme = self._schemes.get(scheme_name)
        if scheme is None:
            return None
        target = hex_value.upper().lstrip("#")
        for name, val in scheme.items():
            if val.upper().lstrip("#") == target:
                return name
        return None

    def apply_to_presentation(self, prs: PptxPresentation, scheme_name: str | None) -> None:
        if scheme_name is None:
            return
        if scheme_name not in self._schemes:
            logger.warning("Color scheme not found, leaving template theme untouched", requested=scheme_name)
            return
        name = scheme_name
        colors = self._schemes[name]
        # find the theme part via the first slide master
        if not prs.slide_masters:
            logger.warning("No slide masters, cannot apply color scheme")
            return
        sm_part = prs.slide_masters[0].part
        theme_part = None
        for rel in sm_part.rels.values():
            if rel.reltype == RT.THEME:
                theme_part = rel.target_part
                break
        if theme_part is None:
            logger.warning("No theme part found, cannot apply color scheme")
            return
        theme = etree.fromstring(theme_part.blob)
        clrScheme = theme.find('.//' + qn('a:clrScheme'))
        if clrScheme is None:
            logger.warning("No clrScheme in theme, cannot apply color scheme")
            return
        for color_name, color_value in colors.items():
            elem = clrScheme.find(qn(f'a:{color_name}'))
            if elem is not None:
                for child in list(elem):
                    elem.remove(child)
                srgb = etree.SubElement(elem, qn('a:srgbClr'))
                srgb.set('val', color_value.upper())
        theme_part._blob = etree.tostring(theme, xml_declaration=True, encoding='UTF-8', standalone=True)
        logger.info("Color scheme applied", name=name)


_scheme_manager: ColorSchemeManager | None = None


def get_color_scheme_manager() -> ColorSchemeManager:
    global _scheme_manager
    if _scheme_manager is None:
        _scheme_manager = ColorSchemeManager()
        _scheme_manager.load()
        logger.debug("ColorSchemeManager singleton created")
    return _scheme_manager
