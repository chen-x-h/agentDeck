import os
import json
import yaml
from pathlib import Path
from ppt_render_engine.log_config import app_logger

logger = app_logger("core.design_style")
_cfg_path = Path(__file__).parents[3] / "config.yaml"


class DesignStyleNotFoundError(KeyError):
    pass


class DesignStyleManager:
    def __init__(self):
        self._styles: dict[str, dict] = {}
        self._default_name: str = ""

    def get_storage_dir(self) -> str:
        try:
            with open(_cfg_path, encoding="utf-8") as f:
                cfg = yaml.safe_load(f) or {}
            path = cfg.get("design_style", {}).get("storage_path", "./storage/design_styles")
        except Exception:
            path = "./storage/design_styles"
        return os.path.abspath(path)

    def load(self, directory: str | None = None) -> None:
        dir_path = directory or self.get_storage_dir()
        if not os.path.isdir(dir_path):
            logger.warning("Design style directory not found", path=dir_path)
            return
        index_path = os.path.join(dir_path, "_index.json")
        if os.path.isfile(index_path):
            with open(index_path, encoding="utf-8") as f:
                meta = json.load(f)
            self._default_name = meta.get("default", "")
            descriptions = meta.get("descriptions", {})
        else:
            self._default_name = ""
            descriptions = {}
        self._styles.clear()
        for fname in sorted(os.listdir(dir_path)):
            if not fname.endswith(".md") or fname == "_index.md":
                continue
            name = fname[:-3]
            md_path = os.path.join(dir_path, fname)
            with open(md_path, encoding="utf-8") as f:
                markdown = f.read()
            self._styles[name] = {
                "name": name,
                "description": descriptions.get(name, ""),
                "markdown": markdown,
            }
        logger.info("Design styles loaded", count=len(self._styles), default=self._default_name)

    def save(self, directory: str | None = None) -> None:
        dir_path = directory or self.get_storage_dir()
        os.makedirs(dir_path, exist_ok=True)
        descriptions = {name: s.get("description", "") for name, s in self._styles.items()}
        with open(os.path.join(dir_path, "_index.json"), "w", encoding="utf-8") as f:
            json.dump({"default": self._default_name, "descriptions": descriptions}, f, ensure_ascii=False, indent=2)
        for name, style in self._styles.items():
            md_path = os.path.join(dir_path, f"{name}.md")
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(style.get("markdown", ""))
        existing = {f for f in os.listdir(dir_path) if f.endswith(".md") and f != "_index.md"}
        expected = {f"{name}.md" for name in self._styles}
        for stale in existing - expected:
            os.remove(os.path.join(dir_path, stale))
        logger.info("Design styles saved", count=len(self._styles))

    def list_styles(self) -> list[dict]:
        return [
            {"name": s["name"], "description": s.get("description", ""),
             "char_count": len(s.get("markdown", ""))}
            for s in self._styles.values()
        ]

    def get_style(self, name: str) -> dict:
        if name not in self._styles:
            raise DesignStyleNotFoundError(name)
        return dict(self._styles[name])

    def get_default_name(self) -> str:
        return self._default_name

    def create_style(self, name: str, description: str, markdown: str) -> None:
        self._styles[name] = {
            "name": name,
            "description": description,
            "markdown": markdown,
        }
        if not self._default_name:
            self._default_name = name
        logger.info("Design style created", name=name)

    def update_style(self, name: str, description: str | None = None, markdown: str | None = None) -> None:
        if name not in self._styles:
            raise DesignStyleNotFoundError(name)
        if description is not None:
            self._styles[name]["description"] = description
        if markdown is not None:
            self._styles[name]["markdown"] = markdown
        logger.info("Design style updated", name=name)

    def delete_style(self, name: str) -> None:
        if name not in self._styles:
            raise DesignStyleNotFoundError(name)
        del self._styles[name]
        if self._default_name == name:
            self._default_name = next(iter(self._styles)) if self._styles else ""
        logger.info("Design style deleted", name=name)

    def set_default(self, name: str) -> None:
        if name not in self._styles:
            raise DesignStyleNotFoundError(name)
        self._default_name = name
        logger.info("Default design style set", name=name)


_design_style_manager: DesignStyleManager | None = None


def get_design_style_manager() -> DesignStyleManager:
    global _design_style_manager
    if _design_style_manager is None:
        _design_style_manager = DesignStyleManager()
        _design_style_manager.load()
        logger.debug("DesignStyleManager singleton created")
    return _design_style_manager