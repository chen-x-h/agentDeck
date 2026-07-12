import yaml
import uvicorn
from pathlib import Path
from ppt_render_engine.log_config import app_logger


def main():
    logger = app_logger("startup")
    config_path = Path(__file__).parent / "config.yaml"
    with open(config_path, encoding="utf-8") as f:
        config = yaml.safe_load(f)

    server_cfg = config.get("server", {})
    host = server_cfg.get("host", "127.0.0.1")
    port = server_cfg.get("port", 8000)

    logger.info("Starting PPT Render Engine", host=host, port=port, config=str(config_path))
    uvicorn.run(
        "ppt_render_engine.main:app",
        host=host,
        port=port,
        reload=server_cfg.get("reload", False),
    )


if __name__ == "__main__":
    main()
