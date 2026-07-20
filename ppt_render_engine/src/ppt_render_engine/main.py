from contextlib import asynccontextmanager
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from ppt_render_engine.api.convert import router as convert_router
from ppt_render_engine.api.preview import router as preview_router
from ppt_render_engine.api.template_mgmt import router as template_router
from ppt_render_engine.api.image import router as image_router
from ppt_render_engine.api.color_scheme import router as color_scheme_router
from ppt_render_engine.api.design_style import router as design_style_router
from ppt_render_engine.api.sync import router as sync_router
from ppt_render_engine.api.agent import router as agent_router
from ppt_render_engine.core.template import get_template_manager
from ppt_render_engine.core.color_scheme import get_color_scheme_manager
from ppt_render_engine.core.design_style import get_design_style_manager
from ppt_render_engine.log_config import app_logger
from ppt_render_engine.temp_util import get_temp_dir

logger = app_logger("ppt_engine")


@asynccontextmanager
async def lifespan(app: FastAPI):
    tm = get_template_manager()
    loaded = tm.preload_all()
    if loaded:
        logger.info("Templates preloaded on startup", count=len(loaded), names=loaded)
    csm = get_color_scheme_manager()
    logger.info("Color schemes loaded", count=len(csm.list_schemes()), default=csm.get_default_name())
    dsm = get_design_style_manager()
    logger.info("Design styles loaded", count=len(dsm.list_styles()), default=dsm.get_default_name())
    get_temp_dir()
    logger.info("Temp directory ready")
    yield


app = FastAPI(
    title="PPT Render Engine",
    description="JSON↔PPTX 高保真双向转换服务，支持模板预加载",
    version="0.2.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(convert_router)
app.include_router(preview_router)
app.include_router(template_router)
app.include_router(image_router)
app.include_router(color_scheme_router)
app.include_router(design_style_router)
app.include_router(sync_router)
app.include_router(agent_router)


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.delete("/temp")
async def cleanup_temp():
    td = get_temp_dir()
    count = 0
    for f in Path(td).iterdir():
        if f.is_file():
            f.unlink()
            count += 1
    logger.info("Temp files cleaned", count=count)
    return {"status": "cleaned", "files_removed": count}


web_dir = Path(__file__).parents[2] / "web"
if web_dir.is_dir():
    @app.get("/convert-pptx")
    async def convert_pptx_page():
        return FileResponse(str(web_dir / "convert-pptx.html"))
    app.mount("/", StaticFiles(directory=str(web_dir), html=True), name="web")

logger.info("PPT Render Engine initialized")
