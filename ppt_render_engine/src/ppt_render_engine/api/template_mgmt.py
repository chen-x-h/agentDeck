from fastapi import APIRouter, UploadFile, File, HTTPException
from ppt_render_engine.core.template import get_template_manager, TemplateNotFoundError
from ppt_render_engine.log_config import app_logger

router = APIRouter(prefix="/template", tags=["Template"])
logger = app_logger("api.template")


@router.post("/upload/{name}")
async def upload_template(name: str, file: UploadFile = File(...)):
    if not file.filename or not file.filename.endswith((".pptx", ".PPTX")):
        raise HTTPException(400, "Only .pptx files are supported")
    logger.info("Uploading template", name=name, file=file.filename)
    try:
        raw = await file.read()
        tm = get_template_manager()
        path = tm.save_to_storage(name, raw)
        tm.load(name, path)
        prs = tm.get_presentation(name)
        layouts = tm.list_layouts(name)
        logger.info("Template saved and loaded", name=name, path=path, layouts=len(layouts))
        return {
            "template": name,
            "slide_width": prs.slide_width,
            "slide_height": prs.slide_height,
            "layouts": layouts,
        }
    except Exception as e:
        logger.error("Template upload failed", name=name, error=str(e))
        raise HTTPException(500, f"Template load failed: {e}")


@router.get("/list")
async def list_templates():
    tm = get_template_manager()
    names = tm.list_templates()
    result = []
    for name in names:
        prs = tm.get_presentation(name)
        layouts = tm.list_layouts(name)
        result.append({
            "name": name,
            "slide_width": prs.slide_width,
            "slide_height": prs.slide_height,
            "layouts": layouts,
        })
    logger.info("Listed templates", count=len(result))
    return {"templates": result}


@router.get("/{name}/layouts")
async def list_layouts(name: str):
    tm = get_template_manager()
    try:
        prs = tm.get_presentation(name)
        layouts = tm.list_layouts(name)
        logger.info("Listed layouts", template=name, count=len(layouts))
        return {
            "template": name,
            "slide_width": prs.slide_width,
            "slide_height": prs.slide_height,
            "layouts": layouts,
        }
    except TemplateNotFoundError:
        raise HTTPException(404, f"Template '{name}' not found")


@router.delete("/{name}")
async def delete_template(name: str):
    tm = get_template_manager()
    if not tm.has_template(name):
        raise HTTPException(404, f"Template '{name}' not found")
    tm.unload(name)
    tm.delete_from_storage(name)
    logger.info("Template deleted from memory and disk", name=name)
    return {"status": "deleted", "template": name}
