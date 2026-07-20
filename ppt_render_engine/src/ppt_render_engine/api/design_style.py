import urllib.parse
from fastapi import APIRouter, HTTPException
from fastapi.responses import PlainTextResponse
from pydantic import BaseModel
from ppt_render_engine.core.design_style import get_design_style_manager, DesignStyleNotFoundError

router = APIRouter(prefix="/agent/design-styles", tags=["design-styles"])


class StyleCreate(BaseModel):
    name: str
    description: str = ""
    markdown: str = ""


class StyleUpdate(BaseModel):
    description: str | None = None
    markdown: str | None = None


@router.get("/list")
async def list_styles():
    mgr = get_design_style_manager()
    return {
        "default": mgr.get_default_name(),
        "styles": mgr.list_styles(),
    }


@router.get("/{name}/download")
async def download_style(name: str):
    mgr = get_design_style_manager()
    try:
        style = mgr.get_style(name)
    except DesignStyleNotFoundError:
        raise HTTPException(status_code=404, detail=f"Design style '{name}' not found")
    return PlainTextResponse(
        style.get("markdown", ""),
        media_type="text/markdown",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{urllib.parse.quote(f'{name}.md', safe='.')}"},
    )


@router.get("/{name}")
async def get_style(name: str):
    mgr = get_design_style_manager()
    try:
        return mgr.get_style(name)
    except DesignStyleNotFoundError:
        raise HTTPException(status_code=404, detail=f"Design style '{name}' not found")


@router.post("/{name}")
async def create_style(name: str, body: StyleCreate):
    if name != body.name:
        raise HTTPException(status_code=400, detail="Name mismatch between path and body")
    mgr = get_design_style_manager()
    mgr.create_style(body.name, body.description, body.markdown)
    mgr.save()
    return {"name": body.name, "description": body.description}


@router.put("/{name}")
async def update_style(name: str, body: StyleUpdate):
    mgr = get_design_style_manager()
    try:
        mgr.update_style(name, description=body.description, markdown=body.markdown)
        mgr.save()
    except DesignStyleNotFoundError:
        raise HTTPException(status_code=404, detail=f"Design style '{name}' not found")
    return {"name": name}


@router.delete("/{name}")
async def delete_style(name: str):
    mgr = get_design_style_manager()
    try:
        mgr.delete_style(name)
        mgr.save()
    except DesignStyleNotFoundError:
        raise HTTPException(status_code=404, detail=f"Design style '{name}' not found")
    return {"deleted": name}


@router.post("/default/{name}")
async def set_default(name: str):
    mgr = get_design_style_manager()
    try:
        mgr.set_default(name)
        mgr.save()
    except DesignStyleNotFoundError:
        raise HTTPException(status_code=404, detail=f"Design style '{name}' not found")
    return {"default": name}
