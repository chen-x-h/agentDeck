from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from ppt_render_engine.core.color_scheme import get_color_scheme_manager, ColorSchemeNotFoundError

router = APIRouter(prefix="/color-scheme", tags=["color-scheme"])


class ColorSchemeCreate(BaseModel):
    name: str
    colors: dict[str, str]


class ColorSchemeDefault(BaseModel):
    name: str


@router.get("/list")
async def list_schemes():
    mgr = get_color_scheme_manager()
    return {
        "default": mgr.get_default_name(),
        "schemes": mgr.list_schemes(),
    }


@router.get("/{name}")
async def get_scheme(name: str):
    mgr = get_color_scheme_manager()
    try:
        colors = mgr.get_colors(name)
        return {"name": name, "colors": colors}
    except ColorSchemeNotFoundError:
        raise HTTPException(status_code=404, detail=f"Color scheme '{name}' not found")


@router.post("/{name}")
async def add_scheme(name: str, body: ColorSchemeCreate):
    if name != body.name:
        raise HTTPException(status_code=400, detail="Name mismatch between path and body")
    mgr = get_color_scheme_manager()
    mgr.add_scheme(body.name, body.colors)
    mgr.save()
    return {"name": body.name, "colors": body.colors}


@router.put("/{name}")
async def update_scheme(name: str, body: ColorSchemeCreate):
    if name != body.name:
        raise HTTPException(status_code=400, detail="Name mismatch between path and body")
    mgr = get_color_scheme_manager()
    try:
        mgr.get_colors(name)
    except ColorSchemeNotFoundError:
        raise HTTPException(status_code=404, detail=f"Color scheme '{name}' not found")
    mgr.add_scheme(body.name, body.colors)
    mgr.save()
    return {"name": body.name, "colors": body.colors}


@router.delete("/{name}")
async def delete_scheme(name: str):
    mgr = get_color_scheme_manager()
    try:
        mgr.remove_scheme(name)
        mgr.save()
    except ColorSchemeNotFoundError:
        raise HTTPException(status_code=404, detail=f"Color scheme '{name}' not found")
    return {"deleted": name}


@router.post("/default/{name}")
async def set_default(name: str):
    mgr = get_color_scheme_manager()
    try:
        mgr.set_default(name)
        mgr.save()
    except ColorSchemeNotFoundError:
        raise HTTPException(status_code=404, detail=f"Color scheme '{name}' not found")
    return {"default": name}
