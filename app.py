from fastapi import FastAPI, Request
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from pathlib import Path
from server.routes.hierarchy_routers import router as hierarchy_router
from server.routes.technical_map_routes import router as technical_map_router
from server.routes.pdf_parser_routes import router as pdf_parser_router
from server.routes.settings_routes import router as settings_router

import logging, sys

logging.basicConfig(
    level=logging.INFO,  # Показывать INFO и выше
    format='%(asctime)s | %(levelname)s | %(name)s | %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)  # Вывод в консоль
    ]
)

app = FastAPI(title="TK AI Generator")

app.include_router(hierarchy_router)
app.include_router(technical_map_router)
app.include_router(pdf_parser_router)
app.include_router(settings_router)

client_dir = Path(__file__).parent / "client"
app.mount("/assets", StaticFiles(directory=str(client_dir / "assets")), name="assets")

templates = Jinja2Templates(directory="client")

@app.get("/")
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/hierarchy/")
async def hierarchy_page(request: Request):
    return templates.TemplateResponse("hierarchy.html", {"request": request})

@app.get("/technical_map/")
async def technical_map_page(request: Request):
    return templates.TemplateResponse("technical_map.html", {"request": request})

@app.get("/pdf_parser/")
async def pdf_parser_page(request: Request):
    return templates.TemplateResponse("pdf_parser.html", {"request": request})
