from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.templating import Jinja2Templates
import os

from app.generator import generate_documents

app = FastAPI(
    title="Income Certification App",
    description=(
        "API para generación automática de documentos de **Certificación Contable** "
        "y **Manifestación de Ingresos** para contribuyentes argentinos "
        "(Monotributistas y Responsables Inscriptos).\n\n"
        "Los documentos se generan en formato **Word (.docx)** y **PDF** "
        "con descarga directa desde el navegador."
    ),
    version="1.0.0",
    contact={
        "name": "Income Certification App",
        "url": "https://github.com/DanielDevProyects/income-certification-app",
    },
    license_info={
        "name": "MIT",
    },
)

templates = Jinja2Templates(directory="app/templates")

MONTHS_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

YEARS_WORDS = {
    2024: "dos mil veinticuatro",
    2025: "dos mil veinticinco",
    2026: "dos mil veintiséis",
    2027: "dos mil veintisiete",
    2028: "dos mil veintiocho",
    2029: "dos mil veintinueve",
    2030: "dos mil treinta",
}

OUTPUT_DIR = os.environ.get("OUTPUT_DIR", "output")


@app.get(
    "/",
    response_class=HTMLResponse,
    summary="Formulario principal",
    description="Devuelve el formulario HTML para ingresar los datos del contribuyente y generar los documentos.",
    tags=["UI"],
    include_in_schema=False,
)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post(
    "/generar",
    summary="Generar documentos",
    description=(
        "Recibe los datos del contribuyente mediante un formulario HTML y genera:\n\n"
        "- Un archivo **Word (.docx)** con Certificación Contable (p.1) y Manifestación de Ingresos (p.2).\n"
        "- Un archivo **PDF** con el mismo contenido.\n\n"
        "Retorna la página de descarga con los enlaces a ambos archivos."
    ),
    tags=["Documentos"],
    include_in_schema=False,
)
async def generar(
    request: Request,
    # Datos del cliente
    nombre_completo: str = Form(...),
    tipo_contribuyente: str = Form("monotributista"),
    dni: str = Form(...),
    cuit: str = Form(...),
    domicilio: str = Form(...),
    ciudad: str = Form(...),
    provincia: str = Form(...),
    actividad: str = Form(...),
    # Período
    mes_inicio: int = Form(...),
    anio_inicio: int = Form(...),
    mes_fin: int = Form(...),
    anio_fin: int = Form(...),
    # Fecha de emisión
    dia_emision: int = Form(...),
    mes_emision: int = Form(...),
    anio_emision: int = Form(...),
    # Ingresos mensuales (12 meses) - se reciben como strings
    m1: str = Form("0"),
    m2: str = Form("0"),
    m3: str = Form("0"),
    m4: str = Form("0"),
    m5: str = Form("0"),
    m6: str = Form("0"),
    m7: str = Form("0"),
    m8: str = Form("0"),
    m9: str = Form("0"),
    m10: str = Form("0"),
    m11: str = Form("0"),
    m12: str = Form("0"),
):
    def parse_amount(val: str) -> float:
        val = val.strip().replace(".", "").replace(",", ".")
        try:
            return float(val)
        except:
            return 0.0

    montos = [parse_amount(x) for x in [m1, m2, m3, m4, m5, m6, m7, m8, m9, m10, m11, m12]]
    total = sum(montos)
    promedio = total / 12

    # Generar etiquetas de meses dinámicamente
    month_labels = []
    m, y = mes_inicio, anio_inicio
    for _ in range(12):
        month_labels.append(f"{MONTHS_ES[m-1].capitalize()} {y}")
        m += 1
        if m > 12:
            m = 1
            y += 1

    data = {
        "nombre_completo": nombre_completo,
        "tipo_contribuyente": tipo_contribuyente,
        "dni": dni,
        "cuit": cuit,
        "domicilio": domicilio,
        "ciudad": ciudad,
        "provincia": provincia,
        "actividad": actividad,
        "mes_inicio_label": MONTHS_ES[mes_inicio - 1],
        "anio_inicio": anio_inicio,
        "mes_fin_label": MONTHS_ES[mes_fin - 1],
        "anio_fin": anio_fin,
        "dia_emision": dia_emision,
        "mes_emision": mes_emision,
        "mes_emision_label": MONTHS_ES[mes_emision - 1],
        "anio_emision": anio_emision,
        "anio_emision_letras": YEARS_WORDS.get(anio_emision, str(anio_emision)),
        "montos": montos,
        "month_labels": month_labels,
        "total": total,
        "promedio": promedio,
    }

    word_path, pdf_path = generate_documents(data)

    return templates.TemplateResponse("resultado.html", {
        "request": request,
        "nombre": nombre_completo,
        "word_file": os.path.basename(word_path),
        "pdf_file": os.path.basename(pdf_path),
    })


@app.get(
    "/download/{filename}",
    summary="Descargar documento generado",
    description=(
        "Descarga un documento previamente generado (`.docx` o `.pdf`) por su nombre de archivo.\n\n"
        "Los archivos se almacenan en el directorio `output/` del servidor."
    ),
    tags=["Documentos"],
    responses={
        200: {"description": "Archivo descargado correctamente"},
        404: {"description": "Archivo no encontrado"},
    },
)
async def download(filename: str):
    filepath = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(filepath):
        return HTMLResponse(content="Archivo no encontrado", status_code=404)
    return FileResponse(filepath, filename=filename)
