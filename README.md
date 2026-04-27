# Income Certification App

Aplicacion web para generar automaticamente documentacion de ingresos para clientes en Argentina.

Genera:
- Un solo archivo Word (.docx) con dos secciones en este orden:
	1. Certificacion Contable
	2. Manifestacion de Ingresos
- Un solo archivo PDF (.pdf) con el mismo orden

La seccion "Tarea realizada" se adapta segun el tipo de contribuyente:
- Monotributista
- Responsable Inscripto

## Requisitos

- Python 3.11+
- pip
- Docker + Docker Compose (opcional, recomendado)

## Estructura del proyecto

- app/main.py: rutas FastAPI y procesamiento del formulario
- app/generator.py: logica de generacion Word/PDF
- app/templates/index.html: formulario
- app/templates/resultado.html: pantalla de descargas
- output/: documentos generados en tiempo de ejecucion

## Ejecucion local (sin Docker)

1. Crear entorno virtual

```powershell
python -m venv .venv
```

2. Activar entorno virtual (PowerShell)

```powershell
.venv\Scripts\Activate.ps1
```

3. Instalar dependencias

```powershell
pip install -r requirements.txt
```

4. Levantar la aplicacion

```powershell
uvicorn app.main:app --reload --port 8000
```

5. Abrir en navegador

http://localhost:8000

## Ejecucion con Docker

1. Construir y levantar

```powershell
docker-compose up --build -d
```

2. Ver estado

```powershell
docker-compose ps
```

3. Ver logs

```powershell
docker-compose logs --tail 50
```

4. Abrir en navegador

http://localhost:8000

5. Detener contenedor

```powershell
docker-compose down
```

## Como funciona el flujo

1. El usuario completa el formulario en index.html
2. FastAPI recibe los datos en POST /generar
3. Se calculan total/promedio y etiquetas de periodo
4. Se generan archivos en output/
5. resultado.html ofrece descargas directas

## Notas importantes

- output/ esta excluido del versionado por .gitignore.
- Se mantiene output/.gitkeep para conservar la carpeta en el repositorio.
- El servidor HTTP de FastAPI dentro del contenedor lo provee Uvicorn.

## Troubleshooting rapido

- Si no abre en navegador:
	- verificar docker-compose ps
	- revisar logs con docker-compose logs --tail 100
- Si aparece error de encoding en pruebas locales:
	- leer archivos Python con encoding='utf-8' en scripts de chequeo
