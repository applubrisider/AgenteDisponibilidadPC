# Agente de Disponibilidad (PC)

Analiza un CSV vertical (Colaborador | Departamento | RUT | Ciudad Residencia | fecha | actividad),
calcula días disponibles en los últimos 30 días, streaks consecutivos y clasifica la criticidad.

## Requisitos
- Python 3.10+ (recomendado)
- Windows (probado), también funciona en Linux/macOS
- (Opcional) Clave API para proveedor IA si quieres usar el módulo de preguntas (OpenAI u Ollama)

## Pasos rápidos (CLI)
```bash
# 1) Crear entorno
python -m venv .venv
# Windows PowerShell:
.\.venv\Scripts\Activate.ps1
# CMD:
# .\.venv\Scripts\activate.bat

# 2) Instalar dependencias
pip install -r requirements.txt

# 3) Copiar y editar variables de entorno
copy .env.example .env   # (en Windows)  |  cp .env.example .env (Linux/macOS)
# Edita .env si usarás IA para preguntas (opcional)

# 4) Ejecutar análisis desde CSV en Google Drive (link de vista):
python src/main.py --csv-url "https://drive.google.com/file/d/1EMhbXo9ptlMqYvZMY2kxBlmPmjIB7uCU/view?usp=drive_link" --out out

# 5) (Opcional) Ejecutar interfaz gráfica
python src/main.py --ui
```

Al terminar, encontrarás:
- `out/resumen_disponibilidad.csv` (por colaborador)
- `out/detalle_disponibilidad.csv` (timeline día a día)
- `out/informe_disponibilidad.html` (reporte en HTML)

## Clasificación por criticidad (configurable en `config.yaml`)
- ALTA si:
  - `dias_disponibles_30d >= 7`, **o**
  - `max_consecutivos >= 3`
- BAJA si:
  - `dias_disponibles_30d <= 4`
- MEDIA en otro caso.

## Actividades consideradas **disponibles**
Por defecto: `DISPONIBLE`, `LIBRE`, `SIN ASIGNACION`, `SIN_ASIGNACION`.
Puedes ajustar o añadir equivalentes en `config.yaml` (detección case-insensitive, por inclusión).

## Empaquetar en .exe (Windows)
```bash
# Desde la raíz del proyecto, con el venv activo
pip install pyinstaller
pyinstaller --noconfirm --onefile --name AgenteDisponibilidadPC --add-data "src/reporting/template.html;reporting" src/main.py

# El ejecutable quedará en dist/AgenteDisponibilidadPC.exe
```

## Inno Setup (opcional)
En `build/installer.iss` tienes una base. Ábrela con Inno Setup y compílala para producir un instalador .exe.

## Módulo IA (opcional)
Puedes hacer preguntas de lenguaje natural al dataset una vez generado `out/resumen_disponibilidad.csv`.
- OpenAI: define `OPENAI_API_KEY` en `.env`.
- Ollama local: define `USE_OLLAMA=true` y `OLLAMA_MODEL=llama3.1` (u otro), con el servicio corriendo.
Usa `--qa` para abrir una consola de preguntas.
