"""
Actualiza docs/meta.json a partir de lo que haya en la carpeta docs/.

Uso: copia tus archivos nuevos dentro de docs/ y ejecuta:

    python generar_meta.py

- Los archivos nuevos se registran con la fecha y hora de hoy.
- Los que ya estaban conservan su fecha original de subida.
- Los que borraste de docs/ se eliminan del registro.
"""
import json
from datetime import datetime
from pathlib import Path

DOCS_DIR = Path(__file__).parent / "docs"
META_FILE = DOCS_DIR / "meta.json"
DOCS_DIR.mkdir(exist_ok=True)

# Registro existente (si lo hay), para no perder las fechas originales
meta_previo = {}
if META_FILE.exists():
    try:
        meta_previo = json.loads(META_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        print("[!] meta.json estaba corrupto; se reconstruye desde cero.")

archivos = sorted(
    f for f in DOCS_DIR.iterdir()
    if f.is_file() and f.name != "meta.json"
)

meta = {}
for ruta in archivos:
    anterior = meta_previo.get(ruta.name)
    if anterior and "fecha" in anterior:
        fecha = anterior["fecha"]          # conserva la fecha original
        estado = "="
    else:
        fecha = datetime.now().isoformat()  # archivo nuevo
        estado = "+"
    meta[ruta.name] = {"fecha": fecha, "size": ruta.stat().st_size}
    print(f"  {estado} {ruta.name}")

eliminados = set(meta_previo) - set(meta)
for nombre in sorted(eliminados):
    print(f"  - {nombre} (ya no esta en docs/)")

META_FILE.write_text(
    json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
)
print(f"\nGuardado: docs/meta.json con {len(meta)} archivo(s).")
print("Leyenda:  + nuevo   = conserva su fecha   - eliminado")
