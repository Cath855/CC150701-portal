"""
Ejecuta este script UNA VEZ en tu computador antes de subir a GitHub.
Genera el docs/meta.json con los 3 archivos iniciales.
"""
import json, os
from datetime import datetime
from pathlib import Path

DOCS_DIR = Path("docs")
DOCS_DIR.mkdir(exist_ok=True)

archivos_iniciales = [
    "Datos_Julio.xlsx",
    "EE109-4.xlsx",
    "EE223-5.xlsx",
]

meta = {}
for nombre in archivos_iniciales:
    ruta = DOCS_DIR / nombre
    if ruta.exists():
        meta[nombre] = {
            "fecha": datetime.now().isoformat(),
            "size": ruta.stat().st_size
        }
        print(f"  + {nombre}")
    else:
        print(f"  [!] No encontrado: docs/{nombre}")

(DOCS_DIR / "meta.json").write_text(
    json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
)
print(f"\nGuardado: docs/meta.json con {len(meta)} archivos")
