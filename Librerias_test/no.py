import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
RUTA_TXT = os.environ.get(
    "EASYGESTION_ULTIMO_EXCEL_TXT",
    str(BASE_DIR / "descargas" / "ultimo_excel.txt"),
)

print(os.path.exists(RUTA_TXT))