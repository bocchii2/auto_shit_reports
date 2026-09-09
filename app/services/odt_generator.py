from __future__ import annotations

import shutil
import subprocess
from pathlib import Path


def _find_soffice() -> str | None:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return candidate
    return None


def convert_docx_to_odt(docx_path: str | Path, output_path: str | Path) -> Path:
    soffice = _find_soffice()
    if not soffice:
        raise FileNotFoundError(
            "No se encontró LibreOffice. Instálalo para generar informes ODT "
            "o selecciona el formato DOCX."
        )

    docx_path = Path(docx_path)
    if not docx_path.exists():
        raise FileNotFoundError(f"No existe el informe DOCX de entrada: {docx_path}")
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "odt",
            "--outdir",
            str(output_path.parent),
            str(docx_path),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    generated = output_path.parent / f"{docx_path.stem}.odt"
    if result.returncode != 0 or not generated.exists():
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(f"No se pudo convertir el informe a ODT{': ' + detail if detail else ''}")
    if generated != output_path:
        generated.replace(output_path)
    return output_path
