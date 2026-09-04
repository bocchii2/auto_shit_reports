"""Extrae actividades desde AThi4.historial.txt (repo CodeCommit AThi-4) hacia un CSV
compatible con app/services/csv_import.py.

Fuente primaria: AThi4.historial.txt (bloques por fecha, curados a mano por el usuario).
git log se usa solo para un reporte de inconsistencias (cross-check), nunca genera filas
de CSV: reintroduciría commits duplicados/amendeados que el historial.txt ya resuelve.

Las fechas del historial vienen en formato MES/DIA/AÑO (US). Se convierten aquí mismo a
ISO (AAAA-MM-DD) antes de escribir el CSV: el parser de csv_import.py usa
dateutil.parser.parse(dayfirst=True, fuzzy=True), que es ambiguo para fechas tipo
"07/05/2026" (¿5-jul o 7-may?). Emitir ISO evita esa ambigüedad por completo.

Uso:
    python scripts/extract_athi_historial.py
    python scripts/extract_athi_historial.py --source /ruta/AThi4.historial.txt --output out.csv
"""
from __future__ import annotations

import argparse
import csv
import re
import subprocess
import sys
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SOURCE = Path("/mnt/c/jostin/app/athi/AThi4.historial.txt")
DEFAULT_REPO = Path("/mnt/c/jostin/app/athi")
DEFAULT_OUTPUT = ROOT / "data" / "imports" / "actividades_codecommit_athi.csv"
DEFAULT_REPORT = ROOT / "data" / "imports" / "actividades_codecommit_athi_crosscheck.txt"

CSV_HEADER = [
    "No.",
    "Actividad",
    "Cliente / Proyecto / Servicio / Aplicación",
    "Fecha desde",
    "Fecha Hasta",
    "Descripción",
    "Estatus *",
]

HEADER_RE = re.compile(
    r"^(?P<date>\d{2}/\d{2}/\d{4})\s+(?P<author>.*?)(?:\s+(?P<time>\d{1,2}:\d{2}:\d{2}))?$"
)
ITEM_RE = re.compile(r"^\s*(?P<num>\d+)\.\s*(?P<text>.+?)\s*$")
MSG_DATE_RE = re.compile(r"^(\d{2}/\d{2}/\d{4})")
KNOWN_AUTHOR_EMAILS = {"jostin.cedeno@altura.com.ec", "jostin9876@gmail.com"}


@dataclass
class HistorialBlock:
    fecha: date
    author_raw: str
    items: list[str]


@dataclass
class ActivityRow:
    numero: int
    actividad: str
    cliente: str
    fecha_desde: str
    fecha_hasta: str
    descripcion: str
    estatus: str


def normalize_author(raw: str) -> str:
    text = unicodedata.normalize("NFKD", raw).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.lower().split())


def parse_historial(path: Path) -> list[HistorialBlock]:
    raw_text = path.read_text(encoding="utf-8").replace("\r\n", "\n")
    blocks: list[HistorialBlock] = []
    for chunk in re.split(r"\n\s*\n", raw_text.strip("\n")):
        lines = [ln for ln in chunk.split("\n")]
        if not lines or not lines[0].strip():
            continue
        header_match = HEADER_RE.match(lines[0].strip())
        if not header_match:
            print(f"Aviso: bloque ignorado, encabezado no reconocido: {lines[0]!r}")
            continue
        try:
            fecha = datetime.strptime(header_match["date"], "%m/%d/%Y").date()
        except ValueError:
            print(f"Aviso: fecha inválida en encabezado: {lines[0]!r}")
            continue

        items: list[str] = []
        for line in lines[1:]:
            if not line.strip():
                continue
            item_match = ITEM_RE.match(line)
            if item_match:
                items.append(item_match["text"])
            elif items:
                items[-1] = f"{items[-1]} {line.strip()}"
            else:
                print(f"Aviso: línea sin ítem previo, ignorada: {line!r}")

        if items:
            blocks.append(HistorialBlock(fecha=fecha, author_raw=header_match["author"], items=items))

    return blocks


def dedup_blocks(blocks: list[HistorialBlock]) -> tuple[list[HistorialBlock], list[HistorialBlock]]:
    seen: set[tuple] = set()
    kept: list[HistorialBlock] = []
    dropped: list[HistorialBlock] = []
    for block in blocks:
        key = (
            block.fecha,
            normalize_author(block.author_raw),
            tuple(item.strip().lower() for item in block.items),
        )
        if key in seen:
            dropped.append(block)
            print(f"Duplicado descartado: {block.fecha} ({block.author_raw}) - {len(block.items)} items")
            continue
        seen.add(key)
        kept.append(block)
    return kept, dropped


def split_actividad(text: str, max_len: int = 90) -> tuple[str, str]:
    descripcion = text
    if len(text) <= max_len:
        return text, descripcion

    window = text[:max_len]
    comma_idx = window.rfind(",")
    if comma_idx > 0:
        candidate = text[:comma_idx].strip().rstrip(" .,;:")
        if candidate:
            return candidate, descripcion

    space_idx = window.rfind(" ")
    if space_idx > 0:
        candidate = text[:space_idx].strip().rstrip(" .,;:")
        if candidate:
            return candidate, descripcion

    return text[:max_len].strip(), descripcion


def build_rows(blocks: list[HistorialBlock], cliente: str, estatus: str) -> list[ActivityRow]:
    pairs = [(block.fecha, item) for block in blocks for item in block.items]
    pairs.sort(key=lambda p: p[0])

    rows: list[ActivityRow] = []
    for numero, (fecha, item_text) in enumerate(pairs, start=1):
        actividad, descripcion = split_actividad(item_text)
        rows.append(
            ActivityRow(
                numero=numero,
                actividad=actividad,
                cliente=cliente,
                fecha_desde=fecha.isoformat(),
                fecha_hasta=fecha.isoformat(),
                descripcion=descripcion,
                estatus=estatus,
            )
        )
    return rows


def write_csv(rows: list[ActivityRow], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with open(output, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADER)
        for row in rows:
            writer.writerow(
                [
                    row.numero,
                    row.actividad,
                    row.cliente,
                    row.fecha_desde,
                    row.fecha_hasta,
                    row.descripcion,
                    row.estatus,
                ]
            )


def run_git_log(repo: Path) -> list[dict]:
    try:
        result = subprocess.run(
            ["git", "-C", str(repo), "log", "--all", "--pretty=format:%h|%an|%ae|%ad|%s", "--date=iso-strict"],
            capture_output=True,
            text=True,
            check=True,
        )
    except (subprocess.CalledProcessError, FileNotFoundError, OSError) as exc:
        print(f"Aviso: no se pudo leer git log ({exc}), se omite cross-check")
        return []

    commits: list[dict] = []
    for line in result.stdout.splitlines():
        if not line.strip():
            continue
        parts = line.split("|", 4)
        if len(parts) != 5:
            continue
        commit_hash, author_name, author_email, push_date_iso, subject = parts

        msg_match = MSG_DATE_RE.match(subject.strip())
        if msg_match:
            try:
                worked_date = datetime.strptime(msg_match.group(1), "%m/%d/%Y").date()
            except ValueError:
                worked_date = datetime.strptime(push_date_iso[:10], "%Y-%m-%d").date()
        else:
            worked_date = datetime.strptime(push_date_iso[:10], "%Y-%m-%d").date()

        is_known_author = (
            "jostin" in author_name.lower()
            or "jotin" in author_name.lower()
            or author_email.lower() in KNOWN_AUTHOR_EMAILS
        )
        commits.append(
            {
                "hash": commit_hash,
                "author_name": author_name,
                "author_email": author_email,
                "worked_date": worked_date,
                "subject": subject,
                "is_known_author": is_known_author,
            }
        )
    return commits


def cross_check(commits: list[dict], kept_blocks: list[HistorialBlock], dropped: list[HistorialBlock], tolerance_days: int) -> str:
    lines: list[str] = []
    lines.append("=== Reporte de cross-check: git log vs AThi4.historial.txt ===")
    lines.append(f"Commits analizados: {len(commits)}")
    lines.append(f"Bloques únicos en historial: {len(kept_blocks)} (duplicados descartados: {len(dropped)})")
    lines.append(f"Tolerancia: {tolerance_days} días")
    lines.append("")

    identities = sorted({(c["author_name"], c["author_email"]) for c in commits if c["is_known_author"]})
    if len(identities) > 1:
        joined = " y ".join(f"'{name} <{email}>'" for name, email in identities)
        lines.append(f"Identidad unificada: {joined} se tratan como la misma persona.")
    lines.append("")

    lines.append("-- (a) Commits sin entrada cercana en el historial --")
    any_a = False
    for c in commits:
        if not c["is_known_author"]:
            continue
        if not kept_blocks:
            diff = None
            nearest = None
        else:
            nearest_block = min(kept_blocks, key=lambda b: abs((c["worked_date"] - b.fecha).days))
            diff = abs((c["worked_date"] - nearest_block.fecha).days)
            nearest = nearest_block.fecha
        if diff is None or diff > tolerance_days:
            any_a = True
            lines.append(
                f"[COMMIT SIN HISTORIAL] {c['hash']} {c['worked_date']} "
                f"(\"{c['subject'][:60]}\") — entrada más cercana: {nearest or 'ninguna'} "
                f"({diff if diff is not None else 'n/a'} días)"
            )
    if not any_a:
        lines.append("(ninguno)")
    lines.append("")

    lines.append("-- (b) Entradas del historial sin commit cercano (prioridad baja) --")
    any_b = False
    known_commits = [c for c in commits if c["is_known_author"]]
    for block in kept_blocks:
        if not known_commits:
            diff = None
        else:
            diff = min(abs((block.fecha - c["worked_date"]).days) for c in known_commits)
        if diff is None or diff > tolerance_days:
            any_b = True
            lines.append(f"[HISTORIAL SIN COMMIT] {block.fecha} ({block.author_raw}) - {len(block.items)} items")
    if not any_b:
        lines.append("(ninguno)")

    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--repo", type=Path, default=DEFAULT_REPO)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--cliente", default="Altura")
    parser.add_argument("--estatus", default="Entregado")
    parser.add_argument("--tolerance-days", type=int, default=5)
    parser.add_argument("--no-crosscheck", action="store_true")
    args = parser.parse_args()

    if not args.source.exists():
        print(f"Error: no existe el archivo fuente: {args.source}", file=sys.stderr)
        sys.exit(1)

    blocks = parse_historial(args.source)
    print(f"Bloques leídos: {len(blocks)}")

    kept, dropped = dedup_blocks(blocks)
    print(f"Bloques únicos: {len(kept)} (descartados: {len(dropped)})")

    rows = build_rows(kept, args.cliente, args.estatus)
    write_csv(rows, args.output)
    print(f"Escritas {len(rows)} filas en {args.output}")

    if not args.no_crosscheck:
        commits = run_git_log(args.repo)
        report_text = cross_check(commits, kept, dropped, args.tolerance_days)
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(report_text, encoding="utf-8")
        print()
        print(report_text)
        print()
        print(f"Reporte de cross-check guardado en {args.report}")


if __name__ == "__main__":
    main()
