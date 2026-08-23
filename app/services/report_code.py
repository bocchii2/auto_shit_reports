from __future__ import annotations

MES_ABREV = {
    1: "ENE",
    2: "FEB",
    3: "MAR",
    4: "ABR",
    5: "MAY",
    6: "JUN",
    7: "JUL",
    8: "AGO",
    9: "SEP",
    10: "OCT",
    11: "NOV",
    12: "DIC",
}

MES_NOMBRE = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}

MES_NOMBRE_TITLE = {k: v.capitalize() for k, v in MES_NOMBRE.items()}

CODIGO_PREFIX = "ALT-INF-ACT-PERS"


def mes_abrev(mes: int) -> str:
    return MES_ABREV[mes]


def build_codigo(mes: int, anio: int) -> str:
    return f"{CODIGO_PREFIX}-{MES_ABREV[mes]}-{anio}"


def build_periodo_label(mes: int, anio: int) -> str:
    """Ej: Del 01 al 31 de julio de 2026"""
    import calendar

    last = calendar.monthrange(anio, mes)[1]
    return f"Del 01 al {last:02d} de {MES_NOMBRE[mes]} de {anio}"


def build_periodo_corto(mes: int, anio: int) -> str:
    """Ej: Julio 2026"""
    return f"{MES_NOMBRE_TITLE[mes]} {anio}"


def format_fecha_informe(day: int, mes: int, anio: int) -> str:
    """Ej: 7  de agosto de 2026"""
    return f"{day}  de {MES_NOMBRE[mes]} de {anio}"
