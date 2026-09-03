from __future__ import annotations

import re
import tempfile
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
from fastapi import UploadFile

CUENTA_EXCLUIDA = "4011010054"
EXTENSIONES_PERMITIDAS = {".xlsx", ".xlsm", ".xls"}


def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def strip_accents(s):
    s = clean_text(s)
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def norm_text(s):
    s = strip_accents(s).lower()
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


def normalize_account(x):
    s = clean_text(x).upper()
    s = re.sub(r"-R$", "", s)
    s = re.sub(r"[^0-9]", "", s)
    return s


def to_number(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)

    s = str(x).strip()
    s = s.replace("S/.", "")
    s = s.replace("S/", "")
    s = s.replace(",", "")
    s = s.replace(" ", "")

    if s.upper() in ["", "-", "NULO", "NULL", "NAN", "NONE"]:
        return 0.0

    try:
        return float(s)
    except Exception:
        return 0.0


def find_col(df, candidates):
    cols = list(df.columns)
    lookup = {norm_text(c): c for c in cols}

    for cand in candidates:
        key = norm_text(cand)
        if key in lookup:
            return lookup[key]

    for cand in candidates:
        key = norm_text(cand)
        for k, col in lookup.items():
            if key in k or k in key:
                return col

    return None


def infer_period_from_filename(filename):
    name = strip_accents(filename).lower()

    month_map = {
        "enero": "ene", "ene": "ene", "jan": "ene",
        "febrero": "feb", "feb": "feb",
        "marzo": "mar", "mar": "mar",
        "abril": "abr", "abr": "abr", "apr": "abr",
        "mayo": "may", "may": "may",
        "junio": "jun", "jun": "jun",
        "julio": "jul", "jul": "jul",
        "agosto": "ago", "ago": "ago", "aug": "ago",
        "septiembre": "sep", "setiembre": "sep", "sep": "sep", "set": "sep",
        "octubre": "oct", "oct": "oct",
        "noviembre": "nov", "nov": "nov",
        "diciembre": "dic", "dic": "dic", "dec": "dic",
    }

    for raw, short in month_map.items():
        m = re.search(rf"\b{raw}\b[\s_\-]*([0-9]{{2,4}})", name)
        if m:
            return f"{short}{m.group(1)[-2:]}"

        m = re.search(rf"\b{raw}([0-9]{{2,4}})\b", name)
        if m:
            return f"{short}{m.group(1)[-2:]}"

    m = re.search(r"20([0-9]{2})", name)
    if m:
        return f"periodo{m.group(1)}"

    return "periodo"


def infer_period_from_names(nombres_originales):
    periodos = []

    for nombre in nombres_originales.values():
        if not nombre:
            continue
        periodo = infer_period_from_filename(nombre)
        if periodo != "periodo":
            periodos.append(periodo)

    unicos = sorted(set(periodos))

    if len(unicos) == 0:
        return "periodo"
    if len(unicos) == 1:
        return unicos[0]

    raise ValueError(
        "Los nombres de los cinco archivos indican más de un periodo: "
        f"{unicos}. En una ejecución deben corresponder al mismo periodo."
    )


def read_excel_auto(path, preferred_sheet=None, markers=None):
    xls = pd.ExcelFile(path)

    if preferred_sheet and preferred_sheet in xls.sheet_names:
        sheet = preferred_sheet
    elif preferred_sheet:
        preferred_norm = norm_text(preferred_sheet)
        candidates = [
            s for s in xls.sheet_names
            if preferred_norm in norm_text(s)
        ]
        sheet = candidates[0] if candidates else xls.sheet_names[0]
    else:
        sheet = xls.sheet_names[0]

    markers = markers or []
    markers_norm = [norm_text(m) for m in markers]

    preview = pd.read_excel(
        path,
        sheet_name=sheet,
        header=None,
        nrows=25,
        dtype=object,
    )

    best_header = 0
    best_score = -1

    for i in range(len(preview)):
        row_values = [clean_text(x) for x in preview.iloc[i].tolist()]
        joined = " ".join(row_values)
        joined_norm = norm_text(joined)

        non_empty = sum(1 for x in row_values if x != "")
        marker_hits = sum(1 for m in markers_norm if m in joined_norm)
        score = marker_hits * 100 + non_empty

        if score > best_score:
            best_score = score
            best_header = i

    df = pd.read_excel(
        path,
        sheet_name=sheet,
        header=best_header,
        dtype=object,
    )

    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    print(
        f"[LECTURA] archivo={Path(path).name} "
        f"hoja={sheet!r} header_fila={best_header + 1}",
        flush=True,
    )

    return df


def aggregate_by_account(df, account_col, amount_col, source_name, sign=1):
    if account_col is None or amount_col is None:
        return pd.DataFrame(columns=["cuenta_contable", source_name])

    temp = df[[account_col, amount_col]].copy()
    temp["cuenta_contable"] = temp[account_col].apply(normalize_account)
    temp[source_name] = temp[amount_col].apply(to_number) * sign

    temp = temp[temp["cuenta_contable"] != ""]
    temp = temp[temp["cuenta_contable"] != CUENTA_EXCLUIDA]

    return (
        temp.groupby("cuenta_contable", as_index=False)[source_name]
        .sum()
    )


def safe_period_label(label):
    return re.sub(r"[^a-zA-Z0-9_]+", "_", label)


def export_excel_ejecutivo(output_excel, periodo, resumen, result, open_cloud):
    resumen_exec = resumen.copy()
    resumen_exec["Valor_num"] = pd.to_numeric(resumen_exec["Valor"], errors="coerce")
    kpi = dict(zip(resumen_exec["Indicador"], resumen_exec["Valor"]))

    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        workbook = writer.book

        c_navy = "#123B5D"
        c_blue = "#2F75B5"
        c_green = "#70AD47"
        c_orange = "#C55A11"
        c_white = "#FFFFFF"
        c_text = "#1F1F1F"

        fmt_title = workbook.add_format({
            "bold": True,
            "font_size": 18,
            "font_color": c_white,
            "bg_color": c_navy,
            "align": "left",
            "valign": "vcenter"
        })
        fmt_subtitle = workbook.add_format({
            "font_size": 10,
            "font_color": "#DCE6F1",
            "bg_color": c_navy,
            "align": "left",
            "valign": "vcenter"
        })
        fmt_section = workbook.add_format({
            "bold": True,
            "font_size": 12,
            "font_color": c_white,
            "bg_color": c_blue,
            "align": "left",
            "valign": "vcenter"
        })
        fmt_header = workbook.add_format({
            "bold": True,
            "font_color": c_white,
            "bg_color": c_navy,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "text_wrap": True
        })
        fmt_default = workbook.add_format({
            "border": 1,
            "font_color": c_text
        })
        fmt_alt = workbook.add_format({
            "border": 1,
            "bg_color": "#FAFBFC",
            "font_color": c_text
        })
        fmt_money = workbook.add_format({
            "num_format": '"S/ " #,##0.00;[Red]-"S/ " #,##0.00',
            "border": 1
        })
        fmt_money_alt = workbook.add_format({
            "num_format": '"S/ " #,##0.00;[Red]-"S/ " #,##0.00',
            "border": 1,
            "bg_color": "#FAFBFC"
        })
        fmt_int = workbook.add_format({
            "num_format": '#,##0',
            "border": 1
        })
        fmt_int_alt = workbook.add_format({
            "num_format": '#,##0',
            "border": 1,
            "bg_color": "#FAFBFC"
        })
        fmt_ok = workbook.add_format({
            "bg_color": "#E2F0D9",
            "font_color": "#375623",
            "bold": True,
            "border": 1,
            "align": "center"
        })
        fmt_bad = workbook.add_format({
            "bg_color": "#FCE4D6",
            "font_color": "#9C0006",
            "bold": True,
            "border": 1,
            "align": "center"
        })
        fmt_card_blue = workbook.add_format({
            "bg_color": c_blue,
            "font_color": c_white,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "font_size": 12,
            "border": 1
        })
        fmt_card_green = workbook.add_format({
            "bg_color": c_green,
            "font_color": c_white,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "font_size": 12,
            "border": 1
        })
        fmt_card_orange = workbook.add_format({
            "bg_color": c_orange,
            "font_color": c_white,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "font_size": 12,
            "border": 1
        })
        fmt_card_navy = workbook.add_format({
            "bg_color": c_navy,
            "font_color": c_white,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "font_size": 12,
            "border": 1
        })
        fmt_helper = workbook.add_format({
            "font_color": "#FFFFFF",
            "bg_color": "#FFFFFF",
            "border": 0
        })
        fmt_helper_money = workbook.add_format({
            "font_color": "#FFFFFF",
            "bg_color": "#FFFFFF",
            "border": 0,
            "num_format": '"S/ " #,##0.00'
        })
        fmt_helper_int = workbook.add_format({
            "font_color": "#FFFFFF",
            "bg_color": "#FFFFFF",
            "border": 0,
            "num_format": '#,##0'
        })

        # HOJA 1: RESUMEN EJECUTIVO
        ws = workbook.add_worksheet("Resumen Ejecutivo")
        writer.sheets["Resumen Ejecutivo"] = ws
        ws.hide_gridlines(2)

        ws.merge_range(
            "A1:H1",
            f"Conciliación de cuentas contables | {periodo}",
            fmt_title
        )
        ws.merge_range(
            "A2:H2",
            "Conciliación entre BDEP Operativa, SAP y Open Cloud (EP + IP - RE), "
            "excluyendo la cuenta 4011010054.",
            fmt_subtitle
        )

        def money_str(x):
            return f"S/ {float(x):,.2f}"

        cuentas_conciliadas = int(float(kpi.get("Cuentas conciliadas", 0)))
        cuentas_ok = int(float(kpi.get("Cuentas OK", 0)))
        cuentas_dif = int(float(kpi.get("Cuentas con diferencia", 0)))

        total_bdep = float(kpi.get("Total BDEP Operativa", 0))
        total_oc = float(kpi.get("Total Open Cloud", 0))
        total_sap = float(kpi.get("Total SAP", 0))

        dif_bdep_oc = float(kpi.get("Diferencia total BDEP vs Open Cloud", 0))
        dif_bdep_sap = float(kpi.get("Diferencia total BDEP vs SAP", 0))

        ws.merge_range("A4:B5", f"Cuentas conciliadas\n{cuentas_conciliadas:,}", fmt_card_blue)
        ws.merge_range("C4:D5", f"Cuentas OK\n{cuentas_ok:,}", fmt_card_green)
        ws.merge_range(
            "E4:F5",
            f"Cuentas con diferencia\n{cuentas_dif:,}",
            fmt_card_orange if cuentas_dif > 0 else fmt_card_green
        )
        ws.merge_range("G4:H5", f"Total conciliado\n{money_str(total_bdep)}", fmt_card_navy)

        ws.merge_range("A6:B7", f"Total BDEP\n{money_str(total_bdep)}", fmt_card_blue)
        ws.merge_range("C6:D7", f"Total Open Cloud\n{money_str(total_oc)}", fmt_card_blue)
        ws.merge_range("E6:F7", f"Total SAP\n{money_str(total_sap)}", fmt_card_blue)
        ws.merge_range(
            "G6:H7",
            "Cuadre total\nOK"
            if abs(dif_bdep_oc) < 0.01 and abs(dif_bdep_sap) < 0.01
            else "Cuadre total\nREVISAR",
            fmt_card_green
            if abs(dif_bdep_oc) < 0.01 and abs(dif_bdep_sap) < 0.01
            else fmt_card_orange
        )

        ws.write("A9", "Resumen numérico", fmt_section)
        ws.write("A10", "Indicador", fmt_header)
        ws.write("B10", "Valor", fmt_header)

        order = [
            "Cuentas conciliadas",
            "Cuentas OK",
            "Cuentas con diferencia",
            "Total BDEP Operativa",
            "Total Open Cloud",
            "Total SAP",
            "Diferencia total BDEP vs Open Cloud",
            "Diferencia total BDEP vs SAP",
            "Diferencia total Open Cloud vs SAP",
        ]

        for i, label in enumerate(order, start=10):
            alt = i % 2 != 0
            ws.write(i, 0, label, fmt_alt if alt else fmt_default)
            val = kpi.get(label, 0)

            if "Total" in label or "Diferencia total" in label:
                ws.write_number(i, 1, float(val), fmt_money_alt if alt else fmt_money)
            else:
                ws.write_number(i, 1, float(val), fmt_int_alt if alt else fmt_int)

        # DATOS AUXILIARES PARA GRÁFICOS
        ws.write("J10", "Fuente", fmt_helper)
        ws.write("K10", "Monto", fmt_helper)

        ws.write("J11", "BDEP Operativa", fmt_helper)
        ws.write("J12", "Open Cloud", fmt_helper)
        ws.write("J13", "SAP", fmt_helper)

        ws.write_number("K11", total_bdep, fmt_helper_money)
        ws.write_number("K12", total_oc, fmt_helper_money)
        ws.write_number("K13", total_sap, fmt_helper_money)

        ws.write("J15", "Estado", fmt_helper)
        ws.write("K15", "Cantidad", fmt_helper)

        if cuentas_dif > 0:
            ws.write("J16", "OK", fmt_helper)
            ws.write("J17", "DIFERENCIA", fmt_helper)
            ws.write_number("K16", cuentas_ok, fmt_helper_int)
            ws.write_number("K17", cuentas_dif, fmt_helper_int)
            estado_first_row = 15
            estado_last_row = 16
        else:
            ws.write("J16", "OK", fmt_helper)
            ws.write_number("K16", max(cuentas_ok, 1), fmt_helper_int)
            estado_first_row = 15
            estado_last_row = 15

        chart1 = workbook.add_chart({"type": "column"})
        chart1.add_series({
            "name": "Monto conciliado",
            "categories": ["Resumen Ejecutivo", 10, 9, 12, 9],
            "values": ["Resumen Ejecutivo", 10, 10, 12, 10],
            "data_labels": {
                "value": True,
                "num_format": '"S/ " #,##0.00'
            }
        })
        chart1.set_title({"name": "Totales por fuente"})
        chart1.set_y_axis({
            "num_format": '"S/ " #,##0',
            "major_gridlines": {"visible": True}
        })
        chart1.set_x_axis({"label_position": "low"})
        chart1.set_legend({"none": True})
        chart1.set_style(10)
        ws.insert_chart("D10", chart1, {"x_scale": 1.25, "y_scale": 1.20})

        chart2 = workbook.add_chart({"type": "doughnut"})
        chart2.add_series({
            "name": "Estado de conciliación",
            "categories": ["Resumen Ejecutivo", estado_first_row, 9, estado_last_row, 9],
            "values": ["Resumen Ejecutivo", estado_first_row, 10, estado_last_row, 10],
            "data_labels": {
                "percentage": True,
                "category": True,
                "position": "best_fit"
            }
        })
        chart2.set_title({"name": "Estado de conciliación"})
        chart2.set_style(10)
        ws.insert_chart("D23", chart2, {"x_scale": 1.15, "y_scale": 1.15})

        ws.set_column("A:A", 44)
        ws.set_column("B:B", 20)
        ws.set_column("C:H", 16)
        ws.set_column("J:K", 2)

        # HOJA 2: DETALLE
        result.to_excel(writer, sheet_name="Detalle", index=False)
        ws = writer.sheets["Detalle"]
        ws.hide_gridlines(2)

        for col_num, col_name in enumerate(result.columns):
            ws.write(0, col_num, col_name, fmt_header)

        money_cols = [
            "base_operativa",
            "carga_OP",
            "descarga_SAP",
            "dif_BDEP_vs_OP",
            "dif_BDEP_vs_SAP",
            "dif_OP_vs_SAP"
        ]

        for r in range(1, len(result) + 1):
            alt = r % 2 == 0

            for c, col_name in enumerate(result.columns):
                val = result.iloc[r - 1, c]

                if col_name == "estado":
                    ws.write(r, c, val, fmt_ok if str(val) == "OK" else fmt_bad)

                elif col_name in money_cols:
                    if pd.isna(val):
                        val = 0
                    ws.write_number(r, c, float(val), fmt_money_alt if alt else fmt_money)

                else:
                    ws.write(r, c, val, fmt_alt if alt else fmt_default)

        for col_num, col_name in enumerate(result.columns):
            if col_name == "cuenta_contable":
                ws.set_column(col_num, col_num, 18)
            elif col_name == "estado":
                ws.set_column(col_num, col_num, 16)
            else:
                ws.set_column(col_num, col_num, 18)

        ws.autofilter(0, 0, len(result), len(result.columns) - 1)
        ws.freeze_panes(1, 0)

        # HOJA 3: OPEN CLOUD DESAGREGADO
        open_cloud.to_excel(writer, sheet_name="OpenCloud_desagregado", index=False)
        ws = writer.sheets["OpenCloud_desagregado"]
        ws.hide_gridlines(2)

        for col_num, col_name in enumerate(open_cloud.columns):
            ws.write(0, col_num, col_name, fmt_header)

        oc_money_cols = [c for c in open_cloud.columns if c != "cuenta_contable"]

        for r in range(1, len(open_cloud) + 1):
            alt = r % 2 == 0

            for c, col_name in enumerate(open_cloud.columns):
                val = open_cloud.iloc[r - 1, c]

                if col_name in oc_money_cols:
                    if pd.isna(val):
                        val = 0
                    ws.write_number(r, c, float(val), fmt_money_alt if alt else fmt_money)
                else:
                    ws.write(r, c, val, fmt_alt if alt else fmt_default)

        for col_num, _ in enumerate(open_cloud.columns):
            ws.set_column(col_num, col_num, 18)

        ws.autofilter(0, 0, len(open_cloud), len(open_cloud.columns) - 1)
        ws.freeze_panes(1, 0)


def process_period(periodo, archivos, output_dir):
    bdep = read_excel_auto(
        archivos["bdep"],
        preferred_sheet="BDEP",
        markers=["Cuenta Contable", "Pérdida bruta del evento"]
    )
    sap = read_excel_auto(
        archivos["sap"],
        markers=[
            "Etiquetas de fila",
            "Cuenta de mayor",
            "Valor de moneda de sociedad",
            "Suma de Valor de moneda de sociedad"
        ]
    )
    ep = read_excel_auto(
        archivos["ep"],
        markers=["ID Interno Evento", "Cuentas Contables", "Pérdida Real"]
    )
    ip = read_excel_auto(
        archivos["ip"],
        markers=["ID Interno Impacto", "Cuentas Contables", "Pérdida Real"]
    )
    recu = read_excel_auto(
        archivos["re"],
        markers=["ID Interno Recupero", "Cuentas Contables", "Importe de Recuperación"]
    )

    print(
        f"[FILAS] periodo={periodo} BDEP={len(bdep)} SAP={len(sap)} "
        f"EP={len(ep)} IP={len(ip)} RE={len(recu)}",
        flush=True,
    )

    bdep_account_col = find_col(bdep, ["Cuenta Contable"])
    bdep_amount_col = find_col(
        bdep,
        ["Pérdida bruta del evento (S/ )", "Monto bruto", "Pérdida bruta del evento"]
    )

    sap_account_col = find_col(sap, ["Etiquetas de fila", "Cuenta de mayor"])
    sap_amount_col = find_col(
        sap,
        [
            "Suma de Valor de moneda de sociedad",
            "Valor de moneda de sociedad",
            "Importe en moneda de sociedad",
            "Importe",
            "Monto",
        ]
    )

    ep_account_col = find_col(ep, ["Cuentas Contables"])
    ip_account_col = find_col(ip, ["Cuentas Contables"])
    re_account_col = find_col(recu, ["Cuentas Contables"])

    ep_amount_col = find_col(
        ep,
        ["Pérdida Real.Importe local", "Importe local", "Monto"]
    )
    ip_amount_col = find_col(
        ip,
        ["Pérdida Real.Importe local", "Importe local", "Monto"]
    )
    re_amount_col = find_col(
        recu,
        ["Importe de Recuperación Real.Importe local", "Importe local", "Monto"]
    )

    print(
        "[COLUMNAS] "
        f"BDEP=({bdep_account_col!r},{bdep_amount_col!r}) "
        f"SAP=({sap_account_col!r},{sap_amount_col!r}) "
        f"EP=({ep_account_col!r},{ep_amount_col!r}) "
        f"IP=({ip_account_col!r},{ip_amount_col!r}) "
        f"RE=({re_account_col!r},{re_amount_col!r})",
        flush=True,
    )

    required = {
        "BDEP cuenta": bdep_account_col,
        "BDEP monto": bdep_amount_col,
        "SAP cuenta": sap_account_col,
        "SAP monto": sap_amount_col,
        "IP cuenta": ip_account_col,
        "IP monto": ip_amount_col,
        "RE cuenta": re_account_col,
        "RE monto": re_amount_col,
    }

    missing = [k for k, v in required.items() if v is None]
    if missing:
        raise ValueError(
            f"No se pudieron identificar estas columnas para {periodo}: {missing}"
        )

    if ep_account_col is None or ep_amount_col is None:
        print(
            f"[NOTA] {periodo}: EP no tiene cuenta/monto utilizable. "
            "Se considerará 0 en EP.",
            flush=True,
        )

    bdep_agg = aggregate_by_account(
        bdep, bdep_account_col, bdep_amount_col, "base_operativa", sign=1
    )
    sap_agg = aggregate_by_account(
        sap, sap_account_col, sap_amount_col, "descarga_SAP", sign=1
    )
    ep_agg = aggregate_by_account(
        ep, ep_account_col, ep_amount_col, "open_cloud_EP", sign=1
    )
    ip_agg = aggregate_by_account(
        ip, ip_account_col, ip_amount_col, "open_cloud_IP", sign=1
    )
    re_agg = aggregate_by_account(
        recu, re_account_col, re_amount_col, "open_cloud_RE", sign=-1
    )

    open_cloud = pd.DataFrame({"cuenta_contable": []})

    for df_part in [ep_agg, ip_agg, re_agg]:
        if len(open_cloud) == 0:
            open_cloud = df_part.copy()
        else:
            open_cloud = open_cloud.merge(
                df_part, on="cuenta_contable", how="outer"
            )

    for col in ["open_cloud_EP", "open_cloud_IP", "open_cloud_RE"]:
        if col not in open_cloud.columns:
            open_cloud[col] = 0.0

    open_cloud = open_cloud.fillna(0)
    open_cloud["carga_OP"] = (
        open_cloud["open_cloud_EP"]
        + open_cloud["open_cloud_IP"]
        + open_cloud["open_cloud_RE"]
    )

    open_cloud_agg = open_cloud[["cuenta_contable", "carga_OP"]].copy()

    result = bdep_agg.merge(open_cloud_agg, on="cuenta_contable", how="outer")
    result = result.merge(sap_agg, on="cuenta_contable", how="outer")
    result = result.fillna(0)

    for col in ["base_operativa", "carga_OP", "descarga_SAP"]:
        result[col] = result[col].round(2)

    result["dif_BDEP_vs_OP"] = (
        result["base_operativa"] - result["carga_OP"]
    ).round(2)
    result["dif_BDEP_vs_SAP"] = (
        result["base_operativa"] - result["descarga_SAP"]
    ).round(2)
    result["dif_OP_vs_SAP"] = (
        result["carga_OP"] - result["descarga_SAP"]
    ).round(2)

    def estado(row):
        if (
            abs(row["dif_BDEP_vs_OP"]) < 0.01
            and abs(row["dif_BDEP_vs_SAP"]) < 0.01
            and abs(row["dif_OP_vs_SAP"]) < 0.01
        ):
            return "OK"
        return "DIFERENCIA"

    result["estado"] = result.apply(estado, axis=1)
    result = result.sort_values("cuenta_contable").reset_index(drop=True)

    resumen = pd.DataFrame(
        [
            ["Cuentas conciliadas", len(result)],
            ["Cuentas OK", int((result["estado"] == "OK").sum())],
            [
                "Cuentas con diferencia",
                int((result["estado"] == "DIFERENCIA").sum()),
            ],
            ["Total BDEP Operativa", result["base_operativa"].sum()],
            ["Total Open Cloud", result["carga_OP"].sum()],
            ["Total SAP", result["descarga_SAP"].sum()],
            [
                "Diferencia total BDEP vs Open Cloud",
                result["base_operativa"].sum() - result["carga_OP"].sum(),
            ],
            [
                "Diferencia total BDEP vs SAP",
                result["base_operativa"].sum() - result["descarga_SAP"].sum(),
            ],
            [
                "Diferencia total Open Cloud vs SAP",
                result["carga_OP"].sum() - result["descarga_SAP"].sum(),
            ],
        ],
        columns=["Indicador", "Valor"],
    )

    output_label = safe_period_label(periodo)
    output_excel = (
        Path(output_dir)
        / f"conciliacion_cuentas_BDEP_SAP_OpenCloud_{output_label}.xlsx"
    )

    export_excel_ejecutivo(
        output_excel=output_excel,
        periodo=periodo,
        resumen=resumen,
        result=result,
        open_cloud=open_cloud,
    )

    return output_excel, resumen, result


def _extension_para_guardar(nombre_original):
    if nombre_original:
        suffix = Path(nombre_original).suffix.lower()
        if suffix in EXTENSIONES_PERMITIDAS:
            return suffix
    return ".xlsx"


async def _guardar_upload(archivo, tipo, nombre_original, workdir):
    contenido = await archivo.read()

    if not contenido:
        raise ValueError(f"{tipo.upper()}: el archivo llegó vacío.")

    extension = _extension_para_guardar(nombre_original)
    path = Path(workdir) / f"{tipo}{extension}"
    path.write_bytes(contenido)

    print(
        f"[ARCHIVO] tipo={tipo.upper()} nombre_original={nombre_original!r} "
        f"filename_connector={archivo.filename!r} bytes={len(contenido)}",
        flush=True,
    )

    return path


async def ejecutar_conciliacion(
    bdep: UploadFile,
    sap: UploadFile,
    ep: UploadFile,
    ip: UploadFile,
    re: UploadFile,
    nombres_originales=None,
):
    nombres_originales = nombres_originales or {}
    periodo = infer_period_from_names(nombres_originales)

    with tempfile.TemporaryDirectory(prefix="conciliacion_") as temp_dir:
        workdir = Path(temp_dir)

        archivos = {
            "bdep": await _guardar_upload(
                bdep, "bdep", nombres_originales.get("bdep"), workdir
            ),
            "sap": await _guardar_upload(
                sap, "sap", nombres_originales.get("sap"), workdir
            ),
            "ep": await _guardar_upload(
                ep, "ep", nombres_originales.get("ep"), workdir
            ),
            "ip": await _guardar_upload(
                ip, "ip", nombres_originales.get("ip"), workdir
            ),
            "re": await _guardar_upload(
                re, "re", nombres_originales.get("re"), workdir
            ),
        }

        output_excel, resumen, result = process_period(
            periodo=periodo,
            archivos=archivos,
            output_dir=workdir,
        )

        # Crear una copia temporal independiente antes de borrar workdir.
        final_dir = Path(tempfile.mkdtemp(prefix="resultado_conciliacion_"))
        final_path = final_dir / output_excel.name
        final_path.write_bytes(output_excel.read_bytes())

        return {
            "archivo": final_path,
            "periodo": periodo,
            "cuentas_conciliadas": int(len(result)),
            "cuentas_ok": int((result["estado"] == "OK").sum()),
            "cuentas_con_diferencia": int(
                (result["estado"] == "DIFERENCIA").sum()
            ),
        }
