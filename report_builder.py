"""
Логика автоматизации:
1. Берём лист "Данные КРГ", строим сводную по столбцу "ТоварнаяГруппа":
   сумма по "Дельта", "СуммаРозница", "СуммаПриход".
2. Копируем лист "Отчет по ТГ" (значения, формулы, форматирование, таблицу
   Excel) в новый лист "Готовый отчет по ТГ".
3. В новом листе:
   - в столбцы "Дельта КРГ, шт / руб. розница / руб. приход" (по умолчанию
     T, U, V) проставляем цифры из сводной по соответствующей Товарной группе;
   - в столбцах с фактическими остатками (по умолчанию I, J, K) там, где
     нашлась КРГ по этой Товарной группе, старое число заменяется формулой
     "=старое_число + соответствующая ячейка в T/U/V" (то есть прибавляем
     дельту КРГ). Там, где данных КРГ по группе нет — ячейка остаётся как
     была.
   - все остальные формулы листа ("Отчет по ТГ") переносятся без изменений,
     ничего не удаляется. Таблица Excel (ListObject) пересоздаётся под новым
     именем, а формулы со структурными ссылками ("Таблица[...]")
     переадресовываются на новую таблицу.
"""
import copy
from io import BytesIO

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo
from openpyxl.worksheet.filters import AutoFilter

YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

SHEET_DATA = "Данные КРГ"
SHEET_REPORT = "Отчет по ТГ"
SHEET_RESULT = "Готовый отчет по ТГ"

# Заголовки, которые ищем на листе "Данные КРГ"
COL_TG = "ТоварнаяГруппа"
COL_DELTA = "Дельта"
COL_SUM_ROZ = "СуммаРозница"
COL_SUM_PRIH = "СуммаПриход"

# Заголовки на листе "Отчет по ТГ", которые нам нужны.
HDR_TG_REPORT = "Товарная группа"
HDR_FACT_SHT = "Фактический остаток товара на момент сплошной инвентаризации, шт."
HDR_FACT_ROZ = "Фактический остаток в средневзвеш.розн. ценах, руб."
HDR_FACT_PRIH = "Фактический остаток в приходных ценах, руб."
HDR_KRG_SHT = "Дельта КРГ, шт"
HDR_KRG_ROZ = "Дельта КРГ в средневзвеш.розн. ценах, руб."
HDR_KRG_PRIH = "Дельта КРГ в приходных ценах, руб."


class ReportBuildError(Exception):
    pass


def _header_map(ws, header_row=1):
    """Возвращает {текст заголовка: буква столбца} для строки заголовков."""
    result = {}
    for cell in ws[header_row]:
        if cell.value is not None:
            result[str(cell.value).strip()] = cell.column_letter
    return result


def _find_col(headers, wanted, required=True, fallback_letter=None):
    if wanted in headers:
        return headers[wanted]
    # попробуем частичное совпадение
    for h, letter in headers.items():
        if wanted.lower() in h.lower() or h.lower() in wanted.lower():
            return letter
    if fallback_letter:
        return fallback_letter
    if required:
        raise ReportBuildError(f"Не найден столбец «{wanted}» на листе.")
    return None


def build_pivot(ws_data):
    """Сводная по Товарной группе: сумма Дельта / СуммаРозница / СуммаПриход."""
    headers = _header_map(ws_data, header_row=1)
    col_tg = _find_col(headers, COL_TG)
    col_delta = _find_col(headers, COL_DELTA)
    col_roz = _find_col(headers, COL_SUM_ROZ)
    col_prih = _find_col(headers, COL_SUM_PRIH)

    pivot = {}
    for row in ws_data.iter_rows(min_row=2):
        tg_cell = row[openpyxl.utils.column_index_from_string(col_tg) - 1]
        if tg_cell.value in (None, ""):
            continue
        tg = str(tg_cell.value).strip()

        def val(col_letter):
            c = row[openpyxl.utils.column_index_from_string(col_letter) - 1]
            v = c.value
            return v if isinstance(v, (int, float)) else 0

        d = pivot.setdefault(tg, {"delta": 0, "roz": 0, "prih": 0})
        d["delta"] += val(col_delta)
        d["roz"] += val(col_roz)
        d["prih"] += val(col_prih)
    return pivot


def _copy_cell_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.number_format = copy.copy(src_cell.number_format)
        dst_cell.protection = copy.copy(src_cell.protection)
        dst_cell.alignment = copy.copy(src_cell.alignment)


def _clone_sheet(wb, src_ws, new_title):
    if new_title in wb.sheetnames:
        del wb[new_title]
    new_ws = wb.create_sheet(new_title)

    # значения + формат ячеек
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = new_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            _copy_cell_style(cell, new_cell)

    # ширины столбцов
    for key, dim in src_ws.column_dimensions.items():
        new_ws.column_dimensions[key].width = dim.width
        new_ws.column_dimensions[key].hidden = dim.hidden

    # высоты строк
    for key, dim in src_ws.row_dimensions.items():
        new_ws.row_dimensions[key].height = dim.height

    # объединённые ячейки
    for merged_range in src_ws.merged_cells.ranges:
        new_ws.merge_cells(str(merged_range))

    # заморозка областей / прочее
    new_ws.freeze_panes = src_ws.freeze_panes
    new_ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines

    # условное форматирование
    for cf_range, rules in src_ws.conditional_formatting._cf_rules.items():
        for rule in rules:
            new_ws.conditional_formatting.add(str(cf_range.sqref), copy.copy(rule))

    return new_ws


def _rewrite_table_formulas(ws, old_table_name, new_table_name):
    prefix_old = f"{old_table_name}["
    prefix_new = f"{new_table_name}["
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                if prefix_old in cell.value:
                    cell.value = cell.value.replace(prefix_old, prefix_new)


def _rebuild_table(ws, src_table, new_table_name):
    """Пересоздаёт объект Excel-таблицы (ListObject) на новом листе."""
    new_columns = []
    for col in src_table.tableColumns:
        new_col = TableColumn(id=col.id, name=col.name)
        new_col.totalsRowFunction = col.totalsRowFunction
        new_col.totalsRowLabel = col.totalsRowLabel
        if col.calculatedColumnFormula is not None:
            f = copy.copy(col.calculatedColumnFormula)
            if f.attr_text:
                f.attr_text = f.attr_text.replace(
                    f"{src_table.name}[", f"{new_table_name}["
                )
            new_col.calculatedColumnFormula = f
        if col.totalsRowFormula is not None:
            f = copy.copy(col.totalsRowFormula)
            if f.attr_text:
                f.attr_text = f.attr_text.replace(
                    f"{src_table.name}[", f"{new_table_name}["
                )
            new_col.totalsRowFormula = f
        new_columns.append(new_col)

    new_table = Table(
        displayName=new_table_name,
        name=new_table_name,
        ref=src_table.ref,
        headerRowCount=src_table.headerRowCount,
        totalsRowCount=src_table.totalsRowCount,
        tableColumns=new_columns,
    )
    if src_table.tableStyleInfo is not None:
        new_table.tableStyleInfo = TableStyleInfo(
            name=src_table.tableStyleInfo.name,
            showFirstColumn=src_table.tableStyleInfo.showFirstColumn,
            showLastColumn=src_table.tableStyleInfo.showLastColumn,
            showRowStripes=src_table.tableStyleInfo.showRowStripes,
            showColumnStripes=src_table.tableStyleInfo.showColumnStripes,
        )

    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(src_table.ref)
    if src_table.totalsRowCount:
        max_row -= 1
    af_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    new_table.autoFilter = AutoFilter(ref=af_ref)

    ws.add_table(new_table)
    return new_table


def build_report(file_bytes: bytes) -> bytes:
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=False)

    missing = [s for s in (SHEET_DATA, SHEET_REPORT) if s not in wb.sheetnames]
    if missing:
        raise ReportBuildError(
            "В файле не найдены обязательные листы: " + ", ".join(missing)
        )

    ws_data = wb[SHEET_DATA]
    ws_report = wb[SHEET_REPORT]

    if not ws_report.tables:
        raise ReportBuildError(
            f"На листе «{SHEET_REPORT}» не найдена таблица Excel (ListObject)."
        )
    src_table_name = list(ws_report.tables.keys())[0]
    src_table = ws_report.tables[src_table_name]

    # 1. сводная по Данные КРГ
    pivot = build_pivot(ws_data)

    # 2. клонируем лист "Отчет по ТГ" -> "Готовый отчет по ТГ"
    new_ws = _clone_sheet(wb, ws_report, SHEET_RESULT)

    # 3. пересчитываем структурные ссылки формул на новую таблицу
    new_table_name = src_table_name
    n = 2
    while new_table_name in {t for ws_ in wb.worksheets for t in ws_.tables.keys()}:
        new_table_name = f"{src_table_name}Готовый{n}"
        n += 1
    _rewrite_table_formulas(new_ws, src_table_name, new_table_name)

    # 4. находим нужные столбцы по заголовкам
    headers = _header_map(new_ws, header_row=1)
    col_tg = _find_col(headers, HDR_TG_REPORT, fallback_letter="C")
    col_fact_sht = _find_col(headers, HDR_FACT_SHT, fallback_letter="I")
    col_fact_roz = _find_col(headers, HDR_FACT_ROZ, fallback_letter="J")
    col_fact_prih = _find_col(headers, HDR_FACT_PRIH, fallback_letter="K")
    col_krg_sht = _find_col(headers, HDR_KRG_SHT, fallback_letter="T")
    col_krg_roz = _find_col(headers, HDR_KRG_ROZ, fallback_letter="U")
    col_krg_prih = _find_col(headers, HDR_KRG_PRIH, fallback_letter="V")

    header_row = 1
    data_first_row = header_row + 1
    data_last_row = data_first_row + src_table.headerRowCount  # placeholder, fixed below

    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(src_table.ref)
    data_last_row = max_row - (src_table.totalsRowCount or 0)

    matched_groups, unmatched_groups = [], []

    for r in range(data_first_row, data_last_row + 1):
        tg_cell = new_ws[f"{col_tg}{r}"]
        tg_name = str(tg_cell.value).strip() if tg_cell.value is not None else ""
        if not tg_name:
            continue
        data = pivot.get(tg_name)
        if data is None:
            unmatched_groups.append(tg_name)
            continue
        matched_groups.append(tg_name)

        new_ws[f"{col_krg_sht}{r}"] = data["delta"]
        new_ws[f"{col_krg_roz}{r}"] = data["roz"]
        new_ws[f"{col_krg_prih}{r}"] = data["prih"]

        for col_fact, col_krg in (
            (col_fact_sht, col_krg_sht),
            (col_fact_roz, col_krg_roz),
            (col_fact_prih, col_krg_prih),
        ):
            fact_cell = new_ws[f"{col_fact}{r}"]
            old_value = fact_cell.value
            if isinstance(old_value, str) and old_value.startswith("="):
                continue
            old_value = old_value if isinstance(old_value, (int, float)) else 0
            fact_cell.value = f"={old_value}+{col_krg}{r}"
            fact_cell.fill = copy.copy(YELLOW_FILL)

    # столбцы, для которых данных КРГ по этой группе не было — не трогаем.

    # 5. заголовков групп, встретившихся в "Данные КРГ", но не найденных
    #    в "Отчет по ТГ" — не игнорируем, поднимаем предупреждение через атрибут результата.
    pivot_tgs = set(pivot.keys())
    report_tgs = set(matched_groups) | set(unmatched_groups)
    not_in_report = sorted(pivot_tgs - report_tgs)

    # 6. пересобираем таблицу Excel на новом листе
    _rebuild_table(new_ws, src_table, new_table_name)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), {
        "matched": sorted(set(matched_groups)),
        "unmatched": sorted(set(unmatched_groups)),
        "not_in_report": not_in_report,
    }