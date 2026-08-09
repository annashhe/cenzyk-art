#!/usr/bin/env python3
"""Сборка шаблона калькуляции проекта (рентабельность + цена КП)."""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

wb = Workbook()

thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
fill_header = PatternFill("solid", fgColor="2F3E46")
fill_section = PatternFill("solid", fgColor="52796F")
fill_input = PatternFill("solid", fgColor="FFF3BF")
fill_calc = PatternFill("solid", fgColor="E8EEF2")
fill_good = PatternFill("solid", fgColor="D8F3DC")
fill_warn = PatternFill("solid", fgColor="FFE8A3")
fill_bad = PatternFill("solid", fgColor="F8D7DA")
fill_soft = PatternFill("solid", fgColor="F0F4F3")
font_header = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
font_section = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
font_title = Font(name="Calibri", bold=True, size=16, color="2F3E46")
font_label = Font(name="Calibri", size=11, color="2F3E46")
font_bold = Font(name="Calibri", bold=True, size=11, color="2F3E46")
font_big = Font(name="Calibri", bold=True, size=14, color="2F3E46")
font_note = Font(name="Calibri", italic=True, size=9, color="666666")
money = '#,##0.00" ₽"'
pct = "0.0%"
align_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_l = Alignment(horizontal="left", vertical="center", wrap_text=True)


def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def header_cell(cell, text, fill=fill_header, font=font_header):
    cell.value = text
    cell.fill = fill
    cell.font = font
    cell.border = thin
    cell.alignment = align_c


def input_cell(cell, value=None, fmt=None):
    if value is not None:
        cell.value = value
    cell.fill = fill_input
    cell.border = thin
    if fmt:
        cell.number_format = fmt


def calc_cell(cell, value=None, fmt=None, bold=False):
    if value is not None:
        cell.value = value
    cell.fill = fill_calc
    cell.border = thin
    cell.font = font_bold if bold else font_label
    if fmt:
        cell.number_format = fmt


# =============================================================================
# Параметры
# =============================================================================
ws = wb.active
ws.title = "Параметры"

ws["B2"] = "ПАРАМЕТРЫ РАСЧЁТА"
ws["B2"].font = font_title
ws.merge_cells("B2:D2")
ws["B3"] = "Меняйте только жёлтые ячейки. Все калькуляции подтягивают значения отсюда."
ws["B3"].font = font_note
ws.merge_cells("B3:F3")

for col, title in [("B", "Параметр"), ("C", "Значение"), ("D", "Комментарий")]:
    header_cell(ws[f"{col}5"], title)

params = [
    (6, "Наценка менеджера на закупки", 0.15, "От цены поставщика. Ваша прибыль как менеджера."),
    (7, "Целевая маржа проекта (от выручки)", 0.20, "Чистая маржа после НПД. Диапазон обычно 10%…30%+."),
    (8, "Минимальная допустимая маржа", 0.10, "Ниже — статус «НИЖЕ МИНИМУМА»."),
    (9, "Налог НПД (самозанятость)", 0.04, "4% — физлица, 6% — юрлица/ИП."),
    (10, "Аренда, ₽ / рабочий день", 2000, "Ставка аренды цеха на рабочий день."),
    (11, "Рабочих дней в месяце", 22, "Справочно для пересчёта постоянных расходов."),
    (12, "Агентские, % от цены КП", 0.00, "Если есть посредник — поставьте %, иначе 0."),
]
for r, name, val, note in params:
    ws[f"B{r}"] = name
    ws[f"B{r}"].fill = fill_soft
    ws[f"B{r}"].border = thin
    input_cell(ws[f"C{r}"], val, pct if r in (6, 7, 8, 9, 12) else "#,##0.00")
    ws[f"C{r}"].font = font_bold
    ws[f"D{r}"] = note
    ws[f"D{r}"].font = font_note
    ws[f"D{r}"].border = thin

ws["B14"] = "Авто-коэффициенты"
ws["B14"].font = font_bold
ws["B15"] = "Множитель цены при целевой марже"
calc_cell(ws["C15"], "=1/(1-C9-C7-C12)", "0.000", bold=True)
ws["D15"] = "Цена КП ≈ Себестоимость × коэффициент"
ws["B16"] = "Множитель цены при минимальной марже"
calc_cell(ws["C16"], "=1/(1-C9-C8-C12)", "0.000", bold=True)
for r in (15, 16):
    ws[f"B{r}"].fill = fill_soft
    ws[f"B{r}"].border = thin
    ws[f"D{r}"].border = thin
    ws[f"D{r}"].font = font_note

ws["B18"] = "Тип изделия (список для шапки)"
ws["B18"].font = font_bold
for i, t in enumerate(["Ворота", "Забор", "Лестница", "Мебель", "Интерьер", "Прочее"], start=19):
    ws[f"B{i}"] = t
    ws[f"B{i}"].fill = fill_soft
    ws[f"B{i}"].border = thin

ws["B26"] = "Правило маржи"
ws["B26"].font = font_bold
ws["B27"] = (
    "Себестоимость C = реальные выплаты (материалы по цене поставщика, труд, транспорт, аренда, расходники, субподряд).\n"
    "Наценка менеджера = % только на тип «закупка» — это часть вашей прибыли, не расход.\n"
    "Цена для целевой маржи P = C / (1 − НПД% − целевая маржа% − агентские%).\n"
    "Прибыль = P − C − P×НПД − P×агентские.  Маржа % = Прибыль / P.\n"
    "Не добавляйте «+15% на закупки» и «+20% маржа» как две наценки на одну базу: 15% уже внутри итоговой прибыли."
)
ws["B27"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("B27:F32")
ws["B27"].fill = fill_soft
ws.row_dimensions[27].height = 96
set_col_widths(ws, {"A": 3, "B": 48, "C": 14, "D": 70, "E": 12, "F": 12})

# =============================================================================
# Сводка
# =============================================================================
ws = wb.create_sheet("Сводка", 0)
ws["B2"] = "СВОДКА ЗАКАЗА — рентабельность и цена КП"
ws["B2"].font = font_title
ws.merge_cells("B2:G2")
ws["B3"] = "Сначала заполните Металл и Калькуляцию. Жёлтое — ввод. Остальное считается автоматически."
ws["B3"].font = font_note
ws.merge_cells("B3:G3")

header_cell(ws["B5"], "ШАПКА ЗАКАЗА", fill_section, font_section)
ws.merge_cells("B5:C5")
for r, lab in [
    (6, "Объект"),
    (7, "Конструкция / изделие"),
    (8, "Тип"),
    (9, "Дата расчёта"),
    (10, "Клиент (опц.)"),
]:
    ws[f"B{r}"] = lab
    ws[f"B{r}"].fill = fill_soft
    ws[f"B{r}"].border = thin
    input_cell(ws[f"C{r}"])
ws["C8"] = "Ворота"
ws["C9"] = "=TODAY()"
ws["C9"].number_format = "DD.MM.YYYY"
dv = DataValidation(type="list", formula1="Параметры!$B$19:$B$24", allow_blank=True)
ws.add_data_validation(dv)
dv.add(ws["C8"])

header_cell(ws["E5"], "ПАРАМЕТРЫ (зеркало)", fill_section, font_section)
ws.merge_cells("E5:F5")
for r, lab, formula, fmt in [
    (6, "Наценка менеджера", "=Параметры!C6", pct),
    (7, "Целевая маржа", "=Параметры!C7", pct),
    (8, "Мин. маржа", "=Параметры!C8", pct),
    (9, "НПД", "=Параметры!C9", pct),
    (10, "Аренда / день", "=Параметры!C10", money),
    (11, "Агентские", "=Параметры!C12", pct),
]:
    ws[f"E{r}"] = lab
    ws[f"E{r}"].fill = fill_soft
    ws[f"E{r}"].border = thin
    calc_cell(ws[f"F{r}"], formula, fmt)

header_cell(ws["B13"], "P&L ПРОЕКТА", fill_header, font_header)
ws.merge_cells("B13:D13")
for r, lab, formula, fmt, bold in [
    (14, "Себестоимость проекта (реальные затраты)", "=Калькуляция!O8", money, True),
    (15, "  в т.ч. закупки у поставщиков", "=Калькуляция!O9", money, False),
    (16, "  в т.ч. труд", "=Калькуляция!O10", money, False),
    (17, "  в т.ч. транспорт", "=Калькуляция!O11", money, False),
    (18, "  в т.ч. аренда цеха", "=Калькуляция!O12", money, False),
    (19, "  в т.ч. услуги / прочее", "=Калькуляция!O13", money, False),
    (20, "Наценка менеджера на закупки (часть прибыли)", "=Калькуляция!O14", money, True),
    (21, "Дни цеха (по строкам аренды)", "=Калькуляция!O15", "0.00", False),
]:
    ws[f"B{r}"] = lab
    ws[f"B{r}"].fill = fill_soft
    ws[f"B{r}"].border = thin
    ws[f"B{r}"].font = font_bold if bold else font_label
    ws.merge_cells(f"B{r}:C{r}")
    calc_cell(ws[f"D{r}"], formula, fmt, bold=bold)

header_cell(ws["B23"], "ЦЕНА И МАРЖА", fill_header, font_header)
ws.merge_cells("B23:D23")

ws["B24"] = "Цена КП (то, что выставите клиенту)"
ws["B24"].fill = fill_input
ws["B24"].border = thin
ws["B24"].font = font_bold
ws.merge_cells("B24:C24")
input_cell(ws["D24"], 0, money)
ws["D24"].font = font_big

rows_price = [
    (25, "Рекомендуемая цена при целевой марже", "=IF((1-F9-F7-F11)<=0,\"Ошибка %\",D14/(1-F9-F7-F11))", fill_good),
    (26, "Минимальная цена (при мин. марже)", "=IF((1-F9-F8-F11)<=0,\"Ошибка %\",D14/(1-F9-F8-F11))", fill_warn),
    (27, "НПД с цены КП", "=D24*F9", fill_calc),
    (28, "Агентские с цены КП", "=D24*F11", fill_calc),
    (29, "Чистая прибыль при цене КП", "=D24-D14-D27-D28", fill_calc),
    (30, "Маржа % при цене КП (от выручки)", "=IF(D24=0,0,D29/D24)", fill_calc),
]
for r, lab, formula, fill in rows_price:
    ws[f"B{r}"] = lab
    ws[f"B{r}"].fill = fill_soft
    ws[f"B{r}"].border = thin
    if r in (29, 30):
        ws[f"B{r}"].font = font_bold
    ws.merge_cells(f"B{r}:C{r}")
    calc_cell(ws[f"D{r}"], formula, pct if r == 30 else money, bold=(r in (25, 29, 30)))
    ws[f"D{r}"].fill = fill
    if r == 25:
        ws[f"D{r}"].font = font_big

ws["B31"] = "Статус"
ws["B31"].fill = fill_soft
ws["B31"].border = thin
ws["B31"].font = font_bold
ws.merge_cells("B31:C31")
calc_cell(
    ws["D31"],
    '=IF(D24=0,"Введите цену КП",IF(D30<F8,"НИЖЕ МИНИМУМА",IF(D30<F7,"НИЖЕ ЦЕЛИ","ЦЕЛЬ ДОСТИГНУТА")))',
    bold=True,
)
ws.conditional_formatting.add("D31", FormulaRule(formula=['$D$31="ЦЕЛЬ ДОСТИГНУТА"'], fill=fill_good))
ws.conditional_formatting.add("D31", FormulaRule(formula=['$D$31="НИЖЕ ЦЕЛИ"'], fill=fill_warn))
ws.conditional_formatting.add("D31", FormulaRule(formula=['$D$31="НИЖЕ МИНИМУМА"'], fill=fill_bad))
ws.conditional_formatting.add("D30", CellIsRule(operator="lessThan", formula=["$F$8"], fill=fill_bad))
ws.conditional_formatting.add("D30", FormulaRule(formula=["AND($D$30>=$F$8,$D$30<$F$7)"], fill=fill_warn))
ws.conditional_formatting.add("D30", CellIsRule(operator="greaterThanOrEqual", formula=["$F$7"], fill=fill_good))

header_cell(ws["B33"], "ГИБКАЯ МАРЖА — таблица цен", fill_section, font_section)
ws.merge_cells("B33:E33")
for col, title in [("B", "Целевая маржа"), ("C", "Цена КП"), ("D", "НПД"), ("E", "Прибыль")]:
    header_cell(ws[f"{col}34"], title)
for i, m in enumerate([0.10, 0.15, 0.20, 0.25, 0.30, 0.35]):
    r = 35 + i
    ws[f"B{r}"] = m
    ws[f"B{r}"].number_format = pct
    ws[f"B{r}"].border = thin
    ws[f"B{r}"].fill = fill_good if abs(m - 0.20) < 1e-9 else fill_soft
    calc_cell(ws[f"C{r}"], f"=IF((1-$F$9-B{r}-$F$11)<=0,0,$D$14/(1-$F$9-B{r}-$F$11))", money)
    calc_cell(ws[f"D{r}"], f"=C{r}*$F$9", money)
    calc_cell(ws[f"E{r}"], f"=C{r}-$D$14-D{r}-C{r}*$F$11", money)
    if abs(m - 0.20) < 1e-9:
        for col in ["C", "D", "E"]:
            ws[f"{col}{r}"].fill = fill_good

ws["B42"] = "Строка 20% — базовый ориентир. Целевую маржу меняйте на листе Параметры — пересчитается рекомендуемая цена."
ws["B42"].font = font_note
ws.merge_cells("B42:E42")

header_cell(ws["B44"], "КОНТРОЛЬ", fill_section, font_section)
ws.merge_cells("B44:D44")
for r, lab, formula, fmt in [
    (45, "Доля наценки менеджера в прибыли (при цене КП)", "=IF(D29<=0,0,D20/D29)", pct),
    (46, "Выручка на день цеха", "=IF(D21=0,0,D24/D21)", money),
    (47, "Прибыль на день цеха", "=IF(D21=0,0,D29/D21)", money),
]:
    ws[f"B{r}"] = lab
    ws[f"B{r}"].fill = fill_soft
    ws[f"B{r}"].border = thin
    ws.merge_cells(f"B{r}:C{r}")
    calc_cell(ws[f"D{r}"], formula, fmt)

ws["B49"] = (
    "Порядок: 1) Металл и Калькуляция → 2) рекомендуемая цена → 3) вписать цену КП в D24 → 4) проверить статус."
)
ws["B49"].font = font_note
ws.merge_cells("B49:G49")
set_col_widths(ws, {"A": 3, "B": 44, "C": 22, "D": 16, "E": 22, "F": 14, "G": 12})
ws.freeze_panes = "B5"

# =============================================================================
# Калькуляция
# =============================================================================
ws = wb.create_sheet("Калькуляция", 1)
ws["B2"] = "КАЛЬКУЛЯЦИЯ ПРОЕКТА"
ws["B2"].font = font_title
ws.merge_cells("B2:L2")
ws["B3"] = (
    "Жёлтые — ввод. Тип «закупка» получает наценку менеджера из Параметров. "
    "Аренда подтягивает ставку/день. Итоги справа — для Сводки."
)
ws["B3"].font = font_note
ws.merge_cells("B3:L3")

header_cell(ws["N5"], "ИТОГИ ДЛЯ СВОДКИ", fill_header, font_header)
ws.merge_cells("N5:O5")
totals = [
    (6, "Контроль суммы себест.", "=SUM(H16:H250)", money),
    (7, "Контроль наценки менеджера", "=SUM(I16:I250)", money),
    (8, "Себестоимость проекта", "=O6", money),
    (9, "Закупки", '=SUMIF(D16:D250,"закупка",H16:H250)', money),
    (10, "Труд", '=SUMIF(D16:D250,"труд",H16:H250)', money),
    (11, "Транспорт", '=SUMIF(D16:D250,"транспорт",H16:H250)', money),
    (12, "Аренда", '=SUMIF(D16:D250,"аренда",H16:H250)', money),
    (13, "Услуги/прочее", "=O6-O9-O10-O11-O12", money),
    (14, "Наценка менеджера", "=O7", money),
    (15, "Дни цеха", '=SUMIF(D16:D250,"аренда",F16:F250)', "0.00"),
]
for r, lab, formula, fmt in totals:
    ws[f"N{r}"] = lab
    ws[f"N{r}"].fill = fill_soft
    ws[f"N{r}"].border = thin
    calc_cell(ws[f"O{r}"], formula, fmt, bold=True)

for col, title in [
    ("B", "Этап"),
    ("C", "Статья / обоснование"),
    ("D", "Тип"),
    ("E", "Цена / ставка"),
    ("F", "Кол-во"),
    ("G", "Ед."),
    ("H", "Себестоимость"),
    ("I", "Наценка менеджера"),
    ("J", "В базу КП"),
    ("L", "Примечание"),
]:
    header_cell(ws[f"{col}15"], title)

dv_type = DataValidation(
    type="list",
    formula1='"закупка,труд,транспорт,аренда,услуга,прочее"',
    allow_blank=True,
)
ws.add_data_validation(dv_type)

rows = []


def add_stage(name):
    rows.append(("__SECTION__", name))


def add(stage, item, typ, price, qty, unit, note="", special=None):
    rows.append((stage, item, typ, price, qty, unit, note, special))


add_stage("1. Проектирование / замер")
add("Проектирование", "Аренда на замер", "аренда", 2000, 0, "день", "", "rent")
add("Проектирование", "Транспорт по городу", "транспорт", 2000, 0, "выезд")
add("Проектирование", "Транспорт по области", "транспорт", 4000, 0, "выезд")
add("Проектирование", "Замер", "услуга", 6000, 0, "услуга")
add("Проектирование", "Консультация", "услуга", 6000, 0, "услуга")
add("Проектирование", "Проектирование, эскизы", "услуга", 13000, 0, "услуга")

add_stage("2. Снять лекала / подготовка")
add("Лекала", "Аренда", "аренда", 2000, 0, "день", "", "rent")
add("Лекала", "Мастер", "труд", 7000, 0, "день")
add("Лекала", "Рабочий", "труд", 5000, 0, "день")
add("Лекала", "Транспорт цех → объект", "транспорт", 2000, 0, "рейс")
add("Лекала", "Транспорт объект → цех", "транспорт", 2000, 0, "рейс")
add("Лекала", "Картоны", "закупка", 300, 0, "шт")
add("Лекала", "Расходники", "закупка", 250, 0, "набор")
add("Лекала", "Дерево", "закупка", 1500, 0, "набор")

add_stage("3. Закуп металла")
add("Металл", "Транспорт: микроавтобус (2ч)", "транспорт", 2000, 0, "рейс")
add("Металл", "Транспорт: бортовая 4м (1ч)", "транспорт", 1500, 0, "рейс")
add("Металл", "Транспорт: бортовая 6м (1ч)", "транспорт", 3500, 0, "рейс")
add("Металл", "Металлопрокат (лист Металл)", "закупка", None, 1, "сводка", "", "metal_cost")
add("Металл", "Лазер / плазма (лист Металл)", "услуга", None, 1, "сводка", "", "metal_plasma")
add("Металл", "Ков. элементы / спец. прокат", "закупка", None, 1, "сводка", "", "metal_forge")
add("Металл", "Токарные работы (лист Металл)", "услуга", None, 1, "сводка", "", "metal_turn")

add_stage("4. Производство")
add("Производство", "Аренда цеха", "аренда", 2000, 0, "день", "", "rent")
add("Производство", "Оплата труда: мастер", "труд", 6000, 0, "день", "ФИО в примечании")
add("Производство", "Оплата труда: рабочий", "труд", 4000, 0, "день")
add("Производство", "Оплата труда: доп.1", "труд", 6000, 0, "день")
add("Производство", "Оплата труда: доп.2", "труд", 4000, 0, "день")

add_stage("5. Предмонтаж")
add("Предмонтаж", "Аренда", "аренда", 2000, 0, "день", "", "rent")
add("Предмонтаж", "Монтажники", "труд", 7300, 0, "день")
add("Предмонтаж", "Помощник", "труд", 5000, 0, "день")
add("Предмонтаж", "Транспорт цех → объект", "транспорт", 2000, 0, "рейс")
add("Предмонтаж", "Транспорт объект → цех", "транспорт", 2000, 0, "рейс")
add("Предмонтаж", "Такси", "транспорт", 1500, 0, "поездка")

add_stage("6. Покраска эмалью")
add("Покраска эмаль", "Аренда", "аренда", 2000, 0, "день", "", "rent")
add("Покраска эмаль", "Маляр", "труд", 3000, 0, "день")
add("Покраска эмаль", "Рабочий", "труд", 5000, 0, "день")
add("Покраска эмаль", "Респиратор", "закупка", 500, 0, "шт")
add("Покраска эмаль", "Грунт", "закупка", 600, 0, "шт")
add("Покраска эмаль", "Лак", "закупка", 1000, 0, "шт")
add("Покраска эмаль", "Эмаль 0,8л", "закупка", 1200, 0, "шт")
add("Покраска эмаль", "Эмаль 2,0л", "закупка", 3500, 0, "шт")
add("Покраска эмаль", "Мет.щётки, кисти", "закупка", 300, 0, "набор")
add("Покраска эмаль", "Растворитель", "закупка", 370, 0, "шт")
add("Покраска эмаль", "Холодный цинк", "закупка", 2500, 0, "шт")
add("Покраска эмаль", "Транспорт микроавтобус", "транспорт", 2000, 0, "рейс")
add("Покраска эмаль", "Транспорт бортовая 4м", "транспорт", 1500, 0, "рейс")
add("Покраска эмаль", "Транспорт бортовая 6м", "транспорт", 3500, 0, "рейс")

add_stage("7. Пескоструй")
add("Пескоструй", "Услуга пескоструй", "услуга", 900, 0, "м")
add("Пескоструй", "Транспорт на пескоструй", "транспорт", 2000, 0, "рейс")
add("Пескоструй", "Транспорт с пескоструя", "транспорт", 2000, 0, "рейс")

add_stage("8. Порошковая покраска")
add("Порошок", "Покраска обычная", "услуга", 2000, 0, "услуга")
add("Порошок", "Покраска с цинком", "услуга", 2400, 0, "услуга")
add("Порошок", "Покраска шагрень", "услуга", 2300, 0, "услуга")
add("Порошок", "Покраска антик (молотковая)", "услуга", 2200, 0, "услуга")
add("Порошок", "Транспорт на покраску", "транспорт", 2000, 0, "рейс")
add("Порошок", "Транспорт с покраски", "транспорт", 2000, 0, "рейс")
add("Порошок", "Транспорт людей", "транспорт", 500, 0, "поездка")

add_stage("9. Монтаж")
add("Монтаж", "Аренда", "аренда", 2000, 0, "день", "", "rent")
add("Монтаж", "Монтажники", "труд", 5000, 0, "день")
add("Монтаж", "Помощник", "труд", 4000, 0, "день")
add("Монтаж", "Крепежи, хим.анкер", "закупка", 2000, 0, "набор")
add("Монтаж", "Транспорт рабочих", "транспорт", 500, 0, "поездка")
add("Монтаж", "Транспорт оборудования на объект", "транспорт", 2000, 0, "рейс")
add("Монтаж", "Транспорт оборудования в цех", "транспорт", 2000, 0, "рейс")

add_stage("10. Расходники в день")
add("Расходники/день", "Абразивы", "закупка", 600, 0, "день")
add("Расходники/день", "Кислород", "закупка", 330, 0, "день", "ориентир: баллон / срок")
add("Расходники/день", "Углекислота", "закупка", 65, 0, "день")
add("Расходники/день", "Пропан (газ)", "закупка", 690, 0, "день")
add("Расходники/день", "Ковка (уголь)", "закупка", 3100, 0, "день")
add("Расходники/день", "Амортизация инструмента", "прочее", 740, 0, "день")

r = 16
for item in rows:
    if item[0] == "__SECTION__":
        ws.merge_cells(f"B{r}:L{r}")
        ws[f"B{r}"] = item[1]
        for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "L"]:
            ws[f"{col}{r}"].fill = fill_section
            ws[f"{col}{r}"].font = font_section
        r += 1
        continue

    stage, name, typ, price, qty, unit, note, special = item
    ws[f"B{r}"] = stage
    ws[f"C{r}"] = name
    ws[f"D{r}"] = typ
    ws[f"F{r}"] = qty
    ws[f"G{r}"] = unit
    ws[f"L{r}"] = note

    if special == "rent":
        ws[f"E{r}"] = "=Параметры!$C$10"
        ws[f"L{r}"] = "Ставка аренды из Параметры"
    elif special == "metal_cost":
        ws[f"E{r}"] = "=Металл!H19"
        ws[f"F{r}"] = 1
        ws[f"L{r}"] = "Себестоимость проката (цена поставщика)"
    elif special == "metal_plasma":
        ws[f"E{r}"] = "=Металл!H28"
        ws[f"F{r}"] = 1
    elif special == "metal_forge":
        ws[f"E{r}"] = "=Металл!H39"
        ws[f"F{r}"] = 1
    elif special == "metal_turn":
        ws[f"E{r}"] = "=Металл!H48"
        ws[f"F{r}"] = 1
    else:
        ws[f"E{r}"] = price if price is not None else 0

    ws[f"H{r}"] = f'=IF(OR(E{r}="",F{r}=""),0,E{r}*F{r})'
    ws[f"I{r}"] = f'=IF(D{r}="закупка",H{r}*Параметры!$C$6,0)'
    ws[f"J{r}"] = f"=H{r}+I{r}"

    for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "L"]:
        ws[f"{col}{r}"].border = thin
    for col in ["C", "D", "E", "F", "G", "L"]:
        ws[f"{col}{r}"].fill = fill_input
    if special in ("rent", "metal_cost", "metal_plasma", "metal_forge", "metal_turn"):
        ws[f"E{r}"].fill = fill_calc
    if special and str(special).startswith("metal"):
        ws[f"F{r}"].fill = fill_calc
    for col in ["H", "I", "J"]:
        ws[f"{col}{r}"].fill = fill_calc
        ws[f"{col}{r}"].number_format = money
    ws[f"E{r}"].number_format = money
    dv_type.add(ws[f"D{r}"])
    r += 1

last_data_row = r - 1
for r in range(last_data_row + 1, last_data_row + 21):
    ws[f"H{r}"] = f'=IF(OR(E{r}="",F{r}=""),0,E{r}*F{r})'
    ws[f"I{r}"] = f'=IF(D{r}="закупка",H{r}*Параметры!$C$6,0)'
    ws[f"J{r}"] = f"=H{r}+I{r}"
    for col in ["B", "C", "D", "E", "F", "G", "H", "I", "J", "L"]:
        ws[f"{col}{r}"].border = thin
    for col in ["C", "D", "E", "F", "G", "L"]:
        ws[f"{col}{r}"].fill = fill_input
    for col in ["H", "I", "J"]:
        ws[f"{col}{r}"].fill = fill_calc
        ws[f"{col}{r}"].number_format = money
    ws[f"E{r}"].number_format = money
    dv_type.add(ws[f"D{r}"])

ws[f"B{last_data_row + 22}"] = (
    "Пустые строки ниже — для своих позиций. Тип «закупка» включает наценку менеджера автоматически."
)
ws[f"B{last_data_row + 22}"].font = font_note

set_col_widths(
    ws,
    {
        "A": 3,
        "B": 16,
        "C": 42,
        "D": 12,
        "E": 14,
        "F": 10,
        "G": 10,
        "H": 14,
        "I": 16,
        "J": 12,
        "K": 3,
        "L": 36,
        "M": 3,
        "N": 28,
        "O": 14,
    },
)
ws.freeze_panes = "B16"
ws.auto_filter.ref = f"B15:L{last_data_row + 20}"

# =============================================================================
# Металл
# =============================================================================
ws = wb.create_sheet("Металл", 2)
ws["B2"] = "МЕТАЛЛ И КОМПЛЕКТУЮЩИЕ"
ws["B2"].font = font_title
ws.merge_cells("B2:I2")
ws["B3"] = (
    "F — цена поставщика (обновляйте по прайсу). D — кол-во под заказ. "
    "Наценка менеджера: себестоимость × % (исправлена ошибка старого шаблона G×15%)."
)
ws["B3"].font = font_note
ws.merge_cells("B3:I3")

header_cell(ws["B4"], "ПРОКАТ", fill_section, font_section)
ws.merge_cells("B4:I4")
for col, title in [
    ("B", "Прокат"),
    ("C", "Сечение"),
    ("D", "Кол-во"),
    ("E", "Ед."),
    ("F", "Цена поставщика"),
    ("G", "С наценкой менеджера"),
    ("H", "Себестоимость"),
    ("I", "Наценка менеджера"),
]:
    header_cell(ws[f"{col}5"], title)

prokat = [
    ("Труба профильная", "15х15х1.5", 0, "шт", 1200),
    ("Труба профильная", "15х15х2.0", 0, "шт", 650),
    ("Труба профильная", "20х20х1.5", 0, "шт", 575.7),
    ("Труба профильная", "20х20х2.0", 0, "шт", None),
    ("Труба профильная", "50х25", 0, "шт", 1200),
    ("Труба профильная", "40х20", 0, "шт", 850),
    ("Круг", "40х40х1.5", 0, "шт", 750),
    ("Квадрат", "40х40х2.0", 0, "шт", 690),
    ("Квадрат", "40х40х3.0", 0, "шт", None),
    ("Лист", "10х10", 0, "шт", None),
    ("Лист", "12х12", 0, "шт", None),
    ("Лист", "14х14", 0, "шт", None),
    ("Лист", "16х16", 0, "шт", None),
]
for i, (name, sect, qty, unit, price) in enumerate(prokat):
    r = 6 + i
    ws[f"B{r}"] = name
    ws[f"C{r}"] = sect
    ws[f"D{r}"] = qty
    ws[f"E{r}"] = unit
    if price is not None:
        ws[f"F{r}"] = price
    ws[f"H{r}"] = f'=IF(OR(D{r}="",F{r}=""),0,D{r}*F{r})'
    ws[f"G{r}"] = f"=H{r}*(1+Параметры!$C$6)"
    ws[f"I{r}"] = f"=H{r}*Параметры!$C$6"
    for col in ["B", "C", "D", "E", "F"]:
        ws[f"{col}{r}"].fill = fill_input
        ws[f"{col}{r}"].border = thin
    for col in ["G", "H", "I"]:
        calc_cell(ws[f"{col}{r}"], fmt=money)
    ws[f"F{r}"].number_format = money

ws["F19"] = "Прокат — себестоимость:"
ws["F19"].font = font_bold
calc_cell(ws["G19"], "=SUM(G6:G18)", money, bold=True)
calc_cell(ws["H19"], "=SUM(H6:H18)", money, bold=True)
calc_cell(ws["I19"], "=SUM(I6:I18)", money, bold=True)

header_cell(ws["B21"], "ПЛАЗМА / ЛАЗЕР", fill_section, font_section)
ws.merge_cells("B21:I21")
for col, title in [
    ("B", "Позиция"),
    ("C", "Сечение"),
    ("D", "Кол-во"),
    ("E", "Ед."),
    ("F", "Цена поставщика"),
    ("G", "Сумма"),
    ("H", "Себестоимость"),
    ("I", "Наценка"),
]:
    header_cell(ws[f"{col}23"], title)
for r in range(24, 28):
    ws[f"D{r}"] = 0
    ws[f"E{r}"] = "шт"
    ws[f"F{r}"] = 0
    ws[f"H{r}"] = f'=IF(OR(D{r}="",F{r}=""),0,D{r}*F{r})'
    ws[f"G{r}"] = f"=H{r}"
    ws[f"I{r}"] = 0
    for col in ["B", "C", "D", "E", "F"]:
        ws[f"{col}{r}"].fill = fill_input
        ws[f"{col}{r}"].border = thin
    for col in ["G", "H", "I"]:
        calc_cell(ws[f"{col}{r}"], fmt=money)
    ws[f"F{r}"].number_format = money
ws["F28"] = "Плазма — себестоимость:"
calc_cell(ws["H28"], "=SUM(H24:H27)", money, bold=True)

header_cell(ws["B30"], "КОВАНЫЕ ЭЛЕМЕНТЫ / СПЕЦ. ПРОКАТ", fill_section, font_section)
ws.merge_cells("B30:I30")
for col, title in [
    ("B", "Элемент"),
    ("C", "Сечение"),
    ("D", "Кол-во"),
    ("E", "Ед."),
    ("F", "Цена поставщика"),
    ("G", "С наценкой"),
    ("H", "Себестоимость"),
    ("I", "Наценка менеджера"),
]:
    header_cell(ws[f"{col}32"], title)
forge = [
    ("11.032", "12х6", 0, "шт", 50),
    ("41.401", "12", 0, "шт", 230),
    ("", "", 0, "шт", 2500),
    ("", "", 0, "шт", 200),
    ("", "", 0, "шт", 0),
    ("", "", 0, "шт", 0),
]
for i, (el, sect, qty, unit, price) in enumerate(forge):
    r = 33 + i
    ws[f"B{r}"] = el
    ws[f"C{r}"] = sect
    ws[f"D{r}"] = qty
    ws[f"E{r}"] = unit
    ws[f"F{r}"] = price
    ws[f"H{r}"] = f'=IF(OR(D{r}="",F{r}=""),0,D{r}*F{r})'
    ws[f"G{r}"] = f"=H{r}*(1+Параметры!$C$6)"
    ws[f"I{r}"] = f"=H{r}*Параметры!$C$6"
    for col in ["B", "C", "D", "E", "F"]:
        ws[f"{col}{r}"].fill = fill_input
        ws[f"{col}{r}"].border = thin
    for col in ["G", "H", "I"]:
        calc_cell(ws[f"{col}{r}"], fmt=money)
    ws[f"F{r}"].number_format = money

ws["F39"] = "Ковка — себестоимость:"
calc_cell(ws["G39"], "=SUM(G33:G38)", money, bold=True)
calc_cell(ws["H39"], "=SUM(H33:H38)", money, bold=True)
calc_cell(ws["I39"], "=SUM(I33:I38)", money, bold=True)

header_cell(ws["B41"], "ТОКАРНЫЕ РАБОТЫ", fill_section, font_section)
ws.merge_cells("B41:I41")
for col, title in [
    ("B", "Позиция"),
    ("C", "Сечение"),
    ("D", "Кол-во"),
    ("E", "Ед."),
    ("F", "Цена"),
    ("G", "Сумма"),
    ("H", "Себестоимость"),
    ("I", "Наценка"),
]:
    header_cell(ws[f"{col}43"], title)
for r in range(44, 48):
    ws[f"D{r}"] = 0
    ws[f"E{r}"] = "шт"
    ws[f"F{r}"] = 0
    ws[f"H{r}"] = f'=IF(OR(D{r}="",F{r}=""),0,D{r}*F{r})'
    ws[f"G{r}"] = f"=H{r}"
    ws[f"I{r}"] = 0
    for col in ["B", "C", "D", "E", "F"]:
        ws[f"{col}{r}"].fill = fill_input
        ws[f"{col}{r}"].border = thin
    for col in ["G", "H", "I"]:
        calc_cell(ws[f"{col}{r}"], fmt=money)
    ws[f"F{r}"].number_format = money
ws["F48"] = "Токарка — себестоимость:"
calc_cell(ws["H48"], "=SUM(H44:H47)", money, bold=True)

ws["F50"] = "ВСЕГО металл — себестоимость:"
ws["F50"].font = font_bold
calc_cell(ws["H50"], "=H19+H28+H39+H48", money, bold=True)
ws["H50"].fill = fill_good
ws["H50"].font = font_big

ws["F51"] = "ВСЕГО наценка менеджера с металла:"
ws["F51"].font = font_bold
calc_cell(ws["H51"], "=I19+I39", money, bold=True)
ws["H51"].fill = fill_good

ws["B53"] = (
    "Следующий этап: вынести цены в отдельный справочник поставщика; "
    "в заказе оставить кол-во. Связь с Калькуляцией уже через H19 / H28 / H39 / H48."
)
ws["B53"].font = font_note
ws.merge_cells("B53:I54")
set_col_widths(ws, {"A": 3, "B": 22, "C": 14, "D": 10, "E": 8, "F": 18, "G": 18, "H": 16, "I": 18})
ws.freeze_panes = "B6"

# =============================================================================
# Касса
# =============================================================================
ws = wb.create_sheet("Касса заказа", 3)
ws["B2"] = "КАССА ЗАКАЗА (не путать с прибылью)"
ws["B2"].font = font_title
ws.merge_cells("B2:D2")
ws["B3"] = "Только деньги: сколько получили и сколько выплатили. Прибыль — на листе Сводка."
ws["B3"].font = font_note

header_cell(ws["B5"], "ПРИХОДЫ ОТ КЛИЕНТА", fill_section, font_section)
ws.merge_cells("B5:D5")
for col, title in [("B", "Дата"), ("C", "Основание"), ("D", "Сумма")]:
    header_cell(ws[f"{col}6"], title)
for r in range(7, 12):
    for col in ["B", "C", "D"]:
        input_cell(ws[f"{col}{r}"])
    ws[f"D{r}"].number_format = money
ws["C12"] = "Итого получено:"
calc_cell(ws["D12"], "=SUM(D7:D11)", money, bold=True)
ws["C13"] = "Цена КП (из Сводки):"
calc_cell(ws["D13"], "=Сводка!D24", money)
ws["C14"] = "Остаток к получению:"
calc_cell(ws["D14"], "=D13-D12", money, bold=True)
ws["D14"].fill = fill_warn

header_cell(ws["B16"], "РАСХОДЫ ПО ЗАКАЗУ (факт выплат)", fill_section, font_section)
ws.merge_cells("B16:D16")
for col, title in [("B", "Дата"), ("C", "Основание"), ("D", "Сумма")]:
    header_cell(ws[f"{col}17"], title)
for r in range(18, 43):
    for col in ["B", "C", "D"]:
        input_cell(ws[f"{col}{r}"])
    ws[f"D{r}"].number_format = money
ws["C44"] = "Итого выплат:"
calc_cell(ws["D44"], "=SUM(D18:D42)", money, bold=True)
ws["C45"] = "Касса заказа (получено − выплачено):"
calc_cell(ws["D45"], "=D12-D44", money, bold=True)
ws["D45"].fill = fill_good
ws["D45"].font = font_big
ws["B47"] = "Это не прибыль. Прибыль = цена КП − себестоимость − НПД − агентские (Сводка)."
ws["B47"].font = font_note
set_col_widths(ws, {"A": 3, "B": 14, "C": 44, "D": 16})

# =============================================================================
# Инструкция
# =============================================================================
ws = wb.create_sheet("Инструкция", 4)
ws["B2"] = "Как работать с файлом"
ws["B2"].font = font_title
steps = [
    "1. «Параметры» — наценка менеджера 15%, целевая маржа 20%, мин. маржа 10%, НПД 4% или 6%, аренда 2000 ₽/день.",
    "2. «Металл» — кол-во и цены поставщика. Наценка менеджера считается как себестоимость × %.",
    "3. «Калькуляция» — количества по этапам. Тип «закупка» → наценка менеджера. Аренда из Параметров.",
    "4. «Сводка» — себестоимость и рекомендуемая цена. Введите цену КП в D24. Смотрите статус.",
    "5. Таблица гибкой маржи 10%…35% — чтобы торговаться и не уходить ниже минимума.",
    "6. «Касса заказа» — авансы и факт выплат. Не смешивать с маржой.",
    "",
    "Уточнение про «нормы»: шаблоны расхода (например, на 1 м забора — X кг профиля и Y часов). "
    "Если их нет — каждый заказ считаем с нуля. Позже можно сделать заготовки под Ворота/Забор/Лестница.",
    "",
    "Следующий этап по вашему запросу: справочник металла с ценами поставщика и связью с заказом.",
]
for i, s in enumerate(steps):
    ws[f"B{4+i}"] = s
    ws[f"B{4+i}"].font = font_label
set_col_widths(ws, {"A": 3, "B": 120})

order = ["Сводка", "Калькуляция", "Металл", "Параметры", "Касса заказа", "Инструкция"]
for i, name in enumerate(order):
    wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

path = "/workspace/templates/Калькуляция_проект_v1.xlsx"
wb.save(path)
print("Saved", path)
print("Sheets:", wb.sheetnames)
print("Last calc template row:", last_data_row)
