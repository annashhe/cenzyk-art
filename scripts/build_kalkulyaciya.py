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

# Каталог проката: цена за 1 шт = штанга 6 м (цена позиции поставщика)
# (группа, сечение, вес_6м_кг, цена_шт)
PROKAT = [
    ("Квадрат", "10х10", 5.1, 428.30),
    ("Квадрат", "12х12", 7.1, 596.26),
    ("Квадрат", "14х14", 9.5, 797.81),
    ("Квадрат", "16х16", 12.6, 1058.15),
    ("Труба квадрат.", "15х15х1.5", 4.0, 383.80),
    ("Труба квадрат.", "20х20х1.5", 6.0, 575.70),
    ("Труба квадрат.", "20х20х2.0", 7.0, 586.53),
    ("Труба квадрат.", "30х30х2.0", 11.0, 921.69),
    ("Труба квадрат.", "40х40х1.5", 12.0, 1151.40),
    ("Труба квадрат.", "40х40х2.0", 15.0, 1256.85),
    ("Труба прямоуг.", "40х20х1.5", 9.0, 863.55),
    ("Труба прямоуг.", "40х20х2.0", 11.0, 921.69),
    ("Труба прямоуг.", "50х25х1.5", 12.0, 1151.40),
    ("Труба прямоуг.", "50х25х2.0", 14.0, 1173.06),
    ("Полоса", "20х4 мм", 4.1, 397.04),
    ("Полоса", "25х4 мм", 5.1, 493.88),
    ("Полоса", "30х4 мм", 6.1, 557.85),
    ("Полоса", "40х4 мм", 8.0, 731.60),
    ("Полоса", "50х4 мм", 10.0, 914.50),
    ("Полоса", "100х10 мм", 49.0, 4687.83),
    ("Труба в/г", "15х2.8", 8.0, 641.52),
    ("Труба в/г", "20х2.8", 11.0, 882.09),
    ("Труба в/г", "25х3.2", 16.0, 1268.80),
    ("Труба в/г", "32х3.2", 20.0, 1586.00),
    ("Труба в/г", "40х3.5", 25.0, 1982.50),
    ("Прокат круглый", "Ø6", 1.6, 156.24),
    ("Прокат круглый", "Ø8", 2.8, 270.87),
    ("Прокат круглый", "Ø10", 4.0, 387.80),
    ("Прокат круглый", "Ø12", 6.0, 581.70),
    ("Прокат круглый", "Ø14", 15.0, 1352.55),
    ("Прокат круглый", "Ø16", 10.0, 901.70),
]

PRODUCT_TYPES = [
    "Ворота",
    "Забор",
    "Лестница",
    "Мебель",
    "Интерьер",
    "Экстерьер",
    "Сувенир",
    "Прочее",
]


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

# Бренд и контакты для листа «Предложение» (как в PDF-референсе)
ws["B18"] = "БРЕНД И КОНТАКТЫ ДЛЯ КП"
ws["B18"].font = font_bold
ws.merge_cells("B18:D18")
brand_params = [
    (19, "Имя в шапке КП", "Кузнец Илай Саматов"),
    (20, "Подзаголовок", "мастер художественной ковки"),
    (21, "ФИО полностью", "Самозанятый Саматов Илай Валерьевич"),
    (22, "Телефон", "+7 (952) 058-92-78"),
    (23, "Сайт", "Cenzyk.art"),
    (24, "Email", "ilja.samatov@yandex.ru"),
    (25, "Адрес", "г. Калининград, ул. Ручейная, 2а"),
    (26, "График Пн–Пт", "9:00-17:00"),
    (27, "График Сб", "10:00-14:00"),
    (28, "График Вс", "выходной"),
    (29, "Стаж", "11 лет"),
    (30, "Профиль (текст)", "Мебель, лестничные и балконные ограждения, элементы декора."),
    (31, "Срок действия КП, дней", 10),
]
for r, name, val in brand_params:
    ws[f"B{r}"] = name
    ws[f"B{r}"].fill = fill_soft
    ws[f"B{r}"].border = thin
    input_cell(ws[f"C{r}"], val)
    ws[f"C{r}"].font = font_bold
    ws.merge_cells(f"C{r}:D{r}")

ws["B33"] = "Тип изделия (список для шапки)"
ws["B33"].font = font_bold
for i, t in enumerate(PRODUCT_TYPES, start=34):
    ws[f"B{i}"] = t
    ws[f"B{i}"].fill = fill_soft
    ws[f"B{i}"].border = thin

ws["B43"] = "Правило маржи"
ws["B43"].font = font_bold
ws["B44"] = (
    "Себестоимость C = реальные выплаты (материалы по цене поставщика, труд, транспорт, аренда, расходники, субподряд).\n"
    "Наценка менеджера = % только на тип «закупка» — это часть вашей прибыли, не расход.\n"
    "Цена для целевой маржи P = C / (1 − НПД% − целевая маржа% − агентские%).\n"
    "Прибыль = P − C − P×НПД − P×агентские.  Маржа % = Прибыль / P.\n"
    "В листе «Предложение» варианты СЗ 4%/6% считаются как Итого×1.04 / Итого×1.06 — как в вашем PDF.\n"
    "КАССА — опциональный учёт денег после сделки, на цену КП не влияет."
)
ws["B44"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("B44:F50")
ws["B44"].fill = fill_soft
ws.row_dimensions[44].height = 110
set_col_widths(ws, {"A": 3, "B": 48, "C": 42, "D": 28, "E": 12, "F": 12})

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
    (10, "Клиент"),
    (11, "№ клиента"),
]:
    ws[f"B{r}"] = lab
    ws[f"B{r}"].fill = fill_soft
    ws[f"B{r}"].border = thin
    input_cell(ws[f"C{r}"])
ws["C8"] = "Ворота"
ws["C9"] = "=TODAY()"
ws["C9"].number_format = "DD.MM.YYYY"
# Тип: список из Параметры B19:B26
dv = DataValidation(type="list", formula1="Параметры!$B$34:$B$41", allow_blank=True)
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
    (25, "Рекомендуемая цена при целевой марже", '=IF((1-F9-F7-F11)<=0,"Ошибка %",D14/(1-F9-F7-F11))', fill_good),
    (26, "Минимальная цена (при мин. марже)", '=IF((1-F9-F8-F11)<=0,"Ошибка %",D14/(1-F9-F8-F11))', fill_warn),
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

ws["B42"] = "Строка 20% — базовый ориентир. Целевую маржу меняйте на листе Параметры."
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
    "Рабочий контур: Металл → Калькуляция → Сводка (цена КП). "
    "Лист «Касса» — только ПОСЛЕ сделки, для авансов/выплат. На расчёт цены НЕ влияет — можно не трогать."
)
ws["B49"].font = font_note
ws.merge_cells("B49:G50")
set_col_widths(ws, {"A": 3, "B": 44, "C": 22, "D": 16, "E": 22, "F": 14, "G": 12})
ws.freeze_panes = "B5"

# =============================================================================
# Металл (сначала — чтобы знать адреса итогов для Калькуляции)
# =============================================================================
ws = wb.create_sheet("Металл", 1)
ws["B2"] = "МЕТАЛЛ И КОМПЛЕКТУЮЩИЕ — каталог"
ws["B2"].font = font_title
ws.merge_cells("B2:I2")
ws["B3"] = (
    "Цена поставщика = за 1 шт (штанга 6 м). Кол-во — сколько штанг в заказ. "
    "Вес 6м — справочно. Наценка менеджера: себестоимость × % из Параметров."
)
ws["B3"].font = font_note
ws.merge_cells("B3:I3")

header_cell(ws["B4"], "ПРОКАТ (каталог поставщика)", fill_section, font_section)
ws.merge_cells("B4:I4")
for col, title in [
    ("B", "Прокат"),
    ("C", "Сечение"),
    ("D", "Кол-во"),
    ("E", "Ед."),
    ("F", "Цена поставщика, ₽/шт (6м)"),
    ("G", "С наценкой менеджера"),
    ("H", "Себестоимость"),
    ("I", "Наценка менеджера"),
    ("J", "Вес 6м, кг"),
]:
    header_cell(ws[f"{col}5"], title)

prokat_first = 6
for i, (group, sect, weight, price) in enumerate(PROKAT):
    r = prokat_first + i
    ws[f"B{r}"] = group
    ws[f"C{r}"] = sect
    ws[f"D{r}"] = 0
    ws[f"E{r}"] = "шт"
    ws[f"F{r}"] = price
    ws[f"J{r}"] = weight
    ws[f"H{r}"] = f'=IF(OR(D{r}="",F{r}=""),0,D{r}*F{r})'
    ws[f"G{r}"] = f"=H{r}*(1+Параметры!$C$6)"
    ws[f"I{r}"] = f"=H{r}*Параметры!$C$6"
    for col in ["B", "C", "D", "E", "F", "J"]:
        ws[f"{col}{r}"].fill = fill_input
        ws[f"{col}{r}"].border = thin
    # цена и вес — каталог: цена можно менять при обновлении прайса
    ws[f"J{r}"].fill = fill_soft
    for col in ["G", "H", "I"]:
        calc_cell(ws[f"{col}{r}"], fmt=money)
    ws[f"F{r}"].number_format = money
    ws[f"J{r}"].number_format = "0.0"

prokat_last = prokat_first + len(PROKAT) - 1  # 36
prokat_total_row = prokat_last + 1  # 37

ws[f"F{prokat_total_row}"] = "Прокат — себестоимость:"
ws[f"F{prokat_total_row}"].font = font_bold
calc_cell(ws[f"G{prokat_total_row}"], f"=SUM(G{prokat_first}:G{prokat_last})", money, bold=True)
calc_cell(ws[f"H{prokat_total_row}"], f"=SUM(H{prokat_first}:H{prokat_last})", money, bold=True)
calc_cell(ws[f"I{prokat_total_row}"], f"=SUM(I{prokat_first}:I{prokat_last})", money, bold=True)
ws[f"B{prokat_total_row}"] = "Итого прокат"
ws[f"B{prokat_total_row}"].font = font_bold

def write_markup_block(ws, start_title_row, title, col_b_header, col_c_header, n_rows=10, seeds=None):
    """Блок с наценкой менеджера как у проката: H=себест, G=с наценкой, I=наценка."""
    header_cell(ws[f"B{start_title_row}"], title, fill_section, font_section)
    ws.merge_cells(f"B{start_title_row}:I{start_title_row}")
    hdr = start_title_row + 2
    for col, text in [
        ("B", col_b_header),
        ("C", col_c_header),
        ("D", "Кол-во"),
        ("E", "Ед."),
        ("F", "Цена поставщика"),
        ("G", "С наценкой менеджера"),
        ("H", "Себестоимость"),
        ("I", "Наценка менеджера"),
    ]:
        header_cell(ws[f"{col}{hdr}"], text)
    first = hdr + 1
    last = first + n_rows - 1
    seeds = seeds or []
    for idx, r in enumerate(range(first, last + 1)):
        if idx < len(seeds):
            b, c, price = seeds[idx]
            ws[f"B{r}"] = b
            ws[f"C{r}"] = c
            ws[f"F{r}"] = price
        else:
            ws[f"F{r}"] = 0
        ws[f"D{r}"] = 0
        ws[f"E{r}"] = "шт"
        ws[f"H{r}"] = f'=IF(OR(D{r}="",F{r}=""),0,D{r}*F{r})'
        ws[f"G{r}"] = f"=H{r}*(1+Параметры!$C$6)"
        ws[f"I{r}"] = f"=H{r}*Параметры!$C$6"
        for col in ["B", "C", "D", "E", "F"]:
            ws[f"{col}{r}"].fill = fill_input
            ws[f"{col}{r}"].border = thin
        for col in ["G", "H", "I"]:
            calc_cell(ws[f"{col}{r}"], fmt=money)
        ws[f"F{r}"].number_format = money
    total_row = last + 1
    short = title.split("(")[0].strip()
    ws[f"F{total_row}"] = f"{short} — себестоимость:"
    calc_cell(ws[f"G{total_row}"], f"=SUM(G{first}:G{last})", money, bold=True)
    calc_cell(ws[f"H{total_row}"], f"=SUM(H{first}:H{last})", money, bold=True)
    calc_cell(ws[f"I{total_row}"], f"=SUM(I{first}:I{last})", money, bold=True)
    return first, last, total_row


# --- Плазма / лазер: 10 позиций + наценка как у проката ---
plasma_first, plasma_last, plasma_total_row = write_markup_block(
    ws,
    prokat_total_row + 2,
    "ПЛАЗМА / ЛАЗЕР (до 10 позиций)",
    "Позиция",
    "Сечение / описание",
)

# --- Ковка: 10 позиций ---
forge_first, forge_last, forge_total_row = write_markup_block(
    ws,
    plasma_total_row + 2,
    "КОВАНЫЕ ЭЛЕМЕНТЫ / СПЕЦ. ПРОКАТ (до 10 позиций)",
    "Элемент / артикул",
    "Сечение",
    seeds=[("11.032", "12х6", 50), ("41.401", "12", 230)],
)

# --- Токарка: 10 позиций + наценка как у проката ---
turn_first, turn_last, turn_total_row = write_markup_block(
    ws,
    forge_total_row + 2,
    "ТОКАРНЫЕ РАБОТЫ (до 10 позиций)",
    "Позиция",
    "Сечение / описание",
)

# --- Доп. комплектующие: 10 пустых позиций ---
extra_first, extra_last, extra_total_row = write_markup_block(
    ws,
    turn_total_row + 2,
    "ДОПОЛНИТЕЛЬНЫЕ КОМПЛЕКТУЮЩИЕ (до 10 позиций)",
    "Комплектующая",
    "Описание / артикул",
)

grand_row = extra_total_row + 2
ws[f"F{grand_row}"] = "ВСЕГО металл — себестоимость:"
ws[f"F{grand_row}"].font = font_bold
calc_cell(
    ws[f"H{grand_row}"],
    (
        f"=H{prokat_total_row}+H{plasma_total_row}+H{forge_total_row}"
        f"+H{turn_total_row}+H{extra_total_row}"
    ),
    money,
    bold=True,
)
ws[f"H{grand_row}"].fill = fill_good
ws[f"H{grand_row}"].font = font_big

ws[f"F{grand_row + 1}"] = "ВСЕГО наценка менеджера с листа Металл:"
ws[f"F{grand_row + 1}"].font = font_bold
calc_cell(
    ws[f"H{grand_row + 1}"],
    (
        f"=I{prokat_total_row}+I{plasma_total_row}+I{forge_total_row}"
        f"+I{turn_total_row}+I{extra_total_row}"
    ),
    money,
    bold=True,
)
ws[f"H{grand_row + 1}"].fill = fill_good

ws[f"B{grand_row + 3}"] = (
    "Обновление прайса: меняйте жёлтые цены в колонке F. Кол-во D — под заказ. "
    "Прокат, плазма, ковка, токарка и комплектующие — с наценкой менеджера. "
    "Итоги уходят в Калькуляцию автоматически."
)
ws[f"B{grand_row + 3}"].font = font_note
ws.merge_cells(f"B{grand_row + 3}:J{grand_row + 4}")

set_col_widths(
    ws,
    {"A": 3, "B": 18, "C": 16, "D": 10, "E": 8, "F": 26, "G": 18, "H": 16, "I": 18, "J": 12},
)
ws.freeze_panes = "B6"
ws.auto_filter.ref = f"B5:J{prokat_last}"

# Сохраняем адреса итогов для связи
METAL_LINKS = {
    "prokat": f"Металл!H{prokat_total_row}",
    "plasma": f"Металл!H{plasma_total_row}",
    "forge": f"Металл!H{forge_total_row}",
    "turn": f"Металл!H{turn_total_row}",
    "extra": f"Металл!H{extra_total_row}",
}
print("Metal totals:", METAL_LINKS, "prokat rows", prokat_first, prokat_last)

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
add("Металл", "Лазер / плазма (лист Металл)", "закупка", None, 1, "сводка", "", "metal_plasma")
add("Металл", "Ков. элементы / спец. прокат", "закупка", None, 1, "сводка", "", "metal_forge")
add("Металл", "Токарные работы (лист Металл)", "закупка", None, 1, "сводка", "", "metal_turn")
add("Металл", "Доп. комплектующие (лист Металл)", "закупка", None, 1, "сводка", "", "metal_extra")

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
        ws[f"E{r}"] = f"={METAL_LINKS['prokat']}"
        ws[f"F{r}"] = 1
        ws[f"L{r}"] = "Себестоимость проката (каталог, цена за шт 6м)"
    elif special == "metal_plasma":
        ws[f"E{r}"] = f"={METAL_LINKS['plasma']}"
        ws[f"F{r}"] = 1
    elif special == "metal_forge":
        ws[f"E{r}"] = f"={METAL_LINKS['forge']}"
        ws[f"F{r}"] = 1
    elif special == "metal_turn":
        ws[f"E{r}"] = f"={METAL_LINKS['turn']}"
        ws[f"F{r}"] = 1
    elif special == "metal_extra":
        ws[f"E{r}"] = f"={METAL_LINKS['extra']}"
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
    if special in ("rent", "metal_cost", "metal_plasma", "metal_forge", "metal_turn", "metal_extra"):
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
    "Пустые строки — для своих позиций. Тип «закупка» включает наценку менеджера автоматически."
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
        "L": 40,
        "M": 3,
        "N": 28,
        "O": 14,
    },
)
ws.freeze_panes = "B16"
ws.auto_filter.ref = f"B15:L{last_data_row + 20}"

# =============================================================================
# Предложение для заказчика — по PDF-референсу Cenzyk.art
# =============================================================================
ws = wb.create_sheet("Предложение", 3)

fill_brand = PatternFill("solid", fgColor="6B9080")
fill_red_tot = PatternFill("solid", fgColor="C94C4C")
fill_green_tot = PatternFill("solid", fgColor="B7D7A8")
fill_yellow_tot = PatternFill("solid", fgColor="F6E59A")
fill_tax = PatternFill("solid", fgColor="D9E8F5")
font_white_bold = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
font_brand = Font(name="Georgia", bold=True, size=18, color="1A1A1A")
font_brand_sub = Font(name="Georgia", size=11, color="1A1A1A")

# Служебная строка (вне печати)
ws["B1"] = (
    "ДЛЯ ВАС: заполните смету и вставьте фото → Файл → Экспорт в PDF (этот лист). "
    "Ориентир цены со Сводки справа. Шаблон как ваш PDF Андрей_Светлана_Краснознаменск."
)
ws["B1"].font = font_note
ws.merge_cells("B1:F1")
ws["G1"] = "Ориентир из Сводки (цена КП):"
ws["G1"].font = font_note
calc_cell(ws["H1"], "=Сводка!D24", money, bold=True)

# ---- ШАПКА БРЕНДА (как в PDF) ----
for col in ["B", "C", "D", "E", "F"]:
    ws[f"{col}3"].fill = fill_brand
    ws[f"{col}4"].fill = fill_brand
    ws[f"{col}5"].fill = fill_brand
ws.merge_cells("B3:C5")
ws["B3"] = '=Параметры!C19&CHAR(10)&Параметры!C20'
ws["B3"].font = font_brand
ws["B3"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws["B3"].fill = fill_brand

ws["D3"] = "Предложение №:"
ws["D3"].fill = fill_brand
ws["D3"].font = font_brand_sub
ws["E3"] = '=IF(Сводка!C11="","—",Сводка!C11)'
ws["E3"].fill = fill_input
ws["E3"].font = font_bold
ws["E3"].border = thin
ws.merge_cells("E3:F3")

ws["D4"] = "Дата:"
ws["D4"].fill = fill_brand
ws["D4"].font = font_brand_sub
ws["E4"] = "=Сводка!C9"
ws["E4"].number_format = '[$-419]D MMMM YYYY;@'
ws["E4"].fill = fill_calc
ws["E4"].border = thin
ws.merge_cells("E4:F4")

ws["D5"] = "Актуально до:"
ws["D5"].fill = fill_brand
ws["D5"].font = font_brand_sub
ws["E5"] = "=E4+Параметры!C31"
ws["E5"].number_format = "DD.MM.YYYY"
ws["E5"].fill = fill_calc
ws["E5"].border = thin
ws.merge_cells("E5:F5")
ws.row_dimensions[3].height = 22
ws.row_dimensions[4].height = 20
ws.row_dimensions[5].height = 20

# Клиент / объект
ws["B7"] = "Клиент:"
ws["C7"] = "=Сводка!C10"
ws["D7"] = "Объект:"
ws["E7"] = "=Сводка!C6"
ws["B8"] = "Изделие:"
ws["C8"] = "=Сводка!C7"
ws["D8"] = "Тип:"
ws["E8"] = "=Сводка!C8"
for r in (7, 8):
    for col in ["B", "D"]:
        ws[f"{col}{r}"].font = font_bold
        ws[f"{col}{r}"].fill = fill_soft
        ws[f"{col}{r}"].border = thin
    for col in ["C", "E"]:
        ws[f"{col}{r}"].fill = fill_calc
        ws[f"{col}{r}"].border = thin
ws.merge_cells("E7:F7")
ws.merge_cells("E8:F8")

# ---- ВИЗУАЛИЗАЦИЯ 1 ----
ws["B10"] = 'Визуализация изделия «'
ws["C10"] = "=Сводка!C7"
ws["D10"] = "»"
ws["B10"].font = Font(name="Calibri", bold=True, size=14, color="2F3E46")
ws["C10"].font = Font(name="Calibri", bold=True, size=14, color="2F3E46")
ws["C10"].fill = fill_input
ws["C10"].border = thin
ws.merge_cells("C10:E10")

ws["B11"] = "ФОТО / ВИЗУАЛИЗАЦИЯ 1\nВставка → Изображение  или  Ctrl+V"
ws["D11"] = "ФОТО / ВИЗУАЛИЗАЦИЯ 2\nВставка → Изображение  или  Ctrl+V"
for addr in ("B11", "D11"):
    ws[addr].fill = PatternFill("solid", fgColor="E8E8E8")
    ws[addr].font = font_note
    ws[addr].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[addr].border = thin
ws.merge_cells("B11:C15")
ws.merge_cells("D11:F15")
for rr in range(11, 16):
    ws.row_dimensions[rr].height = 28

# ---- ВИЗУАЛИЗАЦИЯ 2 (опционально) ----
ws["B17"] = "Визуализация изделия «"
ws["C17"] = ""  # второе изделие — вручную
ws["D17"] = "»  (если одно изделие — оставьте пустым)"
ws["B17"].font = Font(name="Calibri", bold=True, size=14, color="2F3E46")
ws["C17"].fill = fill_input
ws["C17"].border = thin
ws["C17"].font = Font(name="Calibri", bold=True, size=14, color="2F3E46")
ws.merge_cells("C17:E17")
ws["D17"].font = font_note

ws["B18"] = "ФОТО / ВИЗУАЛИЗАЦИЯ 3\nВставка → Изображение"
ws["D18"] = "ФОТО / ВИЗУАЛИЗАЦИЯ 4\nВставка → Изображение"
for addr in ("B18", "D18"):
    ws[addr].fill = PatternFill("solid", fgColor="E8E8E8")
    ws[addr].font = font_note
    ws[addr].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws[addr].border = thin
ws.merge_cells("B18:C22")
ws.merge_cells("D18:F22")
for rr in range(18, 23):
    ws.row_dimensions[rr].height = 28

# ---- СМЕТА (как в PDF) ----
ws["B24"] = "Смета работ и материалов"
ws["B24"].font = Font(name="Calibri", bold=True, size=16, color="2F3E46")
ws.merge_cells("B24:F24")

for col, title in [
    ("B", "Материал / расход / часть"),
    ("C", "ед. изм."),
    ("D", "количество"),
    ("E", "стоимость за шт"),
    ("F", "общее"),
]:
    header_cell(ws[f"{col}25"], title, fill=fill_header, font=font_header)

# 22 строки позиций сметы (жёлтый ввод) — как детальная смета в референсе
smeta_first = 26
smeta_last = 47
for r in range(smeta_first, smeta_last + 1):
    for col in ["B", "C", "D", "E"]:
        input_cell(ws[f"{col}{r}"])
    ws[f"D{r}"] = 0
    ws[f"E{r}"] = 0
    ws[f"E{r}"].number_format = money
    calc_cell(ws[f"F{r}"], f"=IF(OR(D{r}=\"\",E{r}=\"\"),0,D{r}*E{r})", money)
    ws[f"C{r}"] = "шт"

# Красный итог расходов на изготовление
r_mat = smeta_last + 1  # 48
ws[f"B{r_mat}"] = "Итого расходы на изготовление (материал, расходники, покраска, доставка)"
ws[f"B{r_mat}"].font = font_white_bold
ws[f"B{r_mat}"].fill = fill_red_tot
ws.merge_cells(f"B{r_mat}:E{r_mat}")
for col in ["B", "C", "D", "E", "F"]:
    ws[f"{col}{r_mat}"].fill = fill_red_tot
    ws[f"{col}{r_mat}"].border = thin
    ws[f"{col}{r_mat}"].font = font_white_bold
calc_cell(ws[f"F{r_mat}"], f"=SUM(F{smeta_first}:F{smeta_last})", money, bold=True)
ws[f"F{r_mat}"].fill = fill_red_tot
ws[f"F{r_mat}"].font = font_white_bold

# Работы: изготовление + монтаж (пакетные цены для клиента)
r_work1 = r_mat + 1  # 49
r_work2 = r_mat + 2  # 50
ws[f"B{r_work1}"] = "Изготовление конструкции"
ws[f"C{r_work1}"] = "ед."
ws[f"D{r_work1}"] = 1
ws[f"E{r_work1}"] = 0
ws[f"B{r_work2}"] = "Монтаж"
ws[f"C{r_work2}"] = "ед."
ws[f"D{r_work2}"] = 1
ws[f"E{r_work2}"] = 0
for r in (r_work1, r_work2):
    for col in ["B", "C", "D", "E"]:
        input_cell(ws[f"{col}{r}"])
    ws[f"E{r}"].number_format = money
    calc_cell(ws[f"F{r}"], f"=D{r}*E{r}", money)

r_work_tot = r_mat + 3  # 51
ws[f"B{r_work_tot}"] = "Итого за работу"
ws.merge_cells(f"B{r_work_tot}:E{r_work_tot}")
for col in ["B", "C", "D", "E", "F"]:
    ws[f"{col}{r_work_tot}"].fill = fill_green_tot
    ws[f"{col}{r_work_tot}"].border = thin
    ws[f"{col}{r_work_tot}"].font = font_bold
calc_cell(ws[f"F{r_work_tot}"], f"=F{r_work1}+F{r_work2}", money, bold=True)
ws[f"F{r_work_tot}"].fill = fill_green_tot

r_grand = r_mat + 4  # 52
ws[f"B{r_grand}"] = "Итого общая"
ws.merge_cells(f"B{r_grand}:E{r_grand}")
for col in ["B", "C", "D", "E", "F"]:
    ws[f"{col}{r_grand}"].fill = fill_yellow_tot
    ws[f"{col}{r_grand}"].border = thin
    ws[f"{col}{r_grand}"].font = font_big
calc_cell(ws[f"F{r_grand}"], f"=F{r_mat}+F{r_work_tot}", money, bold=True)
ws[f"F{r_grand}"].fill = fill_yellow_tot
ws[f"F{r_grand}"].font = font_big

# Варианты СЗ — как в PDF: Итого × 1.04 / 1.06
r_sz4 = r_mat + 5  # 53
r_sz6 = r_mat + 6  # 54
ws[f"B{r_sz4}"] = "Проведение через СЗ 4% физ. лицо"
ws[f"B{r_sz6}"] = "Проведение через СЗ 6% Юр. лицо"
for r, mult in ((r_sz4, "1.04"), (r_sz6, "1.06")):
    ws.merge_cells(f"B{r}:E{r}")
    for col in ["B", "C", "D", "E", "F"]:
        ws[f"{col}{r}"].fill = fill_tax
        ws[f"{col}{r}"].border = thin
        ws[f"{col}{r}"].font = font_bold
    calc_cell(ws[f"F{r}"], f"=F{r_grand}*{mult}", money, bold=True)
    ws[f"F{r}"].fill = fill_tax

ws[f"B{r_sz6 + 1}"] = (
    "Цены в смете — для заказчика (уже с вашей наценкой). "
    "Строки «СЗ 4%/6%» = Итого общая × 1.04 / 1.06, как в вашем PDF. "
    "После заполнения отправьте лист в PDF вместо скриншота."
)
ws[f"B{r_sz6 + 1}"].font = font_note
ws.merge_cells(f"B{r_sz6 + 1}:F{r_sz6 + 1}")

# ---- КОНТАКТЫ (как в PDF) ----
r_cont = r_sz6 + 3  # 57
ws[f"B{r_cont}"] = "Контакты"
ws[f"B{r_cont}"].font = Font(name="Calibri", bold=True, size=16, color="2F3E46")
ws.merge_cells(f"B{r_cont}:C{r_cont}")
ws[f"E{r_cont}"] = "С уважением,"
ws[f"E{r_cont}"].font = font_label

contacts = [
    (1, '=Параметры!C21'),
    (2, '="Телефон: "&Параметры!C22'),
    (3, '="Сайт: "&Параметры!C23'),
    (4, '="Email: "&Параметры!C24'),
    (5, '="Адрес: "&Параметры!C25'),
    (6, "График работы:"),
    (7, '="Пн-Пт: "&Параметры!C26'),
    (8, '="Сб: "&Параметры!C27'),
    (9, '="Вс: "&Параметры!C28'),
    (10, "Профиль:"),
    (11, '="Стаж работы: "&Параметры!C29'),
    (12, "=Параметры!C30"),
]
for off, formula in contacts:
    r = r_cont + off
    ws[f"B{r}"] = formula
    ws[f"B{r}"].font = font_label
    ws.merge_cells(f"B{r}:C{r}")

ws[f"E{r_cont + 2}"] = "=Параметры!C19"
ws[f"E{r_cont + 2}"].font = font_bold
ws[f"E{r_cont + 3}"] = "=Параметры!C20"
ws[f"E{r_cont + 3}"].font = font_label
ws[f"E{r_cont + 5}"] = "Подпись"
ws[f"E{r_cont + 5}"].font = font_note
ws[f"E{r_cont + 5}"].fill = fill_soft
ws[f"E{r_cont + 5}"].border = thin
ws.merge_cells(f"E{r_cont + 5}:F{r_cont + 8}")
ws[f"E{r_cont + 5}"].alignment = Alignment(horizontal="center", vertical="bottom")

r_valid = r_cont + 14
ws[f"B{r_valid}"] = '="Действительно до: "&TEXT(E5,"DD.MM.YYYY")'
ws[f"B{r_valid}"].font = font_bold
ws.merge_cells(f"B{r_valid}:C{r_valid}")

# Подсветка шапки контактов
for col in ["B", "C", "D", "E", "F"]:
    ws[f"{col}{r_cont}"].fill = fill_soft

set_col_widths(ws, {"A": 3, "B": 54, "C": 12, "D": 12, "E": 16, "F": 16, "G": 18, "H": 14})
ws.print_area = f"B3:F{r_valid}"
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.page_margins.left = 0.4
ws.page_margins.right = 0.4
ws.page_margins.top = 0.4
ws.page_margins.bottom = 0.4
ws.sheet_view.showGridLines = False

print(f"Proposal smeta rows {smeta_first}-{smeta_last}, totals mat={r_mat} work={r_work_tot} grand={r_grand}")

# =============================================================================
# Касса — упрощённая, явно опциональная
# =============================================================================
ws = wb.create_sheet("Касса", 4)
ws["B2"] = "КАССА ЗАКАЗА — опционально (после сделки)"
ws["B2"].font = font_title
ws.merge_cells("B2:E2")

ws["B3"] = (
    "ЭТОТ ЛИСТ НЕ УЧАСТВУЕТ В РАСЧЁТЕ ЦЕНЫ И МАРЖИ.\n"
    "Нужен только чтобы помнить: сколько клиент уже заплатил и сколько вы уже потратили по заказу.\n"
    "Если считаете только КП — этот лист можно полностью игнорировать.\n"
    "\n"
    "Связь со Сводкой: одна — подтягивает «Цену КП» (сколько договорились), чтобы показать остаток к получению.\n"
    "Обратно в Сводку / Калькуляцию / Металл касса НИЧЕГО не отправляет."
)
ws["B3"].font = font_label
ws["B3"].fill = fill_warn
ws["B3"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("B3:E7")
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 18
ws.row_dimensions[5].height = 18
ws.row_dimensions[6].height = 18
ws.row_dimensions[7].height = 18

header_cell(ws["B9"], "1) СКОЛЬКО ПОЛУЧИЛИ ОТ КЛИЕНТА", fill_section, font_section)
ws.merge_cells("B9:D9")
for col, title in [("B", "Дата"), ("C", "Что за платёж"), ("D", "Сумма")]:
    header_cell(ws[f"{col}10"], title)
for r in range(11, 16):
    for col in ["B", "C", "D"]:
        input_cell(ws[f"{col}{r}"])
    ws[f"D{r}"].number_format = money
ws["C16"] = "Всего получено:"
calc_cell(ws["D16"], "=SUM(D11:D15)", money, bold=True)
ws["C17"] = "Цена КП (из Сводки, только для справки):"
calc_cell(ws["D17"], "=Сводка!D24", money)
ws["C18"] = "Ещё должен клиент:"
calc_cell(ws["D18"], "=D17-D16", money, bold=True)
ws["D18"].fill = fill_warn

header_cell(ws["B20"], "2) СКОЛЬКО УЖЕ ВЫПЛАТИЛИ ПО ЗАКАЗУ", fill_section, font_section)
ws.merge_cells("B20:D20")
for col, title in [("B", "Дата"), ("C", "Кому / за что"), ("D", "Сумма")]:
    header_cell(ws[f"{col}21"], title)
for r in range(22, 42):
    for col in ["B", "C", "D"]:
        input_cell(ws[f"{col}{r}"])
    ws[f"D{r}"].number_format = money
ws["C43"] = "Всего выплат:"
calc_cell(ws["D43"], "=SUM(D22:D41)", money, bold=True)
ws["C44"] = "Живые деньги по заказу сейчас (получено − выплаты):"
calc_cell(ws["D44"], "=D16-D43", money, bold=True)
ws["D44"].fill = fill_good
ws["D44"].font = font_big

ws["B46"] = (
    "Пример: КП 100 000, аванс 50 000, купили металл 20 000 → получено 50 000, выплат 20 000, "
    "в кассе заказа 30 000. Прибыль при этом смотрите на Сводке, не здесь."
)
ws["B46"].font = font_note
ws.merge_cells("B46:E47")
set_col_widths(ws, {"A": 3, "B": 14, "C": 48, "D": 16, "E": 14})

# =============================================================================
# Инструкция
# =============================================================================
ws = wb.create_sheet("Инструкция", 5)
ws["B2"] = "Как работать с файлом"
ws["B2"].font = font_title
steps = [
    "ГЛАВНЫЙ КОНТУР (для цены КП):",
    "1. «Параметры» — % менеджера, целевая/мин. маржа, НПД, аренда/день.",
    "2. «Металл» — кол-во штанг проката (шт = 6 м), плюс плазма/ковка/токарка/комплектующие. На все эти блоки — наценка менеджера.",
    "3. «Калькуляция» — количества по этапам (дни, рейсы, материалы).",
    "4. «Сводка» — смотрите рекомендуемую цену, впишите цену КП в жёлтую ячейку, проверьте статус.",
    "5. «Предложение» — по вашему PDF: шапка бренда, визуализации, смета строк, итоги красный/зелёный/жёлтый, СЗ 4%/6%, контакты.",
    "",
    "ВМЕСТО СКРИНШОТА:",
    "• Лист «Предложение» → заполните позиции сметы и вставьте фото → Файл → Экспорт в PDF.",
    "• Отправьте PDF клиенту (WhatsApp / почта). Бренд и контакты правятся на листе Параметры.",
    "",
    "КАССА — НЕ ДЛЯ РАСЧЁТА ЦЕНЫ:",
    "• Нужна после сделки: авансы клиента и факт выплат (металл, ЗП, транспорт).",
    "• Пока готовите КП — лист Касса можно не открывать.",
    "",
    "Типы изделий: Ворота, Забор, Лестница, Мебель, Интерьер, Экстерьер, Сувенир, Прочее.",
    "В шапке Сводки: Клиент + № клиента.",
]
for i, s in enumerate(steps):
    ws[f"B{4 + i}"] = s
    ws[f"B{4 + i}"].font = font_bold if s.endswith(":") else font_label
set_col_widths(ws, {"A": 3, "B": 120})

# Порядок листов
desired = [
    "Сводка",
    "Калькуляция",
    "Металл",
    "Предложение",
    "Параметры",
    "Касса",
    "Инструкция",
]
for i, name in enumerate(desired):
    wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

path = "/workspace/templates/Калькуляция_проект_v1.xlsx"
wb.save(path)
# ASCII copy for easier local open
wb.save("/workspace/templates/Kalkulyaciya_proekt_v1.xlsx")
print("Saved", path)
print("Sheets:", wb.sheetnames)
print(
    f"Prokat {prokat_first}-{prokat_last} total@{prokat_total_row}; "
    f"plasma {plasma_first}-{plasma_last} @{plasma_total_row}; "
    f"forge {forge_first}-{forge_last} @{forge_total_row}; "
    f"turn {turn_first}-{turn_last} @{turn_total_row}; "
    f"extra {extra_first}-{extra_last} @{extra_total_row}"
)
