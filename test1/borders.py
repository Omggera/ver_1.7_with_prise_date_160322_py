from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

# Тонкая граница по низу в диапазоне ячеек
def set_border(ws2, cell_range):
    rows = ws2[cell_range]
    for row in rows:
        row[0].border = Border(bottom=Side(border_style='thin', color='000000'))
        row[-1].border = Border(bottom=Side(border_style='thin', color='000000'))
    for c in rows[0]:
        c.border = Border(bottom=Side(border_style='thin', color='000000'))
    for c in rows[-1]:
        c.border = Border(bottom=Side(border_style='thin', color='000000'))


# Тонкая граница справа в диапазоне ячеек
def set_border2(ws2, cell_range):
    rows = ws2[cell_range]
    for row in rows:
        row[0].border = Border(right=Side(border_style='thin', color='000000'))
        row[-1].border = Border(right=Side(border_style='thin', color='000000'))
    for c in rows[0]:
        c.border = Border(right=Side(border_style='thin', color='000000'))
    for c in rows[-1]:
        c.border = Border(right=Side(border_style='thin', color='000000'))


# Толстая граница справа в диапазоне ячеек
def set_border3(ws2, cell_range):
    rows = ws2[cell_range]
    for row in rows:
        row[0].border = Border(right=Side(border_style='medium', color='000000'))
        row[-1].border = Border(right=Side(border_style='medium', color='000000'))
    for c in rows[0]:
        c.border = Border(right=Side(border_style='medium', color='000000'))
    for c in rows[-1]:
        c.border = Border(right=Side(border_style='medium', color='000000'))


# Толстая по низу в диапазоне ячеек
def set_border4(ws2, cell_range):
    rows = ws2[cell_range]
    for row in rows:
        row[0].border = Border(bottom=Side(border_style='medium', color='000000'))
        row[-1].border = Border(bottom=Side(border_style='medium', color='000000'))
    for c in rows[0]:
        c.border = Border(bottom=Side(border_style='medium', color='000000'))
    for c in rows[-1]:
        c.border = Border(bottom=Side(border_style='medium', color='000000'))


# Тонкая пунктирная по низу в диапазоне ячеек
def set_border5(ws2, cell_range):
    rows = ws2[cell_range]
    for row in rows:
        row[0].border = Border(bottom=Side(border_style='hair', color='000000'))
        row[-1].border = Border(bottom=Side(border_style='hair', color='000000'))
    for c in rows[0]:
        c.border = Border(bottom=Side(border_style='hair', color='000000'))
    for c in rows[-1]:
        c.border = Border(bottom=Side(border_style='hair', color='000000'))