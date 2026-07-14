"""Профессиональный корректировщик каталогов v1.5.0"""

import os
from copy import copy
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

try:
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill

APP_TITLE = "Профессиональный корректировщик каталогов"
APP_VER   = "1.5.0"

MATCH_COLUMNS = [
    "Категория портала ур.1", "Категория портала ур.2",
    "Категория портала ур.3", "Категория портала ур.4",
    "Тема", "Вопрос 1", "Вопрос 2", "Вопрос 3",
    "Вопрос 4", "Вопрос 5", "Сервис", "Бизнес-операция",
]
# Столбцы иерархии раздела каталога — унаследованный контекст (в какой
# раздел добавляется запись), а не содержание самой записи. При добавлении
# целиком новой строки их можно не закрашивать зелёным отдельно (см.
# is_new_row_green) — в отличие от Темы/Вопросов/Сервиса/Бизнес-операции.
CONTEXT_COLUMNS = {
    "Категория портала ур.1", "Категория портала ур.2",
    "Категория портала ур.3", "Категория портала ур.4",
}
# Необязательный столбец: встречается не во всех файлах. Если найден и в
# исходном, и в целевом файле — участвует в сопоставлении строк наравне
# с MATCH_COLUMNS, иначе тихо игнорируется.
LABEL_COL = "Метка"

DELETED_SHEET = "Удаленное"
BIZ_OP_COL    = "Бизнес-операция"
PROD_CAT_COL  = "Продуктовая категоризация уровень 1"
# Заливка целиком новых (зелёных) строк не должна уходить дальше этого
# столбца — служебные столбцы правее (Поддерживающий сервис, Изменения по
# запросу и т.п.) остаются некрашеными, хотя значения в них переносятся как обычно.
TYPE_COL      = "Тип"
# Оба варианта написания: с одной и двумя «с» — встречаются в разных версиях файлов
CHANGE_COL_VARIANTS = ("Изменения по запроссу", "Изменения по запросу")
AUTHOR_COL_VARIANTS = ("Кем внесены изменения",)

RED_RGB    = (255, 0, 0)
YELLOW_RGB = (255, 255, 0)
# Стандартный "Светло-зелёный" из палитры Excel (92D050) — именно этот
# оттенок используется в реальных файлах, а не чистый лаймовый 00FF00.
GREEN_RGB  = (0x92, 0xD0, 0x50)
RED_FILL    = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
GREEN_FILL  = PatternFill(start_color="FF92D050", end_color="FF92D050", fill_type="solid")
NO_FILL     = PatternFill(fill_type="none")


# ── цвет ─────────────────────────────────────────────────────────────────────

def cell_rgb(cell):
    try:
        f = cell.fill
        if f.patternType == "solid":
            c = f.fgColor
            if c.type == "rgb":
                a = c.rgb
                if a != "00000000":
                    return int(a[2:4], 16), int(a[4:6], 16), int(a[6:8], 16)
    except Exception:
        pass
    return None

def is_red(cell):
    return cell_rgb(cell) == RED_RGB

def is_yellow(cell):
    return cell_rgb(cell) == YELLOW_RGB

def is_green(cell):
    return cell_rgb(cell) == GREEN_RGB


# ── заголовок / ключ ─────────────────────────────────────────────────────────

def build_header(ws, row=1):
    return {str(c.value).strip(): c.column for c in ws[row] if c.value is not None}

def find_col(hdr, *variants):
    """Ищет столбец по нескольким возможным названиям; возвращает индекс или None."""
    for name in variants:
        idx = hdr.get(name)
        if idx is not None:
            return idx
    return None

def effective_match_columns(src_h, tgt_h, log=None):
    """
    Базовый MATCH_COLUMNS + необязательный LABEL_COL («Метка»), если он
    найден и в исходном, и в целевом файле. Если хотя бы в одном из
    файлов столбца нет — он тихо пропускается (некритично).
    """
    cols = list(MATCH_COLUMNS)
    if LABEL_COL in src_h and LABEL_COL in tgt_h:
        cols.append(LABEL_COL)
        if log:
            log(f"  [i] Столбец «{LABEL_COL}» найден — учитывается при сопоставлении строк")
    return cols

def row_key(cells, hdr, match_cols):
    m = {c.column: c.value for c in cells}
    return tuple(m.get(hdr.get(n)) for n in match_cols)

def find_row(ws, key, hdr, match_cols):
    for row in ws.iter_rows(min_row=2):
        if row_key(row, hdr, match_cols) == key:
            return row[0].row
    return None

def find_rows_fuzzy(ws, src_row, src_h, tgt_h, ignore_cols, match_cols):
    """
    Ищет в целевом листе строку, МАКСИМАЛЬНО ПОХОЖУЮ на src_row по
    ключевым столбцам match_cols, игнорируя столбцы из ignore_cols
    (изменённые жёлтой/зелёной заливкой — их значения в исходнике и
    цели намеренно различаются либо в цели ещё нет значения).

    Не требует 100%-ного совпадения остальных столбцов: например, если
    зелёная ячейка вставлена в середину последовательности «Вопрос N»
    и «вытолкнула» старое значение в соседний столбец (insert_green_cell
    сдвигает его при записи), у верной строки разойдётся ещё один
    столбец. Поэтому берётся строка с максимальным числом совпадений,
    а не обязательно полное совпадение.

    Сравниваются только столбцы, где у ИСХОДНОЙ строки вообще есть
    значение: пустая ячейка с обеих сторон ни о чём не говорит и не
    должна повышать похожесть случайной строки (иначе при большом
    числе пустых Вопрос-столбцов подворачивается ложное совпадение).
    Совпадение принимается, только если набрано не менее половины
    таких информативных столбцов.

    Возвращает список номеров строк: пусто — подходящих не найдено,
    один элемент — однозначное совпадение, больше одного — несколько
    строк набрали одинаково лучший результат (неоднозначность).
    """
    rev_src = {v: k for k, v in src_h.items()}
    ignore_names  = {rev_src[c] for c in ignore_cols if c in rev_src}
    compare_names = [n for n in match_cols if n not in ignore_names]

    src_vals = {c.column: c.value for c in src_row}
    informative = [n for n in compare_names if src_vals.get(src_h.get(n)) not in (None, "")]
    if not informative:
        return []

    total = len(informative)
    best_score, best_rows = 0, []

    for row in ws.iter_rows(min_row=2):
        tgt_vals = {c.column: c.value for c in row}
        score = 0
        for n in informative:
            t_col = tgt_h.get(n)
            if t_col is None:
                continue
            if src_vals.get(src_h.get(n)) == tgt_vals.get(t_col):
                score += 1
        if score > best_score:
            best_score, best_rows = score, [row[0].row]
        elif score == best_score and score > 0:
            best_rows.append(row[0].row)

    if best_score * 2 < total:
        return []
    return best_rows


# ── снимок ячейки ─────────────────────────────────────────────────────────────

def snap(cell):
    """Снимает все атрибуты ячейки до любых изменений."""
    s = {"v": cell.value, "ok": False}
    try:
        if cell.has_style:
            s.update(ok=True,
                     font=copy(cell.font),
                     border=copy(cell.border),
                     align=copy(cell.alignment),
                     fmt=cell.number_format,
                     prot=copy(cell.protection),
                     fill=copy(cell.fill))
    except Exception:
        pass
    return s

def restore(cell, s, fill_ovr=None):
    """Восстанавливает ячейку из снимка; fill_ovr переопределяет заливку."""
    cell.value = s["v"]
    if s["ok"]:
        try:
            cell.font        = s["font"]
            cell.border      = s["border"]
            cell.alignment   = s["align"]
            cell.number_format = s["fmt"]
            cell.protection  = s["prot"]
            cell.fill        = fill_ovr if fill_ovr is not None else s["fill"]
        except Exception:
            if fill_ovr is not None:
                try:
                    cell.fill = fill_ovr
                except Exception:
                    pass
    elif fill_ovr is not None:
        try:
            cell.fill = fill_ovr
        except Exception:
            pass


# ── операции со строками ──────────────────────────────────────────────────────

def move_to_deleted(tgt_ws, row_idx, del_ws, change_txt, author_txt, del_hdr):
    """
    Копирует строку из tgt_ws на лист Удаленное, заполняет столбцы
    изменений/автора, красит всю строку красным.
    """
    chg_col = find_col(del_hdr, *CHANGE_COL_VARIANTS)
    aut_col = find_col(del_hdr, *AUTHOR_COL_VARIANTS)
    dst_row = del_ws.max_row + 1

    # Итерируем только реальные ячейки строки — не создаём лишних объектов в tgt_ws
    for src in tgt_ws[row_idx]:
        col      = src.column
        dst_cell = del_ws.cell(row=dst_row, column=col)

        # Сначала копируем стиль + красим (restore выставляет cell.value = src.value)
        restore(dst_cell, snap(src), fill_ovr=copy(RED_FILL))

        # Затем перезаписываем value для спецстолбцов (после restore — иначе затрётся)
        if col == chg_col:
            dst_cell.value = change_txt
        elif col == aut_col:
            dst_cell.value = author_txt

    # Если chg/aut столбцы правее использованного диапазона tgt_ws — добавляем отдельно
    for col, txt in [(chg_col, change_txt), (aut_col, author_txt)]:
        if col and col > tgt_ws.max_column:
            c = del_ws.cell(row=dst_row, column=col)
            c.value = txt
            c.fill  = copy(RED_FILL)


def partial_delete(tgt_ws, row_idx, red_cols, boundary, change_txt, author_txt, tgt_h):
    """
    Частичное удаление:
    - до boundary: убирает красные ячейки, сдвигает остальные влево
    - от boundary: только очищает красные ячейки (без сдвига)
    - со всей строки снимается заливка, в столбцы изменений/автора
      записывается информация, указанная пользователем
    """
    row = tgt_ws[row_idx]

    # Снимок всех ячеек до границы (до модификаций)
    before = [
        (c.column, snap(c), c.column in red_cols)
        for c in row if c.column < boundary
    ]
    kept = [(col, s) for col, s, r in before if not r]

    # Очищаем все ячейки до границы
    for col, _, _ in before:
        c = tgt_ws.cell(row=row_idx, column=col)
        c.value = None
        c.fill  = NO_FILL

    # Пишем оставшиеся ячейки сдвинутыми влево, без заливки
    for i, (_, s) in enumerate(kept, start=1):
        restore(tgt_ws.cell(row=row_idx, column=i), s, fill_ovr=NO_FILL)

    # Ячейки от границы и правее — очищаем красные, у остальных снимаем заливку
    for cell in row:
        if cell.column >= boundary:
            if cell.column in red_cols:
                cell.value = None
            cell.fill = NO_FILL

    # Проставляем информацию об изменении в целевые столбцы строки
    chg_col = find_col(tgt_h, *CHANGE_COL_VARIANTS)
    aut_col = find_col(tgt_h, *AUTHOR_COL_VARIANTS)
    if chg_col:
        tgt_ws.cell(row=row_idx, column=chg_col, value=change_txt).fill = NO_FILL
    if aut_col:
        tgt_ws.cell(row=row_idx, column=aut_col, value=author_txt).fill = NO_FILL


def insert_green_cell(tgt_ws, row_idx, target_col, boundary, new_val):
    """
    Вставляет новое значение в target_col строки row_idx:
    - если ячейка там пуста — просто записывает значение;
    - если занята и target_col внутри зоны сдвига (< boundary) —
      сдвигает занятую ячейку и всё, что после неё (до boundary-1),
      на один столбец ближе к границе, освобождая target_col;
    - если сдвинуть некуда (последний столбец зоны уже занят, либо
      target_col вне зоны сдвига и при этом занят) — возвращает False,
      строка не трогается (вызывающий код логирует и пропускает).
    """
    cell = tgt_ws.cell(row=row_idx, column=target_col)
    if cell.value in (None, ""):
        cell.value = new_val
        return True

    if target_col >= boundary:
        return False  # за границей зоны сдвиг не определён

    zone_last = boundary - 1
    last_cell = tgt_ws.cell(row=row_idx, column=zone_last)
    if last_cell.value not in (None, ""):
        return False  # некуда сдвигать — последний столбец зоны уже занят

    for col in range(zone_last, target_col, -1):
        restore(tgt_ws.cell(row=row_idx, column=col),
                snap(tgt_ws.cell(row=row_idx, column=col - 1)))

    tgt_ws.cell(row=row_idx, column=target_col, value=new_val)
    return True


def adjust_row(tgt_ws, row_idx, src_row, src_h, tgt_h, yellow_cols, green_cols,
               boundary, change_txt, author_txt, log=None, context_msg=""):
    """
    Точечная правка строки: жёлтые ячейки (корректировка значения) и/или
    зелёные ячейки (добавление нового значения, при конфликте — со сдвигом).
    - жёлтые: переносит новое значение по имени столбца, без сдвига
    - зелёные: вставляет новое значение через insert_green_cell
    - снимает заливку со всей строки, кроме изменённых ячеек — они
      остаются жёлтыми/зелёными как видимый маркер правки
    - проставляет информацию об изменении
    Возвращает (было ли что-то скорректировано, было ли что-то вставлено).
    """
    rev_src = {v: k for k, v in src_h.items()}
    corrected_cols = set()
    inserted_cols  = set()

    # Сначала зелёные (могут сдвигать ячейки в строке), потом жёлтые —
    # жёлтая правка просто перезаписывает целевую ячейку по имени столбца
    # и не зависит от того, что там было до сдвига.
    for col in sorted(green_cols):
        name    = rev_src.get(col)
        tgt_col = tgt_h.get(name) if name else None
        if tgt_col is None:
            if log:
                log(f"  [!] Зелёный столбец «{name}» не найден в целевом файле — пропущен ({context_msg})")
            continue
        new_val = next((c.value for c in src_row if c.column == col), None)
        if insert_green_cell(tgt_ws, row_idx, tgt_col, boundary, new_val):
            inserted_cols.add(tgt_col)
        elif log:
            log(f"  [!] Нет места для вставки «{name}»={new_val!r} — пропущено ({context_msg})")

    for col in yellow_cols:
        name    = rev_src.get(col)
        tgt_col = tgt_h.get(name) if name else None
        if tgt_col:
            new_val = next((c.value for c in src_row if c.column == col), None)
            tgt_ws.cell(row=row_idx, column=tgt_col, value=new_val)
            corrected_cols.add(tgt_col)

    for cell in tgt_ws[row_idx]:
        if cell.column in corrected_cols:
            cell.fill = copy(YELLOW_FILL)
        elif cell.column in inserted_cols:
            cell.fill = copy(GREEN_FILL)
        else:
            cell.fill = NO_FILL

    chg_col = find_col(tgt_h, *CHANGE_COL_VARIANTS)
    aut_col = find_col(tgt_h, *AUTHOR_COL_VARIANTS)
    if chg_col:
        tgt_ws.cell(row=row_idx, column=chg_col, value=change_txt).fill = NO_FILL
    if aut_col:
        tgt_ws.cell(row=row_idx, column=aut_col, value=author_txt).fill = NO_FILL

    return bool(corrected_cols), bool(inserted_cols)


def is_new_row_green(src_row, src_h, match_cols, green_cols):
    """
    Строка считается добавлением ЦЕЛОЙ новой строки, если ВСЕ непустые
    ключевые ячейки (match_cols), КРОМЕ столбцов иерархии раздела
    (CONTEXT_COLUMNS — унаследованный контекст, не обязательно
    закрашивать отдельно), в исходной строке залиты зелёным — то есть
    найти соответствие в каталоге в принципе не с чем.
    """
    any_green_content = False
    for name in match_cols:
        col = src_h.get(name)
        if col is None:
            continue
        val = next((c.value for c in src_row if c.column == col), None)
        if val in (None, ""):
            continue
        if col not in green_cols:
            if name not in CONTEXT_COLUMNS:
                return False
            continue  # непустой, но некрашеный контекстный столбец — допустимо
        if name not in CONTEXT_COLUMNS:
            any_green_content = True
    return any_green_content


def write_new_row(tgt_ws, dst_row, src_row, src_h, tgt_h, change_txt, author_txt, green_boundary):
    """
    Записывает целиком новую строку (вся строка в исходнике залита
    зелёным) в строку dst_row целевого листа — независимо от того,
    была ли эта строка только что вставлена (insert_rows) или это
    последняя строка листа (добавление в конец). Значения переносятся
    по имени столбца; столбцы, отсутствующие в целевом файле,
    пропускаются.

    Зелёным красятся только столбцы до green_boundary включительно
    (столбец «Тип») — видимый маркер добавления, аналогично тому, как
    удалённая строка целиком красится красным на листе «Удалённое».
    Столбцы правее (служебные — «Поддерживающий сервис», «Изменения по
    запросу» и т.п.) значения получают как обычно, но не красятся.
    """
    rev_src = {v: k for k, v in src_h.items()}

    for cell in src_row:
        name    = rev_src.get(cell.column)
        tgt_col = tgt_h.get(name) if name else None
        if tgt_col is None:
            continue
        dst_cell = tgt_ws.cell(row=dst_row, column=tgt_col)
        fill = copy(GREEN_FILL) if tgt_col <= green_boundary else NO_FILL
        restore(dst_cell, snap(cell), fill_ovr=fill)

    chg_col = find_col(tgt_h, *CHANGE_COL_VARIANTS)
    aut_col = find_col(tgt_h, *AUTHOR_COL_VARIANTS)
    if chg_col:
        c = tgt_ws.cell(row=dst_row, column=chg_col)
        c.value = change_txt
        c.fill  = copy(GREEN_FILL) if chg_col <= green_boundary else NO_FILL
    if aut_col:
        c = tgt_ws.cell(row=dst_row, column=aut_col)
        c.value = author_txt
        c.fill  = copy(GREEN_FILL) if aut_col <= green_boundary else NO_FILL


def append_new_row(tgt_ws, src_row, src_h, tgt_h, change_txt, author_txt, green_boundary):
    """Добавляет целиком новую строку в самый конец целевого листа (запасной
    вариант — используется, когда строку-опору не удалось найти)."""
    dst_row = tgt_ws.max_row + 1
    write_new_row(tgt_ws, dst_row, src_row, src_h, tgt_h, change_txt, author_txt, green_boundary)
    return dst_row


def insert_new_rows_after(tgt_ws, anchor_idx, src_rows, src_h, tgt_h, change_txt, author_txt, green_boundary):
    """
    Вставляет блок целиком новых строк сразу под строкой anchor_idx
    целевого листа (физически раздвигая лист, а не дописывая в конец).
    """
    n = len(src_rows)
    tgt_ws.insert_rows(anchor_idx + 1, amount=n)
    for offset, src_row in enumerate(src_rows):
        write_new_row(tgt_ws, anchor_idx + 1 + offset, src_row, src_h, tgt_h, change_txt, author_txt, green_boundary)


def is_blank_row(row):
    return all(c.value in (None, "") for c in row)


def find_anchor_src_row(src_ws, start_idx, new_row_indices):
    """
    Идёт вверх по исходному листу от start_idx-1, пропуская пустые
    строки и строки, тоже являющиеся целиком новыми (зелёными) — они
    входят в тот же блок вставки и используют общую строку-опору.
    Возвращает номер найденной строки-опоры или None, если дошли до
    начала листа, ничего не найдя.
    """
    idx = start_idx - 1
    while idx >= 2:
        if idx in new_row_indices:
            idx -= 1
            continue
        if is_blank_row(src_ws[idx]):
            idx -= 1
            continue
        return idx
    return None


def resolve_anchor_target_idx(src_ws, anchor_idx, tgt_ws, src_h, tgt_h, match_cols, biz_src, new_row_indices):
    """
    Определяет номер строки в ЦЕЛЕВОМ листе, соответствующей
    строке-опоре anchor_idx исходного листа — той строке, под которой
    нужно вставить блок новых строк. Строка-опора может быть как
    обычной (без заливки), так и скорректированной (частичное
    удаление / точечная жёлтая-зелёная правка) — в этом случае её
    ищут так же, как и при обычной обработке (по итоговому виду,
    достаточно похожему на целевую строку). Если опора сама целиком
    удалена (красная «Бизнес-операция»), поднимается ещё выше в
    поисках существующей строки.
    Возвращает номер строки в tgt_ws или None, если опору найти не удалось.
    """
    idx = anchor_idx
    while idx is not None:
        row = src_ws[idx]
        red    = {c.column for c in row if is_red(c)}
        yellow = {c.column for c in row if is_yellow(c)}
        green  = {c.column for c in row if is_green(c)}

        if red and biz_src in red:
            idx = find_anchor_src_row(src_ws, idx, new_row_indices)
            continue

        if red or yellow or green:
            # Опора уже могла быть скорректирована к этому моменту (опоры
            # резолвятся после основного цикла ops) — частичное удаление,
            # жёлтая правка и/или зелёная вставка, в т.ч. вместе. Их старые
            # (исходные) значения там уже нигде не совпадают — ищем без них.
            ignore_cols = red | yellow | green
            matches = find_rows_fuzzy(tgt_ws, row, src_h, tgt_h, ignore_cols, match_cols)
            return matches[0] if len(matches) == 1 else None

        key = row_key(row, src_h, match_cols)
        return find_row(tgt_ws, key, tgt_h, match_cols)

    return None


# ── основная логика ───────────────────────────────────────────────────────────

def run_processing(src_ws, tgt_ws, del_ws, change_txt, author_txt, log):
    src_h = build_header(src_ws)
    tgt_h = build_header(tgt_ws)
    del_h = build_header(del_ws)
    rev   = {v: k for k, v in src_h.items()}   # col_idx → name

    match_cols = effective_match_columns(src_h, tgt_h, log)

    biz_src  = src_h.get(BIZ_OP_COL)
    boundary = tgt_h.get(PROD_CAT_COL, 9999)
    tema_col = src_h.get("Тема")

    if biz_src is None:
        raise ValueError(f"Столбец «{BIZ_OP_COL}» не найден в исходном файле.")
    if PROD_CAT_COL not in tgt_h:
        log(f"  [!] «{PROD_CAT_COL}» не найден в целевом файле — сдвиг без ограничений")

    green_boundary = tgt_h.get(TYPE_COL, 9999999)
    if TYPE_COL not in tgt_h:
        log(f"  [!] «{TYPE_COL}» не найден в целевом файле — заливка новых строк без ограничения")

    errors, ops, new_rows = [], [], []

    for src_row in src_ws.iter_rows(min_row=2):
        red_src    = {c.column for c in src_row if is_red(c)}
        yellow_src = {c.column for c in src_row if is_yellow(c)}
        green_src  = {c.column for c in src_row if is_green(c)}
        if not red_src and not yellow_src and not green_src:
            continue

        tema = next((c.value for c in src_row if c.column == tema_col), None) if tema_col else None
        n    = src_row[0].row

        if red_src:
            # Ищем строку без учёта жёлтых/зелёных столбцов — их значения в
            # источнике намеренно отличаются от целевого файла (это либо
            # правка, либо новая вставка), точному совпадению мешать не должны.
            ignore_cols = yellow_src | green_src
            matches = find_rows_fuzzy(tgt_ws, src_row, src_h, tgt_h, ignore_cols, match_cols)

            if not matches:
                msg = f"Строка не найдена: Тема='{tema}', исх. строка №{n}"
                errors.append(msg)
                log(f"  [!] {msg}")
                continue
            if len(matches) > 1:
                msg = (f"Неоднозначность ({len(matches)} похожих строк), "
                       f"пропущено: Тема='{tema}', исх. строка №{n}")
                errors.append(msg)
                log(f"  [!] {msg}")
                continue

            idx = matches[0]

            # Если ячейка "Бизнес-операция" красная — удаляем всю строку
            # целиком; жёлтая/зелёная правка в удаляемой строке смысла не
            # имеет и по-прежнему игнорируется.
            if biz_src in red_src:
                log(f"  [✓] УДАЛЕНИЕ: Тема='{tema}' → строка {idx}")
                ops.append((idx, "full", None))
            else:
                # Частичное удаление красных ячеек; если в этой же строке
                # есть жёлтая корректировка и/или зелёная вставка — они
                # применяются вместе (сначала правки, потом удаление, т.к.
                # оно может сдвигать столбцы "Вопрос N" влево).
                names   = {rev[c] for c in red_src if c in rev}
                red_tgt = {tgt_h[nm] for nm in names if nm in tgt_h}
                if yellow_src or green_src:
                    kinds = [f"частичное удаление {sorted(red_tgt)}"]
                    if yellow_src: kinds.append(f"корректировка {sorted(yellow_src)}")
                    if green_src:  kinds.append(f"добавление {sorted(green_src)}")
                    log(f"  [~] {' + '.join(kinds).upper()}: Тема='{tema}' → строка {idx}")
                    ops.append((idx, "partial_adjust", (red_tgt, src_row, yellow_src, green_src)))
                else:
                    log(f"  [~] ЧАСТИЧНОЕ: Тема='{tema}' → строка {idx}, столбцы {sorted(red_tgt)}")
                    ops.append((idx, "partial", red_tgt))
            continue

        if green_src and is_new_row_green(src_row, src_h, match_cols, green_src):
            # Вся строка новая — соответствия в каталоге ещё нет, добавляем в конец
            log(f"  [✓] НОВАЯ СТРОКА: Тема='{tema}', исх. строка №{n}")
            new_rows.append(src_row)
            continue

        # Жёлтая (точечная корректировка) и/или зелёная (добавление ячеек)
        # заливка. Значения изменённых/новых столбцов в источнике и цели
        # намеренно различаются, поэтому ищем строку по совпадению всех
        # ОСТАЛЬНЫХ ключевых столбцов.
        ignore_cols = yellow_src | green_src
        matches = find_rows_fuzzy(tgt_ws, src_row, src_h, tgt_h, ignore_cols, match_cols)

        if not matches:
            msg = f"Похожая строка не найдена: Тема='{tema}', исх. строка №{n}"
            errors.append(msg)
            log(f"  [!] {msg}")
            continue
        if len(matches) > 1:
            msg = (f"Неоднозначность ({len(matches)} похожих строк), "
                   f"пропущено: Тема='{tema}', исх. строка №{n}")
            errors.append(msg)
            log(f"  [!] {msg}")
            continue

        idx = matches[0]
        kinds = []
        if yellow_src: kinds.append(f"корректировка {sorted(yellow_src)}")
        if green_src:  kinds.append(f"добавление {sorted(green_src)}")
        log(f"  [~] {' + '.join(kinds).upper()}: Тема='{tema}' → строка {idx}")
        ops.append((idx, "adjust", (src_row, yellow_src, green_src)))

    # Обрабатываем снизу-вверх: удаления строк не сбивают индексы операций выше
    seen = set()
    ops.sort(key=lambda x: x[0], reverse=True)
    full_cnt = part_cnt = corr_cnt = green_cnt = 0

    for idx, kind, data in ops:
        if idx in seen:
            log(f"  [~] Строка {idx} уже обработана, пропуск")
            continue
        seen.add(idx)

        if kind == "full":
            move_to_deleted(tgt_ws, idx, del_ws, change_txt, author_txt, del_h)
            tgt_ws.delete_rows(idx)
            full_cnt += 1
        elif kind == "partial":
            partial_delete(tgt_ws, idx, data, boundary, change_txt, author_txt, tgt_h)
            part_cnt += 1
        elif kind == "partial_adjust":
            # Сначала жёлтая/зелёная правка (пока столбцы "Вопрос N" ещё в
            # исходном порядке), потом частичное удаление красных ячеек —
            # оно берёт СВЕЖИЙ снимок строки и само сдвигает столбцы влево,
            # так что действует поверх уже применённой правки.
            red_tgt, src_row, yellow_cols, green_cols = data
            tema = next((c.value for c in src_row if c.column == tema_col), None) if tema_col else None
            did_correct, did_insert = adjust_row(
                tgt_ws, idx, src_row, src_h, tgt_h, yellow_cols, green_cols,
                boundary, change_txt, author_txt, log, context_msg=f"Тема='{tema}', строка {idx}",
            )
            partial_delete(tgt_ws, idx, red_tgt, boundary, change_txt, author_txt, tgt_h)
            part_cnt += 1
            if did_correct: corr_cnt += 1
            if did_insert:  green_cnt += 1
        else:
            src_row, yellow_cols, green_cols = data
            tema = next((c.value for c in src_row if c.column == tema_col), None) if tema_col else None
            did_correct, did_insert = adjust_row(
                tgt_ws, idx, src_row, src_h, tgt_h, yellow_cols, green_cols,
                boundary, change_txt, author_txt, log, context_msg=f"Тема='{tema}', строка {idx}",
            )
            if did_correct: corr_cnt += 1
            if did_insert:  green_cnt += 1

    # Целиком новые строки вставляются под своей строкой-опорой (ближайшей
    # предшествующей строкой исходного файла, не входящей в блок новых
    # строк), а не всегда в конец каталога. Несколько подряд идущих новых
    # строк делят одну опору и вставляются вместе, сохраняя порядок.
    new_cnt = 0
    new_row_indices = {r[0].row for r in new_rows}
    groups = {}
    for src_row in new_rows:
        n = src_row[0].row
        anchor = find_anchor_src_row(src_ws, n, new_row_indices)
        groups.setdefault(anchor, []).append(src_row)

    resolved, unresolved = [], []
    for anchor, rows in groups.items():
        tgt_idx = None
        if anchor is not None:
            tgt_idx = resolve_anchor_target_idx(
                src_ws, anchor, tgt_ws, src_h, tgt_h, match_cols, biz_src, new_row_indices,
            )
        if tgt_idx is None:
            unresolved.append(rows)
        else:
            resolved.append((tgt_idx, rows))

    # Вставляем снизу вверх, чтобы вставка одного блока не сбивала уже
    # определённые индексы опор, расположенных выше.
    resolved.sort(key=lambda x: x[0], reverse=True)
    for tgt_idx, rows in resolved:
        temas = ", ".join(
            str(next((c.value for c in r if c.column == tema_col), "")) for r in rows
        ) if tema_col else ""
        log(f"  [✓] НОВАЯ СТРОКА(И) под строкой {tgt_idx}: Тема='{temas}'")
        insert_new_rows_after(tgt_ws, tgt_idx, rows, src_h, tgt_h, change_txt, author_txt, green_boundary)
        new_cnt += len(rows)

    for rows in unresolved:
        for src_row in rows:
            tema = next((c.value for c in src_row if c.column == tema_col), None) if tema_col else None
            log(f"  [!] Строка-опора не найдена — Тема='{tema}' добавлена в конец каталога")
            append_new_row(tgt_ws, src_row, src_h, tgt_h, change_txt, author_txt, green_boundary)
            new_cnt += 1

    return full_cnt, part_cnt, corr_cnt, green_cnt, new_cnt, errors


# ── цветовая схема (тёмная, в духе Telegram) ────────────────────────────────

BG_APP       = "#0e1621"   # фон окна
BG_PANEL     = "#17212b"   # верхняя панель, журнал
BG_INPUT     = "#1e2933"   # поля ввода
BORDER       = "#0b1218"   # тонкая граница полей
TEXT_PRIMARY = "#e5eaf0"   # основной текст
TEXT_SECOND  = "#7d8b99"   # подписи секций
ACCENT       = "#2AABEE"   # фирменный голубой Telegram


def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _blend(c1, c2, t):
    """Смешивает c1→c2 в пропорции t (0..1) — имитация полупрозрачности
    поверх тёмного фона (в Tkinter нет настоящего alpha-blending)."""
    r1, g1, b1 = _hex_to_rgb(c1)
    r2, g2, b2 = _hex_to_rgb(c2)
    return "#%02x%02x%02x" % (
        round(r1 + (r2 - r1) * t),
        round(g1 + (g2 - g1) * t),
        round(b1 + (b2 - b1) * t),
    )


def _rounded_points(x1, y1, x2, y2, r):
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    return [
        x1 + r, y1, x2 - r, y1, x2, y1,
        x2, y1 + r, x2, y2 - r, x2, y2,
        x2 - r, y2, x1 + r, y2, x1, y2,
        x1, y2 - r, x1, y1 + r, x1, y1,
    ]


class GlassButton(tk.Canvas):
    """Кнопка со скруглёнными углами и световым бликом сверху — эмуляция
    эффекта «жидкого стекла» средствами чистого Tkinter (Canvas-полигоны
    со сглаживанием), без сторонних зависимостей."""

    def __init__(self, parent, text, command=None, *, width=180, height=42,
                 radius=13, accent=ACCENT, bg=BG_APP, fg="#ffffff",
                 font=("Segoe UI", 11, "bold")):
        super().__init__(parent, width=width, height=height, bg=bg,
                          highlightthickness=0, bd=0, cursor="hand2")
        self._text, self._command, self._accent = text, command, accent
        self._bg, self._fg, self._font, self._radius = bg, fg, font, radius
        self._state = "normal"
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", lambda e: self._set_state("hover"))
        self.bind("<Leave>", lambda e: self._set_state("normal"))
        self.bind("<ButtonPress-1>", lambda e: self._set_state("press"))
        self.bind("<ButtonRelease-1>", self._on_release)
        self._draw()

    def set_enabled(self, enabled):
        self._state = "normal" if enabled else "disabled"
        self.configure(cursor="hand2" if enabled else "arrow")
        self._draw()

    def _set_state(self, state):
        if self._state == "disabled":
            return
        self._state = state
        self._draw()

    def _on_release(self, event):
        if self._state == "disabled":
            return
        inside = 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height()
        self._state = "hover" if inside else "normal"
        self._draw()
        if inside and self._command:
            self._command()

    def _draw(self):
        self.delete("all")
        w, h = max(self.winfo_width(), 1), max(self.winfo_height(), 1)
        r = min(self._radius, h // 2)

        if self._state == "disabled":
            fill = _blend(self._accent, self._bg, 0.8)
            border = _blend(fill, "#ffffff", 0.1)
            text_fill = _blend(self._fg, self._bg, 0.55)
        else:
            tint = {"normal": 0.55, "hover": 0.4, "press": 0.3}[self._state]
            fill = _blend(self._accent, self._bg, tint)
            border = _blend(fill, "#ffffff", 0.35)
            text_fill = self._fg

        self.create_polygon(_rounded_points(1, 1, w - 1, h - 1, r),
                             smooth=True, fill=fill, outline=border, width=1)

        if self._state != "disabled":
            sheen_h = max(int(h * 0.46), r + 2)
            sheen = _blend(fill, "#ffffff", 0.22 if self._state == "normal" else 0.14)
            self.create_polygon(_rounded_points(2, 2, w - 2, sheen_h, max(r - 1, 0)),
                                 smooth=True, fill=sheen, outline="")

        self.create_text(w // 2, h // 2, text=self._text, fill=text_fill, font=self._font)


# ── интерфейс ─────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  v{APP_VER}")
        self.geometry("760x580")
        self.minsize(620, 480)
        self.configure(bg=BG_APP)
        self.src  = tk.StringVar()
        self.tgt  = tk.StringVar()
        self.chg  = tk.StringVar()
        self.auth = tk.StringVar()
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=BG_PANEL, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text=APP_TITLE, font=("Segoe UI", 12, "bold"),
                 bg=BG_PANEL, fg=TEXT_PRIMARY).pack()

        main = tk.Frame(self, bg=BG_APP, padx=18, pady=10)
        main.pack(fill="both", expand=True)

        self._sec(main, "Файлы")
        self._file_row(main, "Исходный файл:",        self.src,  self._pick_src)
        self._file_row(main, "Корректируемый файл:", self.tgt,  self._pick_tgt)

        self._sec(main, "Информация об изменениях")
        self._entry_row(main, "Изменения по запросу:",   self.chg)
        self._entry_row(main, "Кем внесены изменения:",  self.auth)

        GlassButton(
            main, "▶  Применить корректировки", command=self._run,
            height=44, accent=ACCENT, bg=BG_APP, font=("Segoe UI", 11, "bold"),
        ).pack(fill="x", pady=(14, 4))

        self._sec(main, "Журнал обработки")
        self.log_box = scrolledtext.ScrolledText(
            main, height=10, state="disabled",
            font=("Consolas", 9), bg=BG_PANEL, fg=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY, relief="flat",
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        )
        self.log_box.pack(fill="both", expand=True)
        self.log_box.vbar.configure(
            bg=BG_PANEL, troughcolor=BG_APP, activebackground=ACCENT,
            highlightthickness=0, bd=0, relief="flat", width=10,
        )

    def _sec(self, p, t):
        tk.Label(p, text=t, font=("Segoe UI", 9, "bold"),
                 bg=BG_APP, fg=TEXT_SECOND).pack(anchor="w", pady=(10, 3))

    def _file_row(self, p, lbl, var, cmd):
        f = tk.Frame(p, bg=BG_APP)
        f.pack(fill="x", pady=2)
        tk.Label(f, text=lbl, width=26, anchor="w", bg=BG_APP, fg=TEXT_PRIMARY).pack(side="left")
        tk.Entry(
            f, textvariable=var, state="readonly", width=42, relief="flat", bd=0,
            bg=BG_INPUT, readonlybackground=BG_INPUT, fg=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY, highlightthickness=1,
            highlightbackground=BORDER, highlightcolor=ACCENT,
        ).pack(side="left", padx=4, ipady=3)
        GlassButton(
            f, "Обзор…", command=cmd, width=92, height=30, radius=9,
            accent=ACCENT, bg=BG_APP, font=("Segoe UI", 9),
        ).pack(side="left")

    def _entry_row(self, p, lbl, var):
        f = tk.Frame(p, bg=BG_APP)
        f.pack(fill="x", pady=2)
        tk.Label(f, text=lbl, width=26, anchor="w", bg=BG_APP, fg=TEXT_PRIMARY).pack(side="left")
        tk.Entry(
            f, textvariable=var, width=48, relief="flat", bd=0,
            bg=BG_INPUT, fg=TEXT_PRIMARY, insertbackground=TEXT_PRIMARY,
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        ).pack(side="left", padx=4, ipady=3)

    def _pick(self, var, title):
        p = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel файлы", "*.xlsx *.xls *.xlsm"), ("Все файлы", "*.*")])
        if p:
            var.set(p)

    def _pick_src(self): self._pick(self.src, "Выберите исходный файл")
    def _pick_tgt(self): self._pick(self.tgt, "Выберите корректируемый файл")

    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _run(self):
        src = self.src.get().strip()
        tgt = self.tgt.get().strip()
        if not src or not tgt:
            messagebox.showerror("Ошибка", "Укажите оба файла перед запуском.")
            return
        if not os.path.isfile(src):
            messagebox.showerror("Ошибка", f"Файл не найден:\n{src}")
            return
        if not os.path.isfile(tgt):
            messagebox.showerror("Ошибка", f"Файл не найден:\n{tgt}")
            return

        self._clear_log()
        self._log("=" * 58)
        self._log(f"Исходный:       {os.path.basename(src)}")
        self._log(f"Корректируемый: {os.path.basename(tgt)}")
        self._log("=" * 58)

        try:
            src_wb = load_workbook(src)
            tgt_wb = load_workbook(tgt)

            if DELETED_SHEET not in tgt_wb.sheetnames:
                messagebox.showerror("Ошибка",
                    f"Лист «{DELETED_SHEET}» не найден в корректируемом файле.")
                return

            full, part, corr, green, new, errs = run_processing(
                src_wb.worksheets[0],
                tgt_wb.worksheets[0],
                tgt_wb[DELETED_SHEET],
                self.chg.get().strip(),
                self.auth.get().strip(),
                self._log,
            )

            tgt_wb.save(tgt)
            self._log("-" * 58)
            self._log(
                f"Готово. Удалено строк: {full}  |  "
                f"Частичных удалений: {part}  |  Корректировок: {corr}  |  "
                f"Добавлено ячеек: {green}  |  Новых строк: {new}  |  "
                f"Ошибок: {len(errs)}"
            )

            if errs:
                self._log("\nНе найдены строки:")
                for e in errs:
                    self._log(f"  • {e}")
                messagebox.showwarning(
                    "Завершено с предупреждениями",
                    f"Удалено строк: {full}\nЧастичных удалений: {part}\n"
                    f"Корректировок: {corr}\nДобавлено ячеек: {green}\n"
                    f"Новых строк: {new}\nНе найдено: {len(errs)}\n\n"
                    f"Подробности — в журнале.",
                )
            else:
                messagebox.showinfo(
                    "Готово",
                    f"Обработка завершена.\n"
                    f"Удалено строк: {full}\nЧастичных удалений: {part}\n"
                    f"Корректировок: {corr}\nДобавлено ячеек: {green}\n"
                    f"Новых строк: {new}",
                )

        except PermissionError:
            self._log("\n[ОШИБКА] Файл открыт в Excel — закройте его и повторите.")
            messagebox.showerror(
                "Ошибка доступа",
                "Не удалось сохранить файл.\n"
                "Закройте корректируемый файл в Excel и попробуйте снова.",
            )
        except Exception as exc:
            self._log(f"\n[КРИТИЧЕСКАЯ ОШИБКА] {exc}")
            messagebox.showerror("Критическая ошибка", str(exc))


if __name__ == "__main__":
    App().mainloop()
