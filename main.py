import io
import math
import os
import re
import logging
from dataclasses import dataclass, field
from typing import Dict, Optional, List, Tuple

import pandas as pd
from dotenv import load_dotenv

from telegram import Update, Document
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

load_dotenv()

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(level=LOG_LEVEL, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("wardrobe-bot")

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN is missing in .env")

# Constraints
MAX_SECTION_WIDTH = 1200
MAX_SHELF_SPAN = 800
MAX_FACADE_WIDTH = 600
PARTITION_THRESHOLD = 800

# Типичная плотность ДСП/МДФ (кг/м³)
MATERIAL_DENSITY = 750

SIZE_RE = re.compile(r"(\d+)\s*[xх×*]\s*(\d+)", re.IGNORECASE)
ГАБАРИТ_RE = re.compile(r"(\d{3,4})\s*[xх×*]\s*(\d+)\s*[xх×*]\s*(\d+)", re.IGNORECASE)


@dataclass
class ParsedRow:
    name: str
    thickness_mm: Optional[int] = None
    length_mm: Optional[int] = None
    width_mm: Optional[int] = None
    qty: Optional[float] = None
    material: Optional[str] = None


@dataclass
class FurnitureItem:
    name: str
    code: Optional[str] = None
    qty: Optional[float] = None
    unit: Optional[str] = None


@dataclass
class ParsedSpec:
    source_filename: str
    width_total_mm: int
    depth_mm: int
    height_mm: int
    sections_count: int
    section_width_mm: int
    corpus_rows: List[ParsedRow] = field(default_factory=list)
    furniture_items: List[FurnitureItem] = field(default_factory=list)
    total_weight_kg: float = 0.0


USER_STATE: Dict[int, ParsedSpec] = {}


def _find_sheet_by_keywords(xl, keywords: List[str]) -> Optional[str]:
    """Ищет лист по ключевым словам"""
    sheet_names = xl.sheet_names
    for s in sheet_names:
        s_lower = s.strip().lower()
        if any(kw in s_lower for kw in keywords):
            return s
    return None


def _read_excel_to_sheets(file_bytes: bytes, filename: str) -> Tuple[pd.DataFrame, Optional[pd.DataFrame]]:
    """
    Читает Excel и возвращает (корпус_df, фурнитура_df)
    """
    ext = os.path.splitext(filename.lower())[1]
    bio = io.BytesIO(file_bytes)

    if ext == ".xls":
        xl = pd.ExcelFile(bio, engine="xlrd")
    else:
        xl = pd.ExcelFile(bio, engine="openpyxl")

    # Ищем лист с корпусом
    corpus_sheet = _find_sheet_by_keywords(xl, ["плит", "матер", "корпус", "детал", "дсп"])
    if not corpus_sheet:
        raise ValueError(f"Не найден лист с корпусными деталями. Доступные листы: {xl.sheet_names}")

    df_corpus = xl.parse(corpus_sheet, header=None)

    # Ищем лист с фурнитурой (опционально)
    furniture_sheet = _find_sheet_by_keywords(xl, ["фурнит", "комплект", "метиз"])
    df_furniture = None
    if furniture_sheet:
        df_furniture = xl.parse(furniture_sheet, header=None)

    return df_corpus, df_furniture


def _find_cell_with_text(df: pd.DataFrame, pattern: str) -> Optional[Tuple[int, int]]:
    """Ищет ячейку по регулярному выражению"""
    pat = re.compile(pattern, re.IGNORECASE)
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            v = df.iat[r, c]
            if isinstance(v, str) and pat.search(v):
                return r, c
    return None


def _extract_size_from_text(text: str) -> Optional[Tuple[int, int]]:
    """Извлекает размер вида 2800x600"""
    if not isinstance(text, str):
        return None
    m = SIZE_RE.search(text.replace(" ", ""))
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def _find_column_index(header_row: List[str], keywords: List[str]) -> Optional[int]:
    """Находит индекс колонки по ключевым словам"""
    for i, cell in enumerate(header_row):
        cell_l = str(cell).strip().lower()
        if any(k in cell_l for k in keywords):
            return i
    return None


def _parse_corpus_rows_by_header(df: pd.DataFrame) -> List[ParsedRow]:
    """Парсит корпусные детали по явной строке заголовка."""
    rows: List[ParsedRow] = []
    start_row: Optional[int] = None
    header_row = None

    for r in range(df.shape[0]):
        row_str = " ".join(df.iloc[r].astype(str).str.lower())
        if (
            ("тлщн" in row_str or "толщ" in row_str)
            and "длина" in row_str
            and ("кол-во" in row_str or "кол " in row_str or "колич" in row_str)
        ):
            start_row = r + 1
            header_row = df.iloc[r]
            logger.info(f"Найдена строка заголовков корпуса на позиции {r}: {row_str[:80]}")
            break

    if start_row is None or header_row is None:
        return rows

    header = header_row.astype(str).str.lower().tolist()
    name_idx = next((i for i, h in enumerate(header) if "наимен" in h or "детал" in h or "плита" in h or h == ""), 0)
    thick_idx = next((i for i, h in enumerate(header) if "тлщн" in h or "толщ" in h), None)
    length_idx = next((i for i, h in enumerate(header) if "длина" in h), None)
    width_idx = next((i for i, h in enumerate(header) if "ширина" in h), None)
    qty_idx = next((i for i, h in enumerate(header) if "кол-во" in h or "кол" in h), None)

    for r in range(start_row, df.shape[0]):
        row = df.iloc[r]
        name = str(row.iloc[name_idx]).strip() if name_idx < len(row) else ""
        if not name or name.lower() in ["nan", "итого", "пластик", "ткань", "фурнитура"] or pd.isna(name):
            continue

        thickness_mm = None
        length_mm = None
        width_mm = None
        qty = None

        if thick_idx is not None and thick_idx < len(row):
            thick_val = str(row.iloc[thick_idx])
            m = re.search(r"\d+", thick_val)
            if m:
                thickness_mm = int(m.group(0))

        if length_idx is not None and pd.notna(row.iloc[length_idx]):
            try:
                length_mm = int(float(row.iloc[length_idx]))
            except Exception:
                pass

        if width_idx is not None and pd.notna(row.iloc[width_idx]):
            try:
                width_mm = int(float(row.iloc[width_idx]))
            except Exception:
                pass

        if qty_idx is not None and pd.notna(row.iloc[qty_idx]):
            try:
                qty = float(row.iloc[qty_idx])
            except Exception:
                pass

        if thickness_mm and length_mm and width_mm and qty:
            rows.append(
                ParsedRow(
                    name=name,
                    thickness_mm=thickness_mm,
                    length_mm=length_mm,
                    width_mm=width_mm,
                    qty=qty,
                )
            )

    return rows


def _parse_corpus_rows_heuristic(df: pd.DataFrame) -> List[ParsedRow]:
    """
    Парсит корпусные детали из таблицы.
    Улучшенная версия: ищет строку с "Тлщн" или "Толщ" как начало таблицы
    """
    # ДИАГНОСТИКА: выводим первые 20 строк для понимания структуры
    logger.info(f"DataFrame shape: {df.shape}")
    for r in range(min(20, df.shape[0])):
        row_preview = " | ".join(str(df.iloc[r, c])[:30] for c in range(min(8, df.shape[1])))
        logger.debug(f"Row {r}: {row_preview}")
    
    # Ищем начало таблицы — строку с заголовками
    start_row = None
    for r in range(min(100, df.shape[0])):
        row_str = " ".join(df.iloc[r].astype(str).tolist()).lower()
        # Расширенный список ключевых слов
        keywords = ["тлщн", "толщ", "thickness", "наимен", "детал", "плита", "дсп", "длин", "ширин"]
        if any(kw in row_str for kw in keywords):
            logger.info(f"Найдена строка заголовков на позиции {r}: {row_str[:100]}")
            start_row = r
            break
    
    if start_row is None:
        logger.warning("Не найдена строка заголовков по ключевым словам")
        # Пробуем найти первую строку с числовыми данными
        for r in range(min(50, df.shape[0])):
            row_data = df.iloc[r]
            # Ищем строку где есть хотя бы 2 числа (размеры)
            num_count = sum(1 for v in row_data if pd.notna(v) and str(v).strip().isdigit())
            if num_count >= 2:
                logger.info(f"Найдена потенциальная строка данных на позиции {r}, начинаем оттуда")
                start_row = max(0, r - 1)  # заголовок обычно перед данными
                break
        
        if start_row is None:
            logger.error("Не удалось найти начало таблицы")
            start_row = 0

    header_row = df.iloc[start_row].fillna("").astype(str).tolist()
    logger.info(f"Заголовки: {header_row[:10]}")

    # Определяем колонки
    name_idx = _find_column_index(header_row, ["наимен", "плита", "детал", "описа", "назва"])
    thick_idx = _find_column_index(header_row, ["тлщн", "толщ", "t", "thickness"])
    length_idx = _find_column_index(header_row, ["длин", "length", "l"])
    width_idx = _find_column_index(header_row, ["ширин", "width", "w"])
    size_idx = _find_column_index(header_row, ["размер", "габ", "size"])
    qty_idx = _find_column_index(header_row, ["кол", "кол-во", "количеств", "шт", "qty"])
    mat_idx = _find_column_index(header_row, ["матер", "мдф", "дсп", "material"])

    # Фолбэки
    if name_idx is None:
        name_idx = 0
    if qty_idx is None:
        qty_idx = min(5, df.shape[1] - 1)

    rows: List[ParsedRow] = []
    empty_streak = 0

    for r in range(start_row + 1, df.shape[0]):
        row_data = df.iloc[r]
        name_v = row_data.iloc[name_idx] if name_idx < df.shape[1] else None
        
        if pd.isna(name_v) or (isinstance(name_v, str) and not name_v.strip()):
            empty_streak += 1
            if empty_streak >= 5:
                break
            continue
        
        name = str(name_v).strip()
        
        # Пропускаем итоговые строки
        if any(kw in name.lower() for kw in ["итого", "всего", "total", "сумма"]):
            continue
        
        empty_streak = 0

        # Толщина
        thickness_mm = None
        if thick_idx is not None and thick_idx < df.shape[1]:
            tv = row_data.iloc[thick_idx]
            if pd.notna(tv):
                thick_str = str(tv).strip()
                m = re.search(r"\d+", thick_str)
                if m:
                    thickness_mm = int(m.group(0))

        # Размеры - несколько стратегий
        length_mm = None
        width_mm = None

        # Стратегия 1: отдельные колонки длина и ширина
        if length_idx is not None and length_idx < df.shape[1]:
            lv = row_data.iloc[length_idx]
            if pd.notna(lv):
                try:
                    length_mm = int(float(lv))
                except:
                    pass

        if width_idx is not None and width_idx < df.shape[1]:
            wv = row_data.iloc[width_idx]
            if pd.notna(wv):
                try:
                    width_mm = int(float(wv))
                except:
                    pass

        # Стратегия 2: колонка "Размер" с форматом "2800x600"
        if (length_mm is None or width_mm is None) and size_idx is not None and size_idx < df.shape[1]:
            sv = row_data.iloc[size_idx]
            if pd.notna(sv) and isinstance(sv, str):
                size = _extract_size_from_text(sv)
                if size:
                    length_mm, width_mm = size

        # Стратегия 3: сканируем всю строку на наличие размера
        if length_mm is None or width_mm is None:
            for c in range(df.shape[1]):
                v = row_data.iloc[c]
                if isinstance(v, str):
                    size = _extract_size_from_text(v)
                    if size:
                        length_mm, width_mm = size
                        break

        # Количество
        qty = None
        if qty_idx < df.shape[1]:
            qv = row_data.iloc[qty_idx]
            if pd.notna(qv):
                try:
                    qty = float(qv)
                except:
                    pass

        # Материал
        material = None
        if mat_idx is not None and mat_idx < df.shape[1]:
            mv = row_data.iloc[mat_idx]
            if pd.notna(mv):
                material = str(mv).strip()

        # Добавляем только если есть хоть что-то осмысленное
        if thickness_mm or length_mm or width_mm or qty:
            rows.append(ParsedRow(
                name=name,
                thickness_mm=thickness_mm,
                length_mm=length_mm,
                width_mm=width_mm,
                qty=qty,
                material=material
            ))
            logger.debug(f"Добавлена деталь: {name}, {thickness_mm}мм, {length_mm}×{width_mm}, qty={qty}")

    logger.info(f"Всего распознано деталей: {len(rows)}")
    return rows


def _parse_corpus_rows(df: pd.DataFrame) -> List[ParsedRow]:
    rows = _parse_corpus_rows_by_header(df)
    if rows:
        logger.info(f"Парсинг по заголовку собрал {len(rows)} деталей")
        return rows

    logger.info("Парсинг по заголовку не сработал, пробуем эвристику")
    return _parse_corpus_rows_heuristic(df)


def _parse_furniture_rows(df: pd.DataFrame) -> List[FurnitureItem]:
    """Парсит только реальную фурнитуру, исключает итоги и затраты"""
    items = []
    start_row = None
    for r in range(df.shape[0]):
        row_str = ' '.join(df.iloc[r].astype(str).str.lower())
        if 'код фурнитуры' in row_str or 'наименование фурнитуры' in row_str:
            start_row = r + 1
            header = df.iloc[r].astype(str).str.lower()
            code_idx = header[header.str.contains('код')].index[0] if any(header.str.contains('код')) else None
            name_idx = header[header.str.contains('наимен')].index[0] if any(header.str.contains('наимен')) else 3
            qty_idx = header[header.str.contains('кол')].index[0] if any(header.str.contains('кол')) else None
            unit_idx = header[header.str.contains('ед')].index[0] if any(header.str.contains('ед')) else None
            break

    if start_row is None:
        return items

    for r in range(start_row, df.shape[0]):
        row = df.iloc[r]
        name = str(row.iloc[name_idx]).strip() if name_idx < len(row) else ""
        if (
            not name
            or name.lower() in ['итого', 'рублевая', 'валютная', 'затраты', 'составляющая']
            or pd.isna(name)
            or 'рублев' in name.lower()
            or 'валютн' in name.lower()
            or 'затрат' in name.lower()
        ):
            continue

        code = str(row.iloc[code_idx]).strip() if code_idx is not None and code_idx < len(row) else None
        unit = str(row.iloc[unit_idx]).strip() if unit_idx is not None and unit_idx < len(row) else "шт"

        qty = None
        if qty_idx is not None and qty_idx < len(row):
            try:
                qty = float(row.iloc[qty_idx])
            except:
                pass

        if qty is not None and qty > 0:
            items.append(FurnitureItem(name=name, code=code, qty=qty, unit=unit))

    return items


def _infer_geometry_smart(rows: List[ParsedRow]) -> Tuple[int, int, int, int, int]:
    """
    Умное определение габаритов
    1. Ищем строку с габаритом вида "3000х600х2800"
    2. Анализируем задние стенки
    3. Анализируем крышки/днища для определения глубины
    """
    
    logger.info(f"Начинаем определение габаритов из {len(rows)} строк")
    
    # Стратегия 1: ищем габарит в названии строки
    for row in rows:
        if row.name and isinstance(row.name, str):
            m = ГАБАРИТ_RE.search(row.name)
            if m:
                w, d, h = int(m.group(1)), int(m.group(2)), int(m.group(3))
                logger.info(f"Найден габарит в названии: {w}x{d}x{h}")
                # Пытаемся определить количество секций
                sections = 1
                section_width = w
                back_walls = [r for r in rows if r.name and "задн" in r.name.lower() and r.qty]
                if back_walls and back_walls[0].qty:
                    sections = int(back_walls[0].qty)
                    section_width = w // sections
                return w, d, h, sections, section_width

    # Стратегия 2: анализируем задние стенки
    back_walls = [r for r in rows 
                  if r.name and "задн" in r.name.lower() 
                  and r.length_mm and r.width_mm and r.qty]
    
    logger.info(f"Найдено задних стенок: {len(back_walls)}")
    
    if back_walls:
        bw = back_walls[0]
        logger.info(f"Задняя стенка: {bw.name}, {bw.length_mm}x{bw.width_mm}, qty={bw.qty}")
        # Задняя стенка обычно: высота × ширина_секции
        height = bw.length_mm
        section_width = bw.width_mm
        sections = int(bw.qty)
        width_total = section_width * sections
        
        # Глубину определяем из крышек/дна или боковин
        depth = 600  # дефолт
        top_bottom = [r for r in rows 
                     if r.name and any(kw in r.name.lower() for kw in ["крышк", "дно", "top", "bottom"])
                     and r.width_mm and 300 <= r.width_mm <= 800]
        if top_bottom:
            depth = top_bottom[0].width_mm
            logger.info(f"Глубина из крышки/дна: {depth}")
        
        logger.info(f"Габарит из задней стенки: {width_total}x{depth}x{height}, секций: {sections}")
        return width_total, depth, height, sections, section_width

    # Стратегия 3: общий анализ размеров
    logger.info("Задние стенки не найдены, анализируем все детали")
    
    heights = []
    widths = []
    depths = []
    
    for r in rows:
        if not r.length_mm or not r.width_mm:
            continue
        
        logger.debug(f"Анализ детали: {r.name}, {r.length_mm}x{r.width_mm}")
        
        # Высоты (обычно 2000-3000)
        if 2000 <= r.length_mm <= 3000:
            heights.append(r.length_mm)
        
        # Глубины (обычно 300-700)
        if 300 <= r.width_mm <= 700:
            depths.append(r.width_mm)
        
        # Ширины секций (обычно 600-1200)
        if 600 <= r.width_mm <= 1200:
            widths.append(r.width_mm)
    
    logger.info(f"Найдено высот: {len(heights)}, глубин: {len(depths)}, ширин: {len(widths)}")
    
    if not heights:
        # НОВАЯ СТРАТЕГИЯ: пробуем любые размеры > 1500 как высоту
        for r in rows:
            if r.length_mm and r.length_mm > 1500:
                heights.append(r.length_mm)
        
        if not heights:
            logger.error("Не удалось найти высоту шкафа ни одним способом")
            # Возвращаем дефолтные значения вместо ошибки
            logger.warning("Использую дефолтные габариты: 3000x600x2800")
            return 3000, 600, 2800, 3, 1000
    
    height = max(set(heights), key=heights.count) if heights else 2800
    depth = max(set(depths), key=depths.count) if depths else 600
    section_width = max(set(widths), key=widths.count) if widths else 1000
    
    # Пытаемся угадать количество секций
    sections = 1
    top_bottom = [r for r in rows if r.name and any(kw in r.name.lower() for kw in ["крышк", "дно"]) and r.qty]
    if top_bottom and top_bottom[0].qty:
        sections = max(1, int(top_bottom[0].qty / 2))
    
    width_total = section_width * sections
    
    logger.info(f"Габарит из общего анализа: {width_total}x{depth}x{height}, секций: {sections}")
    return width_total, depth, height, sections, section_width


def _calculate_total_weight(df: pd.DataFrame) -> float:
    """Точный поиск веса — работает с твоими файлами"""
    for r in range(df.shape[0]):
        # Вариант 1: "Вес (кг) =" в колонке A, значение в B
        if str(df.iloc[r, 0]).strip().lower().startswith('вес (кг)'):
            try:
                val = str(df.iloc[r, 1]).strip().replace(',', '.')
                return float(val)
            except:
                pass
        # Вариант 2: в одной ячейке или строке
        for c in range(min(10, df.shape[1])):
            cell = str(df.iloc[r, c])
            m = re.search(r'Вес\s*\(кг\)\s*=\s*(\d+[.,]?\d*)', cell, re.IGNORECASE)
            if m:
                return float(m.group(1).replace(',', '.'))
    return 0.0


def _calculate_total_weight_by_rows(rows: List[ParsedRow]) -> float:
    """Рассчитывает общий вес изделия из геометрии деталей"""
    total_kg = 0.0
    for r in rows:
        if r.length_mm and r.width_mm and r.thickness_mm and r.qty:
            volume_m3 = (r.length_mm / 1000) * (r.width_mm / 1000) * (r.thickness_mm / 1000)
            weight_kg = volume_m3 * MATERIAL_DENSITY * r.qty
            total_kg += weight_kg
    return round(total_kg, 2)


def _split_sections(total_width: int) -> List[int]:
    """Разбивает общую ширину на секции"""
    n = math.ceil(total_width / MAX_SECTION_WIDTH)
    base = total_width // n
    rem = total_width % n
    return [base + (1 if i < rem else 0) for i in range(n)]


def _calc_spans_for_section(section_w: int) -> int:
    """Рассчитывает количество пролётов в секции"""
    spans_by_shelf = math.ceil(section_w / MAX_SHELF_SPAN)
    spans_by_facade = math.ceil(section_w / MAX_FACADE_WIDTH)
    spans = max(spans_by_shelf, spans_by_facade)
    if section_w >= PARTITION_THRESHOLD:
        spans = max(spans, 2)
    return spans


def _recalculate_corpus(spec: ParsedSpec, new_width: int) -> Tuple[List[Dict], float]:
    old_width = spec.width_total_mm
    new_sections = _split_sections(new_width)
    new_sections_count = len(new_sections)

    old_spans = sum(_calc_spans_for_section(spec.section_width_mm) for _ in range(spec.sections_count))
    new_spans = sum(_calc_spans_for_section(w) for w in new_sections)
    span_ratio = new_spans / old_spans if old_spans > 0 else 1

    new_parts = []
    for row in spec.corpus_rows:
        name_low = row.name.lower()
        new_qty = row.qty
        new_length = row.length_mm
        new_width_part = row.width_mm

        if 'полк' in name_low:
            new_qty *= span_ratio
            new_width_part = new_width // new_spans  # Обновляем размер полки по пролёту
        elif 'фасад' in name_low:
            new_qty *= span_ratio
            new_width_part = new_width // new_spans  # Фасад по пролёту
        elif 'задн' in name_low:
            new_qty = new_sections_count
            new_width_part = new_width // new_sections_count
        elif 'крышк' in name_low or 'дно' in name_low:
            new_qty = new_sections_count * 2  # По 2 на секцию? Уточните из оригинала
            new_length = new_width // new_sections_count  # Если крышка по ширине секции
        elif 'боков' in name_low or 'средние' in name_low or 'стенк' in name_low:
            new_qty = new_sections_count + 1  # Боковины + средние
        elif 'цоколь' in name_low:
            new_qty = new_sections_count
            new_length = new_width // new_sections_count
        else:
            new_qty *= (new_width / old_width)

        new_parts.append({
            'name': row.name,
            'thickness': row.thickness_mm,
            'length_mm': new_length,
            'width_mm': new_width_part,
            'qty': math.ceil(new_qty),  # Округление в большую сторону для производства
            'size': f"{new_length}×{new_width_part}"
        })

    # Точный вес по объёму (учтёт все материалы)
    new_weight = 0.0
    for p in new_parts:
        vol_m3 = (p['length_mm'] / 1000) * (p['width_mm'] / 1000) * (p['thickness'] / 1000) * p['qty']
        new_weight += vol_m3 * MATERIAL_DENSITY  # 750 кг/м³ — можно варьировать по материалу, если добавить в ParsedRow

    # Оценка веса фурнитуры (если нужно точно — добавьте вес на позицию в FurnitureItem)
    furn_weight = sum(p['qty'] * 0.05 for p in _recalculate_furniture(spec, new_width))  # ~50 г на шт
    new_weight += furn_weight

    return new_parts, round(new_weight, 2)


def _petals_per_facade(height_mm: int) -> int:
    if height_mm <= 900: return 2
    elif height_mm <= 1400: return 3
    elif height_mm <= 1900: return 4
    elif height_mm <= 2400: return 5
    elif height_mm <= 2800: return 7
    else: return 8


def _recalculate_furniture(spec: ParsedSpec, new_width: int) -> List[dict]:
    old_spans = sum(_calc_spans_for_section(spec.section_width_mm) for _ in range(spec.sections_count))
    new_sections = _split_sections(new_width)
    new_spans = sum(_calc_spans_for_section(w) for w in new_sections)
    span_ratio = new_spans / old_spans if old_spans > 0 else 1
    section_ratio = len(new_sections) / spec.sections_count if spec.sections_count > 0 else 1

    # Фасады из корпуса
    old_facades = next((r.qty for r in spec.corpus_rows if 'фасад' in r.name.lower()), old_spans)
    new_facades = old_facades * span_ratio

    # Высота фасада
    facade_row = next((r for r in spec.corpus_rows if 'фасад' in r.name.lower()), None)
    facade_height = facade_row.length_mm if facade_row else 2700
    petals_per_f = _petals_per_facade(facade_height)

    new_furn = []
    for item in spec.furniture_items:
        name_low = item.name.lower()
        new_qty = item.qty or 0

        if 'петл' in name_low or 'чашк' in name_low or 'заглушка' in name_low and 'петл' in name_low:
            new_qty = new_facades * petals_per_f
        elif 'ручк' in name_low:
            new_qty = new_facades  # 1 на фасад
        elif 'полкодерж' in name_low:
            new_qty *= span_ratio
        elif 'стяжка межсекцион' in name_low:
            new_qty = (len(new_sections) - 1) * (item.qty / (spec.sections_count - 1)) if spec.sections_count > 1 else 0
        elif 'корректор фасада' in name_low:
            new_qty = new_facades
        elif 'винт' in name_low or 'ключ' in name_low:
            new_qty = math.ceil(item.qty * section_ratio)  # По секциям, фиксировано
        elif 'штанг' in name_low:  # Для штанг
            new_qty = len(new_sections)  # По секциям (если штанга на секцию)
        elif 'подсветк' in name_low:  # Для подсветки
            new_qty *= span_ratio  # По пролётам
        else:
            new_qty *= span_ratio  # Остальное по пролётам

        new_furn.append({
            'name': item.name,
            'code': item.code,
            'qty': math.ceil(new_qty),  # Всегда целое, в большую сторону
            'unit': item.unit or 'шт'
        })

    return new_furn


def _format_structure(width_total: int, depth: int, height: int, sections: List[int]) -> str:
    """Форматирует описание структуры"""
    spans_per_section = [_calc_spans_for_section(w) for w in sections]
    total_spans = sum(spans_per_section)
    partitions = sum((s - 1) for w, s in zip(sections, spans_per_section) if w >= PARTITION_THRESHOLD)

    lines = [
        f"📏 Габарит: {width_total}×{depth}×{height} мм (Ш×Г×В)",
        f"📦 Секции: {len(sections)} шт → " + " | ".join(f"{x}мм" for x in sections),
        f"🔲 Пролёты (полка≤{MAX_SHELF_SPAN}, фасад≤{MAX_FACADE_WIDTH}): " +
        " | ".join(f"{w}мм→{s}" for w, s in zip(sections, spans_per_section)) +
        f" (всего {total_spans})",
    ]
    
    if partitions > 0:
        lines.append(f"📐 Вертикальные перегородки внутри секций (при ≥{PARTITION_THRESHOLD}мм): {partitions} шт")
    
    return "\n".join(lines)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "👋 Привет! Я бот для пересчёта спецификаций шкафов.\n\n"
        "📤 Пришли мне Excel-файл (.xls или .xlsx) с калькуляцией шкафа.\n"
        "Я распознаю габариты, структуру и материалы, затем помогу пересчитать под новую ширину.\n\n"
        "Файл должен содержать листы с корпусными деталями и фурнитурой.\n\n"
        "Доступные команды:\n"
        "/start - показать это сообщение\n"
        "/debug - включить подробные логи (для разработчика)"
    )


async def debug_mode(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Включает подробный режим логирования"""
    logging.getLogger().setLevel(logging.DEBUG)
    logger.setLevel(logging.DEBUG)
    await update.message.reply_text("🔧 Режим отладки включен. Теперь в логах будет больше деталей.")



async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    doc: Document = update.message.document
    user_id = update.effective_user.id

    if not doc.file_name.lower().endswith((".xls", ".xlsx")):
        await update.message.reply_text("⚠️ Нужен Excel-файл (.xls или .xlsx)")
        return

    await update.message.reply_text("⏳ Обрабатываю файл...")

    try:
        tg_file = await doc.get_file()
        file_bytes = await tg_file.download_as_bytearray()
        file_bytes = bytes(file_bytes)

        df_corpus, df_furniture = _read_excel_to_sheets(file_bytes, doc.file_name)
        
        corpus_rows = _parse_corpus_rows(df_corpus)
        logger.info(f"Распознано {len(corpus_rows)} строк корпуса")
        
        furniture_items = _parse_furniture_rows(df_furniture) if df_furniture is not None else []
        logger.info(f"Распознано {len(furniture_items)} позиций фурнитуры")
        
        width_total, depth, height, sections, section_width = _infer_geometry_smart(corpus_rows)
        total_weight = _calculate_total_weight(df_corpus)
        if not total_weight:
            total_weight = _calculate_total_weight_by_rows(corpus_rows)
        
        spec = ParsedSpec(
            source_filename=doc.file_name,
            width_total_mm=width_total,
            depth_mm=depth,
            height_mm=height,
            sections_count=sections,
            section_width_mm=section_width,
            corpus_rows=corpus_rows,
            furniture_items=furniture_items,
            total_weight_kg=total_weight
        )
        
        USER_STATE[user_id] = spec

        sections_list = [section_width] * sections
        msg = "✅ Файл успешно обработан!\n\n"
        msg += _format_structure(width_total, depth, height, sections_list)
        msg += f"\n\n📊 Найдено:\n"
        msg += f"  • Корпусных деталей: {len([r for r in corpus_rows if r.qty])} позиций\n"
        msg += f"  • Фурнитуры: {len(furniture_items)} позиций\n"
        msg += f"  • Общий вес: {total_weight} кг\n"
        msg += f"\n💬 Введи новую ширину шкафа в мм (например: 3600)"
        
        await update.message.reply_text(msg)
        
    except Exception as e:
        logger.exception("Failed to process document")
        await update.message.reply_text(f"❌ Ошибка обработки файла:\n{str(e)}\n\nПопробуй другой файл или обратись к разработчику.")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user_id = update.effective_user.id
    text = (update.message.text or "").strip()

    if user_id not in USER_STATE:
        await update.message.reply_text("⚠️ Сначала пришли Excel-файл с калькуляцией.\nИспользуй /start для инструкций.")
        return

    # Парсим число
    m = re.search(r"\d+", text.replace(" ", ""))
    if not m:
        await update.message.reply_text("⚠️ Введи новую ширину числом в мм.\nНапример: 3600")
        return

    new_width = int(m.group(0))
    if new_width < 300 or new_width > 10000:
        await update.message.reply_text("⚠️ Ширина должна быть от 300 до 10000 мм.")
        return

    spec = USER_STATE[user_id]
    
    await update.message.reply_text("🔄 Пересчитываю спецификацию...")

    try:
        sections = _split_sections(new_width)
        corpus_parts, new_weight = _recalculate_corpus(spec, new_width)

        old_spans = sum(_calc_spans_for_section(spec.section_width_mm) for _ in range(spec.sections_count))
        new_spans = sum(_calc_spans_for_section(w) for w in sections)
        new_facades = new_spans

        furniture_items = _recalculate_furniture(spec, new_width)

        # Формируем ответ
        msg = "✅ Пересчёт завершён!\n\n"
        msg += _format_structure(new_width, spec.depth_mm, spec.height_mm, sections)
        msg += f"\n\n⚖️ Вес изделия:\n"
        msg += f"  • Было: {spec.total_weight_kg} кг\n"
        msg += f"  • Стало: {new_weight} кг\n"
        msg += f"  • Разница: {new_weight - spec.total_weight_kg:+.2f} кг\n"
        
        msg += f"\n\n🔨 КОРПУСНЫЕ ДЕТАЛИ ({len(corpus_parts)} поз.):\n"
        for i, p in enumerate(corpus_parts, 1):
            thick_str = f" (т.{p['thickness']}мм)" if p.get('thickness') else ""
            mat_str = f" [{p['material']}]" if p.get('material') else ""
            msg += f"{i}. {p['name']}{thick_str}{mat_str}\n"
            msg += f"   {p['size']} — {p['qty']} шт\n"
        
        if furniture_items:
            msg += f"\n🔩 ФУРНИТУРА ({len(furniture_items)} поз.):\n"
            for i, f in enumerate(furniture_items, 1):
                code_str = f" [{f['code']}]" if f.get('code') else ""
                qty_str = f"{f['qty']:.1f}" if f.get('qty') else "—"
                unit_str = f.get('unit', 'шт')
                msg += f"{i}. {f['name']}{code_str}\n"
                msg += f"   {qty_str} {unit_str}\n"
        
        # Разбиваем на несколько сообщений если слишком длинное
        if len(msg) > 4096:
            # Telegram limit
            parts = []
            current_part = ""
            for line in msg.split('\n'):
                if len(current_part) + len(line) + 1 > 4000:
                    parts.append(current_part)
                    current_part = line + '\n'
                else:
                    current_part += line + '\n'
            if current_part:
                parts.append(current_part)
            
            for part in parts:
                await update.message.reply_text(part)
        else:
            await update.message.reply_text(msg)
        
        # Предложение пересчитать ещё раз
        await update.message.reply_text(
            "💡 Хочешь пересчитать под другую ширину? Просто введи новое значение в мм.\n"
            "Или пришли новый Excel-файл для другого изделия."
        )
        
    except Exception as e:
        logger.exception("Failed to recalculate")
        await update.message.reply_text(f"❌ Ошибка пересчёта:\n{str(e)}")


def main() -> None:
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("debug", debug_mode))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Bot started")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
