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
# Добавляем русскую х и звездочку
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


@dataclass
class SectionType:
    """Тип секции шкафа"""

    width_mm: int
    has_rod: bool = False
    has_shelves: bool = False
    has_lighting: bool = False
    shelf_count: float = 0


def _distribute_items_per_section(total_qty: float, sections_count: int) -> List[float]:
    """Распределяет количество элементов по секциям, сохраняя исходное соотношение."""

    if sections_count <= 0:
        return []

    base_per_section = math.floor(total_qty / sections_count)
    remainder = int(round(total_qty - base_per_section * sections_count))

    distribution = [float(base_per_section) for _ in range(sections_count)]

    for i in range(remainder):
        distribution[i % sections_count] += 1

    return distribution


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


def _infer_material(name: str, material_value: Optional[str] = None) -> Optional[str]:
    """Определяет материал по явному столбцу или по названию детали."""
    for source in (material_value, name):
        if not source:
            continue
        text = str(source).strip().lower()
        if "фанер" in text:
            return "фанера"
        if "мдф" in text:
            return "мдф"
        if "лдсп" in text or "дсп" in text:
            return "лдсп"
    return material_value.strip() if isinstance(material_value, str) and material_value.strip() else None


def _determine_material(name: str, thickness_mm: Optional[int], row_context: Optional[str] = None) -> str:
    """Выбирает материал по толщине, названию и строке исходной таблицы."""

    name_low = (name or "").lower()
    context_low = (row_context or "").lower() or name_low

    material = "ЛДСП"
    if "мдф" in context_low or thickness_mm == 18:
        material = "МДФ"
    elif "фанера" in context_low or thickness_mm in [20, 24, 27]:
        material = "Фанера"
    elif "egger 16" in context_low or thickness_mm == 16:
        material = "ЛДСП Egger 16мм"

    if "фасад" in name_low and material == "МДФ":
        # Если в строке указаны нулевые операции (нет покраски/эмали), считаем фасад ЛДСП
        zero_ops = re.findall(r"\b0[,.]0{2,}\b", context_low)
        if zero_ops:
            material = "ЛДСП"

    return material


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
        row_str = " ".join(str(x) for x in row.tolist())
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

        material = _determine_material(name, thickness_mm, row_str)

        if thickness_mm and length_mm and width_mm and qty:
            rows.append(
                ParsedRow(
                    name=name,
                    thickness_mm=thickness_mm,
                    length_mm=length_mm,
                    width_mm=width_mm,
                    qty=qty,
                    material=material,
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
        material_value = None
        if mat_idx is not None and mat_idx < df.shape[1]:
            mv = row_data.iloc[mat_idx]
            if pd.notna(mv):
                material_value = str(mv).strip()

        row_context = " ".join(str(x) for x in row_data.tolist())
        material = _determine_material(name, thickness_mm, row_context if material_value is None else material_value)

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
    else:
        logger.info("Парсинг по заголовку не сработал, пробуем эвристику")
        rows = _parse_corpus_rows_heuristic(df)

    for r in rows:
        if r.name and "фанера" in r.name.lower() and not r.material:
            r.material = "фанера"

    return rows


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
            material_hint = f"{r.name} {r.material or ''}".lower()
            if "фанер" in material_hint:
                density = 600
            elif "мдф" in material_hint:
                density = 800
            else:
                density = MATERIAL_DENSITY
            volume_m3 = (r.length_mm / 1000) * (r.width_mm / 1000) * (r.thickness_mm / 1000)
            weight_kg = volume_m3 * density * r.qty
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


def _distribute_width_evenly(total_width: int, parts: int) -> List[int]:
    """Равномерно распределяет ширину между частями, сохраняя сумму."""

    if parts <= 0:
        return []

    base, remainder = divmod(total_width, parts)
    return [base + (1 if i < remainder else 0) for i in range(parts)]


def _calculate_span_widths(sections: List[int]) -> List[int]:
    """Возвращает фактические ширины пролётов по секциям."""

    span_widths: List[int] = []
    for sec_width in sections:
        spans = _calc_spans_for_section(sec_width)
        span_widths.extend(_distribute_width_evenly(sec_width, spans))
    return span_widths


def _analyze_section_types(spec: ParsedSpec) -> List[SectionType]:
    """Анализирует функциональные зоны шкафа"""
    sections: List[SectionType] = []

    rods = [r for r in spec.corpus_rows if r.name and 'штанг' in r.name.lower()]
    total_rods = sum(r.qty for r in rods if r.qty) if rods else 0

    shelves = [r for r in spec.corpus_rows if r.name and 'полк' in r.name.lower()]
    total_shelves = sum(r.qty for r in shelves if r.qty) if shelves else 0

    lights = [
        f for f in spec.furniture_items
        if any(kw in f.name.lower() for kw in ['подсвет', 'led', 'освещ'])
    ]
    has_lighting = len(lights) > 0

    rods_per_section = total_rods / spec.sections_count if spec.sections_count > 0 else 0
    shelves_distribution = _distribute_items_per_section(total_shelves, spec.sections_count) if total_shelves else []
    shelves_per_section = total_shelves / spec.sections_count if spec.sections_count > 0 else 0

    for idx in range(spec.sections_count):
        section = SectionType(
            width_mm=spec.section_width_mm,
            has_rod=(rods_per_section > 0),
            has_shelves=(shelves_per_section > 0),
            shelf_count=shelves_distribution[idx] if idx < len(shelves_distribution) else shelves_per_section,
            has_lighting=has_lighting,
        )
        sections.append(section)

    return sections


def _recalculate_corpus(
    spec: ParsedSpec, new_width: int
) -> Tuple[List[Dict], float, List[str], List[str], List[dict]]:
    old_width = spec.width_total_mm
    original_sections_types = _analyze_section_types(spec)
    new_sections = _split_sections(new_width)
    new_sections_count = len(new_sections)
    new_span_widths = _calculate_span_widths(new_sections)

    if new_width == old_width:
        logger.info("Ширина не изменилась — возвращаем исходные данные без пересчёта.")
        corpus_parts = [
            {
                'name': r.name,
                'material': r.material,
                'thickness': r.thickness_mm,
                'length_mm': r.length_mm,
                'width_mm': r.width_mm,
                'qty': r.qty,
                'size': f"{r.length_mm}×{r.width_mm}",
            }
            for r in spec.corpus_rows
        ]

        cut_warnings: List[str] = []
        for part in corpus_parts:
            if part.get('length_mm') and part.get('width_mm'):
                warning = _check_material_sheet_limits(part)
                if warning:
                    cut_warnings.append(warning)

        furn_items = [
            {
                'name': f_item.name,
                'code': f_item.code,
                'qty': f_item.qty,
                'unit': f_item.unit or 'шт',
            }
            for f_item in spec.furniture_items
        ]

        return [
            part | {'widths_mm': []}
            for part in corpus_parts
        ], spec.total_weight_kg, cut_warnings, [], furn_items

    old_spans = sum(_calc_spans_for_section(spec.section_width_mm) for _ in range(spec.sections_count))
    new_spans = sum(_calc_spans_for_section(w) for w in new_sections)
    span_ratio = new_spans / old_spans if old_spans > 0 else 1
    section_ratio = new_sections_count / spec.sections_count if spec.sections_count else 1

    old_polki = next((r.qty for r in spec.corpus_rows if 'полк' in r.name.lower()), 0) or 0

    # Сопоставляем новые секции со старыми типами (пропорционально)
    section_type_map: List[SectionType] = []
    for i, new_sec_width in enumerate(new_sections):
        original_idx = int(i * len(original_sections_types) / len(new_sections)) if original_sections_types else 0
        original_type = original_sections_types[min(original_idx, len(original_sections_types) - 1)] if original_sections_types else SectionType(width_mm=new_sec_width)

        new_type = SectionType(
            width_mm=new_sec_width,
            has_rod=original_type.has_rod,
            has_shelves=original_type.has_shelves,
            has_lighting=original_type.has_lighting,
            shelf_count=original_type.shelf_count,
        )
        section_type_map.append(new_type)

    shelves_plan = [sec.shelf_count for sec in section_type_map]
    if section_type_map:
        logger.info(f"Карта полок по секциям (оригинальные→новые): {shelves_plan}")

    # Карта материалов
    material_map: Dict[str, str] = {}
    for row in spec.corpus_rows:
        name_key = row.name.lower()
        if 'фасад' in name_key:
            material_map['фасад'] = row.material or 'МДФ'
        elif 'боков' in name_key or 'стенк' in name_key:
            if any(kw in name_key for kw in ['видим', 'наружн', 'внешн']):
                material_map['стенка_видимая'] = row.material or 'ЛДСП'
            else:
                material_map['стенка_внутр'] = row.material or 'ЛДСП'
        elif 'полк' in name_key:
            material_map['полка'] = row.material or 'ЛДСП'
        elif 'крышк' in name_key or 'дно' in name_key:
            material_map['крышка'] = row.material or 'ЛДСП'
        elif 'перегород' in name_key or 'стойк' in name_key:
            material_map['перегородка'] = row.material or 'ЛДСП'

    new_parts = []
    for row in spec.corpus_rows:
        name_low = row.name.lower()
        new_qty = row.qty or 0
        new_length = row.length_mm or 0
        new_width_part = row.width_mm or 0
        widths_mm: List[int] = []
        facade_target_qty: Optional[int] = None

        if 'полк' in name_low:
            shelves_from_ratio = sum(shelves_plan) if shelves_plan else 0
            if not shelves_from_ratio and spec.sections_count:
                shelves_from_ratio = (old_polki / spec.sections_count) * new_sections_count
            new_qty = shelves_from_ratio or new_qty
            new_width_part = math.ceil(new_width / new_sections_count) if new_sections_count else new_width_part
        elif 'фасад' in name_low:
            facades_per_span = new_qty / old_spans if old_spans else new_qty
            new_qty = facades_per_span * new_spans
            facades_per_span_int = max(1, int(round(facades_per_span))) if new_spans else 0

            if new_span_widths and facades_per_span_int:
                for span_w in new_span_widths:
                    widths_mm.extend(_distribute_width_evenly(span_w, facades_per_span_int))

            facade_target_qty = math.ceil(new_qty) if new_qty else 0

            if widths_mm:
                if facade_target_qty and len(widths_mm) != facade_target_qty:
                    if len(widths_mm) > facade_target_qty:
                        widths_mm = widths_mm[:facade_target_qty]
                    else:
                        widths_mm.extend([widths_mm[-1]] * (facade_target_qty - len(widths_mm)))
                new_width_part = max(widths_mm)
            else:
                new_width_part = new_width // new_spans if new_spans else new_width_part
        elif 'задн' in name_low:
            new_qty = new_sections_count
            new_width_part = new_width // new_sections_count if new_sections_count else new_width_part
        elif 'крышк' in name_low or 'дно' in name_low:
            # Определяем логику из исходного файла
            pieces_per_section = row.qty / spec.sections_count if spec.sections_count > 0 else 2  # e.g. 6/3=2
            if row.length_mm and row.length_mm > spec.section_width_mm * 1.5:
                # Это ЦЕЛЬНАЯ крышка на весь шкаф (B)
                new_qty = row.qty  # обычно 2 (верх + низ)
                new_length = new_width
            else:
                # Это крышки ПО СЕКЦИЯМ (A)
                new_qty = new_sections_count * pieces_per_section
                new_length = new_width // new_sections_count
        elif 'боков' in name_low:
            new_qty = 2
        elif 'средние' in name_low or 'перегород' in name_low:
            new_qty = new_sections_count - 1 if new_sections_count > 1 else 0
        elif 'стенк' in name_low:
            new_qty = new_sections_count + 1
        elif 'цоколь' in name_low:
            new_qty = new_sections_count
            new_length = new_width // new_sections_count if new_sections_count else new_length
        else:
            new_qty *= (new_width / old_width) if old_width else section_ratio

        inferred_material = row.material
        if not inferred_material:
            if 'фасад' in name_low:
                inferred_material = material_map.get('фасад', 'МДФ')
            elif 'боков' in name_low or 'стенк' in name_low:
                inferred_material = material_map.get('стенка_внутр', 'ЛДСП')
                if any(kw in name_low for kw in ['видим', 'наружн', 'внешн']):
                    inferred_material = material_map.get('стенка_видимая', inferred_material)
            elif 'полк' in name_low:
                inferred_material = material_map.get('полка', 'ЛДСП')
            elif any(kw in name_low for kw in ['крышк', 'дно']):
                inferred_material = material_map.get('крышка', 'ЛДСП')
            elif any(kw in name_low for kw in ['перегород', 'стойк']):
                inferred_material = material_map.get('перегородка', 'ЛДСП')

        new_parts.append({
            'name': row.name,
            'material': inferred_material,
            'thickness': row.thickness_mm,
            'length_mm': new_length,
            'width_mm': new_width_part,
            'widths_mm': widths_mm,
            'qty': facade_target_qty if facade_target_qty is not None else math.ceil(new_qty),
            'size': f"{new_length}×" + (" / ".join(str(w) for w in widths_mm) if widths_mm else f"{new_width_part}")
        })

    logger.debug(new_parts)

    new_weight = 0.0
    for p in new_parts:
        if p['thickness'] and p['length_mm'] and p['qty'] and (p['width_mm'] or p.get('widths_mm')):
            length_adj = p['length_mm']
            material_hint = f"{p['name']} {p.get('material') or ''}".lower()
            if 'фанер' in material_hint:
                density = 600
            elif 'мдф' in material_hint:
                density = 800
            else:
                density = MATERIAL_DENSITY

            widths_for_weight = p.get('widths_mm') or []
            qty_value = p['qty']

            if widths_for_weight:
                for width_item in widths_for_weight[:qty_value]:
                    vol_m3 = (length_adj / 1000) * (width_item / 1000) * (p['thickness'] / 1000)
                    new_weight += vol_m3 * density

                remaining_qty = max(0, qty_value - len(widths_for_weight))
                if remaining_qty:
                    width_adj = widths_for_weight[-1]
                    vol_m3 = (length_adj / 1000) * (width_adj / 1000) * (p['thickness'] / 1000) * remaining_qty
                    new_weight += vol_m3 * density
            else:
                width_adj = p['width_mm']
                vol_m3 = (length_adj / 1000) * (width_adj / 1000) * (p['thickness'] / 1000) * qty_value
                new_weight += vol_m3 * density

    furn_items, furn_warnings, _ = _recalculate_furniture(spec, new_width)
    if new_width == old_width:
        furn_weight = sum(f.qty * 0.05 for f in spec.furniture_items)
    else:
        furn_weight = sum(f['qty'] * 0.05 for f in furn_items)
    new_weight = new_weight + furn_weight

    cut_warnings: List[str] = []
    general_recommendations: List[str] = []
    for p in new_parts:
        warning = _check_material_sheet_limits(p)
        if warning:
            cut_warnings.append(warning)

    if spec.height_mm > 2500:
        general_recommendations.append("⚠️ Устойчивость: добавить антиопрокидывание")

    general_recommendations.extend(furn_warnings)

    return new_parts, round(new_weight, 2), cut_warnings, general_recommendations, furn_items


def _check_material_sheet_limits(part: dict) -> Optional[str]:
    """Проверяет влезает ли деталь в стандартный лист и выдаёт предупреждение"""
    material = (part.get('material') or '').lower()
    length = part['length_mm']
    widths = part.get('widths_mm') or [part['width_mm']]

    if 'лдсп' in material or 'дсп' in part['name'].lower():
        max_l, max_w = 2800, 2070
    elif 'мдф' in material:
        max_l, max_w = 2800, 2070
    elif 'фанер' in material:
        max_l, max_w = 2440, 1220
    else:
        max_l, max_w = 2800, 2070

    for width in widths:
        if max(length, width) > max_l or min(length, width) > max_w:
            return f"⚠️ {part['name']} ({length}×{width}) не влезает в лист {max_l}×{max_w} - нужна стыковка"
    return None


def _petals_per_facade(height_mm: int) -> int:
    if height_mm <= 900: return 2
    elif height_mm <= 1400: return 3
    elif height_mm <= 1900: return 4
    elif height_mm <= 2400: return 5
    elif height_mm <= 2800: return 7
    else: return 8


def _calculate_shelf_counts(spec: ParsedSpec, new_width: int) -> Tuple[float, float]:
    """Возвращает исходное и новое количество полок для пересчёта фурнитуры."""

    old_shelves = sum(r.qty for r in spec.corpus_rows if r.name and 'полк' in r.name.lower() and r.qty)
    new_sections = _split_sections(new_width)
    original_sections_types = _analyze_section_types(spec)

    section_type_map: List[SectionType] = []
    for i, new_sec_width in enumerate(new_sections):
        original_idx = int(i * len(original_sections_types) / len(new_sections)) if original_sections_types else 0
        original_type = (
            original_sections_types[min(original_idx, len(original_sections_types) - 1)]
            if original_sections_types
            else SectionType(width_mm=new_sec_width)
        )

        section_type_map.append(
            SectionType(
                width_mm=new_sec_width,
                has_rod=original_type.has_rod,
                has_shelves=original_type.has_shelves,
                has_lighting=original_type.has_lighting,
                shelf_count=original_type.shelf_count,
            )
        )

    shelves_plan = [sec.shelf_count for sec in section_type_map]
    new_shelves = sum(shelves_plan)

    if not new_shelves and old_shelves and spec.sections_count:
        new_shelves = (old_shelves / spec.sections_count) * len(new_sections)

    return float(old_shelves), float(new_shelves)


def _recalculate_furniture(spec: ParsedSpec, new_width: int) -> Tuple[List[dict], List[str], float]:
    old_spans = sum(_calc_spans_for_section(spec.section_width_mm) for _ in range(spec.sections_count))
    new_sections = _split_sections(new_width)
    new_spans = sum(_calc_spans_for_section(w) for w in new_sections)
    span_ratio = new_spans / old_spans if old_spans > 0 else 1
    section_ratio = len(new_sections) / spec.sections_count if spec.sections_count > 0 else 1

    old_shelves, new_shelves = _calculate_shelf_counts(spec, new_width)

    facade_row = next((r for r in spec.corpus_rows if 'фасад' in r.name.lower()), None)
    old_facades = facade_row.qty if facade_row and facade_row.qty is not None else old_spans
    facades_per_span = old_facades / old_spans if old_spans else old_facades
    new_facades = math.ceil(facades_per_span * new_spans)
    facade_height = facade_row.length_mm if facade_row else 2700
    petals_per_f = _petals_per_facade(facade_height)

    drawers_rows = [r for r in spec.corpus_rows if r.name and 'ящик' in r.name.lower()]
    old_drawers = sum(r.qty for r in drawers_rows if r.qty) if drawers_rows else 0
    recalculated_drawers = old_drawers * section_ratio if old_drawers else 0
    handles_drawer_qty = math.ceil(recalculated_drawers) if recalculated_drawers else 0
    logger.info(
        "Ящики: исходное qty=%s, коэффициент секций=%.2f, пересчитано=%s",
        old_drawers,
        section_ratio,
        handles_drawer_qty,
    )

    new_furn: List[dict] = []
    furn_warnings: List[str] = []
    total_led_power = 0.0
    total_led_length_m = 0.0
    if old_shelves == 0 or new_shelves == 0:
        furn_warnings.append("⚠️ Недостаточно данных по полкам — использован пересчёт по пролётам.")

    spans_per_section = [_calc_spans_for_section(w) for w in new_sections]
    span_width = new_width / new_spans if new_spans else new_width
    handle_drawer_warning_added = False

    for item in spec.furniture_items:
        name_low = item.name.lower()
        base_qty = item.qty or 0
        new_qty = base_qty
        meta: Dict[str, Optional[float]] = {}

        if 'петл' in name_low or 'чашк' in name_low or ('заглушка' in name_low and 'петл' in name_low):
            new_qty = new_facades * petals_per_f
        elif 'ручк' in name_low:
            new_qty = new_facades + handles_drawer_qty
            if old_drawers and handles_drawer_qty and not handle_drawer_warning_added:
                furn_warnings.append(
                    f"ℹ️ Ручки: учтены ящики {old_drawers}→{handles_drawer_qty} (коэф. секций {section_ratio:.2f})"
                )
                handle_drawer_warning_added = True
            logger.info(
                "Ручки: фасады=%s, ящики=%s, итоговое qty=%s (исходное qty=%s)",
                new_facades,
                handles_drawer_qty,
                new_qty,
                base_qty,
            )
        elif 'полкодерж' in name_low:
            if old_shelves > 0 and new_shelves > 0:
                supports_per_shelf = base_qty / old_shelves
                new_qty = supports_per_shelf * new_shelves
            else:
                new_qty *= span_ratio
        elif 'стяжка межсекцион' in name_low:
            stiazki_per_connection = max(1, math.ceil(spec.height_mm / 700))
            new_qty = (len(new_sections) - 1) * stiazki_per_connection if len(new_sections) > 1 else 0
        elif 'корректор фасада' in name_low:
            new_qty = new_facades
        elif 'винт' in name_low or 'ключ' in name_low:
            new_qty = math.ceil(base_qty) if base_qty else 2
        elif 'штанг' in name_low:
            # Штанги ставятся только в секциях-гардеробных
            # Определяем какие секции имели штанги изначально
            original_has_rods = base_qty > 0
            if not original_has_rods:
                new_qty = 0
            else:
                rods_per_original_section = base_qty / spec.sections_count if spec.sections_count > 0 else 1
                new_qty = 0
                lengths_mm: List[int] = []
                for w in new_sections:
                    if rods_per_original_section > 0:
                        rods_in_section = math.ceil(rods_per_original_section)
                        new_qty += rods_in_section
                        rod_length = max(w - 40, 0)
                        lengths_mm.extend([rod_length] * rods_in_section)
                meta['lengths_mm'] = lengths_mm
        elif 'подсветк' in name_low or 'led' in name_low or 'освещен' in name_low:
            new_qty *= span_ratio
            led_length_mm = max(int(span_width - 100), 0)
            total_length_m = (led_length_mm / 1000) * new_qty
            power = total_length_m * 10
            total_led_power += power
            total_led_length_m += total_length_m
            meta['power_w'] = power
            meta['length_mm'] = led_length_mm
        else:
            new_qty *= span_ratio

        new_furn.append({
            'name': item.name,
            'code': item.code,
            'qty': math.ceil(new_qty),
            'unit': item.unit or 'шт',
            **meta
        })

    if total_led_power > 50:
        furn_warnings.append("⚠️ LED: Нужен доп. блок (мощность > 50 Вт)")
    if total_led_length_m > 0:
        blocks_needed = math.ceil(total_led_length_m / 5)
        furn_warnings.append(f"ℹ️ LED: Блок питания x{blocks_needed}, ≥{round(total_led_power * 1.2, 1)} Вт")

    return new_furn, furn_warnings, total_led_power


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
        corpus_parts, new_weight, cut_warnings, general_recommendations, furniture_items = _recalculate_corpus(spec, new_width)

        # Формируем ответ
        msg = "✅ Пересчёт завершён!\n\n"
        msg += _format_structure(new_width, spec.depth_mm, spec.height_mm, sections)
        msg += f"\n\n⚖️ Вес изделия:\n"
        msg += f"  • Было: {spec.total_weight_kg} кг\n"
        msg += f"  • Стало: {new_weight} кг\n"
        msg += f"  • Разница: {new_weight - spec.total_weight_kg:+.2f} кг\n"
        
        msg += f"\n\n🔨 КОРПУСНЫЕ ДЕТАЛИ ({len(corpus_parts)} поз.):\n"
        for i, p in enumerate(corpus_parts, 1):
            thick_str = f"т.{p['thickness']}мм" if p.get('thickness') else ""
            mat_str = f"{p['material']}" if p.get('material') else "ЛДСП"
            attrs = ", ".join([x for x in [thick_str, mat_str] if x])
            msg += f"{i}. {p['name']}\n"
            msg += f"   📐 {p['size']} ({attrs}) × {p['qty']} шт\n"

        if furniture_items:
            msg += f"\n🔩 ФУРНИТУРА ({len(furniture_items)} поз.):\n"
            for i, f in enumerate(furniture_items, 1):
                code_str = f" [{f['code']}]" if f.get('code') else ""
                qty_str = f"{f['qty']:.1f}" if f.get('qty') else "—"
                unit_str = f.get('unit', 'шт')
                meta_parts = []
                if 'power_w' in f:
                    meta_parts.append(f"мощн. {round(f['power_w'], 2)} Вт")
                if 'length_mm' in f:
                    meta_parts.append(f"дл. {f['length_mm']} мм")
                if 'lengths_mm' in f:
                    lengths = ", ".join(str(l) for l in f['lengths_mm'])
                    meta_parts.append(f"длины: {lengths} мм")
                meta_str = f" ({'; '.join(meta_parts)})" if meta_parts else ""
                msg += f"{i}. {f['name']}{code_str}\n"
                msg += f"   🔧 {qty_str} {unit_str}{meta_str}\n"

        if cut_warnings:
            msg += "\n\n⚠️ Предупреждения по раскрою:\n"
            for w in cut_warnings:
                msg += f"  • {w}\n"

        if general_recommendations:
            msg += "\n\nℹ️ Рекомендации:\n"
            for rec in general_recommendations:
                msg += f"  • {rec}\n"
        
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
