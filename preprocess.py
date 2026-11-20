#!/usr/bin/env python
# coding: utf-8

# In[5]:


import os
import re
import glob
import pandas as pd
from __future__ import annotations


# Возможные варианты написания заголовка колонки со ссылкой
ID_COLUMN_CANDIDATES = ["sNO6x href", "sNO6x  href", "href"]

# Точные названия колонок с признаками (как в файле)
COLUMN_MAP = {
    "area":         ["s6KKu 2"],
    "rooms":        ["s6KKu"],
    "floor":        ["s6KKu 3"],
    "metro_minutes":["si5bH"],
    "price":        ["s5oAg"],
}


def _norm(s: str) -> str:
    """Нормализация заголовка: нижний регистр, схлопывание пробелов."""
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def _find_column_by_names(df, candidates: list[str]) -> int | None:
    """Ищет индекс колонки по точному (нормализованному) совпадению заголовка."""
    cand_norm = {_norm(c) for c in candidates}
    for col_idx in df.columns:
        for row_idx in range(len(df)):  # Ограничение на данные изначально было: range(min(5, len(df)))
            cell = _norm(df.iat[row_idx, col_idx])
            if cell in cand_norm:
                return col_idx
    return None


def find_id_column(df: pd.DataFrame) -> int | None:
    """Ищет индекс колонки, содержащей ссылки на объявления domclick."""
    # 1. По точному совпадению заголовка в первых строках
    col = _find_column_by_names(df, ID_COLUMN_CANDIDATES)
    if col is not None:
        return col

    # 2. Fallback: ищем колонку, где большинство значений — ссылки на domclick card
    for col_idx in df.columns:
        col_vals = df[col_idx].dropna().astype(str)
        if len(col_vals) == 0:
            continue
        hits = col_vals.str.contains(r"/card/[a-z_]+__\d+", regex=True).sum()
        if hits >= max(1, len(col_vals) * 0.5):
            return col_idx

    return None


def find_feature_columns(df: pd.DataFrame) -> dict[str, int | None]:
    """Ищет индексы колонок для каждого признака по названиям."""
    result = {}
    for feat, names in COLUMN_MAP.items():
        result[feat] = _find_column_by_names(df, names)
    return result


def extract_id_from_url(text: str) -> str | None:
    """Достаёт числовой id из URL объявления domclick."""
    if not isinstance(text, str):
        return None
    m = re.search(r"sale__flat__(\d+)", text)
    if m:
        return m.group(1)
    m = re.search(r"/card/[a-z_]+__(\d+)", text)
    if m:
        return m.group(1)
    return None


# ---------- Парсеры отдельных признаков ----------

def parse_area(text: str) -> float | None:
    """Площадь (вещественное число). Пример: '50,1 м²' -> 50.1"""
    if not isinstance(text, str):
        return None
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*м[²2]?", text)
    if m:
        try:
            return float(m.group(1).replace(",", "."))
        except ValueError:
            return None
    # если единиц нет — пробуем взять просто число
    m = re.search(r"(\d+(?:[.,]\d+)?)", text)
    if m:
        try:
            return float(m.group(1).replace(",", "."))
        except ValueError:
            return None
    return None


def parse_rooms(text: str) -> int | None:
    """Кол-во комнат. Пример: '3-комн. квартира' -> 3, 'Студия' -> 0"""
    if not isinstance(text, str):
        return None
    m = re.search(r"(\d+)\s*-\s*комн", text, re.IGNORECASE)
    if m:
        return int(m.group(1))
    if re.search(r"студи", text, re.IGNORECASE):
        return 0
    return None


def parse_floor(text: str) -> tuple[int | None, int | None]:
    """Этаж/всего этажей. Пример: '8/9 эт.' -> (8, 9)"""
    if not isinstance(text, str):
        return None, None
    m = re.search(r"(\d+)\s*/\s*(\d+)", text)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def parse_metro_minutes(text: str) -> int | None:
    """Время до метро (целое число минут). Пример: '~18 мин.' -> 18"""
    if not isinstance(text, str):
        return None
    m = re.search(r"(\d+)\s*мин", text, re.IGNORECASE)
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)", text)
    if m:
        return int(m.group(1))
    return None


def parse_price(text: str) -> float | None:
    """Цена (вещественное число). Пример: '17 900 000 ₽' -> 17900000.0"""
    if not isinstance(text, str):
        return None
    # убираем пробелы (включая неразрывные) и знак рубля
    cleaned = re.sub(r"[\s\u00a0\u202f₽]", "", text)
    # берём первое число (возможно с дробной частью)
    m = re.search(r"(\d+(?:[.,]\d+)?)", cleaned)
    if m:
        try:
            return float(m.group(1).replace(",", "."))
        except ValueError:
            return None
    return None


def _get_cell(row: pd.Series, col_idx: int | None) -> str:
    """Безопасно достаёт значение ячейки по индексу колонки."""
    if col_idx is None:
        return ""
    try:
        return str(row.iloc[col_idx])
    except (IndexError, KeyError):
        return ""


def process_excel_file(path: str) -> pd.DataFrame:
    """Читает Excel-файл и возвращает DataFrame с id + признаками."""
    df = pd.read_excel(path, header=None, dtype=str)
    df = df.fillna("")

    id_col = find_id_column(df)
    if id_col is None:
        print(f"  ⚠ не найдена колонка со ссылкой в {os.path.basename(path)}")
        return pd.DataFrame()

    feat_cols = find_feature_columns(df)
    missing = [f for f, c in feat_cols.items() if c is None]
    if missing:
        print(f"  ⚠ не найдены колонки для признаков: {', '.join(missing)} "
              f"(в {os.path.basename(path)})")

    rows = []
    for _, row in df.iterrows():
        id_value = extract_id_from_url(_get_cell(row, id_col))
        if id_value is None:
            continue  # пропускаем строки без id (шапка, пустые и т.п.)

        floor, total_floors = parse_floor(_get_cell(row, feat_cols["floor"]))
        feats = {
            "id": id_value,
            "rooms": parse_rooms(_get_cell(row, feat_cols["rooms"])),
            "area": parse_area(_get_cell(row, feat_cols["area"])),
            "floor": floor,
            "total_floors": total_floors,
            "metro_minutes": parse_metro_minutes(_get_cell(row, feat_cols["metro_minutes"])),
            "price": parse_price(_get_cell(row, feat_cols["price"])),
            "source_file": os.path.basename(path),
        }
        rows.append(feats)

    return pd.DataFrame(rows)


def main(folder: str, out_dir: str = "features", out_name: str = "all_features.csv"):
    files = glob.glob(os.path.join(folder, "*.xls*"))
    if not files:
        print(f"В папке '{folder}' не найдено Excel-файлов.")
        return

    os.makedirs(out_dir, exist_ok=True)

    all_frames = []
    for f in files:
        print(f"Обработка: {f}")
        try:
            df = process_excel_file(f)
            print(f"  → извлечено записей: {len(df)}")
            if not df.empty:
                all_frames.append(df)
        except Exception as e:
            print(f"  ⚠ ошибка: {e}")

    if not all_frames:
        print("Не удалось извлечь данные.")
        return

    result = pd.concat(all_frames, ignore_index=True)
    cols = ["id", "rooms", "area", "floor", "total_floors",
            "metro_minutes", "price", "source_file"]
    result = result[cols]

    if result["id"].notna().any():
        before = len(result)
        result = result.drop_duplicates(subset=["id"], keep="first")
        if before != len(result):
            print(f"Удалено дублей по id: {before - len(result)}")

    combined_path = os.path.join(out_dir, out_name)
    result.to_csv(combined_path, index=False, encoding="utf-8-sig")
    print(f"\nИтого сохранено {len(result)} строк в '{combined_path}'")


if __name__ == "__main__":
    main(folder="./data", out_dir="./features")

