from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import List

import cv2
import numpy as np
import pandas as pd
import pytesseract


@dataclass
class Cell:
    x: int
    y: int
    w: int
    h: int


def preprocess_for_grid(image: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    # Realca linhas da tabela e reduz ruído
    bw = cv2.adaptiveThreshold(
        gray,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        21,
        10,
    )
    return bw


def find_table_cells(binary: np.ndarray) -> List[Cell]:
    h, w = binary.shape

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, w // 35), 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(20, h // 35)))

    horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
    vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel, iterations=1)

    grid = cv2.add(horizontal, vertical)
    grid = cv2.dilate(grid, np.ones((3, 3), np.uint8), iterations=1)

    contours, _ = cv2.findContours(grid, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    cells: List[Cell] = []
    min_area = max(150, (h * w) // 20000)

    for cnt in contours:
        x, y, cw, ch = cv2.boundingRect(cnt)
        area = cw * ch
        if area < min_area:
            continue
        if cw > int(w * 0.95) and ch > int(h * 0.95):
            continue
        if cw < 18 or ch < 14:
            continue
        cells.append(Cell(x, y, cw, ch))

    # remove duplicados aproximados
    dedup: List[Cell] = []
    for c in sorted(cells, key=lambda a: (a.y, a.x, a.w, a.h)):
        if dedup:
            p = dedup[-1]
            if abs(c.x - p.x) <= 3 and abs(c.y - p.y) <= 3 and abs(c.w - p.w) <= 4 and abs(c.h - p.h) <= 4:
                continue
        dedup.append(c)

    return dedup


def cluster_positions(values: List[int], tolerance: int) -> List[int]:
    if not values:
        return []
    values = sorted(values)
    clusters = [[values[0]]]
    for v in values[1:]:
        if abs(v - int(np.mean(clusters[-1]))) <= tolerance:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    return [int(round(float(np.mean(c)))) for c in clusters]


def nearest_index(points: List[int], value: int) -> int:
    return int(np.argmin([abs(value - p) for p in points]))


def clean_text(text: str) -> str:
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip(" |\t")


def ocr_cell(image: np.ndarray) -> str:
    if image.size == 0:
        return ""

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    text = pytesseract.image_to_string(th, config="--oem 3 --psm 6", lang="por+eng")
    return clean_text(text)


def build_matrix(image: np.ndarray, cells: List[Cell]) -> List[List[str]]:
    if not cells:
        return []

    heights = [c.h for c in cells]
    widths = [c.w for c in cells]

    row_tol = max(8, int(np.median(heights) * 0.5))
    col_tol = max(8, int(np.median(widths) * 0.45))

    row_points = cluster_positions([c.y + c.h // 2 for c in cells], row_tol)
    col_points = cluster_positions([c.x + c.w // 2 for c in cells], col_tol)

    matrix = [["" for _ in col_points] for _ in row_points]

    for c in sorted(cells, key=lambda a: (a.y, a.x)):
        r = nearest_index(row_points, c.y + c.h // 2)
        col = nearest_index(col_points, c.x + c.w // 2)

        pad = 2
        y1 = max(0, c.y + pad)
        y2 = min(image.shape[0], c.y + c.h - pad)
        x1 = max(0, c.x + pad)
        x2 = min(image.shape[1], c.x + c.w - pad)

        txt = ocr_cell(image[y1:y2, x1:x2])
        if txt:
            matrix[r][col] = txt if not matrix[r][col] else f"{matrix[r][col]} {txt}".strip()

    # remove linhas totalmente vazias
    matrix = [row for row in matrix if any(cell.strip() for cell in row)]
    return matrix


def find_header_row(rows: List[List[str]]) -> int:
    keywords = ("codigo", "descr", "consumo", "molde")
    best_idx = 0
    best_score = -1
    for i, row in enumerate(rows[:8]):
        line = " ".join(row).lower()
        score = sum(1 for kw in keywords if kw in line)
        if score > best_score:
            best_score = score
            best_idx = i
    return best_idx


def normalize_columns(cols: List[str]) -> List[str]:
    norm = []
    for c in cols:
        c_low = c.lower()
        if "molde" in c_low:
            norm.append("Os molde")
        elif "codigo" in c_low:
            norm.append("Codigo do item")
        elif "descr" in c_low:
            norm.append("Descrição da peça")
        elif "dia" in c_low:
            norm.append("Consumo dia")
        elif "mes" in c_low or "mês" in c_low:
            norm.append("Consumo mês")
        else:
            norm.append(c if c else "coluna")
    # garante nomes únicos
    seen = {}
    unique = []
    for c in norm:
        seen[c] = seen.get(c, 0) + 1
        unique.append(c if seen[c] == 1 else f"{c}_{seen[c]}")
    return unique


def extract_table_to_csv(input_image: Path, output_csv: Path, tesseract_cmd: str | None = None) -> None:
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    image = cv2.imread(str(input_image))
    if image is None:
        raise FileNotFoundError(f"Não foi possível abrir a imagem: {input_image}")

    binary = preprocess_for_grid(image)
    cells = find_table_cells(binary)
    rows = build_matrix(image, cells)

    if not rows:
        raise RuntimeError("Não foi possível detectar células da tabela. Tente imagem com melhor contraste.")

    header_idx = find_header_row(rows)
    header = normalize_columns(rows[header_idx])
    data_rows = rows[header_idx + 1 :]

    # Ajusta largura das linhas para o número de colunas do cabeçalho
    width = len(header)
    fixed_rows = []
    for row in data_rows:
        if len(row) < width:
            row = row + [""] * (width - len(row))
        elif len(row) > width:
            row = row[:width]
        if any(cell.strip() for cell in row):
            fixed_rows.append(row)

    df = pd.DataFrame(fixed_rows, columns=header)

    # Remove linhas claramente fora da tabela (ex.: título)
    if "Codigo do item" in df.columns:
        mask_title = df["Codigo do item"].str.contains("inje", case=False, na=False)
        df = df.loc[~mask_title]

    output_csv.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_csv, index=False, encoding="utf-8-sig")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extrai tabela de imagem e salva em CSV (OCR)."
    )
    parser.add_argument("input_image", type=Path, help="Caminho da imagem (png/jpg)")
    parser.add_argument("output_csv", type=Path, help="Caminho do CSV de saída")
    parser.add_argument(
        "--tesseract-cmd",
        type=str,
        default=None,
        help="Caminho do executável tesseract (se não estiver no PATH)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    extract_table_to_csv(args.input_image, args.output_csv, args.tesseract_cmd)
    print(f"CSV gerado em: {args.output_csv}")


if __name__ == "__main__":
    main()
