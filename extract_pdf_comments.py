import argparse
import re
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd


ART_MAP = {
    "Highlight": "Markierung",
    "StrikeOut": "Durchgestrichen",
    "Underline": "Unterstrichen",
    "Squiggly": "Wellenlinie",
    "Text": "Notiz",
    "Ink": "Freihandmarkierung",
}


def normalize_spaces(text: str) -> str:
    if text is None:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def sort_words(words):
    """
    Sortiert Wörter stabil nach Block, Zeile, Wortnummer.
    PyMuPDF words Format:
    (x0, y0, x1, y1, "word", block_no, line_no, word_no)
    """
    return sorted(words, key=lambda w: (w[5], w[6], w[7], round(w[1], 2), w[0]))


def vertices_to_quads(vertices):
    """
    Wandelt eine Vertices Liste in 4er Blöcke um.
    Bei Highlight und StrikeOut liegt pro Textzeile meist ein Quad vor.
    """
    if not vertices:
        return []

    quads = []
    for i in range(0, len(vertices), 4):
        quad = vertices[i:i + 4]
        if len(quad) == 4:
            quads.append(quad)
    return quads


def quad_to_rect(quad):
    xs = [p[0] for p in quad]
    ys = [p[1] for p in quad]
    return fitz.Rect(min(xs), min(ys), max(xs), max(ys))


def join_words_to_text(words):
    if not words:
        return ""

    words = sort_words(words)

    lines = []
    current_line = []
    last_block = None
    last_line = None

    for w in words:
        block_no, line_no = w[5], w[6]
        token = w[4]

        if last_block is None:
            current_line.append(token)
        elif block_no == last_block and line_no == last_line:
            current_line.append(token)
        else:
            lines.append(" ".join(current_line))
            current_line = [token]

        last_block = block_no
        last_line = line_no

    if current_line:
        lines.append(" ".join(current_line))

    return normalize_spaces(" ".join(lines))


def extract_marked_text(page, annot):
    """
    Extrahiert die markierte Textpassage.
    Für Highlight, StrikeOut, Underline und Squiggly werden bevorzugt die Quads genutzt.
    Falls das nichts liefert, wird auf das Annotationsrechteck zurückgefallen.
    """
    annot_type = annot.type[1]
    supported_markup = {"Highlight", "StrikeOut", "Underline", "Squiggly"}

    if annot_type not in supported_markup:
        return ""

    words = page.get_text("words")
    if not words:
        return ""

    collected = []

    quads = vertices_to_quads(annot.vertices)
    for quad in quads:
        rect = quad_to_rect(quad)
        quad_words = []

        for w in words:
            word_rect = fitz.Rect(w[:4])
            if not word_rect.intersects(rect):
                continue

            overlap = word_rect & rect
            if overlap.is_empty:
                continue

            # Nur Wörter übernehmen, die tatsächlich relevant im markierten Bereich liegen
            if overlap.get_area() >= 0.2 * word_rect.get_area():
                quad_words.append(w)

        if quad_words:
            collected.extend(sort_words(quad_words))

    text = join_words_to_text(collected)

    # Fallback, falls Quads nichts geliefert haben
    if not text:
        rect_words = []
        rect = annot.rect
        for w in words:
            word_rect = fitz.Rect(w[:4])
            if not word_rect.intersects(rect):
                continue

            overlap = word_rect & rect
            if overlap.is_empty:
                continue

            if overlap.get_area() >= 0.2 * word_rect.get_area():
                rect_words.append(w)

        text = join_words_to_text(rect_words)

    return text


def extract_comments(pdf_path, start_page, end_page):
    """
    Extrahiert Kommentare einer PDF im Bereich [start_page, end_page].
    Die Seitenzahlen sind 1-basiert, also wie im PDF-Viewer angezeigt.
    """
    rows = []
    doc = fitz.open(pdf_path)

    start_index = max(start_page - 1, 0)
    end_index = min(end_page - 1, len(doc) - 1)

    for page_index in range(start_index, end_index + 1):
        page = doc[page_index]
        annot = page.first_annot

        while annot:
            info = annot.info or {}
            annot_type_en = annot.type[1]
            annot_type_de = ART_MAP.get(annot_type_en, annot_type_en)

            comment_text = normalize_spaces(info.get("content", ""))
            marked_text = extract_marked_text(page, annot)

            rows.append(
                {
                    "Seite": page_index + 1,
                    "Kommentare": comment_text,
                    "Text": marked_text,
                    "Art": annot_type_de,
                }
            )

            annot = annot.next

    doc.close()
    return pd.DataFrame(rows)


def main():
    parser = argparse.ArgumentParser(
        description="Extrahiert PDF-Kommentare und speichert sie als Excel und CSV."
    )
    parser.add_argument("pdf", help="Pfad zur PDF-Datei")
    parser.add_argument("--start", type=int, required=True, help="Erste PDF-Seite, z.B. 23")
    parser.add_argument("--end", type=int, required=True, help="Letzte PDF-Seite, z.B. 54")
    parser.add_argument(
        "--out",
        default="Kommentare_Extrakt",
        help="Basisname der Ausgabedateien ohne Endung",
    )
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    out_base = Path(args.out)

    df = extract_comments(pdf_path, args.start, args.end)

    xlsx_path = out_base.with_suffix(".xlsx")
    csv_path = out_base.with_suffix(".csv")

    df.to_excel(xlsx_path, index=False)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    print(f"Fertig. {len(df)} Kommentare exportiert.")
    print(f"Excel: {xlsx_path.resolve()}")
    print(f"CSV:   {csv_path.resolve()}")


if __name__ == "__main__":
    main()
