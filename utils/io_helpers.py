import docx
import os
from config import log

def read_docx_text(docx_path: str) -> str:
    """Extracts all text (paragraphs + tables) from a DOCX in logical order."""
    log.info(f"Reading document: {docx_path}")
    doc = docx.Document(docx_path)
    blocks = []
    current = []
    seen = set()

    elements = list(doc.element.body)

    def cell_text(cell):
        return cell.text.strip()

    for el in elements:
        if el.tag.endswith("p"):
            para = el
            txt = para.text.strip()
            if txt:
                if txt not in seen:
                    current.append(txt)
                    seen.add(txt)
            else:
                if current:
                    blocks.append("\n".join(current))
                    current = []
        elif el.tag.endswith("tbl"):
            tbl = next(t for t in doc.tables if t._element == el)
            for row in tbl.rows:
                row_items = []
                row_seen = set()
                for cell in row.cells:
                    ct = cell_text(cell)
                    if ct and ct not in row_seen:
                        row_items.append(ct)
                        row_seen.add(ct)
                if row_items:
                    line = " | ".join(row_items)
                    if line not in seen:
                        current.append(line)
                        seen.add(line)

    if current:
        blocks.append("\n".join(current))
    log.info(f"Extracted {len(blocks)} text blocks")
    return " \n".join(blocks)

def export_text_file(text: str, txt_path: str):
    """Write provided text to a UTF-8 .txt file."""
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(text)

def remove_file(path: str):
    """Safely delete a file if it exists."""
    try:
        os.remove(path)
    except OSError:
        pass
