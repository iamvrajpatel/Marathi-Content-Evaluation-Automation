import os
import re
from spylls.hunspell import Dictionary
from sklearn.feature_extraction.text import CountVectorizer
import pandas as pd
from config import log, TERM_MAPPING

def segment_by_sentences(text: str, sentences_per_segment: int = 15) -> list[str]:
    """Split text into chunks of roughly N sentences (by period counts)."""
    segments, buffer, cnt = [], [], 0
    for ch in text:
        buffer.append(ch)
        if ch == ".":
            cnt += 1
            if cnt >= sentences_per_segment:
                segments.append("".join(buffer).strip())
                buffer, cnt = [], 0
    tail = "".join(buffer).strip()
    if tail:
        segments.append(tail)
    return segments

def check_word_presence(words: list[str], txt_path: str) -> pd.DataFrame:
    """Check which words from the list appear in the text file."""
    results = []
    content = open(txt_path, "r", encoding="utf-8").read()
    for w in words:
        status = "Present" if w in content else "Not Present"
        results.append((w, status))
    return pd.DataFrame(results, columns=["Word", "Status"])

def init_hunspell_dict() -> Dictionary | None:
    """Load Marathi Hunspell dictionary from local folder."""
    base = os.path.join(os.path.dirname(__file__), "mr_IN")
    try:
        return Dictionary.from_files(base)
    except Exception:
        log.warning("Hunspell dictionary load failed", exc_info=True)
        return None

def strip_specials(text: str) -> str:
    """Remove punctuation, digits (both Latin and Devanagari), special quotes."""
    specials = set("-:.?&,!@#$%^*()_+={}[]|\\;\"'<>/`~०१२३४५६७८९”“’‘")
    return "".join(" " if c in specials else c for c in text)

def sanitize_marathi_text(text: str) -> str:
    """Keep only Marathi letters and whitespace."""
    t = re.sub(r'\d+', '', text)
    t = re.sub(r'[^\u0900-\u097F\s]', '', t)
    return re.sub(r'\s+', ' ', t).strip()

def identify_misspellings(text: str, dictionary: Dictionary|None) -> list[str]:
    """Return list of tokens not recognized by the Hunspell dictionary."""
    miss = []
    for token in text.split():
        word = token.strip('.,!?()[]{}:;"\'')
        if dictionary and not dictionary.lookup(word) and word not in TERM_MAPPING.values():
            miss.append(word)
    return sorted(set(miss))

def find_term_replacements(full_text: str) -> pd.DataFrame:
    """List each key in TERM_MAPPING that appears in the text."""
    rows = []
    for src, dst in TERM_MAPPING.items():
        if src in full_text:
            rows.append((src, dst))
    return pd.DataFrame(rows, columns=["Word", "Replacement"])

def detect_repeated_phrases(doc: "docx.Document") -> pd.DataFrame:
    """
    Scan each table cell’s last two tokens,
    flag any that repeat within the same cell.
    """
    snippets, metadata = [], []
    for ti, tbl in enumerate(doc.tables):
        for ri, row in enumerate(tbl.rows):
            for ci, cell in enumerate(row.cells):
                raw = cell.text.strip()
                if not raw:
                    continue
                parts = raw.split(":", 1)
                label = ["before", "after"] if len(parts) == 2 else ["whole"]
                for seg, lbl in zip(parts, label):
                    clean = sanitize_marathi_text(seg)
                    toks = clean.split()
                    if len(toks) < 2:
                        continue
                    last2 = " ".join(toks[-2:])
                    snippets.append(last2)
                    metadata.append((ti, ri, ci, lbl))
    # vectorize and detect repeats
    vec = CountVectorizer(token_pattern=r'[\u0900-\u097F]+')
    X = vec.fit_transform(snippets).toarray()
    features = vec.get_feature_names_out()
    issues = []
    for i, row in enumerate(X):
        duplicates = [features[j] for j, c in enumerate(row) if c > 1]
        if duplicates:
            ti, ri, ci, lbl = metadata[i]
            issues.append({
                "Table": ti, "Row": ri, "Col": ci, "Segment": lbl,
                "Repeated": duplicates
            })
    return pd.DataFrame(issues)
