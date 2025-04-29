import pandas as pd
import docx
from io_helpers import export_text_file, read_docx_text, remove_file
from text_processing import (
    segment_by_sentences, check_word_presence, init_hunspell_dict,
    strip_specials, identify_misspellings, find_term_replacements,
    detect_repeated_phrases
)
from grammar_analysis import analyze_grammar_blocks

def write_excel_report(
    base_name: str,
    docx_path: str,
    txt_path: str,
    words_to_check: list[str]
) -> str:
    """
    Perform all analyses on a single DOCX file,
    write results to an Excel workbook,
    and return the workbook filename.
    """
    # 1) plain text export
    full_text = read_docx_text(docx_path)
    export_text_file(full_text, txt_path)

    # 2) word presence
    presence_df = check_word_presence(words_to_check, txt_path)

    # 3) spell check
    dic = init_hunspell_dict()
    clean = strip_specials(full_text)
    misspelled = identify_misspellings(clean, dic)
    spell_df = pd.DataFrame(misspelled, columns=["Misspelled Word"])

    # 4) repeated-phrase detection
    tbl_df = detect_repeated_phrases(docx.Document(docx_path))

    # 5) replacements
    repl_df = find_term_replacements(full_text)

    # 6) grammar
    segments = segment_by_sentences(full_text)
    gram_reports = analyze_grammar_blocks(segments)
    gram_rows = []
    for rep in gram_reports:
        bn = rep.get("block_number")
        for err in rep.get("grammar_mistakes", []):
            gram_rows.append({
                "Block": bn,
                "Sentence": err["sentence"],
                "Error": err["error"],
                "Suggestion": err["suggestion"]
            })
    gram_df = pd.DataFrame(gram_rows)

    # write to Excel
    out_xl = f"{base_name}_analysis.xlsx"
    with pd.ExcelWriter(out_xl) as writer:
        presence_df.to_excel(writer, sheet_name="Word Presence", index=False)
        spell_df.to_excel(writer, sheet_name="Spell Check", index=False)
        tbl_df.to_excel(writer, sheet_name="Repeated Phrases", index=False)
        repl_df.to_excel(writer, sheet_name="Replacements", index=False)
        gram_df.to_excel(writer, sheet_name="Grammar Check", index=False)

    # cleanup
    remove_file(txt_path)

    return out_xl
