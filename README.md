# Marathi Content Evaluation Automation

This project provides an automated pipeline for evaluating Marathi language `.docx` documents. It performs word presence checks, spell checking, repeated word detection, word replacements, and grammar analysis using OpenAI's GPT models. The results are compiled into Excel reports for each document.

---

## Features

- **Word Presence Check:** Verifies the presence of specific key words/phrases in the document.
- **Spell Check:** Detects misspelled Marathi words using a Hunspell dictionary.
- **Repeated Words Detection:** Identifies repeated word patterns in tables.
- **Word Replacement Suggestions:** Suggests replacements for certain words/phrases based on a predefined dictionary.
- **Grammar Check:** Uses OpenAI GPT to analyze blocks of text (~15 sentences per block) for grammar mistakes and suggests corrections.
- **Batch Processing:** Supports multiple `.docx` files at once.
- **Excel Report Generation:** Outputs results in structured Excel files, one per input document.
- **ZIP Download:** All reports are packaged into a single downloadable ZIP file.

---

## Input

- **File Type:** `.docx` (Microsoft Word) files.
- **Upload Method:** Via the Streamlit web interface (supports multiple files at once).

---

## Output

- **Excel Reports:** For each input file, an Excel file (`*_analysis.xlsx`) is generated with the following sheets:
  - `Word Presence`: Status of key words/phrases.
  - `Spell Check`: List of misspelled words.
  - `Repeated Words`: Contexts where repeated words are detected.
  - `Replacements`: Words/phrases found that have suggested replacements.
  - `Grammar Check`: Detailed grammar mistakes and suggestions per block.
- **ZIP Archive:** All Excel reports are zipped for easy download.

---

## Processing Pipeline

1. **File Upload:** User uploads one or more `.docx` files via the Streamlit interface.
2. **Text Extraction:** Text is extracted from paragraphs and tables, preserving logical blocks.
3. **Word Presence:** Checks for a predefined list of important words/phrases.
4. **Spell Checking:** Uses a Hunspell-based Marathi dictionary to find misspelled words.
5. **Repeated Words:** Scans tables for repeated word patterns, especially in the last two words of each cell segment.
6. **Replacement Suggestions:** Checks for words/phrases that should be replaced and suggests alternatives.
7. **Grammar Analysis:** Splits the document into blocks (~15 sentences each) and sends each block to OpenAI GPT for grammar checking. Results are parsed and structured.
8. **Report Generation:** All results are saved into Excel files, one per input document.
9. **Packaging:** All Excel files are zipped and made available for download.

---

## How to Use

1. **Install Requirements:**
   ```
   pip install -r requirement.txt
   ```
2. **Set OpenAI API Key:**
   - Edit `main_with_session.py` and set your OpenAI API key in the `OPENAI_API_KEY` variable.

3. **Run the App:**
   ```
   streamlit run main_with_session.py
   ```

4. **Upload Files:**
   - Use the web interface to upload `.docx` files.

5. **Start Analysis:**
   - Click "Start Analysis" to begin processing.
   - Progress bars and status messages will be shown.

6. **Download Reports:**
   - Once processing is complete, download the ZIP file containing all Excel reports.

---

## Notes

- **Internet Required:** Grammar checking uses OpenAI's API and requires an internet connection.
- **API Costs:** Each grammar check block uses OpenAI tokens and may incur costs.
- **Dictionary Files:** The Hunspell dictionary files (`mr_IN.dic`, `mr_IN.aff`) must be present in the project directory for spell checking to work.
- **Logging:** Processing logs are saved to `analysis.log` for troubleshooting.

---

## Customization

- **Word List:** Edit the `words` list in `main_with_session.py` to change which words/phrases are checked for presence.
- **Replacement Dictionary:** Update `replacement_dict` in `main_with_session.py` to add or modify word/phrase replacements.
- **Grammar Model:** The OpenAI model used can be changed in the code if needed.

---

## Dependencies

See `requirement.txt` for all Python dependencies.

---

## License

This project is for educational and internal use. Please ensure you comply with OpenAI's terms of service and licensing for dictionary data.

