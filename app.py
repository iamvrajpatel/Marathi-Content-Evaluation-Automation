import time
import zipfile
from io import BytesIO

import streamlit as st
from config import log
from reporting import write_excel_report

# List of target words to check for presence
CHECK_WORDS = [
    "वयोगट","कालावधी","संकल्पना","संकल्पनात्मक वेळ","संवादाची वेळ",
    "मैदानी खेळ","गणिताची वेळ","भाषेची वेळ","सर्जनशीलतेची वेळ",
    "अ‍ॅनिमेटेड व्हिडिओ/ऑडिओ/ डिजिटल फ्लॅशकार्ड","मुद्रित संसाधने",
    "भौतिक संसाधने","किटमधील साहित्य","शिक्षकांकडून संकलित केलेले साहित्य",
    "अध्यान निष्पत्ती","अध्यान क्षेत्रे","३-४ वर्षे","४-५ वर्षे","५-६ वर्षे",
    "वर्ग कार्यपत्रक","गृहपाठ","आय-फेअर","थीम","समारोपाची वेळ","शिकणाऱ्याला"
]

st.title("Marathi Document Analysis")
st.write("Upload one or more .docx files to run presence checks, spell check, repeated‐phrase detection, term replacements, and grammar review.")

if "started" not in st.session_state:
    st.session_state.started = False
if "finished" not in st.session_state:
    st.session_state.finished = False

if st.button("Reset"):
    st.session_state.started = False
    st.session_state.finished = False

uploads = st.file_uploader("Select .docx files", type="docx", accept_multiple_files=True)
if st.button("Run Analysis"):
    st.session_state.started = True
    st.session_state.finished = False

if st.session_state.started and uploads and not st.session_state.finished:
    num = len(uploads)
    zip_buffer = BytesIO()
    start_t = time.time()

    st.info(f"Processing {num} file(s)...")
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, uploaded_file in enumerate(uploads, start=1):
            log.info(f"File {idx}/{num}: {uploaded_file.name}")
            local_doc = f"temp_{uploaded_file.name}"
            with open(local_doc, "wb") as f:
                f.write(uploaded_file.getvalue())

            local_txt = local_doc + ".txt"
            workbook = write_excel_report(
                base_name=uploaded_file.name.replace(".docx",""),
                docx_path=local_doc,
                txt_path=local_txt,
                words_to_check=CHECK_WORDS
            )
            zf.write(workbook, arcname=workbook)

            # cleanup
            try:
                import os
                os.remove(local_doc)
                os.remove(workbook)
            except OSError:
                pass

            st.progress(idx / num)

    elapsed = time.time() - start_t
    st.success(f"Completed {num} file(s) in {elapsed:.1f}s.")
    zip_buffer.seek(0)
    st.download_button("Download All Reports (ZIP)", zip_buffer, "analysis_reports.zip", "application/zip")

    st.session_state.started = False
    st.session_state.finished = True

elif uploads and not st.session_state.started:
    st.info("Click **Run Analysis** to begin.")
elif st.session_state.finished:
    st.info("Analysis complete! You can re‐run or reset.")
else:
    st.info("Awaiting file uploads…")
