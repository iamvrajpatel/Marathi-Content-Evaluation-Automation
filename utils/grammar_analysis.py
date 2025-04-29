import json
import openai
from config import API_KEY, log

openai.api_key = API_KEY

SYSTEM_PROMPT = """
आपण एक मराठी व्याकरण तज्ज्ञ आहात. खाली दिलेल्या मजकुरामधील प्रत्येक ब्लॉकचे सूक्ष्म 
निरीक्षण करा आणि व्याकरणाच्या चुका शोधा. त्या चुका ओळखा आणि स्पष्ट करा. 
प्रत्येक ब्लॉकसाठी एक JSON रिपोर्ट तयार करा ज्यात "block_number" आणि 
"grammar_mistakes" ही फील्ड्स असाव्यात. 
"grammar_mistakes" मध्ये प्रत्येक चुकीसाठी वाक्य, चूक कुठे आहे 
(शब्द/वाक्यरचना), आणि योग्य पर्याय यांचा समावेश असावा.

रिपोर्ट फॉर्मॅट:
{
    "block_number": <ब्लॉक क्रमांक>,
    "grammar_mistakes": [
        {
            "sentence": "<चुकीचं वाक्य>",
            "error": "<त्रुटीचे वर्णन>",
            "suggestion": "<सुधारलेलं वाक्य>"
        },
        ...
    ]
}
"""

def analyze_grammar_blocks(
    segments: list[str],
    progress_callback=None
) -> list[dict]:
    """
    Send each text segment to OpenAI for grammar checking.
    Returns list of parsed JSON reports.
    """
    log.info(f"Starting grammar analysis of {len(segments)} segments")
    reports = []
    total = len(segments)

    for idx, seg in enumerate(segments, start=1):
        log.info(f"Analyzing segment {idx}/{total}")
        user_prompt = (
            f"या दस्तऐवजाच्या ब्लॉक क्रमांक {idx} वरील मजकूर खाली दिला आहे:\n\n"
            f"{seg}\n\nकृपया दिलेल्या पृष्ठातील मराठी व्याकरण तपासा."
        )
        resp = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": user_prompt}
            ],
            temperature=0.0,
            max_tokens=1500
        )
        content = resp.choices[0].message.content
        content = content.replace("```", "").strip()
        try:
            report = json.loads(content)
        except json.JSONDecodeError:
            log.error(f"JSON parse error in segment {idx}", exc_info=True)
            report = {"block_number": idx, "grammar_mistakes": [{
                "sentence": "", "error": "Parse error", "suggestion": content
            }]}
        reports.append(report)

        if progress_callback:
            progress_callback(idx / total)

    log.info("Grammar analysis complete")
    return reports
