import streamlit as st
import PyPDF2
import pdfplumber
import pandas as pd
import json
import re
import os
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openai import OpenAI

# ==================== PAGE CONFIG ====================
st.set_page_config(
    page_title="Multi-Provider Talent CV Extractor",
    page_icon="📄",
    layout="wide"
)

HISTORY_FILE = "session_history.json"
KEYS_FILE = "saved_api_keys.json"

# ==================== PERSISTENT KEYS MANAGEMENT ====================
def load_saved_keys():
    if os.path.exists(KEYS_FILE):
        try:
            with open(KEYS_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_keys_to_file(keys_dict):
    try:
        with open(KEYS_FILE, "w") as f:
            json.dump(keys_dict, f, indent=2)
    except Exception:
        pass

saved_file_keys = load_saved_keys()
api_providers = ["openrouter", "gemini", "groq", "mistral", "together", "cohere"]

for provider in api_providers:
    session_key = f"{provider}_key"
    param_val = st.query_params.get(session_key, "")
    file_val = saved_file_keys.get(session_key, "")
    if session_key not in st.session_state:
        st.session_state[session_key] = param_val or file_val or ""

# ==================== LOCAL HISTORY MANAGEMENT ====================
def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_history(new_results):
    history = load_history()
    updated = new_results + history
    updated = updated[:10]
    with open(HISTORY_FILE, "w") as f:
        json.dump(updated, f, indent=2)

# ==================== EXCEL FORMATTING FUNCTION ====================
def export_formatted_excel(dataframe, filename="Extracted_Candidates.xlsx"):
    dataframe.to_excel(filename, index=False, engine='openpyxl')
    
    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    
    for col in ws.iter_cols(min_row=1, max_row=1):
        for cell in col:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths = {
        'File Name': 20,
        'Name': 22,
        'Email': 25,
        'Phone': 18,
        'Designation': 25,
        'Current Organization': 25,
        'Previous Organizations': 35,
        'Experience Summary': 30,
        'Education': 35,
        'Source': 22
    }

    for col in ws.columns:
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        col_name = str(col[0].value)
        width = col_widths.get(col_name, 22)
        ws.column_dimensions[col_letter].width = width

        for cell in col[1:]:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(filename)
    return filename

# ==================== NAVIGATION TABS ====================
tab_main, tab_api, tab_history = st.tabs(["🚀 CV Extractor", "⚙️ Persistent API Settings", "📜 Local Session History"])

# ==================== TAB 2: API SETTINGS ====================
with tab_api:
    st.header("🔑 Multi-Provider API Configuration")
    st.caption("Keys entered here are auto-saved locally so you DON'T have to re-enter them on refresh!")

    def update_key(provider_name):
        val = st.session_state[f"input_{provider_name}"]
        st.session_state[f"{provider_name}_key"] = val
        st.query_params[f"{provider_name}_key"] = val
        current_keys = load_saved_keys()
        current_keys[f"{provider_name}_key"] = val
        save_keys_to_file(current_keys)

    col_a, col_b = st.columns(2)
    
    with col_a:
        st.subheader("1. OpenRouter")
        st.markdown("[🔗 Get OpenRouter Key](https://openrouter.ai/keys)")
        st.text_input("OpenRouter Key", value=st.session_state["openrouter_key"], type="password", key="input_openrouter", on_change=update_key, args=("openrouter",))

        st.subheader("2. Google Gemini")
        st.markdown("[🔗 Get Gemini Key](https://aistudio.google.com/)")
        st.text_input("Gemini Key", value=st.session_state["gemini_key"], type="password", key="input_gemini", on_change=update_key, args=("gemini",))

        st.subheader("3. Groq AI")
        st.markdown("[🔗 Get Groq Key](https://console.groq.com/keys)")
        st.text_input("Groq Key", value=st.session_state["groq_key"], type="password", key="input_groq", on_change=update_key, args=("groq",))

    with col_b:
        st.subheader("4. Mistral AI")
        st.markdown("[🔗 Get Mistral Key](https://console.mistral.ai/)")
        st.text_input("Mistral Key", value=st.session_state["mistral_key"], type="password", key="input_mistral", on_change=update_key, args=("mistral",))

        st.subheader("5. Together AI")
        st.markdown("[🔗 Get Together AI Key](https://api.together.ai/)")
        st.text_input("Together AI Key", value=st.session_state["together_key"], type="password", key="input_together", on_change=update_key, args=("together",))

        st.subheader("6. Cohere AI")
        st.markdown("[🔗 Get Cohere Key](https://dashboard.cohere.com/api-keys)")
        st.text_input("Cohere Key", value=st.session_state["cohere_key"], type="password", key="input_cohere", on_change=update_key, args=("cohere",))

    st.markdown("---")
    st.subheader("🧪 Test API Connections")
    
    if st.button("Test Saved API Keys"):
        results_status = []
        if st.session_state["openrouter_key"]:
            try:
                c = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=st.session_state["openrouter_key"])
                c.chat.completions.create(model="openrouter/free", messages=[{"role":"user","content":"hi"}], max_tokens=5)
                results_status.append("✅ OpenRouter: Connected successfully!")
            except Exception as e:
                results_status.append(f"❌ OpenRouter Error: {str(e)[:60]}")

        if st.session_state["groq_key"]:
            try:
                c = OpenAI(base_url="https://api.groq.com/openai/v1", api_key=st.session_state["groq_key"])
                c.chat.completions.create(model="llama-3.3-70b-versatile", messages=[{"role":"user","content":"hi"}], max_tokens=5)
                results_status.append("✅ Groq: Connected successfully!")
            except Exception as e:
                results_status.append(f"❌ Groq Error: {str(e)[:60]}")

        if st.session_state["gemini_key"]:
            try:
                c = OpenAI(base_url="https://generativelanguage.googleapis.com/v1beta/openai/", api_key=st.session_state["gemini_key"])
                c.chat.completions.create(model="gemini-2.0-flash", messages=[{"role":"user","content":"hi"}], max_tokens=5)
                results_status.append("✅ Google Gemini: Connected successfully!")
            except Exception as e:
                results_status.append(f"❌ Gemini Error: {str(e)[:60]}")

        if not results_status:
            st.warning("Please enter at least one API key above.")
        else:
            for status in results_status:
                st.write(status)

# ==================== PDF & AI EXTRACTION ENGINE ====================

def extract_text_from_pdf(file):
    text = ""
    file.seek(0)
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text(layout=True) or page.extract_text()
                if page_text:
                    text += page_text + "\n"
        if text.strip():
            return text.strip(), "pdfplumber"
    except Exception:
        pass
    
    try:
        file.seek(0)
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
        if text.strip():
            return text.strip(), "PyPDF2"
    except Exception:
        pass
    
    return "", "none"

def clean_and_parse_json(text):
    text = text.strip()
    text = re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL).strip()
    try:
        return json.loads(text)
    except Exception:
        pass
    
    match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except Exception:
            pass

    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except Exception:
            pass
            
    return None

def run_ai_extraction(text):
    prompt_system = """You are an expert HR Data Parser. Extract CV details strictly into a valid JSON object.
STRICT EXTRACTION RULES:
- Current Organization: Current active job employer or most recent employer.
- Previous Organizations: List all past companies/institutions separated by pipes (e.g., "Concern Universel | Dhaka Ahsania Mission"). If NO previous company is listed, write "None Listed".
- Education: Capture ALL academic degrees/qualifications including JSC, SSC, HSC, MBBS, Trade Certifications, Madrasah degrees, or School names. (e.g., "JSC from Char Kanto Nogor Dimoki High School"). Do NOT return "Not Found" if ANY school, madrasah, or exam is present.
- Experience Summary: Total years or brief breakdown.
- Keys required: Name, Email, Phone, Designation, Current Organization, Previous Organizations, Experience Summary, Education."""

    user_prompt = f"CV TEXT:\n{text[:6000]}"

    providers = []
    if st.session_state.get("openrouter_key"):
        providers.append(("OpenRouter", "https://openrouter.ai/api/v1", st.session_state["openrouter_key"], "openrouter/free"))
    if st.session_state.get("groq_key"):
        providers.append(("Groq", "https://api.groq.com/openai/v1", st.session_state["groq_key"], "llama-3.3-70b-versatile"))
    if st.session_state.get("gemini_key"):
        providers.append(("Gemini", "https://generativelanguage.googleapis.com/v1beta/openai/", st.session_state["gemini_key"], "gemini-2.0-flash"))
    if st.session_state.get("mistral_key"):
        providers.append(("Mistral", "https://api.mistral.ai/v1", st.session_state["mistral_key"], "mistral-tiny"))

    if not providers:
        return None, "No API keys configured. Please add an API key in the 'Persistent API Settings' tab."

    for provider_name, base_url, key, model in providers:
        try:
            client = OpenAI(base_url=base_url, api_key=key)
            res = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": prompt_system},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.1
            )
            raw = res.choices[0].message.content
            parsed = clean_and_parse_json(raw)
            if parsed:
                parsed["Source"] = f"{provider_name} ({model})"
                return parsed, None
        except Exception:
            continue

    return None, "All configured API providers failed."

# ==================== TAB 1: MAIN EXTRACTOR ====================
with tab_main:
    st.write("Upload PDF CVs below to extract candidate data into Excel.")
    uploaded_files = st.file_uploader("Drop CV PDFs here", type=["pdf"], accept_multiple_files=True)

    if st.button("🚀 Process CVs", type="primary") and uploaded_files:
        results = []
        progress = st.progress(0)
        
        for i, file in enumerate(uploaded_files):
            raw_text, extractor_type = extract_text_from_pdf(file)
            
            if not raw_text.strip():
                st.error(f"Could not read {file.name}")
                continue
            
            data, err = run_ai_extraction(raw_text)
            
            if data:
                data["File Name"] = file.name
                results.append(data)
                st.success(f"Successfully processed {file.name}")
            else:
                st.error(f"Extraction failed for {file.name}: {err}")
            
            progress.progress((i + 1) / len(uploaded_files))

        if results:
            st.balloons()
            save_history(results)
            
            df = pd.DataFrame(results)
            expected_cols = ['File Name', 'Name', 'Email', 'Phone', 'Designation', 'Current Organization', 'Previous Organizations', 'Experience Summary', 'Education', 'Source']
            df = df.reindex(columns=[c for c in expected_cols if c in df.columns])
            
            st.dataframe(df, use_container_width=True)
            
            # --- COPY TO CLIPBOARD CODE ---
            tsv_data = df.to_csv(sep="\t", index=False)
            st.subheader("📋 Copy Results directly")
            st.code(tsv_data, language="text")
            st.caption("Tip: Click the copy icon in the top right of the text box above to copy data directly into Excel or Google Sheets!")

            excel_file = export_formatted_excel(df)
            with open(excel_file, "rb") as f:
                st.download_button(
                    label="📥 Download Structured Excel File",
                    data=f.read(),
                    file_name="Candidates_Summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ==================== TAB 3: LOCAL HISTORY ====================
with tab_history:
    st.header("📜 Session History (Last 10 Extractions)")
    history_data = load_history()
    
    if history_data:
        df_hist = pd.DataFrame(history_data)
        st.dataframe(df_hist, use_container_width=True)
        
        tsv_hist = df_hist.to_csv(sep="\t", index=False)
        st.code(tsv_hist, language="text")
        
        if st.button("Clear Saved History"):
            if os.path.exists(HISTORY_FILE):
                os.remove(HISTORY_FILE)
                st.rerun()
    else:
        st.info("No saved local history found yet.")
