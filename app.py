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
    # Prepend new entries and keep only the last 10 sessions/items
    updated = new_results + history
    updated = updated[:10]
    with open(HISTORY_FILE, "w") as f:
        json.dump(updated, f, indent=2)

# ==================== EXCEL FORMATTING FUNCTION ====================
def export_formatted_excel(dataframe, filename="Extracted_Candidates.xlsx"):
    """Saves DataFrame with readable column widths and wrapped text formatting."""
    dataframe.to_excel(filename, index=False, engine='openpyxl')
    
    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    # Header Style
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    
    for col in ws.iter_cols(min_row=1, max_row=1):
        for cell in col:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Column dimensions & text wrapping
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
        'Source': 18
    }

    for col in ws.columns:
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        col_name = str(col[0].value)
        width = col_widths.get(col_name, 22)
        ws.column_dimensions[col_letter].width = width

        # Apply wrap text to body cells
        for cell in col[1:]:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(filename)
    return filename

# ==================== NAVIGATION TABS ====================
tab_main, tab_api, tab_history = st.tabs(["🚀 CV Extractor", "⚙️ API & Model Settings", "📜 Local Session History"])

# ==================== TAB 2: API SETTINGS ====================
with tab_api:
    st.header("🔑 Multi-Provider API Configuration")
    st.caption("Keys are stored locally in your current browser session.")

    col_a, col_b = st.columns(2)
    
    with col_a:
        st.subheader("OpenRouter")
        st.markdown("[🔗 Get OpenRouter Key](https://openrouter.ai/keys)")
        openrouter_key = st.text_input("OpenRouter API Key", type="password", key="openrouter_key")

        st.subheader("Google Gemini")
        st.markdown("[🔗 Get Google Gemini Key](https://aistudio.google.com/)")
        gemini_key = st.text_input("Gemini API Key", type="password", key="gemini_key")

    with col_b:
        st.subheader("Groq (Ultra Fast)")
        st.markdown("[🔗 Get Groq Key](https://console.groq.com/keys)")
        groq_key = st.text_input("Groq API Key", type="password", key="groq_key")

        st.subheader("Mistral AI")
        st.markdown("[🔗 Get Mistral Key](https://console.mistral.ai/)")
        mistral_key = st.text_input("Mistral API Key", type="password", key="mistral_key")

    st.markdown("---")
    st.subheader("🧪 Test API Connections")
    
    if st.button("Test Connected APIs"):
        results_status = []
        
        # Test OpenRouter
        if openrouter_key:
            try:
                c = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=openrouter_key)
                c.chat.completions.create(model="openrouter/free", messages=[{"role":"user","content":"hi"}], max_tokens=5)
                results_status.append("✅ OpenRouter: Connected successfully!")
            except Exception as e:
                results_status.append(f"❌ OpenRouter Error: {str(e)[:60]}")

        # Test Groq
        if groq_key:
            try:
                c = OpenAI(base_url="https://api.groq.com/openai/v1", api_key=groq_key)
                c.chat.completions.create(model="llama-3.3-70b-versatile", messages=[{"role":"user","content":"hi"}], max_tokens=5)
                results_status.append("✅ Groq: Connected successfully!")
            except Exception as e:
                results_status.append(f"❌ Groq Error: {str(e)[:60]}")

        # Test Gemini
        if gemini_key:
            try:
                c = OpenAI(base_url="https://generativelanguage.googleapis.com/v1beta/openai/", api_key=gemini_key)
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
Rules:
- Current Organization: Current active job employer or most recent employer.
- Previous Organizations: List all past companies/institutions separated by commas or pipes (e.g., "BRAC | Khulna Medical College | Afia Clinic").
- Education: Degrees, Majors, and Institutes (e.g., "MBBS from Khulna Medical College, HSC from Govt. BL College"). Do NOT return "Not Found" if education or degrees exist in the text.
- Experience Summary: Total years or brief breakdown.
- Keys required: Name, Email, Phone, Designation, Current Organization, Previous Organizations, Experience Summary, Education."""

    user_prompt = f"CV TEXT:\n{text[:6000]}"

    # Auto-Select Active API Client
    providers = []
    if openrouter_key:
        providers.append(("OpenRouter", "https://openrouter.ai/api/v1", openrouter_key, "openrouter/free"))
    if groq_key:
        providers.append(("Groq", "https://api.groq.com/openai/v1", groq_key, "llama-3.3-70b-versatile"))
    if gemini_key:
        providers.append(("Gemini", "https://generativelanguage.googleapis.com/v1beta/openai/", gemini_key, "gemini-2.0-flash"))

    if not providers:
        return None, "No API key configured in Settings tab."

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
        except Exception as e:
            continue

    return None, "All active API providers failed."

# ==================== TAB 1: MAIN EXTRACTOR ====================
with tab_main:
    st.write("Upload PDF CVs below to extract formatted candidate data into Excel.")
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
        
        if st.button("Clear Saved History"):
            if os.path.exists(HISTORY_FILE):
                os.remove(HISTORY_FILE)
                st.rerun()
    else:
        st.info("No saved local history found yet.")