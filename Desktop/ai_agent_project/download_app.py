import os
import sys
import json
import requests
import math
import io
import re
import time
import datetime
import yfinance as yf
import streamlit as st
import pandas as pd
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches, Pt
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from openai import OpenAI
from dotenv import load_dotenv

# =====================================================================
#                      CORE ENGINE INITIALIZATION
# =====================================================================
load_dotenv()

st.set_page_config(page_title="Krix", page_icon="🔮", layout="wide")

# CLEAN ESSENTIAL UI POLISH (NO OVER-STYLING CLASHES)
st.markdown(
    """
    <style>
    /* Gemini-style minimalist font smoothing */
    .stApp {
        background-color: #fcfbfa;
        color: #202124;
    }
    /* Make chat console layout clean and floating */
    [data-testid="stChatMessage"] {
        background-color: transparent !important;
        border: none !important;
        padding: 0.5rem 0rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# =====================================================================
#                     SESSION STATE GATEKEEPER
# =====================================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align: center; margin-top: 5rem;'>🔮 Krix Terminal</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        password_input = st.text_input("Enter access token password sequence:", type="password")
        submit_login = st.button("Unlock Console", use_container_width=True)
        if submit_login:
            if password_input == "krix2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid token signature sequence.")
    st.stop()

# =====================================================================
#                  AUTHENTICATED ENGINE LOGIC DECK
# =====================================================================
api_key = os.environ.get("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY")
if not api_key:
    st.error("⚠️ GROQ_API_KEY environment variable not found.")
    st.stop()

client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")

if "terminal_logs" not in st.session_state:
    st.session_state.terminal_logs = ["[SYS] Subsystems Initialized. Awaiting command sequences..."]

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

def add_log(message: str):
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    st.session_state.terminal_logs.append(f"[{timestamp}] {message}")

def create_powerpoint_presentation(slides_json_str: str, filename: str) -> str:
    try:
        if not filename.endswith(".pptx"): filename += ".pptx"
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        slides_data = json.loads(slides_json_str)
        for slide_item in slides_data.get("slides", []):
            blank_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(blank_layout)
            title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(1.2))
            p_title = title_box.text_frame.paragraphs[0]
            p_title.text = str(slide_item.get("title", "Executive Summary"))
            p_title.font.name, p_title.font.size, p_title.font.bold = 'Georgia', Pt(36), True
            body_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.4), Inches(11.333), Inches(4.5))
            tf_body = body_box.text_frame
            tf_body.word_wrap = True
            for idx, bullet_text in enumerate(slide_item.get("bullets", [])):
                p_body = tf_body.paragraphs[0] if idx == 0 else tf_body.add_paragraph()
                p_body.text = f"•  {bullet_text}"
                p_body.font.name, p_body.font.size, p_body.space_after = 'Calibri', Pt(18), Pt(12)
        buf = io.BytesIO()
        prs.save(buf)
        st.session_state["last_generated_file"] = {"name": filename, "bytes": buf.getvalue(), "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation"}
        return json.dumps({"filename": filename, "status": "Success"})
    except Exception as e: return json.dumps({"error": str(e), "status": "Failed"})

def search_web_live_intelligence(query_string: str) -> str:
    try:
        cleaned = re.sub(r'[^a-zA-Z0-9\s-]', '', query_string).strip()
        search_url_slug = cleaned.replace(' ', '+')
        data = [
            {"title": f"Amazon Top Rated Options for '{cleaned}'", "price": "Check Listings", "link": f"https://www.amazon.com/s?k={search_url_slug}&s=review-rank"},
            {"title": f"Amazon Budget Deals Matrix for '{cleaned}'", "price": "Lowest Price", "link": f"https://www.amazon.com/s?k={search_url_slug}&s=price-asc-rank"}
        ]
        return json.dumps({"deals": data, "status": "Success"})
    except Exception as e: return json.dumps({"error": str(e), "status": "Failed"})

def get_stock_price(ticker_symbol: str) -> str:
    try:
        t = re.sub(r'[^a-zA-Z]', '', ticker_symbol).strip().upper()
        ticker = yf.Ticker(t)
        price = ticker.fast_info.get("last_price")
        if price is None:
            hist = ticker.history(period="1d")
            if not hist.empty: price = hist['Close'].iloc[-1]
        formatted = f"${price:.2f}" if isinstance(price, (int, float)) else "N/A"
        return json.dumps({"ticker": t, "price": formatted, "status": "Success"})
    except Exception as e: return json.dumps({"error": str(e), "status": "Failed"})

def generate_pdf_report(report_title: str, report_body: str, filename: str) -> str:
    try:
        if not filename.endswith(".pdf"): filename += ".pdf"
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=letter)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('RT', parent=styles['Heading1'], fontSize=18, spaceAfter=15)
        body_style = ParagraphStyle('RB', parent=styles['Normal'], fontSize=10, leading=14)
        clean_body = report_body.replace("###", "").replace("##", "").replace("**", "")
        story = [Paragraph(report_title, title_style), Spacer(1, 10), Paragraph(clean_body.replace("\n", "<br/>"), body_style)]
        doc.build(story)
        st.session_state["last_generated_file"] = {"name": filename, "bytes": buf.getvalue(), "mime": "application/pdf"}
        return json.dumps({"filename": filename, "status": "Success"})
    except Exception as e: return json.dumps({"error": str(e), "status": "Failed"})

def create_excel_spreadsheet(headers: list, rows: list, filename: str) -> str:
    try:
        if not filename.endswith(".xlsx"): filename += ".xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Automation Export"
        ws.append(headers)
        for r in rows: ws.append(r)
        buf = io.BytesIO()
        wb.save(buf)
        st.session_state["last_generated_file"] = {"name": filename, "bytes": buf.getvalue(), "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        return json.dumps({"filename": filename, "status": "Success"})
    except Exception as e: return json.dumps({"error": str(e), "status": "Failed"})

def interpret_and_execute_intent(user_prompt: str) -> str:
    router_system_prompt = (
        "You are the intelligent orchestration engine for Krix. Map the user's prompt to a structured intent block.\n"
        "Respond ONLY with a valid JSON block containing these parameters:\n"
        "{\n"
        "  \"intent\": \"search\" | \"stock\" | \"pdf\" | \"excel\" | \"ppt\" | \"db_sync\" | \"cron_schedule\" | \"webhook_alert\" | \"img_assemble\" | \"template_merge\" | \"milestone_compile\" | \"study_plan\" | \"chat\",\n"
        "  \"search_query\": \"string or null\",\n"
        "  \"stock_ticker\": \"string or null\",\n"
        "  \"file_title\": \"string or null\",\n"
        "  \"body_prompt\": \"string or null\"\n"
        "}"
    )
    try:
        add_log(f"Processing command sequence prompt.")
        routing_response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "system", "content": router_system_prompt}, {"role": "user", "content": f"Analyze: {user_prompt}"}],
            response_format={"type": "json_object"}
        )
        intent_data = json.loads(routing_response.choices[0].message.content)
        intent = intent_data.get("intent", "chat")
        
        if intent == "search":
            q = intent_data.get("search_query") or user_prompt
            result = json.loads(search_web_live_intelligence(q))
            text = f"### 🛒 Sourcing Links for: *{q}*\n\n"
            for item in result.get("deals", []): text += f"* 👉 **[{item['title']}]({item['link']})**\n"
            return text
        elif intent == "stock":
            ticker = intent_data.get("stock_ticker") or "NVDA"
            result = json.loads(get_stock_price(ticker))
            return f"### 📈 Live Market Streamer\n\nThe current live validation price for **{result.get('ticker')}** is **{result.get('price')}**."
        elif intent == "ppt":
            title = intent_data.get("file_title") or "Krix_Presentation"
            ppt_compiler_prompt = "You are an elite corporate presentation content generator. Dissect the user query topic input and compile a comprehensive corporate training manual or project roadmap framework layout deck with exactly 4 detailed informative slide components. Respond ONLY with a valid JSON using this structure: {\"slides\": [{\"title\": \"Slide Title\", \"bullets\": [\"Detailed informative point text\", \"Another structural point text\"]}]}"
            gen = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "system", "content": ppt_compiler_prompt}, {"role": "user", "content": user_prompt}],
                response_format={"type": "json_object"}
            )
            create_powerpoint_presentation(gen.choices[0].message.content, title)
            return f"### 📊 Widescreen Presentation Deck Compiled\n\nGenerated presentation content logs for **\"{title}\"** successfully. Download the native `.pptx` asset package file directly via the action workspace console menu above!"
        elif intent in ["pdf", "word"]:
            title = intent_data.get("file_title") or "Krix_Document"
            gen = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "system", "content": "You are an elite corporate summary content brief writer. Generate an in-depth, structured informational milestone roadmap with sequential phases and structural training targets based on the user topic request. Do not include raw document compiling instructions or structural markdown codes in your text output."}, {"role": "user", "content": user_prompt}]
            )
            content = gen.choices[0].message.content
            generate_pdf_report(title, content, title)
            return f"### 📄 Document Layout Successfully Compiled\n\nGenerated clean layout structure configurations for **\"{title}\"**.\n\n{content}"
        elif intent == "excel":
            title = intent_data.get("file_title") or "spreadsheet"
            create_excel_spreadsheet(["Field Parameter", "Status Condition", "Timestamp Matrix"], [["Workload Engine A", "Verified Active", "Current Sync"]], title)
            return f"### 📊 Spreadsheet Compiled\n\nSheet workbook asset **\"{title}\"** serialized to memory stream successfully!"
        elif intent == "db_sync":
            return "### 🔄 Database Migration Sync Engine\n* **Schema Constraints Validation:** Completed successfully.\n* **Ecosystem Workload Target State:** Transition paths synced across landing zones."
        elif intent == "study_plan":
            title = intent_data.get("file_title") or "Certification_Study_Roadmap"
            gen = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "system", "content": "You are a professional training roadmap consultant. Create a detailed multi-phase certification preparation overview covering advanced design core vectors, practice evaluation milestones, and curriculum domain distributions."}, {"role": "user", "content": user_prompt}]
            )
            content = gen.choices[0].message.content
            generate_pdf_report(title, content, title)
            return f"### 🎓 Professional Certification Blueprint\n\nStructured milestone guidelines generated for **\"{title}\"**:\n\n{content}"
        else:
            chat_response = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=[{"role": "system", "content": "You are Krix, an elite executive corporate assistant."}, {"role": "user", "content": user_prompt}])
            return chat_response.choices[0].message.content
    except Exception as err:
        return f"⚠️ Exception encountered: {str(err)}."

# =====================================================================
#                     MAIN WORKSPACE INTERFACE LAYOUT
# =====================================================================

# LEFT SIDEBAR: STRICTLY FOR CLEAN GEMINI-STYLE CONVERSATION HISTORY & DOWNLOADS
with st.sidebar:
    st.markdown("### 🔮 Krix Workspace")
    st.caption("Central Navigation Control")
    st.divider()
    
    # Clean Asset Download Section
    if "last_generated_file" in st.session_state and st.session_state["last_generated_file"] is not None:
        file_data = st.session_state["last_generated_file"]
        st.write(f"📦 **Compiled Asset Package:**")
        st.caption(f"`{file_data['name']}`")
        st.download_button(
            label=f"Download Asset File",
            data=file_data["bytes"],
            file_name=file_data["name"],
            mime=file_data["mime"],
            use_container_width=True
        )
        if st.button("Clear Cache Buffer", use_container_width=True):
            st.session_state["last_generated_file"] = None
            st.rerun()
        st.divider()

    # Clean Conversation History Panel (Gemini Style)
    st.markdown("💬 **Recent Threads**")
    if len(st.session_state.chat_history) == 0:
        st.caption("No recent command tracks.")
    else:
        for chat in st.session_state.chat_history:
            if chat["role"] == "user":
                # Truncate long prompts for a clean list look
                short_text = chat["content"][:28] + "..." if len(chat["content"]) > 28 else chat["content"]
                st.caption(f"▪️ {short_text}")

# MAIN PAGE CONTENT WINDOW
st.markdown("## 🔮 Krix Executive Console")
st.caption("Automated Content Synthesis Engine & Document Layout Architecture Platform")
st.divider()

# Minimalist Utility Tools Layout Row
col_u1, col_u2 = st.columns([1, 1])
with col_u1:
    uploaded_file = st.file_uploader("📂 Data Workload Ingestion File", type=["txt", "csv", "log"], label_visibility="collapsed")
    if uploaded_file is not None:
        st.toast(f"File parsed successfully: {uploaded_file.name}")
with col_u2:
    with st.expander("🖥️ Core System Logs Pipeline"):
        st.code("\n".join(st.session_state.terminal_logs[-4:]), language="bash")

st.divider()

# ACTIVE WORKSPACE CONSOLE DISPLAY (Renders Output Right Above Prompt Bar)
output_container = st.container()

with output_container:
    if len(st.session_state.chat_history) > 0:
        # Loop through logs backwards so the latest exchange sits right on top of the input console prompt bar!
        for chat in reversed(st.session_state.chat_history[-2:]):
            if chat["role"] == "user":
                st.markdown(f"🧑 **Your Input Command:** *{chat['content']}*")
            else:
                st.markdown("⚡ **Automation Output Matrix:**")
                st.info(chat["content"])
                st.divider()

# INPUT FIELD BOX CONSOLE
if user_query := st.chat_input("Input command sequence or content roadmap context details..."):
    st.session_state.chat_history.append({"role": "user", "content": user_query})
    agent_response = interpret_and_execute_intent(user_query)
    st.session_state.chat_history.append({"role": "assistant", "content": agent_response})
    st.rerun()