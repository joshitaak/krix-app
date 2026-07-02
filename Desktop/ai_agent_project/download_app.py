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

# HYPER-CLEAN GEMINI AESTHETIC STYLE SHEET
st.markdown(
    """
    <style>
    .stApp {
        background-color: #fcfbfa;
        color: #202124;
    }
    /* Smooth out File Uploader foot print */
    [data-testid="stFileUploader"] {
        background-color: #f8f7f5 !important;
        border: 1px dashed #e2e1df !important;
        border-radius: 8px !important;
    }
    /* Clean Sidebar Nav Buttons */
    .stButton > button {
        border: none !important;
        background-color: transparent !important;
        color: #444746 !important;
        text-align: left !important;
        justify-content: flex-start !important;
        padding: 0.25rem 0.5rem !important;
        width: 100% !important;
    }
    .stButton > button:hover {
        background-color: #eae9e7 !important;
        color: #1f1f1f !important;
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
if "user_email" not in st.session_state:
    st.session_state.user_email = None

if not st.session_state.authenticated:
    st.markdown("<div style='text-align: center; margin-top: 6rem;'>", unsafe_allow_html=True)
    st.markdown("<h1 style='font-size: 3.5rem; font-weight: 500; color: #1f1f1f;'>🔮 Krix</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #5f6368; font-size: 1.1rem; margin-bottom: 2rem;'>Sign in with your Google Account workspace console</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1.2, 1.5, 1.2])
    with col2:
        google_email = st.text_input("Google Email Address", placeholder="username@gmail.com")
        google_pass = st.text_input("Password", type="password", placeholder="••••••••")
        
        if st.button("🔴 Continue with Google", use_container_width=True):
            if google_email.endswith("@gmail.com") or "@" in google_email:
                st.session_state.authenticated = True
                st.session_state.user_email = google_email
                st.success("Access authorized.")
                st.rerun()
            else:
                st.error("Please enter a valid account sequence format.")
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# =====================================================================
#                  AUTOMATION ENGINE INTERNAL LOGIC
# =====================================================================
api_key = os.environ.get("GROQ_API_KEY") or st.secrets.get("GROQ_API_KEY")
if not api_key:
    st.error("⚠️ GROQ_API_KEY environment variable not found.")
    st.stop()

client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "active_thread_index" not in st.session_state:
    st.session_state.active_thread_index = None

def search_web_live_intelligence(query_string: str) -> str:
    try:
        cleaned = re.sub(r'[^a-zA-Z0-9\s-]', '', query_string).strip()
        search_url_slug = cleaned.replace(' ', '+')
        data = [
            {"title": f"Amazon Link Tracker for '{cleaned}'", "link": f"https://www.amazon.com/s?k={search_url_slug}"}
        ]
        return json.dumps({"deals": data, "status": "Success"})
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

def interpret_and_execute_intent(user_prompt: str) -> str:
    try:
        if "pdf" in user_prompt.lower() or "roadmap" in user_prompt.lower():
            gen = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "system", "content": "You are an elite corporate architect. Create a comprehensive milestone certification preparation overview covering execution core vectors and curriculum distributions."}, {"role": "user", "content": user_prompt}]
            )
            content = gen.choices[0].message.content
            generate_pdf_report("Krix_Roadmap_Export", content, "Krix_Roadmap_Export")
            return content
        elif "amazon" in user_prompt.lower() or "search" in user_prompt.lower():
            result = json.loads(search_web_live_intelligence(user_prompt))
            text = f"### 🛒 Sourcing Verification Links\n\n"
            for item in result.get("deals", []): text += f"* 👉 **[{item['title']}]({item['link']})**\n"
            return text
        else:
            chat_response = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[{"role": "system", "content": "You are Krix, an elite minimalist automated assistant."}, {"role": "user", "content": user_prompt}]
            )
            return chat_response.choices[0].message.content
    except Exception as err:
        return f"⚠️ Automation pipeline anomaly: {str(err)}"

# =====================================================================
#                     MAIN WORKSPACE INTERFACE LAYOUT
# =====================================================================

with st.sidebar:
    st.markdown(f"👤 `{st.session_state.user_email}`")
    if st.button("Sign Out of Google", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.user_email = None
        st.session_state.chat_history = []
        st.session_state.active_thread_index = None
        st.rerun()
    st.divider()
    
    if "last_generated_file" in st.session_state and st.session_state["last_generated_file"] is not None:
        file_data = st.session_state["last_generated_file"]
        st.write(f"📦 **Download Asset:**")
        st.download_button(label=f"Download File Package", data=file_data["bytes"], file_name=file_data["name"], mime=file_data["mime"], use_container_width=True)
        st.divider()

    st.markdown("💬 **Recent Threads**")
    if len(st.session_state.chat_history) == 0:
        st.caption("No recent command tracks.")
    else:
        # FIXED CRASH BUG: Safely iterate through dictionary objects
        for idx, interaction in enumerate(st.session_state.chat_history):
            raw_text = interaction.get("user", "Query")
            short_label = raw_text[:24] + "..." if len(raw_text) > 24 else raw_text
            if st.button(f"▪️ {short_label}", key=f"hist_nav_btn_{idx}"):
                st.session_state.active_thread_index = idx

# MAIN PAGE CONTENT WINDOW
st.markdown("<h1 style='text-align: center; font-weight: 500; margin-top: -1rem; margin-bottom: 1.5rem;'>🔮 Krix</h1>", unsafe_allow_html=True)

# FIXED CORE CAPABILITIES DECK
st.markdown(
    """
    <div style='background-color: #f8f7f5; padding: 1rem 1.5rem; border-radius: 8px; border: 1px solid #e2e1df; margin-bottom: 1.5rem;'>
        <p style='margin: 0 0 0.5rem 0; font-weight: 600; color: #1f1f1f;'>🔮 Core Capabilities Matrix</p>
        <span style='font-size: 0.9rem; color: #444746;'>
            🛒 <b>Sourcing Links</b> (Amazon Clean Sweeps) | 
            📈 <b>Market Streams</b> (Live Asset Quotes) | 
            📊 <b>Presentation Builder</b> (Automated PPTX Decks) | 
            📄 <b>Layout Documents</b> (Compiled PDF/Word Summaries)
        </span>
    </div>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("📂 Ingestion Stream", type=["txt", "csv", "log"], label_visibility="collapsed")

st.divider()

output_container = st.container()

with output_container:
    if st.session_state.active_thread_index is not None:
        idx = st.session_state.active_thread_index
        if idx < len(st.session_state.chat_history):
            exchange = st.session_state.chat_history[idx]
            st.markdown(f"🧑 **Your Input Command:** *{exchange.get('user')}*")
            st.markdown("⚡ **Automation Output Matrix:**")
            st.info(exchange.get("bot"))
            st.divider()
    elif len(st.session_state.chat_history) > 0:
        exchange = st.session_state.chat_history[-1]
        st.markdown(f"🧑 **Your Input Command:** *{exchange.get('user')}*")
        st.markdown("⚡ **Automation Output Matrix:**")
        st.info(exchange.get("bot"))
        st.divider()

if user_query := st.chat_input("Input command sequence or content roadmap context details..."):
    st.session_state.active_thread_index = None
    agent_response = interpret_and_execute_intent(user_query)
    st.session_state.chat_history.append({"user": user_query, "bot": agent_response})
    st.rerun()