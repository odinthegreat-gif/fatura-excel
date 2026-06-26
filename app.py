import streamlit as st
import pdfplumber
import re
import io
import copy
import json
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl import load_workbook
from google import genai
from google.genai import types
import pandas as pd

# ── Sayfa ayarları ──────────────────────────────────────────────
st.set_page_config(
    page_title="Fatura → Excel Dönüştürücü",
    page_icon="📊",
    layout="wide",
)

# ── CSS ─────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 100%);
        padding: 20px 30px; border-radius: 12px; margin-bottom: 25px;
        color: white;
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; }
    .main-header p  { margin: 5px 0 0; opacity: 0.85; font-size: 0.95rem; }
    .metric-card {
        background: #f8f9fa; border-left: 4px solid #2d6a9f;
        padding: 14px 18px; border-radius: 8px; margin: 6px 0;
    }
    .metric-card .label { font-size: 0.78rem; color: #666; font-weight: 600; text-transform: uppercase; }
    .metric-card .value { font-size: 1.25rem; font-weight: 700; color: #1e3a5f; }
    .success-box {
        background: #d4edda; border: 1px solid #c3e6cb;
        padding: 12px 16px; border-radius: 8px; color: #155724;
    }
    .warning-box {
        background: #fff3cd; border: 1px solid #ffc107;
        padding: 12px 16px; border-radius: 8px; color: #856404;
    }
    .stProgress > div > div { background-color: #2d6a9f !important; }
    div[data-testid="stFileUploader"] {
        border: 2px dashed #2d6a9f !important;
        border-radius: 10px !important; padding: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>📊 Fatura → Excel Dönüştürücü (Gemini)</h1>
  <p>PDF faturalarınızı yükleyin, yapay zeka otomatik olarak verileri okuyup Excel tablosuna aktarsın.</p>
</div>
""", unsafe_allow_html=True)

# ── Sabit veriler ────────────────────────────────────────────────
MONTHS_TR = {
    1:"OCAK", 2:"ŞUBAT", 3:"MART", 4:"NİSAN",
    5:"MAYIS", 6:"HAZİRAN", 7:"TEMMUZ", 8:"AĞUSTOS",
    9:"EYLÜL", 10:"EKİM", 11:"KASIM", 12:"ARALIK",
}
MONTHS_ROW = {v: 4+i for i, v in enumerate(MONTHS_TR.values())}  # OCAK→4 … ARALIK→15

SHEETS = {
    "METYX MNS2 ŞUBESİ": ["MNS2", "METYX MNS2"],
    "MANİSA ŞUBESİ":      ["MANİSA ŞUBESİ", "MANİSA SUBESI", "MUTFAK"],
}

# Excel sütun haritası (A=1 … M=13)
COL_MAP = {
    "elektrik_kwh":   4,   # D
    "elektrik_tl":    5,   # E
    "su_m3":          6,   # F
    "su_tl":          7,   # G
    "dogalgaz_kwh":   8,   # H
    "dogalgaz_tl":    9,   # I
    "atik_adet":     10,   # J
    "atik_tl":       11,   # K
    "atiksu_m3":     12,   # L
    "atiksu_tl":     13,   # M
}

# ── PDF okuma yardımcıları ───────────────────────────────────────
def extract_pdf_text(uploaded_file)
