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
from pydantic import BaseModel, Field
from typing import List, Optional, Literal

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
def extract_pdf_text(uploaded_file) -> str:
    """pdfplumber ile tüm sayfaları metin olarak döndürür."""
    pages = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                pages.append(t)
    return "\n\n--- SAYFA SONU ---\n\n".join(pages)

# ── Gemini Pydantic Şeması (Structured Output için) ────────────────
class FaturaItem(BaseModel):
    tesis: Literal["METYX MNS2 ŞUBESİ", "MANİSA ŞUBESİ"]
    ay: int = Field(description="1-12 arası fatura dönem ayı")
    yil: int = Field(description="4 haneli yıl")
    fatura_tipi: Literal["elektrik", "su", "dogalgaz", "atik", "atiksu", "izin_belgesi", "altyapi", "diger"]
    miktar: Optional[float] = None
    birim: Optional[Literal["KWH", "M3", "ADET"]] = None
    tutar_kdv_dahil: float = Field(description="Ödenecek toplam tutar")

class FaturaResponse(BaseModel):
    faturalar: List[FaturaItem]

# ── Gemini ile fatura ayrıştırma ─────────────────────────────────
SYSTEM_PROMPT = """
Sen bir fatura veri çıkarma uzmanısın. Sana Türkçe OSB (Organize Sanayi Bölgesi) faturaları verilecek.
Fatura tarihi: son ödeme tarihi değil, FATURA TARİHİ'nden ayı belirle.
Mutfak aboneliği MANİSA ŞUBESİ'ne aittir.
MEL fatura numarası = elektrik
MSU fatura numarası = su (şebeke suyu)
MDG fatura numarası = doğalgaz
MKT fatura numarası = evsel katı atık
MAT fatura numarası = atıksu (KOI, AKM, YAĞ&GRES içerir)
GNL/MTA fatura numarası = diğer hizmetler (izin belgesi, altyapı)
"""

def parse_invoices_with_gemini(pdf_text: str) -> list[dict]:
    """Gemini API ile faturaları kesin JSON formatında ayrıştır."""
    # Streamlit Cloud üzerindeki Secrets'ta tanımlı olan GEMINI_API_KEY'i otomatik yakalar.
    client = genai.Client()
    
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=f"Aşağıdaki PDF fatura metinlerini analiz et:\n\n{pdf_text}",
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            response_mime_type="application/json",
            response_schema=FaturaResponse,
            temperature=0.1
        ),
    )
    
    data = json.loads(response.text)
    return data.get("faturalar", [])

# ── Excel doldurma ───────────────────────────────────────────────
def fill_excel(template_bytes: bytes, invoices: list[dict]) -> bytes:
    """Şablonu doldurup dolu Excel döndürür."""
    wb = load_workbook(io.BytesIO(template_bytes))

    for inv in invoices:
        sheet_name = None
        tesis_val = (inv.get("tesis") or "").upper()
        for sn, keywords in SHEETS.items():
            if any(k.upper() in tesis_val for k in keywords):
                sheet_name = sn
                break
        if sheet_name is None or sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        ay_num = inv.get("ay")
        if not ay_num:
            continue
        row = MONTHS_ROW.get(MONTHS_TR.get(ay_num, ""), None)
        if row is None:
            continue

        tip = (inv.get("fatura_tipi") or "").lower()
        miktar = inv.get("miktar")
        tutar = inv.get("tutar_kdv_dahil")

        if tip == "elektrik":
            if miktar:  ws.cell(row=row, column=COL_MAP["elektrik_kwh"]).value = miktar
            if tutar:   ws.cell(row=row, column=COL_MAP["elektrik_tl"]).value  = tutar
        elif tip == "su":
            if miktar:  ws.cell(row=row, column=COL_MAP["su_m3"]).value = miktar
            if tutar:   ws.cell(row=row, column=COL_MAP["su_tl"]).value  = tutar
        elif tip == "dogalgaz":
            prev_kwh = ws.cell(row=row, column=COL_MAP["dogalgaz_kwh"]).value or 0
            prev_tl  = ws.cell(row=row, column=COL_MAP["dogalgaz_tl"]).value  or 0
            if miktar:  ws.cell(row=row, column=COL_MAP["dogalgaz_kwh"]).value = prev_kwh + miktar
            if tutar:   ws.cell(row=row, column=COL_MAP["dogalgaz_tl"]).value  = prev_tl  + tutar
        elif tip == "atik":
            if miktar:  ws.cell(row=row, column=COL_MAP["atik_adet"]).value = miktar
            if tutar:   ws.cell(row=row, column=COL_MAP["atik_tl"]).value   = tutar
        elif tip == "atiksu":
            prev_m3 = ws.cell(row=row, column=COL_MAP["atiksu_m3"]).value or 0
            prev_tl = ws.cell(row=row, column=COL_MAP["atiksu_tl"]).value or 0
            if miktar:  ws.cell(row=row, column=COL_MAP["atiksu_m3"]).value = prev_m3 + miktar
            if tutar:   ws.cell(row=row,
