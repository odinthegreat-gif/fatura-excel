# 📊 Fatura → Excel Dönüştürücü

MANİSA OSB faturalarını otomatik olarak Excel'e aktaran Streamlit uygulaması.

## Kurulum

```bash
pip install -r requirements.txt
```

## Çalıştırma

```bash
streamlit run app.py
```

Tarayıcıda http://localhost:8501 adresine gidin.

## Kullanım

1. **PDF yükle** — Bir veya birden fazla OSB faturası seçin  
2. **Excel şablonu yükle** — `2026_Yılı_Tesis_Verileri_Takip_Tablosu.xlsx` şablonunu yükleyin  
3. **"Faturları İşle" butonuna tıkla** — Claude AI faturları okur ve Excel'e aktarır  
4. **Excel dosyasını indir** — Doldurulmuş dosyayı indirin  

## Desteklenen Fatura Türleri

| Fatura Kodu | Tür           |
|-------------|---------------|
| MEL         | Elektrik      |
| MSU         | Su            |
| MDG         | Doğalgaz      |
| MKT         | Katı Atık     |
| MAT         | Atıksu        |
| GNL / MTA   | Diğer         |

## Gereksinimler

- Python 3.10+
- `ANTHROPIC_API_KEY` ortam değişkeni (Claude API anahtarı)

```bash
export ANTHROPIC_API_KEY="sk-ant-..."
streamlit run app.py
```
