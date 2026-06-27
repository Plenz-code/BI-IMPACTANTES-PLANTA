"""
update_bi_tudo.py  — versão atualizada
1. Baixa planilha do Dropbox (link publico)
2. Salva como planilha.xlsx no repositório
3. NÃO processa nem embute dados no index.html
   → O index.html processa a planilha direto no browser (lógica por turno, estimativas, DF correta)
   → O GitHub Pages serve o index.html + planilha.xlsx; o browser carrega os dois.

ATENÇÃO: o index.html deve estar atualizado no repositório com o código de processamento
         atual (lógica por turno, prodEfetiva, ULTIMO_TS_DATE etc.).
         Este script não toca no index.html.
"""

import os, sys, zipfile, urllib.request, datetime, warnings
warnings.filterwarnings('ignore')

XLSX_FILE   = 'planilha.xlsx'
DROPBOX_URL = os.environ.get('DROPBOX_URL', '')

# ── 1. BAIXAR DO DROPBOX ──────────────────────────────────────────────────────
if DROPBOX_URL:
    print("Baixando planilha do Dropbox...")
    try:
        req = urllib.request.Request(DROPBOX_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = resp.read()
        if len(data) < 10000:
            raise Exception(f"Arquivo muito pequeno ({len(data)} bytes) — possivelmente erro de link")
        with open(XLSX_FILE, 'wb') as f:
            f.write(data)
        print(f"✓ Planilha baixada do Dropbox ({len(data)//1024} KB)")
    except Exception as e:
        print(f"⚠  Dropbox falhou: {e}")
        if not os.path.exists(XLSX_FILE):
            print("Nenhuma planilha disponível — abortando")
            sys.exit(1)
        print("Usando planilha.xlsx existente no repositório")
else:
    print("DROPBOX_URL não configurado — usando planilha.xlsx do repositório")
    if not os.path.exists(XLSX_FILE):
        print("planilha.xlsx não encontrada — abortando")
        sys.exit(1)

# ── 2. VALIDAR ────────────────────────────────────────────────────────────────
try:
    with zipfile.ZipFile(XLSX_FILE, 'r') as z:
        z.testzip()
    print(f"✓ planilha.xlsx válida ({os.path.getsize(XLSX_FILE)//1024} KB)")
except Exception as e:
    print(f"Arquivo inválido: {e}")
    sys.exit(1)

# ── 3. NADA MAIS A FAZER ──────────────────────────────────────────────────────
# O index.html carrega a planilha.xlsx via fetch() e processa tudo no browser.
# O workflow vai fazer git add planilha.xlsx + index.html e commitar.
print("✓ planilha.xlsx pronta — processamento feito pelo browser via index.html")
