"""
update_bi_tudo.py  — versao simplificada
1. Baixa planilha do Dropbox (link publico)
2. Valida que e um .xlsx real
3. Salva como planilha.xlsx no repositorio

NAO processa dados nem reescreve o index.html.
O index.html processa a planilha inteira no navegador (fetch + SheetJS),
entao qualquer mudanca de colunas na planilha so precisa ser ajustada
no JavaScript do index.html, nunca aqui.
"""

import os, sys, re, zipfile, urllib.request

XLSX_FILE   = 'planilha.xlsx'
DROPBOX_URL = os.environ.get('DROPBOX_URL', '')


def normaliza_dropbox_url(url):
    """Garante que o link do Dropbox force download direto do arquivo bruto."""
    if not url:
        return url
    if 'dropbox.com' in url and 'dl=' in url:
        url = re.sub(r'dl=0(&|$)', r'dl=1\1', url)
        if 'dl=1' not in url:
            url += ('&' if '?' in url else '?') + 'dl=1'
    elif 'dropbox.com' in url and 'dl=' not in url and 'dropboxusercontent' not in url:
        url += ('&' if '?' in url else '?') + 'dl=1'
    return url


# ── 1. BAIXAR DO DROPBOX ──────────────────────────────────────────────────────
if DROPBOX_URL:
    DROPBOX_URL = normaliza_dropbox_url(DROPBOX_URL)
    print(f"Baixando do Dropbox... ({DROPBOX_URL[:60]}...)")
    try:
        req = urllib.request.Request(DROPBOX_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = resp.read()
            content_type = resp.headers.get('Content-Type', '?')
        print(f"  Content-Type: {content_type} | Tamanho: {len(data)} bytes")

        # xlsx (zip) sempre comeca com os bytes PK (0x50 0x4B)
        if not data.startswith(b'PK'):
            preview = data[:300].decode('utf-8', errors='replace')
            print("Arquivo baixado NAO e um xlsx valido (nao comeca com assinatura ZIP/PK).")
            print(f"Content-Type recebido: {content_type}")
            print(f"Primeiros bytes (preview): {preview}")
            raise Exception(
                "O Dropbox retornou uma pagina/HTML em vez do arquivo. "
                "Verifique se o link termina com '?dl=1' e se ainda e valido "
                "(links podem mudar se a pasta compartilhada for alterada)."
            )
        if len(data) < 30000:
            raise Exception(f"Arquivo muito pequeno ({len(data)} bytes) — pode estar corrompido ou vazio")

        with open(XLSX_FILE, 'wb') as f:
            f.write(data)
        print(f"✓ Planilha baixada do Dropbox ({len(data)//1024} KB)")
    except Exception as e:
        print(f"⚠ Dropbox falhou: {e}")
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
# Um .xlsx valido com dados reais tem, no minimo, algumas dezenas de KB
# (arquivos de ~200-300 bytes passam no teste de ZIP mas estao vazios/quebrados).
TAMANHO_MINIMO_KB = 30

try:
    tamanho_kb = os.path.getsize(XLSX_FILE) / 1024
    with zipfile.ZipFile(XLSX_FILE, 'r') as z:
        z.testzip()
        nomes = z.namelist()
        tem_planilhas = any('xl/worksheets/' in n for n in nomes)

    if tamanho_kb < TAMANHO_MINIMO_KB:
        raise Exception(
            f"Arquivo suspeito: apenas {tamanho_kb:.1f} KB "
            f"(minimo esperado: {TAMANHO_MINIMO_KB} KB). "
            f"Provavelmente veio vazio ou truncado do Dropbox."
        )
    if not tem_planilhas:
        raise Exception("Arquivo e um ZIP valido mas nao contem planilhas (xl/worksheets/ ausente).")

    print(f"✓ planilha.xlsx válida ({tamanho_kb:.0f} KB, {len(nomes)} arquivos internos)")
except Exception as e:
    print(f"Arquivo inválido: {e}")
    sys.exit(1)

# ── 3. NADA MAIS A FAZER ──────────────────────────────────────────────────────
# index.html faz fetch('planilha.xlsx') e processa tudo no navegador.
print("✓ planilha.xlsx pronta — processamento é feito pelo navegador via index.html")
