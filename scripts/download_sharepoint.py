"""
download_sharepoint.py
Baixa a planilha diretamente do SharePoint usando as
credenciais do usuário (email + senha do Microsoft 365).
Sem Azure AD, sem app registration, sem Premium.

Secrets necessários no GitHub:
  SP_USERNAME  → email Microsoft 365 (ex: dsf.cco01@harpiagold.com.br)
  SP_PASSWORD  → senha do Microsoft 365
  SP_SITE_URL  → https://harpiagold.sharepoint.com/sites/GERENCIAMENTODEMINAEPLANTA
  SP_FILE_PATH → /sites/GERENCIAMENTODEMINAEPLANTA/Documentos/02. Planta/2026/Junho/Planilha_produção_jun_26_otimizada (25).xlsx
"""

import os, sys

try:
    from office365.runtime.auth.user_credential import UserCredential
    from office365.sharepoint.client_context import ClientContext
except ImportError:
    print("Instalando dependências...")
    os.system("pip install Office365-REST-Python-Client --quiet")
    from office365.runtime.auth.user_credential import UserCredential
    from office365.sharepoint.client_context import ClientContext

USERNAME  = os.environ['SP_USERNAME']   # dsf.cco01@harpiagold.com.br
PASSWORD  = os.environ['SP_PASSWORD']   # senha do Microsoft 365
SITE_URL  = os.environ['SP_SITE_URL']   # https://harpiagold.sharepoint.com/sites/...
FILE_PATH = os.environ['SP_FILE_PATH']  # /sites/.../Junho/Planilha_...xlsx

print(f"Conectando ao SharePoint como {USERNAME}...")

try:
    ctx = ClientContext(SITE_URL).with_credentials(
        UserCredential(USERNAME, PASSWORD)
    )
    print("✓ Autenticado")

    print(f"Baixando: {FILE_PATH.split('/')[-1]}")
    with open("planilha.xlsx", "wb") as f:
        ctx.web.get_file_by_server_relative_url(FILE_PATH).download(f).execute_query()

    size = os.path.getsize("planilha.xlsx")
    print(f"✓ Planilha baixada ({size // 1024} KB)")

except Exception as e:
    print(f"✗ Erro: {e}")
    print("\nPossíveis causas:")
    print("  - Senha incorreta ou conta com MFA ativado")
    print("  - Caminho do arquivo incorreto")
    print("  - Sem permissão de acesso à pasta")
    sys.exit(1)
