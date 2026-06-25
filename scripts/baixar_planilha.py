"""
baixar_planilha.py
Baixa planilha.xlsx diretamente da API do GitHub (binario puro).
Usa GITHUB_TOKEN automatico do Actions - sem configuracao extra.
"""
import urllib.request, base64, json, os, sys

token = os.environ.get('GH_TOKEN','')
repo  = os.environ.get('REPO','')

if not token or not repo:
    print("Variaveis GH_TOKEN ou REPO nao encontradas")
    sys.exit(1)

url = f'https://api.github.com/repos/{repo}/contents/planilha.xlsx'
req = urllib.request.Request(url, headers={
    'Authorization': f'Bearer {token}',
    'Accept': 'application/vnd.github+json'
})

try:
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())
    content = base64.b64decode(data['content'].replace('\n',''))
    with open('planilha.xlsx','wb') as f:
        f.write(content)
    print(f"planilha.xlsx baixada ({len(content)//1024} KB)")
except Exception as e:
    print(f"Erro: {e}")
    sys.exit(1)
