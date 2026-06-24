name: Atualizar BI

on:
  repository_dispatch:
    types: [planilha-atualizada]
  schedule:
    - cron: '0 */2 * * *'
  workflow_dispatch:

jobs:
  update:
    runs-on: ubuntu-latest
    permissions:
      contents: write

    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Configurar Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Instalar dependências
        run: pip install Office365-REST-Python-Client openpyxl

      - name: Baixar planilha do SharePoint
        env:
          SP_USERNAME:  ${{ secrets.SP_USERNAME }}
          SP_PASSWORD:  ${{ secrets.SP_PASSWORD }}
          SP_SITE_URL:  ${{ secrets.SP_SITE_URL }}
          SP_FILE_PATH: ${{ secrets.SP_FILE_PATH }}
        run: python scripts/download_sharepoint.py

      - name: Processar dados e atualizar BI
        run: python scripts/update_bi.py

      - name: Publicar index.html
        run: |
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git config user.name  "github-actions[bot]"
          git add index.html
          if git diff --staged --quiet; then
            echo "Sem mudanças nos dados."
          else
            TIMESTAMP=$(date -u '+%d/%m/%Y %H:%M UTC')
            git commit -m "🔄 BI atualizado — ${TIMESTAMP}"
            git push
            echo "✅ BI publicado."
          fi
