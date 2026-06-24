"""
update_bi.py — versão completa sem Azure
Lê planilha.xlsx do repositório e atualiza o index.html.
"""

import json, re, warnings, datetime, io
warnings.filterwarnings('ignore')
from openpyxl import load_workbook

XLSX_FILE = 'planilha.xlsx'
HTML_FILE = 'index.html'
ABA_MAP   = {'Britador': 'brit', 'Linha 1': 'l1', 'Linha 2': 'l2', 'Linha 3': 'l3'}

def fmt_date(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        d = v.date() if isinstance(v, datetime.datetime) else v
        if 2000 <= d.year <= 2100:
            return d.strftime('%Y-%m-%d')
    return None

def to_mins(v):
    if v is None: return 0
    if isinstance(v, datetime.timedelta): return round(v.total_seconds() / 60)
    if isinstance(v, datetime.time):      return v.hour * 60 + v.minute
    if isinstance(v, datetime.datetime):
        if v.year < 2000: return v.hour * 60 + v.minute
        return round((v.timestamp() % 86400) / 60)
    if isinstance(v, float):
        return round((v - int(v)) * 24 * 60)
    return 0

def clean(v):
    return str(v).replace('\xa0', ' ').strip() if v else ''

def process():
    wb = load_workbook(XLSX_FILE, data_only=True)
    print(f"✓ Planilha carregada: {XLSX_FILE}")

    _eventos = {k: [] for k in ABA_MAP.values()}
    _has     = {k: {} for k in ABA_MAP.values()}
    _parada  = {k: {} for k in ABA_MAP.values()}
    balanca  = {k: {} for k in ABA_MAP.values()}
    prod_linha   = {k: {} for k in ABA_MAP.values()}
    material_tot = {k: {} for k in ABA_MAP.values()}

    for sheet, key in ABA_MAP.items():
        if sheet not in wb.sheetnames:
            print(f"  Aba '{sheet}' não encontrada, pulando.")
            continue
        ws = wb[sheet]
        ultima_data = None
        turno_soma  = 0

        for row in ws.iter_rows(min_row=2, values_only=True):
            c0   = row[0]
            data = fmt_date(c0)

            if data:
                ultima_data = data
                tipo  = clean(row[6])
                equip = clean(row[7])
                desc  = clean(row[8])
                hi    = row[9]
                tp    = row[11]
                conch = row[4] if isinstance(row[4], (int, float)) else 0
                prod  = row[5] if isinstance(row[5], (int, float)) else 0
                mat   = clean(row[3])
                t_mins = to_mins(tp)

                if conch > 0 or prod > 0 or tipo or t_mins > 0:
                    _has[key][data] = True
                _parada[key].setdefault(data, 0)
                _parada[key][data] += t_mins

                if t_mins > 0 and hi is not None and (tipo or equip or desc):
                    _eventos[key].append((data, tipo, equip, desc, t_mins))

                prod_linha[key].setdefault(data, {'conch': 0, 'prod': 0})
                prod_linha[key][data]['conch'] += conch
                prod_linha[key][data]['prod']  += prod
                if mat and conch > 0:
                    material_tot[key][mat] = material_tot[key].get(mat, 0) + conch

            elif isinstance(c0, str):
                cu  = c0.upper()
                bal = row[13] if (len(row) > 13 and isinstance(row[13], (int, float))) else 0
                if 'TURNO' in cu:
                    turno_soma += bal
                elif 'TOTAL GLOBAL' in cu:
                    if ultima_data:
                        valor = bal if bal > 0 else turno_soma
                        if valor > 0:
                            balanca[key][ultima_data] = valor
                    turno_soma = 0

    if 'Resumo' in wb.sheetnames:
        for row in wb['Resumo'].iter_rows(min_row=4, max_row=33, values_only=True):
            dr = row[1]
            if not isinstance(dr, (datetime.datetime, datetime.date)): continue
            d  = dr.date() if isinstance(dr, datetime.datetime) else dr
            dt = d.strftime('%Y-%m-%d')
            for i, k in enumerate(['brit', 'l1', 'l2', 'l3'], start=2):
                v = row[i]
                if isinstance(v, (int, float)) and v > 0 and dt not in balanca[k]:
                    balanca[k][dt] = float(v)
    wb.close()

    def build_daily(pm, hm):
        out = {}
        for d, p in sorted(pm.items()):
            if not hm.get(d): continue
            hs = round(min(24, p / 60), 1)
            out[d] = {'prod': round(max(0, 24 - hs), 1), 'parada': hs}
        return out

    def build_mots(evs):
        bd = {}
        for data, tipo, equip, desc, mins in evs:
            bd.setdefault(data, [])
            bd[data].append([f"{tipo} \u2225 {equip} \u2225 {desc}", round(mins / 60, 2)])
        for d in bd:
            bd[d].sort(key=lambda x: -x[1])
        return bd

    def build_rkpi(daily):
        return {d: {'hp': v['prod'], 'hs': v['parada'],
                    'df': round(v['prod'] / (v['prod'] + v['parada']) * 100, 1)
                    if (v['prod'] + v['parada']) > 0 else 0}
                for d, v in daily.items()}

    results = {}
    for key in ABA_MAP.values():
        daily = build_daily(_parada[key], _has[key])
        mots  = build_mots(_eventos[key])
        results[key] = {'mots': mots, 'rkpi': build_rkpi(daily)}
        n = sum(len(v) for v in mots.values())
        print(f"  {key}: {len(daily)} dias, {n} ocorrências")

    todas_datas = sorted(set(d for k in balanca for d in balanca[k]))
    resumo = [{'data': dt, 'dia': int(dt[8:10]),
               'brit': balanca['brit'].get(dt, 0),
               'l1':   balanca['l1'].get(dt, 0),
               'l2':   balanca['l2'].get(dt, 0),
               'l3':   balanca['l3'].get(dt, 0),
               'total': balanca['l1'].get(dt, 0) + balanca['l2'].get(dt, 0) + balanca['l3'].get(dt, 0)}
              for dt in todas_datas]

    prod_planta = {'resumo': resumo, 'prod_linha': prod_linha, 'material_tot': material_tot}
    return results, prod_planta

def inject_auto_init(html):
    AUTO_INIT = """<!-- AUTO_INIT_START -->
<script>
document.addEventListener('DOMContentLoaded',function(){
  if(!R_KPI||!R_KPI.l1||!Object.keys(R_KPI.l1).length)return;
  ['brit','l1','l2','l3'].forEach(function(k){
    DADOS[k+'_daily']={};
    Object.entries(R_KPI[k]||{}).forEach(function(e){DADOS[k+'_daily'][e[0]]={prod:e[1].hp,parada:e[1].hs};});
    var acc={};
    Object.values(R_MOT[k]||{}).forEach(function(l){(l||[]).forEach(function(i){acc[i[0]]=(acc[i[0]]||0)+i[1];});});
    DADOS[k].top_motivos=Object.entries(acc).map(function(e){return[e[0],e[1]];}).sort(function(a,b){return b[1]-a[1];}).slice(0,14);
  });
  setTimeout(function(){var b=document.getElementById('btnDash');if(b&&typeof goTab==='function')goTab('dashboard',b);},150);
});
</script>
<!-- AUTO_INIT_END -->"""
    html = re.sub(r'<!-- AUTO_INIT_START -->[\s\S]*?<!-- AUTO_INIT_END -->', '', html)
    html = html.replace('</body>', AUTO_INIT + '\n</body>', 1)
    return html

def update_html(results, prod_planta):
    with open(HTML_FILE, encoding='utf-8') as f:
        html = f.read()

    rmot = {k: v['mots'] for k, v in results.items()}
    rkpi = {k: v['rkpi'] for k, v in results.items()}
    js   = lambda o: json.dumps(o, ensure_ascii=False, separators=(',', ':'))

    html = re.sub(r'(let|const)\s+R_MOT\s*=\s*\{[\s\S]*?\};',
        f"let R_MOT = {{\n  brit: {js(rmot['brit'])},\n  l1: {js(rmot['l1'])},\n  l2: {js(rmot['l2'])},\n  l3: {js(rmot['l3'])}\n}};", html, count=1)

    html = re.sub(r'(let|const)\s+R_KPI\s*=\s*\{[\s\S]*?\};',
        f"let R_KPI = {{\n  brit: {js(rkpi['brit'])},\n  l1: {js(rkpi['l1'])},\n  l2: {js(rkpi['l2'])},\n  l3: {js(rkpi['l3'])}\n}};", html, count=1)

    html = re.sub(r'(let|const)\s+PROD_PLANTA\s*=\s*\{[\s\S]*?\};',
        f"let PROD_PLANTA = {js(prod_planta)};", html, count=1)

    html = inject_auto_init(html)

    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"✓ {HTML_FILE} atualizado com auto-inicialização")

if __name__ == '__main__':
    print("⚙  Processando planilha...")
    results, prod_planta = process()
    print("📝 Atualizando index.html...")
    update_html(results, prod_planta)
    print("✅ Concluído!")
