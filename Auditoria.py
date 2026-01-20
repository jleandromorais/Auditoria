import pdfplumber
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- CONFIGURAÇÕES ---
ANO_ALVO = "25"
MESES_ALVO = ["OUT", "NOV", "DEZ"] 

# ==============================================================================
# 1. UTILITÁRIOS
# ==============================================================================

def limpar_numero_nf(valor):
    if pd.isna(valor) or str(valor).strip() == "": return ""
    nums = re.findall(r'\d+', str(valor).replace(".", ""))
    if nums: return str(int(nums[0]))
    return ""

def to_float(texto):
    if isinstance(texto, (int, float)): return float(texto)
    if not isinstance(texto, str): return 0.0
    
    # Limpa sujeira e NCMs
    clean = texto.replace(" ", "")
    if "27112100" in clean.replace(".", ""): return 0.0
    
    limpo = re.sub(r'[^\d,]', '', clean)
    if ',' in limpo:
        limpo = limpo.replace('.', '').replace(',', '.')
    
    try:
        val = float(limpo)
        # Filtros: Datas (2025), Chaves (>100M)
        if val > 500000000 or val in [2024.0, 2025.0, 2026.0]: return 0.0
        return val
    except: return 0.0

def remover_area_transporte(texto):
    """
    Remove o bloco de 'TRANSPORTADOR' até 'DADOS DO PRODUTO' 
    para evitar pegar Peso Bruto/Líquido como Volume.
    """
    # Regex que apaga tudo entre "TRANSPORTADOR" e "DADOS DO PRODUTO"
    # O flag re.DOTALL permite que o . pegue quebras de linha
    texto_limpo = re.sub(r'(TRANSPORTADOR.*?DADOS DO PRODUTO)', '', texto, flags=re.IGNORECASE | re.DOTALL)
    
    # Se não achou 'DADOS DO PRODUTO', tenta apagar até 'CÓDIGO' ou 'DESCRIÇÃO'
    if len(texto_limpo) == len(texto):
        texto_limpo = re.sub(r'(TRANSPORTADOR.*?DESCRI)', '', texto, flags=re.IGNORECASE | re.DOTALL)
        
    return texto_limpo

# ==============================================================================
# 2. EXTRAÇÃO DE DADOS (MODO CORRETIVO)
# ==============================================================================

def extrair_dados_final(texto_bruto, nome_arquivo, pdf_obj=None):
    info = {
        'Arquivo': nome_arquivo,
        'Tipo': 'NF-e',
        'Nota': '',
        'Vol': 0.0,
        'Bruto': 0.0,
        'ICMS': 0.0,
        'Liq_Calc': 0.0,
        'Obs': ''
    }
    
    # 1. Identifica Tipo
    if any(x in texto_bruto.upper() for x in ["CONHECIMENTO DE TRANSPORTE", "DACTE", "CT-E"]):
        info['Tipo'] = "CT-e"
    
    # 2. Tenta extrair VOLUME direto da Tabela (Mais seguro)
    # Procura linhas com "GAS", "GNC", "2711"
    if pdf_obj and info['Tipo'] == 'NF-e':
        for page in pdf_obj.pages:
            tabs = page.extract_tables()
            for tab in tabs:
                df = pd.DataFrame(tab)
                # Converte tudo para string
                df = df.astype(str)
                
                for i, row in df.iterrows():
                    linha_str = " ".join(row.values).upper()
                    # Se a linha parece ser do produto Gás
                    if "GAS" in linha_str or "2711" in linha_str.replace(".",""):
                        # Tenta achar o volume nesta linha específica
                        # Procura números com 3 ou 4 casas decimais ou grandes inteiros
                        numeros = re.findall(r'[\d\.]+,\d+', linha_str)
                        for n in numeros:
                            v = to_float(n)
                            # Regra: Volume > 0, Menor que valor financeiro gigante, e não é o NCM
                            if v > 0 and v < 200000000:
                                # Geralmente volume é o menor número grande da linha (Preço Unit < Vol < Preço Total)
                                # Ou simplesmente o primeiro número grande que aparece
                                if info['Vol'] == 0: 
                                    info['Vol'] = v
                                # Se já tem volume, mas achou um maior que não seja o preço total (ex: 15.000.000)
                                # Mantemos o primeiro achado na linha do produto costuma ser a QTD
    
    # 3. Limpeza de Texto (Remove Transportadora)
    texto_limpo = remover_area_transporte(texto_bruto)
    
    # 4. Número da Nota
    match_nf = re.search(r'(?:N[º°o\.]*|NUMERO)\s*[:\.]?\s*(\d+(?:\.\d+)*)', texto_limpo, re.IGNORECASE)
    if match_nf: info['Nota'] = limpar_numero_nf(match_nf.group(1))

    # 5. Volume (Fallback Regex se a tabela falhou)
    if info['Vol'] == 0:
        # Procura M3 especificamente no texto limpo (sem transportadora)
        match_m3 = re.search(r'([\d\.]+,\d{2,4})\s*(?:M3|M³)', texto_limpo, re.IGNORECASE)
        if not match_m3:
             match_m3 = re.search(r'(?:M3|M³)\s*([\d\.]+,\d{2,4})', texto_limpo, re.IGNORECASE)
        
        if match_m3:
            info['Vol'] = to_float(match_m3.group(1))
        else:
            # Tenta QUANTIDADE
            match_qtd = re.search(r'(?:QUANTIDADE|QTD).*?([\d\.]+,\d{2,4})', texto_limpo, re.IGNORECASE | re.DOTALL)
            if match_qtd: info['Vol'] = to_float(match_qtd.group(1))

    # 6. Financeiro (Bruto)
    # Pega todos os valores monetários
    todos_valores = re.findall(r'[\d\.]+,\d{2}', texto_limpo)
    floats = sorted([to_float(v) for v in todos_valores], reverse=True)
    if floats: info['Bruto'] = floats[0] # Maior valor

    # 7. Financeiro (ICMS - Busca Lógica)
    if info['Bruto'] > 0:
        for val in floats:
            if val == info['Bruto']: continue
            ratio = val / info['Bruto']
            if 0.07 <= ratio <= 0.25: # Faixa 7% a 25%
                info['ICMS'] = val
                break

    # 8. Cálculo Líquido
    if info['Bruto'] > 0:
        if info['Tipo'] == "CT-e":
            # Transporte: Líquido = Bruto - ICMS
            info['Liq_Calc'] = info['Bruto'] - info['ICMS']
            info['Obs'] = "Transp"
        else:
            # Gás: Líquido = (Bruto - ICMS) * 0.9075
            # Se não achou ICMS, assume que está embutido ou é zero, mas aplica fator se possível?
            # Melhor: Se não tem ICMS, usa Bruto (para acusar erro se precisar), mas se tiver, aplica fórmula.
            if info['ICMS'] > 0:
                info['Liq_Calc'] = (info['Bruto'] - info['ICMS']) * 0.9075
                info['Obs'] = "Gás (Molécula)"
            else:
                info['Liq_Calc'] = info['Bruto'] # Sem ICMS detectado

    return info

# ==============================================================================
# 3. EXCEL E RELATÓRIO
# ==============================================================================

def carregar_excel(caminho):
    print(f"⏳ Lendo Excel ({MESES_ALVO})...")
    dados = []
    try:
        xls = pd.read_excel(caminho, sheet_name=None, header=None)
        for aba, df in xls.items():
            aba_upper = str(aba).upper()
            if ANO_ALVO not in aba_upper: continue
            if not any(mes in aba_upper for mes in MESES_ALVO): continue

            idx = -1
            for i, row in df.head(60).iterrows():
                linha = [str(x).upper() for x in row.values]
                if any('NOTA' in x for x in linha) and (any('S/TRIBUTOS' in x for x in linha) or any('TOTAL' in x for x in linha)):
                    idx = i
                    break
            
            if idx != -1:
                df.columns = [str(c).upper().strip() for c in df.iloc[idx]]
                df = df[idx+1:].copy()
                
                c_nf = next((c for c in df.columns if any(x in c for x in ['NOTA', 'NF'])), None)
                c_val = next((c for c in df.columns if 'S/TRIBUTOS' in c), None) # Prioridade
                if not c_val: c_val = next((c for c in df.columns if 'VALOR' in c or 'COMPRA' in c), None)
                c_vol = next((c for c in df.columns if any(x in c for x in ['VOL', 'QTD'])), None)

                if c_nf and c_val:
                    temp = df.copy()
                    temp['NF_Clean'] = temp[c_nf].apply(limpar_numero_nf)
                    temp['Vol_Excel'] = temp[c_vol].apply(to_float) if c_vol else 0.0
                    temp['Liq_Excel'] = temp[c_val].apply(to_float)
                    temp['Mes'] = aba
                    temp = temp[temp['NF_Clean'] != ""]
                    if not temp.empty:
                        dados.append(temp[['NF_Clean', 'Vol_Excel', 'Liq_Excel', 'Mes']])
                        print(f"   ✅ Aba {aba}: OK")
        
        if not dados: return pd.DataFrame()
        return pd.concat(dados)
    except Exception as e:
        messagebox.showerror("Erro Excel", str(e))
        return pd.DataFrame()

def gerar_relatorio(lista):
    df = pd.DataFrame(lista)
    cols = ['Arquivo', 'Tipo', 'Mes', 'Nota', 'Vol PDF', 'Vol Excel', 'Diff Vol', 
            'Bruto PDF', 'ICMS PDF', 'Liq PDF (Calc)', 'Liq Excel', 'Diff R$', 'Status', 'Obs']
    for c in cols: 
        if c not in df.columns: df[c] = '-'
    df = df[cols]
    
    ts = datetime.now().strftime("%H%M%S")
    saida = os.path.join(os.environ['USERPROFILE'], 'Downloads', f'Auditoria_Corrigida_{ts}.xlsx')
    
    with pd.ExcelWriter(saida, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultado')
        ws = writer.sheets['Resultado']
        
        header = PatternFill("solid", fgColor="203764")
        font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = header
            cell.font = font
        
        verde = PatternFill("solid", fgColor="C6EFCE")
        vermelho = PatternFill("solid", fgColor="FFC7CE")
        
        for row in ws.iter_rows(min_row=2):
            status = str(row[12].value)
            cor = verde if "OK" in status else vermelho
            for cell in row:
                cell.fill = cor
                cell.border = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
                if isinstance(cell.value, (int, float)):
                    if cell.col_idx in [5,6,7]: cell.number_format = '#,##0.000' # Vol
                    if cell.col_idx >= 8: cell.number_format = 'R$ #,##0.00' # $
        
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 16

    try: os.startfile(saida)
    except: pass
    print(f"\n✅ Relatório: {saida}")

# ==============================================================================
# MAIN
# ==============================================================================

def main():
    root = tk.Tk(); root.withdraw()
    print("--- AUDITORIA: CORREÇÃO VOLUME (IGNORE TRANSPORTE) ---")
    
    pdfs = filedialog.askopenfilenames(title="1. PDFs", filetypes=[("PDF", "*.pdf")])
    if not pdfs: return
    
    excel = filedialog.askopenfilename(title="2. Excel", filetypes=[("Excel", "*.xlsx")])
    if not excel: return
    
    df_base = carregar_excel(excel)
    if df_base.empty: return

    relatorio = []
    print("\n🔍 Analisando...")
    
    for pdf in pdfs:
        # Leitura
        try:
            with pdfplumber.open(pdf) as p:
                texto = ""
                for page in p.pages: texto += page.extract_text() + "\n"
                
                # Passa o objeto PDF para tentar ler tabela se precisar
                info = extrair_dados_final(texto, os.path.basename(pdf), p)
        except: continue
        
        item = info.copy()
        item['Vol Excel'] = 0; item['Liq Excel'] = 0; item['Diff Vol'] = 0; item['Diff R$'] = 0
        item['Status'] = 'Ñ ENCONTRADO ⚠️'
        item['Mes'] = '-'
        
        if info['Nota']:
            match = df_base[df_base['NF_Clean'] == info['Nota']]
            if not match.empty:
                row = match.iloc[0]
                item['Vol Excel'] = row['Vol_Excel']
                item['Liq Excel'] = row['Liq_Excel']
                item['Mes'] = row['Mes']
                
                v_pdf = info['Vol']
                item['Diff Vol'] = v_pdf - item['Vol Excel']
                item['Diff R$'] = info['Liq_Calc'] - item['Liq Excel']
                
                tol_r = 50.0 if info['Tipo'] == 'CT-e' else 5.0
                
                if abs(item['Diff Vol']) < 1.0 and abs(item['Diff R$']) < tol_r:
                    item['Status'] = 'OK ✅'
                else:
                    item['Status'] = 'ERRO VALOR ❌' if abs(item['Diff R$']) >= tol_r else 'ERRO VOL ❌'
        
        item['Vol PDF'] = info['Vol']
        item['Bruto PDF'] = info['Bruto']
        item['ICMS PDF'] = info['ICMS']
        item['Liq PDF (Calc)'] = info['Liq_Calc']
        
        relatorio.append(item)
        print(f"-> {info['Tipo']} {info['Nota']}: {item['Status']}")

    gerar_relatorio(relatorio)

if __name__ == "__main__":
    main()