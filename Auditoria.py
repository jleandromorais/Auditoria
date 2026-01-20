import pdfplumber
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import numpy as np

# --- CONFIGURAÇÕES ---
ANO_ALVO = "25"
MESES_ALVO = ["OUT", "NOV", "DEZ"] 

# ==============================================================================
# 1. LIMPEZA E UTILITÁRIOS
# ==============================================================================

def limpar_numero_nf_bruto(valor):
    if pd.isna(valor) or str(valor).strip() == "": return ""
    texto = str(valor).upper().strip()
    if "-" in texto: texto = texto.split("-")[0]
    if "/" in texto: texto = texto.split("/")[0]
    nums = re.findall(r'\d+', texto.replace(".", ""))
    if nums: return str(int(nums[0]))
    return ""

def to_float(texto):
    if pd.isna(texto) or texto == "": return 0.0
    if isinstance(texto, (int, float)): return float(texto)
    if not isinstance(texto, str): return 0.0
    
    clean = texto.replace(" ", "")
    if "27112100" in clean.replace(".", ""): return 0.0 
    
    limpo = re.sub(r'[^\d,]', '', clean)
    if ',' in limpo:
        limpo = limpo.replace('.', '').replace(',', '.')
    try:
        val = float(limpo)
        if val > 500000000 or val in [2024.0, 2025.0, 2026.0]: return 0.0
        return val
    except: return 0.0

def remover_area_transporte_agressivo(texto):
    texto_novo = re.sub(r'(TRANSPORTADOR.*?DADOS DO PRODUTO)', '', texto, flags=re.IGNORECASE | re.DOTALL)
    if len(texto_novo) == len(texto):
        texto_novo = re.sub(r'(TRANSPORTADOR.*?CÓDIGO)', '', texto, flags=re.IGNORECASE | re.DOTALL)
    return texto_novo

# ==============================================================================
# 2. EXTRAÇÃO
# ==============================================================================

def extrair_dados_tanque_final(texto_bruto, nome_arquivo):
    info = {
        'Arquivo': nome_arquivo, 'Tipo': 'NF-e', 'Nota': '',
        'Vol': 0.0, 'Bruto': 0.0, 'ICMS': 0.0, 'Liq_Calc': 0.0
    }
    
    if any(x in texto_bruto.upper() for x in ["CONHECIMENTO DE TRANSPORTE", "DACTE", "CT-E"]):
        info['Tipo'] = "CT-e"
    
    if info['Tipo'] == 'NF-e':
        texto_analise = remover_area_transporte_agressivo(texto_bruto)
    else:
        texto_analise = texto_bruto

    match_nf = re.search(r'(?:N[º°o\.]*|NUMERO|DOC\.|DOCUMENTO)\s*[:\.]?\s*(\d+(?:\.\d+)*)', texto_analise, re.IGNORECASE)
    if match_nf: 
        info['Nota'] = limpar_numero_nf_bruto(match_nf.group(1))
    
    if not info['Nota']:
        match_chave = re.search(r'(\d{44})', texto_analise.replace(" ", ""))
        if match_chave:
            chave = match_chave.group(1)
            info['Nota'] = str(int(chave[25:34]))

    if info['Tipo'] == 'NF-e':
        padroes_nfe = [
            r'([\d\.]+,\d{1,4})\s*(?:M3|M³|NM3)',
            r'(?:M3|M³|NM3).*?([\d\.]+,\d{1,4})',
            r'(?:QUANTIDADE|QUANT|QTDE|QTD|QTD\.|OTDIE)\s*[:\.]?.*?([\d\.]+,\d{1,4})',
            r'([\d\.]+,\d{1,4})\s*(?:L|LT|LITROS)\b',
            r'PESO\s*L[IÍ]QUIDO.*?([\d\.]+,\d{1,4})'
        ]
        for pat in padroes_nfe:
            match = re.search(pat, texto_analise, re.IGNORECASE)
            if not match and '.*?' in pat:
                match = re.search(pat, texto_analise, re.IGNORECASE | re.DOTALL)
            if match:
                v = to_float(match.group(1))
                if v > 0:
                    info['Vol'] = v
                    break
    else:
        termos_cte = [
            r'([\d\.]+,\d{2,4})\s*mmbtu', r'PESO\s*TAXADO.*?([\d\.]+,\d{2,4})', 
            r'CARGA.*?([\d\.]+,\d{2,4})', r'CUBAGEM.*?([\d\.]+,\d{2,4})', 
            r'PESO\s*AFERIDO.*?([\d\.]+,\d{2,4})'
        ]
        for t in termos_cte:
            m = re.search(t, texto_analise, re.IGNORECASE | re.DOTALL)
            if m: 
                v = to_float(m.group(1))
                if v > 0: 
                    info['Vol'] = v
                    break

    todos_valores = re.findall(r'[\d\.]+,\d{2}', texto_analise)
    floats = sorted([to_float(v) for v in todos_valores], reverse=True)
    if floats: info['Bruto'] = floats[0]

    if info['Bruto'] > 0:
        for val in floats:
            if val == info['Bruto']: continue
            ratio = val / info['Bruto']
            if 0.07 <= ratio <= 0.27: 
                info['ICMS'] = val
                break

    if info['Bruto'] > 0:
        base = info['Bruto']
        if info['ICMS'] > 0: base = base - info['ICMS']
        info['Liq_Calc'] = base * 0.9075

    return info

# ==============================================================================
# 3. EXCEL
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
                c_val = next((c for c in df.columns if 'S/TRIBUTOS' in c), None)
                if not c_val: c_val = next((c for c in df.columns if 'VALOR' in c or 'COMPRA' in c), None)
                c_vol = next((c for c in df.columns if any(x in c for x in ['VOL', 'QTD'])), None)

                if c_nf and c_val:
                    temp = df.copy()
                    temp['NF_Clean'] = temp[c_nf].apply(limpar_numero_nf_bruto)
                    # Força conversão para float e preenche vazios com 0.0
                    temp['Vol_Excel'] = temp[c_vol].apply(to_float) if c_vol else 0.0
                    temp['Liq_Excel'] = temp[c_val].apply(to_float)
                    temp['Mes'] = aba
                    
                    temp = temp[temp['NF_Clean'] != ""]
                    if not temp.empty:
                        dados.append(temp[['NF_Clean', 'Vol_Excel', 'Liq_Excel', 'Mes']])
                        print(f"   ✅ Aba {aba}: Carregada")
        
        if not dados: return pd.DataFrame()
        return pd.concat(dados)
    except Exception as e:
        messagebox.showerror("Erro Excel", str(e))
        return pd.DataFrame()

# ==============================================================================
# 4. RELATÓRIO
# ==============================================================================

def gerar_relatorio(lista):
    df = pd.DataFrame(lista)
    cols = ['Arquivo', 'Tipo', 'Mes', 'Nota', 'Vol PDF', 'Vol Excel', 'Diff Vol', 
            'Bruto PDF', 'ICMS PDF', 'Liq PDF (Calc)', 'Liq Excel', 'Diff R$', 'Status']
    for c in cols: 
        if c not in df.columns: df[c] = '-'
    df = df[cols]
    
    ts = datetime.now().strftime("%H%M%S")
    saida = os.path.join(os.environ['USERPROFILE'], 'Downloads', f'Auditoria_Final_{ts}.xlsx')
    
    try:
        with pd.ExcelWriter(saida, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Resultado')
            ws = writer.sheets['Resultado']
            
            header = PatternFill("solid", fgColor="203764")
            font = Font(bold=True, color="FFFFFF")
            for cell in ws[1]:
                cell.fill = header; cell.font = font; cell.alignment = Alignment(horizontal='center')
            
            verde = PatternFill("solid", fgColor="C6EFCE")
            vermelho = PatternFill("solid", fgColor="FFC7CE")
            
            for row in ws.iter_rows(min_row=2):
                status = str(row[12].value)
                # Pinta de verde se tiver "OK"
                cor = verde if "OK" in status else vermelho
                
                for cell in row:
                    cell.fill = cor
                    if isinstance(cell.value, (int, float)):
                        if cell.col_idx >= 8: cell.number_format = 'R$ #,##0.00'
                        if cell.col_idx in [5,6,7]: cell.number_format = '#,##0.000'
            
            for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = 17

        print(f"\n✅ Relatório: {saida}")
        try: os.startfile(saida)
        except: pass
    except: print("❌ Erro ao salvar.")

# ==============================================================================
# MAIN
# ==============================================================================

def main():
    root = tk.Tk(); root.withdraw()
    print("--- AUDITORIA FINAL (LÓGICA: S/ VOL + $ BATENDO = OK) ---")
    
    pdfs = filedialog.askopenfilenames(title="1. PDFs", filetypes=[("PDF", "*.pdf")])
    if not pdfs: return
    
    excel = filedialog.askopenfilename(title="2. Excel", filetypes=[("Excel", "*.xlsx")])
    if not excel: return
    
    df_base = carregar_excel(excel)
    if df_base.empty: return

    relatorio = []
    print("\n🔍 Analisando...")
    
    for pdf in pdfs:
        try:
            with pdfplumber.open(pdf) as p:
                texto = ""
                for page in p.pages: texto += page.extract_text() + "\n"
                info = extrair_dados_tanque_final(texto, os.path.basename(pdf))
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
                
                # --- TRATAMENTO ROBUSTO PARA ZERO/VAZIO ---
                vol_excel_raw = item['Vol Excel']
                # Se for NaN, None ou vazio, vira 0.0
                if pd.isna(vol_excel_raw) or str(vol_excel_raw).strip() == '':
                    vol_excel_num = 0.0
                else:
                    vol_excel_num = to_float(vol_excel_raw)
                
                item['Diff Vol'] = v_pdf - vol_excel_num
                item['Diff R$'] = info['Liq_Calc'] - item['Liq Excel']
                
                tol_r = 50.0 if info['Tipo'] == 'CT-e' else 5.0

                # --- LÓGICA DE VALIDAÇÃO FINAL ---
                financeiro_ok = abs(item['Diff R$']) < tol_r
                volume_ok = abs(item['Diff Vol']) < 1.0
                excel_sem_vol = (vol_excel_num == 0) # Verifica se é ZERO

                # CONDIÇÃO DO SUCESSO:
                # 1. Financeiro bate
                # 2. E (Volume bate OU Excel está vazio/zero)
                if financeiro_ok and (volume_ok or excel_sem_vol):
                    item['Status'] = 'OK ✅'  # Aqui garante o status verde
                    
                    if excel_sem_vol:
                        # Aviso visual APENAS na coluna de volume, Status continua OK
                        item['Vol Excel'] = "NÃO NO EXCEL"
                        item['Diff Vol'] = "-"
                else:
                    status = []
                    # Só acusa erro de volume se o excel TIVER volume e estiver errado
                    if not volume_ok and not excel_sem_vol: status.append("VOL")
                    if not financeiro_ok: status.append("VALOR")
                    
                    # Fallback para erros não classificados
                    if not status: status.append("VOL (ERRO)")
                    
                    item['Status'] = f"ERRO {'+'.join(status)} ❌"
        
        item['Vol PDF'] = info['Vol']
        item['Bruto PDF'] = info['Bruto']
        item['ICMS PDF'] = info['ICMS']
        item['Liq PDF (Calc)'] = info['Liq_Calc']
        
        relatorio.append(item)
        print(f"-> {info['Tipo']} {info['Nota']}: {item['Status']}")

    gerar_relatorio(relatorio)

if __name__ == "__main__":
    main()