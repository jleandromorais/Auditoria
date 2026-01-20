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

# --- FUNÇÕES ---
def limpar_numero_nf(valor):
    if pd.isna(valor) or str(valor).strip() == "": return ""
    nums = re.findall(r'\d+', str(valor))
    if nums: return str(int(nums[0]))
    return ""

def string_para_float(texto):
    if isinstance(texto, (int, float)): return float(texto)
    if not isinstance(texto, str): return 0.0
    
    # Filtra NCM e códigos
    clean = texto.replace(".", "").replace(" ", "")
    if "27112100" in clean: return 0.0
    
    # Padroniza
    limpo = re.sub(r'[^\d,]', '', texto)
    if ',' in limpo:
        limpo = limpo.replace('.', '').replace(',', '.')
    
    try:
        val = float(limpo)
        if val > 100000000 or val in [2024.0, 2025.0, 2026.0]: return 0.0
        return val
    except: return 0.0

# --- LÓGICA INTELIGENTE DE RESGATE ---
def encontrar_melhor_icms(texto_completo, valor_bruto):
    """
    Em vez de confiar na posição, procura um número que faça sentido matemático.
    O ICMS do gás costuma ser 12% ou 18% do valor bruto.
    """
    if valor_bruto == 0: return 0.0
    
    # Pega todos os números do PDF que parecem dinheiro
    todos_numeros = re.findall(r'[\d\.]+,\d{2}', texto_completo)
    
    candidatos = []
    for num_str in todos_numeros:
        val = string_para_float(num_str)
        if val == 0: continue
        
        ratio = val / valor_bruto
        # Se o valor for entre 11% e 19% do total, é quase certeza que é o ICMS
        if 0.11 <= ratio <= 0.19:
            return val # Achamos o campeão!
            
    return 0.0

def encontrar_volume_gas(texto_completo, valor_bruto):
    """
    Procura volume perto de M3 ou QUANTIDADE.
    Ignora valores iguais ao financeiro.
    """
    # Tenta achar padrão "M3 6.200.000" (comum no seu PDF)
    match_pos_m3 = re.search(r'M3\s*([\d\.]+,\d{1,4})', texto_completo, re.IGNORECASE)
    if match_pos_m3:
        vol = string_para_float(match_pos_m3.group(1))
        if vol > 0 and vol != valor_bruto: return vol

    # Tenta achar padrão "QUANTIDADE ... 6.200.000"
    match_qtd = re.search(r'(?:QUANTIDADE|QTD|VOL).*?([\d\.]+,\d{1,4})', texto_completo, re.IGNORECASE | re.DOTALL)
    if match_qtd:
        vol = string_para_float(match_qtd.group(1))
        if vol > 0 and vol != valor_bruto: return vol
        
    return 0.0

# --- EXCEL ---
def carregar_excel_trimestre(caminho):
    print(f"⏳ Lendo Excel ({MESES_ALVO} / 20{ANO_ALVO})...")
    dados = []
    try:
        xls = pd.read_excel(caminho, sheet_name=None, header=None)
        for aba, df in xls.items():
            aba_upper = str(aba).upper()
            if ANO_ALVO not in aba_upper: continue
            if not any(mes in aba_upper for mes in MESES_ALVO): continue

            idx_header = -1
            for i, row in df.head(50).iterrows():
                linha = [str(x).upper() for x in row.values]
                if any('NOTA' in x for x in linha) and any('S/TRIBUTOS' in x for x in linha):
                    idx_header = i
                    break
            
            if idx_header != -1:
                df.columns = [str(c).upper().strip() for c in df.iloc[idx_header]]
                df = df[idx_header+1:].copy()
                
                c_nf = next((c for c in df.columns if any(x in c for x in ['NOTA', 'NF'])), None)
                c_vol = next((c for c in df.columns if any(x in c for x in ['VOL', 'QTD'])), None)
                c_val = next((c for c in df.columns if 'S/TRIBUTOS' in c), None)
                
                if c_nf and c_vol and c_val:
                    temp = df[[c_nf, c_vol, c_val]].copy()
                    temp.columns = ['NF', 'Vol_Excel', 'Liq_Excel']
                    temp['NF'] = temp['NF'].apply(limpar_numero_nf)
                    temp['Vol_Excel'] = temp['Vol_Excel'].apply(string_para_float)
                    temp['Liq_Excel'] = temp['Liq_Excel'].apply(string_para_float)
                    temp['Mes_Referencia'] = aba 
                    temp = temp[temp['NF'] != ""]
                    if not temp.empty:
                        dados.append(temp)
                        print(f"   ✅ Aba Carregada: {aba}")
        
        if not dados: return pd.DataFrame()
        return pd.concat(dados)
    except Exception as e:
        messagebox.showerror("Erro Excel", str(e))
        return pd.DataFrame()

# --- PDF PROCESSOR ---
def processar_pdf(caminho):
    info = {
        'Arquivo': os.path.basename(caminho),
        'Nota': '',
        'Vol_PDF': 0.0,
        'Bruto': 0.0,
        'ICMS': 0.0,
        'Liq_Calculado': 0.0
    }
    
    texto_full = ""
    try:
        with pdfplumber.open(caminho) as pdf:
            for page in pdf.pages:
                texto_full += page.extract_text() + "\n"
                
                # 1. Busca Bruto na Tabela (Maior Valor)
                for tab in page.extract_tables():
                    df_tab = pd.DataFrame(tab)
                    for col in df_tab.columns:
                        for v in df_tab[col].astype(str):
                            if ',' in v:
                                val = string_para_float(v)
                                if 100000 < val < 100000000: 
                                    if val > info['Bruto']: info['Bruto'] = val

            # 2. Busca Bruto no Texto (Rede de Segurança)
            if info['Bruto'] == 0:
                match_total = re.search(r'(?:VALOR TOTAL|TOTAL PRODUTOS).*?([\d\.]+,\d{2})', texto_full, re.IGNORECASE | re.DOTALL)
                if match_total: info['Bruto'] = string_para_float(match_total.group(1))

            # 3. Busca ICMS Inteligente (Procura os 12%)
            info['ICMS'] = encontrar_melhor_icms(texto_full, info['Bruto'])
            
            # 4. Busca Volume Inteligente
            info['Vol_PDF'] = encontrar_volume_gas(texto_full, info['Bruto'])
            
            # 5. Busca Nota
            match_nf = re.search(r'(?:N[º°]|NF)\.?\s*[:.]?\s*(\d+)', texto_full, re.IGNORECASE)
            if match_nf: info['Nota'] = limpar_numero_nf(match_nf.group(1))

            # --- CÁLCULO FINAL ---
            # Fórmula Gás: (Bruto - ICMS) * 0.9075
            if info['Bruto'] > 0 and info['ICMS'] > 0:
                base_limpa = info['Bruto'] - info['ICMS']
                info['Liq_Calculado'] = base_limpa * 0.9075
            else:
                # Se não achou ICMS, assume Bruto (para mostrar o erro no Excel)
                info['Liq_Calculado'] = info['Bruto']

    except Exception as e:
        print(f"Erro PDF {caminho}: {e}")

    return info

# --- RELATÓRIO ---
def gerar_excel(dados):
    df = pd.DataFrame(dados)
    cols = ['Arquivo', 'Mes_Referencia', 'Nota', 'Vol PDF', 'Vol Excel', 'Diff Vol', 
            'Bruto PDF', 'ICMS PDF', 'Liq PDF (Calc)', 'Liq Excel', 'Diff R$', 'Status']
    
    for c in cols: 
        if c not in df.columns: df[c] = '-'
    df = df[cols]
    
    timestamp = datetime.now().strftime("%H%M%S")
    saida = os.path.join(os.environ['USERPROFILE'], 'Downloads', f'Auditoria_Gás_Final_{timestamp}.xlsx')
    
    try:
        with pd.ExcelWriter(saida, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Resultado')
            ws = writer.sheets['Resultado']
            
            header = PatternFill("solid", fgColor="203764")
            font_h = Font(bold=True, color="FFFFFF")
            border = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
            
            for cell in ws[1]:
                cell.fill = header
                cell.font = font_h
                cell.alignment = Alignment(horizontal='center')
                
            verde = PatternFill("solid", fgColor="C6EFCE")
            vermelho = PatternFill("solid", fgColor="FFC7CE")
            
            for row in ws.iter_rows(min_row=2):
                status = str(row[11].value)
                cor = verde if "OK" in status else vermelho
                for cell in row:
                    cell.fill = cor
                    cell.border = border
                    if isinstance(cell.value, (int, float)):
                        if cell.col_idx in [4,5,6]: cell.number_format = '#,##0.000'
                        if cell.col_idx >= 7: cell.number_format = 'R$ #,##0.00'
            
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 17

        print(f"\n✅ Relatório Criado: {saida}")
        try: os.startfile(saida)
        except: pass
            
    except PermissionError:
        print("\n❌ ERRO: Feche o arquivo Excel anterior!")

# --- MAIN ---
def main():
    root = tk.Tk(); root.withdraw()
    print("--- AUDITORIA GÁS: BUSCA INTELIGENTE DE ICMS ---")
    
    pdfs = filedialog.askopenfilenames(title="1. Selecione os PDFs", filetypes=[("PDF", "*.pdf")])
    if not pdfs: return
    
    excel = filedialog.askopenfilename(title="2. Selecione o Excel", filetypes=[("Excel", "*.xlsx")])
    if not excel: return
    
    df_base = carregar_excel_trimestre(excel)
    if df_base.empty:
        print("❌ Nenhuma aba correspondente encontrada.")
        return

    relatorio = []
    print("\n🔍 Processando...")
    
    for pdf in pdfs:
        dados = processar_pdf(pdf)
        nf = dados['Nota']
        
        item = dados.copy()
        item['Vol Excel'] = 0.0
        item['Liq Excel'] = 0.0
        item['Diff Vol'] = 0.0
        item['Diff R$'] = 0.0
        item['Status'] = 'Ñ ENCONTRADO ⚠️'
        item['Mes_Referencia'] = '-'
        
        if nf:
            match = df_base[df_base['NF'] == nf]
            if not match.empty:
                row = match.iloc[0]
                item['Vol Excel'] = row['Vol_Excel']
                item['Liq Excel'] = row['Liq_Excel']
                item['Mes_Referencia'] = row['Mes_Referencia']
                
                # Tratamento de zero
                v_pdf = item['Vol_PDF'] if item['Vol_PDF'] > 0 else 0
                
                diff_v = v_pdf - item['Vol Excel']
                diff_r = item['Liq_Calculado'] - item['Liq Excel']
                
                item['Diff Vol'] = diff_v
                item['Diff R$'] = diff_r
                
                # Tolerância
                if abs(diff_v) < 1.0 and abs(diff_r) < 5.0:
                    item['Status'] = 'OK ✅'
                else:
                    item['Status'] = 'ERRO VALOR ❌' if abs(diff_r) >= 5.0 else 'ERRO VOL ❌'
        
        # Renomeia
        item['Bruto PDF'] = dados['Bruto']
        item['ICMS PDF'] = dados['ICMS']
        item['Liq PDF (Calc)'] = dados['Liq_Calculado']
        
        relatorio.append(item)
        print(f"Nota {nf}: {item['Status']} (Liq Calc: {item['Liq_Calculado']:,.2f} | ICMS achado: {item['ICMS']:,.2f})")

    gerar_excel(relatorio)

if __name__ == "__main__":
    main()