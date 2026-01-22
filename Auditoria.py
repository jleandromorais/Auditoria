import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import xml.etree.ElementTree as ET
import math
from openpyxl.styles import PatternFill, Font, Alignment

# --- CONFIGURAÇÕES ---
ANO_ALVO = "25"
MESES_ALVO = ["OUT", "NOV", "DEZ"]

# ==============================================================================
# 1. LIMPEZA E UTILITÁRIOS
# ==============================================================================

def limpar_numero_nf_bruto(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return ""
    texto = str(valor).upper().strip()
    if "-" in texto:
        texto = texto.split("-")[0]
    if "/" in texto:
        texto = texto.split("/")[0]
    nums = re.findall(r"\d+", texto.replace(".", ""))
    if nums:
        return str(int(nums[0]))
    return ""

def to_float(texto):
    if texto is None or (isinstance(texto, float) and math.isnan(texto)):
        return 0.0
    if isinstance(texto, (int, float)):
        return float(texto)
    if not isinstance(texto, str):
        texto = str(texto)

    clean = texto.replace(" ", "")
    if clean == "":
        return 0.0

    # Corrige múltiplas vírgulas (separador de milhar errado)
    if clean.count(",") > 1:
        partes = clean.split(",")
        inteiro = "".join(partes[:-1])
        decimal = partes[-1]
        clean = f"{inteiro}.{decimal}"

    limpo = re.sub(r"[^\d.,-]", "", clean)

    if "," in limpo and "." not in limpo:
        limpo = limpo.replace(",", ".")
    elif "," in limpo and "." in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")

    try:
        return float(limpo)
    except:
        return 0.0

def make_unique_columns(cols):
    seen = {}
    out = []
    for c in cols:
        base = str(c).strip()
        if base in seen:
            seen[base] += 1
            out.append(f"{base}__{seen[base]}")
        else:
            seen[base] = 0
            out.append(base)
    return out

def safe_float(x):
    if x is None:
        return 0.0
    try:
        if isinstance(x, float) and math.isnan(x):
            return 0.0
    except:
        pass
    return float(x)

# ==============================================================================
# 2. XML (NF-e e CT-e)
# ==============================================================================

def strip_ns(tag):
    return tag.split("}", 1)[1] if "}" in tag else tag

def iter_elems(root, name):
    for el in root.iter():
        if strip_ns(el.tag) == name:
            yield el

def get_first_text(root, path_names):
    # path_names: sequência de tags ignorando namespace
    def rec(node, idx):
        if idx == len(path_names):
            return node.text
        for ch in node:
            if strip_ns(ch.tag) == path_names[idx]:
                t = rec(ch, idx + 1)
                if t is not None:
                    return t
        return None
    return rec(root, 0)

def parse_nfe(root):
    # procura infNFe (mesmo quando é nfeProc)
    inf = next(iter(iter_elems(root, "infNFe")), None)
    if inf is None:
        return None

    nNF = get_first_text(inf, ["ide", "nNF"])
    nota = re.sub(r"\D", "", nNF or "")
    if nota:
        nota = str(int(nota))

    vNF = get_first_text(inf, ["total", "ICMSTot", "vNF"])
    vICMS = get_first_text(inf, ["total", "ICMSTot", "vICMS"])
    vPIS = get_first_text(inf, ["total", "ICMSTot", "vPIS"])
    vCOFINS = get_first_text(inf, ["total", "ICMSTot", "vCOFINS"])

    bruto = to_float(vNF)
    icms = to_float(vICMS)
    pis = to_float(vPIS)
    cof = to_float(vCOFINS)

    # Volume: soma itens com unidade M3/NM3
    vol = 0.0
    for det in iter_elems(inf, "det"):
        prod = next((ch for ch in det if strip_ns(ch.tag) == "prod"), None)
        if prod is None:
            continue
        uCom = get_first_text(prod, ["uCom"]) or ""
        qCom = get_first_text(prod, ["qCom"])
        u = uCom.upper().replace("³", "3")
        if "M3" in u:   # pega M3 e NM3
            vol += to_float(qCom)

    # fallback (se não achar no item)
    if vol == 0.0:
        qVol = get_first_text(inf, ["transp", "vol", "qVol"])
        if qVol:
            vol = to_float(qVol)

    liq = bruto
    for v in (icms, pis, cof):
        if bruto > 0 and 0 < v < bruto:
            liq -= v
    if liq < 0:
        liq = 0.0

    return {
        "Tipo": "NF-e",
        "Nota": nota,
        "Vol": vol,
        "Bruto": bruto,
        "ICMS": icms,
        "PIS": pis,
        "COFINS": cof,
        "Liq_Calc": liq
    }

def parse_cte(root):
    inf = next(iter(iter_elems(root, "infCte")), None)
    if inf is None:
        return None

    nCT = get_first_text(inf, ["ide", "nCT"])
    nota = re.sub(r"\D", "", nCT or "")
    if nota:
        nota = str(int(nota))

    bruto = to_float(get_first_text(inf, ["vPrest", "vTPrest"]))

    # ICMS pode existir em vários grupos; pegamos o primeiro vICMS
    icms = 0.0
    for el in iter_elems(inf, "vICMS"):
        icms = to_float(el.text)
        break

    # Alguns CT-es não trazem PIS/COFINS no XML (a planilha calcula “conforme XML”/outro critério).
    pis = 0.0
    cof = 0.0
    for el in iter_elems(inf, "vPIS"):
        pis = to_float(el.text); break
    for el in iter_elems(inf, "vCOFINS"):
        cof = to_float(el.text); break

    vol = 0.0
    for infQ in iter_elems(inf, "infQ"):
        q = get_first_text(infQ, ["qCarga"])
        if q:
            vol = to_float(q)
            if vol > 0:
                break

    liq = bruto
    for v in (icms, pis, cof):
        if bruto > 0 and 0 < v < bruto:
            liq -= v
    if liq < 0:
        liq = 0.0

    return {
        "Tipo": "CT-e",
        "Nota": nota,
        "Vol": vol,
        "Bruto": bruto,
        "ICMS": icms,
        "PIS": pis,
        "COFINS": cof,
        "Liq_Calc": liq
    }

def parse_xml_file(path):
    tree = ET.parse(path)
    root = tree.getroot()

    root_tag = strip_ns(root.tag).lower()
    tags = set(strip_ns(el.tag) for el in root.iter())

    # IMPORTANTE: CT-e pode conter "infNFe" dentro do CT-e (doc transportado),
    # então damos prioridade ao CT-e quando existir infCte e o root for do namespace do CT-e.
    if "infCte" in tags and ("cte" in root_tag or "portalfiscal.inf.br/cte" in root.tag.lower()):
        return parse_cte(root)

    if "infNFe" in tags:
        return parse_nfe(root)

    if "infCte" in tags:
        return parse_cte(root)

    return None

# ==============================================================================
# 3. EXCEL (LÊ NOTAS + VOLUME + S/TRIBUTOS + IMPOSTOS)
# ==============================================================================

def carregar_excel(caminho):
    dados = []
    xls = pd.read_excel(caminho, sheet_name=None, header=None)

    for aba, df in xls.items():
        aba_upper = str(aba).upper()
        if ANO_ALVO not in aba_upper:
            continue
        if not any(mes in aba_upper for mes in MESES_ALVO):
            continue

        idx = -1
        for i, row in df.head(120).iterrows():
            linha = [str(x).upper() for x in row.values]
            if any("NOTA" in x for x in linha) and (
                any("S/TRIBUTOS" in x for x in linha) or
                any("C/TRIBUTOS" in x for x in linha) or
                any("TOTAL" in x for x in linha)
            ):
                idx = i
                break

        if idx == -1:
            continue

        cols = make_unique_columns([str(c).upper().strip() for c in df.iloc[idx]])
        df = df[idx + 1:].copy()
        df.columns = cols

        c_nf = next((c for c in df.columns if "NOTA" in c or c == "NF"), None)
        c_liq = next((c for c in df.columns if "S/TRIBUTOS" in c), None)
        c_vol = next((c for c in df.columns if "VOL" in c or "M³" in c or "M3" in c or "QTDE" in c or "QTD" in c or "QUANT" in c), None)

        c_icms = [c for c in df.columns if c.startswith("ICMS")]
        c_pis  = [c for c in df.columns if c.startswith("PIS")]
        c_cof  = [c for c in df.columns if c.startswith("COFINS")]

        c_icms = c_icms[-1] if c_icms else None
        c_pis  = c_pis[-1]  if c_pis  else None
        c_cof  = c_cof[-1]  if c_cof  else None

        if not (c_nf and c_liq):
            continue

        temp = df.copy()
        temp["NF_Clean"] = temp[c_nf].apply(limpar_numero_nf_bruto)
        temp["Vol_Excel"] = temp[c_vol].apply(to_float) if c_vol else 0.0
        temp["Liq_Excel"] = temp[c_liq].apply(to_float)

        temp["ICMS_Excel"] = temp[c_icms].apply(to_float) if c_icms else 0.0
        temp["PIS_Excel"] = temp[c_pis].apply(to_float) if c_pis else 0.0
        temp["COFINS_Excel"] = temp[c_cof].apply(to_float) if c_cof else 0.0

        temp["Mes"] = aba
        temp = temp[temp["NF_Clean"] != ""]

        if not temp.empty:
            dados.append(temp[["NF_Clean", "Vol_Excel", "Liq_Excel", "ICMS_Excel", "PIS_Excel", "COFINS_Excel", "Mes"]])

    return pd.concat(dados, ignore_index=True) if dados else pd.DataFrame()

# ==============================================================================
# 4. RELATÓRIO
# ==============================================================================

def gerar_relatorio(lista, saida=None):
    df = pd.DataFrame(lista)
    cols = [
        "Arquivo", "Tipo", "Mes", "Nota",
        "Vol XML", "Vol Excel", "Diff Vol",
        "Bruto XML", "ICMS XML", "PIS", "COFINS",
        "ICMS Excel", "PIS Excel", "COFINS Excel",
        "Liq XML (Calc)", "Liq Excel", "Diff R$", "Status",
        "Obs"
    ]
    for c in cols:
        if c not in df.columns:
            df[c] = "-"
    df = df[cols]

    ts = datetime.now().strftime("%H%M%S")
    if saida is None:
        saida = os.path.join(os.environ.get("USERPROFILE", os.getcwd()), "Downloads", f"Auditoria_XML_{ts}.xlsx")

    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
        ws = writer.sheets["Resultado"]

        header_fill = PatternFill("solid", fgColor="203764")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        verde = PatternFill("solid", fgColor="C6EFCE")
        vermelho = PatternFill("solid", fgColor="FFC7CE")

        status_col = cols.index("Status") + 1
        for row in ws.iter_rows(min_row=2):
            status = str(row[status_col - 1].value)
            cor = verde if "OK" in status else vermelho
            for cell in row:
                cell.fill = cor
                if isinstance(cell.value, (int, float)):
                    if cell.col_idx >= cols.index("Bruto XML") + 1:
                        cell.number_format = 'R$ #,##0.00'
                    if cell.col_idx in [cols.index("Vol XML")+1, cols.index("Vol Excel")+1, cols.index("Diff Vol")+1]:
                        cell.number_format = '#,##0.000'

        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22

    try:
        os.startfile(saida)
    except:
        pass

    return saida

# ==============================================================================
# MAIN (AUDITORIA XML)
# ==============================================================================

def auditar(xmls, excel_path, saida=None):
    df_base = carregar_excel(excel_path)
    if df_base.empty:
        raise RuntimeError("Não consegui carregar as abas do Excel alvo (verifique ANO_ALVO e MESES_ALVO).")

    relatorio = []

    for xml in xmls:
        nome = os.path.basename(xml)
        try:
            info = parse_xml_file(xml)
        except Exception as e:
            info = None

        item = {
            "Arquivo": nome,
            "Tipo": info["Tipo"] if info else "-",
            "Mes": "-",
            "Nota": info["Nota"] if info else "",
            "Vol XML": info["Vol"] if info else 0.0,
            "Bruto XML": info["Bruto"] if info else 0.0,
            "ICMS XML": info["ICMS"] if info else 0.0,
            "PIS": info["PIS"] if info else 0.0,
            "COFINS": info["COFINS"] if info else 0.0,
            "Liq XML (Calc)": info["Liq_Calc"] if info else 0.0,
            "Vol Excel": 0.0,
            "Liq Excel": 0.0,
            "ICMS Excel": 0.0,
            "PIS Excel": 0.0,
            "COFINS Excel": 0.0,
            "Diff Vol": 0.0,
            "Diff R$": 0.0,
            "Status": "ERRO PARSE ❌" if not info else "Ñ ENCONTRADO ⚠️",
            "Obs": ""
        }

        if info and info["Nota"]:
            match = df_base[df_base["NF_Clean"] == info["Nota"]]
            if not match.empty:
                row = match.iloc[0]
                item["Mes"] = row["Mes"]

                vol_excel = safe_float(row["Vol_Excel"])
                liq_excel = safe_float(row["Liq_Excel"])
                icms_excel = safe_float(row["ICMS_Excel"])
                pis_excel = safe_float(row["PIS_Excel"])
                cof_excel = safe_float(row["COFINS_Excel"])

                item["Vol Excel"] = vol_excel if vol_excel != 0 else "NÃO NO EXCEL"
                item["Liq Excel"] = liq_excel
                item["ICMS Excel"] = icms_excel
                item["PIS Excel"] = pis_excel
                item["COFINS Excel"] = cof_excel

                # Ajuste do líquido:
                # - NF-e: usa o XML direto.
                # - CT-e: se o XML não trouxer PIS/COFINS (comum), usa os valores do Excel para comparar com "S/TRIBUTOS".
                icms = info["ICMS"]
                pis = info["PIS"]
                cof = info["COFINS"]
                if info["Tipo"] == "CT-e" and pis == 0 and cof == 0 and (pis_excel != 0 or cof_excel != 0):
                    pis = pis_excel
                    cof = cof_excel
                    item["Obs"] = "CT-e sem PIS/COFINS no XML; usei valores do Excel p/ calcular líquido."

                liq_calc = info["Bruto"] - sum(v for v in (icms, pis, cof) if 0 < v < info["Bruto"])
                if liq_calc < 0:
                    liq_calc = 0.0
                item["Liq XML (Calc)"] = liq_calc

                item["Diff Vol"] = "-" if vol_excel == 0 else (info["Vol"] - vol_excel)
                item["Diff R$"] = liq_calc - liq_excel

                tol_r = 50.0 if info["Tipo"] == "CT-e" else 5.0
                financeiro_ok = abs(item["Diff R$"]) < tol_r
                volume_ok = True if vol_excel == 0 else abs(float(item["Diff Vol"])) < 1.0

                if financeiro_ok and volume_ok:
                    item["Status"] = "OK ✅"
                else:
                    status = []
                    if not volume_ok and vol_excel != 0:
                        status.append("VOL")
                    if not financeiro_ok:
                        status.append("VALOR")
                    item["Status"] = f"ERRO {'+'.join(status)} ❌"

        relatorio.append(item)

    return gerar_relatorio(relatorio, saida=saida)

def main():
    root = tk.Tk()
    root.withdraw()

    xmls = filedialog.askopenfilenames(title="1. XMLs (NF-e / CT-e)", filetypes=[("XML", "*.xml;*.XML")])
    if not xmls:
        return

    excel = filedialog.askopenfilename(title="2. Excel", filetypes=[("Excel", "*.xlsx")])
    if not excel:
        return

    try:
        auditar(xmls, excel)
    except Exception as e:
        messagebox.showerror("Erro", str(e))

if __name__ == "__main__":
    main()
