# 🧾 Auditoria (NF-e / CT-e) — XML x Excel

## 📌 Sobre o Projeto
O **Auditoria** é uma ferramenta em **Python** que realiza a auditoria de documentos fiscais **NF-e** e **CT-e** a partir de arquivos **XML**, comparando os valores com uma planilha **Excel** (abas por mês/ano) e gerando um **relatório final em Excel (.xlsx)** com status de conferência.

O objetivo é facilitar conferências fiscais/administrativas, identificando diferenças de **volume** e **valores líquidos** (descontando impostos como ICMS, PIS e COFINS).

---

## 🚀 O que o sistema faz
1. Você seleciona os **XMLs** (NF-e e/ou CT-e)
2. Você seleciona o **Excel** base (com abas do ano/mês alvo)
3. O sistema:
   - Lê os XMLs e extrai: **nota**, **volume**, **bruto**, **ICMS**, **PIS**, **COFINS**
   - Lê o Excel e encontra a linha correspondente pela **NF**
   - Calcula o **líquido** e compara com o Excel
   - Gera um relatório em `.xlsx` com:
     - Diferença de volume
     - Diferença financeira (R$)
     - Status (**OK / ERRO / não encontrado / erro de parse**)
     - Formatação com cores (verde/vermelho)

---

## ✅ Funcionalidades
- 📂 Leitura de XMLs **NF-e** e **CT-e**
- 🧠 Identificação automática do tipo (NF-e / CT-e)
- 🧾 Extração de:
  - Nota (nNF / nCT)
  - Volume (M3/NM3 ou fallback no XML)
  - Bruto (vNF / vTPrest)
  - ICMS, PIS, COFINS
- 📊 Leitura de Excel com abas filtradas por:
  - `ANO_ALVO` (ex: `"25"`)
  - `MESES_ALVO` (ex: `["OUT", "NOV", "DEZ"]`)
- 🧮 Cálculo do **Líquido** (Bruto - impostos válidos)
- 🧾 Ajuste especial para **CT-e** quando não houver PIS/COFINS no XML:
  - usa os valores do Excel para comparar corretamente
- 📄 Geração automática de relatório `.xlsx` com:
  - Cabeçalho formatado
  - Linhas verdes para **OK**
  - Linhas vermelhas para **ERRO**
  - Formatação numérica (R$ e volumes)
- 🖥️ Interface simples por janelas (Tkinter: seleção de arquivos)

---

## 🛠 Tecnologias Utilizadas
- **Python**
- **Pandas**
- **Tkinter**
- **ElementTree (xml.etree.ElementTree)**
- **OpenPyXL** (formatação do relatório Excel)
- **Regex (re)**

---

## ⚙️ Configurações Importantes
No topo do código existem filtros de abas do Excel:

```python
ANO_ALVO = "25"
MESES_ALVO = ["OUT", "NOV", "DEZ"]
✅ O sistema só processa abas que contenham:

o ANO_ALVO no nome (ex: 2025 OUT)

e algum dos meses em MESES_ALVO

📥 Como usar
1) Instalar dependências
pip install pandas openpyxl
O tkinter geralmente já vem com o Python no Windows.

2) Executar
python Auditoria.py
3) Fluxo na tela
Selecione os XMLs (NF-e / CT-e)

Selecione o arquivo Excel (.xlsx)

O relatório será gerado automaticamente e salvo em:

Downloads/Auditoria_XML_<hora>.xlsx

📄 Saída (Relatório)
O relatório final contém colunas como:

Arquivo, Tipo, Mês, Nota

Vol XML / Vol Excel / Diff Vol

Bruto XML, ICMS XML, PIS, COFINS

ICMS Excel, PIS Excel, COFINS Excel

Líq XML (Calc) / Líq Excel / Diff R$

Status e Observações

✅ Status possíveis
OK ✅ → valores dentro da tolerância

ERRO VOL ❌ → volume divergente

ERRO VALOR ❌ → valor líquido divergente

ERRO VOL+VALOR ❌ → ambos divergentes

Ñ ENCONTRADO ⚠️ → não achou a NF no Excel

ERRO PARSE ❌ → falha ao ler o XML

🎯 Regras de tolerância
NF-e: tolerância financeira de R$ 5,00

CT-e: tolerância financeira de R$ 50,00

Volume: diferença < 1.0 (quando houver volume no Excel)

📌 Possíveis Melhorias Futuras
Barra de progresso (UI)

Exportação de relatório em PDF

Log detalhado de processamento

Processamento por pasta (selecionar diretório de XMLs)

Configurar tolerâncias pela interface

Suporte a mais layouts de planilhas

📄 Licença
Este projeto está sob a licença MIT.

👤 Autor
Leandro Morais
GitHub: https://github.com/jleandromorais
