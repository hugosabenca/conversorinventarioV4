import streamlit as st
import pandas as pd
import os
import glob
from openpyxl.utils import get_column_letter
from datetime import datetime
import json
import tempfile
import io

# =====================================================================================
# FUNÇÃO UNIFICADA: PROCESSA O ARQUIVO INDEPENDENTE DO TIPO
# =====================================================================================
def processar_arquivo_generico(caminho_csv):
    """
    Lê um arquivo CSV e tenta identificar automaticamente se é PRODUTO ACABADO
    ou BOBINA baseando-se na estrutura dos dados.
    """
    try:
        try:
            df = pd.read_csv(caminho_csv, header=None, encoding='utf-8', dtype=str)
        except UnicodeDecodeError:
            df = pd.read_csv(caminho_csv, header=None, encoding='latin1', dtype=str)
    except Exception as e:
        return None, f"Erro ao ler arquivo: {e}"

    dados_processados = []
    
    if df.empty:
        return pd.DataFrame(), None

    for index, row in df.iterrows():
        if len(row) < 4:
            continue

        col_zero = str(row[0]).strip()
        
        # Filtro de sujeira (ignora o cabeçalho "DATE,..." ou linhas vazias)
        if "date" in col_zero.lower() or not col_zero or not col_zero[0].isdigit():
            continue

        # ====================================================================
        # DETECTOR AUTOMÁTICO DE FORMATO DO CSV (CELULARES/APPS DIFERENTES)
        # ====================================================================
        # Se a primeira coluna tem espaço e é grande (ex: "2026-06-21 11:40:14 -0300")
        if len(col_zero) > 10 and " " in col_zero:
            # FORMATO NOVO (Ex: N6F7)
            partes_dt = col_zero.split(" ")
            dt_leitura_original = partes_dt[0] # Pega só o '2026-06-21'
            hr_leitura = partes_dt[1] if len(partes_dt) > 1 else ""
            
            # O dado lido (CODECONTENT) está na coluna 4 (índice 3)
            dados_lidos = str(row[3]).strip() if len(row) > 3 else ""
            coluna_tipo = str(row[4]).strip() if len(row) > 4 else ""
            
            # Converte a data de YYYY-MM-DD para DD/MM/YYYY
            try:
                obj_date = datetime.strptime(dt_leitura_original, '%Y-%m-%d')
                dt_formatada = obj_date.strftime('%d/%m/%Y')
            except:
                dt_formatada = dt_leitura_original

        else:
            # FORMATO ANTIGO (Ex: N6F4)
            dt_leitura_original = col_zero
            hr_leitura = str(row[1]).strip() if len(row) > 1 else ""
            
            # O dado lido (CODECONTENT) está na coluna 5 (índice 4)
            coluna_tipo = str(row[3]).strip() if len(row) > 3 else ""
            dados_lidos = str(row[4]).strip() if len(row) > 4 else ""

            # Converte a data de MM-DD-YYYY para DD/MM/YYYY
            try:
                obj_date = datetime.strptime(dt_leitura_original, '%m-%d-%Y')
                dt_formatada = obj_date.strftime('%d/%m/%Y')
            except:
                dt_formatada = dt_leitura_original


        # Estrutura padrão da linha
        nova_linha = {
            "Data da Leitura": dt_formatada,
            "Hora da Leitura": hr_leitura,
            "Filial": None,
            "Código": None,
            "Armazém": None,
            "Lote": None,
            "Peso": None,
            "Localização": os.path.splitext(os.path.basename(caminho_csv))[0]
        }

        # ====================================================================
        # TESTE 1: É PRODUTO ACABADO CLÁSSICO? (Padrão: "XXX-XXX - YYY")
        # ====================================================================
        processado_como_pa = False
        if " -" in dados_lidos:
            try:
                partes_maiores = dados_lidos.split(" -", 1)
                
                parte_esq = partes_maiores[0].split("-")
                filial = parte_esq[0].strip() if len(parte_esq) > 0 else ""
                codigo = parte_esq[1].strip() if len(parte_esq) > 1 else ""
                
                parte_dir = partes_maiores[1].split("-") if len(partes_maiores) > 1 else []
                
                armazem = parte_dir[0].strip() if len(parte_dir) > 0 else ""
                lote = parte_dir[1].strip() if len(parte_dir) > 1 else ""
                peso_str = parte_dir[2].strip() if len(parte_dir) > 2 else "0"
                
                try:
                    peso_val = float(peso_str) / 1000.0
                except:
                    peso_val = None

                nova_linha["Filial"] = filial
                nova_linha["Código"] = codigo
                nova_linha["Armazém"] = armazem
                nova_linha["Lote"] = lote
                nova_linha["Peso"] = peso_val
                
                dados_processados.append(nova_linha)
                processado_como_pa = True
            except Exception:
                pass
        
        if processado_como_pa:
            continue

        # ====================================================================
        # TESTE 2: É BOBINA? (Ou Bobina com estrutura completa)
        # ====================================================================
        
        lote_b = "erro"
        peso_b = None 

        # CASO 1: CODE128 (Asteriscos)
        if coluna_tipo == 'Code128' or '*' in dados_lidos:
            if ' ' in dados_lidos:
                 lote_b, peso_b = "erro de leitura", None
            elif '*' in dados_lidos:
                try:
                    partes = dados_lidos.split('*')
                    if dados_lidos.startswith('*'): 
                        lote_b = partes[3].strip()
                        peso_b = float(partes[2].strip()) / 1000.0
                    else: 
                        lote_b = partes[2].strip()
                        peso_b = float(partes[1].strip()) / 1000.0
                except:
                    lote_b, peso_b = "erro Code128/*", None
            elif dados_lidos.isdigit() and len(dados_lidos) <= 5:
                 try:
                    peso_b, lote_b = float(dados_lidos)/1000.0, ""
                 except:
                    peso_b, lote_b = None, dados_lidos
            else:
                 lote_b, peso_b = dados_lidos, None

        # CASO 2: QR CODE / DATAMATRIX
        elif coluna_tipo in ['QR_CODE', 'QR', 'CODE_39', 'CODE_128'] or '{' in dados_lidos or ',' in dados_lidos:
            
            # 2.1 JSON
            if '{' in dados_lidos and '}' in dados_lidos:
                try:
                    partes = dados_lidos.split('{', 1)
                    identificador = partes[0].strip('"-')
                    dados_json = json.loads('{' + partes[1])
                    peso_b = float(dados_json.get('peso', 0))
                    lote_b = identificador
                except:
                    lote_b = "erro QR/JSON"

            # 2.2 NOVO FORMATO COM VÍRGULA
            elif ',' in dados_lidos and '-' in dados_lidos:
                try:
                    partes_virgula = dados_lidos.split(',')
                    
                    if len(partes_virgula) > 1 and partes_virgula[-1].replace('.', '', 1).isdigit():
                        peso_decimal = partes_virgula[-1].strip()
                        
                        parte_lote_completa = ','.join(partes_virgula[:-1])
                        partes_hifen = parte_lote_completa.split('-')
                        
                        # Extração de Filial/Código/Armazém se disponível
                        if len(partes_hifen) >= 4:
                            nova_linha["Filial"] = partes_hifen[0].strip()
                            nova_linha["Código"] = partes_hifen[1].strip()
                            nova_linha["Armazém"] = partes_hifen[2].strip()

                        lote_b = partes_hifen[-2].strip()
                        
                        peso_completo_str = f"{partes_hifen[-1].strip()},{peso_decimal}"
                        peso_b = float(peso_completo_str.replace(',', '.'))
                        
                    else:
                        raise ValueError("Formato virgula invalido")
                except:
                    try:
                        partes = dados_lidos.split('-')
                        lote_b = partes[3].strip()
                        peso_b = float(partes[-1].strip()) / 1000.0
                    except:
                        lote_b = "erro QR/FormatoVirgula"
            
            # 2.3 Formato Antigo (Só hifens)
            else:
                 try:
                     partes = dados_lidos.split('-')
                     if len(partes) >= 4:
                        lote_b = partes[3].strip() 
                        peso_b = float(partes[-1].strip()) / 1000.0
                     else:
                        lote_b = dados_lidos
                        peso_b = None
                 except:
                     lote_b = dados_lidos
                     peso_b = None

        else:
            lote_b = dados_lidos
            peso_b = None 
        
        nova_linha["Lote"] = lote_b
        nova_linha["Peso"] = peso_b
        dados_processados.append(nova_linha)

    return pd.DataFrame(dados_processados), None

# =====================================================================================
# LEITOR DE MÃO (Tomate MDK 103S): FUNÇÕES DE PROCESSAMENTO
# =====================================================================================
PREFIXO_LOC = "LOC:"
COR_DUPLICADO = "FFF6CC"   # amarelo
COR_SEM_LOC = "FFE8CC"     # laranja
COR_ERRO = "FFD9D9"        # vermelho

def interpretar_leitura_mao(texto):
    """
    Recebe o texto de UMA leitura e devolve um dicionário com o tipo:
      - {"tipo": "local", "local": "N4F1"}
      - {"tipo": "item", "Filial":..., "Código":..., "Armazém":..., "Lote":..., "Peso":...}
      - {"tipo": "erro", "motivo": "..."}
    Padrão do item: FILIAL-PRODUTO -ARMAZÉM-LOTE-PESO
    """
    texto = str(texto).strip()

    # É uma localização? (QR com prefixo LOC:)
    if texto.upper().startswith(PREFIXO_LOC):
        local = texto[len(PREFIXO_LOC):].strip()
        if local:
            return {"tipo": "local", "local": local}
        return {"tipo": "erro", "motivo": "QR de localização sem nome"}

    # É um item? Precisa ter o " -" que separa o produto do armazém
    if " -" not in texto:
        return {"tipo": "erro", "motivo": f'Fora do padrão: "{texto}"'}

    lado_esq, lado_dir = texto.split(" -", 1)
    partes_esq = lado_esq.split("-", 1)
    partes_dir = lado_dir.split("-")

    if len(partes_esq) < 2 or len(partes_dir) < 3:
        return {"tipo": "erro", "motivo": f'Fora do padrão: "{texto}"'}

    filial = partes_esq[0].strip()
    codigo = partes_esq[1].strip()
    armazem = partes_dir[0].strip()
    lote = "-".join(partes_dir[1:-1]).strip()   # tudo entre o armazém e o peso
    peso_txt = partes_dir[-1].strip().replace(",", ".")

    try:
        peso = float(peso_txt) / 1000.0
    except ValueError:
        return {"tipo": "erro", "motivo": f'Peso inválido: "{texto}"'}

    if not all([filial, codigo, armazem, lote]):
        return {"tipo": "erro", "motivo": f'Campo vazio: "{texto}"'}

    return {"tipo": "item", "Filial": filial, "Código": codigo, "Armazém": armazem,
            "Lote": lote, "Peso": peso}


def processar_planilha_leitor(arquivo):
    """
    Lê a planilha gerada pelo leitor de mão (todas as abas, célula por célula,
    de cima para baixo) e devolve um DataFrame com os itens já separados.
    Cada aba/arquivo começa SEM localização (não herda do anterior).
    """
    abas = pd.read_excel(arquivo, sheet_name=None, header=None, dtype=str)
    linhas = []

    for _, df in abas.items():
        local_atual = None
        for _, row in df.iterrows():
            for valor in row:
                if pd.isna(valor) or not str(valor).strip():
                    continue
                r = interpretar_leitura_mao(valor)

                if r["tipo"] == "local":
                    local_atual = r["local"]
                    continue

                if r["tipo"] == "erro":
                    linhas.append({"Filial": "", "Código": "", "Armazém": "", "Lote": "",
                                   "Peso": None, "Localização": local_atual or "SEM LOCALIZAÇÃO",
                                   "Observação": r["motivo"], "_status": "erro"})
                    continue

                linhas.append({"Filial": r["Filial"], "Código": r["Código"],
                               "Armazém": r["Armazém"], "Lote": r["Lote"], "Peso": r["Peso"],
                               "Localização": local_atual or "SEM LOCALIZAÇÃO",
                               "Observação": "" if local_atual else "Lido antes de qualquer QR de localização",
                               "_status": "ok" if local_atual else "semloc"})

    return pd.DataFrame(linhas, columns=["Filial", "Código", "Armazém", "Lote", "Peso",
                                         "Localização", "Observação", "_status"])


def marcar_duplicados(df):
    """Marca os lotes lidos mais de uma vez (todas as ocorrências)."""
    df = df.copy()
    df["_dup"] = False
    itens = df["_status"] != "erro"
    contagem = df.loc[itens, "Lote"].value_counts()
    repetidos = contagem[contagem > 1]

    for lote, qtd in repetidos.items():
        mask = itens & (df["Lote"] == lote)
        df.loc[mask, "_dup"] = True
        aviso = f"Lote duplicado ({qtd} leituras)"
        df.loc[mask, "Observação"] = df.loc[mask, "Observação"].apply(
            lambda o: f"{o} | {aviso}" if o else aviso)
    return df


def cor_da_linha(status, dup):
    """Prioridade: erro (vermelho) > sem localização (laranja) > duplicado (amarelo)."""
    if status == "erro":
        return COR_ERRO
    if status == "semloc":
        return COR_SEM_LOC
    if dup:
        return COR_DUPLICADO
    return None


def resumo_por_local(df):
    """Quantidade de itens por localização, na ordem em que foram lidos."""
    itens = df[df["_status"] != "erro"]
    if itens.empty:
        return pd.DataFrame(columns=["Localização", "Itens"])
    ordem = list(dict.fromkeys(itens["Localização"]))
    if "SEM LOCALIZAÇÃO" in ordem:
        ordem.remove("SEM LOCALIZAÇÃO")
        ordem.insert(0, "SEM LOCALIZAÇÃO")
    qtd = itens["Localização"].value_counts()
    return pd.DataFrame({"Localização": ordem, "Itens": [int(qtd[l]) for l in ordem]})


def gerar_excel_leitor(df, df_locais, data_inventario=""):
    """Monta o Excel final: aba 'Inventario Geral' (com cores) + aba 'Itens por Localização' + aba 'Info'."""
    from openpyxl.styles import PatternFill, Font

    colunas = ["Filial", "Código", "Armazém", "Lote", "Peso", "Localização", "Observação"]
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df[colunas].to_excel(writer, index=False, sheet_name="Inventario Geral")
        ws = writer.sheets["Inventario Geral"]

        # Cabeçalho em negrito + congelado
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"

        for i, (status, dup) in enumerate(zip(df["_status"], df["_dup"]), start=2):
            # Colunas de texto ficam como TEXTO (mantém zeros à esquerda)
            for col in (1, 2, 3, 4, 6, 7):
                ws.cell(row=i, column=col).number_format = "@"
            ws.cell(row=i, column=5).number_format = "0.000"

            cor = cor_da_linha(status, dup)
            if cor:
                fill = PatternFill(start_color=cor, end_color=cor, fill_type="solid")
                for col in range(1, len(colunas) + 1):
                    ws.cell(row=i, column=col).fill = fill

        for col in ws.columns:
            max_len = max((len(str(c.value)) for c in col if c.value is not None), default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 60)

        # Aba 2: itens por localização
        df_locais.to_excel(writer, index=False, sheet_name="Itens por Localização")
        ws2 = writer.sheets["Itens por Localização"]
        for cell in ws2[1]:
            cell.font = Font(bold=True)
        ws2.column_dimensions["A"].width = 24
        ws2.column_dimensions["B"].width = 10
        for i, local in enumerate(df_locais["Localização"], start=2):
            if local == "SEM LOCALIZAÇÃO":
                fill = PatternFill(start_color=COR_SEM_LOC, end_color=COR_SEM_LOC, fill_type="solid")
                ws2.cell(row=i, column=1).fill = fill
                ws2.cell(row=i, column=2).fill = fill
        total = len(df_locais) + 2
        ws2.cell(row=total, column=1, value="Total").font = Font(bold=True)
        ws2.cell(row=total, column=2, value=int(df_locais["Itens"].sum())).font = Font(bold=True)

        # Aba 3: informações do inventário (usada pela aba Ordem de Carregamento)
        pd.DataFrame({"Data do inventário": [data_inventario],
                      "Gerado em": [datetime.now().strftime("%d/%m/%Y %H:%M")]}
                     ).to_excel(writer, index=False, sheet_name="Info")
        ws3 = writer.sheets["Info"]
        ws3.column_dimensions["A"].width = 20
        ws3.column_dimensions["B"].width = 20

    output.seek(0)
    return output


# =====================================================================================
# ETIQUETAS DE LOCALIZAÇÃO: QR CODE + PDF
# =====================================================================================
LAYOUTS_ETIQUETA = {
    "1 por folha (A4 inteira)": (1, 1),
    "2 por folha (1 x 2)": (1, 2),
    "6 por folha (2 x 3)": (2, 3),
    "12 por folha (3 x 4)": (3, 4),
}

def gerar_qr_png(texto, escala=10):
    """Gera o QR Code em PNG (bytes)."""
    import segno
    buf = io.BytesIO()
    segno.make_qr(texto, error="m").save(buf, kind="png", scale=escala, border=4)
    buf.seek(0)
    return buf


def limpar_lista_locais(texto):
    """Uma localização por linha, sem vazios, sem repetidos, sem o prefixo LOC: se alguém digitou."""
    locais = []
    for linha in texto.splitlines():
        nome = linha.strip()
        if nome.upper().startswith(PREFIXO_LOC):
            nome = nome[len(PREFIXO_LOC):].strip()
        if nome and nome not in locais:
            locais.append(nome)
    return locais


def gerar_pdf_etiquetas(locais, colunas, linhas):
    """Monta o PDF A4 com as etiquetas (QR com 'LOC:' + nome grande embaixo)."""
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader

    largura, altura = A4
    margem = 12 * 2.835  # 12 mm em pontos
    cel_w = (largura - 2 * margem) / colunas
    cel_h = (altura - 2 * margem) / linhas
    por_pagina = colunas * linhas

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle("Etiquetas de Localização")

    for idx, local in enumerate(locais):
        pos = idx % por_pagina
        if idx > 0 and pos == 0:
            c.showPage()
        col = pos % colunas
        lin = pos // colunas
        x = margem + col * cel_w
        y = altura - margem - (lin + 1) * cel_h

        # Borda tracejada para recorte
        c.setDash(4, 3)
        c.setStrokeColorRGB(0.6, 0.6, 0.6)
        c.rect(x + 3, y + 3, cel_w - 6, cel_h - 6)
        c.setDash()

        # Tamanho do texto proporcional à etiqueta
        tam_fonte = max(14, min(cel_w, cel_h) * 0.16)
        espaco_texto = tam_fonte * 1.9

        # QR o maior possível dentro da etiqueta
        lado_qr = min(cel_w * 0.80, cel_h - espaco_texto - 20)
        qr_x = x + (cel_w - lado_qr) / 2
        qr_y = y + espaco_texto + (cel_h - espaco_texto - lado_qr) / 2
        c.drawImage(ImageReader(gerar_qr_png(PREFIXO_LOC + local)), qr_x, qr_y, lado_qr, lado_qr)

        # Nome da localização embaixo
        c.setFont("Helvetica-Bold", tam_fonte)
        c.setFillColorRGB(0.15, 0.15, 0.2)
        c.drawCentredString(x + cel_w / 2, y + tam_fonte * 0.9, local)

    c.save()
    buf.seek(0)
    return buf


# =====================================================================================
# INTERFACE DO STREAMLIT (UI)
# =====================================================================================

st.set_page_config(page_title="Conversor de Inventário Dox", layout="wide")
st.title("Conversor de Inventário Unificado")
st.markdown("---")

aba_celular, aba_mao, aba_etiquetas = st.tabs(["📱 App Celular", "🔫 Leitor de Mão", "🏷️ Etiquetas de Localização"])

# =====================================================================================
# ABA 1: APP CELULAR (sem alterações)
# =====================================================================================
with aba_celular:

    # --- LÓGICA DO BOTÃO LIMPAR (SESSION STATE) ---
    if 'uploader_key' not in st.session_state:
        st.session_state.uploader_key = 0

    def limpar_lista():
        st.session_state.uploader_key += 1

    col1, col2 = st.columns([2, 1])

    with col1:
        # O parametro 'key' muda cada vez que clicamos em Limpar, resetando o widget
        uploaded_files = st.file_uploader(
            "Importar arquivos .csv (Aceita Bobina e Produto Acabado misturados)",
            type="csv",
            accept_multiple_files=True,
            key=str(st.session_state.uploader_key)
        )
        
        # --- FRASE DE CONTAGEM E BOTÃO LIMPAR ---
        if uploaded_files:
            qtd_arquivos = len(uploaded_files)
            col_msg, col_btn = st.columns([3, 1])
            with col_msg:
                st.info(f"📂 **{qtd_arquivos} arquivos anexados.**")
            with col_btn:
                st.button("🗑️ Limpar Lista", on_click=limpar_lista, type="secondary", use_container_width=True)

    with col2:
        st.info("Configurações de Saída")
        nome_arquivo_usuario = st.text_input("Nome do Arquivo Final (sem .xlsx):", value="")

    # --- BOTÃO CONVERTER ---
    st.markdown("###") # Espaçamento
    if st.button("Converter Arquivos", type="primary"):
        if not uploaded_files:
            st.warning("⚠️ Por favor, carregue pelo menos um arquivo .csv.")
        else:
            with st.spinner("Processando..."):
                try:
                    todos_dfs = []
                    for uploaded_file in uploaded_files:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp_file:
                            tmp_file.write(uploaded_file.getbuffer())
                            tmp_path = tmp_file.name
                        
                        df_temp, erro = processar_arquivo_generico(tmp_path)
                        
                        if erro:
                            st.error(f"Erro no arquivo {uploaded_file.name}: {erro}")
                        elif not df_temp.empty:
                            df_temp["Localização"] = uploaded_file.name.replace('.csv', '')
                            todos_dfs.append(df_temp)
                        
                        os.unlink(tmp_path)

                    if todos_dfs:
                        df_final = pd.concat(todos_dfs, ignore_index=True)
                        
                        # Tratamento estético
                        if "Armazém" in df_final.columns:
                            df_final["Armazém"] = df_final["Armazém"].fillna('').apply(lambda x: str(x).split('.')[0].zfill(2) if str(x).replace('.','').isdigit() else str(x))
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_final.to_excel(writer, index=False, sheet_name='Inventario Geral')
                            ws = writer.sheets['Inventario Geral']
                            for col in ws.columns:
                                max_len = max((len(str(cell.value)) for cell in col if cell.value is not None), default=0)
                                ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 4
                                
                                if col[0].value == "Peso":
                                    for cell in col[1:]:
                                        if cell.value is not None:
                                            cell.number_format = '0.000'

                        output.seek(0)
                        nome_download = f"{nome_arquivo_usuario.strip()}.xlsx" if nome_arquivo_usuario.strip() else "Inventario.xlsx"

                        st.success(f"✅ Sucesso! {len(todos_dfs)} arquivos processados.")
                        st.download_button(
                            label="📥 Baixar Excel Consolidado",
                            data=output,
                            file_name=nome_download,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("Os arquivos foram lidos, mas nenhum dado válido foi encontrado.")

                except Exception as e:
                    st.error(f"Ocorreu um erro crítico: {e}")

# =====================================================================================
# ABA 2: LEITOR DE MÃO
# =====================================================================================
with aba_mao:
    if 'uploader_key_mao' not in st.session_state:
        st.session_state.uploader_key_mao = 0

    def limpar_lista_mao():
        st.session_state.uploader_key_mao += 1
        st.session_state.pop("resultado_mao", None)

    with st.expander("ℹ️ Como usar o Leitor de Mão", expanded=False):
        st.markdown(
            "1. Abra uma planilha Excel em branco e clique na célula **A1**.\n"
            "2. Ao chegar em cada local, **leia primeiro o QR da localização** (ex: `LOC:N4F1`) "
            "e depois todas as etiquetas daquele local.\n"
            "3. Ao terminar, salve a planilha e anexe abaixo.\n\n"
            "Formato esperado da etiqueta: `FILIAL-PRODUTO -ARMAZÉM-LOTE-PESO`\n\n"
            "🎨 **Legenda das cores:** 🟨 lote lido mais de uma vez · "
            "🟧 item sem localização · 🟥 leitura fora do padrão"
        )

    col1_m, col2_m = st.columns([2, 1])

    with col1_m:
        arquivos_mao = st.file_uploader(
            "Importar planilhas do leitor (.xlsx)",
            type="xlsx",
            accept_multiple_files=True,
            key=f"mao_{st.session_state.uploader_key_mao}",
            help="Planilhas Excel geradas pelo leitor de mão. Pode anexar várias de uma vez. "
                 "Cada planilha começa sem localização: leia o QR do local antes dos itens."
        )
        if arquivos_mao:
            col_msg_m, col_btn_m = st.columns([3, 1])
            with col_msg_m:
                st.info(f"📂 **{len(arquivos_mao)} arquivo(s) anexado(s).**")
            with col_btn_m:
                st.button("🗑️ Limpar Lista", on_click=limpar_lista_mao, type="secondary",
                          use_container_width=True, key="limpar_mao")

    with col2_m:
        st.info("Configurações de Saída")
        nome_arquivo_mao = st.text_input(
            "Nome do Arquivo Final (sem .xlsx):", value="", key="nome_mao",
            help='Se ficar em branco, o arquivo será salvo como "Inventario_LeitorMao.xlsx".'
        )
        data_inventario = st.date_input(
            "Data do inventário:", value=datetime.now().date(), format="DD/MM/YYYY", key="data_inv",
            help="Dia em que as leituras foram feitas. Fica gravada no Excel e aparece no Roteiro de "
                 "Separação da aba Ordem de Carregamento."
        )

    st.markdown("###")
    if st.button("Converter Planilhas", type="primary", key="converter_mao"):
        if not arquivos_mao:
            st.warning("⚠️ Por favor, carregue pelo menos uma planilha .xlsx.")
        else:
            with st.spinner("Processando..."):
                try:
                    dfs = []
                    for arq in arquivos_mao:
                        try:
                            df_arq = processar_planilha_leitor(arq)
                            if df_arq.empty:
                                st.warning(f"⚠️ O arquivo **{arq.name}** não tem nenhuma leitura.")
                            else:
                                dfs.append(df_arq)
                        except Exception as e:
                            st.error(f"Erro ao ler o arquivo {arq.name}: {e}")

                    if dfs:
                        df_mao = marcar_duplicados(pd.concat(dfs, ignore_index=True))
                        df_locais = resumo_por_local(df_mao)
                        nome = nome_arquivo_mao.strip() or "Inventario_LeitorMao"
                        st.session_state.resultado_mao = {
                            "df": df_mao,
                            "locais": df_locais,
                            "excel": gerar_excel_leitor(df_mao, df_locais,
                                                        data_inventario.strftime("%d/%m/%Y")).getvalue(),
                            "nome": f"{nome}.xlsx",
                            "qtd_arquivos": len(dfs),
                        }
                    else:
                        st.session_state.pop("resultado_mao", None)
                except Exception as e:
                    st.error(f"Ocorreu um erro crítico: {e}")

    # --- RESULTADO (fica na tela mesmo depois de clicar em baixar) ---
    res = st.session_state.get("resultado_mao")
    if res:
        df_mao = res["df"]
        st.success(f"✅ Conversão concluída! {res['qtd_arquivos']} arquivo(s) processado(s).")

        st.subheader("📊 Resumo da conversão")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Itens lidos", len(df_mao),
                  help="Total de leituras de etiquetas (sem contar os QR de localização).")
        m2.metric("Localizações", int((res["locais"]["Localização"] != "SEM LOCALIZAÇÃO").sum()),
                  help="Quantidade de locais diferentes lidos (QR com LOC:).")
        m3.metric("Lotes duplicados", int(df_mao.loc[df_mao["_dup"], "Lote"].nunique()),
                  help="Lotes lidos mais de uma vez. Todas as leituras repetidas ficam em amarelo.")
        m4.metric("Sem localização", int((df_mao["_status"] == "semloc").sum()),
                  help="Itens lidos antes de qualquer QR de localização na planilha. Ficam em laranja.")
        m5.metric("Leituras com erro", int((df_mao["_status"] == "erro").sum()),
                  help="Leituras fora do padrão FILIAL-PRODUTO -ARMAZÉM-LOTE-PESO. Ficam em vermelho.")

        st.markdown("#### 📍 Itens por localização")
        st.info("💡 Confira se as quantidades fazem sentido. Um local com muito mais itens que o normal, "
                "ou um local que deveria aparecer e não está na lista, pode indicar que alguém "
                "esqueceu de bipar a placa de localização.")
        df_loc = res["locais"]
        st.dataframe(
            df_loc.style.apply(
                lambda r: ["background-color: #FFE8CC" if r["Localização"] == "SEM LOCALIZAÇÃO" else ""] * len(r),
                axis=1),
            column_config={"Itens": st.column_config.ProgressColumn(
                "Itens", format="%d", min_value=0,
                max_value=int(df_loc["Itens"].max()) if not df_loc.empty else 1)},
            hide_index=True, use_container_width=False, width=520,
        )
        st.caption(f"Total: {int(df_loc['Itens'].sum())} itens. As leituras com erro não entram nesta "
                   "tabela. Os locais aparecem na ordem em que foram lidos.")

        st.markdown("#### 👀 Prévia do arquivo")
        st.caption("🟨 Lote lido mais de uma vez · 🟧 Sem localização · 🟥 Leitura fora do padrão")
        colunas_vis = ["Filial", "Código", "Armazém", "Lote", "Peso", "Localização", "Observação"]

        def _colorir(row):
            cor = cor_da_linha(df_mao.at[row.name, "_status"], df_mao.at[row.name, "_dup"])
            return [f"background-color: #{cor}" if cor else ""] * len(row)

        df_vis = df_mao[colunas_vis].copy()
        df_vis["Peso"] = df_vis["Peso"].apply(lambda v: "" if pd.isna(v) else f"{v:.3f}".replace(".", ","))
        st.dataframe(df_vis.style.apply(_colorir, axis=1), hide_index=True, use_container_width=True)
        st.caption("O Excel baixado sai com as mesmas cores, a coluna Observação e uma segunda aba "
                   "com os itens por localização.")

        st.download_button(
            label="📥 Baixar Excel Consolidado",
            data=res["excel"],
            file_name=res["nome"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="baixar_mao",
        )

# =====================================================================================
# ABA 3: ETIQUETAS DE LOCALIZAÇÃO
# =====================================================================================
with aba_etiquetas:
    st.info("Gere aqui os QR Codes das localizações para imprimir e fixar nos locais. "
            "Cada QR já sai com o prefixo `LOC:`, que o sistema usa para reconhecer que a "
            "leitura é um local e não um item.")

    col1_e, col2_e = st.columns([2, 1])
    with col1_e:
        texto_locais = st.text_area(
            "Localizações (uma por linha)", height=180, key="lista_locais",
            placeholder="N4F1\nN4F2\nN4F3",
            help="Digite ou cole as localizações, uma por linha. Pode colar direto de uma coluna "
                 "do Excel. Linhas em branco e nomes repetidos são ignorados."
        )
    with col2_e:
        layout_escolhido = st.selectbox(
            "Etiquetas por folha A4", list(LAYOUTS_ETIQUETA.keys()), index=2, key="layout_etq",
            help="Quanto menos etiquetas por folha, maior o QR. Etiquetas maiores são mais fáceis "
                 "de ler à distância."
        )
        nome_pdf = st.text_input("Nome do PDF (sem .pdf)", value="", key="nome_pdf",
                                 help='Se ficar em branco, será salvo como "Etiquetas_Localizacao.pdf".')

    if st.button("Gerar Etiquetas", type="primary", key="gerar_etq"):
        locais = limpar_lista_locais(texto_locais)
        if not locais:
            st.warning("⚠️ Digite pelo menos uma localização.")
            st.session_state.pop("resultado_etq", None)
        else:
            with st.spinner("Gerando etiquetas..."):
                cols_l, lins_l = LAYOUTS_ETIQUETA[layout_escolhido]
                paginas = -(-len(locais) // (cols_l * lins_l))
                st.session_state.resultado_etq = {
                    "locais": locais,
                    "pdf": gerar_pdf_etiquetas(locais, cols_l, lins_l).getvalue(),
                    "nome": f"{nome_pdf.strip() or 'Etiquetas_Localizacao'}.pdf",
                    "paginas": paginas,
                }

    res_e = st.session_state.get("resultado_etq")
    if res_e:
        st.success(f"✅ {len(res_e['locais'])} etiqueta(s) gerada(s) em {res_e['paginas']} página(s).")
        st.download_button("📥 Baixar PDF para Impressão", data=res_e["pdf"],
                           file_name=res_e["nome"], mime="application/pdf", key="baixar_etq")

        st.markdown("#### 👀 Prévia das etiquetas")
        LIMITE_PREVIA = 8
        mostrar = res_e["locais"][:LIMITE_PREVIA]
        cols_prev = st.columns(4)
        for i, local in enumerate(mostrar):
            with cols_prev[i % 4]:
                st.image(gerar_qr_png(PREFIXO_LOC + local, escala=6), width=150)
                st.markdown(f"<div style='font-size:24px;font-weight:700;margin-top:-8px'>{local}</div>",
                            unsafe_allow_html=True)
        if len(res_e["locais"]) > LIMITE_PREVIA:
            st.caption(f"Mostrando {LIMITE_PREVIA} de {len(res_e['locais'])}. O PDF contém todas.")                    