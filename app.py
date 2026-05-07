import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Consolidador de Folha de Pagamento", layout="wide")

st.title("📑 Consolidar Folhas de Pagamento")
st.markdown("""
Arraste os arquivos de folha (.xlsx) para consolidar os proventos e descontos por setor e empresa.
""")

COLUNAS_PROCESSADAS = ['EMPRESA', 'COMPETENCIA', 'SETOR', 'TIPO', 'RUBRICA', 'VALOR']
COLUNAS_LAYOUT_EVENTOS_SEPARADOS = {
    'nome_emp',
    'cp_competencia',
    'cp_desr_sep',
    'cp_nome_eve_p',
    'cp_eve_val_p',
    'cp_nome_eve_d',
    'cp_eve_val_d',
}


def normalizar_texto(serie):
    texto = serie.astype("string").str.replace(r"\s+", " ", regex=True).str.strip()
    return texto.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})


def normalizar_setor(serie):
    texto = normalizar_texto(serie)
    return texto.apply(lambda x: x.split('-')[-1].strip().upper() if pd.notna(x) else x)


def normalizar_competencia(serie):
    texto_original = normalizar_texto(serie)
    datas = pd.to_datetime(serie, errors='coerce', dayfirst=True)

    competencia = texto_original.copy()
    competencia.loc[datas.notna()] = datas.loc[datas.notna()].dt.strftime('%d/%m/%Y')
    return competencia


def normalizar_valor(serie):
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors='coerce')

    texto = (
        serie.astype("string")
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
    )
    texto = texto.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    return pd.to_numeric(texto, errors='coerce')


def padronizar_dataframe(df):
    df = df.copy()
    df['EMPRESA'] = normalizar_texto(df['EMPRESA'])
    df['COMPETENCIA'] = normalizar_competencia(df['COMPETENCIA'])
    df['SETOR'] = normalizar_setor(df['SETOR'])
    df['TIPO'] = normalizar_texto(df['TIPO']).str.upper()
    df['RUBRICA'] = normalizar_texto(df['RUBRICA']).str.upper()
    df['VALOR'] = normalizar_valor(df['VALOR'])

    df = df[df['TIPO'].isin(['P', 'D'])].copy()
    df = df.dropna(subset=['EMPRESA', 'COMPETENCIA', 'SETOR', 'RUBRICA', 'VALOR'])
    df = df[~df['RUBRICA'].str.contains('DEPENDENTE.*IRRF.*MENSAL', case=False, regex=True, na=False)]
    df['VALOR'] = np.where(df['TIPO'] == 'D', -df['VALOR'], df['VALOR'])

    return df[COLUNAS_PROCESSADAS]


def extrair_eventos_de_colunas_separadas(df):
    setor_coluna = 'cp_desr_sep' if 'cp_desr_sep' in df.columns else 'cp_depto'
    base_colunas = ['nome_emp', 'cp_competencia', setor_coluna]
    eventos = []

    for tipo, rubrica_coluna, valor_coluna in [
        ('P', 'cp_nome_eve_p', 'cp_eve_val_p'),
        ('D', 'cp_nome_eve_d', 'cp_eve_val_d'),
    ]:
        if rubrica_coluna not in df.columns or valor_coluna not in df.columns:
            continue

        evento = df[base_colunas + [rubrica_coluna, valor_coluna]].copy()
        evento.columns = ['EMPRESA', 'COMPETENCIA', 'SETOR', 'RUBRICA', 'VALOR']
        evento['TIPO'] = tipo
        eventos.append(evento[COLUNAS_PROCESSADAS])

    if not eventos:
        raise ValueError("O arquivo nao possui colunas de proventos/descontos reconhecidas.")

    return pd.concat(eventos, ignore_index=True)


def ler_arquivo_folha(arquivo):
    df = pd.read_excel(arquivo)
    colunas = {str(col).strip() for col in df.columns}

    if COLUNAS_LAYOUT_EVENTOS_SEPARADOS.issubset(colunas):
        return padronizar_dataframe(extrair_eventos_de_colunas_separadas(df))

    if set(COLUNAS_PROCESSADAS).issubset(colunas):
        return padronizar_dataframe(df[COLUNAS_PROCESSADAS])

    raise ValueError(
        "Estrutura de planilha nao reconhecida. O arquivo precisa conter as colunas do exportador de folha "
        "ou as colunas padrao EMPRESA, COMPETENCIA, SETOR, TIPO, RUBRICA e VALOR."
    )


def processar_folhas(arquivos_carregados):
    dados_consolidados = []
    
    for arquivo in arquivos_carregados:
        df = ler_arquivo_folha(arquivo)
        dados_consolidados.append(df)

    if not dados_consolidados:
        raise ValueError("Nenhum arquivo valido foi enviado para processamento.")

    df_completo = pd.concat(dados_consolidados, ignore_index=True)
    df_completo = df_completo.dropna(subset=['RUBRICA'])
    
    proventos_cols = sorted(df_completo[df_completo['TIPO'] == 'P']['RUBRICA'].dropna().unique().tolist())
    descontos_cols = sorted(df_completo[df_completo['TIPO'] == 'D']['RUBRICA'].dropna().unique().tolist())
    
    df_pivot = df_completo.pivot_table(
        index=['EMPRESA', 'SETOR', 'COMPETENCIA'],
        columns='RUBRICA',
        values='VALOR',
        aggfunc='sum'
    ).reset_index()
    
    colunas_ordem = ['EMPRESA', 'SETOR', 'COMPETENCIA'] + proventos_cols + descontos_cols
    df_pivot = df_pivot.reindex(columns=colunas_ordem)
    df_pivot.fillna(0, inplace=True)
    
    df_pivot.insert(3, 'SALDO LIQUIDO', df_pivot[proventos_cols + descontos_cols].sum(axis=1))
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_pivot.to_excel(writer, index=False, startrow=1, sheet_name='Consolidado')
        
        wb = writer.book
        ws = wb['Consolidado']
        
        fill_provento = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fill_desconto = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        fonte_negrito = Font(bold=True)
        alinhamento_centro = Alignment(horizontal="center", vertical="center")
        
        letra_inicio_prov = 5 
        letra_fim_prov = letra_inicio_prov + len(proventos_cols) - 1
        letra_inicio_desc = letra_fim_prov + 1
        letra_fim_desc = letra_inicio_desc + len(descontos_cols) - 1
        
        if len(proventos_cols) > 0:
            celula_prov = ws.cell(row=1, column=letra_inicio_prov, value="PROVENTOS")
            celula_prov.font = fonte_negrito
            celula_prov.fill = fill_provento
            celula_prov.alignment = alinhamento_centro
            ws.merge_cells(start_row=1, start_column=letra_inicio_prov, end_row=1, end_column=letra_fim_prov)
            
        if len(descontos_cols) > 0:
            celula_desc = ws.cell(row=1, column=letra_inicio_desc, value="DESCONTOS")
            celula_desc.font = fonte_negrito
            celula_desc.fill = fill_desconto
            celula_desc.alignment = alinhamento_centro
            ws.merge_cells(start_row=1, start_column=letra_inicio_desc, end_row=1, end_column=letra_fim_desc)
        
        for col in range(1, ws.max_column + 1):
            celula = ws.cell(row=2, column=col)
            celula.font = fonte_negrito
            celula.value = str(celula.value).upper()

    return df_pivot, output.getvalue()

arquivos = st.file_uploader("Suba os arquivos de Folha de Pagamento", type=['xlsx'], accept_multiple_files=True)

if arquivos:
    if st.button("🚀 Gerar Relatório Consolidado"):
        with st.spinner("Processando..."):
            try:
                df_final, excel_binario = processar_folhas(arquivos)
            except Exception as erro:
                st.error(f"Falha ao processar a planilha: {erro}")
            else:
                st.success("Relatório gerado com sucesso!")
                
                st.dataframe(df_final.replace(0, ''), use_container_width=True)
                
                st.download_button(
                    label="📥 Baixar Relatório Consolidado (Excel)",
                    data=excel_binario,
                    file_name="RELATORIO_CONSOLIDADO_MES.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
