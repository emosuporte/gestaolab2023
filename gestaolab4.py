import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Função para carregar os dados do arquivo Excel
def carregar_dados_arquivo(nome_arquivo):
    try:
        dados = pd.read_excel(nome_arquivo)
        return dados
    except FileNotFoundError:
        return None

# Função para gerar o arquivo docx com base no template
def gerar_documento_docx(template, dados):
    doc = Document(template)

    for coluna in dados.columns:
        for paragrafo in doc.paragraphs:
            for run in paragrafo.runs:
                run.text = run.text.replace(coluna, str(dados[coluna].values[0]))
                run.font.size = Pt(12)
            paragrafo.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    return doc

# Página de registro de novas coletas
def pagina_registro_coletas():
    st.title('Registro de Novas Coletas')

    data = st.date_input('Data')
    motivo = st.text_input('Motivo')
    pedido = st.text_input('Pedido')
    atendimento = st.text_input('Atendimento')
    exames = st.text_input('Exames')

    if st.button('Salvar'):
        dados = pd.DataFrame({'Data': [data], 'Motivo': [motivo], 'Pedido': [pedido], 'Atendimento': [atendimento], 'Exames': [exames]})
        dados.to_excel('registro_coletas.xlsx', index=False)
        st.success('Registro salvo com sucesso!')

# Página de registro de não conformidades
def pagina_registro_nao_conformidades():
    st.title('Registro de Não Conformidades')

    dados = pd.DataFrame()

    numero_registro = st.text_input('Número de Registro', value='1', key='registro')
    data_registro = st.date_input('Data do Registro')
    data_fato = st.date_input('Data do Fato')
    aberto_por = st.text_input('A não conformidade aberta por')
    numero_pedido = st.text_input('Número do Pedido')
    tipo_nao_conformidade = st.selectbox('Tipo de Não Conformidade', ['Coleta: Troca de paciente',
                                                                      'Coleta: Troca de etiquetas',
                                                                      'Coleta: Coleta em tubo inadequado',
                                                                      'Coleta: Material sem identificação do paciente',
                                                                      'Coleta: Material não coleta',
                                                                      'Secretaria: Erro de cadastro',
                                                                      'Secretaria: Troca de etiquetas nos tubos/frascos do mesmo paciente',
                                                                      'Secretaria: Troca de etiquetas nos tubos/frascos de pacientes diferentes',
                                                                      'Área Técnica: Exames não realizados',
                                                                      'Área Técnica: Erro na liberação do exame',
                                                                      'Área Técnica: Controle interno fora das especificações',
                                                                      'Área Técnica: Equipamentos',
                                                                      'Área Técnica: Outro'])
    fato = st.text_area('Descreva o fato')
    acao_corretiva = st.text_area('Ação corretiva imediata')
    responsavel_acao_corretiva = st.text_input('Responsável pela ação corretiva')

    if st.button('Salvar'):
        dados = pd.DataFrame({'Número de Registro': [numero_registro],
                              'Data do Registro': [data_registro],
                              'Data do Fato': [data_fato],
                              'A não conformidade aberta por': [aberto_por],
                              'Número do Pedido': [numero_pedido],
                              'Tipo de Não Conformidade': [tipo_nao_conformidade],
                              'Fato': [fato],
                              'Ação Corretiva Imediata': [acao_corretiva],
                              'Responsável pela Ação Corretiva': [responsavel_acao_corretiva]})

        dados.to_excel('registro_nao_conformidades.xlsx', index=False)

        template = 'template_rnc.docx'
        documento = gerar_documento_docx(template, dados)

        nome_arquivo = f"registro_nao_conformidades_{numero_registro}.docx"
        documento.save(nome_arquivo)

        st.success('Registro salvo com sucesso!')

# Página de indicadores de coletas
def pagina_indicadores_coletas():
    st.title('Indicadores de Coletas')

    dados_coletas = carregar_dados_arquivo('registro_coletas.xlsx')

    if dados_coletas is not None:
        dados_coletas['Data'] = pd.to_datetime(dados_coletas['Data'])
        coletas_por_dia = dados_coletas.groupby(dados_coletas['Data'].dt.date).size()
        coletas_por_mes = dados_coletas.groupby(dados_coletas['Data'].dt.to_period('M')).size()
        coletas_por_ano = dados_coletas.groupby(dados_coletas['Data'].dt.to_period('Y')).size()

        st.subheader('Coletas por Dia')
        st.bar_chart(coletas_por_dia)

        st.subheader('Coletas por Mês')
        st.bar_chart(coletas_por_mes)

        st.subheader('Coletas por Ano')
        st.bar_chart(coletas_por_ano)
    else:
        st.warning('Nenhum dado de coleta encontrado.')

# Página de indicadores de não conformidades
def pagina_indicadores_nao_conformidades():
    st.title('Indicadores de Não Conformidades')

    dados_nao_conformidades = carregar_dados_arquivo('registro_nao_conformidades.xlsx')

    if dados_nao_conformidades is not None:
        dados_nao_conformidades['Data do Registro'] = pd.to_datetime(dados_nao_conformidades['Data do Registro'])
        nao_conformidades_por_dia = dados_nao_conformidades.groupby(dados_nao_conformidades['Data do Registro'].dt.date).size()
        nao_conformidades_por_mes = dados_nao_conformidades.groupby(dados_nao_conformidades['Data do Registro'].dt.to_period('M')).size()
        nao_conformidades_por_ano = dados_nao_conformidades.groupby(dados_nao_conformidades['Data do Registro'].dt.to_period('Y')).size()

        st.subheader('Não Conformidades por Dia')
        st.bar_chart(nao_conformidades_por_dia)

        st.subheader('Não Conformidades por Mês')
        st.bar_chart(nao_conformidades_por_mes)

        st.subheader('Não Conformidades por Ano')
        st.bar_chart(nao_conformidades_por_ano)
    else:
        st.warning('Nenhum dado de não conformidade encontrado.')

# Página principal
def pagina_principal():
    st.title('Gestão do Laboratório')

    st.sidebar.title('Menu')
    opcao = st.sidebar.selectbox('Selecione uma opção', ['Registro de Novas Coletas',
                                                         'Registro de Não Conformidades',
                                                         'Indicadores de Coletas',
                                                         'Indicadores de Não Conformidades'])

    if opcao == 'Registro de Novas Coletas':
        pagina_registro_coletas()
    elif opcao == 'Registro de Não Conformidades':
        pagina_registro_nao_conformidades()
    elif opcao == 'Indicadores de Coletas':
        pagina_indicadores_coletas()
    elif opcao == 'Indicadores de Não Conformidades':
        pagina_indicadores_nao_conformidades()

if __name__ == '__main__':
    pagina_principal()
