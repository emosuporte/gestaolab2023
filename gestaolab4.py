import streamlit as st
import pandas as pd
from datetime import date
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx2pdf import convert
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Função para carregar os dados do arquivo Excel
def carregar_dados_arquivo(nome_arquivo):
    try:
        dados = pd.read_excel(nome_arquivo)
        return dados
    except FileNotFoundError:
        return None

# Função para gerar o número automático de registro
def gerar_numero_registro():
    dados_nao_conformidades = carregar_dados_arquivo('registro_nao_conformidades.xlsx')
    if dados_nao_conformidades is not None:
        if 'Número de Registro' in dados_nao_conformidades.columns:
            ultimo_registro = dados_nao_conformidades['Número de Registro'].max()
            if pd.isnull(ultimo_registro):
                numero_registro = 1
            else:
                numero_registro = int(ultimo_registro) + 1
        else:
            numero_registro = 1
    else:
        numero_registro = 1
    return numero_registro



# Função para gerar o arquivo docx com base no template
def gerar_documento_docx(template, dados):
    doc = Document(template)

    for coluna in dados.columns:
        for paragrafo in doc.paragraphs:
            for run in paragrafo.runs:
                run.text = run.text.replace(f"{{{{{coluna}}}}}", str(dados[coluna].values[0]))
                run.font.size = Pt(12)
            paragrafo.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    return doc

# Função para gerar o arquivo PDF
def gerar_documento_pdf(template, dados, numero_registro):
    doc = gerar_documento_docx(template, dados)
    nome_arquivo_temp = f'registro_nao_conformidades_{numero_registro}_{date.today().strftime("%Y%m%d_%H%M%S")}.docx'
    doc.save(nome_arquivo_temp)
    convert(nome_arquivo_temp, f'registro_nao_conformidades_{numero_registro}_{date.today().strftime("%Y%m%d_%H%M%S")}.pdf')
    return f'registro_nao_conformidades_{numero_registro}_{date.today().strftime("%Y%m%d_%H%M%S")}.pdf'

# Função para enviar o email com o PDF anexado
def enviar_email(destinatario, assunto, corpo, nome_arquivo_pdf):
    remetente = 'app.ezzap@gmail.com'  # Insira seu endereço de email
    senha = 'krfknvlfkdpnirmi'  # Insira a senha do seu email
    servidor_smtp = 'smtp.gmail.com'
    porta_smtp = 587

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto

    corpo_email = MIMEText(corpo, 'plain')
    msg.attach(corpo_email)

    # Anexar o PDF ao email
    with open(nome_arquivo_pdf, 'rb') as arquivo_pdf:
        anexo = MIMEBase('application', 'pdf')
        anexo.set_payload(arquivo_pdf.read())
        encoders.encode_base64(anexo)
        anexo.add_header('Content-Disposition', f'attachment; filename={nome_arquivo_pdf}')
        msg.attach(anexo)

    # Enviar o email
    try:
        smtp_obj = smtplib.SMTP(servidor_smtp, porta_smtp)
        smtp_obj.starttls()
        smtp_obj.login(remetente, senha)
        smtp_obj.sendmail(remetente, destinatario, msg.as_string())
        smtp_obj.quit()
        st.success('Email enviado com sucesso!')
    except Exception as e:
        st.error(f'Ocorreu um erro ao enviar o email: {str(e)}')

# Função para página de registro de novas coletas
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

# Função para página de registro de não conformidades
def pagina_registro_nao_conformidades():
    st.title('Registro de Não Conformidades')

    numero_registro = gerar_numero_registro()
    data_registro = date.today()
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
        novo_registro = pd.DataFrame({'Número de Registro': [numero_registro],
                                      'Data do Registro': [data_registro],
                                      'Data do Fato': [data_fato],
                                      'A não conformidade aberta por': [aberto_por],
                                      'Número do Pedido': [numero_pedido],
                                      'Tipo de Não Conformidade': [tipo_nao_conformidade],
                                      'Fato': [fato],
                                      'Ação Corretiva Imediata': [acao_corretiva],
                                      'Responsável pela Ação Corretiva': [responsavel_acao_corretiva]})

        dados_nao_conformidades = carregar_dados_arquivo('registro_nao_conformidades.xlsx')

        if dados_nao_conformidades is not None:
            dados_nao_conformidades = pd.concat([dados_nao_conformidades, novo_registro], ignore_index=True)
        else:
            dados_nao_conformidades = novo_registro

        dados_nao_conformidades.to_excel('registro_nao_conformidades.xlsx', index=False)

        template = 'template_rnc.docx'
        nome_arquivo_pdf = gerar_documento_pdf(template, novo_registro, numero_registro)

        st.success('Registro salvo com sucesso!')
        st.markdown(f"Baixe o arquivo PDF: [registro_nao_conformidades_{numero_registro}_{date.today().strftime('%Y%m%d_%H%M%S')}.pdf]({nome_arquivo_pdf})")

        # Enviar o email com o PDF anexado
        destinatario = 'emo.suporte@gmail.com'  # Email fixo
        assunto = 'Registro de Não Conformidade'
        corpo = 'Corpo do email'
        enviar_email(destinatario, assunto, corpo, nome_arquivo_pdf)


# Função para página de indicadores de coletas
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

# Função para página de indicadores de não conformidades
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

# Função para página principal
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

# Executar a página principal
pagina_principal()
