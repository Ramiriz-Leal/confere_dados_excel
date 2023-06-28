import pandas as pd
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Configurações de conexão com o servidor de e-mail
smtp_host = 'smtp.dominio.com.br'
smtp_port = 'portasmtp'
smtp_username = 'email@email.com.br'
smtp_password = 'senha'

# Caminho do arquivo
caminho_arquivo = r"\\pasta\subpasta\arquivo.xlsx"

# Obter a data e hora atual
data_hora_atual = datetime.datetime.now()

# Converter para uma string formatada
data_hora_formatada = data_hora_atual.strftime('%d/%m/%Y %H:%M:%S')

# Carregar o arquivo Excel usando pandas
df = pd.read_excel(caminho_arquivo, sheet_name='Produtos')

# Filtrar as linhas com estoque abaixo de 4 para o item "Toner"
estoque_toner = df[(df['Itens p/ uso'] == 'Toner') & (df['Quantidade'] < 4)]

# Filtrar as linhas com estoque entre 0 e 2 para os demais itens
estoque_critico = df[(df['Itens p/ uso'] != 'Toner') & (df['Quantidade'] >= 0) & (df['Quantidade'] < 2)]

# Combina estoque de toner e estoque dos demais itens na mesma coluna do estoque
estoque_critico = pd.concat([estoque_toner, estoque_critico])

# Verificar se há produtos em estoque crítico
if not estoque_critico.empty:
    # Configurações de email
    remetente = 'email@email.com.br'
    destinatario = ['email@email.com.br','email@email.com.br']
    assunto = 'Estoque T.I'

    # Construir o corpo do email
    mensagem = MIMEMultipart()
    mensagem['From'] = remetente
    mensagem['To'] = ', '.join(destinatario)
    mensagem['Subject'] = assunto

    # Criar tabela HTML com os produtos em estoque crítico
    tabela_html = "<h2>Estoque de produtos T.I de acordo com a data: {data_hora_formatada}</h2>"
    tabela_html += estoque_critico[['Itens p/ uso', 'Quantidade']].to_html(index=False)

    # Adicionar a tabela HTML à mensagem do email
    mensagem.attach(MIMEText(tabela_html, 'html'))

    # Criar conexão com o servidor SMTP
    servidor_smtp = smtplib.SMTP_SSL(smtp_host, smtp_port)

    # Fazer login no servidor SMTP
    servidor_smtp.login(smtp_username, smtp_password)

    # Enviar o email
    servidor_smtp.sendmail(remetente, destinatario, mensagem.as_string())

    # Fechar a conexão com o servidor SMTP
    servidor_smtp.quit()

    print("Email enviado com sucesso!")
else:
    print("Nenhum produto em estoque crítico.")