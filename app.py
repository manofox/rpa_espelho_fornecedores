# ///////////////////////////////////////////////////////////////////////
<<<<<<< HEAD
# Projeto bot Espelho de notas (Envio por e-mail para os fornecedores)
# V: 1.1
# 
# Por: Renalt S(https://github.com/NaltAirCode/NaltAirCode)
# 
# Melhorias realizadas Por: Marcelo N (https://github.com/manofox)
# Data da Melhoria: 15/09/2023
# LISTA DE MELHORIAS
# 1 - Imports Ordenados: Imports em ordem alfabética para melhor legibilidade.
# 2 - Tratamento de Exceção Específico: Para evitar a captura de todas as exceções usando except OSError:. Em vez disso, use um tratamento de exceção mais específico, por exemplo, except FileExistsError: para capturar apenas exceções relacionadas ao diretório já existente.
# 3 - Comentários e Documentação: Adicionados comentários para explicar o propósito das seções do código e documentei as funções/métodos.
# 4 - Divisão em Funções: Para melhor organização e reutilização do código, é aconselhável dividir o código em funções.
# 5 - Tratamento de Erros Adequado: Lide com exceções de forma adequada, adicionando tratamentos de erro para lidar com cenários inesperados.
# 6 - Limpeza de Recursos: Certifique-se de liberar os recursos apropriados, como fechar a conexão SMTP, mesmo em caso de erro.
#
# CONSIDERE SEMPRE
# 1 - - Encapsulamento das Credenciais: Não é uma prática segura incluir senhas diretamente no código. Você pode usar variáveis de ambiente ou um arquivo de configuração externo para armazenar informações sensíveis.
#
=======
# Por: Renalt Silva (https://github.com/NaltAirCode/NaltAirCode)
# Melhoria realizada Por: Marcelo Nogueira (https://github.com/manofox)
# Data da Melhoria: 15/09/2023
# Projeto bot Espelho de notas (Envio por e-mail para os fornecedores)
# V: 1.1
>>>>>>> d532ec4530bcc222ea95e5424579554021c6f331
# # //////////////////////////////////////////////////////////////////////

import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import timedelta
import locale
import pandas as pd

def criar_diretorio(diretorio):
    try:
        os.mkdir(diretorio)
    except FileExistsError:
        print('A pasta já existe!')

def enviar_email(user, password, host, port, email, mensagem):
    try:
        server = smtplib.SMTP(host, port)
        server.ehlo()
        server.starttls()
        server.login(user, password)

        email_msg = MIMEMultipart()
        email_msg['From'] = user
        email_msg['To'] = email
        email_msg['Subject'] = 'Espelho de NF'
        email_msg.attach(MIMEText(mensagem, 'html'))

        server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
        print('Mensagem enviada!')
    except Exception as e:
        print('Erro ao enviar email:', str(e))
    finally:
        print('Desconectando o servidor')
        server.quit()

def main():
    locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
    locale.setlocale(locale.LC_MONETARY, 'pt_BR.UTF-8')

    data = datetime.date.today()
    dfdata = data.strftime('%Y-%m-%d')
    ano = data.year
    mes = data.month

    mes_ant = datetime.datetime.now() - timedelta(31)
    mes_ant_str = mes_ant.strftime('%B')
    mes_atual_str = data.strftime('%m')

     #use o método to_datetime para converter a string em data
    dfdata = pd.to_datetime(dfdata)
    #use o método strftime para formatar a data
    dfdata = dfdata.strftime('%d/%m/%Y')
    
    diretorio = "C:/python/espelhos/temp"
    criar_diretorio(diretorio)

    dados = pd.read_excel(os.path.join(diretorio, "matriz.xlsx"), sheet_name="Sheet1", engine='openpyxl')
    fornecedor = dados[['Referência', 'Fornecedor', 'Email', 'Empresa', 'Imposto', 'Valor']]
    dados_df = fornecedor.loc[dados['Referência'].dt.month == mes]

    dados_df.to_excel(os.path.join(diretorio, 'DadosdoMes.xlsx'), index=False)
    ndados = pd.read_excel(os.path.join(diretorio, 'DadosdoMes.xlsx'), sheet_name="Sheet1", engine='openpyxl')

    for contador in range(len(ndados)):
        referencia = ndados.loc[contador, 'Referência']
        consultor = ndados.loc[contador, 'Fornecedor']
        empresa = ndados.loc[contador, 'Empresa']
        email = ndados.loc[contador, 'Email']
        valor = ndados.loc[contador, 'Valor']
        imposto = ndados.loc[contador, 'Imposto']

        mensagem = f'<p>Olá, <b>{consultor}</b>,</p>'
        mensagem += f'<p>Segue o espelho para emissão de Nota Fiscal referente aos serviços prestados no mês de {mes_ant_str} de {ano}.</p>'
        mensagem += '<p>Lembramos que o pagamento somente será efetuado após apresentação da Nota Fiscal.</p>'
        mensagem += '<p>Por favor, encaminhe a cópia do documento em formato PDF para o e-mail: seuemail@gmail.com.</p>'
        mensagem += '<p>RAZÃO SOCIAL DA EMPRESA - CNPJ 99.999.999/0001-99</p>'
        mensagem += f'<p>Vencimento: 20/{mes_atual_str}/{ano}</p>'
        mensagem += f'''
            <table>
                <tr align="center">
                    <td width="225" nowrap style="width:169.0pt;border:solid black 1.0pt;background:#DADADA;padding:0cm 5.4pt 0cm 5.4pt;height:17.0pt">Mês de Referência</td>
                    <td width="405" nowrap style="width:304.0pt;border:solid black 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:17.0pt"><span style="font-size:10.0pt;font-family:"Segoe UI Symbol",sans-serif;color:black">{mes_ant_str} de {ano}</span></td>
                </tr>
                <tr align="center">
                    <td width="225" nowrap style="width:169.0pt;border:solid black 1.0pt;background:#DADADA;padding:0cm 5.4pt 0cm 5.4pt;height:17.0pt">Empresa</td>
                    <td width="405" nowrap style="width:304.0pt;border:solid black 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:17.0pt"><span style="font-size:10.0pt;font-family:"Segoe UI Symbol",sans-serif;color:black">{empresa}</span></td>
                </tr>
                <tr align="center">
                    <td width="225" nowrap style="width:169.0pt;border:solid black 1.0pt;background:#DADADA;padding:0cm 5.4pt 0cm 5.4pt;height:17.0pt">Impostos Retidos na Fonte</td>
                    <td width="405" nowrap style="width:304.0pt;border:solid black 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:17.0pt"><span style="font-size:10.0pt;font-family:"Segoe UI Symbol",sans-serif;color:black">{imposto}</span></td>
                </tr>
                <tr align="center">
                    <td width="225" nowrap style="width:169.0pt;border:solid black 1.0pt;background:#DADADA;padding:0cm 5.4pt 0cm 5.4pt;height:17.0pt">Valor da Nfe</td>
                    <td width="405" nowrap style="width:304.0pt;border:solid black 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:17.0pt"><span style="font-size:10.0pt;font-family:"Segoe UI Symbol",sans-serif;color:black">{valor}</span></td>
                </tr>
            </table>
        '''
        mensagem += '<p><span style="font-size:12.0pt;font-family:"Segoe UI",sans-serif;color:#303030"</span>Atenciosamente,</p>'
        mensagem += '<h4 style="font-family:"Segoe UI",sans-serif;color:#FF9801">OrangeFox - Gestão Administrativa</h4>'

        enviar_email('seue-mail@gmail.com', 'sua senha', 'smtp.gmail.com', 587, email, mensagem)

if __name__ == '__main__':
    print('+++++++++++++++++++++++++++ Espelho de Fornecedores +++++++++++++++++++++++++++++++')
    print('+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++')
    print('++++++++++++++++++++++++++ RESULTADO DA CONSULTA ++++++++++++++++++++++++++++++')
    main()
