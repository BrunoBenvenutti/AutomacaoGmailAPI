import os
import base64
import time
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from pathlib import Path
from datetime import datetime, timedelta, timezone


# Esse programa é uma automação que monitora e baixa anexos de contas do Gmail. Ele suporta multi contas, ou seja,
# vc pode deixar baixando e monitorando várias contas gmail, porém somente uma conta por vez, para não exceder o limite
# de uso da API do Gmail e nem ir contra as diretrizes de uso definidas pela Google.
# O programa baixará somente os emails da INBOX, que é a caixa de entrada principal do Gmail. Os emails devem estar
# marcados como unread(Não lidos). Emails lidos seram ignorados.

# Mais funcionalidades seram implementadas ao longo do tempo.

# Esse programa baixa anexos de email utilizando a API do Google.
# É necessário que a api esteja ativada somente em uma conta, e as outras contas irão fazer login atráves dessa chave
# de api gerada na conta principal. Toda vez que uma conta fizer login, é gerado um token temporário que garante as
# permissões necessárias para que esse programa acesse as funcionalidades do Gmail da conta em questão.

# A documentação da API se encontra em: https://developers.google.com/gmail/api/guides?hl=pt-br

# Gostaria de contribuir e me incentivar a continuar disponibilizando automações de forma gratuita?
# Envie sua contribuição para meu GitHub Sponsors
# Para feedback ou contato para trabalhos e soluções mande-me um email para: brkas_dev@proton.me

# Configurações
user = Path.home()
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
# CREDENTIALS_FILE = 'credentials.json'
# QUERY = 'is:unread has:attachment'  # Filtro para buscar emails não lidos com anexos
SAVE_DIR = os.path.join('savedir.txt')

# A primeira vez que o programa for iniciado, será solicitado a escolha de um diretório
if not os.path.exists(SAVE_DIR):
    dir = str(input("Δ Essa é sua primeira vez executando o programa, digite caminho do diretório onde os anexos "
                    "seram salvos: "))
    if not os.path.exists(dir):
        try:
            os.makedirs(dir)
            with open(SAVE_DIR, 'w') as diretorio:
                diretorio.write(dir)
            print("Δ Diretório criado com sucesso.")

        except Exception as e:
            print(f"Δ Erro ao criar diretório: {e}")
    else:
        print("Δ Diretório já existe. Salvando em savedir.txt")
        with open(SAVE_DIR, 'w') as diretorio:
            diretorio.write(dir)

# 1) Autenticar fazendo login no google
def gerar_token():
    creds = None
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)

    email_associado = obter_email(creds)

    if email_associado:
        print(f'Δ O e-mail associado à conta é: {email_associado}')
    else:
        print('Δ Não foi possível obter o endereço de e-mail.')

    return creds

def obter_email(credentials):
    service = build('gmail', 'v1', credentials=credentials)
    info_conta = service.users().getProfile(userId='me').execute()
    email = info_conta.get('emailAddress')
    return email

def listar_contas():
    diretorio = os.getcwd()
    palavra_de_busca = 'token'
    arquivos = os.listdir(diretorio)

    # Filtra os arquivos que contêm a palavra desejada no nome
    arquivos_com_palavra = [arquivo for arquivo in arquivos if palavra_de_busca in arquivo]

    contas = arquivos_com_palavra
    return contas

def verificar_conta(arquivo_escolhido):
    creds = None

    # O arquivo token.json armazena as credenciais do usuário e é criado automaticamente após a primeira execução
    if os.path.exists(arquivo_escolhido):
        creds = Credentials.from_authorized_user_file(arquivo_escolhido)

    # Se não houver credenciais válidas disponíveis, solicita que o usuário faça login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            credentials = gerar_token()  # Pega o valor de creds e joga na variavel "credentials"
            email = obter_email(credentials)  # Constroi a API com build usando a variavel credentials para obter o email

            # Salve as credenciais para a próxima vez
            salvar_credencial_unica(credentials, email)  # Com o valor das duas variaveis acima, salva o token em JSON

    # Construa o serviço Gmail
    service = build('gmail', 'v1', credentials=creds)
    return service

def salvar_credencial_unica(credentials, email):

    email_json = str(email + '_token.json')

    with open(email_json, 'w') as token_file:
        token_file.write(credentials.to_json())

    print(f"Δ Token para {email} criado com sucesso")

    return email_json

def menu():

    try:
        gmails = listar_contas()

        print("Δ Escolha o número da conta na qual deseja iniciar o download dos anexos: ")

        for i, arquivo in enumerate(gmails, start=1):
            print(f"{i}. {arquivo}")

        numero_escolhido = int(
            input(" Outras opções:\n"
                  "( 0 ) Baixar anexos de uma única conta (Monitora uma unica conta) \n"
                  "(-1 ) Adicionar nova conta \n"
                  "(-2 ) Escolher data, hora e ordem em que baixar os anexos \n"
                  "( ? ) Digite qualquer letra para encerrar o programa \n"
                  "Digite aqui: "))

        time.sleep(5)

        if numero_escolhido >= 1:
            escolher_ordem = int(input(
                "Δ Qual a ordem em que você quer baixar os anexos? \n"
                "1) (antigo-->recente) Do mais antigo para o mais recente \n"
                "2) (recente-->antigo) Do mais recente para o mais antigo \n"
                "Digite: "))
        else:
            escolher_ordem = 0


        return numero_escolhido, gmails, escolher_ordem


    except ValueError:
        print("Δ Você digitou uma STRING. Encerrando programa")
        exit()
    except KeyboardInterrupt:
        print("Δ KeyboardInterrupt. Encerrando programa.")
        exit()

def marcar_como_lido(service, user_id, message_id):
    try:
        service.users().messages().modify(userId=user_id, id=message_id, body={'removeLabelIds': ['UNREAD']}).execute()
        print("Δ E-mail marcado como lido")
    except Exception as e:
        print(f"Δ Erro ao marcar email como lido: {e}")


def retornando_diretorio():
    with open(SAVE_DIR, 'r') as arquivo:
        diretorio = arquivo.read()
    return diretorio

def baixar_anexos(service, user_id, message_id, diretorio):
    print("Δ message id:", message_id)

    try:
        message = service.users().messages().get(userId=user_id, id=message_id).execute()

        print("Δ Detalhes do email: ", message)

        for part in message['payload']['parts']:
            if 'filename' in part and part['filename']:
                nome_anexo = part['filename']

                # Construir o novo nome do arquivo com data e hora
                timestamp = datetime.fromtimestamp(int(message['internalDate']) / 1000.0)
                timestamp_str = timestamp.strftime('%Y-%m-%d_%H-%M-%S')
                novo_nome_arquivo = f"{timestamp_str}_{nome_anexo}"

                # Obter os dados binários do anexo
                if 'body' in part and 'attachmentId' in part['body']:
                    attachment = service.users().messages().attachments().get(
                        userId=user_id,
                        messageId=message_id,
                        id=part['body']['attachmentId']
                    ).execute()

                    dados_base64 = attachment['data']
                    dados_bytes = base64.urlsafe_b64decode(dados_base64)

                    # Construir o caminho completo para salvar o anexo
                    full_path = os.path.join(diretorio, novo_nome_arquivo)

                    # Decodificar e salvar o conteúdo do anexo
                    with open(full_path, 'wb') as arquivo:
                        arquivo.write(dados_bytes)
                        print(f"Δ Anexo '{nome_anexo}' baixado com sucesso e salvo em: {full_path}")

        return True

    except Exception as e:
        print(f"Δ Erro ao baixar anexo: {e}")
        return False


def search_messages(service, query):

    result = service.users().messages().list(userId='me',q=query).execute()
    messages = [ ]
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId='me',q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
        result_size = result.get('resultSizeEstimate', [])
        print("Δ Resultados encontrados: ", result_size)
    print(messages)

    return messages

def date_to_seconds(date):
    dt = datetime.strptime(date, "%Y-%m-%dT%H:%M:%S.%fZ")
    return dt.timestamp()

def filtros(service, antigo):
    # Ajuste de acordo com o seu fuso horário
    data_hoje = datetime.now(timezone(timedelta(hours=-3))).replace(hour=0, minute=0, second=0, microsecond=0)

    data_formatada = data_hoje.strftime('%Y-%m-%dT%H:%M:%S.%f') + 'Z'

    seconds = date_to_seconds(data_formatada)


    if antigo == 1:

        query = f"is:unread after:{int(seconds)}"
        resultados = search_messages(service, query)

        return resultados


    elif antigo == 2:
        # Baixar os anexos na ordem normal, ou seja, do mais recente para o mais antigo
        pass

def main():
    while True:
        n = 0
        try:
            escolha, gmails, antigo = menu()
            diretorio = retornando_diretorio()

            if os.path.exists(diretorio):
                print("Δ Salvando anexos em: ", diretorio)
            else:
                print("Δ Erro no diretório. Necessário especificar o caminho novamente.")
                os.remove(SAVE_DIR)
                print("Δ Arquivo savedir.txt removido. Reinicie o programa para escolher um novo diretório.")
                time.sleep(5)
                exit()

            while True:
                print("Δ O download das proximas contas será feito automaticamente em todos emails que possuam um anexo.")
                if 1 <= escolha <= len(gmails):
                    arquivo = gmails[escolha - 1]
                    service = verificar_conta(arquivo)
                    remover_string = '_token.json'
                    nome_arquivo = arquivo.replace(remover_string, "")
                    print(f"Δ Baixando anexos de: {nome_arquivo}")

                    while True:
                        messages = filtros(service, antigo)

                        if not messages:
                            print("Δ Não foram encontrados emails não lidos.")
                            print(
                                f"Δ Aguardando 60 segundos para adquirir nova lista de emails na conta: {nome_arquivo}")
                            time.sleep(60)
                            messages = filtros(service, antigo)
                            n += 1
                            while n<=3:
                                print(f"Tentativas: {n} de 3")
                                break
                            if n > 2:
                                print(f"Não foram encontrados emails não lidos em: {nome_arquivo}")
                                time.sleep(3)
                                print("Analisando a próxima conta...")
                                time.sleep(5)
                                escolha += 1
                                n = 0
                                if escolha > len(gmails):
                                    escolha = 1
                                break
                        else:
                            print("Δ Foram encontrados emails não lidos.")
                            n=0
                            time.sleep(3)
                            break  # Sair do loop interno quando houver mensagens

                    for message in messages:
                        last_message = messages[-1]

                        success = baixar_anexos(service, 'me', last_message['id'], diretorio)

                        # Marcar e-mail como lido dependendo do sucesso do download
                        if success:
                            try:
                                marcar_como_lido(service, 'me', last_message['id'])
                            except Exception as e:
                                print(f"Δ Erro ao marcar email como lido: {e}, email id:{last_message['id']}")
                        messages.pop(-1)
                    time.sleep(4)

                elif escolha == -1:
                    creds = gerar_token()
                    email = obter_email(creds)
                    salvar_credencial_unica(creds, email)
                    break

                else:
                    print(
                        "Δ Número inválido. Por favor, escolha um número válido. Ou certifique-se de que o número escolhido é referente a uma conta."
                        "Caso não haja contas, adicione uma antes de prosseguir.")
                    time.sleep(10)
                    break

        except KeyboardInterrupt as k:
            print(f"Programa foi encerrado: {k}")




if __name__ == '__main__':
    main()







