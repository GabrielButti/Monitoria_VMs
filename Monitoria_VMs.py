import os
import socket
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import win32com.client
from datetime import datetime

# Escopo para o Google Sheets
ESCOPOS = [
    'https://www.googleapis.com/auth/spreadsheets'
]

# Função para autenticação no Google
def autenticar_google():
    credenciais = None
    if os.path.exists("token.json"):
        credenciais = Credentials.from_authorized_user_file("token.json", ESCOPOS)
    if not credenciais or not credenciais.valid:
        if credenciais and credenciais.expired and credenciais.refresh_token:
            credenciais.refresh(Request())
        else:
            fluxo = InstalledAppFlow.from_client_secrets_file(
                "./credentials.json", ESCOPOS)
            credenciais = fluxo.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credenciais.to_json())
    return credenciais

# Função para obter o status da tarefa agendada
def obter_status_tarefa(nome_tarefa):
    try:
        agendador = win32com.client.Dispatch("Schedule.Service")
        agendador.Connect()
        pasta_raiz = agendador.GetFolder("\\Caminho Ficticio")  # Caminho fictício
        tarefa = pasta_raiz.GetTask(nome_tarefa)

        mapa_status = {
            0: "Inativo",
            1: "Tarefa desativada",
            2: "Tarefa na fila",
            3: "Tarefa pronta para execução",
            4: "Tarefa em execução",
            5: "Tarefa falhou"
        }
        estado_tarefa = tarefa.State
        status = mapa_status.get(estado_tarefa, "Estado desconhecido")

        ultima_execucao = (tarefa.LastRunTime.strftime("%d/%m/%Y %H:%M:%S") if tarefa.LastRunTime else "Nunca foi executada")
        proxima_execucao = (tarefa.NextRunTime.strftime("%d/%m/%Y %H:%M:%S") if tarefa.NextRunTime else "Sem próxima execução agendada")
        resultado_ultima_execucao = tarefa.LastTaskResult

        return {
            "nome_tarefa": nome_tarefa,
            "status": status,
            "ultima_execucao": ultima_execucao,
            "proxima_execucao": proxima_execucao,
            "resultado_ultima_execucao": resultado_ultima_execucao
        }

    except Exception as e:
        return {"erro": str(e)}

# Função para obter o IP da máquina
def obter_ip_maquina():
    try:
        nome_maquina = socket.gethostname()
        ip_maquina = socket.gethostbyname(nome_maquina)
        return ip_maquina
    except Exception as e:
        return f"Erro ao obter IP: {str(e)}"

# Função para mapear o resultado da última execução
def mapear_resultado_execucao(resultado_execucao):
    mapa_erros = {
        0: "Sucesso",
        1: "Erro genérico",
        0x41301: "Tarefa em execução",
        0x41302: "Tarefa pronta",
        0x41303: "Tarefa em pausa",
        0x41304: "Nenhuma instância disponível",
        0x41305: "Tarefa foi encerrada",
        0x80070002: "Arquivo ou diretório não encontrado",
        0x8007010B: "Nome do diretório é inválido",
        0xC000013A: "Interrupção pelo usuário",
        0x8004131F: "Configuração inválida",
        0x8007052E: "Falha de autenticação",
        0xC0000005: "Violação de acesso"
    }
    return mapa_erros.get(resultado_execucao, "Resultado desconhecido")

# Função principal
def atualizar_planilha_google(id_planilha, dados):
    credenciais = autenticar_google()
    servico = build('sheets', 'v4', credentials=credenciais)
    planilha = servico.spreadsheets()

    # Nome da aba e cabeçalho
    nome_aba = "Relatorio"
    cabecalho = [
        "IP da Máquina", "Nome da Automação", "Nome da Tarefa", "Status", "Resultado Última Execução", 
        "Primeiro Resultado do Dia", "Última Execução", "Próxima Execução", "Última Hora de Execução"
    ]

    # Verificar e criar a aba "Relatorio" se necessário
    metadados_planilha = planilha.get(spreadsheetId=id_planilha).execute()
    nomes_abas = [aba['properties']['title'] for aba in metadados_planilha['sheets']]

    if nome_aba not in nomes_abas:
        planilha.batchUpdate(
            spreadsheetId=id_planilha,
            body={
                "requests": [
                    {
                        "addSheet": {
                            "properties": {"title": nome_aba}
                        }
                    }
                ]
            }
        ).execute()
        # Adicionar cabeçalho à nova aba
        planilha.values().update(
            spreadsheetId=id_planilha,
            range=f"{nome_aba}!A1:I1",
            valueInputOption="RAW",
            body={"values": [cabecalho]}
        ).execute()

    # Obter os dados existentes na planilha
    try:
        resultado = planilha.values().get(spreadsheetId=id_planilha, range=f"{nome_aba}!A2:I").execute()
        valores_existentes = resultado.get('values', [])
    except Exception:
        valores_existentes = []

    mapa_existente = {(linha[1], linha[0]): idx + 2 for idx, linha in enumerate(valores_existentes)}  # (Nome Tarefa, IP Máquina) -> Linha

    hoje = datetime.now().strftime("%d/%m/%Y")
    atualizacoes = []
    novos_dados = []

    for linha in dados:
        linha[4] = mapear_resultado_execucao(linha[4])

        try:
            ultima_execucao_horario = datetime.strptime(linha[6], "%d/%m/%Y %H:%M:%S").time()
        except ValueError:
            ultima_execucao_horario = None

        chave = (linha[1], linha[0])  
        if chave in mapa_existente:
            numero_linha = mapa_existente[chave]
            linha_existente = valores_existentes[numero_linha - 2]

            if linha_existente[5].startswith(hoje):
                linha.insert(5, linha_existente[5])  
            else:
                linha.insert(5, f"{hoje} - {linha[4]}")
            atualizacoes.append((numero_linha, linha))
        else:
            linha.insert(5, f"{hoje} - {linha[4]}")
            novos_dados.append(linha)

    ultima_hora_execucao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    for numero_linha, linha in atualizacoes:
        linha.append(ultima_hora_execucao)
        intervalo = f"{nome_aba}!A{numero_linha}:I{numero_linha}"
        planilha.values().update(
            spreadsheetId=id_planilha,
            range=intervalo,
            valueInputOption="RAW",
            body={"values": [linha]}
        ).execute()

    if novos_dados:
        for linha in novos_dados:
            linha.append(ultima_hora_execucao)
        proxima_linha = len(valores_existentes) + 2
        planilha.values().update(
            spreadsheetId=id_planilha,
            range=f"{nome_aba}!A{proxima_linha}:I",
            valueInputOption="RAW",
            body={"values": novos_dados}
        ).execute()

    print(f"{len(atualizacoes)} registros atualizados.")
    print(f"{len(novos_dados)} novos registros adicionados.")

if __name__ == "__main__":
    nomes_tarefas = [
        "Tarefa A", "Tarefa B", "Tarefa C", "Tarefa D"
    ]

    nome_automacao = "Monitoramento Automatizado"
    ip_maquina = obter_ip_maquina()
    dados_para_enviar = []

    for nome_tarefa in nomes_tarefas:
        log = obter_status_tarefa(nome_tarefa)
        if "erro" in log:
            print(f"Erro ao verificar a tarefa {nome_tarefa}: {log['erro']}")
        else:
            dados_para_enviar.append([
                ip_maquina,
                log["nome_tarefa"],
                nome_automacao,                
                log["status"],
                log["resultado_ultima_execucao"],
                log["ultima_execucao"],
                log["proxima_execucao"]
            ])

    id_planilha = "Seu_ID"
    atualizar_planilha_google(id_planilha, dados_para_enviar)
