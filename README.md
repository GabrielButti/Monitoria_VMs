# 🛠️ Automação de Monitoramento e Integração com Google Sheets  

Com o avanço no desenvolvimento de automações em Python e o agendamento de tarefas em máquinas virtuais (VMs), o monitoramento das execuções pode se tornar desafiador, especialmente ao lidar com múltiplas tarefas distribuídas. Este projeto tem como objetivo centralizar e facilitar esse processo, registrando informações relevantes diretamente em uma planilha do Google Sheets.  

## 🚀 Funcionalidades  
- **Monitoramento de tarefas agendadas**  
  - Coleta de informações como status, última execução e próximo agendamento.  
- **Atualização em tempo quase real**  
  - O script é executado automaticamente a cada 15 minutos, garantindo que as informações estejam sempre atualizadas.  
- **Mapeamento e categorização de resultados**  
  - Tratamento de códigos de erro para identificação precisa de falhas.  
- **Integração com Google Sheets**  
  - Atualização dinâmica de dados, incluindo preservação de histórico e registros diários.  

## 📈 Benefícios  
- **Ideal para Dashboards**: A frequência de atualização torna essa solução perfeita para visualizações em tempo quase real, melhorando o monitoramento e a tomada de decisões.  
- **Centralização e Rastreabilidade**: Reduz intervenções manuais e consolida informações em um único lugar.  
- **Agilidade na Detecção de Problemas**: Fornece relatórios confiáveis para identificar falhas rapidamente.  

## 🛠️ Tecnologias Utilizadas  
- **Python**: Desenvolvimento do script principal.  
- **Google Sheets API**: Integração para registro e atualização de informações.  
- **Tarefas Agendadas**: Monitoramento automatizado via Windows Task Scheduler.  

## 📋 Pré-requisitos  
- Python 3.9+  
- Biblioteca `google-auth` e dependências relacionadas  
- Credenciais configuradas para a API do Google Sheets  

## ⚙️ Configuração  
1. Clone este repositório.  
2. Instale as dependências do projeto:  
   ```bash
   pip install -r requirements.txt
