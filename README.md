# Automação de lembretes pelo WhatsApp

Projeto em Python criado para automatizar lembretes de atendimentos a partir de uma planilha Excel. A aplicação lê os dados, valida telefone, data e horário, personaliza a mensagem e prepara o envio pelo WhatsApp Web.

## O que o projeto demonstra

- Automação de tarefas repetitivas com Python
- Leitura e validação de dados em Excel com OpenPyXL
- Geração de mensagens e links personalizados para o WhatsApp
- Registro estruturado de falhas em CSV
- Uso de modo de simulação para evitar envios acidentais
- Empacotamento possível como executável para Windows

## Tecnologias

- Python 3.11+
- OpenPyXL
- PyAutoGUI
- Excel
- WhatsApp Web

## Como executar

1. Crie um ambiente virtual e instale as dependências:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   ```

2. Execute primeiro no modo seguro de simulação:

   ```powershell
   python app.py
   ```

3. Revise os links exibidos. Para realizar envios pelo WhatsApp Web:

   ```powershell
   python app.py --enviar
   ```

O modo `--enviar` controla o navegador e confirma mensagens automaticamente. Mantenha o WhatsApp Web conectado e não use o computador durante o processo.

## Estrutura da planilha

A planilha deve possuir uma aba com estas cinco colunas, nesta ordem:

| paciente | responsavel | telefone | data | horario |
|---|---|---|---|---|
| Ana Souza | Carla Souza | 5534999990001 | 25/07/2026 | 09:00 |

Os dados incluídos no repositório são totalmente fictícios. Nunca publique nomes, telefones ou informações reais de pacientes.

## Privacidade e uso responsável

Este projeto é uma demonstração educacional. Para uso comercial ou em maior escala, prefira a API oficial do WhatsApp Business, obtenha consentimento dos destinatários e cumpra a LGPD e as regras da plataforma.

## Próximas melhorias

- Interface gráfica para selecionar a planilha
- Confirmação individual antes de cada envio
- Relatório final de enviados e não enviados
- Integração com a API oficial do WhatsApp Business
- Agendamento de lembretes

## Autor

Felipe Bastos — [LinkedIn](https://www.linkedin.com/in/felipe-b-077a97230/) · [GitHub](https://github.com/Felipe-Bastos97)
