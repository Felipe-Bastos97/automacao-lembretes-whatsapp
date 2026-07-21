# Automação de mensagens pelo WhatsApp

Solução em Python criada para automatizar mensagens personalizadas a partir de uma planilha Excel. O projeto demonstra como uma mesma base pode atender diferentes processos de negócio com validação de dados e modo seguro de simulação.

## Módulos

### Lembretes de atendimentos

O arquivo `app.py` lê paciente, responsável, telefone, data e horário na aba `Atendimentos`, preparando lembretes personalizados.

```powershell
python app.py
```

### Avisos de vencimento

O arquivo `cobrancas.py` lê cliente, telefone, vencimento e link de pagamento na aba `Cobrancas`, preparando avisos personalizados.

```powershell
python cobrancas.py
```

## O que o projeto demonstra

- Automação de tarefas repetitivas com Python
- Leitura e validação de dados em Excel com OpenPyXL
- Geração de lembretes de atendimentos e avisos de vencimento
- Reutilização da mesma arquitetura em diferentes casos de uso
- Geração de links personalizados para o WhatsApp
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

2. Execute o módulo desejado primeiro no modo seguro de simulação:

   ```powershell
   python app.py
   python cobrancas.py
   ```

3. Revise os links exibidos. Para realizar envios pelo WhatsApp Web:

   ```powershell
   python app.py --enviar
   python cobrancas.py --enviar
   ```

O modo `--enviar` controla o navegador e confirma mensagens automaticamente. Mantenha o WhatsApp Web conectado e não use o computador durante o processo.

## Estrutura da planilha

A aba `Atendimentos` deve possuir estas cinco colunas:

| paciente | responsavel | telefone | data | horario |
|---|---|---|---|---|
| Ana Souza | Carla Souza | 5534999990001 | 25/07/2026 | 09:00 |

Os dados incluídos no repositório são totalmente fictícios. Nunca publique nomes, telefones ou informações reais de pacientes.

A aba `Cobrancas` deve possuir estas quatro colunas:

| cliente | telefone | vencimento | link_pagamento |
|---|---|---|---|
| Empresa Exemplo | 5534999990004 | 28/07/2026 | https://pagamentos.exemplo/fatura-001 |

Todos os nomes, telefones e links das duas abas são fictícios.

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
