"""Automação de lembretes de atendimento pelo WhatsApp Web.

O modo padrão é uma simulação segura: os links são gerados e exibidos, mas
nenhuma mensagem é enviada. Use ``--enviar`` somente após revisar a planilha.
"""

from __future__ import annotations

import argparse
import csv
import logging
import re
import time
import webbrowser
from dataclasses import dataclass
from datetime import date, datetime, time as horario
from pathlib import Path
from urllib.parse import quote

import openpyxl


PASTA_PROJETO = Path(__file__).resolve().parent
PLANILHA_PADRAO = PASTA_PROJETO / "outputs" / "exemplo" / "clientes_exemplo.xlsx"
ARQUIVO_ERROS = PASTA_PROJETO / "erros.csv"
COLUNAS_OBRIGATORIAS = ("paciente", "responsavel", "telefone", "data", "horario")


@dataclass(frozen=True)
class Atendimento:
    paciente: str
    responsavel: str
    telefone: str
    data: date
    horario: horario


def somente_digitos(valor: object) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def converter_data(valor: object) -> date:
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    return datetime.strptime(str(valor).strip(), "%d/%m/%Y").date()


def converter_horario(valor: object) -> horario:
    if isinstance(valor, datetime):
        return valor.time().replace(second=0, microsecond=0)
    if isinstance(valor, horario):
        return valor.replace(second=0, microsecond=0)
    return datetime.strptime(str(valor).strip(), "%H:%M").time()


def carregar_atendimentos(caminho: Path) -> list[Atendimento]:
    if not caminho.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {caminho}")

    workbook = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    planilha = workbook.active
    cabecalhos = [str(celula.value or "").strip().lower() for celula in planilha[1]]

    if tuple(cabecalhos[:5]) != COLUNAS_OBRIGATORIAS:
        raise ValueError(
            "Cabeçalhos inválidos. Use: " + ", ".join(COLUNAS_OBRIGATORIAS)
        )

    atendimentos: list[Atendimento] = []
    for numero_linha, valores in enumerate(planilha.iter_rows(min_row=2, values_only=True), 2):
        if not any(valores[:5]):
            continue
        try:
            telefone = somente_digitos(valores[2])
            if len(telefone) < 12:
                raise ValueError("telefone deve incluir código do país e DDD")
            atendimentos.append(
                Atendimento(
                    paciente=str(valores[0]).strip(),
                    responsavel=str(valores[1]).strip(),
                    telefone=telefone,
                    data=converter_data(valores[3]),
                    horario=converter_horario(valores[4]),
                )
            )
        except (TypeError, ValueError) as erro:
            logging.warning("Linha %s ignorada: %s", numero_linha, erro)

    workbook.close()
    return atendimentos


def criar_mensagem(atendimento: Atendimento) -> str:
    return (
        f"Olá {atendimento.responsavel}, tudo bem? Aqui é da Clínica Exemplo.\n"
        f"Gostaríamos de lembrar que o atendimento de {atendimento.paciente} "
        f"está agendado para {atendimento.data:%d/%m/%Y} às "
        f"{atendimento.horario:%H:%M}.\n"
        "Se precisar reagendar, estamos à disposição."
    )


def criar_link(atendimento: Atendimento) -> str:
    return (
        "https://web.whatsapp.com/send?"
        f"phone={atendimento.telefone}&text={quote(criar_mensagem(atendimento))}"
    )


def registrar_erro(atendimento: Atendimento, motivo: str) -> None:
    novo_arquivo = not ARQUIVO_ERROS.exists()
    with ARQUIVO_ERROS.open("a", newline="", encoding="utf-8") as arquivo:
        writer = csv.writer(arquivo)
        if novo_arquivo:
            writer.writerow(("paciente", "telefone", "motivo"))
        writer.writerow((atendimento.paciente, atendimento.telefone, motivo))


def processar(atendimentos: list[Atendimento], enviar: bool, intervalo: int) -> None:
    if enviar:
        try:
            import pyautogui
        except ImportError as erro:
            raise RuntimeError("Instale as dependências antes de usar --enviar") from erro

    for atendimento in atendimentos:
        link = criar_link(atendimento)
        if not enviar:
            print(f"[SIMULAÇÃO] {atendimento.paciente}: {link}")
            continue

        try:
            webbrowser.open(link)
            time.sleep(intervalo)
            pyautogui.press("enter")
            time.sleep(3)
            pyautogui.hotkey("ctrl", "w")
        except Exception as erro:  # falhas externas devem ser registradas por contato
            logging.exception("Falha no envio para %s", atendimento.paciente)
            registrar_erro(atendimento, str(erro))


def argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--planilha", type=Path, default=PLANILHA_PADRAO)
    parser.add_argument(
        "--enviar",
        action="store_true",
        help="abre o WhatsApp Web e confirma cada mensagem; sem esta opção, apenas simula",
    )
    parser.add_argument("--intervalo", type=int, default=12)
    return parser.parse_args()


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    opcoes = argumentos()
    atendimentos = carregar_atendimentos(opcoes.planilha)
    logging.info("%s atendimento(s) válido(s) carregado(s)", len(atendimentos))
    processar(atendimentos, opcoes.enviar, opcoes.intervalo)


if __name__ == "__main__":
    main()
