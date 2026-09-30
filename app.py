import random
import re
import shutil
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import segno
from neonize.client import NewClient
from neonize.events import ConnectedEv


CAMINHO_PLANILHA = Path(r"C:\Users\deocl\Downloads\teste.xlsx")
CAMINHO_SESSAO = Path(__file__).with_name("whatsapp_session")
LINHA_CABECALHO = 8
INTERVALO_SEGUNDOS = 5

MENSAGENS_PREVENTIVO = [
    "Bom dia, tudo bem? Entro em contato para verificar sobre o pagamento de amanhã, está tudo certo?",
    "Olá, bom dia! Passando para confirmar se está tudo certo para o pagamento agendado para amanhã.",
    "Bom dia! Tudo bem com você? Gostaria apenas de confirmar nosso compromisso de pagamento amanhã.",
    "Olá! Bom dia. Entrando em contato rapidinho para checar se o pagamento de amanhã está confirmado.",
    "Bom dia, como vai? Passando apenas para lembrar do pagamento de amanhã e verificar se precisa de algo.",
    "Olá, tudo bem? Gostaria de saber se está tudo ok para a liquidação agendada para o dia de amanhã.",
    "Bom dia! Entro em contato para confirmar o pagamento previsto para amanhã. Tudo certo por aí?",
    "Olá, bom dia! Tudo bem? Passando para lembrar do vencimento de amanhã e verificar se está tudo certo.",
    "Bom dia! Checando apenas para garantir que está tudo ok referente ao pagamento agendado para amanhã.",
    "Olá, como vai? Entro em contato para confirmar a previsão do pagamento de amanhã. Conte comigo se precisar de algo!",
]

MENSAGENS_VENCIMENTO_HOJE = [
    "Bom dia, tudo bem? Está tudo certo para o pagamento hoje? Aguardo o comprovante.",
    "Olá, bom dia! Passando para confirmar se está tudo certo para o pagamento de hoje. Fico no aguardo do comprovante.",
    "Bom dia! Como vai? O pagamento previsto para hoje está confirmado? Aguardo o comprovante.",
    "Olá, tudo bem? Gostaria de confirmar se o pagamento de hoje será realizado. Pode me enviar o comprovante, por favor?",
    "Bom dia! Só confirmando se está tudo certo para o pagamento com vencimento hoje. Aguardo o comprovante.",
    "Olá! O pagamento de hoje está tudo certo? Quando puder, me encaminhe o comprovante, por favor.",
    "Bom dia, tudo bem? Estou passando para confirmar o pagamento de hoje. Fico aguardando o comprovante.",
    "Olá, bom dia! Consegue confirmar se o pagamento de hoje está programado? Aguardo o comprovante.",
    "Bom dia! Está tudo certo para realizar o pagamento hoje? Por favor, me envie o comprovante quando fizer.",
    "Olá, tudo bem? Como o vencimento é hoje, gostaria de confirmar o pagamento. Aguardo o comprovante.",
]


def normalizar_telefone(valor):
    digitos = re.sub(r"\D", "", str(valor).split(".")[0])
    if digitos.startswith("55") and len(digitos) in (12, 13):
        return digitos
    if len(digitos) in (10, 11):
        return f"55{digitos}"
    return None


def ler_envios():
    try:
        planilha = pd.read_excel(
            CAMINHO_PLANILHA,
            skiprows=LINHA_CABECALHO - 1,
            header=0,
        )
    except Exception as erro:
        raise RuntimeError(f"Não foi possível abrir a planilha: {erro}") from erro

    colunas = {
        str(coluna).strip().casefold(): coluna
        for coluna in planilha.columns
    }
    necessarias = ("vencimento", "tel", "status")
    if not all(coluna in colunas for coluna in necessarias):
        raise ValueError(
            "A planilha precisa conter as colunas 'vencimento', 'tel' e 'status'."
        )

    hoje = date.today()
    amanha = hoje + timedelta(days=1)
    envios = []

    for _, linha in planilha.iterrows():
        data_raw = linha[colunas["vencimento"]]
        telefone_raw = linha[colunas["tel"]]
        status_raw = linha[colunas["status"]]

        if pd.isna(data_raw) or pd.isna(telefone_raw) or pd.isna(status_raw):
            continue
        if str(status_raw).strip().casefold() != "no prazo":
            continue

        telefone = normalizar_telefone(telefone_raw)
        if telefone is None:
            print(f"[IGNORADO] Telefone inválido na planilha: {telefone_raw}")
            continue

        try:
            vencimento = pd.to_datetime(data_raw, dayfirst=True).date()
        except (TypeError, ValueError):
            continue

        if vencimento == hoje:
            mensagem = random.choice(MENSAGENS_VENCIMENTO_HOJE)
        elif vencimento == amanha:
            mensagem = random.choice(MENSAGENS_PREVENTIVO)
        else:
            continue

        envios.append(
            {
                "telefone": telefone,
                "vencimento": vencimento,
                "status": str(status_raw).strip(),
                "mensagem": mensagem,
            }
        )

    print("--- Iniciando Verificação ---")
    print(f"Buscando vencimentos para hoje: {hoje} e amanhã: {amanha}\n")
    return envios


def renderizar_qr_terminal(dados_qr):
    conteudo = dados_qr.decode("utf-8") if isinstance(dados_qr, bytes) else str(dados_qr)
    qr = segno.make_qr(conteudo)
    matriz = qr.matrix
    margem = 4
    largura_terminal = (len(matriz[0]) + margem * 2) * 2

    if shutil.get_terminal_size(fallback=(120, 24)).columns < largura_terminal:
        print(
            "Aumente a largura da janela do terminal para o QR não quebrar de linha."
        )

    print("Escaneie este QR em WhatsApp > Dispositivos conectados > Conectar dispositivo:\n")
    linha_vazia = " " * largura_terminal
    for _ in range(margem):
        print(linha_vazia)
    for linha in matriz:
        print(
            " " * (margem * 2)
            + "".join("##" if modulo else "  " for modulo in linha)
            + " " * (margem * 2)
        )
    for _ in range(margem):
        print(linha_vazia)
    print()


def main():
    envios = ler_envios()
    if not envios:
        print("Nenhum registro elegível para envio hoje ou amanhã.")
        return

    cliente = NewClient(str(CAMINHO_SESSAO))
    cliente.qr(lambda _cliente, dados_qr: renderizar_qr_terminal(dados_qr))

    @cliente.event(ConnectedEv)
    def ao_conectar(cliente_conectado, _evento):
        print("WhatsApp conectado.\n")
        for indice, envio in enumerate(envios):
            print(
                f"Encontrado: Telefone {envio['telefone']} | "
                f"Vencimento: {envio['vencimento']} | Status: {envio['status']}"
            )
            try:
                resultados = cliente_conectado.is_on_whatsapp(envio["telefone"])
                destinatario = next(
                    (
                        resultado
                        for resultado in resultados
                        if resultado.Query == envio["telefone"]
                    ),
                    None,
                )
                if destinatario is None or not destinatario.IsIn:
                    print(
                        f"[IGNORADO] {envio['telefone']} não foi encontrado no WhatsApp."
                    )
                else:
                    cliente_conectado.send_message(
                        destinatario.JID,
                        envio["mensagem"],
                    )
                    print(f"[SUCESSO] Mensagem enviada para {envio['telefone']}")
            except Exception as erro:
                print(f"[ERRO] Falha ao enviar para {envio['telefone']}: {erro}")

            if indice < len(envios) - 1:
                print(f"Aguardando {INTERVALO_SEGUNDOS} segundos para o próximo envio...\n")
                time.sleep(INTERVALO_SEGUNDOS)

        cliente_conectado.stop()

    try:
        cliente.connect()
    except KeyboardInterrupt:
        cliente.stop()
        print("\nExecução interrompida.")
    except Exception as erro:
        cliente.stop()
        print(f"[ERRO] Não foi possível conectar ao WhatsApp: {erro}")


if __name__ == "__main__":
    main()