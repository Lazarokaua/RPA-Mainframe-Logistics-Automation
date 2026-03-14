import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src')))


from datetime import datetime
import MainframeAutomation as mainframe
from logar import logar
from utils_captura_tela import capturar_area as ca
import pandas as pd


# import streamlit as st


def coleta_infos_pesado(matricula, senha, data_entrega, grupo_rota, output_folder):

    hoje = datetime.now().strftime("%d.%m.%Y")

    robo = mainframe.MainframeAutomation("A")

    print("Logando no Mainframe")
    logar(robo)

    print("Tela Inicial")
    robo.digitar_texto("texto", 1, 1)
    robo.enviar_tecla("[enter]")
    robo.aguardar_pronto()

    print("Tela Deposito")
    operacao = robo.ler_tela(1, 2, 4)

    if operacao == "texto":
        print(operacao)
        robo.digitar_texto("texto", 3, 37)
        robo.digitar_texto("texto", 3, 52)
        robo.digitar_texto("texto", 3, 79)
        robo.digitar_texto("texto", 5, 5)
        robo.enviar_tecla("[enter]")
    else:
        print("Erro: Não logou no Mainframe")
        return
    robo.aguardar_pronto()

    print("Tela Roteirização")
    operacao_b = robo.ler_tela(1, 2, 4)
    if operacao_b == "texto":
        print(operacao_b)
        robo.digitar_texto("texto", 3, 51)
        robo.digitar_texto(matricula, 3, 54)
        robo.digitar_texto(senha, 3, 70)
        robo.digitar_texto("texto", 5, 4)
        robo.digitar_texto("1", 21, 2)
        robo.enviar_tecla("[enter]")
    else:
        print("Erro: Não logou no Mainframe")
        return

    robo.aguardar_pronto()

    erro_notas_denegadas = robo.ler_tela(10, 6, 80)
    print(f"DEBUG: Texto lido na linha 24: '{erro_notas_denegadas}'")

    if "Essas" in erro_notas_denegadas.strip():
        print("Erro: Notas Denegadas")
        robo.digitar_texto("n", 16, 46)
        robo.enviar_tecla("[enter]")
        robo.aguardar_pronto()

    robo.aguardar_pronto()

    print("Tela Carga Entrega")
    robo.digitar_texto(data_entrega, 4, 34)
    print("Data inserida", data_entrega)
    robo.digitar_texto(grupo_rota, 5, 34)
    print("Grupo inserido", grupo_rota)
    robo.enviar_tecla("[enter]")
    robo.aguardar_pronto()

    # Leitura de tela
    print("====================")
    campos = []

    while True:
        # 1. Lê os dados da página ATUAL (Começando da coluna 1 para garantir alinhamento perfeito)
        print("Lendo dados da página atual")
        ler_dados = ca(robo, 12, 1, 21, 80)
        campos.append(ler_dados)

        # 2. Verifica se apareceu a mensagem de fim
        msg_fim = ca(robo, 24, 2, 24, 15)[0].strip()

        if "FIM" in msg_fim:
            break

        # 3. Se não é fim, vira a página
        robo.enviar_tecla("[PF8]")
        robo.aguardar_pronto()

    robo.enviar_tecla("[PF9]")
    robo.aguardar_pronto()
    robo.digitar_texto("disc", 23, 14)
    robo.enviar_tecla("[enter]")
    robo.aguardar_pronto()

    # Flattening
    todas_linhas = [linha for pagina in campos for linha in pagina]

    print("Iniciando processamento de dados")
    dados_processados = []

    for linha in todas_linhas:

        dados_processados.append({
            'box': linha[0:6].strip(),
            'carga': linha[6:16].strip(),
            'status': linha[16:36].strip(),
            'rota': linha[53:62].strip(),
            'total_pedidos': linha[62:69].strip() if len(linha) > 62 else '',
            'cubagem': linha[69:].strip() if len(linha) > 69 else ''
        })

    print("Processamento de dados finalizado")
    # Filter out box '999' using list comprehension to avoid modification during iteration issues
    dados_processados = [item for item in dados_processados if item['box'] != "999" and item['status'] != "ENCERRADA"]


    df = pd.DataFrame(dados_processados)
    print("Iniciando salvamento de dados")
    output_file = os.path.join(output_folder, "output.xlsx")
    print("Salvando dados")
    df.to_excel(output_file, sheet_name=f"data_{hoje}", index=False)
    print(f"Coleta realizada com sucesso! salvo em: {output_file}")

    return output_file

if __name__ == "__main__":
    # Test path
    test_path = os.path.join(os.path.dirname(__file__), "..", "output")
    coleta_infos_pesado("minhaMatricula", "minhaSenha", "19012026", "minhaRota", test_path)
