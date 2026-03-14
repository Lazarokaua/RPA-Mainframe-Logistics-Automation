import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', 'src')))

import MainframeAutomation as mainframe
from logar import logar
import utils_captura_tela as ct

def executar_rotina_B(matricula, senha, data):
    print(f"Iniciando rotina com Matricula={matricula}, Data={data}")
    robo = mainframe.MainframeAutomation("A")

    logar(robo)
    print("Logado com sucesso")

    robo.digitar_texto("texto", 1, 1)
    print("texto preenchido com sucesso")

    robo.enviar_tecla("[enter]")
    robo.aguardar_pronto()
    print("Enter enviado")

    robo.digitar_texto("21", 4, 8)

    print("Empresa preenchida com sucesso")
    robo.digitar_texto("minhaFilial", 4, 16)

    print("Filial preenchida com sucesso")
    robo.digitar_texto("d", 4, 31)
    robo.digitar_texto("242", 4, 38)
    robo.digitar_texto("29", 4, 51)


    # Matricula
    robo.digitar_texto(matricula, 4, 54)


    # Senha
    robo.digitar_texto(senha, 4, 70)


    robo.digitar_texto("3", 7, 13)
    robo.digitar_texto("6", 21, 3)

    mensagem = robo.ler_tela(linha=24, coluna=1, comprimento=80)
    print(f"Mensagem da tela: {mensagem}")

    # Ou se quiser ler apenas a parte da mensagem de erro:
    mensagem_erro = robo.ler_tela(linha=24, coluna=1, comprimento=45)
    if "0262" in mensagem_erro or "ULTRAPASSOU" in mensagem_erro:
            print("ATENÇÃO: Período de trabalho ultrapassou jornada!")
            # Talvez lançar exceção ou retornar status?
            # raise Exception("Período de trabalho ultrapassou jornada!")

    robo.enviar_tecla("[enter]")
    robo.aguardar_pronto()


    robo.digitar_texto("x", 9, 7)
    robo.enviar_tecla("[enter]")
    robo.aguardar_pronto()

    robo.digitar_texto("rotina_b", 4, 11)
    robo.enviar_tecla("[enter]")
    robo.aguardar_pronto()

    robo.digitar_texto("x", 3, 49)
    robo.enviar_tecla("[enter]")
    robo.aguardar_pronto()
    print("Tela anterior processada. Iniciando preenchimento final...")

    robo.enviar_tecla("[enter]")
    robo.aguardar_pronto()

    # Data
    print(f"Digitando data: {data}...")
    robo.digitar_texto(data, 10, 13)

    # IMPORTANTE: Aguardar o sistema processar antes de ler mensagens
    robo.aguardar_pronto()

    # Agora sim, ler a mensagem da linha 24
    try:
        read = ct.salvar_tela_excel(robo, "teste.xlsx")
        print("Tela salva em excel com sucesso! ")

    except Exception as e:
        print(f"erro ao salvar a tela em excel {e}")


    robo.digitar_texto("        ", 10, 22)
    robo.enviar_tecla("[PF4]")
    robo.aguardar_pronto()
    robo.digitar_texto("S", 20, 49)
    robo.enviar_tecla("[enter]")
    print(f"Rotina rodada para a data {data}, com sucesso!")

if __name__ == "__main__":
    # Exemplo de uso para teste local
    executar_rotina_B("12345678", "senha_teste", "01012026")
