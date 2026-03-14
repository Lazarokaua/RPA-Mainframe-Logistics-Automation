from datetime import datetime
import MainframeAutomation as mainframe
from logar import logar

def tela_carga_entrega(matricula, senha, data_entrega, grupo_rota):
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
        robo.digitar_texto("texto", 21, 2)
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

if __name__ == "__main__":
    tela_carga_entrega()
