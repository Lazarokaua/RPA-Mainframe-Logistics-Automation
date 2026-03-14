from MainframeAutomation import MainframeAutomation
from utils_captura_tela import capturar_area

def logar(robo: MainframeAutomation):

        try:
            robo.aguardar_pronto()
            robo.enviar_tecla("[PF9]")
            erro_disc = robo.ler_tela(24, 9, 6)
            print("Erro Disc", erro_disc)
            if erro_disc != "(DISC)":
                robo.enviar_tecla("[PF9]")
                robo.aguardar_pronto()
                robo.digitar_texto("disc", 23, 14)
                robo.enviar_tecla("[enter]")
                robo.aguardar_pronto()
        except Exception as e:
            print("Erro ao tentar deslogar", e)


        print("Enviando primeiro Enter...")
        robo.enviar_tecla("[enter]")
        robo.aguardar_pronto()

        try:
            print("Tentando selecionar opção 1...")
            tentativas = ["1", "2", "3"]
            for tentativa in tentativas:
                robo.digitar_texto(tentativa, 23, 14)
                robo.aguardar_pronto()


                print(f"Opção {tentativa} enviada.")
                robo.enviar_tecla("[enter]")
                robo.aguardar_pronto()

                erro = robo.ler_tela(24, 9, 6)
                print("Erro nas opções: ", erro)

                if erro.strip() != '(DISC)':
                    print(f"Sucesso detectado ao usar a opção {tentativa}.")
                    return


                print(f"Ainda não foi logado com a opção {tentativa}!")



        except Exception as e:
            print(f"Erro durante a tentativa {tentativa}: {e}")



if __name__ == "__main__":
    logar()
