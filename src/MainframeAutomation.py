import time
import win32com.client

class MainframeAutomation:
    def __init__(self, session_id='A'):
        self.session_id = session_id
        try:
            # Conexão com a API do MainframeM

            # self.win_metrics = win32com.client.Dispatch("MainframeM.autECLWinMetrics")
            # self.win_metrics.SetConnectionByName(self.session_id)
            # Conexão com a API do Mainframe
            self.conn_list = win32com.client.Dispatch("Mainframe.autECLConnList")
            self.session = win32com.client.Dispatch("Mainframe.autECLSession")

            self.win_metrics = win32com.client.Dispatch("Mainframe.autECLWinMetrics")
            self.win_metrics.SetConnectionByName(self.session_id)

            self.conn_list.Refresh()
            self.session.SetConnectionByName(self.session_id)

            if not self.session.Ready:
                raise Exception(f"A sessão {self.session_id} não está pronta.")

            print(f"Conectado à sessão {self.session_id}.")

        except Exception as e:
            print(f"Erro de conexão: {e}")
            raise

    def mover_cursor(self, linha, coluna):
        """Move o cursor para uma posição específica antes de digitar."""
        ps = self.session.autECLPS
        ps.SetCursorPos(linha, coluna)

    def ler_tela(self, linha, coluna, comprimento):
        ps = self.session.autECLPS
        texto = ps.GetText(linha, coluna, comprimento)
        return texto # Retorna RAW (com espaços) para manter alinhamento

    def enviar_tecla(self, tecla):
        """
        Envia teclas ou texto como se fosse digitação real.
        Ex: '1[enter]' ou apenas '[enter]'
        """
        ps = self.session.autECLPS
        ps.SendKeys(tecla)

    def digitar_texto_longo(self, texto, linha=None, coluna=None):
        """
        Usa SetText para preencher campos longos (ex: descrições).
        """
        ps = self.session.autECLPS

        if linha and coluna:
            # Tenta mover o cursor para garantir
            try:
                self.mover_cursor(linha, coluna)
            except:
                pass # Se falhar mover, tenta escrever mesmo assim

            ps.SetText(texto, linha, coluna)
        else:
            # Se não der linha, escreve onde o cursor estiver
            ps.SendKeys(texto)

    def aguardar_pronto(self, timeout_ms=10000):
        """Espera o reloginho (X-Clock) sumir."""
        oia = self.session.autECLOIA
        if not oia.WaitForInputReady(timeout_ms):
            raise TimeoutError("Mainframe demorou demais.")
        oia.WaitForAppAvailable(timeout_ms)

    def digitar_texto(self, texto, linha=None, coluna=None):
        """
        Move cursor → escreve → aguarda.
        Se linha/coluna não fornecidos, usa a posição atual do cursor.
        """
        ps = self.session.autECLPS

        # Se forneceu linha/coluna, move o cursor
        if linha is not None and coluna is not None:
             ps.SetCursorPos(linha, coluna)

        # Escreve o texto (instantâneo, como no VBA)
        ps.SetText(texto)

        # Garante sincronia sem gastar tempo com sleep fixo
        self.aguardar_pronto()



    def definir_visibilidade(self, visivel: bool):
        """
        Define se a janela do Mainframe deve ficar visível ou oculta.
        visivel=False aumenta MUITO a performance pois o Windows não precisa desenhar a tela.
        """
        self.win_metrics.Visible = visivel






def executar_fluxo_basico():
    robo = MainframeAutomation('A')
    LINHA_COMMAND = 23
    COLUNA_COMMAND = 14

    try:
        robo.aguardar_pronto()
        robo.enviar_tecla("[enter]")
        robo.aguardar_pronto()

        robo.digitar_texto("1", LINHA_COMMAND, COLUNA_COMMAND)
        robo.aguardar_pronto()

        robo.enviar_tecla("[enter]")
        robo.aguardar_pronto()

        robo.digitar_texto("texto", linha=1, coluna=1)
        robo.aguardar_pronto()

        robo.enviar_tecla("[enter]")
        robo.aguardar_pronto()

        robo.digitar_texto("21", linha=3, coluna=22)
        robo.aguardar_pronto()
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    executar_fluxo_basico()
