from gerar_resumo import read_base
import customtkinter as ctk
import threading

from datetime import datetime
import sys
import os
from tkinter import filedialog

from rota_pesado_coleta import coleta_infos_pesado
from processamento_pesado import processamento_base, base_filiais, get_daily_filenames
from tela_carga_entrega import tela_carga_entrega
from CTkMessagebox import CTkMessagebox


from rotinas.Rotina_A import executar_rotina_A
from rotinas.Rotina_B import executar_rotina_B

initial_date = datetime.now().strftime("%d%m%Y")

class App(ctk.CTk):
    def __init__(self, geometry = "600x650"):
        super().__init__()
        self.title("RPA - Cálculo de Cubagem")
        self.geometry(geometry)

        # Define absolute path for icon
        base_dir = os.path.dirname(__file__)
        icon_path = os.path.join(base_dir, "assets", "GRUPO.ico")

        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)



        # 1. Configurações Iniciais
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        self.geometry(geometry)

        # Label e Campo de Email
        self.label_matricula = ctk.CTkLabel(self, text="Matricula:", fg_color="transparent")

        self.label_erro = ctk.CTkLabel(self,  text="", text_color="red")

        self.matricula = ctk.CTkEntry(
            self,
            placeholder_text="ex: minhaMatricula",
            width=300,
            height=35,
            corner_radius=10
        )

        # Label e Campo de Senha
        self.label_senha = ctk.CTkLabel(self, text="Senha:", fg_color="transparent")

        self.senha = ctk.CTkEntry(
            self,
            placeholder_text="Digite sua senha",
            width=300,
            height=35,
            corner_radius=10,
            show="*"
        )

        self.grupo_label = ctk.CTkLabel(
            self,
            text="Grupo Rota: ",
            fg_color="transparent"
        )

        self.grupo_rota = ctk.CTkEntry(
            self,
            placeholder_text="ex: minhaRota",
            width=150,
            height=35,
            corner_radius=10
        )
        self.grupo_rota.insert(0, "minhaRota")

        self.data_inicial = ctk.CTkEntry(
            self,
            width=150,
            placeholder_text=f"{initial_date}"
        )


        self.data_inicial.insert(0, initial_date)

        # Frame para botões
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")

        # Botão de Execução
        self.btn_run = ctk.CTkButton(
            self.button_frame,
            text="Calcular Cubagem",
            width=150,
            height=40,
            corner_radius=10,
            command=self.iniciar_processo_calculo,
            state="disabled"
        )

        self.btn_tela_carga_entrega = ctk.CTkButton(
            self.button_frame,
            text="Tela de Carga/Entrega",
            width=150,
            height=40,
            corner_radius=10,
            command=self.iniciar_processo_tela_carga_entrega,
            state="disabled"
        )

        self.btn_limpar_terminal = ctk.CTkButton(
            self.button_frame,
            text="🧹",
            width=40,
            height=40,
            corner_radius=10,
            command=self.limpar_terminal,
            state="normal"
        )

        self.btn_rotina_A = ctk.CTkButton(
            self.button_frame,
            text="Rodar Rotina A",
            width=150,
            height=40,
            corner_radius=10,
            command=self.executar_rotina_A,
            state="disabled"
        )

        self.btn_rotina_B = ctk.CTkButton(
            self.button_frame,
            text="Rodar Rotina B",
            width=150,
            height=40,
            corner_radius=10,
            command=self.executar_rotina_B,
            state="disabled"
        )

        self.alert_area = ctk.CTkTextbox(
            self,
            width=380,
            height=200,
            corner_radius=10
        )


        self.matricula_bind = self.matricula.bind("<KeyRelease>", lambda e: self.validar_login())

        self.senha_bind = self.senha.bind("<KeyRelease>", lambda e: self.validar_login())

        self.data_inicial_bind = self.data_inicial.bind("<KeyRelease>", lambda e: self.validar_data())

        self.grupo_rota_bind = self.grupo_rota.bind("<KeyRelease>", lambda e: self.validar_login())

        sys.stdout = ConsoleRedirector(self.alert_area)
        sys.stderr = ConsoleRedirector(self.alert_area)

        self.grid_columnconfigure(0, weight=1)

        self.label_matricula.grid(row=0, column=0, pady=(5, 0), sticky="n")
        self.label_erro.grid(row=1, column=0, pady=(2, 0))
        self.matricula.grid(row=2, column=0, pady=2)

        self.label_senha.grid(row=3, column=0, pady=(5, 0))
        self.senha.grid(row=4, column=0, pady=2)

        self.grupo_label.grid(row=5, column=0, pady=(5, 0))
        self.grupo_rota.grid(row=6, column=0, pady=2)
        self.data_inicial.grid(row=7, column=0, pady=2)

        self.button_frame.grid(row=8, column=0, pady=5)
        self.btn_run.grid(row=0, column=0, padx=10)
        self.btn_tela_carga_entrega.grid(row=0, column=1, padx=10)
        self.btn_limpar_terminal.grid(row=0, column=2, padx=5)
        self.btn_rotina_A.grid(row=1, column=0, padx=10, pady=5)
        self.btn_rotina_B.grid(row=1, column=1, pady=5)

        self.alert_area.grid(row=10, column=0, pady=5)

        self.copy_right = ctk.CTkLabel(
            self,
            text=f"© {datetime.now().year} - Lázaro Kauã - RPA - Cálculo de Cubagem",
            fg_color="transparent"
        )

        self.copy_right.grid(row=11, column=0, pady=2)


    def validar_login(self, event=None):
        matricula = self.matricula.get()
        senha = self.senha.get()
        grupo_de_rota = self.grupo_rota.get()

        senha_valida = len(senha) >= 8
        matricula_valida = len(matricula) <= 8
        grupo_de_rota_valido = len(grupo_de_rota) <= 4 and grupo_de_rota != ""

        if matricula_valida and senha_valida and grupo_de_rota_valido:
            self.btn_run.configure(state="normal")
            self.btn_tela_carga_entrega.configure(state="normal")
            self.btn_rotina_A.configure(state="normal")
            self.btn_rotina_B.configure(state="normal")
        else:
            self.btn_run.configure(state="disabled")
            self.btn_tela_carga_entrega.configure(state="disabled")
            self.btn_rotina_A.configure(state="disabled")
            self.btn_rotina_B.configure(state="disabled")

        if matricula and not matricula_valida:
            self.label_erro.configure(text="Matricula deve ter no máximo 8 caracteres")
        elif senha and not senha_valida:
            self.label_erro.configure(text="Senha deve ter no mínimo 8 caracteres")
        else:
            self.label_erro.configure(text="")

    def validar_data(self, event=None):
        pass

    def iniciar_processo_calculo(self):
        matricula = self.matricula.get()
        senha = self.senha.get()
        data_inicial = self.data_inicial.get()
        grupo_de_rota = self.grupo_rota.get()

        pasta_user = self.selecionar_pasta()

        if not pasta_user:
            print("Nenhuma pasta selecionada. Operação cancelada.")
            return

        self.btn_run.configure(state="disabled")
        self.btn_tela_carga_entrega.configure(state="disabled")

        thread = threading.Thread(target=self.executar_calculo, args=(matricula, senha, data_inicial, grupo_de_rota, pasta_user))
        thread.start()

        self.btn_run.configure(state="normal")
        self.btn_tela_carga_entrega.configure(state="normal")
        self.btn_rotina_A.configure(state="normal")
        self.btn_rotina_B.configure(state="normal")

    def executar_calculo(self, matricula, senha, data_inicial, grupo_de_rota, pasta_user):
        try:
            print(f"Iniciando coleta. Pasta selecionada: {pasta_user}")

            arquivo_coletado = coleta_infos_pesado(matricula, senha, data_inicial, grupo_de_rota, pasta_user)
            print(f"Coleta finalizada. Arquivo gerado: {arquivo_coletado}")

            print("Iniciando processamento e alocação...")
            result_df, erro_df = processamento_base(arquivo_coletado, base_filiais, pasta_user)

            print("Adicionando aba de resumo na planilha de indicações...")
            _, indicacao_box_path, _ = get_daily_filenames(pasta_user)
            read_base(result_df, indicacao_box_path)
            print("✅ Aba de resumo adicionada com sucesso!")
            self.after(0, lambda: CTkMessagebox(title="Sucesso", message="Processamento e resumo concluídos com sucesso!", icon="check"))

        except Exception as e:
            print(f"ERRO CRÍTICO: {str(e)}")
            import traceback
            traceback.print_exc()

        finally:
            self.after(0, lambda: self.btn_run.configure(state="normal"))

    def selecionar_pasta(self):
        return filedialog.askdirectory(title="Selecione a pasta para Salvar os Arquivos", mustexist=True)

    def iniciar_processo_tela_carga_entrega(self):
        grupo_de_rota = self.grupo_rota.get()
        data_entrega = self.data_inicial.get()
        matricula = self.matricula.get()
        senha = self.senha.get()
        tela_carga_entrega(matricula, senha, data_entrega, grupo_de_rota)

    def limpar_terminal(self):
        self.alert_area.delete("1.0", "end")
    def executar_rotina_A(self):
        data_entrega = self.data_inicial.get()
        matricula = self.matricula.get()
        senha = self.senha.get()
        executar_rotina_A(matricula, senha, data_entrega)
    def executar_rotina_B(self):
        data_entrega = self.data_inicial.get()
        matricula = self.matricula.get()
        senha = self.senha.get()
        executar_rotina_B(matricula, senha, data_entrega)




class ConsoleRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        def _update():
            self.text_widget.insert("end", text)
            self.text_widget.see("end")

        self.text_widget.after(0, _update)

    def flush(self):
        pass

if __name__ == "__main__":

    app = App()
    app.mainloop()
