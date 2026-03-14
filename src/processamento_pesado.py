import pandas as pd
import sys
import os
from datetime import datetime

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from utils.resource_handler import get_resource_path


file_analisar_default = get_resource_path(os.path.join("output", "output.xlsx"))
base_filiais_default = get_resource_path(os.path.join("output", "base_filiais.xlsx"))

base_filiais = base_filiais_default


def get_daily_filenames(folder):
    """Gera os nomes de arquivo baseados na data atual."""
    hoje = datetime.now().strftime("%d-%m-%Y")
    indicacao_box_file = os.path.join(folder, f"indicação_box_{hoje}.xlsx")
    cargas_mainframe_file = os.path.join(folder, f"cargas_mainframe_{hoje}.xlsx")
    return cargas_mainframe_file, indicacao_box_file, hoje

def salvar_cargas_mainframe(df_cargas, filepath):
    """Salva o histórico diário de cargas do Mainframe."""
    print(f"💾 ETAPA 2: Salvando/atualizando arquivo de cargas do Mainframe: {filepath}")
    df_cargas.to_excel(filepath, index=False)
    print("✅ Arquivo de cargas do Mainframe salvo/atualizado.")

def carregar_alocacoes_existentes(filepath):
    """Carrega alocações já realizadas no dia, se houver."""
    if os.path.exists(filepath):
        print(f"🔍 ETAPA 3: Verificando alocações existentes de hoje: {filepath}")
        try:
            df_existente = pd.read_excel(filepath, sheet_name="resultado")
            print(f"⚠️ 2ª+ Puxada: Encontradas {len(df_existente)} cargas já alocadas hoje.")
            return df_existente
        except Exception as e:
            print(f"⚠️ Erro ao ler arquivo existente (iniciando como 1ª puxada): {e}")
            return pd.DataFrame()
    else:
        print("🆕 Modo: 1ª Puxada (nenhum arquivo de indicações encontrado para hoje)")
        return pd.DataFrame()

def processamento_base(arquivo, base_filiais, pasta_user):
    """
    Função principal de processamento com suporte a múltiplas execuções diárias ("puxadas").
    Agora utiliza o arquivo 'indicação_box_{data}.xlsx' como fonte da verdade para capacidade restante.
    """
    print("================================================================================")
    print("PROCESSAMENTO DE ALOCAÇÃO DE CARGAS")
    print("================================================================================")

    # Nomes de arquivos diários
    cargas_mainframe_path, indicacao_box_path, hoje_str = get_daily_filenames(pasta_user)

    # 1. Leitura do arquivo Mainframe (output.xlsx)
    print("📦 ETAPA 1: Lendo cargas do Mainframe...")
    print(f"   Arquivo: {arquivo}")
    try:
        df_cargas = pd.read_excel(arquivo, engine='openpyxl')
        print(f"   Total de cargas lidas: {len(df_cargas)}")
    except FileNotFoundError:
        print(f"❌ Erro ao ler arquivo Mainframe: {arquivo}")
        return pd.DataFrame(), pd.DataFrame()
    except Exception as e:
        print(f"❌ Erro ao ler arquivo Mainframe: {e}")
        return pd.DataFrame(), pd.DataFrame()

    # 2. Salvar histórico do dia (sobrescreve para manter o status mais atual)
    salvar_cargas_mainframe(df_cargas, cargas_mainframe_path)
    print(f"✅ Arquivo de cargas salvo/atualizado: {cargas_mainframe_path}")
    
    # 3. Carregar Base de Filiais e Capacity
    print("📋 ETAPA 3: Carregando base de rotas e boxes...")
    try:
        df_rotas = pd.read_excel(base_filiais, sheet_name="db") # Explicitando sheet name 'db' pois é o padrão
        df_boxes_config = pd.read_excel(base_filiais, sheet_name="capacity_filiais")

        # GARANTIA DE TIPOS NA BASE
        df_rotas["filial"] = df_rotas["filial"].astype(str)
        df_boxes_config["filial"] = df_boxes_config["filial"].astype(str)
        df_boxes_config["box"] = df_boxes_config["box"].astype(str)

        # Limpar nans
        df_boxes_config = df_boxes_config.dropna(subset=["filial", "box", "capacity"])

        print(f"   Rotas carregadas: {len(df_rotas)}")
        print(f"   Configurações de boxes: {len(df_boxes_config)}")
    except Exception as e:
        print(f"❌ Erro ao ler base_filiais.xlsx: {e}")
        # Tenta ler sem sheet name se falhar, ou retorna erro
        return pd.DataFrame(), pd.DataFrame()

    # -------------------------------------------------------------------------
    # Recuperação de Estado
    # -------------------------------------------------------------------------
    print("📦 ETAPA 4: Recuperando estado atual dos boxes...")

    # Inicializa capacitade atual com o TOTAL configurado
    # Estrutura: { (filial, box): capacity_restante }
    estado_boxes = {}

    # Preenche com capacidade total inicial
    for _, row in df_boxes_config.iterrows():
        key = (str(row["filial"]), str(row["box"]))
        estado_boxes[key] = float(row["capacity"])

    df_historico = pd.DataFrame()
    cargas_ja_alocadas_ids = set()

    # Se existe arquivo do dia, vamos ler para atualizar o "restante"
    if os.path.exists(indicacao_box_path):
        print(f"   Arquivo de indicações encontrado: {indicacao_box_path}")
        try:
            df_historico = pd.read_excel(indicacao_box_path, sheet_name="resultado")

            # Garantir tipos no histórico
            if not df_historico.empty:
                df_historico["filial"] = df_historico["filial"].astype(str)
                df_historico["box"] = df_historico["box"].astype(str)
                df_historico["carga"] = df_historico["carga"].astype(str)

                # Lista de cargas já alocadas para não duplicar
                cargas_ja_alocadas_ids = set(df_historico["carga"].unique())
                print(f"   Carregadas {len(df_historico)} alocações anteriores.")

                # Iterar e atualizar o estado_boxes com o valor mais recente
                if "restante_capacity" in df_historico.columns:
                     for _, row in df_historico.iterrows():
                        key = (str(row["filial"]), str(row["box"]))
                        restante = row.get("restante_capacity")

                        if pd.notna(restante):
                            estado_boxes[key] = float(restante)
                else:
                    print("⚠️ Coluna 'restante_capacity' não encontrada no histórico. Recalculando do zero (fallback).")
                    # Fallback: se não tiver a coluna, subtraímos cubagem
                    for _, row in df_historico.iterrows():
                        key = (str(row["filial"]), str(row["box"]))
                        cubagem = row.get("cubagem", 0)
                        if key in estado_boxes:
                            estado_boxes[key] -= float(cubagem)

        except Exception as e:
            print(f"⚠️ Erro ao ler histórico diário: {e}. Iniciando como se fosse vazio.")
            df_historico = pd.DataFrame()

    else:
        print("   Nenhum histórico encontrado (1ª Puxada do dia).")

    # 🔧 ETAPA 5: Preparando dados Mainframe
    print("🔧 ETAPA 5: Processando cargas Mainframe...")

    # Tratamento básico Mainframe
    df_cargas["rota"] = df_cargas["rota"].astype(str).str.replace(r"\D", "", regex=True)
    df_cargas["cubagem"] = df_cargas["cubagem"].astype(str).str.replace(",", ".")
    df_cargas["cubagem"] = pd.to_numeric(df_cargas["cubagem"], errors="coerce")

    # Merge com Rotas para pegar Filial e Tipo
    df_rotas["rota"] = df_rotas["rota"].astype(str)
    # Remove dups df_rotas
    df_rotas = df_rotas.drop_duplicates(subset=['rota'])

    df_cargas_completo = df_cargas.merge(
        df_rotas[["rota", "filial", "tipo"]],
        on="rota",
        how="left"
    )

    # Filtrar apenas NOVAS Cargas FECHADAS
    # Status deve ser FECHADA
    # Carga ID não deve estar em cargas_ja_alocadas_ids

    novas_cargas_lista = []

    for _, row in df_cargas_completo.iterrows():
        status = str(row.get("status", "")).upper()
        carga_id = str(row.get("carga", ""))

        # Regra básica: Só queremos alocar o que está FECHADA
        if status != "FECHADA":
            continue

        # Regra de duplicidade: Já processei essa carga hoje?
        if carga_id in cargas_ja_alocadas_ids:
            continue

        # Validar dados mínimos
        filial = str(row.get("filial", ""))
        if filial.lower() == "nan" or not filial:
            continue

        novas_cargas_lista.append(row)

    df_novas = pd.DataFrame(novas_cargas_lista)
    print(f"   Cargas FECHADAS Novas a processar: {len(df_novas)}")

    # 🎲 ETAPA 6: Alocação
    print("🎲 ETAPA 6: Realizando alocação das novas cargas...")

    novas_alocacoes = []
    erros_alocacao = []

    if not df_novas.empty:
        # Agrupar por filial para processar
        cargas_por_filial = df_novas.groupby("filial")

        for filial, grupo in cargas_por_filial:
            # Ordenar cargas: Maior cubagem primeiro (Best Fit)
            cargas_da_filial = grupo.sort_values(by="cubagem", ascending=False).to_dict('records')

            # Identificar boxes dessa filial
            # Filtra df_boxes_config para pegar boxes dessa filial
            boxes_da_filial_config = df_boxes_config[df_boxes_config["filial"] == filial].to_dict('records')

            if not boxes_da_filial_config:
                # Erro: Filial sem boxes
                for c in cargas_da_filial:
                    erro = c.copy()
                    erro["motivo"] = "Filial sem boxes cadastrados na base"
                    erro["box"] = "INFORMAR_BOX_MANUALMENTE"
                    erros_alocacao.append(erro)
                continue

            # Priorização de Boxes (Regra 9999 mantida)
            # Precisamos criar uma lista de objetos box que contenham o estado atual
            lista_boxes_estado = []
            for b_conf in boxes_da_filial_config:
                box_id = str(b_conf["box"])
                key = (str(row["filial"]), str(row["box"]))
                cap_atual = estado_boxes.get(key, 0.0)

                # Se negativo, zera (segurança)
                if cap_atual < 0: cap_atual = 0

                lista_boxes_estado.append({
                    "box": box_id,
                    "filial": filial,
                    "restante": cap_atual,
                    "original_capacity": b_conf["capacity"] # Informativo
                })

            # Ordenação dos boxes (Prioridade)
            # REGRA 9999: Prioriza BOX_A e BOX_B -> depois BOX_C (ou regra puxada especial se único)
            modo_puxada_especial = (len(cargas_por_filial) == 1 and str(filial) == "9999")

            # A ordenação será dinâmica dentro do loop de cargas, pois depende do tipo da carga (Leves vs Pesado)
            # lista_boxes_estado.sort(key=sort_key_boxes)

            for carga in cargas_da_filial:
                alocado = False
                cubagem_carga = float(carga["cubagem"])
                tipo_carga = str(carga["tipo"]) # .strip() ?

                # DEFINIR ORDEM DE PRIORIDADE DOS BOXES PARA ESTA CARGA ESPECÍFICA
                def sort_key_current_load(b):
                    bx = str(b["box"])
                    prio = 10 # Default baixa prioridade
                    if str(filial) == "9999":
                        if modo_puxada_especial:
                            # REGRA PUXADA ESPECIAL (SÓ 9999 NO DIA)
                            if tipo_carga == "Leves":
                                # Prioridade: BOX_D -> BOX_C -> Resto
                                if bx == "BOX_D": prio = 0
                                elif bx == "BOX_C": prio = 1
                                elif bx in ["BOX_A", "BOX_B"]: prio = 2
                                else: prio = 3
                            else:
                                # Normal (Pesado): BOX_C -> BOX_A/BOX_B -> Resto
                                if bx == "BOX_C": prio = 0
                                elif bx in ["BOX_A", "BOX_B"]: prio = 1
                                else: prio = 2
                        else:
                            # MODO MISTO (Tem outras filiais)
                            if tipo_carga == "Leves":
                                # Prioridade: BOX_E -> Resto
                                if bx == "BOX_E": prio = 0
                                else: prio = 1
                            else:
                                # Pesado: BOX_A/BOX_B -> BOX_C -> Resto
                                if bx in ["BOX_A", "BOX_B"]: prio = 0
                                elif bx == "BOX_C": prio = 1
                                else: prio = 2
                    else:
                         prio = 0

                    return prio

                # Cria uma copia da lista ordenada para esta iteração
                boxes_para_tentar = sorted(lista_boxes_estado, key=sort_key_current_load)

                for box_obj in boxes_para_tentar:
                    box_id = str(box_obj["box"])

                    # Verifica Capacidade
                    if cubagem_carga <= box_obj["restante"]:
                        # ALOCA
                        box_obj["restante"] -= cubagem_carga

                        # Atualiza estado global (opcional, já que estamos usando lista local, mas bom manter sync)
                        estado_boxes[(str(filial), box_id)] = box_obj["restante"]

                        sucesso = carga.copy()
                        sucesso["box"] = box_id
                        sucesso["restante_capacity"] = box_obj["restante"] # AQUI O CAMPO PEDIDO

                        # Campos para output limpo
                        # Garantir que tenhamos apenas campos úteis
                        sucesso_clean = {
                            "filial": filial,
                            "box": box_id,
                            "carga": carga["carga"],
                            "pedidos": carga.get("total_pedidos"),
                            "cubagem": cubagem_carga,
                            "tipo": tipo_carga,
                            "rota": carga["rota"],
                            "status": carga["status"],
                            "restante_capacity": box_obj["restante"],
                            "timestamp_leitura": datetime.now().strftime("%d/%m/%Y %H:%M")
                        }

                        novas_alocacoes.append(sucesso_clean)
                        alocado = True
                        print(f"   ✓ Carga {carga['carga']} -> Box {box_id} | Resta: {box_obj['restante']:.2f}")
                        break

                if not alocado:
                    erro = carga.copy()
                    erro["motivo"] = "Sem capacidade disponível nos boxes"
                    erro["box"] = "INFORMAR_BOX_MANUALMENTE"
                    # GARANTIA: copiar o campo total_pedidos para pedidos, caso não exista
                    erro["pedidos"] = carga.get("total_pedidos")
                    erros_alocacao.append(erro)
                    print(f"   ❌ Carga {carga['carga']} (Filial {filial}) não alocada. {cubagem_carga}m³")

    # 🔗 ETAPA 7: Consolidando e Salvando
    print("💾 ETAPA 7: Salvando arquivo de indicações...")

    df_novos_resultados = pd.DataFrame(novas_alocacoes)
    df_erros_novos = pd.DataFrame(erros_alocacao)

    # Combinar histórico + novos
    if not df_historico.empty and not df_novos_resultados.empty:
        # Garantir colunas iguais
        # Pode haver divergência de colunas se o histórico antigo não tinha 'restante_capacity'
        df_final = pd.concat([df_historico, df_novos_resultados], ignore_index=True)
    elif not df_novos_resultados.empty:
        df_final = df_novos_resultados
    else:
        df_final = df_historico

    if not df_final.empty:
        colunas_ordenadas = ["filial", "box", "carga", "pedidos", "cubagem", "tipo", "rota", "status", "restante_capacity", "timestamp_leitura"]
        # Filtrar colunas existentes e adicionar que faltam
        cols_existentes = [c for c in colunas_ordenadas if c in df_final.columns]
        cols_extras = [c for c in df_final.columns if c not in cols_existentes]

        df_final = df_final[cols_existentes + cols_extras]

        with pd.ExcelWriter(indicacao_box_path) as writer:
            df_final.to_excel(writer, sheet_name="resultado", index=False)

            # Erros (sempre da rodada atual ou acumula? Geralmente erro é resolvido na hora, vamos salvar só o atual)
            if not df_erros_novos.empty:
                # Reordena colunas para ficar igual a aba resultado
                colunas_ordenadas_erros = ["filial", "box", "carga", "pedidos", "cubagem", "tipo", "rota", "status", "motivo", "timestamp_leitura"]
                cols_existentes_erros = [c for c in colunas_ordenadas_erros if c in df_erros_novos.columns]
                # Excluir 'total_pedidos' explicitamente se existir, pois já copiamos para 'pedidos'
                cols_extras_erros = [c for c in df_erros_novos.columns if c not in cols_existentes_erros and c != "total_pedidos"]
                df_erros_novos = df_erros_novos[cols_existentes_erros + cols_extras_erros]

                df_erros_novos.to_excel(writer, sheet_name="erros_rodada_atual", index=False)

        print("✅ Arquivo atualizado com sucesso!")
        print(f"   Total linhas histórico: {len(df_final)}")
        print(f"   Novas alocações: {len(df_novos_resultados)}")
    else:
        print("⚠️ Nenhum dado final para salvar.")

    return df_final, df_erros_novos


if __name__ == "__main__":
    # Define test output path when running directly
    output_test_path = get_resource_path("output")
    # Ensure output exists
    if not os.path.exists(output_test_path):
        os.makedirs(output_test_path)

    processamento_base(file_analisar_default, base_filiais_default, output_test_path)
