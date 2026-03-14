"""
Utilitários para Captura de Tela do MainframeM

Este módulo contém funções para ler e salvar diferentes partes da tela do mainframe.

Autor: Equipe de Automação
Data: 27/11/2025
"""

from datetime import datetime
from pathlib import Path
import json


# ============================================================================
# CAPTURA DE TELA COMPLETA
# ============================================================================

def capturar_tela_completa(robo):
    """
    Captura toda a tela (24 linhas x 80 colunas).

    Args:
        robo: Instância do MainframeAutomation

    Returns:
        list: Lista com 24 strings (uma por linha)

    Example:
        >>> linhas = capturar_tela_completa(robo)
        >>> print(linhas[0])  # Primeira linha
        >>> print(linhas[23])  # Última linha (mensagens)
    """
    tela = []

    for linha in range(1, 25):  # Linhas 1 a 24
        texto = robo.ler_tela(linha, 1, 80)
        tela.append(texto)

    return tela


def salvar_tela_completa(robo, nome_arquivo=None, diretorio="capturas"):
    """
    Salva toda a tela em um arquivo de texto.

    Args:
        robo: Instância do MainframeAutomation
        nome_arquivo (str, optional): Nome do arquivo. Se None, usa timestamp.
        diretorio (str): Diretório onde salvar

    Returns:
        Path: Caminho do arquivo salvo

    Example:
        >>> arquivo = salvar_tela_completa(robo)
        >>> print(f"Tela salva em: {arquivo}")
    """
    # Criar diretório se não existir
    dir_path = Path(diretorio)
    dir_path.mkdir(exist_ok=True)

    # Gerar nome do arquivo se não fornecido
    if nome_arquivo is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nome_arquivo = f"tela_completa_{timestamp}.txt"

    arquivo_path = dir_path / nome_arquivo

    # Capturar tela
    tela = capturar_tela_completa(robo)

    # Salvar em arquivo
    with open(arquivo_path, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write(f"CAPTURA DE TELA - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write("=" * 80 + "\n\n")

        for i, linha in enumerate(tela, 1):
            f.write(f"{i:02d}: {linha}\n")

        f.write("\n" + "=" * 80 + "\n")

    print(f"✓ Tela completa salva em: {arquivo_path}")
    return arquivo_path



def capturar_area(robo, linha_inicial, coluna_inicial, linha_final, coluna_final):
    """
    Captura uma área retangular da tela.

    Args:
        robo: Instância do MainframeAutomation
        linha_inicial (int): Linha inicial (1-24)
        coluna_inicial (int): Coluna inicial (1-80)
        linha_final (int): Linha final (1-24)
        coluna_final (int): Coluna final (1-80)

    Returns:
        list: Lista de strings, uma por linha da área

    Example:
        >>> # Capturar área de dados (linhas 10-20, colunas 1-60)
        >>> area = capturar_area(robo, 10, 1, 20, 60)
        >>> for linha in area:
        ...     print(linha)
    """
    area = []
    largura = coluna_final - coluna_inicial + 1

    for linha in range(linha_inicial, linha_final + 1):
        texto = robo.ler_tela(linha, coluna_inicial, largura)
        area.append(texto)

    return area


def salvar_area(robo, linha_inicial, coluna_inicial, linha_final, coluna_final,
                nome_arquivo=None, diretorio="capturas"):
    """
    Salva uma área específica da tela em arquivo.

    Args:
        robo: Instância do MainframeAutomation
        linha_inicial (int): Linha inicial
        coluna_inicial (int): Coluna inicial
        linha_final (int): Linha final
        coluna_final (int): Coluna final
        nome_arquivo (str, optional): Nome do arquivo
        diretorio (str): Diretório onde salvar

    Returns:
        Path: Caminho do arquivo salvo

    Example:
        >>> # Salvar área de dados
        >>> arquivo = salvar_area(robo, 10, 1, 20, 60, "dados_pedido.txt")
    """
    # Criar diretório
    dir_path = Path(diretorio)
    dir_path.mkdir(exist_ok=True)

    # Gerar nome do arquivo
    if nome_arquivo is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nome_arquivo = f"area_L{linha_inicial}-{linha_final}_C{coluna_inicial}-{coluna_final}_{timestamp}.txt"

    arquivo_path = dir_path / nome_arquivo

    # Capturar área
    area = capturar_area(robo, linha_inicial, coluna_inicial, linha_final, coluna_final)

    # Salvar
    with open(arquivo_path, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write(f"CAPTURA DE ÁREA - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write(f"Linhas: {linha_inicial} a {linha_final}\n")
        f.write(f"Colunas: {coluna_inicial} a {coluna_final}\n")
        f.write("=" * 80 + "\n\n")

        for i, linha in enumerate(area, linha_inicial):
            f.write(f"{i:02d}: {linha}\n")

        f.write("\n" + "=" * 80 + "\n")

    print(f"✓ Área salva em: {arquivo_path}")
    return arquivo_path


def capturar_tela_filtrada(robo, ignorar_vazias=True):
    """
    Captura tela ignorando linhas vazias.

    Args:
        robo: Instância do MainframeAutomation
        ignorar_vazias (bool): Se True, ignora linhas vazias

    Returns:
        dict: Dicionário {numero_linha: texto}

    Example:
        >>> tela = capturar_tela_filtrada(robo)
        >>> for num_linha, texto in tela.items():
        ...     print(f"Linha {num_linha}: {texto}")
    """
    tela = {}

    for linha in range(1, 25):
        texto = robo.ler_tela(linha, 1, 80)

        if ignorar_vazias and not texto.strip():
            continue

        tela[linha] = texto

    return tela


def salvar_tela_json(robo, nome_arquivo=None, diretorio="capturas"):
    """
    Salva tela em formato JSON para fácil processamento.

    Args:
        robo: Instância do MainframeAutomation
        nome_arquivo (str, optional): Nome do arquivo
        diretorio (str): Diretório onde salvar

    Returns:
        Path: Caminho do arquivo salvo

    Example:
        >>> arquivo = salvar_tela_json(robo)
        >>> # Depois você pode ler:
        >>> import json
        >>> with open(arquivo) as f:
        ...     dados = json.load(f)
    """
    # Criar diretório
    dir_path = Path(diretorio)
    dir_path.mkdir(exist_ok=True)

    # Gerar nome do arquivo
    if nome_arquivo is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nome_arquivo = f"tela_{timestamp}.json"

    arquivo_path = dir_path / nome_arquivo

    # Capturar tela
    tela_completa = capturar_tela_completa(robo)

    # Criar estrutura JSON
    dados = {
        'timestamp': datetime.now().isoformat(),
        'linhas': {}
    }

    for i, linha in enumerate(tela_completa, 1):
        if linha.strip():  # Só adiciona linhas com conteúdo
            dados['linhas'][str(i)] = linha

    # Salvar JSON
    with open(arquivo_path, 'w', encoding='utf-8') as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    print(f"✓ Tela salva em JSON: {arquivo_path}")
    return arquivo_path


def salvar_tela_excel(robo, nome_arquivo=None, diretorio="capturas"):
    """
    Salva tela em formato Excel (requer pandas).

    Args:
        robo: Instância do MainframeAutomation
        nome_arquivo (str, optional): Nome do arquivo
        diretorio (str): Diretório onde salvar

    Returns:
        Path: Caminho do arquivo salvo

    Example:
        >>> arquivo = salvar_tela_excel(robo)
    """
    try:
        import pandas as pd
    except ImportError:
        print("❌ Erro: pandas não instalado. Execute: pip install pandas openpyxl")
        return None

    # Criar diretório
    dir_path = Path(diretorio)
    dir_path.mkdir(exist_ok=True)

    # Gerar nome do arquivo
    if nome_arquivo is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nome_arquivo = f"tela_{timestamp}.xlsx"

    arquivo_path = dir_path / nome_arquivo

    # Capturar tela
    tela = capturar_tela_completa(robo)

    # Criar DataFrame
    df = pd.DataFrame({
        'Linha': range(1, 25),
        'Conteúdo': tela
    })

    # Salvar Excel
    df.to_excel(arquivo_path, index=False)

    print(f"✓ Tela salva em Excel: {arquivo_path}")
    return arquivo_path


def comparar_telas(tela_antes, tela_depois):
    """
    Compara duas capturas de tela e mostra diferenças.

    Args:
        tela_antes (list): Primeira captura
        tela_depois (list): Segunda captura

    Returns:
        dict: Dicionário com linhas que mudaram

    Example:
        >>> tela1 = capturar_tela_completa(robo)
        >>> # ... fazer algo ...
        >>> tela2 = capturar_tela_completa(robo)
        >>> diferencas = comparar_telas(tela1, tela2)
    """
    diferencas = {}

    for i, (antes, depois) in enumerate(zip(tela_antes, tela_depois), 1):
        if antes != depois:
            diferencas[i] = {
                'antes': antes,
                'depois': depois
            }

    return diferencas


def imprimir_tela_completa(robo, mostrar_vazias=False):
    """
    Imprime tela completa no console de forma formatada.

    Args:
        robo: Instância do MainframeAutomation
        mostrar_vazias (bool): Se True, mostra linhas vazias

    Example:
        >>> imprimir_tela_completa(robo)
    """
    print("\n" + "=" * 80)
    print(f"TELA COMPLETA - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 80 + "\n")

    tela = capturar_tela_completa(robo)

    for i, linha in enumerate(tela, 1):
        if not mostrar_vazias and not linha.strip():
            continue
        print(f"{i:02d}: {linha}")

    print("\n" + "=" * 80 + "\n")


def imprimir_area(robo, linha_inicial, coluna_inicial, linha_final, coluna_final):
    """
    Imprime área específica no console.

    Args:
        robo: Instância do MainframeAutomation
        linha_inicial (int): Linha inicial
        coluna_inicial (int): Coluna inicial
        linha_final (int): Linha final
        coluna_final (int): Coluna final

    Example:
        >>> imprimir_area(robo, 10, 1, 20, 60)
    """
    print("\n" + "=" * 80)
    print(f"ÁREA: Linhas {linha_inicial}-{linha_final}, Colunas {coluna_inicial}-{coluna_final}")
    print("=" * 80 + "\n")

    area = capturar_area(robo, linha_inicial, coluna_inicial, linha_final, coluna_final)

    for i, linha in enumerate(area, linha_inicial):
        print(f"{i:02d}: {linha}")

    print("\n" + "=" * 80 + "\n")


if __name__ == "__main__":
    import MainframeAutomation as mainframe
    from logar import logar

    print("=" * 80)
    print("EXEMPLOS DE CAPTURA DE TELA")
    print("=" * 80)

    # Conectar
    robo = mainframe.MainframeAutomation("A")
    logar(robo)

    # Aguardar tela carregar
    robo.aguardar_pronto()

    # ========================================================================
    # EXEMPLO 1: Salvar tela completa
    # ========================================================================
    print("\n1. Salvando tela completa...")
    arquivo1 = salvar_tela_completa(robo, "minha_tela.txt")

    # ========================================================================
    # EXEMPLO 2: Salvar área específica
    # ========================================================================
    print("\n2. Salvando área específica (linhas 10-20, colunas 1-60)...")
    arquivo2 = salvar_area(robo, 10, 1, 20, 60, "area_dados.txt")

    # ========================================================================
    # EXEMPLO 3: Imprimir tela no console
    # ========================================================================
    print("\n3. Imprimindo tela no console...")
    imprimir_tela_completa(robo)

    # ========================================================================
    # EXEMPLO 4: Salvar em JSON
    # ========================================================================
    print("\n4. Salvando em JSON...")
    arquivo3 = salvar_tela_json(robo)

    # ========================================================================
    # EXEMPLO 5: Capturar apenas linhas com conteúdo
    # ========================================================================
    print("\n5. Capturando apenas linhas com conteúdo...")
    tela_filtrada = capturar_tela_filtrada(robo)
    print(f"Encontradas {len(tela_filtrada)} linhas com conteúdo")

    print("\n" + "=" * 80)
    print("EXEMPLOS CONCLUÍDOS")
    print("=" * 80)
