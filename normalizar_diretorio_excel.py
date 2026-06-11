import pandas as pd
import unicodedata
import re
import os
import logging
from datetime import datetime
from openpyxl.utils import get_column_letter

# =======================================================
# CONFIGURAÇÃO DOS DIRETÓRIOS — ALTERE AQUI
# =======================================================
PASTA_ORIGEM = r"C:\Users\guilherme.gfw\Documents\EFETIVOS"
PASTA_DESTINO = r"C:\Users\guilherme.gfw\Documents\EFETIVOS_PADRONIZADOS_R01\EFETIVO 2006"

# =======================================================
# LISTA DE COLUNAS ALVO (FILTRO ESTRITO)
# =======================================================
COLUNAS_ALVO = [
    # Matrículas
    'NR_MATRICULA_SIAPE', 'NR_MATRICULA', 'NR_MATRICULA_DPF',
    # Pessoa
    'NR_CPF', 'NO_SERVIDOR',
    # Lotação (pode ser NO_, SG_ ou CD_)
    'NO_LOTACAO', 'SG_LOTACAO', 'CD_LOTACAO',
    # Cargo (pode ser CD_, NO_ ou SG_)
    'CD_CARGO', 'NO_CARGO', 'SG_CARGO', 'SG_CARGOS', 
    'NO_CARGO_EFETIVO',
    # Classe e Padrão
    'CD_CLASSE', 'CD_PADRAO'
]

# =======================================================
# FUNÇÕES DE APOIO
# =======================================================
def classificar_por_conteudo(serie):
    """Analisa os 5 primeiros registros válidos para definir o prefixo correto."""
    amostra = serie.dropna().head(5)
    if amostra.empty: return "NO_"
    
    if pd.api.types.is_numeric_dtype(amostra):
        return "CD_"
        
    amostra_str = amostra.astype(str).str.strip()
    amostra_limpa = amostra_str.str.replace(r'\.0$', '', regex=True)
    
    if amostra_limpa.str.isnumeric().all():
        return "CD_"
        
    if amostra_str.str.len().max() <= 3: 
        return "SG_"
        
    return "NO_"

def normalizar_coluna(nome):
    """Canonicaliza o nome da coluna ANTES de buscar no dicionário."""
    if pd.isna(nome): return ""
    if not isinstance(nome, str): nome = str(nome)
    
    # 1. LIMPEZA PESADA
    nome_limpo = nome.strip().upper()
    nome_limpo = unicodedata.normalize('NFKD', nome_limpo).encode('ASCII', 'ignore').decode('utf-8')
    nome_limpo = re.sub(r'[ \.\-]+', '_', nome_limpo) 
    nome_limpo = ''.join(c for c in nome_limpo if c.isalnum() or c == '_')
    nome_limpo = re.sub(r'_+', '_', nome_limpo).strip('_')

    # 2. DICIONÁRIO ENXUTO E CANONICALIZADO
    excecoes = {
        # --- PADRÕES DE IDENTIFICAÇÃO ---
        "USUARIO": "LOGIN", "SISTEMA_ELETRONICO": "SEI",
        "CPF": "NR_CPF", "MATRICULA": "NR_MATRICULA",
        "MATRICULA_1": "NR_MATRICULA", "MAT": "NR_MATRICULA", "CD_MATRICULA": "NR_MATRICULA",
        "MATRICULA_SIAPE": "NR_MATRICULA_SIAPE", 
        "MATRICULA_DPF": "NR_MATRICULA_DPF",
        "MATR_DPF": "NR_MATRICULA_DPF", 
        "MATR_SIAPE": "NR_MATRICULA_SIAPE",
        "TITULO_ELEITORAL": "NR_TITULO_ELEITORAL", "NO": "NR_NUMERO_ORDEM",
        
        # --- PESSOAS ---
        "SERVIDOR": "NO_SERVIDOR", "NM_SERVIDOR": "NO_SERVIDOR", "NOME": "NO_SERVIDOR",
        "NO_PESSOA": "NO_SERVIDOR", "PESSOA": "NO_SERVIDOR",
        
        # --- CARGOS E ATIVIDADES ---
        "ATIVIDADE": "NO_ATIVIDADE", 
        "CARGO": "CARGO", 
        "GRUPO_CARGO_EMPREGO": "CARGO",
        "CARGO_DO_SERVIDOR": "CARGO", 
        "CARGO_EMPREGO": "CARGO", 
        "CO_CARGO_EMPREGO": "CARGO", 
        "CODIGO_DO_CARGO": "CARGO",
        "CO_GRUPO_CARGO": "CARGO", 
        "CO_GRUPO_CARGO_EMPREGO": "CARGO",
        "CARGOS": "SG_CARGOS", 
        "CARGO_EFETIVO": "NO_CARGO_EFETIVO", 
        "CARRREIRA": "NO_CARREIRA", "CARREIRA": "NO_CARREIRA",
        "FUNCAO": "NO_FUNCAO", "NOVA_FUNCAO": "NO_NOVA_FUNCAO", 
        "CARGO_SIGLA": "SG_CARGO", "CARGO_CLASSE": "SG_CARGOCLASSE",
        
        # --- LOTAÇÃO ---
        "LOT":"LOTACAO",
        "LOTACAO": "LOTACAO", "LOT_1": "LOTACAO",
        "LOTACAO_DO_SERVIDOR": "LOTACAO", 
        "CO_UORG_LOTACAO_SERVIDOR": "LOTACAO",
        "COD_LOTACAO_SERVIDOR": "LOTACAO", 
        "UORG_LOTACAO_SERVIDOR": "LOTACAO",
        "UORG_EXERCICIO": "LOTACAO", 
        "SG_UNIDADE_LOTACAO": "SG_LOTACAO", 
        "SG_UNIDADE_EXERCICIO": "SG_LOTACAO",
        "CD_UORG_EXERCICIO": "CD_LOTACAO", 
        "CD_UORG": "CD_LOTACAO",
        "UNIDADE_LOTACAO": "LOTACAO", "UNIDADE_EXERCICIO": "LOTACAO",
        
        # --- CLASSES E NÍVEIS ---
        "CLASSE": "CD_CLASSE", "CO_CLASSE": "CD_CLASSE", 
        "PADRAO": "CD_PADRAO", "CO_PADRAO": "CD_PADRAO"
    }

    if nome_limpo in excecoes: 
        return excecoes[nome_limpo]
    
    nome = re.sub(r'^(CO_|COD_|CODIGO_)', 'CD_', nome_limpo)
    nome = re.sub(r'^(DA_|DATA_|DTA_)', 'DT_', nome)
    nome = re.sub(r'^(NUM_|NUMERO_|NU_)', 'NR_', nome)
    nome = re.sub(r'^(VAL_|VALOR_|VL_)', 'VR_', nome)
    nome = re.sub(r'^(IND_|INDICADOR_|ID_)', 'IC_', nome)
    nome = re.sub(r'^(NOME_|NM_)', 'NO_', nome)
    
    return nome[:30]


# =======================================================
# LÓGICA DE PROCESSAMENTO DE ARQUIVO INDIVIDUAL
# =======================================================
def processar_arquivo(caminho_origem, pasta_destino):
    try:
        xls = pd.ExcelFile(caminho_origem)
        linha_cabecalho, melhor_match, aba_correta = 0, 0, None
        
        for aba in xls.sheet_names:
            df_topo = pd.read_excel(xls, sheet_name=aba, nrows=5, header=None)
            for idx, row in df_topo.iterrows():
                simuladas = [normalizar_coluna(str(val).rsplit('.', 1)[0] if '.' in str(val) and str(val).rsplit('.', 1)[-1].isdigit() else val) for val in row.values]
                match = [c for c in COLUNAS_ALVO if c in simuladas]
                if len(match) > melhor_match:
                    melhor_match = len(match)
                    linha_cabecalho = idx
                    aba_correta = aba

        if melhor_match == 0:
            logging.warning(f"[IGNORADO] Nenhuma coluna alvo em: {os.path.basename(caminho_origem)}")
            return False

        logging.info(f"[LENDO] {os.path.basename(caminho_origem)} | Aba: '{aba_correta}' | Linha: {linha_cabecalho}")
        df = pd.read_excel(xls, sheet_name=aba_correta, header=linha_cabecalho)
        
        # 1. Inteligência de Conteúdo por ÍNDICE (iloc)
        novos_nomes = []
        for idx, col in enumerate(df.columns):
            col_str = str(col)
            base_col = col_str.rsplit('.', 1)[0] if '.' in col_str and col_str.rsplit('.', 1)[-1].isdigit() else col_str
            nome_norm = normalizar_coluna(base_col)
            
            if nome_norm in ['NO_CARGO', 'CD_CARGO', 'CARGO']:
                prefixo = classificar_por_conteudo(df.iloc[:, idx])
                novos_nomes.append(prefixo + 'CARGO')
            elif nome_norm in ['LOTACAO', 'NO_LOTACAO', 'CD_LOTACAO']:
                prefixo = classificar_por_conteudo(df.iloc[:, idx])
                novos_nomes.append(prefixo + 'LOTACAO')
            else:
                novos_nomes.append(nome_norm)
                
        df.columns = novos_nomes

        # 2. Regra da barra (/)
        novos_nomes_barra = []
        for idx, col in enumerate(df.columns):
            if col.startswith('NO_') and ('CARGO' in col or 'LOTACAO' in col):
                if df.iloc[:, idx].dropna().astype(str).str.contains(r'/').any():
                    novos_nomes_barra.append(col.replace('NO_', 'SG_'))
                else:
                    novos_nomes_barra.append(col)
            else:
                novos_nomes_barra.append(col)
        df.columns = novos_nomes_barra

        # 3. Regras semânticas por tamanho
        REGRAS_SEMANTICAS = {'NO_CLASSE': ('CD_CLASSE', 4), 'NO_PADRAO': ('CD_PADRAO', 5)}
        novos_nomes_semantica = []
        for idx, col in enumerate(df.columns):
            if col in REGRAS_SEMANTICAS:
                col_novo, comp_max = REGRAS_SEMANTICAS[col]
                mediana = df.iloc[:, idx].dropna().astype(str).str.len().median()
                novos_nomes_semantica.append(col_novo if mediana <= comp_max else col)
            else:
                novos_nomes_semantica.append(col)
        df.columns = novos_nomes_semantica

        # 4. Remove duplicados residuais
        df = df.loc[:, ~df.columns.duplicated()].copy()

        # ======================================================================
        # SOLUÇÃO DEFINITIVA DA NOTAÇÃO CIENTÍFICA (CONVERSÃO PARA TEXTO PURO)
        # ======================================================================
        for col in df.columns:
            # Pega tudo que é identificador (Matrículas, CPFs, Códigos)
            if col.startswith('NR_') or col.startswith('CD_'):
                # 1. Força a transformação para texto
                df[col] = df[col].astype(str)
                # 2. Remove o ".0" do final caso o Pandas tenha lido como Float
                df[col] = df[col].str.replace(r'\.0$', '', regex=True)
                # 3. Limpa células vazias que viraram as palavras "nan" ou "<NA>"
                df[col] = df[col].replace(['nan', '<NA>', 'NaN', 'None'], '')
                # 4. Remove eventuais espaços nas bordas
                df[col] = df[col].str.strip()
        # ======================================================================

        # 5. Exportação com AutoFit
        cols_finais = [c for c in COLUNAS_ALVO if c in df.columns]
        df_final = df[cols_finais].copy()
        
        nome_saida = os.path.splitext(os.path.basename(caminho_origem))[0] + '_PADRONIZADO.xlsx'
        caminho_completo_saida = os.path.join(pasta_destino, nome_saida)

        with pd.ExcelWriter(caminho_completo_saida, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Dados')
            worksheet = writer.sheets['Dados']
            
            for idx, col in enumerate(df_final.columns):
                tamanho_maximo = len(str(col))
                if not df_final[col].dropna().empty:
                    maior_conteudo = df_final[col].dropna().astype(str).map(len).max()
                    tamanho_maximo = max(tamanho_maximo, maior_conteudo)
                
                largura_ajustada = tamanho_maximo + 2
                col_letra = get_column_letter(idx + 1)
                worksheet.column_dimensions[col_letra].width = largura_ajustada

        logging.info(f"[SALVO] Sucesso: {nome_saida} com {len(cols_finais)} colunas.")
        return True

    except Exception as e:
        logging.error(f"[ERRO] Falha ao processar {caminho_origem}: {e}", exc_info=True)
        return False


# =======================================================
# LÓGICA DE NAVEGAÇÃO EM LOTE E ESPELHAMENTO
# =======================================================
def processar_pastas(origem, destino_raiz):
    os.makedirs(destino_raiz, exist_ok=True)
    
    log_path = os.path.join(destino_raiz, f"log_processamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
        
    logging.basicConfig(level=logging.INFO, format='%(message)s', handlers=[
        logging.FileHandler(log_path, 'w', 'utf-8'),
        logging.StreamHandler()
    ])

    logging.info("==================================================")
    logging.info(f"INICIANDO PROCESSAMENTO EM LOTE")
    logging.info(f"Origem : {origem}")
    logging.info(f"Destino: {destino_raiz}")
    logging.info("==================================================\n")

    arquivos_processados = 0
    arquivos_ignorados = 0

    for pasta_atual, _, arquivos in os.walk(origem):
        for arquivo in arquivos:
            if arquivo.lower().endswith(('.xls', '.xlsx')) and not arquivo.startswith('~$'):
                caminho_origem_completo = os.path.join(pasta_atual, arquivo)
                
                caminho_relativo = os.path.relpath(pasta_atual, origem)
                pasta_destino_atual = destino_raiz if caminho_relativo == '.' else os.path.join(destino_raiz, caminho_relativo)
                os.makedirs(pasta_destino_atual, exist_ok=True)
                
                if processar_arquivo(caminho_origem_completo, pasta_destino_atual):
                    arquivos_processados += 1
                else:
                    arquivos_ignorados += 1

    logging.info("\n==================================================")
    logging.info(f"RESUMO DO PROCESSAMENTO")
    logging.info(f"Arquivos Padronizados com Sucesso: {arquivos_processados}")
    logging.info(f"Arquivos Ignorados/Com Falhas    : {arquivos_ignorados}")
    logging.info(f"Log detalhado salvo em           : {log_path}")
    logging.info("==================================================")


if __name__ == "__main__":
    processar_pastas(PASTA_ORIGEM, PASTA_DESTINO)