import pandas as pd
import unicodedata
import re
import os
import logging
from openpyxl.utils import get_column_letter

# =======================================================
# CONFIGURAÇÃO DOS DIRETÓRIOS — ALTERE AQUI
# =======================================================
ARQUIVO_ORIGEM = r"C:\Users\guilherme.gfw\Documents\EFETIVOS_PADRONIZADOS_R01\EFETIVO 2006\LIS - 12_DEZEMBRO.xls"

# O arquivo de saída será gerado na mesma pasta do arquivo de origem,
# com o sufixo _PADRONIZADO.xlsx. Altere ARQUIVO_DESTINO se preferir outro local.
ARQUIVO_DESTINO = None  # Exemplo: r"C:\Outra_Pasta\Resultado.xlsx"

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
# FUNÇÕES DE APOIO DA INTELIGÊNCIA
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
    
    # 1. LIMPEZA PESADA (Canonicalização)
    nome_limpo = nome.strip().upper()
    nome_limpo = unicodedata.normalize('NFKD', nome_limpo).encode('ASCII', 'ignore').decode('utf-8')
    nome_limpo = re.sub(r'[ \.\-]+', '_', nome_limpo) 
    nome_limpo = ''.join(c for c in nome_limpo if c.isalnum() or c == '_')
    nome_limpo = re.sub(r'_+', '_', nome_limpo).strip('_')

    # 2. DICIONÁRIO ENXUTO
    excecoes = {
        "USUARIO": "LOGIN", "SISTEMA_ELETRONICO": "SEI",
        "CPF": "NR_CPF", "MATRICULA": "NR_MATRICULA",
        "MATRICULA_1": "NR_MATRICULA", "MAT": "NR_MATRICULA", "CD_MATRICULA": "NR_MATRICULA",
        "MATRICULA_SIAPE": "NR_MATRICULA_SIAPE", 
        "MATRICULA_DPF": "NR_MATRICULA_DPF",
        "MATR_DPF": "NR_MATRICULA_DPF", 
        "MATR_SIAPE": "NR_MATRICULA_SIAPE",
        "TITULO_ELEITORAL": "NR_TITULO_ELEITORAL", "NO": "NR_NUMERO_ORDEM",
        
        "SERVIDOR": "NO_SERVIDOR", "NM_SERVIDOR": "NO_SERVIDOR", "NOME": "NO_SERVIDOR",
        "NO_PESSOA": "NO_SERVIDOR", "PESSOA": "NO_SERVIDOR",
        
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
# LÓGICA PRINCIPAL (ARQUIVO ÚNICO)
# =======================================================
def processar_arquivo_unico(caminho_origem, caminho_destino=None):
    # Configurar LOG na tela
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    logging.info("="*50)
    logging.info(f"INICIANDO PROCESSAMENTO: {os.path.basename(caminho_origem)}")
    logging.info("="*50)

    # Define o caminho de destino caso não tenha sido fornecido
    if not caminho_destino:
        diretorio, nome_arquivo = os.path.split(caminho_origem)
        nome_sem_extensao = os.path.splitext(nome_arquivo)[0]
        caminho_destino = os.path.join(diretorio, f"{nome_sem_extensao}_PADRONIZADO.xlsx")

    try:
        xls = pd.ExcelFile(caminho_origem)
        linha_cabecalho, melhor_match, aba_correta = 0, 0, None
        
        # Leitura Multi-Aba
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
            logging.warning("[IGNORADO] Nenhuma coluna alvo encontrada neste arquivo.")
            return

        logging.info(f"[LENDO] Aba: '{aba_correta}' | Linha Cabeçalho: {linha_cabecalho}")
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
            if col.startswith('NR_') or col.startswith('CD_'):
                df[col] = df[col].astype(str)
                df[col] = df[col].str.replace(r'\.0$', '', regex=True)
                df[col] = df[col].replace(['nan', '<NA>', 'NaN', 'None'], '')
                df[col] = df[col].str.strip()
        # ======================================================================

        # 5. Exportação com AutoFit
        cols_finais = [c for c in COLUNAS_ALVO if c in df.columns]
        col_faltantes = [c for c in COLUNAS_ALVO if c not in df.columns]
        df_final = df[cols_finais].copy()
        
        with pd.ExcelWriter(caminho_destino, engine='openpyxl') as writer:
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

        logging.info(f"\n[SUCESSO] Arquivo salvo em:\n -> {caminho_destino}")
        logging.info(f"Colunas mantidas ({len(cols_finais)}): {cols_finais}")
        if col_faltantes:
            logging.info(f"Colunas ausentes no original: {col_faltantes}")
        logging.info("="*50)

    except Exception as e:
        logging.error(f"[ERRO CRÍTICO] Falha ao processar arquivo: {e}", exc_info=True)


if __name__ == "__main__":
    processar_arquivo_unico(ARQUIVO_ORIGEM, ARQUIVO_DESTINO)