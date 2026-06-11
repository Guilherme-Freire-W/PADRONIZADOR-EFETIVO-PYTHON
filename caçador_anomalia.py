import pandas as pd
import unicodedata
import re
import os

# =======================================================
# CONFIGURAÇÃO DOS DIRETÓRIOS
# =======================================================
PASTA_ORIGEM = r"C:\Users\guilherme.gfw\Documents\EFETIVOS"

# =======================================================
# LISTA DE COLUNAS CONHECIDAS (O caçador vai ignorar estas)
# Inclui as bases "LOTACAO" e "CARGO" que a inteligência resolve depois
# =======================================================
COLUNAS_CONHECIDAS = [
    'NR_MATRICULA_SIAPE', 'NR_MATRICULA', 'NR_MATRICULA_PF',
    'NR_CPF', 'NO_SERVIDOR',
    'NO_LOTACAO', 'SG_LOTACAO', 'CD_LOTACAO', 'LOTACAO',
    'CD_CARGO', 'NO_CARGO', 'SG_CARGO', 'SG_CARGOS', 'NO_CARGO_EFETIVO', 'CARGO',
    'CD_CLASSE', 'CD_PADRAO'
]

def normalizar_coluna(nome):
    """Exatamente o mesmo motor de limpeza do script de padronização."""
    if pd.isna(nome): return ""
    if not isinstance(nome, str): nome = str(nome)
    nome_limpo = nome.strip().upper()

    excecoes = {
        "USUARIO": "LOGIN", "SISTEMA_ELETRONICO": "SEI",
        "CPF": "NR_CPF", "MATRICULA": "NR_MATRICULA",
        "MATRICULA_1": "NR_MATRICULA", "MATRÍCULA": "NR_MATRICULA", "MAT": "NR_MATRICULA",
        "MATRICULA                .1": "NR_MATRICULA", 
        "MATRICULA_SIAPE": "NR_MATRICULA_SIAPE", "MATRICULA_DPF": "NR_MATRICULA_DPF",
        "MATRÍCULA SIAPE": "NR_MATRICULA_DPF", "MATRICULA-SIAPE": "NR_MATRICULA_SIAPE",
        "MATRICULA SIAPE ": "NR_MATRICULA_SIAPE",
        "TITULO_ELEITORAL": "NR_TITULO_ELEITORAL", "NO": "NR_NUMERO_ORDEM",
        
        "SERVIDOR": "NO_SERVIDOR", "NM_SERVIDOR": "NO_SERVIDOR", "NOME": "NO_SERVIDOR",
        "NO PESSOA": "NO_SERVIDOR", "NO_PESSOA": "NO_SERVIDOR", "PESSOA": "NO_SERVIDOR",
        
        "ATIVIDADE": "NO_ATIVIDADE", 
        "CARGO": "NO_CARGO", "GRUPO CARGO EMPREGO": "NO_CARGO",
        "CARGOS": "SG_CARGOS", 
        "CARGO_EFETIVO": "NO_CARGO_EFETIVO", 
        "CARRREIRA": "NO_CARREIRA", "CARREIRA": "NO_CARREIRA",
        "FUNCAO": "NO_FUNCAO", "NOVA_FUNCAO": "NO_NOVA_FUNCAO", 
        "CARGO_SIGLA": "SG_CARGO", "CARGO_CLASSE": "SG_CARGOCLASSE",
        
        "LOTACAO": "LOTACAO", "LOT": "LOTACAO", "LOT_1": "LOTACAO",
        "LOTACAO_DO_SERVIDOR": "LOTACAO", 
        "CO-UORG-LOTACAO-SERVIDOR": "LOTACAO",
        "SG UNIDADE LOTACAO": "SG_LOTACAO", "SG_UNIDADE_LOTACAO": "SG_LOTACAO",
        "SG UNIDADE EXERCICIO": "SG_LOTACAO", "SG_UNIDADE_EXERCICIO": "SG_LOTACAO",
        "CD UORG EXERCICIO": "CD_LOTACAO", "CD_UORG_EXERCICIO": "CD_LOTACAO",
        "CD UORG": "CD_LOTACAO", "CD_UORG": "CD_LOTACAO",
        "UNIDADE LOTACAO": "LOTACAO", "UNIDADE EXERCICIO": "LOTACAO",
        
        "CLASSE": "CD_CLASSE", "CO-CLASSE": "CD_CLASSE", 
        "PADRAO": "CD_PADRAO", "CO-PADRAO": "CD_PADRAO"
    }

    if nome_limpo in excecoes: return excecoes[nome_limpo]

    nome = unicodedata.normalize('NFKD', nome_limpo).encode('ASCII', 'ignore').decode('utf-8')
    nome = nome.replace(' ', '_').replace('.', '_').replace('-', '_')
    nome = ''.join(c for c in nome if c.isalnum() or c == '_')
    while '__' in nome: nome = nome.replace('__', '_')
    
    nome = re.sub(r'^(CO_|COD_|CODIGO_)', 'CD_', nome)
    nome = re.sub(r'^(DA_|DATA_|DTA_)', 'DT_', nome)
    nome = re.sub(r'^(NUM_|NUMERO_|NU_)', 'NR_', nome)
    nome = re.sub(r'^(VAL_|VALOR_|VL_)', 'VR_', nome)
    nome = re.sub(r'^(IND_|INDICADOR_|ID_)', 'IC_', nome)
    nome = re.sub(r'^(NOME_|NM_)', 'NO_', nome)
    
    return nome[:30].rstrip('_')


def buscar_anomalias(pasta_raiz):
    print("=== INICIANDO CAÇADA DE ANOMALIAS NAS PLANILHAS ===")
    print("Aguarde, isso pode levar alguns minutos...")
    
    anomalias = {}
    arquivos_lidos = 0

    for pasta_atual, _, arquivos in os.walk(pasta_raiz):
        for arquivo in arquivos:
            if arquivo.lower().endswith(('.xls', '.xlsx')) and not arquivo.startswith('~$'):
                caminho_completo = os.path.join(pasta_atual, arquivo)
                
                try:
                    xls = pd.ExcelFile(caminho_completo)
                    arquivos_lidos += 1
                    
                    # Varre todas as abas
                    for aba in xls.sheet_names:
                        df_topo = pd.read_excel(xls, sheet_name=aba, nrows=5, header=None)
                        
                        # Tenta encontrar a linha de cabeçalho
                        linha_cabecalho, melhor_match = 0, 0
                        for idx, row in df_topo.iterrows():
                            simuladas = [normalizar_coluna(val) for val in row.values]
                            match = [c for c in COLUNAS_CONHECIDAS if c in simuladas]
                            if len(match) > melhor_match:
                                melhor_match, linha_cabecalho = len(match), idx
                        
                        # Se achou algo que parece um cabeçalho, avalia as colunas
                        if melhor_match > 0:
                            df = pd.read_excel(xls, sheet_name=aba, header=linha_cabecalho, nrows=0)
                            
                            # Limpa sufixos do Pandas
                            colunas_reais = [str(col).rsplit('.', 1)[0] if '.' in str(col) and str(col).rsplit('.', 1)[-1].isdigit() else str(col) for col in df.columns]
                            
                            for col_original in colunas_reais:
                                nome_normalizado = normalizar_coluna(col_original)
                                
                                # Verifica se a coluna caiu fora do padrão e não é nula/vazia
                                if nome_normalizado and nome_normalizado not in COLUNAS_CONHECIDAS:
                                    if nome_normalizado not in anomalias:
                                        anomalias[nome_normalizado] = {'originais': set(), 'arquivos': set()}
                                    
                                    anomalias[nome_normalizado]['originais'].add(col_original)
                                    # Salva apenas o nome do arquivo e a pasta para o log não ficar gigante
                                    anomalias[nome_normalizado]['arquivos'].add(f"{os.path.basename(pasta_atual)}/{arquivo}")

                except Exception as e:
                    print(f"Erro ao ler {arquivo}: {e}")

    # Geração do Relatório
    log_path = os.path.join(pasta_raiz, "relatorio_anomalias_detectadas.txt")
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write(f"RELATÓRIO DE ANOMALIAS (COLUNAS NÃO MAPEADAS)\n")
        f.write(f"Total de arquivos verificados: {arquivos_lidos}\n")
        f.write("="*60 + "\n")
        
        # Ordena alfabeticamente para facilitar a leitura
        for nome_norm in sorted(anomalias.keys()):
            info = anomalias[nome_norm]
            f.write(f"\n[ANOMALIA DETECTADA]: {nome_norm}\n")
            f.write(f"  > Como estava escrito no Excel: {', '.join(info['originais'])}\n")
            # Mostra até 3 arquivos de exemplo para você saber onde achar
            f.write(f"  > Visto em (exemplos): {list(info['arquivos'])[:3]}\n")
            
    print(f"\nVarredura concluída! {arquivos_lidos} arquivos analisados.")
    print(f"Abra o arquivo: '{log_path}' para ver os resultados.")

if __name__ == "__main__":
    buscar_anomalias(PASTA_ORIGEM)