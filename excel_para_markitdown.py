from pathlib import Path
from markitdown import MarkItDown
#SALVA TABELAS EXCEL EM ARQUIVOS .TXT
# 1. Defina o caminho do arquivo original usando 'r' antes das aspas para evitar erros com as barras do Windows
caminho_excel = Path(r"C:\Users\guilherme.gfw\Documents\EFETIVOS_PADRONIZADOS_R01\EFETIVO 2006\LIS - 12_DEZEMBRO.xlsx")

# 2. Cria o caminho do TXT substituindo a extensão original por .txt (mantendo a mesma pasta e nome)
caminho_txt = caminho_excel.with_suffix(".txt")

# 3. Inicializa o conversor do MarkItDown
conversor = MarkItDown()

# 4. Executa a conversão passando o caminho como string
resultado = conversor.convert(str(caminho_excel))

# 5. Salva o conteúdo de texto diretamente no novo arquivo (usando utf-8 para evitar erros com acentos)
caminho_txt.write_text(resultado.text_content, encoding="utf-8")

print(f"Texto extraído e salvo com sucesso em:\n{caminho_txt}")