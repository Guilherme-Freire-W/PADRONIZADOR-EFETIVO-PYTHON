Padronizador de Efetivos (Python)

Este projeto contém um conjunto de scripts em Python desenvolvidos para automatizar a extração, normalização e padronização de planilhas de efetivos (arquivos .xls e .xlsx). O objetivo principal é transformar dados brutos e inconsistentes em bases de dados limpas, estruturadas e prontas para análise.

🚀 Funcionalidades Principais
Canonicalização Inteligente: Utiliza algoritmos de limpeza de strings para identificar colunas, mesmo que existam variações de digitação, acentuação ou pontuação.

Normalização Semântica: Identifica automaticamente o conteúdo das colunas (matrículas, cargos, lotações) para aplicar prefixos padronizados (NR_, CD_, SG_, NO_).

Tratamento de Notação Científica: Garante que matrículas e códigos numéricos sejam tratados como texto puro, evitando a perda de precisão ou formatação comum no Excel.

Processamento Multi-Aba: O sistema varre automaticamente diferentes abas de uma planilha até encontrar o cabeçalho correto.

Ajuste Automático: Exporta os arquivos finais com colunas formatadas (AutoFit), facilitando a leitura imediata.

Detecção de Anomalias: Inclui um script dedicado (caçador_anomalia.py) que audita grandes diretórios para identificar colunas não mapeadas ou inconsistências nos dados históricos.

🛠 Tecnologias Utilizadas
Python 3.x

Pandas: Manipulação e processamento de dados.

OpenPyXL: Leitura/escrita de arquivos Excel e formatação visual.

Unicodedata & Re: Limpeza e normalização de strings.

📁 Estrutura dos Scripts
normalizar_diretorio_excel.py: Processa pastas inteiras, gerando logs de auditoria e tratando centenas de arquivos em lote.

normalizar_arquivo_excel.py: Versão otimizada para o processamento pontual de arquivos individuais.

caçador_anomalia.py: Ferramenta de auditoria para identificar padrões divergentes em toda a base de dados histórica.

excel_para_markitdown.py: Utilitário de conversão para preparação de relatórios.

Como utilizar
Para rodar os scripts, certifique-se de ter as bibliotecas instaladas:

Bash
pip install pandas openpyxl 

Basta configurar o caminho do arquivo ou diretório de origem na variável ARQUIVO_ORIGEM dentro do script desejado e executar:

Bash 
python normalizar_arquivo_excel.py 
