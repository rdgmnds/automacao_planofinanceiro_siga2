# Automação de sistema web para a FGV

## Essa automação foi desenvolvida para gerar os relatórios financeiros dos alunos. Os relatórios são extraídos no sistema web integrado da empresa. O projeto economizou mais de 90% de trabalho.

### Principais tecnologias utilizadas
1. Python como linguagem;
2. PlayWright para automação web;
3. Openpyxl para manipular Excel;
4. Streamlit para interface do usuário.

### Objetivos da automação
1. Economizar tempo na geração de relatórios;
2. Diminuir a quantidade de funcionários responsáveis pela emissão dos relatórios;
3. Permitir que o funcionário atenda outras demandas enquanto os relatórios são gerados pela automação.

### Fluxo
1. O funcionário importa o relatório .xlsx com as informações dos alunos na ferramenta;
2. A automação acessa o sistema integrado da empresa;
3. Itera a planilha com as informações dos alunos e extrai o relatório de cada um;
4. Guarda cada relatório na pasta de rede da empresa;
5. Envia uma mensagem de sucesso ao terminar.

### Tempo de desenvolvimento
O projeto foi concluído em dois dias de trabalho.

### Desenvolvedores responsáveis
Rodrigo Mendes Vasconcelos
