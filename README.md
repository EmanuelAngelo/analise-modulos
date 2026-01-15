# Web-Analise

## Descrição

Este projeto é uma aplicação web para análise e comparação de bases de dados em Excel, focada em comparar informações entre diferentes fontes: SAP, ORÇAFASCIO e CADERNO. O sistema processa um arquivo .xlsx enviado pelo usuário, gera um relatório de divergências e disponibiliza um resumo visual e o download do resultado.

## Funcionalidades
- Upload de arquivo Excel (.xlsx) contendo as abas esperadas:
  - BASE ( SE AUT LD)
  - BASE ( SAP )
  - BASE ( ORÇAFASCIO)
  - BASE ( CADERNO)
- Análise automática das colunas: COD SAP, DESCRICAO, UNIDADE
- Geração de relatório de divergências entre as bases
- Visualização gráfica dos resultados (CORRETO × A VERIFICAR)
- Download do relatório processado em Excel

## Como usar
1. Instale as dependências:
   ```bash
   pip install -r requirements.txt.txt
   ```
2. Execute a aplicação:
   ```bash
   python app.py
   ```
3. Acesse `http://127.0.0.1:5000` no navegador.
4. Envie seu arquivo .xlsx conforme o modelo esperado.
5. Visualize o resumo e baixe o relatório gerado.

## Estrutura do Projeto
- `app.py`: Aplicação Flask, rotas e lógica web
- `analyze_core.py`: Núcleo de análise e comparação dos dados
- `templates/index.html`: Interface web (frontend)
- `static/style.css`: Estilos da interface
- `uploads/`: Arquivos enviados
- `outputs/`: Relatórios gerados

## Requisitos
- Python 3.8+
- Flask
- pandas
- openpyxl
- Werkzeug

## Observações
- O arquivo Excel enviado deve conter as abas e colunas conforme especificado.
- O sistema prioriza performance e clareza visual dos resultados.

## Licença
Este projeto é de uso interno e não possui licença aberta definida.
