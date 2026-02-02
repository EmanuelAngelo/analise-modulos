
# Web-Analise

## Descrição Geral
Web-Analise é uma aplicação web para análise, comparação e visualização de divergências entre múltiplas bases de dados Excel, com foco em projetos de engenharia e gestão de materiais. O sistema compara informações entre SAP, ORÇAFASCIO e CADERNO, gera relatórios detalhados e gráficos interativos, e conta com integração de Inteligência Artificial (Google Gemini) para sugestões automáticas de visualizações.

## Principais Funcionalidades
- **Upload de Excel (.xlsx)**: O usuário envia um arquivo contendo as abas:
  - BASE ( SE AUT LD) *(módulo)*
  - BASE ( SAP )
  - BASE ( ORÇAFASCIO)
  - BASE ( CADERNO)
- **Análise automática** das colunas: `COD SAP`, `DESCRICAO`, `UNIDADE`.
- **Comparação cruzada** entre as bases, identificando divergências de descrição e unidade.
- **Geração de relatório Excel** com três abas: `resumo`, `erros`, `analysis`.
- **Visualização gráfica** (stacked bar) dos resultados (CORRETO × A VERIFICAR).
- **Download automático** do relatório processado.
- **Sugestão de gráficos por IA**: Utiliza Google Gemini para sugerir visualizações adicionais a partir dos dados analisados.

## Fluxo de Uso
1. Instale as dependências:
   ```bash
   pip install -r requirements.txt.txt
   ```
2. Execute a aplicação:
   ```bash
   python app.py
   ```
3. Acesse `http://127.0.0.1:5000` no navegador.
4. Envie seu arquivo `.xlsx` conforme o modelo esperado.
5. Visualize o resumo gráfico e baixe o relatório gerado.
6. (Opcional) Utilize a integração IA para obter sugestões de gráficos customizados.

## Estrutura do Projeto
- `app.py`: Backend Flask, rotas, upload, processamento e download.
- `analyze_core.py`: Núcleo de análise, normalização, comparação e exportação dos dados.
- `ai_service.py`: Integração com Google Gemini para análise e sugestões de gráficos.
- `templates/index.html`: Interface web moderna, frontend responsivo e interativo.
- `static/style.css`: Estilos visuais customizados.
- `uploads/`: Armazena arquivos enviados temporariamente.
- `outputs/`: Relatórios Excel gerados para download.
- `requirements.txt.txt`: Dependências do projeto.

## Exemplo de Estrutura do Excel Esperado
Cada aba deve conter as colunas mínimas:

| COD SAP | DESCRICAO | UNIDADE |
|---------|-----------|---------|

## Exemplo de Uso (CLI)
```bash
pip install -r requirements.txt.txt
python app.py
# Acesse http://127.0.0.1:5000 e envie seu arquivo Excel
```

## Requisitos
- Python 3.8+
- Flask
- pandas
- openpyxl
- Werkzeug
- google-generativeai *(para integração IA)*

## Observações Importantes
- O arquivo Excel enviado deve conter as abas e colunas conforme especificado acima.
- O sistema prioriza performance, clareza visual e facilidade de uso.
- A integração IA é opcional, mas recomenda-se configurar a chave de API do Google Gemini para uso completo.

## Licença
Projeto de uso interno. Licença não definida.
