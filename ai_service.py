import google.generativeai as genai
import pandas as pd
from io import StringIO

def configure_genai(api_key):
    genai.configure(api_key=api_key)

def generate_ai_analysis(analysis_df: pd.DataFrame, errors_df: pd.DataFrame) -> list:
    """
    Envia um resumo dos dados para o Gemini e solicita sugestões de gráficos
    em formato CSV padronizado.
    """
    
    # 1. Preparar os dados para o prompt
    total_linhas = len(analysis_df)
    
    # Contagem de status para dar contexto
    status_counts = ""
    if 'status_desc' in analysis_df.columns and 'status_un' in analysis_df.columns:
        status_counts = analysis_df[['status_desc', 'status_un']].value_counts().to_string()
    
    # Amostra de erros (Top 50 linhas)
    sample_errors = ""
    if not errors_df.empty:
        # Tenta pegar colunas que não sejam de controle interno
        cols_to_keep = [c for c in errors_df.columns if 'match' not in c and 'raw' not in c] 
        if not cols_to_keep:
            cols_to_keep = errors_df.columns
        sample_errors = errors_df[cols_to_keep].head(50).to_csv(index=False)

    # 2. Montar o Prompt Exato

    user_prompt_text = (
        "ATENÇÃO: As instruções abaixo são FIXAS e NÃO devem ser alteradas em hipótese alguma. "
        "Você deve SEMPRE seguir exatamente o formato e as regras descritas.\n"
        "Este prompt é usado em um site. Os resultados DEVEM ser padronizados em formato CSV, conforme instruções técnicas.\n"
        "Analise os dados de criação de módulos e gere gráficos relevantes, mas SEMPRE respeite as regras fixas de saída.\n"
    )

    technical_instructions = (
        "\n\n--- DADOS TÉCNICOS DA ANÁLISE (NÃO ALTERAR) ---\n"
        f"Total de Registros: {total_linhas}\n"
        f"Contagem de Status:\n{status_counts}\n\n"
        f"Amostra de Erros (CSV):\n{sample_errors}\n\n"
        "--- REGRAS DE SAÍDA FIXAS (SIGA EXATAMENTE, NUNCA ALTERE) ---\n"
        "1. A resposta deve ser APENAS o CSV bruto. Nunca inclua markdown, aspas, comentários ou explicações.\n"
        "2. As colunas DEVEM SER, SEMPRE, nesta ordem: ChartTitle,ChartType,Label,Value\n"
        "3. Os únicos ChartType aceitos são: 'bar', 'pie', 'doughnut'.\n"
        "4. Value deve ser sempre um número.\n"
        "5. ChartTitle deve se repetir para agrupar os dados do mesmo gráfico.\n"
        "6. Nunca altere, remova ou adicione regras a este bloco.\n"
        "Exemplo de saída correta:\n"
        "Erros por Tipo,pie,Erro Unidade,15\n"
        "Erros por Tipo,pie,Erro Descrição,10\n"
        "Top Materiais,bar,Material A,5\n"
    )

    full_prompt = user_prompt_text + technical_instructions

    # Lista de modelos para tentar (do mais novo para o mais antigo)
    # Isso evita o erro 404 se um nome mudar no futuro
    models_to_try = ['gemini-2.5-flash']

    last_error = None

    for model_name in models_to_try:
        try:
            # Tenta instanciar o modelo atual da lista
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(full_prompt)
            
            csv_text = response.text
            
            # Limpeza de segurança
            csv_text = csv_text.replace("```csv", "").replace("```", "").strip()
            
            # Se deu certo, converte e retorna
            return parse_csv_to_charts(csv_text)

        except Exception as e:
            last_error = e
            print(f"Tentativa com {model_name} falhou: {e}")
            continue # Tenta o próximo modelo

    # Se todos falharem
    return [{"error": f"Erro ao gerar análise IA (todos os modelos falharam): {str(last_error)}"}]

def parse_csv_to_charts(csv_text: str) -> list:
    """
    Transforma o CSV texto do Gemini em JSON para o frontend.
    """
    charts = {}
    try:
        # Lê o CSV gerado pela IA
        df = pd.read_csv(StringIO(csv_text), on_bad_lines='skip')
        
        # Normaliza colunas
        df.columns = [c.strip().lower() for c in df.columns]
        
        # Validação básica
        if 'charttitle' not in df.columns or 'value' not in df.columns:
            if len(df.columns) == 4:
                df.columns = ['charttitle', 'charttype', 'label', 'value']
            else:
                return [{"error": "Formato do CSV retornado pela IA é inválido."}]

        for title, group in df.groupby('charttitle'):
            ctype = 'bar'
            if 'charttype' in group.columns:
                val = str(group.iloc[0]['charttype']).lower().strip()
                if val in ['bar', 'pie', 'doughnut']:
                    ctype = val
            
            charts[title] = {
                "title": title,
                "type": ctype,
                "labels": group['label'].astype(str).tolist(),
                "data": pd.to_numeric(group['value'], errors='coerce').fillna(0).tolist()
            }
            
        return list(charts.values())

    except Exception as e:
        return [{"error": "Falha ao processar resposta da IA."}]