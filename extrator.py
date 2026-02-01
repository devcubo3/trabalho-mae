"""
Módulo de extração de dados de extratos bancários usando OpenAI Vision.
Converte PDF em imagens e usa GPT-4o para extrair lançamentos estruturados.
"""

import base64
import json
import io
import fitz  # PyMuPDF
from openai import OpenAI


SYSTEM_PROMPT = """Você é um especialista em extrair dados de extratos bancários brasileiros.

Sua tarefa é analisar a imagem de uma página de extrato bancário e extrair TODOS os lançamentos.

REGRAS CRÍTICAS:
1. Extraia ABSOLUTAMENTE TODOS os lançamentos visíveis na página. Não pule nenhum.
2. Mantenha a ordem exata em que aparecem no extrato.
3. Se um lançamento tem valor na coluna "Crédito (R$)", o tipo é "C".
4. Se um lançamento tem valor na coluna "Débito (R$)", o tipo é "D".
5. A data pode estar em branco se for a mesma data do lançamento anterior - nesse caso, repita a data anterior.
6. O valor deve ser extraído SEM o "R$", apenas o número com vírgula (ex: "34.695,00").
7. Na descrição, inclua o texto principal E os detalhes (como DEST:, CONTR:, etc.) em uma única linha.
8. Ignore linhas de cabeçalho, rodapé, saldo anterior e saldo do dia.
9. ESTORNO deve ser tratado como crédito (C) pois devolve dinheiro.

Retorne APENAS um JSON válido no seguinte formato, sem markdown, sem ```json, apenas o JSON puro:
{
  "lancamentos": [
    {
      "data": "DD/MM/AAAA",
      "tipo": "C" ou "D",
      "descricao": "Descrição completa do lançamento",
      "valor": "1.234,56"
    }
  ],
  "pagina_tem_continuacao": true/false,
  "observacoes": "qualquer observação relevante sobre a extração"
}

Se a página não contiver lançamentos (ex: capa, resumo), retorne:
{
  "lancamentos": [],
  "pagina_tem_continuacao": false,
  "observacoes": "Página sem lançamentos"
}
"""


def pdf_para_imagens(pdf_path: str, dpi: int = 300) -> list[bytes]:
    """Converte cada página do PDF em imagem PNG de alta resolução."""
    doc = fitz.open(pdf_path)
    imagens = []

    for pagina in doc:
        # Zoom para alta resolução (300 DPI)
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = pagina.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        imagens.append(img_bytes)

    doc.close()
    return imagens


def extrair_lancamentos_pagina(client: OpenAI, imagem_bytes: bytes, numero_pagina: int, data_anterior: str = "") -> dict:
    """Extrai lançamentos de uma única página usando GPT-4o Vision."""

    img_base64 = base64.b64encode(imagem_bytes).decode("utf-8")

    user_message = f"Analise esta página {numero_pagina} do extrato bancário e extraia todos os lançamentos."
    if data_anterior:
        user_message += f"\nA última data da página anterior foi: {data_anterior}. Use-a para lançamentos sem data explícita."

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": user_message},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{img_base64}",
                            "detail": "high"
                        }
                    }
                ]
            }
        ],
        max_tokens=4096,
        temperature=0.0,  # Máxima precisão, sem criatividade
    )

    resposta_texto = response.usage
    conteudo = response.choices[0].message.content.strip()

    # Limpar possíveis markdown wrappers
    if conteudo.startswith("```"):
        conteudo = conteudo.split("\n", 1)[1] if "\n" in conteudo else conteudo[3:]
        if conteudo.endswith("```"):
            conteudo = conteudo[:-3]
        conteudo = conteudo.strip()

    try:
        dados = json.loads(conteudo)
    except json.JSONDecodeError:
        # Tenta extrair JSON de dentro do texto
        inicio = conteudo.find("{")
        fim = conteudo.rfind("}") + 1
        if inicio != -1 and fim > inicio:
            dados = json.loads(conteudo[inicio:fim])
        else:
            raise ValueError(f"Página {numero_pagina}: Resposta não é JSON válido: {conteudo[:200]}")

    return dados


def extrair_extrato_completo(pdf_path: str, api_key: str, progresso_callback=None) -> list[dict]:
    """
    Extrai todos os lançamentos de um extrato bancário em PDF.

    Args:
        pdf_path: Caminho do arquivo PDF
        api_key: Chave da API OpenAI
        progresso_callback: Função opcional chamada com (pagina_atual, total_paginas, mensagem)

    Returns:
        Lista de lançamentos ordenados por data
    """
    client = OpenAI(api_key=api_key)

    if progresso_callback:
        progresso_callback(0, 0, "Convertendo PDF em imagens...")

    imagens = pdf_para_imagens(pdf_path)
    total_paginas = len(imagens)

    if progresso_callback:
        progresso_callback(0, total_paginas, f"PDF tem {total_paginas} páginas. Iniciando extração...")

    todos_lancamentos = []
    data_anterior = ""

    for i, img in enumerate(imagens):
        numero_pagina = i + 1

        if progresso_callback:
            progresso_callback(numero_pagina, total_paginas, f"Analisando página {numero_pagina} de {total_paginas}...")

        try:
            resultado = extrair_lancamentos_pagina(client, img, numero_pagina, data_anterior)
            lancamentos = resultado.get("lancamentos", [])

            if lancamentos:
                # Atualiza data_anterior para a próxima página
                for lanc in reversed(lancamentos):
                    if lanc.get("data"):
                        data_anterior = lanc["data"]
                        break

                todos_lancamentos.extend(lancamentos)

                if progresso_callback:
                    progresso_callback(
                        numero_pagina, total_paginas,
                        f"Página {numero_pagina}: {len(lancamentos)} lançamentos extraídos."
                    )
            else:
                if progresso_callback:
                    obs = resultado.get("observacoes", "Sem lançamentos")
                    progresso_callback(numero_pagina, total_paginas, f"Página {numero_pagina}: {obs}")

        except Exception as e:
            if progresso_callback:
                progresso_callback(numero_pagina, total_paginas, f"ERRO na página {numero_pagina}: {str(e)}")
            # Continua com as próximas páginas mesmo com erro

    # Preencher datas vazias
    ultima_data = ""
    for lanc in todos_lancamentos:
        if lanc.get("data"):
            ultima_data = lanc["data"]
        else:
            lanc["data"] = ultima_data

    if progresso_callback:
        progresso_callback(total_paginas, total_paginas, f"Extração completa! {len(todos_lancamentos)} lançamentos encontrados.")

    return todos_lancamentos
