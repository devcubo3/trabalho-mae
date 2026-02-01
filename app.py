"""
App principal - Interface web para extração de extratos bancários.
Usa Flask com Server-Sent Events para progresso em tempo real.
"""

import os
import json
import uuid
from flask import Flask, render_template, request, send_file, Response
from extrator import extrair_extrato_completo
from gerador_docx import gerar_docx


def carregar_env():
    """Carrega variáveis do .env.local"""
    env_path = os.path.join(os.path.dirname(__file__), ".env.local")
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for linha in f:
                linha = linha.strip()
                if linha and not linha.startswith("#") and "=" in linha:
                    chave, valor = linha.split("=", 1)
                    os.environ[chave.strip()] = valor.strip()


carregar_env()

OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")

app = Flask(__name__)

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
RESULT_DIR = os.path.join(os.path.dirname(__file__), "resultados")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/processar", methods=["POST"])
def processar():
    pdf = request.files.get("pdf")
    api_key = request.form.get("api_key", "").strip() or OPENAI_API_KEY
    banco = request.form.get("banco", "BRADESCO").strip()
    agencia = request.form.get("agencia", "3050").strip()
    conta = request.form.get("conta", "7223-0").strip()

    if not pdf or not api_key:
        return Response(
            f"data: {json.dumps({'tipo': 'erro', 'mensagem': 'PDF e chave API são obrigatórios'})}\n\n",
            mimetype="text/event-stream"
        )

    # Salvar PDF temporário
    job_id = str(uuid.uuid4())[:8]
    pdf_path = os.path.join(UPLOAD_DIR, f"{job_id}.pdf")
    pdf.save(pdf_path)

    def gerar_eventos():
        mensagens = []

        def callback_progresso(pagina, total, mensagem):
            mensagens.append({
                "tipo": "progresso",
                "pagina": pagina,
                "total": total,
                "mensagem": mensagem
            })

        try:
            # Extrair lançamentos (o callback é síncrono, chamado durante o processamento)
            # Como precisamos de streaming, vamos processar de forma diferente
            yield f"data: {json.dumps({'tipo': 'progresso', 'pagina': 0, 'total': 0, 'mensagem': 'Iniciando processamento...'})}\n\n"

            lancamentos = []
            erro_ocorreu = False

            # Importar funções internas para controle granular
            from extrator import pdf_para_imagens, extrair_lancamentos_pagina
            from openai import OpenAI

            client = OpenAI(api_key=api_key)

            yield f"data: {json.dumps({'tipo': 'progresso', 'pagina': 0, 'total': 0, 'mensagem': 'Convertendo PDF em imagens de alta resolução...'})}\n\n"

            imagens = pdf_para_imagens(pdf_path)
            total = len(imagens)

            yield f"data: {json.dumps({'tipo': 'progresso', 'pagina': 0, 'total': total, 'mensagem': f'PDF tem {total} páginas. Enviando para análise da IA...'})}\n\n"

            data_anterior = ""

            for i, img in enumerate(imagens):
                num_pag = i + 1

                yield f"data: {json.dumps({'tipo': 'progresso', 'pagina': num_pag, 'total': total, 'mensagem': f'Analisando página {num_pag} de {total} com GPT-4o Vision...'})}\n\n"

                try:
                    resultado = extrair_lancamentos_pagina(client, img, num_pag, data_anterior)
                    lancs = resultado.get("lancamentos", [])

                    if lancs:
                        for lanc in reversed(lancs):
                            if lanc.get("data"):
                                data_anterior = lanc["data"]
                                break
                        lancamentos.extend(lancs)

                        yield f"data: {json.dumps({'tipo': 'progresso', 'pagina': num_pag, 'total': total, 'mensagem': f'Página {num_pag}: {len(lancs)} lançamentos extraídos.'})}\n\n"
                    else:
                        obs = resultado.get("observacoes", "Sem lançamentos")
                        yield f"data: {json.dumps({'tipo': 'progresso', 'pagina': num_pag, 'total': total, 'mensagem': f'Página {num_pag}: {obs}'})}\n\n"

                except Exception as e:
                    yield f"data: {json.dumps({'tipo': 'progresso', 'pagina': num_pag, 'total': total, 'mensagem': f'ERRO página {num_pag}: {str(e)}'})}\n\n"

            if not lancamentos:
                yield f"data: {json.dumps({'tipo': 'erro', 'mensagem': 'Nenhum lançamento encontrado no PDF.'})}\n\n"
                return

            # Preencher datas vazias
            ultima_data = ""
            for lanc in lancamentos:
                if lanc.get("data"):
                    ultima_data = lanc["data"]
                else:
                    lanc["data"] = ultima_data

            yield f"data: {json.dumps({'tipo': 'progresso', 'pagina': total, 'total': total, 'mensagem': f'Gerando documento DOCX com {len(lancamentos)} lançamentos...'})}\n\n"

            # Gerar DOCX
            nome_arquivo = f"movimentacao_{job_id}.docx"
            docx_path = os.path.join(RESULT_DIR, nome_arquivo)
            gerar_docx(lancamentos, docx_path, agencia=agencia, conta=conta, banco=banco)

            yield f"data: {json.dumps({'tipo': 'concluido', 'mensagem': f'Concluído! {len(lancamentos)} lançamentos processados.', 'arquivo': nome_arquivo, 'total_lancamentos': len(lancamentos)})}\n\n"

        except Exception as e:
            yield f"data: {json.dumps({'tipo': 'erro', 'mensagem': f'Erro geral: {str(e)}'})}\n\n"

        finally:
            # Limpar PDF temporário
            try:
                os.remove(pdf_path)
            except OSError:
                pass

    return Response(gerar_eventos(), mimetype="text/event-stream")


@app.route("/download/<nome_arquivo>")
def download(nome_arquivo):
    # Sanitizar nome do arquivo
    nome_arquivo = os.path.basename(nome_arquivo)
    caminho = os.path.join(RESULT_DIR, nome_arquivo)

    if not os.path.exists(caminho):
        return "Arquivo não encontrado", 404

    return send_file(
        caminho,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 50)
    print("  MOVIMENTAÇÃO BANCÁRIA - Extrator de Extratos")
    print(f"  Acesse: http://localhost:{port}")
    print("=" * 50)
    app.run(debug=True, port=port, threaded=True)
