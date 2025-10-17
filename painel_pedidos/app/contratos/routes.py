# Arquivo: app/contratos/routes.py (VERSÃO FINAL, 100% COMPLETA E MESCLADA)

import re
from datetime import datetime
from flask import Blueprint, request, jsonify, send_file
from app import db
from app.models import Pedido

# --- IMPORTAÇÕES ATUALIZADAS ---
# Importamos 'Dict' e 'Any' para as anotações de tipo (type hints),
# o que deixa o código mais claro para a equipe.
from typing import Dict, Any

# Importamos nossas novas e renomeadas funções do Extractor.py.
# Os nomes são claros e refletem suas funções específicas.
from app.Extractor import (
    extrair_dados_do_contrato_por_tipo,
    gerar_contrato_docx,
    gerar_contrato_pdf_direto
)

contratos_bp = Blueprint('contratos', __name__, url_prefix='/api/contracts')

# ==============================================================================
# ROTA DE UPLOAD E ANÁLISE (ATUALIZADA)
# ==============================================================================
@contratos_bp.route('/upload', methods=['POST'])
def upload_contract():
    """
    Recebe um arquivo PDF via upload, extrai os dados e os retorna ao frontend
    para que o usuário possa revisar antes de salvar como um novo pedido.
    """
    # Validação de segurança e de dados de entrada
    user_id = request.headers.get('X-User-Id')
    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401

    if 'file' not in request.files:
        return jsonify({'message': 'Nenhum arquivo enviado.'}), 400

    file = request.files['file']
    if file.filename == '' or not file.filename.endswith('.pdf'):
        return jsonify({'message': 'Nenhum arquivo PDF selecionado.'}), 400

    # MUDANÇA IMPORTANTE: Lê o 'tipo_analise' que o frontend envia
    # com base no botão que o usuário clicou ('sistema' ou 'padrao').
    tipo_analise = request.form.get('tipo_analise', 'padrao')

    try:
        # OTIMIZAÇÃO: Lê o arquivo em memória (file.read()) em vez de salvá-lo no disco.
        # É mais rápido e não deixa lixo no servidor.
        pdf_bytes = file.read()

        # CHAMA NOSSA FUNÇÃO INTELIGENTE: Ela escolhe o método de extração correto (Regex ou IA).
        dados_extraidos = extrair_dados_do_contrato_por_tipo(pdf_bytes, tipo_analise)

        if not dados_extraidos:
            return jsonify({'message': 'Não foi possível extrair dados do contrato.'}), 500

        # Retorna os dados para o frontend para o usuário revisar.
        # A estrutura deste JSON corresponde ao que o JavaScript espera.
        return jsonify({
            'message': 'Dados extraídos com sucesso! Revise para salvar.',
            'extractedData': dados_extraidos,
        }), 200

    except Exception as e:
        print(f"[ERRO] Falha na rota /upload: {e}")
        return jsonify({'message': f'Erro ao processar o contrato: {str(e)}'}), 500

# ==============================================================================
# ROTA PARA GERAR NOVOS CONTRATOS (ATUALIZADA)
# ==============================================================================
@contratos_bp.route('/gerar-contrato', methods=['POST'])
def gerar_contrato():
    """
    Recebe dados de um formulário via JSON, escolhe o formato de documento
    desejado (DOCX ou PDF) e chama a função apropriada para gerar e
    enviar o arquivo de volta para o usuário.
    """
    user_id = request.headers.get('X-User-Id')
    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401

    data = request.json
    # Lê o formato que o frontend envia com base no botão clicado (ex: 'pdf' ou 'docx').
    formato_desejado = data.get('formato_desejado', 'docx')

    # Mapeamento dos dados do formulário para o formato padrão que nossas funções geradoras esperam.
    # Isso garante consistência e desacopla o frontend do backend.
    contratante_info = {
        'Nome': data.get('contratanteNome'), 'RG': data.get('contratanteRg'),
        'CPF': data.get('contratanteCpf'), 'Endereco': data.get('contratanteEndereco'),
        'Telefone': data.get('contratanteTelefone'), 'Email': data.get('contratanteEmail'),
    }
    dados_para_contrato = {
        'Contratante': contratante_info,
        'Data do Evento': data.get('dataEvento'),
        'Local do Evento': data.get('localEvento'),
        'Produtos Contratados': data.get('produtosContratados', []),
        'Valor Total do Pedido': data.get('valorTotalPedidoContrato'),
        'Data de Pagamento': data.get('dataPagamentoContrato'),
        'Forma de Pagamento': data.get('formaPagamento'),
        'Como nos conheceu': data.get('comoConheceu'),
        'Responsavel': data.get('responsavelContrato'),
    }

    try:
        # Delega a lógica de geração e envio para uma função auxiliar para manter a rota limpa.
        response, error_message = _processar_e_enviar_contrato(dados_para_contrato, formato_desejado)
        if error_message:
            return jsonify({'message': error_message}), 500
        return response
        
    except Exception as e:
        print(f"[ERRO] Não foi possível gerar ou enviar o contrato: {e}")
        return jsonify({'message': f"Erro ao gerar o contrato: {str(e)}"}), 500

# --- FUNÇÃO AUXILIAR PARA GERAÇÃO E ENVIO DE DOCUMENTOS ---
def _processar_e_enviar_contrato(dados: Dict[str, Any], formato: str):
    """
    Função interna que chama o gerador correto (DOCX ou PDF) e envia o arquivo.
    """
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    # Limpa o nome do cliente para usar em nomes de arquivo seguros (evita erros com espaços ou caracteres especiais).
    nome_cliente_safe = re.sub(r'[^\w-]', '_', dados['Contratante'].get('Nome', 'Contrato')).lower()
    nome_base = f"contrato_{nome_cliente_safe}_{timestamp}"

    if formato == 'pdf':
        # --- NOVA LÓGICA DE GERAÇÃO DIRETA DE PDF ---
        print("[INFO] Gerando contrato em formato PDF via WeasyPrint...")
        pdf_stream = gerar_contrato_pdf_direto(dados)
        if not pdf_stream:
            return None, "Erro interno ao gerar o documento PDF."
        
        return send_file(
            pdf_stream,
            as_attachment=True,
            download_name=f"{nome_base}.pdf",
            mimetype='application/pdf'
        ), None
    
    else: # O padrão é 'docx'
        # --- LÓGICA DE GERAÇÃO DE DOCX ---
        print("[INFO] Gerando contrato em formato DOCX...")
        doc_stream = gerar_contrato_docx(dados)
        if not doc_stream:
            return None, "Erro interno ao gerar o documento DOCX."

        return send_file(
            doc_stream,
            as_attachment=True,
            download_name=f"{nome_base}.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        ), None