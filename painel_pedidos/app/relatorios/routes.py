import os
import time
from datetime import datetime
import openpyxl
import json
from flask import Blueprint, request, jsonify, current_app, send_file, after_this_request
from app import db
from app.models import Pedido
from app.Extractor import gerar_relatorio_entrega

relatorios_bp = Blueprint('relatorios', __name__, url_prefix='/api/reports')

@relatorios_bp.route('/export-planilha', methods=['GET'])
def export_planilha():
    user_id = request.headers.get('X-User-Id')
    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401

    data_inicio_str = request.args.get('dataInicio')
    data_fim_str = request.args.get('dataFim')
    tipo_produto = request.args.get('tipoProduto', '')

    if not data_inicio_str or not data_fim_str:
        return jsonify({'message': 'Datas de início e fim são obrigatórias.'}), 400

    try:
        # Construímos a consulta base
        query = Pedido.query.filter_by(user_id=user_id, status='confirmado')

        # Adicionamos os filtros de data e produto diretamente na consulta
        start_date = datetime.fromisoformat(data_inicio_str)
        end_date = datetime.fromisoformat(data_fim_str)
        query = query.filter(Pedido.dataRetirada.between(data_inicio_str, data_fim_str))

        if tipo_produto:
            query = query.filter_by(tipoPedido=tipo_produto)
        
        # Executamos a consulta no banco de dados
        pedidos_filtrados = query.all()

        print("Sql gerado para exportaçao")
        print(str(query))

    except ValueError:
        return jsonify({'message': 'Formato de data inválido. Use AAAA-MM-DD.'}), 400

    if not pedidos_filtrados:
        return jsonify({'message': 'Nenhum pedido confirmado encontrado para os filtros selecionados.'}), 404

    # A lógica de criação do Excel continua a mesma, mas agora com objetos Pedido
    excel_data_list = []
    for p in pedidos_filtrados:
        # Usamos nosso método to_dict() para facilitar, mas poderíamos acessar os atributos diretamente
        pedido_dict = p.to_dict()
        excel_data_list.append({
            'ID do Pedido': pedido_dict.get('id', ''),
            'Nome do Cliente': pedido_dict.get('clienteNome', ''),
            'Produto': pedido_dict.get('tipoPedido', ''),
            'Quantidade': pedido_dict.get('quantidade', ''),
            'Sabor': pedido_dict.get('sabores', ''),
            'Data de Retirada': pedido_dict.get('dataRetirada', ''),
            # Adicione outros campos conforme necessário
        })
    
    # ... (O restante da lógica de criar e enviar o arquivo Excel continua igual)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Pedidos Confirmados"
    headers = list(excel_data_list[0].keys()) if excel_data_list else []
    if headers: sheet.append(headers)
    for item in excel_data_list: sheet.append([item.get(h, '') for h in headers])

    excel_filename = f"pedidos_confirmados_{data_inicio_str}_a_{data_fim_str}.xlsx"
    temp_filepath = os.path.join(current_app.config.get('UPLOAD_FOLDER'), excel_filename)
    workbook.save(filename=temp_filepath)

    @after_this_request
    def remove_file(response):
        try: os.remove(temp_filepath)
        except Exception as e: print(f"Erro ao remover arquivo temporário: {e}")
        return response

    return send_file(temp_filepath, as_attachment=True, download_name=excel_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@relatorios_bp.route('/export-selected-pedidos', methods=['POST'])
def export_selected_pedidos():
    user_id = request.headers.get('X-User-Id')
    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401

    data = request.json
    selected_pedido_ids = data.get('pedido_ids', [])

    if not selected_pedido_ids:
        return jsonify({'message': 'Nenhum ID de pedido selecionado.'}), 400

    # Usamos o operador 'in_' para buscar múltiplos IDs de uma só vez
    pedidos_para_exportar = Pedido.query.filter(
        Pedido.id.in_(selected_pedido_ids),
        Pedido.user_id == user_id
    ).all()

    if not pedidos_para_exportar:
        return jsonify({'message': 'Nenhum pedido encontrado com os IDs fornecidos.'}), 404

    # ... (O restante da lógica de criar e enviar o arquivo Excel continua igual)
    excel_data_list = [p.to_dict() for p in pedidos_para_exportar]
    # Simplificação da criação do excel
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Pedidos Selecionados"
    headers = list(excel_data_list[0].keys()) if excel_data_list else []
    if headers: sheet.append(headers)
    for item in excel_data_list: sheet.append([item.get(h, '') for h in headers])

    excel_filename = f"pedidos_selecionados_{datetime.now().strftime('%Y%m%d')}.xlsx"
    temp_filepath = os.path.join(current_app.config.get('UPLOAD_FOLDER'), excel_filename)
    workbook.save(filename=temp_filepath)

    @after_this_request
    def remove_file(response):
        try: os.remove(temp_filepath)
        except Exception as e: print(f"Erro ao remover arquivo temporário: {e}")
        return response
        
    return send_file(temp_filepath, as_attachment=True, download_name=excel_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@relatorios_bp.route('', methods=['GET'])
def get_relatorios():
    user_id = request.headers.get('X-User-Id')
    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401

    # Buscamos todos os pedidos do usuário uma única vez
    user_pedidos = Pedido.query.filter_by(user_id=user_id).all()

    # A lógica de cálculo em Python continua a mesma, mas agora sobre dados vindos do DB
    current_month = datetime.now().month
    current_year = datetime.now().year
    pedidos_mes = 0
    for pedido in user_pedidos:
        # createdAt agora é um objeto datetime, não precisamos converter
        if pedido.createdAt.month == current_month and pedido.createdAt.year == current_year:
            pedidos_mes += 1

    produto_counts = {}
    for pedido in user_pedidos:
        tipo = pedido.tipoPedido or 'Não especificado'
        produto_counts[tipo] = produto_counts.get(tipo, 0) + pedido.quantidade
    sorted_produtos = sorted(produto_counts.items(), key=lambda item: item[1], reverse=True)
    produtos_mais_pedidos = f"{sorted_produtos[0][0]} ({sorted_produtos[0][1]})" if sorted_produtos else 'N/A'

    cliente_counts = {}
    for pedido in user_pedidos:
        cliente = pedido.clienteNome or 'Anônimo'
        cliente_counts[cliente] = cliente_counts.get(cliente, 0) + 1
    sorted_clientes = sorted(cliente_counts.items(), key=lambda item: item[1], reverse=True)
    clientes_mais_pedidos = f"{sorted_clientes[0][0]} ({sorted_clientes[0][1]} pedidos)" if sorted_clientes else 'N/A'

    return jsonify({
        'totalPedidosMes': pedidos_mes,
        'produtosMaisPedidos': produtos_mais_pedidos,
        'clientesMaisPedidos': clientes_mais_pedidos,
        'evolucaoSemanalMensal': [60, 80, 40, 90, 70] # Dados mockados mantidos
    }), 200


# Em app/relatorios/routes.py

@relatorios_bp.route('/generate-delivery-report/<int:pedido_id>', methods=['GET'])
def generate_delivery_report(pedido_id):
    user_id = request.headers.get('X-User-Id')
    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401

    pedido_encontrado = Pedido.query.filter_by(id=pedido_id, user_id=user_id).first()

    if not pedido_encontrado:
        return jsonify({'message': 'Pedido não encontrado ou não autorizado.'}), 404


    
    # Mapear os dados do objeto Pedido para o formato que gerar_relatorio_entrega espera
    contratante_info = {
        'Nome': pedido_encontrado.clienteNome,
        'RG': pedido_encontrado.clienteRG,
        'CPF': pedido_encontrado.clienteCPF,
    }

    produtos_list = []
    try:
        if pedido_encontrado.produtosContratadosJson:
            produtos_list = json.loads(pedido_encontrado.produtosContratadosJson)
    except json.JSONDecodeError:
        print(f"Erro ao decodificar produtosContratadosJson para o pedido {pedido_id}")

    dados_formatados = {
        'Contratante': contratante_info,
        'Data do Evento': pedido_encontrado.dataEvento,
        'Local do Evento': pedido_encontrado.localEvento,
        'Valor Total do Pedido': pedido_encontrado.valorTotalPedidoContrato,
        'Data de Pagamento': pedido_encontrado.dataPagamentoContrato,
        'Produtos Contratados': produtos_list,
        'Observacoes': pedido_encontrado.observacoes,
        # No caso de futuramente o relatorio precisar de outros campos esse é um dos dicionarios que eu preciso alterar
    }


    report_filename = f"comprovante_retirada_{pedido_id}.docx"
    temp_filepath = os.path.join(current_app.config.get('UPLOAD_FOLDER'), report_filename)
    
    try:
        gerar_relatorio_entrega(dados_formatados, nome_arquivo=temp_filepath)
        
        @after_this_request
        def remove_file(response):
            try: os.remove(temp_filepath)
            except Exception as e: print(f"Erro ao remover arquivo temporário: {e}")
            return response
        
        return send_file(temp_filepath, as_attachment=True, download_name=report_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    except Exception as e:
        return jsonify({'message': f"Erro ao gerar o comprovante: {str(e)}"}), 500
# @relatorios_bp.route('/generate-delivery-report/<int:pedido_id>', methods=['GET'])
# def generate_delivery_report(pedido_id):
#     user_id = request.headers.get('X-User-Id')
#     if not user_id:
#         return jsonify({'message': 'Usuário não autenticado.'}), 401

#     # Substituímos o loop inteiro por uma única consulta ao banco
#     pedido_encontrado = Pedido.query.filter_by(id=pedido_id, user_id=user_id).first()

#     if not pedido_encontrado:
#         return jsonify({'message': 'Pedido não encontrado ou não autorizado.'}), 404

#     # A lógica de mapeamento para o gerador de DOCX continua a mesma,
#     # mas agora acessamos atributos do objeto (ex: pedido_encontrado.clienteNome)
#     # em vez de usar .get() em um dicionário.
#     dados_para_relatorio = pedido_encontrado.to_dict() # Usamos nosso to_dict para facilitar!
    
#     # A função de gerar o docx precisa de um formato específico, então vamos mapear
#     contratante_info = {'Nome': dados_para_relatorio.get('clienteNome'), 'RG': dados_para_relatorio.get('clienteRG'), 'CPF': dados_para_relatorio.get('clienteCPF')}
#     produtos_list = json.loads(dados_para_relatorio.get('produtosContratadosJson', '[]'))
    
#     dados_formatados = {
#         'Contratante': contratante_info,
#         'Produtos Contratados': produtos_list,
#         # Adicione outros campos que a função gerar_relatorio_entrega precise
#     }

#     report_filename = f"comprovante_retirada_{pedido_id}.docx"
#     temp_filepath = os.path.join(current_app.config.get('UPLOAD_FOLDER'), report_filename)
    
#     try:
#         gerar_relatorio_entrega(dados_formatados, nome_arquivo=temp_filepath)
        
#         @after_this_request
#         def remove_file(response):
#             try: os.remove(temp_filepath)
#             except Exception as e: print(f"Erro ao remover arquivo temporário: {e}")
#             return response
        
#         return send_file(temp_filepath, as_attachment=True, download_name=report_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

#     except Exception as e:
#         return jsonify({'message': f"Erro ao gerar o comprovante: {str(e)}"}), 500