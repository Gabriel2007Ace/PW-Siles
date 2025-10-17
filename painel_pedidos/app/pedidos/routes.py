# Arquivo: app/pedidos/routes.py (Versão Refatorada com SQLAlchemy)

from flask import Blueprint, request, jsonify
from datetime import datetime
from app import db                  # Importamos a instância do banco de dados
from app.models import Pedido, User # Importamos nossos modelos

pedidos_bp = Blueprint('pedidos', __name__, url_prefix='/api/pedidos')

@pedidos_bp.route('/cadastro', methods=['POST'])
def create_pedido_cadastro():
    data = request.json
    user_id = request.headers.get('X-User-Id')

    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401
    
    # Validação dos campos continua igual...
    required_fields = ['clienteNome', 'dataEvento', 'quantidade', 'tipoPedido', 'dataRetirada', 'horarioRetirada']
    if not all(field in data and data[field] for field in required_fields):
        return jsonify({'message': 'Campos obrigatórios faltando.'}), 400

    # Em vez de criar um dicionário, criamos uma instância do nosso modelo Pedido
    new_pedido = Pedido(
        clienteNome=data['clienteNome'],
        dataEvento=data['dataEvento'],
        dataRetirada=data['dataRetirada'],
        horarioRetirada=data['horarioRetirada'],
        tipoPedido=data['tipoPedido'],
        quantidade=int(data['quantidade']),
        sabores=data.get('sabores', ''),
        tipoEmbalagem=data.get('tipoEmbalagem', ''),
        observacoes=data.get('observacoes', ''),
        status='pendente',
        prioridade=data.get('prioridade', 'normal'),
        responsavel=data.get('responsavel', None),
        user_id=user_id, # Ligamos o pedido ao usuário
        # Campos do contrato
        clienteRG=data.get('clienteRG', ''),
        clienteCPF=data.get('clienteCPF', ''),
        nomeContratado=data.get('nomeContratado', ''),
        cnpjContratado=data.get('cnpjContratado', ''),
        valorTotalPedidoContrato=data.get('valorTotalPedidoContrato', ''),
        dataPagamentoContrato=data.get('dataPagamentoContrato', ''),
        localEvento=data.get('localEvento', ''),
        produtosContratadosJson=data.get('produtosContratadosJson', '[]')
    )

    # Adicionamos à sessão e salvamos no banco de dados
    db.session.add(new_pedido)
    db.session.commit()

    return jsonify({'message': 'Pedido salvo com sucesso!', 'pedido': new_pedido.to_dict()}), 201



@pedidos_bp.route('', methods=['POST'])
def create_pedido():
    """Essa função lida com o pedido que vem do upload contrato"""
    data = request.json
    user_id = request.headers.get('X-User-Id')

    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401
    
    # Validação dos campos continua igual...
    required_fields = ['clienteNome', 'dataEvento', 'quantidade', 'tipoPedido', 'dataRetirada', 'horarioRetirada']
    if not all(field in data and data[field] for field in required_fields):
        return jsonify({'message': 'Campos obrigatórios faltando.'}), 400

    # ALTERAÇÃO QUE SALVA TODAS AS "DATAEVENTO" NO FORMATO YYYY-MM-DD

    print("antes da formatacao", type(data["dataEvento"]))
    
    # dataEvento = datetime.strptime(data['dataEvento'], '%d-%m-%Y').date()
    dataEvento = data["dataEvento"]
    # dataEvento = slice(dataEvento)
    # print(dataEvento)
    dias = dataEvento[0:2]
    mes = dataEvento[3:5]
    ano = dataEvento[6:10]
    dataEvento = ano + "-" + mes +"-"+dias
    data["dataEvento"] = dataEvento
    # print(data["dataEvento"])




    # Em vez de criar um dicionário, criamos uma instância do nosso modelo Pedido
    new_pedido = Pedido(
        clienteNome=data['clienteNome'],
        dataEvento=data['dataEvento'],
        dataRetirada=data['dataRetirada'],
        horarioRetirada=data['horarioRetirada'],
        tipoPedido=data['tipoPedido'],
        quantidade=int(data['quantidade']),
        sabores=data.get('sabores', ''),
        tipoEmbalagem=data.get('tipoEmbalagem', ''),
        observacoes=data.get('observacoes', ''),
        status='pendente',
        prioridade=data.get('prioridade', 'normal'),
        responsavel=data.get('responsavel', None),
        user_id=user_id, # Ligamos o pedido ao usuário
        # Campos do contrato
        clienteRG=data.get('clienteRG', ''),
        clienteCPF=data.get('clienteCPF', ''),
        nomeContratado=data.get('nomeContratado', ''),
        cnpjContratado=data.get('cnpjContratado', ''),
        valorTotalPedidoContrato=data.get('valorTotalPedidoContrato', ''),
        dataPagamentoContrato=data.get('dataPagamentoContrato', ''),
        localEvento=data.get('localEvento', ''),
        produtosContratadosJson=data.get('produtosContratadosJson', '[]')
    )

    # Adicionamos à sessão e salvamos no banco de dados
    db.session.add(new_pedido)
    db.session.commit()

    return jsonify({'message': 'Pedido salvo com sucesso!', 'pedido': new_pedido.to_dict()}), 201

@pedidos_bp.route('', methods=['GET'])
def get_pedidos():
    user_id = request.headers.get('X-User-Id')
    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401

    # Construímos a consulta ao banco de dados passo a passo
    query = Pedido.query.filter_by(user_id=user_id)

    # Aplicamos os filtros vindos da URL
    filtro_cliente = request.args.get('cliente', '').lower()
    if filtro_cliente:
        query = query.filter(Pedido.clienteNome.ilike(f'%{filtro_cliente}%'))

    filtro_data_evento = request.args.get('dataEvento', '')
    if filtro_data_evento:
        query = query.filter_by(dataEvento=filtro_data_evento)

    filtro_status = request.args.get('status', '')
    if filtro_status:
        if ',' in filtro_status:
            status_list = filtro_status.split(',')
            query = query.filter(Pedido.status.in_(status_list))
        else:
            query = query.filter_by(status=filtro_status)
    
    # Executamos a consulta final no banco de dados
    pedidos = query.order_by(Pedido.createdAt.desc()).all()
    
    # Convertemos a lista de objetos Pedido para uma lista de dicionários
    pedidos_dict = [pedido.to_dict() for pedido in pedidos]

    return jsonify(pedidos_dict), 200

@pedidos_bp.route('/<int:pedido_id>', methods=['GET'])
def get_pedido_details(pedido_id):
    user_id = request.headers.get('X-User-Id')
    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401

    # Buscamos o pedido específico pelo ID e ID do usuário
    pedido = Pedido.query.filter_by(id=pedido_id, user_id=user_id).first()

    if pedido:
        return jsonify(pedido.to_dict()), 200
    return jsonify({'message': 'Pedido não encontrado ou não autorizado.'}), 404

@pedidos_bp.route('/<int:pedido_id>', methods=['PUT'])
def update_pedido(pedido_id):
    user_id = request.headers.get('X-User-Id')
    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401

    pedido = Pedido.query.filter_by(id=pedido_id, user_id=user_id).first()
    if not pedido:
        return jsonify({'message': 'Pedido não encontrado ou não autorizado.'}), 404

    data = request.json
    
    # Atualizamos os campos do objeto pedido diretamente
    for key, value in data.items():
        if hasattr(pedido, key) and key not in ['id', 'userId', 'createdAt']:
            setattr(pedido, key, value)
    
    db.session.commit() # Salvamos as alterações no banco
    return jsonify({'message': 'Pedido atualizado com sucesso!', 'pedido': pedido.to_dict()}), 200

@pedidos_bp.route('/<int:pedido_id>', methods=['DELETE'])
def delete_pedido(pedido_id):
    user_id = request.headers.get('X-User-Id')
    if not user_id:
        return jsonify({'message': 'Usuário não autenticado.'}), 401

    pedido = Pedido.query.filter_by(id=pedido_id, user_id=user_id).first()
    if not pedido:
        return jsonify({'message': 'Pedido não encontrado ou não autorizado.'}), 404
        
    db.session.delete(pedido) # Marcamos o objeto para exclusão
    db.session.commit() # Confirmamos a exclusão no banco
    return jsonify({'message': 'Pedido excluído com sucesso!'}), 200