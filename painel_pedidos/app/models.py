

from app import db # Importa a instância 'db' do __init__.py
from datetime import datetime

# O 'db.Model' é a classe base para todos os modelos do Flask-SQLAlchemy
class User(db.Model):
    """Modelo para a tabela de Usuários."""
    id = db.Column(db.Integer, primary_key=True)
    nome_completo = db.Column(db.String(150), nullable=False) # ALTERAÇÕES FINDAS COM O FRONT-END 
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False) # Aumentado para hashes mais longos
    tipo_usuario = db.Column(db.String(50), nullable=False, default='funcionario') # <<< NOVO CAMPO

    
    # Relacionamento: Um usuário pode ter vários pedidos.
    # 'backref' cria um atributo 'user' no modelo Pedido para acessar o usuário correspondente.
    pedidos = db.relationship('Pedido', backref='user', lazy=True)

    def __repr__(self):
        return f"User('{self.email}{self.nome_completo}')"

class Pedido(db.Model):
    """Modelo para a tabela de Pedidos."""
    id = db.Column(db.Integer, primary_key=True)
    clienteNome = db.Column(db.String(150), nullable=False)
    dataEvento = db.Column(db.String(10), nullable=False) # Armazenado como YYYY-MM-DD
    dataRetirada = db.Column(db.String(10), nullable=False)
    horarioRetirada = db.Column(db.String(5), nullable=False)
    tipoPedido = db.Column(db.String(100), nullable=False)
    quantidade = db.Column(db.Integer, nullable=False)
    sabores = db.Column(db.Text, nullable=True)
    tipoEmbalagem = db.Column(db.String(100), nullable=True)
    observacoes = db.Column(db.Text, nullable=True)
    status = db.Column(db.String(50), nullable=False, default='pendente')
    prioridade = db.Column(db.String(50), nullable=False, default='normal')
    responsavel = db.Column(db.String(150), nullable=True) # alteração de responsavel pelo pedido que Damaris pediu
    createdAt = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    
    # --- CAMPOS VINDOS DO CONTRATO ---
    clienteRG = db.Column(db.String(20), nullable=True)
    clienteCPF = db.Column(db.String(20), nullable=True)
    nomeContratado = db.Column(db.String(150), nullable=True)
    cnpjContratado = db.Column(db.String(20), nullable=True)
    valorTotalPedidoContrato = db.Column(db.String(50), nullable=True)
    dataPagamentoContrato = db.Column(db.String(10), nullable=True)
    localEvento = db.Column(db.String(200), nullable=True)
    produtosContratadosJson = db.Column(db.Text, nullable=True) # Armazena a lista de produtos como JSON

    # --- Chave Estrangeira ---
    # Liga este pedido a um usuário específico. 'user.id' refere-se à tabela 'user' e coluna 'id'.
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)

    def __repr__(self):
        return f"Pedido(ID: {self.id}, Cliente: '{self.clienteNome}', Status: '{self.status}')"
    
    def to_dict(self):
        """Converte a instância do Pedido em um dicionário serializável."""
        return {
            'id': self.id,
            'clienteNome': self.clienteNome,
            'dataEvento': self.dataEvento,
            'dataRetirada': self.dataRetirada,
            'horarioRetirada': self.horarioRetirada,
            'tipoPedido': self.tipoPedido,
            'quantidade': self.quantidade,
            'sabores': self.sabores,
            'tipoEmbalagem': self.tipoEmbalagem,
            'observacoes': self.observacoes,
            'status': self.status,
            'prioridade': self.prioridade,
            'responsavel': self.responsavel, #alteração de responsavel pelo pedido que a Damaris pediu
            'createdAt': self.createdAt.isoformat(),
            'userId': self.user_id,
            'clienteRG': self.clienteRG,
            'clienteCPF': self.clienteCPF,
            'nomeContratado': self.nomeContratado,
            'cnpjContratado': self.cnpjContratado,
            'valorTotalPedidoContrato': self.valorTotalPedidoContrato,
            'dataPagamentoContrato': self.dataPagamentoContrato,
            'localEvento': self.localEvento,
            'produtosContratadosJson': self.produtosContratadosJson
        }