from flask import Blueprint, request, jsonify
from app import db                  # Importamos a instância do banco de dados
from app.models import User         # Importamos nosso modelo User
from werkzeug.security import generate_password_hash, check_password_hash # Ferramentas para hashing

auth_bp = Blueprint('auth', __name__, url_prefix='/api/auth')

# Em app/auth/routes.py

@auth_bp.route('/register', methods=['POST'])
def register_user():
    data = request.json
    email = data.get('email')
    password = data.get('password')
    nome_completo = data.get('nome_completo') # <<< ALTERACAO FINDA DO INTEGRACAO DO FRONT-END
    tipo_usuario = data.get('tipo_usuario', 'funcionario') # <<< ALTERACAO FINDA DO INTEGRACAO DO FRONT-END

    # A validação agora inclui o nome completo
    if not email or not password or not nome_completo:
        return jsonify({'message': 'Nome, e-mail e senha são obrigatórios.'}), 400

    if User.query.filter_by(email=email).first():
        return jsonify({'message': 'E-mail já cadastrado.'}), 409

    hashed_password = generate_password_hash(password, method='pbkdf2:sha256')

    # Criamos o novo usuário com todos os campos
    new_user = User(
        email=email, 
        nome_completo=nome_completo, 
        tipo_usuario=tipo_usuario, 
        password_hash=hashed_password
    )

    db.session.add(new_user)
    db.session.commit()

    return jsonify({'message': 'Usuário cadastrado com sucesso!'}), 201

@auth_bp.route('/login', methods=['POST'])
def login_user():
    data = request.json
    email = data.get('email')
    password = data.get('password')

    if not email or not password:
        return jsonify({'message': 'E-mail e senha são obrigatórios.'}), 400

    # Busca o usuário pelo e-mail no banco de dados
    user = User.query.filter_by(email=email).first()

    # Verifica se o usuário existe E se a senha fornecida corresponde ao hash guardado
    if user and check_password_hash(user.password_hash, password):
        # O login é bem-sucedido!
        # Agora retornamos o ID numérico do usuário, que é mais seguro e padrão.
        return jsonify({'message': 'Login bem-sucedido!', 'userId': user.id}), 200
    else:
        # Se o usuário não existe ou a senha está errada, a mensagem é a mesma por segurança.
        return jsonify({'message': 'E-mail ou senha incorretos.'}), 401