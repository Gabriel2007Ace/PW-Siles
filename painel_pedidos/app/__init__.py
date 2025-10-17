# Arquivo: app/__init__.py (Versão final e mesclada)

import os
from flask import Flask, jsonify
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy # Importamos SQLAlchemy
from dotenv import load_dotenv

# 1. CRIAMOS A INSTÂNCIA DO BANCO DE DADOS AQUI FORA
# Esta variável 'db' será importada por outros arquivos, como o models.py
load_dotenv()

db = SQLAlchemy()

def create_app():
    """Função Application Factory: configura e retorna a instância da aplicação Flask."""
    app = Flask(__name__, instance_relative_config=True)
    CORS(app)

    # Carregar configurações do arquivo config.py
    app.config.from_object('app.config.Config')

    # --- DEBUG: Esta linha do esqueleto é útil para confirmar o endereço do banco de dados ---
    print("*" * 80)
    print(f"INFO: Conectando ao banco de dados: {app.config['SQLALCHEMY_DATABASE_URI']}")
    print("*" * 80)

    # 2. INICIALIZAMOS O BANCO DE DADOS COM A NOSSA APLICAÇÃO
    # Esta linha conecta o SQLAlchemy ao Flask
    db.init_app(app)

    # Cria a pasta de upload (lógica mantida do seu projeto padrão)
    UPLOAD_FOLDER = os.path.join(app.root_path, 'uploads_temp')
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

    # --- Registro de Blueprints (lógica mantida do seu projeto padrão) ---
    from app.main_pages.routes import main_pages_bp
    from app.auth.routes import auth_bp
    from app.pedidos.routes import pedidos_bp
    from app.contratos.routes import contratos_bp
    from app.relatorios.routes import relatorios_bp
    from app.calculadora.routes import calculadora_bp

    app.register_blueprint(main_pages_bp)
    app.register_blueprint(auth_bp)
    app.register_blueprint(pedidos_bp)
    app.register_blueprint(contratos_bp)
    app.register_blueprint(relatorios_bp)
    app.register_blueprint(calculadora_bp)

    # 3. CRIAMOS AS TABELAS NO BANCO DE DADOS
    # Este bloco de código lê seus models.py e cria as tabelas no arquivo site.db
    with app.app_context():
        db.create_all()
        print("Banco de dados inicializado e tabelas criadas (se necessário).")

    # Rota de status (mantida)
    @app.route('/status', methods=['GET'])
    def status_check():
        return jsonify({'status': 'online', 'message': 'Backend online e pronto!'}), 200

    return app
