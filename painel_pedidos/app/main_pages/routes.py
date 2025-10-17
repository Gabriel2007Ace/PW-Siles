from flask import Blueprint, render_template

# Este Blueprint servirá todas as suas páginas HTML.
# Não terá um prefixo de URL para que as rotas sejam simples, ex: /login, /pedidos.
main_pages_bp = Blueprint('main_pages', __name__)

@main_pages_bp.route('/')
@main_pages_bp.route('/login')
def login_page():
    """Rota para a página de login."""
    return render_template('login.html')

@main_pages_bp.route('/pedidos')
def list_pedidos_page():
    """Rota para a página de listagem de pedidos."""
    return render_template('pedidos.html')

@main_pages_bp.route('/pedidos/novo')
def new_pedido_page():
    """Rota para a página de criação de um novo pedido."""
    return render_template('novo_pedido.html')

@main_pages_bp.route('/pedidos/<pedido_id>')
def details_pedido_page(pedido_id):
    """Rota para a página de detalhes de um pedido específico."""
    # O 'pedido_id' pode ser usado pelo JavaScript na página para buscar os dados.
    return render_template('pedido_detalhes.html')

@main_pages_bp.route('/contratos')
def upload_contratos_page():
    """Rota para a página de upload de contratos."""
    return render_template('contratos.html')

@main_pages_bp.route('/exportar')
def exportar_planilha_page():
    """Rota para a página de exportação de planilhas."""
    return render_template('exportar.html')

@main_pages_bp.route('/relatorios')
def relatorios_page():
    """Rota para a página de relatórios e indicadores."""
    return render_template('relatorios.html')

@main_pages_bp.route('/contratos/novo')
def novo_contrato_page():
    """Rota para a página de criação de um novo contrato."""
    return render_template('novo_contrato.html')


@main_pages_bp.route('/calculadora')
def calculadora_custos_page():
    """Rota para a página da calculadora de custos."""
    return render_template('calculadora.html')

@main_pages_bp.route('/equipe')
def pagina_equipe():
    """ Rota para página da equipe. """
    return render_template("equipe.html")