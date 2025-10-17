import os
import json
from flask import Blueprint, request, jsonify, current_app
# from app.__init__ import load_data # Usaremos as funções auxiliares existentes


calculadora_bp = Blueprint('calculadora', __name__, url_prefix='/api/calculadora')
RECIPES_FILE = 'recipes.json'

# Função auxiliar para carregar e salvar receitas
def load_recipes():
    pass
    return jsonify({'message': 'em manuntenção'}),503
    # filepath = os.path.join(current_app.root_path, RECIPES_FILE)
    # if not os.path.exists(filepath):
    #     # Cria o arquivo com receitas de exemplo se ele não existir
    #     with open(filepath, 'w', encoding='utf-8') as f:
    #         json.dump([], f, indent=4, ensure_ascii=False)
    #     return []
    # with open(filepath, 'r', encoding='utf-8') as f:
    #     return json.load(f)

def save_recipes(recipes_list):
    pass
    return jsonify({'message': 'em manuntenção'}),503
    # filepath = os.path.join(current_app.root_path, RECIPES_FILE)
    # with open(filepath, 'w', encoding='utf-8') as f:
    #     json.dump(recipes_list, f, indent=4, ensure_ascii=False)


@calculadora_bp.route('/receitas', methods=['GET'])
def get_receitas():
    pass
    return jsonify({'message': 'em manuntenção'}),503
    # """Retorna a lista de todas as receitas cadastradas."""
    # recipes = load_recipes()
    # return jsonify(recipes), 200

@calculadora_bp.route('/receitas', methods=['POST'])
def add_receita():
    pass
    return jsonify({'message': 'em manuntenção'}),503
    # """Adiciona uma nova receita à lista."""
    # data = request.json
    # required_fields = ['name', 'sellingPrice', 'ingredients']
    # if not all(field in data for field in required_fields):
    #     return jsonify({'message': 'Nome, preço de venda e ingredientes são obrigatórios.'}), 400
    
    # recipes_list = load_recipes()
    # new_recipe_id = f"recipe-{len(recipes_list) + 1}"
    # data['id'] = new_recipe_id
    
    # recipes_list.append(data)
    # save_recipes(recipes_list)
    
    # return jsonify({'message': 'Receita cadastrada com sucesso!', 'recipeId': new_recipe_id}), 201