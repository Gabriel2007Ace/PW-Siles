import os

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'uma_chave_secreta_padrao_para_desenvolvimento'
    # Configurações do Banco de Dados (exemplo para SQLite, você usará seu DB relacional)
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or 'sqlite:///site.db'
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    # Outras configurações globais podem vir aqui, ex:
    # WHATSAPP_API_KEY = os.environ.get('WHATSAPP_API_KEY')
    # PDF_EXTRACTION_MODEL = 'v1'