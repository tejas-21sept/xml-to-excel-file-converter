from flask import Flask
from app.blueprints.excel_converter import excel_converter_bp

def create_app(config_class='config.Config'):
    app = Flask(__name__)
    app.config.from_object(config_class)

    # Register Blueprints
    app.register_blueprint(excel_converter_bp, url_prefix='/excel_converter')

    return app
