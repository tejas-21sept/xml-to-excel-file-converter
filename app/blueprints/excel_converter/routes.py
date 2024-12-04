from app.blueprints.excel_converter import excel_converter_bp
from app.blueprints.excel_converter.views import ConerterAPI

converter_view = ConerterAPI.as_view("converter_api")

excel_converter_bp.add_url_rule("/xml_to_xlsx", view_func=converter_view, methods=["POST"])
