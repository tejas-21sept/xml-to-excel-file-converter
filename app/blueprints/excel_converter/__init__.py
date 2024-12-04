from flask import Blueprint

excel_converter_bp = Blueprint("excel_converter", __name__)

from . import routes
