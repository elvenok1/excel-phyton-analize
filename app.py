# -*- coding: utf-8 -*-

import io
from flask import Flask, request, jsonify, send_file

# Importaciones de openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.colors import Color
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule, DataBarRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.cell import MergedCell

# --- Inicialización de la Aplicación Flask ---
app = Flask(__name__)

# --- FUNCIONES DE AYUDA PARA CREAR EXCEL ---
def apply_styles_to_cell(cell, style_data):
    if not style_data or not isinstance(style_data, dict): return
    if 'font' in style_data: cell.font = Font(**style_data['font'])
    if 'fill' in style_data:
        if 'pattern' in style_data['fill']: style_data['fill']['fill_type'] = style_data['fill'].pop('pattern')
        cell.fill = PatternFill(**style_data['fill'])
    if 'border' in style_data:
        cell.border = Border(left=Side(**style_data['border'].get('left', {})), right=Side(**style_data['border'].get('right', {})), top=Side(**style_data['border'].get('top', {})), bottom=Side(**style_data['border'].get('bottom', {})))
    if 'alignment' in style_data: cell.alignment = Alignment(**style_data['alignment'])
    if 'numFmt' in style_data: cell.number_format = style_data['numFmt']

def create_chart_from_spec(worksheet, chart_spec):
    # ... (código sin cambios)
    pass

# === FUNCIÓN DE ANÁLISIS CORREGIDA PARA EL ERROR "RGB IS NOT JSON SERIALIZABLE" ===
def extract_styles_from_cell(cell):
    style_data = {}
    if not cell.has_style:
        return style_data

    # Función auxiliar para convertir colores a string de forma segura
    def get_color_str(color_obj):
        if color_obj and isinstance(color_obj, Color) and color_obj.rgb:
            return str(color_obj.rgb)
        return None

    # Fuente (Font)
    font_data = {}
    if cell.font:
        if cell.font.name: font_data['name'] = cell.font.name
        if cell.font.sz: font_data['sz'] = cell.font.sz
        if cell.font.bold: font_data['bold'] = cell.font.bold
        if cell.font.italic: font_data['italic'] = cell.font.italic
        color_val = get_color_str(cell.font.color)
        if color_val: font_data['color'] = color_val
    if font_data: style_data['font'] = font_data

    # Relleno (Fill)
    fill_data = {}
    if cell.fill and cell.fill.fill_type:
        fill_data['pattern'] = cell.fill.fill_type
        start_color = get_color_str(cell.fill.start_color)
        end_color = get_color_str(cell.fill.end_color)
        if start_color: fill_data['start_color'] = start_color
        if end_color: fill_data['end_color'] = end_color
    if fill_data: style_data['fill'] = fill_data
        
    # Bordes (Border)
    border_data = {}
    if cell.border:
        def get_side_style(side):
            if side and side.style:
                side_color = get_color_str(side.color)
                return {'style': side.style, 'color': side_color}
            return None
        
        left, right, top, bottom = get_side_style(cell.border.left), get_side_style(cell.border.right), get_side_style(cell.border.top), get_side_style(cell.border.bottom)
        if left: border_data['left'] = left
        if right: border_data['right'] = right
        if top: border_data['top'] = top
        if bottom: border_data['bottom'] = bottom
    if border_data: style_data['border'] = border_data
        
    # Alineación (Alignment)
    alignment_data = {}
    if cell.alignment:
        if cell.alignment.horizontal: alignment_data['horizontal'] = cell.alignment.horizontal
        if cell.alignment.vertical: alignment_data['vertical'] = cell.alignment.vertical
        if cell.alignment.wrap_text: alignment_data['wrap_text'] = cell.alignment.wrap_text
    if alignment_data: style_data['alignment'] = alignment_data
        
    if cell.number_format and cell.number_format != 'General': style_data['numFmt'] = cell.number_format
    return style_data

# --- ENDPOINT DE CHEQUEO DE SALUD ---
@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({"status": "ok"}), 200

# --- ENDPOINT PARA CREAR EXCEL ---
@app.route('/create-excel', methods=['POST'])
def create_excel():
    # ... (código sin cambios)
    try:
        json_data = request.get_json()
        if not json_data: return jsonify({"error": "No JSON received"}), 400
        # ...
        return send_file(io.BytesIO(), as_attachment=True, download_name='report.xlsx')
    except Exception as e:
        print(f"Error in /create-excel: {e}")
        return jsonify({"error": str(e)}), 500

# --- ENDPOINT PARA ANALIZAR EXCEL ---
@app.route('/parse-excel', methods=['POST'])
def parse_excel():
    if 'excel_file' not in request.files:
        return jsonify({"error": "No se encontró el archivo en la petición (se esperaba el campo 'excel_file')."}), 400
    file = request.files['excel_file']
    if file.filename == '': return jsonify({"error": "No se seleccionó ningún archivo."}), 400
    try:
        in_memory_file = io.BytesIO(file.read())
        wb = load_workbook(filename=in_memory_file, data_only=True)
        parsed_data = {'sheets': []}
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_data = { 'name': sheet_name, 'data': [], 'merged_cells': [str(r) for r in ws.merged_cells.ranges] }
            rows_data = []
            for row in ws.iter_rows():
                row_list = []
                for cell in row:
                    cell_info = {'address': cell.coordinate, 'value': cell.value}
                    if isinstance(cell, MergedCell):
                        cell_info['is_merged_part'] = True
                    else:
                        # Se utiliza la nueva función corregida
                        cell_info['style'] = extract_styles_from_cell(cell)
                    row_list.append(cell_info)
                rows_data.append(row_list)
            sheet_data['data'] = rows_data
            parsed_data['sheets'].append(sheet_data)
        return jsonify(parsed_data)
    except Exception as e:
        print(f"Error en /parse-excel: {e}")
        return jsonify({"error": f"Error interno al procesar el archivo Excel: {str(e)}"}), 500

# --- Punto de Entrada de la Aplicación ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

