from flask import Flask, render_template, request
import os


from flask import Flask, render_template, request, redirect
import os
import base64
from io import BytesIO
from PIL import Image
from docx import Document
from docx.shared import Inches

app = Flask(__name__)

UPLOAD_FOLDER = 'firmas_temp'
CURSOS_FILE = 'cursos.txt'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def get_cursos():
    if not os.path.exists(CURSOS_FILE):
        return []
    with open(CURSOS_FILE, 'r', encoding='utf-8') as f:
        return [line.strip() for line in f if line.strip()]


def find_table_with_headers(doc):
    """Buscar una tabla que contenga encabezados relacionados con Nombre, Codigo y Firma.
    Devuelve (table, name_idx, code_idx, firma_idx) o (None, None, None, None).
    """
    for table in doc.tables:
        # leer textos de la primera fila
        if len(table.rows) == 0:
            continue
        headers = [cell.text.strip().upper().replace(' ', '') for cell in table.rows[0].cells]
        # buscar coincidencias parciales
        has_name = any('NOMBRE' in h or 'NOMBRECOMPLETO' in h for h in headers)
        has_code = any('CODIG' in h for h in headers)
        has_firma = any('FIRMA' in h for h in headers)
        if has_name and has_code and has_firma:
            # obtener índices
            name_idx = next((i for i,h in enumerate(headers) if 'NOMBRE' in h or 'NOMBRECOMPLETO' in h), 0)
            code_idx = next((i for i,h in enumerate(headers) if 'CODIG' in h), 1)
            firma_idx = next((i for i,h in enumerate(headers) if 'FIRMA' in h), 2)
            return table, name_idx, code_idx, firma_idx
    return None, None, None, None

@app.route('/')
def index():
    cursos = get_cursos()
    return render_template('form.html', cursos=cursos)

@app.route('/crear_curso', methods=['GET', 'POST'])
def crear_curso():
    if request.method == 'POST':
        nombre_curso = request.form['nombre_curso'].strip()
        if nombre_curso:
            with open(CURSOS_FILE, 'a', encoding='utf-8') as f:
                f.write(nombre_curso + '\n')
        return redirect('/crear_curso')
    cursos = get_cursos()
    return render_template('crear_curso.html', cursos=cursos)


@app.route('/submit', methods=['POST'])
def submit():
    curso = request.form['curso']
    nombre = request.form['nombre']
    codigo = request.form['codigo']
    firma_data = request.form['firma']

    # Documento Word por curso con formato F6_Acta_{curso}.docx
    safe_curso = curso.replace(' ', '_')
    word_path = f"F6_Acta_{safe_curso}.docx"

    # Cargar documento o crearlo si no existe
    if os.path.exists(word_path):
        doc = Document(word_path)
    else:
        doc = Document()
        doc.add_heading(f"Acta F6 - {curso}", 0)

    table, name_idx, code_idx, firma_idx = find_table_with_headers(doc)

    # Si no existe la tabla con los encabezados, crearla al final
    if table is None:
        table = doc.add_table(rows=1, cols=3)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'NOMBRE COMPLETO'
        hdr_cells[1].text = 'CODIGO'
        hdr_cells[2].text = 'FIRMA'
        name_idx, code_idx, firma_idx = 0, 1, 2

    # Verificar duplicado en la columna de código
    codigo_duplicado = False
    for row in table.rows[1:]:
        # proteger si la fila tiene menos celdas
        if len(row.cells) > code_idx:
            if row.cells[code_idx].text.strip() == codigo:
                codigo_duplicado = True
                break

    if codigo_duplicado:
        cursos = get_cursos()
        error = f"El código {codigo} ya está registrado en el curso {curso}."
        return render_template('form.html', cursos=cursos, error=error)

    # Procesar imagen de la firma
    header, encoded = firma_data.split(',', 1)
    img_bytes = base64.b64decode(encoded)
    img = Image.open(BytesIO(img_bytes))
    img_path = os.path.join(UPLOAD_FOLDER, f"{codigo}_firma.png")
    img.save(img_path)

    # Agregar nueva fila usando los índices de columnas encontrados
    new_cells = table.add_row().cells
    # Asegurar que la fila tenga al menos 3 celdas
    while len(new_cells) < 3:
        new_cells.append(new_cells[-1])
    new_cells[name_idx].text = nombre
    new_cells[code_idx].text = codigo
    # Insertar imagen en la celda de firma
    run = new_cells[firma_idx].paragraphs[0].add_run()
    run.add_picture(img_path, width=Inches(1.5))
    doc.save(word_path)

    return render_template('gracias.html')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5006)