import os
import io
import re
import zipfile
import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image
from flask import Flask, request, render_template, session, jsonify, send_file

# --- CONFIGURACIÓN ---
app = Flask(__name__)
app.secret_key = 'una-clave-secreta-muy-robusta' # Cambia esto en producción

# ¡IMPORTANTE! Si estás en Windows, probablemente necesites descomentar y ajustar la siguiente línea:
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# --- FUNCIONES DE PROCESAMIENTO MEJORADAS ---

def extraer_texto_con_ocr(stream_pdf):
    """
    Extrae texto de un PDF. Primero intenta la extracción directa.
    Si el texto es mínimo, aplica OCR asumiendo que es un PDF escaneado.
    Retorna una tupla: (texto_extraido, None) en éxito, o (None, error_message) en fallo.
    """
    texto_completo = ""
    try:
        doc = fitz.open(stream=stream_pdf, filetype="pdf")
        
        # Iterar por cada página para decidir si usar OCR
        for num_pagina, pagina in enumerate(doc):
            # 1. Intento de extracción directa
            texto_directo = pagina.get_text().strip()
            
            # 2. Heurística: Si el texto directo es muy corto, probablemente es una imagen.
            if len(texto_directo) < 100: # Ajusta este umbral según tus documentos
                try:
                    # Convertir página a imagen para OCR
                    pix = pagina.get_pixmap(dpi=300) # Mayor DPI para mejor calidad de OCR
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data))
                    
                    # 3. Aplicar OCR en español
                    texto_ocr = pytesseract.image_to_string(img, lang='spa') # 'spa' para español
                    texto_completo += texto_ocr + "\n"
                except Exception as ocr_error:
                    # Si el OCR falla en una página, lo registramos y continuamos
                    print(f"Advertencia: OCR falló en la página {num_pagina + 1}: {ocr_error}")
                    texto_completo += texto_directo + "\n" # Usamos el poco texto que se pudo sacar
            else:
                texto_completo += texto_directo + "\n"
                
        doc.close()
        return (texto_completo, None)

    except Exception as e:
        error_msg = f"Error Crítico al procesar el PDF: {e}"
        print(error_msg)
        return (None, error_msg)

def analizar_contrato(texto_contrato):
    """Aplica regex para extraer datos de un texto, usando flags importantes."""

    # Usa el nuevo diccionario de patrones
    patrones = {
        'num_contrato': r'(?:CONTRATO|ORDEN DE SERVICIO|ORDEN DE COMPRA)\s*(?:N[°º.\sro]+[:\s]*)?([\w\d\-\/.]+)',
        'fecha_suscripcion_contrato': r'(?:en la ciudad de|firmado el|celebrado a los|Fecha:|a los)\s*.*?(\d{1,2}\s+(?:de\s+)?\w+\s+(?:de\s+)?\d{4})',
        'proceso_seleccion': r'(Adjudicaci[oó]n Simplificada|Licitaci[oó]n P[úu]blica|Concurso P[úu]blico)(?:\s*N[°º.\sro]+[:\s]*([\w\d\-\/]+))?',
        'vigencia_final': r'vigencia.*?\s+hasta\s+(.*?)(?:\.|\n|CL[ÁA]USULA)',
        'monto_contratado_total': r'(?:Monto|Valor Total|asciende a)\s*.*?(?:S\/\.?|R\$|USD)?\s*([\d,]+(?:\.\d{2})?)',
        'ruc_contratista': r'CONTRATISTA[\s\S]*?RUC\s*N[°º.\sro]*\s*(\d{11})',
        'ruc_entidad': r'(?:LA\s+ENTIDAD|CONTRATANTE)[\s\S]*?RUC\s*N[°º.\sro]*\s*(\d{11})',
        'resolucion': r'RESOLUCI[OÓ]N\s*(?:DECANAL|RECTORAL|DIRECTORAL)?\s*N[°º.\sro]+[:\s]*([\w\d\-\/.-]+)',
        'num_item': r'(?:Ítem|Item)\s*(?:N[°º.\sro]+)?\s*[:]?\s*(\d+)',
        
        # Patrones que NECESITAN re.DOTALL para ser efectivos
        'objeto_contrato': r'CL[ÁA]USULA\s+PRIMERA\s*:\s*OBJETO[\s\S]*?([\s\S]*?)(?=CL[ÁA]USULA\s+SEGUNDA|II\.\s+BASE\s+LEGAL)',
        'plazo_ejecucion': r'PLAZO\s+DE\s+(?:EJECUCI[OÓ]N|ENTREGA)[\s\S]*?([\s\S]*?)(?=CL[ÁA]USULA|INICIO\s+DEL\s+PLAZO)',
        'representante_legal_contratista': r'representad[oa]\s+(?:legalmente\s+)?por\s+([^,]+(?:,\s*Jr\.|,\s*S\.A\.C\.)?),' # Captura nombres con comas como "Juan Perez, Jr."
    }
    
    datos_extraidos = {}
    for clave, patron in patrones.items():
        # ¡IMPORTANTE! Usamos re.DOTALL para que '.' incluya saltos de línea
        # y re.IGNORECASE para que no distinga entre mayúsculas y minúsculas.
        match = re.search(patron, texto_contrato, re.IGNORECASE | re.DOTALL)
        
        if match:
            # Si hay múltiples grupos de captura, unimos los que no estén vacíos.
            # Esto es útil para 'proceso_seleccion'.
            grupos_validos = [g for g in match.groups() if g is not None]
            resultado = ' '.join(grupos_validos)
            
            # Limpiamos el resultado de espacios extra y saltos de línea
            resultado_limpio = ' '.join(resultado.split()).strip()
            datos_extraidos[clave] = resultado_limpio
        else:
            datos_extraidos[clave] = "No encontrado"
            
    return datos_extraidos

# --- RUTAS DE LA APLICACIÓN WEB ---

@app.route('/', methods=['GET'])
def index():
    """Sirve la página principal estática."""
    return render_template('index.html')

# En app.py

# ... (otras funciones no cambian) ...

@app.route('/process', methods=['POST'])
def process_files():
    if 'upload_files' not in request.files: # Usamos el nuevo nombre del input
        return jsonify({'success': False, 'errors': ['No se enviaron archivos.']}), 400

    files = request.files.getlist('upload_files') # Usamos el nuevo nombre del input
    
    resultados_ok = []
    errores_proceso = []

    for file in files:
        nombre_archivo = file.filename
        
        # --- INICIO DE LA LÓGICA PARA ZIP ---
        if nombre_archivo.endswith('.zip'):
            try:
                # Abrir el ZIP en memoria sin guardarlo en disco
                zip_stream = io.BytesIO(file.read())
                if not zipfile.is_zipfile(zip_stream):
                    errores_proceso.append(f"'{nombre_archivo}' no es un archivo ZIP válido.")
                    continue

                with zipfile.ZipFile(zip_stream) as zip_ref:
                    # Iterar sobre cada archivo dentro del ZIP
                    for item_name in zip_ref.namelist():
                        # Procesar solo si es un PDF y no una carpeta de macOS (ignora __MACOSX)
                        if item_name.lower().endswith('.pdf') and not item_name.startswith('__MACOSX'):
                            try:
                                # Leer el contenido del PDF desde el ZIP
                                pdf_data_in_zip = zip_ref.read(item_name)
                                texto, error = extraer_texto_con_ocr(pdf_data_in_zip)
                                
                                if error:
                                    errores_proceso.append(f"Error en '{item_name}' (dentro de {nombre_archivo}): {error}")
                                    continue
                                
                                datos = analizar_contrato(texto)
                                datos['Archivo'] = item_name # Usamos el nombre del archivo dentro del ZIP
                                resultados_ok.append(datos)
                            except Exception as e_pdf:
                                errores_proceso.append(f"Error procesando '{item_name}' (dentro de {nombre_archivo}): {e_pdf}")
            except Exception as e_zip:
                errores_proceso.append(f"Error al procesar el archivo ZIP '{nombre_archivo}': {e_zip}")

        # --- FIN DE LA LÓGICA PARA ZIP ---
        
        # Mantenemos la lógica para PDFs individuales
        elif nombre_archivo.endswith('.pdf'):
            try:
                pdf_stream = file.read()
                texto, error = extraer_texto_con_ocr(pdf_stream)
                
                if error:
                    errores_proceso.append(f"'{nombre_archivo}': {error}")
                    continue

                datos = analizar_contrato(texto)
                datos['Archivo'] = nombre_archivo
                resultados_ok.append(datos)

            except Exception as e:
                errores_proceso.append(f"Error inesperado procesando '{nombre_archivo}': {e}")

    # El resto de la función (crear el DataFrame, el JSON de respuesta, etc.) no cambia
    # ...
    if not resultados_ok:
        return jsonify({'success': False, 'errors': ['No se pudo extraer información de ningún archivo.'] + errores_proceso})

    df = pd.DataFrame(resultados_ok)
    columnas = ['Archivo'] + [col for col in df.columns if col != 'Archivo']
    df = df[columnas]
    
    output = io.BytesIO()
    df.to_excel(output, index=False, sheet_name='Resultados')
    output.seek(0)
    session['excel_data'] = output.getvalue()

    return jsonify({
        'success': True,
        'table_html': df.to_html(classes='table table-striped table-hover', index=False, border=0),
        'errors': errores_proceso
    })

@app.route('/download')
def download_excel():
    """Sirve el archivo Excel guardado en la sesión."""
    excel_data = session.pop('excel_data', None)
    if not excel_data:
        return "No hay datos para descargar o la sesión ha expirado.", 404
        
    return send_file(
        io.BytesIO(excel_data),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='resultados_contratos.xlsx'
    )


if __name__ == '__main__':
    app.run(debug=True)