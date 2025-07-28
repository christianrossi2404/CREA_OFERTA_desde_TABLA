import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Pt
import os

# --- Reemplazo de placeholders manteniendo formato ---
def reemplazar_placeholders_mejorado(doc, replacements):
    def replace_text_in_paragraph_or_cell(container, key, value):
        full_container_text = "".join([run.text for run in container.runs])
        if key not in full_container_text:
            return

        first_run_font_properties = None
        if container.runs:
            first_run = container.runs[0]
            first_run_font_properties = {
                'bold': first_run.font.bold,
                'italic': first_run.font.italic,
                'underline': first_run.font.underline,
                'color': first_run.font.color.rgb if first_run.font.color else None,
                'size': first_run.font.size,
                'name': first_run.font.name
            }

        new_full_text = full_container_text.replace(key, str(value))
        for run in container.runs:
            run.text = ""
        new_run = container.add_run(new_full_text)

        if first_run_font_properties:
            if first_run_font_properties['bold'] is not None:
                new_run.font.bold = first_run_font_properties['bold']
            if first_run_font_properties['italic'] is not None:
                new_run.font.italic = first_run_font_properties['italic']
            if first_run_font_properties['underline'] is not None:
                new_run.font.underline = first_run_font_properties['underline']
            if first_run_font_properties['color'] is not None:
                new_run.font.color.rgb = first_run_font_properties['color']
            if first_run_font_properties['size'] is not None:
                new_run.font.size = (first_run_font_properties['size']
                                     if isinstance(first_run_font_properties['size'], Pt)
                                     else Pt(first_run_font_properties['size']))
            if first_run_font_properties['name'] is not None:
                new_run.font.name = first_run_font_properties['name']

    for p in doc.paragraphs:
        for key, value in replacements.items():
            replace_text_in_paragraph_or_cell(p, key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in replacements.items():
                        replace_text_in_paragraph_or_cell(p, key, value)

# --- Combinar documentos Word sin saltos de página ---
def combinar_documentos_word(documentos_ruta, ruta_salida_combinado):
    documento_combinado = Document(documentos_ruta[0])
    for doc_path in documentos_ruta[1:]:
        doc_temp = Document(doc_path)
        for elemento in doc_temp.element.body:
            documento_combinado.element.body.append(elemento)
    documento_combinado.save(ruta_salida_combinado)
    print(f"\nDocumento combinado guardado en: {ruta_salida_combinado}")

# --- Función principal ---
def generar_documentos_word():
    root = tk.Tk()
    root.withdraw()

    excel_file_path = filedialog.askopenfilename(
        title="Selecciona el archivo Excel con los datos",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if not excel_file_path:
        messagebox.showwarning("Ningún archivo seleccionado", "No se ha seleccionado ningún archivo Excel.")
        return

    output_directory = os.path.dirname(excel_file_path)

    # Rutas a plantillas
    word_template_insertable = r"C:\Users\Christian Rossi\PRUEBAS\CREA_WORD_OFERTA\modelo_A_AE.docx"
    word_template_centralizado = r"C:\Users\Christian Rossi\PRUEBAS\CREA_WORD_OFERTA\modelo_E2.docx"

    try:
        # Obtener número de oferta desde el nombre del archivo
        excel_file_name = os.path.basename(excel_file_path)
        offer_number_raw = os.path.splitext(excel_file_name)[0]
        offer_number = offer_number_raw[len("Of-"):] if offer_number_raw.startswith("Of-") else offer_number_raw

        # Leer encabezados (fila 3) y datos desde fila 5
        headers_df = pd.read_excel(excel_file_path, skiprows=2, nrows=1, header=None)
        column_names = headers_df.iloc[0].tolist()
        data_df = pd.read_excel(excel_file_path, skiprows=4, header=None)
        data_df.columns = column_names
        df = data_df

        # Mapeo de columnas a marcadores
        column_mapping = {
            '{{ITEM}}': 'ITEM',
            '{{CAUDAL}}': 'CAUDAL',
            '{{TEMP}}': 'TEMP',
            '{{PRES}}': 'PRESION',
            '{{SUPF}}': 'SUPERFICIE\nFILTRANTE',
            '{{MODELO}}': 'FILTRO\nMODELO',
            '{{TIPO}}': 'TIPO \nFILTRO',
            '{{KW}}': 'POTENCIA',
            '{{RPM}}': 'VELOCIDAD \nRODETE',
            '{{---}}': 'WEIGTH',
            '{{PVP}}': 'PVP'
        }

        num_generated_docs = 0
        documentos_generados = []

        for index, row_data in df.iterrows():
            contador_fila = index + 1

            # Seleccionar plantilla según modelo de filtro
            modelo_filtro = str(row_data.get('TIPO \nFILTRO', '')).strip().lower()
            print(modelo_filtro)
            if modelo_filtro == "insertable":
                plantilla_usada = word_template_insertable
            elif modelo_filtro == "centralizado":
                plantilla_usada = word_template_centralizado
            else:
                plantilla_usada = word_template_insertable  # por defecto

            doc = Document(plantilla_usada)

            replacements = {}
            for placeholder, col_excel_name in column_mapping.items():
                valor = row_data.get(col_excel_name)
                replacements[placeholder] = "" if pd.isna(valor) else str(valor)

            replacements['{{CONT}}'] = str(contador_fila)
            replacements['{{UNIDAD}}'] = "mmca"
            replacements['{{NOF}}'] = offer_number
            replacements['{{CONTADOR}}'] = str(contador_fila)

            reemplazar_placeholders_mejorado(doc, replacements)

            output_file_name = f"documento_generado_{contador_fila}.docx"
            full_output_path = os.path.join(output_directory, output_file_name)
            doc.save(full_output_path)

            documentos_generados.append(full_output_path)
            num_generated_docs += 1
            print(f"Documento '{full_output_path}' generado con éxito.")

        # Combinar todos los documentos generados
        if documentos_generados:
            ruta_doc_combinado = os.path.join(output_directory, f"Of-{offer_number}.docx")
            combinar_documentos_word(documentos_generados, ruta_doc_combinado)

            print("\nLista de documentos Word generados:")
            for ruta in documentos_generados:
                print(f"- {ruta}")

        messagebox.showinfo("Proceso Completado", f"Se han generado {num_generated_docs} documentos Word en:\n{output_directory}")

    except FileNotFoundError:
        messagebox.showerror("Error de archivo", "No se encontró alguna de las plantillas Word. Verifica las rutas.")
    except KeyError as e:
        messagebox.showerror("Error en columnas Excel", f"Columna faltante en el Excel: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")


    # Eliminar todos los archivos que empiecen por "documento_generado_"
    for archivo in os.listdir(output_directory):
        if archivo.startswith("documento_generado_") and archivo.endswith(".docx"):
            try:
                os.remove(os.path.join(output_directory, archivo))
                print(f"Archivo temporal eliminado: {archivo}")
            except Exception as e:
                print(f"No se pudo eliminar {archivo}: {e}")


if __name__ == "__main__":
    generar_documentos_word()