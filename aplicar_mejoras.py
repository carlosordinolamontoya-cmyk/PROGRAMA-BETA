import os
import re
import shutil

def patch_file():
    file_path = "app_kardex.py"
    if not os.path.exists(file_path):
        print(f"Error: No se encontró el archivo '{file_path}' en el directorio actual.")
        return

    # 1. Crear copia de seguridad
    shutil.copy(file_path, "app_kardex_backup.py")
    print("Copia de seguridad creada con éxito: app_kardex_backup.py")

    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    # 2. Arreglar rutas de guardado para evitar errores de permisos en Windows
    content = content.replace(
        'REPORT_DIR = os.path.join(BASE_DIR, "reportes")',
        'USER_DOCS = os.path.join(os.path.expanduser("~"), "Documents", "ERP_Kardex")\nREPORT_DIR = os.path.join(USER_DOCS, "reportes")'
    )
    content = content.replace(
        'VALE_DIR = os.path.join(BASE_DIR, "vales_impresos")',
        'VALE_DIR = os.path.join(USER_DOCS, "vales_impresos")'
    )
    content = content.replace(
        'OC_DIR = os.path.join(BASE_DIR, "ordenes_compra_impresas")',
        'OC_DIR = os.path.join(USER_DOCS, "ordenes_compra_impresas")'
    )

    # 3. Arreglar geometrías muy grandes para pantallas de laptop
    content = re.sub(r'self\.geometry\("1340x760\+25\+25"\)\s+self\.minsize\(1220,\s*680\)', 'self.geometry("1024x680+25+25")\n        self.minsize(1024, 600)', content)
    content = re.sub(r'self\.geometry\("1280x760\+35\+20"\)\s+self\.minsize\(1180,\s*680\)', 'self.geometry("1024x680+35+20")\n        self.minsize(1024, 600)', content)
    content = re.sub(r'self\.geometry\("1260x780\+45\+25"\)\s+self\.minsize\(1120,\s*700\)', 'self.geometry("1024x680+45+25")\n        self.minsize(1024, 600)', content)
    
    # Atrapar otras posibles resoluciones gigantes por seguridad
    content = re.sub(r'self\.geometry\("1[1-9]\d{2}x[7-9]\d{2}\+[0-9]+\+[0-9]+"\)', 'self.geometry("1024x680+30+30")', content)
    content = re.sub(r'self\.minsize\(1[1-9]\d{2},\s*[6-9]\d{2}\)', 'self.minsize(1024, 600)', content)

    # 4. Debounce Autocomplete (Arreglar lag al escribir)
    content = content.replace('self._after_id = self.widget.after(120, self.show)', 'self._after_id = self.widget.after(300, self.show)')

    # 5. Bug de Input Fantasma en la Matriz de Cotización (Scroll wheel fix)
    target_ghost = 'self.tree.bind("<Button-3>", self.on_right_click)\n        self.refresh_tree()'
    replacement_ghost = 'self.tree.bind("<Button-3>", self.on_right_click)\n        self.tree.bind("<MouseWheel>", self.cancel)\n        self.refresh_tree()'
    content = content.replace(target_ghost, replacement_ghost)

    # 6. Atrapar la fila exacta de error en la subida Masiva de Excel
    target_bulk = '''        clean_rows = []
        for row in rows:
            codigo, desc, um, tipo = row[:4]
            minimo_a = row[4] if len(row) > 4 else 0
            minimo_b = row[5] if len(row) > 5 else None
            if minimo_b is None or minimo_b == "":
                clean_rows.append((codigo, desc, um, tipo, minimo_a or 0, None, None))
            else:
                clean_rows.append((codigo, desc, um, tipo, 0, minimo_a or 0, minimo_b or 0))'''
    
    replacement_bulk = '''        clean_rows = []
        for idx_row, row in enumerate(rows, start=2):
            try:
                codigo, desc, um, tipo = row[:4]
                minimo_a = row[4] if len(row) > 4 else 0
                minimo_b = row[5] if len(row) > 5 else None
                if minimo_b is None or minimo_b == "":
                    clean_rows.append((codigo, desc, um, tipo, minimo_a or 0, None, None))
                else:
                    clean_rows.append((codigo, desc, um, tipo, 0, minimo_a or 0, minimo_b or 0))
            except Exception as e:
                raise ValueError(f"Error en la fila {idx_row} del Excel. Detalle: {e}")'''
    content = content.replace(target_bulk, replacement_bulk)

    # Guardar los cambios
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(content)
    
    print("\n=======================================================")
    print(" ¡Modificaciones aplicadas con éxito en app_kardex.py! ")
    print("=======================================================\n")

if __name__ == "__main__":
    patch_file()