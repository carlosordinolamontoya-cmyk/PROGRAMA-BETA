import os
import sys
import sqlite3
import hashlib
import secrets
import json
import subprocess
import urllib.request
import urllib.parse
import urllib.error
from copy import deepcopy
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

from erp_core import (
    D, q_qty, q_unit, q_tc, q_money, calc_price_values as decimal_calc_price_values,
    money_sum, money_fmt, iso_now, normalize_datetime_text,
    AnnualCorrelativeService, configure_app_logger, install_global_exception_logging,
    apply_sqlite_runtime_pragmas, ensure_performance_indexes,
)

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
except Exception:
    Workbook = None
    load_workbook = None

try:
    from docx import Document
except Exception:
    Document = None

APP_TITLE = "Sistema Kardex de Almacén"
BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "kardex_almacen.db")
PRODUCTS_FILE = os.path.join(BASE_DIR, "productos_iniciales.json")
REPORT_DIR = os.path.join(BASE_DIR, "reportes")
VALE_DIR = os.path.join(BASE_DIR, "vales_impresos")
OC_DIR = os.path.join(BASE_DIR, "ordenes_compra_impresas")
OC_WORD_TEMPLATE = os.path.join(BASE_DIR, "Orden_Plantilla_ERP_Final_TIPO_DOC.docx")
VALE_WORD_TEMPLATE = os.path.join(BASE_DIR, "Vale_Almacen_Plantilla.docx")
VALE_ITEMS_PER_PAGE = 12
OC_BASE_ITEM_ROWS = 16
DATE_FMT = "%d/%m/%Y"
DATETIME_FMT = "%Y-%m-%d %H:%M:%S"
ISO_FMT = "%Y-%m-%d"

CENTROS_COSTO = {
    "01": "ALMACEN",
    "02": "OPERACIONES",
    "03": "RRHH / SSOMA",
    "04": "LOGISTICA",
    "05": "TALLER PLANTA",
    "06": "MANT. VEHICULAR",
}

ALMACENES = {
    "AQP": "ALMACEN AREQUIPA",
    "MIN": "ALMACEN MINA",
}

ESTADOS_REQUERIMIENTO = ["PENDIENTE", "DESPACHO PARCIAL", "ATENDIDO", "ANULADO"]

TIPOS_ORDEN_COMPRA = ["NACIONAL", "SERVICIO"]
ESTADOS_OC = ["PENDIENTE", "EN COTIZACION", "APROBADA", "APROBADA PARCIAL", "DENEGADA", "ANULADA", "PAGO PARCIAL", "PAGADA", "EN TRANSITO", "ATENDIDA PARCIALMENTE", "ATENDIDA", "LIQUIDADA"]
ESTADOS_PAGO_OC = ["POR PAGAR", "PAGO PARCIAL", "PAGADA", "VENCIDA"]
MONEDAS = ["SOLES", "DOLARES"]
BANCOS = ["01 BCP", "02 BBVA", "03 INTERBANK", "04 SCOTIABANK", "05 OTROS"]
BANCOS_MAP = {"01": "BCP", "02": "BBVA", "03": "INTERBANK", "04": "SCOTIABANK", "05": "OTROS"}
OPERADORES_LOGISTICOS = ["FLORES", "SHALOM", "MARVISUR", "COMITE 4", "OTROS"]
IGV_RATE = 0.18
FORMAS_PAGO = ["CONTADO", "CREDITO 15 DIAS", "CREDITO 30 DIAS", "50% INICIAL/50% FINAL"]

API_CONFIG_DEFAULTS = {
    "ruc_api_url": "https://api.apis.net.pe/v2/sunat/ruc?numero={ruc}",
    "ruc_api_token": "",
    "tc_api_url": "https://api.apis.net.pe/v2/sunat/tipo-cambio?date={fecha}",
    "tc_api_token": "",
    "api_timeout": "10",
}

FUENTES_TC = ["API_SUNAT", "SUNAT", "SBS", "BCRP", "MANUAL"]

TIPOS_MOVIMIENTO = {
    "CN": "INGRESO",
    "CO": "SALIDA",
    "AJ": "AJUSTE",
}

UNIDADES_PERMITIDAS = ["UND", "KG", "MT", "M3", "CM3", "LT", "KIT", "GL", "M2", "PAR", "JGO", "ROLLO", "BALDE", "VARILLA"]
TIPOS_ARTICULO = ["CONSUMIBLE", "HERRAMIENTA", "SERVICIO"]
ARTICULO_DESC_MAX = 120

ROLES_ESCRITURA = ("ADMIN", "OPERADOR", "ALMACEN", "LOGISTICA", "FINANZAS")
ROLES_ADMIN = ("ADMIN",)
ROLES_ALMACEN = ("ADMIN", "OPERADOR", "ALMACEN")
ROLES_LOGISTICA = ("ADMIN", "OPERADOR", "LOGISTICA")
ROLES_FINANZAS = ("ADMIN", "OPERADOR", "FINANZAS")
ROLES_VALIDOS = ["ADMIN", "OPERADOR", "ALMACEN", "LOGISTICA", "FINANZAS", "CONSULTA"]


# --- Control visual de mensajes ---
# Evita que una advertencia/error mande al usuario al menú principal.
# Si el mensaje pertenece a una ventana hija, al cerrar el mensaje la misma ventana vuelve al frente.
def keep_window_front(parent):
    if parent is None:
        return
    try:
        parent.lift()
        parent.focus_force()
        parent.attributes("-topmost", True)
        parent.after(250, lambda: parent.attributes("-topmost", False))
    except Exception:
        pass


_original_showinfo = messagebox.showinfo
_original_showwarning = messagebox.showwarning
_original_showerror = messagebox.showerror
_original_askyesno = messagebox.askyesno


def _wrap_messagebox(func):
    def wrapper(title=None, message=None, **options):
        parent = options.get("parent")
        keep_window_front(parent)
        result = func(title, message, **options)
        keep_window_front(parent)
        return result
    return wrapper


messagebox.showinfo = _wrap_messagebox(_original_showinfo)
messagebox.showwarning = _wrap_messagebox(_original_showwarning)
messagebox.showerror = _wrap_messagebox(_original_showerror)
messagebox.askyesno = _wrap_messagebox(_original_askyesno)


def bind_as_child_window(win, parent=None):
    try:
        if parent is not None:
            win.transient(parent)
        win.lift()
        win.focus_force()
        # UX: Escape cierra ventanas secundarias sin usar mouse.
        try:
            win.bind("<Escape>", lambda e: win.destroy())
        except Exception:
            pass
    except Exception:
        pass


class AutocompleteHelper:
    """Autocompletado liviano para Tkinter.
    Muestra sugerencias mientras el usuario escribe y permite seleccionar con Enter/doble clic.
    No reemplaza los cuadros Shift+F1; los complementa para capturar más rápido.
    """
    def __init__(self, widget, get_rows, on_select, max_rows=12, min_chars=1, width_chars=70):
        self.widget = widget
        self.get_rows = get_rows
        self.on_select = on_select
        self.max_rows = max_rows
        self.min_chars = min_chars
        self.width_chars = width_chars
        self.popup = None
        self.listbox = None
        self.rows = []
        self._after_id = None
        self._selecting = False
        # Mantener referencia para que el GC no elimine el helper.
        try:
            current = getattr(widget, "_autocomplete_helpers", [])
            current.append(self)
            widget._autocomplete_helpers = current
        except Exception:
            pass
        widget.bind("<KeyRelease>", self._on_keyrelease, add="+")
        widget.bind("<Down>", self._move_down, add="+")
        widget.bind("<Up>", self._move_up, add="+")
        widget.bind("<Right>", self._move_right, add="+")
        widget.bind("<Left>", self._move_left, add="+")
        widget.bind("<Return>", self._on_return, add="+")
        widget.bind("<Escape>", self._on_escape, add="+")
        widget.bind("<FocusOut>", self._on_focus_out, add="+")

    def _widget_text(self):
        try:
            return self.widget.get()
        except Exception:
            return ""

    def _state_disabled(self):
        try:
            return str(self.widget.cget("state")) in ("disabled", "readonly")
        except Exception:
            return False

    def _on_keyrelease(self, event=None):
        if self._selecting or self._state_disabled():
            self.hide()
            return None
        if event and event.keysym in ("Return", "Escape", "Up", "Down", "Left", "Right", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return None
        if self._after_id:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
        self._after_id = self.widget.after(120, self.show)
        return None

    def _on_return(self, event=None):
        if self.popup and self.popup.winfo_exists() and self.rows:
            self.select_current()
            return "break"
        return None

    def _on_escape(self, event=None):
        self.hide()
        return None

    def _on_focus_out(self, event=None):
        # Se retrasa para permitir clic en la lista de sugerencias.
        try:
            self.widget.after(180, self._hide_if_focus_out)
        except Exception:
            self.hide()
        return None

    def _hide_if_focus_out(self):
        try:
            focus = self.widget.focus_get()
            if focus not in (self.widget, self.listbox):
                self.hide()
        except Exception:
            self.hide()

    def _move_down(self, event=None):
        if not (self.popup and self.popup.winfo_exists() and self.rows):
            self.show()
            return "break"
        cur = self.listbox.curselection()
        idx = (cur[0] + 1) if cur else 0
        if idx >= len(self.rows): idx = 0
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(idx)
        self.listbox.activate(idx)
        self.listbox.see(idx)
        return "break"

    def _move_up(self, event=None):
        if not (self.popup and self.popup.winfo_exists() and self.rows):
            self.show()
            return "break"
        cur = self.listbox.curselection()
        idx = (cur[0] - 1) if cur else len(self.rows) - 1
        if idx < 0: idx = len(self.rows) - 1
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(idx)
        self.listbox.activate(idx)
        self.listbox.see(idx)
        return "break"

    def _move_right(self, event=None):
        # En listas verticales, derecha avanza a la siguiente sugerencia; Enter confirma.
        if self.popup and self.popup.winfo_exists() and self.rows:
            return self._move_down(event)
        return None

    def _move_left(self, event=None):
        # En listas verticales, izquierda retrocede a la sugerencia anterior.
        if self.popup and self.popup.winfo_exists() and self.rows:
            return self._move_up(event)
        return None

    def show(self):
        text = self._widget_text()
        if self._state_disabled() or len(normalize_text(text)) < self.min_chars:
            self.hide()
            return
        try:
            rows = list(self.get_rows(text))[:self.max_rows]
        except Exception:
            rows = []
        if not rows:
            self.hide()
            return
        self.rows = rows
        if not self.popup or not self.popup.winfo_exists():
            self.popup = tk.Toplevel(self.widget)
            self.popup.overrideredirect(True)
            try:
                self.popup.transient(self.widget.winfo_toplevel())
            except Exception:
                pass
            self.listbox = tk.Listbox(self.popup, height=min(self.max_rows, len(rows)), width=self.width_chars)
            self.listbox.pack(fill="both", expand=True)
            self.listbox.bind("<Double-Button-1>", lambda e: self.select_current())
            self.listbox.bind("<Return>", lambda e: self.select_current())
            self.listbox.bind("<Escape>", lambda e: self.hide())
            self.listbox.bind("<ButtonRelease-1>", lambda e: self.select_current())
        else:
            self.listbox.delete(0, tk.END)
            self.listbox.configure(height=min(self.max_rows, len(rows)))
        for row in rows:
            self.listbox.insert(tk.END, row[0])
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(0)
        self.listbox.activate(0)
        try:
            x = self.widget.winfo_rootx()
            y = self.widget.winfo_rooty() + self.widget.winfo_height()
            w = max(self.widget.winfo_width(), 420)
            h = min(28 * len(rows) + 6, 300)
            self.popup.geometry(f"{w}x{h}+{x}+{y}")
            self.popup.lift()
        except Exception:
            pass

    def select_current(self):
        if not self.rows:
            return "break"
        try:
            cur = self.listbox.curselection()
            idx = cur[0] if cur else 0
            display, value, payload = self.rows[idx]
        except Exception:
            return "break"
        self._selecting = True
        try:
            self.on_select(value, payload)
        finally:
            self._selecting = False
            self.hide()
            try:
                self.widget.icursor(tk.END)
            except Exception:
                pass
        return "break"

    def hide(self):
        try:
            if self.popup and self.popup.winfo_exists():
                self.popup.destroy()
        except Exception:
            pass
        self.popup = None
        self.listbox = None
        self.rows = []


def attach_solicitante_autocomplete(widget, db, variable, on_select=None):
    def rows(term):
        q = normalize_text(term)
        out = []
        for r in db.list_solicitantes(True):
            nombre = r["nombre"]
            if q in normalize_text(nombre):
                out.append((nombre, nombre, r))
        return out
    def select(value, payload):
        variable.set(value)
        if on_select:
            on_select(value)
    return AutocompleteHelper(widget, rows, select, width_chars=52)


def attach_articulo_autocomplete(widget, db, variable, on_select=None):
    def rows(term):
        out = []
        for r in db.search_articulos(term):
            display = f"{r['codigo']} - {r['descripcion']} | UM {r['unidad_medida']} | AQP {float(r['stock_aqp'] or 0):g} | MIN {float(r['stock_min'] or 0):g}"
            out.append((display, r["codigo"], r))
        return out
    def select(value, payload):
        variable.set(value)
        if on_select:
            on_select(value)
    return AutocompleteHelper(widget, rows, select, width_chars=90)


def attach_proveedor_autocomplete(widget, db, variable, on_select=None):
    def rows(term):
        out = []
        for r in db.list_proveedores(term, active_only=True):
            display = f"{r['ruc']} - {r['razon_social']}"
            if r["banco"]:
                display += f" | {r['banco']}"
            out.append((display, r["ruc"], r))
        return out
    def select(value, payload):
        variable.set(value)
        if on_select:
            on_select(value)
    return AutocompleteHelper(widget, rows, select, width_chars=82)


def now_text():
    return datetime.now().strftime(DATETIME_FMT)


def today_text():
    return date.today().strftime(DATE_FMT)


MESES_ES = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def fecha_larga_es(fecha_iso_o_texto):
    try:
        if "-" in str(fecha_iso_o_texto):
            f = datetime.strptime(str(fecha_iso_o_texto), ISO_FMT).date()
        else:
            f = parse_fecha(str(fecha_iso_o_texto))
        return f"{f.day} de {MESES_ES[f.month]} de {f.year}"
    except Exception:
        return str(fecha_iso_o_texto or "")


def parse_fecha(texto):
    try:
        return datetime.strptime(texto.strip(), DATE_FMT).date()
    except Exception:
        raise ValueError("La fecha debe tener formato DD/MM/AAAA.")


def fecha_iso_from_text(texto, allow_future=False):
    f = parse_fecha(texto)
    if f > date.today() and not allow_future:
        raise ValueError("No se permite registrar movimientos con fecha futura.")
    return f.strftime(ISO_FMT)


def fecha_iso_consulta_from_text(texto):
    """Convierte fechas de filtros/reportes. Permite fechas futuras porque no registra movimientos."""
    return fecha_iso_from_text(texto, allow_future=True)


def fecha_text_from_iso(value):
    if not value:
        return ""
    try:
        return datetime.strptime(value[:10], ISO_FMT).strftime(DATE_FMT)
    except Exception:
        return value


def normalize_text(value):
    return str(value or "").strip().upper()


def clean_ot(value):
    value = normalize_text(value).replace("OT-", "").replace("OT ", "").strip()
    return f"OT-{value}" if value else ""


def only_digits(value):
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def normalizar_banco(value, banco_otro=""):
    """Normaliza banco por código para pagos/proveedores.
    Acepta '01', '01 BCP', 'BCP' o '05 OTROS'.
    Si es OTROS, concatena el nombre escrito por el usuario.
    """
    raw = normalize_text(value)
    otro = normalize_text(banco_otro)
    if not raw:
        return ""
    code = raw.split()[0] if raw.split() else raw
    if code in BANCOS_MAP:
        base = f"{code} {BANCOS_MAP[code]}"
    else:
        inv = {v: k for k, v in BANCOS_MAP.items()}
        code = inv.get(raw, "")
        base = f"{code} {raw}" if code else raw
    if base.startswith("05") and otro:
        return f"05 OTROS - {otro}"
    return base


def banco_codigo_guardado(value):
    raw = normalize_text(value)
    if not raw:
        return ""
    code = raw.split()[0]
    if code in BANCOS_MAP:
        return code
    inv = {v: k for k, v in BANCOS_MAP.items()}
    return inv.get(raw, "")


def money_to_words_es(amount, moneda="SOLES"):
    """Convierte un importe a una frase simple para impresión de OC."""
    try:
        amount = round(float(amount or 0), 2)
    except Exception:
        amount = 0.0
    entero = int(amount)
    centimos = int(round((amount - entero) * 100))
    unidades = ["CERO", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE"]
    decenas = ["", "", "VEINTI", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
    centenas = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"]

    def n_to_words(n):
        n = int(n)
        if n <= 20:
            return unidades[n]
        if n < 30:
            return "VEINTI" + unidades[n-20].lower().upper()
        if n < 100:
            d, u = divmod(n, 10)
            return decenas[d] if u == 0 else f"{decenas[d]} Y {unidades[u]}"
        if n == 100:
            return "CIEN"
        if n < 1000:
            c, r = divmod(n, 100)
            return centenas[c] if r == 0 else f"{centenas[c]} {n_to_words(r)}"
        if n < 1000000:
            m, r = divmod(n, 1000)
            pref = "MIL" if m == 1 else f"{n_to_words(m)} MIL"
            return pref if r == 0 else f"{pref} {n_to_words(r)}"
        mill, r = divmod(n, 1000000)
        pref = "UN MILLON" if mill == 1 else f"{n_to_words(mill)} MILLONES"
        return pref if r == 0 else f"{pref} {n_to_words(r)}"

    return f"{n_to_words(entero)} CON {centimos:02d}/100 {normalize_text(moneda)}"


def iter_docx_paragraphs(document):
    for paragraph in document.paragraphs:
        yield paragraph
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph
                for nested in cell.tables:
                    for nrow in nested.rows:
                        for ncell in nrow.cells:
                            for paragraph in ncell.paragraphs:
                                yield paragraph


def replace_docx_placeholders(document, mapping):
    """Reemplaza marcadores {{CAMPO}} en párrafos y tablas de una plantilla Word.
    Si Word divide un marcador en varios runs, reconstruye el párrafo conservando el estilo base.
    """
    for paragraph in iter_docx_paragraphs(document):
        text = paragraph.text
        if not text or "{{" not in text:
            continue
        new_text = text
        for key, value in mapping.items():
            new_text = new_text.replace(key, str(value or ""))
        if new_text != text:
            if paragraph.runs:
                paragraph.runs[0].text = new_text
                for run in paragraph.runs[1:]:
                    run.text = ""
            else:
                paragraph.add_run(new_text)


def _docx_set_paragraph_text(paragraph, value):
    """Reemplaza el texto visible de un párrafo conservando el formato del primer run."""
    value = "" if value is None else str(value)
    if paragraph.runs:
        paragraph.runs[0].text = value
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(value)


def _docx_set_cell_text(cell, value):
    if not cell.paragraphs:
        cell.add_paragraph()
    _docx_set_paragraph_text(cell.paragraphs[0], value)
    for paragraph in cell.paragraphs[1:]:
        _docx_set_paragraph_text(paragraph, "")


def _docx_remove_trailing_empty_body_paragraph(document):
    """Evita que un párrafo vacío al final de una plantilla empuje la página siguiente."""
    body = document.element.body
    elements = list(body)
    if len(elements) < 2:
        return
    candidate = elements[-2]  # antes de sectPr
    if candidate.tag.endswith('}p') and not ''.join(candidate.itertext()).strip():
        body.remove(candidate)


def _docx_append_body(destination, source):
    """Anexa el cuerpo de otro DOCX antes del sectPr, sin duplicar secciones."""
    from docx.oxml.ns import qn
    body = destination.element.body
    sect_pr = body.sectPr
    for element in list(source.element.body):
        if element.tag == qn('w:sectPr'):
            continue
        sect_pr.addprevious(deepcopy(element))


def _find_docx_table(document, required_headers):
    wanted = {normalize_text(x).upper() for x in required_headers}
    for table in document.tables:
        if not table.rows:
            continue
        first = {normalize_text(cell.text).upper() for cell in table.rows[0].cells}
        if wanted.issubset(first):
            return table
    return None


def require_docx_document(parent=None):
    """Carga python-docx y, si falta, intenta instalarlo automáticamente.
    Evita que el usuario se quede bloqueado con el mensaje "Instale python-docx".
    """
    global Document
    if Document is not None:
        return Document
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx>=1.1.0"], stdout=subprocess.DEVNULL)
        from docx import Document as _Doc
        Document = _Doc
        return Document
    except Exception as e:
        raise ValueError(
            "No se pudo cargar python-docx.\n\n"
            "Solución manual: abra la terminal de VS Code en esta carpeta y ejecute:\n"
            "python -m pip install -r requirements_app.txt\n\n"
            f"Detalle técnico: {e}"
        )



def fmt_money(value):
    try:
        return money_fmt(value)
    except Exception:
        return "0.00"


def currency_symbol(moneda):
    mon = normalize_text(moneda)
    if mon in ("DOLARES", "DOLARES AMERICANOS", "USD", "US$"):
        return "US$"
    return "S/"


def fmt_money_currency(value, moneda):
    return f"{currency_symbol(moneda)} {fmt_money(value)}"



class KardexDB:
    def __init__(self, db_path=DB_FILE):
        self.db_path = db_path
        self.conn = sqlite3.connect(self.db_path)
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.row_factory = sqlite3.Row
        self.create_tables()
        self.migrate_tables()
        self.ensure_default_user()
        self.ensure_user_warehouse_defaults()
        self.load_initial_products()

    def create_tables(self):
        cur = self.conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT UNIQUE NOT NULL,
                nombre TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                rol TEXT NOT NULL DEFAULT 'OPERADOR',
                activo INTEGER NOT NULL DEFAULT 1,
                creado_en TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS articulos (
                codigo TEXT PRIMARY KEY,
                descripcion TEXT NOT NULL,
                unidad_medida TEXT NOT NULL,
                tipo_articulo TEXT NOT NULL,
                stock_minimo REAL NOT NULL DEFAULT 0,
                stock_minimo_aqp REAL NOT NULL DEFAULT 0,
                stock_minimo_min REAL NOT NULL DEFAULT 0,
                fecha_registro TEXT NOT NULL,
                activo INTEGER NOT NULL DEFAULT 1
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS solicitantes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                activo INTEGER NOT NULL DEFAULT 1,
                creado_en TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS movimientos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                vale TEXT UNIQUE NOT NULL,
                fecha TEXT NOT NULL,
                fecha_operacion TEXT NOT NULL,
                tipo_codigo TEXT NOT NULL,
                tipo_nombre TEXT NOT NULL,
                documento TEXT NOT NULL,
                ot TEXT,
                centro_costo_codigo TEXT NOT NULL,
                centro_costo_nombre TEXT NOT NULL,
                solicitante TEXT NOT NULL DEFAULT '',
                observacion TEXT NOT NULL DEFAULT '',
                almacen_codigo TEXT NOT NULL DEFAULT 'AQP',
                almacen_nombre TEXT NOT NULL DEFAULT 'ALMACEN AREQUIPA',
                requerimiento TEXT NOT NULL DEFAULT '',
                usuario TEXT NOT NULL DEFAULT '',
                creado_en TEXT NOT NULL,
                modificado_en TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS movimiento_detalle (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                movimiento_id INTEGER NOT NULL,
                item INTEGER NOT NULL,
                codigo TEXT NOT NULL,
                descripcion TEXT NOT NULL,
                unidad_medida TEXT NOT NULL,
                cantidad REAL NOT NULL,
                precio_unitario_sin_igv REAL NOT NULL DEFAULT 0,
                subtotal_sin_igv REAL NOT NULL DEFAULT 0,
                FOREIGN KEY(movimiento_id) REFERENCES movimientos(id) ON DELETE RESTRICT,
                FOREIGN KEY(codigo) REFERENCES articulos(codigo) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS requerimientos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                requerimiento TEXT UNIQUE NOT NULL,
                fecha TEXT NOT NULL,
                fecha_requerimiento TEXT NOT NULL,
                almacen_codigo TEXT NOT NULL DEFAULT 'AQP',
                almacen_nombre TEXT NOT NULL DEFAULT 'ALMACEN AREQUIPA',
                solicitante TEXT NOT NULL,
                centro_costo_codigo TEXT NOT NULL,
                centro_costo_nombre TEXT NOT NULL,
                ot TEXT,
                observacion TEXT NOT NULL DEFAULT '',
                estado TEXT NOT NULL DEFAULT 'PENDIENTE',
                usuario_creador TEXT NOT NULL DEFAULT '',
                creado_en TEXT NOT NULL,
                modificado_en TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS requerimiento_detalle (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                requerimiento_id INTEGER NOT NULL,
                item INTEGER NOT NULL,
                codigo TEXT NOT NULL,
                descripcion TEXT NOT NULL,
                unidad_medida TEXT NOT NULL,
                cantidad_requerida REAL NOT NULL,
                cantidad_atendida REAL NOT NULL DEFAULT 0,
                observacion TEXT NOT NULL DEFAULT '',
                FOREIGN KEY(requerimiento_id) REFERENCES requerimientos(id) ON DELETE RESTRICT,
                FOREIGN KEY(codigo) REFERENCES articulos(codigo) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS requerimiento_atenciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                requerimiento_id INTEGER NOT NULL,
                vale TEXT NOT NULL,
                fecha_atencion TEXT NOT NULL,
                usuario TEXT NOT NULL,
                estado_atencion TEXT NOT NULL,
                observacion TEXT NOT NULL DEFAULT '',
                FOREIGN KEY(requerimiento_id) REFERENCES requerimientos(id) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS transferencias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                transferencia TEXT UNIQUE NOT NULL,
                fecha TEXT NOT NULL,
                fecha_operacion TEXT NOT NULL,
                almacen_origen TEXT NOT NULL,
                almacen_destino TEXT NOT NULL,
                vale_salida TEXT NOT NULL,
                vale_ingreso TEXT NOT NULL,
                solicitante TEXT NOT NULL DEFAULT '',
                centro_costo_codigo TEXT NOT NULL DEFAULT '',
                centro_costo_nombre TEXT NOT NULL DEFAULT '',
                observacion TEXT NOT NULL DEFAULT '',
                numero_guia TEXT NOT NULL DEFAULT '',
                usuario TEXT NOT NULL DEFAULT '',
                creado_en TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS movimiento_reversos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                vale_original TEXT UNIQUE NOT NULL,
                vale_reverso TEXT UNIQUE NOT NULL,
                motivo TEXT NOT NULL,
                usuario TEXT NOT NULL,
                creado_en TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS usuario_almacen_permiso (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario_id INTEGER NOT NULL,
                almacen_codigo TEXT NOT NULL,
                puede_mover INTEGER NOT NULL DEFAULT 1,
                UNIQUE(usuario_id, almacen_codigo),
                FOREIGN KEY(usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS proveedores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ruc TEXT UNIQUE NOT NULL,
                razon_social TEXT NOT NULL,
                direccion TEXT NOT NULL DEFAULT '',
                telefono TEXT NOT NULL DEFAULT '',
                email TEXT NOT NULL DEFAULT '',
                banco TEXT NOT NULL DEFAULT '',
                numero_cuenta TEXT NOT NULL DEFAULT '',
                cci TEXT NOT NULL DEFAULT '',
                contacto TEXT NOT NULL DEFAULT '',
                moneda_preferida TEXT NOT NULL DEFAULT 'SOLES',
                activo INTEGER NOT NULL DEFAULT 1,
                creado_en TEXT NOT NULL,
                modificado_en TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS sistema_config (
                clave TEXT PRIMARY KEY,
                valor TEXT NOT NULL DEFAULT '',
                descripcion TEXT NOT NULL DEFAULT '',
                modificado_por TEXT NOT NULL DEFAULT '',
                modificado_en TEXT NOT NULL DEFAULT ''
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS tipo_cambio (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL,
                moneda TEXT NOT NULL DEFAULT 'DOLARES',
                compra REAL NOT NULL DEFAULT 0,
                venta REAL NOT NULL,
                fuente TEXT NOT NULL DEFAULT 'MANUAL',
                es_manual INTEGER NOT NULL DEFAULT 0,
                motivo_manual TEXT NOT NULL DEFAULT '',
                usuario_registro TEXT NOT NULL DEFAULT '',
                fecha_registro TEXT NOT NULL,
                UNIQUE(fecha, moneda, fuente)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS orden_compra (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                numero TEXT UNIQUE NOT NULL,
                fecha TEXT NOT NULL,
                fecha_orden TEXT NOT NULL,
                almacen_codigo TEXT NOT NULL DEFAULT 'AQP',
                almacen_nombre TEXT NOT NULL DEFAULT 'ALMACEN AREQUIPA',
                proveedor_id INTEGER NOT NULL,
                ruc TEXT NOT NULL,
                razon_social TEXT NOT NULL,
                direccion TEXT NOT NULL DEFAULT '',
                telefono TEXT NOT NULL DEFAULT '',
                tipo_orden TEXT NOT NULL DEFAULT 'NACIONAL',
                requerimiento TEXT NOT NULL DEFAULT '',
                cotizacion TEXT NOT NULL DEFAULT '',
                solicitante TEXT NOT NULL DEFAULT '',
                ot TEXT NOT NULL DEFAULT '',
                centro_costo_codigo TEXT NOT NULL DEFAULT '',
                centro_costo_nombre TEXT NOT NULL DEFAULT '',
                observacion TEXT NOT NULL DEFAULT '',
                moneda TEXT NOT NULL DEFAULT 'SOLES',
                tipo_cambio REAL NOT NULL DEFAULT 1,
                tipo_cambio_compra REAL NOT NULL DEFAULT 1,
                tipo_cambio_venta REAL NOT NULL DEFAULT 1,
                tipo_cambio_fuente TEXT NOT NULL DEFAULT '',
                tipo_cambio_fecha TEXT NOT NULL DEFAULT '',
                precio_incluye_igv INTEGER NOT NULL DEFAULT 0,
                subtotal REAL NOT NULL DEFAULT 0,
                igv REAL NOT NULL DEFAULT 0,
                total REAL NOT NULL DEFAULT 0,
                estado TEXT NOT NULL DEFAULT 'PENDIENTE',
                avance INTEGER NOT NULL DEFAULT 0,
                forma_pago TEXT NOT NULL DEFAULT 'CONTADO',
                fecha_vencimiento_pago TEXT NOT NULL DEFAULT '',
                estado_pago TEXT NOT NULL DEFAULT 'POR PAGAR',
                monto_pagado_cache REAL NOT NULL DEFAULT 0,
                fecha_entrega TEXT NOT NULL DEFAULT '',
                lugar_entrega TEXT NOT NULL DEFAULT '',
                usuario_creador TEXT NOT NULL DEFAULT '',
                aprobado_por TEXT NOT NULL DEFAULT '',
                fecha_aprobacion TEXT NOT NULL DEFAULT '',
                motivo_decision TEXT NOT NULL DEFAULT '',
                creado_en TEXT NOT NULL,
                modificado_en TEXT,
                FOREIGN KEY(proveedor_id) REFERENCES proveedores(id) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS orden_compra_detalle (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                oc_id INTEGER NOT NULL,
                item INTEGER NOT NULL,
                codigo TEXT NOT NULL DEFAULT '',
                descripcion TEXT NOT NULL,
                unidad_medida TEXT NOT NULL,
                cantidad_solicitada REAL NOT NULL,
                cantidad_aprobada REAL NOT NULL DEFAULT 0,
                cantidad_ingresada REAL NOT NULL DEFAULT 0,
                precio_unitario_sin_igv REAL NOT NULL DEFAULT 0,
                precio_unitario_con_igv REAL NOT NULL DEFAULT 0,
                subtotal REAL NOT NULL DEFAULT 0,
                igv REAL NOT NULL DEFAULT 0,
                total REAL NOT NULL DEFAULT 0,
                centro_costo_codigo TEXT NOT NULL DEFAULT '',
                ot TEXT NOT NULL DEFAULT '',
                solicitante TEXT NOT NULL DEFAULT '',
                FOREIGN KEY(oc_id) REFERENCES orden_compra(id) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS orden_compra_estado_historial (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                oc_id INTEGER NOT NULL,
                estado_anterior TEXT NOT NULL DEFAULT '',
                estado_nuevo TEXT NOT NULL,
                usuario TEXT NOT NULL,
                fecha_hora TEXT NOT NULL,
                detalle TEXT NOT NULL DEFAULT '',
                FOREIGN KEY(oc_id) REFERENCES orden_compra(id) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS orden_compra_pagos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                oc_id INTEGER NOT NULL,
                fecha_pago TEXT NOT NULL,
                numero_operacion TEXT NOT NULL,
                numero_factura TEXT NOT NULL DEFAULT '',
                banco TEXT NOT NULL,
                monto REAL NOT NULL,
                observacion TEXT NOT NULL DEFAULT '',
                usuario TEXT NOT NULL,
                creado_en TEXT NOT NULL,
                FOREIGN KEY(oc_id) REFERENCES orden_compra(id) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS orden_compra_transito (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                oc_id INTEGER NOT NULL,
                operador_logistico TEXT NOT NULL,
                fecha_despacho TEXT NOT NULL,
                fecha_estimada_llegada TEXT NOT NULL DEFAULT '',
                guia_transportista TEXT NOT NULL DEFAULT '',
                observacion TEXT NOT NULL DEFAULT '',
                usuario TEXT NOT NULL,
                creado_en TEXT NOT NULL,
                FOREIGN KEY(oc_id) REFERENCES orden_compra(id) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS orden_compra_recepciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                oc_id INTEGER NOT NULL,
                vale TEXT NOT NULL,
                documento_recepcion TEXT NOT NULL,
                fecha_recepcion TEXT NOT NULL,
                observacion TEXT NOT NULL DEFAULT '',
                usuario TEXT NOT NULL,
                creado_en TEXT NOT NULL,
                FOREIGN KEY(oc_id) REFERENCES orden_compra(id) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS orden_compra_recepcion_detalle (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                recepcion_id INTEGER NOT NULL,
                oc_detalle_id INTEGER NOT NULL,
                codigo TEXT NOT NULL DEFAULT '',
                descripcion TEXT NOT NULL,
                unidad_medida TEXT NOT NULL,
                cantidad REAL NOT NULL,
                FOREIGN KEY(recepcion_id) REFERENCES orden_compra_recepciones(id) ON DELETE RESTRICT,
                FOREIGN KEY(oc_detalle_id) REFERENCES orden_compra_detalle(id) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS orden_compra_liquidacion (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                oc_id INTEGER UNIQUE NOT NULL,
                factura TEXT NOT NULL DEFAULT '',
                guia TEXT NOT NULL DEFAULT '',
                fecha_liquidacion TEXT NOT NULL,
                observacion TEXT NOT NULL DEFAULT '',
                usuario TEXT NOT NULL,
                creado_en TEXT NOT NULL,
                FOREIGN KEY(oc_id) REFERENCES orden_compra(id) ON DELETE RESTRICT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS auditoria (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tabla TEXT NOT NULL,
                registro TEXT NOT NULL,
                accion TEXT NOT NULL,
                usuario TEXT NOT NULL,
                fecha_hora TEXT NOT NULL,
                detalle TEXT NOT NULL
            )
        """)
        self.conn.commit()

    def columns(self, table):
        cur = self.conn.cursor()
        cur.execute(f"PRAGMA table_info({table})")
        return {row[1] for row in cur.fetchall()}

    def add_column_if_missing(self, table, column, definition):
        if column not in self.columns(table):
            self.conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
            self.conn.commit()

    def migrate_tables(self):
        self.add_column_if_missing("articulos", "stock_minimo", "REAL NOT NULL DEFAULT 0")
        self.add_column_if_missing("articulos", "stock_minimo_aqp", "REAL NOT NULL DEFAULT 0")
        self.add_column_if_missing("articulos", "stock_minimo_min", "REAL NOT NULL DEFAULT 0")
        self.add_column_if_missing("articulos", "valoriza", "INTEGER NOT NULL DEFAULT 1")
        self.add_column_if_missing("articulos", "permite_decimales", "INTEGER NOT NULL DEFAULT 1")
        self.add_column_if_missing("movimiento_detalle", "precio_unitario_sin_igv", "REAL NOT NULL DEFAULT 0")
        self.add_column_if_missing("movimiento_detalle", "subtotal_sin_igv", "REAL NOT NULL DEFAULT 0")
        self.add_column_if_missing("transferencias", "numero_guia", "TEXT NOT NULL DEFAULT ''")
        self.conn.execute("UPDATE articulos SET stock_minimo_aqp = stock_minimo WHERE IFNULL(stock_minimo_aqp,0)=0 AND IFNULL(stock_minimo,0)>0")
        self.conn.execute("UPDATE articulos SET stock_minimo_min = stock_minimo WHERE IFNULL(stock_minimo_min,0)=0 AND IFNULL(stock_minimo,0)>0")
        self.add_column_if_missing("movimientos", "fecha_operacion", "TEXT")
        self.add_column_if_missing("movimientos", "solicitante", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("movimientos", "observacion", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("movimientos", "almacen_codigo", "TEXT NOT NULL DEFAULT 'AQP'")
        self.add_column_if_missing("movimientos", "almacen_nombre", "TEXT NOT NULL DEFAULT 'ALMACEN AREQUIPA'")
        self.add_column_if_missing("movimientos", "requerimiento", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("movimientos", "usuario", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("movimientos", "modificado_en", "TEXT")
        self.conn.execute("UPDATE movimientos SET fecha_operacion = fecha WHERE fecha_operacion IS NULL OR fecha_operacion = ''")
        self.conn.execute("UPDATE movimientos SET almacen_codigo = 'AQP' WHERE almacen_codigo IS NULL OR almacen_codigo = ''")
        self.conn.execute("UPDATE movimientos SET almacen_nombre = 'ALMACEN AREQUIPA' WHERE almacen_nombre IS NULL OR almacen_nombre = ''")
        self.add_column_if_missing("requerimientos", "tipo_orden", "TEXT NOT NULL DEFAULT 'NACIONAL'")
        self.add_column_if_missing("requerimientos", "fecha_entrega_solicitada", "TEXT NOT NULL DEFAULT ''")
        self.conn.execute("UPDATE requerimientos SET tipo_orden='NACIONAL' WHERE tipo_orden IS NULL OR tipo_orden=''")
        # Campos tributarios para proveedores consultados por API/SUNAT
        self.add_column_if_missing("proveedores", "estado_contribuyente", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("proveedores", "condicion_domicilio", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("proveedores", "ubigeo", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("proveedores", "fuente_ruc", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("proveedores", "fecha_validacion_ruc", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("proveedores", "respuesta_ruc_json", "TEXT NOT NULL DEFAULT ''")
        # Tipo de cambio histórico congelado en la OC emitida
        self.add_column_if_missing("orden_compra", "tipo_cambio", "REAL NOT NULL DEFAULT 1")
        self.add_column_if_missing("orden_compra", "tipo_cambio_compra", "REAL NOT NULL DEFAULT 1")
        self.add_column_if_missing("orden_compra", "tipo_cambio_venta", "REAL NOT NULL DEFAULT 1")
        self.add_column_if_missing("orden_compra", "tipo_cambio_fuente", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("orden_compra", "tipo_cambio_fecha", "TEXT NOT NULL DEFAULT ''")
        # Finanzas v22: estado de pago independiente del estado operativo/logístico.
        # Una OC puede estar EN TRANSITO, ATENDIDA o LIQUIDADA y seguir POR PAGAR si es a crédito.
        self.add_column_if_missing("orden_compra", "fecha_vencimiento_pago", "TEXT NOT NULL DEFAULT ''")
        self.add_column_if_missing("orden_compra", "estado_pago", "TEXT NOT NULL DEFAULT 'POR PAGAR'")
        self.add_column_if_missing("orden_compra", "monto_pagado_cache", "REAL NOT NULL DEFAULT 0")
        self.add_column_if_missing("orden_compra_pagos", "numero_factura", "TEXT NOT NULL DEFAULT ''")
        self.conn.execute("UPDATE orden_compra SET estado_pago='PAGADA' WHERE estado IN ('PAGADA') AND IFNULL(estado_pago,'') IN ('','POR PAGAR')")
        self.conn.execute("UPDATE orden_compra SET estado_pago='PAGO PARCIAL' WHERE estado IN ('PAGO PARCIAL') AND IFNULL(estado_pago,'') IN ('','POR PAGAR')")
        self.conn.execute("UPDATE orden_compra SET fecha_vencimiento_pago=fecha_orden WHERE IFNULL(fecha_vencimiento_pago,'')='' AND forma_pago='CONTADO'")
        self.ensure_default_config()
        self.conn.commit()
        self.ensure_user_warehouse_defaults()

    def ensure_default_config(self):
        cur = self.conn.cursor()
        for clave, valor in API_CONFIG_DEFAULTS.items():
            cur.execute(
                "INSERT OR IGNORE INTO sistema_config(clave, valor, descripcion, modificado_por, modificado_en) VALUES (?, ?, ?, '', ?)",
                (clave, valor, f"Configuración {clave}", now_text())
            )
        self.conn.commit()

    def hash_password(self, password, salt=None):
        salt = salt or secrets.token_hex(16)
        digest = hashlib.sha256((salt + password).encode("utf-8")).hexdigest()
        return digest, salt

    def ensure_default_user(self):
        cur = self.conn.cursor()
        cur.execute("SELECT COUNT(*) AS total FROM usuarios")
        if cur.fetchone()["total"] == 0:
            digest, salt = self.hash_password("admin123")
            cur.execute("""
                INSERT INTO usuarios(usuario, nombre, password_hash, salt, rol, activo, creado_en)
                VALUES (?, ?, ?, ?, ?, 1, ?)
            """, ("admin", "ADMINISTRADOR", digest, salt, "ADMIN", now_text()))
            self.conn.commit()

    def verify_user(self, usuario, password):
        usuario = normalize_text(usuario).lower()
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM usuarios WHERE LOWER(usuario)=LOWER(?) AND activo=1", (usuario,))
        row = cur.fetchone()
        if not row:
            return None
        digest, _ = self.hash_password(password, row["salt"])
        return row if digest == row["password_hash"] else None

    def create_user(self, usuario, nombre, password, rol="OPERADOR"):
        usuario = str(usuario or "").strip().lower()
        nombre = normalize_text(nombre)
        rol = normalize_text(rol) or "OPERADOR"
        if not usuario or not nombre or not password:
            raise ValueError("Complete usuario, nombre y contraseña.")
        if len(password) < 4:
            raise ValueError("La contraseña debe tener mínimo 4 caracteres.")
        if rol not in ROLES_VALIDOS:
            raise ValueError("Rol inválido.")
        digest, salt = self.hash_password(password)
        cur = self.conn.cursor()
        cur.execute("""
            INSERT INTO usuarios(usuario, nombre, password_hash, salt, rol, activo, creado_en)
            VALUES (?, ?, ?, ?, ?, 1, ?)
        """, (usuario, nombre, digest, salt, rol, now_text()))
        user_id = cur.lastrowid
        for alm in ALMACENES:
            cur.execute("INSERT OR IGNORE INTO usuario_almacen_permiso(usuario_id, almacen_codigo, puede_mover) VALUES (?, ?, 1)", (user_id, alm))
        self.conn.commit()

    def list_users(self):
        cur = self.conn.cursor()
        cur.execute("SELECT id, usuario, nombre, rol, activo, creado_en FROM usuarios ORDER BY usuario")
        return cur.fetchall()

    def update_user_password(self, user_id, password):
        if len(password) < 4:
            raise ValueError("La contraseña debe tener mínimo 4 caracteres.")
        digest, salt = self.hash_password(password)
        self.conn.execute("UPDATE usuarios SET password_hash=?, salt=? WHERE id=?", (digest, salt, user_id))
        self.conn.commit()

    def set_user_active(self, user_id, active):
        self.conn.execute("UPDATE usuarios SET activo=? WHERE id=?", (1 if active else 0, user_id))
        self.conn.commit()

    def ensure_user_warehouse_defaults(self):
        """Evita bloquear usuarios existentes al migrar. El ADMIN y los roles operativos reciben AQP/MIN por defecto."""
        cur = self.conn.cursor()
        cur.execute("SELECT id, rol FROM usuarios")
        for r in cur.fetchall():
            cur.execute("SELECT COUNT(*) AS n FROM usuario_almacen_permiso WHERE usuario_id=?", (r["id"],))
            if int(cur.fetchone()["n"] or 0) == 0:
                default_perm = 1
                for alm in ALMACENES:
                    cur.execute("INSERT OR IGNORE INTO usuario_almacen_permiso(usuario_id, almacen_codigo, puede_mover) VALUES (?, ?, ?)", (r["id"], alm, default_perm))
        self.conn.commit()

    def set_user_almacen_permissions(self, user_id, almacenes_permitidos):
        almacenes_permitidos = {normalize_text(a) for a in (almacenes_permitidos or []) if normalize_text(a) in ALMACENES}
        cur = self.conn.cursor()
        for alm in ALMACENES:
            cur.execute("INSERT OR IGNORE INTO usuario_almacen_permiso(usuario_id, almacen_codigo, puede_mover) VALUES (?, ?, 0)", (user_id, alm))
            cur.execute("UPDATE usuario_almacen_permiso SET puede_mover=? WHERE usuario_id=? AND almacen_codigo=?", (1 if alm in almacenes_permitidos else 0, user_id, alm))
        self.conn.commit()

    def user_almacen_permissions(self, user_id):
        cur = self.conn.cursor()
        cur.execute("SELECT almacen_codigo, puede_mover FROM usuario_almacen_permiso WHERE usuario_id=?", (user_id,))
        rows = cur.fetchall()
        if not rows:
            self.ensure_user_warehouse_defaults()
            cur.execute("SELECT almacen_codigo, puede_mover FROM usuario_almacen_permiso WHERE usuario_id=?", (user_id,))
            rows = cur.fetchall()
        return {r["almacen_codigo"]: bool(r["puede_mover"]) for r in rows}

    def user_can_access_almacen(self, user_row_or_id, almacen_codigo):
        almacen_codigo = normalize_text(almacen_codigo or "AQP")
        if almacen_codigo not in ALMACENES:
            return False
        if isinstance(user_row_or_id, sqlite3.Row):
            if user_row_or_id["rol"] == "ADMIN":
                return True
            user_id = user_row_or_id["id"]
        else:
            user_id = int(user_row_or_id)
            cur = self.conn.cursor()
            cur.execute("SELECT rol FROM usuarios WHERE id=?", (user_id,))
            u = cur.fetchone()
            if u and u["rol"] == "ADMIN":
                return True
        perms = self.user_almacen_permissions(user_id)
        return bool(perms.get(almacen_codigo, False))

    def audit(self, tabla, registro, accion, usuario, detalle, commit=True):
        self.conn.execute("""
            INSERT INTO auditoria(tabla, registro, accion, usuario, fecha_hora, detalle)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (tabla, str(registro), accion, usuario, now_text(), detalle))
        if commit:
            self.conn.commit()

    def list_solicitantes(self, active_only=True):
        cur = self.conn.cursor()
        if active_only:
            cur.execute("SELECT id, nombre, activo, creado_en FROM solicitantes WHERE activo=1 ORDER BY nombre")
        else:
            cur.execute("SELECT id, nombre, activo, creado_en FROM solicitantes ORDER BY nombre")
        return cur.fetchall()

    def solicitante_names(self):
        return [r["nombre"] for r in self.list_solicitantes(True)]

    def add_solicitante(self, nombre):
        nombre = normalize_text(nombre)
        if not nombre:
            raise ValueError("Ingrese el nombre del solicitante.")
        cur = self.conn.cursor()
        cur.execute("SELECT id, activo FROM solicitantes WHERE UPPER(nombre)=UPPER(?)", (nombre,))
        existing = cur.fetchone()
        if existing:
            if existing["activo"]:
                raise ValueError("El solicitante ya existe.")
            self.conn.execute("UPDATE solicitantes SET activo=1 WHERE id=?", (existing["id"],))
        else:
            self.conn.execute("INSERT INTO solicitantes(nombre, activo, creado_en) VALUES (?, 1, ?)", (nombre, now_text()))
        self.conn.commit()

    def deactivate_solicitante(self, solicitante_id):
        self.conn.execute("UPDATE solicitantes SET activo=0 WHERE id=?", (solicitante_id,))
        self.conn.commit()

    def load_initial_products(self):
        cur = self.conn.cursor()
        cur.execute("SELECT COUNT(*) AS total FROM articulos")
        if cur.fetchone()["total"] > 0 or not os.path.exists(PRODUCTS_FILE):
            return
        with open(PRODUCTS_FILE, "r", encoding="utf-8") as f:
            products = json.load(f)
        now = now_text()
        for p in products:
            codigo = normalize_text(p.get("codigo"))
            descripcion = normalize_text(p.get("descripcion"))[:ARTICULO_DESC_MAX]
            unidad = normalize_text(p.get("unidad_medida")) or "UND"
            tipo = normalize_text(p.get("tipo_articulo")) or "CONSUMIBLE"
            if codigo and descripcion:
                cur.execute("""
                    INSERT OR IGNORE INTO articulos(codigo, descripcion, unidad_medida, tipo_articulo, stock_minimo, stock_minimo_aqp, stock_minimo_min, fecha_registro)
                    VALUES (?, ?, ?, ?, 0, 0, 0, ?)
                """, (codigo, descripcion, unidad, tipo, now))
        self.conn.commit()

    def get_articulo(self, codigo):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM articulos WHERE UPPER(codigo)=UPPER(?) AND activo=1", (normalize_text(codigo),))
        return cur.fetchone()

    def validar_cantidad_articulo(self, codigo, cantidad, contexto="cantidad"):
        """Valida cantidad contra la configuración del artículo.

        Si el artículo no permite decimales, bloquea cantidades como 1.5.
        Se usa en requerimientos y también queda disponible para otros flujos.
        """
        codigo = normalize_text(codigo)
        articulo = self.get_articulo(codigo)
        if not articulo:
            raise ValueError(f"El código {codigo} no existe.")
        try:
            raw_qty = D(cantidad)
            qty_d = q_qty(raw_qty)
        except Exception:
            raise ValueError(f"{contexto.capitalize()} inválida para {codigo}.")
        if raw_qty != qty_d:
            raise ValueError(f"La {contexto} de {codigo} admite como máximo 3 decimales.")
        qty = float(qty_d)
        if qty_d <= 0:
            raise ValueError(f"La {contexto} de {codigo} debe ser mayor a cero.")
        permite = 1
        try:
            permite = int(articulo["permite_decimales"] if "permite_decimales" in articulo.keys() else 1)
        except Exception:
            permite = 1
        if not permite and abs(qty - round(qty)) > 0.0000001:
            raise ValueError(
                f"El artículo {codigo} - {articulo['descripcion']} no permite cantidades decimales. "
                f"Ingrese un número entero."
            )
        return qty, articulo

    def all_articulos(self):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT codigo, descripcion, unidad_medida, tipo_articulo, stock_minimo, stock_minimo_aqp, stock_minimo_min, valoriza, permite_decimales, fecha_registro
            FROM articulos WHERE activo=1
            ORDER BY codigo
        """)
        return cur.fetchall()

    def search_articulos(self, term=""):
        term = f"%{normalize_text(term)}%"
        cur = self.conn.cursor()
        cur.execute("""
            WITH stock AS (
                SELECT d.codigo,
                       COALESCE(SUM(d.cantidad), 0) AS stock_total,
                       COALESCE(SUM(CASE WHEN m.almacen_codigo='AQP' THEN d.cantidad ELSE 0 END), 0) AS stock_aqp,
                       COALESCE(SUM(CASE WHEN m.almacen_codigo='MIN' THEN d.cantidad ELSE 0 END), 0) AS stock_min
                FROM movimiento_detalle d
                JOIN movimientos m ON m.id = d.movimiento_id
                GROUP BY d.codigo
            ), pendiente AS (
                SELECT d.codigo,
                       COALESCE(SUM(CASE WHEN r.almacen_codigo='AQP' THEN d.cantidad_requerida - d.cantidad_atendida ELSE 0 END), 0) AS comprometido_aqp,
                       COALESCE(SUM(CASE WHEN r.almacen_codigo='MIN' THEN d.cantidad_requerida - d.cantidad_atendida ELSE 0 END), 0) AS comprometido_min
                FROM requerimiento_detalle d
                JOIN requerimientos r ON r.id = d.requerimiento_id
                WHERE r.estado IN ('PENDIENTE','DESPACHO PARCIAL')
                GROUP BY d.codigo
            )
            SELECT a.codigo, a.descripcion, a.unidad_medida, a.tipo_articulo,
                   a.stock_minimo, a.stock_minimo_aqp, a.stock_minimo_min, a.valoriza, a.permite_decimales, a.fecha_registro,
                   COALESCE(s.stock_total, 0) AS stock,
                   COALESCE(s.stock_aqp, 0) AS stock_aqp,
                   COALESCE(s.stock_min, 0) AS stock_min,
                   COALESCE(p.comprometido_aqp, 0) AS comprometido_aqp,
                   COALESCE(p.comprometido_min, 0) AS comprometido_min,
                   (COALESCE(s.stock_aqp, 0) - COALESCE(p.comprometido_aqp, 0)) AS libre_aqp,
                   (COALESCE(s.stock_min, 0) - COALESCE(p.comprometido_min, 0)) AS libre_min,
                   (COALESCE(s.stock_total, 0) - COALESCE(p.comprometido_aqp, 0) - COALESCE(p.comprometido_min, 0)) AS libre_total,
                   CASE
                       WHEN a.stock_minimo_aqp > 0 AND COALESCE(s.stock_aqp, 0) <= a.stock_minimo_aqp THEN 'AQP BAJO'
                       WHEN a.stock_minimo_min > 0 AND COALESCE(s.stock_min, 0) <= a.stock_minimo_min THEN 'MIN BAJO'
                       WHEN a.stock_minimo > 0 AND COALESCE(s.stock_total, 0) <= a.stock_minimo THEN 'STOCK BAJO'
                       ELSE 'OK'
                   END AS estado_stock,
                   CASE WHEN a.stock_minimo_aqp > 0 AND COALESCE(s.stock_aqp, 0) <= a.stock_minimo_aqp THEN 'STOCK BAJO' ELSE 'OK' END AS estado_aqp,
                   CASE WHEN a.stock_minimo_min > 0 AND COALESCE(s.stock_min, 0) <= a.stock_minimo_min THEN 'STOCK BAJO' ELSE 'OK' END AS estado_min
            FROM articulos a
            LEFT JOIN stock s ON s.codigo = a.codigo
            LEFT JOIN pendiente p ON p.codigo = a.codigo
            WHERE a.activo=1 AND (UPPER(a.codigo) LIKE ? OR UPPER(a.descripcion) LIKE ?)
            ORDER BY a.descripcion, a.codigo
            LIMIT 500
        """, (term, term))
        return cur.fetchall()

    def _normalizar_articulo(self, codigo, descripcion, unidad, tipo, stock_minimo=0, stock_minimo_aqp=None, stock_minimo_min=None, valoriza=1, permite_decimales=1):
        codigo = normalize_text(codigo)
        descripcion = normalize_text(descripcion)
        unidad = normalize_text(unidad)
        tipo = normalize_text(tipo)
        try:
            stock_minimo = float(str(stock_minimo or 0).replace(",", "."))
            stock_minimo_aqp = stock_minimo if stock_minimo_aqp is None or stock_minimo_aqp == "" else float(str(stock_minimo_aqp or 0).replace(",", "."))
            stock_minimo_min = stock_minimo if stock_minimo_min is None or stock_minimo_min == "" else float(str(stock_minimo_min or 0).replace(",", "."))
        except Exception:
            raise ValueError("Stock mínimo inválido.")
        valoriza = 1 if bool(valoriza) else 0
        permite_decimales = 1 if bool(permite_decimales) else 0
        if stock_minimo > 0:
            if stock_minimo_aqp == 0:
                stock_minimo_aqp = stock_minimo
            if stock_minimo_min == 0:
                stock_minimo_min = stock_minimo
        if not codigo or not descripcion or not unidad or not tipo:
            raise ValueError("Complete todos los campos del artículo.")
        if len(descripcion) > ARTICULO_DESC_MAX:
            raise ValueError(f"La descripción no puede exceder {ARTICULO_DESC_MAX} caracteres.")
        if unidad not in UNIDADES_PERMITIDAS:
            raise ValueError("Unidad de medida no permitida.")
        if tipo not in TIPOS_ARTICULO:
            raise ValueError("Tipo de artículo no permitido.")
        return codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min, valoriza, permite_decimales

    def add_articulo(self, codigo, descripcion, unidad, tipo, stock_minimo=0, stock_minimo_aqp=None, stock_minimo_min=None, valoriza=1, permite_decimales=1):
        codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min, valoriza, permite_decimales = self._normalizar_articulo(codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min, valoriza, permite_decimales)
        cur = self.conn.cursor()
        cur.execute("SELECT codigo FROM articulos WHERE UPPER(codigo)=UPPER(?)", (codigo,))
        if cur.fetchone():
            raise ValueError("El código ya existe.")
        cur.execute("""
            INSERT INTO articulos(codigo, descripcion, unidad_medida, tipo_articulo, stock_minimo, stock_minimo_aqp, stock_minimo_min, valoriza, permite_decimales, fecha_registro)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min, valoriza, permite_decimales, now_text()))
        self.conn.commit()

    def add_articulos_bulk(self, rows, usuario=""):
        if not rows:
            raise ValueError("No hay artículos para cargar.")
        normalized = []
        vistos = set()
        for i, raw in enumerate(rows, start=2):
            codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min = self._normalizar_articulo(*raw)
            if codigo in vistos:
                raise ValueError(f"Código duplicado dentro del Excel en la fila {i}: {codigo}.")
            vistos.add(codigo)
            if self.get_articulo(codigo):
                raise ValueError(f"El código {codigo} ya existe. Corrija el Excel antes de cargar.")
            normalized.append((codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min))
        try:
            for codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min in normalized:
                self.conn.execute("""
                    INSERT INTO articulos(codigo, descripcion, unidad_medida, tipo_articulo, stock_minimo, stock_minimo_aqp, stock_minimo_min, fecha_registro)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min, now_text()))
                self.audit("articulos", codigo, "CREACION MASIVA", usuario, "Carga masiva de artículos", commit=False)
            self.conn.commit()
            return len(normalized)
        except Exception:
            self.conn.rollback()
            raise

    def update_articulo(self, codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp=None, stock_minimo_min=None, valoriza=1, permite_decimales=1):
        codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min, valoriza, permite_decimales = self._normalizar_articulo(codigo, descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min, valoriza, permite_decimales)
        cur = self.conn.cursor()
        cur.execute("""
            UPDATE articulos SET descripcion=?, unidad_medida=?, tipo_articulo=?, stock_minimo=?, stock_minimo_aqp=?, stock_minimo_min=?, valoriza=?, permite_decimales=?
            WHERE UPPER(codigo)=UPPER(?) AND activo=1
        """, (descripcion, unidad, tipo, stock_minimo, stock_minimo_aqp, stock_minimo_min, valoriza, permite_decimales, codigo))
        if cur.rowcount == 0:
            raise ValueError("No se encontró el artículo activo.")
        self.conn.commit()

    def next_vale(self, tipo_codigo=None, almacen_codigo=None, year=None):
        # V44: correlativo anual global. Ingresos/salidas comparten VALE;
        # los ajustes llevan su propia serie AJ. El almacen queda en campos
        # separados y ya no forma parte del numero documental.
        tipo_codigo = normalize_text(tipo_codigo or "")
        doc_type = "AJ" if tipo_codigo == "AJ" else "VALE"
        year = int(year or date.today().year)
        return AnnualCorrelativeService(self.db_path, self.conn).next(doc_type, year)

    def next_requerimiento(self, almacen_codigo="AQP", year=None):
        # El almacen se conserva como dato del requerimiento, no en el correlativo.
        year = int(year or date.today().year)
        return AnnualCorrelativeService(self.db_path, self.conn).next("REQ", year)

    def get_stock(self, codigo, almacen_codigo=None):
        cur = self.conn.cursor()
        if almacen_codigo:
            cur.execute("""
                SELECT COALESCE(SUM(d.cantidad),0) AS stock
                FROM movimiento_detalle d
                JOIN movimientos m ON m.id = d.movimiento_id
                WHERE d.codigo=? AND m.almacen_codigo=?
            """, (normalize_text(codigo), normalize_text(almacen_codigo)))
        else:
            cur.execute("SELECT COALESCE(SUM(cantidad),0) AS stock FROM movimiento_detalle WHERE codigo=?", (normalize_text(codigo),))
        return float(cur.fetchone()["stock"])

    def ultimo_costo_articulo(self, codigo, almacen_codigo=None):
        """Devuelve un costo referencial sin IGV en soles para autollenar salidas manuales.
        Prioriza el último ingreso/ajuste positivo valorizado del almacén; si no existe, busca globalmente.
        """
        codigo = normalize_text(codigo)
        almacen_codigo = normalize_text(almacen_codigo or "")
        cur = self.conn.cursor()
        def buscar(almacen=None):
            sql = """
                SELECT d.precio_unitario_sin_igv AS precio
                FROM movimiento_detalle d
                JOIN movimientos m ON m.id=d.movimiento_id
                WHERE UPPER(d.codigo)=UPPER(?)
                  AND d.cantidad > 0
                  AND COALESCE(d.precio_unitario_sin_igv,0) > 0
            """
            params=[codigo]
            if almacen:
                sql += " AND m.almacen_codigo=?"
                params.append(almacen)
            sql += " ORDER BY m.fecha_operacion DESC, m.id DESC, d.id DESC LIMIT 1"
            cur.execute(sql, params)
            r=cur.fetchone()
            return float(r["precio"] or 0) if r else 0.0
        precio = buscar(almacen_codigo) if almacen_codigo else 0.0
        if precio <= 0:
            precio = buscar(None)
        if precio <= 0:
            cur.execute("""
                SELECT precio_unitario_sin_igv AS precio
                FROM orden_compra_detalle
                WHERE UPPER(codigo)=UPPER(?) AND COALESCE(precio_unitario_sin_igv,0)>0
                ORDER BY id DESC LIMIT 1
            """, (codigo,))
            r=cur.fetchone(); precio=float(r["precio"] or 0) if r else 0.0
        return round(precio, 6)

    def duplicate_exists(self, fecha_iso, documento, ot, tipo_codigo, almacen_codigo=None):
        cur = self.conn.cursor()
        if almacen_codigo:
            cur.execute("""
                SELECT vale FROM movimientos
                WHERE fecha_operacion=? AND UPPER(documento)=UPPER(?) AND IFNULL(UPPER(ot),'')=IFNULL(UPPER(?),'') AND tipo_codigo=? AND almacen_codigo=?
                LIMIT 1
            """, (fecha_iso, normalize_text(documento), ot, tipo_codigo, normalize_text(almacen_codigo)))
        else:
            cur.execute("""
                SELECT vale FROM movimientos
                WHERE fecha_operacion=? AND UPPER(documento)=UPPER(?) AND IFNULL(UPPER(ot),'')=IFNULL(UPPER(?),'') AND tipo_codigo=?
                LIMIT 1
            """, (fecha_iso, normalize_text(documento), ot, tipo_codigo))
        return cur.fetchone()

    def save_movimiento(self, fecha_texto, tipo_codigo, documento, ot_num, cc_codigo, solicitante, observacion, detalles, usuario, almacen_codigo="AQP", requerimiento="", commit=True):
        fecha_iso = fecha_iso_from_text(fecha_texto)
        fecha_visible = fecha_text_from_iso(fecha_iso)
        tipo_codigo = normalize_text(tipo_codigo)
        documento = normalize_text(documento)
        cc_codigo = normalize_text(cc_codigo)
        solicitante = normalize_text(solicitante)
        observacion = normalize_text(observacion)
        almacen_codigo = normalize_text(almacen_codigo or "AQP")
        requerimiento = normalize_text(requerimiento)
        ot = clean_ot(ot_num)
        if tipo_codigo not in TIPOS_MOVIMIENTO:
            raise ValueError("Tipo inválido. Use CN para ingreso, CO para salida o AJ para ajuste.")
        if not documento:
            raise ValueError("Ingrese documento.")
        if not cc_codigo or cc_codigo not in CENTROS_COSTO:
            raise ValueError("Centro de costo inválido.")
        if almacen_codigo not in ALMACENES:
            raise ValueError("Almacén inválido.")
        if not solicitante:
            raise ValueError("Ingrese el nombre del solicitante.")
        if not observacion:
            raise ValueError("Ingrese una observación para mantener trazabilidad.")
        if not detalles:
            raise ValueError("Debe ingresar al menos un artículo.")
        dup = self.duplicate_exists(fecha_iso, documento, ot, tipo_codigo, almacen_codigo)
        if dup:
            raise ValueError(f"Documento duplicado. Ya existe un vale con la misma fecha, documento, OT y tipo: {dup['vale']}")
        signed = []
        stock_proyectado = {}
        for d in detalles:
            codigo = normalize_text(d.get("codigo"))
            try:
                raw_cantidad = D(d.get("cantidad"))
                cantidad_d = q_qty(raw_cantidad)
                if raw_cantidad != cantidad_d:
                    raise ValueError(f"La cantidad de {codigo} admite como máximo 3 decimales.")
                cantidad = float(cantidad_d)
            except ValueError:
                raise
            except Exception:
                raise ValueError(f"Cantidad inválida para {codigo}.")
            articulo = self.get_articulo(codigo)
            if not articulo:
                raise ValueError(f"El código {codigo} no existe.")
            if tipo_codigo in ["CN", "CO"] and cantidad <= 0:
                raise ValueError(f"La cantidad del código {codigo} debe ser mayor a cero.")
            if tipo_codigo == "AJ" and cantidad == 0:
                raise ValueError(f"El ajuste del código {codigo} no puede ser cero.")
            cantidad_guardar = cantidad
            if tipo_codigo == "CO":
                cantidad_guardar = -abs(cantidad)
            elif tipo_codigo == "CN":
                cantidad_guardar = abs(cantidad)
            try:
                raw_precio = D(d.get("precio_unitario_sin_igv", d.get("precio", 0)) or 0)
                precio_d = q_unit(raw_precio)
                precio = float(precio_d)
            except ValueError:
                raise
            except Exception:
                raise ValueError(f"Precio inválido para {codigo}.")
            if precio_d < 0:
                raise ValueError(f"El precio del código {codigo} no puede ser negativo.")
            subtotal = float(q_money(abs(D(cantidad_guardar)) * precio_d))

            if codigo not in stock_proyectado:
                stock_proyectado[codigo] = self.get_stock(codigo, almacen_codigo)
            stock_proyectado[codigo] += cantidad_guardar
            if stock_proyectado[codigo] < -0.00001:
                stock_actual = self.get_stock(codigo, almacen_codigo)
                total_salida = stock_actual - stock_proyectado[codigo]
                raise ValueError(
                    f"Stock insuficiente para {codigo}. Stock actual: {stock_actual:g}. "
                    f"Salida/ajuste acumulado en este documento: {total_salida:g}."
                )
            signed.append((d.get("item", len(signed) + 1), codigo, articulo["descripcion"], articulo["unidad_medida"], cantidad_guardar, precio, subtotal))
        vale = self.next_vale(tipo_codigo, almacen_codigo, int(fecha_iso[:4]))
        cur = self.conn.cursor()
        cur.execute("""
            INSERT INTO movimientos(vale, fecha, fecha_operacion, tipo_codigo, tipo_nombre, documento, ot, centro_costo_codigo, centro_costo_nombre, solicitante, observacion, almacen_codigo, almacen_nombre, requerimiento, usuario, creado_en)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (vale, fecha_visible, fecha_iso, tipo_codigo, TIPOS_MOVIMIENTO[tipo_codigo], documento, ot, cc_codigo, CENTROS_COSTO[cc_codigo], solicitante, observacion, almacen_codigo, ALMACENES[almacen_codigo], requerimiento, usuario, now_text()))
        mov_id = cur.lastrowid
        for item, codigo, desc, um, cantidad, precio, subtotal in signed:
            cur.execute("""
                INSERT INTO movimiento_detalle(movimiento_id, item, codigo, descripcion, unidad_medida, cantidad, precio_unitario_sin_igv, subtotal_sin_igv)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (mov_id, item, codigo, desc, um, cantidad, precio, subtotal))
        self.audit("movimientos", vale, "CREACION", usuario, f"Registro {TIPOS_MOVIMIENTO[tipo_codigo]} - Documento {documento}", commit=False)
        if commit:
            self.conn.commit()
        return vale

    def next_transferencia(self, year=None):
        year = int(year or date.today().year)
        return AnnualCorrelativeService(self.db_path, self.conn).next("TRF", year)

    def transferir_stock(self, fecha_texto, almacen_origen, almacen_destino, cc_codigo, solicitante, observacion, detalles, usuario, numero_guia=""):
        fecha_iso = fecha_iso_from_text(fecha_texto)
        fecha_visible = fecha_text_from_iso(fecha_iso)
        almacen_origen = normalize_text(almacen_origen)
        almacen_destino = normalize_text(almacen_destino)
        cc_codigo = normalize_text(cc_codigo)
        solicitante = normalize_text(solicitante)
        observacion = normalize_text(observacion)
        numero_guia = normalize_text(numero_guia)
        if not numero_guia:
            raise ValueError("Ingrese número de guía para la transferencia.")
        if almacen_origen not in ALMACENES or almacen_destino not in ALMACENES:
            raise ValueError("Seleccione almacenes válidos.")
        if almacen_origen == almacen_destino:
            raise ValueError("El almacén origen y destino deben ser diferentes.")
        if cc_codigo not in CENTROS_COSTO:
            raise ValueError("Centro de costo inválido.")
        if not solicitante:
            raise ValueError("Ingrese solicitante/responsable de la transferencia.")
        if not observacion:
            raise ValueError("Ingrese una observación para la transferencia.")
        transferencia = self.next_transferencia(int(fecha_iso[:4]))
        try:
            obs = f"TRANSFERENCIA {transferencia}: {observacion}"
            vale_salida = self.save_movimiento(fecha_texto, "CO", f"{transferencia}-SAL", "", cc_codigo, solicitante, obs, detalles, usuario, almacen_origen, transferencia, commit=False)
            vale_ingreso = self.save_movimiento(fecha_texto, "CN", f"{transferencia}-ING", "", cc_codigo, solicitante, obs, detalles, usuario, almacen_destino, transferencia, commit=False)
            self.conn.execute("""
                INSERT INTO transferencias(transferencia, fecha, fecha_operacion, almacen_origen, almacen_destino, vale_salida, vale_ingreso, solicitante, centro_costo_codigo, centro_costo_nombre, observacion, numero_guia, usuario, creado_en)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (transferencia, fecha_visible, fecha_iso, almacen_origen, almacen_destino, vale_salida, vale_ingreso, solicitante, cc_codigo, CENTROS_COSTO[cc_codigo], observacion, numero_guia, usuario, now_text()))
            self.audit("transferencias", transferencia, "CREACION", usuario, f"{almacen_origen} -> {almacen_destino} | Salida {vale_salida} | Ingreso {vale_ingreso}", commit=False)
            self.conn.commit()
            return transferencia, vale_salida, vale_ingreso
        except Exception:
            self.conn.rollback()
            raise

    def reverse_movimiento(self, vale_original, motivo, usuario):
        motivo = normalize_text(motivo)
        if not motivo:
            raise ValueError("Ingrese el motivo del reverso.")
        mov, det = self.get_movimiento_by_vale(vale_original)
        if not mov:
            raise ValueError("No se encontró el vale a reversar.")
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM movimiento_reversos WHERE vale_original=? OR vale_reverso=?", (mov["vale"], mov["vale"]))
        if cur.fetchone():
            raise ValueError("Este vale ya fue reversado o corresponde a un reverso. No se puede reversar nuevamente.")
        if mov["tipo_codigo"] == "CN":
            tipo_reverso = "CO"
            detalles_rev = [{"item": r["item"], "codigo": r["codigo"], "cantidad": abs(float(r["cantidad"]))} for r in det]
        elif mov["tipo_codigo"] == "CO":
            tipo_reverso = "CN"
            detalles_rev = [{"item": r["item"], "codigo": r["codigo"], "cantidad": abs(float(r["cantidad"]))} for r in det]
        elif mov["tipo_codigo"] == "AJ":
            tipo_reverso = "AJ"
            detalles_rev = [{"item": r["item"], "codigo": r["codigo"], "cantidad": -float(r["cantidad"])} for r in det]
        else:
            raise ValueError("Tipo de movimiento no soportado para reverso.")
        try:
            documento = f"REV-{mov['vale']}-{datetime.now().strftime('%H%M%S')}"
            observacion = f"REVERSO DE {mov['vale']}: {motivo}"
            vale_rev = self.save_movimiento(
                today_text(), tipo_reverso, documento, (mov["ot"] or "").replace("OT-", ""),
                mov["centro_costo_codigo"], mov["solicitante"], observacion, detalles_rev, usuario,
                mov["almacen_codigo"], f"REVERSO:{mov['vale']}", commit=False
            )

            # Si el vale original estaba vinculado a una OC o requerimiento, el reverso también
            # debe devolver el documento de origen a su saldo pendiente correcto.
            origen = normalize_text(mov["requerimiento"] or "")
            if origen.startswith("OC-"):
                oc, oc_det = self.get_orden_compra(origen)
                if oc:
                    cur2 = self.conn.cursor()
                    for r in det:
                        cantidad_rev = abs(float(r["cantidad"] or 0))
                        codigo_rev = normalize_text(r["codigo"])
                        pendientes = [d for d in oc_det if normalize_text(d["codigo"]) == codigo_rev and float(d["cantidad_ingresada"] or 0) > 0]
                        restante = cantidad_rev
                        for d in pendientes:
                            if restante <= 0:
                                break
                            actual_ing = float(d["cantidad_ingresada"] or 0)
                            bajar = min(actual_ing, restante)
                            cur2.execute("UPDATE orden_compra_detalle SET cantidad_ingresada = cantidad_ingresada - ? WHERE id=?", (bajar, d["id"]))
                            restante -= bajar
                    cur2.execute("SELECT COALESCE(SUM(cantidad_solicitada),0) sol, COALESCE(SUM(cantidad_aprobada),0) aprob, COALESCE(SUM(cantidad_ingresada),0) ing FROM orden_compra_detalle WHERE oc_id=?", (oc["id"],))
                    sums = cur2.fetchone()
                    aprobado = float(sums["aprob"] or 0)
                    ingresado = float(sums["ing"] or 0)
                    solicitado = float(sums["sol"] or 0)
                    cur2.execute("SELECT COALESCE(SUM(monto),0) pagado FROM orden_compra_pagos WHERE oc_id=?", (oc["id"],))
                    pagado = float(cur2.fetchone()["pagado"] or 0)
                    cur2.execute("SELECT COUNT(*) c FROM orden_compra_transito WHERE oc_id=?", (oc["id"],))
                    tiene_transito = int(cur2.fetchone()["c"] or 0) > 0
                    if ingresado > 0:
                        estado_oc = "ATENDIDA" if aprobado > 0 and abs(ingresado - aprobado) <= 0.00001 else "ATENDIDA PARCIALMENTE"
                    elif tiene_transito:
                        estado_oc = "EN TRANSITO"
                    elif pagado > 0:
                        estado_oc = "PAGADA" if pagado + 0.01 >= float(oc["total"] or 0) else "PAGO PARCIAL"
                    else:
                        estado_oc = "APROBADA" if aprobado > 0 and abs(aprobado - solicitado) <= 0.00001 else "APROBADA PARCIAL"
                    self._set_oc_state(oc["id"], estado_oc, usuario, f"Reverso de vale {mov['vale']} ajustó cantidad ingresada")
            elif origen.startswith("REQ-"):
                req, req_det = self.find_requerimiento(origen)
                if req:
                    cur2 = self.conn.cursor()
                    for r in det:
                        cantidad_rev = abs(float(r["cantidad"] or 0))
                        codigo_rev = normalize_text(r["codigo"])
                        pendientes = [d for d in req_det if normalize_text(d["codigo"]) == codigo_rev and float(d["cantidad_atendida"] or 0) > 0]
                        restante = cantidad_rev
                        for d in pendientes:
                            if restante <= 0:
                                break
                            actual_at = float(d["cantidad_atendida"] or 0)
                            bajar = min(actual_at, restante)
                            cur2.execute("UPDATE requerimiento_detalle SET cantidad_atendida = cantidad_atendida - ? WHERE id=?", (bajar, d["id"]))
                            restante -= bajar
                    cur2.execute("SELECT COALESCE(SUM(cantidad_requerida),0) req, COALESCE(SUM(cantidad_atendida),0) at FROM requerimiento_detalle WHERE requerimiento_id=?", (req["id"],))
                    sums = cur2.fetchone()
                    total_req = float(sums["req"] or 0)
                    total_at = float(sums["at"] or 0)
                    estado_req = "PENDIENTE" if total_at <= 0.00001 else ("ATENDIDO" if abs(total_at - total_req) <= 0.00001 else "DESPACHO PARCIAL")
                    cur2.execute("UPDATE requerimientos SET estado=?, modificado_en=? WHERE id=?", (estado_req, now_text(), req["id"]))
                    cur2.execute("INSERT INTO requerimiento_atenciones(requerimiento_id, vale, fecha_atencion, usuario, estado_atencion, observacion) VALUES (?, ?, ?, ?, ?, ?)", (req["id"], vale_rev, fecha_iso_from_text(today_text()), usuario, estado_req, f"REVERSO DE {mov['vale']}: {motivo}"))
            self.conn.execute("""
                INSERT INTO movimiento_reversos(vale_original, vale_reverso, motivo, usuario, creado_en)
                VALUES (?, ?, ?, ?, ?)
            """, (mov["vale"], vale_rev, motivo, usuario, now_text()))
            self.audit("movimientos", mov["vale"], "REVERSO", usuario, f"Vale reverso {vale_rev}. Motivo: {motivo}", commit=False)
            self.conn.commit()
            return vale_rev
        except Exception:
            self.conn.rollback()
            raise

    def get_movimiento_by_vale(self, vale):
        q = str(vale or "").strip()
        if not q:
            return None, []
        candidates = [q]
        dg = only_digits(q)
        if dg:
            candidates.extend([dg, f"{int(dg):06d}"])
        cur = self.conn.cursor()
        mov = None
        for c in dict.fromkeys(candidates):
            cur.execute("SELECT * FROM movimientos WHERE UPPER(vale)=UPPER(?)", (c,))
            mov = cur.fetchone()
            if mov:
                break
        if not mov:
            search_terms = [q.upper()]
            if dg:
                search_terms.append(f"{int(dg):06d}")
            for term in search_terms:
                like = f"%{term}%"
                cur.execute("""
                    SELECT * FROM movimientos
                    WHERE UPPER(vale) LIKE ? OR UPPER(documento) LIKE ? OR UPPER(IFNULL(ot,'')) LIKE ?
                    ORDER BY id DESC LIMIT 1
                """, (like, like, like))
                mov = cur.fetchone()
                if mov:
                    break
        if not mov:
            return None, []
        cur.execute("SELECT * FROM movimiento_detalle WHERE movimiento_id=? ORDER BY item", (mov["id"],))
        return mov, cur.fetchall()

    def list_vales(self, term="", almacen_codigo="", tipo_codigo=""):
        """Lista vales para cuadros de ayuda/consulta.
        No depende de datos demo; muestra hasta 500 vales reales registrados.
        Acepta búsquedas por vale, documento, OT, centro de costo, solicitante, usuario o requerimiento.
        """
        q = normalize_text(term)
        params = []
        where = []
        if q:
            like = f"%{q}%"
            where.append("(" + " OR ".join([
                "UPPER(IFNULL(vale,'')) LIKE ?",
                "UPPER(IFNULL(documento,'')) LIKE ?",
                "UPPER(IFNULL(ot,'')) LIKE ?",
                "UPPER(IFNULL(centro_costo_codigo,'')) LIKE ?",
                "UPPER(IFNULL(centro_costo_nombre,'')) LIKE ?",
                "UPPER(IFNULL(solicitante,'')) LIKE ?",
                "UPPER(IFNULL(usuario,'')) LIKE ?",
                "UPPER(IFNULL(requerimiento,'')) LIKE ?"
            ]) + ")")
            params.extend([like] * 8)
        alm = normalize_text(almacen_codigo)
        if alm and alm != "TODOS":
            where.append("UPPER(IFNULL(almacen_codigo,'')) = ?")
            params.append(alm)
        tipo = normalize_text(tipo_codigo)
        if tipo and tipo != "TODOS":
            where.append("UPPER(IFNULL(tipo_codigo,'')) = ?")
            params.append(tipo)
        sql = """
            SELECT vale, fecha, fecha_operacion, tipo_codigo, tipo_nombre, documento, ot,
                   centro_costo_codigo, centro_costo_nombre, solicitante, almacen_codigo,
                   requerimiento, usuario, creado_en
            FROM movimientos
        """
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY id DESC LIMIT 500"
        cur = self.conn.cursor()
        cur.execute(sql, params)
        return cur.fetchall()

    def update_movimiento_header(self, vale, ot_num, cc_codigo, solicitante, observacion, usuario):
        cc_codigo = normalize_text(cc_codigo)
        solicitante = normalize_text(solicitante)
        observacion = normalize_text(observacion)
        if cc_codigo not in CENTROS_COSTO:
            raise ValueError("Centro de costo inválido.")
        if not solicitante:
            raise ValueError("El solicitante no puede quedar vacío.")
        if not observacion:
            raise ValueError("La observación no puede quedar vacía.")
        mov, _ = self.get_movimiento_by_vale(vale)
        if not mov:
            raise ValueError("No se encontró el vale.")
        ot = clean_ot(ot_num)
        old = f"OT={mov['ot']} | CC={mov['centro_costo_codigo']} | SOL={mov['solicitante']} | OBS={mov['observacion']}"
        new = f"OT={ot} | CC={cc_codigo} | SOL={solicitante} | OBS={observacion}"
        self.conn.execute("""
            UPDATE movimientos SET ot=?, centro_costo_codigo=?, centro_costo_nombre=?, solicitante=?, observacion=?, modificado_en=?
            WHERE id=?
        """, (ot, cc_codigo, CENTROS_COSTO[cc_codigo], solicitante, observacion, now_text(), mov["id"]))
        self.conn.commit()
        self.audit("movimientos", mov["vale"], "MODIFICACION", usuario, f"Antes: {old} / Después: {new}")

    def report_month(self):
        first = date.today().replace(day=1)
        if first.month == 12:
            next_month = date(first.year + 1, 1, 1)
        else:
            next_month = date(first.year, first.month + 1, 1)
        cur = self.conn.cursor()
        cur.execute("""
            SELECT m.vale, m.fecha, m.fecha_operacion, m.tipo_codigo, m.tipo_nombre, m.documento, m.ot,
                   m.centro_costo_codigo, m.centro_costo_nombre, m.solicitante, m.observacion, m.almacen_codigo, m.almacen_nombre, m.requerimiento, m.usuario, m.creado_en,
                   d.item, d.codigo, d.descripcion, d.unidad_medida, d.cantidad
            FROM movimientos m
            JOIN movimiento_detalle d ON d.movimiento_id = m.id
            WHERE m.fecha_operacion >= ? AND m.fecha_operacion < ?
            ORDER BY m.fecha_operacion, m.id, d.item
        """, (first.strftime(ISO_FMT), next_month.strftime(ISO_FMT)))
        return cur.fetchall()


    def reporte_valorizacion_mensual(self, anio, mes, almacen_codigo="TODOS"):
        """Reporte mensual valorizado por promedio ponderado móvil.

        La valorización se calcula solo para artículos marcados como valorizables.
        Entradas/AJ positivos toman el precio registrado sin IGV. Las salidas se valorizan
        al costo promedio móvil disponible antes de la salida. Las transferencias intentan
        trasladar el costo promedio del almacén origen al ingreso del almacén destino.
        """
        try:
            anio = int(anio)
            mes = int(mes)
            first = date(anio, mes, 1)
        except Exception:
            raise ValueError("Seleccione un año y mes válidos.")
        if mes == 12:
            next_month = date(anio + 1, 1, 1)
        else:
            next_month = date(anio, mes + 1, 1)
        almacen_codigo = normalize_text(almacen_codigo or "TODOS")
        if almacen_codigo not in ("TODOS", "AQP", "MIN"):
            raise ValueError("Almacén inválido para el reporte.")

        # Se procesan todos los almacenes para que las transferencias puedan trasladar
        # su costo desde el almacén origen al destino. El filtro de almacén se aplica
        # solo a las filas finales del reporte.
        params = [next_month.strftime(ISO_FMT)]
        almacen_filter = ""
        cur = self.conn.cursor()
        cur.execute(f"""
            SELECT m.id AS movimiento_id, m.vale, m.fecha_operacion, m.tipo_codigo, m.documento,
                   m.almacen_codigo, m.almacen_nombre, m.requerimiento, m.centro_costo_codigo, m.ot,
                   d.id AS detalle_id, d.item, d.codigo, d.descripcion, d.unidad_medida,
                   d.cantidad, COALESCE(d.precio_unitario_sin_igv,0) AS precio_unitario_sin_igv,
                   a.valoriza, a.tipo_articulo
            FROM movimientos m
            JOIN movimiento_detalle d ON d.movimiento_id = m.id
            JOIN articulos a ON UPPER(a.codigo)=UPPER(d.codigo)
            WHERE m.fecha_operacion < ?
              AND a.activo=1
              AND COALESCE(a.valoriza,1)=1
              {almacen_filter}
            ORDER BY m.fecha_operacion, m.id, d.item, d.id
        """, params)
        rows = cur.fetchall()

        periodo = f"{anio:04d}-{mes:02d}"
        state = {}  # (almacen,codigo) -> {qty,value,last_price,meta}
        summaries = {}
        transfer_unit_cost = {}  # (transferencia,codigo) -> costo unitario promedio del origen
        warnings = {}

        def key_for(r):
            return (r["almacen_codigo"], normalize_text(r["codigo"]))

        def include_key(k):
            return almacen_codigo == "TODOS" or k[0] == almacen_codigo

        def ensure_state(r):
            k = key_for(r)
            if k not in state:
                state[k] = {
                    "qty": 0.0,
                    "value": 0.0,
                    "last_price": 0.0,
                    "codigo": normalize_text(r["codigo"]),
                    "descripcion": r["descripcion"],
                    "unidad_medida": r["unidad_medida"],
                    "almacen_codigo": r["almacen_codigo"],
                    "almacen_nombre": r["almacen_nombre"],
                }
            return state[k]

        def ensure_summary_from_state(k):
            st = state[k]
            if k not in summaries:
                summaries[k] = {
                    "periodo": periodo,
                    "almacen_codigo": st["almacen_codigo"],
                    "almacen_nombre": st["almacen_nombre"],
                    "codigo": st["codigo"],
                    "descripcion": st["descripcion"],
                    "unidad_medida": st["unidad_medida"],
                    "stock_inicial": st["qty"],
                    "valor_inicial": st["value"],
                    "entradas": 0.0,
                    "valor_entradas": 0.0,
                    "salidas": 0.0,
                    "valor_salidas": 0.0,
                    "ajustes_entrada": 0.0,
                    "valor_ajustes_entrada": 0.0,
                    "ajustes_salida": 0.0,
                    "valor_ajustes_salida": 0.0,
                    "stock_final": 0.0,
                    "costo_promedio_final": 0.0,
                    "valor_final": 0.0,
                    "precio_prom_sin_igv_soles": 0.0,
                    "precio_prom_con_igv_soles": 0.0,
                    "precio_prom_sin_igv_dolares": 0.0,
                    "precio_prom_con_igv_dolares": 0.0,
                    "destino_centro_costo": "",
                    "destino_ot": "",
                    "metodo": "PROMEDIO PONDERADO MOVIL",
                    "observacion": "",
                    "_cc_set": set(),
                    "_ot_set": set(),
                    "_tc_ref": 0.0,
                }
            return summaries[k]

        def avg_cost(st):
            qty = float(st["qty"] or 0)
            val = float(st["value"] or 0)
            if qty > 0.00001:
                return val / qty
            return float(st.get("last_price") or 0)

        def note(k, txt):
            if not txt:
                return
            warnings.setdefault(k, set()).add(txt)

        def apply_row(r, in_period=False):
            k = key_for(r)
            st = ensure_state(r)
            qty_signed = float(r["cantidad"] or 0)
            qty_abs = abs(qty_signed)
            precio = float(r["precio_unitario_sin_igv"] or 0)
            tipo = normalize_text(r["tipo_codigo"])
            req_doc = normalize_text(r["requerimiento"] or "")
            avg = avg_cost(st)
            summ = ensure_summary_from_state(k) if (in_period and include_key(k)) else None
            if summ is not None:
                if r["centro_costo_codigo"]:
                    summ["_cc_set"].add(str(r["centro_costo_codigo"]))
                if r["ot"]:
                    summ["_ot_set"].add(str(r["ot"]))

            if qty_signed > 0:
                unit_cost = precio
                if unit_cost <= 0 and req_doc.startswith(("TR-", "TRF-")):
                    unit_cost = transfer_unit_cost.get((req_doc, normalize_text(r["codigo"])), 0.0)
                if unit_cost <= 0 and avg > 0:
                    unit_cost = avg
                if unit_cost <= 0:
                    note(k, "Entrada sin precio valorizable")
                mov_value = qty_abs * unit_cost
                st["qty"] += qty_abs
                st["value"] += mov_value
                if unit_cost > 0:
                    st["last_price"] = unit_cost
                if summ is not None:
                    if tipo == "AJ":
                        summ["ajustes_entrada"] += qty_abs
                        summ["valor_ajustes_entrada"] += mov_value
                    else:
                        summ["entradas"] += qty_abs
                        summ["valor_entradas"] += mov_value
            elif qty_signed < 0:
                unit_cost = avg
                if unit_cost <= 0:
                    note(k, "Salida sin costo promedio previo")
                mov_value = qty_abs * unit_cost
                if req_doc.startswith(("TR-", "TRF-")):
                    transfer_unit_cost[(req_doc, normalize_text(r["codigo"]))] = unit_cost
                st["qty"] -= qty_abs
                st["value"] -= mov_value
                if abs(st["qty"]) <= 0.00001:
                    st["qty"] = 0.0
                    st["value"] = 0.0
                if st["value"] < -0.01:
                    note(k, "Valor negativo: revisar stock/precios históricos")
                if summ is not None:
                    if tipo == "AJ":
                        summ["ajustes_salida"] += qty_abs
                        summ["valor_ajustes_salida"] += mov_value
                    else:
                        summ["salidas"] += qty_abs
                        summ["valor_salidas"] += mov_value

        # 1) Estado inicial: movimientos anteriores al mes.
        first_iso = first.strftime(ISO_FMT)
        next_iso = next_month.strftime(ISO_FMT)
        for r in rows:
            if str(r["fecha_operacion"]) < first_iso:
                apply_row(r, in_period=False)

        # 2) Crea resumen para artículos con saldo inicial, aunque no tengan movimiento mensual.
        for k in list(state.keys()):
            st = state[k]
            if include_key(k) and (abs(st["qty"]) > 0.00001 or abs(st["value"]) > 0.01):
                ensure_summary_from_state(k)

        # 3) Movimientos del mes.
        for r in rows:
            f = str(r["fecha_operacion"])
            if first_iso <= f < next_iso:
                apply_row(r, in_period=True)

        # 4) Saldos finales.
        result = []
        for k, summ in summaries.items():
            st = state.get(k, {})
            qty_final = float(st.get("qty") or 0)
            val_final = float(st.get("value") or 0)
            if abs(qty_final) <= 0.00001:
                qty_final = 0.0
                val_final = 0.0
            summ["stock_final"] = qty_final
            summ["valor_final"] = val_final
            summ["costo_promedio_final"] = (val_final / qty_final) if qty_final > 0.00001 else 0.0
            summ["precio_prom_sin_igv_soles"] = summ["costo_promedio_final"]
            summ["precio_prom_con_igv_soles"] = summ["costo_promedio_final"] * (1 + IGV_RATE)
            # Referencia en USD usando el último TC venta disponible del mes; si no existe, queda 0 para evitar datos inventados.
            cur.execute("SELECT venta FROM tipo_cambio WHERE moneda='DOLARES' AND fecha < ? ORDER BY fecha DESC LIMIT 1", (next_iso,))
            tcrow = cur.fetchone(); tc_ref = float(tcrow["venta"] or 0) if tcrow else 0.0
            summ["precio_prom_sin_igv_dolares"] = (summ["precio_prom_sin_igv_soles"] / tc_ref) if tc_ref > 0 else 0.0
            summ["precio_prom_con_igv_dolares"] = (summ["precio_prom_con_igv_soles"] / tc_ref) if tc_ref > 0 else 0.0
            summ["destino_centro_costo"] = ", ".join(sorted(summ.get("_cc_set", set())))
            summ["destino_ot"] = ", ".join(sorted(summ.get("_ot_set", set())))
            obs = sorted(warnings.get(k, set()))
            summ["observacion"] = "; ".join(obs)
            # Redondeo solo para salida; internamente el cálculo mantiene precisión.
            for campo in ["stock_inicial", "valor_inicial", "entradas", "valor_entradas", "salidas", "valor_salidas", "ajustes_entrada", "valor_ajustes_entrada", "ajustes_salida", "valor_ajustes_salida", "stock_final", "costo_promedio_final", "valor_final", "precio_prom_sin_igv_soles", "precio_prom_con_igv_soles", "precio_prom_sin_igv_dolares", "precio_prom_con_igv_dolares"]:
                summ[campo] = round(float(summ[campo] or 0), 4 if "stock" in campo or campo in ("entradas", "salidas", "ajustes_entrada", "ajustes_salida") else 2)
            if any(abs(float(summ[c] or 0)) > 0.00001 for c in ["stock_inicial", "valor_inicial", "entradas", "valor_entradas", "salidas", "valor_salidas", "ajustes_entrada", "valor_ajustes_entrada", "ajustes_salida", "valor_ajustes_salida", "stock_final", "valor_final"]):
                summ.pop("_cc_set", None); summ.pop("_ot_set", None); summ.pop("_tc_ref", None)
                result.append(summ)
        result.sort(key=lambda x: (x["almacen_codigo"], x["codigo"]))
        return result

    def stock_rows(self):
        return self.search_articulos("")


    def create_requerimiento(self, fecha_texto, almacen_codigo, solicitante, cc_codigo, ot_num, observacion, detalles, usuario, fecha_entrega_solicitada=""):
        fecha_iso = fecha_iso_from_text(fecha_texto)
        fecha_visible = fecha_text_from_iso(fecha_iso)
        almacen_codigo = normalize_text(almacen_codigo or "AQP")
        solicitante = normalize_text(solicitante)
        cc_codigo = normalize_text(cc_codigo)
        observacion = normalize_text(observacion)
        ot = clean_ot(ot_num)
        if almacen_codigo not in ALMACENES:
            raise ValueError("Seleccione un almacén válido: AQP o MIN.")
        if not solicitante:
            raise ValueError("Ingrese solicitante.")
        if cc_codigo not in CENTROS_COSTO:
            raise ValueError("Centro de costo inválido.")
        if not observacion:
            raise ValueError("Ingrese una observación.")
        if not detalles:
            raise ValueError("Debe ingresar al menos un artículo al requerimiento.")
        fecha_entrega_iso = fecha_iso_consulta_from_text(fecha_entrega_solicitada) if str(fecha_entrega_solicitada or "").strip() else ""
        clean = []
        seen = {}
        for d in detalles:
            codigo = normalize_text(d.get("codigo"))
            qty, articulo = self.validar_cantidad_articulo(codigo, d.get("cantidad"), "cantidad requerida")
            if codigo in seen:
                nueva_qty = seen[codigo]["cantidad"] + qty
                self.validar_cantidad_articulo(codigo, nueva_qty, "cantidad requerida acumulada")
                seen[codigo]["cantidad"] = nueva_qty
            else:
                seen[codigo] = {"codigo": codigo, "descripcion": articulo["descripcion"], "um": articulo["unidad_medida"], "cantidad": qty, "observacion": normalize_text(d.get("observacion", ""))}
        clean = list(seen.values())
        req = self.next_requerimiento(almacen_codigo, int(fecha_iso[:4]))
        cur = self.conn.cursor()
        cur.execute("""
            INSERT INTO requerimientos(requerimiento, fecha, fecha_requerimiento, fecha_entrega_solicitada, almacen_codigo, almacen_nombre, solicitante, centro_costo_codigo, centro_costo_nombre, ot, observacion, estado, usuario_creador, creado_en)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDIENTE', ?, ?)
        """, (req, fecha_visible, fecha_iso, fecha_entrega_iso, almacen_codigo, ALMACENES[almacen_codigo], solicitante, cc_codigo, CENTROS_COSTO[cc_codigo], ot, observacion, usuario, now_text()))
        req_id = cur.lastrowid
        for i, d in enumerate(clean, start=1):
            cur.execute("""
                INSERT INTO requerimiento_detalle(requerimiento_id, item, codigo, descripcion, unidad_medida, cantidad_requerida, cantidad_atendida, observacion)
                VALUES (?, ?, ?, ?, ?, ?, 0, ?)
            """, (req_id, i, d["codigo"], d["descripcion"], d["um"], d["cantidad"], d.get("observacion", "")))
        self.conn.commit()
        self.audit("requerimientos", req, "CREACION", usuario, "Registro de requerimiento")
        return req

    def find_requerimiento(self, value):
        q = str(value or "").strip()
        if not q:
            return None, []
        candidates = [q.upper()]
        dg = only_digits(q)
        if dg:
            candidates.extend([
                f"REQ-{date.today().year}-{int(dg):06d}",
                f"REQ-AQP-{int(dg):06d}",
                f"REQ-MIN-{int(dg):06d}",
                f"%{int(dg):06d}%",
            ])
        cur = self.conn.cursor()
        req = None
        for c in dict.fromkeys(candidates):
            if "%" in c:
                cur.execute("SELECT * FROM requerimientos WHERE UPPER(requerimiento) LIKE ? ORDER BY id DESC LIMIT 1", (c,))
            else:
                cur.execute("SELECT * FROM requerimientos WHERE UPPER(requerimiento)=UPPER(?)", (c,))
            req = cur.fetchone()
            if req:
                break
        if not req:
            return None, []
        cur.execute("SELECT *, (cantidad_requerida - cantidad_atendida) AS pendiente FROM requerimiento_detalle WHERE requerimiento_id=? ORDER BY item", (req["id"],))
        return req, cur.fetchall()

    def list_requerimientos(self, estados=None, term=""):
        like = f"%{normalize_text(term)}%"
        cur = self.conn.cursor()
        sql = """
            SELECT r.*, 
                   COALESCE(SUM(d.cantidad_requerida),0) AS total_requerido,
                   COALESCE(SUM(d.cantidad_atendida),0) AS total_atendido,
                   COALESCE(SUM(d.cantidad_requerida - d.cantidad_atendida),0) AS total_pendiente
            FROM requerimientos r
            LEFT JOIN requerimiento_detalle d ON d.requerimiento_id = r.id
            WHERE (UPPER(r.requerimiento) LIKE ? OR UPPER(r.solicitante) LIKE ? OR UPPER(IFNULL(r.ot,'')) LIKE ? OR UPPER(r.observacion) LIKE ?)
        """
        params = [like, like, like, like]
        if estados:
            placeholders = ",".join("?" for _ in estados)
            sql += f" AND r.estado IN ({placeholders})"
            params += list(estados)
        sql += " GROUP BY r.id ORDER BY r.id DESC LIMIT 500"
        cur.execute(sql, params)
        return cur.fetchall()

    def list_requerimientos_filtered(self, desde="", hasta="", estado="TODOS", almacen="TODOS", solicitante="", cc_codigo="", term=""):
        """Consulta consolidada para pantalla única de requerimientos."""
        cur = self.conn.cursor()
        sql = """
            SELECT r.*,
                   COALESCE(SUM(d.cantidad_requerida),0) AS total_requerido,
                   COALESCE(SUM(d.cantidad_atendida),0) AS total_atendido,
                   COALESCE(SUM(d.cantidad_requerida - d.cantidad_atendida),0) AS total_pendiente,
                   COUNT(d.id) AS total_items
            FROM requerimientos r
            LEFT JOIN requerimiento_detalle d ON d.requerimiento_id = r.id
            WHERE 1=1
        """
        params = []
        if desde:
            sql += " AND r.fecha_requerimiento >= ?"
            params.append(fecha_iso_consulta_from_text(desde))
        if hasta:
            sql += " AND r.fecha_requerimiento <= ?"
            params.append(fecha_iso_consulta_from_text(hasta))
        estado = normalize_text(estado)
        if estado and estado != "TODOS":
            sql += " AND r.estado = ?"
            params.append(estado)
        almacen = normalize_text(almacen)
        if almacen and almacen != "TODOS":
            sql += " AND r.almacen_codigo = ?"
            params.append(almacen)
        if solicitante:
            sql += " AND UPPER(r.solicitante) LIKE ?"
            params.append(f"%{normalize_text(solicitante)}%")
        if cc_codigo:
            sql += " AND r.centro_costo_codigo = ?"
            params.append(normalize_text(cc_codigo))
        if term:
            like = f"%{normalize_text(term)}%"
            sql += " AND (UPPER(r.requerimiento) LIKE ? OR UPPER(IFNULL(r.ot,'')) LIKE ? OR UPPER(r.observacion) LIKE ?)"
            params.extend([like, like, like])
        sql += " GROUP BY r.id ORDER BY r.fecha_requerimiento DESC, r.id DESC LIMIT 1000"
        cur.execute(sql, params)
        return cur.fetchall()

    def requerimiento_export_rows_filtered(self, desde="", hasta="", estado="TODOS", almacen="TODOS", solicitante="", cc_codigo="", term=""):
        cur = self.conn.cursor()
        sql = """
            SELECT r.requerimiento, r.fecha_requerimiento, r.fecha_entrega_solicitada, r.almacen_codigo, r.almacen_nombre, r.estado,
                   r.solicitante, r.centro_costo_codigo, r.centro_costo_nombre, r.ot, r.observacion,
                   d.item, d.codigo, d.descripcion, d.unidad_medida, d.cantidad_requerida, d.cantidad_atendida,
                   (d.cantidad_requerida - d.cantidad_atendida) AS pendiente,
                   r.usuario_creador, r.creado_en, r.modificado_en
            FROM requerimientos r
            JOIN requerimiento_detalle d ON d.requerimiento_id = r.id
            WHERE 1=1
        """
        params = []
        if desde:
            sql += " AND r.fecha_requerimiento >= ?"
            params.append(fecha_iso_consulta_from_text(desde))
        if hasta:
            sql += " AND r.fecha_requerimiento <= ?"
            params.append(fecha_iso_consulta_from_text(hasta))
        estado = normalize_text(estado)
        if estado and estado != "TODOS":
            sql += " AND r.estado = ?"
            params.append(estado)
        almacen = normalize_text(almacen)
        if almacen and almacen != "TODOS":
            sql += " AND r.almacen_codigo = ?"
            params.append(almacen)
        if solicitante:
            sql += " AND UPPER(r.solicitante) LIKE ?"
            params.append(f"%{normalize_text(solicitante)}%")
        if cc_codigo:
            sql += " AND r.centro_costo_codigo = ?"
            params.append(normalize_text(cc_codigo))
        if term:
            like = f"%{normalize_text(term)}%"
            sql += " AND (UPPER(r.requerimiento) LIKE ? OR UPPER(IFNULL(r.ot,'')) LIKE ? OR UPPER(r.observacion) LIKE ?)"
            params.extend([like, like, like])
        sql += " ORDER BY r.fecha_requerimiento DESC, r.requerimiento, d.item"
        cur.execute(sql, params)
        return cur.fetchall()

    def requerimiento_atenciones(self, requerimiento_id):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT vale, fecha_atencion, usuario, estado_atencion, observacion
            FROM requerimiento_atenciones
            WHERE requerimiento_id=?
            ORDER BY id
        """, (requerimiento_id,))
        return cur.fetchall()

    def attend_requerimiento(self, requerimiento_num, fecha_texto, cantidades, observacion_atencion, usuario):
        req, det = self.find_requerimiento(requerimiento_num)
        if not req:
            raise ValueError("No se encontró el requerimiento.")
        if req["estado"] in ["ATENDIDO", "ANULADO"]:
            raise ValueError(f"El requerimiento está {req['estado']} y no puede atenderse.")
        observacion_atencion = normalize_text(observacion_atencion) or "ATENCION DE REQUERIMIENTO"
        detalles_mov = []
        total_pendiente_final = 0
        total_atender = 0
        stock_proyectado = {}
        # cantidades: dict codigo -> atender
        for row in det:
            pendiente = float(row["cantidad_requerida"]) - float(row["cantidad_atendida"])
            if pendiente <= 0:
                continue
            codigo = row["codigo"]
            try:
                atender = float(str(cantidades.get(codigo, 0)).replace(",", "."))
            except Exception:
                raise ValueError(f"Cantidad atendida inválida para {codigo}.")
            if atender < 0:
                raise ValueError(f"La cantidad atendida de {codigo} no puede ser negativa.")
            if atender > pendiente:
                raise ValueError(f"No puede atender {atender:g} de {codigo}; pendiente máximo: {pendiente:g}.")
            if atender > 0:
                if codigo not in stock_proyectado:
                    stock_proyectado[codigo] = self.get_stock(codigo, req["almacen_codigo"])
                stock_proyectado[codigo] -= atender
                if stock_proyectado[codigo] < -0.00001:
                    stock = self.get_stock(codigo, req["almacen_codigo"])
                    raise ValueError(f"Stock insuficiente para {codigo}. Stock disponible en {req['almacen_codigo']}: {stock:g}. La atención acumulada supera el stock.")
                detalles_mov.append({"item": len(detalles_mov)+1, "codigo": codigo, "cantidad": atender})
                total_atender += atender
            total_pendiente_final += (pendiente - atender)
        if total_atender <= 0:
            raise ValueError("No hay cantidades para atender. Ingrese al menos una cantidad mayor a cero.")
        # Un requerimiento puede tener varias atenciones. Para no chocar con la regla de documento duplicado,
        # cada atención usa un documento interno correlativo, manteniendo el requerimiento original vinculado.
        cur = self.conn.cursor()
        cur.execute("SELECT COUNT(*) AS n FROM requerimiento_atenciones WHERE requerimiento_id=?", (req["id"],))
        n_atencion = int(cur.fetchone()["n"]) + 1
        documento_atencion = f"{req['requerimiento']}-AT{n_atencion:02d}"
        try:
            vale = self.save_movimiento(fecha_texto, "CO", documento_atencion, (req["ot"] or "").replace("OT-", ""), req["centro_costo_codigo"], req["solicitante"], observacion_atencion, detalles_mov, usuario, req["almacen_codigo"], req["requerimiento"], commit=False)
            for row in det:
                codigo = row["codigo"]
                atender = float(str(cantidades.get(codigo, 0)).replace(",", ".")) if codigo in cantidades else 0
                if atender > 0:
                    cur.execute("UPDATE requerimiento_detalle SET cantidad_atendida = cantidad_atendida + ? WHERE id=?", (atender, row["id"]))
            new_state = "ATENDIDO" if total_pendiente_final <= 0.00001 else "DESPACHO PARCIAL"
            cur.execute("UPDATE requerimientos SET estado=?, modificado_en=? WHERE id=?", (new_state, now_text(), req["id"]))
            cur.execute("""
                INSERT INTO requerimiento_atenciones(requerimiento_id, vale, fecha_atencion, usuario, estado_atencion, observacion)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (req["id"], vale, now_text(), usuario, new_state, observacion_atencion))
            self.audit("requerimientos", req["requerimiento"], "ATENCION", usuario, f"Vale {vale} - Estado {new_state}", commit=False)
            self.conn.commit()
            return vale, new_state
        except Exception:
            self.conn.rollback()
            raise

    def anular_requerimiento(self, requerimiento_num, motivo, usuario):
        req, det = self.find_requerimiento(requerimiento_num)
        if not req:
            raise ValueError("No se encontró el requerimiento.")
        if req["estado"] == "ANULADO":
            raise ValueError("El requerimiento ya está anulado.")
        if req["estado"] == "ATENDIDO":
            raise ValueError("No se puede anular un requerimiento atendido. Revierta primero los vales asociados si corresponde.")
        atendido = sum(float(r["cantidad_atendida"] or 0) for r in det)
        if atendido > 0.00001:
            raise ValueError("No se puede anular un requerimiento con despacho parcial. Revierta primero los vales de atención.")
        motivo = normalize_text(motivo)
        if not motivo:
            raise ValueError("Ingrese el motivo de anulación.")
        try:
            self.conn.execute("UPDATE requerimientos SET estado='ANULADO', observacion=observacion || ' | ANULADO: ' || ?, modificado_en=? WHERE id=?", (motivo, now_text(), req["id"]))
            self.audit("requerimientos", req["requerimiento"], "ANULACION", usuario, motivo, commit=False)
            self.conn.commit()
            return req["requerimiento"]
        except Exception:
            self.conn.rollback()
            raise

    def requerimiento_report_rows(self, estados=None):
        cur = self.conn.cursor()
        sql = """
            SELECT r.requerimiento, r.fecha_requerimiento, r.fecha_entrega_solicitada, r.almacen_codigo, r.almacen_nombre, r.estado,
                   r.solicitante, r.centro_costo_codigo, r.centro_costo_nombre, r.ot, r.observacion,
                   d.item, d.codigo, d.descripcion, d.unidad_medida, d.cantidad_requerida, d.cantidad_atendida,
                   (d.cantidad_requerida - d.cantidad_atendida) AS pendiente,
                   r.usuario_creador, r.creado_en, r.modificado_en
            FROM requerimientos r
            JOIN requerimiento_detalle d ON d.requerimiento_id = r.id
        """
        params = []
        if estados:
            placeholders = ",".join("?" for _ in estados)
            sql += f" WHERE r.estado IN ({placeholders})"
            params.extend(estados)
        sql += " ORDER BY r.fecha_requerimiento, r.requerimiento, d.item"
        cur.execute(sql, params)
        return cur.fetchall()

    def pendientes_por_articulo(self):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT r.almacen_codigo, d.codigo, d.descripcion, d.unidad_medida,
                   SUM(d.cantidad_requerida - d.cantidad_atendida) AS pendiente_total
            FROM requerimientos r
            JOIN requerimiento_detalle d ON d.requerimiento_id = r.id
            WHERE r.estado IN ('PENDIENTE','DESPACHO PARCIAL')
            GROUP BY r.almacen_codigo, d.codigo, d.descripcion, d.unidad_medida
            HAVING pendiente_total > 0
            ORDER BY r.almacen_codigo, d.descripcion
        """)
        return cur.fetchall()

    # =========================
    # CONFIGURACIÓN, API RUC Y TIPO DE CAMBIO
    # =========================

    def get_config(self, clave, default=""):
        cur = self.conn.cursor()
        cur.execute("SELECT valor FROM sistema_config WHERE clave=?", (clave,))
        row = cur.fetchone()
        if row:
            return row["valor"]
        if clave in API_CONFIG_DEFAULTS:
            return API_CONFIG_DEFAULTS[clave]
        return default

    def set_config(self, clave, valor, usuario="SISTEMA", descripcion=""):
        clave = str(clave or "").strip()
        valor = str(valor or "").strip()
        if not clave:
            raise ValueError("Clave de configuración inválida.")
        self.conn.execute("""
            INSERT INTO sistema_config(clave, valor, descripcion, modificado_por, modificado_en)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor, descripcion=excluded.descripcion, modificado_por=excluded.modificado_por, modificado_en=excluded.modificado_en
        """, (clave, valor, descripcion or f"Configuración {clave}", usuario, now_text()))
        self.conn.commit()

    def api_config(self):
        return {k: self.get_config(k, v) for k, v in API_CONFIG_DEFAULTS.items()}

    def _http_get_json(self, url, token="", timeout=10):
        headers = {"Accept": "application/json", "User-Agent": "KardexERP/1.0"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        req = urllib.request.Request(url, headers=headers, method="GET")
        try:
            with urllib.request.urlopen(req, timeout=float(timeout or 10)) as resp:
                raw = resp.read().decode("utf-8", errors="replace")
                if not raw.strip():
                    raise ValueError("La API respondió vacío.")
                return json.loads(raw)
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace") if hasattr(e, "read") else ""
            raise ValueError(f"Error HTTP {e.code} consultando API. {body[:180]}")
        except urllib.error.URLError as e:
            raise ValueError(f"No hay conexión o la API no respondió: {e.reason}")
        except json.JSONDecodeError:
            raise ValueError("La API no devolvió JSON válido.")

    def _ruc_payload_candidates(self, data):
        """Devuelve posibles diccionarios de contribuyente desde respuestas comunes de APIs RUC.
        Algunas APIs devuelven {data:{...}}, {result:{...}}, listas o incluso metadata.
        """
        candidates = []
        def add(obj):
            if isinstance(obj, dict):
                candidates.append(obj)
            elif isinstance(obj, list):
                for it in obj:
                    add(it)
        add(data)
        if isinstance(data, dict):
            for key in ("data", "result", "results", "response", "contribuyente", "empresa", "sunat"):
                if key in data:
                    add(data.get(key))
        return candidates

    def _get_any(self, data, keys):
        if not isinstance(data, dict):
            return ""
        lower = {str(k).lower(): v for k, v in data.items()}
        for k in keys:
            if k in data and data.get(k) not in (None, ""):
                return data.get(k)
            lk = k.lower()
            if lk in lower and lower.get(lk) not in (None, ""):
                return lower.get(lk)
        return ""

    def consultar_ruc_api(self, ruc):
        ruc = only_digits(ruc)
        if len(ruc) != 11:
            raise ValueError("Para consulta SUNAT/API el RUC debe tener 11 dígitos.")
        cfg = self.api_config()
        url_tpl = cfg.get("ruc_api_url") or API_CONFIG_DEFAULTS["ruc_api_url"]
        if "{ruc}" in url_tpl:
            url = url_tpl.replace("{ruc}", urllib.parse.quote(ruc))
        else:
            sep = "&" if "?" in url_tpl else "?"
            url = f"{url_tpl}{sep}numero={urllib.parse.quote(ruc)}"
        data = self._http_get_json(url, cfg.get("ruc_api_token", ""), cfg.get("api_timeout", "10"))

        best = None
        for cand in self._ruc_payload_candidates(data):
            cand_ruc = only_digits(self._get_any(cand, ["ruc", "numeroDocumento", "numero", "documento", "numRuc", "num_ruc", "RUC"]))
            if cand_ruc == ruc:
                best = cand
                break
            if best is None and any(self._get_any(cand, [k]) for k in ("razonSocial", "razon_social", "nombre_o_razon_social", "nombre", "denominacion", "business_name")):
                best = cand

        if best is None:
            raise ValueError("La API respondió JSON, pero no se encontró un bloque válido de contribuyente. Revise la URL configurada para consulta RUC.")

        cand_ruc = only_digits(self._get_any(best, ["ruc", "numeroDocumento", "numero", "documento", "numRuc", "num_ruc", "RUC"]))
        if cand_ruc and cand_ruc != ruc:
            raise ValueError(f"La API devolvió información de otro RUC ({cand_ruc}). Revise la URL/token; no se guardó el proveedor.")

        razon = self._get_any(best, [
            "razonSocial", "razon_social", "razon", "nombre_o_razon_social", "nombreORazonSocial",
            "nombreRazonSocial", "nombre", "denominacion", "denominacionSocial", "business_name",
            "ddp_nombre", "apellidosNombres", "full_name"
        ])
        direccion = self._get_any(best, [
            "direccion", "direccionFiscal", "domicilioFiscal", "direccionCompleta", "address",
            "ddp_nomvia", "domicilio"
        ])
        estado = self._get_any(best, ["estado", "estadoContribuyente", "estado_contribuyente", "status"])
        condicion = self._get_any(best, ["condicion", "condicionDomicilio", "condicion_domicilio", "condition"])
        ubigeo = self._get_any(best, ["ubigeo", "ubigeoSunat", "ubigeo_sunat", "codigoUbigeo"])

        razon_norm = normalize_text(razon)
        if not razon_norm:
            raise ValueError("La API no devolvió razón social. Revise si el endpoint corresponde a consulta RUC y no a otra API.")
        # Evita guardar respuestas de perfil/proveedor API cuando el endpoint no trae el RUC consultado.
        if cand_ruc == "" and razon_norm in {"REXTIE S.A.C.", "REXTIE SAC", "DECOLECTA", "APIS.NET"}:
            raise ValueError("La API parece devolver datos del proveedor de API y no del RUC consultado. Revise URL de consulta RUC y token.")

        return {
            "ruc": ruc,
            "razon_social": razon_norm,
            "direccion": normalize_text(direccion),
            "estado_contribuyente": normalize_text(estado),
            "condicion_domicilio": normalize_text(condicion),
            "ubigeo": str(ubigeo or "").strip(),
            "fuente_ruc": "API_RUC",
            "fecha_validacion_ruc": now_text(),
            "respuesta_ruc_json": json.dumps(data, ensure_ascii=False)[:4000],
        }

    def cargar_proveedor_desde_ruc_api(self, ruc, usuario="SISTEMA"):
        info = self.consultar_ruc_api(ruc)
        if not info["razon_social"]:
            raise ValueError("La API no devolvió razón social para este RUC.")
        actual = self.get_proveedor_by_ruc(info["ruc"], active_only=False)
        banco = actual["banco"] if actual else ""
        cuenta = actual["numero_cuenta"] if actual else ""
        cci = actual["cci"] if actual else ""
        contacto = actual["contacto"] if actual else ""
        telefono = actual["telefono"] if actual else ""
        email = actual["email"] if actual else ""
        moneda = actual["moneda_preferida"] if actual else "SOLES"
        proveedor_id = self.upsert_proveedor(
            info["ruc"], info["razon_social"], info["direccion"], telefono, email,
            banco, cuenta, cci, contacto, moneda, 1,
            estado_contribuyente=info["estado_contribuyente"],
            condicion_domicilio=info["condicion_domicilio"],
            ubigeo=info["ubigeo"], fuente_ruc=info["fuente_ruc"],
            fecha_validacion_ruc=info["fecha_validacion_ruc"],
            respuesta_ruc_json=info["respuesta_ruc_json"], usuario=usuario
        )
        self.audit("proveedores", info["ruc"], "CONSULTA RUC API", usuario, f"{info['razon_social']} | Estado {info['estado_contribuyente']} | Condición {info['condicion_domicilio']}")
        return proveedor_id, info

    def registrar_tipo_cambio_manual(self, fecha_texto, compra, venta, fuente="MANUAL", motivo="", usuario="SISTEMA"):
        fecha_iso = fecha_iso_consulta_from_text(fecha_texto)
        raw_compra = D(compra or 0); compra_d = q_tc(raw_compra)
        raw_venta = D(venta or 0); venta_d = q_tc(raw_venta)
        compra = float(compra_d); venta = float(venta_d)
        fuente = normalize_text(fuente or "MANUAL")
        motivo = normalize_text(motivo)
        if venta <= 0:
            raise ValueError("El tipo de cambio venta debe ser mayor a cero.")
        if compra < 0:
            raise ValueError("El tipo de cambio compra no puede ser negativo.")
        if fuente not in FUENTES_TC:
            fuente = "MANUAL"
        self.conn.execute("""
            INSERT INTO tipo_cambio(fecha, moneda, compra, venta, fuente, es_manual, motivo_manual, usuario_registro, fecha_registro)
            VALUES (?, 'DOLARES', ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(fecha, moneda, fuente) DO UPDATE SET compra=excluded.compra, venta=excluded.venta, es_manual=excluded.es_manual, motivo_manual=excluded.motivo_manual, usuario_registro=excluded.usuario_registro, fecha_registro=excluded.fecha_registro
        """, (fecha_iso, compra, venta, fuente, 1 if fuente == "MANUAL" else 0, motivo, usuario, now_text()))
        self.audit("tipo_cambio", fecha_iso, "REGISTRO MANUAL", usuario, f"Compra {compra:g} | Venta {venta:g} | Fuente {fuente}", commit=False)
        self.conn.commit()
        return self.get_tipo_cambio(fecha_iso, fuente=fuente)

    def get_tipo_cambio(self, fecha_iso_or_text, fuente=None):
        fecha = fecha_iso_or_text
        if "/" in str(fecha):
            fecha = fecha_iso_consulta_from_text(fecha)
        fuente = normalize_text(fuente or "")
        cur = self.conn.cursor()
        if fuente:
            cur.execute("SELECT * FROM tipo_cambio WHERE fecha=? AND moneda='DOLARES' AND fuente=?", (fecha, fuente))
        else:
            cur.execute("""
                SELECT * FROM tipo_cambio
                WHERE fecha=? AND moneda='DOLARES'
                ORDER BY CASE fuente WHEN 'API_SUNAT' THEN 1 WHEN 'SUNAT' THEN 2 WHEN 'SBS' THEN 3 WHEN 'BCRP' THEN 4 WHEN 'MANUAL' THEN 5 ELSE 9 END, id DESC
                LIMIT 1
            """, (fecha,))
        return cur.fetchone()

    def list_tipo_cambio(self, desde="", hasta=""):
        cur = self.conn.cursor()
        sql = "SELECT * FROM tipo_cambio WHERE 1=1"
        params = []
        if desde:
            sql += " AND fecha>=?"; params.append(fecha_iso_consulta_from_text(desde))
        if hasta:
            sql += " AND fecha<=?"; params.append(fecha_iso_consulta_from_text(hasta))
        sql += " ORDER BY fecha DESC, fuente"
        cur.execute(sql, params)
        return cur.fetchall()

    def consultar_tipo_cambio_api(self, fecha_texto, usuario="SISTEMA"):
        fecha_iso = fecha_iso_consulta_from_text(fecha_texto) if "/" in str(fecha_texto) else str(fecha_texto)[:10]
        cfg = self.api_config()
        url_tpl = cfg.get("tc_api_url") or API_CONFIG_DEFAULTS["tc_api_url"]
        if "{fecha}" in url_tpl:
            url = url_tpl.replace("{fecha}", urllib.parse.quote(fecha_iso))
        else:
            sep = "&" if "?" in url_tpl else "?"
            url = f"{url_tpl}{sep}date={urllib.parse.quote(fecha_iso)}"
        data = self._http_get_json(url, cfg.get("tc_api_token", ""), cfg.get("api_timeout", "10"))
        compra = data.get("compra") or data.get("precioCompra") or data.get("buy") or data.get("tcCompra") or 0
        venta = data.get("venta") or data.get("precioVenta") or data.get("sell") or data.get("tcVenta") or 0
        try:
            compra = float(q_tc(compra))
            venta = float(q_tc(venta))
        except Exception:
            raise ValueError("La API no devolvió compra/venta numéricos.")
        if venta <= 0:
            raise ValueError("La API no devolvió tipo de cambio venta válido.")
        self.conn.execute("""
            INSERT INTO tipo_cambio(fecha, moneda, compra, venta, fuente, es_manual, motivo_manual, usuario_registro, fecha_registro)
            VALUES (?, 'DOLARES', ?, ?, 'API_SUNAT', 0, '', ?, ?)
            ON CONFLICT(fecha, moneda, fuente) DO UPDATE SET compra=excluded.compra, venta=excluded.venta, usuario_registro=excluded.usuario_registro, fecha_registro=excluded.fecha_registro
        """, (fecha_iso, compra, venta, usuario, now_text()))
        self.audit("tipo_cambio", fecha_iso, "CONSULTA API", usuario, f"Compra {compra:g} | Venta {venta:g}", commit=False)
        self.conn.commit()
        return self.get_tipo_cambio(fecha_iso, fuente="API_SUNAT")

    def obtener_tipo_cambio_para_oc(self, fecha_iso, moneda, usuario="SISTEMA", permitir_api=True):
        moneda = normalize_text(moneda or "SOLES")
        if moneda == "SOLES":
            return {"venta": 1.0, "compra": 1.0, "fuente": "SOLES", "fecha": fecha_iso}
        if moneda != "DOLARES":
            raise ValueError("Moneda inválida.")
        row = self.get_tipo_cambio(fecha_iso)
        if not row and permitir_api:
            try:
                row = self.consultar_tipo_cambio_api(fecha_text_from_iso(fecha_iso), usuario)
            except Exception as e:
                raise ValueError(f"No se encontró tipo de cambio para {fecha_text_from_iso(fecha_iso)} y la API no respondió. Registre el TC manualmente. Detalle: {e}")
        if not row:
            raise ValueError(f"No existe tipo de cambio para {fecha_text_from_iso(fecha_iso)}. Regístrelo manualmente antes de emitir OC en dólares.")
        return {"venta": float(row["venta"]), "compra": float(row["compra"] or 0), "fuente": row["fuente"], "fecha": row["fecha"]}

    # =========================
    # MÓDULO ERP/MRP: PROVEEDORES, ÓRDENES DE COMPRA, FINANZAS, TRÁNSITO Y LIQUIDACIÓN
    # =========================

    def upsert_proveedor(self, ruc, razon_social, direccion="", telefono="", email="", banco="", numero_cuenta="", cci="", contacto="", moneda_preferida="SOLES", activo=1, estado_contribuyente="", condicion_domicilio="", ubigeo="", fuente_ruc="", fecha_validacion_ruc="", respuesta_ruc_json="", usuario="SISTEMA"):
        ruc = only_digits(ruc)
        razon_social = normalize_text(razon_social)
        direccion = normalize_text(direccion)
        telefono = str(telefono or "").strip()
        email = str(email or "").strip().lower()
        banco = normalize_text(banco)
        numero_cuenta = str(numero_cuenta or "").strip()
        cci = str(cci or "").strip()
        contacto = normalize_text(contacto)
        moneda_preferida = normalize_text(moneda_preferida) or "SOLES"
        estado_contribuyente = normalize_text(estado_contribuyente)
        condicion_domicilio = normalize_text(condicion_domicilio)
        ubigeo = str(ubigeo or "").strip()
        fuente_ruc = normalize_text(fuente_ruc)
        fecha_validacion_ruc = str(fecha_validacion_ruc or "").strip()
        respuesta_ruc_json = str(respuesta_ruc_json or "")[:4000]
        if not ruc or len(ruc) not in (8, 11):
            raise ValueError("Ingrese RUC/DNI válido del proveedor.")
        if len(ruc) == 11 and not ruc.startswith(("10", "15", "17", "20")):
            raise ValueError("El RUC peruano debe iniciar con 10, 15, 17 o 20.")
        if not razon_social:
            raise ValueError("Ingrese razón social del proveedor.")
        banco = normalizar_banco(banco)
        if banco and (banco_codigo_guardado(banco) not in BANCOS_MAP and not banco.startswith("05 OTROS")):
            raise ValueError("Banco inválido.")
        if moneda_preferida not in MONEDAS:
            raise ValueError("Moneda preferida inválida.")
        cur = self.conn.cursor()
        cur.execute("SELECT id FROM proveedores WHERE ruc=?", (ruc,))
        row = cur.fetchone()
        if row:
            cur.execute("""
                UPDATE proveedores
                SET razon_social=?, direccion=?, telefono=?, email=?, banco=?, numero_cuenta=?, cci=?, contacto=?, moneda_preferida=?, activo=?,
                    estado_contribuyente=?, condicion_domicilio=?, ubigeo=?, fuente_ruc=?, fecha_validacion_ruc=?, respuesta_ruc_json=?, modificado_en=?
                WHERE id=?
            """, (razon_social, direccion, telefono, email, banco, numero_cuenta, cci, contacto, moneda_preferida, 1 if activo else 0,
                  estado_contribuyente, condicion_domicilio, ubigeo, fuente_ruc, fecha_validacion_ruc, respuesta_ruc_json, now_text(), row["id"]))
            proveedor_id = row["id"]
            accion = "MODIFICACION"
        else:
            cur.execute("""
                INSERT INTO proveedores(ruc, razon_social, direccion, telefono, email, banco, numero_cuenta, cci, contacto, moneda_preferida, activo,
                                        estado_contribuyente, condicion_domicilio, ubigeo, fuente_ruc, fecha_validacion_ruc, respuesta_ruc_json, creado_en)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (ruc, razon_social, direccion, telefono, email, banco, numero_cuenta, cci, contacto, moneda_preferida, 1 if activo else 0,
                  estado_contribuyente, condicion_domicilio, ubigeo, fuente_ruc, fecha_validacion_ruc, respuesta_ruc_json, now_text()))
            proveedor_id = cur.lastrowid
            accion = "CREACION"
        self.audit("proveedores", ruc, accion, usuario or "SISTEMA", razon_social, commit=False)
        self.conn.commit()
        return proveedor_id

    def get_proveedor_by_ruc(self, ruc, active_only=True):
        cur = self.conn.cursor()
        sql = "SELECT * FROM proveedores WHERE ruc=?"
        params = [only_digits(ruc)]
        if active_only:
            sql += " AND activo=1"
        cur.execute(sql, params)
        return cur.fetchone()

    def list_proveedores(self, term="", active_only=True):
        like = f"%{normalize_text(term)}%"
        cur = self.conn.cursor()
        sql = """
            SELECT * FROM proveedores
            WHERE (UPPER(ruc) LIKE ? OR UPPER(razon_social) LIKE ? OR UPPER(contacto) LIKE ?)
        """
        params = [like, like, like]
        if active_only:
            sql += " AND activo=1"
        sql += " ORDER BY razon_social LIMIT 500"
        cur.execute(sql, params)
        return cur.fetchall()

    def next_oc_number(self, tipo_orden="NACIONAL", almacen_codigo="AQP", year=None):
        tipo_orden = normalize_text(tipo_orden) or "NACIONAL"
        doc_type = "OS" if tipo_orden == "SERVICIO" else "OC"
        year = int(year or date.today().year)
        return AnnualCorrelativeService(self.db_path, self.conn).next(doc_type, year)

    def oc_avance_from_estado(self, estado):
        estado = normalize_text(estado)
        if estado in ("PENDIENTE", "EN COTIZACION", "DENEGADA", "ANULADA"):
            return 0
        if estado in ("APROBADA", "APROBADA PARCIAL"):
            return 25
        if estado in ("PAGO PARCIAL", "PAGADA"):
            return 50
        if estado == "EN TRANSITO":
            return 80
        if estado in ("ATENDIDA PARCIALMENTE", "ATENDIDA", "LIQUIDADA"):
            return 100
        return 0

    def _calc_price_values(self, cantidad, precio_unitario, precio_incluye_igv):
        precio_sin, precio_con, subtotal, igv, total = decimal_calc_price_values(
            cantidad, precio_unitario, bool(precio_incluye_igv)
        )
        # SQLite V44 mantiene columnas REAL por compatibilidad beta, pero todos
        # los calculos y redondeos se realizan primero con Decimal.
        return float(precio_sin), float(precio_con), float(subtotal), float(igv), float(total)


    def oc_tiene_movimientos_operativos(self, oc_id):
        cur = self.conn.cursor()
        checks = {}
        for name, table in [("pagos", "orden_compra_pagos"), ("transito", "orden_compra_transito"), ("recepciones", "orden_compra_recepciones"), ("liquidacion", "orden_compra_liquidacion")]:
            cur.execute(f"SELECT COUNT(*) c FROM {table} WHERE oc_id=?", (oc_id,))
            checks[name] = int(cur.fetchone()["c"] or 0)
        return checks

    def update_orden_compra_pendiente(self, numero, ruc, cotizacion, observacion, forma_pago, fecha_entrega, lugar_entrega, usuario):
        oc, det = self.get_orden_compra(numero)
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        if oc["estado"] not in ("PENDIENTE", "EN COTIZACION"):
            raise ValueError("Solo se puede modificar una OC pendiente o en cotización.")
        proveedor = self.get_proveedor_by_ruc(ruc, active_only=True)
        if not proveedor:
            raise ValueError("Proveedor no encontrado o inactivo.")
        forma_pago = normalize_text(forma_pago)
        if forma_pago not in FORMAS_PAGO:
            raise ValueError("Forma de pago inválida.")
        fecha_entrega_iso = fecha_iso_consulta_from_text(fecha_entrega) if fecha_entrega else ""
        venc = self.fecha_vencimiento_pago(oc["fecha_orden"], forma_pago, fecha_entrega_iso)
        old = f"Proveedor={oc['ruc']} {oc['razon_social']} | Cot={oc['cotizacion']} | Forma={oc['forma_pago']}"
        new = f"Proveedor={proveedor['ruc']} {proveedor['razon_social']} | Cot={normalize_text(cotizacion)} | Forma={forma_pago}"
        self.conn.execute("""
            UPDATE orden_compra SET proveedor_id=?, ruc=?, razon_social=?, direccion=?, telefono=?, cotizacion=?, observacion=?, forma_pago=?, fecha_vencimiento_pago=?, fecha_entrega=?, lugar_entrega=?, modificado_en=?
            WHERE id=?
        """, (proveedor["id"], proveedor["ruc"], proveedor["razon_social"], proveedor["direccion"], proveedor["telefono"], normalize_text(cotizacion), normalize_text(observacion), forma_pago, venc, fecha_entrega_iso, normalize_text(lugar_entrega), now_text(), oc["id"]))
        self.audit("orden_compra", numero, "MODIFICACION PENDIENTE", usuario, f"Antes: {old} / Después: {new}", commit=False)
        self.conn.commit()
        return numero

    def desaprobar_orden_compra(self, numero, motivo, usuario):
        oc, det = self.get_orden_compra(numero)
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        if oc["estado"] not in ("APROBADA", "APROBADA PARCIAL"):
            raise ValueError("Solo se puede desaprobar una OC aprobada sin movimientos posteriores.")
        checks = self.oc_tiene_movimientos_operativos(oc["id"])
        usados = [k for k, v in checks.items() if v > 0]
        if usados:
            raise ValueError("No se puede desaprobar porque la OC ya tiene movimientos: " + ", ".join(usados) + ".")
        motivo = normalize_text(motivo)
        if not motivo:
            raise ValueError("Ingrese motivo de desaprobación.")
        try:
            self.conn.execute("UPDATE orden_compra_detalle SET cantidad_aprobada=0 WHERE oc_id=?", (oc["id"],))
            self.conn.execute("UPDATE orden_compra SET aprobado_por='', fecha_aprobacion='', motivo_decision=?, modificado_en=? WHERE id=?", (motivo, now_text(), oc["id"]))
            self._set_oc_state(oc["id"], "PENDIENTE", usuario, f"DESAPROBACION: {motivo}")
            self.conn.commit()
            return "PENDIENTE"
        except Exception:
            self.conn.rollback()
            raise

    def requerimiento_estado_operativo(self, requerimiento_num):
        req, det = self.find_requerimiento(requerimiento_num)
        if not req:
            raise ValueError("No se encontró el requerimiento.")
        cur = self.conn.cursor()
        rows = []
        for d in det:
            codigo = d["codigo"]
            stock_alm = self.get_stock(codigo, req["almacen_codigo"])
            cur.execute("""
                SELECT oc.numero, oc.estado, oc.estado_pago, oc.avance, oc.razon_social, od.cantidad_solicitada, od.cantidad_aprobada, od.cantidad_ingresada
                FROM orden_compra oc
                JOIN orden_compra_detalle od ON od.oc_id = oc.id
                WHERE oc.requerimiento=? AND od.codigo=? AND oc.estado NOT IN ('ANULADA','DENEGADA')
                ORDER BY oc.id DESC
            """, (req["requerimiento"], codigo))
            ocs = cur.fetchall()
            rows.append({"detalle": d, "stock_almacen": stock_alm, "ocs": ocs})
        return req, rows

    def oc_mapa_proceso(self, numero):
        oc, det = self.get_orden_compra(numero)
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM orden_compra_pagos WHERE oc_id=? ORDER BY fecha_pago, id", (oc["id"],)); pagos = cur.fetchall()
        cur.execute("SELECT * FROM orden_compra_transito WHERE oc_id=? ORDER BY fecha_despacho, id", (oc["id"],)); transito = cur.fetchall()
        cur.execute("SELECT * FROM orden_compra_recepciones WHERE oc_id=? ORDER BY fecha_recepcion, id", (oc["id"],)); recepciones = cur.fetchall()
        cur.execute("SELECT * FROM orden_compra_liquidacion WHERE oc_id=?", (oc["id"],)); liquidacion = cur.fetchone()
        req = None
        if oc["requerimiento"]:
            req, _ = self.find_requerimiento(oc["requerimiento"])
        return {"oc": oc, "det": det, "req": req, "pagos": pagos, "transito": transito, "recepciones": recepciones, "liquidacion": liquidacion, "historial": self.oc_historial(oc["id"])}

    def fecha_vencimiento_pago(self, fecha_orden_iso, forma_pago, fecha_entrega_texto=""):
        forma_pago = normalize_text(forma_pago or "CONTADO")
        try:
            base = datetime.strptime(fecha_orden_iso, ISO_FMT).date()
        except Exception:
            base = date.today()
        if forma_pago == "CREDITO 15 DIAS":
            return (base + timedelta(days=15)).strftime(ISO_FMT)
        if forma_pago == "CREDITO 30 DIAS":
            return (base + timedelta(days=30)).strftime(ISO_FMT)
        if forma_pago == "50% INICIAL/50% FINAL":
            # La cuota final vence con la entrega estimada. Si no se informó entrega, se usa la fecha de emisión.
            if fecha_entrega_texto:
                try:
                    return fecha_iso_consulta_from_text(fecha_entrega_texto)
                except Exception:
                    pass
            return fecha_orden_iso
        return fecha_orden_iso

    def _calc_estado_pago(self, total, pagado, fecha_vencimiento_iso):
        total = float(total or 0)
        pagado = float(pagado or 0)
        saldo = max(0.0, total - pagado)
        if saldo <= 0.01 and total > 0:
            return "PAGADA"
        if fecha_vencimiento_iso:
            try:
                venc = datetime.strptime(fecha_vencimiento_iso, ISO_FMT).date()
                if venc < date.today() and saldo > 0.01:
                    return "VENCIDA"
            except Exception:
                pass
        if pagado > 0.01:
            return "PAGO PARCIAL"
        return "POR PAGAR"

    def _sync_oc_payment_status(self, oc_id, commit=False):
        cur = self.conn.cursor()
        cur.execute("SELECT total, fecha_vencimiento_pago FROM orden_compra WHERE id=?", (oc_id,))
        oc = cur.fetchone()
        if not oc:
            return "POR PAGAR", 0.0, 0.0
        cur.execute("SELECT COALESCE(SUM(monto),0) AS pagado FROM orden_compra_pagos WHERE oc_id=?", (oc_id,))
        pagado = float(cur.fetchone()["pagado"] or 0)
        total = float(oc["total"] or 0)
        saldo = max(0.0, total - pagado)
        estado_pago = self._calc_estado_pago(total, pagado, oc["fecha_vencimiento_pago"])
        cur.execute("UPDATE orden_compra SET estado_pago=?, monto_pagado_cache=?, modificado_en=? WHERE id=?", (estado_pago, pagado, now_text(), oc_id))
        if commit:
            self.conn.commit()
        return estado_pago, pagado, saldo

    def refresh_payment_statuses(self):
        cur = self.conn.cursor()
        cur.execute("SELECT id FROM orden_compra WHERE estado NOT IN ('ANULADA','DENEGADA')")
        ids = [r["id"] for r in cur.fetchall()]
        for oc_id in ids:
            self._sync_oc_payment_status(oc_id, commit=False)
        self.conn.commit()

    def _recalc_oc_totals(self, oc_id):
        cur = self.conn.cursor()
        cur.execute("SELECT COALESCE(SUM(subtotal),0) AS subtotal, COALESCE(SUM(igv),0) AS igv, COALESCE(SUM(total),0) AS total FROM orden_compra_detalle WHERE oc_id=?", (oc_id,))
        t = cur.fetchone()
        cur.execute("UPDATE orden_compra SET subtotal=?, igv=?, total=?, modificado_en=? WHERE id=?", (float(t["subtotal"]), float(t["igv"]), float(t["total"]), now_text(), oc_id))
        return float(t["subtotal"]), float(t["igv"]), float(t["total"])

    def _add_oc_historial(self, oc_id, estado_anterior, estado_nuevo, usuario, detalle=""):
        self.conn.execute("""
            INSERT INTO orden_compra_estado_historial(oc_id, estado_anterior, estado_nuevo, usuario, fecha_hora, detalle)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (oc_id, estado_anterior or "", estado_nuevo, usuario, now_text(), detalle or ""))

    def _set_oc_state(self, oc_id, estado_nuevo, usuario, detalle=""):
        estado_nuevo = normalize_text(estado_nuevo)
        if estado_nuevo not in ESTADOS_OC:
            raise ValueError("Estado de orden de compra inválido.")
        cur = self.conn.cursor()
        cur.execute("SELECT estado, numero FROM orden_compra WHERE id=?", (oc_id,))
        oc = cur.fetchone()
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        estado_anterior = oc["estado"]
        cur.execute("UPDATE orden_compra SET estado=?, avance=?, modificado_en=? WHERE id=?", (estado_nuevo, self.oc_avance_from_estado(estado_nuevo), now_text(), oc_id))
        self._add_oc_historial(oc_id, estado_anterior, estado_nuevo, usuario, detalle)
        self.audit("orden_compra", oc["numero"], "CAMBIO ESTADO", usuario, f"{estado_anterior} -> {estado_nuevo}. {detalle}", commit=False)

    def create_orden_compra(self, fecha_texto, almacen_codigo, ruc, tipo_orden, requerimiento, cotizacion, solicitante, ot_num, cc_codigo, observacion, moneda, precio_incluye_igv, detalles, usuario, forma_pago="CONTADO", fecha_entrega="", lugar_entrega="", tipo_cambio=None, tipo_cambio_fuente="", tipo_cambio_fecha="", commit=True):

        fecha_iso = fecha_iso_from_text(fecha_texto)
        fecha_visible = fecha_text_from_iso(fecha_iso)
        almacen_codigo = normalize_text(almacen_codigo or "AQP")
        tipo_orden = normalize_text(tipo_orden or "NACIONAL")
        requerimiento = normalize_text(requerimiento)
        cotizacion = normalize_text(cotizacion)
        solicitante = normalize_text(solicitante)
        cc_codigo = normalize_text(cc_codigo)
        observacion = normalize_text(observacion)
        moneda = normalize_text(moneda or "SOLES")
        forma_pago = normalize_text(forma_pago or "CONTADO")
        lugar_entrega = normalize_text(lugar_entrega or ALMACENES.get(almacen_codigo, ""))
        if almacen_codigo not in ALMACENES:
            raise ValueError("Almacén inválido para OC.")
        if tipo_orden not in TIPOS_ORDEN_COMPRA:
            raise ValueError("Tipo de orden inválido.")
        if moneda not in MONEDAS:
            raise ValueError("Moneda inválida.")
        if forma_pago not in FORMAS_PAGO:
            raise ValueError("Forma de pago inválida.")
        if moneda == "SOLES":
            tc_info = {"venta": 1.0, "compra": 1.0, "fuente": "SOLES", "fecha": fecha_iso}
        else:
            if tipo_cambio not in (None, ""):
                tc_val = float(str(tipo_cambio).replace(",", "."))
                if tc_val <= 0:
                    raise ValueError("El tipo de cambio debe ser mayor a cero para OC en dólares.")
                tc_info = {"venta": tc_val, "compra": tc_val, "fuente": normalize_text(tipo_cambio_fuente or "MANUAL"), "fecha": tipo_cambio_fecha or fecha_iso}
            else:
                tc_info = self.obtener_tipo_cambio_para_oc(fecha_iso, moneda, usuario, permitir_api=True)
        if not cc_codigo or cc_codigo not in CENTROS_COSTO:
            raise ValueError("Centro de costo inválido.")
        if not solicitante:
            raise ValueError("Ingrese solicitante.")
        if not observacion:
            raise ValueError("Ingrese observación.")
        if not detalles:
            raise ValueError("Debe ingresar al menos un item.")
        proveedor = self.get_proveedor_by_ruc(ruc, active_only=True)
        if not proveedor:
            raise ValueError("Proveedor no existe o está inactivo. Regístrelo antes de generar la OC.")
        if requerimiento:
            cur = self.conn.cursor()
            cur.execute("""
                SELECT numero FROM orden_compra
                WHERE requerimiento=? AND ruc=? AND estado NOT IN ('ANULADA','DENEGADA','LIQUIDADA')
                LIMIT 1
            """, (requerimiento, proveedor["ruc"]))
            dup = cur.fetchone()
            if dup:
                raise ValueError(f"Ya existe una OC activa para este requerimiento y proveedor: {dup['numero']}. Si se anula, podrá generarse nuevamente.")
        numero = self.next_oc_number(tipo_orden, almacen_codigo, int(fecha_iso[:4]))
        ot = clean_ot(ot_num)
        fecha_vencimiento_pago = self.fecha_vencimiento_pago(fecha_iso, forma_pago, fecha_entrega)
        try:
            cur = self.conn.cursor()
            cur.execute("""
                INSERT INTO orden_compra(numero, fecha, fecha_orden, almacen_codigo, almacen_nombre, proveedor_id, ruc, razon_social, direccion, telefono, tipo_orden, requerimiento, cotizacion, solicitante, ot, centro_costo_codigo, centro_costo_nombre, observacion, moneda, tipo_cambio, tipo_cambio_compra, tipo_cambio_venta, tipo_cambio_fuente, tipo_cambio_fecha, precio_incluye_igv, forma_pago, fecha_vencimiento_pago, estado_pago, monto_pagado_cache, fecha_entrega, lugar_entrega, usuario_creador, creado_en)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (numero, fecha_visible, fecha_iso, almacen_codigo, ALMACENES[almacen_codigo], proveedor["id"], proveedor["ruc"], proveedor["razon_social"], proveedor["direccion"], proveedor["telefono"], tipo_orden, requerimiento, cotizacion, solicitante, ot, cc_codigo, CENTROS_COSTO[cc_codigo], observacion, moneda, float(tc_info["venta"]), float(tc_info.get("compra") or 0), float(tc_info["venta"]), tc_info.get("fuente", ""), tc_info.get("fecha", fecha_iso), 1 if precio_incluye_igv else 0, forma_pago, fecha_vencimiento_pago, "POR PAGAR", 0.0, fecha_entrega, lugar_entrega, usuario, now_text()))
            oc_id = cur.lastrowid
            for idx, d in enumerate(detalles, start=1):
                codigo = normalize_text(d.get("codigo", ""))
                descripcion = normalize_text(d.get("descripcion", ""))
                unidad = normalize_text(d.get("unidad_medida") or d.get("um") or "UND")
                raw_cantidad = D(d.get("cantidad") or d.get("cantidad_solicitada") or 0)
                cantidad_d = q_qty(raw_cantidad)
                if raw_cantidad != cantidad_d:
                    raise ValueError(f"La cantidad del item {idx} admite como máximo 3 decimales.")
                raw_precio = D(d.get("precio") or d.get("precio_unitario") or 0)
                precio_d = q_unit(raw_precio)
                cantidad = float(cantidad_d)
                precio = float(precio_d)
                cc_det = normalize_text(d.get("centro_costo_codigo") or cc_codigo)
                ot_det = clean_ot(d.get("ot") or ot)
                sol_det = normalize_text(d.get("solicitante") or solicitante)
                if tipo_orden == "NACIONAL":
                    art = self.get_articulo(codigo)
                    if not art:
                        raise ValueError(f"El artículo {codigo} no existe para OC nacional.")
                    descripcion = art["descripcion"]
                    unidad = art["unidad_medida"]
                if not descripcion:
                    raise ValueError("Ingrese descripción del item.")
                if unidad not in UNIDADES_PERMITIDAS:
                    raise ValueError(f"Unidad inválida en item {idx}.")
                if cc_det and cc_det not in CENTROS_COSTO:
                    raise ValueError(f"Centro de costo inválido en item {idx}.")
                precio_sin, precio_con, subtotal, igv, total = self._calc_price_values(cantidad, precio, precio_incluye_igv)
                cur.execute("""
                    INSERT INTO orden_compra_detalle(oc_id, item, codigo, descripcion, unidad_medida, cantidad_solicitada, cantidad_aprobada, cantidad_ingresada, precio_unitario_sin_igv, precio_unitario_con_igv, subtotal, igv, total, centro_costo_codigo, ot, solicitante)
                    VALUES (?, ?, ?, ?, ?, ?, 0, 0, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (oc_id, idx, codigo, descripcion, unidad, cantidad, precio_sin, precio_con, subtotal, igv, total, cc_det, ot_det, sol_det))
            self._recalc_oc_totals(oc_id)
            self._add_oc_historial(oc_id, "", "PENDIENTE", usuario, "Creación de OC")
            self.audit("orden_compra", numero, "CREACION", usuario, f"Proveedor {proveedor['razon_social']} | Tipo {tipo_orden} | Moneda {moneda} | TC {tc_info['venta']:g} {tc_info.get('fuente','')}", commit=False)
            if commit:
                self.conn.commit()
            return numero
        except Exception:
            self.conn.rollback()
            raise

    def create_oc_from_requerimiento(self, requerimiento_num, ruc, cotizacion, precios_por_codigo, moneda, precio_incluye_igv, usuario):
        req_nums_txt = requerimiento_num
        if isinstance(requerimiento_num, str) and (";" in requerimiento_num or "+" in requerimiento_num):
            req, det = self.collect_requerimientos_for_oc(requerimiento_num)
        elif isinstance(requerimiento_num, (list, tuple)):
            req, det = self.collect_requerimientos_for_oc(requerimiento_num)
            req_nums_txt = req["requerimiento"]
        else:
            req, det = self.find_requerimiento(requerimiento_num)
            if not req:
                raise ValueError("No se encontró el requerimiento.")
            activa = self.requerimiento_tiene_oc_activa(req["requerimiento"])
            if activa:
                raise ValueError(f"El requerimiento {req['requerimiento']} ya tiene una OC activa asociada: {activa['numero']} ({activa['estado']}).")
            if req["estado"] in ("ANULADO", "ATENDIDO"):
                raise ValueError(f"No se puede generar OC desde un requerimiento {req['estado']}.")
        detalles = []
        for row in det:
            pendiente = float(row["cantidad_requerida"] or 0) - float(row["cantidad_atendida"] or 0)
            if pendiente <= 0:
                continue
            codigo = row["codigo"]
            precio = float(str(precios_por_codigo.get(codigo, 0)).replace(",", ".")) if codigo in precios_por_codigo else 0
            if precio <= 0:
                continue
            detalles.append({"codigo": codigo, "descripcion": row["descripcion"], "um": row["unidad_medida"], "cantidad": pendiente, "precio": precio, "centro_costo_codigo": req["centro_costo_codigo"], "ot": req["ot"], "solicitante": req["solicitante"]})
        if not detalles:
            raise ValueError("Ingrese precio mayor a cero para al menos un item pendiente del requerimiento.")
        return self.create_orden_compra(today_text(), req["almacen_codigo"], ruc, req["tipo_orden"] if "tipo_orden" in req.keys() else "NACIONAL", req["requerimiento"], cotizacion, req["solicitante"], (req["ot"] or "").replace("OT-", ""), req["centro_costo_codigo"], f"OC GENERADA DESDE {req['requerimiento']}", moneda, precio_incluye_igv, detalles, usuario)


    def collect_requerimientos_for_oc(self, requerimientos):
        if isinstance(requerimientos, str):
            req_nums = [x.strip() for x in requerimientos.replace("+", ";").split(";") if x.strip()]
        else:
            req_nums = [str(x).strip() for x in requerimientos if str(x).strip()]
        if not req_nums:
            raise ValueError("Seleccione al menos un requerimiento.")
        base = None
        base_ot = None
        agg = {}
        for num in req_nums:
            req, det = self.find_requerimiento(num)
            if not req:
                raise ValueError(f"No se encontró el requerimiento {num}.")
            activa = self.requerimiento_tiene_oc_activa(num)
            if activa:
                raise ValueError(f"El requerimiento {num} ya tiene una OC activa asociada: {activa['numero']} ({activa['estado']}).")
            if req["estado"] in ("ANULADO", "ATENDIDO"):
                raise ValueError(f"El requerimiento {num} está {req['estado']} y no puede incluirse en una OC.")
            ot = normalize_text(req["ot"] or "")
            if base is None:
                base = dict(req)
                base_ot = ot
            elif ot != base_ot:
                raise ValueError("Solo se pueden unir requerimientos que pertenecen a la misma OT.")
            # Para que la OC tenga una cabecera consistente, se mantiene almacén/CC/tipo del primer requerimiento.
            for r in det:
                pend = float(r["cantidad_requerida"] or 0) - float(r["cantidad_atendida"] or 0)
                if pend <= 0:
                    continue
                cod = r["codigo"]
                if cod not in agg:
                    agg[cod] = {
                        "codigo": cod,
                        "descripcion": r["descripcion"],
                        "unidad_medida": r["unidad_medida"],
                        "cantidad_requerida": 0.0,
                        "cantidad_atendida": 0.0,
                    }
                agg[cod]["cantidad_requerida"] += pend
        if not agg:
            raise ValueError("Los requerimientos seleccionados no tienen cantidades pendientes.")
        base["requerimiento"] = ";".join(req_nums)
        return base, list(agg.values())


    def requerimiento_tiene_oc_activa(self, requerimiento_num):
        """Indica si un requerimiento ya está vinculado a una OC activa.

        Se considera activa toda OC que no esté ANULADA ni DENEGADA.
        La búsqueda soporta OC generadas con un solo requerimiento o con varios
        requerimientos unidos por punto y coma.
        """
        req = normalize_text(requerimiento_num)
        if not req:
            return None
        cur = self.conn.cursor()
        cur.execute("""
            SELECT numero, estado
            FROM orden_compra
            WHERE estado NOT IN ('ANULADA','DENEGADA')
              AND (
                    UPPER(requerimiento)=UPPER(?)
                 OR (';' || UPPER(REPLACE(IFNULL(requerimiento,''), '+', ';')) || ';') LIKE ?
              )
            ORDER BY id DESC
            LIMIT 1
        """, (req, f"%;{req};%"))
        return cur.fetchone()


    def create_ocs_from_requerimiento_multiproveedor(self, requerimiento_num, proveedores_precios, precio_incluye_igv, usuario):
        """Genera una OC por proveedor desde uno o varios requerimientos.

        Reglas críticas:
        - Solo incluye en cada OC los items con precio > 0 para ese proveedor.
        - Antes de insertar, valida todos los proveedores y precios del lote.
        - La generación es atómica: si un proveedor o item falla, no queda ninguna OC parcial.
        """
        if isinstance(requerimiento_num, str) and (";" in requerimiento_num or "+" in requerimiento_num):
            req, det = self.collect_requerimientos_for_oc(requerimiento_num)
        elif isinstance(requerimiento_num, (list, tuple)):
            req, det = self.collect_requerimientos_for_oc(requerimiento_num)
        else:
            req, det = self.find_requerimiento(requerimiento_num)
            if not req:
                raise ValueError("No se encontró el requerimiento.")
            activa = self.requerimiento_tiene_oc_activa(req["requerimiento"])
            if activa:
                raise ValueError(f"El requerimiento {req['requerimiento']} ya tiene una OC activa asociada: {activa['numero']} ({activa['estado']}).")
            if req["estado"] in ("ANULADO", "ATENDIDO"):
                raise ValueError(f"No se puede generar OC desde un requerimiento {req['estado']}.")

        if not proveedores_precios:
            raise ValueError("Agregue al menos un proveedor con precios.")

        detalles_por_codigo = {normalize_text(r["codigo"]): r for r in det}
        preparados = []

        # Prevalidación completa antes de insertar cualquier OC. Esto evita que un lote
        # multiproveedor deje documentos huérfanos si el segundo/tercer proveedor falla.
        for prov in proveedores_precios:
            ruc = only_digits(prov.get("ruc"))
            proveedor = self.get_proveedor_by_ruc(ruc, active_only=True)
            if not proveedor:
                raise ValueError(f"Proveedor {ruc or '(vacío)'} no existe o está inactivo. Regístrelo antes de generar la OC.")
            moneda = normalize_text(prov.get("moneda") or "SOLES")
            if moneda not in MONEDAS:
                raise ValueError(f"Moneda inválida para proveedor {proveedor['razon_social']}.")
            precios = prov.get("precios") or {}
            detalles_oc = []
            for codigo, raw_precio in precios.items():
                codigo = normalize_text(codigo)
                if codigo not in detalles_por_codigo:
                    continue
                try:
                    precio = float(str(raw_precio or 0).replace(",", "."))
                except Exception:
                    raise ValueError(f"Precio inválido para artículo {codigo} del proveedor {ruc}.")
                if precio <= 0:
                    continue
                row = detalles_por_codigo[codigo]
                pendiente = float(row["cantidad_requerida"] or 0) - float(row["cantidad_atendida"] or 0)
                if pendiente <= 0:
                    continue
                detalles_oc.append({
                    "codigo": codigo,
                    "descripcion": row["descripcion"],
                    "um": row["unidad_medida"],
                    "cantidad": pendiente,
                    "precio": precio,
                    "centro_costo_codigo": req["centro_costo_codigo"],
                    "ot": req["ot"],
                    "solicitante": req["solicitante"],
                })
            if detalles_oc:
                preparados.append((prov, detalles_oc))

        if not preparados:
            raise ValueError("No hay precios mayores a cero para generar órdenes de compra.")

        started_tx = False
        try:
            if not self.conn.in_transaction:
                self.conn.execute("BEGIN")
                started_tx = True
            generadas = []
            for prov, detalles_oc in preparados:
                numero = self.create_orden_compra(
                    today_text(), req["almacen_codigo"], only_digits(prov.get("ruc")), prov.get("tipo_orden") or (req["tipo_orden"] if "tipo_orden" in req.keys() else "NACIONAL"),
                    req["requerimiento"], normalize_text(prov.get("cotizacion") or ""), req["solicitante"], (req["ot"] or "").replace("OT-", ""),
                    req["centro_costo_codigo"], f"OC GENERADA DESDE {req['requerimiento']}", normalize_text(prov.get("moneda") or "SOLES"),
                    bool(prov.get("precio_incluye_igv", precio_incluye_igv)), detalles_oc, usuario,
                    forma_pago=prov.get("forma_pago") or "CONTADO",
                    tipo_cambio=prov.get("tipo_cambio"),
                    tipo_cambio_fuente=prov.get("tipo_cambio_fuente") or "",
                    tipo_cambio_fecha=prov.get("tipo_cambio_fecha") or "",
                    commit=False,
                )
                generadas.append(numero)
            self.conn.commit()
            return generadas
        except Exception:
            if self.conn.in_transaction or started_tx:
                self.conn.rollback()
            raise

    def get_orden_compra(self, numero):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM orden_compra WHERE UPPER(numero)=UPPER(?)", (normalize_text(numero),))
        oc = cur.fetchone()
        if not oc:
            return None, []
        cur.execute("SELECT * FROM orden_compra_detalle WHERE oc_id=? ORDER BY item", (oc["id"],))
        return oc, cur.fetchall()

    def ultimo_precio_compra(self, codigo, ruc=None):
        cur = self.conn.cursor()
        codigo = normalize_text(codigo)
        params = [codigo]
        sql = """
            SELECT d.precio_unitario_sin_igv AS precio, oc.moneda, oc.tipo_cambio_venta, oc.fecha_orden, oc.numero, oc.ruc
            FROM orden_compra_detalle d
            JOIN orden_compra oc ON oc.id = d.oc_id
            WHERE d.codigo=? AND d.precio_unitario_sin_igv>0 AND oc.estado NOT IN ('ANULADA','DENEGADA')
        """
        if ruc:
            sql += " AND oc.ruc=?"
            params.append(only_digits(ruc))
        sql += " ORDER BY oc.fecha_orden DESC, oc.id DESC LIMIT 1"
        cur.execute(sql, params)
        return cur.fetchone()


    def list_ordenes_compra(self, term="", estado="TODOS", tipo="TODOS", desde="", hasta="", orden="fecha_desc"):
        cur = self.conn.cursor()
        sql = "SELECT * FROM orden_compra WHERE 1=1"
        params = []
        if term:
            like = f"%{normalize_text(term)}%"
            sql += " AND (UPPER(numero) LIKE ? OR UPPER(razon_social) LIKE ? OR UPPER(ruc) LIKE ? OR UPPER(requerimiento) LIKE ? OR UPPER(cotizacion) LIKE ?)"
            params.extend([like, like, like, like, like])
        estado = normalize_text(estado or "TODOS")
        if estado and estado != "TODOS":
            sql += " AND estado=?"
            params.append(estado)
        tipo = normalize_text(tipo or "TODOS")
        if tipo and tipo != "TODOS":
            sql += " AND tipo_orden=?"
            params.append(tipo)
        if desde:
            sql += " AND fecha_orden>=?"
            params.append(fecha_iso_consulta_from_text(desde))
        if hasta:
            sql += " AND fecha_orden<=?"
            params.append(fecha_iso_consulta_from_text(hasta))
        if orden == "fecha_asc":
            sql += " ORDER BY fecha_orden ASC, numero ASC"
        elif orden == "importe_asc":
            sql += " ORDER BY total ASC"
        elif orden == "importe_desc":
            sql += " ORDER BY total DESC"
        else:
            sql += " ORDER BY fecha_orden DESC, id DESC"
        cur.execute(sql, params)
        return cur.fetchall()

    def list_oc_finanzas(self, term="", estado_pago="TODOS", tipo="TODOS", solo_saldo=False, orden="vencimiento_asc"):
        self.refresh_payment_statuses()
        cur = self.conn.cursor()
        sql = """
            SELECT oc.*,
                   COALESCE((SELECT SUM(monto) FROM orden_compra_pagos p WHERE p.oc_id=oc.id),0) AS pagado,
                   (oc.total - COALESCE((SELECT SUM(monto) FROM orden_compra_pagos p WHERE p.oc_id=oc.id),0)) AS saldo
            FROM orden_compra oc
            WHERE oc.estado NOT IN ('ANULADA','DENEGADA')
        """
        params = []
        if term:
            like = f"%{normalize_text(term)}%"
            sql += " AND (UPPER(oc.numero) LIKE ? OR UPPER(oc.razon_social) LIKE ? OR UPPER(oc.ruc) LIKE ? OR UPPER(oc.requerimiento) LIKE ?)"
            params.extend([like, like, like, like])
        estado_pago = normalize_text(estado_pago or "TODOS")
        if estado_pago and estado_pago != "TODOS":
            sql += " AND oc.estado_pago=?"
            params.append(estado_pago)
        tipo = normalize_text(tipo or "TODOS")
        if tipo and tipo != "TODOS":
            sql += " AND oc.tipo_orden=?"
            params.append(tipo)
        if solo_saldo:
            sql += " AND (oc.total - COALESCE((SELECT SUM(monto) FROM orden_compra_pagos p WHERE p.oc_id=oc.id),0)) > 0.01"
        if orden == "saldo_desc":
            sql += " ORDER BY saldo DESC, oc.fecha_vencimiento_pago ASC"
        elif orden == "fecha_desc":
            sql += " ORDER BY oc.fecha_orden DESC, oc.id DESC"
        else:
            sql += " ORDER BY CASE WHEN IFNULL(oc.fecha_vencimiento_pago,'')='' THEN '9999-12-31' ELSE oc.fecha_vencimiento_pago END ASC, oc.numero ASC"
        cur.execute(sql, params)
        return cur.fetchall()

    def oc_historial(self, oc_id):
        cur = self.conn.cursor()
        cur.execute("SELECT estado_anterior, estado_nuevo, usuario, fecha_hora, detalle FROM orden_compra_estado_historial WHERE oc_id=? ORDER BY id", (oc_id,))
        return cur.fetchall()

    def approve_orden_compra(self, numero, cantidades_aprobadas, usuario, motivo=""):
        oc, det = self.get_orden_compra(numero)
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        if oc["estado"] not in ("PENDIENTE", "EN COTIZACION"):
            raise ValueError("Solo se puede aprobar una OC pendiente o en cotización.")
        motivo = normalize_text(motivo)
        total_solic = 0.0
        total_aprob = 0.0
        try:
            cur = self.conn.cursor()
            for row in det:
                req = float(row["cantidad_solicitada"])
                val = cantidades_aprobadas.get(row["id"], cantidades_aprobadas.get(str(row["id"]), cantidades_aprobadas.get(row["codigo"], req)))
                aprob = float(str(val).replace(",", "."))
                if aprob < 0:
                    raise ValueError(f"Cantidad aprobada negativa en item {row['item']}.")
                if aprob - req > 0.00001:
                    raise ValueError(f"No se puede aprobar más de lo solicitado en item {row['item']}.")
                total_solic += req
                total_aprob += aprob
                cur.execute("UPDATE orden_compra_detalle SET cantidad_aprobada=? WHERE id=?", (aprob, row["id"]))
            if total_aprob <= 0:
                if not motivo:
                    raise ValueError("Ingrese motivo para denegar la OC.")
                estado = "DENEGADA"
            elif abs(total_aprob - total_solic) <= 0.00001:
                estado = "APROBADA"
            else:
                if not motivo:
                    raise ValueError("Ingrese motivo para aprobación parcial.")
                estado = "APROBADA PARCIAL"
            cur.execute("UPDATE orden_compra SET aprobado_por=?, fecha_aprobacion=?, motivo_decision=? WHERE id=?", (usuario, now_text(), motivo, oc["id"]))
            self._set_oc_state(oc["id"], estado, usuario, motivo or "Aprobación completa")
            self.conn.commit()
            return estado
        except Exception:
            self.conn.rollback()
            raise

    def deny_orden_compra(self, numero, motivo, usuario):
        oc, det = self.get_orden_compra(numero)
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        if oc["estado"] not in ("PENDIENTE", "EN COTIZACION"):
            raise ValueError("Solo se puede denegar una OC pendiente o en cotización.")
        motivo = normalize_text(motivo)
        if not motivo:
            raise ValueError("Ingrese motivo para denegar.")
        try:
            self.conn.execute("UPDATE orden_compra_detalle SET cantidad_aprobada=0 WHERE oc_id=?", (oc["id"],))
            self.conn.execute("UPDATE orden_compra SET aprobado_por=?, fecha_aprobacion=?, motivo_decision=? WHERE id=?", (usuario, now_text(), motivo, oc["id"]))
            self._set_oc_state(oc["id"], "DENEGADA", usuario, motivo)
            self.conn.commit()
            return "DENEGADA"
        except Exception:
            self.conn.rollback()
            raise

    def anular_orden_compra(self, numero, motivo, usuario):
        oc, det = self.get_orden_compra(numero)
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        if oc["estado"] in ("ANULADA", "LIQUIDADA"):
            raise ValueError("La OC ya está anulada o liquidada.")
        if any(float(r["cantidad_ingresada"] or 0) > 0 for r in det):
            raise ValueError("No se puede anular una OC con ingresos de almacén. Revierta primero los vales de ingreso.")
        motivo = normalize_text(motivo)
        if not motivo:
            raise ValueError("Ingrese motivo de anulación.")
        try:
            self._set_oc_state(oc["id"], "ANULADA", usuario, motivo)
            self.conn.commit()
            return "ANULADA"
        except Exception:
            self.conn.rollback()
            raise

    def registrar_pago_oc(self, numero, fecha_pago_texto, numero_operacion, banco, monto, observacion, usuario, numero_factura=""):
        oc, det = self.get_orden_compra(numero)
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        # Finanzas se controla de forma independiente al avance operativo.
        # Una OC atendida/liquidada a crédito puede seguir pendiente de pago.
        if oc["estado"] in ("PENDIENTE", "EN COTIZACION", "DENEGADA", "ANULADA"):
            raise ValueError("La OC debe estar aprobada para registrar pago.")
        fecha_iso = fecha_iso_from_text(fecha_pago_texto)
        numero_operacion = normalize_text(numero_operacion)
        banco = normalizar_banco(banco)
        observacion = normalize_text(observacion)
        numero_factura = normalize_text(numero_factura)
        monto = float(str(monto or 0).replace(",", "."))
        if not numero_operacion:
            raise ValueError("Ingrese número de operación de pago.")
        if not banco or (banco_codigo_guardado(banco) not in BANCOS_MAP and not banco.startswith("05 OTROS")):
            raise ValueError("Banco inválido.")
        if monto <= 0:
            raise ValueError("El monto pagado debe ser mayor a cero.")
        cur = self.conn.cursor()
        cur.execute("SELECT COALESCE(SUM(monto),0) AS pagado FROM orden_compra_pagos WHERE oc_id=?", (oc["id"],))
        pagado_actual = float(cur.fetchone()["pagado"] or 0)
        if pagado_actual + monto - float(oc["total"] or 0) > 0.01:
            raise ValueError("El pago acumulado no puede superar el total de la OC.")
        try:
            cur.execute("""
                INSERT INTO orden_compra_pagos(oc_id, fecha_pago, numero_operacion, numero_factura, banco, monto, observacion, usuario, creado_en)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (oc["id"], fecha_iso, numero_operacion, numero_factura, banco, monto, observacion, usuario, now_text()))
            estado_pago, pagado_nuevo, saldo = self._sync_oc_payment_status(oc["id"], commit=False)
            factura_txt = f" factura {numero_factura}" if numero_factura else ""
            self._add_oc_historial(oc["id"], oc["estado"], oc["estado"], usuario, f"Pago financiero {estado_pago}: operación {numero_operacion}{factura_txt} por {monto:g}. Saldo {saldo:.2f}")
            self.audit("orden_compra", oc["numero"], "PAGO", usuario, f"Estado pago {estado_pago}: operación {numero_operacion}{factura_txt} por {monto:g}. Avance operativo se mantiene en {oc['estado']} {oc['avance']}%", commit=False)
            self.conn.commit()
            return estado_pago
        except Exception:
            self.conn.rollback()
            raise

    def pagos_oc(self, oc_id):
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM orden_compra_pagos WHERE oc_id=? ORDER BY fecha_pago, id", (oc_id,))
        return cur.fetchall()

    def fecha_registro_entrada_oc(self, oc_id):
        cur = self.conn.cursor()
        cur.execute("SELECT MIN(fecha_recepcion) AS fecha FROM orden_compra_recepciones WHERE oc_id=?", (oc_id,))
        r = cur.fetchone()
        return r["fecha"] if r and r["fecha"] else ""

    def registrar_transito_oc(self, numero, operador, fecha_despacho_texto, fecha_estimada_texto, guia_transportista, observacion, usuario):
        oc, det = self.get_orden_compra(numero)
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        if oc["tipo_orden"] == "SERVICIO":
            raise ValueError("Una OC de servicio no debe pasar por tránsito de almacén.")
        if oc["estado"] not in ("APROBADA", "APROBADA PARCIAL", "PAGO PARCIAL", "PAGADA"):
            raise ValueError("Solo puede poner en tránsito una OC aprobada o pagada.")
        operador = normalize_text(operador)
        if operador not in OPERADORES_LOGISTICOS:
            raise ValueError("Operador logístico inválido.")
        fecha_despacho = fecha_iso_from_text(fecha_despacho_texto)
        fecha_estimada = fecha_iso_consulta_from_text(fecha_estimada_texto) if fecha_estimada_texto else ""
        try:
            self.conn.execute("""
                INSERT INTO orden_compra_transito(oc_id, operador_logistico, fecha_despacho, fecha_estimada_llegada, guia_transportista, observacion, usuario, creado_en)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (oc["id"], operador, fecha_despacho, fecha_estimada, normalize_text(guia_transportista), normalize_text(observacion), usuario, now_text()))
            self._set_oc_state(oc["id"], "EN TRANSITO", usuario, f"Operador {operador}")
            self.conn.commit()
            return "EN TRANSITO"
        except Exception:
            self.conn.rollback()
            raise

    def recepcionar_orden_compra(self, numero, fecha_texto, documento_recepcion, observacion, cantidades_recibir, usuario):
        oc, det = self.get_orden_compra(numero)
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        if oc["tipo_orden"] == "SERVICIO":
            raise ValueError("Una orden de servicio no ingresa al Kardex.")
        if oc["estado"] not in ("APROBADA", "APROBADA PARCIAL", "PAGO PARCIAL", "PAGADA", "EN TRANSITO", "ATENDIDA PARCIALMENTE"):
            raise ValueError("La OC no está disponible para ingreso a almacén.")
        documento_recepcion = normalize_text(documento_recepcion)
        observacion = normalize_text(observacion) or f"INGRESO POR OC {numero}"
        if not documento_recepcion:
            raise ValueError("Ingrese factura o guía de remisión.")
        detalles_mov = []
        recep_rows = []
        total_recibir = 0.0
        total_pendiente_final = 0.0
        for row in det:
            aprobado = float(row["cantidad_aprobada"] or 0)
            ingresado = float(row["cantidad_ingresada"] or 0)
            pendiente = aprobado - ingresado
            if pendiente <= 0:
                continue
            val = cantidades_recibir.get(row["id"], cantidades_recibir.get(str(row["id"]), cantidades_recibir.get(row["codigo"], 0)))
            recibir = float(str(val or 0).replace(",", "."))
            if recibir < 0:
                raise ValueError(f"Cantidad recibida negativa en item {row['item']}.")
            if recibir - pendiente > 0.00001:
                raise ValueError(f"No puede recibir más de lo pendiente en item {row['item']}. Pendiente: {pendiente:g}.")
            if recibir > 0:
                # Para valorización de almacén, el ingreso por OC debe llevar el costo unitario
                # sin IGV en soles. Si la OC está en dólares, se convierte usando el TC congelado
                # de la orden. Esto evita ingresos valorizables con precio cero en el reporte mensual.
                precio_oc = float(row["precio_unitario_sin_igv"] or 0)
                if normalize_text(oc["moneda"] or "SOLES") == "DOLARES":
                    precio_oc = round(precio_oc * float(oc["tipo_cambio"] or 1), 6)
                detalles_mov.append({"item": len(detalles_mov)+1, "codigo": row["codigo"], "cantidad": recibir, "precio_unitario_sin_igv": precio_oc})
                recep_rows.append((row, recibir))
                total_recibir += recibir
            total_pendiente_final += (pendiente - recibir)
        if total_recibir <= 0:
            raise ValueError("Ingrese al menos una cantidad recibida mayor a cero.")
        try:
            vale = self.save_movimiento(fecha_texto, "CN", f"{numero}-{documento_recepcion}", (oc["ot"] or "").replace("OT-", ""), oc["centro_costo_codigo"], oc["solicitante"], observacion, detalles_mov, usuario, oc["almacen_codigo"], numero, commit=False)
            cur = self.conn.cursor()
            cur.execute("""
                INSERT INTO orden_compra_recepciones(oc_id, vale, documento_recepcion, fecha_recepcion, observacion, usuario, creado_en)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (oc["id"], vale, documento_recepcion, fecha_iso_from_text(fecha_texto), observacion, usuario, now_text()))
            recep_id = cur.lastrowid
            for row, recibir in recep_rows:
                cur.execute("UPDATE orden_compra_detalle SET cantidad_ingresada = cantidad_ingresada + ? WHERE id=?", (recibir, row["id"]))
                cur.execute("""
                    INSERT INTO orden_compra_recepcion_detalle(recepcion_id, oc_detalle_id, codigo, descripcion, unidad_medida, cantidad)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (recep_id, row["id"], row["codigo"], row["descripcion"], row["unidad_medida"], recibir))
            estado = "ATENDIDA" if total_pendiente_final <= 0.00001 else "ATENDIDA PARCIALMENTE"
            self._set_oc_state(oc["id"], estado, usuario, f"Recepción {documento_recepcion} | Vale {vale}")
            self.conn.commit()
            return vale, estado
        except Exception:
            self.conn.rollback()
            raise

    def liquidar_orden_compra(self, numero, fecha_texto, factura, guia, observacion, usuario):
        oc, det = self.get_orden_compra(numero)
        if not oc:
            raise ValueError("No se encontró la orden de compra.")
        if oc["estado"] in ("ANULADA", "DENEGADA", "LIQUIDADA"):
            raise ValueError("Esta OC no puede liquidarse.")
        if oc["tipo_orden"] == "NACIONAL" and oc["estado"] not in ("ATENDIDA", "ATENDIDA PARCIALMENTE"):
            raise ValueError("Una OC nacional solo puede liquidarse después de registrar ingreso de almacén.")
        if oc["tipo_orden"] == "SERVICIO" and oc["estado"] not in ("APROBADA", "APROBADA PARCIAL", "PAGO PARCIAL", "PAGADA"):
            raise ValueError("Una orden de servicio debe estar aprobada o pagada para liquidarse.")
        factura = normalize_text(factura)
        guia = normalize_text(guia)
        observacion = normalize_text(observacion)
        if not observacion:
            raise ValueError("Ingrese observación/conformidad de liquidación.")
        try:
            self.conn.execute("""
                INSERT INTO orden_compra_liquidacion(oc_id, factura, guia, fecha_liquidacion, observacion, usuario, creado_en)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (oc["id"], factura, guia, fecha_iso_from_text(fecha_texto), observacion, usuario, now_text()))
            self._set_oc_state(oc["id"], "LIQUIDADA", usuario, observacion)
            self.conn.commit()
            return "LIQUIDADA"
        except sqlite3.IntegrityError:
            self.conn.rollback()
            raise ValueError("Esta OC ya fue liquidada.")
        except Exception:
            self.conn.rollback()
            raise

    def reporte_oc_rows(self, tipo="TODOS", estado="TODOS"):
        return self.list_ordenes_compra(term="", estado=estado, tipo=tipo, orden="fecha_asc")

    def audit_rows(self):
        cur = self.conn.cursor()
        cur.execute("SELECT tabla, registro, accion, usuario, fecha_hora, detalle FROM auditoria ORDER BY id DESC LIMIT 1000")
        return cur.fetchall()


class SplashScreen(tk.Toplevel):
    """Pantalla simple de carga para evitar que el usuario piense que el sistema se quedó colgado."""
    def __init__(self, root):
        super().__init__(root)
        self.title("Cargando sistema")
        self.configure(bg="#f4f4f4")
        self.resizable(False, False)
        self.geometry("520x280+420+220")
        self.overrideredirect(True)

        frame = tk.Frame(self, bg="#f4f4f4", bd=2, relief="ridge")
        frame.pack(fill="both", expand=True, padx=2, pady=2)

        tk.Label(frame, text="SISTEMA KARDEX DE ALMACÉN", bg="#f4f4f4", fg="#111111", font=("Arial", 18, "bold")).pack(pady=(34, 8))
        tk.Label(frame, text="Cargando módulos de almacén y logística...", bg="#f4f4f4", fg="#333333", font=("Arial", 11)).pack(pady=(0, 16))

        self.status = tk.StringVar(value="Inicializando sistema")
        tk.Label(frame, textvariable=self.status, bg="#f4f4f4", fg="#555555", font=("Arial", 10)).pack(pady=(0, 8))

        self.progress = ttk.Progressbar(frame, mode="indeterminate", length=360)
        self.progress.pack(pady=(0, 18))
        self.progress.start(12)

        tk.Label(frame, text="Espere un momento...", bg="#f4f4f4", fg="#777777", font=("Arial", 9)).pack()
        self.lift()
        self.focus_force()
        try:
            self.attributes("-topmost", True)
            self.after(800, lambda: self.attributes("-topmost", False))
        except Exception:
            pass
        self.update_idletasks()
        self.update()

    def set_status(self, text):
        self.status.set(text)
        self.update_idletasks()
        self.update()

    def close(self):
        try:
            self.progress.stop()
        except Exception:
            pass
        self.destroy()


class LoginDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.result = None
        self.title("Ingreso al sistema")
        self.geometry("440x315+430+180")
        # Login visible: no usar transient sobre ventana principal oculta
        self.update_idletasks()
        self.deiconify()
        self.lift()
        self.focus_force()
        try:
            self.attributes("-topmost", True)
            self.after(800, lambda: self.attributes("-topmost", False))
        except Exception:
            pass
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.build()
        self.update_idletasks()
        self.deiconify()
        self.lift()
        self.focus_force()
        self.grab_set()
        self.usuario_entry.focus_force()

    def build(self):
        frm = ttk.Frame(self, padding=24)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Sistema Kardex de Almacén", font=("Arial", 15, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 18))
        ttk.Label(frm, text="Usuario:").grid(row=1, column=0, sticky="w", pady=8)
        self.usuario = tk.StringVar(value="admin")
        self.password = tk.StringVar()
        self.almacen = tk.StringVar(value="AQP")
        self.usuario_entry = ttk.Entry(frm, textvariable=self.usuario, width=28)
        self.usuario_entry.grid(row=1, column=1, pady=8)
        ttk.Label(frm, text="Contraseña:").grid(row=2, column=0, sticky="w", pady=8)
        pass_entry = ttk.Entry(frm, textvariable=self.password, width=28, show="*")
        pass_entry.grid(row=2, column=1, pady=8)
        ttk.Label(frm, text="Almacén de trabajo:").grid(row=3, column=0, sticky="w", pady=8)
        alm_combo = ttk.Combobox(frm, textvariable=self.almacen, values=list(ALMACENES.keys()), width=25, state="readonly")
        alm_combo.grid(row=3, column=1, pady=8, sticky="w")
        ttk.Button(frm, text="Ingresar", command=self.login).grid(row=4, column=1, sticky="e", pady=18)
        ttk.Label(frm, text="Usuario inicial: admin | Contraseña: admin123", foreground="#666666").grid(row=5, column=0, columnspan=2, sticky="w")
        self.usuario_entry.bind("<Return>", lambda e: pass_entry.focus_set())
        pass_entry.bind("<Return>", lambda e: alm_combo.focus_set())
        alm_combo.bind("<Return>", lambda e: self.login())

    def login(self):
        user = self.db.verify_user(self.usuario.get(), self.password.get())
        if not user:
            messagebox.showerror("Acceso denegado", "Usuario o contraseña incorrectos.", parent=self)
            return
        almacen = normalize_text(self.almacen.get() or "AQP")
        if not self.db.user_can_access_almacen(user, almacen):
            messagebox.showerror("Acceso denegado", "Usted no está autorizado para realizar movimientos para este almacén.", parent=self)
            return
        self.result = {"user": user, "almacen": almacen}
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


class CenterCostDialog(tk.Toplevel):
    def __init__(self, parent, on_select):
        super().__init__(parent)
        self.on_select = on_select
        self.title("Consulta de Centros de Costo")
        self.geometry("440x300+260+160")
        bind_as_child_window(self, self.master if hasattr(self, "master") else None)
        self.build()

    def build(self):
        ttk.Label(self, text="Enter o doble clic selecciona | Esc cierra", padding=8).pack(anchor="w")
        cols = ("codigo", "descripcion")
        box = ttk.Frame(self)
        box.pack(fill="both", expand=True, padx=10, pady=10)
        box.rowconfigure(0, weight=1); box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(box, columns=cols, show="headings", height=8)
        self.tree.heading("codigo", text="Código")
        self.tree.heading("descripcion", text="Centro de Costo")
        self.tree.column("codigo", width=90, anchor="center", stretch=False)
        self.tree.column("descripcion", width=300, anchor="w", stretch=False)
        vsb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        for cod, desc in CENTROS_COSTO.items():
            self.tree.insert("", "end", values=(cod, desc))
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.tree.focus(children[0])
        self.tree.bind("<Double-1>", self.choose)
        self.tree.bind("<Return>", self.choose)
        self.bind("<Escape>", lambda e: self.destroy())
        ttk.Button(self, text="Seleccionar", command=self.choose).pack(side="left", padx=10, pady=6)
        ttk.Button(self, text="Cerrar", command=self.destroy).pack(side="right", padx=10, pady=6)

    def choose(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return "break"
        cod, desc = self.tree.item(sel[0], "values")
        self.on_select(cod, desc)
        self.destroy()
        return "break"


class MovimientoTipoDialog(tk.Toplevel):
    """Ayuda rápida para seleccionar tipo de movimiento sin abrir vales por error."""
    def __init__(self, parent, on_select):
        super().__init__(parent)
        self.on_select = on_select
        self.title("Seleccionar tipo de movimiento")
        self.geometry("520x300+300+180")
        bind_as_child_window(self, parent)
        self.build()

    def build(self):
        ttk.Label(self, text="Seleccione con Enter o doble clic", padding=8).pack(anchor="w")
        cols = ("codigo", "descripcion")
        box = ttk.Frame(self)
        box.pack(fill="both", expand=True, padx=10, pady=8)
        box.rowconfigure(0, weight=1); box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(box, columns=cols, show="headings", height=6)
        self.tree.heading("codigo", text="Tipo"); self.tree.column("codigo", width=90, anchor="center")
        self.tree.heading("descripcion", text="Descripción"); self.tree.column("descripcion", width=380, anchor="w")
        for cod, nombre in TIPOS_MOVIMIENTO.items():
            self.tree.insert("", "end", values=(cod, nombre))
        vsb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns")
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0]); self.tree.focus(children[0])
        self.tree.bind("<Double-1>", self.choose)
        self.tree.bind("<Return>", self.choose)
        ttk.Button(self, text="Seleccionar", command=self.choose).pack(side="left", padx=12, pady=8)
        ttk.Button(self, text="Cerrar", command=self.destroy).pack(side="right", padx=12, pady=8)

    def choose(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return "break"
        codigo = self.tree.item(sel[0], "values")[0]
        self.on_select(codigo)
        self.destroy()
        return "break"


class ValeListDialog(tk.Toplevel):
    """Cuadro único de ayuda para buscar vales generados.
    Corrige el problema de listas vacías por filtros y evita duplicidad de diálogos.
    """
    def __init__(self, parent, db, on_select, almacen_codigo="TODOS", tipo_codigo="TODOS"):
        super().__init__(parent)
        self.db = db
        self.on_select = on_select
        self.almacen_codigo = tk.StringVar(value=almacen_codigo or "TODOS")
        self.tipo_codigo = tk.StringVar(value=tipo_codigo or "TODOS")
        self.title("Consulta de Vales Generados")
        self.geometry("980x520+150+90")
        self.minsize(920, 460)
        bind_as_child_window(self, parent)
        self.term = tk.StringVar()
        self.status = tk.StringVar(value="")
        self.build()
        self.search()
        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(parent)
        self.lift()
        self.focus_force()

    def build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Buscar vale, documento, OT, centro, solicitante o usuario:").grid(row=0, column=0, sticky="w")
        e = ttk.Entry(top, textvariable=self.term, width=42)
        e.grid(row=0, column=1, padx=8, sticky="w")
        e.bind("<KeyRelease>", lambda e: self.search())
        e.bind("<Return>", self.choose)
        ttk.Label(top, text="Almacén:").grid(row=0, column=2, padx=(12, 2), sticky="e")
        alm = ttk.Combobox(top, textvariable=self.almacen_codigo, values=["TODOS"] + list(ALMACENES.keys()), width=9, state="readonly")
        alm.grid(row=0, column=3, sticky="w")
        alm.bind("<<ComboboxSelected>>", lambda e: self.search())
        ttk.Label(top, text="Tipo:").grid(row=0, column=4, padx=(12, 2), sticky="e")
        tipo = ttk.Combobox(top, textvariable=self.tipo_codigo, values=["TODOS"] + list(TIPOS_MOVIMIENTO.keys()), width=9, state="readonly")
        tipo.grid(row=0, column=5, sticky="w")
        tipo.bind("<<ComboboxSelected>>", lambda e: self.search())
        ttk.Button(top, text="Actualizar", command=self.search).grid(row=0, column=6, padx=8)
        ttk.Label(top, text="Enter o doble clic selecciona | Esc cierra", foreground="#666666").grid(row=1, column=0, columnspan=7, sticky="w", pady=(6, 0))
        e.focus_set()

        cols = ("vale", "fecha", "almacen", "tipo", "documento", "ot", "cc", "solicitante", "usuario")
        tree_box = ttk.Frame(self)
        tree_box.pack(fill="both", expand=True, padx=10, pady=8)
        tree_box.rowconfigure(0, weight=1)
        tree_box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_box, columns=cols, show="headings", height=16)
        specs = [
            ("vale", "Vale", 125), ("fecha", "Fecha", 95), ("almacen", "Alm.", 60),
            ("tipo", "Tipo", 60), ("documento", "Documento", 160), ("ot", "OT", 105),
            ("cc", "Centro", 170), ("solicitante", "Solicitante", 170), ("usuario", "Usuario", 105)
        ]
        for col, text, width in specs:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor="center" if col in ["vale", "fecha", "almacen", "tipo"] else "w", stretch=False)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self.choose)
        self.tree.bind("<Return>", self.choose)

        bottom = ttk.Frame(self, padding=(10, 0, 10, 10))
        bottom.pack(fill="x")
        ttk.Label(bottom, textvariable=self.status, foreground="#555555").pack(side="left")
        ttk.Button(bottom, text="Seleccionar", command=self.choose).pack(side="right", padx=5)
        ttk.Button(bottom, text="Imprimir / Ver vale", command=self.print_selected).pack(side="right", padx=5)
        ttk.Button(bottom, text="Cerrar", command=self.destroy).pack(side="right", padx=5)

    def search(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        rows = list(self.db.list_vales(self.term.get(), self.almacen_codigo.get(), self.tipo_codigo.get()))
        first = None
        for r in rows:
            iid = self.tree.insert("", "end", values=(
                r["vale"], fecha_text_from_iso(r["fecha_operacion"]), r["almacen_codigo"],
                r["tipo_codigo"], r["documento"], r["ot"] or "",
                r["centro_costo_nombre"] or r["centro_costo_codigo"], r["solicitante"] or "", r["usuario"] or ""
            ))
            if first is None:
                first = iid
        if first:
            self.tree.selection_set(first)
            self.tree.focus(first)
            self.status.set(f"{len(rows)} vale(s) encontrados.")
        else:
            self.status.set("No hay vales para el filtro actual. Cree un ingreso/salida o limpie los filtros.")

    def choose(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return "break"
        vale = self.tree.item(sel[0], "values")[0]
        self.on_select(vale)
        self.destroy()
        return "break"

    def print_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Vale", "Seleccione un vale para visualizar o imprimir.", parent=self)
            return
        vale = self.tree.item(sel[0], "values")[0]
        mov, det = self.db.get_movimiento(vale)
        if not mov:
            messagebox.showerror("Vale", "No se encontró el vale seleccionado.", parent=self)
            return
        PrintPreviewDialog(self, mov, det)


class StockDialog(tk.Toplevel):
    def __init__(self, app, parent=None, on_select=None):
        super().__init__(parent or app.root)
        self.app = app
        self.db = app.db
        self.on_select = on_select
        self.title("Consulta de Stock")
        self.geometry("1180x560+120+90")
        bind_as_child_window(self, parent or app.root)
        self.build()
        self.search()

    def build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        self.term = tk.StringVar()
        ttk.Label(top, text="Buscar por código o descripción:").pack(side="left")
        e = ttk.Entry(top, textvariable=self.term, width=50)
        e.pack(side="left", padx=8)
        e.bind("<KeyRelease>", lambda e: self.search())
        e.bind("<Return>", self.choose_first)
        e.focus_set()
        ttk.Label(top, text="Enter selecciona | Doble clic selecciona | Esc cierra", foreground="#666").pack(side="left", padx=8)
        cols = ("codigo", "descripcion", "um", "stock_aqp", "comp_aqp", "libre_aqp", "min_aqp", "est_aqp", "stock_min", "comp_min", "libre_min", "min_min", "est_min")
        tree_box = ttk.Frame(self)
        tree_box.pack(fill="both", expand=True, padx=10, pady=10)
        tree_box.rowconfigure(0, weight=1); tree_box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_box, columns=cols, show="headings", height=16)
        specs = [("codigo", "Código", 95), ("descripcion", "Descripción", 300), ("um", "UM", 60), ("stock_aqp", "Stock AQP", 80), ("comp_aqp", "Comp. AQP", 80), ("libre_aqp", "Libre AQP", 80), ("min_aqp", "Mín. AQP", 75), ("est_aqp", "Estado AQP", 90), ("stock_min", "Stock MIN", 80), ("comp_min", "Comp. MIN", 80), ("libre_min", "Libre MIN", 80), ("min_min", "Mín. MIN", 75), ("est_min", "Estado MIN", 90)]
        for c, t, w in specs:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c != "descripcion" else "w", stretch=False)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self.choose)
        self.tree.bind("<Return>", self.choose)
        ttk.Button(self, text="Seleccionar", command=self.choose).pack(side="left", padx=10, pady=8)
        ttk.Button(self, text="Cerrar", command=self.destroy).pack(side="right", padx=10, pady=8)

    def search(self):
        for x in self.tree.get_children():
            self.tree.delete(x)
        first = None
        for r in self.db.search_articulos(self.term.get()):
            iid = self.tree.insert("", "end", values=(r["codigo"], r["descripcion"], r["unidad_medida"], f"{r['stock_aqp']:g}", f"{r['comprometido_aqp']:g}", f"{r['libre_aqp']:g}", f"{r['stock_minimo_aqp']:g}", r["estado_aqp"], f"{r['stock_min']:g}", f"{r['comprometido_min']:g}", f"{r['libre_min']:g}", f"{r['stock_minimo_min']:g}", r["estado_min"]))
            if first is None:
                first = iid
        if first:
            self.tree.selection_set(first)
            self.tree.focus(first)

    def choose_first(self, event=None):
        if self.tree.get_children():
            self.tree.focus_set()
            return self.choose()
        return "break"

    def choose(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return "break"
        codigo = self.tree.item(sel[0], "values")[0]
        if self.on_select:
            self.on_select(codigo)
        self.destroy()
        return "break"


class EntryDialog(tk.Toplevel):
    def __init__(self, app, mode="new", vale=None):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.mode = mode
        self.detalles = []
        self.selected_index = None
        self.items_enabled = False
        self._help_dialogs = {}
        self.title("Ingreso / Salida" if mode == "new" else "Consultar Registro")
        self.geometry("1080x700+100+45")
        bind_as_child_window(self, self.master if hasattr(self, "master") else None)
        self.minsize(1100, 650)
        self.configure(bg="#f4f6f8")
        self.build()
        if self.mode == "new":
            self.after(200, self.tipo_entry.focus_set)
        self.bind("<Shift-F2>", lambda e: self.open_stock_help())
        self.bind("<Shift-F1>", self.shift_f1_handler)
        if vale:
            self.vale_var.set(vale)
            self.buscar_vale()

    def build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        if self.mode == "view":
            ttk.Label(top, text="Número de Vale:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
            self.vale_var = tk.StringVar()
            ent_vale = ttk.Entry(top, textvariable=self.vale_var, width=18)
            ent_vale.grid(row=0, column=1, sticky="w", padx=4)
            ent_vale.bind("<Return>", lambda e: self.buscar_vale())
            ent_vale.bind("<Shift-F1>", lambda e: self.open_vale_list())
            ttk.Button(top, text="Buscar", command=self.buscar_vale).grid(row=0, column=2, padx=4)
            ttk.Label(top, text="Shift + F1: ver lista de vales", foreground="#666666").grid(row=0, column=3, columnspan=3, sticky="w")
        else:
            self.vale_var = tk.StringVar()
        self.fecha_var = tk.StringVar(value=today_text())
        self.fecha_entrega_var = tk.StringVar(value=today_text())
        self.edit_fecha = tk.BooleanVar(value=False)
        self.almacen_var = tk.StringVar(value=getattr(self.app, "almacen_actual", "AQP"))
        self.tipo_var = tk.StringVar()
        self.doc_var = tk.StringVar()
        self.ot_var = tk.StringVar()
        self.cc_code_var = tk.StringVar()
        self.cc_name_var = tk.StringVar()
        self.solicitante_var = tk.StringVar()
        self.obs_var = tk.StringVar()
        row = 1 if self.mode == "view" else 0
        ttk.Label(top, text="Fecha almacén:").grid(row=row, column=0, sticky="w", padx=4, pady=6)
        self.fecha_entry = ttk.Entry(top, textvariable=self.fecha_var, width=13, state="readonly")
        self.fecha_entry.grid(row=row, column=1, sticky="w", padx=4)
        if self.mode == "new":
            ttk.Checkbutton(top, text="Editar fecha", variable=self.edit_fecha, command=self.toggle_fecha).grid(row=row, column=2, sticky="w", padx=4)
        ttk.Label(top, text="Tipo CN/CO/AJ:").grid(row=row, column=3, sticky="w", padx=4, pady=6)
        self.tipo_entry = ttk.Entry(top, textvariable=self.tipo_var, width=8, state="readonly" if self.mode == "view" else "normal")
        self.tipo_entry.grid(row=row, column=4, sticky="w", padx=4)
        self.tipo_entry.bind("<Shift-F1>", lambda e: self.open_tipo_help())
        ttk.Label(top, text="Documento:").grid(row=row, column=5, sticky="w", padx=4, pady=6)
        self.doc_entry = ttk.Entry(top, textvariable=self.doc_var, width=24, state="readonly" if self.mode == "view" else "normal")
        self.doc_entry.grid(row=row, column=6, sticky="w", padx=4)
        ttk.Label(top, text="Almacén:").grid(row=row, column=7, sticky="w", padx=4, pady=6)
        self.almacen_combo = ttk.Combobox(top, textvariable=self.almacen_var, values=list(ALMACENES.keys()), width=6, state="readonly" if self.mode == "view" else "readonly")
        self.almacen_combo.grid(row=row, column=8, sticky="w", padx=4)
        row2 = row + 1
        ttk.Label(top, text="OT-").grid(row=row2, column=0, sticky="w", padx=4, pady=6)
        self.ot_entry = ttk.Entry(top, textvariable=self.ot_var, width=16)
        self.ot_entry.grid(row=row2, column=1, sticky="w", padx=4)
        ttk.Label(top, text="Centro Costo:").grid(row=row2, column=2, sticky="w", padx=4, pady=6)
        self.cc_entry = ttk.Entry(top, textvariable=self.cc_code_var, width=8)
        self.cc_entry.grid(row=row2, column=3, sticky="w", padx=4)
        self.cc_entry.bind("<KeyRelease>", self.fill_cc_from_typed)
        self.cc_entry.bind("<Shift-F1>", lambda e: self.open_cc_help())
        ttk.Entry(top, textvariable=self.cc_name_var, width=27, state="readonly").grid(row=row2, column=4, columnspan=2, sticky="w", padx=4)
        ttk.Label(top, text="Shift + F1 para consultar CC", foreground="#666666").grid(row=row2, column=6, sticky="w")
        row3 = row + 2
        ttk.Label(top, text="Solicitante:").grid(row=row3, column=0, sticky="w", padx=4, pady=6)
        self.solicitante_combo = ttk.Combobox(top, textvariable=self.solicitante_var, values=self.db.solicitante_names(), width=36, state="normal")
        self.solicitante_combo.grid(row=row3, column=1, columnspan=3, sticky="w", padx=4)
        self.solicitante_combo.bind("<Shift-F1>", lambda e: self.open_solicitante_help())
        attach_solicitante_autocomplete(self.solicitante_combo, self.db, self.solicitante_var, self.set_solicitante_from_help)
        ttk.Label(top, text="Puede escribir o seleccionar de la lista | Shift + F1 busca", foreground="#666666").grid(row=row3, column=4, columnspan=3, sticky="w")
        row4 = row + 3
        ttk.Label(top, text="Observación:").grid(row=row4, column=0, sticky="w", padx=4, pady=6)
        self.obs_entry = ttk.Entry(top, textvariable=self.obs_var, width=88)
        self.obs_entry.grid(row=row4, column=1, columnspan=6, sticky="we", padx=4)
        # Flujo de teclado: ENTER trabaja como TAB, validando cada etapa.
        # Secuencia: Tipo -> Documento -> OT -> Centro de costo -> Solicitante -> Observación -> Código.
        self.tipo_entry.bind("<Return>", self.enter_tipo)
        self.doc_entry.bind("<Return>", self.enter_doc)
        self.ot_entry.bind("<Return>", self.enter_ot)
        self.cc_entry.bind("<Return>", self.enter_cc)
        self.solicitante_combo.bind("<Return>", self.enter_solicitante)
        self.obs_entry.bind("<Return>", self.enter_obs)
        self.fecha_entry.bind("<Return>", self.focus_next)
        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        columns = ("item", "codigo", "descripcion", "um", "stock", "cantidad", "precio", "subtotal")
        self.tree = ttk.Treeview(body, columns=columns, show="headings", height=14)
        headers = [("item", "Item", 55), ("codigo", "Código", 110), ("descripcion", "Descripción", 360), ("um", "UM", 70), ("stock", "Stock actual", 95), ("cantidad", "Cantidad", 95), ("precio", "P.Unit s/IGV", 110), ("subtotal", "Subtotal s/IGV", 115)]
        for col, text, width in headers:
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor="center" if col != "descripcion" else "w")
        body.rowconfigure(0, weight=1); body.columnconfigure(0, weight=1)
        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(body, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.select_row)
        if self.mode == "new":
            form = ttk.Frame(self, padding=10)
            form.pack(fill="x")
            self.cod_var = tk.StringVar()
            self.desc_var = tk.StringVar()
            self.um_var = tk.StringVar()
            self.qty_var = tk.StringVar()
            self.stock_var = tk.StringVar(value="0")
            self.price_var = tk.StringVar(value="0")
            ttk.Label(form, text="Código:").grid(row=0, column=0, padx=4)
            self.cod_entry = ttk.Entry(form, textvariable=self.cod_var, width=14, state="disabled")
            self.cod_entry.grid(row=0, column=1, padx=4)
            self.cod_entry.bind("<Return>", self.load_code)
            self.cod_entry.bind("<Shift-F1>", lambda e: self.open_stock_help())
            self.cod_entry.bind("<Shift-F2>", lambda e: self.open_stock_help())
            attach_articulo_autocomplete(self.cod_entry, self.db, self.cod_var, self.set_code_from_stock)
            ttk.Label(form, text="Descripción:").grid(row=0, column=2, padx=4)
            ttk.Entry(form, textvariable=self.desc_var, width=45, state="readonly").grid(row=0, column=3, padx=4)
            ttk.Label(form, text="UM:").grid(row=0, column=4, padx=4)
            ttk.Entry(form, textvariable=self.um_var, width=8, state="readonly").grid(row=0, column=5, padx=4)
            ttk.Label(form, text="Stock:").grid(row=0, column=6, padx=4)
            ttk.Entry(form, textvariable=self.stock_var, width=10, state="readonly").grid(row=0, column=7, padx=4)
            ttk.Label(form, text="Cantidad:").grid(row=0, column=8, padx=4)
            self.qty_entry = ttk.Entry(form, textvariable=self.qty_var, width=10, state="disabled")
            self.qty_entry.grid(row=0, column=9, padx=4)
            self.qty_entry.bind("<Return>", lambda e: self.price_entry.focus_set())
            ttk.Label(form, text="Precio s/IGV S/:").grid(row=0, column=10, padx=4)
            self.price_entry = ttk.Entry(form, textvariable=self.price_var, width=12, state="disabled")
            self.price_entry.grid(row=0, column=11, padx=4)
            self.price_entry.bind("<Return>", self.add_item)
            ttk.Label(form, text="Para CO el precio se jala del último costo valorizado y queda bloqueado. Para CN/AJ puede registrarse manualmente.", foreground="#666666").grid(row=1, column=1, columnspan=11, sticky="w", pady=4)
        buttons = ttk.Frame(self, padding=10)
        buttons.pack(fill="x")
        if self.mode == "new":
            ttk.Button(buttons, text="✖", width=5, command=self.delete_item).pack(side="left", padx=5)
            ttk.Button(buttons, text="📖", width=5, command=self.modify_item).pack(side="left", padx=5)
            ttk.Button(buttons, text="Consultar Stock Shift+F2", command=self.open_stock_help).pack(side="right", padx=5)
            ttk.Button(buttons, text="💾 Guardar", command=self.save).pack(side="right", padx=5)
        else:
            ttk.Button(buttons, text="💾 Guardar modificación", command=self.save_mod).pack(side="right", padx=5)

    def focus_next(self, event):
        event.widget.tk_focusNext().focus()
        return "break"

    def enter_tipo(self, event=None):
        tipo = normalize_text(self.tipo_var.get())
        if tipo not in TIPOS_MOVIMIENTO:
            messagebox.showwarning("Tipo inválido", "Ingrese un tipo válido: CN = Ingreso, CO = Salida, AJ = Ajuste.", parent=self)
            self.after(100, self.tipo_entry.focus_set)
            return "break"
        self.tipo_var.set(tipo)
        self.doc_entry.focus_set()
        return "break"

    def enter_doc(self, event=None):
        self.doc_var.set(self.doc_var.get().strip().upper())
        self.ot_entry.focus_set()
        return "break"

    def enter_ot(self, event=None):
        # La OT no define centro de costo. Puede quedar vacía.
        self.ot_var.set(self.ot_var.get().strip().upper().replace("OT-", ""))
        self.cc_entry.focus_set()
        return "break"

    def enter_cc(self, event=None):
        self.fill_cc_from_typed()
        if not self.cc_name_var.get():
            messagebox.showwarning("Centro de costo", "Ingrese un centro de costo válido. Ejemplo: 01, 02, 03.\nTambién puede presionar Shift + F1 para consultar.", parent=self)
            self.after(100, self.cc_entry.focus_set)
            try:
                self.cc_entry.selection_range(0, tk.END)
            except Exception:
                pass
            return "break"
        self.solicitante_combo.focus_set()
        return "break"

    def enter_solicitante(self, event=None):
        self.solicitante_var.set(self.solicitante_var.get().strip().upper())
        if not self.solicitante_var.get():
            messagebox.showwarning("Solicitante obligatorio", "Ingrese el nombre del solicitante antes de continuar.", parent=self)
            self.after(100, self.solicitante_combo.focus_set)
            return "break"
        self.obs_entry.focus_set()
        return "break"

    def enter_obs(self, event=None):
        if not self.obs_var.get().strip():
            messagebox.showwarning("Observación obligatoria", "Ingrese una observación antes de registrar artículos.", parent=self)
            self.after(100, self.obs_entry.focus_set)
            return "break"
        self.obs_var.set(self.obs_var.get().strip().upper())
        self.enable_item_inputs()
        self.cod_entry.focus_set()
        return "break"

    def enable_item_inputs(self):
        if self.mode != "new":
            return
        self.items_enabled = True
        self.cod_entry.configure(state="normal")
        self.qty_entry.configure(state="normal")
        self.price_entry.configure(state="normal")

    def disable_item_inputs(self):
        if self.mode != "new":
            return
        self.items_enabled = False
        self.cod_var.set("")
        self.desc_var.set("")
        self.um_var.set("")
        self.qty_var.set("")
        self.stock_var.set("0")
        self.price_var.set("0")
        self.cod_entry.configure(state="disabled")
        self.qty_entry.configure(state="disabled")
        self.price_entry.configure(state="disabled")

    def toggle_fecha(self):
        self.fecha_entry.configure(state="normal" if self.edit_fecha.get() else "readonly")
        if not self.edit_fecha.get():
            self.fecha_var.set(today_text())

    def shift_f1_handler(self, event=None):
        widget = self.focus_get()
        if widget == getattr(self, "tipo_entry", None):
            return self.open_tipo_help()
        if widget == getattr(self, "cc_entry", None):
            return self.open_cc_help()
        if widget == getattr(self, "solicitante_combo", None):
            return self.open_solicitante_help()
        if widget == getattr(self, "cod_entry", None):
            return self.open_stock_help()
        # En consulta de registros, Shift+F1 abre vales solo cuando no está el foco en otra ayuda.
        if self.mode == "view":
            return self.open_vale_list()
        return "break"

    def _open_unique_help(self, key, factory):
        win = self._help_dialogs.get(key)
        try:
            if win is not None and win.winfo_exists():
                win.lift()
                win.focus_force()
                return "break"
        except Exception:
            pass
        win = factory()
        self._help_dialogs[key] = win
        def _clear(event=None, k=key, w=win):
            if self._help_dialogs.get(k) is w:
                self._help_dialogs.pop(k, None)
        try:
            win.bind("<Destroy>", _clear, add="+")
        except Exception:
            pass
        return "break"

    def open_tipo_help(self):
        return self._open_unique_help("tipo", lambda: MovimientoTipoDialog(self, self.set_tipo_from_help))

    def set_tipo_from_help(self, codigo):
        self.tipo_var.set(codigo)
        self.doc_entry.focus_set()

    def open_cc_help(self):
        return self._open_unique_help("cc", lambda: CenterCostDialog(self, self.set_cc))

    def open_solicitante_help(self):
        return self._open_unique_help("solicitante", lambda: SolicitanteSelectDialog(self, self.db, self.set_solicitante_from_help))

    def set_solicitante_from_help(self, nombre):
        self.solicitante_var.set(nombre)
        self.obs_entry.focus_set()

    def open_stock_help(self):
        return self._open_unique_help("stock", lambda: StockDialog(self.app, parent=self, on_select=self.set_code_from_stock))

    def set_code_from_stock(self, codigo):
        self.cod_var.set(codigo)
        self.load_code()

    def set_cc(self, codigo, nombre):
        self.cc_code_var.set(codigo)
        self.cc_name_var.set(nombre)
        self.solicitante_combo.focus_set()

    def fill_cc_from_typed(self, event=None):
        code = normalize_text(self.cc_code_var.get())
        if code in CENTROS_COSTO:
            self.cc_code_var.set(code)
            self.cc_name_var.set(CENTROS_COSTO[code])
        else:
            self.cc_name_var.set("")

    def open_vale_list(self):
        return self._open_unique_help("vales", lambda: ValeListDialog(self, self.db, self.set_vale_from_list, almacen_codigo="TODOS", tipo_codigo="TODOS"))

    def set_vale_from_list(self, vale):
        self.vale_var.set(vale)
        self.buscar_vale()

    def load_code(self, event=None):
        if not self.items_enabled:
            messagebox.showwarning("Complete cabecera", "Primero complete Tipo, Documento, OT, Centro de costo, Solicitante y Observación.\nLuego podrá ingresar los códigos.", parent=self)
            self.after(100, self.obs_entry.focus_set)
            return "break"
        art = self.db.get_articulo(self.cod_var.get())
        if not art:
            self.desc_var.set("")
            self.um_var.set("")
            self.qty_var.set("")
            messagebox.showwarning("No existe", "El código no existe. Puede registrarlo en Registrar Artículo.\nEl registro actual continuará abierto.", parent=self)
            self.after(100, self.cod_entry.focus_set)
            try:
                self.cod_entry.selection_range(0, tk.END)
            except Exception:
                pass
            return "break"
        self.cod_var.set(art["codigo"])
        self.desc_var.set(art["descripcion"])
        self.um_var.set(art["unidad_medida"])
        stock = self.db.get_stock(art["codigo"], self.almacen_var.get())
        self.stock_var.set(f"{stock:g}")
        tipo = normalize_text(self.tipo_var.get())
        if tipo == "CO":
            precio = self.db.ultimo_costo_articulo(art["codigo"], self.almacen_var.get())
            self.price_var.set(f"{precio:.2f}" if precio else "0.00")
            self.price_entry.configure(state="readonly")
        else:
            self.price_entry.configure(state="normal")
        self.qty_entry.focus_set()
        return "break"

    def add_item(self, event=None):
        if not self.items_enabled:
            return self.load_code()
        if not self.desc_var.get():
            self.load_code()
            if not self.desc_var.get():
                return "break"
        try:
            qty = float(self.qty_var.get().replace(",", "."))
            tipo = normalize_text(self.tipo_var.get())
            if tipo != "AJ" and qty <= 0:
                raise ValueError
            if tipo == "AJ" and qty == 0:
                raise ValueError
            price = float(str(self.price_var.get() or 0).replace(",", "."))
            if price < 0:
                raise ValueError
        except Exception:
            messagebox.showwarning("Cantidad / precio inválido", "Cantidad o precio inválido. El precio debe ser mayor o igual a cero y sin IGV en soles.", parent=self)
            return "break"
        item = len(self.detalles) + 1
        d = {"item": item, "codigo": normalize_text(self.cod_var.get()), "descripcion": self.desc_var.get(), "um": self.um_var.get(), "stock_actual": float(str(self.stock_var.get() or 0).replace(",", ".")), "cantidad": qty, "precio_unitario_sin_igv": price}
        self.detalles.append(d)
        self.refresh_items()
        self.cod_var.set("")
        self.desc_var.set("")
        self.um_var.set("")
        self.stock_var.set("0")
        self.qty_var.set("")
        self.price_var.set("0")
        self.cod_entry.focus_set()
        return "break"

    def refresh_items(self):
        for x in self.tree.get_children():
            self.tree.delete(x)
        for i, d in enumerate(self.detalles, start=1):
            d["item"] = i
            precio = float(d.get("precio_unitario_sin_igv", d.get("precio", 0)) or 0)
            subtotal = abs(float(d["cantidad"] or 0)) * precio
            self.tree.insert("", "end", values=(i, d["codigo"], d["descripcion"], d["um"], f"{float(d.get('stock_actual', self.db.get_stock(d['codigo'], self.almacen_var.get()))):g}", f"{d['cantidad']:g}", f"{precio:.2f}", f"{subtotal:.2f}"))

    def select_row(self, event=None):
        sel = self.tree.selection()
        self.selected_index = self.tree.index(sel[0]) if sel else None

    def delete_item(self):
        if self.selected_index is None:
            messagebox.showinfo("Seleccione", "Seleccione un item para eliminar.", parent=self)
            self.after(100, self.focus_force)
            return
        if messagebox.askyesno("Eliminar", "¿Eliminar el artículo seleccionado?", parent=self):
            self.detalles.pop(self.selected_index)
            self.selected_index = None
            self.refresh_items()

    def modify_item(self):
        if self.selected_index is None:
            messagebox.showinfo("Seleccione", "Seleccione un item para modificar.", parent=self)
            self.after(100, self.focus_force)
            return
        d = self.detalles[self.selected_index]
        new_qty = simpledialog.askstring("Modificar cantidad", f"Código {d['codigo']} - Nueva cantidad:", initialvalue=str(d["cantidad"]), parent=self)
        if new_qty is None:
            return
        try:
            qty = float(new_qty.replace(",", "."))
            tipo = normalize_text(self.tipo_var.get())
            if tipo != "AJ" and qty <= 0:
                raise ValueError
            if tipo == "AJ" and qty == 0:
                raise ValueError
            d["cantidad"] = qty
            self.refresh_items()
        except Exception:
            messagebox.showwarning("Inválido", "Cantidad inválida.", parent=self)
            self.after(100, self.focus_force)

    def save(self):
        try:
            if not self.app.require_almacen(self, self.almacen_var.get()):
                return
            self.fill_cc_from_typed()
            if not self.cc_name_var.get():
                messagebox.showwarning("Centro de costo", "Ingrese un centro de costo válido antes de guardar.", parent=self)
                self.cc_entry.focus_set()
                return
            if not self.solicitante_var.get().strip():
                messagebox.showwarning("Solicitante obligatorio", "Ingrese el nombre del solicitante antes de guardar.", parent=self)
                self.solicitante_combo.focus_set()
                return
            if not self.obs_var.get().strip():
                messagebox.showwarning("Observación obligatoria", "Ingrese una observación antes de guardar.", parent=self)
                self.obs_entry.focus_set()
                return
            vale = self.db.save_movimiento(self.fecha_var.get(), self.tipo_var.get(), self.doc_var.get(), self.ot_var.get(), self.cc_code_var.get(), self.solicitante_var.get(), self.obs_var.get(), self.detalles, self.app.usuario_actual, self.almacen_var.get())
            messagebox.showinfo("Registro completo", f"Registro guardado correctamente.\nVale generado: {vale}", parent=self)
            if messagebox.askyesno("Imprimir documento", "¿Desea imprimir documento?", parent=self):
                mov, det = self.db.get_movimiento_by_vale(vale)
                if mov:
                    PrintPreviewDialog(self, mov, det)
            self.fecha_var.set(today_text())
            self.edit_fecha.set(False)
            self.toggle_fecha()
            self.almacen_var.set("AQP")
            self.tipo_var.set("")
            self.doc_var.set("")
            self.ot_var.set("")
            self.cc_code_var.set("")
            self.cc_name_var.set("")
            self.solicitante_var.set("")
            self.obs_var.set("")
            self.detalles = []
            self.refresh_items()
            self.disable_item_inputs()
            self.tipo_entry.focus_set()
        except Exception as e:
            messagebox.showerror("No se pudo guardar", str(e), parent=self)
            self.after(100, self.focus_force)

    def buscar_vale(self):
        mov, det = self.db.get_movimiento_by_vale(self.vale_var.get())
        if not mov:
            messagebox.showwarning("No encontrado", "No se encontró el vale.", parent=self)
            self.after(100, self.focus_force)
            return
        self.fecha_var.set(fecha_text_from_iso(mov["fecha_operacion"]))
        self.tipo_var.set(mov["tipo_codigo"])
        self.doc_var.set(mov["documento"])
        self.almacen_var.set(mov["almacen_codigo"] if "almacen_codigo" in mov.keys() else "AQP")
        self.ot_var.set((mov["ot"] or "").replace("OT-", ""))
        self.cc_code_var.set(mov["centro_costo_codigo"])
        self.cc_name_var.set(mov["centro_costo_nombre"])
        self.solicitante_var.set(mov["solicitante"])
        self.obs_var.set(mov["observacion"])
        self.detalles = [{"item": r["item"], "codigo": r["codigo"], "descripcion": r["descripcion"], "um": r["unidad_medida"], "stock_actual": self.db.get_stock(r["codigo"], mov["almacen_codigo"] if "almacen_codigo" in mov.keys() else None), "cantidad": r["cantidad"], "precio_unitario_sin_igv": (r["precio_unitario_sin_igv"] if "precio_unitario_sin_igv" in r.keys() else 0)} for r in det]
        self.refresh_items()

    def save_mod(self):
        try:
            if not self.app.require_write(self):
                return
            self.fill_cc_from_typed()
            self.db.update_movimiento_header(self.vale_var.get(), self.ot_var.get(), self.cc_code_var.get(), self.solicitante_var.get(), self.obs_var.get(), self.app.usuario_actual)
            messagebox.showinfo("Modificación Completa", "Modificación Completa", parent=self)
            self.vale_var.set("")
            self.fecha_var.set("")
            self.almacen_var.set("AQP")
            self.tipo_var.set("")
            self.doc_var.set("")
            self.ot_var.set("")
            self.cc_code_var.set("")
            self.cc_name_var.set("")
            self.solicitante_var.set("")
            self.obs_var.set("")
            self.detalles = []
            self.refresh_items()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)
            self.after(100, self.focus_force)


class TransferDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.detalles = []
        self.selected_index = None
        self.title("Transferencia entre Almacenes")
        self.geometry("1040x650+110+60")
        self.minsize(980, 600)
        bind_as_child_window(self, app.root)
        self.build()

    def build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        self.fecha_var = tk.StringVar(value=today_text())
        self.origen_var = tk.StringVar(value="AQP")
        self.destino_var = tk.StringVar(value="MIN")
        self.cc_code_var = tk.StringVar(value="01")
        self.cc_name_var = tk.StringVar(value=CENTROS_COSTO.get("01", ""))
        self.solicitante_var = tk.StringVar()
        self.obs_var = tk.StringVar(value="TRANSFERENCIA ENTRE ALMACENES")
        self.guia_var = tk.StringVar()
        ttk.Label(top, text="Fecha:").grid(row=0, column=0, sticky="w", padx=4, pady=6)
        ttk.Entry(top, textvariable=self.fecha_var, width=13).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(top, text="Origen:").grid(row=0, column=2, sticky="w", padx=4)
        ttk.Combobox(top, textvariable=self.origen_var, values=list(ALMACENES.keys()), width=8, state="readonly").grid(row=0, column=3, sticky="w", padx=4)
        ttk.Label(top, text="Destino:").grid(row=0, column=4, sticky="w", padx=4)
        ttk.Combobox(top, textvariable=self.destino_var, values=list(ALMACENES.keys()), width=8, state="readonly").grid(row=0, column=5, sticky="w", padx=4)
        ttk.Label(top, text="CC:").grid(row=1, column=0, sticky="w", padx=4, pady=6)
        self.cc_entry = ttk.Entry(top, textvariable=self.cc_code_var, width=8)
        self.cc_entry.grid(row=1, column=1, sticky="w", padx=4)
        self.cc_entry.bind("<KeyRelease>", self.fill_cc_from_typed)
        self.cc_entry.bind("<Shift-F1>", lambda e: self.open_cc_help())
        ttk.Entry(top, textvariable=self.cc_name_var, width=28, state="readonly").grid(row=1, column=2, columnspan=2, sticky="w", padx=4)
        ttk.Label(top, text="Responsable/Solicitante:").grid(row=1, column=4, sticky="w", padx=4)
        self.solicitante_combo = ttk.Combobox(top, textvariable=self.solicitante_var, values=self.db.solicitante_names(), width=30, state="normal")
        self.solicitante_combo.grid(row=1, column=5, columnspan=2, sticky="w", padx=4)
        self.solicitante_combo.bind("<Shift-F1>", lambda e: self.open_solicitante_help())
        attach_solicitante_autocomplete(self.solicitante_combo, self.db, self.solicitante_var, self.set_solicitante_from_help)
        ttk.Label(top, text="Nro. guía:").grid(row=2, column=0, sticky="w", padx=4, pady=6)
        ttk.Entry(top, textvariable=self.guia_var, width=24).grid(row=2, column=1, sticky="w", padx=4)
        ttk.Label(top, text="Observación:").grid(row=2, column=2, sticky="w", padx=4, pady=6)
        ttk.Entry(top, textvariable=self.obs_var, width=72).grid(row=2, column=3, columnspan=4, sticky="we", padx=4)
        ttk.Label(top, text="Shift+F1 en CC o Código para buscar", foreground="#666").grid(row=3, column=0, columnspan=7, sticky="w", padx=4)

        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        cols = ("item", "codigo", "descripcion", "um", "cantidad", "stock_origen")
        self.tree = ttk.Treeview(body, columns=cols, show="headings", height=14)
        specs = [("item", "Item", 60), ("codigo", "Código", 120), ("descripcion", "Descripción", 470), ("um", "UM", 80), ("cantidad", "Cantidad", 100), ("stock_origen", "Stock origen", 120)]
        for c, t, w in specs:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c != "descripcion" else "w")
        body.rowconfigure(0, weight=1); body.columnconfigure(0, weight=1)
        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(body, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.select_row)

        form = ttk.Frame(self, padding=10)
        form.pack(fill="x")
        self.cod_var = tk.StringVar()
        self.desc_var = tk.StringVar()
        self.um_var = tk.StringVar()
        self.qty_var = tk.StringVar()
        ttk.Label(form, text="Código:").grid(row=0, column=0, padx=4)
        self.cod_entry = ttk.Entry(form, textvariable=self.cod_var, width=14)
        self.cod_entry.grid(row=0, column=1, padx=4)
        self.cod_entry.bind("<Return>", self.load_code)
        self.cod_entry.bind("<Shift-F1>", lambda e: self.open_articulo_help())
        self.cod_entry.bind("<Shift-F2>", lambda e: self.open_stock_help())
        attach_articulo_autocomplete(self.cod_entry, self.db, self.cod_var, self.set_codigo_from_help)
        ttk.Label(form, text="Descripción:").grid(row=0, column=2, padx=4)
        ttk.Entry(form, textvariable=self.desc_var, width=42, state="readonly").grid(row=0, column=3, padx=4)
        ttk.Label(form, text="UM:").grid(row=0, column=4, padx=4)
        ttk.Entry(form, textvariable=self.um_var, width=8, state="readonly").grid(row=0, column=5, padx=4)
        ttk.Label(form, text="Cantidad:").grid(row=0, column=6, padx=4)
        self.qty_entry = ttk.Entry(form, textvariable=self.qty_var, width=12)
        self.qty_entry.grid(row=0, column=7, padx=4)
        self.qty_entry.bind("<Return>", self.add_item)

        buttons = ttk.Frame(self, padding=10)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="✖ Quitar", command=self.delete_item).pack(side="left", padx=5)
        ttk.Button(buttons, text="📖 Modificar cantidad", command=self.modify_item).pack(side="left", padx=5)
        ttk.Button(buttons, text="💾 Guardar transferencia", command=self.save).pack(side="right", padx=5)
        ttk.Button(buttons, text="Consultar Stock Shift+F2", command=lambda: StockDialog(self.app, parent=self, on_select=self.set_codigo_from_help)).pack(side="right", padx=5)
        self.cod_entry.focus_set()

    def open_cc_help(self):
        CenterCostDialog(self, self.set_cc_from_help)
        return "break"

    def open_solicitante_help(self):
        SolicitanteSelectDialog(self, self.db, self.set_solicitante_from_help)
        return "break"

    def set_solicitante_from_help(self, nombre):
        self.solicitante_var.set(nombre)
        self.cod_entry.focus_set()

    def open_articulo_help(self):
        ArticuloSelectDialog(self, self.db, self.set_codigo_from_help)
        return "break"

    def open_stock_help(self):
        StockDialog(self.app, parent=self, on_select=self.set_codigo_from_help)
        return "break"

    def fill_cc_from_typed(self, event=None):
        code = normalize_text(self.cc_code_var.get())
        self.cc_name_var.set(CENTROS_COSTO.get(code, ""))

    def set_cc_from_help(self, codigo, nombre):
        self.cc_code_var.set(codigo)
        self.cc_name_var.set(nombre)
        self.solicitante_combo.focus_set()

    def set_codigo_from_help(self, codigo):
        self.cod_var.set(codigo)
        self.load_code()

    def load_code(self, event=None):
        art = self.db.get_articulo(self.cod_var.get())
        if not art:
            self.desc_var.set(""); self.um_var.set(""); self.qty_var.set("")
            messagebox.showwarning("No existe", "El código no existe.", parent=self)
            self.cod_entry.focus_set()
            return "break"
        self.cod_var.set(art["codigo"])
        self.desc_var.set(art["descripcion"])
        self.um_var.set(art["unidad_medida"])
        self.qty_entry.focus_set()
        return "break"

    def add_item(self, event=None):
        if not self.desc_var.get():
            self.load_code()
            if not self.desc_var.get():
                return "break"
        try:
            qty = float(self.qty_var.get().replace(",", "."))
            if qty <= 0:
                raise ValueError
        except Exception:
            messagebox.showwarning("Cantidad inválida", "La cantidad debe ser mayor a cero.", parent=self)
            self.qty_entry.focus_set()
            return "break"
        codigo = normalize_text(self.cod_var.get())
        for d in self.detalles:
            if d["codigo"] == codigo:
                d["cantidad"] += qty
                self.refresh_items()
                self.cod_var.set(""); self.desc_var.set(""); self.um_var.set(""); self.qty_var.set("")
                self.cod_entry.focus_set()
                return "break"
        self.detalles.append({"item": len(self.detalles)+1, "codigo": codigo, "descripcion": self.desc_var.get(), "um": self.um_var.get(), "cantidad": qty})
        self.refresh_items()
        self.cod_var.set(""); self.desc_var.set(""); self.um_var.set(""); self.qty_var.set("")
        self.cod_entry.focus_set()
        return "break"

    def refresh_items(self):
        for x in self.tree.get_children():
            self.tree.delete(x)
        for i, d in enumerate(self.detalles, start=1):
            d["item"] = i
            stock = self.db.get_stock(d["codigo"], self.origen_var.get())
            self.tree.insert("", "end", values=(i, d["codigo"], d["descripcion"], d["um"], f"{d['cantidad']:g}", f"{stock:g}"))

    def select_row(self, event=None):
        sel = self.tree.selection()
        self.selected_index = self.tree.index(sel[0]) if sel else None

    def delete_item(self):
        if self.selected_index is None:
            messagebox.showinfo("Seleccione", "Seleccione un item para eliminar.", parent=self); return
        self.detalles.pop(self.selected_index); self.selected_index = None; self.refresh_items()

    def modify_item(self):
        if self.selected_index is None:
            messagebox.showinfo("Seleccione", "Seleccione un item para modificar.", parent=self); return
        d = self.detalles[self.selected_index]
        new_qty = simpledialog.askstring("Modificar cantidad", f"Código {d['codigo']} - Nueva cantidad:", initialvalue=str(d["cantidad"]), parent=self)
        if new_qty is None:
            return
        try:
            qty, _art = self.db.validar_cantidad_articulo(d["codigo"], new_qty, "cantidad requerida")
            d["cantidad"] = qty; self.refresh_items()
        except Exception as e:
            messagebox.showwarning("Inválido", str(e), parent=self)

    def save(self):
        try:
            if not self.app.require_almacen(self, self.origen_var.get()):
                return
            self.fill_cc_from_typed()
            transferencia, vale_salida, vale_ingreso = self.db.transferir_stock(
                self.fecha_var.get(), self.origen_var.get(), self.destino_var.get(), self.cc_code_var.get(),
                self.solicitante_var.get(), self.obs_var.get(), self.detalles, self.app.usuario_actual, self.guia_var.get()
            )
            messagebox.showinfo("Transferencia registrada", f"Transferencia: {transferencia}\nSalida origen: {vale_salida}\nIngreso destino: {vale_ingreso}", parent=self)
            self.detalles = []; self.refresh_items(); self.guia_var.set(""); self.cod_entry.focus_set()
        except Exception as e:
            messagebox.showerror("No se pudo transferir", str(e), parent=self)
            self.focus_force()


class PrintPreviewDialog(tk.Toplevel):
    def __init__(self, parent, movimiento, detalles):
        super().__init__(parent)
        self.movimiento = movimiento
        self.detalles = detalles
        self.generated_path = None
        self.title("Previsualización de documento A5 horizontal")
        self.geometry("620x760+180+35")
        self.minsize(540, 650)
        bind_as_child_window(self, parent)
        self.build()

    def document_title(self):
        tipo = self.movimiento["tipo_codigo"]
        if tipo == "CO":
            return "Vale de Salida de Almacen"
        if tipo == "CN":
            return "Ingreso de Almacen"
        return "Ajuste de Inventario"

    def display_qty(self, row):
        qty = float(row["cantidad"])
        if self.movimiento["tipo_codigo"] == "CO":
            return abs(qty)
        return qty

    def build(self):
        main = ttk.Frame(self, padding=18)
        main.pack(fill="both", expand=True)
        ttk.Label(main, text=self.document_title(), font=("Arial", 15, "bold")).pack(pady=(0, 10))
        info = ttk.Frame(main)
        info.pack(fill="x", pady=6)
        data = [
            ("Vale", self.movimiento["vale"]),
            ("Fecha", fecha_text_from_iso(self.movimiento["fecha_operacion"])),
            ("Documento", self.movimiento["documento"]),
            ("OT", self.movimiento["ot"] or ""),
            ("Centro de costo", f"{self.movimiento['centro_costo_codigo']} - {self.movimiento['centro_costo_nombre']}"),
            ("Solicitante", self.movimiento["solicitante"]),
            ("Usuario", self.movimiento["usuario"]),
        ]
        for i, (label, value) in enumerate(data):
            r, c = divmod(i, 2)
            ttk.Label(info, text=f"{label}:", font=("Arial", 10, "bold")).grid(row=r, column=c*2, sticky="w", padx=(0, 6), pady=3)
            ttk.Label(info, text=str(value)).grid(row=r, column=c*2+1, sticky="w", padx=(0, 30), pady=3)
        ttk.Label(main, text=f"Observación: {self.movimiento['observacion']}", wraplength=540).pack(anchor="w", pady=(6, 6))
        cols = ("item", "codigo", "descripcion", "um", "cantidad")
        tree = ttk.Treeview(main, columns=cols, show="headings", height=12)
        specs = [("item", "Item", 45), ("codigo", "Código", 85), ("descripcion", "Descripción", 280), ("um", "UM", 55), ("cantidad", "Cant.", 70)]
        for c, t, w in specs:
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="center" if c != "descripcion" else "w")
        tree_box = ttk.Frame(main)
        tree_box.pack(fill="both", expand=True, pady=10)
        tree_box.rowconfigure(0, weight=1); tree_box.columnconfigure(0, weight=1)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(in_=tree_box, row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        for r in self.detalles:
            tree.insert("", "end", values=(r["item"], r["codigo"], r["descripcion"], r["unidad_medida"], f"{self.display_qty(r):g}"))
        firma = ttk.Frame(main)
        firma.pack(fill="x", pady=(20, 8))
        ttk.Label(firma, text="______________________________\nFirma del Solicitante", justify="center").pack(side="left", expand=True)
        ttk.Label(firma, text="______________________________\nVB° Responsable", justify="center").pack(side="right", expand=True)
        buttons = ttk.Frame(main)
        buttons.pack(fill="x", pady=8)
        ttk.Button(buttons, text="Guardar Word", command=self.save_word_as).pack(side="right", padx=6)
        ttk.Button(buttons, text="Imprimir Word", command=self.print_document).pack(side="right", padx=6)
        ttk.Button(buttons, text="Cerrar", command=self.destroy).pack(side="right", padx=6)

    def default_path(self):
        os.makedirs(VALE_DIR, exist_ok=True)
        return os.path.join(VALE_DIR, f"vale_{self.movimiento['vale']}.xlsx")

    def default_word_path(self):
        os.makedirs(VALE_DIR, exist_ok=True)
        return os.path.join(VALE_DIR, f"vale_{self.movimiento['vale']}.docx")

    def build_word_document(self):
        Doc = require_docx_document(parent=self)
        if not os.path.exists(VALE_WORD_TEMPLATE):
            # Respaldo simple si la plantilla no está disponible.
            doc = Doc()
            section = doc.sections[0]
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = 457200
            title = doc.add_paragraph()
            title.alignment = 1
            run = title.add_run(self.document_title().upper())
            run.bold = True
            table = doc.add_table(rows=1, cols=5)
            table.style = 'Table Grid'
            for i, h in enumerate(["ITEM", "CÓDIGO", "DESCRIPCIÓN", "UM", "CANTIDAD"]):
                table.cell(0, i).text = h
            for d in self.detalles:
                cells = table.add_row().cells
                vals = [d["item"], d["codigo"], d["descripcion"], d["unidad_medida"], f"{self.display_qty(d):g}"]
                for i, value in enumerate(vals):
                    cells[i].text = str(value)
            return doc

        # La plantilla oficial dispone de 12 líneas de detalle por página.
        # Cada bloque adicional se convierte en una nueva hoja completa del mismo vale,
        # repitiendo encabezado, datos generales, cabecera de detalle y firmas.
        max_items_per_page = VALE_ITEMS_PER_PAGE
        detalles = list(self.detalles or [])
        chunks = [detalles[i:i + max_items_per_page] for i in range(0, len(detalles), max_items_per_page)] or [[]]
        total_pages = len(chunks)
        tipo_mov = {"CO": "SALIDA", "CN": "INGRESO", "AJ": "AJUSTE"}.get(self.movimiento["tipo_codigo"], "MOVIMIENTO")

        def build_page(page_items, page_number):
            page_doc = Doc(VALE_WORD_TEMPLATE)
            mapping = {
                "{{TIPO_MOV}}": tipo_mov,
                "{{VALE}}": self.movimiento["vale"],
                "{{FECHA}}": fecha_text_from_iso(self.movimiento["fecha_operacion"]),
                "{{DOCUMENTO}}": self.movimiento["documento"],
                "{{OT}}": self.movimiento["ot"] or "",
                "{{CENTRO_COSTO}}": f"{self.movimiento['centro_costo_codigo']} - {self.movimiento['centro_costo_nombre']}",
                "{{ALMACEN}}": self.movimiento["almacen_codigo"] if "almacen_codigo" in self.movimiento.keys() else "",
                "{{SOLICITANTE}}": self.movimiento["solicitante"],
                "{{USUARIO}}": self.movimiento["usuario"],
                "{{OBSERVACION}}": self.movimiento["observacion"],
            }
            for i in range(1, max_items_per_page + 1):
                if i <= len(page_items):
                    d = page_items[i - 1]
                    mapping[f"{{{{I{i:02d}}}}}"] = str(d["item"])
                    mapping[f"{{{{COD{i:02d}}}}}"] = str(d["codigo"])
                    mapping[f"{{{{DESC{i:02d}}}}}"] = str(d["descripcion"])
                    mapping[f"{{{{UM{i:02d}}}}}"] = str(d["unidad_medida"])
                    mapping[f"{{{{CANT{i:02d}}}}}"] = f"{self.display_qty(d):g}"
                else:
                    for pref in ("I", "COD", "DESC", "UM", "CANT"):
                        mapping[f"{{{{{pref}{i:02d}}}}}"] = ""
            replace_docx_placeholders(page_doc, mapping)

            # Numeración real de páginas: 01 de 02, 02 de 02, etc.
            for table in page_doc.tables:
                for row in table.rows:
                    row_text = [normalize_text(cell.text).upper() for cell in row.cells]
                    if any(text.startswith("PÁGINA:") or text.startswith("PAGINA:") for text in row_text):
                        _docx_set_cell_text(row.cells[-1], f"{page_number:02d} de {total_pages:02d}")
                        break

            # Quitar líneas vacías de ítems. Esta corrección evita que un vale de 1, 2, etc.
            # productos mande la tabla de firmas a una segunda hoja.
            items_table = _find_docx_table(page_doc, ["ITEM", "CÓDIGO", "DESCRIPCIÓN", "UM", "CANTIDAD"])
            if items_table is not None:
                desired_rows = 1 + len(page_items)  # cabecera + detalle real
                while len(items_table.rows) > desired_rows:
                    tr = items_table.rows[-1]._tr
                    items_table._tbl.remove(tr)

            # La plantilla trae un párrafo vacío después de las firmas. Al concatenar páginas
            # ese párrafo podía crear una hoja totalmente en blanco.
            _docx_remove_trailing_empty_body_paragraph(page_doc)
            return page_doc

        doc = build_page(chunks[0], 1)
        for page_number, page_items in enumerate(chunks[1:], start=2):
            _docx_append_body(doc, build_page(page_items, page_number))
        return doc

    def save_word_file(self, path):
        doc = self.build_word_document()
        doc.save(path)
        self.generated_path = path
        return path

    def save_word_as(self):
        path = filedialog.asksaveasfilename(parent=self, title="Guardar vale en Word", defaultextension=".docx", initialfile=f"vale_{self.movimiento['vale']}.docx", filetypes=[("Word", "*.docx")])
        if not path:
            return
        try:
            self.save_word_file(path)
            messagebox.showinfo("Documento guardado", f"Documento Word guardado:\n{path}", parent=self)
        except Exception as e:
            messagebox.showerror("No se pudo guardar", str(e), parent=self)

    def build_workbook(self):
        if Workbook is None:
            raise ValueError("Instale openpyxl: pip install openpyxl")
        wb = Workbook()
        ws = wb.active
        ws.title = "Vale A5"

        # Formato de impresión: A5 VERTICAL
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = 11  # A5
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.30
        ws.page_margins.bottom = 0.30
        ws.page_margins.header = 0.10
        ws.page_margins.footer = 0.10

        thin = Side(style="thin", color="999999")
        border = Border(top=thin, bottom=thin, left=thin, right=thin)
        fill = PatternFill("solid", fgColor="D9EAF7")

        ws.merge_cells("A1:E1")
        ws["A1"] = self.document_title()
        ws["A1"].font = Font(size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24

        # Cabecera compacta para hoja A5 vertical
        meta = [
            ("Vale", self.movimiento["vale"], "Fecha", fecha_text_from_iso(self.movimiento["fecha_operacion"])),
            ("Doc.", self.movimiento["documento"], "OT", self.movimiento["ot"] or ""),
            ("C. Costo", f"{self.movimiento['centro_costo_codigo']} - {self.movimiento['centro_costo_nombre']}", "", ""),
            ("Solicitante", self.movimiento["solicitante"], "", ""),
            ("Usuario", self.movimiento["usuario"], "", ""),
        ]
        row = 3
        for a, b, c, d in meta:
            ws.cell(row=row, column=1, value=a).font = Font(bold=True, size=9)
            ws.cell(row=row, column=2, value=b).font = Font(size=9)
            if c:
                ws.cell(row=row, column=4, value=c).font = Font(bold=True, size=9)
                ws.cell(row=row, column=5, value=d).font = Font(size=9)
            row += 1

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        obs_cell = ws.cell(row=row, column=1, value=f"Observación: {self.movimiento['observacion']}")
        obs_cell.font = Font(size=9)
        obs_cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 2

        headers = ["Item", "Código", "Descripción", "UM", "Cant."]
        for col, h in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=h)
            cell.font = Font(bold=True, size=9)
            cell.fill = fill
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
        row += 1

        for d in self.detalles:
            values = [d["item"], d["codigo"], d["descripcion"], d["unidad_medida"], self.display_qty(d)]
            for col, value in enumerate(values, start=1):
                cell = ws.cell(row=row, column=col, value=value)
                cell.font = Font(size=8)
                cell.border = border
                cell.alignment = Alignment(horizontal="center" if col != 3 else "left", vertical="center", wrap_text=(col == 3))
            ws.row_dimensions[row].height = 24
            row += 1

        row += 2
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
        ws.cell(row=row, column=1, value="________________________")
        ws.cell(row=row, column=4, value="________________________")
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=4).alignment = Alignment(horizontal="center")
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
        ws.cell(row=row, column=1, value="Firma del Solicitante")
        ws.cell(row=row, column=4, value="VB° Responsable")
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=4).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=1).font = Font(size=9)
        ws.cell(row=row, column=4).font = Font(size=9)

        widths = {"A": 6, "B": 13, "C": 33, "D": 8, "E": 10}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        ws.print_area = f"A1:E{row}"
        return wb

    def save_file(self, path):
        wb = self.build_workbook()
        wb.save(path)
        self.generated_path = path
        return path

    def save_as(self):
        # Alias de seguridad: los vales vigentes se generan en Word.
        return self.save_word_as()

    def print_document(self):
        try:
            path = self.generated_path or self.save_word_file(self.default_word_path())
            if os.name == "nt":
                os.startfile(path, "print")
                messagebox.showinfo("Impresión", "El formato Word fue enviado a la impresora predeterminada.", parent=self)
            else:
                messagebox.showinfo("Impresión", f"Documento Word generado. Ábralo para imprimir:\n{path}", parent=self)
        except Exception as e:
            messagebox.showerror("No se pudo imprimir", str(e), parent=self)



class RegisterArticleDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.title("Registrar Artículo")
        self.geometry("680x470+180+100")
        bind_as_child_window(self, self.master if hasattr(self, "master") else None)
        self.configure(bg="#f4f6f8")
        self.build()

    def build(self):
        frm = ttk.Frame(self, padding=20)
        frm.pack(fill="both", expand=True)
        self.fecha = tk.StringVar(value=now_text())
        self.codigo = tk.StringVar(value=self.suggest_code())
        self.desc = tk.StringVar()
        self.um = tk.StringVar()
        self.tipo = tk.StringVar(value="CONSUMIBLE")
        self.stock_minimo = tk.StringVar(value="0")
        self.stock_minimo_aqp = tk.StringVar(value="0")
        self.stock_minimo_min = tk.StringVar(value="0")
        self.valoriza = tk.BooleanVar(value=True)
        self.permite_decimales = tk.BooleanVar(value=True)
        fields = [("Fecha de registro:", self.fecha, "readonly"), ("Código:", self.codigo, "normal"), ("Descripción:", self.desc, "normal"), ("Stock mínimo general:", self.stock_minimo, "normal"), ("Stock mínimo AQP:", self.stock_minimo_aqp, "normal"), ("Stock mínimo MIN:", self.stock_minimo_min, "normal")]
        for r, (lab, var, state) in enumerate(fields):
            ttk.Label(frm, text=lab).grid(row=r, column=0, sticky="w", pady=8)
            ttk.Entry(frm, textvariable=var, width=34, state=state).grid(row=r, column=1, sticky="w", pady=8)
        ttk.Label(frm, text="Unidad de medida:").grid(row=6, column=0, sticky="w", pady=8)
        ttk.Combobox(frm, textvariable=self.um, values=UNIDADES_PERMITIDAS, state="readonly", width=31).grid(row=6, column=1, sticky="w", pady=8)
        ttk.Label(frm, text="Tipo de artículo:").grid(row=7, column=0, sticky="w", pady=8)
        ttk.Combobox(frm, textvariable=self.tipo, values=TIPOS_ARTICULO, state="readonly", width=31).grid(row=7, column=1, sticky="w", pady=8)
        ttk.Checkbutton(frm, text="Artículo valorizable", variable=self.valoriza).grid(row=8, column=1, sticky="w", pady=4)
        ttk.Checkbutton(frm, text="Permite decimales", variable=self.permite_decimales).grid(row=9, column=1, sticky="w", pady=4)
        for child in frm.winfo_children():
            if isinstance(child, (ttk.Entry, ttk.Combobox)):
                child.bind("<Return>", self.focus_next)
        ttk.Label(frm, text="Si AQP/MIN quedan en 0, se usará el mínimo general. Servicios pueden marcarse como no valorizables.", foreground="#666666").grid(row=10, column=1, sticky="w", pady=4)
        ttk.Button(frm, text="💾 Guardar", command=self.save).grid(row=11, column=1, sticky="e", pady=20)

    def suggest_code(self):
        nums = []
        for r in self.db.all_articulos():
            try:
                nums.append(int(only_digits(r["codigo"])))
            except Exception:
                pass
        return str(max(nums) + 1 if nums else 1)

    def focus_next(self, event):
        event.widget.tk_focusNext().focus()
        return "break"

    def save(self):
        try:
            if not self.app.require_write(self):
                return
            self.db.add_articulo(self.codigo.get(), self.desc.get(), self.um.get(), self.tipo.get(), self.stock_minimo.get(), self.stock_minimo_aqp.get(), self.stock_minimo_min.get(), self.valoriza.get(), self.permite_decimales.get())
            self.db.audit("articulos", self.codigo.get(), "CREACION", self.app.usuario_actual, "Registro de artículo")
            messagebox.showinfo("Registro completo", "El registro se completó correctamente.", parent=self)
            self.fecha.set(now_text())
            self.codigo.set(self.suggest_code())
            self.desc.set("")
            self.um.set("")
            self.tipo.set("CONSUMIBLE")
            self.stock_minimo.set("0")
            self.stock_minimo_aqp.set("0")
            self.stock_minimo_min.set("0")
            self.valoriza.set(True)
            self.permite_decimales.set(True)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)
            self.after(100, self.focus_force)


class BulkArticleDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.title("Registro Artículo Masivo")
        self.geometry("720x330+180+120")
        bind_as_child_window(self, self.master if hasattr(self, "master") else None)
        self.file_path = tk.StringVar()
        self.build()

    def build(self):
        frm = ttk.Frame(self, padding=20)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Fecha de registro:").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Entry(frm, width=32, state="readonly", textvariable=tk.StringVar(value=now_text())).grid(row=0, column=1, sticky="w")
        ttk.Label(frm, text="Cargar documento:").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(frm, textvariable=self.file_path, width=56, state="readonly").grid(row=1, column=1, sticky="w")
        ttk.Button(frm, text="Buscar", command=self.pick_file).grid(row=1, column=2, padx=6)
        ttk.Label(frm, text="Formato obligatorio: CODIGO | DESCRIPCION | UNIDAD DE MEDIDA | TIPO DE ARTICULO\nOpcional: STOCK MINIMO o STOCK MINIMO AQP | STOCK MINIMO MIN. Máximo 15 artículos.\nCarga atómica: si una fila falla, no se guarda ninguna.", foreground="#555555").grid(row=2, column=1, sticky="w", pady=12)
        ttk.Button(frm, text="💾 Guardar carga", command=self.save).grid(row=3, column=1, sticky="e", pady=20)

    def pick_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if p:
            self.file_path.set(p)

    def save(self):
        if not self.app.require_write(self):
            return
        if load_workbook is None:
            messagebox.showerror("Falta librería", "Instale openpyxl: pip install openpyxl", parent=self)
            return
        try:
            if not self.file_path.get():
                raise ValueError("Seleccione un archivo.")
            wb = load_workbook(self.file_path.get(), data_only=True)
            ws = wb.active
            headers = [normalize_text(c.value) for c in ws[1][:6]]
            expected = ["CODIGO", "DESCRIPCION", "UNIDAD DE MEDIDA", "TIPO DE ARTICULO"]
            if headers[:4] != expected:
                raise ValueError("Documento inválido. Las columnas deben ser: CODIGO, DESCRIPCION, UNIDAD DE MEDIDA, TIPO DE ARTICULO")
            rows = []
            for row in ws.iter_rows(min_row=2, max_col=6, values_only=True):
                if not any(row):
                    continue
                rows.append(row)
            if len(rows) > 15:
                raise ValueError("El archivo supera el máximo de 15 artículos.")
            clean_rows = []
            for row in rows:
                codigo, desc, um, tipo = row[:4]
                minimo_a = row[4] if len(row) > 4 else 0
                minimo_b = row[5] if len(row) > 5 else None
                if minimo_b is None or minimo_b == "":
                    clean_rows.append((codigo, desc, um, tipo, minimo_a or 0, None, None))
                else:
                    clean_rows.append((codigo, desc, um, tipo, 0, minimo_a or 0, minimo_b or 0))
            total = self.db.add_articulos_bulk(clean_rows, self.app.usuario_actual)
            messagebox.showinfo("Ingreso Satisfactorio", f"Ingreso Satisfactorio. Artículos cargados: {total}", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Documento inválido", str(e), parent=self)


class ArticleQueryDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.title("Consulta / Modificación de Artículos")
        self.geometry("980x560+140+80")
        bind_as_child_window(self, self.master if hasattr(self, "master") else None)
        self.build()
        self.search()

    def build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        self.term = tk.StringVar()
        ttk.Label(top, text="Buscar:").pack(side="left")
        e = ttk.Entry(top, textvariable=self.term, width=45)
        e.pack(side="left", padx=8)
        e.bind("<KeyRelease>", lambda e: self.search())
        cols = ("codigo", "descripcion", "um", "tipo", "valoriza", "decimales", "minimo", "min_aqp", "min_min", "fecha")
        tree_box = ttk.Frame(self)
        tree_box.pack(fill="both", expand=True, padx=10, pady=10)
        tree_box.rowconfigure(0, weight=1); tree_box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_box, columns=cols, show="headings", height=13)
        specs = [("codigo", "Código", 90), ("descripcion", "Descripción", 300), ("um", "UM", 70), ("tipo", "Tipo", 120), ("valoriza", "Valoriza", 75), ("decimales", "Decimales", 80), ("minimo", "Mín. Gral", 90), ("min_aqp", "Mín. AQP", 90), ("min_min", "Mín. MIN", 90), ("fecha", "Fecha Registro", 150)]
        for c, t, w in specs:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c != "descripcion" else "w", stretch=False)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.load_selected)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="x")
        self.codigo = tk.StringVar()
        self.desc = tk.StringVar()
        self.um = tk.StringVar()
        self.tipo = tk.StringVar()
        self.minimo = tk.StringVar()
        self.minimo_aqp = tk.StringVar()
        self.minimo_min = tk.StringVar()
        self.valoriza = tk.BooleanVar(value=True)
        self.permite_decimales = tk.BooleanVar(value=True)
        ttk.Label(frm, text="Código:").grid(row=0, column=0)
        ttk.Entry(frm, textvariable=self.codigo, width=12, state="readonly").grid(row=0, column=1, padx=5)
        ttk.Label(frm, text="Descripción:").grid(row=0, column=2)
        ttk.Entry(frm, textvariable=self.desc, width=30).grid(row=0, column=3, padx=5)
        ttk.Label(frm, text="UM:").grid(row=0, column=4)
        ttk.Combobox(frm, textvariable=self.um, values=UNIDADES_PERMITIDAS, state="readonly", width=10).grid(row=0, column=5, padx=5)
        ttk.Label(frm, text="Tipo:").grid(row=0, column=6)
        ttk.Combobox(frm, textvariable=self.tipo, values=TIPOS_ARTICULO, state="readonly", width=15).grid(row=0, column=7, padx=5)
        ttk.Label(frm, text="Mín. general:").grid(row=1, column=0, pady=8)
        ttk.Entry(frm, textvariable=self.minimo, width=12).grid(row=1, column=1, padx=5, pady=8)
        ttk.Label(frm, text="Mín. AQP:").grid(row=1, column=2, pady=8)
        ttk.Entry(frm, textvariable=self.minimo_aqp, width=12).grid(row=1, column=3, padx=5, pady=8)
        ttk.Label(frm, text="Mín. MIN:").grid(row=1, column=4, pady=8)
        ttk.Entry(frm, textvariable=self.minimo_min, width=12).grid(row=1, column=5, padx=5, pady=8)
        ttk.Checkbutton(frm, text="Valoriza", variable=self.valoriza).grid(row=2, column=1, pady=6, sticky="w")
        ttk.Checkbutton(frm, text="Permite decimales", variable=self.permite_decimales).grid(row=2, column=3, pady=6, sticky="w")
        ttk.Button(frm, text="💾 Guardar", command=self.save).grid(row=2, column=7, padx=8, sticky="e")

    def search(self):
        for x in self.tree.get_children():
            self.tree.delete(x)
        term = normalize_text(self.term.get())
        for r in self.db.search_articulos(term):
            self.tree.insert("", "end", values=(r["codigo"], r["descripcion"], r["unidad_medida"], r["tipo_articulo"], "SI" if r["valoriza"] else "NO", "SI" if r["permite_decimales"] else "NO", f"{r['stock_minimo']:g}", f"{r['stock_minimo_aqp']:g}", f"{r['stock_minimo_min']:g}", r["fecha_registro"]))

    def load_selected(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel[0], "values")
        self.codigo.set(values[0])
        self.desc.set(values[1])
        self.um.set(values[2])
        self.tipo.set(values[3])
        self.valoriza.set(values[4] == "SI")
        self.permite_decimales.set(values[5] == "SI")
        self.minimo.set(values[6])
        self.minimo_aqp.set(values[7])
        self.minimo_min.set(values[8])

    def save(self):
        try:
            if not self.app.require_write(self):
                return
            before = self.db.get_articulo(self.codigo.get())
            self.db.update_articulo(self.codigo.get(), self.desc.get(), self.um.get(), self.tipo.get(), self.minimo.get(), self.minimo_aqp.get(), self.minimo_min.get(), self.valoriza.get(), self.permite_decimales.get())
            self.db.audit("articulos", self.codigo.get(), "MODIFICACION", self.app.usuario_actual, f"Antes: {dict(before) if before else ''}")
            messagebox.showinfo("Completo", "Artículo modificado correctamente.", parent=self)
            self.search()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)
            self.after(100, self.focus_force)




class RequerimientoListDialog(tk.Toplevel):
    def __init__(self, parent, db, on_select, estados=("PENDIENTE", "DESPACHO PARCIAL")):
        super().__init__(parent)
        self.db = db
        self.on_select = on_select
        self.estados = estados
        self.term = tk.StringVar()
        self.title("Consulta de Requerimientos")
        self.geometry("950x470+150+120")
        bind_as_child_window(self, parent)
        self.build()
        self.search()

    def build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Buscar requerimiento / solicitante / OT:").pack(side="left")
        e = ttk.Entry(top, textvariable=self.term, width=42)
        e.pack(side="left", padx=8)
        e.bind("<KeyRelease>", lambda e: self.search())
        e.bind("<Return>", lambda e: self.choose())
        e.focus_set()
        cols = ("req", "fecha", "almacen", "estado", "sol", "cc", "ot", "pend")
        tree_box = ttk.Frame(self)
        tree_box.pack(fill="both", expand=True, padx=10, pady=10)
        tree_box.rowconfigure(0, weight=1); tree_box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_box, columns=cols, show="headings", height=15)
        specs = [("req", "Requerimiento", 130), ("fecha", "Fecha", 95), ("almacen", "Almacén", 80), ("estado", "Estado", 130), ("sol", "Solicitante", 210), ("cc", "CC", 130), ("ot", "OT", 110), ("pend", "Pendiente", 90)]
        for c, t, w in specs:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, minwidth=w, anchor="center" if c not in ["sol", "cc"] else "w", stretch=False)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self.choose)
        self.tree.bind("<Return>", self.choose)
        ttk.Button(self, text="Cerrar", command=self.destroy).pack(pady=6)

    def search(self):
        for x in self.tree.get_children():
            self.tree.delete(x)
        first = None
        for r in self.db.list_requerimientos(self.estados, self.term.get()):
            iid = self.tree.insert("", "end", values=(r["requerimiento"], fecha_text_from_iso(r["fecha_requerimiento"]), r["almacen_codigo"], r["estado"], r["solicitante"], f"{r['centro_costo_codigo']} - {r['centro_costo_nombre']}", r["ot"] or "", f"{r['total_pendiente']:g}"))
            if first is None:
                first = iid
        if first:
            self.tree.selection_set(first)
            self.tree.focus(first)

    def choose(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return "break"
        req = self.tree.item(sel[0], "values")[0]
        self.on_select(req)
        self.destroy()
        return "break"


class RequerimientoMultiSelectDialog(tk.Toplevel):
    """Selecciona uno o varios requerimientos pendientes de la misma OT para generar OC.
    UX v36: usa columna de check. Al marcar un requerimiento, solo deja marcar otros de la misma OT.
    """
    def __init__(self, parent, db, on_select):
        super().__init__(parent)
        self.db = db
        self.on_select = on_select
        self.term = tk.StringVar()
        today = date.today()
        self.mes = tk.StringVar(value=f"{today.month:02d}")
        self.anio = tk.StringVar(value=str(today.year))
        self.checked = set()
        self.locked_ot = None
        self.row_ot = {}
        self.title("Seleccionar requerimientos para OC")
        self.geometry("1120x590+100+70")
        self.minsize(1040, 520)
        bind_as_child_window(self, parent)
        self.build(); self.search()

    def build(self):
        top = ttk.Frame(self, padding=10); top.pack(fill="x")
        ttk.Label(top, text="Mes:").pack(side="left")
        ttk.Combobox(top, textvariable=self.mes, values=[f"{i:02d}" for i in range(1,13)], state="readonly", width=5).pack(side="left", padx=4)
        ttk.Label(top, text="Año:").pack(side="left")
        ttk.Entry(top, textvariable=self.anio, width=7).pack(side="left", padx=4)
        ttk.Label(top, text="Buscar:").pack(side="left", padx=(12,2))
        e = ttk.Entry(top, textvariable=self.term, width=38); e.pack(side="left", padx=4)
        e.bind("<KeyRelease>", lambda e: self.search())
        e.bind("<Return>", lambda e: self.toggle_current())
        ttk.Button(top, text="Filtrar", command=self.search).pack(side="left", padx=4)
        ttk.Button(top, text="Limpiar selección", command=self.clear_checks).pack(side="left", padx=4)
        self.msg = tk.StringVar(value="Marque un requerimiento. Luego solo podrá marcar requerimientos de la misma OT.")
        ttk.Label(top, textvariable=self.msg, foreground="#555").pack(side="left", padx=10)
        cols=("sel","req","fecha","entrega","estado","sol","cc","ot","pend")
        box=ttk.Frame(self); box.pack(fill="both", expand=True, padx=10, pady=8); box.rowconfigure(0, weight=1); box.columnconfigure(0, weight=1)
        self.tree=ttk.Treeview(box, columns=cols, show="headings", height=17, selectmode="browse")
        specs=[("sel","✓",45),("req","Requerimiento",135),("fecha","Fecha",90),("entrega","Entrega solicitada",125),("estado","Estado",130),("sol","Solicitante",210),("cc","CC",160),("ot","OT",120),("pend","Pendiente",90)]
        for c,t,w in specs:
            self.tree.heading(c,text=t); self.tree.column(c,width=w,minwidth=w,anchor="center" if c not in ("sol","cc") else "w",stretch=False)
        self.tree.tag_configure("locked", foreground="#9a9a9a")
        self.tree.tag_configure("checked", background="#dff3df")
        vsb=ttk.Scrollbar(box,orient="vertical",command=self.tree.yview); hsb=ttk.Scrollbar(box,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        self.tree.bind("<Double-1>", self.toggle_current)
        self.tree.bind("<space>", self.toggle_current)
        self.tree.bind("<Return>", self.choose)
        self.tree.bind("<Up>", self._arrow)
        self.tree.bind("<Down>", self._arrow)
        btn=ttk.Frame(self,padding=10); btn.pack(fill="x")
        ttk.Label(btn, text="Doble clic o Espacio: marcar/desmarcar. Enter: continuar.", foreground="#666").pack(side="left", padx=4)
        ttk.Button(btn,text="Seleccionar marcados",command=self.choose).pack(side="right",padx=4)
        ttk.Button(btn,text="Cerrar",command=self.destroy).pack(side="right",padx=4)

    def _arrow(self, event=None):
        # Deja que Treeview se mueva y luego asegura que la fila enfocada quede seleccionada.
        self.after(30, lambda: self.tree.selection_set(self.tree.focus()) if self.tree.focus() else None)
        return None

    def clear_checks(self):
        self.checked.clear(); self.locked_ot = None; self.refresh_marks()

    def search(self):
        self.checked.clear(); self.locked_ot = None; self.row_ot = {}
        for x in self.tree.get_children(): self.tree.delete(x)
        try:
            desde = f"01/{self.mes.get()}/{self.anio.get()}"
            f = parse_fecha(desde)
            hasta_date = date(f.year, 12, 31) if f.month == 12 else date(f.year, f.month + 1, 1) - timedelta(days=1)
            hasta = hasta_date.strftime(DATE_FMT)
        except Exception:
            desde, hasta = "", ""
        rows = self.db.list_requerimientos_filtered(desde, hasta, estado="TODOS", term=self.term.get())
        first=None
        for r in rows:
            if r["estado"] in ("ATENDIDO", "ANULADO") or float(r["total_pendiente"] or 0) <= 0:
                continue
            # Si el requerimiento ya generó una OC activa, deja de aparecer como pendiente de compra.
            if self.db.requerimiento_tiene_oc_activa(r["requerimiento"]):
                continue
            vals=("☐", r["requerimiento"], fecha_text_from_iso(r["fecha_requerimiento"]), fecha_text_from_iso(r["fecha_entrega_solicitada"]) if "fecha_entrega_solicitada" in r.keys() and r["fecha_entrega_solicitada"] else "", r["estado"], r["solicitante"], f"{r['centro_costo_codigo']} - {r['centro_costo_nombre']}", r["ot"] or "", f"{float(r['total_pendiente'] or 0):g}")
            iid=self.tree.insert("","end",values=vals)
            self.row_ot[iid] = r["ot"] or ""
            if first is None: first=iid
        if first:
            self.tree.selection_set(first); self.tree.focus(first)
        self.refresh_marks()

    def refresh_marks(self):
        for iid in self.tree.get_children():
            vals=list(self.tree.item(iid,"values"))
            vals[0] = "☑" if iid in self.checked else "☐"
            ot = self.row_ot.get(iid, "")
            tags=[]
            if iid in self.checked:
                tags.append("checked")
            elif self.locked_ot is not None and ot != self.locked_ot:
                tags.append("locked")
            self.tree.item(iid, values=vals, tags=tuple(tags))
        if self.locked_ot:
            self.msg.set(f"Selección bloqueada a OT: {self.locked_ot}. Solo puede marcar requerimientos con esa misma OT.")
        else:
            self.msg.set("Marque un requerimiento. Luego solo podrá marcar requerimientos de la misma OT.")

    def toggle_current(self, event=None):
        iid = self.tree.focus() or (self.tree.selection()[0] if self.tree.selection() else None)
        if not iid:
            return "break"
        ot = self.row_ot.get(iid, "")
        if iid in self.checked:
            self.checked.remove(iid)
            if not self.checked:
                self.locked_ot = None
        else:
            if self.locked_ot is not None and ot != self.locked_ot:
                messagebox.showwarning("OT diferente", "Solo se pueden unir requerimientos que pertenecen a la misma OT.", parent=self)
                return "break"
            self.locked_ot = ot
            self.checked.add(iid)
        self.refresh_marks()
        return "break"

    def choose(self, event=None):
        if not self.checked:
            self.toggle_current()
        if not self.checked:
            return "break"
        vals = [self.tree.item(i,"values") for i in self.checked]
        ots = {v[7] for v in vals}
        if len(ots) > 1:
            messagebox.showwarning("OT diferente", "Solo se pueden unir requerimientos de la misma OT.", parent=self)
            return "break"
        reqs = [v[1] for v in vals]
        self.on_select(reqs)
        self.destroy()
        return "break"


class RequerimientoDetailDialog(tk.Toplevel):
    def __init__(self, parent, app, requerimiento):
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self.req, self.det = self.db.find_requerimiento(requerimiento)
        self.title(f"Detalle de Requerimiento - {requerimiento}")
        self.geometry("1050x620+120+70")
        self.minsize(980, 560)
        bind_as_child_window(self, parent)
        self.build()

    def build(self):
        if not self.req:
            ttk.Label(self, text="No se encontró el requerimiento.").pack(padx=20, pady=20)
            return
        top = ttk.LabelFrame(self, text="Cabecera", padding=10)
        top.pack(fill="x", padx=10, pady=8)
        data = [
            ("Requerimiento", self.req["requerimiento"]),
            ("Fecha solicitud", fecha_text_from_iso(self.req["fecha_requerimiento"])),
            ("Fecha entrega solicitada", fecha_text_from_iso(self.req["fecha_entrega_solicitada"]) if "fecha_entrega_solicitada" in self.req.keys() else ""),
            ("Almacén", f"{self.req['almacen_codigo']} - {self.req['almacen_nombre']}"),
            ("Estado", self.req["estado"]),
            ("Solicitante", self.req["solicitante"]),
            ("Centro Costo", f"{self.req['centro_costo_codigo']} - {self.req['centro_costo_nombre']}"),
            ("OT", self.req["ot"] or ""),
            ("Usuario", self.req["usuario_creador"]),
            ("Observación", self.req["observacion"]),
        ]
        for i, (k, v) in enumerate(data):
            r = i // 3
            c = (i % 3) * 2
            ttk.Label(top, text=f"{k}:", font=("Arial", 9, "bold")).grid(row=r, column=c, sticky="w", padx=4, pady=3)
            ttk.Label(top, text=str(v), wraplength=260).grid(row=r, column=c+1, sticky="w", padx=4, pady=3)

        mid = ttk.LabelFrame(self, text="Artículos", padding=8)
        mid.pack(fill="both", expand=True, padx=10, pady=6)
        cols = ("item", "codigo", "descripcion", "um", "req", "atend", "pend", "obs")
        frame = ttk.Frame(mid)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=10)
        ysb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        xsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        specs = [("item", "Item", 55), ("codigo", "Código", 100), ("descripcion", "Descripción", 360), ("um", "UM", 70), ("req", "Requerido", 90), ("atend", "Atendido", 90), ("pend", "Pendiente", 90), ("obs", "Obs.", 220)]
        for c, t, w in specs:
            tree.heading(c, text=t); tree.column(c, width=w, anchor="center" if c not in ["descripcion", "obs"] else "w")
        tree.grid(row=0, column=0, sticky="nsew"); ysb.grid(row=0, column=1, sticky="ns"); xsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1); frame.columnconfigure(0, weight=1)
        for r in self.det:
            pend = float(r["cantidad_requerida"]) - float(r["cantidad_atendida"])
            tree.insert("", "end", values=(r["item"], r["codigo"], r["descripcion"], r["unidad_medida"], f"{float(r['cantidad_requerida']):g}", f"{float(r['cantidad_atendida']):g}", f"{pend:g}", r["observacion"] or ""))

        hist = ttk.LabelFrame(self, text="Historial de atención", padding=8)
        hist.pack(fill="both", expand=True, padx=10, pady=6)
        cols2 = ("vale", "fecha", "usuario", "estado", "obs")
        frame2 = ttk.Frame(hist)
        frame2.pack(fill="both", expand=True)
        tree2 = ttk.Treeview(frame2, columns=cols2, show="headings", height=6)
        y2 = ttk.Scrollbar(frame2, orient="vertical", command=tree2.yview)
        x2 = ttk.Scrollbar(frame2, orient="horizontal", command=tree2.xview)
        tree2.configure(yscrollcommand=y2.set, xscrollcommand=x2.set)
        specs2 = [("vale", "Vale salida", 135), ("fecha", "Fecha/Hora", 150), ("usuario", "Usuario", 120), ("estado", "Estado atención", 150), ("obs", "Observación", 420)]
        for c, t, w in specs2:
            tree2.heading(c, text=t); tree2.column(c, width=w, anchor="center" if c != "obs" else "w")
        tree2.grid(row=0, column=0, sticky="nsew"); y2.grid(row=0, column=1, sticky="ns"); x2.grid(row=1, column=0, sticky="ew")
        frame2.rowconfigure(0, weight=1); frame2.columnconfigure(0, weight=1)
        for a in self.db.requerimiento_atenciones(self.req["id"]):
            tree2.insert("", "end", values=(a["vale"], a["fecha_atencion"], a["usuario"], a["estado_atencion"], a["observacion"]))
        ttk.Button(self, text="Cerrar (Esc)", command=self.destroy).pack(pady=8)


class ConsultaRequerimientosDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.title("Consulta / Reporte de Requerimientos")
        self.geometry("1180x700+70+35")
        self.minsize(1080, 620)
        bind_as_child_window(self, app.root)
        self.build()
        self.set_mes_actual()

    def build(self):
        filters = ttk.LabelFrame(self, text="Filtros", padding=10)
        filters.pack(fill="x", padx=10, pady=8)
        self.desde_var = tk.StringVar()
        self.hasta_var = tk.StringVar()
        self.estado_var = tk.StringVar(value="TODOS")
        self.almacen_var = tk.StringVar(value="TODOS")
        self.sol_var = tk.StringVar()
        self.cc_var = tk.StringVar()
        self.term_var = tk.StringVar()

        ttk.Label(filters, text="Desde:").grid(row=0, column=0, padx=4, pady=4, sticky="w")
        ttk.Entry(filters, textvariable=self.desde_var, width=12).grid(row=0, column=1, padx=4, sticky="w")
        ttk.Label(filters, text="Hasta:").grid(row=0, column=2, padx=4, sticky="w")
        ttk.Entry(filters, textvariable=self.hasta_var, width=12).grid(row=0, column=3, padx=4, sticky="w")
        ttk.Label(filters, text="Estado:").grid(row=0, column=4, padx=4, sticky="w")
        ttk.Combobox(filters, textvariable=self.estado_var, values=["TODOS"] + ESTADOS_REQUERIMIENTO, width=18, state="readonly").grid(row=0, column=5, padx=4, sticky="w")
        ttk.Label(filters, text="Almacén:").grid(row=0, column=6, padx=4, sticky="w")
        ttk.Combobox(filters, textvariable=self.almacen_var, values=["TODOS"] + list(ALMACENES.keys()), width=10, state="readonly").grid(row=0, column=7, padx=4, sticky="w")
        ttk.Button(filters, text="Hoy", command=self.set_hoy).grid(row=0, column=8, padx=4)
        ttk.Button(filters, text="Mes actual", command=self.set_mes_actual).grid(row=0, column=9, padx=4)
        ttk.Button(filters, text="Todos", command=self.set_todos).grid(row=0, column=10, padx=4)

        ttk.Label(filters, text="Solicitante:").grid(row=1, column=0, padx=4, pady=4, sticky="w")
        ttk.Entry(filters, textvariable=self.sol_var, width=22).grid(row=1, column=1, columnspan=2, padx=4, sticky="w")
        ttk.Label(filters, text="CC:").grid(row=1, column=3, padx=4, sticky="w")
        ttk.Entry(filters, textvariable=self.cc_var, width=8).grid(row=1, column=4, padx=4, sticky="w")
        ttk.Label(filters, text="Buscar Req/OT/Obs:").grid(row=1, column=5, padx=4, sticky="w")
        ttk.Entry(filters, textvariable=self.term_var, width=34).grid(row=1, column=6, columnspan=2, padx=4, sticky="w")
        ttk.Button(filters, text="Buscar", command=self.search).grid(row=1, column=8, padx=4)
        ttk.Button(filters, text="Limpiar", command=self.clear_filters).grid(row=1, column=9, padx=4)
        ttk.Button(filters, text="Exportar Excel", command=self.export_excel).grid(row=1, column=10, padx=4)
        for v in [self.estado_var, self.almacen_var]:
            pass
        for child in filters.winfo_children():
            try:
                child.bind("<Return>", lambda e: self.search())
            except Exception:
                pass

        body = ttk.Frame(self, padding=(10, 2, 10, 2))
        body.pack(fill="both", expand=True)
        cols = ("req", "fecha", "entrega", "almacen", "estado", "sol", "cc", "ot", "items", "reqtot", "atend", "pend")
        frame = ttk.Frame(body)
        frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=18)
        ysb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        specs = [("req", "Requerimiento", 140), ("fecha", "Fecha solicitud", 105), ("entrega", "Entrega solic.", 105), ("almacen", "Almacén", 80), ("estado", "Estado", 145), ("sol", "Solicitante", 210), ("cc", "Centro Costo", 190), ("ot", "OT", 120), ("items", "Items", 65), ("reqtot", "Req.", 80), ("atend", "Atend.", 80), ("pend", "Pend.", 80)]
        for c, t, w in specs:
            self.tree.heading(c, text=t); self.tree.column(c, width=w, anchor="center" if c not in ["sol", "cc"] else "w")
        self.tree.grid(row=0, column=0, sticky="nsew"); ysb.grid(row=0, column=1, sticky="ns"); xsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1); frame.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", lambda e: self.view_detail())
        self.req_menu = tk.Menu(self, tearoff=0)
        self.req_menu.add_command(label="Ver detalle", command=self.view_detail)
        self.req_menu.add_command(label="Ver estado operativo", command=self.view_estado_operativo)
        self.tree.bind("<Button-3>", self.show_req_menu)

        buttons = ttk.Frame(self, padding=10)
        buttons.pack(fill="x")
        self.count_lbl = ttk.Label(buttons, text="")
        self.count_lbl.pack(side="left", padx=5)
        ttk.Button(buttons, text="Ver detalle", command=self.view_detail).pack(side="right", padx=5)
        ttk.Button(buttons, text="Cerrar (Esc)", command=self.destroy).pack(side="right", padx=5)

    def set_hoy(self):
        self.desde_var.set(today_text()); self.hasta_var.set(today_text()); self.search()

    def set_mes_actual(self):
        first = date.today().replace(day=1)
        if first.month == 12:
            next_month = date(first.year + 1, 1, 1)
        else:
            next_month = date(first.year, first.month + 1, 1)
        last = next_month - timedelta(days=1)
        self.desde_var.set(first.strftime(DATE_FMT)); self.hasta_var.set(last.strftime(DATE_FMT)); self.search()

    def set_todos(self):
        self.desde_var.set(""); self.hasta_var.set(""); self.search()

    def clear_filters(self):
        self.desde_var.set(""); self.hasta_var.set(""); self.estado_var.set("TODOS"); self.almacen_var.set("TODOS")
        self.sol_var.set(""); self.cc_var.set(""); self.term_var.set(""); self.search()

    def _rows(self):
        return self.db.list_requerimientos_filtered(
            self.desde_var.get().strip(), self.hasta_var.get().strip(), self.estado_var.get(), self.almacen_var.get(),
            self.sol_var.get().strip(), self.cc_var.get().strip(), self.term_var.get().strip()
        )

    def search(self):
        try:
            for x in self.tree.get_children(): self.tree.delete(x)
            rows = self._rows()
            total_pend = 0
            for r in rows:
                total_pend += float(r["total_pendiente"] or 0)
                self.tree.insert("", "end", values=(r["requerimiento"], fecha_text_from_iso(r["fecha_requerimiento"]), fecha_text_from_iso(r["fecha_entrega_solicitada"]) if "fecha_entrega_solicitada" in r.keys() and r["fecha_entrega_solicitada"] else "", r["almacen_codigo"], r["estado"], r["solicitante"], f"{r['centro_costo_codigo']} - {r['centro_costo_nombre']}", r["ot"] or "", r["total_items"], f"{float(r['total_requerido']):g}", f"{float(r['total_atendido']):g}", f"{float(r['total_pendiente']):g}"))
            self.count_lbl.configure(text=f"Registros: {len(rows)} | Pendiente total: {total_pend:g}")
        except Exception as e:
            messagebox.showerror("Error en consulta", str(e), parent=self)
            self.focus_force()

    def selected_req(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Seleccione", "Seleccione un requerimiento.", parent=self)
            return None
        return self.tree.item(sel[0], "values")[0]

    def view_detail(self):
        req = self.selected_req()
        if req:
            RequerimientoDetailDialog(self, self.app, req)

    def show_req_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self.req_menu.tk_popup(event.x_root, event.y_root)

    def view_estado_operativo(self):
        req = self.selected_req()
        if req:
            RequerimientoEstadoOperativoDialog(self, self.app, req)

    def export_excel(self):
        try:
            headers = ["requerimiento", "fecha_requerimiento", "fecha_entrega_solicitada", "almacen_codigo", "almacen_nombre", "estado", "solicitante", "centro_costo_codigo", "centro_costo_nombre", "ot", "observacion", "item", "codigo", "descripcion", "unidad_medida", "cantidad_requerida", "cantidad_atendida", "pendiente", "usuario_creador", "creado_en", "modificado_en"]
            titles = ["Requerimiento", "Fecha solicitud", "Fecha entrega solicitada", "Almacén", "Almacén nombre", "Estado", "Solicitante", "CC", "Centro costo", "OT", "Observación", "Item", "Código", "Descripción", "UM", "Cant. requerida", "Cant. atendida", "Pendiente", "Usuario creador", "Creado en", "Modificado en"]
            rows = self.db.requerimiento_export_rows_filtered(self.desde_var.get().strip(), self.hasta_var.get().strip(), self.estado_var.get(), self.almacen_var.get(), self.sol_var.get().strip(), self.cc_var.get().strip(), self.term_var.get().strip())
            self.app.export_excel(rows, headers, titles, f"reporte_general_requerimientos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        except Exception as e:
            messagebox.showerror("No se pudo exportar", str(e), parent=self)
            self.focus_force()


class RequerimientoDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.detalles = []
        self.selected_index = None
        self.items_enabled = False
        self.title("Logística - Requerimientos")
        self.geometry("1120x720+80+40")
        self.minsize(1040, 650)
        bind_as_child_window(self, app.root)
        self.build()
        self.after(200, self.almacen_combo.focus_set)
        self.bind("<Shift-F2>", lambda e: (StockDialog(self.app, parent=self), "break")[1])
        self.bind("<Shift-F1>", self.shift_f1_handler)

    def build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        self.fecha_var = tk.StringVar(value=today_text())
        self.fecha_entrega_var = tk.StringVar(value=today_text())
        self.edit_fecha = tk.BooleanVar(value=False)
        self.almacen_var = tk.StringVar(value=getattr(self.app, "almacen_actual", "AQP"))
        self.solicitante_var = tk.StringVar()
        self.cc_code_var = tk.StringVar()
        self.cc_name_var = tk.StringVar()
        self.ot_var = tk.StringVar()
        self.obs_var = tk.StringVar()
        ttk.Label(top, text="Fecha requerimiento:").grid(row=0, column=0, padx=4, pady=6, sticky="w")
        self.fecha_entry = ttk.Entry(top, textvariable=self.fecha_var, width=13, state="readonly")
        self.fecha_entry.grid(row=0, column=1, padx=4, sticky="w")
        ttk.Checkbutton(top, text="Editar fecha", variable=self.edit_fecha, command=self.toggle_fecha).grid(row=0, column=2, sticky="w", padx=4)
        ttk.Label(top, text="Fecha entrega solicitada:").grid(row=0, column=3, padx=4, sticky="w")
        self.fecha_entrega_entry = ttk.Entry(top, textvariable=self.fecha_entrega_var, width=13)
        self.fecha_entrega_entry.grid(row=0, column=4, padx=4, sticky="w")
        ttk.Label(top, text="Almacén destino:").grid(row=0, column=5, padx=4, sticky="w")
        self.almacen_combo = ttk.Combobox(top, textvariable=self.almacen_var, values=list(ALMACENES.keys()), width=8, state="readonly")
        self.almacen_combo.grid(row=0, column=6, padx=4, sticky="w")
        ttk.Label(top, text="Solicitante:").grid(row=1, column=0, padx=4, sticky="w")
        self.solicitante_combo = ttk.Combobox(top, textvariable=self.solicitante_var, values=self.db.solicitante_names(), width=34, state="normal")
        self.solicitante_combo.grid(row=1, column=1, columnspan=3, padx=4, sticky="w")
        self.solicitante_combo.bind("<Shift-F1>", lambda e: self.open_solicitante_help())
        attach_solicitante_autocomplete(self.solicitante_combo, self.db, self.solicitante_var, self.set_solicitante_from_help)
        ttk.Label(top, text="Centro Costo:").grid(row=1, column=4, padx=4, pady=6, sticky="w")
        self.cc_entry = ttk.Entry(top, textvariable=self.cc_code_var, width=8)
        self.cc_entry.grid(row=1, column=5, padx=4, sticky="w")
        self.cc_entry.bind("<KeyRelease>", self.fill_cc_from_typed)
        self.cc_entry.bind("<Shift-F1>", lambda e: self.open_cc_help())
        ttk.Entry(top, textvariable=self.cc_name_var, width=28, state="readonly").grid(row=1, column=6, columnspan=2, padx=4, sticky="w")
        ttk.Label(top, text="OT-").grid(row=2, column=0, padx=4, sticky="w")
        self.ot_entry = ttk.Entry(top, textvariable=self.ot_var, width=18)
        self.ot_entry.grid(row=2, column=1, padx=4, sticky="w")
        ttk.Label(top, text="Observación:").grid(row=2, column=2, padx=4, pady=6, sticky="w")
        self.obs_entry = ttk.Entry(top, textvariable=self.obs_var, width=78)
        self.obs_entry.grid(row=2, column=3, columnspan=5, padx=4, sticky="we")
        self.fecha_entrega_entry.bind("<Return>", lambda e: self.almacen_combo.focus_set())
        self.almacen_combo.bind("<Return>", lambda e: self.solicitante_combo.focus_set())
        self.solicitante_combo.bind("<Return>", self.enter_solicitante)
        self.cc_entry.bind("<Return>", self.enter_cc)
        self.ot_entry.bind("<Return>", lambda e: self.obs_entry.focus_set() or "break")
        self.obs_entry.bind("<Return>", self.enter_obs)

        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        columns = ("item", "codigo", "descripcion", "um", "cantidad")
        tree_frame = ttk.Frame(body)
        tree_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=14)
        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        specs = [("item", "Item", 60), ("codigo", "Código", 120), ("descripcion", "Descripción", 570), ("um", "UM", 90), ("cantidad", "Cantidad requerida", 150)]
        for c, t, w in specs:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c != "descripcion" else "w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1); tree_frame.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self.select_row)
        form = ttk.Frame(self, padding=10)
        form.pack(fill="x")
        self.cod_var = tk.StringVar(); self.desc_var = tk.StringVar(); self.um_var = tk.StringVar(); self.qty_var = tk.StringVar()
        ttk.Label(form, text="Código:").grid(row=0, column=0, padx=4)
        self.cod_entry = ttk.Entry(form, textvariable=self.cod_var, width=14, state="disabled")
        self.cod_entry.grid(row=0, column=1, padx=4)
        self.cod_entry.bind("<Return>", self.load_code)
        self.cod_entry.bind("<Shift-F1>", lambda e: self.open_articulo_help())
        self.cod_entry.bind("<Shift-F2>", lambda e: self.open_stock_help())
        attach_articulo_autocomplete(self.cod_entry, self.db, self.cod_var, self.set_codigo_from_help)
        ttk.Label(form, text="Descripción:").grid(row=0, column=2, padx=4)
        ttk.Entry(form, textvariable=self.desc_var, width=50, state="readonly").grid(row=0, column=3, padx=4)
        ttk.Label(form, text="UM:").grid(row=0, column=4, padx=4)
        ttk.Entry(form, textvariable=self.um_var, width=8, state="readonly").grid(row=0, column=5, padx=4)
        ttk.Label(form, text="Cantidad:").grid(row=0, column=6, padx=4)
        self.qty_entry = ttk.Entry(form, textvariable=self.qty_var, width=12, state="disabled")
        self.qty_entry.grid(row=0, column=7, padx=4)
        self.qty_entry.bind("<Return>", self.add_item)
        buttons = ttk.Frame(self, padding=10)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="✖ Eliminar", command=self.delete_item).pack(side="left", padx=5)
        ttk.Button(buttons, text="📖 Modificar", command=self.modify_item).pack(side="left", padx=5)
        ttk.Button(buttons, text="💾 Guardar Requerimiento", command=self.save).pack(side="right", padx=5)

    def toggle_fecha(self):
        self.fecha_entry.configure(state="normal" if self.edit_fecha.get() else "readonly")
        if not self.edit_fecha.get():
            self.fecha_var.set(today_text())

    def shift_f1_handler(self, event=None):
        widget = self.focus_get()
        if widget == self.cc_entry:
            return self.open_cc_help()
        if widget == self.solicitante_combo:
            return self.open_solicitante_help()
        if widget == self.cod_entry:
            return self.open_stock_help()
        return "break"

    def open_cc_help(self):
        CenterCostDialog(self, self.set_cc)
        return "break"

    def open_solicitante_help(self):
        SolicitanteSelectDialog(self, self.db, self.set_solicitante_from_help)
        return "break"

    def set_solicitante_from_help(self, nombre):
        self.solicitante_var.set(nombre)
        self.cc_entry.focus_set()

    def open_articulo_help(self):
        ArticuloSelectDialog(self, self.db, self.set_codigo_from_help)
        return "break"

    def open_stock_help(self):
        StockDialog(self.app, parent=self, on_select=self.set_codigo_from_help)
        return "break"

    def set_codigo_from_help(self, codigo):
        self.cod_var.set(codigo)
        self.load_code()

    def set_cc(self, codigo, nombre):
        self.cc_code_var.set(codigo); self.cc_name_var.set(nombre); self.ot_entry.focus_set()

    def fill_cc_from_typed(self, event=None):
        code = normalize_text(self.cc_code_var.get())
        self.cc_name_var.set(CENTROS_COSTO.get(code, ""))

    def enter_solicitante(self, event=None):
        self.solicitante_var.set(self.solicitante_var.get().strip().upper())
        if not self.solicitante_var.get():
            messagebox.showwarning("Solicitante", "Ingrese solicitante.", parent=self); self.solicitante_combo.focus_set(); return "break"
        self.cc_entry.focus_set(); return "break"

    def enter_cc(self, event=None):
        self.fill_cc_from_typed()
        if not self.cc_name_var.get():
            messagebox.showwarning("Centro de costo", "Ingrese un centro de costo válido o use Shift + F1.", parent=self); self.cc_entry.focus_set(); return "break"
        self.ot_entry.focus_set(); return "break"

    def enter_obs(self, event=None):
        if not self.obs_var.get().strip():
            messagebox.showwarning("Observación", "Ingrese una observación.", parent=self); self.obs_entry.focus_set(); return "break"
        self.obs_var.set(self.obs_var.get().strip().upper())
        self.items_enabled = True
        self.cod_entry.configure(state="normal")
        self.qty_entry.configure(state="normal")
        self.cod_entry.focus_set(); return "break"

    def load_code(self, event=None):
        if not self.items_enabled:
            messagebox.showwarning("Complete cabecera", "Complete la cabecera antes de ingresar artículos.", parent=self); return "break"
        art = self.db.get_articulo(self.cod_var.get())
        if not art:
            self.desc_var.set(""); self.um_var.set(""); self.qty_var.set("")
            messagebox.showwarning("No existe", "El código no existe. El requerimiento continuará abierto.", parent=self)
            self.cod_entry.focus_set(); return "break"
        self.cod_var.set(art["codigo"]); self.desc_var.set(art["descripcion"]); self.um_var.set(art["unidad_medida"])
        self.qty_entry.focus_set(); return "break"

    def add_item(self, event=None):
        if not self.desc_var.get():
            self.load_code()
            if not self.desc_var.get():
                return "break"
        codigo = normalize_text(self.cod_var.get())
        try:
            qty, _art = self.db.validar_cantidad_articulo(codigo, self.qty_var.get(), "cantidad requerida")
        except Exception as e:
            messagebox.showwarning("Cantidad inválida", str(e), parent=self); self.qty_entry.focus_set(); return "break"
        self.detalles.append({"item": len(self.detalles)+1, "codigo": codigo, "descripcion": self.desc_var.get(), "um": self.um_var.get(), "cantidad": qty})
        self.refresh_items(); self.cod_var.set(""); self.desc_var.set(""); self.um_var.set(""); self.qty_var.set(""); self.cod_entry.focus_set(); return "break"

    def refresh_items(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        for i, d in enumerate(self.detalles, start=1):
            d["item"] = i
            self.tree.insert("", "end", values=(i, d["codigo"], d["descripcion"], d["um"], f"{d['cantidad']:g}"))

    def select_row(self, event=None):
        sel = self.tree.selection(); self.selected_index = self.tree.index(sel[0]) if sel else None

    def delete_item(self):
        if self.selected_index is None:
            messagebox.showinfo("Seleccione", "Seleccione un item para eliminar.", parent=self); return
        if messagebox.askyesno("Eliminar", "¿Eliminar el artículo seleccionado?", parent=self):
            self.detalles.pop(self.selected_index); self.selected_index = None; self.refresh_items()

    def modify_item(self):
        if self.selected_index is None:
            messagebox.showinfo("Seleccione", "Seleccione un item para modificar.", parent=self); return
        d = self.detalles[self.selected_index]
        new_qty = simpledialog.askstring("Modificar cantidad", f"Código {d['codigo']} - Nueva cantidad requerida:", initialvalue=str(d["cantidad"]), parent=self)
        if new_qty is None: return
        try:
            qty = float(new_qty.replace(",", "."))
            if qty <= 0: raise ValueError
            d["cantidad"] = qty; self.refresh_items()
        except Exception:
            messagebox.showwarning("Inválido", "Cantidad inválida.", parent=self)

    def save(self):
        try:
            if not self.app.require_write(self):
                return
            self.fill_cc_from_typed()
            req = self.db.create_requerimiento(self.fecha_var.get(), self.almacen_var.get(), self.solicitante_var.get(), self.cc_code_var.get(), self.ot_var.get(), self.obs_var.get(), self.detalles, self.app.usuario_actual, self.fecha_entrega_var.get())
            messagebox.showinfo("Requerimiento generado", f"Requerimiento guardado correctamente:\n{req}\nEstado: PENDIENTE", parent=self)
            self.fecha_var.set(today_text()); self.fecha_entrega_var.set(today_text()); self.edit_fecha.set(False); self.toggle_fecha(); self.almacen_var.set("AQP")
            self.solicitante_var.set(""); self.cc_code_var.set(""); self.cc_name_var.set(""); self.ot_var.set(""); self.obs_var.set("")
            self.detalles = []; self.refresh_items(); self.items_enabled = False
            self.cod_var.set(""); self.desc_var.set(""); self.um_var.set(""); self.qty_var.set("")
            self.cod_entry.configure(state="disabled"); self.qty_entry.configure(state="disabled")
            self.almacen_combo.focus_set()
        except Exception as e:
            messagebox.showerror("No se pudo guardar", str(e), parent=self); self.focus_force()


class AtencionRequerimientoDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.req = None
        self.detalles = []
        self.vars_atender = {}
        self.title("Almacén - Atención de Requerimientos")
        self.geometry("1120x720+80+40")
        self.minsize(1040, 650)
        bind_as_child_window(self, app.root)
        self.build()
        self.req_entry.focus_set()
        self.bind("<Shift-F1>", lambda e: self.open_req_list())

    def build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        self.req_var = tk.StringVar()
        self.fecha_var = tk.StringVar(value=today_text())
        self.obs_atencion_var = tk.StringVar(value="ATENCION DE REQUERIMIENTO")
        ttk.Label(top, text="N° Requerimiento:").grid(row=0, column=0, padx=4, pady=6, sticky="w")
        self.req_entry = ttk.Entry(top, textvariable=self.req_var, width=22)
        self.req_entry.grid(row=0, column=1, padx=4, sticky="w")
        self.req_entry.bind("<Return>", lambda e: self.load_req())
        self.req_entry.bind("<Shift-F1>", lambda e: self.open_req_list())
        ttk.Button(top, text="Buscar", command=self.open_req_list).grid(row=0, column=2, padx=4)
        ttk.Label(top, text="Enter carga el requerimiento | Shift + F1 busca pendientes/parciales", foreground="#666666").grid(row=0, column=3, columnspan=4, sticky="w")
        ttk.Label(top, text="Fecha atención:").grid(row=1, column=0, padx=4, pady=6, sticky="w")
        ttk.Entry(top, textvariable=self.fecha_var, width=13).grid(row=1, column=1, padx=4, sticky="w")
        self.info_lbl = ttk.Label(top, text="Cargue un requerimiento pendiente o parcial.", foreground="#333333")
        self.info_lbl.grid(row=2, column=0, columnspan=7, sticky="w", padx=4, pady=6)
        ttk.Label(top, text="Observación atención:").grid(row=3, column=0, padx=4, pady=6, sticky="w")
        ttk.Entry(top, textvariable=self.obs_atencion_var, width=92).grid(row=3, column=1, columnspan=6, padx=4, sticky="we")
        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        cols = ("item", "codigo", "descripcion", "um", "req", "atend", "pend", "stock", "cant")
        tree_frame = ttk.Frame(body)
        tree_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15)
        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        specs = [("item", "Item", 55), ("codigo", "Código", 100), ("descripcion", "Descripción", 410), ("um", "UM", 70), ("req", "Requerido", 95), ("atend", "Atendido", 95), ("pend", "Pendiente", 95), ("stock", "Stock", 90), ("cant", "A atender", 100)]
        for c, t, w in specs:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c != "descripcion" else "w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1); tree_frame.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self.modify_selected)
        buttons = ttk.Frame(self, padding=10)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Modificar cantidad seleccionada", command=self.modify_selected).pack(side="left", padx=5)
        ttk.Button(buttons, text="💾 Guardar atención / Generar CO", command=self.save).pack(side="right", padx=5)

    def open_req_list(self):
        RequerimientoListDialog(self, self.db, self.set_req_from_list)
        return "break"

    def set_req_from_list(self, req):
        self.req_var.set(req); self.load_req()

    def load_req(self):
        try:
            req, det = self.db.find_requerimiento(self.req_var.get())
            if not req:
                messagebox.showwarning("No encontrado", "No se encontró el requerimiento.", parent=self); self.req_entry.focus_set(); return
            if req["estado"] in ["ATENDIDO", "ANULADO"]:
                messagebox.showwarning("No permitido", f"El requerimiento está {req['estado']} y no puede atenderse.", parent=self); self.req_entry.focus_set(); return
            self.req = req
            self.detalles = [r for r in det if (float(r["cantidad_requerida"]) - float(r["cantidad_atendida"])) > 0]
            self.vars_atender = {}
            for r in self.detalles:
                pendiente = float(r["cantidad_requerida"]) - float(r["cantidad_atendida"])
                stock = self.db.get_stock(r["codigo"], req["almacen_codigo"])
                self.vars_atender[r["codigo"]] = min(pendiente, stock)
            self.info_lbl.configure(text=f"{req['requerimiento']} | {req['almacen_codigo']} | Estado: {req['estado']} | Solicitante: {req['solicitante']} | CC: {req['centro_costo_codigo']} - {req['centro_costo_nombre']} | OT: {req['ot'] or ''}")
            self.refresh_items()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def refresh_items(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        if not self.req:
            return
        for r in self.detalles:
            pendiente = float(r["cantidad_requerida"]) - float(r["cantidad_atendida"])
            stock = self.db.get_stock(r["codigo"], self.req["almacen_codigo"])
            atender = self.vars_atender.get(r["codigo"], 0)
            self.tree.insert("", "end", values=(r["item"], r["codigo"], r["descripcion"], r["unidad_medida"], f"{float(r['cantidad_requerida']):g}", f"{float(r['cantidad_atendida']):g}", f"{pendiente:g}", f"{stock:g}", f"{atender:g}"))

    def modify_selected(self, event=None):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Seleccione", "Seleccione un item para modificar la cantidad a atender.", parent=self); return
        values = self.tree.item(sel[0], "values")
        codigo = values[1]
        actual = self.vars_atender.get(codigo, 0)
        new_qty = simpledialog.askstring("Cantidad a atender", f"Código {codigo}\nCantidad a atender:", initialvalue=str(actual), parent=self)
        if new_qty is None:
            return
        try:
            qty = float(new_qty.replace(",", "."))
            if qty < 0: raise ValueError
            pendiente = 0
            for r in self.detalles:
                if r["codigo"] == codigo:
                    pendiente = float(r["cantidad_requerida"]) - float(r["cantidad_atendida"])
                    break
            if qty > pendiente:
                raise ValueError(f"No puede atender más que lo pendiente: {pendiente:g}")
            self.vars_atender[codigo] = qty
            self.refresh_items()
        except Exception as e:
            messagebox.showwarning("Cantidad inválida", str(e), parent=self)

    def save(self):
        try:
            if not self.app.require_write(self):
                return
            if not self.req:
                messagebox.showwarning("Sin requerimiento", "Primero cargue un requerimiento.", parent=self); self.req_entry.focus_set(); return
            vale, estado = self.db.attend_requerimiento(self.req["requerimiento"], self.fecha_var.get(), self.vars_atender, self.obs_atencion_var.get(), self.app.usuario_actual)
            messagebox.showinfo("Atención registrada", f"Vale de salida generado: {vale}\nEstado del requerimiento: {estado}", parent=self)
            if messagebox.askyesno("Imprimir documento", "¿Desea imprimir documento?", parent=self):
                mov, det = self.db.get_movimiento_by_vale(vale)
                if mov:
                    PrintPreviewDialog(self, mov, det)
            self.req = None; self.detalles = []; self.vars_atender = {}
            self.req_var.set(""); self.fecha_var.set(today_text()); self.obs_atencion_var.set("ATENCION DE REQUERIMIENTO")
            self.info_lbl.configure(text="Cargue un requerimiento pendiente o parcial.")
            self.refresh_items(); self.req_entry.focus_set()
        except Exception as e:
            messagebox.showerror("No se pudo atender", str(e), parent=self); self.focus_force()

class SolicitudesDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.title("Solicitantes")
        self.geometry("620x420+240+130")
        self.selected_id = None
        bind_as_child_window(self, app.root)
        self.build()
        self.refresh()

    def build(self):
        ttk.Label(self, text="Agregar o eliminar solicitantes para el registro de Ingreso/Salida", padding=10).pack(anchor="w")
        cols = ("id", "nombre", "activo", "creado")
        tree_box = ttk.Frame(self)
        tree_box.pack(fill="both", expand=True, padx=10, pady=8)
        tree_box.rowconfigure(0, weight=1); tree_box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_box, columns=cols, show="headings", height=10)
        specs = [("id", "ID", 50), ("nombre", "Solicitante", 300), ("activo", "Activo", 80), ("creado", "Creado", 150)]
        for c, t, w in specs:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c != "nombre" else "w", stretch=False)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.select)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="x")
        self.nombre = tk.StringVar()
        ttk.Label(frm, text="Nombre:").grid(row=0, column=0, padx=4)
        entry = ttk.Entry(frm, textvariable=self.nombre, width=35)
        entry.grid(row=0, column=1, padx=4)
        entry.bind("<Return>", lambda e: self.add())
        ttk.Button(frm, text="Agregar", command=self.add).grid(row=0, column=2, padx=4)
        ttk.Button(frm, text="Eliminar / Desactivar", command=self.delete).grid(row=0, column=3, padx=4)
        entry.focus_set()

    def refresh(self):
        for x in self.tree.get_children():
            self.tree.delete(x)
        for r in self.db.list_solicitantes(False):
            self.tree.insert("", "end", values=(r["id"], r["nombre"], "SI" if r["activo"] else "NO", r["creado_en"]))

    def select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            self.selected_id = None
            return
        values = self.tree.item(sel[0], "values")
        self.selected_id = int(values[0])
        self.nombre.set(values[1])

    def add(self):
        if not self.app.require_write(self):
            return
        try:
            nuevo_solicitante = normalize_text(self.nombre.get())
            self.db.add_solicitante(nuevo_solicitante)
            self.db.audit("solicitantes", nuevo_solicitante, "CREACION", self.app.usuario_actual, "Registro de solicitante")
            messagebox.showinfo("Completo", "Solicitante agregado correctamente.", parent=self)
            self.nombre.set("")
            self.selected_id = None
            self.refresh()
            # Resaltar el solicitante recién creado/reactivado en la lista.
            for iid in self.tree.get_children():
                vals = self.tree.item(iid, "values")
                if vals and str(vals[1]).strip().lower() == nuevo_solicitante.strip().lower():
                    self.tree.selection_set(iid)
                    self.tree.focus(iid)
                    self.tree.see(iid)
                    break
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)
            self.after(100, self.focus_force)

    def delete(self):
        if not self.app.require_write(self):
            return
        if not self.selected_id:
            messagebox.showinfo("Seleccione", "Seleccione un solicitante.", parent=self)
            return
        if not messagebox.askyesno("Confirmar", "¿Eliminar/desactivar el solicitante seleccionado?", parent=self):
            return
        try:
            nombre = self.nombre.get()
            self.db.deactivate_solicitante(self.selected_id)
            self.db.audit("solicitantes", nombre, "DESACTIVACION", self.app.usuario_actual, "Desactivación de solicitante")
            self.nombre.set("")
            self.selected_id = None
            self.refresh()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)



class SolicitanteSelectDialog(tk.Toplevel):
    def __init__(self, parent, db, on_select):
        super().__init__(parent)
        self.db = db
        self.on_select = on_select
        self.term = tk.StringVar()
        self.title("Buscar solicitante")
        self.geometry("560x420+260+130")
        bind_as_child_window(self, parent)
        self.build(); self.search()

    def build(self):
        top = ttk.Frame(self, padding=10); top.pack(fill="x")
        ttk.Label(top, text="Buscar:").pack(side="left")
        e = ttk.Entry(top, textvariable=self.term, width=40); e.pack(side="left", padx=6); e.focus_set()
        e.bind("<KeyRelease>", lambda e: self.search())
        e.bind("<Return>", lambda e: self.choose())
        tree_box = ttk.Frame(self)
        tree_box.pack(fill="both", expand=True, padx=10, pady=8)
        tree_box.rowconfigure(0, weight=1); tree_box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_box, columns=("id", "nombre"), show="headings", height=14)
        self.tree.heading("id", text="ID"); self.tree.column("id", width=60, anchor="center", stretch=False)
        self.tree.heading("nombre", text="Solicitante"); self.tree.column("nombre", width=430, anchor="w", stretch=False)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self.choose)
        self.tree.bind("<Return>", self.choose)
        ttk.Button(self, text="Seleccionar", command=self.choose).pack(side="left", padx=10, pady=8)
        ttk.Button(self, text="Cerrar", command=self.destroy).pack(side="right", padx=10, pady=8)

    def search(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        term = normalize_text(self.term.get())
        first = None
        for r in self.db.list_solicitantes(True):
            if term in normalize_text(r["nombre"]):
                iid = self.tree.insert("", "end", values=(r["id"], r["nombre"]))
                if first is None:
                    first = iid
        if first:
            self.tree.selection_set(first)
            self.tree.focus(first)

    def choose(self, event=None):
        sel = self.tree.selection()
        if not sel: return "break"
        nombre = self.tree.item(sel[0], "values")[1]
        self.on_select(nombre)
        self.destroy()
        return "break"


class CentroCostoSelectDialog(tk.Toplevel):
    def __init__(self, parent, on_select):
        super().__init__(parent)
        self.on_select = on_select
        self.term = tk.StringVar()
        self.title("Buscar centro de costo")
        self.geometry("520x360+280+150")
        bind_as_child_window(self, parent)
        self.build(); self.search()

    def build(self):
        top = ttk.Frame(self, padding=10); top.pack(fill="x")
        ttk.Label(top, text="Buscar:").pack(side="left")
        e = ttk.Entry(top, textvariable=self.term, width=36); e.pack(side="left", padx=6); e.focus_set()
        e.bind("<KeyRelease>", lambda e: self.search())
        e.bind("<Return>", lambda e: self.choose())
        tree_box = ttk.Frame(self)
        tree_box.pack(fill="both", expand=True, padx=10, pady=8)
        tree_box.rowconfigure(0, weight=1); tree_box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_box, columns=("codigo", "nombre"), show="headings", height=10)
        self.tree.heading("codigo", text="Código"); self.tree.column("codigo", width=80, anchor="center", stretch=False)
        self.tree.heading("nombre", text="Centro de costo"); self.tree.column("nombre", width=370, anchor="w", stretch=False)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self.choose)
        self.tree.bind("<Return>", self.choose)
        ttk.Button(self, text="Seleccionar", command=self.choose).pack(side="left", padx=10, pady=8)
        ttk.Button(self, text="Cerrar", command=self.destroy).pack(side="right", padx=10, pady=8)

    def search(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        term = normalize_text(self.term.get())
        first = None
        for cod, nombre in CENTROS_COSTO.items():
            if term in normalize_text(cod + " " + nombre):
                iid = self.tree.insert("", "end", values=(cod, nombre))
                if first is None:
                    first = iid
        if first:
            self.tree.selection_set(first)
            self.tree.focus(first)

    def choose(self, event=None):
        sel = self.tree.selection()
        if not sel: return "break"
        codigo = self.tree.item(sel[0], "values")[0]
        self.on_select(codigo)
        self.destroy()
        return "break"


class ArticuloSelectDialog(tk.Toplevel):
    def __init__(self, parent, db, on_select):
        super().__init__(parent)
        self.db = db
        self.on_select = on_select
        self.term = tk.StringVar()
        self.title("Buscar artículo")
        self.geometry("900x500+130+100")
        bind_as_child_window(self, parent)
        self.build(); self.search()

    def build(self):
        top = ttk.Frame(self, padding=10); top.pack(fill="x")
        ttk.Label(top, text="Código / descripción:").pack(side="left")
        e = ttk.Entry(top, textvariable=self.term, width=46); e.pack(side="left", padx=6); e.focus_set()
        e.bind("<KeyRelease>", lambda e: self.search())
        e.bind("<Return>", lambda e: self.choose())
        cols=("codigo","descripcion","um","stock_aqp","stock_min")
        tree_box=ttk.Frame(self); tree_box.pack(fill="both", expand=True, padx=10, pady=8)
        tree_box.rowconfigure(0,weight=1); tree_box.columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(tree_box, columns=cols, show="headings", height=16)
        for c,t,w in [("codigo","Código",100),("descripcion","Descripción",360),("um","UM",70),("stock_aqp","Stock AQP",100),("stock_min","Stock MIN",100)]:
            self.tree.heading(c,text=t); self.tree.column(c,width=w,minwidth=w,anchor="center" if c!="descripcion" else "w",stretch=False)
        vsb=ttk.Scrollbar(tree_box,orient="vertical",command=self.tree.yview); hsb=ttk.Scrollbar(tree_box,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        self.tree.bind("<Double-1>", self.choose)
        self.tree.bind("<Return>", self.choose)
        btn=ttk.Frame(self,padding=8); btn.pack(fill="x")
        ttk.Button(btn,text="Seleccionar",command=self.choose).pack(side="left",padx=4)
        ttk.Button(btn,text="Cerrar",command=self.destroy).pack(side="right",padx=4)

    def search(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        first = None
        for r in self.db.search_articulos(self.term.get()):
            iid = self.tree.insert("", "end", values=(r["codigo"], r["descripcion"], r["unidad_medida"], f"{r['stock_aqp']:g}", f"{r['stock_min']:g}"))
            if first is None:
                first = iid
        if first:
            self.tree.selection_set(first)
            self.tree.focus(first)

    def choose(self, event=None):
        sel = self.tree.selection()
        if not sel: return "break"
        codigo = self.tree.item(sel[0], "values")[0]
        self.on_select(codigo)
        self.destroy()
        return "break"


class OCSelectDialog(tk.Toplevel):
    def __init__(self, parent, db, on_select, estados=None, tipo="TODOS", titulo="Buscar orden de compra"):
        super().__init__(parent)
        self.db = db
        self.on_select = on_select
        self.estados = tuple(estados or [])
        self.tipo = tipo
        self.term = tk.StringVar()
        self.title(titulo)
        self.geometry("980x500+120+110")
        bind_as_child_window(self, parent)
        self.build(); self.search()

    def build(self):
        top=ttk.Frame(self,padding=10); top.pack(fill="x")
        ttk.Label(top,text="Buscar OC / proveedor / requerimiento:").pack(side="left")
        e=ttk.Entry(top,textvariable=self.term,width=46); e.pack(side="left",padx=6); e.focus_set()
        e.bind("<KeyRelease>", lambda e: self.search())
        e.bind("<Return>", lambda e: self.choose())
        cols=("numero","fecha","proveedor","tipo","estado","total","req")
        tree_box=ttk.Frame(self); tree_box.pack(fill="both",expand=True,padx=10,pady=8)
        tree_box.rowconfigure(0,weight=1); tree_box.columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(tree_box,columns=cols,show="headings",height=15)
        for c,t,w in [("numero","OC",135),("fecha","Fecha",90),("proveedor","Proveedor",280),("tipo","Tipo",90),("estado","Estado",150),("total","Total",90),("req","Req/Cot",130)]:
            self.tree.heading(c,text=t); self.tree.column(c,width=w,minwidth=w,anchor="center" if c not in ("proveedor",) else "w",stretch=False)
        vsb=ttk.Scrollbar(tree_box,orient="vertical",command=self.tree.yview); hsb=ttk.Scrollbar(tree_box,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        self.tree.bind("<Double-1>", self.choose)
        self.tree.bind("<Return>", self.choose)
        btn=ttk.Frame(self,padding=8); btn.pack(fill="x")
        ttk.Button(btn,text="Seleccionar",command=self.choose).pack(side="left",padx=4)
        ttk.Button(btn,text="Cerrar",command=self.destroy).pack(side="right",padx=4)

    def search(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        first = None
        rows = self.db.list_ordenes_compra(self.term.get(), estado="TODOS", tipo=self.tipo, orden="fecha_desc")
        for r in rows:
            if self.estados and r["estado"] not in self.estados:
                continue
            iid = self.tree.insert("", "end", values=(r["numero"], fecha_text_from_iso(r["fecha_orden"]), r["razon_social"], r["tipo_orden"], r["estado"], f"{r['total']:.2f}", r["requerimiento"] or r["cotizacion"]))
            if first is None:
                first = iid
        if first:
            self.tree.selection_set(first)
            self.tree.focus(first)

    def choose(self, event=None):
        sel=self.tree.selection()
        if not sel: return "break"
        numero=self.tree.item(sel[0],"values")[0]
        self.on_select(numero)
        self.destroy()
        return "break"


class OCItemEditorDialog(tk.Toplevel):
    def __init__(self, parent, db, tipo_orden, defaults, item=None, on_save=None):
        super().__init__(parent)
        self.db=db; self.tipo_orden=normalize_text(tipo_orden or "NACIONAL"); self.defaults=defaults or {}; self.item=item or {}; self.on_save=on_save
        self.title("Agregar / modificar item de OC")
        self.geometry("720x420+230+130")
        bind_as_child_window(self,parent)
        self.codigo=tk.StringVar(value=self.item.get("codigo", ""))
        self.descripcion=tk.StringVar(value=self.item.get("descripcion", ""))
        self.um=tk.StringVar(value=self.item.get("um", "UND"))
        self.cantidad=tk.StringVar(value=str(self.item.get("cantidad", 1)))
        self.precio=tk.StringVar(value=str(self.item.get("precio", 0)))
        self.cc=tk.StringVar(value=self.item.get("centro_costo_codigo") or self.defaults.get("cc", "04"))
        self.cc_nombre=tk.StringVar(value=CENTROS_COSTO.get(normalize_text(self.cc.get()),""))
        self.ot=tk.StringVar(value=self.item.get("ot") or self.defaults.get("ot", ""))
        self.solicitante=tk.StringVar(value=self.item.get("solicitante") or self.defaults.get("solicitante", ""))
        self.build()

    def build(self):
        frm=ttk.LabelFrame(self,text="Datos del item",padding=12); frm.pack(fill="both",expand=True,padx=12,pady=12)
        rows=[("Código",self.codigo,0),("Descripción",self.descripcion,1),("Unidad",self.um,2),("Cantidad",self.cantidad,3),("Precio unitario",self.precio,4),("Centro costo",self.cc,5),("OT",self.ot,6),("Solicitante",self.solicitante,7)]
        for lab,var,r in rows:
            ttk.Label(frm,text=lab+":").grid(row=r,column=0,sticky="w",pady=5,padx=4)
            if lab=="Unidad":
                ttk.Combobox(frm,textvariable=var,values=UNIDADES_PERMITIDAS,state="readonly",width=18).grid(row=r,column=1,sticky="w",pady=5,padx=4)
            elif lab=="Descripción":
                state="normal" if self.tipo_orden=="SERVICIO" else "readonly"
                ttk.Entry(frm,textvariable=var,width=58,state=state).grid(row=r,column=1,columnspan=3,sticky="we",pady=5,padx=4)
            elif lab=="Centro costo":
                ent=ttk.Entry(frm,textvariable=var,width=18); ent.grid(row=r,column=1,sticky="w",pady=5,padx=4); ent.bind("<Shift-F1>", lambda e: self.open_cc())
                ttk.Entry(frm,textvariable=self.cc_nombre,state="readonly",width=32).grid(row=r,column=2,sticky="w",pady=5,padx=4)
                ttk.Button(frm,text="Buscar",command=self.open_cc).grid(row=r,column=3,sticky="w",padx=4)
            elif lab=="Solicitante":
                ent=ttk.Entry(frm,textvariable=var,width=34); ent.grid(row=r,column=1,sticky="w",pady=5,padx=4); ent.bind("<Shift-F1>", lambda e: self.open_solicitante())
                attach_solicitante_autocomplete(ent, self.db, self.solicitante)
                ttk.Button(frm,text="Buscar",command=self.open_solicitante).grid(row=r,column=2,sticky="w",padx=4)
            else:
                ent=ttk.Entry(frm,textvariable=var,width=22); ent.grid(row=r,column=1,sticky="w",pady=5,padx=4)
                if lab=="Código" and self.tipo_orden=="NACIONAL":
                    ent.bind("<Shift-F1>", lambda e: self.open_articulo())
                    attach_articulo_autocomplete(ent, self.db, self.codigo, self.set_articulo)
                    ttk.Button(frm,text="Buscar artículo",command=self.open_articulo).grid(row=r,column=2,sticky="w",padx=4)
                    ttk.Button(frm,text="Validar código",command=self.load_articulo).grid(row=r,column=3,sticky="w",padx=4)
        self.cc.trace_add("write", lambda *a: self.cc_nombre.set(CENTROS_COSTO.get(normalize_text(self.cc.get()),"")))
        btn=ttk.Frame(self,padding=10); btn.pack(fill="x")
        ttk.Button(btn,text="Guardar item",command=self.save).pack(side="right",padx=4)
        ttk.Button(btn,text="Cerrar",command=self.destroy).pack(side="right",padx=4)
        if self.tipo_orden=="NACIONAL" and self.codigo.get():
            self.load_articulo(silent=True)

    def open_articulo(self):
        ArticuloSelectDialog(self, self.db, self.set_articulo)

    def set_articulo(self,codigo):
        self.codigo.set(codigo); self.load_articulo()

    def load_articulo(self, silent=False):
        art=self.db.get_articulo(self.codigo.get())
        if not art:
            if not silent: messagebox.showwarning("Artículo","No se encontró el artículo.",parent=self)
            return
        self.codigo.set(art["codigo"]); self.descripcion.set(art["descripcion"]); self.um.set(art["unidad_medida"])

    def open_cc(self):
        CentroCostoSelectDialog(self, self.cc.set)

    def open_solicitante(self):
        SolicitanteSelectDialog(self, self.db, self.solicitante.set)

    def save(self):
        try:
            if self.tipo_orden=="NACIONAL":
                art=self.db.get_articulo(self.codigo.get())
                if not art: raise ValueError("Seleccione un artículo válido.")
                codigo=art["codigo"]; descripcion=art["descripcion"]; um=art["unidad_medida"]
            else:
                codigo=normalize_text(self.codigo.get())
                descripcion=normalize_text(self.descripcion.get())
                um=normalize_text(self.um.get()) or "UND"
            cantidad=float(str(self.cantidad.get() or 0).replace(",","."))
            precio=float(str(self.precio.get() or 0).replace(",","."))
            cc=normalize_text(self.cc.get())
            if not descripcion: raise ValueError("Ingrese descripción.")
            if um not in UNIDADES_PERMITIDAS: raise ValueError("Unidad de medida inválida.")
            if cantidad<=0: raise ValueError("La cantidad debe ser mayor a cero.")
            if precio<0: raise ValueError("El precio no puede ser negativo.")
            if cc not in CENTROS_COSTO: raise ValueError("Centro de costo inválido.")
            data={"codigo":codigo,"descripcion":descripcion,"um":um,"cantidad":cantidad,"precio":precio,"centro_costo_codigo":cc,"ot":self.ot.get(),"solicitante":self.solicitante.get()}
            if self.on_save:
                self.on_save(data)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Item",str(e),parent=self)




class ProviderListDialog(tk.Toplevel):
    def __init__(self, app, on_select=None):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.on_select = on_select
        self.title("Proveedores / Nuevo proveedor")
        self.geometry("1180x720+60+30")
        self.minsize(1050, 620)
        bind_as_child_window(self, self.master if hasattr(self, "master") else None)
        self.term = tk.StringVar()
        self.build()
        self.refresh()

    def build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Buscar RUC / razón social:").pack(side="left")
        e = ttk.Entry(top, textvariable=self.term, width=42)
        e.pack(side="left", padx=8)
        e.bind("<KeyRelease>", lambda e: self.refresh())
        ttk.Button(top, text="Filtrar", command=self.refresh).pack(side="left", padx=4)
        ttk.Button(top, text="Nuevo proveedor", command=self.clear_form).pack(side="left", padx=4)
        ttk.Button(top, text="Consultar RUC API", command=self.consultar_ruc_api).pack(side="left", padx=4)
        ttk.Button(top, text="Guardar proveedor", command=self.save).pack(side="left", padx=12)

        cols = ("ruc", "razon", "telefono", "banco", "cuenta", "moneda", "estado", "condicion", "activo")
        tree_box = ttk.Frame(self)
        tree_box.pack(fill="both", expand=True, padx=10, pady=8)
        tree_box.rowconfigure(0, weight=1); tree_box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_box, columns=cols, show="headings", height=9)
        specs = [
            ("ruc", "RUC", 110), ("razon", "Razón social", 310), ("telefono", "Teléfono", 95),
            ("banco", "Banco", 95), ("cuenta", "Cuenta", 145), ("moneda", "Moneda", 80),
            ("estado", "Estado SUNAT", 110), ("condicion", "Condición", 110), ("activo", "Activo", 65)
        ]
        for c, t, w in specs:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, minwidth=w, anchor="center" if c != "razon" else "w", stretch=False)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.load_selected)
        self.tree.bind("<Double-1>", self.choose)

        frm = ttk.LabelFrame(self, text="Datos de proveedor", padding=10)
        frm.pack(fill="x", padx=10, pady=6)
        self.ruc = tk.StringVar(); self.razon = tk.StringVar(); self.direccion = tk.StringVar(); self.telefono = tk.StringVar(); self.email = tk.StringVar()
        self.banco = tk.StringVar(value="01 BCP"); self.banco_otro = tk.StringVar(); self.cuenta = tk.StringVar(); self.cci = tk.StringVar(); self.contacto = tk.StringVar(); self.moneda = tk.StringVar(value="SOLES"); self.activo = tk.BooleanVar(value=True)
        self.estado_sunat = tk.StringVar(); self.condicion_sunat = tk.StringVar(); self.ubigeo = tk.StringVar(); self.fuente_ruc = tk.StringVar(); self.fecha_val = tk.StringVar()
        labels = [
            ("RUC", self.ruc, 0, 0), ("Razón social", self.razon, 0, 2),
            ("Dirección", self.direccion, 1, 0), ("Teléfono", self.telefono, 1, 2),
            ("Email", self.email, 2, 0), ("Banco", self.banco, 2, 2),
            ("Cuenta", self.cuenta, 3, 0), ("CCI", self.cci, 3, 2),
            ("Contacto", self.contacto, 4, 0), ("Moneda", self.moneda, 4, 2),
            ("Banco otros", self.banco_otro, 5, 0), ("Estado SUNAT", self.estado_sunat, 5, 2),
            ("Condición", self.condicion_sunat, 6, 0), ("Ubigeo", self.ubigeo, 6, 2),
            ("Fuente/Fecha", self.fuente_ruc, 7, 0),
        ]
        for lab, var, r, c in labels:
            ttk.Label(frm, text=lab+":").grid(row=r, column=c, sticky="w", padx=4, pady=4)
            if lab == "Banco":
                cb = ttk.Combobox(frm, textvariable=var, values=BANCOS, state="readonly", width=24)
                cb.grid(row=r, column=c+1, sticky="w", padx=4, pady=4)
                cb.bind("<<ComboboxSelected>>", lambda e: self.toggle_banco_otro())
            elif lab == "Moneda":
                ttk.Combobox(frm, textvariable=var, values=MONEDAS, state="readonly", width=24).grid(row=r, column=c+1, sticky="w", padx=4, pady=4)
            elif lab in ("Estado SUNAT", "Condición", "Ubigeo", "Fuente/Fecha"):
                ent = ttk.Entry(frm, textvariable=var, width=28 if c else 36, state="readonly")
                ent.grid(row=r, column=c+1, sticky="w", padx=4, pady=4)
            else:
                ent = ttk.Entry(frm, textvariable=var, width=28 if c else 36)
                ent.grid(row=r, column=c+1, sticky="w", padx=4, pady=4)
                if lab == "RUC":
                    self.ruc_entry = ent
                    ent.bind("<Return>", lambda e: self.consultar_ruc_api())
        ttk.Checkbutton(frm, text="Activo", variable=self.activo).grid(row=8, column=1, sticky="w", padx=4, pady=4)
        ttk.Label(frm, text="Nota: SUNAT/API no trae datos bancarios; cuenta, CCI y banco se registran manualmente.", foreground="#555555").grid(row=8, column=2, columnspan=3, sticky="w")

        # Botones dentro del bloque de datos para que siempre sean visibles aun con pantallas pequeñas.
        action_bar = ttk.Frame(frm)
        action_bar.grid(row=9, column=0, columnspan=5, sticky="ew", padx=4, pady=(10, 2))
        ttk.Button(action_bar, text="Guardar proveedor", command=self.save).pack(side="left", padx=5)
        ttk.Button(action_bar, text="Seleccionar", command=self.choose).pack(side="left", padx=5)
        ttk.Button(action_bar, text="Limpiar / Nuevo", command=self.clear_form).pack(side="left", padx=5)
        ttk.Button(action_bar, text="Cerrar", command=self.destroy).pack(side="left", padx=5)

        self.bind("<Control-s>", lambda e: self.save())
        self.bind("<Control-S>", lambda e: self.save())

    def toggle_banco_otro(self):
        # El campo está visible permanentemente para evitar confusión; solo se usa si el banco es 05 OTROS.
        if not self.banco.get().startswith("05"):
            self.banco_otro.set("")

    def clear_form(self):
        self.ruc.set(""); self.razon.set(""); self.direccion.set(""); self.telefono.set(""); self.email.set("")
        self.banco.set("01 BCP"); self.banco_otro.set(""); self.cuenta.set(""); self.cci.set(""); self.contacto.set(""); self.moneda.set("SOLES"); self.activo.set(True)
        self.estado_sunat.set(""); self.condicion_sunat.set(""); self.ubigeo.set(""); self.fuente_ruc.set(""); self.fecha_val.set("")
        self.toggle_banco_otro()
        self.tree.selection_remove(self.tree.selection())

    def refresh(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        for r in self.db.list_proveedores(self.term.get(), active_only=False):
            self.tree.insert("", "end", values=(
                r["ruc"], r["razon_social"], r["telefono"], r["banco"], r["numero_cuenta"], r["moneda_preferida"],
                r["estado_contribuyente"], r["condicion_domicilio"], "SI" if r["activo"] else "NO"
            ))

    def load_selected(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        ruc = self.tree.item(sel[0], "values")[0]
        r = self.db.get_proveedor_by_ruc(ruc, active_only=False)
        if not r: return
        self.ruc.set(r["ruc"]); self.razon.set(r["razon_social"]); self.direccion.set(r["direccion"]); self.telefono.set(r["telefono"]); self.email.set(r["email"])
        banco_guardado = r["banco"] or "01 BCP"
        if str(banco_guardado).startswith("05"):
            self.banco.set("05 OTROS")
            self.banco_otro.set(str(banco_guardado).replace("05 OTROS", "").replace("-", "").strip())
        else:
            cod = banco_codigo_guardado(banco_guardado)
            self.banco.set(f"{cod} {BANCOS_MAP[cod]}" if cod else banco_guardado)
            self.banco_otro.set("")
        self.cuenta.set(r["numero_cuenta"]); self.cci.set(r["cci"]); self.contacto.set(r["contacto"]); self.moneda.set(r["moneda_preferida"] or "SOLES"); self.activo.set(bool(r["activo"]))
        self.toggle_banco_otro()
        self.estado_sunat.set(r["estado_contribuyente"]); self.condicion_sunat.set(r["condicion_domicilio"]); self.ubigeo.set(r["ubigeo"])
        fuente = " ".join(x for x in [r["fuente_ruc"], r["fecha_validacion_ruc"]] if x)
        self.fuente_ruc.set(fuente)

    def consultar_ruc_api(self):
        if not self.app.require_action("modificar_proveedor", self): return
        ruc = only_digits(self.ruc.get() or self.term.get())
        if not ruc:
            ruc = simpledialog.askstring("Consultar RUC", "Ingrese RUC de 11 dígitos:", parent=self)
        if not ruc: return
        try:
            info = self.db.consultar_ruc_api(ruc)
            actual = self.db.get_proveedor_by_ruc(info["ruc"], active_only=False)
            self.ruc.set(info["ruc"]); self.razon.set(info["razon_social"]); self.direccion.set(info["direccion"])
            if actual:
                self.telefono.set(actual["telefono"] or self.telefono.get())
                self.email.set(actual["email"] or self.email.get())
                self.banco.set(actual["banco"] or self.banco.get() or "01 BCP")
                self.cuenta.set(actual["numero_cuenta"] or self.cuenta.get())
                self.cci.set(actual["cci"] or self.cci.get())
                self.contacto.set(actual["contacto"] or self.contacto.get())
                self.moneda.set(actual["moneda_preferida"] or self.moneda.get() or "SOLES")
                self.activo.set(bool(actual["activo"]))
            self.estado_sunat.set(info["estado_contribuyente"]); self.condicion_sunat.set(info["condicion_domicilio"]); self.ubigeo.set(info["ubigeo"]); self.fuente_ruc.set(f"{info['fuente_ruc']} {info['fecha_validacion_ruc']}")
            msg = f"Datos cargados desde API en pantalla. Complete teléfono/cuenta/CCI/contacto y pulse Guardar proveedor.\nEstado: {info['estado_contribuyente']}\nCondición: {info['condicion_domicilio']}"
            if info["estado_contribuyente"] and info["estado_contribuyente"] != "ACTIVO":
                msg += "\n\nAdvertencia: el contribuyente no figura ACTIVO."
            if info["condicion_domicilio"] and info["condicion_domicilio"] != "HABIDO":
                msg += "\nAdvertencia: la condición no figura HABIDO."
            messagebox.showinfo("Consulta RUC", msg, parent=self)
        except Exception as e:
            messagebox.showerror("Consulta RUC", str(e), parent=self)

    def save(self):
        if not self.app.require_action("modificar_proveedor", self): return
        try:
            fuente_txt = (self.fuente_ruc.get() or "").strip()
            parts = fuente_txt.split()
            fuente = parts[0] if parts else "MANUAL"
            fecha_val = " ".join(parts[1:]) if len(parts) > 1 else (now_text() if fuente != "MANUAL" else "")
            self.db.upsert_proveedor(
                self.ruc.get(), self.razon.get(), self.direccion.get(), self.telefono.get(), self.email.get(),
                normalizar_banco(self.banco.get(), self.banco_otro.get()), self.cuenta.get(), self.cci.get(), self.contacto.get(), self.moneda.get(), self.activo.get(),
                estado_contribuyente=self.estado_sunat.get(), condicion_domicilio=self.condicion_sunat.get(), ubigeo=self.ubigeo.get(),
                fuente_ruc=fuente,
                fecha_validacion_ruc=fecha_val,
                usuario=self.app.usuario_actual
            )
            self.db.audit("proveedores", self.ruc.get(), "GUARDAR", self.app.usuario_actual, self.razon.get())
            messagebox.showinfo("Proveedor", "Proveedor guardado correctamente.", parent=self)
            self.refresh()
        except Exception as e:
            messagebox.showerror("Proveedor", str(e), parent=self)

    def choose(self, event=None):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Seleccione", "Seleccione un proveedor.", parent=self); return
        ruc = self.tree.item(sel[0], "values")[0]
        if self.on_select:
            self.on_select(ruc)
            self.destroy()


class OCDetailDialog(tk.Toplevel):
    def __init__(self, parent, db, numero):
        super().__init__(parent)
        self.db = db
        self.numero = numero
        self.title(f"Detalle Orden de Compra {numero}")
        self.geometry("1260x780+45+25")
        self.minsize(1120, 700)
        bind_as_child_window(self, parent)
        self.build()

    def build(self):
        oc, det = self.db.get_orden_compra(self.numero)
        if not oc:
            ttk.Label(self, text="No se encontró la OC").pack(padx=20, pady=20)
            return

        # Refuerzo v33: consultar el detalle directamente y renderizarlo en una zona fija.
        # En algunas ventanas el total de la OC se veía, pero el Treeview de items quedaba sin filas visibles.
        try:
            cur = self.db.conn.cursor()
            cur.execute("SELECT * FROM orden_compra_detalle WHERE oc_id=? ORDER BY item, id", (oc["id"],))
            det = cur.fetchall()
        except Exception:
            det = det or []

        estado_pago, pagado, saldo = self.db._sync_oc_payment_status(oc["id"], commit=True)

        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=3)
        main.rowconfigure(5, weight=2)

        titulo = f"{oc['numero']} - {oc['estado']} - Avance operativo {oc['avance']}% | Pago: {estado_pago} | Saldo: {saldo:.2f}"
        ttk.Label(main, text=titulo, font=("Arial", 14, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))

        meta = ttk.LabelFrame(main, text="Cabecera", padding=8)
        meta.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        meta.columnconfigure(1, weight=1)
        meta.columnconfigure(3, weight=1)

        data = [
            ("Fecha", fecha_text_from_iso(oc["fecha_orden"])),
            ("Proveedor", f"{oc['ruc']} - {oc['razon_social']}"),
            ("Tipo", oc["tipo_orden"]),
            ("Moneda", oc["moneda"]),
            ("TC", f"{float(oc['tipo_cambio_venta'] or 0):.4f} {oc['tipo_cambio_fuente'] or ''}"),
            ("Requerimiento", oc["requerimiento"] or ""),
            ("Cotización", oc["cotizacion"] or ""),
            ("Solicitante", oc["solicitante"] or ""),
            ("OT", oc["ot"] or ""),
            ("Centro costo", f"{oc['centro_costo_codigo']} - {oc['centro_costo_nombre']}"),
            ("Almacén", f"{oc['almacen_codigo']} - {oc['almacen_nombre']}"),
            ("Forma pago", oc["forma_pago"] or ""),
            ("Vencimiento pago", fecha_text_from_iso(oc["fecha_vencimiento_pago"]) if oc["fecha_vencimiento_pago"] else ""),
            ("Pagado", f"{pagado:.2f}"),
            ("Saldo", f"{saldo:.2f}"),
            ("Observación", oc["observacion"] or ""),
        ]
        for i, (k, v) in enumerate(data):
            r, c = divmod(i, 2)
            ttk.Label(meta, text=k + ":", font=("Arial", 9, "bold")).grid(row=r, column=c*2, sticky="w", padx=4, pady=2)
            ttk.Label(meta, text=str(v), wraplength=430).grid(row=r, column=c*2 + 1, sticky="w", padx=4, pady=2)

        # Área de items: etiqueta visible + grilla con encabezados siempre visibles.
        items_box = ttk.LabelFrame(main, text=f"Items de la orden ({len(det)} registro(s))", padding=6)
        items_box.grid(row=2, column=0, sticky="nsew", pady=(0, 4))
        items_box.rowconfigure(0, weight=1)
        items_box.columnconfigure(0, weight=1)

        cols = ("item", "codigo", "descripcion", "um", "sol", "aprob", "ing", "pend", "punit", "subtotal", "igv", "total")
        tree = ttk.Treeview(items_box, columns=cols, show="headings", height=11, selectmode="browse")
        tree["displaycolumns"] = cols
        specs = [
            ("item", "Item", 45, "center"),
            ("codigo", "Código", 100, "center"),
            ("descripcion", "Descripción", 420, "w"),
            ("um", "UM", 55, "center"),
            ("sol", "Solicitada", 85, "e"),
            ("aprob", "Aprobada", 85, "e"),
            ("ing", "Ingresada", 85, "e"),
            ("pend", "Pendiente", 85, "e"),
            ("punit", "P.Unit s/IGV", 110, "e"),
            ("subtotal", "Subtotal", 100, "e"),
            ("igv", "IGV", 90, "e"),
            ("total", "Total", 100, "e"),
        ]
        for c, t, w, anchor in specs:
            tree.heading(c, text=t)
            tree.column(c, width=w, minwidth=40, anchor=anchor, stretch=False)
        vsb = ttk.Scrollbar(items_box, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(items_box, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        if not det:
            tree.insert("", "end", values=("", "", "SIN ITEMS REGISTRADOS EN ESTA OC. Revise si la orden fue generada sin detalle.", "", "", "", "", "", "", "", "", ""))
        else:
            for d in det:
                pend = float(d["cantidad_aprobada"] or 0) - float(d["cantidad_ingresada"] or 0)
                tree.insert("", "end", values=(
                    d["item"], d["codigo"], d["descripcion"], d["unidad_medida"],
                    f"{float(d['cantidad_solicitada'] or 0):g}",
                    f"{float(d['cantidad_aprobada'] or 0):g}",
                    f"{float(d['cantidad_ingresada'] or 0):g}",
                    f"{pend:g}",
                    f"{float(d['precio_unitario_sin_igv'] or 0):.2f}",
                    f"{float(d['subtotal'] or 0):.2f}",
                    f"{float(d['igv'] or 0):.2f}",
                    f"{float(d['total'] or 0):.2f}",
                ))

        totals = ttk.Frame(main)
        totals.grid(row=3, column=0, sticky="ew", pady=(6, 6))
        ttk.Label(
            totals,
            text=f"Subtotal: {float(oc['subtotal'] or 0):.2f}   IGV 18%: {float(oc['igv'] or 0):.2f}   Total: {float(oc['total'] or 0):.2f}",
            font=("Arial", 11, "bold")
        ).pack(side="right")

        ttk.Label(main, text="Historial de estados:", font=("Arial", 10, "bold")).grid(row=4, column=0, sticky="w", pady=(4, 2))
        hist_box = ttk.Frame(main)
        hist_box.grid(row=5, column=0, sticky="nsew")
        hist_box.rowconfigure(0, weight=1)
        hist_box.columnconfigure(0, weight=1)

        htxt = tk.Text(hist_box, height=8, wrap="word")
        hscroll = ttk.Scrollbar(hist_box, orient="vertical", command=htxt.yview)
        htxt.configure(yscrollcommand=hscroll.set)
        htxt.grid(row=0, column=0, sticky="nsew")
        hscroll.grid(row=0, column=1, sticky="ns")
        hist = self.db.oc_historial(oc["id"])
        if not hist:
            htxt.insert("end", "Sin historial registrado.\n")
        for h in hist:
            htxt.insert("end", f"{h['fecha_hora']} | {h['usuario']} | {h['estado_anterior']} -> {h['estado_nuevo']} | {h['detalle']}\n")
        htxt.configure(state="disabled")

        ttk.Button(main, text="Cerrar", command=self.destroy).grid(row=6, column=0, sticky="e", pady=(8, 0))


class OCModificarDialog(tk.Toplevel):
    """Modifica solo cabecera editable de una OC pendiente/en cotización.
    No toca solicitante, centro de costo, artículos ni cantidades heredadas del requerimiento.
    """
    def __init__(self, app, numero, on_saved=None):
        super().__init__(app.root)
        self.app = app; self.db = app.db; self.numero = numero; self.on_saved = on_saved
        self.title(f"Modificar OC pendiente - {numero}")
        self.geometry("820x430+210+110")
        bind_as_child_window(self, app.root)
        self.ruc=tk.StringVar(); self.razon=tk.StringVar(); self.cotizacion=tk.StringVar(); self.observacion=tk.StringVar()
        self.forma=tk.StringVar(value="CONTADO"); self.fecha_entrega=tk.StringVar(); self.lugar_entrega=tk.StringVar()
        self.info=tk.StringVar()
        self.build(); self.load()

    def build(self):
        frm=ttk.Frame(self,padding=18); frm.pack(fill="both",expand=True)
        ttk.Label(frm,textvariable=self.info,foreground="#444",wraplength=760).grid(row=0,column=0,columnspan=4,sticky="w",pady=(0,10))
        rows=[("RUC proveedor",self.ruc), ("Razón social",self.razon), ("Cotización",self.cotizacion), ("Observación",self.observacion), ("Fecha entrega",self.fecha_entrega), ("Lugar entrega",self.lugar_entrega)]
        for i,(lab,var) in enumerate(rows, start=1):
            ttk.Label(frm,text=lab+":").grid(row=i,column=0,sticky="w",pady=5)
            ent=ttk.Entry(frm,textvariable=var,width=48,state="readonly" if lab=="Razón social" else "normal")
            ent.grid(row=i,column=1,columnspan=2,sticky="w",pady=5)
            if lab=="RUC proveedor":
                ent.bind("<Shift-F1>", lambda e: self.buscar_proveedor())
                attach_proveedor_autocomplete(ent, self.db, self.ruc, self.set_proveedor)
        ttk.Button(frm,text="Buscar proveedor",command=self.buscar_proveedor).grid(row=1,column=3,padx=8)
        ttk.Label(frm,text="Forma pago:").grid(row=7,column=0,sticky="w",pady=5)
        ttk.Combobox(frm,textvariable=self.forma,values=FORMAS_PAGO,state="readonly",width=30).grid(row=7,column=1,sticky="w",pady=5)
        ttk.Label(frm,text="Restricción: solo mientras la OC esté pendiente/en cotización. No modifica artículos ni datos heredados del requerimiento.",foreground="#666").grid(row=8,column=0,columnspan=4,sticky="w",pady=8)
        ttk.Button(frm,text="Guardar cambios",command=self.save).grid(row=9,column=2,sticky="e",pady=12)
        ttk.Button(frm,text="Cerrar",command=self.destroy).grid(row=9,column=3,sticky="w",pady=12,padx=8)

    def load(self):
        oc,det=self.db.get_orden_compra(self.numero)
        if not oc:
            self.info.set("OC no encontrada."); return
        self.info.set(f"{oc['numero']} | Estado: {oc['estado']} | Requerimiento: {oc['requerimiento']} | Total: {oc['total']:.2f}")
        self.ruc.set(oc['ruc']); self.razon.set(oc['razon_social']); self.cotizacion.set(oc['cotizacion']); self.observacion.set(oc['observacion'])
        self.forma.set(oc['forma_pago'] or 'CONTADO'); self.fecha_entrega.set(oc['fecha_entrega'] or ''); self.lugar_entrega.set(oc['lugar_entrega'] or '')

    def buscar_proveedor(self):
        ProviderListDialog(self.app, on_select=self.set_proveedor)

    def set_proveedor(self, ruc):
        r=self.db.get_proveedor_by_ruc(ruc, active_only=True)
        if not r: return
        self.ruc.set(r['ruc']); self.razon.set(r['razon_social'])

    def save(self):
        if not self.app.require_action("crear_oc", self): return
        try:
            self.db.update_orden_compra_pendiente(self.numero, self.ruc.get(), self.cotizacion.get(), self.observacion.get(), self.forma.get(), self.fecha_entrega.get(), self.lugar_entrega.get(), self.app.usuario_actual)
            messagebox.showinfo("OC", "OC pendiente actualizada correctamente.", parent=self)
            if self.on_saved: self.on_saved()
            self.destroy()
        except Exception as e:
            messagebox.showerror("No se pudo modificar", str(e), parent=self)


class OCManualDialog(tk.Toplevel):
    def __init__(self, app, on_saved=None):
        super().__init__(app.root)
        self.app=app; self.db=app.db; self.on_saved=on_saved; self.detalles=[]; self.selected_index=None
        self.title("Nueva Orden de Compra")
        self.geometry("1280x760+35+20")
        self.minsize(1180, 680)
        bind_as_child_window(self, self.master if hasattr(self,"master") else None)
        self.field_entries={}
        self.build()

    def build(self):
        top=ttk.LabelFrame(self,text="Cabecera",padding=10); top.pack(fill="x",padx=10,pady=8)
        self.fecha=tk.StringVar(value=today_text())
        self.almacen=tk.StringVar(value=getattr(self.app,"almacen_actual","AQP"))
        self.ruc=tk.StringVar(); self.razon=tk.StringVar(); self.direccion=tk.StringVar(); self.telefono=tk.StringVar()
        self.tipo=tk.StringVar(value="NACIONAL"); self.requerimiento=tk.StringVar(); self.cotizacion=tk.StringVar()
        self.solicitante=tk.StringVar(); self.ot=tk.StringVar(); self.cc=tk.StringVar(value="04")
        self.cc_nombre=tk.StringVar(value=CENTROS_COSTO.get("04","")); self.obs=tk.StringVar()
        self.moneda=tk.StringVar(value="SOLES"); self.igv_mode=tk.StringVar(value="SIN IGV")
        self.forma_pago=tk.StringVar(value="CONTADO"); self.fecha_entrega=tk.StringVar(value=today_text())
        self.lugar_entrega=tk.StringVar(value=ALMACENES.get(self.almacen.get(),""))
        self.tipo_cambio=tk.StringVar(value="1.0000"); self.tc_fuente=tk.StringVar(value="SOLES"); self.tc_fecha=tk.StringVar(value="")
        fields=[
            ("Fecha",self.fecha,0,0),("Almacén",self.almacen,0,2),("RUC",self.ruc,0,4),
            ("Razón social",self.razon,1,0),("Dirección",self.direccion,1,2),("Teléfono",self.telefono,1,4),
            ("Tipo",self.tipo,2,0),("Requerimiento",self.requerimiento,2,2),("Cot/Cotiz.",self.cotizacion,2,4),
            ("Solicitante",self.solicitante,3,0),("OT",self.ot,3,2),("Centro costo",self.cc,3,4),
            ("Observaciones",self.obs,4,0),("Moneda",self.moneda,5,0),("Precio",self.igv_mode,5,2),
            ("Forma pago",self.forma_pago,5,4),("Entrega",self.fecha_entrega,6,0),("Lugar entrega",self.lugar_entrega,6,2),
            ("Tipo cambio",self.tipo_cambio,7,0),("Fuente TC",self.tc_fuente,7,2)
        ]
        for lab,var,r,c in fields:
            ttk.Label(top,text=lab+":").grid(row=r,column=c,sticky="w",padx=4,pady=3)
            if lab=="Almacén":
                widget=ttk.Combobox(top,textvariable=var,values=list(ALMACENES.keys()),state="readonly",width=18)
            elif lab=="Tipo":
                widget=ttk.Combobox(top,textvariable=var,values=TIPOS_ORDEN_COMPRA,state="readonly",width=18)
            elif lab=="Moneda":
                widget=ttk.Combobox(top,textvariable=var,values=MONEDAS,state="readonly",width=18)
            elif lab=="Precio":
                widget=ttk.Combobox(top,textvariable=var,values=["SIN IGV","CON IGV"],state="readonly",width=18)
            elif lab=="Forma pago":
                widget=ttk.Combobox(top,textvariable=var,values=FORMAS_PAGO,state="readonly",width=22)
            elif lab in ("Tipo cambio", "Fuente TC"):
                widget=ttk.Entry(top,textvariable=var,width=22)
            elif lab=="Observaciones":
                widget=ttk.Entry(top,textvariable=var,width=52); widget.grid(row=r,column=c+1,columnspan=5,sticky="we",padx=4,pady=3); self.field_entries[lab]=widget; continue
            elif lab=="Lugar entrega":
                widget=ttk.Entry(top,textvariable=var,width=34); widget.grid(row=r,column=c+1,columnspan=3,sticky="we",padx=4,pady=3); self.field_entries[lab]=widget; continue
            else:
                widget=ttk.Entry(top,textvariable=var,width=22)
            widget.grid(row=r,column=c+1,sticky="w",padx=4,pady=3)
            self.field_entries[lab]=widget
        self.field_entries["RUC"].bind("<Return>", lambda e: (self.load_provider(), "break")[1])
        self.field_entries["RUC"].bind("<Shift-F1>", lambda e: self.open_provider())
        attach_proveedor_autocomplete(self.field_entries["RUC"], self.db, self.ruc, self.set_provider)
        self.field_entries["Requerimiento"].bind("<Shift-F1>", lambda e: self.open_req())
        self.field_entries["Solicitante"].bind("<Shift-F1>", lambda e: self.open_solicitante())
        attach_solicitante_autocomplete(self.field_entries["Solicitante"], self.db, self.solicitante)
        self.field_entries["Centro costo"].bind("<Shift-F1>", lambda e: self.open_cc())
        ttk.Button(top,text="Buscar proveedor",command=self.open_provider).grid(row=0,column=6,padx=4)
        ttk.Button(top,text="Cargar RUC",command=self.load_provider).grid(row=0,column=7,padx=4)
        ttk.Button(top,text="Nuevo proveedor",command=lambda: ProviderListDialog(self.app)).grid(row=1,column=6,padx=4)
        ttk.Button(top,text="Buscar Req",command=self.open_req).grid(row=2,column=6,padx=4)
        ttk.Button(top,text="Buscar Sol.",command=self.open_solicitante).grid(row=3,column=6,padx=4)
        ttk.Button(top,text="Buscar CC",command=self.open_cc).grid(row=3,column=7,padx=4)
        ttk.Button(top,text="Cargar TC",command=self.load_tipo_cambio).grid(row=7,column=4,padx=4)
        ttk.Button(top,text="TC manual",command=lambda: TipoCambioDialog(self.app, fecha_inicial=self.fecha.get(), on_select=self.set_tipo_cambio)).grid(row=7,column=5,padx=4)
        ttk.Label(top,text="Tip: SHIFT + F1 busca requerimiento, solicitante, centro de costo y artículos.",foreground="#555").grid(row=8,column=0,columnspan=8,sticky="w",pady=(4,0))
        self.cc.trace_add("write", lambda *a: self.cc_nombre.set(CENTROS_COSTO.get(normalize_text(self.cc.get()),"")))
        self.almacen.trace_add("write", lambda *a: self.lugar_entrega.set(ALMACENES.get(normalize_text(self.almacen.get()), self.lugar_entrega.get())))
        self.moneda.trace_add("write", lambda *a: self.on_moneda_changed())
        self.fecha.trace_add("write", lambda *a: self.on_moneda_changed())
        ttk.Entry(top,textvariable=self.cc_nombre,state="readonly",width=18).grid(row=3,column=8,columnspan=2,sticky="w",padx=4)

        body=ttk.LabelFrame(self,text="Items",padding=10); body.pack(fill="both",expand=True,padx=10,pady=6)
        cols=("item","codigo","descripcion","um","cantidad","precio","cc","ot","sol")
        tree_box=ttk.Frame(body); tree_box.pack(fill="both",expand=True)
        tree_box.rowconfigure(0,weight=1); tree_box.columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(tree_box,columns=cols,show="headings",height=12)
        specs=[("item","Item",50),("codigo","Código",100),("descripcion","Descripción",390),("um","UM",60),("cantidad","Cantidad",85),("precio","Precio",85),("cc","CC",65),("ot","OT",90),("sol","Solicitante",135)]
        for c,t,w in specs:
            self.tree.heading(c,text=t); self.tree.column(c,width=w,minwidth=w,anchor="center" if c!="descripcion" else "w",stretch=False)
        vsb=ttk.Scrollbar(tree_box,orient="vertical",command=self.tree.yview); hsb=ttk.Scrollbar(tree_box,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        self.tree.bind("<<TreeviewSelect>>",self.select_item)
        self.tree.bind("<Double-1>",lambda e: self.modify_item())
        btn=ttk.Frame(self,padding=10); btn.pack(fill="x")
        ttk.Button(btn,text="Agregar item",command=self.add_item).pack(side="left",padx=4)
        ttk.Button(btn,text="Modificar item",command=self.modify_item).pack(side="left",padx=4)
        ttk.Button(btn,text="Eliminar item",command=self.delete_item).pack(side="left",padx=4)
        ttk.Button(btn,text="Guardar OC",command=self.save).pack(side="right",padx=4)
        ttk.Button(btn,text="Cerrar",command=self.destroy).pack(side="right",padx=4)

    def open_provider(self):
        ProviderListDialog(self.app, self.set_provider)
    def set_provider(self,ruc):
        self.ruc.set(ruc); self.load_provider()
    def load_provider(self):
        r=self.db.get_proveedor_by_ruc(self.ruc.get(),active_only=True)
        if not r:
            if messagebox.askyesno("Proveedor","No se encontró proveedor activo. ¿Desea abrir Nuevo proveedor / consultar RUC API?",parent=self):
                ProviderListDialog(self.app, self.set_provider)
            return
        self.ruc.set(r["ruc"]); self.razon.set(r["razon_social"]); self.direccion.set(r["direccion"]); self.telefono.set(r["telefono"]); self.moneda.set(r["moneda_preferida"] or "SOLES")
        if r["estado_contribuyente"] and r["estado_contribuyente"] != "ACTIVO":
            messagebox.showwarning("Proveedor", f"Advertencia: estado SUNAT {r['estado_contribuyente']}.", parent=self)
        if r["condicion_domicilio"] and r["condicion_domicilio"] != "HABIDO":
            messagebox.showwarning("Proveedor", f"Advertencia: condición SUNAT {r['condicion_domicilio']}.", parent=self)
        self.on_moneda_changed()

    def open_req(self):
        RequerimientoListDialog(self,self.db,self.set_requerimiento)
    def set_requerimiento(self, req_num):
        req, det = self.db.find_requerimiento(req_num)
        if not req:
            messagebox.showwarning("Requerimiento","No se encontró el requerimiento.",parent=self); return
        self.requerimiento.set(req["requerimiento"])
        self.solicitante.set(req["solicitante"])
        self.cc.set(req["centro_costo_codigo"])
        self.ot.set((req["ot"] or "").replace("OT-", ""))
        self.almacen.set(req["almacen_codigo"])
        try:
            self.tipo.set(req["tipo_orden"] if "tipo_orden" in req.keys() else "NACIONAL")
        except Exception:
            self.tipo.set("NACIONAL")
        self.obs.set(f"OC GENERADA DESDE {req['requerimiento']}")

    def open_solicitante(self):
        SolicitanteSelectDialog(self, self.db, self.solicitante.set)
    def open_cc(self):
        CentroCostoSelectDialog(self, self.cc.set)

    def on_moneda_changed(self):
        if normalize_text(self.moneda.get()) == "SOLES":
            self.tipo_cambio.set("1.0000"); self.tc_fuente.set("SOLES"); self.tc_fecha.set("")
        else:
            if self.tipo_cambio.get() in ("", "1.0000"):
                self.tc_fuente.set("PENDIENTE")

    def set_tipo_cambio(self, row):
        if not row: return
        self.tipo_cambio.set(f"{float(row['venta']):.4f}")
        self.tc_fuente.set(row["fuente"])
        self.tc_fecha.set(row["fecha"])

    def load_tipo_cambio(self):
        if normalize_text(self.moneda.get()) == "SOLES":
            self.tipo_cambio.set("1.0000"); self.tc_fuente.set("SOLES"); self.tc_fecha.set(""); return
        try:
            fecha_iso = fecha_iso_consulta_from_text(self.fecha.get())
            info = self.db.obtener_tipo_cambio_para_oc(fecha_iso, "DOLARES", self.app.usuario_actual, permitir_api=True)
            self.tipo_cambio.set(f"{info['venta']:.4f}"); self.tc_fuente.set(info["fuente"]); self.tc_fecha.set(info["fecha"])
            messagebox.showinfo("Tipo de cambio", f"TC cargado: venta {info['venta']:.4f}\nFuente: {info['fuente']}", parent=self)
        except Exception as e:
            messagebox.showerror("Tipo de cambio", str(e), parent=self)

    def select_item(self,event=None):
        sel=self.tree.selection(); self.selected_index=self.tree.index(sel[0]) if sel else None
    def refresh_items(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        for i,d in enumerate(self.detalles,1):
            d["item"]=i
            self.tree.insert("","end",values=(i,d.get("codigo",""),d.get("descripcion",""),d.get("um","UND"),f"{d.get('cantidad',0):g}",f"{d.get('precio',0):.2f}",d.get("centro_costo_codigo",""),d.get("ot",""),d.get("solicitante","")))
    def _item_defaults(self):
        return {"cc": self.cc.get(), "ot": self.ot.get(), "solicitante": self.solicitante.get()}
    def add_item(self):
        OCItemEditorDialog(self, self.db, self.tipo.get(), self._item_defaults(), on_save=self._append_item)
    def _append_item(self, data):
        self.detalles.append(data); self.refresh_items()
    def modify_item(self):
        if self.selected_index is None:
            messagebox.showinfo("Seleccione","Seleccione un item.",parent=self); return
        OCItemEditorDialog(self, self.db, self.tipo.get(), self._item_defaults(), item=self.detalles[self.selected_index], on_save=self._replace_item)
    def _replace_item(self, data):
        if self.selected_index is not None:
            self.detalles[self.selected_index]=data; self.refresh_items()
    def delete_item(self):
        if self.selected_index is None: return
        self.detalles.pop(self.selected_index); self.selected_index=None; self.refresh_items()
    def save(self):
        if not self.app.require_action("crear_oc", self): return
        try:
            n=self.db.create_orden_compra(self.fecha.get(),self.almacen.get(),self.ruc.get(),self.tipo.get(),self.requerimiento.get(),self.cotizacion.get(),self.solicitante.get(),self.ot.get(),self.cc.get(),self.obs.get(),self.moneda.get(),self.igv_mode.get()=="CON IGV",self.detalles,self.app.usuario_actual,self.forma_pago.get(),self.fecha_entrega.get(),self.lugar_entrega.get(), self.tipo_cambio.get() if normalize_text(self.moneda.get())=="DOLARES" else None, self.tc_fuente.get(), self.tc_fecha.get())
            messagebox.showinfo("OC generada",f"Orden generada: {n}",parent=self)
            if self.on_saved: self.on_saved()
            OCPrintPreviewDialog(self, self.db, n)
            self.destroy()
        except Exception as e: messagebox.showerror("No se pudo guardar",str(e),parent=self)


class OCListDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.app=app; self.db=app.db
        self.title("Registro / Consulta de Órdenes de Compra")
        self.geometry("1180x620+60+60")
        bind_as_child_window(self,self.master if hasattr(self,"master") else None)
        self.term=tk.StringVar(); self.estado=tk.StringVar(value="TODOS"); self.tipo=tk.StringVar(value="TODOS"); self.orden=tk.StringVar(value="fecha_desc")
        self.build(); self.refresh()
    def build(self):
        top=ttk.Frame(self,padding=10); top.pack(fill="x")
        ttk.Label(top,text="Buscar:").pack(side="left"); ttk.Entry(top,textvariable=self.term,width=28).pack(side="left",padx=4)
        ttk.Label(top,text="Estado:").pack(side="left",padx=(10,2)); ttk.Combobox(top,textvariable=self.estado,values=["TODOS"]+ESTADOS_OC,state="readonly",width=19).pack(side="left")
        ttk.Label(top,text="Tipo:").pack(side="left",padx=(10,2)); ttk.Combobox(top,textvariable=self.tipo,values=["TODOS"]+TIPOS_ORDEN_COMPRA,state="readonly",width=12).pack(side="left")
        ttk.Label(top,text="Orden:").pack(side="left",padx=(10,2)); ttk.Combobox(top,textvariable=self.orden,values=["fecha_desc","fecha_asc","importe_asc","importe_desc"],state="readonly",width=14).pack(side="left")
        ttk.Button(top,text="Filtrar",command=self.refresh).pack(side="left",padx=6)
        cols=("numero","fecha","proveedor","moneda","importe","estado","pago","avance","tipo","cotreq")
        tree_box=ttk.Frame(self); tree_box.pack(fill="both",expand=True,padx=10,pady=8)
        tree_box.rowconfigure(0,weight=1); tree_box.columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(tree_box,columns=cols,show="headings",height=18)
        specs=[("numero","Nro. Documento",135),("fecha","Fecha",90),("proveedor","Proveedor",250),("moneda","Moneda",75),("importe","Importe",90),("estado","Estado operativo",145),("pago","Estado pago",115),("avance","% total",65),("tipo","Tipo",90),("cotreq","Cotiza/Requerimiento",180)]
        for c,t,w in specs: self.tree.heading(c,text=t); self.tree.column(c,width=w,minwidth=w,anchor="center" if c!="proveedor" else "w",stretch=False)
        vsb=ttk.Scrollbar(tree_box,orient="vertical",command=self.tree.yview); hsb=ttk.Scrollbar(tree_box,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        btn=ttk.Frame(self,padding=10); btn.pack(fill="x")
        ttk.Button(btn,text="Nueva OC",command=lambda: OCManualDialog(self.app,self.refresh)).pack(side="left",padx=4)
        ttk.Button(btn,text="Visualizar",command=self.view).pack(side="left",padx=4)
        ttk.Button(btn,text="Modificar",command=self.modify).pack(side="left",padx=4)
        ttk.Button(btn,text="Imprimir",command=self.print_oc).pack(side="left",padx=4)
        ttk.Button(btn,text="Anular",command=self.annul).pack(side="left",padx=4)
        ttk.Button(btn,text="Actualizar",command=self.refresh).pack(side="right",padx=4)
    def refresh(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        for r in self.db.list_ordenes_compra(self.term.get(),self.estado.get(),self.tipo.get(),orden=self.orden.get()):
            cotreq = r["cotizacion"] or r["requerimiento"]
            estado_pago, pagado, saldo = self.db._sync_oc_payment_status(r["id"], commit=True)
            avance_total = max(int(r["avance"] or 0), 50 if estado_pago in ("PAGO PARCIAL", "PAGADA") else 0)
            self.tree.insert("","end",values=(r["numero"],fecha_text_from_iso(r["fecha_orden"]),r["razon_social"],r["moneda"],f"{r['total']:.2f}",r["estado"],estado_pago,f"{avance_total}%",r["tipo_orden"],cotreq))
    def selected_num(self):
        sel=self.tree.selection(); return self.tree.item(sel[0],"values")[0] if sel else None
    def view(self):
        n=self.selected_num();
        if n: OCDetailDialog(self,self.db,n)
    def modify(self):
        if not self.app.require_logistica(self): return
        n=self.selected_num()
        if n: OCModificarDialog(self.app, n, self.refresh)
    def print_oc(self):
        n=self.selected_num();
        if n: OCPrintPreviewDialog(self,self.db,n)
    def annul(self):
        if not self.app.require_action("anular_oc", self): return
        n=self.selected_num();
        if not n: return
        m=simpledialog.askstring("Anular OC","Motivo de anulación:",parent=self)
        if not m: return
        try:
            self.db.anular_orden_compra(n,m,self.app.usuario_actual); messagebox.showinfo("OC","Orden anulada.",parent=self); self.refresh()
        except Exception as e: messagebox.showerror("OC",str(e),parent=self)


class OCAprobacionDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.app=app; self.db=app.db
        self.title("Aprobación de Órdenes de Compra"); self.geometry("1120x610+80+65"); bind_as_child_window(self,self.master if hasattr(self,"master") else None)
        self.term=tk.StringVar(); self.fecha_modo=tk.StringVar(value="MES"); self.fecha=tk.StringVar(value=today_text())
        today=date.today(); self.mes=tk.StringVar(value=f"{today.month:02d}"); self.anio=tk.StringVar(value=str(today.year))
        self.ver_aprobadas=tk.BooleanVar(value=True)
        self.build(); self.refresh()
    def build(self):
        top=ttk.Frame(self,padding=10); top.pack(fill="x")
        ttk.Label(top,text="OC/RUC/Proveedor:").pack(side="left")
        ent=ttk.Entry(top,textvariable=self.term,width=26); ent.pack(side="left",padx=4); ent.bind("<Return>",lambda e:self.refresh())
        ttk.Label(top,text="Filtro fecha:").pack(side="left",padx=(10,2))
        ttk.Combobox(top,textvariable=self.fecha_modo,values=["DIA","MES","TODOS"],state="readonly",width=7).pack(side="left")
        ttk.Label(top,text="Día:").pack(side="left",padx=(8,2)); ttk.Entry(top,textvariable=self.fecha,width=11).pack(side="left")
        ttk.Label(top,text="Mes:").pack(side="left",padx=(8,2)); ttk.Combobox(top,textvariable=self.mes,values=[f"{i:02d}" for i in range(1,13)],state="readonly",width=5).pack(side="left")
        ttk.Label(top,text="Año:").pack(side="left",padx=(8,2)); ttk.Entry(top,textvariable=self.anio,width=7).pack(side="left")
        ttk.Checkbutton(top,text="Ver aprobadas",variable=self.ver_aprobadas).pack(side="left",padx=10)
        ttk.Button(top,text="Filtrar",command=self.refresh).pack(side="left",padx=4)
        cols=("numero","fecha","ruc","proveedor","tipo","total","estado","aprobador")
        tree_box=ttk.Frame(self); tree_box.pack(fill="both",expand=True,padx=10,pady=10)
        tree_box.rowconfigure(0,weight=1); tree_box.columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(tree_box,columns=cols,show="headings",height=17)
        specs=[("numero","OC",135),("fecha","Fecha",90),("ruc","RUC",110),("proveedor","Proveedor",285),("tipo","Tipo",90),("total","Total",100),("estado","Estado",150),("aprobador","Aprobado por",130)]
        for c,t,w in specs: self.tree.heading(c,text=t); self.tree.column(c,width=w,minwidth=w,anchor="center" if c not in ("proveedor",) else "w",stretch=False)
        vsb=ttk.Scrollbar(tree_box,orient="vertical",command=self.tree.yview); hsb=ttk.Scrollbar(tree_box,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        self.tree.bind("<Double-1>", lambda e:self.view())
        self.tree.bind("<Return>", lambda e:self.view())
        btn=ttk.Frame(self,padding=10); btn.pack(fill="x")
        ttk.Button(btn,text="Visualizar",command=self.view).pack(side="left",padx=4)
        ttk.Button(btn,text="Aprobar total",command=self.approve_total).pack(side="left",padx=4)
        ttk.Button(btn,text="Aprobar parcial",command=self.approve_partial).pack(side="left",padx=4)
        ttk.Button(btn,text="Denegar",command=self.deny).pack(side="left",padx=4)
        ttk.Button(btn,text="Desaprobar / devolver a pendiente",command=self.unapprove).pack(side="left",padx=4)
        ttk.Button(btn,text="Actualizar",command=self.refresh).pack(side="right",padx=4)
    def _date_range(self):
        modo=normalize_text(self.fecha_modo.get())
        if modo=="TODOS": return "",""
        if modo=="DIA": return self.fecha.get(), self.fecha.get()
        try:
            f=parse_fecha(f"01/{self.mes.get()}/{self.anio.get()}")
            hasta=date(f.year,12,31) if f.month==12 else date(f.year,f.month+1,1)-timedelta(days=1)
            return f.strftime(DATE_FMT), hasta.strftime(DATE_FMT)
        except Exception:
            return "",""
    def refresh(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        estados=["PENDIENTE","EN COTIZACION"]
        if self.ver_aprobadas.get(): estados += ["APROBADA","APROBADA PARCIAL"]
        desde,hasta=self._date_range()
        first=None
        for estado in estados:
            for r in self.db.list_ordenes_compra(term=self.term.get(), estado=estado, desde=desde, hasta=hasta):
                iid=self.tree.insert("","end",values=(r["numero"],fecha_text_from_iso(r["fecha_orden"]),r["ruc"],r["razon_social"],r["tipo_orden"],f"{r['total']:.2f}",r["estado"],r["aprobado_por"] or ""))
                if first is None: first=iid
        if first:
            self.tree.selection_set(first); self.tree.focus(first)
    def selected(self):
        sel=self.tree.selection(); return self.tree.item(sel[0],"values")[0] if sel else None
    def view(self):
        n=self.selected();
        if n: OCDetailDialog(self,self.db,n)
    def approve_total(self):
        if not self.app.require_action("aprobar_oc", self): return
        n=self.selected();
        if not n: return
        oc,det=self.db.get_orden_compra(n)
        try:
            self.db.approve_orden_compra(n,{r["id"]:r["cantidad_solicitada"] for r in det},self.app.usuario_actual)
            messagebox.showinfo("Aprobación","OC aprobada.",parent=self); self.refresh()
        except Exception as e: messagebox.showerror("Aprobación",str(e),parent=self)
    def approve_partial(self):
        if not self.app.require_action("aprobar_oc", self): return
        n=self.selected()
        if not n: return
        OCAprobacionParcialDialog(self, self.app, n, after_save=self.refresh)
    def deny(self):
        if not self.app.require_action("aprobar_oc", self): return
        n=self.selected();
        if not n: return
        motivo=simpledialog.askstring("Denegar","Motivo:",parent=self)
        if not motivo: return
        try:
            self.db.deny_orden_compra(n,motivo,self.app.usuario_actual); messagebox.showinfo("OC","OC denegada.",parent=self); self.refresh()
        except Exception as e: messagebox.showerror("OC",str(e),parent=self)
    def unapprove(self):
        if not self.app.require_action("aprobar_oc", self): return
        n=self.selected()
        if not n: return
        motivo=simpledialog.askstring("Desaprobar OC", "Motivo de la desaprobación/devolución a pendiente:", parent=self)
        if not motivo: return
        try:
            estado=self.db.desaprobar_orden_compra(n,motivo,self.app.usuario_actual)
            messagebox.showinfo("OC", f"OC devuelta a estado {estado}.", parent=self)
            self.refresh()
        except Exception as e:
            messagebox.showerror("No se pudo desaprobar", str(e), parent=self)


class OCAprobacionParcialDialog(tk.Toplevel):
    """Aprobación parcial dinámica por ítem, sin preguntar item por item."""
    def __init__(self, parent, app, numero, after_save=None):
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self.numero = numero
        self.after_save = after_save
        self.title(f"Aprobación parcial - {numero}")
        self.geometry("1120x620+80+70")
        self.minsize(1040, 560)
        bind_as_child_window(self, parent)
        self.qtys = {}
        self.build()

    def build(self):
        oc, det = self.db.get_orden_compra(self.numero)
        self.oc, self.det = oc, det
        if not oc:
            ttk.Label(self, text="OC no encontrada").pack(padx=20, pady=20)
            return
        top = ttk.LabelFrame(self, text="Cabecera", padding=10)
        top.pack(fill="x", padx=10, pady=8)
        info = f"{oc['numero']} | Proveedor: {oc['razon_social']} | Total: {oc['moneda']} {float(oc['total'] or 0):.2f} | Usuario aprobador: {self.app.usuario_actual}"
        ttk.Label(top, text=info, font=("Arial", 10, "bold"), wraplength=980).pack(anchor="w")
        ttk.Label(top, text="Seleccione un item y use 'Aprobar parcial' o edite con doble clic la columna Cant. aprobada.", foreground="#555").pack(anchor="w", pady=(4,0))

        box = ttk.Frame(self)
        box.pack(fill="both", expand=True, padx=10, pady=8)
        box.rowconfigure(0, weight=1); box.columnconfigure(0, weight=1)
        cols = ("id", "item", "codigo", "descripcion", "um", "solicitado", "aprobado", "precio", "total")
        self.tree = ttk.Treeview(box, columns=cols, show="headings", height=16)
        specs = [("id","ID",55),("item","Item",55),("codigo","Código",105),("descripcion","Descripción",390),("um","UM",65),("solicitado","Solicitado",95),("aprobado","Cant. aprobada",110),("precio","P.Unit",90),("total","Total",100)]
        for c,t,w in specs:
            self.tree.heading(c, text=t); self.tree.column(c, width=w, minwidth=w, anchor="center" if c!="descripcion" else "w", stretch=False)
        vsb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview); hsb = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self.edit_selected)
        for r in det:
            val = float(r["cantidad_solicitada"] or 0)
            self.qtys[r["id"]] = val
            self.tree.insert("", "end", iid=str(r["id"]), values=(r["id"], r["item"], r["codigo"], r["descripcion"], r["unidad_medida"], f"{val:g}", f"{val:g}", f"{float(r['precio_unitario_sin_igv'] or 0):.2f}", f"{float(r['total'] or 0):.2f}"))

        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")
        ttk.Label(bottom, text="Motivo:").pack(side="left", padx=4)
        self.motivo = tk.StringVar(value="APROBACIÓN PARCIAL")
        ttk.Entry(bottom, textvariable=self.motivo, width=45).pack(side="left", padx=4)
        ttk.Button(bottom, text="Aprobar parcial item", command=self.edit_selected).pack(side="left", padx=4)
        ttk.Button(bottom, text="Aprobar todo", command=self.set_all_full).pack(side="left", padx=4)
        ttk.Button(bottom, text="Guardar aprobación", command=self.save).pack(side="right", padx=4)
        ttk.Button(bottom, text="Cerrar", command=self.destroy).pack(side="right", padx=4)

    def edit_selected(self, event=None):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Seleccione", "Seleccione un item.", parent=self); return
        iid = sel[0]
        vals = self.tree.item(iid, "values")
        solicitado = float(str(vals[5]).replace(",", "."))
        actual = self.qtys.get(int(iid), solicitado)
        val = simpledialog.askstring("Cantidad aprobada", f"Item {vals[1]} - {vals[3]}\nSolicitado: {solicitado:g}\nAprobar:", initialvalue=f"{actual:g}", parent=self)
        if val is None:
            return
        try:
            aprob = float(str(val).replace(",", "."))
            if aprob < 0 or aprob - solicitado > 0.00001:
                raise ValueError("La cantidad aprobada debe estar entre 0 y lo solicitado.")
            self.qtys[int(iid)] = aprob
            vals = list(vals); vals[6] = f"{aprob:g}"
            self.tree.item(iid, values=vals)
        except Exception as e:
            messagebox.showerror("Cantidad", str(e), parent=self)

    def set_all_full(self):
        for r in self.det:
            val = float(r["cantidad_solicitada"] or 0)
            self.qtys[r["id"]] = val
            iid = str(r["id"])
            vals = list(self.tree.item(iid, "values")); vals[6] = f"{val:g}"
            self.tree.item(iid, values=vals)

    def save(self):
        try:
            estado = self.db.approve_orden_compra(self.numero, self.qtys, self.app.usuario_actual, self.motivo.get())
            messagebox.showinfo("Aprobación", f"Estado: {estado}", parent=self)
            if self.after_save:
                self.after_save()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Aprobación", str(e), parent=self)



class OCRecepcionDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.app=app; self.db=app.db; self.det=[]
        self.title("Registro de Entrada con Orden de Compra"); self.geometry("1220x620+60+70"); bind_as_child_window(self,self.master if hasattr(self,"master") else None)
        self.numero=tk.StringVar(); self.documento=tk.StringVar(); self.obs=tk.StringVar(value="INGRESO POR ORDEN DE COMPRA"); self.fecha=tk.StringVar(value=today_text())
        self.build()
    def build(self):
        top=ttk.Frame(self,padding=10); top.pack(fill="x")
        self.oc_entry = None
        for lab,var,w in [("OC",self.numero,20),("Fecha",self.fecha,14),("Factura/Guía",self.documento,24),("Observación",self.obs,30)]:
            ttk.Label(top,text=lab+":").pack(side="left",padx=3)
            ent=ttk.Entry(top,textvariable=var,width=w); ent.pack(side="left",padx=3)
            if lab=="OC":
                self.oc_entry=ent; ent.bind("<Shift-F1>", lambda e: self.open_oc()); ent.bind("<Return>", lambda e: self.load())
        ttk.Button(top,text="Buscar OC",command=self.open_oc).pack(side="left",padx=5)
        ttk.Label(top,text="Enter en OC carga automáticamente",foreground="#666").pack(side="left",padx=5)
        cols=("id","item","codigo","descripcion","um","aprob","ing","pend")
        tree_box=ttk.Frame(self); tree_box.pack(fill="both",expand=True,padx=10,pady=8)
        tree_box.rowconfigure(0,weight=1); tree_box.columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(tree_box,columns=cols,show="headings",height=16)
        specs=[("id","ID",50),("item","Item",50),("codigo","Código",100),("descripcion","Descripción",420),("um","UM",60),("aprob","Aprob.",90),("ing","Ingresado",90),("pend","Pendiente",90)]
        for c,t,w in specs: self.tree.heading(c,text=t); self.tree.column(c,width=w,minwidth=w,anchor="center" if c!="descripcion" else "w",stretch=False)
        vsb=ttk.Scrollbar(tree_box,orient="vertical",command=self.tree.yview); hsb=ttk.Scrollbar(tree_box,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        btn=ttk.Frame(self,padding=10); btn.pack(fill="x")
        ttk.Button(btn,text="Grabar",command=self.receive_all).pack(side="left",padx=4)
        ttk.Button(btn,text="Ingreso parcial",command=self.receive_partial).pack(side="left",padx=4)
        ttk.Button(btn,text="Visualizar OC",command=self.view_oc).pack(side="left",padx=4)
        ttk.Button(btn,text="Cerrar",command=self.destroy).pack(side="right",padx=4)
    def open_oc(self):
        OCSelectDialog(self, self.db, self.set_oc, estados=("APROBADA","APROBADA PARCIAL","PAGO PARCIAL","PAGADA","EN TRANSITO","ATENDIDA PARCIALMENTE"), tipo="NACIONAL", titulo="Buscar OC para ingreso a almacén")
    def set_oc(self, numero):
        self.numero.set(numero); self.load()
    def view_oc(self):
        if not self.numero.get():
            messagebox.showinfo("OC","Seleccione una OC.",parent=self); return
        OCDetailDialog(self,self.db,self.numero.get())
    def load(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        oc,self.det=self.db.get_orden_compra(self.numero.get())
        if not oc: messagebox.showwarning("OC","No encontrada.",parent=self); return
        for r in self.det:
            pend=float(r["cantidad_aprobada"] or 0)-float(r["cantidad_ingresada"] or 0)
            if pend>0:
                self.tree.insert("","end",values=(r["id"],r["item"],r["codigo"],r["descripcion"],r["unidad_medida"],f"{r['cantidad_aprobada']:g}",f"{r['cantidad_ingresada']:g}",f"{pend:g}"))
    def receive_all(self):
        if not self.det:
            messagebox.showinfo("OC", "Cargue primero una orden de compra.", parent=self)
            return
        qty={r["id"]:float(r["cantidad_aprobada"] or 0)-float(r["cantidad_ingresada"] or 0) for r in self.det}
        self._receive(qty)
    def receive_partial(self):
        if not self.det:
            messagebox.showinfo("OC", "Cargue primero una orden de compra.", parent=self)
            return
        qty={}
        for r in self.det:
            pend=float(r["cantidad_aprobada"] or 0)-float(r["cantidad_ingresada"] or 0)
            if pend<=0: continue
            val=simpledialog.askstring("Recibir",f"{r['descripcion']}\nPendiente: {pend:g}\nRecibir:",initialvalue=str(pend),parent=self)
            if val is None: return
            try:
                qty[r["id"]]=float(str(val).replace(",","."))
            except Exception:
                messagebox.showerror("Cantidad inválida", f"Ingrese una cantidad numérica válida para {r['descripcion']}.", parent=self)
                return
        self._receive(qty)
    def _receive(self,qty):
        # V42: el permiso debe validarse contra el almacén REAL de la OC,
        # no contra el almacén seleccionado al iniciar sesión.
        oc, _ = self.db.get_orden_compra(self.numero.get())
        if not oc:
            messagebox.showwarning("OC", "No encontrada.", parent=self)
            return
        if not self.app.require_almacen(self, oc["almacen_codigo"]):
            return
        try:
            vale,estado=self.db.recepcionar_orden_compra(self.numero.get(),self.fecha.get(),self.documento.get(),self.obs.get(),qty,self.app.usuario_actual)
            messagebox.showinfo("Ingreso registrado",f"Vale: {vale}\nEstado OC: {estado}",parent=self); self.load()
        except Exception as e: messagebox.showerror("Ingreso OC",str(e),parent=self)


class OCPagoDialog(tk.Toplevel):
    def __init__(self, app, close_on_save=False, after_save=None):
        super().__init__(app.root); self.app=app; self.db=app.db; self.close_on_save=close_on_save; self.after_save=after_save
        self.title("Finanzas - Registrar Pago de OC"); self.geometry("820x390+220+120"); bind_as_child_window(self,self.master if hasattr(self,"master") else None)
        self.numero=tk.StringVar(); self.fecha=tk.StringVar(value=today_text()); self.factura=tk.StringVar(); self.operacion=tk.StringVar(); self.banco=tk.StringVar(value="01 BCP"); self.banco_otro=tk.StringVar(); self.monto=tk.StringVar(); self.obs=tk.StringVar(value="PAGO DE ORDEN DE COMPRA"); self.info=tk.StringVar(value="")
        self.build()
    def build(self):
        frm=ttk.Frame(self,padding=18); frm.pack(fill="both",expand=True)
        rows=[("OC",self.numero),("Fecha pago",self.fecha),("Nro factura",self.factura),("Nro operación",self.operacion),("Monto",self.monto),("Observación",self.obs)]
        self.oc_entry=None
        for i,(lab,var) in enumerate(rows):
            ttk.Label(frm,text=lab+":").grid(row=i,column=0,sticky="w",pady=6)
            ent=ttk.Entry(frm,textvariable=var,width=35); ent.grid(row=i,column=1,sticky="w",pady=6)
            if lab=="OC":
                self.oc_entry=ent; ent.bind("<Shift-F1>", lambda e: self.open_oc())
        ttk.Label(frm,text="Banco:").grid(row=3,column=2,sticky="w",padx=8); ttk.Combobox(frm,textvariable=self.banco,values=BANCOS,state="readonly",width=18).grid(row=3,column=3,sticky="w")
        ttk.Label(frm,text="Banco otros:").grid(row=4,column=2,sticky="w",padx=8); ttk.Entry(frm,textvariable=self.banco_otro,width=20).grid(row=4,column=3,sticky="w")
        ttk.Button(frm,text="Buscar OC",command=self.open_oc).grid(row=0,column=2,padx=8)
        ttk.Label(frm,text="Enter en OC carga automáticamente",foreground="#666").grid(row=0,column=3,padx=8,sticky="w")
        if self.oc_entry:
            self.oc_entry.bind("<Return>", lambda e: self.load())
        ttk.Label(frm,textvariable=self.info,wraplength=700,foreground="#444").grid(row=7,column=0,columnspan=4,sticky="w",pady=8)
        ttk.Button(frm,text="Registrar pago",command=self.save).grid(row=8,column=3,sticky="e",pady=14)
    def open_oc(self):
        OCSelectDialog(self, self.db, self.set_oc, estados=("APROBADA","APROBADA PARCIAL","PAGO PARCIAL","PAGADA","EN TRANSITO","ATENDIDA PARCIALMENTE","ATENDIDA","LIQUIDADA"), titulo="Buscar OC para pago")
    def set_oc(self, numero):
        self.numero.set(numero); self.load()
    def load(self):
        oc,det=self.db.get_orden_compra(self.numero.get())
        if not oc: self.info.set("OC no encontrada."); return
        estado_pago, pagado, saldo = self.db._sync_oc_payment_status(oc["id"], commit=True)
        self.monto.set(f"{saldo:.2f}")
        vence = fecha_text_from_iso(oc["fecha_vencimiento_pago"]) if "fecha_vencimiento_pago" in oc.keys() and oc["fecha_vencimiento_pago"] else ""
        self.info.set(f"{oc['numero']} | {oc['razon_social']} | Estado operativo {oc['estado']} ({oc['avance']}%) | Pago {estado_pago} | Vence {vence} | Total {oc['total']:.2f} | Pagado {pagado:.2f} | Saldo {saldo:.2f}")
    def clear_form(self):
        self.numero.set(""); self.fecha.set(today_text()); self.factura.set(""); self.operacion.set(""); self.banco.set("01 BCP"); self.banco_otro.set(""); self.monto.set(""); self.obs.set("PAGO DE ORDEN DE COMPRA"); self.info.set("")
    def save(self):
        if not self.app.require_action("registrar_pago", self): return
        try:
            banco = normalizar_banco(self.banco.get(), self.banco_otro.get())
            estado=self.db.registrar_pago_oc(self.numero.get(),self.fecha.get(),self.operacion.get(),banco,self.monto.get(),self.obs.get(),self.app.usuario_actual, numero_factura=self.factura.get())
            messagebox.showinfo("Pago",f"Pago registrado. Estado de pago: {estado}",parent=self)
            if self.after_save:
                self.after_save()
            if self.close_on_save:
                self.destroy()
            else:
                self.clear_form()
        except Exception as e: messagebox.showerror("Pago",str(e),parent=self)


class OCTransitoDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.app=app; self.db=app.db
        self.title("Logística - Registrar OC en Tránsito"); self.geometry("760x350+220+120"); bind_as_child_window(self,self.master if hasattr(self,"master") else None)
        self.numero=tk.StringVar(); self.operador=tk.StringVar(value="FLORES"); self.fecha=tk.StringVar(value=today_text()); self.llegada=tk.StringVar(value=today_text()); self.guia=tk.StringVar(); self.obs=tk.StringVar(value="MATERIAL EN TRANSITO")
        self.build()
    def build(self):
        frm=ttk.Frame(self,padding=18); frm.pack(fill="both",expand=True)
        fields=[("OC",self.numero),("Fecha despacho",self.fecha),("Llegada estimada",self.llegada),("Guía transportista",self.guia),("Observación",self.obs)]
        for i,(lab,var) in enumerate(fields):
            ttk.Label(frm,text=lab+":").grid(row=i,column=0,sticky="w",pady=6)
            ent=ttk.Entry(frm,textvariable=var,width=35); ent.grid(row=i,column=1,sticky="w",pady=6)
            if lab=="OC": ent.bind("<Shift-F1>", lambda e: self.open_oc())
        ttk.Label(frm,text="Operador:").grid(row=1,column=2,sticky="w",padx=8); ttk.Combobox(frm,textvariable=self.operador,values=OPERADORES_LOGISTICOS,state="readonly",width=18).grid(row=1,column=3,sticky="w")
        ttk.Button(frm,text="Buscar OC",command=self.open_oc).grid(row=0,column=2,padx=8)
        ttk.Button(frm,text="Visualizar OC",command=self.view_oc).grid(row=0,column=3,padx=8)
        ttk.Button(frm,text="Registrar tránsito",command=self.save).grid(row=6,column=3,sticky="e",pady=14)
    def open_oc(self):
        OCSelectDialog(self, self.db, self.numero.set, estados=("APROBADA","APROBADA PARCIAL","PAGO PARCIAL","PAGADA"), tipo="NACIONAL", titulo="Buscar OC para tránsito")
    def view_oc(self):
        if not self.numero.get():
            messagebox.showinfo("OC","Seleccione una OC.",parent=self); return
        OCDetailDialog(self,self.db,self.numero.get())
    def save(self):
        if not self.app.require_logistica(self): return
        try:
            estado=self.db.registrar_transito_oc(self.numero.get(),self.operador.get(),self.fecha.get(),self.llegada.get(),self.guia.get(),self.obs.get(),self.app.usuario_actual)
            messagebox.showinfo("Tránsito",f"Estado: {estado}",parent=self)
            self.numero.set(""); self.guia.set(""); self.obs.set("MATERIAL EN TRANSITO")
        except Exception as e: messagebox.showerror("Tránsito",str(e),parent=self)


class OCLiquidacionDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.app=app; self.db=app.db
        self.title("Liquidación de Orden de Compra"); self.geometry("740x350+220+120"); bind_as_child_window(self,self.master if hasattr(self,"master") else None)
        self.numero=tk.StringVar(); self.fecha=tk.StringVar(value=today_text()); self.factura=tk.StringVar(); self.guia=tk.StringVar(); self.obs=tk.StringVar(value="CONFORME PARA LIQUIDACION")
        self.build()
    def build(self):
        frm=ttk.Frame(self,padding=18); frm.pack(fill="both",expand=True)
        for i,(lab,var) in enumerate([("OC",self.numero),("Fecha",self.fecha),("Factura",self.factura),("Guía",self.guia),("Observación/Conformidad",self.obs)]):
            ttk.Label(frm,text=lab+":").grid(row=i,column=0,sticky="w",pady=6)
            ent=ttk.Entry(frm,textvariable=var,width=42); ent.grid(row=i,column=1,sticky="w",pady=6)
            if lab=="OC": ent.bind("<Shift-F1>", lambda e: self.open_oc())
        ttk.Button(frm,text="Buscar OC",command=self.open_oc).grid(row=0,column=2,padx=8)
        ttk.Button(frm,text="Visualizar OC",command=lambda: OCDetailDialog(self,self.db,self.numero.get())).grid(row=0,column=3,padx=8)
        ttk.Button(frm,text="Liquidar",command=self.save).grid(row=6,column=3,sticky="e",pady=14)
    def open_oc(self):
        OCSelectDialog(self, self.db, self.numero.set, estados=("APROBADA","APROBADA PARCIAL","PAGO PARCIAL","PAGADA","ATENDIDA PARCIALMENTE","ATENDIDA"), titulo="Buscar OC para liquidación")
    def save(self):
        if not self.app.require_logistica(self): return
        try:
            estado=self.db.liquidar_orden_compra(self.numero.get(),self.fecha.get(),self.factura.get(),self.guia.get(),self.obs.get(),self.app.usuario_actual)
            messagebox.showinfo("Liquidación",f"Estado: {estado}",parent=self)
            self.numero.set(""); self.fecha.set(today_text()); self.factura.set(""); self.guia.set(""); self.obs.set("CONFORME PARA LIQUIDACION")
        except Exception as e: messagebox.showerror("Liquidación",str(e),parent=self)



class OCRequerimientoDialog(tk.Toplevel):
    """Primera pantalla: solo selección/carga del requerimiento.
    La matriz de cotización se abre en una segunda ventana para que el usuario trabaje los precios como tabla.
    """
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.title("Orden de Compra por Requerimiento")
        self.geometry("760x260+220+140")
        bind_as_child_window(self, self.master if hasattr(self, "master") else None)
        self.req = tk.StringVar()
        self.req_nums = []
        self.solicitante = tk.StringVar()
        self.cc = tk.StringVar()
        self.ot = tk.StringVar()
        self.tipo = tk.StringVar()
        self.almacen = tk.StringVar()
        self.build()

    def build(self):
        frm = ttk.LabelFrame(self, text="Seleccione requerimiento", padding=14)
        frm.pack(fill="both", expand=True, padx=12, pady=12)
        ttk.Label(frm, text="Requerimiento:").grid(row=0, column=0, sticky="w", padx=4, pady=6)
        ent = ttk.Entry(frm, textvariable=self.req, width=28)
        ent.grid(row=0, column=1, sticky="w", padx=4, pady=6)
        ent.bind("<Shift-F1>", lambda e: self.open_req())
        ttk.Button(frm, text="Buscar Req", command=self.open_req).grid(row=0, column=2, padx=4)
        ttk.Button(frm, text="Continuar", command=self.load_req).grid(row=0, column=3, padx=4)
        rows = [("Solicitante", self.solicitante), ("Centro de costo", self.cc), ("OT", self.ot), ("Tipo", self.tipo), ("Almacén", self.almacen)]
        for i, (lab, var) in enumerate(rows, start=1):
            ttk.Label(frm, text=lab + ":").grid(row=i, column=0, sticky="w", padx=4, pady=4)
            ttk.Entry(frm, textvariable=var, state="readonly", width=55).grid(row=i, column=1, columnspan=3, sticky="we", padx=4, pady=4)
        ttk.Label(frm, text="Tip: presione SHIFT + F1 en Requerimiento para buscarlo.", foreground="#555").grid(row=6, column=0, columnspan=4, sticky="w", padx=4, pady=(8, 0))
        btn = ttk.Frame(self, padding=(12, 0, 12, 12)); btn.pack(fill="x")
        ttk.Button(btn, text="Visualizar requerimiento", command=self.view_req).pack(side="left", padx=4)
        ttk.Button(btn, text="Cerrar", command=self.destroy).pack(side="right", padx=4)
        ent.focus_set()

    def open_req(self):
        RequerimientoMultiSelectDialog(self, self.db, self.set_req)

    def set_req(self, req):
        if isinstance(req, (list, tuple)):
            self.req_nums = list(req)
            self.req.set(";".join(self.req_nums))
        else:
            self.req_nums = [req]
            self.req.set(req)
        self.preview_req()

    def preview_req(self):
        try:
            if self.req_nums and len(self.req_nums) > 1:
                req, det = self.db.collect_requerimientos_for_oc(self.req_nums)
            else:
                req, det = self.db.find_requerimiento(self.req.get())
        except Exception as e:
            messagebox.showwarning("Requerimiento", str(e), parent=self)
            return
        if not req:
            return
        self.solicitante.set(req["solicitante"])
        self.cc.set(f"{req['centro_costo_codigo']} - {req['centro_costo_nombre']}")
        self.ot.set(req["ot"] or "")
        self.tipo.set(req["tipo_orden"] if "tipo_orden" in req.keys() else "NACIONAL")
        self.almacen.set(f"{req['almacen_codigo']} - {req['almacen_nombre']}")

    def load_req(self):
        try:
            if self.req_nums and len(self.req_nums) > 1:
                req, det = self.db.collect_requerimientos_for_oc(self.req_nums)
            else:
                req, det = self.db.find_requerimiento(self.req.get())
        except Exception as e:
            messagebox.showwarning("Req", str(e), parent=self)
            return
        if not req:
            messagebox.showwarning("Req", "No se encontró el requerimiento.", parent=self)
            return
        if req["estado"] in ("ANULADO", "ATENDIDO"):
            messagebox.showwarning("Req", f"No se puede generar OC de un requerimiento {req['estado']}.", parent=self)
            return
        items = []
        for r in det:
            pend = float(r["cantidad_requerida"] or 0) - float(r["cantidad_atendida"] or 0)
            if pend > 0:
                items.append({
                    "codigo": r["codigo"],
                    "descripcion": r["descripcion"],
                    "um": r["unidad_medida"],
                    "cantidad": pend,
                })
        if not items:
            messagebox.showinfo("Req", "El requerimiento no tiene cantidades pendientes.", parent=self)
            return
        OCRequerimientoMatrizDialog(self.app, self.db, req, items)
        self.destroy()

    def view_req(self):
        if not self.req.get():
            messagebox.showinfo("Requerimiento", "Seleccione un requerimiento.", parent=self)
            return
        RequerimientoDetailDialog(self, self.app, self.req.get())


class OCRequerimientoMatrizDialog(tk.Toplevel):
    """Matriz editable: ITEM/DESCRIPCIÓN/UND/CANTIDAD + columnas dinámicas por proveedor.
    El precio se edita directamente sobre la celda, sin ventanas flotantes.
    """
    def __init__(self, app, db, req, items):
        super().__init__(app.root)
        self.app = app
        self.db = db
        self.req_row = req
        self.det = items
        self.providers = []
        self.igv = tk.StringVar(value="SIN IGV")  # valor por defecto para nuevos proveedores
        self.forma_pago = tk.StringVar(value="CONTADO")
        self.fecha_tc = tk.StringVar(value=today_text())
        self.context_provider_idx = None
        self.context_row_id = None
        self.editor = None
        self.title(f"Matriz de cotización - {req['requerimiento']}")
        self.geometry("1340x760+25+25")
        self.minsize(1220, 680)
        bind_as_child_window(self, self.master if hasattr(self, "master") else None)
        self.build()

    @property
    def requerimiento(self):
        return self.req_row["requerimiento"]

    def build(self):
        top = ttk.LabelFrame(self, text="Datos del requerimiento", padding=10)
        top.pack(fill="x", padx=10, pady=8)
        info = [
            ("Requerimiento", self.req_row["requerimiento"]),
            ("Solicitante", self.req_row["solicitante"]),
            ("Centro costo", f"{self.req_row['centro_costo_codigo']} - {self.req_row['centro_costo_nombre']}"),
            ("OT", self.req_row["ot"] or ""),
            ("Tipo", self.req_row["tipo_orden"] if "tipo_orden" in self.req_row.keys() else "NACIONAL"),
            ("Almacén", f"{self.req_row['almacen_codigo']} - {self.req_row['almacen_nombre']}"),
        ]
        for i, (lab, val) in enumerate(info):
            r = i // 3
            c = (i % 3) * 2
            ttk.Label(top, text=lab + ":", font=("Arial", 9, "bold")).grid(row=r, column=c, sticky="w", padx=4, pady=3)
            ttk.Label(top, text=str(val), wraplength=300).grid(row=r, column=c+1, sticky="w", padx=4, pady=3)
        ttk.Label(top, text="Forma pago:").grid(row=2, column=0, sticky="w", padx=4, pady=6)
        ttk.Combobox(top, textvariable=self.forma_pago, values=FORMAS_PAGO, state="readonly", width=22).grid(row=2, column=1, sticky="w", padx=4, pady=6)
        ttk.Label(top, text="Fecha TC:").grid(row=2, column=2, sticky="w", padx=4, pady=6)
        ttk.Entry(top, textvariable=self.fecha_tc, width=12).grid(row=2, column=3, sticky="w", padx=4, pady=6)
        ttk.Label(top, text="IGV, moneda y tipo de OC se configuran dentro de cada proveedor, para evitar datos contradictorios en la matriz.", foreground="#555").grid(row=3, column=0, columnspan=6, sticky="w", padx=4, pady=2)
        ttk.Label(top, text="Doble clic en una celda de proveedor para escribir el precio. Clic derecho en la matriz para agregar proveedor.", foreground="#555").grid(row=4, column=0, columnspan=6, sticky="w", padx=4, pady=2)

        bar = ttk.Frame(self, padding=(10, 0, 10, 4))
        bar.pack(fill="x")
        ttk.Button(bar, text="Agregar proveedor", command=self.add_provider).pack(side="left", padx=4)
        ttk.Button(bar, text="Modificar proveedor", command=self.modify_provider).pack(side="left", padx=4)
        ttk.Button(bar, text="Eliminar proveedor", command=self.remove_provider).pack(side="left", padx=4)
        self.provider_info = tk.StringVar(value="Sin proveedores agregados")
        ttk.Label(bar, textvariable=self.provider_info, foreground="#333", wraplength=850).pack(side="left", padx=12)

        grid = ttk.LabelFrame(self, text="Precios por proveedor", padding=6)
        grid.pack(fill="both", expand=True, padx=10, pady=6)
        self.tree_frame = grid
        self.tree = None
        self.rebuild_tree()

        btn = ttk.Frame(self, padding=10)
        btn.pack(fill="x")
        ttk.Button(btn, text="Visualizar requerimiento", command=self.view_req).pack(side="left", padx=4)
        ttk.Label(btn, text="Después de generar se abrirá opción de Word y se cerrará esta matriz.", foreground="#555").pack(side="left", padx=12)
        ttk.Button(btn, text="Generar OC", command=self.save).pack(side="right", padx=4)
        ttk.Button(btn, text="Cerrar", command=self.destroy).pack(side="right", padx=4)

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Agregar proveedor", command=self.add_provider)
        self.menu.add_command(label="Modificar proveedor", command=self.modify_provider)
        self.menu.add_command(label="Eliminar proveedor", command=self.remove_provider)
        self.menu.add_separator()
        self.menu.add_command(label="Limpiar precio de esta celda", command=self.clear_context_price)

    def update_provider_info(self):
        if not self.providers:
            self.provider_info.set("Sin proveedores agregados")
            return
        txt = []
        for i, p in enumerate(self.providers, start=1):
            tc = "TC 1.0000" if p["moneda"] == "SOLES" else f"TC {float(p.get('tipo_cambio') or 0):.4f}"
            txt.append(f"{i}. {p['razon']} [{p.get('tipo_orden','NACIONAL')} | {p['moneda']} | {p.get('igv_mode','SIN IGV')} | {tc}] Cot:{p['cotizacion'] or '-'}")
        self.provider_info.set("  |  ".join(txt))

    def rebuild_tree(self):
        for widget in self.tree_frame.winfo_children():
            widget.destroy()
        self.tree = None
        cols = ["item", "descripcion", "um", "cantidad"] + [f"p{i}" for i in range(len(self.providers))]
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show="headings", height=18)
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree_frame.rowconfigure(0, weight=1)
        self.tree_frame.columnconfigure(0, weight=1)
        base = [("item", "ITEM", 70), ("descripcion", "DESCRIPCIÓN", 420), ("um", "UND", 70), ("cantidad", "CANTIDAD", 95)]
        for c, t, w in base:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, minwidth=w, anchor="center" if c != "descripcion" else "w", stretch=False)
        for i, p in enumerate(self.providers):
            c = f"p{i}"
            titulo = f"{p['razon'][:20]}\n{p.get('tipo_orden','NACIONAL')} {p['moneda']}"
            self.tree.heading(c, text=titulo)
            self.tree.column(c, width=135, minwidth=120, anchor="center", stretch=False)
        self.tree.tag_configure("sin_precio", background="#f0f0f0")
        self.tree.tag_configure("con_precio", background="#dff3df")
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.on_right_click)
        self.refresh_tree()

    def refresh_tree(self):
        if not self.tree:
            return
        for x in self.tree.get_children():
            self.tree.delete(x)
        for idx, d in enumerate(self.det, start=1):
            vals = [idx, d["descripcion"], d["um"], f"{d['cantidad']:g}"]
            priced = False
            for p in self.providers:
                precio = p["precios"].get(d["codigo"], "")
                if precio not in ("", None, 0, 0.0):
                    priced = True
                    vals.append(f"{float(precio):.2f}")
                else:
                    vals.append("")
            self.tree.insert("", "end", values=vals, tags=("con_precio" if priced else "sin_precio",))

    def cell_provider_info(self, event):
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not col or col == "#0":
            return row_id, None
        try:
            col_idx = int(col.replace("#", "")) - 1
        except Exception:
            return row_id, None
        provider_idx = col_idx - 4
        if provider_idx < 0 or provider_idx >= len(self.providers):
            return row_id, None
        return row_id, provider_idx

    def on_right_click(self, event):
        row_id, provider_idx = self.cell_provider_info(event)
        self.context_row_id = row_id
        self.context_provider_idx = provider_idx
        self.menu.entryconfig("Modificar proveedor", state=("normal" if provider_idx is not None else "disabled"))
        self.menu.entryconfig("Eliminar proveedor", state=("normal" if provider_idx is not None else "disabled"))
        self.menu.entryconfig("Limpiar precio de esta celda", state=("normal" if row_id and provider_idx is not None else "disabled"))
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def on_double_click(self, event):
        row_id, provider_idx = self.cell_provider_info(event)
        if provider_idx is None:
            if not self.providers:
                self.add_provider()
            return
        if not row_id:
            return
        self.edit_cell(row_id, provider_idx)

    def edit_cell(self, row_id, provider_idx):
        if self.editor is not None:
            try:
                self.editor.destroy()
            except Exception:
                pass
            self.editor = None
        try:
            row_index = self.tree.index(row_id)
            item = self.det[row_index]
        except Exception:
            return
        column = f"#{provider_idx + 5}"
        bbox = self.tree.bbox(row_id, column)
        if not bbox:
            return
        x, y, w, h = bbox
        provider = self.providers[provider_idx]
        current = provider["precios"].get(item["codigo"], "")
        var = tk.StringVar(value="" if current in (None, "", 0, 0.0) else str(current))
        entry = ttk.Entry(self.tree, textvariable=var, justify="center")
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        entry.select_range(0, "end")
        self.editor = entry
        saved = {"ok": False}

        def save_and_close(event=None):
            if saved["ok"]:
                return
            saved["ok"] = True
            raw = str(var.get()).strip()
            try:
                if raw == "":
                    provider["precios"].pop(item["codigo"], None)
                else:
                    val = float(raw.replace(",", "."))
                    if val < 0:
                        raise ValueError("El precio no puede ser negativo.")
                    if val == 0:
                        provider["precios"].pop(item["codigo"], None)
                    else:
                        provider["precios"][item["codigo"]] = val
                entry.destroy()
                self.editor = None
                self.refresh_tree()
            except Exception as e:
                saved["ok"] = False
                messagebox.showerror("Precio", str(e), parent=self)
                entry.focus_set()

        def cancel(event=None):
            saved["ok"] = True
            entry.destroy()
            self.editor = None

        entry.bind("<Return>", save_and_close)
        entry.bind("<FocusOut>", save_and_close)
        entry.bind("<Escape>", cancel)

    def selected_provider_index(self):
        if self.context_provider_idx is not None:
            return self.context_provider_idx
        if not self.providers:
            return None
        labels = [f"{i+1}. {p['razon']} ({p['moneda']})" for i, p in enumerate(self.providers)]
        val = simpledialog.askinteger("Proveedor", "Número de proveedor:\n" + "\n".join(labels), minvalue=1, maxvalue=len(self.providers), parent=self)
        return None if val is None else val - 1

    def add_provider(self):
        ProviderListDialog(self.app, self.open_provider_config)

    def open_provider_config(self, ruc, existing_idx=None):
        ruc = only_digits(ruc)
        if existing_idx is None and any(p["ruc"] == ruc for p in self.providers):
            messagebox.showwarning("Proveedor", "Ese proveedor ya fue agregado.", parent=self)
            return
        prov = self.db.get_proveedor_by_ruc(ruc)
        if not prov:
            messagebox.showwarning("Proveedor", "Proveedor no encontrado o inactivo.", parent=self)
            return
        current = self.providers[existing_idx] if existing_idx is not None else {}
        win = tk.Toplevel(self)
        win.title("Proveedor de cotización")
        win.geometry("660x420+290+120")
        bind_as_child_window(win, self)
        cot = tk.StringVar(value=current.get("cotizacion", ""))
        tipo_orden = tk.StringVar(value=current.get("tipo_orden") or (self.req_row["tipo_orden"] if "tipo_orden" in self.req_row.keys() else "NACIONAL"))
        igv_mode = tk.StringVar(value=current.get("igv_mode") or self.igv.get() or "SIN IGV")
        moneda = tk.StringVar(value=current.get("moneda") or prov["moneda_preferida"] or "SOLES")
        tc = tk.StringVar(value=str(current.get("tipo_cambio") or (1 if moneda.get() == "SOLES" else "")))
        tc_fuente = tk.StringVar(value=current.get("tipo_cambio_fuente") or ("SOLES" if moneda.get() == "SOLES" else "MANUAL"))
        tc_fecha = tk.StringVar(value=current.get("tipo_cambio_fecha") or self.fecha_tc.get() or today_text())
        frm = ttk.Frame(win, padding=14); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Proveedor:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        ttk.Label(frm, text=f"{prov['ruc']} - {prov['razon_social']}", wraplength=420).grid(row=0, column=1, columnspan=3, sticky="w", pady=5)
        ttk.Label(frm, text="Cotización:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=cot, width=24).grid(row=1, column=1, sticky="w", pady=5)
        ttk.Label(frm, text="Tipo OC:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Combobox(frm, textvariable=tipo_orden, values=TIPOS_ORDEN_COMPRA, state="readonly", width=18).grid(row=2, column=1, sticky="w", pady=5)
        ttk.Label(frm, text="Precio:").grid(row=2, column=2, sticky="w", pady=5, padx=(10,0))
        ttk.Combobox(frm, textvariable=igv_mode, values=["SIN IGV", "CON IGV"], state="readonly", width=12).grid(row=2, column=3, sticky="w", pady=5)
        ttk.Label(frm, text="Moneda:").grid(row=3, column=0, sticky="w", pady=5)
        cmb = ttk.Combobox(frm, textvariable=moneda, values=MONEDAS, state="readonly", width=18)
        cmb.grid(row=3, column=1, sticky="w", pady=5)
        ttk.Label(frm, text="Fecha TC:").grid(row=4, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=tc_fecha, width=18).grid(row=4, column=1, sticky="w", pady=5)
        ttk.Label(frm, text="Tipo de cambio:").grid(row=5, column=0, sticky="w", pady=5)
        ent_tc = ttk.Entry(frm, textvariable=tc, width=18)
        ent_tc.grid(row=5, column=1, sticky="w", pady=5)
        ttk.Label(frm, text="Fuente TC:").grid(row=6, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=tc_fuente, width=24).grid(row=6, column=1, sticky="w", pady=5)

        def sync_tc_state(event=None):
            mon = normalize_text(moneda.get())
            if mon == "SOLES":
                tc.set("1.0000")
                tc_fuente.set("SOLES")
            elif tc.get() in ("", "1", "1.0", "1.0000"):
                tc.set("")
                if tc_fuente.get() == "SOLES":
                    tc_fuente.set("MANUAL")

        def cargar_tc():
            try:
                if normalize_text(moneda.get()) == "SOLES":
                    tc.set("1.0000"); tc_fuente.set("SOLES"); return
                info = self.db.obtener_tipo_cambio_para_oc(fecha_iso_from_text(tc_fecha.get()), "DOLARES", self.app.usuario_actual, permitir_api=True)
                tc.set(f"{float(info['venta']):.4f}")
                tc_fuente.set(info.get("fuente", "API"))
                tc_fecha.set(fecha_text_from_iso(info.get("fecha") or fecha_iso_from_text(tc_fecha.get())))
            except Exception as e:
                messagebox.showerror("Tipo de cambio", str(e), parent=win)

        cmb.bind("<<ComboboxSelected>>", sync_tc_state)
        ttk.Button(frm, text="Cargar TC", command=cargar_tc).grid(row=5, column=2, sticky="w", padx=8)
        ttk.Label(frm, text="Al guardar se precargará el último precio comprado por artículo si existe.", foreground="#555").grid(row=7, column=0, columnspan=4, sticky="w", pady=(8, 0))

        def save_provider():
            try:
                mon = normalize_text(moneda.get())
                if mon not in MONEDAS:
                    raise ValueError("Moneda inválida.")
                tc_val = 1.0 if mon == "SOLES" else float(str(tc.get()).replace(",", "."))
                if tc_val <= 0:
                    raise ValueError("El tipo de cambio debe ser mayor a cero.")
                precios_actuales = dict(current.get("precios", {}) if existing_idx is not None else {})
                # Sugerir último precio comprado por proveedor/artículo solo si la celda aún está vacía.
                for item in self.det:
                    if item["codigo"] not in precios_actuales:
                        last = self.db.ultimo_precio_compra(item["codigo"], prov["ruc"]) or self.db.ultimo_precio_compra(item["codigo"])
                        if last and float(last["precio"] or 0) > 0:
                            precios_actuales[item["codigo"]] = float(last["precio"] or 0)
                data = {
                    "ruc": prov["ruc"],
                    "razon": prov["razon_social"],
                    "tipo_orden": normalize_text(tipo_orden.get() or "NACIONAL"),
                    "igv_mode": normalize_text(igv_mode.get() or "SIN IGV"),
                    "moneda": mon,
                    "cotizacion": normalize_text(cot.get()),
                    "tipo_cambio": tc_val,
                    "tipo_cambio_fuente": normalize_text(tc_fuente.get() or ("SOLES" if mon == "SOLES" else "MANUAL")),
                    "tipo_cambio_fecha": fecha_text_from_iso(fecha_iso_from_text(tc_fecha.get())),
                    "precios": precios_actuales,
                }
                if existing_idx is None:
                    self.providers.append(data)
                else:
                    self.providers[existing_idx] = data
                self.context_provider_idx = None
                self.update_provider_info()
                self.rebuild_tree()
                win.destroy()
            except Exception as e:
                messagebox.showerror("Proveedor", str(e), parent=win)

        btn = ttk.Frame(win, padding=10); btn.pack(fill="x")
        ttk.Button(btn, text="Guardar proveedor", command=save_provider).pack(side="right", padx=4)
        ttk.Button(btn, text="Cancelar", command=win.destroy).pack(side="right", padx=4)
        sync_tc_state()

    def modify_provider(self):
        idx = self.selected_provider_index()
        if idx is None:
            return
        self.open_provider_config(self.providers[idx]["ruc"], existing_idx=idx)

    def remove_provider(self):
        idx = self.selected_provider_index()
        if idx is None:
            return
        if not messagebox.askyesno("Proveedor", f"¿Eliminar {self.providers[idx]['razon']} de la matriz?", parent=self):
            return
        self.providers.pop(idx)
        self.context_provider_idx = None
        self.update_provider_info()
        self.rebuild_tree()

    def clear_context_price(self):
        if not self.context_row_id or self.context_provider_idx is None:
            return
        try:
            row_index = self.tree.index(self.context_row_id)
            item = self.det[row_index]
            self.providers[self.context_provider_idx]["precios"].pop(item["codigo"], None)
            self.refresh_tree()
        except Exception:
            pass

    def view_req(self):
        RequerimientoDetailDialog(self, self.app, self.requerimiento)

    def save(self):
        if not self.app.require_action("crear_oc", self):
            return
        try:
            proveedores = []
            for p in self.providers:
                precios = {}
                for cod, precio in p["precios"].items():
                    try:
                        val = float(precio or 0)
                    except Exception:
                        raise ValueError(f"Precio inválido del proveedor {p['razon']} para artículo {cod}.")
                    if val > 0:
                        precios[cod] = val
                if precios:
                    proveedores.append({
                        "ruc": p["ruc"],
                        "moneda": p["moneda"],
                        "cotizacion": p["cotizacion"],
                        "tipo_cambio": p.get("tipo_cambio"),
                        "tipo_cambio_fuente": p.get("tipo_cambio_fuente"),
                        "tipo_cambio_fecha": p.get("tipo_cambio_fecha"),
                        "forma_pago": self.forma_pago.get(),
                        "tipo_orden": p.get("tipo_orden") or (self.req_row["tipo_orden"] if "tipo_orden" in self.req_row.keys() else "NACIONAL"),
                        "precio_incluye_igv": (p.get("igv_mode") == "CON IGV"),
                        "precios": precios,
                    })
            duplicados = []
            for item in self.det:
                cargados = [p["razon"] for p in self.providers if float(str(p["precios"].get(item["codigo"], 0) or 0).replace(",", ".")) > 0]
                if len(cargados) > 1:
                    duplicados.append(f"{item['descripcion']} -> " + ", ".join(cargados))
            if duplicados:
                msg = "Hay items con precio en más de un proveedor. Si continúa, se generarán varias OC para el mismo item.\n\n" + "\n".join(duplicados[:8])
                if len(duplicados) > 8:
                    msg += f"\n... y {len(duplicados)-8} más"
                if not messagebox.askyesno("Validar adjudicación", msg + "\n\n¿Desea continuar?", parent=self):
                    return
            nums = self.db.create_ocs_from_requerimiento_multiproveedor(self.requerimiento, proveedores, False, self.app.usuario_actual)
            msg = "Órdenes generadas:\n" + "\n".join(nums)
            messagebox.showinfo("OC generadas", msg, parent=self)
            if nums and messagebox.askyesno("Imprimir / guardar", "¿Desea abrir la vista del formato Word de las OC generadas?", parent=self):
                # Se usa la ventana principal como padre para que las vistas Word no se cierren
                # cuando se cierre la matriz de cotización.
                for n in nums:
                    OCPrintPreviewDialog(self.app.root, self.db, n)
            self.destroy()
        except Exception as e:
            messagebox.showerror("OC por requerimiento", str(e), parent=self)


class OCFinanzasSeguimientoDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.app=app; self.db=app.db
        self.title("Finanzas - Seguimiento de pagos OC"); self.geometry("1220x620+60+60"); bind_as_child_window(self,self.master if hasattr(self,"master") else None)
        self.term=tk.StringVar(); self.estado_pago=tk.StringVar(value="TODOS"); self.tipo=tk.StringVar(value="TODOS"); self.solo_saldo=tk.BooleanVar(value=False); self.orden=tk.StringVar(value="vencimiento_asc")
        self.build(); self.refresh()
    def build(self):
        top=ttk.Frame(self,padding=10); top.pack(fill="x")
        ttk.Label(top,text="Buscar:").pack(side="left"); ttk.Entry(top,textvariable=self.term,width=28).pack(side="left",padx=4)
        ttk.Label(top,text="Estado pago:").pack(side="left",padx=(10,2)); ttk.Combobox(top,textvariable=self.estado_pago,values=["TODOS"]+ESTADOS_PAGO_OC,state="readonly",width=16).pack(side="left")
        ttk.Label(top,text="Tipo:").pack(side="left",padx=(10,2)); ttk.Combobox(top,textvariable=self.tipo,values=["TODOS"]+TIPOS_ORDEN_COMPRA,state="readonly",width=12).pack(side="left")
        ttk.Checkbutton(top,text="Solo pendientes/saldo",variable=self.solo_saldo).pack(side="left",padx=10)
        ttk.Label(top,text="Orden:").pack(side="left",padx=(8,2)); ttk.Combobox(top,textvariable=self.orden,values=["vencimiento_asc","saldo_desc","fecha_desc"],state="readonly",width=16).pack(side="left")
        ttk.Button(top,text="Filtrar",command=self.refresh).pack(side="left",padx=6)
        ttk.Button(top,text="Exportar Excel",command=self.export).pack(side="right",padx=4)
        cols=("numero","fecha","proveedor","forma","vencimiento","total","pagado","saldo","estado_pago","estado","avance","tipo")
        box=ttk.Frame(self); box.pack(fill="both",expand=True,padx=10,pady=8); box.rowconfigure(0,weight=1); box.columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(box,columns=cols,show="headings",height=20)
        specs=[("numero","OC",135),("fecha","Fecha OC",85),("proveedor","Proveedor",260),("forma","Forma pago",145),("vencimiento","Vence",90),("total","Total",90),("pagado","Pagado",90),("saldo","Saldo",90),("estado_pago","Estado pago",115),("estado","Estado operativo",155),("avance","Avance",70),("tipo","Tipo",85)]
        for c,t,w in specs:
            self.tree.heading(c,text=t); self.tree.column(c,width=w,minwidth=w,anchor="center" if c not in ("proveedor",) else "w",stretch=False)
        vsb=ttk.Scrollbar(box,orient="vertical",command=self.tree.yview); hsb=ttk.Scrollbar(box,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Mapa de proceso", command=self.mapa_proceso)
        self.menu.add_command(label="Registrar pago", command=self.pay_selected)
        self.menu.add_command(label="Visualizar OC", command=self.view_selected)
        self.tree.bind("<Button-3>", self.show_menu)
        btn=ttk.Frame(self,padding=10); btn.pack(fill="x")
        ttk.Button(btn,text="Registrar pago",command=self.pay_selected).pack(side="left",padx=4)
        ttk.Button(btn,text="Visualizar OC",command=self.view_selected).pack(side="left",padx=4)
        ttk.Button(btn,text="Actualizar",command=self.refresh).pack(side="right",padx=4)
    def refresh(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        rows=self.db.list_oc_finanzas(self.term.get(), self.estado_pago.get(), self.tipo.get(), self.solo_saldo.get(), self.orden.get())
        for r in rows:
            self.tree.insert("","end",values=(r["numero"],fecha_text_from_iso(r["fecha_orden"]),r["razon_social"],r["forma_pago"],fecha_text_from_iso(r["fecha_vencimiento_pago"]) if r["fecha_vencimiento_pago"] else "",f"{r['total']:.2f}",f"{r['pagado']:.2f}",f"{r['saldo']:.2f}",r["estado_pago"],r["estado"],f"{r['avance']}%",r["tipo_orden"]))
    def selected_num(self):
        sel=self.tree.selection(); return self.tree.item(sel[0],"values")[0] if sel else None
    def show_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self.menu.tk_popup(event.x_root, event.y_root)
    def mapa_proceso(self):
        n=self.selected_num()
        if n: OCMapaProcesoDialog(self, self.db, n)
    def pay_selected(self):
        n=self.selected_num()
        dlg=OCPagoDialog(self.app, close_on_save=True, after_save=self.refresh)
        if n:
            dlg.numero.set(n); dlg.load()
    def view_selected(self):
        n=self.selected_num()
        if n: OCDetailDialog(self,self.db,n)
    def export(self):
        rows=self.db.list_oc_finanzas(self.term.get(), self.estado_pago.get(), self.tipo.get(), self.solo_saldo.get(), self.orden.get())
        data=[]
        for r in rows:
            data.append([r["numero"],fecha_text_from_iso(r["fecha_orden"]),r["ruc"],r["razon_social"],r["forma_pago"],fecha_text_from_iso(r["fecha_vencimiento_pago"]) if r["fecha_vencimiento_pago"] else "",r["moneda"],r["total"],r["pagado"],r["saldo"],r["estado_pago"],r["estado"],r["avance"],r["tipo_orden"],r["requerimiento"],r["almacen_codigo"]])
        self.app.export_excel(data,["Nro OC","Fecha OC","RUC","Proveedor","Forma pago","Vencimiento pago","Moneda","Total","Pagado","Saldo","Estado pago","Estado operativo","Avance operativo","Tipo","Requerimiento","Almacén"],["Seguimiento financiero de OC"],"seguimiento_pagos_oc.xlsx")


class ValorizacionMensualDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.title("Reporte de valorización mensual de almacén")
        self.geometry("1180x660+70+45")
        self.minsize(1050, 560)
        bind_as_child_window(self, app.root)
        self.rows = []
        self.build()

    def build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        today = date.today()
        self.anio_var = tk.StringVar(value=str(today.year))
        self.mes_var = tk.StringVar(value=f"{today.month:02d}")
        self.almacen_var = tk.StringVar(value="TODOS")
        self.moneda_reporte = tk.StringVar(value="SOLES")
        ttk.Label(top, text="Año:").grid(row=0, column=0, sticky="w", padx=4)
        ttk.Combobox(top, textvariable=self.anio_var, values=[str(y) for y in range(today.year - 5, today.year + 2)], width=8, state="readonly").grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(top, text="Mes:").grid(row=0, column=2, sticky="w", padx=4)
        ttk.Combobox(top, textvariable=self.mes_var, values=[f"{m:02d}" for m in range(1, 13)], width=6, state="readonly").grid(row=0, column=3, sticky="w", padx=4)
        ttk.Label(top, text="Almacén:").grid(row=0, column=4, sticky="w", padx=4)
        ttk.Combobox(top, textvariable=self.almacen_var, values=["TODOS", "AQP", "MIN"], width=10, state="readonly").grid(row=0, column=5, sticky="w", padx=4)
        ttk.Label(top, text="Moneda reporte:").grid(row=0, column=6, sticky="w", padx=4)
        ttk.Combobox(top, textvariable=self.moneda_reporte, values=["SOLES", "DOLARES"], width=9, state="readonly").grid(row=0, column=7, sticky="w", padx=4)
        ttk.Label(top, text="Método: Promedio ponderado móvil", foreground="#555").grid(row=1, column=0, columnspan=4, sticky="w", padx=4, pady=(6,0))
        ttk.Button(top, text="Vista previa", command=self.preview).grid(row=0, column=8, padx=5)
        ttk.Button(top, text="Exportar Excel", command=self.export).grid(row=0, column=9, padx=5)
        ttk.Button(top, text="Cerrar", command=self.destroy).grid(row=0, column=10, padx=5)

        note = ttk.Frame(self, padding=(10, 0, 10, 4))
        note.pack(fill="x")
        ttk.Label(note, text="Incluye artículos valorizables. Muestra costo promedio en soles y referencia en dólares si hay TC del mes; entradas sin IGV y con IGV, destino CC/OT.", foreground="#555").pack(anchor="w")

        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        self.cols = (
            "almacen_codigo", "codigo", "descripcion", "unidad_medida",
            "stock_inicial", "valor_inicial", "entradas", "valor_entradas",
            "salidas", "valor_salidas", "ajustes_entrada", "valor_ajustes_entrada",
            "ajustes_salida", "valor_ajustes_salida", "stock_final", "costo_promedio_final",
            "precio_prom_sin_igv_soles", "precio_prom_con_igv_soles", "precio_prom_sin_igv_dolares", "precio_prom_con_igv_dolares",
            "valor_final", "destino_centro_costo", "destino_ot", "observacion"
        )
        self.tree = ttk.Treeview(body, columns=self.cols, show="headings", height=18)
        headings = {
            "almacen_codigo": "Alm.", "codigo": "Código", "descripcion": "Descripción", "unidad_medida": "UM",
            "stock_inicial": "Stock inicial", "valor_inicial": "Valor inicial", "entradas": "Entradas", "valor_entradas": "Valor entradas",
            "salidas": "Salidas", "valor_salidas": "Valor salidas", "ajustes_entrada": "Aj. +", "valor_ajustes_entrada": "Valor Aj. +",
            "ajustes_salida": "Aj. -", "valor_ajustes_salida": "Valor Aj. -", "stock_final": "Stock final", "costo_promedio_final": "Costo prom. final",
            "precio_prom_sin_igv_soles": "P.S/ sIGV", "precio_prom_con_igv_soles": "P.S/ cIGV",
            "precio_prom_sin_igv_dolares": "P.US$ sIGV", "precio_prom_con_igv_dolares": "P.US$ cIGV",
            "valor_final": "Valor final", "destino_centro_costo": "Destino CC", "destino_ot": "Destino OT", "observacion": "Observación"
        }
        widths = {
            "almacen_codigo": 65, "codigo": 110, "descripcion": 280, "unidad_medida": 70,
            "stock_inicial": 105, "valor_inicial": 110, "entradas": 90, "valor_entradas": 115,
            "salidas": 90, "valor_salidas": 110, "ajustes_entrada": 80, "valor_ajustes_entrada": 105,
            "ajustes_salida": 80, "valor_ajustes_salida": 105, "stock_final": 95, "costo_promedio_final": 120,
            "precio_prom_sin_igv_soles": 105, "precio_prom_con_igv_soles": 105, "precio_prom_sin_igv_dolares": 105, "precio_prom_con_igv_dolares": 105,
            "valor_final": 110, "destino_centro_costo": 130, "destino_ot": 130, "observacion": 260
        }
        for c in self.cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], anchor="e" if c not in ("almacen_codigo", "codigo", "descripcion", "unidad_medida", "observacion") else ("w" if c in ("descripcion", "observacion") else "center"))
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(body, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.total_var = tk.StringVar(value="Sin datos cargados")
        ttk.Label(self, textvariable=self.total_var, padding=(10, 0, 10, 10), font=("Arial", 10, "bold")).pack(fill="x")
        self.preview()

    def _get_rows(self):
        return self.db.reporte_valorizacion_mensual(self.anio_var.get(), self.mes_var.get(), self.almacen_var.get())

    def preview(self):
        try:
            self.rows = self._get_rows()
            for item in self.tree.get_children():
                self.tree.delete(item)
            total_final = 0.0
            for r in self.rows:
                total_final += float(r.get("valor_final") or 0)
                values = []
                for c in self.cols:
                    v = r.get(c, "")
                    if isinstance(v, float):
                        if c in ("stock_inicial", "entradas", "salidas", "ajustes_entrada", "ajustes_salida", "stock_final"):
                            v = f"{v:,.4f}".rstrip("0").rstrip(".")
                        else:
                            v = f"{v:,.2f}"
                    values.append(v)
                self.tree.insert("", "end", values=values)
            if self.moneda_reporte.get() == "DOLARES":
                tc_ref = self._tc_referencia_mes(strict=True)
                total_usd = total_final / tc_ref if tc_ref else 0.0
                self.total_var.set(f"Filas: {len(self.rows)} | Valor final total: US$ {total_usd:,.2f} | TC ref. venta: {tc_ref:,.4f}")
            else:
                self.total_var.set(f"Filas: {len(self.rows)} | Valor final total: S/ {total_final:,.2f}")
        except Exception as e:
            messagebox.showerror("Valorización mensual", str(e), parent=self)

    def _tc_referencia_mes(self, strict=False):
        try:
            first, last = date(int(self.anio_var.get()), int(self.mes_var.get()), 1), None
            m = int(self.mes_var.get()); y = int(self.anio_var.get())
            last = (date(y + 1, 1, 1) if m == 12 else date(y, m + 1, 1)) - timedelta(days=1)
            rows = list(self.db.list_tipo_cambio(fecha_text_from_iso(first.strftime(ISO_FMT)), fecha_text_from_iso(last.strftime(ISO_FMT))))
            vals = [float(r["venta"] or 0) for r in rows if float(r["venta"] or 0) > 0]
            if vals:
                return sum(vals)/len(vals)
            if strict:
                raise ValueError(f"No existe tipo de cambio registrado para {self.mes_var.get()}/{self.anio_var.get()}. Registre el TC antes de emitir reportes valorizados en dólares.")
            return 1.0
        except ValueError:
            raise
        except Exception:
            if strict:
                raise ValueError("No se pudo validar el tipo de cambio del mes para el reporte en dólares.")
            return 1.0

    def export(self):
        try:
            if self.moneda_reporte.get() == "DOLARES":
                self._tc_referencia_mes(strict=True)
            rows = self._get_rows()
            headers = [
                "periodo", "almacen_codigo", "almacen_nombre", "codigo", "descripcion", "unidad_medida",
                "stock_inicial", "valor_inicial", "entradas", "valor_entradas", "salidas", "valor_salidas",
                "ajustes_entrada", "valor_ajustes_entrada", "ajustes_salida", "valor_ajustes_salida",
                "stock_final", "costo_promedio_final", "precio_prom_sin_igv_soles", "precio_prom_con_igv_soles", "precio_prom_sin_igv_dolares", "precio_prom_con_igv_dolares", "valor_final", "destino_centro_costo", "destino_ot", "metodo", "observacion"
            ]
            titles = [
                "Periodo", "Almacén", "Almacén nombre", "Código", "Descripción", "UM",
                "Stock inicial", "Valor inicial", "Entradas", "Valor entradas", "Salidas", "Valor salidas",
                "Ajustes entrada", "Valor ajustes entrada", "Ajustes salida", "Valor ajustes salida",
                "Stock final", "Costo promedio final", "Precio soles sin IGV", "Precio soles con IGV", "Precio dólares sin IGV", "Precio dólares con IGV", "Valor final", "Destino centro costo", "Destino OT", "Método", "Observación"
            ]
            filename = f"reporte_valorizacion_almacen_{self.anio_var.get()}{self.mes_var.get()}_{self.almacen_var.get().lower()}_{self.moneda_reporte.get().lower()}.xlsx"
            self.app.export_excel(rows, headers, titles, filename)
        except Exception as e:
            messagebox.showerror("Exportar valorización", str(e), parent=self)


class OCReportePendientesDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.app=app; self.db=app.db
        self.title("Reporte de Órdenes Pendientes"); self.geometry("520x230+300+170"); bind_as_child_window(self,self.master if hasattr(self,"master") else None)
        self.tipo=tk.StringVar(value="TODOS"); self.estado=tk.StringVar(value="PENDIENTE"); self.estado_pago=tk.StringVar(value="TODOS")
        frm=ttk.Frame(self,padding=20); frm.pack(fill="both",expand=True)
        ttk.Label(frm,text="Tipo:").grid(row=0,column=0,sticky="w",pady=8); ttk.Combobox(frm,textvariable=self.tipo,values=["TODOS"]+TIPOS_ORDEN_COMPRA,state="readonly",width=22).grid(row=0,column=1,sticky="w")
        ttk.Label(frm,text="Estado operativo:").grid(row=1,column=0,sticky="w",pady=8); ttk.Combobox(frm,textvariable=self.estado,values=["TODOS"]+ESTADOS_OC,state="readonly",width=22).grid(row=1,column=1,sticky="w")
        ttk.Label(frm,text="Estado pago:").grid(row=2,column=0,sticky="w",pady=8); ttk.Combobox(frm,textvariable=self.estado_pago,values=["TODOS"]+ESTADOS_PAGO_OC,state="readonly",width=22).grid(row=2,column=1,sticky="w")
        ttk.Button(frm,text="Exportar Excel",command=self.export).grid(row=4,column=1,sticky="e",pady=18)
    def export(self):
        rows=self.db.reporte_oc_rows(self.tipo.get(),self.estado.get())
        data=[]
        for r in rows:
            if self.estado_pago.get() != "TODOS" and r["estado_pago"] != self.estado_pago.get():
                continue
            estado_pago, pagado, saldo = self.db._sync_oc_payment_status(r["id"], commit=True)
            req_fecha = ""; req_entrega = ""
            if r["requerimiento"]:
                req_row, _ = self.db.find_requerimiento(r["requerimiento"])
                if req_row:
                    req_fecha = fecha_text_from_iso(req_row["fecha_requerimiento"])
                    req_entrega = fecha_text_from_iso(req_row["fecha_entrega_solicitada"]) if "fecha_entrega_solicitada" in req_row.keys() and req_row["fecha_entrega_solicitada"] else ""
            fecha_entrada = fecha_text_from_iso(self.db.fecha_registro_entrada_oc(r["id"]))
            data.append([r["numero"],fecha_text_from_iso(r["fecha_orden"]),r["ruc"],r["razon_social"],r["tipo_orden"],r["moneda"],r["subtotal"],r["igv"],r["total"],r["forma_pago"],fecha_text_from_iso(r["fecha_vencimiento_pago"]) if r["fecha_vencimiento_pago"] else "",pagado,saldo,estado_pago,r["estado"],r["avance"],r["requerimiento"],req_fecha,req_entrega,fecha_entrada,r["centro_costo_codigo"],r["ot"],r["almacen_codigo"],fecha_text_from_iso(r["fecha_entrega"]) if r["fecha_entrega"] else ""])
        headers=["nro_oc","fecha_oc","ruc","proveedor","tipo","moneda","subtotal","igv","total","forma_pago","vence_pago","pagado","saldo","estado_pago","estado_operativo","avance_operativo","requerimiento","fecha_solicitud_req","fecha_entrega_solicitada_req","fecha_registro_entrada","cc","ot","almacen","fecha_entrega_oc"]
        titles=["Nro OC","Fecha OC","RUC","Proveedor","Tipo","Moneda","Subtotal","IGV","Total","Forma pago","Vence pago","Pagado","Saldo","Estado pago","Estado operativo","Avance operativo","Requerimiento","Fecha solicitud requerimiento","Fecha entrega solicitada requerimiento","Fecha registro entrada","CC","OT","Almacén","Fecha entrega OC"]
        self.app.export_excel(data,headers,titles,"reporte_ordenes_compra.xlsx")


class OCPrintPreviewDialog(tk.Toplevel):
    def __init__(self, parent, db, numero):
        super().__init__(parent); self.db=db; self.numero=numero; self.generated_path=None
        self.title("Formato Orden de Compra / Servicio"); self.geometry("720x620+170+60"); bind_as_child_window(self,parent); self.build()
    def build(self):
        oc,det=self.db.get_orden_compra(self.numero)
        main=ttk.Frame(self,padding=15); main.pack(fill="both",expand=True)
        if not oc: ttk.Label(main,text="OC no encontrada").pack(); return
        ttk.Label(main,text=("ORDEN DE SERVICIO" if oc["tipo_orden"]=="SERVICIO" else "ORDEN DE COMPRA"),font=("Arial",16,"bold")).pack(pady=4)
        ttk.Label(main,text="Vista previa del contenido que se enviará al formato Word editable.",foreground="#666").pack(pady=(0,4))
        estado_pago, pagado, saldo = self.db._sync_oc_payment_status(oc["id"], commit=True)
        ttk.Label(main,text=f"NRO: {oc['numero']}   Estado operativo: {oc['estado']}   Avance: {oc['avance']}%   Pago: {estado_pago}   Saldo: {fmt_money_currency(saldo, oc['moneda'])}   Total: {fmt_money_currency(oc['total'], oc['moneda'])}   TC: {oc['tipo_cambio_venta']:.4f}",font=("Arial",11,"bold")).pack(pady=4)
        txt=tk.Text(main,height=8,wrap="word"); txt.pack(fill="x",pady=6)
        txt.insert("end",f"Proveedor: {oc['razon_social']}\nRUC: {oc['ruc']}\nDirección: {oc['direccion']}\nSolicitante: {oc['solicitante']} | CC: {oc['centro_costo_codigo']} {oc['centro_costo_nombre']} | OT: {oc['ot']}\nObservaciones: {oc['observacion']}\n")
        txt.configure(state="disabled")
        ttk.Label(main, text=f"Detalle de ítems ({len(det)} registro(s))", font=("Arial",10,"bold")).pack(anchor="w", pady=(4,0))
        cols=("item","descripcion","um","cantidad","punit","total")
        tree=ttk.Treeview(main,columns=cols,show="headings",height=10)
        for c,t,w in [("item","Item",50),("descripcion","Descripción",330),("um","UM",70),("cantidad","Cant.",70),("punit","P.Unit",90),("total","P.Total",90)]:
            tree.heading(c,text=t); tree.column(c,width=w,anchor="center" if c!="descripcion" else "w")
        tree_box=ttk.Frame(main); tree_box.pack(fill="both",expand=True,pady=6)
        tree_box.rowconfigure(0,weight=1); tree_box.columnconfigure(0,weight=1)
        vsb=ttk.Scrollbar(tree_box,orient="vertical",command=tree.yview); hsb=ttk.Scrollbar(tree_box,orient="horizontal",command=tree.xview)
        tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        tree.grid(in_=tree_box,row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        if det:
            for d in det:
                tree.insert("","end",values=(d["item"],d["descripcion"],d["unidad_medida"],f"{d['cantidad_solicitada']:g}",fmt_money_currency(d['precio_unitario_sin_igv'], oc['moneda']),fmt_money_currency(d['subtotal'], oc['moneda'])))
        else:
            tree.insert("","end",values=("-","Sin ítems registrados","","","",""))
        btn=ttk.Frame(main); btn.pack(fill="x",pady=8)
        ttk.Button(btn,text="Guardar Word",command=self.save_word_as).pack(side="right",padx=4)
        ttk.Button(btn,text="Imprimir Word",command=self.print_document).pack(side="right",padx=4)
        ttk.Button(btn,text="Cerrar",command=self.destroy).pack(side="right",padx=4)
    def default_path(self):
        # Compatibilidad: el formato vigente para OC/OS es Word, no Excel.
        return self.default_word_path()
    def build_workbook(self):
        if Workbook is None:
            raise ValueError("Instale openpyxl: pip install openpyxl")
        oc,det=self.db.get_orden_compra(self.numero)
        if not oc:
            raise ValueError("No se encontró la OC.")
        wb=Workbook(); ws=wb.active; ws.title="Orden"
        ws.page_setup.orientation="portrait"; ws.page_setup.paperSize=9; ws.page_setup.fitToWidth=1; ws.sheet_properties.pageSetUpPr.fitToPage=True
        thin=Side(style="thin",color="888888"); border=Border(top=thin,bottom=thin,left=thin,right=thin); fill=PatternFill("solid",fgColor="EDEDED")

        def put(row, col, value="", bold=False, size=None, align=None, fill_cell=None, border_cell=None, wrap=False):
            cell = ws.cell(row=row, column=col)
            cell.value = value
            if bold or size:
                cell.font = Font(bold=bold, size=size or 11)
            if align:
                cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
            elif wrap:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            if fill_cell:
                cell.fill = fill_cell
            if border_cell:
                cell.border = border_cell
            return cell

        ws.merge_cells("A1:F1"); put(1,1,"MEDSURSA - MINERA ESPAÑOLITA DEL SUR S.A.",bold=True,size=14,align="center")
        ws.merge_cells("A2:F2"); put(2,1,"PANAMERICANA SUR KM 616 QUEBRADA HUANCA (PLANTA LA ENCAÑADA)",align="center")
        ws.merge_cells("A4:F4"); put(4,1,"ORDEN DE SERVICIO" if oc["tipo_orden"]=="SERVICIO" else "ORDEN DE COMPRA",bold=True,size=16,align="center")
        put(5,5,"NRO:",bold=True); put(5,6,oc["numero"])
        rows=[("DATOS DE PROVEEDOR",""),("PROVEEDOR",oc["razon_social"]),("DIRECCIÓN",oc["direccion"]),("RUC",oc["ruc"]),("TELÉFONO",oc["telefono"]),("CONDICIONES DE VENTA","") ,("FORMA DE PAGO",oc["forma_pago"]),("VENCIMIENTO DE PAGO",fecha_text_from_iso(oc["fecha_vencimiento_pago"]) if oc["fecha_vencimiento_pago"] else ""),("LUGAR DE ENTREGA",oc["lugar_entrega"]),("MONEDA",oc["moneda"]),("TIPO DE CAMBIO",f"{oc['tipo_cambio_venta']:.4f} {oc['tipo_cambio_fuente']}"),("FECHA DE ENTREGA",oc["fecha_entrega"]),("REFERENCIAS","") ,("SOLICITANTE",oc["solicitante"]),("CENTRO DE COSTO",f"{oc['centro_costo_codigo']} {oc['centro_costo_nombre']}"),("OT",oc["ot"]),("OBSERVACIONES",oc["observacion"])]
        r=7
        for k,v in rows:
            ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
            ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6)
            put(r,1,k,bold=True,fill_cell=fill if not v else None)
            put(r,3,v,wrap=True)
            for c in range(1,7): ws.cell(r,c).border=border
            r+=1
        headers=["ITEM","DESCRIPCIÓN","U.MED.","CANTIDAD","P.UNIT.","P.TOTAL"]
        for c,h in enumerate(headers,1):
            put(r,c,h,bold=True,align="center",fill_cell=fill,border_cell=border)
        r+=1
        for d in det:
            vals=[d["item"],d["descripcion"],d["unidad_medida"],d["cantidad_solicitada"],d["precio_unitario_sin_igv"],d["subtotal"]]
            for c,v in enumerate(vals,1):
                put(r,c,v,align="center" if c!=2 else "left",border_cell=border,wrap=(c==2))
            r+=1
        r+=1
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=4)
        put(r,1,money_to_words_es(oc["total"], oc["moneda"]),wrap=True)
        put(r,5,"VALOR VENTA"); put(r,6,oc["subtotal"]); r+=1
        put(r,5,"I.G.V. 18%"); put(r,6,oc["igv"]); r+=1
        put(r,5,"TOTAL",bold=True); put(r,6,oc["total"],bold=True); r+=3
        for c,title in [(1,"ELABORADO POR"),(3,"APROBADO POR"),(5,"ACEPTADO PROVEEDOR")]:
            ws.merge_cells(start_row=r,start_column=c,end_row=r,end_column=c+1)
            put(r,c,title,bold=True,align="center")
            ws.merge_cells(start_row=r+1,start_column=c,end_row=r+4,end_column=c+1)
            nombre = oc['usuario_creador'] if c==1 else oc['aprobado_por'] if c==3 else ''
            # Importante: en rangos combinados solo se escribe en la celda superior izquierda.
            put(r+1,c,f"\n\nNOMBRE: {nombre}\nFECHA:",align="left",wrap=True)
            for rr in range(r, r+5):
                for cc in range(c, c+2):
                    ws.cell(rr,cc).border=border
        widths={"A":8,"B":42,"C":10,"D":12,"E":14,"F":14}
        for col,w in widths.items(): ws.column_dimensions[col].width=w
        ws.print_area=f"A1:F{r+5}"
        return wb
    def save_file(self,path):
        wb=self.build_workbook(); wb.save(path); self.generated_path=path; return path

    def default_word_path(self):
        os.makedirs(OC_DIR, exist_ok=True)
        return os.path.join(OC_DIR, f"{self.numero}.docx")

    def build_word_document(self):
        Doc = require_docx_document(parent=self)
        if not os.path.exists(OC_WORD_TEMPLATE):
            raise ValueError(f"No se encontró la plantilla Word:\n{OC_WORD_TEMPLATE}")
        oc, det = self.db.get_orden_compra(self.numero)
        if not oc:
            raise ValueError("No se encontró la OC.")
        doc = Doc(OC_WORD_TEMPLATE)
        prov = self.db.get_proveedor_by_ruc(oc["ruc"], active_only=False)
        tipo_doc = "ORDEN DE SERVICIO" if oc["tipo_orden"] == "SERVICIO" else "ORDEN DE COMPRA"
        destino = oc["lugar_entrega"] or f"{oc['almacen_codigo']} - {oc['almacen_nombre']}"
        mapping = {
            "{{TIPO_DOC}}": tipo_doc,
            "{{OC}}": oc["numero"],
            "{{FEC_LARGA}}": fecha_larga_es(oc["fecha_orden"]),
            "{{FECHA}}": fecha_text_from_iso(oc["fecha_orden"]),
            "{{PROVEEDOR}}": oc["razon_social"],
            "{{DIR_PROV}}": oc["direccion"],
            "{{COTIZ}}": oc["cotizacion"] or "",
            "{{RUC_PROV}}": oc["ruc"],
            "{{TEL_PROV}}": oc["telefono"] or "",
            "{{DESTINO}}": destino,
            "{{TERM_PAGO}}": oc["forma_pago"],
            "{{MONEDA}}": oc["moneda"],
            "{{BANCO}}": prov["banco"] if prov and "banco" in prov.keys() else "",
            "{{CTA}}": prov["numero_cuenta"] if prov and "numero_cuenta" in prov.keys() else "",
            "{{CCI}}": prov["cci"] if prov and "cci" in prov.keys() else "",
            "{{SUBTOTAL}}": fmt_money_currency(oc["subtotal"], oc["moneda"]),
            "{{IGV}}": fmt_money_currency(oc["igv"], oc["moneda"]),
            "{{TOTAL}}": fmt_money_currency(oc["total"], oc["moneda"]),
            "{{ELAB}}": oc["usuario_creador"] or "",
            "{{SOLIC}}": oc["solicitante"] or "",
            "{{AUTORIZ}}": oc["aprobado_por"] or "",
        }
        # La plantilla trae 16 líneas base. Si la OC/OS tiene más ítems, se insertan
        # filas adicionales DENTRO de la misma tabla. No se duplica la cabecera general
        # ni el bloque final de totales/firmas: Word continúa el documento naturalmente.
        items_table = _find_docx_table(doc, ["CANT", "UNID", "DESCRIPCION", "VALOR UNIT.", "IMPORTE"])
        if items_table is not None and len(det) > OC_BASE_ITEM_ROWS:
            total_anchor_index = None
            for idx, row in enumerate(items_table.rows):
                text = " ".join(normalize_text(cell.text).upper() for cell in row.cells)
                if "SUB TOTAL" in text:
                    total_anchor_index = idx
                    break
            if total_anchor_index is not None and total_anchor_index > 1:
                template_tr = deepcopy(items_table.rows[total_anchor_index - 1]._tr)
                anchor_tr = items_table.rows[total_anchor_index]._tr
                for extra_index, d in enumerate(det[OC_BASE_ITEM_ROWS:], start=OC_BASE_ITEM_ROWS + 1):
                    new_tr = deepcopy(template_tr)
                    anchor_tr.addprevious(new_tr)
                    # La fila recién insertada queda inmediatamente antes del bloque de totales.
                    inserted_row = items_table.rows[extra_index]
                    cantidad = float(d["cantidad_aprobada"] or 0) if float(d["cantidad_aprobada"] or 0) > 0 else float(d["cantidad_solicitada"] or 0)
                    values = [
                        f"{cantidad:g}",
                        d["unidad_medida"],
                        d["descripcion"],
                        fmt_money_currency(d["precio_unitario_sin_igv"], oc["moneda"]),
                        fmt_money_currency(d["subtotal"], oc["moneda"]),
                    ]
                    for cell, value in zip(inserted_row.cells, values):
                        _docx_set_cell_text(cell, value)

        # Se llenan las 16 líneas originales y se limpian las restantes cuando no se usan.
        for i in range(1, OC_BASE_ITEM_ROWS + 1):
            if i <= len(det):
                d = det[i-1]
                cantidad = float(d["cantidad_aprobada"] or 0) if float(d["cantidad_aprobada"] or 0) > 0 else float(d["cantidad_solicitada"] or 0)
                mapping[f"{{{{C{i:02d}}}}}"] = f"{cantidad:g}"
                mapping[f"{{{{U{i:02d}}}}}"] = d["unidad_medida"]
                mapping[f"{{{{DESC{i:02d}}}}}"] = d["descripcion"]
                mapping[f"{{{{PU{i:02d}}}}}"] = fmt_money_currency(d["precio_unitario_sin_igv"], oc["moneda"])
                mapping[f"{{{{IMP{i:02d}}}}}"] = fmt_money_currency(d["subtotal"], oc["moneda"])
            else:
                mapping[f"{{{{C{i:02d}}}}}"] = ""
                mapping[f"{{{{U{i:02d}}}}}"] = ""
                mapping[f"{{{{DESC{i:02d}}}}}"] = ""
                mapping[f"{{{{PU{i:02d}}}}}"] = ""
                mapping[f"{{{{IMP{i:02d}}}}}"] = ""
        replace_docx_placeholders(doc, mapping)
        # Quitar el párrafo vacío final de la plantilla evita páginas en blanco cuando
        # el detalle crece hasta el límite físico de la hoja.
        _docx_remove_trailing_empty_body_paragraph(doc)
        return doc

    def save_word_file(self, path):
        doc = self.build_word_document()
        doc.save(path)
        return path

    def save_word_as(self):
        path = filedialog.asksaveasfilename(parent=self, title="Guardar OC en Word", defaultextension=".docx", initialfile=f"{self.numero}.docx", filetypes=[("Word", "*.docx")])
        if path:
            try:
                self.save_word_file(path)
                messagebox.showinfo("OC", f"Documento Word guardado:\n{path}", parent=self)
            except Exception as e:
                messagebox.showerror("OC Word", str(e), parent=self)

    def save_as(self):
        # Alias de seguridad: ya no se generan órdenes de compra en Excel.
        return self.save_word_as()
    def print_document(self):
        try:
            path = self.default_word_path()
            self.save_word_file(path)
            self.generated_path = path
            if os.name=="nt":
                os.startfile(path, "print")
                messagebox.showinfo("Impresión", "Formato Word enviado a impresora.", parent=self)
            else:
                messagebox.showinfo("Impresión", f"Documento Word generado. Ábralo para imprimir:\n{path}", parent=self)
        except Exception as e:
            messagebox.showerror("Impresión", str(e), parent=self)



class TipoCambioDialog(tk.Toplevel):
    """Registro diario de tipo de cambio con grilla mensual y edición rápida."""
    def __init__(self, app, fecha_inicial=None, on_select=None):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.on_select = on_select
        self.title("Tipo de Cambio Diario")
        self.geometry("980x700+120+45")
        self.minsize(940, 650)
        bind_as_child_window(self, self.master if hasattr(self, "master") else None)
        base_fecha = fecha_inicial or today_text()
        self.fecha = tk.StringVar(value=base_fecha)
        try:
            f = parse_fecha(base_fecha)
        except Exception:
            f = date.today()
        self.anio = tk.StringVar(value=str(f.year))
        self.mes = tk.StringVar(value=f"{f.month:02d}")
        self.fuente = tk.StringVar(value="MANUAL")
        self.compra = tk.StringVar(value="")
        self.venta = tk.StringVar(value="")
        self.motivo = tk.StringVar(value="REGISTRO MANUAL")
        self.build()
        self.refresh_month()

    def build(self):
        top = ttk.LabelFrame(self, text="Periodo y fuente", padding=10)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text="Fecha:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.fecha_entry = ttk.Entry(top, textvariable=self.fecha, width=14)
        self.fecha_entry.grid(row=0, column=1, sticky="w", padx=4, pady=4)
        self.fecha_entry.bind("<Return>", lambda e: self.set_period_from_fecha())

        ttk.Label(top, text="Fuente:").grid(row=0, column=2, sticky="w", padx=12, pady=4)
        ttk.Combobox(top, textvariable=self.fuente, values=FUENTES_TC, state="readonly", width=18).grid(row=0, column=3, sticky="w", padx=4, pady=4)

        ttk.Label(top, text="Año:").grid(row=0, column=4, sticky="w", padx=12, pady=4)
        ttk.Entry(top, textvariable=self.anio, width=8).grid(row=0, column=5, sticky="w", padx=4, pady=4)
        ttk.Label(top, text="Mes:").grid(row=0, column=6, sticky="w", padx=8, pady=4)
        ttk.Combobox(top, textvariable=self.mes, values=[f"{i:02d}" for i in range(1,13)], state="readonly", width=5).grid(row=0, column=7, sticky="w", padx=4, pady=4)

        ttk.Button(top, text="Ver mes", command=self.refresh_month).grid(row=0, column=8, padx=8)
        ttk.Button(top, text="Cargar día desde API", command=self.cargar_api).grid(row=1, column=1, padx=4, pady=8, sticky="w")
        ttk.Button(top, text="Usar seleccionado", command=self.choose).grid(row=1, column=3, padx=4, pady=8, sticky="w")
        ttk.Label(top, text="La API llena compra/venta del día seleccionado. Si falla, registre manualmente abajo.", foreground="#555").grid(row=1, column=4, columnspan=5, sticky="w", padx=4)

        cols = ("dia", "fecha", "compra", "venta", "fuente", "manual", "usuario", "registro")
        box = ttk.Frame(self)
        box.pack(fill="both", expand=True, padx=10, pady=8)
        box.rowconfigure(0, weight=1); box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(box, columns=cols, show="headings", height=18)
        specs = [
            ("dia", "Día", 55), ("fecha", "Fecha", 90), ("compra", "Compra", 95), ("venta", "Venta", 95),
            ("fuente", "Fuente", 110), ("manual", "Manual", 75), ("usuario", "Usuario", 115), ("registro", "Registro", 160)
        ]
        for c, t, w in specs:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, minwidth=w, anchor="center", stretch=False)
        vsb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.load_selected)
        self.tree.bind("<Double-1>", self.choose)
        self.tree.bind("<Return>", self.choose)

        bottom = ttk.LabelFrame(self, text="Registro manual del día seleccionado", padding=10)
        bottom.pack(fill="x", padx=10, pady=8)

        ttk.Label(bottom, text="Compra:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.compra_entry = ttk.Entry(bottom, textvariable=self.compra, width=16)
        self.compra_entry.grid(row=0, column=1, sticky="w", padx=4, pady=4)
        self.compra_entry.bind("<Return>", lambda e: (self.venta_entry.focus_set(), "break")[1])
        ttk.Label(bottom, text="Venta:").grid(row=0, column=2, sticky="w", padx=12, pady=4)
        self.venta_entry = ttk.Entry(bottom, textvariable=self.venta, width=16)
        self.venta_entry.grid(row=0, column=3, sticky="w", padx=4, pady=4)
        self.venta_entry.bind("<Return>", lambda e: (self.motivo_entry.focus_set(), "break")[1])
        ttk.Label(bottom, text="Motivo:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.motivo_entry = ttk.Entry(bottom, textvariable=self.motivo, width=54)
        self.motivo_entry.grid(row=1, column=1, columnspan=4, sticky="we", padx=4, pady=4)
        self.motivo_entry.bind("<Return>", lambda e: (self.guardar_manual(), "break")[1])
        bottom.columnconfigure(4, weight=1)
        ttk.Button(bottom, text="Guardar manual", command=self.guardar_manual).grid(row=1, column=5, padx=8, pady=4, sticky="e")
        ttk.Button(bottom, text="Cerrar", command=self.destroy).grid(row=1, column=6, padx=4, pady=4, sticky="e")
        ttk.Label(bottom, text="Flujo: Compra + Enter → Venta + Enter → Motivo + Enter → Guardar", foreground="#555").grid(row=2, column=0, columnspan=7, sticky="w", padx=4, pady=(4,0))

    def set_period_from_fecha(self):
        try:
            f = parse_fecha(self.fecha.get())
            self.anio.set(str(f.year))
            self.mes.set(f"{f.month:02d}")
            self.refresh_month(select_fecha=f.strftime(ISO_FMT))
        except Exception as e:
            messagebox.showerror("Fecha", str(e), parent=self)
        return "break"

    def _month_range(self):
        try:
            y = int(self.anio.get())
            m = int(self.mes.get())
            first = date(y, m, 1)
            if m == 12:
                nxt = date(y + 1, 1, 1)
            else:
                nxt = date(y, m + 1, 1)
            last = nxt - timedelta(days=1)
            return first, last
        except Exception:
            raise ValueError("Ingrese año y mes válidos.")

    def refresh_month(self, select_fecha=None):
        for x in self.tree.get_children():
            self.tree.delete(x)
        try:
            first, last = self._month_range()
            rows = {r["fecha"]: r for r in self.db.list_tipo_cambio(fecha_text_from_iso(first.strftime(ISO_FMT)), fecha_text_from_iso(last.strftime(ISO_FMT)))}
            selected = None
            d = first
            while d <= last:
                iso = d.strftime(ISO_FMT)
                r = rows.get(iso)
                if r:
                    iid = self.tree.insert("", "end", values=(
                        f"{d.day:02d}", fecha_text_from_iso(iso), f"{r['compra']:.4f}", f"{r['venta']:.4f}",
                        r["fuente"], "SI" if r["es_manual"] else "NO", r["usuario_registro"], r["fecha_registro"]
                    ))
                else:
                    iid = self.tree.insert("", "end", values=(f"{d.day:02d}", fecha_text_from_iso(iso), "", "", "", "", "", ""))
                if select_fecha and iso == select_fecha:
                    selected = iid
                d += timedelta(days=1)
            if selected:
                self.tree.selection_set(selected)
                self.tree.focus(selected)
                self.tree.see(selected)
                self.load_selected()
            elif self.tree.get_children():
                first_iid = self.tree.get_children()[0]
                self.tree.selection_set(first_iid)
                self.tree.focus(first_iid)
                self.load_selected()
        except Exception as e:
            messagebox.showerror("Tipo de cambio", str(e), parent=self)

    def load_selected(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return "break"
        v = self.tree.item(sel[0], "values")
        self.fecha.set(v[1])
        self.compra.set(v[2] or "")
        self.venta.set(v[3] or "")
        if v[4]:
            self.fuente.set(v[4])
        return "break"

    def cargar_api(self):
        if not self.app.require_finanzas(self):
            return
        try:
            row = self.db.consultar_tipo_cambio_api(self.fecha.get(), self.app.usuario_actual)
            self.compra.set(f"{row['compra']:.4f}")
            self.venta.set(f"{row['venta']:.4f}")
            self.fuente.set(row["fuente"])
            self.set_period_from_fecha()
            messagebox.showinfo("Tipo de cambio", "Tipo de cambio cargado desde API.", parent=self)
        except Exception as e:
            messagebox.showerror("Tipo de cambio", str(e), parent=self)

    def guardar_manual(self):
        if not self.app.require_finanzas(self):
            return
        try:
            self.db.registrar_tipo_cambio_manual(self.fecha.get(), self.compra.get(), self.venta.get(), self.fuente.get(), self.motivo.get(), self.app.usuario_actual)
            self.set_period_from_fecha()
            messagebox.showinfo("Tipo de cambio", "Tipo de cambio guardado correctamente.", parent=self)
        except Exception as e:
            messagebox.showerror("Tipo de cambio", str(e), parent=self)

    def choose(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return "break"
        v = self.tree.item(sel[0], "values")
        if not v[2] or not v[3]:
            if self.on_select:
                messagebox.showinfo("Tipo de cambio", "Ese día no tiene tipo de cambio registrado.", parent=self)
            return "break"
        if self.on_select:
            fecha_iso = fecha_iso_consulta_from_text(v[1])
            self.on_select({"fecha": fecha_iso, "compra": float(v[2]), "venta": float(v[3]), "fuente": v[4] or self.fuente.get()})
            self.destroy()
        return "break"


class APIConfigDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app=app; self.db=app.db
        self.title("Configuración de APIs")
        self.geometry("860x420+180+100")
        bind_as_child_window(self, self.master if hasattr(self,"master") else None)
        cfg=self.db.api_config()
        self.ruc_url=tk.StringVar(value=cfg.get("ruc_api_url","")); self.ruc_token=tk.StringVar(value=cfg.get("ruc_api_token",""))
        self.tc_url=tk.StringVar(value=cfg.get("tc_api_url","")); self.tc_token=tk.StringVar(value=cfg.get("tc_api_token","")); self.timeout=tk.StringVar(value=cfg.get("api_timeout","10"))
        self.build()

    def build(self):
        frm=ttk.Frame(self,padding=16); frm.pack(fill="both",expand=True)
        ttk.Label(frm,text="Las URLs pueden usar {ruc} o {fecha}. Si el proveedor de API exige token, colóquelo aquí.",foreground="#555555").grid(row=0,column=0,columnspan=3,sticky="w",pady=(0,10))
        data=[("URL API RUC",self.ruc_url,1),("Token API RUC",self.ruc_token,2),("URL API Tipo Cambio",self.tc_url,3),("Token API Tipo Cambio",self.tc_token,4),("Timeout segundos",self.timeout,5)]
        for lab,var,row in data:
            ttk.Label(frm,text=lab+":").grid(row=row,column=0,sticky="w",padx=4,pady=6)
            ttk.Entry(frm,textvariable=var,width=92,show="*" if "Token" in lab else "").grid(row=row,column=1,columnspan=2,sticky="we",padx=4,pady=6)
        btn=ttk.Frame(frm); btn.grid(row=7,column=1,columnspan=2,sticky="e",pady=16)
        ttk.Button(btn,text="Restaurar valores por defecto",command=self.defaults).pack(side="left",padx=5)
        ttk.Button(btn,text="Guardar",command=self.save).pack(side="left",padx=5)
        ttk.Button(btn,text="Cerrar",command=self.destroy).pack(side="left",padx=5)

    def defaults(self):
        self.ruc_url.set(API_CONFIG_DEFAULTS["ruc_api_url"]); self.tc_url.set(API_CONFIG_DEFAULTS["tc_api_url"]); self.timeout.set(API_CONFIG_DEFAULTS["api_timeout"])

    def save(self):
        if not self.app.require_admin(self): return
        try:
            float(self.timeout.get() or 10)
            self.db.set_config("ruc_api_url", self.ruc_url.get(), self.app.usuario_actual, "URL consulta RUC")
            self.db.set_config("ruc_api_token", self.ruc_token.get(), self.app.usuario_actual, "Token consulta RUC")
            self.db.set_config("tc_api_url", self.tc_url.get(), self.app.usuario_actual, "URL tipo de cambio")
            self.db.set_config("tc_api_token", self.tc_token.get(), self.app.usuario_actual, "Token tipo de cambio")
            self.db.set_config("api_timeout", self.timeout.get(), self.app.usuario_actual, "Timeout API")
            messagebox.showinfo("Configuración", "Configuración de APIs guardada.", parent=self)
        except Exception as e:
            messagebox.showerror("Configuración", str(e), parent=self)


class UsersDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.title("Usuarios del Sistema")
        self.geometry("760x480+200+120")
        bind_as_child_window(self, self.master if hasattr(self, "master") else None)
        self.selected_user_id = None
        self.build()
        self.refresh()

    def build(self):
        cols = ("id", "usuario", "nombre", "rol", "activo", "creado")
        tree_box = ttk.Frame(self)
        tree_box.pack(fill="both", expand=True, padx=10, pady=10)
        tree_box.rowconfigure(0, weight=1); tree_box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(tree_box, columns=cols, show="headings", height=12)
        specs = [("id", "ID", 50), ("usuario", "Usuario", 110), ("nombre", "Nombre", 220), ("rol", "Rol", 110), ("activo", "Activo", 80), ("creado", "Creado", 160)]
        for c, t, w in specs:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c != "nombre" else "w", stretch=False)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0, column=1, sticky="ns"); hsb.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", self.select)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="x")
        self.usuario = tk.StringVar()
        self.nombre = tk.StringVar()
        self.password = tk.StringVar()
        self.rol = tk.StringVar(value="OPERADOR")
        ttk.Label(frm, text="Usuario:").grid(row=0, column=0)
        ttk.Entry(frm, textvariable=self.usuario, width=14).grid(row=0, column=1, padx=4)
        ttk.Label(frm, text="Nombre:").grid(row=0, column=2)
        ttk.Entry(frm, textvariable=self.nombre, width=22).grid(row=0, column=3, padx=4)
        ttk.Label(frm, text="Contraseña:").grid(row=0, column=4)
        ttk.Entry(frm, textvariable=self.password, width=14, show="*").grid(row=0, column=5, padx=4)
        ttk.Label(frm, text="Rol:").grid(row=1, column=0, pady=8)
        ttk.Combobox(frm, textvariable=self.rol, values=ROLES_VALIDOS, state="readonly", width=12).grid(row=1, column=1, padx=4, pady=8)
        self.perm_aqp = tk.BooleanVar(value=True)
        self.perm_min = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="Permite AQP", variable=self.perm_aqp).grid(row=1, column=2, padx=4, pady=8)
        ttk.Checkbutton(frm, text="Permite MIN", variable=self.perm_min).grid(row=1, column=3, padx=4, pady=8)
        ttk.Button(frm, text="Crear usuario", command=self.create_user).grid(row=1, column=4, padx=4)
        ttk.Button(frm, text="Cambiar contraseña", command=self.change_password).grid(row=1, column=5, padx=4)
        ttk.Button(frm, text="Activar/Desactivar", command=self.toggle_active).grid(row=2, column=4, padx=4)
        ttk.Button(frm, text="Guardar permisos almacén", command=self.save_permissions).grid(row=2, column=5, padx=4)

    def refresh(self):
        for x in self.tree.get_children():
            self.tree.delete(x)
        for r in self.db.list_users():
            self.tree.insert("", "end", values=(r["id"], r["usuario"], r["nombre"], r["rol"], "SI" if r["activo"] else "NO", r["creado_en"]))

    def select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            self.selected_user_id = None
            return
        values = self.tree.item(sel[0], "values")
        self.selected_user_id = int(values[0])
        self.usuario.set(values[1])
        self.nombre.set(values[2])
        self.rol.set(values[3])
        self.password.set("")
        try:
            perms = self.db.user_almacen_permissions(self.selected_user_id)
            self.perm_aqp.set(bool(perms.get("AQP", False)))
            self.perm_min.set(bool(perms.get("MIN", False)))
        except Exception:
            pass

    def create_user(self):
        if not self.app.require_admin(self):
            return
        try:
            nuevo_usuario = self.usuario.get()
            self.db.create_user(nuevo_usuario, self.nombre.get(), self.password.get(), self.rol.get())
            user = self.db.verify_user(nuevo_usuario, self.password.get())
            if user:
                perms = []
                if self.perm_aqp.get(): perms.append("AQP")
                if self.perm_min.get(): perms.append("MIN")
                self.db.set_user_almacen_permissions(user["id"], perms)
            self.db.audit("usuarios", self.usuario.get(), "CREACION", self.app.usuario_actual, "Creación de usuario")
            messagebox.showinfo("Completo", "Usuario creado correctamente.", parent=self)
            self.usuario.set("")
            self.nombre.set("")
            self.password.set("")
            self.rol.set("OPERADOR")
            self.refresh()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)
            self.after(100, self.focus_force)

    def change_password(self):
        if not self.app.require_admin(self):
            return
        if not self.selected_user_id:
            messagebox.showinfo("Seleccione", "Seleccione un usuario.", parent=self)
            return
        try:
            self.db.update_user_password(self.selected_user_id, self.password.get())
            self.db.audit("usuarios", self.selected_user_id, "CAMBIO CLAVE", self.app.usuario_actual, "Cambio de contraseña")
            messagebox.showinfo("Completo", "Contraseña actualizada.", parent=self)
            self.password.set("")
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)
            self.after(100, self.focus_force)

    def toggle_active(self):
        if not self.app.require_admin(self):
            return
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Seleccione", "Seleccione un usuario.", parent=self)
            return
        values = self.tree.item(sel[0], "values")
        active = values[4] != "SI"
        self.db.set_user_active(int(values[0]), active)
        self.db.audit("usuarios", values[1], "ACTIVACION" if active else "DESACTIVACION", self.app.usuario_actual, "Cambio de estado de usuario")
        self.refresh()

    def save_permissions(self):
        if not self.app.require_admin(self):
            return
        if not self.selected_user_id:
            messagebox.showinfo("Seleccione", "Seleccione un usuario.", parent=self)
            return
        perms = []
        if self.perm_aqp.get(): perms.append("AQP")
        if self.perm_min.get(): perms.append("MIN")
        self.db.set_user_almacen_permissions(self.selected_user_id, perms)
        self.db.audit("usuarios", self.selected_user_id, "PERMISOS ALMACEN", self.app.usuario_actual, ",".join(perms) or "SIN PERMISOS")
        messagebox.showinfo("Permisos", "Permisos de almacén actualizados.", parent=self)




class RequerimientoEstadoOperativoDialog(tk.Toplevel):
    def __init__(self, parent, app, requerimiento):
        super().__init__(parent); self.app=app; self.db=app.db; self.requerimiento=requerimiento
        self.title(f"Estado operativo del requerimiento {requerimiento}"); self.geometry("1120x560+90+80"); bind_as_child_window(self, parent)
        self.build()
    def build(self):
        req, rows = self.db.requerimiento_estado_operativo(self.requerimiento)
        ttk.Label(self, text=f"{req['requerimiento']} | {req['estado']} | Almacén {req['almacen_codigo']} | Solicitante {req['solicitante']} | CC {req['centro_costo_codigo']}", font=("Arial", 12, "bold")).pack(anchor="w", padx=12, pady=10)
        cols=("item","codigo","descripcion","req","atend","pend","stock","oc_estado")
        box=ttk.Frame(self); box.pack(fill="both", expand=True, padx=10, pady=8); box.rowconfigure(0,weight=1); box.columnconfigure(0,weight=1)
        tree=ttk.Treeview(box, columns=cols, show="headings", height=16)
        specs=[("item","Item",50),("codigo","Código",100),("descripcion","Descripción",330),("req","Req.",80),("atend","Atend.",80),("pend","Pend.",80),("stock",f"Stock {req['almacen_codigo']}",95),("oc_estado","OC / Estado compra",300)]
        for c,t,w in specs:
            tree.heading(c,text=t); tree.column(c,width=w,minwidth=w,anchor="center" if c!="descripcion" and c!="oc_estado" else "w",stretch=False)
        vsb=ttk.Scrollbar(box,orient="vertical",command=tree.yview); hsb=ttk.Scrollbar(box,orient="horizontal",command=tree.xview)
        tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        for r in rows:
            d=r['detalle']; pend=float(d['cantidad_requerida'] or 0)-float(d['cantidad_atendida'] or 0)
            oc_txt="SIN OC ACTIVA"
            if r['ocs']:
                parts=[]
                for oc in r['ocs'][:3]:
                    parts.append(f"{oc['numero']} {oc['estado']} / Pago {oc['estado_pago']} / Ing {float(oc['cantidad_ingresada'] or 0):g}")
                oc_txt=" | ".join(parts)
            tree.insert("","end",values=(d['item'],d['codigo'],d['descripcion'],f"{float(d['cantidad_requerida']):g}",f"{float(d['cantidad_atendida']):g}",f"{pend:g}",f"{float(r['stock_almacen']):g}",oc_txt))
        ttk.Button(self,text="Cerrar",command=self.destroy).pack(side="right",padx=12,pady=8)


class OCMapaProcesoDialog(tk.Toplevel):
    def __init__(self, parent, db, numero):
        super().__init__(parent); self.db=db; self.numero=numero
        self.title(f"Mapa de proceso OC {numero}"); self.geometry("980x620+140+60"); bind_as_child_window(self, parent)
        self.build()
    def build(self):
        data=self.db.oc_mapa_proceso(self.numero); oc=data['oc']
        ttk.Label(self,text=f"Mapa de proceso: {oc['numero']} | {oc['razon_social']} | Operativo {oc['estado']} {oc['avance']}% | Pago {oc['estado_pago']}",font=("Arial",12,"bold")).pack(anchor="w",padx=12,pady=10)
        canvas=ttk.Frame(self,padding=10); canvas.pack(fill="x",padx=10)
        steps=[
            ("1. Requerimiento", oc['requerimiento'] or 'OC manual'),
            ("2. Compra", f"{oc['estado']} | Total {oc['total']:.2f}"),
            ("3. Pago", f"{oc['estado_pago']} | Pagado {oc['monto_pagado_cache']:.2f}"),
            ("4. Transporte", f"Registros: {len(data['transito'])}"),
            ("5. Almacén/Despacho", f"Recepciones: {len(data['recepciones'])}"),
            ("6. Liquidación", "SI" if data['liquidacion'] else "NO"),
        ]
        for i,(t,sub) in enumerate(steps):
            card=ttk.LabelFrame(canvas,text=t,padding=8)
            card.grid(row=0,column=i,padx=4,sticky="nsew")
            ttk.Label(card,text=sub,wraplength=130,justify="center").pack()
        detail=ttk.Notebook(self); detail.pack(fill="both",expand=True,padx=10,pady=10)
        def add_tab(title, lines):
            frame=ttk.Frame(detail); detail.add(frame,text=title)
            txt=tk.Text(frame,wrap="word",height=12); scr=ttk.Scrollbar(frame,orient="vertical",command=txt.yview); txt.configure(yscrollcommand=scr.set)
            txt.pack(side="left",fill="both",expand=True); scr.pack(side="right",fill="y")
            txt.insert("end", "\n".join(lines) if lines else "Sin registros."); txt.configure(state="disabled")
        add_tab("Requerimiento", [f"{data['req']['requerimiento']} | {data['req']['estado']} | {data['req']['solicitante']}" if data['req'] else "OC manual sin requerimiento"])
        add_tab("Compra", [f"{d['item']} {d['codigo']} {d['descripcion']} | Sol {d['cantidad_solicitada']:g} | Aprob {d['cantidad_aprobada']:g} | Ing {d['cantidad_ingresada']:g}" for d in data['det']])
        add_tab("Pagos", [f"{fecha_text_from_iso(p['fecha_pago'])} | {p['banco']} | Fact {p['numero_factura'] if 'numero_factura' in p.keys() else ''} | Op {p['numero_operacion']} | {p['monto']:.2f} | {p['observacion']}" for p in data['pagos']])
        add_tab("Transporte", [f"{fecha_text_from_iso(t['fecha_despacho'])} | {t['operador_logistico']} | Guía {t['guia_transportista']} | Llegada {fecha_text_from_iso(t['fecha_estimada_llegada'])}" for t in data['transito']])
        add_tab("Recepciones", [f"{fecha_text_from_iso(r['fecha_recepcion'])} | Vale {r['vale']} | Doc {r['documento_recepcion']} | {r['observacion']}" for r in data['recepciones']])
        add_tab("Historial", [f"{h['fecha_hora']} | {h['estado_anterior']} -> {h['estado_nuevo']} | {h['usuario']} | {h['detalle']}" for h in data['historial']])
        ttk.Button(self,text="Cerrar",command=self.destroy).pack(side="right",padx=12,pady=8)


class ReversarValeDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.app=app; self.db=app.db
        self.title("Reversar / Anular Vale"); self.geometry("650x300+260+150"); bind_as_child_window(self, app.root)
        self.vale=tk.StringVar(); self.motivo=tk.StringVar()
        self.build()
    def build(self):
        frm=ttk.Frame(self,padding=18); frm.pack(fill="both",expand=True)
        ttk.Label(frm,text="Vale:").grid(row=0,column=0,sticky="w",pady=8)
        ent=ttk.Entry(frm,textvariable=self.vale,width=28); ent.grid(row=0,column=1,sticky="w",pady=8)
        ent.bind("<Shift-F2>", lambda e: self.buscar_vale())
        ent.bind("<Shift-F1>", lambda e: self.buscar_vale())
        ttk.Button(frm,text="Buscar vale Shift+F2",command=self.buscar_vale).grid(row=0,column=2,padx=8)
        ttk.Label(frm,text="Motivo:").grid(row=1,column=0,sticky="nw",pady=8)
        ttk.Entry(frm,textvariable=self.motivo,width=62).grid(row=1,column=1,columnspan=2,sticky="w",pady=8)
        ttk.Label(frm,text="El reverso no borra el vale original: genera un movimiento contrario con trazabilidad.",foreground="#666",wraplength=580).grid(row=2,column=0,columnspan=3,sticky="w",pady=8)
        ttk.Button(frm,text="Reversar vale",command=self.save).grid(row=3,column=1,sticky="e",pady=15)
        ttk.Button(frm,text="Cerrar",command=self.destroy).grid(row=3,column=2,sticky="w",pady=15,padx=8)
    def buscar_vale(self):
        ValeListDialog(self, self.db, self.set_vale)
    def set_vale(self, vale):
        self.vale.set(vale)
    def save(self):
        if not self.app.require_action("reversar_vale", self): return
        try:
            vale_rev=self.db.reverse_movimiento(self.vale.get(), self.motivo.get(), self.app.usuario_actual)
            messagebox.showinfo("Reverso registrado", f"Vale reverso generado: {vale_rev}", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror("No se pudo reversar", str(e), parent=self)


class AnularRequerimientoDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.app=app; self.db=app.db
        self.title("Anular requerimiento"); self.geometry("680x310+250+140"); bind_as_child_window(self, app.root)
        self.req=tk.StringVar(); self.motivo=tk.StringVar()
        self.build()
    def build(self):
        frm=ttk.Frame(self,padding=18); frm.pack(fill="both",expand=True)
        ttk.Label(frm,text="Requerimiento:").grid(row=0,column=0,sticky="w",pady=8)
        ent=ttk.Entry(frm,textvariable=self.req,width=28); ent.grid(row=0,column=1,sticky="w",pady=8)
        ent.bind("<Shift-F2>", lambda e: self.buscar_req())
        ent.bind("<Shift-F1>", lambda e: self.buscar_req())
        ttk.Button(frm,text="Buscar requerimiento Shift+F2",command=self.buscar_req).grid(row=0,column=2,padx=8)
        ttk.Label(frm,text="Motivo:").grid(row=1,column=0,sticky="nw",pady=8)
        ttk.Entry(frm,textvariable=self.motivo,width=62).grid(row=1,column=1,columnspan=2,sticky="w",pady=8)
        ttk.Label(frm,text="Solo se anulan requerimientos sin atención. Si ya tiene vales, primero revierta los movimientos asociados.",foreground="#666",wraplength=600).grid(row=2,column=0,columnspan=3,sticky="w",pady=8)
        ttk.Button(frm,text="Anular requerimiento",command=self.save).grid(row=3,column=1,sticky="e",pady=15)
        ttk.Button(frm,text="Cerrar",command=self.destroy).grid(row=3,column=2,sticky="w",pady=15,padx=8)
    def buscar_req(self):
        RequerimientoListDialog(self, self.db, self.set_req, estados=("PENDIENTE", "DESPACHO PARCIAL"))
    def set_req(self, req):
        self.req.set(req)
    def save(self):
        if not self.app.require_write(self): return
        try:
            req_num=self.db.anular_requerimiento(self.req.get(), self.motivo.get(), self.app.usuario_actual)
            messagebox.showinfo("Requerimiento anulado", f"Requerimiento anulado: {req_num}", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror("No se pudo anular", str(e), parent=self)


class KardexApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.withdraw()
        self.root.configure(bg="#eef2f6")

        splash = SplashScreen(self.root)
        try:
            splash.set_status("Preparando base de datos local...")
            self.db = KardexDB()
            splash.set_status("Cargando usuarios, artículos, almacenes y requerimientos...")
            self.root.after(350)
            splash.set_status("Validando configuración inicial...")
            self.root.after(250)
        finally:
            splash.close()

        self.current_user = None
        self.usuario_actual = ""
        self.almacen_actual = "AQP"
        self.authenticate()

    def authenticate(self):
        login = LoginDialog(self)
        self.root.wait_window(login)
        if not login.result:
            self.root.destroy()
            return
        self.current_user = login.result["user"]
        self.almacen_actual = login.result["almacen"]
        self.usuario_actual = self.current_user["usuario"]
        self.build_ui()
        self.root.deiconify()
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.attributes("-zoomed", True)

    def user_role(self):
        return self.current_user["rol"] if self.current_user else "CONSULTA"

    def can_write(self):
        return self.user_role() in ROLES_ESCRITURA

    def can_almacen(self):
        return self.user_role() in ROLES_ALMACEN

    def can_logistica(self):
        return self.user_role() in ROLES_LOGISTICA

    def can_finanzas(self):
        return self.user_role() in ROLES_FINANZAS

    def is_admin(self):
        return self.user_role() in ROLES_ADMIN

    def require_write(self, parent=None):
        if self.can_write():
            return True
        messagebox.showerror("Permiso denegado", "Su rol es CONSULTA. Puede revisar información y reportes, pero no registrar ni modificar datos.", parent=parent or self.root)
        return False

    def require_almacen(self, parent=None, almacen_codigo=None):
        if not self.can_almacen():
            messagebox.showerror("Permiso denegado", "Esta operación requiere rol de almacén u operador.", parent=parent or self.root)
            return False
        almacen_codigo = normalize_text(almacen_codigo or self.almacen_actual)
        if not self.db.user_can_access_almacen(self.current_user, almacen_codigo):
            messagebox.showerror("Permiso denegado", "Usted no está autorizado para realizar movimientos para este almacén.", parent=parent or self.root)
            return False
        return True

    def require_logistica(self, parent=None):
        if self.can_logistica():
            return True
        messagebox.showerror("Permiso denegado", "Esta operación requiere rol LOGISTICA, OPERADOR o ADMIN.", parent=parent or self.root)
        return False

    def require_finanzas(self, parent=None):
        if self.can_finanzas():
            return True
        messagebox.showerror("Permiso denegado", "Esta operación requiere rol FINANZAS, OPERADOR o ADMIN.", parent=parent or self.root)
        return False

    def require_admin(self, parent=None):
        if self.is_admin():
            return True
        messagebox.showerror("Permiso denegado", "Esta operación requiere rol ADMIN.", parent=parent or self.root)
        return False

    def require_action(self, accion, parent=None):
        if self.db.has_action_permission(self.usuario_actual, accion):
            return True
        etiquetas = dict(ACCIONES_PERMISO) if "ACCIONES_PERMISO" in globals() else {}
        messagebox.showerror(
            "Permiso denegado",
            f"No tiene autorización para {etiquetas.get(accion, accion)}.",
            parent=parent or self.root,
        )
        return False

    def reverse_vale_dialog(self):
        if not self.require_action("reversar_vale", self.root):
            return
        ReversarValeDialog(self)

    def annul_req_dialog(self):
        if not self.require_write(self.root):
            return
        AnularRequerimientoDialog(self)

    def build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        write_state = "normal" if self.can_write() else "disabled"
        almacen_state = "normal" if self.can_almacen() else "disabled"
        logistica_state = "normal" if self.can_logistica() else "disabled"
        finanzas_state = "normal" if self.can_finanzas() else "disabled"
        admin_state = "normal" if self.is_admin() else "disabled"
        crear_oc_state = "normal" if self.db.has_action_permission(self.usuario_actual, "crear_oc") else "disabled"
        aprobar_oc_state = "normal" if self.db.has_action_permission(self.usuario_actual, "aprobar_oc") else "disabled"
        reversar_vale_state = "normal" if self.db.has_action_permission(self.usuario_actual, "reversar_vale") else "disabled"
        registrar_pago_state = "normal" if self.db.has_action_permission(self.usuario_actual, "registrar_pago") else "disabled"
        proveedor_state = "normal" if self.db.has_action_permission(self.usuario_actual, "modificar_proveedor") else "disabled"

        menubar = tk.Menu(self.root)
        almacen = tk.Menu(menubar, tearoff=0)
        almacen.add_command(label="Ingreso / Salida Manual", command=lambda: EntryDialog(self), state=almacen_state)
        almacen.add_command(label="Atención de Requerimientos", command=lambda: AtencionRequerimientoDialog(self), state=almacen_state)
        almacen.add_command(label="Transferencia entre Almacenes", command=lambda: TransferDialog(self), state=almacen_state)
        almacen.add_command(label="Registro de Entrada con Orden de Compra", command=lambda: OCRecepcionDialog(self), state=almacen_state)
        almacen.add_separator()
        almacen.add_command(label="Registrar Artículo", command=lambda: RegisterArticleDialog(self), state=almacen_state)
        almacen.add_command(label="Registro Artículo Masivo", command=lambda: BulkArticleDialog(self), state=almacen_state)
        almacen.add_command(label="Consultar Vale", command=lambda: EntryDialog(self, mode="view"))
        almacen.add_command(label="Consultar Artículo", command=lambda: ArticleQueryDialog(self))
        almacen.add_separator()
        almacen.add_command(label="Reversar / Anular Vale", command=self.reverse_vale_dialog, state=reversar_vale_state)
        menubar.add_cascade(label="Almacén", menu=almacen)

        req_menu = tk.Menu(menubar, tearoff=0)
        req_menu.add_command(label="Requerimientos", command=lambda: RequerimientoDialog(self), state=logistica_state)
        req_menu.add_command(label="Consulta de Requerimientos", command=lambda: ConsultaRequerimientosDialog(self))
        req_menu.add_command(label="Anular Requerimiento", command=self.annul_req_dialog, state=logistica_state)
        menubar.add_cascade(label="Requerimientos", menu=req_menu)

        logist = tk.Menu(menubar, tearoff=0)
        logist.add_command(label="Registro Orden de Compra", command=lambda: OCListDialog(self), state=crear_oc_state)
        logist.add_command(label="Liquidación de Orden Compra", command=lambda: OCLiquidacionDialog(self), state=logistica_state)
        logist.add_command(label="Orden de Compra por Requerimiento", command=lambda: OCRequerimientoDialog(self), state=crear_oc_state)
        logist.add_command(label="Registro de Entrada con Orden de Compra", command=lambda: OCRecepcionDialog(self), state=almacen_state)
        logist.add_separator()
        logist.add_command(label="Aprobación Orden de Compra", command=lambda: OCAprobacionDialog(self), state=aprobar_oc_state)
        logist.add_command(label="Consulta de Orden de Compra", command=lambda: OCListDialog(self))
        logist.add_command(label="Reporte de Ordenes Pendientes", command=lambda: OCReportePendientesDialog(self))
        logist.add_separator()
        logist.add_command(label="Proveedores / Nuevo proveedor", command=lambda: ProviderListDialog(self), state=proveedor_state)
        logist.add_command(label="Registrar OC en Tránsito", command=lambda: OCTransitoDialog(self), state=logistica_state)
        menubar.add_cascade(label="Logística", menu=logist)

        finanzas = tk.Menu(menubar, tearoff=0)
        finanzas.add_command(label="Seguimiento de Pagos OC", command=lambda: OCFinanzasSeguimientoDialog(self), state=finanzas_state)
        finanzas.add_command(label="Registrar Pago de Orden de Compra", command=lambda: OCPagoDialog(self), state=registrar_pago_state)
        finanzas.add_command(label="Tipo de Cambio", command=lambda: TipoCambioDialog(self), state=finanzas_state)
        finanzas.add_command(label="Consulta de Ordenes de Compra", command=lambda: OCListDialog(self))
        finanzas.add_command(label="Reporte OC Pagadas/Pendientes", command=lambda: OCReportePendientesDialog(self))
        menubar.add_cascade(label="Finanzas", menu=finanzas)

        reportes = tk.Menu(menubar, tearoff=0)
        reportes.add_command(label="Reporte de Almacén Mensual", command=self.export_month_report)
        reportes.add_command(label="Reporte de Valorización Mensual", command=lambda: ValorizacionMensualDialog(self))
        reportes.add_command(label="Reporte de Stock", command=self.export_stock_report)
        reportes.add_command(label="Reporte de Artículos", command=self.export_articles_report)
        reportes.add_command(label="Reporte de Auditoría", command=self.export_audit_report)
        reportes.add_separator()
        reportes.add_command(label="Reporte General de Requerimientos", command=lambda: ConsultaRequerimientosDialog(self))
        menubar.add_cascade(label="Reportes", menu=reportes)

        config = tk.Menu(menubar, tearoff=0)
        config.add_command(label="Consultar Stock Shift+F2", command=lambda: StockDialog(self))
        config.add_command(label="Solicitantes", command=lambda: SolicitudesDialog(self), state=write_state)
        config.add_command(label="Usuarios", command=lambda: UsersDialog(self), state=admin_state)
        config.add_command(label="Tipo de Cambio", command=lambda: TipoCambioDialog(self), state=finanzas_state)
        config.add_command(label="Configuración de APIs", command=lambda: APIConfigDialog(self), state=admin_state)
        config.add_separator()
        config.add_command(label="Salir", command=self.root.destroy)
        menubar.add_cascade(label="Configuración", menu=config)
        self.root.config(menu=menubar)
        self.root.bind("<Shift-F2>", lambda e: (StockDialog(self), "break")[1])
        container = ttk.Frame(self.root, padding=30)
        container.pack(fill="both", expand=True)
        logo_path = os.path.join(BASE_DIR, "logo_empresa.png")
        if os.path.exists(logo_path):
            try:
                self.logo_img = tk.PhotoImage(file=logo_path)
                ttk.Label(container, image=self.logo_img).pack(pady=20)
            except Exception:
                ttk.Label(container, text="LOGO DE LA EMPRESA", font=("Arial", 30, "bold")).pack(pady=20)
        else:
            ttk.Label(container, text="LOGO DE LA EMPRESA", font=("Arial", 30, "bold")).pack(pady=20)
            ttk.Label(container, text="Para usar tu logo, guarda la imagen como logo_empresa.png en la misma carpeta del programa.").pack(pady=5)
        ttk.Label(container, text="Sistema Kardex de Almacén", font=("Arial", 24, "bold")).pack(pady=20)
        ttk.Label(container, text=f"Usuario conectado: {self.current_user['nombre']} ({self.current_user['rol']}) | Almacén: {self.almacen_actual} - {ALMACENES.get(self.almacen_actual, '')}", font=("Arial", 12)).pack(pady=5)
        ttk.Label(container, text="Use el menú superior para Kardex, requerimientos, compras, finanzas, consultas y reportes.", font=("Arial", 13)).pack(pady=10)
        ttk.Label(container, text="Atajos: Shift + F2 consulta stock | Shift + F1 ayuda en centro de costo, vales o requerimientos.", font=("Arial", 11), foreground="#555555").pack(pady=5)

    def export_excel(self, rows, headers, titles, filename):
        if Workbook is None:
            messagebox.showerror("Falta librería", "Instale openpyxl: pip install openpyxl", parent=self.root)
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "Reporte"
        ws.append(titles)
        header_fill = PatternFill("solid", fgColor="D9EAF7")
        thin = Side(style="thin", color="CCCCCC")
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.border = Border(bottom=thin)
            cell.alignment = Alignment(horizontal="center")
        for r in rows:
            row_values = []
            for idx, h in enumerate(headers):
                if isinstance(r, sqlite3.Row) and h in r.keys():
                    value = r[h]
                elif isinstance(r, dict):
                    value = r.get(h, "")
                elif isinstance(r, (list, tuple)) and idx < len(r):
                    value = r[idx]
                else:
                    value = ""
                if h in ("fecha_operacion", "fecha_orden", "fecha_requerimiento"):
                    value = fecha_text_from_iso(value)
                row_values.append(value)
            ws.append(row_values)
        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 10), 45)
        os.makedirs(REPORT_DIR, exist_ok=True)
        path = os.path.join(REPORT_DIR, filename)
        wb.save(path)
        messagebox.showinfo("Reporte generado", f"Reporte exportado:\n{path}", parent=self.root)

    def export_month_report(self):
        headers = ["vale", "fecha_operacion", "almacen_codigo", "tipo_codigo", "tipo_nombre", "documento", "requerimiento", "ot", "centro_costo_codigo", "centro_costo_nombre", "solicitante", "observacion", "usuario", "item", "codigo", "descripcion", "unidad_medida", "cantidad", "creado_en"]
        titles = ["Vale", "Fecha almacén", "Almacén", "Tipo", "Movimiento", "Documento", "Requerimiento", "OT", "CC", "Centro de costo", "Solicitante", "Observación", "Usuario", "Item", "Código", "Descripción", "UM", "Cantidad", "Registrado en"]
        self.export_excel(self.db.report_month(), headers, titles, f"reporte_almacen_mensual_{datetime.now().strftime('%Y%m')}.xlsx")

    def export_stock_report(self):
        rows = self.db.stock_rows()
        headers = ["codigo", "descripcion", "unidad_medida", "tipo_articulo", "stock", "stock_aqp", "comprometido_aqp", "libre_aqp", "stock_minimo_aqp", "estado_aqp", "stock_min", "comprometido_min", "libre_min", "stock_minimo_min", "estado_min", "libre_total", "estado_stock"]
        titles = ["Código", "Descripción", "UM", "Tipo", "Stock total", "Stock AQP", "Comprometido AQP", "Libre AQP", "Mínimo AQP", "Estado AQP", "Stock MIN", "Comprometido MIN", "Libre MIN", "Mínimo MIN", "Estado MIN", "Libre total", "Estado general"]
        self.export_excel(rows, headers, titles, f"reporte_stock_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

    def export_articles_report(self):
        headers = ["codigo", "descripcion", "unidad_medida", "tipo_articulo", "stock_minimo", "stock_minimo_aqp", "stock_minimo_min", "fecha_registro"]
        titles = ["Código", "Descripción", "UM", "Tipo", "Stock mínimo general", "Stock mínimo AQP", "Stock mínimo MIN", "Fecha registro"]
        self.export_excel(self.db.all_articulos(), headers, titles, f"reporte_articulos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

    def export_audit_report(self):
        headers = ["tabla", "registro", "accion", "usuario", "fecha_hora", "detalle"]
        titles = ["Tabla", "Registro", "Acción", "Usuario", "Fecha/Hora", "Detalle"]
        self.export_excel(self.db.audit_rows(), headers, titles, f"reporte_auditoria_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

    def export_requerimientos_report(self, estados):
        estado_txt = "_".join(e.lower().replace(" ", "_") for e in estados)
        headers = ["requerimiento", "fecha_requerimiento", "fecha_entrega_solicitada", "almacen_codigo", "almacen_nombre", "estado", "solicitante", "centro_costo_codigo", "centro_costo_nombre", "ot", "observacion", "item", "codigo", "descripcion", "unidad_medida", "cantidad_requerida", "cantidad_atendida", "pendiente", "usuario_creador", "creado_en", "modificado_en"]
        titles = ["Requerimiento", "Fecha solicitud", "Fecha entrega solicitada", "Almacén", "Almacén nombre", "Estado", "Solicitante", "CC", "Centro costo", "OT", "Observación", "Item", "Código", "Descripción", "UM", "Cant. requerida", "Cant. atendida", "Pendiente", "Usuario creador", "Creado en", "Modificado en"]
        self.export_excel(self.db.requerimiento_report_rows(estados), headers, titles, f"reporte_requerimientos_{estado_txt}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

    def export_pendientes_articulo_report(self):
        headers = ["almacen_codigo", "codigo", "descripcion", "unidad_medida", "pendiente_total"]
        titles = ["Almacén", "Código", "Descripción", "UM", "Pendiente total"]
        self.export_excel(self.db.pendientes_por_articulo(), headers, titles, f"reporte_pendientes_por_articulo_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

    def run(self):
        if self.root.winfo_exists():
            self.root.mainloop()


# ============================================================
# V40 - Consolidación ERP/MRP: reportes, cierre, permisos y auditoría
# ============================================================
ACCIONES_PERMISO = [
    ("crear_oc", "Crear orden de compra"),
    ("aprobar_oc", "Aprobar/desaprobar orden de compra"),
    ("anular_oc", "Anular orden de compra"),
    ("reversar_vale", "Reversar/anular vale de almacén"),
    ("registrar_pago", "Registrar pago de OC"),
    ("cerrar_mes", "Cerrar mes de almacén"),
    ("modificar_proveedor", "Crear / modificar proveedor"),
]


# --- Parches de base de datos -------------------------------------------------
_ORIG_KARDEX_CREATE_TABLES = KardexDB.create_tables
_ORIG_KARDEX_MIGRATE_TABLES = KardexDB.migrate_tables
_ORIG_SAVE_MOVIMIENTO = KardexDB.save_movimiento
_ORIG_REVERSE_MOVIMIENTO = KardexDB.reverse_movimiento
_ORIG_CREATE_OC = KardexDB.create_orden_compra
_ORIG_REGISTRAR_PAGO_OC = KardexDB.registrar_pago_oc
_ORIG_LIQUIDAR_OC = KardexDB.liquidar_orden_compra
_ORIG_RECEPCIONAR_OC = KardexDB.recepcionar_orden_compra


def _v40_create_tables(self):
    _ORIG_KARDEX_CREATE_TABLES(self)
    cur = self.conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS cierre_mensual_almacen (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            anio INTEGER NOT NULL,
            mes INTEGER NOT NULL,
            almacen_codigo TEXT NOT NULL DEFAULT 'TODOS',
            cerrado_por TEXT NOT NULL,
            fecha_cierre TEXT NOT NULL,
            snapshot_json TEXT NOT NULL DEFAULT '',
            observacion TEXT NOT NULL DEFAULT '',
            UNIQUE(anio, mes, almacen_codigo)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS historial_precios_articulo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT NOT NULL,
            descripcion TEXT NOT NULL DEFAULT '',
            ruc TEXT NOT NULL DEFAULT '',
            proveedor TEXT NOT NULL DEFAULT '',
            fecha TEXT NOT NULL,
            oc TEXT NOT NULL DEFAULT '',
            moneda TEXT NOT NULL DEFAULT 'SOLES',
            precio_sin_igv REAL NOT NULL DEFAULT 0,
            precio_con_igv REAL NOT NULL DEFAULT 0,
            tipo_cambio REAL NOT NULL DEFAULT 1,
            precio_equivalente_soles REAL NOT NULL DEFAULT 0,
            creado_en TEXT NOT NULL,
            UNIQUE(codigo, ruc, oc, precio_sin_igv)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS usuario_permiso_accion (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_id INTEGER NOT NULL,
            accion TEXT NOT NULL,
            permitido INTEGER NOT NULL DEFAULT 0,
            modificado_por TEXT NOT NULL DEFAULT '',
            modificado_en TEXT NOT NULL DEFAULT '',
            UNIQUE(usuario_id, accion),
            FOREIGN KEY(usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_hist_precio_codigo ON historial_precios_articulo(codigo, fecha)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_oc_pago_factura ON orden_compra_pagos(numero_factura)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_oc_pago_operacion ON orden_compra_pagos(banco, numero_operacion, monto)")
    self.conn.commit()


def _v40_migrate_tables(self):
    _ORIG_KARDEX_MIGRATE_TABLES(self)
    # Asegura tablas v40 también si la base ya existía antes.
    _v40_create_tables(self)
    self.ensure_default_action_permissions()


KardexDB.create_tables = _v40_create_tables
KardexDB.migrate_tables = _v40_migrate_tables


def _db_is_admin_user(self, usuario):
    cur = self.conn.cursor()
    cur.execute("SELECT rol FROM usuarios WHERE UPPER(usuario)=UPPER(?)", (str(usuario or '').strip(),))
    r = cur.fetchone()
    return bool(r and normalize_text(r["rol"]) == "ADMIN")


def _db_user_id(self, usuario):
    cur = self.conn.cursor()
    cur.execute("SELECT id FROM usuarios WHERE UPPER(usuario)=UPPER(?)", (str(usuario or '').strip(),))
    r = cur.fetchone()
    return int(r["id"]) if r else None


def _db_has_action_permission(self, usuario, accion):
    """Permiso por acción con fallback por rol para no romper bases antiguas."""
    accion = normalize_text(accion).lower()
    cur = self.conn.cursor()
    cur.execute("SELECT id, rol FROM usuarios WHERE UPPER(usuario)=UPPER(?)", (str(usuario or '').strip(),))
    r = cur.fetchone()
    if not r:
        # Compatibilidad con pruebas y scripts operativos que pasan el nombre del rol como usuario.
        rol_text = normalize_text(usuario)
        if accion == "registrar_pago" and rol_text in ("FINANZAS", "OPERADOR", "ADMIN"):
            return True
        if accion == "crear_oc" and rol_text in ("LOGISTICA", "OPERADOR", "ADMIN"):
            return True
        if accion == "reversar_vale" and rol_text in ("ALMACEN", "OPERADOR", "ADMIN"):
            return True
        if accion in ("cerrar_mes", "aprobar_oc") and rol_text == "ADMIN":
            return True
        return False
    if normalize_text(r["rol"]) == "ADMIN":
        return True
    cur.execute("SELECT permitido FROM usuario_permiso_accion WHERE usuario_id=? AND accion=?", (r["id"], accion))
    p = cur.fetchone()
    if p is not None:
        return bool(p["permitido"])
    rol = normalize_text(r["rol"])
    defaults = {
        "crear_oc": rol in ("LOGISTICA", "OPERADOR"),
        "aprobar_oc": False,
        "anular_oc": rol in ("LOGISTICA", "OPERADOR"),
        "reversar_vale": rol in ("ALMACEN", "OPERADOR"),
        "registrar_pago": rol in ("FINANZAS", "OPERADOR"),
        "cerrar_mes": False,
        "modificar_proveedor": rol in ("LOGISTICA", "OPERADOR"),
    }
    return bool(defaults.get(accion, False))


def _db_set_action_permission(self, usuario_id, accion, permitido, modificado_por):
    self.conn.execute("""
        INSERT INTO usuario_permiso_accion(usuario_id, accion, permitido, modificado_por, modificado_en)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(usuario_id, accion) DO UPDATE SET permitido=excluded.permitido, modificado_por=excluded.modificado_por, modificado_en=excluded.modificado_en
    """, (usuario_id, accion, 1 if permitido else 0, modificado_por, now_text()))
    self.conn.commit()


def _db_ensure_default_action_permissions(self):
    cur = self.conn.cursor()
    cur.execute("SELECT id, rol FROM usuarios")
    rows = cur.fetchall()
    for u in rows:
        rol = normalize_text(u["rol"])
        for accion, _desc in ACCIONES_PERMISO:
            if rol == "ADMIN":
                default = 1
            elif accion == "crear_oc":
                default = 1 if rol in ("LOGISTICA", "OPERADOR") else 0
            elif accion == "anular_oc":
                default = 1 if rol in ("LOGISTICA", "OPERADOR") else 0
            elif accion == "reversar_vale":
                default = 1 if rol in ("ALMACEN", "OPERADOR") else 0
            elif accion == "registrar_pago":
                default = 1 if rol in ("FINANZAS", "OPERADOR") else 0
            elif accion == "modificar_proveedor":
                default = 1 if rol in ("LOGISTICA", "OPERADOR") else 0
            else:
                default = 0
            cur.execute("INSERT OR IGNORE INTO usuario_permiso_accion(usuario_id, accion, permitido, modificado_por, modificado_en) VALUES (?, ?, ?, 'SISTEMA', ?)", (u["id"], accion, default, now_text()))
    self.conn.commit()


def _db_period_is_closed(self, fecha_iso, almacen_codigo="AQP"):
    try:
        d = datetime.strptime(str(fecha_iso), ISO_FMT).date()
    except Exception:
        return False
    cur = self.conn.cursor()
    cur.execute("""
        SELECT * FROM cierre_mensual_almacen
        WHERE anio=? AND mes=? AND (almacen_codigo=? OR almacen_codigo='TODOS')
        LIMIT 1
    """, (d.year, d.month, normalize_text(almacen_codigo or "TODOS")))
    return cur.fetchone() is not None


def _db_cerrar_periodo(self, anio, mes, almacen_codigo, usuario, observacion=""):
    anio, mes = int(anio), int(mes)
    almacen_codigo = normalize_text(almacen_codigo or "TODOS")
    if almacen_codigo not in ("TODOS", "AQP", "MIN"):
        raise ValueError("Almacén inválido para cierre.")
    if not self.has_action_permission(usuario, "cerrar_mes"):
        raise ValueError("No tiene permiso para cerrar mes de almacén.")
    rows = self.reporte_valorizacion_mensual(anio, mes, almacen_codigo)
    snapshot = []
    for r in rows:
        snapshot.append({k: (list(v) if isinstance(v, set) else v) for k, v in dict(r).items() if not str(k).startswith("_")})
    try:
        self.conn.execute("""
            INSERT INTO cierre_mensual_almacen(anio, mes, almacen_codigo, cerrado_por, fecha_cierre, snapshot_json, observacion)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (anio, mes, almacen_codigo, usuario, now_text(), json.dumps(snapshot, ensure_ascii=False), normalize_text(observacion)))
        self.audit("cierre_mensual_almacen", f"{anio:04d}-{mes:02d}-{almacen_codigo}", "CIERRE", usuario, f"Periodo cerrado con {len(snapshot)} línea(s). {observacion}", commit=False)
        self.conn.commit()
    except sqlite3.IntegrityError:
        self.conn.rollback()
        raise ValueError("El periodo seleccionado ya se encuentra cerrado.")


def _db_list_cierres(self):
    cur = self.conn.cursor()
    cur.execute("SELECT * FROM cierre_mensual_almacen ORDER BY anio DESC, mes DESC, almacen_codigo")
    return cur.fetchall()


def _db_registrar_historial_precios_oc(self, numero):
    oc, det = self.get_orden_compra(numero)
    if not oc:
        return 0
    count = 0
    cur = self.conn.cursor()
    for d in det:
        precio_sin = float(d["precio_unitario_sin_igv"] or 0)
        precio_con = float(d["precio_unitario_con_igv"] or 0)
        tc = float(oc["tipo_cambio"] or oc["tipo_cambio_venta"] or 1)
        moneda = normalize_text(oc["moneda"] or "SOLES")
        equiv = precio_sin * tc if moneda == "DOLARES" else precio_sin
        try:
            cur.execute("""
                INSERT INTO historial_precios_articulo(codigo, descripcion, ruc, proveedor, fecha, oc, moneda, precio_sin_igv, precio_con_igv, tipo_cambio, precio_equivalente_soles, creado_en)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (d["codigo"], d["descripcion"], oc["ruc"], oc["razon_social"], oc["fecha_orden"], oc["numero"], moneda, precio_sin, precio_con, tc, equiv, now_text()))
            count += 1
        except sqlite3.IntegrityError:
            pass
    self.conn.commit()
    return count


def _db_historial_precios_rows(self, codigo="", ruc=""):
    sql = "SELECT * FROM historial_precios_articulo WHERE 1=1"
    params = []
    if codigo:
        sql += " AND UPPER(codigo) LIKE ?"
        params.append(f"%{normalize_text(codigo)}%")
    if ruc:
        sql += " AND ruc LIKE ?"
        params.append(f"%{only_digits(ruc)}%")
    sql += " ORDER BY fecha DESC, id DESC LIMIT 1000"
    cur = self.conn.cursor()
    cur.execute(sql, params)
    return cur.fetchall()


def _db_reporte_kardex_valorizado_articulo(self, desde="", hasta="", almacen_codigo="TODOS", codigo="", moneda="SOLES"):
    """Kardex valorizado con promedio móvil por artículo/almacén."""
    desde_iso = fecha_iso_consulta_from_text(desde) if desde else "0001-01-01"
    hasta_iso = fecha_iso_consulta_from_text(hasta) if hasta else "9999-12-31"
    moneda = normalize_text(moneda or "SOLES")
    tc_ref = 1.0
    if moneda == "DOLARES":
        # Para reporte con conversión se exige TC del día final o del día actual si no hay fecha final.
        fecha_ref = hasta_iso if hasta else date.today().strftime(ISO_FMT)
        tc_ref = float(self.require_tipo_cambio_para_fecha(fecha_ref, "DOLARES")["venta"] or 0)
        if tc_ref <= 0:
            raise ValueError("No existe tipo de cambio válido para convertir el reporte a dólares.")
    params = [hasta_iso]
    where = ["m.fecha_operacion<=?"]
    alm = normalize_text(almacen_codigo or "TODOS")
    if alm != "TODOS":
        where.append("m.almacen_codigo=?")
        params.append(alm)
    if codigo:
        where.append("UPPER(d.codigo) LIKE ?")
        params.append(f"%{normalize_text(codigo)}%")
    cur = self.conn.cursor()
    cur.execute(f"""
        SELECT m.fecha_operacion, m.vale, m.tipo_codigo, m.tipo_nombre, m.documento, m.almacen_codigo, m.ot, m.centro_costo_codigo,
               d.codigo, d.descripcion, d.unidad_medida, d.cantidad, COALESCE(d.precio_unitario_sin_igv,0) precio
        FROM movimientos m
        JOIN movimiento_detalle d ON d.movimiento_id=m.id
        JOIN articulos a ON UPPER(a.codigo)=UPPER(d.codigo)
        WHERE {' AND '.join(where)} AND COALESCE(a.valoriza,1)=1
        ORDER BY d.codigo, m.almacen_codigo, m.fecha_operacion, m.id, d.item
    """, params)
    rows = cur.fetchall()
    state = {}
    out = []
    for r in rows:
        k = (r["almacen_codigo"], normalize_text(r["codigo"]))
        st = state.setdefault(k, {"qty": 0.0, "value": 0.0, "last": 0.0})
        qty = float(r["cantidad"] or 0)
        avg = (st["value"] / st["qty"]) if st["qty"] > 0.00001 else st["last"]
        if qty > 0:
            unit = float(r["precio"] or 0) or avg
            val = abs(qty) * unit
            st["qty"] += abs(qty); st["value"] += val
            if unit > 0: st["last"] = unit
            entrada, salida = abs(qty), 0.0
        else:
            unit = avg
            val = abs(qty) * unit
            st["qty"] -= abs(qty); st["value"] -= val
            entrada, salida = 0.0, abs(qty)
        if r["fecha_operacion"] >= desde_iso:
            factor = (1/tc_ref) if moneda == "DOLARES" else 1
            out.append({
                "fecha": fecha_text_from_iso(r["fecha_operacion"]),
                "documento": r["vale"],
                "tipo_movimiento": f"{r['tipo_codigo']} - {r['tipo_nombre']}",
                "codigo": r["codigo"],
                "descripcion": r["descripcion"],
                "entrada": round(entrada, 6),
                "salida": round(salida, 6),
                "stock_saldo": round(st["qty"], 6),
                "costo_unitario": round(unit * factor, 6),
                "valor_saldo": round(st["value"] * factor, 2),
                "moneda_reporte": moneda,
                "almacen": r["almacen_codigo"],
                "ot": r["ot"] or "",
                "centro_costo": r["centro_costo_codigo"] or "",
            })
    return out


def _db_reporte_compras_por_proveedor(self, desde="", hasta="", moneda="TODOS"):
    params = []
    where = ["estado NOT IN ('ANULADA','DENEGADA')"]
    if desde:
        where.append("fecha_orden>=?"); params.append(fecha_iso_consulta_from_text(desde))
    if hasta:
        where.append("fecha_orden<=?"); params.append(fecha_iso_consulta_from_text(hasta))
    mon = normalize_text(moneda or "TODOS")
    if mon != "TODOS":
        where.append("moneda=?"); params.append(mon)
    cur = self.conn.cursor()
    cur.execute(f"""
        SELECT razon_social AS proveedor, ruc, moneda, estado_pago, COUNT(*) cantidad_oc,
               SUM(subtotal) subtotal, SUM(igv) igv, SUM(total) total
        FROM orden_compra
        WHERE {' AND '.join(where)}
        GROUP BY razon_social, ruc, moneda, estado_pago
        ORDER BY SUM(total) DESC
    """, params)
    return cur.fetchall()


def _db_reporte_consumo_cc_ot(self, desde="", hasta="", almacen_codigo="TODOS"):
    params = []
    # V42: consumo = salida operativa real. No debe incluir transferencias internas
    # entre almacenes ni movimientos de reverso/corrección.
    where = [
        "d.cantidad<0",
        "m.tipo_codigo='CO'",
        "NOT EXISTS (SELECT 1 FROM transferencias t WHERE t.vale_salida=m.vale)",
        "NOT EXISTS (SELECT 1 FROM movimiento_reversos rv WHERE rv.vale_reverso=m.vale)",
    ]
    if desde:
        where.append("m.fecha_operacion>=?"); params.append(fecha_iso_consulta_from_text(desde))
    if hasta:
        where.append("m.fecha_operacion<=?"); params.append(fecha_iso_consulta_from_text(hasta))
    alm = normalize_text(almacen_codigo or "TODOS")
    if alm != "TODOS":
        where.append("m.almacen_codigo=?"); params.append(alm)
    cur = self.conn.cursor()
    cur.execute(f"""
        SELECT m.centro_costo_codigo, m.centro_costo_nombre, m.ot, d.codigo, d.descripcion,
               SUM(ABS(d.cantidad)) cantidad_consumida,
               SUM(ABS(d.cantidad) * COALESCE(d.precio_unitario_sin_igv,0)) valor_consumido,
               MIN(m.fecha_operacion) desde, MAX(m.fecha_operacion) hasta
        FROM movimientos m JOIN movimiento_detalle d ON d.movimiento_id=m.id
        WHERE {' AND '.join(where)}
        GROUP BY m.centro_costo_codigo, m.centro_costo_nombre, m.ot, d.codigo, d.descripcion
        ORDER BY m.centro_costo_codigo, m.ot, d.descripcion
    """, params)
    rows=[]
    for r in cur.fetchall():
        rows.append({
            "centro_costo": f"{r['centro_costo_codigo']} - {r['centro_costo_nombre']}",
            "ot": r["ot"] or "",
            "articulo": f"{r['codigo']} - {r['descripcion']}",
            "cantidad_consumida": float(r["cantidad_consumida"] or 0),
            "valor_consumido": float(r["valor_consumido"] or 0),
            "periodo": f"{fecha_text_from_iso(r['desde'])} - {fecha_text_from_iso(r['hasta'])}",
        })
    return rows


def _db_reporte_cuentas_pagar_aging(self, filtro="TODOS", proveedor="", moneda="TODOS"):
    self.refresh_payment_statuses()
    params=[]
    where=["estado NOT IN ('ANULADA','DENEGADA')"]
    if proveedor:
        like=f"%{normalize_text(proveedor)}%"
        where.append("(UPPER(razon_social) LIKE ? OR UPPER(ruc) LIKE ? OR UPPER(numero) LIKE ?)")
        params.extend([like, like, like])
    mon=normalize_text(moneda or "TODOS")
    if mon != "TODOS":
        where.append("moneda=?"); params.append(mon)
    cur=self.conn.cursor()
    cur.execute(f"""
        SELECT oc.*, COALESCE((SELECT SUM(monto) FROM orden_compra_pagos p WHERE p.oc_id=oc.id),0) pagado,
               (SELECT GROUP_CONCAT(NULLIF(numero_factura,''), ', ') FROM orden_compra_pagos p WHERE p.oc_id=oc.id) factura
        FROM orden_compra oc
        WHERE {' AND '.join(where)}
        ORDER BY fecha_vencimiento_pago ASC, numero
    """, params)
    today=date.today()
    out=[]
    filt=normalize_text(filtro or "TODOS")
    for r in cur.fetchall():
        total=float(r["total"] or 0); pagado=float(r["pagado"] or 0); saldo=round(total-pagado,2)
        venc_iso=r["fecha_vencimiento_pago"] or r["fecha_orden"]
        try: venc=datetime.strptime(venc_iso, ISO_FMT).date()
        except Exception: venc=today
        dias=(today-venc).days if saldo>0 else 0
        estado_pago=r["estado_pago"] or ("PAGADA" if saldo<=0.01 else "POR PAGAR")
        if filt == "POR VENCER" and not (saldo>0 and dias<=0): continue
        if filt == "VENCIDAS" and not (saldo>0 and dias>0): continue
        if filt == "PAGADAS" and saldo>0.01: continue
        if filt == "PAGO PARCIAL" and estado_pago != "PAGO PARCIAL": continue
        out.append({"proveedor": r["razon_social"], "oc": r["numero"], "factura": r["factura"] or "", "fecha_emision": fecha_text_from_iso(r["fecha_orden"]), "fecha_vencimiento": fecha_text_from_iso(venc_iso), "dias_vencidos": max(dias,0), "total": total, "pagado": pagado, "saldo": saldo, "estado": estado_pago, "moneda": r["moneda"]})
    return out


def _db_trazabilidad_documento(self, tipo, numero):
    tipo=normalize_text(tipo or "OC")
    numero=normalize_text(numero)
    rows=[]
    cur=self.conn.cursor()
    if tipo == "OC":
        oc, det = self.get_orden_compra(numero)
        if not oc: return []
        rows.append({"fecha": oc["creado_en"], "proceso": "CREACIÓN", "usuario": oc["usuario_creador"], "detalle": f"OC {oc['numero']} creada para {oc['razon_social']} por {fmt_money_currency(oc['total'], oc['moneda'])}"})
        for h in self.oc_historial(oc["id"]):
            rows.append({"fecha": h["fecha_hora"], "proceso": "ESTADO", "usuario": h["usuario"], "detalle": f"{h['estado_anterior']} -> {h['estado_nuevo']} | {h['detalle']}"})
        for p in self.pagos_oc(oc["id"]):
            rows.append({"fecha": p["creado_en"], "proceso": "PAGO", "usuario": p["usuario"], "detalle": f"Factura {p['numero_factura']} | Operación {p['numero_operacion']} | {p['banco']} | Monto {fmt_money_currency(p['monto'], oc['moneda'])}"})
        cur.execute("SELECT * FROM orden_compra_transito WHERE oc_id=? ORDER BY id", (oc["id"],))
        for t in cur.fetchall():
            rows.append({"fecha": t["creado_en"], "proceso": "TRÁNSITO", "usuario": t["usuario"], "detalle": f"{t['operador_logistico']} | Guía {t['guia_transportista']} | Estimada {fecha_text_from_iso(t['fecha_estimada_llegada'])}"})
        cur.execute("SELECT * FROM orden_compra_recepciones WHERE oc_id=? ORDER BY id", (oc["id"],))
        for rec in cur.fetchall():
            rows.append({"fecha": rec["creado_en"], "proceso": "RECEPCIÓN", "usuario": rec["usuario"], "detalle": f"Vale {rec['vale']} | Documento {rec['documento_recepcion']}"})
        cur.execute("SELECT * FROM orden_compra_liquidacion WHERE oc_id=?", (oc["id"],))
        for liq in cur.fetchall():
            rows.append({"fecha": liq["creado_en"], "proceso": "LIQUIDACIÓN", "usuario": liq["usuario"], "detalle": f"Factura {liq['factura']} | Guía {liq['guia']} | {liq['observacion']}"})
    elif tipo == "VALE":
        mov, det = self.get_movimiento_by_vale(numero)
        if not mov: return []
        rows.append({"fecha": mov["creado_en"], "proceso": "MOVIMIENTO", "usuario": mov["usuario"], "detalle": f"{mov['tipo_codigo']} {mov['tipo_nombre']} | Documento {mov['documento']} | Almacén {mov['almacen_codigo']}"})
        cur.execute("SELECT * FROM movimiento_reversos WHERE vale_original=? OR vale_reverso=?", (mov["vale"], mov["vale"]))
        for rev in cur.fetchall():
            rows.append({"fecha": rev["creado_en"], "proceso": "REVERSO", "usuario": rev["usuario"], "detalle": f"Original {rev['vale_original']} | Reverso {rev['vale_reverso']} | Motivo {rev['motivo']}"})
    else:
        cur.execute("SELECT * FROM auditoria WHERE UPPER(registro)=UPPER(?) ORDER BY fecha_hora", (numero,))
        for a in cur.fetchall():
            rows.append({"fecha": a["fecha_hora"], "proceso": a["accion"], "usuario": a["usuario"], "detalle": a["detalle"]})
    return sorted(rows, key=lambda x: str(x.get("fecha") or ""))


KardexDB.is_admin_user = _db_is_admin_user
KardexDB.user_id_by_login = _db_user_id
KardexDB.has_action_permission = _db_has_action_permission
KardexDB.set_action_permission = _db_set_action_permission
KardexDB.ensure_default_action_permissions = _db_ensure_default_action_permissions
KardexDB.period_is_closed = _db_period_is_closed
KardexDB.cerrar_periodo = _db_cerrar_periodo
KardexDB.list_cierres = _db_list_cierres
KardexDB.registrar_historial_precios_oc = _db_registrar_historial_precios_oc
KardexDB.historial_precios_rows = _db_historial_precios_rows
KardexDB.reporte_kardex_valorizado_articulo = _db_reporte_kardex_valorizado_articulo
KardexDB.reporte_compras_por_proveedor = _db_reporte_compras_por_proveedor
KardexDB.reporte_consumo_cc_ot = _db_reporte_consumo_cc_ot
KardexDB.reporte_cuentas_pagar_aging = _db_reporte_cuentas_pagar_aging
KardexDB.trazabilidad_documento = _db_trazabilidad_documento


def _db_require_tipo_cambio_para_fecha(self, fecha_iso_or_text, moneda="DOLARES"):
    moneda = normalize_text(moneda or "DOLARES")
    if moneda != "DOLARES":
        return {"fecha": fecha_iso_or_text, "compra": 1.0, "venta": 1.0, "fuente": "SOLES"}
    fecha_iso = fecha_iso_consulta_from_text(fecha_iso_or_text) if "/" in str(fecha_iso_or_text) else str(fecha_iso_or_text)
    row = self.get_tipo_cambio(fecha_iso)
    if not row:
        raise ValueError(f"No existe tipo de cambio registrado para la fecha {fecha_text_from_iso(fecha_iso)}. Regístrelo en Finanzas > Tipo de Cambio antes de emitir reportes con conversión.")
    return row

KardexDB.require_tipo_cambio_para_fecha = _db_require_tipo_cambio_para_fecha


def _v40_save_movimiento(self, fecha_texto, tipo_codigo, documento, ot_num, cc_codigo, solicitante, observacion, detalles, usuario, almacen_codigo="AQP", requerimiento="", commit=True):
    fecha_iso = fecha_iso_from_text(fecha_texto)
    tipo = normalize_text(tipo_codigo)
    if self.period_is_closed(fecha_iso, almacen_codigo) and not (self.is_admin_user(usuario) and tipo == "AJ"):
        raise ValueError(f"El periodo {fecha_iso[:7]} del almacén {almacen_codigo} se encuentra cerrado. Solo un ADMIN puede registrar ajustes autorizados.")
    return _ORIG_SAVE_MOVIMIENTO(self, fecha_texto, tipo_codigo, documento, ot_num, cc_codigo, solicitante, observacion, detalles, usuario, almacen_codigo, requerimiento, commit)


def _v40_reverse_movimiento(self, vale_original, motivo, usuario):
    mov, _ = self.get_movimiento_by_vale(vale_original)
    if mov and self.period_is_closed(mov["fecha_operacion"], mov["almacen_codigo"]) and not self.is_admin_user(usuario):
        raise ValueError("El vale pertenece a un periodo cerrado. Solo un ADMIN puede autorizar reversos sobre periodos cerrados.")
    return _ORIG_REVERSE_MOVIMIENTO(self, vale_original, motivo, usuario)


def _v40_create_orden_compra(self, *args, **kwargs):
    numero = _ORIG_CREATE_OC(self, *args, **kwargs)
    try:
        self.registrar_historial_precios_oc(numero)
    except Exception as e:
        self.audit("historial_precios_articulo", numero, "ADVERTENCIA", kwargs.get("usuario", "SISTEMA"), f"No se pudo registrar historial de precios: {e}")
    return numero


def _v40_recepcionar_orden_compra(self, numero, fecha_texto, documento_recepcion, observacion, cantidades_recibir, usuario):
    oc, _det = self.get_orden_compra(numero)
    if not oc:
        raise ValueError("No se encontró la orden de compra.")
    # V42: defensa en profundidad. Aunque la UI se equivoque o la función sea
    # invocada desde otro módulo, el usuario debe tener permiso sobre el almacén
    # al que pertenece la OC que se intenta recibir.
    if not self.is_admin_user(usuario):
        user_id = self.user_id_by_login(usuario)
        if user_id is None:
            raise ValueError("Usuario no válido para registrar recepción de OC.")
        if not self.user_can_access_almacen(user_id, oc["almacen_codigo"]):
            raise ValueError(f"No tiene permiso para registrar recepciones en el almacén {oc['almacen_codigo']}.")
    if self.period_is_closed(fecha_iso_from_text(fecha_texto), oc["almacen_codigo"]):
        raise ValueError(f"El periodo {fecha_iso_from_text(fecha_texto)[:7]} del almacén {oc['almacen_codigo']} está cerrado. No se puede registrar recepción por OC.")
    result = _ORIG_RECEPCIONAR_OC(self, numero, fecha_texto, documento_recepcion, observacion, cantidades_recibir, usuario)
    try:
        self.registrar_historial_precios_oc(numero)
    except Exception:
        pass
    return result


def _v40_registrar_pago_oc(self, numero, fecha_pago_texto, numero_operacion, banco, monto, observacion, usuario, numero_factura=""):
    oc, _ = self.get_orden_compra(numero)
    if not oc:
        raise ValueError("No se encontró la orden de compra.")
    cur = self.conn.cursor()
    op = normalize_text(numero_operacion)
    banco_n = normalize_text(banco)
    monto_f = float(str(monto or 0).replace(",", "."))
    if op:
        cur.execute("""
            SELECT oc.numero FROM orden_compra_pagos p JOIN orden_compra oc ON oc.id=p.oc_id
            WHERE UPPER(p.banco)=UPPER(?) AND UPPER(p.numero_operacion)=UPPER(?) AND ABS(p.monto-?)<0.01
            LIMIT 1
        """, (banco_n, op, monto_f))
        dup = cur.fetchone()
        if dup:
            raise ValueError(f"Pago duplicado: ya existe la operación {op} en {banco_n} para la OC {dup['numero']}.")
    fac = normalize_text(numero_factura)
    if fac:
        # V42: una factura puede pagarse en varias cuotas dentro de la MISMA OC.
        # Se conserva la alerta de duplicidad cuando la misma factura del proveedor
        # aparece asociada a una OC diferente.
        cur.execute("""
            SELECT oc.numero FROM orden_compra_pagos p JOIN orden_compra oc ON oc.id=p.oc_id
            WHERE UPPER(oc.ruc)=UPPER(?) AND UPPER(p.numero_factura)=UPPER(?)
              AND p.oc_id<>?
            LIMIT 1
        """, (oc["ruc"], fac, oc["id"]))
        dup = cur.fetchone()
        if dup:
            raise ValueError(f"Factura duplicada para el mismo proveedor. Ya figura en la OC {dup['numero']}.")
    if not self.has_action_permission(usuario, "registrar_pago"):
        raise ValueError("No tiene permiso para registrar pagos de OC.")
    return _ORIG_REGISTRAR_PAGO_OC(self, numero, fecha_pago_texto, numero_operacion, banco, monto, observacion, usuario, numero_factura)


def _v40_liquidar_orden_compra(self, numero, fecha_texto, factura, guia, observacion, usuario):
    oc, _ = self.get_orden_compra(numero)
    if oc and oc["estado"] == "ATENDIDA PARCIALMENTE" and not self.is_admin_user(usuario):
        raise ValueError("La OC tiene recepción parcial. Solo un ADMIN puede liquidarla con recepción pendiente.")
    return _ORIG_LIQUIDAR_OC(self, numero, fecha_texto, factura, guia, observacion, usuario)


KardexDB.save_movimiento = _v40_save_movimiento
KardexDB.reverse_movimiento = _v40_reverse_movimiento
KardexDB.create_orden_compra = _v40_create_orden_compra
KardexDB.recepcionar_orden_compra = _v40_recepcionar_orden_compra
KardexDB.registrar_pago_oc = _v40_registrar_pago_oc
KardexDB.liquidar_orden_compra = _v40_liquidar_orden_compra


# --- Diálogos v40 -------------------------------------------------------------
class SimpleReportDialog(tk.Toplevel):
    title_text = "Reporte"
    columns = []
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.db = app.db
        self.rows = []
        self.title(self.title_text)
        self.geometry("1200x620+70+60")
        self.minsize(980, 520)
        bind_as_child_window(self, app.root)
        self.build_base()

    def build_filters(self, parent):
        pass

    def fetch_rows(self):
        return []

    def build_base(self):
        self.filter_frame = ttk.LabelFrame(self, text="Filtros", padding=8)
        self.filter_frame.pack(fill="x", padx=10, pady=8)
        self.build_filters(self.filter_frame)
        mid = ttk.Frame(self)
        mid.pack(fill="both", expand=True, padx=10, pady=6)
        mid.rowconfigure(0, weight=1); mid.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(mid, columns=[c[0] for c in self.columns], show="headings", height=18)
        for key, title, width, anchor in self.columns:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor=anchor, stretch=False)
        vsb=ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        hsb=ttk.Scrollbar(mid, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        bottom=ttk.Frame(self, padding=10); bottom.pack(fill="x")
        self.status=tk.StringVar(value="")
        ttk.Label(bottom,textvariable=self.status,foreground="#555").pack(side="left")
        ttk.Button(bottom,text="Actualizar",command=self.load).pack(side="right",padx=4)
        ttk.Button(bottom,text="Exportar Excel",command=self.export).pack(side="right",padx=4)
        ttk.Button(bottom,text="Cerrar",command=self.destroy).pack(side="right",padx=4)
        self.load()

    def load(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        try:
            self.rows = list(self.fetch_rows())
            for r in self.rows:
                values=[]
                for key, _title, _width, _anchor in self.columns:
                    if isinstance(r, sqlite3.Row): values.append(r[key] if key in r.keys() else "")
                    else: values.append(r.get(key, ""))
                self.tree.insert("", "end", values=values)
            self.status.set(f"{len(self.rows)} registro(s).")
        except Exception as e:
            self.status.set("Error")
            messagebox.showerror("Reporte", str(e), parent=self)

    def export(self):
        headers=[c[0] for c in self.columns]
        titles=[c[1] for c in self.columns]
        filename=self.title_text.lower().replace(" ","_").replace("/","_") + "_" + datetime.now().strftime("%Y%m%d_%H%M") + ".xlsx"
        self.app.export_excel(self.rows, headers, titles, filename)


class ReporteKardexValorizadoArticuloDialog(SimpleReportDialog):
    title_text = "Kardex valorizado por artículo"
    columns = [("fecha","Fecha",95,"center"),("documento","Documento",130,"center"),("tipo_movimiento","Tipo movimiento",150,"w"),("codigo","Código",100,"center"),("descripcion","Descripción",300,"w"),("entrada","Entrada",90,"e"),("salida","Salida",90,"e"),("stock_saldo","Stock saldo",95,"e"),("costo_unitario","Costo unitario",110,"e"),("valor_saldo","Valor saldo",110,"e"),("moneda_reporte","Moneda",90,"center"),("almacen","Almacén",80,"center"),("ot","OT",110,"w"),("centro_costo","Centro costo",100,"center")]
    def build_filters(self, f):
        self.desde=tk.StringVar(); self.hasta=tk.StringVar(); self.codigo=tk.StringVar(); self.almacen=tk.StringVar(value="TODOS"); self.moneda=tk.StringVar(value="SOLES")
        for i,(txt,var,w) in enumerate([("Desde",self.desde,12),("Hasta",self.hasta,12),("Artículo",self.codigo,18)]):
            ttk.Label(f,text=txt+":").grid(row=0,column=i*2,sticky="w",padx=3); ttk.Entry(f,textvariable=var,width=w).grid(row=0,column=i*2+1,padx=3)
        ttk.Label(f,text="Almacén:").grid(row=0,column=6); ttk.Combobox(f,textvariable=self.almacen,values=["TODOS"]+list(ALMACENES.keys()),width=10,state="readonly").grid(row=0,column=7)
        ttk.Label(f,text="Moneda:").grid(row=0,column=8); ttk.Combobox(f,textvariable=self.moneda,values=MONEDAS,width=10,state="readonly").grid(row=0,column=9)
    def fetch_rows(self):
        return self.db.reporte_kardex_valorizado_articulo(self.desde.get(), self.hasta.get(), self.almacen.get(), self.codigo.get(), self.moneda.get())


class ReporteComprasProveedorDialog(SimpleReportDialog):
    title_text = "Compras por proveedor"
    columns = [("proveedor","Proveedor",320,"w"),("ruc","RUC",110,"center"),("cantidad_oc","Cantidad OC",90,"e"),("subtotal","Subtotal",120,"e"),("igv","IGV",110,"e"),("total","Total",120,"e"),("moneda","Moneda",90,"center"),("estado_pago","Estado pago",120,"center")]
    def build_filters(self, f):
        self.desde=tk.StringVar(); self.hasta=tk.StringVar(); self.moneda=tk.StringVar(value="TODOS")
        ttk.Label(f,text="Desde:").grid(row=0,column=0); ttk.Entry(f,textvariable=self.desde,width=12).grid(row=0,column=1,padx=4)
        ttk.Label(f,text="Hasta:").grid(row=0,column=2); ttk.Entry(f,textvariable=self.hasta,width=12).grid(row=0,column=3,padx=4)
        ttk.Label(f,text="Moneda:").grid(row=0,column=4); ttk.Combobox(f,textvariable=self.moneda,values=["TODOS"]+MONEDAS,width=10,state="readonly").grid(row=0,column=5)
    def fetch_rows(self):
        return self.db.reporte_compras_por_proveedor(self.desde.get(), self.hasta.get(), self.moneda.get())


class ReporteConsumoCCOTDialog(SimpleReportDialog):
    title_text = "Consumo por centro de costo y OT"
    columns = [("centro_costo","Centro de costo",260,"w"),("ot","OT",120,"w"),("articulo","Artículo",380,"w"),("cantidad_consumida","Cantidad consumida",130,"e"),("valor_consumido","Valor consumido",130,"e"),("periodo","Periodo",180,"center")]
    def build_filters(self, f):
        self.desde=tk.StringVar(); self.hasta=tk.StringVar(); self.almacen=tk.StringVar(value="TODOS")
        ttk.Label(f,text="Desde:").grid(row=0,column=0); ttk.Entry(f,textvariable=self.desde,width=12).grid(row=0,column=1,padx=4)
        ttk.Label(f,text="Hasta:").grid(row=0,column=2); ttk.Entry(f,textvariable=self.hasta,width=12).grid(row=0,column=3,padx=4)
        ttk.Label(f,text="Almacén:").grid(row=0,column=4); ttk.Combobox(f,textvariable=self.almacen,values=["TODOS"]+list(ALMACENES.keys()),width=10,state="readonly").grid(row=0,column=5)
    def fetch_rows(self):
        return self.db.reporte_consumo_cc_ot(self.desde.get(), self.hasta.get(), self.almacen.get())


class ReporteAgingCuentasPagarDialog(SimpleReportDialog):
    title_text = "Cuentas por pagar / Aging"
    columns = [("proveedor","Proveedor",280,"w"),("oc","OC",130,"center"),("factura","Factura",140,"w"),("fecha_emision","Fecha emisión",110,"center"),("fecha_vencimiento","Fecha vencimiento",120,"center"),("dias_vencidos","Días vencidos",110,"e"),("total","Total",115,"e"),("pagado","Pagado",115,"e"),("saldo","Saldo",115,"e"),("estado","Estado",120,"center"),("moneda","Moneda",90,"center")]
    def build_filters(self, f):
        self.filtro=tk.StringVar(value="TODOS"); self.proveedor=tk.StringVar(); self.moneda=tk.StringVar(value="TODOS")
        ttk.Label(f,text="Filtro:").grid(row=0,column=0); ttk.Combobox(f,textvariable=self.filtro,values=["TODOS","POR VENCER","VENCIDAS","PAGADAS","PAGO PARCIAL"],width=15,state="readonly").grid(row=0,column=1,padx=4)
        ttk.Label(f,text="Proveedor/OC:").grid(row=0,column=2); ttk.Entry(f,textvariable=self.proveedor,width=25).grid(row=0,column=3,padx=4)
        ttk.Label(f,text="Moneda:").grid(row=0,column=4); ttk.Combobox(f,textvariable=self.moneda,values=["TODOS"]+MONEDAS,width=10,state="readonly").grid(row=0,column=5)
    def fetch_rows(self):
        return self.db.reporte_cuentas_pagar_aging(self.filtro.get(), self.proveedor.get(), self.moneda.get())


class HistorialPreciosArticuloDialog(SimpleReportDialog):
    title_text = "Historial de precios por artículo"
    columns = [("codigo","Código",100,"center"),("descripcion","Descripción",300,"w"),("proveedor","Proveedor",280,"w"),("ruc","RUC",110,"center"),("fecha","Fecha",100,"center"),("oc","OC",130,"center"),("moneda","Moneda",90,"center"),("precio_sin_igv","Precio sin IGV",120,"e"),("precio_con_igv","Precio con IGV",120,"e"),("tipo_cambio","TC",90,"e"),("precio_equivalente_soles","Equiv. soles",120,"e")]
    def build_filters(self, f):
        self.codigo=tk.StringVar(); self.ruc=tk.StringVar()
        ttk.Label(f,text="Código artículo:").grid(row=0,column=0); ttk.Entry(f,textvariable=self.codigo,width=18).grid(row=0,column=1,padx=4)
        ttk.Label(f,text="RUC proveedor:").grid(row=0,column=2); ttk.Entry(f,textvariable=self.ruc,width=18).grid(row=0,column=3,padx=4)
    def fetch_rows(self):
        return self.db.historial_precios_rows(self.codigo.get(), self.ruc.get())


class TrazabilidadDocumentoDialog(tk.Toplevel):
    def __init__(self, app, tipo="OC", numero=""):
        super().__init__(app.root)
        self.app=app; self.db=app.db
        self.title("Trazabilidad de documento")
        self.geometry("1100x570+100+80")
        bind_as_child_window(self, app.root)
        self.tipo=tk.StringVar(value=tipo); self.numero=tk.StringVar(value=numero)
        self.build(); self.load()
    def build(self):
        top=ttk.Frame(self,padding=10); top.pack(fill="x")
        ttk.Label(top,text="Tipo:").pack(side="left"); ttk.Combobox(top,textvariable=self.tipo,values=["OC","VALE","AUDITORIA"],width=12,state="readonly").pack(side="left",padx=4)
        ttk.Label(top,text="Número:").pack(side="left",padx=(12,2)); ttk.Entry(top,textvariable=self.numero,width=28).pack(side="left",padx=4)
        ttk.Button(top,text="Buscar",command=self.load).pack(side="left",padx=6)
        box=ttk.Frame(self); box.pack(fill="both",expand=True,padx=10,pady=8); box.rowconfigure(0,weight=1); box.columnconfigure(0,weight=1)
        cols=("fecha","proceso","usuario","detalle"); self.tree=ttk.Treeview(box,columns=cols,show="headings",height=18)
        for c,t,w in [("fecha","Fecha/hora",160),("proceso","Proceso",130),("usuario","Usuario",130),("detalle","Detalle",650)]: self.tree.heading(c,text=t); self.tree.column(c,width=w,anchor="w",stretch=False)
        vsb=ttk.Scrollbar(box,orient="vertical",command=self.tree.yview); hsb=ttk.Scrollbar(box,orient="horizontal",command=self.tree.xview); self.tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns"); hsb.grid(row=1,column=0,sticky="ew")
        ttk.Button(self,text="Cerrar",command=self.destroy).pack(side="right",padx=10,pady=8)
    def load(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        rows=self.db.trazabilidad_documento(self.tipo.get(), self.numero.get()) if self.numero.get() else []
        for r in rows: self.tree.insert("","end",values=(r.get("fecha",""),r.get("proceso",""),r.get("usuario",""),r.get("detalle","")))


class CierreMensualDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app=app; self.db=app.db
        self.title("Cierre mensual de almacén")
        self.geometry("980x540+130+80")
        bind_as_child_window(self, app.root)
        self.anio=tk.IntVar(value=date.today().year); self.mes=tk.IntVar(value=date.today().month); self.almacen=tk.StringVar(value="TODOS"); self.obs=tk.StringVar()
        self.build(); self.load()
    def build(self):
        top=ttk.LabelFrame(self,text="Cerrar periodo",padding=10); top.pack(fill="x",padx=10,pady=8)
        ttk.Label(top,text="Año:").grid(row=0,column=0); ttk.Entry(top,textvariable=self.anio,width=8).grid(row=0,column=1,padx=4)
        ttk.Label(top,text="Mes:").grid(row=0,column=2); ttk.Combobox(top,textvariable=self.mes,values=list(range(1,13)),width=8,state="readonly").grid(row=0,column=3,padx=4)
        ttk.Label(top,text="Almacén:").grid(row=0,column=4); ttk.Combobox(top,textvariable=self.almacen,values=["TODOS"]+list(ALMACENES.keys()),width=10,state="readonly").grid(row=0,column=5,padx=4)
        ttk.Label(top,text="Observación:").grid(row=1,column=0,sticky="w",pady=6); ttk.Entry(top,textvariable=self.obs,width=80).grid(row=1,column=1,columnspan=5,sticky="ew",pady=6)
        ttk.Button(top,text="Cerrar mes",command=self.cerrar).grid(row=0,column=6,padx=8)
        ttk.Label(top,text="Al cerrar se bloquean movimientos anteriores. Solo ADMIN puede registrar ajustes autorizados.",foreground="#555").grid(row=2,column=0,columnspan=7,sticky="w")
        box=ttk.Frame(self); box.pack(fill="both",expand=True,padx=10,pady=8); box.rowconfigure(0,weight=1); box.columnconfigure(0,weight=1)
        cols=("anio","mes","almacen","cerrado_por","fecha_cierre","obs"); self.tree=ttk.Treeview(box,columns=cols,show="headings")
        for c,t,w in [("anio","Año",70),("mes","Mes",70),("almacen","Almacén",90),("cerrado_por","Cerrado por",120),("fecha_cierre","Fecha cierre",160),("obs","Observación",420)]: self.tree.heading(c,text=t); self.tree.column(c,width=w,anchor="w")
        vsb=ttk.Scrollbar(box,orient="vertical",command=self.tree.yview); self.tree.configure(yscrollcommand=vsb.set); self.tree.grid(row=0,column=0,sticky="nsew"); vsb.grid(row=0,column=1,sticky="ns")
        ttk.Button(self,text="Actualizar",command=self.load).pack(side="right",padx=8,pady=8); ttk.Button(self,text="Cerrar ventana",command=self.destroy).pack(side="right",padx=8,pady=8)
    def load(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for r in self.db.list_cierres(): self.tree.insert("","end",values=(r["anio"],r["mes"],r["almacen_codigo"],r["cerrado_por"],r["fecha_cierre"],r["observacion"]))
    def cerrar(self):
        if not messagebox.askyesno("Cerrar periodo", "¿Confirma cerrar el periodo seleccionado? Esta acción bloqueará movimientos del mes.", parent=self): return
        try:
            self.db.cerrar_periodo(self.anio.get(), self.mes.get(), self.almacen.get(), self.app.usuario_actual, self.obs.get())
            messagebox.showinfo("Cierre mensual", "Periodo cerrado correctamente.", parent=self)
            self.load()
        except Exception as e:
            messagebox.showerror("No se pudo cerrar", str(e), parent=self)


class PermisosAccionDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.app=app; self.db=app.db
        self.title("Permisos por acción"); self.geometry("840x520+180+100"); bind_as_child_window(self, app.root)
        self.users=[]; self.vars={}; self.user_var=tk.StringVar(); self.build(); self.load_users()
    def build(self):
        top=ttk.Frame(self,padding=10); top.pack(fill="x")
        ttk.Label(top,text="Usuario:").pack(side="left"); self.combo=ttk.Combobox(top,textvariable=self.user_var,width=40,state="readonly"); self.combo.pack(side="left",padx=6); self.combo.bind("<<ComboboxSelected>>",lambda e:self.load_perms())
        frame=ttk.LabelFrame(self,text="Acciones permitidas",padding=12); frame.pack(fill="both",expand=True,padx=10,pady=8)
        for i,(accion,desc) in enumerate(ACCIONES_PERMISO):
            self.vars[accion]=tk.BooleanVar(value=False)
            ttk.Checkbutton(frame,text=desc,variable=self.vars[accion]).grid(row=i,column=0,sticky="w",pady=5)
        bottom=ttk.Frame(self,padding=10); bottom.pack(fill="x")
        ttk.Button(bottom,text="Guardar permisos",command=self.save).pack(side="right",padx=5); ttk.Button(bottom,text="Cerrar",command=self.destroy).pack(side="right",padx=5)
    def load_users(self):
        cur=self.db.conn.cursor(); cur.execute("SELECT id, usuario, nombre, rol FROM usuarios ORDER BY usuario"); self.users=cur.fetchall(); vals=[f"{u['id']} | {u['usuario']} | {u['rol']} | {u['nombre']}" for u in self.users]; self.combo['values']=vals
        if vals: self.combo.current(0); self.load_perms()
    def selected_user_id(self):
        val=self.user_var.get().split('|')[0].strip(); return int(val) if val.isdigit() else None
    def load_perms(self):
        uid=self.selected_user_id();
        if not uid: return
        cur=self.db.conn.cursor(); cur.execute("SELECT rol FROM usuarios WHERE id=?",(uid,)); u=cur.fetchone(); rol=normalize_text(u["rol"]) if u else ""
        cur.execute("SELECT accion, permitido FROM usuario_permiso_accion WHERE usuario_id=?",(uid,)); data={r['accion']:bool(r['permitido']) for r in cur.fetchall()}
        for accion,_ in ACCIONES_PERMISO: self.vars[accion].set(True if rol=="ADMIN" else data.get(accion, False))
    def save(self):
        uid=self.selected_user_id()
        if not uid: return
        cur=self.db.conn.cursor(); cur.execute("SELECT rol FROM usuarios WHERE id=?",(uid,)); u=cur.fetchone(); rol=normalize_text(u["rol"]) if u else ""
        if rol=="ADMIN":
            for accion,_ in ACCIONES_PERMISO: self.db.set_action_permission(uid, accion, True, self.app.usuario_actual)
            self.load_perms()
            messagebox.showinfo("Permisos", "El rol ADMIN conserva acceso total por seguridad del sistema.", parent=self)
            return
        for accion,_ in ACCIONES_PERMISO: self.db.set_action_permission(uid, accion, self.vars[accion].get(), self.app.usuario_actual)
        messagebox.showinfo("Permisos", "Permisos actualizados.", parent=self)


class PruebaOperativaControladaDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root); self.title("Prueba operativa controlada"); self.geometry("760x520+200+90"); bind_as_child_window(self, app.root)
        txt=tk.Text(self, wrap="word", height=22); txt.pack(fill="both",expand=True,padx=12,pady=12)
        txt.insert("end", "Secuencia recomendada para validar el sistema antes de producción:\n\n")
        pasos=["Crear o revisar artículos valorizables y no valorizables.","Registrar solicitantes y proveedores.","Registrar tipo de cambio del día si se trabajará con dólares o reportes convertidos.","Registrar ingreso manual CN con precio sin IGV.","Registrar salida CO y verificar bloqueo de stock negativo.","Crear requerimiento con fecha de entrega solicitada.","Generar OC por requerimiento y verificar que el requerimiento ya no aparezca como pendiente de OC activa.","Aprobar OC total o parcial.","Registrar pago contado/crédito y verificar aging.","Registrar tránsito.","Registrar entrada por OC y verificar vale Word.","Liquidar OC.","Emitir Kardex valorizado por artículo, compras por proveedor, consumo por CC/OT y cuentas por pagar.","Cerrar el mes y comprobar bloqueo de movimientos antiguos.","Consultar trazabilidad de OC y vale."]
        for i,p in enumerate(pasos,1): txt.insert("end", f"{i}. {p}\n")
        txt.configure(state="disabled")
        ttk.Button(self,text="Cerrar",command=self.destroy).pack(pady=8)


# --- Menú v40 -----------------------------------------------------------------
_ORIG_APP_BUILD_UI = KardexApp.build_ui

def _v40_app_has_action(self, accion):
    return self.db.has_action_permission(self.usuario_actual, accion)


def _v40_build_ui(self):
    _ORIG_APP_BUILD_UI(self)
    try:
        menubar = self.root.nametowidget(self.root.cget("menu"))
    except Exception:
        return
    control = tk.Menu(menubar, tearoff=0)
    control.add_command(label="Prueba operativa controlada", command=lambda: PruebaOperativaControladaDialog(self))
    control.add_separator()
    control.add_command(label="Kardex valorizado por artículo", command=lambda: ReporteKardexValorizadoArticuloDialog(self))
    control.add_command(label="Compras por proveedor", command=lambda: ReporteComprasProveedorDialog(self))
    control.add_command(label="Consumo por centro de costo y OT", command=lambda: ReporteConsumoCCOTDialog(self))
    control.add_command(label="Cuentas por pagar / Aging", command=lambda: ReporteAgingCuentasPagarDialog(self))
    control.add_command(label="Historial de precios por artículo", command=lambda: HistorialPreciosArticuloDialog(self))
    control.add_separator()
    control.add_command(label="Cierre mensual de almacén", command=lambda: CierreMensualDialog(self), state=("normal" if self._v40_has_action("cerrar_mes") else "disabled"))
    control.add_command(label="Ver trazabilidad de documento", command=lambda: TrazabilidadDocumentoDialog(self))
    control.add_separator()
    control.add_command(label="Permisos por acción", command=lambda: PermisosAccionDialog(self), state=("normal" if self.is_admin() else "disabled"))
    menubar.add_cascade(label="Control ERP", menu=control)


KardexApp._v40_has_action = _v40_app_has_action
KardexApp.build_ui = _v40_build_ui



# =============================================================================
# V43 - Seguridad, respaldos, restauracion y permisos efectivos
# =============================================================================
V43_SCHEMA_VERSION = 43
V43_PBKDF2_ITERATIONS = 310000


def _v43_schema_version_from_file(db_path):
    if not db_path or db_path == ":memory:" or not os.path.exists(db_path) or os.path.getsize(db_path) == 0:
        return 0
    con = None
    try:
        con = sqlite3.connect(db_path)
        cur = con.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='schema_version'")
        if not cur.fetchone():
            return 0
        cur.execute("SELECT MAX(version) FROM schema_version")
        row = cur.fetchone()
        return int((row[0] if row else 0) or 0)
    except Exception:
        return 0
    finally:
        if con is not None:
            con.close()


def _v43_integrity_check_file(db_path):
    con = None
    try:
        con = sqlite3.connect(db_path)
        row = con.execute("PRAGMA integrity_check").fetchone()
        result = str(row[0] if row else "").strip().lower()
        if result != "ok":
            raise ValueError(f"La base de datos no supera la verificacion de integridad: {result or 'sin respuesta'}")
        return True
    except sqlite3.DatabaseError as exc:
        raise ValueError(f"El archivo seleccionado no es una base SQLite valida o esta corrupto: {exc}") from exc
    finally:
        if con is not None:
            con.close()


def _v43_copy_sqlite_database(source_path, destination_path):
    os.makedirs(os.path.dirname(os.path.abspath(destination_path)), exist_ok=True)
    src = sqlite3.connect(source_path)
    dst = sqlite3.connect(destination_path)
    try:
        src.backup(dst)
        dst.commit()
    finally:
        dst.close()
        src.close()
    _v43_integrity_check_file(destination_path)
    return destination_path


def _v43_prepare_pre_migration_backup(db_path):
    """Crea un respaldo consistente antes de migrar una BD anterior a V43."""
    if not db_path or db_path == ":memory:" or not os.path.exists(db_path) or os.path.getsize(db_path) == 0:
        return ""
    if _v43_schema_version_from_file(db_path) >= V43_SCHEMA_VERSION:
        return ""
    backup_dir = os.path.join(os.path.dirname(os.path.abspath(db_path)), "backups", "pre_migracion")
    os.makedirs(backup_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    dest = os.path.join(backup_dir, f"kardex_pre_v43_{stamp}.db")
    return _v43_copy_sqlite_database(db_path, dest)


_ORIG_V43_DB_INIT = KardexDB.__init__


def _v43_db_init(self, db_path=DB_FILE):
    pre_backup = ""
    try:
        pre_backup = _v43_prepare_pre_migration_backup(db_path)
        _ORIG_V43_DB_INIT(self, db_path)
        self.ensure_schema_version(V43_SCHEMA_VERSION, "V43 seguridad, respaldos y permisos efectivos")
        self.ensure_daily_backup()
    except Exception as exc:
        # Si una migracion falla, restaurar automaticamente la BD previa.
        try:
            if hasattr(self, "conn") and self.conn:
                self.conn.close()
        except Exception:
            pass
        if pre_backup and os.path.exists(pre_backup):
            try:
                tmp_restore = str(db_path) + ".restore_v43_tmp"
                if os.path.exists(tmp_restore):
                    os.remove(tmp_restore)
                _v43_copy_sqlite_database(pre_backup, tmp_restore)
                os.replace(tmp_restore, db_path)
            except Exception:
                pass
        raise RuntimeError(f"No se pudo inicializar/migrar la base de datos de forma segura: {exc}") from exc


def _db_ensure_schema_version(self, version=V43_SCHEMA_VERSION, descripcion=""):
    self.conn.execute("""
        CREATE TABLE IF NOT EXISTS schema_version (
            version INTEGER PRIMARY KEY,
            aplicado_en TEXT NOT NULL,
            descripcion TEXT NOT NULL DEFAULT ''
        )
    """)
    self.conn.execute(
        "INSERT OR IGNORE INTO schema_version(version, aplicado_en, descripcion) VALUES (?, ?, ?)",
        (int(version), now_text(), str(descripcion or "")[:300]),
    )
    self.conn.commit()
    return int(version)


def _db_get_schema_version(self):
    try:
        row = self.conn.execute("SELECT MAX(version) AS v FROM schema_version").fetchone()
        return int((row["v"] if row else 0) or 0)
    except Exception:
        return 0


def _db_integrity_check(self):
    row = self.conn.execute("PRAGMA integrity_check").fetchone()
    result = str(row[0] if row else "").strip()
    if result.lower() != "ok":
        raise ValueError(f"La base de datos presenta problemas de integridad: {result or 'sin respuesta'}")
    return "ok"


def _db_backup_database(self, destination=None, reason="MANUAL", usuario="SISTEMA", audit_event=True):
    if not self.db_path or self.db_path == ":memory:":
        raise ValueError("No se puede respaldar una base de datos en memoria.")
    self.integrity_check()
    base_dir = os.path.dirname(os.path.abspath(self.db_path))
    backup_dir = os.path.join(base_dir, "backups", "manuales")
    os.makedirs(backup_dir, exist_ok=True)
    if not destination:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        destination = os.path.join(backup_dir, f"kardex_backup_{stamp}.db")
    destination = os.path.abspath(destination)
    if os.path.abspath(self.db_path) == destination:
        raise ValueError("El respaldo no puede sobrescribir la base de datos activa.")
    os.makedirs(os.path.dirname(destination), exist_ok=True)
    if os.path.exists(destination):
        os.remove(destination)
    target = sqlite3.connect(destination)
    try:
        self.conn.backup(target)
        target.commit()
    finally:
        target.close()
    _v43_integrity_check_file(destination)
    if audit_event:
        try:
            self.audit("sistema", os.path.basename(destination), "BACKUP", usuario or "SISTEMA", f"Motivo: {reason}")
        except Exception:
            pass
    return destination


def _db_ensure_daily_backup(self, keep_days=30):
    if not self.db_path or self.db_path == ":memory:" or not os.path.exists(self.db_path):
        return ""
    backup_dir = os.path.join(os.path.dirname(os.path.abspath(self.db_path)), "backups", "diarios")
    os.makedirs(backup_dir, exist_ok=True)
    today_key = datetime.now().strftime("%Y%m%d")
    destination = os.path.join(backup_dir, f"kardex_diario_{today_key}.db")
    if not os.path.exists(destination):
        self.backup_database(destination, reason="AUTO_DIARIO", usuario="SISTEMA", audit_event=False)
    # Retencion simple por antiguedad del archivo.
    cutoff = datetime.now() - timedelta(days=max(int(keep_days), 1))
    for name in os.listdir(backup_dir):
        if not name.lower().endswith(".db"):
            continue
        path = os.path.join(backup_dir, name)
        try:
            if datetime.fromtimestamp(os.path.getmtime(path)) < cutoff:
                os.remove(path)
        except Exception:
            pass
    return destination


def _db_restore_database(self, source_path, usuario="SISTEMA"):
    if not self.db_path or self.db_path == ":memory:":
        raise ValueError("No se puede restaurar sobre una base de datos en memoria.")
    source_path = os.path.abspath(str(source_path or ""))
    if not os.path.isfile(source_path):
        raise ValueError("No se encontro el archivo de respaldo seleccionado.")
    _v43_integrity_check_file(source_path)

    # Respaldo de emergencia antes de cualquier restauracion.
    emergency = self.backup_database(reason="PRE_RESTAURACION", usuario=usuario, audit_event=False)
    tmp_path = os.path.abspath(self.db_path) + ".restore_tmp"
    if os.path.exists(tmp_path):
        os.remove(tmp_path)
    try:
        _v43_copy_sqlite_database(source_path, tmp_path)
        try:
            self.conn.close()
        except Exception:
            pass
        os.replace(tmp_path, self.db_path)
        self.conn = sqlite3.connect(self.db_path)
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.row_factory = sqlite3.Row
        self.create_tables()
        self.migrate_tables()
        self.ensure_default_user()
        self.ensure_user_warehouse_defaults()
        self.ensure_schema_version(V43_SCHEMA_VERSION, "Restauracion compatible con V43")
        self.integrity_check()
        try:
            self.audit("sistema", os.path.basename(source_path), "RESTAURACION", usuario or "SISTEMA", f"Respaldo de seguridad previo: {emergency}")
        except Exception:
            pass
        return emergency
    except Exception as exc:
        try:
            if hasattr(self, "conn") and self.conn:
                self.conn.close()
        except Exception:
            pass
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        # Recuperacion automatica al estado anterior si la restauracion falla.
        try:
            recovery_tmp = os.path.abspath(self.db_path) + ".recovery_tmp"
            if os.path.exists(recovery_tmp):
                os.remove(recovery_tmp)
            _v43_copy_sqlite_database(emergency, recovery_tmp)
            os.replace(recovery_tmp, self.db_path)
            self.conn = sqlite3.connect(self.db_path)
            self.conn.execute("PRAGMA foreign_keys = ON")
            self.conn.row_factory = sqlite3.Row
        except Exception:
            pass
        raise ValueError(f"No se pudo restaurar el respaldo. Se intento recuperar la BD anterior. Detalle: {exc}") from exc


# Contraseñas: se mantiene admin123 en beta, pero se almacena con PBKDF2-HMAC-SHA256.
def _v43_hash_password(self, password, salt=None):
    salt = salt or secrets.token_hex(16)
    iterations = V43_PBKDF2_ITERATIONS
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        str(password).encode("utf-8"),
        str(salt).encode("utf-8"),
        iterations,
    ).hex()
    return f"pbkdf2_sha256${iterations}${digest}", salt


def _v43_verify_user(self, usuario, password):
    usuario = normalize_text(usuario).lower()
    cur = self.conn.cursor()
    cur.execute("SELECT * FROM usuarios WHERE LOWER(usuario)=LOWER(?) AND activo=1", (usuario,))
    row = cur.fetchone()
    if not row:
        return None
    stored = str(row["password_hash"] or "")
    salt = str(row["salt"] or "")
    valid = False
    legacy = False
    if stored.startswith("pbkdf2_sha256$"):
        try:
            _scheme, iter_text, expected = stored.split("$", 2)
            iterations = int(iter_text)
            calculated = hashlib.pbkdf2_hmac(
                "sha256", str(password).encode("utf-8"), salt.encode("utf-8"), iterations
            ).hex()
            valid = secrets.compare_digest(calculated, expected)
        except Exception:
            valid = False
    else:
        # Compatibilidad V42 y anteriores: SHA-256(salt + password).
        calculated = hashlib.sha256((salt + str(password)).encode("utf-8")).hexdigest()
        valid = secrets.compare_digest(calculated, stored)
        legacy = valid
    if not valid:
        return None
    if legacy:
        # Migracion transparente al primer inicio de sesion exitoso.
        new_hash, new_salt = _v43_hash_password(self, password)
        self.conn.execute("UPDATE usuarios SET password_hash=?, salt=? WHERE id=?", (new_hash, new_salt, row["id"]))
        self.conn.commit()
        cur.execute("SELECT * FROM usuarios WHERE id=?", (row["id"],))
        row = cur.fetchone()
    return row


# Permisos efectivos: la capa de negocio es la autoridad final.
def _db_require_action_permission(self, usuario, accion):
    if not self.has_action_permission(usuario, accion):
        etiquetas = dict(ACCIONES_PERMISO)
        raise ValueError(f"Permiso denegado: no tiene autorizacion para {etiquetas.get(accion, accion)}.")
    return True


def _v43_user_from_call(args, kwargs, positional_index, default="SISTEMA"):
    if "usuario" in kwargs:
        return kwargs.get("usuario") or default
    if len(args) > positional_index:
        return args[positional_index] or default
    return default


_V43_PRE_CREATE_OC = KardexDB.create_orden_compra
_V43_PRE_APPROVE_OC = KardexDB.approve_orden_compra
_V43_PRE_DENY_OC = KardexDB.deny_orden_compra
_V43_PRE_DESAPPROVE_OC = KardexDB.desaprobar_orden_compra
_V43_PRE_ANNUL_OC = KardexDB.anular_orden_compra
_V43_PRE_REVERSE_VALE = KardexDB.reverse_movimiento
_V43_PRE_UPSERT_PROVIDER = KardexDB.upsert_proveedor
_V43_PRE_UPDATE_PENDING_OC = KardexDB.update_orden_compra_pendiente
_V43_PRE_SET_ACTION_PERMISSION = KardexDB.set_action_permission


def _v43_create_orden_compra(self, *args, **kwargs):
    usuario = _v43_user_from_call(args, kwargs, 13)
    self.require_action_permission(usuario, "crear_oc")
    return _V43_PRE_CREATE_OC(self, *args, **kwargs)


def _v43_approve_orden_compra(self, numero, cantidades_aprobadas, usuario, motivo=""):
    self.require_action_permission(usuario, "aprobar_oc")
    return _V43_PRE_APPROVE_OC(self, numero, cantidades_aprobadas, usuario, motivo)


def _v43_deny_orden_compra(self, numero, motivo, usuario):
    self.require_action_permission(usuario, "aprobar_oc")
    return _V43_PRE_DENY_OC(self, numero, motivo, usuario)


def _v43_desaprobar_orden_compra(self, numero, motivo, usuario):
    self.require_action_permission(usuario, "aprobar_oc")
    return _V43_PRE_DESAPPROVE_OC(self, numero, motivo, usuario)


def _v43_anular_orden_compra(self, numero, motivo, usuario):
    self.require_action_permission(usuario, "anular_oc")
    return _V43_PRE_ANNUL_OC(self, numero, motivo, usuario)


def _v43_reverse_movimiento(self, vale_original, motivo, usuario):
    self.require_action_permission(usuario, "reversar_vale")
    return _V43_PRE_REVERSE_VALE(self, vale_original, motivo, usuario)


def _v43_upsert_proveedor(self, *args, **kwargs):
    usuario = _v43_user_from_call(args, kwargs, 17)
    self.require_action_permission(usuario, "modificar_proveedor")
    return _V43_PRE_UPSERT_PROVIDER(self, *args, **kwargs)


def _v43_update_orden_compra_pendiente(self, numero, ruc, cotizacion, observacion, forma_pago, fecha_entrega, lugar_entrega, usuario):
    self.require_action_permission(usuario, "crear_oc")
    return _V43_PRE_UPDATE_PENDING_OC(self, numero, ruc, cotizacion, observacion, forma_pago, fecha_entrega, lugar_entrega, usuario)


def _v43_set_action_permission(self, usuario_id, accion, permitido, modificado_por):
    if not self.is_admin_user(modificado_por):
        raise ValueError("Solo un ADMIN puede modificar permisos por acción.")
    return _V43_PRE_SET_ACTION_PERMISSION(self, usuario_id, accion, permitido, modificado_por)


KardexDB.ensure_schema_version = _db_ensure_schema_version
KardexDB.get_schema_version = _db_get_schema_version
KardexDB.integrity_check = _db_integrity_check
KardexDB.backup_database = _db_backup_database
KardexDB.ensure_daily_backup = _db_ensure_daily_backup
KardexDB.restore_database = _db_restore_database
KardexDB.hash_password = _v43_hash_password
KardexDB.verify_user = _v43_verify_user
KardexDB.require_action_permission = _db_require_action_permission
KardexDB.create_orden_compra = _v43_create_orden_compra
KardexDB.approve_orden_compra = _v43_approve_orden_compra
KardexDB.deny_orden_compra = _v43_deny_orden_compra
KardexDB.desaprobar_orden_compra = _v43_desaprobar_orden_compra
KardexDB.anular_orden_compra = _v43_anular_orden_compra
KardexDB.reverse_movimiento = _v43_reverse_movimiento
KardexDB.upsert_proveedor = _v43_upsert_proveedor
KardexDB.update_orden_compra_pendiente = _v43_update_orden_compra_pendiente
KardexDB.set_action_permission = _v43_set_action_permission
KardexDB.__init__ = _v43_db_init


# UI V43: respaldo/restauracion e integridad accesibles solo al ADMIN.
def _v43_manual_backup_ui(self):
    if not self.require_admin(self.root):
        return
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    initial = f"kardex_backup_{stamp}.db"
    path = filedialog.asksaveasfilename(
        parent=self.root,
        title="Guardar respaldo de la base de datos",
        defaultextension=".db",
        filetypes=[("Base SQLite", "*.db"), ("Todos los archivos", "*.*")],
        initialfile=initial,
    )
    if not path:
        return
    try:
        result = self.db.backup_database(path, reason="MANUAL", usuario=self.usuario_actual)
        messagebox.showinfo("Respaldo", f"Respaldo creado y verificado correctamente:\n{result}", parent=self.root)
    except Exception as exc:
        messagebox.showerror("Respaldo", str(exc), parent=self.root)


def _v43_restore_backup_ui(self):
    if not self.require_admin(self.root):
        return
    path = filedialog.askopenfilename(
        parent=self.root,
        title="Seleccionar respaldo para restaurar",
        filetypes=[("Base SQLite", "*.db"), ("Todos los archivos", "*.*")],
    )
    if not path:
        return
    if not messagebox.askyesno(
        "Restaurar respaldo",
        "Se reemplazara la base de datos actual por el respaldo seleccionado.\n\nAntes de hacerlo se creara automaticamente un respaldo de emergencia.\n\n¿Desea continuar?",
        parent=self.root,
    ):
        return
    try:
        emergency = self.db.restore_database(path, usuario=self.usuario_actual)
        messagebox.showinfo(
            "Restauracion completada",
            f"La base fue restaurada y verificada.\n\nRespaldo de emergencia previo:\n{emergency}\n\nEl programa se cerrara para completar la recarga de forma segura.",
            parent=self.root,
        )
        self.root.destroy()
    except Exception as exc:
        messagebox.showerror("Restauracion", str(exc), parent=self.root)


def _v43_integrity_ui(self):
    if not self.require_admin(self.root):
        return
    try:
        result = self.db.integrity_check()
        version = self.db.get_schema_version()
        daily = self.db.ensure_daily_backup()
        messagebox.showinfo(
            "Estado de la base de datos",
            f"Integridad SQLite: {result.upper()}\nVersion de esquema: {version}\nRespaldo diario: {daily}",
            parent=self.root,
        )
    except Exception as exc:
        messagebox.showerror("Estado de la base de datos", str(exc), parent=self.root)


_ORIG_V43_BUILD_UI = KardexApp.build_ui


def _v43_build_ui(self):
    _ORIG_V43_BUILD_UI(self)
    try:
        menubar = self.root.nametowidget(self.root.cget("menu"))
    except Exception:
        return
    security = tk.Menu(menubar, tearoff=0)
    state = "normal" if self.is_admin() else "disabled"
    security.add_command(label="Crear respaldo ahora", command=lambda: self._v43_manual_backup(), state=state)
    security.add_command(label="Restaurar respaldo", command=lambda: self._v43_restore_backup(), state=state)
    security.add_separator()
    security.add_command(label="Verificar integridad de la base", command=lambda: self._v43_integrity_check(), state=state)
    menubar.add_cascade(label="Seguridad / Respaldo", menu=security)


KardexApp._v43_manual_backup = _v43_manual_backup_ui
KardexApp._v43_restore_backup = _v43_restore_backup_ui
KardexApp._v43_integrity_check = _v43_integrity_ui
KardexApp.build_ui = _v43_build_ui




# =============================================================================
# V44 - Robustez SQLite, correlativos anuales, Decimal, fechas ISO y modularidad
# =============================================================================
V44_SCHEMA_VERSION = 44
APP_LOGGER = configure_app_logger(BASE_DIR)
install_global_exception_logging(APP_LOGGER)


def _v44_prepare_pre_migration_backup(db_path):
    if not db_path or db_path == ":memory:" or not os.path.exists(db_path) or os.path.getsize(db_path) == 0:
        return ""
    if _v43_schema_version_from_file(db_path) >= V44_SCHEMA_VERSION:
        return ""
    backup_dir = os.path.join(os.path.dirname(os.path.abspath(db_path)), "backups", "pre_migracion")
    os.makedirs(backup_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    dest = os.path.join(backup_dir, f"kardex_pre_v44_{stamp}.db")
    return _v43_copy_sqlite_database(db_path, dest)


def _v44_normalize_internal_timestamps(conn):
    # Solo columnas de auditoria/sistema; las fechas visibles de negocio se
    # conservan en DD/MM/AAAA y las fechas operativas ya usan YYYY-MM-DD.
    columns = {
        "usuarios": ["creado_en"],
        "articulos": ["fecha_registro"],
        "solicitantes": ["creado_en"],
        "movimientos": ["creado_en", "modificado_en"],
        "movimiento_reversos": ["creado_en"],
        "requerimientos": ["creado_en", "modificado_en"],
        "transferencias": ["creado_en"],
        "proveedores": ["creado_en", "modificado_en"],
        "orden_compra": ["creado_en", "modificado_en", "fecha_aprobacion"],
        "orden_compra_estado_historial": ["fecha_hora"],
        "orden_compra_pagos": ["creado_en"],
        "orden_compra_transito": ["creado_en"],
        "orden_compra_recepciones": ["creado_en"],
        "orden_compra_liquidacion": ["creado_en"],
        "historial_precios_articulo": ["creado_en"],
        "auditoria": ["fecha_hora"],
        "sistema_config": ["modificado_en"],
        "usuario_permiso_accion": ["modificado_en"],
        "cierre_mensual_almacen": ["fecha_cierre"],
    }
    tables = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
    changed = 0
    for table, cols in columns.items():
        if table not in tables:
            continue
        existing = {r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()}
        for col in cols:
            if col not in existing:
                continue
            rows = conn.execute(f"SELECT rowid, {col} FROM {table} WHERE IFNULL({col},'')<>''").fetchall()
            for row in rows:
                old = row[1]
                new = normalize_datetime_text(old)
                if new != old:
                    conn.execute(f"UPDATE {table} SET {col}=? WHERE rowid=?", (new, row[0]))
                    changed += 1
    return changed


def _v44_apply_schema(self):
    AnnualCorrelativeService.ensure_table(self.conn)
    index_count = ensure_performance_indexes(self.conn)
    normalized = _v44_normalize_internal_timestamps(self.conn)
    # Configuracion declarativa que facilitara PostgreSQL sin cambiar la UI.
    config_rows = [
        ("db_backend", "sqlite", "Motor de base durante fase beta"),
        ("precision_cantidad", "3", "Decimales maximos de cantidades"),
        ("precision_precio_unitario", "4", "Decimales de precio/costo unitario"),
        ("precision_tipo_cambio", "4", "Decimales de tipo de cambio"),
        ("precision_importe", "2", "Decimales de importes finales"),
        ("correlativos_anuales", "1", "Correlativos reiniciados por tipo y anio"),
    ]
    existing_cols = self.columns("sistema_config") if "sistema_config" else set()
    if existing_cols:
        for clave, valor, descripcion in config_rows:
            self.conn.execute("""
                INSERT OR IGNORE INTO sistema_config(clave, valor, descripcion, modificado_por, modificado_en)
                VALUES (?, ?, ?, 'SISTEMA', ?)
            """, (clave, valor, descripcion, iso_now()))
    self.ensure_schema_version(V44_SCHEMA_VERSION, "V44 robustez SQLite, correlativos, Decimal, fechas ISO y modularidad")
    self.conn.commit()
    APP_LOGGER.info("Migracion V44 aplicada | indices=%s | timestamps_normalizados=%s", index_count, normalized)
    return {"indices": index_count, "timestamps": normalized}


_V44_PRE_DB_INIT = KardexDB.__init__


def _v44_db_init(self, db_path=DB_FILE):
    pre_backup = _v44_prepare_pre_migration_backup(db_path)
    try:
        _V44_PRE_DB_INIT(self, db_path)
        apply_sqlite_runtime_pragmas(self.conn, self.db_path, busy_timeout_ms=15000)
        if self.get_schema_version() < V44_SCHEMA_VERSION:
            _v44_apply_schema(self)
        else:
            # Los indices/pragmas se aseguran tambien en aperturas normales.
            ensure_performance_indexes(self.conn)
            self.conn.commit()
        APP_LOGGER.info("Base abierta | path=%s | schema=%s", self.db_path, self.get_schema_version())
    except Exception as exc:
        APP_LOGGER.exception("Fallo inicializando V44 | db=%s", db_path)
        try:
            if hasattr(self, "conn") and self.conn:
                self.conn.close()
        except Exception:
            pass
        if pre_backup and os.path.exists(pre_backup):
            try:
                tmp_restore = str(db_path) + ".restore_v44_tmp"
                if os.path.exists(tmp_restore):
                    os.remove(tmp_restore)
                _v43_copy_sqlite_database(pre_backup, tmp_restore)
                os.replace(tmp_restore, db_path)
            except Exception:
                APP_LOGGER.exception("No se pudo recuperar automaticamente el backup pre-V44")
        raise RuntimeError(f"No se pudo inicializar/migrar V44 de forma segura: {exc}") from exc


def _v44_calc_estado_pago(self, total, pagado, fecha_vencimiento_iso):
    total_d = q_money(total)
    pagado_d = q_money(pagado)
    saldo = q_money(max(D("0"), total_d - pagado_d))
    if total_d > 0 and saldo == 0:
        return "PAGADA"
    if fecha_vencimiento_iso:
        try:
            venc = datetime.strptime(fecha_vencimiento_iso, ISO_FMT).date()
            if venc < date.today() and saldo > 0:
                return "VENCIDA"
        except Exception:
            pass
    if pagado_d > 0:
        return "PAGO PARCIAL"
    return "POR PAGAR"


def _v44_sync_oc_payment_status(self, oc_id, commit=False):
    cur = self.conn.cursor()
    oc = cur.execute("SELECT total, fecha_vencimiento_pago FROM orden_compra WHERE id=?", (oc_id,)).fetchone()
    if not oc:
        return "POR PAGAR", 0.0, 0.0
    pagos = cur.execute("SELECT monto FROM orden_compra_pagos WHERE oc_id=?", (oc_id,)).fetchall()
    pagado_d = money_sum([r["monto"] for r in pagos]) if pagos else q_money(0)
    total_d = q_money(oc["total"] or 0)
    saldo_d = q_money(max(D("0"), total_d - pagado_d))
    estado = self._calc_estado_pago(total_d, pagado_d, oc["fecha_vencimiento_pago"])
    cur.execute(
        "UPDATE orden_compra SET estado_pago=?, monto_pagado_cache=?, modificado_en=? WHERE id=?",
        (estado, float(pagado_d), iso_now(), oc_id),
    )
    if commit:
        self.conn.commit()
    return estado, float(pagado_d), float(saldo_d)


def _v44_recalc_oc_totals(self, oc_id):
    rows = self.conn.execute(
        "SELECT subtotal, igv, total FROM orden_compra_detalle WHERE oc_id=? ORDER BY id", (oc_id,)
    ).fetchall()
    subtotal = money_sum([r["subtotal"] for r in rows]) if rows else q_money(0)
    igv = money_sum([r["igv"] for r in rows]) if rows else q_money(0)
    total = money_sum([r["total"] for r in rows]) if rows else q_money(0)
    self.conn.execute(
        "UPDATE orden_compra SET subtotal=?, igv=?, total=?, modificado_en=? WHERE id=?",
        (float(subtotal), float(igv), float(total), iso_now(), oc_id),
    )
    return float(subtotal), float(igv), float(total)


def _v44_registrar_pago_oc(self, numero, fecha_pago_texto, numero_operacion, banco, monto, observacion, usuario, numero_factura=""):
    self.require_action_permission(usuario, "registrar_pago")
    oc, _ = self.get_orden_compra(numero)
    if not oc:
        raise ValueError("No se encontro la orden de compra.")
    if oc["estado"] in ("PENDIENTE", "EN COTIZACION", "DENEGADA", "ANULADA"):
        raise ValueError("La OC debe estar aprobada para registrar pago.")

    fecha_iso = fecha_iso_from_text(fecha_pago_texto)
    op = normalize_text(numero_operacion)
    banco_n = normalizar_banco(banco)
    obs = normalize_text(observacion)
    fac = normalize_text(numero_factura)
    monto_d = q_money(monto)
    if not op:
        raise ValueError("Ingrese numero de operacion de pago.")
    if not banco_n or (banco_codigo_guardado(banco_n) not in BANCOS_MAP and not banco_n.startswith("05 OTROS")):
        raise ValueError("Banco invalido.")
    if monto_d <= 0:
        raise ValueError("El monto pagado debe ser mayor a cero.")

    cur = self.conn.cursor()
    cur.execute("""
        SELECT oc.numero FROM orden_compra_pagos p JOIN orden_compra oc ON oc.id=p.oc_id
        WHERE UPPER(p.banco)=UPPER(?) AND UPPER(p.numero_operacion)=UPPER(?) AND ABS(p.monto-?)<0.005
        LIMIT 1
    """, (banco_n, op, float(monto_d)))
    dup = cur.fetchone()
    if dup:
        raise ValueError(f"Pago duplicado: ya existe la operacion {op} en {banco_n} para la OC {dup['numero']}.")
    if fac:
        cur.execute("""
            SELECT oc.numero FROM orden_compra_pagos p JOIN orden_compra oc ON oc.id=p.oc_id
            WHERE UPPER(oc.ruc)=UPPER(?) AND UPPER(p.numero_factura)=UPPER(?) AND p.oc_id<>?
            LIMIT 1
        """, (oc["ruc"], fac, oc["id"]))
        dup = cur.fetchone()
        if dup:
            raise ValueError(f"Factura duplicada para el mismo proveedor. Ya figura en la OC {dup['numero']}.")

    pagos = cur.execute("SELECT monto FROM orden_compra_pagos WHERE oc_id=?", (oc["id"],)).fetchall()
    pagado_actual = money_sum([r["monto"] for r in pagos]) if pagos else q_money(0)
    total_oc = q_money(oc["total"] or 0)
    if q_money(pagado_actual + monto_d) > total_oc:
        raise ValueError("El pago acumulado no puede superar el total de la OC.")

    try:
        cur.execute("""
            INSERT INTO orden_compra_pagos(oc_id, fecha_pago, numero_operacion, numero_factura, banco, monto, observacion, usuario, creado_en)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (oc["id"], fecha_iso, op, fac, banco_n, float(monto_d), obs, usuario, iso_now()))
        estado_pago, pagado_nuevo, saldo = self._sync_oc_payment_status(oc["id"], commit=False)
        factura_txt = f" factura {fac}" if fac else ""
        self._add_oc_historial(
            oc["id"], oc["estado"], oc["estado"], usuario,
            f"Pago financiero {estado_pago}: operacion {op}{factura_txt} por {monto_d:.2f}. Saldo {saldo:.2f}",
        )
        self.audit(
            "orden_compra", oc["numero"], "PAGO", usuario,
            f"Estado pago {estado_pago}: operacion {op}{factura_txt} por {monto_d:.2f}. Avance operativo se mantiene en {oc['estado']} {oc['avance']}%",
            commit=False,
        )
        self.conn.commit()
        APP_LOGGER.info("Pago OC | oc=%s | usuario=%s | monto=%s | estado=%s", numero, usuario, monto_d, estado_pago)
        return estado_pago
    except Exception:
        self.conn.rollback()
        APP_LOGGER.exception("Fallo registrando pago | oc=%s | usuario=%s", numero, usuario)
        raise


_V44_PRE_RESTORE_DATABASE = KardexDB.restore_database


def _v44_restore_database(self, source_path, usuario="SISTEMA"):
    emergency = _V44_PRE_RESTORE_DATABASE(self, source_path, usuario=usuario)
    apply_sqlite_runtime_pragmas(self.conn, self.db_path, busy_timeout_ms=15000)
    if self.get_schema_version() < V44_SCHEMA_VERSION:
        _v44_apply_schema(self)
    else:
        ensure_performance_indexes(self.conn)
        self.conn.commit()
    self.integrity_check()
    APP_LOGGER.info("Restauracion completada y migrada a V44 | source=%s | usuario=%s", source_path, usuario)
    return emergency


def _v44_tk_exception(self, exc, val, tb):
    APP_LOGGER.exception("Error en callback Tkinter", exc_info=(exc, val, tb))
    try:
        messagebox.showerror(
            "Error interno",
            "Ocurrio un error inesperado. El detalle tecnico fue registrado en la carpeta logs.",
            parent=self.root,
        )
    except Exception:
        pass


_V44_PRE_APP_INIT = KardexApp.__init__


def _v44_app_init(self):
    _V44_PRE_APP_INIT(self)
    try:
        self.root.report_callback_exception = lambda exc, val, tb: _v44_tk_exception(self, exc, val, tb)
    except Exception:
        pass


KardexDB.__init__ = _v44_db_init
KardexDB._calc_estado_pago = _v44_calc_estado_pago
KardexDB._sync_oc_payment_status = _v44_sync_oc_payment_status
KardexDB._recalc_oc_totals = _v44_recalc_oc_totals
KardexDB.registrar_pago_oc = _v44_registrar_pago_oc
KardexDB.restore_database = _v44_restore_database
KardexApp.__init__ = _v44_app_init


if __name__ == "__main__":
    app = KardexApp()
    app.run()
