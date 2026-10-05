class AnnualCorrelativeService:
    """Generador anual atomico de correlativos para SQLite.

    Usa la conexion activa del ERP. Si no existe una transaccion inicia
    BEGIN IMMEDIATE y confirma solo la reserva del numero. Si ya existe una
    transaccion (p. ej. una transferencia con dos vales), participa en ella.
    Esto evita duplicados entre procesos y no abre una segunda conexion que
    pueda bloquear la transaccion actual.
    """

    def __init__(self, db_path=None, connection=None, timeout_ms=15000):
        self.db_path = db_path
        self.connection = connection
        self.timeout_ms = int(timeout_ms)

    @staticmethod
    def format(doc_type, year, number):
        return f"{str(doc_type).upper()}-{int(year):04d}-{int(number):06d}"

    @staticmethod
    def ensure_table(conn):
        conn.execute("""
            CREATE TABLE IF NOT EXISTS correlativos (
                tipo TEXT NOT NULL,
                anio INTEGER NOT NULL,
                ultimo INTEGER NOT NULL DEFAULT 0,
                PRIMARY KEY (tipo, anio)
            )
        """)

    def next(self, doc_type, year):
        doc_type = str(doc_type or "").strip().upper()
        if not doc_type:
            raise ValueError("Tipo de correlativo vacio.")
        year = int(year)
        conn = self.connection
        if conn is None:
            raise ValueError("No hay conexion disponible para generar el correlativo.")

        conn.execute(f"PRAGMA busy_timeout={self.timeout_ms}")
        already_in_tx = bool(getattr(conn, "in_transaction", False))
        try:
            if not already_in_tx:
                conn.execute("BEGIN IMMEDIATE")
            self.ensure_table(conn)
            row = conn.execute(
                "SELECT ultimo FROM correlativos WHERE tipo=? AND anio=?",
                (doc_type, year),
            ).fetchone()
            number = (int(row[0]) if row else 0) + 1
            if row:
                conn.execute(
                    "UPDATE correlativos SET ultimo=? WHERE tipo=? AND anio=?",
                    (number, doc_type, year),
                )
            else:
                conn.execute(
                    "INSERT INTO correlativos(tipo, anio, ultimo) VALUES (?, ?, ?)",
                    (doc_type, year, number),
                )
            if not already_in_tx:
                conn.commit()
            return self.format(doc_type, year, number)
        except Exception:
            if not already_in_tx:
                conn.rollback()
            raise
