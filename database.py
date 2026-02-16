import sqlite3
from models import Site, Collection, Assemblage, Material, Level, Specimen

class Database:
    def __init__(self, db_path, create_tables=False):
        """Inicializa a conexão com o banco de dados. Cria as tabelas apenas se necessário."""
        self.db_path = db_path
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
        self.cursor.execute("PRAGMA foreign_keys = ON;") # Ativa suporte a chaves estrangeiras no SQLite
        
        if create_tables:
            self.create_tables()  # Só cria tabelas se solicitado

    def create_tables(self):
        """Cria as tabelas no banco de dados."""
        self.cursor.execute("""
        CREATE TABLE IF NOT EXISTS Site (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            state TEXT,
            city TEXT,
            location TEXT,
            number INTEGER,
            longitude REAL,
            latitude REAL
        )
        """)

        self.cursor.execute("""
        CREATE TABLE IF NOT EXISTS Collection (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            longitude REAL,
            latitude REAL,
            FOREIGN KEY (site_id) REFERENCES Site(id) ON DELETE CASCADE
        )
        """)

        self.cursor.execute("""
        CREATE TABLE IF NOT EXISTS Assemblage (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            collection_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            longitude REAL,
            latitude REAL,
            screenSize TEXT,
            FOREIGN KEY (collection_id) REFERENCES Collection(id) ON DELETE CASCADE
        )
        """)

        # Tabela Unidade de Escavação (ExcavationUnit)
        self.cursor.execute("""
        CREATE TABLE IF NOT EXISTS ExcavationUnit (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            assemblage_id INTEGER NOT NULL,
            name TEXT,
            unit_type TEXT,
            size TEXT,
            latitude REAL,
            longitude REAL,
            nivel_holotipo TEXT,
            profundidade_inicial REAL,
            profundidade_final REAL,
            camada_geologica TEXT,
            metodo_escavacao TEXT,
            metodo_peneiramento TEXT,
            responsavel_escavacao TEXT,
            data_inicio TEXT,
            data_conclusao TEXT,
            fotos_registro TEXT,
            observacao TEXT,
            FOREIGN KEY (assemblage_id) REFERENCES Assemblage(id) ON DELETE CASCADE
        );
        """)

        self.cursor.execute("""
        CREATE TABLE IF NOT EXISTS Level (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            excavation_unit_id INTEGER NOT NULL,
            level TEXT,
            start_depth REAL,
            end_depth REAL,
            color TEXT,
            texture TEXT,
            description TEXT,
            FOREIGN KEY (excavation_unit_id) REFERENCES ExcavationUnit(id) ON DELETE CASCADE
        )
        """)

        self.cursor.execute("""
        CREATE TABLE IF NOT EXISTS Material (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            excavation_unit_id INTEGER NOT NULL,
            level_id INTEGER,
            uuid TEXT,
            field_serial TEXT,
            lab_serial TEXT,
            material_type TEXT,
            material_description TEXT,
            measurements REAL,
            weight REAL,
            quantity INTEGER,
            x_coord REAL,
            y_coord REAL,
            z_coord REAL,
            notes TEXT,
            photos TEXT,
            user TEXT,
            cataloging_date TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (excavation_unit_id) REFERENCES ExcavationUnit(id) ON DELETE CASCADE,
            FOREIGN KEY (level_id) REFERENCES Level(id) ON DELETE SET NULL
        );
        """)

        self.cursor.execute("""
        CREATE TABLE IF NOT EXISTS Specimen (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            material_id INTEGER,
            excavation_unit_id INTEGER,
            level_id INTEGER,
            uuid TEXT,
            field_serial TEXT,
            lab_serial TEXT,
            class TEXT,
            bio_order TEXT,
            family TEXT,
            genus TEXT,
            species TEXT,
            taxon TEXT,
            element TEXT,
            symmetry TEXT,
            portion TEXT,
            completion TEXT,
            sex TEXT,
            weight REAL,
            fusion TEXT,
            weathering TEXT,
            burning TEXT,
            gnawing TEXT,
            butchering TEXT,
            pathology TEXT,
            dentition TEXT,
            tooth_wear TEXT,
            measurements REAL,
            x_coord REAL,
            y_coord REAL,
            z_coord REAL,
            notes TEXT,
            photos TEXT,
            user TEXT,
            cataloging_date TEXT,
            FOREIGN KEY (material_id) REFERENCES Material(id) ON DELETE CASCADE,
            FOREIGN KEY (excavation_unit_id) REFERENCES ExcavationUnit(id) ON DELETE SET NULL,
            FOREIGN KEY (level_id) REFERENCES Level(id) ON DELETE SET NULL
        );
        """)

        self.conn.commit()

    def insert(self, table, data):
        """Insere um novo registro na tabela."""
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['?' for _ in data])
        query = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
        self.cursor.execute(query, tuple(data.values()))
        self.conn.commit()
        return self.cursor.lastrowid

    def execute_query(self, query, params=()):
        """Executa uma query SQL bruta e retorna os resultados (útil para agregações)."""
        self.cursor.execute(query, params)
        return self.cursor.fetchall()

    def update(self, table, data, condition):
        """Atualiza registros na tabela baseado na condição."""
        set_clause = ', '.join([f"{key} = ?" for key in data.keys()])
        condition_clause = ' AND '.join([f"{key} = ?" for key in condition.keys()])
        query = f"UPDATE {table} SET {set_clause} WHERE {condition_clause}"
        self.cursor.execute(query, tuple(data.values()) + tuple(condition.values()))
        self.conn.commit()

    def delete(self, table, condition):
        """Deleta registros da tabela baseado na condição."""
        condition_clause = ' AND '.join([f"{key} = ?" for key in condition.keys()])
        query = f"DELETE FROM {table} WHERE {condition_clause}"
        self.cursor.execute(query, tuple(condition.values()))
        self.conn.commit()

    def fetch(self, table, columns='*', condition=None):
        """Busca registros na tabela baseado em uma condição opcional."""
        condition_clause = ''
        values = []
        if condition:
            condition_clause = 'WHERE ' + ' AND '.join([f"{key} = ?" for key in condition.keys()])
            values = list(condition.values())
        
        query = f"SELECT {columns} FROM {table} {condition_clause}"
        self.cursor.execute(query, values)
        return self.cursor.fetchall()

    def close(self):
        """Fecha a conexão com o banco de dados."""
        self.conn.close()
