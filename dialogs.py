import uuid
from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLineEdit, QPushButton, QMessageBox, 
    QComboBox, QScrollArea, QWidget, QGridLayout, QLabel, QHBoxLayout
)

class NewDatabaseDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Novo Banco de Dados")
        self.setWindowIcon(QIcon("./icon/iconTray.png"))
        self.layout = QVBoxLayout(self)
        self.input_name = QLineEdit(self)
        self.input_name.setPlaceholderText("Nome do banco de dados")
        self.layout.addWidget(self.input_name)
        self.button_ok = QPushButton("Criar", self)
        self.button_ok.clicked.connect(self.accept)
        self.layout.addWidget(self.button_ok)

    def get_db_name(self):
        return self.input_name.text().strip()

class AddSiteDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Adicionar Sítio")
        self.setWindowIcon(QIcon("./icon/iconTray.png"))
        layout = QVBoxLayout(self)
        self.inputs = {}
        campos = ["name", "state", "city", "location", "number", "longitude", "latitude"]
        for campo in campos:
            input_field = QLineEdit(self)
            input_field.setPlaceholderText(campo)
            layout.addWidget(input_field)
            self.inputs[campo.lower()] = input_field
        button_save = QPushButton("Salvar")
        button_save.clicked.connect(self.save_site)
        layout.addWidget(button_save)

    def save_site(self):
        data = {k: v.text() for k, v in self.inputs.items()}
        if not data["name"]:
            QMessageBox.warning(self, "Erro", "O campo Nome é obrigatório.")
            return
        self.db.insert("Site", data)
        self.accept()

class AddCollectionDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Adicionar Coleção")
        self.setWindowIcon(QIcon("./icon/iconTray.png"))
        layout = QVBoxLayout(self)
        self.input_collection_name = QLineEdit(self)
        self.input_collection_name.setPlaceholderText("Nome da Coleção")
        layout.addWidget(self.input_collection_name)
        self.input_site_id = QComboBox(self)
        self.input_site_id.setPlaceholderText("Selecione o Sítio")
        layout.addWidget(self.input_site_id)
        sites = self.db.fetch("Site", "id, name")
        for site_id, site_name in sites:
            self.input_site_id.addItem(str(site_name), site_id)
        self.input_collection_longitude = QLineEdit(self)
        self.input_collection_longitude.setPlaceholderText("Longitude")
        layout.addWidget(self.input_collection_longitude)
        self.input_collection_latitude = QLineEdit(self)
        self.input_collection_latitude.setPlaceholderText("Latitude")
        layout.addWidget(self.input_collection_latitude)
        self.button_save = QPushButton("Salvar", self)
        self.button_save.clicked.connect(self.add_collection)
        layout.addWidget(self.button_save)

    def add_collection(self):
        name = self.input_collection_name.text()
        site_id = self.input_site_id.currentData()
        longitude = self.input_collection_longitude.text()
        latitude = self.input_collection_latitude.text()
        if not name or site_id is None:
            QMessageBox.warning(self, "Erro", "O campo Nome é obrigatório.")
            return
        self.db.insert("Collection", {"site_id": site_id, "name": name, "longitude": longitude, "latitude": latitude})
        self.accept()

class AddSampleDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Adicionar Amostra")
        self.setWindowIcon(QIcon("./icon/iconTray.png"))
        layout = QVBoxLayout(self)
        self.input_sample_collection_id = QComboBox(self)
        self.input_sample_collection_id.setPlaceholderText("Selecione a Coleção")
        layout.addWidget(self.input_sample_collection_id)
        collections = self.db.fetch("Collection", "id, name")
        for col_id, col_name in collections:
            self.input_sample_collection_id.addItem(str(col_name), col_id)
        self.input_sample_name = QLineEdit(self)
        self.input_sample_name.setPlaceholderText("Nome da Amostra")
        layout.addWidget(self.input_sample_name)
        self.input_sample_longitude = QLineEdit(self)
        self.input_sample_longitude.setPlaceholderText("Longitude")
        layout.addWidget(self.input_sample_longitude)
        self.input_sample_latitude = QLineEdit(self)
        self.input_sample_latitude.setPlaceholderText("Latitude")
        layout.addWidget(self.input_sample_latitude)
        self.input_sample_screen_size = QLineEdit(self)
        self.input_sample_screen_size.setPlaceholderText("Tamanho da Tela")
        layout.addWidget(self.input_sample_screen_size)
        button_add = QPushButton("Adicionar", self)
        button_add.clicked.connect(self.add_sample)
        layout.addWidget(button_add)
        button_cancel = QPushButton("Cancelar", self)
        button_cancel.clicked.connect(self.reject)
        layout.addWidget(button_cancel)

    def add_sample(self):
        collection_id = self.input_sample_collection_id.currentData()
        name = self.input_sample_name.text().strip()
        longitude = self.input_sample_longitude.text().strip()
        latitude = self.input_sample_latitude.text().strip()
        screen_size = self.input_sample_screen_size.text().strip()
        if collection_id is None or not name:
            QMessageBox.warning(self, "Erro", "Os campos 'ID da Coleção' e 'Nome' são obrigatórios.")
            return
        self.db.insert("Assemblage", {
            "collection_id": collection_id,
            "name": name,
            "longitude": longitude,
            "latitude": latitude,
            "screenSize": screen_size
        })
        self.accept()

class AddLevelDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Adicionar Nível")
        self.setWindowIcon(QIcon("./icon/iconTray.png"))
        layout = QVBoxLayout(self)

        self.input_unit = QComboBox(self)
        self.input_unit.setPlaceholderText("Selecione a Unidade de Escavação")
        units = self.db.fetch("ExcavationUnit", "id, name")
        for u_id, u_name in units:
            self.input_unit.addItem(str(u_name), u_id)
        layout.addWidget(self.input_unit)

        self.inputs = {}
        campos = [("level", "Nível"), ("start_depth", "Profundidade Inicial"), 
                  ("end_depth", "Profundidade Final"), ("color", "Cor"), 
                  ("texture", "Textura"), ("description", "Descrição")]
        
        for key, placeholder in campos:
            inp = QLineEdit(self)
            inp.setPlaceholderText(placeholder)
            layout.addWidget(inp)
            self.inputs[key] = inp

        btn_save = QPushButton("Salvar")
        btn_save.clicked.connect(self.save_level)
        layout.addWidget(btn_save)

    def save_level(self):
        unit_id = self.input_unit.currentData()
        if unit_id is None:
            QMessageBox.warning(self, "Erro", "Selecione uma unidade.")
            return
        data = {k: v.text() for k, v in self.inputs.items()}
        data["excavation_unit_id"] = unit_id
        self.db.insert("Level", data)
        self.accept()

class AddMaterialDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Adicionar Material")
        self.setWindowIcon(QIcon("./icon/iconTray.png"))
        self.resize(800, 600)
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        content_widget = QWidget()
        grid_layout = QGridLayout(content_widget)
        self.input_fields = {}

        # Filtro de Sítio (Não é salvo no banco, apenas para filtrar unidades)
        label_site = QLabel("Filtrar por Sítio:")
        self.combo_site = QComboBox()
        self.combo_site.setPlaceholderText("Selecione o Sítio")
        self.combo_site.addItem("Todos os Sítios", None)
        sites = self.db.fetch("Site", "id, name")
        for s_id, s_name in sites:
            self.combo_site.addItem(str(s_name), s_id)
        self.combo_site.currentIndexChanged.connect(self.filter_units)
        grid_layout.addWidget(label_site, 0, 0)
        grid_layout.addWidget(self.combo_site, 0, 1)

        field_names = [
            ("excavation_unit_id", "Unidade de Escavação"), ("level_id", "Nível"),
            ("field_serial", "Serial de Campo"), ("lab_serial", "Serial de Laboratório"),
            ("material_type", "Tipo de Material"), ("material_description", "Descrição do Material"),
            ("measurements", "Medidas"), ("weight", "Peso (g)"), ("quantity", "Quantidade"),
            ("x_coord", "X"), ("y_coord", "Y"), ("z_coord", "Z"), ("notes", "Observação"), ("photos", "Fotos"),
            ("user", "Usuário"),
        ]
        for i, (key, placeholder) in enumerate(field_names):
            label = QLabel(placeholder)
            if key == "excavation_unit_id":
                input_field = QComboBox()
                input_field.setPlaceholderText("Selecione a Unidade de Escavação")
                units = self.db.fetch("ExcavationUnit", "id, name")
                for u_id, u_name in units:
                    input_field.addItem(str(u_name), u_id)
                input_field.currentIndexChanged.connect(self.update_levels)
            elif key == "level_id":
                input_field = QComboBox()
                input_field.setPlaceholderText("Selecione o Nível")
            else:
                input_field = QLineEdit()
                input_field.setPlaceholderText(placeholder)
            self.input_fields[key] = input_field
            col = i % 2
            row = (i // 2) + 1 # +1 por causa do filtro de sítio
            grid_layout.addWidget(label, row, col * 2)
            grid_layout.addWidget(input_field, row, col * 2 + 1)
        button_layout = QHBoxLayout()
        button_add = QPushButton("Adicionar")
        button_add.clicked.connect(self.add_material)
        button_layout.addWidget(button_add)
        button_cancel = QPushButton("Cancelar")
        button_cancel.clicked.connect(self.reject)
        button_layout.addWidget(button_cancel)
        layout = QVBoxLayout(self)
        scroll_area.setWidget(content_widget)
        layout.addWidget(scroll_area)
        layout.addLayout(button_layout)

    def filter_units(self):
        """Filtra as unidades de escavação com base no sítio selecionado."""
        site_id = self.combo_site.currentData()
        unit_combo = self.input_fields["excavation_unit_id"]
        unit_combo.blockSignals(True)
        unit_combo.clear()
        
        if site_id is None:
            # Carrega todas as unidades
            units = self.db.fetch("ExcavationUnit", "id, name")
        else:
            # Carrega unidades filtradas pelo sítio
            query = """
                SELECT u.id, u.name 
                FROM ExcavationUnit u 
                JOIN Assemblage a ON u.assemblage_id = a.id 
                JOIN Collection c ON a.collection_id = c.id 
                WHERE c.site_id = ?
            """
            cursor = self.db.conn.execute(query, (site_id,))
            units = cursor.fetchall()

        for u_id, u_name in units:
            unit_combo.addItem(str(u_name), u_id)
        
        unit_combo.blockSignals(False)
        self.update_levels() # Atualiza níveis para a nova lista de unidades (ou seleção vazia)

    def update_levels(self):
        unit_id = self.input_fields["excavation_unit_id"].currentData()
        level_combo = self.input_fields["level_id"]
        level_combo.clear()
        if unit_id:
            levels = self.db.fetch("Level", "id, level", {"excavation_unit_id": unit_id})
            for l_id, l_name in levels:
                level_combo.addItem(str(l_name), l_id)

    def add_material(self):
        material_data = {}
        optional_fields = ["x_coord", "y_coord", "z_coord"]
        for key, field in self.input_fields.items():
            if isinstance(field, QComboBox):
                value = field.currentData()
                if value is None and key not in optional_fields:
                    QMessageBox.warning(self, "Erro", f"O campo '{field.placeholderText()}' é obrigatório.")
                    return
                material_data[key] = value
            else:
                value = field.text().strip()
                if not value and key not in optional_fields:
                    QMessageBox.warning(self, "Erro", f"O campo '{field.placeholderText()}' é obrigatório.")
                    return
                material_data[key] = value
        material_data["uuid"] = str(uuid.uuid4())
        try:
            self.db.insert("Material", material_data)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao inserir no banco de dados: {str(e)}")


class AddUnitDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Adicionar Unidade de Escavação")
        self.setWindowIcon(QIcon("./icon/iconTray.png"))
        layout = QVBoxLayout(self)

        # Selecionar Amostra
        self.input_assemblage = QComboBox(self)
        self.input_assemblage.setPlaceholderText("Selecione a Amostra")
        assemblages = self.db.fetch("Assemblage", "id, name")
        for asm_id, asm_name in assemblages:
            self.input_assemblage.addItem(str(asm_name), asm_id)
        layout.addWidget(self.input_assemblage)

        # Campos básicos
        self.input_name = QLineEdit(self)
        self.input_name.setPlaceholderText("Nome")
        layout.addWidget(self.input_name)

        self.input_type = QLineEdit(self)
        self.input_type.setPlaceholderText("Tipo")
        layout.addWidget(self.input_type)

        self.input_size = QLineEdit(self)
        self.input_size.setPlaceholderText("Tamanho")
        layout.addWidget(self.input_size)

        self.input_lat = QLineEdit(self)
        self.input_lat.setPlaceholderText("Latitude")
        layout.addWidget(self.input_lat)

        self.input_lon = QLineEdit(self)
        self.input_lon.setPlaceholderText("Longitude")
        layout.addWidget(self.input_lon)

        self.input_nivel = QLineEdit(self)
        self.input_nivel.setPlaceholderText("Nível Holótipo")
        layout.addWidget(self.input_nivel)

        self.input_profund_inicio = QLineEdit(self)
        self.input_profund_inicio.setPlaceholderText("Profundidade Inicial")
        layout.addWidget(self.input_profund_inicio)

        self.input_profund_fim = QLineEdit(self)
        self.input_profund_fim.setPlaceholderText("Profundidade Final")
        layout.addWidget(self.input_profund_fim)

        self.input_camada = QLineEdit(self)
        self.input_camada.setPlaceholderText("Camada Geológica")
        layout.addWidget(self.input_camada)

        self.input_metodo_esc = QLineEdit(self)
        self.input_metodo_esc.setPlaceholderText("Método Escavação")
        layout.addWidget(self.input_metodo_esc)

        self.input_metodo_pen = QLineEdit(self)
        self.input_metodo_pen.setPlaceholderText("Método Peneiramento")
        layout.addWidget(self.input_metodo_pen)

        self.input_responsavel = QLineEdit(self)
        self.input_responsavel.setPlaceholderText("Responsável Escavação")
        layout.addWidget(self.input_responsavel)

        self.input_data_inicio = QLineEdit(self)
        self.input_data_inicio.setPlaceholderText("Data Início")
        layout.addWidget(self.input_data_inicio)

        self.input_data_fim = QLineEdit(self)
        self.input_data_fim.setPlaceholderText("Data Conclusão")
        layout.addWidget(self.input_data_fim)

        self.input_fotos = QLineEdit(self)
        self.input_fotos.setPlaceholderText("Fotos Registro")
        layout.addWidget(self.input_fotos)

        self.input_obs = QLineEdit(self)
        self.input_obs.setPlaceholderText("Observação")
        layout.addWidget(self.input_obs)

        button_add = QPushButton("Adicionar")
        button_add.clicked.connect(self.add_unit)
        layout.addWidget(button_add)

        button_cancel = QPushButton("Cancelar")
        button_cancel.clicked.connect(self.reject)
        layout.addWidget(button_cancel)

    def add_unit(self):
        assemblage_id = self.input_assemblage.currentData()
        if assemblage_id is None:
            QMessageBox.warning(self, "Erro", "Selecione a Amostra (Amostra obrigatória).")
            return

        data = {
            "assemblage_id": assemblage_id,
            "name": self.input_name.text().strip(),
            "unit_type": self.input_type.text().strip(),
            "size": self.input_size.text().strip(),
            "latitude": self.input_lat.text().strip(),
            "longitude": self.input_lon.text().strip(),
            "nivel_holotipo": self.input_nivel.text().strip(),
            "profundidade_inicial": self.input_profund_inicio.text().strip(),
            "profundidade_final": self.input_profund_fim.text().strip(),
            "camada_geologica": self.input_camada.text().strip(),
            "metodo_escavacao": self.input_metodo_esc.text().strip(),
            "metodo_peneiramento": self.input_metodo_pen.text().strip(),
            "responsavel_escavacao": self.input_responsavel.text().strip(),
            "data_inicio": self.input_data_inicio.text().strip(),
            "data_conclusao": self.input_data_fim.text().strip(),
            "fotos_registro": self.input_fotos.text().strip(),
            "observacao": self.input_obs.text().strip(),
        }

        try:
            self.db.insert("ExcavationUnit", data)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao inserir Unidade de Escavação: {str(e)}")
