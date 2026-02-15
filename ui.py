import os
from PySide6.QtCore import Qt, QCoreApplication, QSize, QMetaObject
from PySide6.QtGui import QIcon, QFont, QPixmap
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPushButton, QLabel,
    QVBoxLayout, QGridLayout, QStackedWidget, QTextBrowser, QLayout,
    QLineEdit, QTableWidget, QTableWidgetItem, QMessageBox, QDialog, QFileDialog, QSpacerItem, QSizePolicy,
    QHBoxLayout, QScrollArea, QComboBox, QTabWidget, QHeaderView, QFrame, QDoubleSpinBox
)
from database import Database
import dialogs
import io_handlers
import resources_rc
import statistics
import traceback

class Ui_MainWindow(object):
    
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1920, 1080)
        MainWindow.showMaximized()  # Abre a janela maximizada
        self.MainWindow = MainWindow # Referência para aplicar tema

        # Ícone da janela
        icon = QIcon()
        icon.addFile("./icon/iconTray.png", QSize(), QIcon.Mode.Normal, QIcon.State.Off)
        MainWindow.setWindowIcon(icon)
        
        # Definir título da janela
        MainWindow.setWindowTitle("ArchaeoTrack")
        
        # Estilo Moderno e Científico (CSS) and theme definitions
        self.current_lang = "pt"
        self.current_theme = "light"

        # --- Definição de Temas ---
        self.themes = {
            "light": """
            QMainWindow {
                background-color: #f0f2f5;
            }
            QTabWidget::pane {
                border: 1px solid #dcdcdc;
                background: white;
                border-radius: 4px;
            }
            QTabBar::tab {
                background: #e1e4e8;
                border: 1px solid #dcdcdc;
                padding: 10px 20px;
                margin-bottom: 2px;
                border-top-left-radius: 4px;
                border-bottom-left-radius: 4px;
                font-weight: bold;
                color: #555;
            }
            QTabBar::tab:selected {
                background: white;
                border-right-color: white; /* Faz parecer conectado ao painel */
                color: #2c3e50;
                border-left: 4px solid #4a90e2; /* Acento azul */
            }
            QPushButton {
                background-color: #ffffff;
                border: 1px solid #dcdcdc;
                border-radius: 4px;
                padding: 6px 12px;
                color: #333;
            }
            QPushButton:hover {
                background-color: #f5f5f5;
                border-color: #adadad;
            }
            QPushButton:pressed {
                background-color: #e6e6e6;
            }
            QLineEdit {
                border: 1px solid #dcdcdc;
                border-radius: 4px;
                padding: 4px;
                background: white;
            }
            QTableWidget {
                background-color: white;
                gridline-color: #e0e0e0;
                border: 1px solid #dcdcdc;
                selection-background-color: #4a90e2;
                selection-color: white;
            }
            QHeaderView::section {
                background-color: #f8f9fa;
                padding: 6px;
                border: 1px solid #e0e0e0;
                font-weight: bold;
                color: #444;
            }
            QComboBox {
                border: 1px solid #dcdcdc;
                border-radius: 4px;
                padding: 4px;
                background: white;
                color: #333;
            }
            QComboBox:hover {
                border-color: #adadad;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #dcdcdc;
                selection-background-color: #4a90e2;
                selection-color: white;
                background-color: white;
            }
            """,
            "dark": """
            QMainWindow {
                background-color: #1e1e1e;
                color: #e0e0e0;
            }
            QWidget {
                color: #e0e0e0;
            }
            QTabWidget::pane {
                border: 1px solid #444;
                background: #2d2d2d;
                border-radius: 4px;
            }
            QTabBar::tab {
                background: #333;
                border: 1px solid #444;
                padding: 10px 20px;
                margin-bottom: 2px;
                border-top-left-radius: 4px;
                border-bottom-left-radius: 4px;
                font-weight: bold;
                color: #aaa;
            }
            QTabBar::tab:selected {
                background: #2d2d2d;
                border-right-color: #2d2d2d;
                color: #fff;
                border-left: 4px solid #4a90e2;
            }
            QPushButton {
                background-color: #333;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 6px 12px;
                color: #e0e0e0;
            }
            QPushButton:hover {
                background-color: #444;
                border-color: #666;
            }
            QPushButton:pressed {
                background-color: #222;
            }
            QLineEdit {
                border: 1px solid #555;
                border-radius: 4px;
                padding: 4px;
                background: #333;
                color: #fff;
            }
            QTableWidget, QTableView {
                background-color: #2d2d2d;
                alternate-background-color: #262626;
                gridline-color: #444;
                border: 1px solid #444;
                selection-background-color: #4a90e2;
                selection-color: white;
                color: #e0e0e0;
            }
            QHeaderView::section {
                background-color: #333;
                padding: 6px;
                border: 1px solid #444;
                font-weight: bold;
                color: #e0e0e0;
            }
            QComboBox {
                border: 1px solid #555;
                border-radius: 4px;
                padding: 4px;
                background: #333;
                color: #e0e0e0;
            }
            QComboBox:hover {
                border-color: #777;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #555;
                selection-background-color: #4a90e2;
                selection-color: white;
                background-color: #333;
                color: #e0e0e0;
            }
            """
        }

        # Aplica tema inicial
        self.apply_theme(self.current_theme)
        self._populating_tables = False

        MainWindow.setIconSize(QSize(256, 256))

        self.centralwidget = QWidget(MainWindow)
        self.gridLayout_2 = QGridLayout(self.centralwidget)

        self.stackedWidget = QStackedWidget(self.centralwidget)

        spacer = QSpacerItem(20, 40, QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)

        # Mapas para filtro hierárquico
        self.assemblage_to_collection = {}
        self.collection_to_site = {}
        
        #########################################################################################
        # Página principal
        self.Main = QWidget()
        self.verticalLayout_4 = QVBoxLayout(self.Main)
        self.gridLayout = QGridLayout()
        self.gridLayout.setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize)
        self.gridLayout.setVerticalSpacing(7)

        # Ícone
        self.label_2 = QLabel(self.Main)
        self.label_2.setMaximumSize(QSize(200, 200))
        self.label_2.setPixmap(QPixmap("./icon/icon_no_bckgrnd.png"))
        self.label_2.setScaledContents(True)
        self.gridLayout.addWidget(self.label_2, 0, 2, 1, 1)

        # Título
        self.label = QLabel("ArchaeoTrack", self.Main)
        self.label.setMaximumSize(QSize(200, 50))
        font = QFont()
        font.setFamily("PMingLiU-ExtB")
        font.setPointSize(28)
        self.label.setFont(font)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.gridLayout.addWidget(self.label, 1, 2, 1, 1)
        
        # Botão Novo Banco de Dados
        self.button_newDB = QPushButton("Novo Banco de Dados", self.Main)
        self.button_newDB.clicked.connect(self.create_new_database)
        self.gridLayout.addWidget(self.button_newDB, 2, 2, 1, 1)

        # Botão Abrir Banco de Dados
        self.button_openDB = QPushButton("Abrir Banco de Dados", self.Main)
        self.button_openDB.clicked.connect(self.open_database)
        self.gridLayout.addWidget(self.button_openDB, 3, 2, 1, 1)

        # Botão Importar Excel
        self.button_import_xlsx = QPushButton("Criar DB de Excel", self.Main)
        self.button_import_xlsx.clicked.connect(self.import_database_from_excel)
        self.gridLayout.addWidget(self.button_import_xlsx, 4, 2, 1, 1)

        # Botão Configurações
        self.button_settings = QPushButton("Configurações", self.Main)
        self.button_settings.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.Configuracoes))
        self.gridLayout.addWidget(self.button_settings, 5, 2, 1, 1)

        # Botão Sobre
        self.button_about = QPushButton("Sobre", self.Main)
        self.button_about.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.Sobre))
        self.gridLayout.addWidget(self.button_about, 6, 2, 1, 1)

        # Botão Sair
        self.button_exit = QPushButton("Sair", self.Main)
        self.button_exit.clicked.connect(MainWindow.close)
        self.gridLayout.addWidget(self.button_exit, 7, 2, 1, 1)
        
        # spacer
        self.gridLayout.addItem(spacer, 8, 2, 1, 1)          
        self.gridLayout.addItem(spacer, 9, 2, 1, 1)          

        self.verticalLayout_4.addLayout(self.gridLayout)
        self.stackedWidget.addWidget(self.Main)

        #########################################################################################
        # Página Configurações
        self.Configuracoes = QWidget()
        self.layout_config = QVBoxLayout(self.Configuracoes)
        self.layout_config.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.label_config_title = QLabel("Configurações", self.Configuracoes)
        self.label_config_title.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        self.label_config_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout_config.addWidget(self.label_config_title)

        # Container para opções
        self.config_container = QWidget()
        self.config_grid = QGridLayout(self.config_container)
        self.config_grid.setSpacing(20)

        # Idioma
        self.label_lang = QLabel("Idioma / Language:", self.Configuracoes)
        self.combo_lang = QComboBox(self.Configuracoes)
        self.combo_lang.addItem("Português", "pt")
        self.combo_lang.addItem("English", "en")
        self.combo_lang.addItem("Español", "es")
        self.combo_lang.currentIndexChanged.connect(self.change_language)
        self.config_grid.addWidget(self.label_lang, 0, 0)
        self.config_grid.addWidget(self.combo_lang, 0, 1)

        # Tema
        self.label_theme = QLabel("Tema / Theme:", self.Configuracoes)
        self.combo_theme = QComboBox(self.Configuracoes)
        self.combo_theme.addItem("Light Mode", "light")
        self.combo_theme.addItem("Dark Mode", "dark")
        self.combo_theme.currentIndexChanged.connect(self.change_theme)
        self.config_grid.addWidget(self.label_theme, 1, 0)
        self.config_grid.addWidget(self.combo_theme, 1, 1)

        self.layout_config.addWidget(self.config_container)

        # Botão Voltar
        self.btn_config_back = QPushButton("Voltar", self.Configuracoes)
        self.btn_config_back.setFixedSize(200, 40)
        self.btn_config_back.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.Main))
        self.layout_config.addWidget(self.btn_config_back)

        self.stackedWidget.addWidget(self.Configuracoes)

        #########################################################################################
        # Página Sobre
        self.Sobre = QWidget()
        self.verticalLayout = QVBoxLayout(self.Sobre)     
        
        self.labelSobre = QLabel("ArchaeoTrack", self.Main)
        self.labelSobre.setMaximumSize(QSize(200, 50))
        font = QFont()
        font.setFamily("PMingLiU-ExtB")
        font.setPointSize(26)
        self.labelSobre.setFont(font)
        self.labelSobre.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.verticalLayout.addWidget(self.labelSobre)
        
        # Botão Voltar
        self.sobre_Voltar = QPushButton("Voltar", self.Sobre)
        self.sobre_Voltar.setFixedSize(200, 25)
        self.sobre_Voltar.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.Main))
        self.verticalLayout.addWidget(self.sobre_Voltar)

        self.stackedWidget.addWidget(self.Sobre)
        self.gridLayout_2.addWidget(self.stackedWidget, 0, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)

        self.stackedWidget.setCurrentWidget(self.Main)
        QMetaObject.connectSlotsByName(MainWindow)
        
        self.textBrowser = QTextBrowser(self.Sobre)
        self.textBrowser.setHtml("""
        <h2 data-start="66" data-end="92"><strong data-start="69" data-end="90">Sobre o ArchaeoTrack</strong></h2>
        <p data-start="94" data-end="371">O <strong data-start="96" data-end="109">ArchaeoTrack</strong> &eacute; um software desenvolvido para facilitar o gerenciamento e an&aacute;lise de dados em zooarqueologia. Ele permite a cataloga&ccedil;&atilde;o, visualiza&ccedil;&atilde;o e manipula&ccedil;&atilde;o de informa&ccedil;&otilde;es sobre vest&iacute;gios &oacute;sseos, auxiliando pesquisadores na organiza&ccedil;&atilde;o e interpreta&ccedil;&atilde;o de seus dados.</p>
        <h3 data-start="373" data-end="403"><strong data-start="377" data-end="401">Principais Recursos:</strong></h3>
        <p data-start="404" data-end="725">✔ Interface intuitiva para inser&ccedil;&atilde;o e consulta de dados</p>
        <div class="markdown-heading" dir="auto">
        <h3 data-start="373" data-end="403"><strong data-start="377" data-end="401">TO DO:</strong></h3>
        <a id="user-content-to-do" class="anchor" href="https://github.com/geraldopmj/ArchaeoTrack#to-do" aria-label="Permalink: TO DO"></a></div>
        <ul dir="auto">
        <li>export to csv</li>
        <li>export to pdf (Ficha de esp&eacute;cime)</li>
        <li>add Users, manager and permission</li>
        <li>add Manager panel to check user's performance</li>
        <li>Estat&iacute;sticas</li>
        <li>Connection to Cloud</li>
        <li>Criptography and change file extension to bntk</li>
        </ul>
        <h3 data-start="1011" data-end="1038"><strong data-start="1015" data-end="1036">Contato e Suporte</strong></h3>
        <p data-start="1039" data-end="1179">Para d&uacute;vidas, sugest&otilde;es ou suporte, entre em contato pelo e-mail geraldo.pmj@gmail.com ou acesse nosso reposit&oacute;rio no GitHub: [seu link do GitHub].</p>
        <h3 data-start="727" data-end="762"><strong data-start="731" data-end="760">Desenvolvimento e Licen&ccedil;a</strong></h3>
        <p data-start="763" data-end="1009">O ArchaeoTrack foi criado para atender &agrave;s necessidades de profissionais da zooarqueologia, proporcionando uma ferramenta eficiente e acess&iacute;vel. O software segue a licen&ccedil;a GNU AGPL-3.0, permitindo seu uso conforme os termos estabelecidos abaixo:</p>
        <p>GNU AFFERO GENERAL PUBLIC LICENSE<br />Version 3, 19 November 2007</p>
        <p>Copyright (C) 2007 Free Software Foundation, Inc. &lt;https://fsf.org/&gt;<br />Everyone is permitted to copy and distribute verbatim copies<br />of this license document, but changing it is not allowed.</p>
        <p>Preamble</p>
        <p>The GNU Affero General Public License is a free, copyleft license for<br />software and other kinds of works, specifically designed to ensure<br />cooperation with the community in the case of network server software.</p>
        <p>The licenses for most software and other practical works are designed<br />to take away your freedom to share and change the works. By contrast,<br />our General Public Licenses are intended to guarantee your freedom to<br />share and change all versions of a program--to make sure it remains free<br />software for all its users.</p>
        <p>When we speak of free software, we are referring to freedom, not<br />price. Our General Public Licenses are designed to make sure that you<br />have the freedom to distribute copies of free software (and charge for<br />them if you wish), that you receive source code or can get it if you<br />want it, that you can change the software or use pieces of it in new<br />free programs, and that you know you can do these things.</p>
        <p>Developers that use our General Public Licenses protect your rights<br />with two steps: (1) assert copyright on the software, and (2) offer<br />you this License which gives you legal permission to copy, distribute<br />and/or modify the software.</p>
        <p>A secondary benefit of defending all users' freedom is that<br />improvements made in alternate versions of the program, if they<br />receive widespread use, become available for other developers to<br />incorporate. Many developers of free software are heartened and<br />encouraged by the resulting cooperation. However, in the case of<br />software used on network servers, this result may fail to come about.<br />The GNU General Public License permits making a modified version and<br />letting the public access it on a server without ever releasing its<br />source code to the public.</p>
        <p>The GNU Affero General Public License is designed specifically to<br />ensure that, in such cases, the modified source code becomes available<br />to the community. It requires the operator of a network server to<br />provide the source code of the modified version running there to the<br />users of that server. Therefore, public use of a modified version, on<br />a publicly accessible server, gives the public access to the source<br />code of the modified version.</p>
        <p>An older license, called the Affero General Public License and<br />published by Affero, was designed to accomplish similar goals. This is<br />a different license, not a version of the Affero GPL, but Affero has<br />released a new version of the Affero GPL which permits relicensing under<br />this license.</p>
        <p>The precise terms and conditions for copying, distribution and<br />modification follow.</p>
        <p>TERMS AND CONDITIONS</p>
        <p>0. Definitions.</p>
        <p>"This License" refers to version 3 of the GNU Affero General Public License.</p>
        <p>"Copyright" also means copyright-like laws that apply to other kinds of<br />works, such as semiconductor masks.</p>
        <p>"The Program" refers to any copyrightable work licensed under this<br />License. Each licensee is addressed as "you". "Licensees" and<br />"recipients" may be individuals or organizations.</p>
        <p>To "modify" a work means to copy from or adapt all or part of the work<br />in a fashion requiring copyright permission, other than the making of an<br />exact copy. The resulting work is called a "modified version" of the<br />earlier work or a work "based on" the earlier work.</p>
        <p>A "covered work" means either the unmodified Program or a work based<br />on the Program.</p>
        <p>To "propagate" a work means to do anything with it that, without<br />permission, would make you directly or secondarily liable for<br />infringement under applicable copyright law, except executing it on a<br />computer or modifying a private copy. Propagation includes copying,<br />distribution (with or without modification), making available to the<br />public, and in some countries other activities as well.</p>
        <p>To "convey" a work means any kind of propagation that enables other<br />parties to make or receive copies. Mere interaction with a user through<br />a computer network, with no transfer of a copy, is not conveying.</p>
        <p>An interactive user interface displays "Appropriate Legal Notices"<br />to the extent that it includes a convenient and prominently visible<br />feature that (1) displays an appropriate copyright notice, and (2)<br />tells the user that there is no warranty for the work (except to the<br />extent that warranties are provided), that licensees may convey the<br />work under this License, and how to view a copy of this License. If<br />the interface presents a list of user commands or options, such as a<br />menu, a prominent item in the list meets this criterion.</p>
        <p>1. Source Code.</p>
        <p>The "source code" for a work means the preferred form of the work<br />for making modifications to it. "Object code" means any non-source<br />form of a work.</p>
        <p>A "Standard Interface" means an interface that either is an official<br />standard defined by a recognized standards body, or, in the case of<br />interfaces specified for a particular programming language, one that<br />is widely used among developers working in that language.</p>
        <p>The "System Libraries" of an executable work include anything, other<br />than the work as a whole, that (a) is included in the normal form of<br />packaging a Major Component, but which is not part of that Major<br />Component, and (b) serves only to enable use of the work with that<br />Major Component, or to implement a Standard Interface for which an<br />implementation is available to the public in source code form. A<br />"Major Component", in this context, means a major essential component<br />(kernel, window system, and so on) of the specific operating system<br />(if any) on which the executable work runs, or a compiler used to<br />produce the work, or an object code interpreter used to run it.</p>
        <p>The "Corresponding Source" for a work in object code form means all<br />the source code needed to generate, install, and (for an executable<br />work) run the object code and to modify the work, including scripts to<br />control those activities. However, it does not include the work's<br />System Libraries, or general-purpose tools or generally available free<br />programs which are used unmodified in performing those activities but<br />which are not part of the work. For example, Corresponding Source<br />includes interface definition files associated with source files for<br />the work, and the source code for shared libraries and dynamically<br />linked subprograms that the work is specifically designed to require,<br />such as by intimate data communication or control flow between those<br />subprograms and other parts of the work.</p>
        <p>The Corresponding Source need not include anything that users<br />can regenerate automatically from other parts of the Corresponding<br />Source.</p>
        <p>The Corresponding Source for a work in source code form is that<br />same work.</p>
        <p>2. Basic Permissions.</p>
        <p>All rights granted under this License are granted for the term of<br />copyright on the Program, and are irrevocable provided the stated<br />conditions are met. This License explicitly affirms your unlimited<br />permission to run the unmodified Program. The output from running a<br />covered work is covered by this License only if the output, given its<br />content, constitutes a covered work. This License acknowledges your<br />rights of fair use or other equivalent, as provided by copyright law.</p>
        <p>You may make, run and propagate covered works that you do not<br />convey, without conditions so long as your license otherwise remains<br />in force. You may convey covered works to others for the sole purpose<br />of having them make modifications exclusively for you, or provide you<br />with facilities for running those works, provided that you comply with<br />the terms of this License in conveying all material for which you do<br />not control copyright. Those thus making or running the covered works<br />for you must do so exclusively on your behalf, under your direction<br />and control, on terms that prohibit them from making any copies of<br />your copyrighted material outside their relationship with you.</p>
        <p>Conveying under any other circumstances is permitted solely under<br />the conditions stated below. Sublicensing is not allowed; section 10<br />makes it unnecessary.</p>
        <p>3. Protecting Users' Legal Rights From Anti-Circumvention Law.</p>
        <p>No covered work shall be deemed part of an effective technological<br />measure under any applicable law fulfilling obligations under article<br />11 of the WIPO copyright treaty adopted on 20 December 1996, or<br />similar laws prohibiting or restricting circumvention of such<br />measures.</p>
        <p>When you convey a covered work, you waive any legal power to forbid<br />circumvention of technological measures to the extent such circumvention<br />is effected by exercising rights under this License with respect to<br />the covered work, and you disclaim any intention to limit operation or<br />modification of the work as a means of enforcing, against the work's<br />users, your or third parties' legal rights to forbid circumvention of<br />technological measures.</p>
        <p>4. Conveying Verbatim Copies.</p>
        <p>You may convey verbatim copies of the Program's source code as you<br />receive it, in any medium, provided that you conspicuously and<br />appropriately publish on each copy an appropriate copyright notice;<br />keep intact all notices stating that this License and any<br />non-permissive terms added in accord with section 7 apply to the code;<br />keep intact all notices of the absence of any warranty; and give all<br />recipients a copy of this License along with the Program.</p>
        <p>You may charge any price or no price for each copy that you convey,<br />and you may offer support or warranty protection for a fee.</p>
        <p>5. Conveying Modified Source Versions.</p>
        <p>You may convey a work based on the Program, or the modifications to<br />produce it from the Program, in the form of source code under the<br />terms of section 4, provided that you also meet all of these conditions:</p>
        <p>a) The work must carry prominent notices stating that you modified<br />it, and giving a relevant date.</p>
        <p>b) The work must carry prominent notices stating that it is<br />released under this License and any conditions added under section<br />7. This requirement modifies the requirement in section 4 to<br />"keep intact all notices".</p>
        <p>c) You must license the entire work, as a whole, under this<br />License to anyone who comes into possession of a copy. This<br />License will therefore apply, along with any applicable section 7<br />additional terms, to the whole of the work, and all its parts,<br />regardless of how they are packaged. This License gives no<br />permission to license the work in any other way, but it does not<br />invalidate such permission if you have separately received it.</p>
        <p>d) If the work has interactive user interfaces, each must display<br />Appropriate Legal Notices; however, if the Program has interactive<br />interfaces that do not display Appropriate Legal Notices, your<br />work need not make them do so.</p>
        <p>A compilation of a covered work with other separate and independent<br />works, which are not by their nature extensions of the covered work,<br />and which are not combined with it such as to form a larger program,<br />in or on a volume of a storage or distribution medium, is called an<br />"aggregate" if the compilation and its resulting copyright are not<br />used to limit the access or legal rights of the compilation's users<br />beyond what the individual works permit. Inclusion of a covered work<br />in an aggregate does not cause this License to apply to the other<br />parts of the aggregate.</p>
        <p>6. Conveying Non-Source Forms.</p>
        <p>You may convey a covered work in object code form under the terms<br />of sections 4 and 5, provided that you also convey the<br />machine-readable Corresponding Source under the terms of this License,<br />in one of these ways:</p>
        <p>a) Convey the object code in, or embodied in, a physical product<br />(including a physical distribution medium), accompanied by the<br />Corresponding Source fixed on a durable physical medium<br />customarily used for software interchange.</p>
        <p>b) Convey the object code in, or embodied in, a physical product<br />(including a physical distribution medium), accompanied by a<br />written offer, valid for at least three years and valid for as<br />long as you offer spare parts or customer support for that product<br />model, to give anyone who possesses the object code either (1) a<br />copy of the Corresponding Source for all the software in the<br />product that is covered by this License, on a durable physical<br />medium customarily used for software interchange, for a price no<br />more than your reasonable cost of physically performing this<br />conveying of source, or (2) access to copy the<br />Corresponding Source from a network server at no charge.</p>
        <p>c) Convey individual copies of the object code with a copy of the<br />written offer to provide the Corresponding Source. This<br />alternative is allowed only occasionally and noncommercially, and<br />only if you received the object code with such an offer, in accord<br />with subsection 6b.</p>
        <p>d) Convey the object code by offering access from a designated<br />place (gratis or for a charge), and offer equivalent access to the<br />Corresponding Source in the same way through the same place at no<br />further charge. You need not require recipients to copy the<br />Corresponding Source along with the object code. If the place to<br />copy the object code is a network server, the Corresponding Source<br />may be on a different server (operated by you or a third party)<br />that supports equivalent copying facilities, provided you maintain<br />clear directions next to the object code saying where to find the<br />Corresponding Source. Regardless of what server hosts the<br />Corresponding Source, you remain obligated to ensure that it is<br />available for as long as needed to satisfy these requirements.</p>
        <p>e) Convey the object code using peer-to-peer transmission, provided<br />you inform other peers where the object code and Corresponding<br />Source of the work are being offered to the general public at no<br />charge under subsection 6d.</p>
        <p>A separable portion of the object code, whose source code is excluded<br />from the Corresponding Source as a System Library, need not be<br />included in conveying the object code work.</p>
        <p>A "User Product" is either (1) a "consumer product", which means any<br />tangible personal property which is normally used for personal, family,<br />or household purposes, or (2) anything designed or sold for incorporation<br />into a dwelling. In determining whether a product is a consumer product,<br />doubtful cases shall be resolved in favor of coverage. For a particular<br />product received by a particular user, "normally used" refers to a<br />typical or common use of that class of product, regardless of the status<br />of the particular user or of the way in which the particular user<br />actually uses, or expects or is expected to use, the product. A product<br />is a consumer product regardless of whether the product has substantial<br />commercial, industrial or non-consumer uses, unless such uses represent<br />the only significant mode of use of the product.</p>
        <p>"Installation Information" for a User Product means any methods,<br />procedures, authorization keys, or other information required to install<br />and execute modified versions of a covered work in that User Product from<br />a modified version of its Corresponding Source. The information must<br />suffice to ensure that the continued functioning of the modified object<br />code is in no case prevented or interfered with solely because<br />modification has been made.</p>
        <p>If you convey an object code work under this section in, or with, or<br />specifically for use in, a User Product, and the conveying occurs as<br />part of a transaction in which the right of possession and use of the<br />User Product is transferred to the recipient in perpetuity or for a<br />fixed term (regardless of how the transaction is characterized), the<br />Corresponding Source conveyed under this section must be accompanied<br />by the Installation Information. But this requirement does not apply<br />if neither you nor any third party retains the ability to install<br />modified object code on the User Product (for example, the work has<br />been installed in ROM).</p>
        <p>The requirement to provide Installation Information does not include a<br />requirement to continue to provide support service, warranty, or updates<br />for a work that has been modified or installed by the recipient, or for<br />the User Product in which it has been modified or installed. Access to a<br />network may be denied when the modification itself materially and<br />adversely affects the operation of the network or violates the rules and<br />protocols for communication across the network.</p>
        <p>Corresponding Source conveyed, and Installation Information provided,<br />in accord with this section must be in a format that is publicly<br />documented (and with an implementation available to the public in<br />source code form), and must require no special password or key for<br />unpacking, reading or copying.</p>
        <p>7. Additional Terms.</p>
        <p>"Additional permissions" are terms that supplement the terms of this<br />License by making exceptions from one or more of its conditions.<br />Additional permissions that are applicable to the entire Program shall<br />be treated as though they were included in this License, to the extent<br />that they are valid under applicable law. If additional permissions<br />apply only to part of the Program, that part may be used separately<br />under those permissions, but the entire Program remains governed by<br />this License without regard to the additional permissions.</p>
        <p>When you convey a copy of a covered work, you may at your option<br />remove any additional permissions from that copy, or from any part of<br />it. (Additional permissions may be written to require their own<br />removal in certain cases when you modify the work.) You may place<br />additional permissions on material, added by you to a covered work,<br />for which you have or can give appropriate copyright permission.</p>
        <p>Notwithstanding any other provision of this License, for material you<br />add to a covered work, you may (if authorized by the copyright holders of<br />that material) supplement the terms of this License with terms:</p>
        <p>a) Disclaiming warranty or limiting liability differently from the<br />terms of sections 15 and 16 of this License; or</p>
        <p>b) Requiring preservation of specified reasonable legal notices or<br />author attributions in that material or in the Appropriate Legal<br />Notices displayed by works containing it; or</p>
        <p>c) Prohibiting misrepresentation of the origin of that material, or<br />requiring that modified versions of such material be marked in<br />reasonable ways as different from the original version; or</p>
        <p>d) Limiting the use for publicity purposes of names of licensors or<br />authors of the material; or</p>
        <p>e) Declining to grant rights under trademark law for use of some<br />trade names, trademarks, or service marks; or</p>
        <p>f) Requiring indemnification of licensors and authors of that<br />material by anyone who conveys the material (or modified versions of<br />it) with contractual assumptions of liability to the recipient, for<br />any liability that these contractual assumptions directly impose on<br />those licensors and authors.</p>
        <p>All other non-permissive additional terms are considered "further<br />restrictions" within the meaning of section 10. If the Program as you<br />received it, or any part of it, contains a notice stating that it is<br />governed by this License along with a term that is a further<br />restriction, you may remove that term. If a license document contains<br />a further restriction but permits relicensing or conveying under this<br />License, you may add to a covered work material governed by the terms<br />of that license document, provided that the further restriction does<br />not survive such relicensing or conveying.</p>
        <p>If you add terms to a covered work in accord with this section, you<br />must place, in the relevant source files, a statement of the<br />additional terms that apply to those files, or a notice indicating<br />where to find the applicable terms.</p>
        <p>Additional terms, permissive or non-permissive, may be stated in the<br />form of a separately written license, or stated as exceptions;<br />the above requirements apply either way.</p>
        <p>8. Termination.</p>
        <p>You may not propagate or modify a covered work except as expressly<br />provided under this License. Any attempt otherwise to propagate or<br />modify it is void, and will automatically terminate your rights under<br />this License (including any patent licenses granted under the third<br />paragraph of section 11).</p>
        <p>However, if you cease all violation of this License, then your<br />license from a particular copyright holder is reinstated (a)<br />provisionally, unless and until the copyright holder explicitly and<br />finally terminates your license, and (b) permanently, if the copyright<br />holder fails to notify you of the violation by some reasonable means<br />prior to 60 days after the cessation.</p>
        <p>Moreover, your license from a particular copyright holder is<br />reinstated permanently if the copyright holder notifies you of the<br />violation by some reasonable means, this is the first time you have<br />received notice of violation of this License (for any work) from that<br />copyright holder, and you cure the violation prior to 30 days after<br />your receipt of the notice.</p>
        <p>Termination of your rights under this section does not terminate the<br />licenses of parties who have received copies or rights from you under<br />this License. If your rights have been terminated and not permanently<br />reinstated, you do not qualify to receive new licenses for the same<br />material under section 10.</p>
        <p>9. Acceptance Not Required for Having Copies.</p>
        <p>You are not required to accept this License in order to receive or<br />run a copy of the Program. Ancillary propagation of a covered work<br />occurring solely as a consequence of using peer-to-peer transmission<br />to receive a copy likewise does not require acceptance. However,<br />nothing other than this License grants you permission to propagate or<br />modify any covered work. These actions infringe copyright if you do<br />not accept this License. Therefore, by modifying or propagating a<br />covered work, you indicate your acceptance of this License to do so.</p>
        <p>10. Automatic Licensing of Downstream Recipients.</p>
        <p>Each time you convey a covered work, the recipient automatically<br />receives a license from the original licensors, to run, modify and<br />propagate that work, subject to this License. You are not responsible<br />for enforcing compliance by third parties with this License.</p>
        <p>An "entity transaction" is a transaction transferring control of an<br />organization, or substantially all assets of one, or subdividing an<br />organization, or merging organizations. If propagation of a covered<br />work results from an entity transaction, each party to that<br />transaction who receives a copy of the work also receives whatever<br />licenses to the work the party's predecessor in interest had or could<br />give under the previous paragraph, plus a right to possession of the<br />Corresponding Source of the work from the predecessor in interest, if<br />the predecessor has it or can get it with reasonable efforts.</p>
        <p>You may not impose any further restrictions on the exercise of the<br />rights granted or affirmed under this License. For example, you may<br />not impose a license fee, royalty, or other charge for exercise of<br />rights granted under this License, and you may not initiate litigation<br />(including a cross-claim or counterclaim in a lawsuit) alleging that<br />any patent claim is infringed by making, using, selling, offering for<br />sale, or importing the Program or any portion of it.</p>
        <p>11. Patents.</p>
        <p>A "contributor" is a copyright holder who authorizes use under this<br />License of the Program or a work on which the Program is based. The<br />work thus licensed is called the contributor's "contributor version".</p>
        <p>A contributor's "essential patent claims" are all patent claims<br />owned or controlled by the contributor, whether already acquired or<br />hereafter acquired, that would be infringed by some manner, permitted<br />by this License, of making, using, or selling its contributor version,<br />but do not include claims that would be infringed only as a<br />consequence of further modification of the contributor version. For<br />purposes of this definition, "control" includes the right to grant<br />patent sublicenses in a manner consistent with the requirements of<br />this License.</p>
        <p>Each contributor grants you a non-exclusive, worldwide, royalty-free<br />patent license under the contributor's essential patent claims, to<br />make, use, sell, offer for sale, import and otherwise run, modify and<br />propagate the contents of its contributor version.</p>
        <p>In the following three paragraphs, a "patent license" is any express<br />agreement or commitment, however denominated, not to enforce a patent<br />(such as an express permission to practice a patent or covenant not to<br />sue for patent infringement). To "grant" such a patent license to a<br />party means to make such an agreement or commitment not to enforce a<br />patent against the party.</p>
        <p>If you convey a covered work, knowingly relying on a patent license,<br />and the Corresponding Source of the work is not available for anyone<br />to copy, free of charge and under the terms of this License, through a<br />publicly available network server or other readily accessible means,<br />then you must either (1) cause the Corresponding Source to be so<br />available, or (2) arrange to deprive yourself of the benefit of the<br />patent license for this particular work, or (3) arrange, in a manner<br />consistent with the requirements of this License, to extend the patent<br />license to downstream recipients. "Knowingly relying" means you have<br />actual knowledge that, but for the patent license, your conveying the<br />covered work in a country, or your recipient's use of the covered work<br />in a country, would infringe one or more identifiable patents in that<br />country that you have reason to believe are valid.</p>
        <p>If, pursuant to or in connection with a single transaction or<br />arrangement, you convey, or propagate by procuring conveyance of, a<br />covered work, and grant a patent license to some of the parties<br />receiving the covered work authorizing them to use, propagate, modify<br />or convey a specific copy of the covered work, then the patent license<br />you grant is automatically extended to all recipients of the covered<br />work and works based on it.</p>
        <p>A patent license is "discriminatory" if it does not include within<br />the scope of its coverage, prohibits the exercise of, or is<br />conditioned on the non-exercise of one or more of the rights that are<br />specifically granted under this License. You may not convey a covered<br />work if you are a party to an arrangement with a third party that is<br />in the business of distributing software, under which you make payment<br />to the third party based on the extent of your activity of conveying<br />the work, and under which the third party grants, to any of the<br />parties who would receive the covered work from you, a discriminatory<br />patent license (a) in connection with copies of the covered work<br />conveyed by you (or copies made from those copies), or (b) primarily<br />for and in connection with specific products or compilations that<br />contain the covered work, unless you entered into that arrangement,<br />or that patent license was granted, prior to 28 March 2007.</p>
        <p>Nothing in this License shall be construed as excluding or limiting<br />any implied license or other defenses to infringement that may<br />otherwise be available to you under applicable patent law.</p>
        <p>12. No Surrender of Others' Freedom.</p>
        <p>If conditions are imposed on you (whether by court order, agreement or<br />otherwise) that contradict the conditions of this License, they do not<br />excuse you from the conditions of this License. If you cannot convey a<br />covered work so as to satisfy simultaneously your obligations under this<br />License and any other pertinent obligations, then as a consequence you may<br />not convey it at all. For example, if you agree to terms that obligate you<br />to collect a royalty for further conveying from those to whom you convey<br />the Program, the only way you could satisfy both those terms and this<br />License would be to refrain entirely from conveying the Program.</p>
        <p>13. Remote Network Interaction; Use with the GNU General Public License.</p>
        <p>Notwithstanding any other provision of this License, if you modify the<br />Program, your modified version must prominently offer all users<br />interacting with it remotely through a computer network (if your version<br />supports such interaction) an opportunity to receive the Corresponding<br />Source of your version by providing access to the Corresponding Source<br />from a network server at no charge, through some standard or customary<br />means of facilitating copying of software. This Corresponding Source<br />shall include the Corresponding Source for any work covered by version 3<br />of the GNU General Public License that is incorporated pursuant to the<br />following paragraph.</p>
        <p>Notwithstanding any other provision of this License, you have<br />permission to link or combine any covered work with a work licensed<br />under version 3 of the GNU General Public License into a single<br />combined work, and to convey the resulting work. The terms of this<br />License will continue to apply to the part which is the covered work,<br />but the work with which it is combined will remain governed by version<br />3 of the GNU General Public License.</p>
        <p>14. Revised Versions of this License.</p>
        <p>The Free Software Foundation may publish revised and/or new versions of<br />the GNU Affero General Public License from time to time. Such new versions<br />will be similar in spirit to the present version, but may differ in detail to<br />address new problems or concerns.</p>
        <p>Each version is given a distinguishing version number. If the<br />Program specifies that a certain numbered version of the GNU Affero General<br />Public License "or any later version" applies to it, you have the<br />option of following the terms and conditions either of that numbered<br />version or of any later version published by the Free Software<br />Foundation. If the Program does not specify a version number of the<br />GNU Affero General Public License, you may choose any version ever published<br />by the Free Software Foundation.</p>
        <p>If the Program specifies that a proxy can decide which future<br />versions of the GNU Affero General Public License can be used, that proxy's<br />public statement of acceptance of a version permanently authorizes you<br />to choose that version for the Program.</p>
        <p>Later license versions may give you additional or different<br />permissions. However, no additional obligations are imposed on any<br />author or copyright holder as a result of your choosing to follow a<br />later version.</p>
        <p>15. Disclaimer of Warranty.</p>
        <p>THERE IS NO WARRANTY FOR THE PROGRAM, TO THE EXTENT PERMITTED BY<br />APPLICABLE LAW. EXCEPT WHEN OTHERWISE STATED IN WRITING THE COPYRIGHT<br />HOLDERS AND/OR OTHER PARTIES PROVIDE THE PROGRAM "AS IS" WITHOUT WARRANTY<br />OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO,<br />THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR<br />PURPOSE. THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE PROGRAM<br />IS WITH YOU. SHOULD THE PROGRAM PROVE DEFECTIVE, YOU ASSUME THE COST OF<br />ALL NECESSARY SERVICING, REPAIR OR CORRECTION.</p>
        <p>16. Limitation of Liability.</p>
        <p>IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN WRITING<br />WILL ANY COPYRIGHT HOLDER, OR ANY OTHER PARTY WHO MODIFIES AND/OR CONVEYS<br />THE PROGRAM AS PERMITTED ABOVE, BE LIABLE TO YOU FOR DAMAGES, INCLUDING ANY<br />GENERAL, SPECIAL, INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING OUT OF THE<br />USE OR INABILITY TO USE THE PROGRAM (INCLUDING BUT NOT LIMITED TO LOSS OF<br />DATA OR DATA BEING RENDERED INACCURATE OR LOSSES SUSTAINED BY YOU OR THIRD<br />PARTIES OR A FAILURE OF THE PROGRAM TO OPERATE WITH ANY OTHER PROGRAMS),<br />EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN ADVISED OF THE POSSIBILITY OF<br />SUCH DAMAGES.</p>
        <p>17. Interpretation of Sections 15 and 16.</p>
        <p>If the disclaimer of warranty and limitation of liability provided<br />above cannot be given local legal effect according to their terms,<br />reviewing courts shall apply local law that most closely approximates<br />an absolute waiver of all civil liability in connection with the<br />Program, unless a warranty or assumption of liability accompanies a<br />copy of the Program in return for a fee.</p>
        <p>END OF TERMS AND CONDITIONS</p>
        <p>How to Apply These Terms to Your New Programs</p>
        <p>If you develop a new program, and you want it to be of the greatest<br />possible use to the public, the best way to achieve this is to make it<br />free software which everyone can redistribute and change under these terms.</p>
        <p>To do so, attach the following notices to the program. It is safest<br />to attach them to the start of each source file to most effectively<br />state the exclusion of warranty; and each file should have at least<br />the "copyright" line and a pointer to where the full notice is found.</p>
        <p>&lt;one line to give the program's name and a brief idea of what it does.&gt;<br />Copyright (C) &lt;year&gt; &lt;name of author&gt;</p>
        <p>This program is free software: you can redistribute it and/or modify<br />it under the terms of the GNU Affero General Public License as published<br />by the Free Software Foundation, either version 3 of the License, or<br />(at your option) any later version.</p>
        <p>This program is distributed in the hope that it will be useful,<br />but WITHOUT ANY WARRANTY; without even the implied warranty of<br />MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the<br />GNU Affero General Public License for more details.</p>
        <p>You should have received a copy of the GNU Affero General Public License<br />along with this program. If not, see &lt;https://www.gnu.org/licenses/&gt;.</p>
        <p>Also add information on how to contact you by electronic and paper mail.</p>
        <p>If your software can interact with users remotely through a computer<br />network, you should also make sure that it provides a way for users to<br />get its source. For example, if your program is a web application, its<br />interface could display a "Source" link that leads users to an archive<br />of the code. There are many ways you could offer source, and different<br />solutions will be better for different programs; see section 13 for the<br />specific requirements.</p>
        <p>You should also get your employer (if you work as a programmer) or school,<br />if any, to sign a "copyright disclaimer" for the program, if necessary.<br />For more information on this, and how to apply and follow the GNU AGPL, see<br />&lt;https://www.gnu.org/licenses/&gt;.</p>
        <p data-start="1039" data-end="1179">&nbsp;</p>
        """)
        self.verticalLayout.addWidget(self.textBrowser)
        
        
        #########################################################################################
        # Página Dashboard (Contém as Abas e Toolbar Global)
        self.DashboardPage = QWidget()
        self.dashboard_layout = QVBoxLayout(self.DashboardPage)
        self.dashboard_layout.setContentsMargins(10, 10, 10, 10)
        self.dashboard_layout.setSpacing(10)

        # --- Toolbar Global (Topo) ---
        self.global_toolbar = QHBoxLayout()
        
        # Botão Fechar Banco
        self.btn_close_db = QPushButton("Fechar Banco", self.DashboardPage)
        #self.btn_close_db.setIcon(QIcon("./icon/iconTray.png")) # Placeholder icon
        self.btn_close_db.clicked.connect(self.close_database)
        self.global_toolbar.addWidget(self.btn_close_db)

        self.global_toolbar.addStretch() # Espaço flexível

        # Botões de Exportação
        self.btn_export_pdf = QPushButton("Gerar PDF (Materiais)", self.DashboardPage)
        self.btn_export_pdf.clicked.connect(self.export_pdf)
        self.global_toolbar.addWidget(self.btn_export_pdf)

        self.btn_export_unit_pdf = QPushButton("Gerar PDF (Unidades)", self.DashboardPage)
        self.btn_export_unit_pdf.clicked.connect(self.export_units_pdf)
        self.global_toolbar.addWidget(self.btn_export_unit_pdf)

        self.btn_export_xlsx = QPushButton("Exportar Excel Completo", self.DashboardPage)
        self.btn_export_xlsx.clicked.connect(self.export_xlsx)
        self.global_toolbar.addWidget(self.btn_export_xlsx)

        self.dashboard_layout.addLayout(self.global_toolbar)

        # --- Área Principal com Abas ---
        self.tabWidget = QTabWidget(self.DashboardPage)
        self.tabWidget.setTabPosition(QTabWidget.TabPosition.West) # Abas na esquerda
        self.tabWidget.setDocumentMode(False)
        
        # 1. Aba Sítios
        self.NovoBanco = QWidget() # Mantendo o nome da variável para compatibilidade, mas agora é a aba Sítios
        self.layout_novo = QVBoxLayout(self.NovoBanco)
        
        # Cabeçalho da Aba Sítios
        header_site = QHBoxLayout()
        self.label_site = QLabel()
        self.label_site.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        header_site.addWidget(self.label_site)
        header_site.addStretch()
        self.layout_novo.addLayout(header_site)
        
        # Layout Botoes Ação
        self.action_layout_site = QHBoxLayout()
        self.action_layout_site.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Botão Adicionar Sítio
        self.button_add_site = QPushButton("Adicionar Sítio", self.NovoBanco)
        self.button_add_site.setFixedSize(200, 25)
        self.button_add_site.clicked.connect(self.open_add_site_dialog)
        self.action_layout_site.addWidget(self.button_add_site)

        # Botão Deletar Sítio Selecionado
        self.button_delete_site = QPushButton("Deletar Sítio Selecionado", self.NovoBanco)
        self.button_delete_site.setFixedSize(200, 25)
        self.button_delete_site.clicked.connect(self.delete_selected_site)
        self.action_layout_site.addWidget(self.button_delete_site)
        
        self.layout_novo.addLayout(self.action_layout_site)
        
        # Layout Botoes
        self.buttons_layout_site = QGridLayout()
        
        # Campo para digitar o filtro
        self.filter_input_site = QLineEdit(self.NovoBanco)
        self.filter_input_site.setFixedSize(850, 25)  # Largura 300px e altura 30px
        self.filter_input_site.setPlaceholderText("Digite o filtro para buscar...")
        self.buttons_layout_site.addWidget(self.filter_input_site, 0, 0, 1, 3, Qt.AlignmentFlag.AlignLeft)

        # Botão para aplicar o filtro
        self.button_apply_filter_site = QPushButton("Aplicar Filtro", self.NovoBanco)
        self.button_apply_filter_site.setFixedSize(200, 25)  # Largura 200px e altura 50px
        self.button_apply_filter_site.clicked.connect(self.apply_filter)
        self.buttons_layout_site.addWidget(self.button_apply_filter_site, 0, 4, Qt.AlignmentFlag.AlignCenter)

        # Botão para limpar o filtro
        self.button_clear_filter_site = QPushButton("Limpar Filtro", self.NovoBanco)
        self.button_clear_filter_site.setFixedSize(200, 25)  # Largura 200px e altura 50px
        self.button_clear_filter_site.clicked.connect(self.clear_filter)
        self.buttons_layout_site.addWidget(self.button_clear_filter_site, 0, 5, Qt.AlignmentFlag.AlignLeft)
        
        self.layout_novo.addLayout(self.buttons_layout_site)

        # Fim Layout Botoes
        
        # Tabela de sítios
        self.table_sites = QTableWidget(self.NovoBanco)
        self.table_sites.setColumnCount(8)
        self.table_sites.setHorizontalHeaderLabels(
            ["ID", "Nome", "Estado", "Cidade", "Localização", "Número", "Longitude", "Latitude"]
        )
        self.style_table(self.table_sites) # Aplica estilo moderno
        # Registrar mapeamento para commits de edição
        self.register_table(self.table_sites, 'Site', [
            'id', 'name', 'state', 'city', 'location', 'number', 'longitude', 'latitude'
        ])
        self.layout_novo.addWidget(self.table_sites)

        self.tabWidget.addTab(self.NovoBanco, "Sítios")
        self.tabWidget.addTab(self.NovoBanco, "") # Texto definido em update_ui_text
   
        #########################################################################################
        # 2. Aba Coleções
        self.Colecoes = QWidget()
        self.layout_colecoes = QVBoxLayout(self.Colecoes)

        # Cabeçalho
        header_col = QHBoxLayout()
        self.label_col = QLabel()
        self.label_col.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        header_col.addWidget(self.label_col)
        header_col.addStretch()
        self.layout_colecoes.addLayout(header_col)

        # Layout Botoes Ação
        self.action_layout_collection = QHBoxLayout()
        self.action_layout_collection.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Botão Adicionar Coleção
        self.button_add_collection = QPushButton("Adicionar Coleção", self.Colecoes)
        self.button_add_collection.setFixedSize(200, 25) # Largura 200px e altura 50px
        self.button_add_collection.clicked.connect(self.show_add_collection_dialog)
        self.action_layout_collection.addWidget(self.button_add_collection)
        
        # Botão Deletar Coleção Selecionada
        self.button_delete_collection = QPushButton("Deletar Selecionado", self.Colecoes)
        self.button_delete_collection.setFixedSize(200, 25) # Largura 200px e altura 50px
        self.button_delete_collection.clicked.connect(self.delete_selected_collection)
        self.action_layout_collection.addWidget(self.button_delete_collection)
        
        self.layout_colecoes.addLayout(self.action_layout_collection)
        
        # Layout de Filtros Hierárquicos para Coleções (Site -> Coleção)
        self.filter_hierarchy_layout_col = QHBoxLayout()
        self.combo_filter_site_col = QComboBox(self.Colecoes)
        self.combo_filter_site_col.setPlaceholderText("Filtrar por Sítio")
        self.combo_filter_site_col.addItem("Todos os Sítios", None)
        self.combo_filter_site_col.currentIndexChanged.connect(self.update_collection_filter_collections)
        self.combo_filter_site_col.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_col.addWidget(self.combo_filter_site_col)

        self.combo_filter_collection_col = QComboBox(self.Colecoes)
        self.combo_filter_collection_col.setPlaceholderText("Filtrar por Coleção")
        self.combo_filter_collection_col.addItem("Todas as Coleções", None)
        self.combo_filter_collection_col.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_col.addWidget(self.combo_filter_collection_col)

        self.layout_colecoes.addLayout(self.filter_hierarchy_layout_col)

        # Layout Botoes
        self.buttons_layout_collection = QGridLayout()
        
        # Campo para digitar o filtro
        self.filter_input_collection = QLineEdit(self.Colecoes)
        self.filter_input_collection.setFixedSize(850, 25)  # Largura 300px e altura 30px
        self.filter_input_collection.setPlaceholderText("Digite o filtro para buscar...")
        self.buttons_layout_collection.addWidget(self.filter_input_collection, 0, 0, 1, 3, Qt.AlignmentFlag.AlignLeft)

        # Botão para aplicar o filtro
        self.button_apply_filter_collection = QPushButton("Aplicar Filtro", self.Colecoes)
        self.button_apply_filter_collection.setFixedSize(200, 25)  # Largura 200px e altura 50px
        self.button_apply_filter_collection.clicked.connect(self.apply_filter)
        self.buttons_layout_collection.addWidget(self.button_apply_filter_collection, 0, 4, Qt.AlignmentFlag.AlignCenter)

        # Botão para limpar o filtro
        self.button_clear_filter_collection = QPushButton("Limpar Filtro", self.Colecoes)
        self.button_clear_filter_collection.setFixedSize(200, 25)  # Largura 200px e altura 50px
        self.button_clear_filter_collection.clicked.connect(self.clear_filter)
        self.buttons_layout_collection.addWidget(self.button_clear_filter_collection, 0, 5, Qt.AlignmentFlag.AlignLeft)
        
        self.layout_colecoes.addLayout(self.buttons_layout_collection)

        # Fim Layout Botoes

        # Tabela de Coleções
        self.table_collections = QTableWidget(self.Colecoes)
        self.table_collections.setColumnCount(5)
        self.table_collections.setHorizontalHeaderLabels(["ID", "Sítio", "Nome", "Longitude", "Latitude"])
        self.style_table(self.table_collections)
        self.layout_colecoes.addWidget(self.table_collections)

        self.tabWidget.addTab(self.Colecoes, "Coleções")
        self.tabWidget.addTab(self.Colecoes, "")
        
        #########################################################################################
        # 3. Aba Amostras
        self.Amostras = QWidget()
        self.layout_amostras = QVBoxLayout(self.Amostras)

        # Cabeçalho
        header_samp = QHBoxLayout()
        self.label_samp = QLabel()
        self.label_samp.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        header_samp.addWidget(self.label_samp)
        header_samp.addStretch()
        self.layout_amostras.addLayout(header_samp)

        # Layout Botoes Ação
        self.action_layout_sample = QHBoxLayout()
        self.action_layout_sample.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Botão Adicionar Amostra
        self.button_add_sample = QPushButton("Adicionar Amostra", self.Amostras)
        self.button_add_sample.setFixedSize(200, 25) # Largura 200px e altura 50px
        self.button_add_sample.clicked.connect(self.show_add_sample_dialog)
        self.action_layout_sample.addWidget(self.button_add_sample)

        # Botão Deletar Amostra Selecionada
        self.button_delete_sample = QPushButton("Deletar Selecionado", self.Amostras)
        self.button_delete_sample.setFixedSize(200, 25) # Largura 200px e altura 50px
        self.button_delete_sample.clicked.connect(self.delete_selected_sample)
        self.action_layout_sample.addWidget(self.button_delete_sample)
        
        self.layout_amostras.addLayout(self.action_layout_sample)
        
        # Layout de Filtros Hierárquicos para Amostras (Site -> Coleção -> Amostra)
        self.filter_hierarchy_layout_samp = QHBoxLayout()
        self.combo_filter_site_samp = QComboBox(self.Amostras)
        self.combo_filter_site_samp.setPlaceholderText("Filtrar por Sítio")
        self.combo_filter_site_samp.addItem("Todos os Sítios", None)
        self.combo_filter_site_samp.currentIndexChanged.connect(self.update_collection_filter_samples)
        self.combo_filter_site_samp.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_samp.addWidget(self.combo_filter_site_samp)

        self.combo_filter_collection_samp = QComboBox(self.Amostras)
        self.combo_filter_collection_samp.setPlaceholderText("Filtrar por Coleção")
        self.combo_filter_collection_samp.addItem("Todas as Coleções", None)
        self.combo_filter_collection_samp.currentIndexChanged.connect(self.update_sample_filter_samples)
        self.combo_filter_collection_samp.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_samp.addWidget(self.combo_filter_collection_samp)

        self.combo_filter_sample_samp = QComboBox(self.Amostras)
        self.combo_filter_sample_samp.setPlaceholderText("Filtrar por Amostra")
        self.combo_filter_sample_samp.addItem("Todas as Amostras", None)
        self.combo_filter_sample_samp.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_samp.addWidget(self.combo_filter_sample_samp)

        self.layout_amostras.addLayout(self.filter_hierarchy_layout_samp)

        # Layout Botoes
        self.buttons_layout_sample = QGridLayout()
        
        # Campo para digitar o filtro
        self.filter_input_sample = QLineEdit(self.Amostras)
        self.filter_input_sample.setFixedSize(850, 25)  # Largura 300px e altura 30px
        self.filter_input_sample.setPlaceholderText("Digite o filtro para buscar...")
        self.buttons_layout_sample.addWidget(self.filter_input_sample, 0, 0, 1, 3, Qt.AlignmentFlag.AlignLeft)

        # Botão para aplicar o filtro
        self.button_apply_filter_sample = QPushButton("Aplicar Filtro", self.Amostras)
        self.button_apply_filter_sample.setFixedSize(200, 25)  # Largura 200px e altura 50px
        self.button_apply_filter_sample.clicked.connect(self.apply_filter)
        self.buttons_layout_sample.addWidget(self.button_apply_filter_sample, 0, 4, Qt.AlignmentFlag.AlignCenter)

        # Botão para limpar o filtro
        self.button_clear_filter_sample = QPushButton("Limpar Filtro", self.Amostras)
        self.button_clear_filter_sample.setFixedSize(200, 25)  # Largura 200px e altura 50px
        self.button_clear_filter_sample.clicked.connect(self.clear_filter)
        self.buttons_layout_sample.addWidget(self.button_clear_filter_sample, 0, 5, Qt.AlignmentFlag.AlignLeft)
        
        self.layout_amostras.addLayout(self.buttons_layout_sample)

        # Fim Layout Botoes

        # Tabela de Amostras
        self.table_samples = QTableWidget(self.Amostras)
        self.table_samples.setColumnCount(6)
        self.table_samples.setHorizontalHeaderLabels(["ID", "Coleção", "Nome", "Longitude", "Latitude", "Malha da Peneira"])
        self.style_table(self.table_samples)
        self.register_table(self.table_samples, 'Assemblage', [
            'id', 'collection_id', 'name', 'longitude', 'latitude', 'screenSize'
        ])
        self.layout_amostras.addWidget(self.table_samples)

        self.tabWidget.addTab(self.Amostras, "Amostras")
        self.tabWidget.addTab(self.Amostras, "")

        #########################################################################################
        # 3.1 Aba Unidade de Escavação
        self.Unidades = QWidget()
        self.layout_unidades = QVBoxLayout(self.Unidades)

        header_unit = QHBoxLayout()
        self.label_unit = QLabel()
        self.label_unit.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        header_unit.addWidget(self.label_unit)
        header_unit.addStretch()
        self.layout_unidades.addLayout(header_unit)

        # Ações
        self.action_layout_unit = QHBoxLayout()
        self.action_layout_unit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.button_add_unit = QPushButton("Adicionar Unidade", self.Unidades)
        self.button_add_unit.setFixedSize(200,25)
        self.button_add_unit.clicked.connect(self.show_add_unit_dialog)
        self.action_layout_unit.addWidget(self.button_add_unit)

        self.button_delete_unit = QPushButton("Deletar Selecionado", self.Unidades)
        self.button_delete_unit.setFixedSize(200,25)
        self.button_delete_unit.clicked.connect(self.delete_selected_unit)
        self.action_layout_unit.addWidget(self.button_delete_unit)

        self.layout_unidades.addLayout(self.action_layout_unit)

        # Layout de Filtros Hierárquicos para Unidades (Site -> Coleção -> Amostra)
        self.filter_hierarchy_layout_unit = QHBoxLayout()
        self.combo_filter_site_unit = QComboBox(self.Unidades)
        self.combo_filter_site_unit.setPlaceholderText("Filtrar por Sítio")
        self.combo_filter_site_unit.addItem("Todos os Sítios", None)
        self.combo_filter_site_unit.currentIndexChanged.connect(self.update_collection_filter_units)
        self.combo_filter_site_unit.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_unit.addWidget(self.combo_filter_site_unit)

        self.combo_filter_collection_unit = QComboBox(self.Unidades)
        self.combo_filter_collection_unit.setPlaceholderText("Filtrar por Coleção")
        self.combo_filter_collection_unit.addItem("Todas as Coleções", None)
        self.combo_filter_collection_unit.currentIndexChanged.connect(self.update_sample_filter_units)
        self.combo_filter_collection_unit.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_unit.addWidget(self.combo_filter_collection_unit)

        self.combo_filter_sample_unit = QComboBox(self.Unidades)
        self.combo_filter_sample_unit.setPlaceholderText("Filtrar por Amostra")
        self.combo_filter_sample_unit.addItem("Todas as Amostras", None)
        self.combo_filter_sample_unit.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_unit.addWidget(self.combo_filter_sample_unit)

        self.layout_unidades.addLayout(self.filter_hierarchy_layout_unit)

        # Botões de filtro
        self.buttons_layout_unit = QGridLayout()
        self.filter_input_unit = QLineEdit(self.Unidades)
        self.filter_input_unit.setFixedSize(850,25)
        self.filter_input_unit.setPlaceholderText("Digite o filtro para buscar...")
        self.buttons_layout_unit.addWidget(self.filter_input_unit, 0, 0, 1, 3, Qt.AlignmentFlag.AlignLeft)

        self.button_apply_filter_unit = QPushButton("Aplicar Filtro", self.Unidades)
        self.button_apply_filter_unit.setFixedSize(200,25)
        self.button_apply_filter_unit.clicked.connect(self.apply_filter)
        self.buttons_layout_unit.addWidget(self.button_apply_filter_unit, 0, 4, Qt.AlignmentFlag.AlignCenter)

        self.button_clear_filter_unit = QPushButton("Limpar Filtro", self.Unidades)
        self.button_clear_filter_unit.setFixedSize(200,25)
        self.button_clear_filter_unit.clicked.connect(self.clear_filter)
        self.buttons_layout_unit.addWidget(self.button_clear_filter_unit, 0, 5, Qt.AlignmentFlag.AlignLeft)

        self.layout_unidades.addLayout(self.buttons_layout_unit)

        # Tabela Unidades
        self.table_units = QTableWidget(self.Unidades)
        headers_unit = [
            "ID", "Amostra", "Nome", "Tipo", "Tamanho", "Latitude", "Longitude",
            "Nível Holótipo", "Profundidade Inicial", "Profundidade Final", "Camada Geológica",
            "Método Escavação", "Método Peneiramento", "Responsável Escavação", "Data Início",
            "Data Conclusão", "Fotos", "Observação"
        ]
        self.table_units.setColumnCount(len(headers_unit))
        self.table_units.setHorizontalHeaderLabels(headers_unit)
        self.table_units.setColumnCount(18)
        self.style_table(self.table_units)
        self.register_table(self.table_units, 'ExcavationUnit', [
            'id', 'assemblage_id', 'name', 'unit_type', 'size', 'latitude', 'longitude',
            'nivel_holotipo', 'profundidade_inicial', 'profundidade_final', 'camada_geologica',
            'metodo_escavacao', 'metodo_peneiramento', 'responsavel_escavacao', 'data_inicio',
            'data_conclusao', 'fotos_registro', 'observacao'
        ])
        self.layout_unidades.addWidget(self.table_units)

        self.tabWidget.addTab(self.Unidades, "Unidade de Escavação")
        self.tabWidget.addTab(self.Unidades, "")

        #########################################################################################
        # 3.2 Aba Níveis
        self.Niveis = QWidget()
        self.layout_niveis = QVBoxLayout(self.Niveis)

        header_lvl = QHBoxLayout()
        self.label_lvl = QLabel()
        self.label_lvl.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        header_lvl.addWidget(self.label_lvl)
        header_lvl.addStretch()
        self.layout_niveis.addLayout(header_lvl)

        # Ações
        self.action_layout_lvl = QHBoxLayout()
        self.action_layout_lvl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.button_add_lvl = QPushButton("Adicionar Nível", self.Niveis)
        self.button_add_lvl.setFixedSize(200, 25)
        self.button_add_lvl.clicked.connect(self.show_add_level_dialog)
        self.action_layout_lvl.addWidget(self.button_add_lvl)

        self.button_delete_lvl = QPushButton("Deletar Selecionado", self.Niveis)
        self.button_delete_lvl.setFixedSize(200, 25)
        self.button_delete_lvl.clicked.connect(self.delete_selected_level)
        self.action_layout_lvl.addWidget(self.button_delete_lvl)
        self.layout_niveis.addLayout(self.action_layout_lvl)

        # Layout de Filtros Hierárquicos para Níveis (Site -> Coleção -> Amostra -> Unidade)
        self.filter_hierarchy_layout_lvl = QHBoxLayout()
        
        self.combo_filter_site_lvl = QComboBox(self.Niveis)
        self.combo_filter_site_lvl.setPlaceholderText("Filtrar por Sítio")
        self.combo_filter_site_lvl.addItem("Todos os Sítios", None)
        self.combo_filter_site_lvl.currentIndexChanged.connect(self.update_collection_filter_levels)
        self.combo_filter_site_lvl.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_lvl.addWidget(self.combo_filter_site_lvl)

        self.combo_filter_collection_lvl = QComboBox(self.Niveis)
        self.combo_filter_collection_lvl.setPlaceholderText("Filtrar por Coleção")
        self.combo_filter_collection_lvl.addItem("Todas as Coleções", None)
        self.combo_filter_collection_lvl.currentIndexChanged.connect(self.update_sample_filter_levels)
        self.combo_filter_collection_lvl.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_lvl.addWidget(self.combo_filter_collection_lvl)

        self.combo_filter_sample_lvl = QComboBox(self.Niveis)
        self.combo_filter_sample_lvl.setPlaceholderText("Filtrar por Amostra")
        self.combo_filter_sample_lvl.addItem("Todas as Amostras", None)
        self.combo_filter_sample_lvl.currentIndexChanged.connect(self.update_unit_filter_levels)
        self.combo_filter_sample_lvl.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_lvl.addWidget(self.combo_filter_sample_lvl)

        self.combo_filter_unit_lvl = QComboBox(self.Niveis)
        self.combo_filter_unit_lvl.setPlaceholderText("Filtrar por Unidade")
        self.combo_filter_unit_lvl.addItem("Todas as Unidades", None)
        self.combo_filter_unit_lvl.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout_lvl.addWidget(self.combo_filter_unit_lvl)

        self.layout_niveis.addLayout(self.filter_hierarchy_layout_lvl)

        # Filtros
        self.buttons_layout_lvl = QGridLayout()
        self.filter_input_lvl = QLineEdit(self.Niveis)
        self.filter_input_lvl.setFixedSize(850, 25)
        self.filter_input_lvl.setPlaceholderText("Digite o filtro para buscar...")
        self.buttons_layout_lvl.addWidget(self.filter_input_lvl, 0, 0, 1, 3, Qt.AlignmentFlag.AlignLeft)

        self.button_apply_filter_lvl = QPushButton("Aplicar Filtro", self.Niveis)
        self.button_apply_filter_lvl.setFixedSize(200, 25)
        self.button_apply_filter_lvl.clicked.connect(self.apply_filter)
        self.buttons_layout_lvl.addWidget(self.button_apply_filter_lvl, 0, 4, Qt.AlignmentFlag.AlignCenter)

        self.button_clear_filter_lvl = QPushButton("Limpar Filtro", self.Niveis)
        self.button_clear_filter_lvl.setFixedSize(200, 25)
        self.button_clear_filter_lvl.clicked.connect(self.clear_filter)
        self.buttons_layout_lvl.addWidget(self.button_clear_filter_lvl, 0, 5, Qt.AlignmentFlag.AlignLeft)
        self.layout_niveis.addLayout(self.buttons_layout_lvl)

        # Tabela
        self.table_levels = QTableWidget(self.Niveis)
        self.table_levels.setColumnCount(8)
        self.table_levels.setHorizontalHeaderLabels(["ID", "Unidade", "Nível", "Prof. Inicial", "Prof. Final", "Cor", "Textura", "Descrição"])
        self.style_table(self.table_levels)
        self.register_table(self.table_levels, 'Level', [
            'id', 'excavation_unit_id', 'level', 'start_depth', 'end_depth', 'color', 'texture', 'description'
        ])
        self.layout_niveis.addWidget(self.table_levels)
        self.tabWidget.addTab(self.Niveis, "Níveis")
        self.tabWidget.addTab(self.Niveis, "")

        #########################################################################################
        # 4. Aba Materiais
        self.Materiais = QWidget()
        self.layout_materiais = QVBoxLayout(self.Materiais)

        # Cabeçalho
        header_mat = QHBoxLayout()
        self.label_mat = QLabel()
        self.label_mat.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        header_mat.addWidget(self.label_mat)
        header_mat.addStretch()
        self.layout_materiais.addLayout(header_mat)

        # Layout Botoes Ação
        self.action_layout_material = QHBoxLayout()
        self.action_layout_material.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Botão Adicionar Material
        self.button_add_material = QPushButton("Adicionar Material", self.Materiais)
        self.button_add_material.setFixedSize(200, 25) # Largura 200px e altura 50px
        self.button_add_material.clicked.connect(self.show_add_material_dialog)
        self.action_layout_material.addWidget(self.button_add_material)

        # Botão Deletar Material Selecionado
        self.button_delete_material = QPushButton("Deletar Selecionado", self.Materiais)
        self.button_delete_material.setFixedSize(200, 25) # Largura 200px e altura 50px
        self.button_delete_material.clicked.connect(self.delete_selected_material)
        self.action_layout_material.addWidget(self.button_delete_material)
        
        # Botão Importar/Atualizar Excel
        self.button_update_material_xlsx = QPushButton("Importar/Atualizar Excel", self.Materiais)
        self.button_update_material_xlsx.setFixedSize(200, 25)
        self.button_update_material_xlsx.clicked.connect(self.update_database_from_excel)
        self.action_layout_material.addWidget(self.button_update_material_xlsx)

        self.layout_materiais.addLayout(self.action_layout_material)

        # Layout de Filtros Hierárquicos (Dropdowns)
        self.filter_hierarchy_layout = QHBoxLayout()
        
        self.combo_filter_site = QComboBox(self.Materiais)
        self.combo_filter_site.setPlaceholderText("Filtrar por Sítio")
        self.combo_filter_site.addItem("Todos os Sítios", None)
        self.combo_filter_site.currentIndexChanged.connect(self.update_collection_filter)
        self.combo_filter_site.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout.addWidget(self.combo_filter_site)

        self.combo_filter_collection = QComboBox(self.Materiais)
        self.combo_filter_collection.setPlaceholderText("Filtrar por Coleção")
        self.combo_filter_collection.addItem("Todas as Coleções", None)
        self.combo_filter_collection.currentIndexChanged.connect(self.update_sample_filter)
        self.combo_filter_collection.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout.addWidget(self.combo_filter_collection)

        self.combo_filter_sample = QComboBox(self.Materiais)
        self.combo_filter_sample.setPlaceholderText("Filtrar por Amostra")
        self.combo_filter_sample.addItem("Todas as Amostras", None)
        self.combo_filter_sample.currentIndexChanged.connect(self.update_unit_filter)
        self.combo_filter_sample.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout.addWidget(self.combo_filter_sample)

        self.combo_filter_unit = QComboBox(self.Materiais)
        self.combo_filter_unit.setPlaceholderText("Filtrar por Unidade")
        self.combo_filter_unit.addItem("Todas as Unidades", None)
        self.combo_filter_unit.currentIndexChanged.connect(self.apply_filter)
        self.filter_hierarchy_layout.addWidget(self.combo_filter_unit)

        self.layout_materiais.addLayout(self.filter_hierarchy_layout)
        
        # Layout Botoes
        self.buttons_layout_material = QGridLayout()
        
        # Campo para digitar o filtro
        self.filter_input_material = QLineEdit(self.Materiais)
        self.filter_input_material.setFixedSize(850, 25)  # Largura 300px e altura 30px
        self.filter_input_material.setPlaceholderText("Digite o filtro para buscar...")
        self.buttons_layout_material.addWidget(self.filter_input_material, 0, 0, 1, 3, Qt.AlignmentFlag.AlignLeft)

        # Botão para aplicar o filtro
        self.button_apply_filter_material = QPushButton("Aplicar Filtro", self.Materiais)
        self.button_apply_filter_material.setFixedSize(200, 25)  # Largura 200px e altura 50px
        self.button_apply_filter_material.clicked.connect(self.apply_filter)
        self.buttons_layout_material.addWidget(self.button_apply_filter_material, 0, 4, Qt.AlignmentFlag.AlignCenter)

        # Botão para limpar o filtro
        self.button_clear_filter_material = QPushButton("Limpar Filtro", self.Materiais)
        self.button_clear_filter_material.setFixedSize(200, 25)  # Largura 200px e altura 50px
        self.button_clear_filter_material.clicked.connect(self.clear_filter)
        self.buttons_layout_material.addWidget(self.button_clear_filter_material, 0, 5, Qt.AlignmentFlag.AlignLeft)
        
        self.layout_materiais.addLayout(self.buttons_layout_material)

        # Fim Layout Botoes
        
        # Tabela de Materiais
        self.table_material = QTableWidget(self.Materiais)
        self.table_material.setColumnCount(19)
        self.table_material.setHorizontalHeaderLabels([
            "ID", "Sítio", "Unidade", "Nível", "UUID", "Serial de Campo", "Serial de Laboratório",
            "Tipo de Material", "Descrição", "Medidas", "Peso (g)", "Quantidade",
            "X", "Y", "Z", "Notas", "Fotos", "Usuário", "Data de Catalogação"
        ])

        self.style_table(self.table_material)
        # 'None' para a coluna Sítio, pois ela não existe na tabela Material do DB
        self.register_table(self.table_material, 'Material', [
            'id', None, 'excavation_unit_id', 'level_id', 'uuid', 'field_serial', 'lab_serial',
            'material_type', 'material_description', 'measurements', 'weight', 'quantity',
            'x_coord', 'y_coord', 'z_coord', 'notes', 'photos', 'user', 'cataloging_date'
        ])
        self.layout_materiais.addWidget(self.table_material)

        self.tabWidget.addTab(self.Materiais, "Materiais")
        self.tabWidget.addTab(self.Materiais, "")
        
        #########################################################################################
        # 5. Aba Estatísticas
        self.Estatistica = QWidget()
        self.layout_stats = QVBoxLayout(self.Estatistica)
        
        # Cabeçalho
        header_stats = QHBoxLayout()
        self.label_stats = QLabel("Estatísticas e Gráficos")
        self.label_stats.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        header_stats.addWidget(self.label_stats)
        header_stats.addStretch()
        self.layout_stats.addLayout(header_stats)

        # Sub-abas para Materiais e Unidades
        self.tab_stats_sub = QTabWidget()
        self.layout_stats.addWidget(self.tab_stats_sub)

        # --- Sub-aba Estatísticas de Materiais ---
        self.tab_stats_material = QWidget()
        self.layout_stats_mat = QVBoxLayout(self.tab_stats_material)
        
        # Filtros (Reutilizando lógica hierárquica)
        self.stats_filter_layout_mat = QHBoxLayout()
        self.combo_stats_site_mat = QComboBox()
        self.combo_stats_col_mat = QComboBox()
        self.combo_stats_samp_mat = QComboBox()
        self.combo_stats_unit_mat = QComboBox()
        
        self.stats_filter_layout_mat.addWidget(self.combo_stats_site_mat)
        self.stats_filter_layout_mat.addWidget(self.combo_stats_col_mat)
        self.stats_filter_layout_mat.addWidget(self.combo_stats_samp_mat)
        self.stats_filter_layout_mat.addWidget(self.combo_stats_unit_mat)
        self.layout_stats_mat.addLayout(self.stats_filter_layout_mat)

        # Seleção de Tipo de Gráfico e Ações
        self.stats_action_layout_mat = QHBoxLayout()
        self.combo_chart_type_mat = QComboBox()
        self.combo_chart_type_mat.addItems(["Tipo de Material", "Descrição de Material", "Quantidade"])
        
        # Controles de Tamanho e Largura da Barra
        self.spin_chart_w_mat = QDoubleSpinBox()
        self.spin_chart_w_mat.setRange(1, 50)
        self.spin_chart_w_mat.setValue(10)
        self.spin_chart_w_mat.setPrefix("W: ")
        self.spin_chart_h_mat = QDoubleSpinBox()
        self.spin_chart_h_mat.setRange(1, 50)
        self.spin_chart_h_mat.setValue(6)
        self.spin_chart_h_mat.setPrefix("H: ")
        self.spin_bar_w_mat = QDoubleSpinBox()
        self.spin_bar_w_mat.setRange(0.1, 1.0)
        self.spin_bar_w_mat.setSingleStep(0.1)
        self.spin_bar_w_mat.setValue(0.8)
        self.spin_bar_w_mat.setPrefix("Bar: ")

        self.btn_generate_chart_mat = QPushButton("Gerar Gráfico")
        self.btn_export_chart_mat = QPushButton("Baixar PNG")
        
        self.stats_action_layout_mat.addWidget(QLabel("Tipo de Gráfico:"))
        self.stats_action_layout_mat.addWidget(self.combo_chart_type_mat)
        self.stats_action_layout_mat.addWidget(self.spin_chart_w_mat)
        self.stats_action_layout_mat.addWidget(self.spin_chart_h_mat)
        self.stats_action_layout_mat.addWidget(self.spin_bar_w_mat)
        self.stats_action_layout_mat.addWidget(self.btn_generate_chart_mat)
        self.stats_action_layout_mat.addWidget(self.btn_export_chart_mat)
        self.stats_action_layout_mat.addStretch()
        self.layout_stats_mat.addLayout(self.stats_action_layout_mat)

        # Área do Gráfico
        self.chart_canvas_mat = statistics.MplCanvas(self.tab_stats_material, width=5, height=4, dpi=100)
        self.layout_stats_mat.addWidget(self.chart_canvas_mat)
        
        self.tab_stats_sub.addTab(self.tab_stats_material, "Materiais")

        # --- Sub-aba Estatísticas de Unidades ---
        self.tab_stats_unit = QWidget()
        self.layout_stats_unit = QVBoxLayout(self.tab_stats_unit)

        # Filtros Unidade (Site -> Col -> Samp)
        self.stats_filter_layout_unit = QHBoxLayout()
        self.combo_stats_site_unit = QComboBox()
        self.combo_stats_col_unit = QComboBox()
        self.combo_stats_samp_unit = QComboBox()
        
        self.stats_filter_layout_unit.addWidget(self.combo_stats_site_unit)
        self.stats_filter_layout_unit.addWidget(self.combo_stats_col_unit)
        self.stats_filter_layout_unit.addWidget(self.combo_stats_samp_unit)
        self.layout_stats_unit.addLayout(self.stats_filter_layout_unit)

        # Ações Unidade
        self.stats_action_layout_unit = QHBoxLayout()
        
        self.spin_chart_w_unit = QDoubleSpinBox()
        self.spin_chart_w_unit.setRange(1, 50)
        self.spin_chart_w_unit.setValue(10)
        self.spin_chart_w_unit.setPrefix("W: ")
        self.spin_chart_h_unit = QDoubleSpinBox()
        self.spin_chart_h_unit.setRange(1, 50)
        self.spin_chart_h_unit.setValue(6)
        self.spin_chart_h_unit.setPrefix("H: ")
        self.spin_bar_w_unit = QDoubleSpinBox()
        self.spin_bar_w_unit.setRange(0.1, 1.0)
        self.spin_bar_w_unit.setSingleStep(0.1)
        self.spin_bar_w_unit.setValue(0.8)
        self.spin_bar_w_unit.setPrefix("Bar: ")

        self.btn_generate_chart_unit = QPushButton("Gerar Gráfico (Densidade)")
        self.btn_export_chart_unit = QPushButton("Baixar PNG")
        self.stats_action_layout_unit.addWidget(self.spin_chart_w_unit)
        self.stats_action_layout_unit.addWidget(self.spin_chart_h_unit)
        self.stats_action_layout_unit.addWidget(self.spin_bar_w_unit)
        self.stats_action_layout_unit.addWidget(self.btn_generate_chart_unit)
        self.stats_action_layout_unit.addWidget(self.btn_export_chart_unit)
        self.stats_action_layout_unit.addStretch()
        self.layout_stats_unit.addLayout(self.stats_action_layout_unit)
        
        self.chart_canvas_unit = statistics.MplCanvas(self.tab_stats_unit, width=5, height=4, dpi=100)
        self.layout_stats_unit.addWidget(self.chart_canvas_unit)

        self.tab_stats_sub.addTab(self.tab_stats_unit, "Unidades")

        self.tabWidget.addTab(self.Estatistica, "Estatísticas")
        self.tabWidget.addTab(self.Estatistica, "")

        # Conexões de Sinais para Estatísticas
        self.setup_stats_connections()

        # Adiciona o TabWidget ao layout do Dashboard
        self.dashboard_layout.addWidget(self.tabWidget)
        
        # Adiciona a página Dashboard ao StackedWidget
        self.stackedWidget.addWidget(self.DashboardPage)

        # Inicializa textos
        self.update_ui_text()
        
        #########################################################################################
        #########################################################################################

    def change_language(self, index):
        """Muda o idioma da interface."""
        lang_code = self.combo_lang.itemData(index)
        if lang_code:
            self.current_lang = lang_code
            self.update_ui_text()

    def change_theme(self, index):
        """Muda o tema da interface."""
        theme_code = self.combo_theme.itemData(index)
        if theme_code:
            self.current_theme = theme_code
            self.apply_theme(theme_code)

    def apply_theme(self, theme_name):
        """Aplica o CSS correspondente ao tema."""
        if theme_name in self.themes:
            self.MainWindow.setStyleSheet(self.themes[theme_name])

    def update_ui_text(self):
        """Atualiza todos os textos da interface com base no idioma selecionado."""
        
        # Dicionário de Traduções
        translations = {
            "pt": {
                "new_db": "Novo Banco de Dados", "open_db": "Abrir Banco de Dados", "import_xlsx": "Criar DB de Excel",
                "settings": "Configurações", "about": "Sobre", "exit": "Sair", "back": "Voltar",
                "close_db": "Fechar Banco", "export_pdf_mat": "Gerar PDF (Materiais)", "export_pdf_unit": "Gerar PDF (Unidades)",
                "export_xlsx": "Exportar Excel Completo",
                "tab_sites": "Sítios", "tab_collections": "Coleções", "tab_samples": "Amostras",
                "tab_units": "Unidade de Escavação", "tab_levels": "Níveis", "tab_specimens": "Materiais",
                "tab_stats": "Estatísticas",
                "header_sites": "Gerenciar Sítios", "header_collections": "Gerenciar Coleções", "header_samples": "Gerenciar Amostras",
                "header_units": "Gerenciar Unidades", "header_levels": "Gerenciar Níveis", "header_specimens": "Gerenciar Materiais",
                "btn_add_site": "Adicionar Sítio", "btn_del_site": "Deletar Sítio Selecionado",
                "btn_add_col": "Adicionar Coleção", "btn_del_col": "Deletar Selecionado",
                "btn_add_samp": "Adicionar Amostra", "btn_del_samp": "Deletar Selecionado",
                "btn_add_unit": "Adicionar Unidade", "btn_del_unit": "Deletar Selecionado",
                "btn_add_lvl": "Adicionar Nível", "btn_del_lvl": "Deletar Selecionado",
                "btn_add_spec": "Adicionar Material", "btn_del_spec": "Deletar Selecionado", "btn_upd_xlsx": "Importar/Atualizar Excel",
                "filter_ph": "Digite o filtro para buscar...", "btn_apply": "Aplicar Filtro", "btn_clear": "Limpar Filtro",
                "ph_site": "Filtrar por Sítio", "ph_col": "Filtrar por Coleção", "ph_samp": "Filtrar por Amostra", "ph_unit": "Filtrar por Unidade",
                "all_sites": "Todos os Sítios", "all_cols": "Todas as Coleções", "all_samps": "Todas as Amostras", "all_units": "Todas as Unidades",
                "lang_label": "Idioma / Language:", "theme_label": "Tema / Theme:", "config_title": "Configurações"
            },
            "en": {
                "new_db": "New Database", "open_db": "Open Database", "import_xlsx": "Create DB from Excel",
                "settings": "Settings", "about": "About", "exit": "Exit", "back": "Back",
                "close_db": "Close Database", "export_pdf_mat": "Generate PDF (Materials)", "export_pdf_unit": "Generate PDF (Units)",
                "export_xlsx": "Export Full Excel",
                "tab_sites": "Sites", "tab_collections": "Collections", "tab_samples": "Assemblages",
                "tab_units": "Excavation Units", "tab_levels": "Levels", "tab_specimens": "Materials",
                "tab_stats": "Statistics",
                "header_sites": "Manage Sites", "header_collections": "Manage Collections", "header_samples": "Manage Assemblages",
                "header_units": "Manage Units", "header_levels": "Manage Levels", "header_specimens": "Manage Materials",
                "btn_add_site": "Add Site", "btn_del_site": "Delete Selected Site",
                "btn_add_col": "Add Collection", "btn_del_col": "Delete Selected",
                "btn_add_samp": "Add Assemblage", "btn_del_samp": "Delete Selected",
                "btn_add_unit": "Add Unit", "btn_del_unit": "Delete Selected",
                "btn_add_lvl": "Add Level", "btn_del_lvl": "Delete Selected",
                "btn_add_spec": "Add Material", "btn_del_spec": "Delete Selected", "btn_upd_xlsx": "Import/Update Excel",
                "filter_ph": "Type filter to search...", "btn_apply": "Apply Filter", "btn_clear": "Clear Filter",
                "ph_site": "Filter by Site", "ph_col": "Filter by Collection", "ph_samp": "Filter by Assemblage", "ph_unit": "Filter by Unit",
                "all_sites": "All Sites", "all_cols": "All Collections", "all_samps": "All Assemblages", "all_units": "All Units",
                "lang_label": "Language:", "theme_label": "Theme:", "config_title": "Settings"
            },
            "es": {
                "new_db": "Nueva Base de Datos", "open_db": "Abrir Base de Datos", "import_xlsx": "Crear BD desde Excel",
                "settings": "Configuración", "about": "Acerca de", "exit": "Salir", "back": "Volver",
                "close_db": "Cerrar Base de Datos", "export_pdf_mat": "Generar PDF (Materiales)", "export_pdf_unit": "Generar PDF (Unidades)",
                "export_xlsx": "Exportar Excel Completo",
                "tab_sites": "Sitios", "tab_collections": "Colecciones", "tab_samples": "Muestras",
                "tab_units": "Unidades de Excavación", "tab_levels": "Niveles", "tab_specimens": "Materiales",
                "tab_stats": "Estadísticas",
                "header_sites": "Gestionar Sitios", "header_collections": "Gestionar Colecciones", "header_samples": "Gestionar Muestras",
                "header_units": "Gestionar Unidades", "header_levels": "Gestionar Niveles", "header_specimens": "Gestionar Materiales",
                "btn_add_site": "Añadir Sitio", "btn_del_site": "Eliminar Sitio Seleccionado",
                "btn_add_col": "Añadir Colección", "btn_del_col": "Eliminar Seleccionado",
                "btn_add_samp": "Añadir Muestra", "btn_del_samp": "Eliminar Seleccionado",
                "btn_add_unit": "Añadir Unidad", "btn_del_unit": "Eliminar Seleccionado",
                "btn_add_lvl": "Añadir Nivel", "btn_del_lvl": "Eliminar Seleccionado",
                "btn_add_spec": "Añadir Material", "btn_del_spec": "Eliminar Seleccionado", "btn_upd_xlsx": "Importar/Actualizar Excel",
                "filter_ph": "Escriba filtro para buscar...", "btn_apply": "Aplicar Filtro", "btn_clear": "Limpiar Filtro",
                "ph_site": "Filtrar por Sitio", "ph_col": "Filtrar por Colección", "ph_samp": "Filtrar por Muestra", "ph_unit": "Filtrar por Unidad",
                "all_sites": "Todos los Sitios", "all_cols": "Todas las Colecciones", "all_samps": "Todas las Muestras", "all_units": "Todas las Unidades",
                "lang_label": "Idioma / Language:", "theme_label": "Tema / Theme:", "config_title": "Configuración"
            }
        }

        t = translations.get(self.current_lang, translations["pt"])

        # Menu Principal
        self.button_newDB.setText(t["new_db"])
        self.button_openDB.setText(t["open_db"])
        self.button_import_xlsx.setText(t["import_xlsx"])
        self.button_settings.setText(t["settings"])
        self.button_about.setText(t["about"])
        self.button_exit.setText(t["exit"])
        self.sobre_Voltar.setText(t["back"])
        self.btn_config_back.setText(t["back"])
        self.label_config_title.setText(t["config_title"])
        self.label_lang.setText(t["lang_label"])
        self.label_theme.setText(t["theme_label"])

        # Dashboard Toolbar
        self.btn_close_db.setText(t["close_db"])
        self.btn_export_pdf.setText(t["export_pdf_mat"])
        self.btn_export_unit_pdf.setText(t["export_pdf_unit"])
        self.btn_export_xlsx.setText(t["export_xlsx"])

        # Abas
        self.tabWidget.setTabText(0, t["tab_sites"])
        self.tabWidget.setTabText(1, t["tab_collections"])
        self.tabWidget.setTabText(2, t["tab_samples"])
        self.tabWidget.setTabText(3, t["tab_units"])
        self.tabWidget.setTabText(4, t["tab_levels"])
        self.tabWidget.setTabText(5, t["tab_specimens"])
        self.tabWidget.setTabText(6, t["tab_stats"])

        # Headers
        self.label_site.setText(t["header_sites"])
        self.label_col.setText(t["header_collections"])
        self.label_samp.setText(t["header_samples"])
        self.label_unit.setText(t["header_units"])
        self.label_lvl.setText(t["header_levels"])
        self.label_mat.setText(t["header_specimens"])

        # Botões de Ação
        self.button_add_site.setText(t["btn_add_site"])
        self.button_delete_site.setText(t["btn_del_site"])
        self.button_add_collection.setText(t["btn_add_col"])
        self.button_delete_collection.setText(t["btn_del_col"])
        self.button_add_sample.setText(t["btn_add_samp"])
        self.button_delete_sample.setText(t["btn_del_samp"])
        self.button_add_unit.setText(t["btn_add_unit"])
        self.button_delete_unit.setText(t["btn_del_unit"])
        self.button_add_lvl.setText(t["btn_add_lvl"])
        self.button_delete_lvl.setText(t["btn_del_lvl"])
        self.button_add_material.setText(t["btn_add_spec"])
        self.button_delete_material.setText(t["btn_del_spec"])
        self.button_update_material_xlsx.setText(t["btn_upd_xlsx"])

        # Filtros
        for btn in [self.button_apply_filter_site, self.button_apply_filter_collection, self.button_apply_filter_sample,
                self.button_apply_filter_unit, self.button_apply_filter_lvl, self.button_apply_filter_material]:
            btn.setText(t["btn_apply"])
        
        for btn in [self.button_clear_filter_site, self.button_clear_filter_collection, self.button_clear_filter_sample,
                self.button_clear_filter_unit, self.button_clear_filter_lvl, self.button_clear_filter_material]:
            btn.setText(t["btn_clear"])

        for inp in [self.filter_input_site, self.filter_input_collection, self.filter_input_sample,
                self.filter_input_unit, self.filter_input_lvl, self.filter_input_material]:
            inp.setPlaceholderText(t["filter_ph"])

        # Dropdowns Placeholders (Atualiza texto do item 0 se for o placeholder)
        # Nota: QComboBox não tem placeholder nativo editável facilmente sem customização, 
        # mas usamos setPlaceholderText.
        for cb in [self.combo_filter_site_col, self.combo_filter_site_samp, self.combo_filter_site_unit, 
                   self.combo_filter_site_lvl, self.combo_filter_site]:
            cb.setPlaceholderText(t["ph_site"])
            if cb.count() > 0: cb.setItemText(0, t["all_sites"])
            
        for cb in [self.combo_filter_collection_col, self.combo_filter_collection_samp, self.combo_filter_collection_unit,
                   self.combo_filter_collection_lvl, self.combo_filter_collection]:
            cb.setPlaceholderText(t["ph_col"])
            if cb.count() > 0: cb.setItemText(0, t["all_cols"])

        for cb in [self.combo_filter_sample_samp, self.combo_filter_sample_unit, self.combo_filter_sample_lvl, self.combo_filter_sample]:
            cb.setPlaceholderText(t["ph_samp"])
            if cb.count() > 0: cb.setItemText(0, t["all_samps"])

        for cb in [self.combo_filter_unit_lvl, self.combo_filter_unit]:
            cb.setPlaceholderText(t["ph_unit"])
            if cb.count() > 0: cb.setItemText(0, t["all_units"])

        # Atualizar Cabeçalhos das Tabelas
        if self.current_lang == "pt":
            self.table_sites.setHorizontalHeaderLabels(["ID", "Nome", "Estado", "Cidade", "Localização", "Número", "Longitude", "Latitude"])
            self.table_collections.setHorizontalHeaderLabels(["ID", "Sítio", "Nome", "Longitude", "Latitude"])
            self.table_samples.setHorizontalHeaderLabels(["ID", "Coleção", "Nome", "Longitude", "Latitude", "Malha da Peneira"])
            self.table_units.setHorizontalHeaderLabels([
                "ID", "Amostra", "Nome", "Tipo", "Tamanho", "Latitude", "Longitude",
                "Nível Holótipo", "Profundidade Inicial", "Profundidade Final", "Camada Geológica",
                "Método Escavação", "Método Peneiramento", "Responsável Escavação", "Data Início",
                "Data Conclusão", "Fotos", "Observação"
            ])
            self.table_levels.setHorizontalHeaderLabels(["ID", "Unidade", "Nível", "Prof. Inicial", "Prof. Final", "Cor", "Textura", "Descrição"])
            self.table_material.setHorizontalHeaderLabels([
                "ID", "Sítio", "Unidade", "Nível", "UUID", "Serial de Campo", "Serial de Laboratório",
                "Tipo de Material", "Descrição", "Medidas", "Peso (g)", "Quantidade",
                "X", "Y", "Z", "Notas", "Fotos", "Usuário", "Data de Catalogação"
            ])
        elif self.current_lang == "en":
            self.table_sites.setHorizontalHeaderLabels(["ID", "Name", "State", "City", "Location", "Number", "Longitude", "Latitude"])
            self.table_collections.setHorizontalHeaderLabels(["ID", "Site", "Name", "Longitude", "Latitude"])
            self.table_samples.setHorizontalHeaderLabels(["ID", "Collection", "Name", "Longitude", "Latitude", "Screen Size"])
            self.table_units.setHorizontalHeaderLabels([
                "ID", "Assemblage", "Name", "Type", "Size", "Latitude", "Longitude",
                "Holotype Level", "Start Depth", "End Depth", "Geological Layer",
                "Excavation Method", "Screening Method", "Excavator", "Start Date",
                "End Date", "Photos", "Notes"
            ])
            self.table_levels.setHorizontalHeaderLabels(["ID", "Unit", "Level", "Start Depth", "End Depth", "Color", "Texture", "Description"])
            self.table_material.setHorizontalHeaderLabels([
                "ID", "Site", "Unit", "Level", "UUID", "Field Serial", "Lab Serial",
                "Material Type", "Description", "Measurements", "Weight (g)", "Quantity",
                "X", "Y", "Z", "Notes", "Photos", "User", "Catalog Date"
            ])
        elif self.current_lang == "es":
            self.table_sites.setHorizontalHeaderLabels(["ID", "Nombre", "Estado", "Ciudad", "Ubicación", "Número", "Longitud", "Latitud"])
            self.table_collections.setHorizontalHeaderLabels(["ID", "Sitio", "Nombre", "Longitud", "Latitud"])
            self.table_samples.setHorizontalHeaderLabels(["ID", "Colección", "Nombre", "Longitud", "Latitud", "Tamaño Malla"])
            self.table_units.setHorizontalHeaderLabels([
                "ID", "Muestra", "Nombre", "Tipo", "Tamaño", "Latitud", "Longitud",
                "Nivel Holotipo", "Profundidad Inicial", "Profundidad Final", "Capa Geológica",
                "Método Excavación", "Método Cribado", "Responsable", "Fecha Inicio",
                "Fecha Fin", "Fotos", "Observación"
            ])
            self.table_levels.setHorizontalHeaderLabels(["ID", "Unidad", "Nivel", "Prof. Inicial", "Prof. Final", "Color", "Textura", "Descripción"])
            self.table_material.setHorizontalHeaderLabels([
                "ID", "Sitio", "Unidad", "Nivel", "UUID", "Serial Campo", "Serial Lab",
                "Tipo de Material", "Descripción", "Medidas", "Peso (g)", "Cantidad",
                "X", "Y", "Z", "Notas", "Fotos", "Usuario", "Fecha Catálogo"
            ])
    
    def style_table(self, table_widget):
        """Aplica estilo moderno e científico à tabela."""
        table_widget.setAlternatingRowColors(True)
        table_widget.setShowGrid(True)
        table_widget.verticalHeader().setVisible(False) # Esconde números de linha laterais
        header = table_widget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)

    def register_table(self, table_widget, table_name, column_keys):
        """Registra mapeamento entre um QTableWidget e a tabela do DB.

        column_keys deve ser uma lista de nomes de colunas do banco na mesma ordem das colunas
        exibidas na tabela (incluindo `id` como primeira coluna).
        """
        if not hasattr(self, 'table_mappings'):
            self.table_mappings = {}
        self.table_mappings[table_widget] = (table_name, column_keys)
        table_widget.itemChanged.connect(lambda item, tbl=table_widget: self.on_table_item_changed(tbl, item))

    def on_table_item_changed(self, table_widget, item):
        """Quando uma célula é editada, efetua o commit no banco de dados.

        Ignora alterações na coluna ID e quando a tabela está sendo populada.
        """
        # Evita reagir durante populações programáticas
        if getattr(self, '_populating_tables', False):
            return

        # Se o banco não estiver aberto/associado, ignora (ex.: import temporário)
        if not hasattr(self, 'db') or self.db is None:
            return

        row = item.row()
        col = item.column()
        # Não atualizar coluna ID
        if col == 0:
            return

        id_item = table_widget.item(row, 0)
        if id_item is None:
            return
        record_id = id_item.text()
        if record_id is None or record_id == "":
            return

        mapping = self.table_mappings.get(table_widget)
        if not mapping:
            return
        table_name, col_keys = mapping
        if col >= len(col_keys):
            return
        column_key = col_keys[col]
        if column_key is None: # Coluna apenas visual (ex: Sítio na tabela de espécimes)
            return
            
        new_value = item.text()

        # Tenta converter inteiros/float quando aplicável
        try:
            if new_value.isdigit():
                cast_value = int(new_value)
            else:
                cast_value = float(new_value)
        except Exception:
            cast_value = new_value

        try:
            # Usa a API de Database
            self.db.update(table_name, {column_key: cast_value}, {'id': int(record_id)})
        except Exception as e:
            # Log completo para diagnóstico no console
            try:
                traceback.print_exc()
            except Exception:
                pass
            # Mensagem mais específica para ajudar diagnóstico
            QMessageBox.warning(None, "Erro ao salvar", f"Falha ao salvar alteração ({type(e).__name__}): {str(e)}")

    def create_new_database(self):
        """Abre uma janela para inserir o nome do novo banco de dados, cria o banco e o abre."""
        dialog = dialogs.NewDatabaseDialog()
        if dialog.exec():  # Se o usuário confirmou
            db_name = dialog.get_db_name()
            if not db_name:
                QMessageBox.warning(None, "Erro", "O nome do banco de dados não pode estar vazio.")
                return
            
            base_dir = os.path.dirname(os.path.abspath(__file__))
            db_path = os.path.join(base_dir, "database", f"{db_name}.db")
            os.makedirs(os.path.dirname(db_path), exist_ok=True)

            # Verifica se o banco já existe
            if os.path.exists(db_path):
                msgbox = QMessageBox()
                msgbox.setIcon(QMessageBox.Icon.Warning)
                msgbox.setWindowTitle("Banco já existe")
                msgbox.setText(f"O banco de dados '{db_name}.db' já existe.\nDeseja sobrescrevê-lo?")
                msgbox.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                msgbox.setWindowIcon(QIcon("./icon/iconTray.png"))  # <- Define o ícone da janela
                reply = msgbox.exec()

                if reply == QMessageBox.StandardButton.No:
                    return  # Se o usuário escolher "Não", interrompe a criação

                # Se escolher "Sim", remove o banco antigo
                os.remove(db_path)

            # Criar o novo banco de dados
            self.db = Database(db_path, create_tables=True)
            self.load_sites()
            self.load_collections()
            self.load_samples()
            self.load_units()
            self.load_levels()
            self.load_material()
            self.populate_filter_dropdowns()
            self.populate_stats_dropdowns()
            self.stackedWidget.setCurrentWidget(self.DashboardPage)  # Vai para o Dashboard

    def close_database(self):
        """Fecha o banco de dados e volta para a tela inicial."""
        self.db = None
        self.stackedWidget.setCurrentWidget(self.Main)
    
    def open_add_site_dialog(self):
        """Abre uma janela para inserir os dados de um novo sítio."""
        dialog = dialogs.AddSiteDialog(self.db)
        if dialog.exec():
            self.load_sites()
        
    
    def delete_selected_site(self):
        """Deleta o sítio selecionado na tabela."""
        selected_row = self.table_sites.currentRow()
        if selected_row == -1:
            QMessageBox.warning(None, "Erro", "Selecione um sítio para deletar.")
            return

        site_name = self.table_sites.item(selected_row, 1).text()
        site_id = self.table_sites.item(selected_row, 0).text()
        confirm = QMessageBox.question(
            None, "Confirmar Exclusão", f"Tem certeza que deseja excluir o sítio {site_name}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if confirm == QMessageBox.StandardButton.Yes:
            self.db.delete("Site", {"id": site_id})
            self.load_sites()  # Atualiza a tabela
        
    def load_sites(self):
        """Carrega os sítios cadastrados na tabela"""
        # Evita disparar commits enquanto popula a tabela
        self._populating_tables = True
        self.table_sites.setRowCount(0)
        sites = self.db.fetch("Site")
        for row_number, row_data in enumerate(sites):
            self.table_sites.insertRow(row_number)
            for col_number, col_data in enumerate(row_data):
                self.table_sites.setItem(row_number, col_number, QTableWidgetItem(str(col_data)))
        self._populating_tables = False

    def open_database(self):
        """Abre uma janela para selecionar um banco de dados existente e carrega os dados."""
        file_path, _ = QFileDialog.getOpenFileName(None, "Abrir Banco de Dados", "", "SQLite Database (*.db)")
        if file_path:
            # Evita que handlers de edição façam commits enquanto carregamos o DB
            self._populating_tables = True
            try:
                self.db = Database(file_path)  # Abre o banco de dados selecionado
                # Carrega dados nas tabelas (cada load também gerencia _populating_tables)
                self.load_sites()
                self.load_collections()
                self.load_samples()
                self.load_units()
                self.load_levels()
                self.load_material()
                self.populate_filter_dropdowns()
                self.populate_stats_dropdowns()
                self.stackedWidget.setCurrentWidget(self.DashboardPage)  # Vai para o Dashboard
            finally:
                # Garante reset mesmo em caso de erro durante o load
                self._populating_tables = False
            
    def load_collections(self):
        """Carrega os sítios cadastrados na tabela"""
        self._populating_tables = True
        self.table_collections.setRowCount(0)
        self.collection_to_site = {} # Reset map
        
        # Map sites id -> name
        sites = self.db.fetch("Site", "id, name")
        site_map = {s[0]: s[1] for s in sites}

        collection = self.db.fetch("Collection")
        for row_number, row_data in enumerate(collection):
            self.table_collections.insertRow(row_number)
            for col_number, col_data in enumerate(row_data):
                if col_number == 1: # Site ID column
                    item = QTableWidgetItem(str(site_map.get(col_data, col_data)))
                    item.setData(Qt.UserRole, col_data) # Store ID
                    self.table_collections.setItem(row_number, col_number, item)
                    continue
                self.table_collections.setItem(row_number, col_number, QTableWidgetItem(str(col_data)))
            self.collection_to_site[str(row_data[0])] = str(row_data[1]) # id -> site_id
        self._populating_tables = False
            
    def show_add_collection_dialog(self):
        """Abre uma janela para adicionar uma nova coleção."""
        dialog = dialogs.AddCollectionDialog(self.db)
        if dialog.exec():
            self.load_collections()
        
    def delete_selected_collection(self):
        """Deleta a coleção selecionada na tabela."""
        selected_row = self.table_collections.currentRow()
        if selected_row == -1:
            QMessageBox.warning(None, "Erro", "Nenhuma coleção selecionada.")
            return

        collection_id = self.table_collections.item(selected_row, 0).text()
        self.db.delete("Collection", {"id": collection_id})
        self.load_collections()

    def show_add_sample_dialog(self):
        """Abre uma janela para adicionar uma nova amostra."""
        dialog = dialogs.AddSampleDialog(self.db)
        if dialog.exec():
            self.load_samples()

    def delete_selected_sample(self):
        """Deleta a amostra selecionada na tabela."""
        selected_row = self.table_samples.currentRow()
        if selected_row == -1:
            QMessageBox.warning(None, "Erro", "Selecione uma amostra para deletar.")
            return

        sample_id = self.table_samples.item(selected_row, 0).text()
        
        confirmation = QMessageBox.question(
            None, "Confirmação", f"Tem certeza que deseja excluir a amostra ID {sample_id}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if confirmation == QMessageBox.StandardButton.Yes:
            self.db.delete("Assemblage", {"id": sample_id})
            self.load_samples()


    def load_samples(self):
        """Carrega as amostras cadastradas na tabela."""
        self._populating_tables = True
        self.table_samples.setRowCount(0)
        self.assemblage_to_collection = {} # Reset map
        
        # Map collections id -> name
        cols = self.db.fetch("Collection", "id, name")
        col_map = {c[0]: c[1] for c in cols}

        samples = self.db.fetch("Assemblage")

        for row_number, row_data in enumerate(samples):
            self.table_samples.insertRow(row_number)
            for col_number, col_data in enumerate(row_data):
                if col_number == 1: # Collection ID column
                    item = QTableWidgetItem(str(col_map.get(col_data, col_data)))
                    item.setData(Qt.UserRole, col_data)
                    self.table_samples.setItem(row_number, col_number, item)
                    continue
                self.table_samples.setItem(row_number, col_number, QTableWidgetItem(str(col_data)))
            self.assemblage_to_collection[str(row_data[0])] = str(row_data[1]) # id -> collection_id
        self._populating_tables = False
                
    def show_add_material_dialog(self):
        """Abre uma janela para adicionar um novo material com barra de rolagem e layout em duas colunas."""
        dialog = dialogs.AddMaterialDialog(self.db)
        if dialog.exec():
            self.load_material()
        
    def show_add_unit_dialog(self):
        dialog = dialogs.AddUnitDialog(self.db)
        if dialog.exec():
            self.load_units()

    def show_add_level_dialog(self):
        dialog = dialogs.AddLevelDialog(self.db)
        if dialog.exec():
            self.load_levels()

    def load_levels(self):
        self._populating_tables = True
        self.table_levels.setRowCount(0)

        # Map units id -> name
        units = self.db.fetch("ExcavationUnit", "id, name")
        unit_map = {u[0]: u[1] for u in units}

        try:
            levels = self.db.fetch("Level")
            for row_number, row_data in enumerate(levels):
                self.table_levels.insertRow(row_number)
                for col_number, col_data in enumerate(row_data):
                    if col_number == 1: # Unit ID column
                        item = QTableWidgetItem(str(unit_map.get(col_data, col_data)))
                        item.setData(Qt.UserRole, col_data)
                        self.table_levels.setItem(row_number, col_number, item)
                        continue
                    self.table_levels.setItem(row_number, col_number, QTableWidgetItem(str(col_data)))
        finally:
            self._populating_tables = False

    def load_units(self):
        """Carrega as unidades de escavação na tabela."""
        self._populating_tables = True
        self.table_units.setRowCount(0)

        try:
            # Map samples id -> name
            samps = self.db.fetch("Assemblage", "id, name")
            samp_map = {s[0]: s[1] for s in samps}

            units = self.db.fetch("ExcavationUnit")
            self.unit_to_assemblage = {}
            for row_number, row_data in enumerate(units):
                self.table_units.insertRow(row_number)
                for col_number, col_data in enumerate(row_data):
                    if col_number == 1: # Assemblage ID column
                        item = QTableWidgetItem(str(samp_map.get(col_data, col_data)))
                        item.setData(Qt.UserRole, col_data)
                        self.table_units.setItem(row_number, col_number, item)
                        continue
                    self.table_units.setItem(row_number, col_number, QTableWidgetItem(str(col_data)))
                # row_data[0]=id, row_data[1]=assemblage_id
                try:
                    self.unit_to_assemblage[str(row_data[0])] = str(row_data[1])
                except Exception:
                    pass
        finally:
            self._populating_tables = False

    def delete_selected_unit(self):
        selected_row = self.table_units.currentRow()
        if selected_row == -1:
            QMessageBox.warning(None, "Erro", "Selecione uma unidade para deletar.")
            return
        unit_id = self.table_units.item(selected_row, 0).text()
        confirm = QMessageBox.question(None, "Confirmação", f"Deseja excluir a unidade ID {unit_id}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            self.db.delete("ExcavationUnit", {"id": unit_id})
            self.load_units()

    def delete_selected_level(self):
        selected_row = self.table_levels.currentRow()
        if selected_row == -1:
            QMessageBox.warning(None, "Erro", "Selecione um nível para deletar.")
            return
        lvl_id = self.table_levels.item(selected_row, 0).text()
        if QMessageBox.question(None, "Confirmação", f"Excluir nível ID {lvl_id}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self.db.delete("Level", {"id": lvl_id})
            self.load_levels()

    def load_material(self):
        """Carrega os materiais cadastrados na tabela."""
        self._populating_tables = True
        self.table_material.setRowCount(0)

        columns = (
            "id, excavation_unit_id, level_id, uuid, field_serial, lab_serial, material_type, "
            "material_description, measurements, weight, quantity, x_coord, y_coord, z_coord, "
            "notes, photos, user, cataloging_date"
        )
        samples = self.db.fetch("Material", columns=columns)

        # Maps
        units = self.db.fetch("ExcavationUnit", "id, name")
        unit_map = {u[0]: u[1] for u in units}

        levels = self.db.fetch("Level", "id, level")
        level_map = {l[0]: l[1] for l in levels}

        # Mapeamento para Sítio
        # Unit -> Assemblage -> Collection -> Site
        sites = {s[0]: s[1] for s in self.db.fetch("Site", "id, name")}
        cols = {c[0]: c[1] for c in self.db.fetch("Collection", "id, site_id")} # id -> site_id
        asms = {a[0]: a[1] for a in self.db.fetch("Assemblage", "id, collection_id")} # id -> col_id
        units_data = self.db.fetch("ExcavationUnit", "id, assemblage_id")

        unit_site_map = {}
        for u_id, u_asm_id in units_data:
            if u_asm_id in asms:
                col_id = asms[u_asm_id]
                if col_id in cols:
                    site_id = cols[col_id]
                    if site_id in sites:
                        unit_site_map[u_id] = sites[site_id]

        for row_number, row_data in enumerate(samples):
            self.table_material.insertRow(row_number)

            # Coluna 0: ID
            self.table_material.setItem(row_number, 0, QTableWidgetItem(str(row_data[0])))

            # Coluna 1: Sítio (Calculado)
            unit_id = row_data[1]
            site_name = unit_site_map.get(unit_id, "")
            self.table_material.setItem(row_number, 1, QTableWidgetItem(str(site_name)))

            # Restante das colunas (deslocadas por 1 devido à coluna Sítio)
            for col_number, col_data in enumerate(row_data):
                if col_number == 0: continue # Já inserido

                target_col = col_number + 1 # Desloca +1

                if col_number == 1: # Unit ID
                    item = QTableWidgetItem(str(unit_map.get(col_data, col_data)))
                    item.setData(Qt.UserRole, col_data)
                    self.table_material.setItem(row_number, target_col, item)
                    continue
                if col_number == 2: # Level ID
                    item = QTableWidgetItem(str(level_map.get(col_data, col_data) if col_data else ""))
                    item.setData(Qt.UserRole, col_data)
                    self.table_material.setItem(row_number, target_col, item)
                    continue
                self.table_material.setItem(row_number, target_col, QTableWidgetItem(str(col_data)))
        self._populating_tables = False

    def delete_selected_material(self):
        """Deleta o material selecionado na tabela."""
        selected_row = self.table_material.currentRow()
        if selected_row == -1:
            QMessageBox.warning(None, "Erro", "Selecione um material para deletar.")
            return

        material_id = self.table_material.item(selected_row, 0).text()
        
        confirmation = QMessageBox.question(
            None, "Confirmação", f"Tem certeza que deseja excluir o material ID: {material_id}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if confirmation == QMessageBox.StandardButton.Yes:
            self.db.delete("Material", {"id": material_id})
            self.load_material()

    def get_active_filter_widgets(self):
        """Retorna a tabela e o input de filtro ativos com base na página atual."""
        # Verifica se estamos na página do Dashboard
        if self.stackedWidget.currentWidget() == self.DashboardPage:
            current_tab = self.tabWidget.currentWidget()
            
            if current_tab == self.NovoBanco:
                return self.table_sites, self.filter_input_site
            elif current_tab == self.Colecoes:
                return self.table_collections, self.filter_input_collection
            elif current_tab == self.Amostras:
                return self.table_samples, self.filter_input_sample
            elif current_tab == self.Unidades:
                return self.table_units, self.filter_input_unit
            elif current_tab == self.Niveis:
                return self.table_levels, self.filter_input_lvl
            elif current_tab == self.Materiais:
                return self.table_material, self.filter_input_material
        return None, None
        
    def populate_filter_dropdowns(self):
        """Popula os dropdowns de filtro com dados do banco."""
        if not self.db:
            return
            
        # Sítios
        self.combo_filter_site.blockSignals(True)
        self.combo_filter_site.clear()
        self.combo_filter_site.addItem("Todos os Sítios", None)
        sites = self.db.fetch("Site", "id, name")
        for s_id, s_name in sites:
            self.combo_filter_site.addItem(str(s_name), s_id)
        self.combo_filter_site.blockSignals(False)
        
        # Atualiza os outros em cascata
        self.update_collection_filter()

        # Popula dropdowns para Coleções (coleções tab)
        try:
            self.combo_filter_site_col.blockSignals(True)
            self.combo_filter_site_col.clear()
            self.combo_filter_site_col.addItem("Todos os Sítios", None)
            for s_id, s_name in sites:
                self.combo_filter_site_col.addItem(str(s_name), s_id)
            self.combo_filter_site_col.blockSignals(False)
            self.update_collection_filter_collections()
        except Exception:
            pass

        # Popula dropdowns para Amostras (amostras tab)
        try:
            self.combo_filter_site_samp.blockSignals(True)
            self.combo_filter_site_samp.clear()
            self.combo_filter_site_samp.addItem("Todos os Sítios", None)
            for s_id, s_name in sites:
                self.combo_filter_site_samp.addItem(str(s_name), s_id)
            self.combo_filter_site_samp.blockSignals(False)
            self.update_collection_filter_samples()
        except Exception:
            pass

        # Popula dropdowns para Unidades (unidades tab)
        try:
            self.combo_filter_site_unit.blockSignals(True)
            self.combo_filter_site_unit.clear()
            self.combo_filter_site_unit.addItem("Todos os Sítios", None)
            for s_id, s_name in sites:
                self.combo_filter_site_unit.addItem(str(s_name), s_id)
            self.combo_filter_site_unit.blockSignals(False)
            self.update_collection_filter_units()
        except Exception:
            pass

        # Popula dropdowns para Níveis (níveis tab)
        try:
            self.combo_filter_site_lvl.blockSignals(True)
            self.combo_filter_site_lvl.clear()
            self.combo_filter_site_lvl.addItem("Todos os Sítios", None)
            for s_id, s_name in sites:
                self.combo_filter_site_lvl.addItem(str(s_name), s_id)
            self.combo_filter_site_lvl.blockSignals(False)
            self.update_collection_filter_levels()
        except Exception:
            pass
            
        # Popula dropdowns para Estatísticas
        self.populate_stats_dropdowns()

    def populate_stats_dropdowns(self):
        """Popula os dropdowns da aba de estatísticas."""
        if not self.db: return
        
        sites = self.db.fetch("Site", "id, name")
        
        # Espécimes
        self.combo_stats_site_mat.blockSignals(True)
        self.combo_stats_site_mat.clear()
        self.combo_stats_site_mat.addItem("Todos os Sítios", None)
        for s_id, s_name in sites:
            self.combo_stats_site_mat.addItem(str(s_name), s_id)
        self.combo_stats_site_mat.blockSignals(False)
        self.update_stats_collection_mat()

        # Unidades
        self.combo_stats_site_unit.blockSignals(True)
        self.combo_stats_site_unit.clear()
        self.combo_stats_site_unit.addItem("Todos os Sítios", None)
        for s_id, s_name in sites:
            self.combo_stats_site_unit.addItem(str(s_name), s_id)
        self.combo_stats_site_unit.blockSignals(False)
        self.update_stats_collection_unit()

    def setup_stats_connections(self):
        # Filtros Espécimes
        self.combo_stats_site_mat.currentIndexChanged.connect(self.update_stats_collection_mat)
        self.combo_stats_col_mat.currentIndexChanged.connect(self.update_stats_sample_mat)
        self.combo_stats_samp_mat.currentIndexChanged.connect(self.update_stats_unit_mat)
        
        # Botões Espécimes
        self.btn_generate_chart_mat.clicked.connect(self.generate_material_chart)
        self.btn_export_chart_mat.clicked.connect(lambda: self.export_chart(self.chart_canvas_mat))

        # Filtros Unidades
        self.combo_stats_site_unit.currentIndexChanged.connect(self.update_stats_collection_unit)
        self.combo_stats_col_unit.currentIndexChanged.connect(self.update_stats_sample_unit)
        
        # Botões Unidades
        self.btn_generate_chart_unit.clicked.connect(self.generate_unit_chart)
        self.btn_export_chart_unit.clicked.connect(lambda: self.export_chart(self.chart_canvas_unit))

    def update_stats_collection_mat(self):
        self._update_stats_combo(self.combo_stats_site_mat, self.combo_stats_col_mat, "Collection", "site_id")
        self.update_stats_sample_mat()

    def update_stats_sample_mat(self):
        self._update_stats_combo(self.combo_stats_col_mat, self.combo_stats_samp_mat, "Assemblage", "collection_id")
        self.update_stats_unit_mat()

    def update_stats_unit_mat(self):
        self._update_stats_combo(self.combo_stats_samp_mat, self.combo_stats_unit_mat, "ExcavationUnit", "assemblage_id")

    def update_stats_collection_unit(self):
        self._update_stats_combo(self.combo_stats_site_unit, self.combo_stats_col_unit, "Collection", "site_id")
        self.update_stats_sample_unit()

    def update_stats_sample_unit(self):
        self._update_stats_combo(self.combo_stats_col_unit, self.combo_stats_samp_unit, "Assemblage", "collection_id")

    def _update_stats_combo(self, parent_combo, child_combo, table, fk_field):
        parent_id = parent_combo.currentData()
        child_combo.blockSignals(True)
        child_combo.clear()
        child_combo.addItem("Todos", None)
        
        if parent_id:
            items = self.db.fetch(table, "id, name", {fk_field: parent_id})
        else:
            # Se não tem pai selecionado, idealmente mostraria todos ou limparia. 
            # Para simplificar, mostra todos se o banco não for gigante, ou limpa.
            # Vamos limpar para forçar hierarquia ou mostrar todos se for o primeiro nível.
            # Aqui, assumimos mostrar todos se parent for None (comportamento padrão das outras abas)
            items = self.db.fetch(table, "id, name")
            
        for i_id, i_name in items:
            child_combo.addItem(str(i_name), i_id)
        child_combo.blockSignals(False)

    def generate_material_chart(self):
        if not self.db: return
        filters = {
            'site_id': self.combo_stats_site_mat.currentData(),
            'collection_id': self.combo_stats_col_mat.currentData(),
            'assemblage_id': self.combo_stats_samp_mat.currentData(),
            'excavation_unit_id': self.combo_stats_unit_mat.currentData()
        }
        settings = {
            'width': self.spin_chart_w_mat.value(),
            'height': self.spin_chart_h_mat.value(),
            'bar_width': self.spin_bar_w_mat.value()
        }
        chart_type = self.combo_chart_type_mat.currentText()
        
        if "Tipo de Material" in chart_type or "Type" in chart_type:
            statistics.plot_material_types(self.db, filters, self.chart_canvas_mat, settings)
        elif "Descrição" in chart_type or "Description" in chart_type:
            statistics.plot_material_descriptions(self.db, filters, self.chart_canvas_mat, settings)
        elif "Quantidade" in chart_type or "Quantity" in chart_type:
            statistics.plot_material_quantities(self.db, filters, self.chart_canvas_mat, settings)

    def generate_unit_chart(self):
        if not self.db: return
        filters = {
            'site_id': self.combo_stats_site_unit.currentData(),
            'collection_id': self.combo_stats_col_unit.currentData(),
            'assemblage_id': self.combo_stats_samp_unit.currentData()
        }
        settings = {
            'width': self.spin_chart_w_unit.value(),
            'height': self.spin_chart_h_unit.value(),
            'bar_width': self.spin_bar_w_unit.value()
        }
        # Por enquanto apenas um tipo
        statistics.plot_unit_counts(self.db, filters, self.chart_canvas_unit, settings)

    def export_chart(self, canvas):
        """Exporta o gráfico atual para PNG."""
        file_path, _ = QFileDialog.getSaveFileName(None, "Salvar Gráfico", "grafico.png", "PNG Files (*.png)")
        if file_path:
            try:
                canvas.fig.savefig(file_path)
                QMessageBox.information(None, "Sucesso", "Gráfico salvo com sucesso!")
            except Exception as e:
                QMessageBox.critical(None, "Erro", f"Erro ao salvar gráfico: {str(e)}")

    def update_collection_filter(self):
        """Atualiza o dropdown de coleções baseado no sítio selecionado."""
        site_id = self.combo_filter_site.currentData()
        
        self.combo_filter_collection.blockSignals(True)
        self.combo_filter_collection.clear()
        self.combo_filter_collection.addItem("Todas as Coleções", None)
        
        collections = self.db.fetch("Collection", "id, name, site_id")
        for col_id, name, s_id in collections:
            if site_id is None or str(s_id) == str(site_id):
                self.combo_filter_collection.addItem(str(name), col_id)
        
        self.combo_filter_collection.blockSignals(False)
        self.update_sample_filter()

    def update_collection_filter_collections(self):
        """Atualiza o dropdown de coleções na aba Coleções baseado no sítio selecionado nesta aba."""
        try:
            site_id = self.combo_filter_site_col.currentData()
            self.combo_filter_collection_col.blockSignals(True)
            self.combo_filter_collection_col.clear()
            self.combo_filter_collection_col.addItem("Todas as Coleções", None)
            collections = self.db.fetch("Collection", "id, name, site_id")
            for col_id, name, s_id in collections:
                if site_id is None or str(s_id) == str(site_id):
                    self.combo_filter_collection_col.addItem(str(name), col_id)
            self.combo_filter_collection_col.blockSignals(False)
        except Exception:
            pass

    def update_collection_filter_samples(self):
        """Atualiza o dropdown de coleções na aba Amostras baseado no sítio selecionado nesta aba."""
        try:
            site_id = self.combo_filter_site_samp.currentData()
            self.combo_filter_collection_samp.blockSignals(True)
            self.combo_filter_collection_samp.clear()
            self.combo_filter_collection_samp.addItem("Todas as Coleções", None)
            collections = self.db.fetch("Collection", "id, name, site_id")
            for col_id, name, s_id in collections:
                if site_id is None or str(s_id) == str(site_id):
                    self.combo_filter_collection_samp.addItem(str(name), col_id)
            self.combo_filter_collection_samp.blockSignals(False)
            self.update_sample_filter_samples()
        except Exception:
            pass

    def update_sample_filter_samples(self):
        """Atualiza o dropdown de amostras na aba Amostras baseado na coleção (e sítio) selecionada."""
        try:
            col_id = self.combo_filter_collection_samp.currentData()
            site_id = self.combo_filter_site_samp.currentData()
            self.combo_filter_sample_samp.blockSignals(True)
            self.combo_filter_sample_samp.clear()
            self.combo_filter_sample_samp.addItem("Todas as Amostras", None)
            samples = self.db.fetch("Assemblage", "id, name, collection_id")
            valid_col_ids = None
            if site_id is not None and col_id is None:
                cols = self.db.fetch("Collection", "id", {"site_id": site_id})
                valid_col_ids = [c[0] for c in cols]

            for samp_id, name, c_id in samples:
                if col_id is not None:
                    if str(c_id) == str(col_id):
                        self.combo_filter_sample_samp.addItem(str(name), samp_id)
                elif valid_col_ids is not None:
                    if c_id in valid_col_ids:
                        self.combo_filter_sample_samp.addItem(str(name), samp_id)
                else:
                    self.combo_filter_sample_samp.addItem(str(name), samp_id)

            self.combo_filter_sample_samp.blockSignals(False)
        except Exception:
            pass

    def update_sample_filter(self):
        """Atualiza o dropdown de amostras baseado na coleção (e sítio) selecionada."""
        col_id = self.combo_filter_collection.currentData()
        site_id = self.combo_filter_site.currentData()
        
        self.combo_filter_sample.blockSignals(True)
        self.combo_filter_sample.clear()
        self.combo_filter_sample.addItem("Todas as Amostras", None)
        
        samples = self.db.fetch("Assemblage", "id, name, collection_id")
        
        # Se não tem coleção selecionada, mas tem sítio, precisamos filtrar amostras cujas coleções são daquele sítio
        valid_col_ids = None
        if site_id is not None and col_id is None:
             cols = self.db.fetch("Collection", "id", {"site_id": site_id})
             valid_col_ids = [c[0] for c in cols]

        for samp_id, name, c_id in samples:
            if col_id is not None:
                if str(c_id) == str(col_id):
                    self.combo_filter_sample.addItem(str(name), samp_id)
            elif valid_col_ids is not None:
                if c_id in valid_col_ids:
                    self.combo_filter_sample.addItem(str(name), samp_id)
            else:
                self.combo_filter_sample.addItem(str(name), samp_id)
        
        self.combo_filter_sample.blockSignals(False)
        self.update_unit_filter()

    def update_unit_filter(self):
        """Atualiza o dropdown de unidades na aba Espécimes baseado na amostra selecionada."""
        try:
            sample_id = self.combo_filter_sample.currentData()
            col_id = self.combo_filter_collection.currentData()
            site_id = self.combo_filter_site.currentData()

            self.combo_filter_unit.blockSignals(True)
            self.combo_filter_unit.clear()
            self.combo_filter_unit.addItem("Todas as Unidades", None)
            
            units = self.db.fetch("ExcavationUnit", "id, name, assemblage_id")
            
            valid_sample_ids = None
            
            if sample_id is None:
                if col_id is not None:
                    samps = self.db.fetch("Assemblage", "id", {"collection_id": col_id})
                    valid_sample_ids = [str(s[0]) for s in samps]
                elif site_id is not None:
                    cols = self.db.fetch("Collection", "id", {"site_id": site_id})
                    col_ids = [c[0] for c in cols]
                    if col_ids:
                        all_samples = self.db.fetch("Assemblage", "id, collection_id")
                        valid_sample_ids = [str(s[0]) for s in all_samples if s[1] in col_ids]
                    else:
                        valid_sample_ids = []

            for u_id, u_name, u_samp_id in units:
                if sample_id is not None:
                    if str(u_samp_id) == str(sample_id):
                        self.combo_filter_unit.addItem(str(u_name), u_id)
                elif valid_sample_ids is not None:
                    if str(u_samp_id) in valid_sample_ids:
                        self.combo_filter_unit.addItem(str(u_name), u_id)
                else:
                    self.combo_filter_unit.addItem(str(u_name), u_id)

            self.combo_filter_unit.blockSignals(False)
        except Exception:
            pass

    def update_collection_filter_units(self):
        """Atualiza o dropdown de coleções na aba Unidades baseado no sítio selecionado nesta aba."""
        try:
            site_id = self.combo_filter_site_unit.currentData()
            self.combo_filter_collection_unit.blockSignals(True)
            self.combo_filter_collection_unit.clear()
            self.combo_filter_collection_unit.addItem("Todas as Coleções", None)
            collections = self.db.fetch("Collection", "id, name, site_id")
            for col_id, name, s_id in collections:
                if site_id is None or str(s_id) == str(site_id):
                    self.combo_filter_collection_unit.addItem(str(name), col_id)
            self.combo_filter_collection_unit.blockSignals(False)
            self.update_sample_filter_units()
        except Exception:
            pass

    def update_sample_filter_units(self):
        """Atualiza o dropdown de amostras na aba Unidades baseado na coleção (e sítio) selecionada."""
        try:
            col_id = self.combo_filter_collection_unit.currentData()
            site_id = self.combo_filter_site_unit.currentData()
            self.combo_filter_sample_unit.blockSignals(True)
            self.combo_filter_sample_unit.clear()
            self.combo_filter_sample_unit.addItem("Todas as Amostras", None)
            samples = self.db.fetch("Assemblage", "id, name, collection_id")
            
            valid_col_ids = None
            if site_id is not None and col_id is None:
                 cols = self.db.fetch("Collection", "id", {"site_id": site_id})
                 valid_col_ids = [c[0] for c in cols]

            for samp_id, name, c_id in samples:
                if col_id is not None:
                    if str(c_id) == str(col_id):
                        self.combo_filter_sample_unit.addItem(str(name), samp_id)
                elif valid_col_ids is not None:
                    if c_id in valid_col_ids:
                        self.combo_filter_sample_unit.addItem(str(name), samp_id)
                else:
                    self.combo_filter_sample_unit.addItem(str(name), samp_id)
            
            self.combo_filter_sample_unit.blockSignals(False)
        except Exception:
            pass

    def update_collection_filter_levels(self):
        """Atualiza o dropdown de coleções na aba Níveis baseado no sítio selecionado."""
        try:
            site_id = self.combo_filter_site_lvl.currentData()
            self.combo_filter_collection_lvl.blockSignals(True)
            self.combo_filter_collection_lvl.clear()
            self.combo_filter_collection_lvl.addItem("Todas as Coleções", None)
            collections = self.db.fetch("Collection", "id, name, site_id")
            for col_id, name, s_id in collections:
                if site_id is None or str(s_id) == str(site_id):
                    self.combo_filter_collection_lvl.addItem(str(name), col_id)
            self.combo_filter_collection_lvl.blockSignals(False)
            self.update_sample_filter_levels()
        except Exception:
            pass

    def update_sample_filter_levels(self):
        """Atualiza o dropdown de amostras na aba Níveis baseado na coleção selecionada."""
        try:
            col_id = self.combo_filter_collection_lvl.currentData()
            site_id = self.combo_filter_site_lvl.currentData()
            self.combo_filter_sample_lvl.blockSignals(True)
            self.combo_filter_sample_lvl.clear()
            self.combo_filter_sample_lvl.addItem("Todas as Amostras", None)
            samples = self.db.fetch("Assemblage", "id, name, collection_id")
            
            valid_col_ids = None
            if site_id is not None and col_id is None:
                 cols = self.db.fetch("Collection", "id", {"site_id": site_id})
                 valid_col_ids = [c[0] for c in cols]

            for samp_id, name, c_id in samples:
                if col_id is not None:
                    if str(c_id) == str(col_id):
                        self.combo_filter_sample_lvl.addItem(str(name), samp_id)
                elif valid_col_ids is not None:
                    if c_id in valid_col_ids:
                        self.combo_filter_sample_lvl.addItem(str(name), samp_id)
                else:
                    self.combo_filter_sample_lvl.addItem(str(name), samp_id)
            
            self.combo_filter_sample_lvl.blockSignals(False)
            self.update_unit_filter_levels()
        except Exception:
            pass

    def update_unit_filter_levels(self):
        """Atualiza o dropdown de unidades na aba Níveis baseado na amostra selecionada."""
        try:
            sample_id = self.combo_filter_sample_lvl.currentData()
            col_id = self.combo_filter_collection_lvl.currentData()
            site_id = self.combo_filter_site_lvl.currentData()

            self.combo_filter_unit_lvl.blockSignals(True)
            self.combo_filter_unit_lvl.clear()
            self.combo_filter_unit_lvl.addItem("Todas as Unidades", None)
            
            units = self.db.fetch("ExcavationUnit", "id, name, assemblage_id")
            
            valid_sample_ids = None
            
            if sample_id is None:
                if col_id is not None:
                    samps = self.db.fetch("Assemblage", "id", {"collection_id": col_id})
                    valid_sample_ids = [str(s[0]) for s in samps]
                elif site_id is not None:
                    cols = self.db.fetch("Collection", "id", {"site_id": site_id})
                    col_ids = [c[0] for c in cols]
                    if col_ids:
                        all_samples = self.db.fetch("Assemblage", "id, collection_id")
                        valid_sample_ids = [str(s[0]) for s in all_samples if s[1] in col_ids]
                    else:
                        valid_sample_ids = []

            for u_id, u_name, u_samp_id in units:
                if sample_id is not None:
                    if str(u_samp_id) == str(sample_id):
                        self.combo_filter_unit_lvl.addItem(str(u_name), u_id)
                elif valid_sample_ids is not None:
                    if str(u_samp_id) in valid_sample_ids:
                        self.combo_filter_unit_lvl.addItem(str(u_name), u_id)
                else:
                    self.combo_filter_unit_lvl.addItem(str(u_name), u_id)

            self.combo_filter_unit_lvl.blockSignals(False)
        except Exception:
            pass

    # Função para aplicar o filtro
    def apply_filter(self):
        table, filter_input = self.get_active_filter_widgets()
        
        if not table or not filter_input:
            return

        filter_text = filter_input.text().strip().lower()
        
        # Lógica de aspas para correspondência exata
        exact_match = False
        if filter_text.startswith('"') and filter_text.endswith('"') and len(filter_text) >= 2:
            filter_text = filter_text[1:-1] # Remove as aspas
            exact_match = True

        # Filtros Hierárquicos (Apenas para tabela de Materiais)
        selected_site_id = None
        selected_col_id = None
        selected_sample_id = None
        selected_unit_id = None
        
        if table == self.table_material:
            if self.combo_filter_site.currentData():
                selected_site_id = str(self.combo_filter_site.currentData())
            if self.combo_filter_collection.currentData():
                selected_col_id = str(self.combo_filter_collection.currentData())
            if self.combo_filter_sample.currentData():
                selected_sample_id = str(self.combo_filter_sample.currentData())
            if hasattr(self, 'combo_filter_unit') and self.combo_filter_unit.currentData():
                selected_unit_id = str(self.combo_filter_unit.currentData())
        
        for row in range(table.rowCount()):
            match = False
            
            # Verifica filtros hierárquicos primeiro (se for tabela de espécimes)
            if table == self.table_material:
                # Coluna 2 é ID da Unidade de Escavação (ExcavationUnit ID) - Ajustado para nova estrutura
                item_unit_id_widget = table.item(row, 2)
                item_unit_id = str(item_unit_id_widget.data(Qt.UserRole)) if item_unit_id_widget else ""

                if selected_unit_id and item_unit_id != selected_unit_id:
                    table.setRowHidden(row, True)
                    continue

                # Mapear para assemblage (amostra) usando unit_to_assemblage
                item_assemblage_id = None
                if hasattr(self, 'unit_to_assemblage'):
                    item_assemblage_id = self.unit_to_assemblage.get(item_unit_id)

                # Se o filtro estiver na amostra (Assemblage), compara com assemblage_id da unidade
                if selected_sample_id and (not item_assemblage_id or item_assemblage_id != selected_sample_id):
                    table.setRowHidden(row, True)
                    continue

                if selected_col_id:
                    sample_col_id = None
                    if item_assemblage_id:
                        sample_col_id = self.assemblage_to_collection.get(item_assemblage_id)
                    if sample_col_id != selected_col_id:
                        table.setRowHidden(row, True)
                        continue

                if selected_site_id:
                    sample_col_id = None
                    if item_assemblage_id:
                        sample_col_id = self.assemblage_to_collection.get(item_assemblage_id)
                    col_site_id = self.collection_to_site.get(sample_col_id)
                    if col_site_id != selected_site_id:
                        table.setRowHidden(row, True)
                        continue

            # Filtros Hierárquicos para Níveis
            if table == self.table_levels:
                lvl_selected_site = None
                lvl_selected_col = None
                lvl_selected_sample = None
                lvl_selected_unit = None
                
                if hasattr(self, 'combo_filter_site_lvl') and self.combo_filter_site_lvl.currentData():
                    lvl_selected_site = str(self.combo_filter_site_lvl.currentData())
                if hasattr(self, 'combo_filter_collection_lvl') and self.combo_filter_collection_lvl.currentData():
                    lvl_selected_col = str(self.combo_filter_collection_lvl.currentData())
                if hasattr(self, 'combo_filter_sample_lvl') and self.combo_filter_sample_lvl.currentData():
                    lvl_selected_sample = str(self.combo_filter_sample_lvl.currentData())
                if hasattr(self, 'combo_filter_unit_lvl') and self.combo_filter_unit_lvl.currentData():
                    lvl_selected_unit = str(self.combo_filter_unit_lvl.currentData())

                # Coluna 1 = ID da Unidade
                item_unit_id_widget = table.item(row, 1)
                item_unit_id = str(item_unit_id_widget.data(Qt.UserRole)) if item_unit_id_widget else ""

                if lvl_selected_unit and item_unit_id != lvl_selected_unit:
                    table.setRowHidden(row, True)
                    continue
                
                if hasattr(self, 'unit_to_assemblage'):
                    unit_sample_id = self.unit_to_assemblage.get(item_unit_id)
                    
                    if lvl_selected_sample and unit_sample_id != lvl_selected_sample:
                        table.setRowHidden(row, True)
                        continue
                        
                    if lvl_selected_col:
                        sample_col_id = self.assemblage_to_collection.get(unit_sample_id)
                        if sample_col_id != lvl_selected_col:
                            table.setRowHidden(row, True)
                            continue
                            
                    if lvl_selected_site:
                        sample_col_id = self.assemblage_to_collection.get(unit_sample_id)
                        col_site_id = self.collection_to_site.get(sample_col_id)
                        if col_site_id != lvl_selected_site:
                            table.setRowHidden(row, True)
                            continue

            # Filtros Hierárquicos para Unidades
            if table == self.table_units:
                unit_selected_site = None
                unit_selected_col = None
                unit_selected_sample = None
                if hasattr(self, 'combo_filter_site_unit') and self.combo_filter_site_unit.currentData():
                    unit_selected_site = str(self.combo_filter_site_unit.currentData())
                if hasattr(self, 'combo_filter_collection_unit') and self.combo_filter_collection_unit.currentData():
                    unit_selected_col = str(self.combo_filter_collection_unit.currentData())
                if hasattr(self, 'combo_filter_sample_unit') and self.combo_filter_sample_unit.currentData():
                    unit_selected_sample = str(self.combo_filter_sample_unit.currentData())

                # Coluna 1 = ID da Amostra (Assemblage ID)
                item_assemblage_id_widget = table.item(row, 1)
                item_assemblage_id = str(item_assemblage_id_widget.data(Qt.UserRole)) if item_assemblage_id_widget else ""

                if unit_selected_sample and item_assemblage_id != unit_selected_sample:
                    table.setRowHidden(row, True)
                    continue

                if unit_selected_col:
                    sample_col_id = self.assemblage_to_collection.get(item_assemblage_id)
                    if sample_col_id != unit_selected_col:
                        table.setRowHidden(row, True)
                        continue

                if unit_selected_site:
                    sample_col_id = self.assemblage_to_collection.get(item_assemblage_id)
                    col_site_id = self.collection_to_site.get(sample_col_id)
                    if col_site_id != unit_selected_site:
                        table.setRowHidden(row, True)
                        continue

            # Filtros Hierárquicos para Amostras
            if table == self.table_samples:
                samp_selected_site = None
                samp_selected_col = None
                samp_selected_sample = None
                if hasattr(self, 'combo_filter_site_samp') and self.combo_filter_site_samp.currentData():
                    samp_selected_site = str(self.combo_filter_site_samp.currentData())
                if hasattr(self, 'combo_filter_collection_samp') and self.combo_filter_collection_samp.currentData():
                    samp_selected_col = str(self.combo_filter_collection_samp.currentData())
                if hasattr(self, 'combo_filter_sample_samp') and self.combo_filter_sample_samp.currentData():
                    samp_selected_sample = str(self.combo_filter_sample_samp.currentData())

                # Coluna 0 = ID da Amostra, Coluna 1 = collection_id
                item_sample_id_widget = table.item(row, 0)
                item_sample_id = item_sample_id_widget.text() if item_sample_id_widget else ""
                item_collection_id_widget = table.item(row, 1)
                item_collection_id = str(item_collection_id_widget.data(Qt.UserRole)) if item_collection_id_widget else ""

                if samp_selected_sample and item_sample_id != samp_selected_sample:
                    table.setRowHidden(row, True)
                    continue

                if samp_selected_col and item_collection_id != samp_selected_col:
                    table.setRowHidden(row, True)
                    continue

                if samp_selected_site:
                    col_site_id = self.collection_to_site.get(item_collection_id)
                    if col_site_id != samp_selected_site:
                        table.setRowHidden(row, True)
                        continue

            # Filtros Hierárquicos para Coleções
            if table == self.table_collections:
                col_selected_site = None
                col_selected_collection = None
                if hasattr(self, 'combo_filter_site_col') and self.combo_filter_site_col.currentData():
                    col_selected_site = str(self.combo_filter_site_col.currentData())
                if hasattr(self, 'combo_filter_collection_col') and self.combo_filter_collection_col.currentData():
                    col_selected_collection = str(self.combo_filter_collection_col.currentData())

                # Coluna 0 = ID, Coluna 1 = site_id
                item_collection_id_widget = table.item(row, 0)
                item_collection_id = item_collection_id_widget.text() if item_collection_id_widget else ""
                item_site_id_widget = table.item(row, 1)
                item_site_id = str(item_site_id_widget.data(Qt.UserRole)) if item_site_id_widget else ""

                if col_selected_collection and item_collection_id != col_selected_collection:
                    table.setRowHidden(row, True)
                    continue

                if col_selected_site and item_site_id != col_selected_site:
                    table.setRowHidden(row, True)
                    continue

            # Se o texto do filtro estiver vazio, mostra tudo
            if not filter_text:
                match = True
            else:
                for column in range(table.columnCount()):
                    item = table.item(row, column)
                    # Verifica se o item existe e se o texto corresponde
                    if item:
                        item_text = item.text().lower()
                        if exact_match:
                            if item_text == filter_text:
                                match = True
                                break
                        else:
                            if filter_text in item_text:
                                match = True
                                break
            table.setRowHidden(row, not match)

    # Função para limpar o filtro
    def clear_filter(self):
        table, filter_input = self.get_active_filter_widgets()
        
        if not table or not filter_input:
            return

        filter_input.clear()
        for row in range(table.rowCount()):
            table.setRowHidden(row, False)

    # --- Funções de Exportação (Integradas do teste.py) ---

    def export_pdf(self):
        """Gera um PDF com todos os materiais."""
        if not self.db:
            return
        file_path, _ = QFileDialog.getSaveFileName(None, "Salvar PDF", "Todos_Materiais.pdf", "PDF Files (*.pdf)")
        if not file_path:
            return
        
        # Obter IDs visíveis
        visible_ids = []
        for row in range(self.table_material.rowCount()):
            if not self.table_material.isRowHidden(row):
                item = self.table_material.item(row, 0)
                if item:
                    try:
                        visible_ids.append(int(item.text()))
                    except ValueError:
                        pass

        # Obter nomes das colunas da tabela para o PDF
        colunas = [self.table_material.horizontalHeaderItem(i).text() for i in range(self.table_material.columnCount())]
        success, msg = io_handlers.export_pdf(self.db, file_path, colunas, visible_ids)
        if success:
            QMessageBox.information(None, "Sucesso", msg)
        else:
            QMessageBox.critical(None, "Erro", msg)

    def export_units_pdf(self):
        """Gera um PDF com todas as unidades e níveis."""
        if not self.db:
            return
        file_path, _ = QFileDialog.getSaveFileName(None, "Salvar PDF de Unidades", "Unidades_Escavacao.pdf", "PDF Files (*.pdf)")
        if not file_path:
            return
        
        # Obter IDs visíveis
        visible_ids = []
        for row in range(self.table_units.rowCount()):
            if not self.table_units.isRowHidden(row):
                item = self.table_units.item(row, 0)
                if item:
                    try:
                        visible_ids.append(int(item.text()))
                    except ValueError:
                        pass

        success, msg = io_handlers.export_units_pdf(self.db, file_path, visible_ids)
        if success:
            QMessageBox.information(None, "Sucesso", msg)
        else:
            QMessageBox.critical(None, "Erro", msg)

    def export_xlsx(self):
        """Exporta todo o banco de dados para Excel."""
        if not self.db:
            return
        file_path, _ = QFileDialog.getSaveFileName(None, "Salvar Excel", "Banco_Completo.xlsx", "Excel Files (*.xlsx)")
        if not file_path:
            return
        
        success, msg = io_handlers.export_xlsx(self.db, file_path)
        if success:
            QMessageBox.information(None, "Sucesso", msg)
        else:
            QMessageBox.critical(None, "Erro", msg)

    def update_database_from_excel(self):
        """Atualiza o banco de dados atual a partir de um arquivo Excel."""
        if not self.db:
            QMessageBox.warning(None, "Erro", "Nenhum banco de dados aberto.")
            return
        xlsx_path, _ = QFileDialog.getOpenFileName(None, "Selecionar Arquivo Excel para Atualização", "", "Excel Files (*.xlsx)")
        if not xlsx_path:
            return
        
        success, msg = io_handlers.update_database_from_excel(self.db, xlsx_path)
        if success:
            QMessageBox.information(None, "Sucesso", msg)
            self.load_sites()
            self.load_collections()
            self.load_samples()
            self.load_units()
            self.load_levels()
            self.load_material()
            self.populate_filter_dropdowns()
            self.populate_stats_dropdowns()
        else:
            QMessageBox.critical(None, "Erro", msg)

    def import_database_from_excel(self):
        """Cria ou atualiza um banco de dados a partir de um arquivo Excel."""
        xlsx_path, _ = QFileDialog.getOpenFileName(None, "Selecionar Arquivo Excel", "", "Excel Files (*.xlsx)")
        if not xlsx_path:
            return

        db_path, _ = QFileDialog.getSaveFileName(None, "Salvar/Selecionar Banco de Dados", "", "SQLite Database (*.db)")
        if not db_path:
            return
            
        # Garante a extensão .db
        if not db_path.lower().endswith(".db"):
            db_path += ".db"

        success, msg = io_handlers.import_database_from_excel(xlsx_path, db_path)
        
        if success:
            QMessageBox.information(None, "Sucesso", msg)
            # Abre o banco de dados na interface
            self.db = Database(db_path)
            self.load_sites()
            self.load_collections()
            self.load_samples()
            self.load_units()
            self.load_levels()
            self.load_material()
            self.populate_filter_dropdowns()
            self.populate_stats_dropdowns()
            self.stackedWidget.setCurrentWidget(self.DashboardPage)
        else:
            QMessageBox.critical(None, "Erro", msg)