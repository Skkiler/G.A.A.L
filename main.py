#main.py
import sys, os, json, re
from functools import partial
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QLineEdit, QPushButton, QCheckBox,
    QRadioButton, QButtonGroup, QSlider, QTableWidget,
    QTableWidgetItem, QFileDialog, QScrollArea, QAbstractItemView,
    QHeaderView, QProgressBar, QMessageBox, QComboBox, QInputDialog,
    QDialog, QColorDialog, QFormLayout, QListWidget, QSizePolicy, QStackedWidget, QFrame
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QPropertyAnimation, QSequentialAnimationGroup, QEasingCurve, QSize, QTimeLine
from PySide6.QtGui import QColor, QIcon, QMovie, QPixmap
import pandas as pd

# ================== Splash Integrada ===================
def show_splash_and_run(main_window_class):
    app = QApplication(sys.argv)
    app.setApplicationName("GAELIS")
    icon_path = 'splash.png'
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    splash_widget = QWidget(None, Qt.SplashScreen | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
    splash_widget.setAttribute(Qt.WA_TranslucentBackground)
    layout = QVBoxLayout(splash_widget)
    splash_label = QLabel(splash_widget)
    splash_label.setAlignment(Qt.AlignCenter)

    screen = app.primaryScreen().availableGeometry()
    max_w = screen.width() // 7  # um pouco maior
    gif_path = 'splash.gif'
    if os.path.exists(gif_path):
        movie = QMovie(gif_path)
        movie.jumpToFrame(0)
        orig = movie.currentPixmap().size()
        scaled_h = int(orig.height() * (max_w / orig.width()))
        movie.setScaledSize(QSize(max_w, scaled_h))
        splash_label.setMovie(movie)
        splash_label.setFixedSize(max_w, scaled_h)
        movie.start()
    else:
        pix = QPixmap('splash.png')
        pix = pix.scaledToWidth(max_w, Qt.SmoothTransformation)
        splash_label.setPixmap(pix)
        splash_label.setFixedSize(pix.size())
    layout.addWidget(splash_label)
    splash_widget.setLayout(layout)
    splash_widget.adjustSize()
    geom = splash_widget.frameGeometry()
    center_point = screen.center()
    geom.moveCenter(center_point)
    splash_widget.move(geom.topLeft())
    splash_widget.show()

    fade_in = QPropertyAnimation(splash_widget, b"windowOpacity")
    fade_in.setDuration(800)
    fade_in.setStartValue(0.0)
    fade_in.setEndValue(1.0)
    fade_in.setEasingCurve(QEasingCurve.InOutQuad)
    fade_out = QPropertyAnimation(splash_widget, b"windowOpacity")
    fade_out.setDuration(900)
    fade_out.setStartValue(1.0)
    fade_out.setEndValue(0.0)
    fade_out.setEasingCurve(QEasingCurve.InOutQuad)
    seq = QSequentialAnimationGroup()
    seq.addAnimation(fade_in)
    seq.addPause(1200)
    seq.addAnimation(fade_out)

    def start_app():
        splash_widget.close()
        window = main_window_class()
        window.show()

    seq.finished.connect(start_app)
    QTimer.singleShot(0, seq.start)
    # Só um app.exec()
    sys.exit(app.exec())

# ================== Utilidade Flag + Info ===================
FLAG_INFOS = {
    "chk_sub": (
        "Cria subpastas de destino com base nas condições definidas para cada arquivo processado. "
        "Se marcada junto com 'Hierarquia', cria subpastas aninhadas conforme os campos principais das condições. "
        "Se marcada sem 'Hierarquia', cria uma única subpasta (por condições) para cada arquivo. "
        "Se desmarcada, mantém a estrutura física da origem caso 'Hierarquia' esteja marcada."
    ),
    "chk_hierarchy": (
        "Mantém a estrutura de pastas original da origem ao copiar/mover arquivos. "
        "Se 'Criar Subpasta' estiver marcada, a hierarquia é baseada nas condições e campos principais, "
        "em vez da estrutura física original."
    ),
    "chk_multiply": (
        "Permite que um mesmo arquivo seja copiado/movido para múltiplos destinos, "
        "caso satisfaça mais de uma condição."
    ),
    "chk_copydirs": (
        "Inclui também pastas e subpastas (além dos arquivos) na cópia/movimentação."
    ),
    "chk_sobra": (
        "Define uma pasta específica para arquivos que não se encaixam em nenhuma condição. "
        "A hierarquia física será mantida apenas se 'Hierarquia' estiver ativa e 'Criar Subpasta' estiver desmarcada."
    ),
    "chk_recursive": (
        "Busca arquivos dentro de todas as subpastas, além da pasta principal de origem."
    ),
    "chk_extract": (
        "Extrai arquivos de dentro de arquivos ZIP, caso existam, antes de processar."
    ),
    "chk_zip": (
        "Compacta toda a pasta de destino em um arquivo ZIP ao final do processamento."
    ),
    "chk_findsub": (
        "Procura uma subpasta existente com o nome correspondente às condições principais, e move o arquivo para ela. Se não encontrar, pode criar ou copiar para pasta sobra (veja as outras opções)."
    )
}

def add_flag_with_info(layout, checkbox, info_text):
    row = QHBoxLayout()
    row.addWidget(checkbox)
    btn_info = QPushButton("(!)")
    btn_info.setObjectName("infoButton")
    btn_info.setFixedSize(24, 24)
    btn_info.setToolTip("Clique para ver detalhes sobre esta opção")
    btn_info.clicked.connect(lambda: QMessageBox.information(checkbox, "Informação", info_text))
    row.addWidget(btn_info)
    row.addStretch()
    layout.addLayout(row)

# ================== SETTINGS E RESTANTE DA APP ===================

SETTINGS_PATH = Path(__file__).parent / "settings.json"

def carregar_settings():
    defaults = {
        "theme": "Dark",
        "themes": {
            "Dark": {
                "bg_start": "#2f2f2f", "bg_end": "#3f3f3f",
                "text": "#e5e5e5",    "input_bg": "#141111",
                "btn": "#cc2727",     "btn_hover": "#e11717",
                "btn_text": "#ffffff","checkbox": "#e11717",
                "slider": "#e11717",  "cond_selected": "#505050",
                "box_border": "#888888"
            },
            "Light": {
                "bg_start": "#e1e0e0", "bg_end": "#eeeeee",
                "text": "#183544",     "input_bg": "#eef0f2",
                "btn": "#3f7ad1",      "btn_hover": "#347de9",
                "btn_text": "#ffffff", "checkbox": "#347de9",
                "slider": "#347de9",   "cond_selected": "#cccccc",
                "box_border": "#444444"
            }
        }
    }
    if SETTINGS_PATH.exists():
        try:
            s = json.load(SETTINGS_PATH.open("r", encoding="utf-8"))
            for k, v in defaults.items():
                s.setdefault(k, v)
            for name, props in defaults["themes"].items():
                s["themes"].setdefault(name, {}).update(props)
            if s["theme"] not in s["themes"]:
                s["theme"] = next(iter(s["themes"]))
            return s
        except:
            return defaults
    return defaults

def salvar_settings(s):
    with SETTINGS_PATH.open("w", encoding="utf-8") as f:
        json.dump(s, f, indent=2, ensure_ascii=False)

# --- THREAD DE EXECUÇÃO COM CANCELAMENTO E RELATÓRIO ---

class ExecutorThread(QThread):
    progress = Signal(int, int)    # Processados, total
    error    = Signal(str)
    finished = Signal(list)        # Lista de dicionários de relatório
    canceled = Signal()

    def __init__(self, config: dict):
        super().__init__()
        self.config = config
        self.cancel_requested = False
        self._report = []

    def run(self):
        try:
            from executor import Executor
        except Exception as e:
            self.error.emit(f"Erro ao importar executor: {e}")
            return
        executor = Executor(
            self.config,
            max_workers=self.config.get("max_workers", 4),
            progress_callback=self._progress_callback,
            error_callback=lambda e: self.error.emit(str(e)),
            cancel_checker=lambda: self.cancel_requested,
            report_callback=self._append_report
        )
        executor.run()
        if self.cancel_requested:
            self.canceled.emit()
        else:
            self.finished.emit(self._report)

    def cancel(self):
        self.cancel_requested = True

    def _progress_callback(self, p, t):
        self.progress.emit(p, t)

    def _append_report(self, item):
        self._report.append(item)

# ========== Report Dialog com Exportação ==========

class ReportDialog(QDialog):
    
    def __init__(self, report, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Relatório de Execução")
        self.report = report
        layout = QVBoxLayout(self)
        self.table = QTableWidget(len(report), 4)
        self.table.setHorizontalHeaderLabels(["Arquivo", "Origem", "Destino", "Ação"])
        for i, r in enumerate(report):
            self.table.setItem(i, 0, QTableWidgetItem(r.get("arquivo", "")))
            self.table.setItem(i, 1, QTableWidgetItem(r.get("origem", "")))
            self.table.setItem(i, 2, QTableWidgetItem(r.get("destino", "")))
            self.table.setItem(i, 3, QTableWidgetItem(r.get("acao", "")))
        self.table.resizeColumnsToContents()
        self.table.setSelectionMode(QAbstractItemView.NoSelection)
        layout.addWidget(self.table)
        h = QHBoxLayout()
        btn_save = QPushButton("Salvar Excel")
        btn_save.clicked.connect(self.save_excel)
        btn_import = QPushButton("Importar Excel")
        btn_import.clicked.connect(self.import_excel)
        btn_close = QPushButton("Fechar")
        btn_close.clicked.connect(self.accept)
        h.addWidget(btn_save)
        h.addWidget(btn_import)
        h.addStretch()
        h.addWidget(btn_close)
        layout.addLayout(h)

    def save_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Salvar relatório como", "relatorio.xlsx", "Excel (*.xlsx)")
        if path:
            df = pd.DataFrame(self.report)
            df.to_excel(path, index=False)

    def import_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir relatório Excel", "", "Excel (*.xlsx *.xls)")
        if path:
            df = pd.read_excel(path)
            self.table.setRowCount(len(df))
            for i, (_, row) in enumerate(df.iterrows()):
                self.table.setItem(i, 0, QTableWidgetItem(str(row.get("arquivo", ""))))
                self.table.setItem(i, 1, QTableWidgetItem(str(row.get("origem", ""))))
                self.table.setItem(i, 2, QTableWidgetItem(str(row.get("destino", ""))))
                self.table.setItem(i, 3, QTableWidgetItem(str(row.get("acao", ""))))

class ThemeEditorDialog(QDialog):
    
    def __init__(self, name, props, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Editar Tema: {name}")
        self.props = props.copy()
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.buttons = {}
        fields = {
            "Fundo Início": "bg_start", "Fundo Fim": "bg_end",
            "Texto": "text",        "Entrada": "input_bg",
            "Botão": "btn",         "Botão Hover": "btn_hover",
            "Texto Botão": "btn_text","Checkbox": "checkbox",
            "Slider": "slider",     "Condição Selecionada": "cond_selected",
            "Contorno Boxes": "box_border"
        }
        for label, key in fields.items():
            btn = QPushButton()
            btn.setFixedSize(30, 30)
            btn.clicked.connect(partial(self._pick_color, key))
            form.addRow(label + ":", btn)
            self.buttons[key] = btn
        layout.addLayout(form)

        self.preview = QWidget()
        self.preview.setFixedHeight(100)
        layout.addWidget(self.preview)

        row = QHBoxLayout()
        ok     = QPushButton("Aplicar");    ok.clicked.connect(self.accept)
        cancel = QPushButton("Cancelar"); cancel.clicked.connect(self.reject)
        row.addStretch(); row.addWidget(ok); row.addWidget(cancel)
        layout.addLayout(row)
        self._update_ui()

    def _pick_color(self, key):
        inicial = QColor(self.props.get(key, "#ffffff"))
        col     = QColorDialog.getColor(inicial, self)
        if col.isValid():
            self.props[key] = col.name()
            self._update_ui()

    def _update_ui(self):
        for k, btn in self.buttons.items():
            btn.setStyleSheet(f"background:{self.props[k]}")
        grad = (
            f"qlineargradient(x1:0,y1:0,x2:1,y2:0,"
            f" stop:0 {self.props['bg_start']}, stop:1 {self.props['bg_end']})"
        )
        qss = (
            f"QWidget#preview {{ background: {grad}; border:1px solid {self.props['box_border']}; }}\n"
            f"QWidget {{ color: {self.props['text']}; }}\n"
            f"QLineEdit {{ background: {self.props['input_bg']}; }}\n"
            f"QPushButton {{ background: {self.props['btn']}; color: {self.props['btn_text']}; }}\n"
            f"QPushButton:hover {{ background: {self.props['btn_hover']}; }}\n"
            f"QCheckBox::indicator:checked {{ background: {self.props['checkbox']}; }}\n"
            f"QTableWidget::item:selected {{ background: {self.props['cond_selected']}; }}\n"
            f"QGroupBox {{ border:1px solid {self.props['box_border']}; margin-top:10px; }}")
        self.preview.setStyleSheet(qss)

class ThemeManagerDialog(QDialog):
    
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Gerenciar Temas")
        self.parent = parent
        self.themes = parent.themes
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        self.list = QListWidget()
        self.list.addItems(self.themes.keys())
        layout.addWidget(self.list)
        btns = QHBoxLayout()
        btns.addWidget(QPushButton("➕ Novo",    clicked=self._new))
        btns.addWidget(QPushButton("✏️ Editar", clicked=self._edit))
        btns.addWidget(QPushButton("🗑️ Deletar",clicked=self._delete))
        btns.addStretch()
        btns.addWidget(QPushButton("Fechar", clicked=self.accept))
        layout.addLayout(btns)

    def _new(self):
        name, ok = QInputDialog.getText(self, "Novo Tema", "Nome:")
        if ok and name and name not in self.themes:
            cur = self.parent.current_theme
            self.themes[name] = self.themes[cur].copy()
            self.list.addItem(name)
            salvar_settings(self.parent.settings)
        self.parent.update_theme_combo()

    def _edit(self):
        item = self.list.currentItem()
        if not item: return
        key = item.text()
        dlg = ThemeEditorDialog(key, self.themes[key], self)
        if dlg.exec() == QDialog.Accepted:
            self.themes[key] = dlg.props
            salvar_settings(self.parent.settings)
            if key == self.parent.current_theme:
                self.parent._apply_theme(key)
        self.parent.update_theme_combo()

    def _delete(self):
        item = self.list.currentItem()
        if not item: return
        key = item.text()
        if len(self.themes) <= 1:
            QMessageBox.warning(self, "Aviso", "Não pode excluir o único tema.")
            return
        if QMessageBox.question(self, "Confirmar", f"Excluir '{key}'?") == QMessageBox.Yes:
            del self.themes[key]
            salvar_settings(self.parent.settings)
            self.list.takeItem(self.list.row(item))
            if key == self.parent.current_theme:
                new = next(iter(self.themes))
                self.parent.combo_theme.blockSignals(True)
                self.parent.combo_theme.setCurrentText(new)
                self.parent.combo_theme.blockSignals(False)
                self.parent._apply_theme(new)
        self.parent.update_theme_combo()

class DropLineEdit(QLineEdit):
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptDrops(True)
        self.setMinimumWidth(120)
        self.setMaximumWidth(300)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.accept()
        else:
            super().dragEnterEvent(e)
    
    def dropEvent(self, e):
        urls = e.mimeData().urls()
        if urls:
            self.setText(urls[0].toLocalFile())
        e.accept()

class ConditionTable(QTableWidget):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.verticalHeader().setVisible(False)
        self.setColumnCount(4)
        self.setHorizontalHeaderLabels(["", "Condição", "Índice", "Principal"])

        self.setColumnWidth(0, 28)

        m = self.model()
        m.rowsInserted.connect(self.refresh_move_column)
        m.rowsRemoved.connect(self.refresh_move_column)
        self.selectionModel().currentRowChanged.connect(lambda *_: self.refresh_move_column())

    def refresh_move_column(self):
        """Coloca em cada célula da coluna 0 o widget com as duas setas."""
        for row in range(self.rowCount()):
            # limpa widget antigo, se houver
            if self.cellWidget(row, 0):
                self.removeCellWidget(row, 0)

            # container horizontal sem espaçamento
            w = QWidget()
            hb = QHBoxLayout(w)
            hb.setContentsMargins(0, 4, 0, 4)  # 4px padding vertical para centralizar
            hb.setSpacing(0)

            # seta pra cima
            up = QPushButton("▲", w)
            up.setObjectName("arrowBtn")
            up.setFixedSize(8, 11)
            up.clicked.connect(partial(self.window().move_condition, row, row - 1))

            # seta pra baixo
            dn = QPushButton("▼", w)
            dn.setObjectName("arrowBtn")
            dn.setFixedSize(8, 11)
            dn.clicked.connect(partial(self.window().move_condition, row, row + 1))

            hb.addWidget(up)
            hb.addWidget(dn)

            w.setFixedWidth(28)
            self.setCellWidget(row, 0, w)

        # força repintar
        self.viewport().update()

class OrigemDialog(QDialog):
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Tipo de Origem")
        self.setModal(True)
        self.selecionado = None

        layout = QVBoxLayout(self)
        label = QLabel("O que deseja selecionar?")
        label.setWordWrap(True)
        label.setAlignment(Qt.AlignLeft)
        layout.addWidget(label)

        btn_layout = QHBoxLayout()
        self.btn_pasta = QPushButton("Pasta")
        self.btn_zip = QPushButton("Arquivo ZIP")
        self.btn_cancelar = QPushButton("Cancelar")
        # Ajusta tamanho pelo texto + padding generoso
        for btn in [self.btn_pasta, self.btn_zip, self.btn_cancelar]:
            btn.setMinimumWidth(btn.fontMetrics().boundingRect(btn.text()).width() + 38)
            btn.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Preferred)
        btn_layout.addWidget(self.btn_pasta)
        btn_layout.addWidget(self.btn_zip)
        btn_layout.addWidget(self.btn_cancelar)
        layout.addLayout(btn_layout)

        # Eventos
        self.btn_pasta.clicked.connect(self._escolhe_pasta)
        self.btn_zip.clicked.connect(self._escolhe_zip)
        self.btn_cancelar.clicked.connect(self.reject)

        self.setFixedHeight(self.sizeHint().height() + 12)
        self.setFixedWidth(max(320, self.sizeHint().width() + 12))

    def _escolhe_pasta(self):
        self.selecionado = "pasta"
        self.accept()
    
    def _escolhe_zip(self):
        self.selecionado = "zip"
        self.accept()

# ========== MainWindow Alterações ==========

class MainWindow(QMainWindow):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("G.A.A.L - Gerenciador Avançado de Arquivos por Lógica")
        self.resize(950, 750)
        self.settings = carregar_settings()
        self.current_theme = self.settings["theme"]
        self.themes = self.settings["themes"]
        self.max_workers = os.cpu_count() or 4

        scroll = QScrollArea()
        container = QWidget()
        container.setObjectName("centralwidget")
        self.layout = QVBoxLayout(container)
        self.layout.setContentsMargins(15, 15, 15, 15)
        self.layout.setSpacing(10)
        scroll.setWidgetResizable(True)
        scroll.setWidget(container)
        self.setCentralWidget(scroll)

        self._last_edit = None
        QApplication.instance().focusChanged.connect(self._on_focus_changed)

        self._setup_theme_header()
        self.init_paths_group()
        self.init_options_group()
        self.init_threads_group()
        self.init_conditions_group()
        self.init_execution_group()
        self.layout.addStretch()
        self._apply_theme(self.current_theme)
        self.thread = None

    def validate_config(self, cfg) -> bool:
        # 1) pelo menos uma origem
        if not cfg["origens"]:
            QMessageBox.warning(self, "Validação", "Você precisa informar ao menos uma origem.")
            return False

        # 2) cada origem deve existir
        for o in cfg["origens"]:
            if not Path(o).exists():
                QMessageBox.warning(self, "Validação", f"Origem inválida:\n{o}")
                return False

        # 3) destino preenchido, existente e gravável
        dest = cfg["destino"].strip()
        if not dest:
            QMessageBox.warning(self, "Validação", "Informe uma pasta de destino.")
            return False
        dest_path = Path(dest)
        if not dest_path.exists():
            QMessageBox.warning(self, "Validação", f"Pasta de destino inválida:\n{dest}")
            return False
        if not os.access(dest, os.W_OK):
            QMessageBox.warning(self, "Validação", f"Sem permissão de escrita em destino:\n{dest}")
            return False

        # 4) se condições estão ativas mas nenhuma condição nem expressão
        if cfg["use_conditions"]:
            no_cols = not cfg["colunas"]
            no_expr = not cfg["condition_expression"].strip()
            if no_cols and no_expr:
                QMessageBox.warning(
                    self, "Validação",
                    "Você ativou o uso de condições, mas não definiu nenhuma condição\n"
                    "nem escreveu nenhuma expressão lógica."
                )
                return False

        # 5) valida formato das datas, se marcadas
        def ok_date(txt):
            return bool(re.match(r"^\d{2}-\d{2}-\d{4}$", txt) or
                        re.match(r"^\d{4}-\d{2}-\d{2}$", txt))
        if self.chk_dt1.isChecked() and not ok_date(self.in_dt1.text().strip()):
            QMessageBox.warning(self, "Validação", "Data inicial inválida. Use dd-mm-aaaa ou aaaa-mm-dd.")
            return False
        if self.chk_dt2.isChecked() and not ok_date(self.in_dt2.text().strip()):
            QMessageBox.warning(self, "Validação", "Data final inválida. Use dd-mm-aaaa ou aaaa-mm-dd.")
            return False

        # 6) se Excel, cheque arquivo
        if cfg["use_conditions"] and cfg["condition_mode"] == "excel":
            path = cfg["excel"].strip()
            if not path or not Path(path).is_file():
                QMessageBox.warning(self, "Validação", "Arquivo Excel de condições inválido ou não selecionado.")
                return False

        # 7) se modo pastas, cheque existência da pasta de condições
        if cfg["use_conditions"] and cfg["condition_mode"] == "folders":
            path = cfg["cond_folder"].strip()
            if not path or not Path(path).is_dir():
                QMessageBox.warning(self, "Validação", "Pasta de condições inválida ou não selecionada.")
                return False

        # tudo ok
        return True

    # ─── Tema ──────────────────────────────────────────────────────────────
    
    def _setup_theme_header(self):
        h = QHBoxLayout()
        h.setSpacing(8)
        h.setContentsMargins(0, 0, 0, 0)
        h.addWidget(QLabel("Tema:"))
        self.combo_theme = QComboBox()
        self.combo_theme.addItems(self.themes.keys())
        self.combo_theme.setCurrentText(self.current_theme)
        self.combo_theme.currentTextChanged.connect(self._apply_theme)
        h.addWidget(self.combo_theme)
        btn = QPushButton("Gerenciar Temas")
        btn.setFixedHeight(28)
        btn.setFixedWidth(140)
        btn.clicked.connect(lambda: ThemeManagerDialog(self).exec())
        h.addWidget(btn)
        h.addStretch()
        self.layout.addLayout(h)

    def _apply_theme(self, name):
        # --- Função auxiliar para interpolar cores ---
        def lerp_color(a: str, b: str, t: float) -> str:
            c1, c2 = QColor(a), QColor(b)
            r = c1.red() + (c2.red() - c1.red()) * t
            g = c1.green() + (c2.green() - c1.green()) * t
            b = c1.blue() + (c2.blue() - c1.blue()) * t
            return QColor(int(r), int(g), int(b)).name()

        def adjust_color_contrast(color: str) -> str:
            """Clareia se o fundo for escuro, escurece se o fundo for claro."""
            c = QColor(color)
            luminance = (0.299 * c.red() + 0.587 * c.green() + 0.114 * c.blue())
            factor = 1.2 if luminance < 128 else 0.8
            r = min(max(int(c.red() * factor), 0), 255)
            g = min(max(int(c.green() * factor), 0), 255)
            b = min(max(int(c.blue() * factor), 0), 255)
            return QColor(r, g, b).name()

        # --- Seleção do tema ---
        if name not in self.themes:
            name = next(iter(self.themes))
        new_th = self.themes[name]
        old_th = self.themes.get(self.current_theme, new_th)

        # --- Função que aplica QSS interpolando cores ---
        def aplicar_qss(t: float):
            th = {k: lerp_color(old_th[k], new_th[k], t) for k in new_th.keys()}
            slider_dark = QColor(th['slider']).darker(150).name()
            arrow_bg = adjust_color_contrast(th['input_bg'])
            grad = f"qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 {th['bg_start']}, stop:1 {th['bg_end']})"
            qss = f"""
            QMainWindow, QWidget#centralwidget {{
                background: {grad};
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 12px;
            }}
            QPushButton#arrowBtn {{
                background: transparent;
                color: {th['btn_text']};
                border: none;
                padding: 0px;
                margin: 0px;
                min-width: 11px;
                min-height: 8px;
            }}
            QPushButton#arrowBtn:hover, QPushButton#arrowBtn:pressed {{
                color: {th['btn_hover']};
            }}
            QTableWidget::item:selected {{
                background-color: {th['cond_selected']};
                color: {th['text']}
            }}
            QGroupBox {{
                border: 1px solid {th['box_border']};
                border-radius: 8px;
                margin-top: 22px;
                padding: 14px 12px 12px 12px;
                background: rgba(0,0,0,0.03);
            }}
            QGroupBox:title {{
                subcontrol-origin: content;
                subcontrol-position: top left;
                left: 12px;
                top: -14px;
                background: transparent;
                color: {th['text']};
                font-weight: bold;
                font-size: 13px;
                padding: 0 8px;
            }}
            QLabel, QCheckBox, QRadioButton {{
                color: {th['text']};
                font-size: 12px;
            }}
            QLineEdit, QComboBox, QTableWidget {{
                background: {th['input_bg']};
                color: {th['text']};
                border: 1.2px solid {th['box_border']};
                border-radius: 8px;
                padding: 2px 6px;
                min-height: 18px;
                font-size: 12px;
            }}
            QComboBox::drop-down {{
                border: none;
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 24px;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
                background: {arrow_bg};
            }}
            QComboBox::down-arrow {{
                width: 0px;
                height: 0px;
                border-left: 6px solid transparent;
                border-right: 6px solid transparent;
                border-top: 7px solid {th['text']};
                margin-right: 7px;
            }}
            QFrame#tableCondContainer {{
                border-radius: 10px;
                border: 1.5px solid {th['box_border']};
                background: {th['input_bg']};
                padding: 0px;
                margin: 1px 1px 1px 1px;
            }}
            QTableWidget {{
                background: transparent;
                border: none;
                border-radius: 10px;
                selection-background-color: {th['cond_selected']};
                outline: none;
                font-size: 12px;
                gridline-color: {th['box_border']};
            }}
            QTableWidget::item {{
                min-height: 22px;
                padding: 4px 4px;
            }}
            QHeaderView::section {{
                background: {th['input_bg']};
                color: {th['text']};
                border: none;
                padding: 6px 0 6px 0;
                font-weight: bold;
                font-size: 12px;
                min-height: 22px;
                max-height: 24px;
            }}
            QPushButton#infoButton {{
                color: {th['text']};
                background: transparent;
                font-size: 13px;
                font-family: 'Consolas', 'Segoe UI', Arial, sans-serif;
                font-weight: 500;
                border: none;
            }}
            QPushButton {{
                background-color: {th['btn']};
                color: {th['btn_text']};
                border-radius: 7px;
                padding: 3px 14px;
                font-weight: 500;
                min-width: 22px;
                min-height: 16px;
                font-size: 12px;
                border: none;
            }}
            QPushButton:hover, QPushButton:pressed {{
                background-color: {th['btn_hover']};
                color: {th['btn_text']};
            }}

            QCheckBox::indicator, QRadioButton::indicator {{
                width: 14px; height: 14px;
                border-radius: 7px;
                border: 1.2px solid {th['box_border']};
                background: {th['input_bg']};
                margin-right: 5px;
            }}
            QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
                background: {th['checkbox']};
                border: 1.2px solid {th['checkbox']};
            }}

            QTableWidget::item:selected {{
                background: {th['cond_selected']};
                color: {th['text']};
            }}

            QProgressBar {{
                background: {th['input_bg']};
                border-radius: 7px;
                text-align: center;
                color: {th['text']};
                font-weight: bold;
                height: 18px;
            }}
            QProgressBar::chunk {{
                background: {th['slider']};
                border-radius: 7px;
            }}

            QSlider::groove:horizontal {{
                border-radius: 4px;
                height: 8px;
                background: {th['input_bg']};
            }}
            QSlider::sub-page:horizontal {{
                background: {th['slider']};
                border-radius: 4px;
                height: 8px;
            }}
            QSlider::add-page:horizontal {{
                background: {th['input_bg']};
                border-radius: 4px;
                height: 8px;
            }}
            QSlider::handle:horizontal {{
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 {th['slider']}, stop:1 {slider_dark});
                border-radius: 8px;
                width: 16px;
                margin: -4px 0;
            }}
            """
            QApplication.instance().setStyleSheet(qss)

        # --- Animação usando QTimeLine ---
        timeline = QTimeLine(500, self)
        timeline.setFrameRange(0, 100)
        timeline.frameChanged.connect(lambda f: aplicar_qss(f / 100.0))
        timeline.start()

        # Atualiza o estado final
        self.current_theme = name
        self.settings["theme"] = name
        salvar_settings(self.settings)

    def update_theme_combo(self):
        self.combo_theme.blockSignals(True)
        self.combo_theme.clear()
        self.combo_theme.addItems(self.themes.keys())
        self.combo_theme.setCurrentText(self.current_theme)
        self.combo_theme.blockSignals(False)

    # ─── Multíplas Origens ─────────────────────────────────────────────────
    
    def init_paths_group(self):
        grp = QGroupBox("Fontes de Origem (Pasta ou ZIP)")
        lo = QVBoxLayout(grp)
        lo.setSpacing(6)
        lo.setContentsMargins(10, 8, 10, 8)
        self.paths_lo = lo

        self.add_origin_btn = QPushButton("➕ Adicionar Origem")
        self.add_origin_btn.setMinimumHeight(24)
        self.add_origin_btn.setStyleSheet("font-size: 12px; padding: 4px 18px; font-weight: 500;")
        self.add_origin_btn.clicked.connect(self.add_origin_row)
        self.add_origin_btn.adjustSize()
        lo.addWidget(self.add_origin_btn)

        self.origin_rows = []
        self.add_origin_row()

        h1 = QHBoxLayout(); h1.setSpacing(8)
        h1.addWidget(QLabel("Extrair ZIPs:"))
        self.chk_extract = QCheckBox()
        h1.addWidget(self.chk_extract)
        h1.addStretch()
        lo.addLayout(h1)

        h2 = QHBoxLayout(); h2.setSpacing(8)
        h2.addWidget(QLabel("Destino:"))
        self.input_dest = DropLineEdit()
        h2.addWidget(self.input_dest)
        b2 = QPushButton("🔍")
        b2.setMinimumHeight(24)
        b2.setStyleSheet("font-size: 15px; padding: 4px 10px; font-weight: 700;")
        b2.clicked.connect(lambda: self._select_folder(self.input_dest))
        b2.adjustSize()
        h2.addWidget(b2)
        h2.addStretch()
        lo.addLayout(h2)
        self.layout.addWidget(grp)

    def add_origin_row(self):
        hl = QHBoxLayout(); hl.setSpacing(6)
        le = DropLineEdit()
        b1 = QPushButton("🔍")
        b1.setMinimumHeight(24)
        b1.setStyleSheet("font-size: 15px; padding: 4px 10px; font-weight: 700;")
        b1.clicked.connect(lambda _, w=le: self._select_origin(w))
        b1.adjustSize()

        b2 = QPushButton("🗑")
        b2.setMinimumHeight(24)
        b2.setStyleSheet("font-size: 15px; padding: 4px 10px; font-weight: 700;")
        b2.clicked.connect(lambda _, r=hl: self._remove_origin_row(r))
        b2.adjustSize()

        le.setMinimumWidth(120)
        le.setMaximumWidth(270)
        hl.addWidget(le)
        hl.addWidget(b1)
        hl.addWidget(b2)
        hl.addStretch()
        idx = self.paths_lo.indexOf(self.add_origin_btn)
        self.paths_lo.insertLayout(idx, hl)
        self.origin_rows.append((hl, le))

    def _remove_origin_row(self, hl):
        if len(self.origin_rows) <= 1: return
        for r, le in self.origin_rows:
            if r is hl:
                while r.count():
                    it = r.takeAt(0)
                    w = it.widget()
                    if w: w.deleteLater()
                self.origin_rows.remove((r, le))
                break

    def _select_origin(self, w):
        dlg = OrigemDialog(self)
        if dlg.exec() == QDialog.Accepted:
            if dlg.selecionado == "pasta":
                path = QFileDialog.getExistingDirectory(self, "Selecione uma pasta de origem")
                if path:
                    w.setText(path)
            elif dlg.selecionado == "zip":
                path, _ = QFileDialog.getOpenFileName(self, "Selecione um arquivo ZIP", "", "Arquivo ZIP (*.zip);;Todos Arquivos (*)")
                if path:
                    w.setText(path)

    def _select_folder(self, w):
        path = QFileDialog.getExistingDirectory(self, "Selecione uma pasta")
        if path:
            w.setText(path)

    # ─── Opções de Subpasta e Ação ─────────────────────────────────────────
    
    def init_options_group(self):
        grp = QGroupBox("Opções de Subpasta e Ação")
        lo = QVBoxLayout(grp); lo.setSpacing(8); lo.setContentsMargins(10, 8, 10, 8)
        self.chk_sub       = QCheckBox("Criar Subpasta")
        self.chk_find_sub = QCheckBox("Procurar Subpasta")
        self.chk_hierarchy = QCheckBox("Hierarquia")
        self.chk_multiply  = QCheckBox("Multiplicar")
        self.chk_copydirs  = QCheckBox("Copiar Pastas/Subpastas")
        add_flag_with_info(lo, self.chk_sub,       FLAG_INFOS["chk_sub"])
        add_flag_with_info(lo, self.chk_find_sub, FLAG_INFOS["chk_findsub"])
        add_flag_with_info(lo, self.chk_hierarchy, FLAG_INFOS["chk_hierarchy"])
        add_flag_with_info(lo, self.chk_multiply,  FLAG_INFOS["chk_multiply"])
        add_flag_with_info(lo, self.chk_copydirs,  FLAG_INFOS["chk_copydirs"])
        h_sobra = QHBoxLayout(); h_sobra.setSpacing(8)
        self.chk_sobra = QCheckBox("Pasta Sobra")
        self.le_sobra  = QLineEdit(); self.le_sobra.setEnabled(False)
        self.le_sobra.setMinimumWidth(120); self.le_sobra.setMaximumWidth(270)
        self.chk_sobra.toggled.connect(self.le_sobra.setEnabled)
        h_sobra.addWidget(self.chk_sobra)
        btn_info_sobra = QPushButton("(!)")
        btn_info_sobra.setObjectName("infoButton")
        btn_info_sobra.setFixedSize(24, 24)
        btn_info_sobra.setToolTip("Clique para ver detalhes sobre esta opção")
        btn_info_sobra.clicked.connect(lambda: QMessageBox.information(self.chk_sobra, "Informação", FLAG_INFOS["chk_sobra"]))
        h_sobra.addWidget(btn_info_sobra)
        h_sobra.addWidget(self.le_sobra); h_sobra.addStretch()
        lo.addLayout(h_sobra)
        self.chk_recursive = QCheckBox("Busca Recursiva")
        add_flag_with_info(lo, self.chk_recursive, FLAG_INFOS["chk_recursive"])
        h_act = QHBoxLayout(); h_act.setSpacing(8)
        h_act.addWidget(QLabel("Ação:"))
        self.rb_move   = QRadioButton("Mover")
        self.rb_copy   = QRadioButton("Copiar"); self.rb_copy.setChecked(True)
        self.rb_delete = QRadioButton("Excluir")
        bg = QButtonGroup(); bg.addButton(self.rb_move); bg.addButton(self.rb_copy); bg.addButton(self.rb_delete)
        h_act.addWidget(self.rb_move); h_act.addWidget(self.rb_copy); h_act.addWidget(self.rb_delete); h_act.addStretch()
        lo.addLayout(h_act)
        self.layout.addWidget(grp)

    # ─── Threads ──────────────────────────────────────────────────────────
    
    def init_threads_group(self):
        grp = QGroupBox("Configurações de Threads")
        lo = QHBoxLayout(grp); lo.setSpacing(8); lo.setContentsMargins(10, 8, 10, 8)
        lo.addWidget(QLabel("Número de Threads:"))
        self.slider_threads = QSlider(Qt.Horizontal)
        self.slider_threads.setMinimum(1)
        self.slider_threads.setMaximum(self.max_workers)
        self.slider_threads.setValue(min(4, self.max_workers))
        self.slider_threads.setTickPosition(QSlider.TicksBelow)
        self.slider_threads.setTickInterval(1)
        self.slider_threads.setMinimumWidth(200)
        self.slider_threads.setMaximumWidth(400)
        self.slider_threads.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.label_threads = QLabel(f"{self.slider_threads.value()} / {self.max_workers}")
        self.slider_threads.valueChanged.connect(lambda v: self.label_threads.setText(f"{v} / {self.max_workers}"))
        lo.addWidget(self.slider_threads)
        lo.addWidget(self.label_threads)
        lo.addStretch()
        self.layout.addWidget(grp)
    
    # ─── Condições ─────────────────────────────────────────────────────────
    
    def init_conditions_group(self):
        grp = QGroupBox("Condições")
        form = QFormLayout(grp)
        form.setSpacing(5)
        form.setContentsMargins(10, 8, 10, 8)
        self.chk_none = QCheckBox("Sem Condições")
        self.chk_none.toggled.connect(self._toggle_conditions)
        form.addRow(self.chk_none)

        # Radios
        self.rb_excel   = QRadioButton("De Excel")
        self.rb_folders = QRadioButton("De Subpastas")
        self.rb_excel.setChecked(True)
        mg = QButtonGroup(); mg.addButton(self.rb_excel); mg.addButton(self.rb_folders)
        radios_h = QHBoxLayout()
        radios_h.addWidget(self.rb_excel); radios_h.addWidget(self.rb_folders); radios_h.addStretch()
        form.addRow("Origem:", radios_h)

        # Input dinâmico
        self.stacked_input = QStackedWidget()
        excel_w = QWidget(); hl1 = QHBoxLayout(excel_w); hl1.setContentsMargins(0,0,0,0); hl1.setSpacing(5)
        self.le_excel = DropLineEdit()
        self.le_excel.textChanged.connect(self.load_conditions_from_excel)
        def on_excel_browse():
            self._select_file(self.le_excel, "Excel (*.xlsx *.xls)")
            self.load_conditions_from_excel()
        btn_exc = QPushButton("🔍")
        btn_exc.setFixedHeight(18)
        btn_exc.setFixedWidth(24)
        btn_exc.clicked.connect(on_excel_browse)
        hl1.addWidget(self.le_excel)
        hl1.addWidget(btn_exc)
        self.stacked_input.addWidget(excel_w)
        folder_w = QWidget(); hl2 = QHBoxLayout(folder_w); hl2.setContentsMargins(0,0,0,0); hl2.setSpacing(5)
        self.le_folder = DropLineEdit()
        btn_fol = QPushButton("🔍"); btn_fol.setFixedHeight(18); btn_fol.setFixedWidth(24)
        btn_fol.clicked.connect(lambda: self._select_folder(self.le_folder))
        hl2.addWidget(self.le_folder); hl2.addWidget(btn_fol)
        self.stacked_input.addWidget(folder_w)
        form.addRow("Arquivo/Pasta:", self.stacked_input)
        self.rb_excel.toggled.connect(lambda x: self.stacked_input.setCurrentIndex(0 if x else 1))
        self.stacked_input.setCurrentIndex(0)

        self.le_sep = QLineEdit("_")
        self.le_sep.setMaximumWidth(46)
        self.le_sep.setEnabled(False)
        def update_sep_enabled():
            enabled = self.rb_folders.isChecked() or self.chk_sub.isChecked()
            self.le_sep.setEnabled(enabled)
        self.rb_folders.toggled.connect(lambda _: update_sep_enabled())
        self.chk_sub.toggled.connect(lambda _: update_sep_enabled())
        update_sep_enabled()  # chamada inicial
        form.addRow("Separador:", self.le_sep)

        # --- Tabela de condições (já preparada para expansão de campos extras) ---
        self.table = ConditionTable(0, 4)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setColumnWidth(0, 44)
        self.table.setColumnWidth(2, 52)
        self.table.setColumnWidth(3, 58)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(26)
        self.table_container = QFrame()
        self.table_container.setObjectName("tableCondContainer")
        container_layout = QVBoxLayout(self.table_container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)
        container_layout.addWidget(self.table)
        max_visible = 5
        def adjust_table_height():
            rows = self.table.rowCount()
            visible = min(max(rows, 1), max_visible)
            row_height = self.table.verticalHeader().defaultSectionSize()
            header_height = self.table.horizontalHeader().height()
            padding = 12
            total = (visible * row_height) + header_height + padding
            self.table.setMinimumHeight(total)
            self.table.setMaximumHeight(total)
            self.table_container.setMinimumHeight(total)
            self.table_container.setMaximumHeight(total)
        self.table.model().rowsInserted.connect(lambda *_: adjust_table_height())
        self.table.model().rowsRemoved.connect(lambda *_: adjust_table_height())
        adjust_table_height()
        form.addRow("Defina Condições:", self.table_container)

        # Botões adicionar/remover
        h_btns = QHBoxLayout(); h_btns.setSpacing(5)
        b1 = QPushButton("Adicionar")
        b1.setMinimumHeight(22)
        b1.setStyleSheet("font-size: 12px; font-weight: 500;")
        b1.clicked.connect(self.add_condition_row)
        b1.adjustSize()
        b2 = QPushButton("Remover")
        b2.setMinimumHeight(22)
        b2.setStyleSheet("font-size: 12px; font-weight: 500;")
        b2.clicked.connect(self.remove_condition_row)
        b2.adjustSize()
        h_btns.addWidget(b1); h_btns.addWidget(b2); h_btns.addStretch()
        form.addRow("", h_btns)

        self.le_expr = QLineEdit(); self.le_expr.setToolTip("Use & para E, | para OU, () para agrupar")
        form.addRow("Expressão:", self.le_expr)

        # Atalhos - atualização em tempo real (edição de célula também)
        self.shortcuts = QHBoxLayout()
        sw = QWidget(); sw.setLayout(self.shortcuts)
        form.addRow("Inserir:", sw)

        def update_shortcuts():
            while self.shortcuts.count():
                w = self.shortcuts.takeAt(0).widget()
                if w: w.deleteLater()
            for i in range(self.table.rowCount()):
                it = self.table.item(i,1)
                if it:
                    txt = f"!{it.text()}!"
                    btn = QPushButton(txt)
                    btn.clicked.connect(partial(self._insert_text, txt))
                    btn.setMinimumHeight(22)
                    btn.setStyleSheet("font-size:12px; padding:4px 14px; font-weight:500;")
                    btn.adjustSize()
                    self.shortcuts.addWidget(btn)
            shortcut_buttons = [
                ("(", "("),
                (")", ")"),
                ("E", "&"),
                ("EXCETO", "{}"),
                ("OU", "|"), ("FIXO", '""')
            ]
            for label, val in shortcut_buttons:
                btn = QPushButton(label)
                btn.clicked.connect(partial(self._insert_text, val))
                btn.setMinimumHeight(22)
                btn.setStyleSheet("font-size:12px; padding:4px 18px; font-weight:500;")
                btn.adjustSize()
                self.shortcuts.addWidget(btn)
    
        self.update_shortcuts = update_shortcuts

        # --- Botão de Filtros Avançados ---
        self.btn_advfilters = QPushButton("Filtros Avançados ▼")
        self.btn_advfilters.setCheckable(True)
        self.btn_advfilters.setChecked(False)
        form.addRow("", self.btn_advfilters)

        # --- Painel de Filtros Avançados (inicialmente oculto) ---
        self.advfilters_panel = QWidget()
        adv_lo = QVBoxLayout(self.advfilters_panel)
        adv_lo.setContentsMargins(0,0,0,0)
        adv_lo.setSpacing(4)

        # Linha 1: Filtro de extensão
        ext_layout = QHBoxLayout()
        self.chk_ext = QCheckBox("Filtrar por extensão")
        self.in_ext = QLineEdit()
        self.in_ext.setPlaceholderText("ex: pdf, xlsx, jpg (separados por vírgula)")
        self.in_ext.setMinimumWidth(150)
        self.in_ext.setToolTip("Exemplo: pdf, xlsx, jpg (separe por vírgulas)")
        ext_layout.addWidget(self.chk_ext)
        ext_layout.addWidget(self.in_ext)
        ext_layout.addStretch()
        adv_lo.addLayout(ext_layout)

        # Linha 2: Filtro de data inicial
        dt1_layout = QHBoxLayout()
        self.chk_dt1 = QCheckBox("Data Inicial (modificação)")
        self.in_dt1 = QLineEdit()
        self.in_dt1.setPlaceholderText("dd-mm-aaaa ou yyyy-mm-dd")
        self.in_dt1.setMinimumWidth(120)
        self.in_dt1.setToolTip("Data inicial para considerar arquivos (formato dd-mm-aaaa ou yyyy-mm-dd)")
        dt1_layout.addWidget(self.chk_dt1)
        dt1_layout.addWidget(self.in_dt1)
        dt1_layout.addStretch()
        adv_lo.addLayout(dt1_layout)

        # Linha 3: Filtro de data final
        dt2_layout = QHBoxLayout()
        self.chk_dt2 = QCheckBox("Data Final (modificação)")
        self.in_dt2 = QLineEdit()
        self.in_dt2.setPlaceholderText("dd-mm-aaaa ou yyyy-mm-dd")
        self.in_dt2.setMinimumWidth(120)
        self.in_dt2.setToolTip("Data final para considerar arquivos (formato dd-mm-aaaa ou yyyy-mm-dd)")
        dt2_layout.addWidget(self.chk_dt2)
        dt2_layout.addWidget(self.in_dt2)
        dt2_layout.addStretch()
        adv_lo.addLayout(dt2_layout)

        # ---- UX MELHORIA: Campos só editáveis quando o checkbox correspondente estiver marcado ----
        self.in_ext.setEnabled(False)
        self.in_dt1.setEnabled(False)
        self.in_dt2.setEnabled(False)
        self.chk_ext.toggled.connect(self.in_ext.setEnabled)
        self.chk_dt1.toggled.connect(self.in_dt1.setEnabled)
        self.chk_dt2.toggled.connect(self.in_dt2.setEnabled)

        # ---- UX MELHORIA: Limpa o campo ao desmarcar ----
        def clear_if_unchecked(chk, field):
            if not chk.isChecked():
                field.clear()
                field.setStyleSheet("")
        self.chk_ext.toggled.connect(lambda: clear_if_unchecked(self.chk_ext, self.in_ext))
        self.chk_dt1.toggled.connect(lambda: clear_if_unchecked(self.chk_dt1, self.in_dt1))
        self.chk_dt2.toggled.connect(lambda: clear_if_unchecked(self.chk_dt2, self.in_dt2))

        # ---- MELHORIA: Validação visual de datas ----
        def validate_date(lineedit):
            val = lineedit.text().strip()
            if not val:
                lineedit.setStyleSheet("")
                return
            import re
            pat1 = r"^\d{2}-\d{2}-\d{4}$"
            pat2 = r"^\d{4}-\d{2}-\d{2}$"
            if re.match(pat1, val) or re.match(pat2, val):
                lineedit.setStyleSheet("")  # Normal caso válido
            else:
                # Vermelho escuro + texto branco para erro
                lineedit.setStyleSheet("background-color: #aa2020; color: #fff;")

        self.in_dt1.textChanged.connect(lambda: validate_date(self.in_dt1))
        self.in_dt2.textChanged.connect(lambda: validate_date(self.in_dt2))

        self.advfilters_panel.setVisible(False)
        form.addRow("", self.advfilters_panel)

        def toggle_advanced():
            st = self.btn_advfilters.isChecked()
            self.advfilters_panel.setVisible(st)
            self.btn_advfilters.setText("Filtros Avançados ▲" if st else "Filtros Avançados ▼")
        self.btn_advfilters.toggled.connect(toggle_advanced)

        # === Seção de Renomeação ===
        self.chk_rename = QCheckBox("Habilitar Renomeação")
        self.le_rename = QLineEdit()
        self.le_rename.setPlaceholderText('Ex: "cliente"!Cond1!"data"')
        self.le_rename.setEnabled(False)
        self.chk_rename.toggled.connect(self.le_rename.setEnabled)
        form.addRow("Renomear Arquivos:", self.chk_rename)
        form.addRow("Padrão de Nome:",  self.le_rename)

        self.table.selectionModel().currentRowChanged.connect(self.table.refresh_move_column)
        self.table.itemChanged.connect(lambda _: self.update_shortcuts())  # Atualiza atalhos ao editar qualquer célula
        self.update_shortcuts()
        self.layout.addWidget(grp)

    def _toggle_conditions(self, no_cond):
        for w in (
            self.rb_excel, self.rb_folders,
            self.stacked_input, self.le_sep,
            self.le_expr, self.table
        ):
            w.setEnabled(not no_cond)

    def _select_file(self, w, filt):
        p,_ = QFileDialog.getOpenFileName(self, "Selecione Arquivo", filter=filt)
        if p: w.setText(p)

    def add_condition_row(self):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.refresh_move_column()
        self.table.setItem(r,1, QTableWidgetItem(f"Cond{r+1}"))
        self.table.setItem(r,2, QTableWidgetItem(str(r+1)))
        chk = QTableWidgetItem(); chk.setFlags(chk.flags()|Qt.ItemIsUserCheckable); chk.setCheckState(Qt.Unchecked)
        self.table.setItem(r,3, chk)
        self.table.selectRow(r)
        self.update_shortcuts()
        self.table.viewport().update()

    def remove_condition_row(self):
        r = self.table.currentRow()
        if r >= 0:
            self.table.removeRow(r)
            self.table.refresh_move_column()
            if self.table.rowCount():
                idx = min(r, self.table.rowCount() - 1)
                self.table.selectRow(idx)
            self.update_shortcuts()
            self.table.viewport().update()

    def move_condition(self, src, dst):
        cnt = self.table.rowCount()
        if dst < 0 or dst >= cnt:
            return
        for c in range(self.table.columnCount()):
            item_src = self.table.takeItem(src, c)
            item_dst = self.table.takeItem(dst, c)
            self.table.setItem(dst, c, item_src)
            self.table.setItem(src, c, item_dst)
        self.table.selectRow(dst)
        self.table.refresh_move_column()
        self.table.viewport().update()

    def load_conditions_from_excel(self):
        path = self.le_excel.text().strip()
        if not path or not os.path.isfile(path):
            return
        try:
            import pandas as pd
            df = pd.read_excel(path, nrows=1)
            headers = list(df.columns)
            self.table.setRowCount(0)
            for i, col in enumerate(headers):
                self.table.insertRow(i)
                # Coluna 1: Nome da condição (igual ao cabeçalho)
                self.table.setItem(i, 1, QTableWidgetItem(str(col)))
                # Coluna 2: Índice (apenas para compatibilidade, mas pode ser o número da coluna +1)
                self.table.setItem(i, 2, QTableWidgetItem(str(i + 1)))
                # Coluna 3: Checkbox "Principal"
                chk = QTableWidgetItem()
                chk.setFlags(chk.flags() | Qt.ItemIsUserCheckable)
                chk.setCheckState(Qt.Unchecked)
                self.table.setItem(i, 3, chk)
            self.update_shortcuts()
        except Exception as e:
            QMessageBox.warning(self, "Erro", f"Falha ao ler cabeçalhos do Excel: {e}")
    
    # ─── Execução ──────────────────────────────────────────────────────────
    
    def init_execution_group(self):
        grp = QGroupBox("Execução")
        v   = QVBoxLayout(grp)
        h = QHBoxLayout()
        self.btn_execute = QPushButton("▶ Executar")
        self.btn_execute.setMinimumHeight(24)
        self.btn_execute.clicked.connect(self.start_execution)
        self.btn_execute.adjustSize()
        self.btn_cancel = QPushButton("❌ Cancelar")
        self.btn_cancel.setMinimumHeight(24)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.cancel_execution)
        self.btn_cancel.adjustSize()
        self.btn_report = QPushButton("📊 Relatório")
        self.btn_report.setMinimumHeight(24)
        self.btn_report.setEnabled(False)
        self.btn_report.clicked.connect(self.show_report)
        self.btn_report.adjustSize()
        h.addWidget(self.btn_execute)
        h.addWidget(self.btn_cancel)
        h.addWidget(self.btn_report)
        h.addStretch()
        v.addLayout(h)
        self.chk_zip   = QCheckBox("Compactar destino ao concluir")
        add_flag_with_info(v, self.chk_zip, FLAG_INFOS["chk_zip"])
        self.progress  = QProgressBar()
        self.progress.setVisible(False)
        v.addWidget(self.progress)
        self.layout.addWidget(grp)
        self._last_report = []

    def _insert_text(self, txt):
        w = self._last_edit if self._last_edit in (self.le_expr, self.le_rename) else self.le_expr
        w.setFocus()
        pos = w.cursorPosition()
        cur = w.text()
        w.setText(cur[:pos] + txt + cur[pos:])
        w.setCursorPosition(pos + len(txt))

    def _on_focus_changed(self, old: QWidget, new: QWidget):
        # guarda só QLineEdits
        if isinstance(new, QLineEdit):
            self._last_edit = new

    def collect_config(self) -> dict:
        # Define a ação principal
        if self.rb_move.isChecked():
            action = "move"
        elif self.rb_copy.isChecked():
            action = "copy"
        elif self.rb_delete.isChecked():
            action = "delete"
        else:
            action = "copy"

        col_map, princ = {}, []

        # Se modo Excel: lê o cabeçalho real da planilha e faz o mapeamento nome → nome
        if self.rb_excel.isChecked() and self.le_excel.text():
            try:
                import pandas as pd
                df = pd.read_excel(self.le_excel.text(), dtype=str, nrows=1)
                headers = list(df.columns)
                for r in range(self.table.rowCount()):
                    it_n = self.table.item(r, 1)
                    it_p = self.table.item(r, 3)
                    if it_n:
                        nm = it_n.text()
                        if nm in headers:
                            col_map[nm] = nm  # mapeia o nome para ele mesmo
                            if it_p and it_p.checkState() == Qt.Checked:
                                princ.append(nm)
            except Exception as e:
                print(f"[collect_config] Erro ao ler cabeçalho Excel: {e}")
        else:
            # Modo subpasta (índice)
            for r in range(self.table.rowCount()):
                it_n = self.table.item(r, 1)
                it_i = self.table.item(r, 2)
                it_p = self.table.item(r, 3)
                if it_n and it_i:
                    nm, idx = it_n.text(), int(it_i.text())
                    col_map[nm] = idx
                    if it_p and it_p.checkState() == Qt.Checked:
                        princ.append(nm)

        # Filtros de arquivo
        file_filters = {}
        if self.chk_ext.isChecked() and self.in_ext.text().strip():
            tipos = [x.strip().lower() for x in self.in_ext.text().split(",") if x.strip()]
            if tipos:
                file_filters["types"] = tipos
        if self.chk_dt1.isChecked() and self.in_dt1.text().strip():
            file_filters["date_start"] = self.in_dt1.text().strip()
        if self.chk_dt2.isChecked() and self.in_dt2.text().strip():
            file_filters["date_end"] = self.in_dt2.text().strip()

        return {
            "origens":            [le.text() for _, le in self.origin_rows if le.text()],
            "find_subpasta":      self.chk_find_sub.isChecked(),
            "destino":            self.input_dest.text(),
            "action":             action,
            "extract_zips":       self.chk_extract.isChecked(),
            "recursivo":          self.chk_recursive.isChecked(),
            "criar_subpasta":     self.chk_sub.isChecked(),
            "hierarchy":          self.chk_hierarchy.isChecked(),
            "multiply":           self.chk_multiply.isChecked(),
            "sobra_enabled":      self.chk_sobra.isChecked(),
            "sobra":              self.le_sobra.text() if self.chk_sobra.isChecked() else None,
            "zip_dest":           self.chk_zip.isChecked(),
            "max_workers":        self.slider_threads.value(),
            "use_conditions":     not self.chk_none.isChecked(),
            "condition_mode":     "folders" if self.rb_folders.isChecked() else "excel",
            "excel":              self.le_excel.text(),
            "cond_folder":        self.le_folder.text(),
            "cond_sep":           self.le_sep.text(),
            "colunas":            col_map,
            "principais":         princ,
            "condition_expression": self.le_expr.text(),
            "copy_dirs":          self.chk_copydirs.isChecked(),
            "file_filters":       file_filters,
            # ==== RENOMEAÇÃO ====
            "rename": {
                "enabled": self.chk_rename.isChecked(),
                "pattern": self.le_rename.text().strip()
            }
        }

    def start_execution(self):
        cfg = self.collect_config()
        if not self.validate_config(cfg):
            return
        self.btn_execute.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.btn_report.setEnabled(False)
        self.progress.setValue(0)
        self.progress.setVisible(True)
        self.progress.setFormat("%v de %m arquivos processados")
        self._set_all_enabled(False)
        self.thread = ExecutorThread(cfg)
        self.thread.progress.connect(self._on_progress)
        self.thread.error.connect(lambda msg: QMessageBox.warning(self, "Erro", msg))
        self.thread.finished.connect(self.execution_finished)
        self.thread.canceled.connect(self.execution_canceled)
        self.thread.start()
    
    def _on_progress(self, value, total):
        self.progress.setMaximum(total)
        self.progress.setValue(value)
        self.progress.setFormat(f"{value} de {total} arquivos processados")
    
    def cancel_execution(self):
        if self.thread:
            self.thread.cancel()
            self.btn_cancel.setEnabled(False)
    
    def execution_finished(self, report):
        self.btn_execute.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.btn_report.setEnabled(True)
        self.progress.setVisible(False)
        self._set_all_enabled(True)
        self._last_report = report
        QMessageBox.information(self, "Concluído", "Execução finalizada.")
        self.show_report()
    
    def execution_canceled(self):
        self.btn_execute.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.btn_report.setEnabled(bool(self._last_report))
        self.progress.setVisible(False)
        self._set_all_enabled(True)
        QMessageBox.information(self, "Cancelado", "Execução foi cancelada pelo usuário.")

    def _set_all_enabled(self, enabled):
        widgets = [
            self.combo_theme, self.add_origin_btn, self.input_dest, self.chk_extract,
            self.chk_sub, self.chk_hierarchy, self.chk_multiply, self.chk_sobra,
            self.le_sobra, self.chk_recursive, self.rb_move, self.rb_copy, self.rb_delete,
            self.slider_threads, self.chk_none, self.rb_excel, self.rb_folders,
            self.le_excel, self.le_folder, self.le_sep,
            self.table, self.le_expr, self.chk_zip
        ]
        for w in widgets:
            try: w.setEnabled(enabled)
            except: pass
        self.btn_cancel.setEnabled(not enabled and self.thread is not None)
    
    def show_report(self):
        if self._last_report:
            dlg = ReportDialog(self._last_report, self)
            dlg.exec()
        else:
            QMessageBox.information(self, "Relatório", "Nenhum relatório disponível.")

if __name__ == "__main__":
    show_splash_and_run(MainWindow)
