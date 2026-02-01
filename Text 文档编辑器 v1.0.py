import sys
import os
import json
import platform
import base64
import re
import imghdr
from urllib.parse import unquote, urlparse
from functools import partial

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QAction, QFileDialog, QToolBar,
    QFontComboBox, QComboBox, QSpinBox, QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QLabel, QListWidget, QLineEdit, QMessageBox, QListWidgetItem,
    QInputDialog, QColorDialog
)
from PyQt5.QtGui import (
    QIcon, QFont, QTextCharFormat, QTextCursor, QTextBlockFormat,
    QTextImageFormat, QImage, QTextTableFormat, QColor, QBrush
)
from PyQt5.QtCore import Qt, QSize

# ---------------------------
# Combo popup & menu QSS (hover/selected styles)
# ---------------------------
COMBO_POPUP_QSS = """
QListView {
  background-color: #FFFFFF;
  outline: 0;
}
QListView::item {
  height: 28px;
  padding: 4px 10px;
  color: #213447;
}
QListView::item:selected, QListView::item:hover {
  background-color: #2E8BFF;
  color: #FFFFFF;
}
"""

GLOBAL_QSS_MENU = """
QMenu {
  background-color: #FFFFFF;
  border: 1px solid #cfe8fb;
}
QMenu::item {
  padding: 6px 24px;
  color: #213447;
}
QMenu::item:selected {
  background-color: #2E8BFF;
  color: #FFFFFF;
}
QComboBox QAbstractItemView {
  background-color: #FFFFFF;
  selection-background-color: #2E8BFF;
  selection-color: #FFFFFF;
}
"""

# ---------------------------
# Translations (expanded)
# ---------------------------
TRANSLATIONS = {
    "en-US": {
        "app_title": "Text Editor",
        "menu_file": "File",
        "menu_edit": "Edit",
        "menu_insert": "Insert",
        "new": "New",
        "open": "Open",
        "save": "Save",
        "save_as": "Save As",
        "exit": "Exit",
        "bold": "Bold",
        "italic": "Italic",
        "underline": "Underline",
        "align_left": "Align Left",
        "align_center": "Align Center",
        "align_right": "Align Right",
        "search": "Search .voc",
        "desktop_search": "Desktop",
        "filename_prompt": "Filename (without extension):",
        "saved": "Saved.",
        "open_error": "Failed to open file.",
        "not_windows": "Not running on Windows. Desktop save will instead use your home directory.",
        "no_file_selected": "No file selected.",
        "confirm_clear": "Clear editor? Unsaved changes will be lost.",
        "insert_image": "Insert Image",
        "insert_table": "Insert Table",
        "style": "Style",
        "undo": "Undo",
        "redo": "Redo",
        "style_label": "Style:",
        "style_normal": "Normal",
        "style_title": "Title",
        "style_subtitle": "Subtitle",
        "style_subtitle2": "Subtitle 2",
        "style_h1": "Heading 1",
        "style_h2": "Heading 2",
        "style_quote": "Quote",
        "style_code": "Code",
        "style_mono": "Monospace",
        "style_small": "Small",
        "style_large": "Large",
        "style_emphasis": "Emphasis",
        "rows_label": "Rows:",
        "cols_label": "Columns:",
        "insert_image_error": "Cannot load image:",
        "table_title": "Insert Table",
        "text_color": "Text Color",
        "bg_color": "Highlight Color"
    },
    "en-GB": {
        "app_title": "Text Editor",
        "menu_file": "File",
        "menu_edit": "Edit",
        "menu_insert": "Insert",
        "new": "New",
        "open": "Open",
        "save": "Save",
        "save_as": "Save As",
        "exit": "Exit",
        "bold": "Bold",
        "italic": "Italic",
        "underline": "Underline",
        "align_left": "Align Left",
        "align_center": "Align Centre",
        "align_right": "Align Right",
        "search": "Search .voc",
        "desktop_search": "Desktop",
        "filename_prompt": "Filename (without extension):",
        "saved": "Saved.",
        "open_error": "Failed to open file.",
        "not_windows": "Not running on Windows. Desktop save will instead use your home directory.",
        "no_file_selected": "No file selected.",
        "confirm_clear": "Clear editor? Unsaved changes will be lost.",
        "insert_image": "Insert Image",
        "insert_table": "Insert Table",
        "style": "Style",
        "undo": "Undo",
        "redo": "Redo",
        "style_label": "Style:",
        "style_normal": "Normal",
        "style_title": "Title",
        "style_subtitle": "Subtitle",
        "style_subtitle2": "Subtitle 2",
        "style_h1": "Heading 1",
        "style_h2": "Heading 2",
        "style_quote": "Quote",
        "style_code": "Code",
        "style_mono": "Monospace",
        "style_small": "Small",
        "style_large": "Large",
        "style_emphasis": "Emphasis",
        "rows_label": "Rows:",
        "cols_label": "Columns:",
        "insert_image_error": "Cannot load image:",
        "table_title": "Insert Table",
        "text_color": "Text Color",
        "bg_color": "Highlight Color"
    },
    "zh-CN": {
        "app_title": "Text 文档编辑器",
        "menu_file": "文件",
        "menu_edit": "编辑",
        "menu_insert": "插入",
        "new": "新建",
        "open": "打开",
        "save": "保存",
        "save_as": "另存为",
        "exit": "退出",
        "bold": "加粗",
        "italic": "斜体",
        "underline": "下划线",
        "align_left": "左对齐",
        "align_center": "居中",
        "align_right": "右对齐",
        "search": "查找 .voc",
        "desktop_search": "桌面",
        "filename_prompt": "文件名（不含扩展名）：",
        "saved": "已保存。",
        "open_error": "打开文件失败。",
        "not_windows": "非 Windows 系统，保存目录将使用用户主目录。",
        "no_file_selected": "未选择文件。",
        "confirm_clear": "清空编辑区？未保存的更改将丢失。",
        "insert_image": "插入图片",
        "insert_table": "插入表格",
        "style": "样式",
        "undo": "撤销",
        "redo": "重做",
        "style_label": "样式：",
        "style_normal": "正文",
        "style_title": "标题",
        "style_subtitle": "副标题",
        "style_subtitle2": "副标题2",
        "style_h1": "标题 1",
        "style_h2": "标题 2",
        "style_quote": "引用",
        "style_code": "代码",
        "style_mono": "等宽",
        "style_small": "小号",
        "style_large": "大号",
        "style_emphasis": "强调",
        "rows_label": "行数：",
        "cols_label": "列数：",
        "insert_image_error": "无法加载图片：",
        "table_title": "插入表格",
        "text_color": "文字颜色",
        "bg_color": "背景颜色"
    },
    "zh-TW": {
        "app_title": "Text 文件編輯器",
        "menu_file": "文件",
        "menu_edit": "編輯",
        "menu_insert": "插入",
        "new": "新建",
        "open": "打開",
        "save": "保存",
        "save_as": "另存為",
        "exit": "退出",
        "bold": "加粗",
        "italic": "斜體",
        "underline": "底線",
        "align_left": "靠左",
        "align_center": "置中",
        "align_right": "靠右",
        "search": "查找 .voc",
        "desktop_search": "桌面",
        "filename_prompt": "檔名（不含副檔名）：",
        "saved": "已保存。",
        "open_error": "打開檔案失敗。",
        "not_windows": "非 Windows 系統，保存目錄將使用使用者主目錄。",
        "no_file_selected": "未選擇檔案。",
        "confirm_clear": "清空編輯區？未保存的更改將遺失。",
        "insert_image": "插入圖片",
        "insert_table": "插入表格",
        "style": "樣式",
        "undo": "復原",
        "redo": "重做",
        "style_label": "樣式：",
        "style_normal": "正文",
        "style_title": "標題",
        "style_subtitle": "副標題",
        "style_subtitle2": "副標題2",
        "style_h1": "標題 1",
        "style_h2": "標題 2",
        "style_quote": "引用",
        "style_code": "程式碼",
        "style_mono": "等寬",
        "style_small": "小號",
        "style_large": "大號",
        "style_emphasis": "強調",
        "rows_label": "行數：",
        "cols_label": "列數：",
        "insert_image_error": "無法載入圖片：",
        "table_title": "插入表格",
        "text_color": "文字顏色",
        "bg_color": "背景顏色"
    },
    "ja-JP": {
        "app_title": "Text エディタ",
        "menu_file": "ファイル",
        "menu_edit": "編集",
        "menu_insert": "挿入",
        "new": "新規",
        "open": "開く",
        "save": "保存",
        "save_as": "名前を付けて保存",
        "exit": "終了",
        "bold": "太字",
        "italic": "斜体",
        "underline": "下線",
        "align_left": "左揃え",
        "align_center": "中央揃え",
        "align_right": "右揃え",
        "search": ".voc を検索",
        "desktop_search": "デスクトップ",
        "filename_prompt": "ファイル名（拡張子なし）：",
        "saved": "保存しました。",
        "open_error": "ファイルを開けませんでした。",
        "not_windows": "Windows ではありません。保存先はホームディレクトリになります。",
        "no_file_selected": "ファイルが選択されていません。",
        "confirm_clear": "エディタをクリアしますか？未保存の変更は失われます。",
        "insert_image": "画像を挿入",
        "insert_table": "表を挿入",
        "style": "スタイル",
        "undo": "元に戻す",
        "redo": "やり直し",
        "style_label": "スタイル：",
        "style_normal": "標準",
        "style_title": "タイトル",
        "style_subtitle": "サブタイトル",
        "style_subtitle2": "サブタイトル2",
        "style_h1": "見出し 1",
        "style_h2": "見出し 2",
        "style_quote": "引用",
        "style_code": "コード",
        "style_mono": "等幅",
        "style_small": "小",
        "style_large": "大",
        "style_emphasis": "強調",
        "rows_label": "行：",
        "cols_label": "列：",
        "insert_image_error": "画像を読み込めません：",
        "table_title": "表を挿入",
        "text_color": "文字色",
        "bg_color": "背景色"
    },
    "es-ES": {
        "app_title": "Editor Text",
        "menu_file": "Archivo",
        "menu_edit": "Editar",
        "menu_insert": "Insertar",
        "new": "Nuevo",
        "open": "Abrir",
        "save": "Guardar",
        "save_as": "Guardar como",
        "exit": "Salir",
        "bold": "Negrita",
        "italic": "Cursiva",
        "underline": "Subrayado",
        "align_left": "Alinear a la izquierda",
        "align_center": "Centrar",
        "align_right": "Alinear a la derecha",
        "search": "Buscar .voc",
        "desktop_search": "Escritorio",
        "filename_prompt": "Nombre de archivo (sin extensión):",
        "saved": "Guardado.",
        "open_error": "Error al abrir el archivo.",
        "not_windows": "No está en Windows. El guardado usará el directorio del usuario.",
        "no_file_selected": "Ningún archivo seleccionado.",
        "confirm_clear": "¿Borrar editor? Los cambios no guardados se perderán.",
        "insert_image": "Insertar imagen",
        "insert_table": "Insertar tabla",
        "style": "Estilo",
        "undo": "Deshacer",
        "redo": "Rehacer",
        "style_label": "Estilo:",
        "style_normal": "Normal",
        "style_title": "Título",
        "style_subtitle": "Subtítulo",
        "style_subtitle2": "Subtítulo2",
        "style_h1": "Encabezado 1",
        "style_h2": "Encabezado 2",
        "style_quote": "Cita",
        "style_code": "Código",
        "style_mono": "Monospace",
        "style_small": "Pequeño",
        "style_large": "Grande",
        "style_emphasis": "Énfasis",
        "rows_label": "Filas:",
        "cols_label": "Columnas:",
        "insert_image_error": "No se puede cargar la imagen:",
        "table_title": "Insertar tabla",
        "text_color": "Color de texto",
        "bg_color": "Color de fondo"
    }
}

LANG_CHOICES = [
    ("English (US)", "en-US"),
    ("English (UK)", "en-GB"),
    ("简体中文", "zh-CN"),
    ("繁體中文", "zh-TW"),
    ("日本語", "ja-JP"),
    ("Español", "es-ES"),
]
def get_desktop_path():
    if platform.system() == "Windows":
        return os.path.join(os.path.expanduser("~"), "Desktop")
    else:
        return os.path.expanduser("~")


def ensure_extension(filename):
    if not filename.lower().endswith(".voc"):
        return filename + ".voc"
    return filename


def _detect_image_mime(path):
    t = imghdr.what(path)
    if t:
        if t == 'jpeg':
            return 'image/jpeg'
        return f'image/{t}'
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.jpg', '.jpeg'):
        return 'image/jpeg'
    if ext == '.png':
        return 'image/png'
    if ext == '.gif':
        return 'image/gif'
    return 'application/octet-stream'


def _encode_file_to_data_uri(path):
    try:
        with open(path, 'rb') as f:
            data = f.read()
        b64 = base64.b64encode(data).decode('ascii')
        mime = _detect_image_mime(path)
        return f"data:{mime};base64,{b64}"
    except Exception:
        return None


def _replace_file_src_with_data_uris(html):
    """
    Replace file:// image src with data URIs
    """
    pattern = re.compile(r'src=(["\'])(file:///?[^"\']+)\1', re.IGNORECASE)

    def repl(m):
        quote = m.group(1)
        file_url = m.group(2)
        parsed = urlparse(file_url)
        path = parsed.path
        if platform.system() == "Windows" and re.match(r'^/[a-zA-Z]:', path):
            path = path.lstrip('/')
        path = unquote(path)
        if os.path.exists(path):
            data_uri = _encode_file_to_data_uri(path)
            if data_uri:
                return f'src={quote}{data_uri}{quote}'
        return m.group(0)

    return pattern.sub(repl, html)
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_filepath = None
        self.lang = "en-US"
        self.trans = TRANSLATIONS[self.lang]

        self.setWindowTitle(self.trans["app_title"])
        self.resize(1200, 780)

        central = QWidget()
        central_layout = QVBoxLayout()
        central.setLayout(central_layout)
        self.setCentralWidget(central)

        # Toolbar
        self.toolbar = QToolBar("Formatting")
        self.toolbar.setIconSize(QSize(18, 18))
        self.toolbar.setStyleSheet("background: #FFFFFF; border-bottom: 1px solid #cfe8fb;")
        self.addToolBar(self.toolbar)

        # Editor
        self.editor = QTextEdit()
        self.editor.setStyleSheet("background: #E7F0FA; padding: 10px;")
        self.editor.setAcceptRichText(True)
        central_layout.addWidget(self.editor)

        # Bottom: search and file list
        bottom_widget = QWidget()
        bottom_layout = QHBoxLayout()
        bottom_widget.setLayout(bottom_layout)
        central_layout.addWidget(bottom_widget, stretch=0)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(self.trans["search"])
        bottom_layout.addWidget(self.search_input)

        self.search_btn = QPushButton(self.trans["search"])
        self.search_btn.clicked.connect(self.search_voc_files)
        bottom_layout.addWidget(self.search_btn)

        self.search_dir_combo = QComboBox()
        self.search_dir_combo.addItem(self.trans["desktop_search"], get_desktop_path())
        self.search_dir_combo.addItem("Home", os.path.expanduser("~"))
        bottom_layout.addWidget(self.search_dir_combo)

        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(160)
        self.files_list.itemDoubleClicked.connect(self.open_voc_from_list)
        central_layout.addWidget(self.files_list)

        # Build actions and UI
        self._create_actions()
        self._create_format_toolbar()
        self._create_menus()
        self.retranslate_ui()

    # ---------------------------
    # Actions
    # ---------------------------
    def _create_actions(self):
        # File actions
        self.act_new = QAction("", self)
        self.act_new.setShortcut("Ctrl+N")
        self.act_new.triggered.connect(self.action_new)

        self.act_open = QAction("", self)
        self.act_open.setShortcut("Ctrl+O")
        self.act_open.triggered.connect(self.action_open)

        self.act_save = QAction("", self)
        self.act_save.setShortcut("Ctrl+S")
        self.act_save.triggered.connect(self.action_save)

        self.act_save_as = QAction("", self)
        self.act_save_as.triggered.connect(self.action_save_as)

        self.act_exit = QAction("", self)
        self.act_exit.triggered.connect(self.close)

        # Formatting actions
        self.act_bold = QAction("B", self)
        self.act_bold.setCheckable(True)
        self.act_bold.setShortcut("Ctrl+B")
        self.act_bold.triggered.connect(self.toggle_bold)

        self.act_italic = QAction("I", self)
        self.act_italic.setCheckable(True)
        self.act_italic.setShortcut("Ctrl+I")
        self.act_italic.triggered.connect(self.toggle_italic)

        self.act_underline = QAction("U", self)
        self.act_underline.setCheckable(True)
        self.act_underline.setShortcut("Ctrl+U")
        self.act_underline.triggered.connect(self.toggle_underline)

        self.act_align_left = QAction("L", self)
        self.act_align_left.triggered.connect(partial(self.set_alignment, Qt.AlignLeft))

        self.act_align_center = QAction("C", self)
        self.act_align_center.setShortcut("Ctrl+E")
        self.act_align_center.triggered.connect(partial(self.set_alignment, Qt.AlignCenter))

        self.act_align_right = QAction("R", self)
        self.act_align_right.triggered.connect(partial(self.set_alignment, Qt.AlignRight))

        # Undo/Redo
        self.act_undo = QAction("", self)
        self.act_undo.setShortcut("Ctrl+Z")
        self.act_undo.triggered.connect(self.editor.undo)

        self.act_redo = QAction("", self)
        self.act_redo.setShortcut("Ctrl+Y")
        self.act_redo.triggered.connect(self.editor.redo)

        # Insert image / table
        self.act_insert_image = QAction("", self)
        self.act_insert_image.triggered.connect(self.insert_image)

        self.act_insert_table = QAction("", self)
        self.act_insert_table.triggered.connect(self.insert_table)

        # Styles
        self.style_label = QLabel(self.trans.get("style_label", "Style:"))
        self.style_combo = QComboBox()
        self.style_combo.currentIndexChanged.connect(self.apply_style_from_combo)

        # Color actions (use buttons)
        self.text_color_btn = QPushButton("A")
        self.text_color_btn.setToolTip(self.trans.get("text_color", "Text Color"))
        self.text_color_btn.clicked.connect(self.choose_text_color)
        self.text_color_btn.setFixedWidth(32)
        self.text_color_btn.setStyleSheet("font-weight: bold;")

        self.bg_color_btn = QPushButton("▯")
        self.bg_color_btn.setToolTip(self.trans.get("bg_color", "Highlight Color"))
        self.bg_color_btn.clicked.connect(self.choose_bg_color)
        self.bg_color_btn.setFixedWidth(32)

    def _create_format_toolbar(self):
        # Font family
        self.font_combo = QFontComboBox()
        self.font_combo.currentFontChanged.connect(self.set_font_family)
        self.toolbar.addWidget(self.font_combo)
        try:
            self.font_combo.view().setStyleSheet(COMBO_POPUP_QSS)
        except Exception:
            pass

        # Font size
        self.font_size = QSpinBox()
        self.font_size.setRange(6, 96)
        self.font_size.setValue(12)
        self.font_size.valueChanged.connect(self.set_font_size)
        self.toolbar.addWidget(self.font_size)

        # Format buttons
        self.toolbar.addAction(self.act_bold)
        self.toolbar.addAction(self.act_italic)
        self.toolbar.addAction(self.act_underline)

        self.toolbar.addSeparator()
        self.toolbar.addAction(self.act_align_left)
        self.toolbar.addAction(self.act_align_center)
        self.toolbar.addAction(self.act_align_right)

        self.toolbar.addSeparator()
        # Style label and combo
        self.toolbar.addWidget(self.style_label)
        self.toolbar.addWidget(self.style_combo)
        try:
            self.style_combo.view().setStyleSheet(COMBO_POPUP_QSS)
        except Exception:
            pass

        self.toolbar.addSeparator()
        self.toolbar.addAction(self.act_insert_image)
        self.toolbar.addAction(self.act_insert_table)

        self.toolbar.addSeparator()
        self.toolbar.addAction(self.act_undo)
        self.toolbar.addAction(self.act_redo)

        self.toolbar.addSeparator()
        # Color buttons
        self.toolbar.addWidget(QLabel(" "))
        self.toolbar.addWidget(self.text_color_btn)
        self.toolbar.addWidget(self.bg_color_btn)

        self.toolbar.addSeparator()
        # Language combo
        self.lang_combo = QComboBox()
        for name, code in LANG_CHOICES:
            self.lang_combo.addItem(name, code)
        self.lang_combo.setCurrentIndex(0)
        self.lang_combo.currentIndexChanged.connect(self.lang_changed)
        self.toolbar.addWidget(QLabel(" "))
        self.toolbar.addWidget(self.lang_combo)
        try:
            self.lang_combo.view().setStyleSheet(COMBO_POPUP_QSS)
        except Exception:
            pass

        # Ensure search_dir_combo popup style
        try:
            self.search_dir_combo.view().setStyleSheet(COMBO_POPUP_QSS)
        except Exception:
            pass

    def _create_menus(self):
        menubar = self.menuBar()
        # create menus but set titles in retranslate_ui for translation support
        self.menu_file = menubar.addMenu("")
        self.menu_file.addAction(self.act_new)
        self.menu_file.addAction(self.act_open)
        self.menu_file.addAction(self.act_save)
        self.menu_file.addAction(self.act_save_as)
        self.menu_file.addSeparator()
        self.menu_file.addAction(self.act_exit)

        self.menu_edit = menubar.addMenu("")
        self.menu_edit.addAction(self.act_undo)
        self.menu_edit.addAction(self.act_redo)

        self.menu_insert = menubar.addMenu("")
        self.menu_insert.addAction(self.act_insert_image)
        self.menu_insert.addAction(self.act_insert_table)

    # ---------------------------
    # Formatting helpers
    # ---------------------------
    def toggle_bold(self):
        fmt = QTextCharFormat()
        current_weight = self.editor.fontWeight()
        if current_weight == QFont.Bold:
            fmt.setFontWeight(QFont.Normal)
            self.act_bold.setChecked(False)
        else:
            fmt.setFontWeight(QFont.Bold)
            self.act_bold.setChecked(True)
        cursor = self.editor.textCursor()
        cursor.mergeCharFormat(fmt)
        self.editor.mergeCurrentCharFormat(fmt)

    def toggle_italic(self):
        cursor = self.editor.textCursor()
        fmt = QTextCharFormat()
        fmt.setFontItalic(not self.editor.fontItalic())
        self.act_italic.setChecked(fmt.fontItalic())
        cursor.mergeCharFormat(fmt)
        self.editor.mergeCurrentCharFormat(fmt)

    def toggle_underline(self):
        cursor = self.editor.textCursor()
        fmt = QTextCharFormat()
        fmt.setFontUnderline(not self.editor.fontUnderline())
        self.act_underline.setChecked(fmt.fontUnderline())
        cursor.mergeCharFormat(fmt)
        self.editor.mergeCurrentCharFormat(fmt)

    def set_alignment(self, alignment):
        self.editor.setAlignment(alignment)

    def set_font_family(self, qfont):
        fmt = QTextCharFormat()
        fmt.setFontFamily(qfont.family())
        cursor = self.editor.textCursor()
        cursor.mergeCharFormat(fmt)
        self.editor.mergeCurrentCharFormat(fmt)

    def set_font_size(self, size):
        fmt = QTextCharFormat()
        fmt.setFontPointSize(size)
        cursor = self.editor.textCursor()
        cursor.mergeCharFormat(fmt)
        self.editor.mergeCurrentCharFormat(fmt)

    # ---------------------------
    # Styles (extended)
    # ---------------------------
    def apply_style_from_combo(self, index):
        style = self.style_combo.itemText(index)
        self.apply_style(style)

    def apply_style(self, style_name):
        cursor = self.editor.textCursor()
        if not cursor:
            return
        block_fmt = QTextBlockFormat()
        char_fmt = QTextCharFormat()

        # Match against translated labels and common English names
        title_names = (self.trans.get("style_title"), "Title", "标题", "タイトル")
        subtitle_names = (self.trans.get("style_subtitle"), "Subtitle", "副标题", "サブタイトル")
        subtitle2_names = (self.trans.get("style_subtitle2"), "Subtitle 2")
        h1_names = (self.trans.get("style_h1"), "Heading 1")
        h2_names = (self.trans.get("style_h2"), "Heading 2")
        quote_names = (self.trans.get("style_quote"), "Quote")
        code_names = (self.trans.get("style_code"), "Code", "代码")
        mono_names = (self.trans.get("style_mono"), "Monospace", "等宽")
        small_names = (self.trans.get("style_small"), "Small")
        large_names = (self.trans.get("style_large"), "Large")
        emphasis_names = (self.trans.get("style_emphasis"), "Emphasis", "强调")

        if style_name in title_names:
            block_fmt.setAlignment(Qt.AlignCenter)
            char_fmt.setFontPointSize(28)
            char_fmt.setFontWeight(QFont.Bold)
        elif style_name in subtitle_names:
            block_fmt.setAlignment(Qt.AlignCenter)
            char_fmt.setFontPointSize(20)
            char_fmt.setFontItalic(True)
        elif style_name in subtitle2_names:
            block_fmt.setAlignment(Qt.AlignCenter)
            char_fmt.setFontPointSize(16)
        elif style_name in h1_names:
            block_fmt.setAlignment(Qt.AlignLeft)
            char_fmt.setFontPointSize(24)
            char_fmt.setFontWeight(QFont.Bold)
        elif style_name in h2_names:
            block_fmt.setAlignment(Qt.AlignLeft)
            char_fmt.setFontPointSize(18)
            char_fmt.setFontWeight(QFont.Bold)
        elif style_name in quote_names:
            block_fmt.setLeftMargin(24)
            char_fmt.setFontItalic(True)
            char_fmt.setFontPointSize(12)
        elif style_name in code_names:
            # monospace look with gray background
            char_fmt.setFontFamily("Courier New")
            char_fmt.setFontPointSize(11)
            char_fmt.setFontFixedPitch(True)
            char_fmt.setBackground(QBrush(QColor("#F5F5F5")))
        elif style_name in mono_names:
            char_fmt.setFontFamily("Courier New")
            char_fmt.setFontPointSize(12)
        elif style_name in small_names:
            char_fmt.setFontPointSize(10)
        elif style_name in large_names:
            char_fmt.setFontPointSize(16)
        elif style_name in emphasis_names:
            char_fmt.setFontItalic(True)
            char_fmt.setFontWeight(QFont.DemiBold)
        else:
            # Normal
            block_fmt.setAlignment(Qt.AlignLeft)
            char_fmt.setFontPointSize(12)
            char_fmt.setFontWeight(QFont.Normal)
            char_fmt.setFontItalic(False)

        cursor.beginEditBlock()
        cursor.setBlockFormat(block_fmt)
        cursor.mergeCharFormat(char_fmt)
        self.editor.mergeCurrentCharFormat(char_fmt)
        cursor.endEditBlock()

    # ---------------------------
    # Color features
    # ---------------------------
    def choose_text_color(self):
        color = QColorDialog.getColor(parent=self, title=self.trans.get("text_color", "Text Color"))
        if not color.isValid():
            return
        fmt = QTextCharFormat()
        fmt.setForeground(QBrush(color))
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            cursor.mergeCharFormat(fmt)
        else:
            # apply for typing forward
            self.editor.mergeCurrentCharFormat(fmt)
        # update button background to indicate color
        self.text_color_btn.setStyleSheet(f"background: {color.name()}; color: white; font-weight:bold;")

    def choose_bg_color(self):
        color = QColorDialog.getColor(parent=self, title=self.trans.get("bg_color", "Highlight Color"))
        if not color.isValid():
            return
        fmt = QTextCharFormat()
        fmt.setBackground(QBrush(color))
        cursor = self.editor.textCursor()
        if cursor.hasSelection():
            cursor.mergeCharFormat(fmt)
        else:
            self.editor.mergeCurrentCharFormat(fmt)
        self.bg_color_btn.setStyleSheet(f"background: {color.name()}; color: white;")

    # ---------------------------
    # Insert image / table (robust)
    # ---------------------------
    def insert_image(self):
        path, _ = QFileDialog.getOpenFileName(self, self.trans["insert_image"], "", "Images (*.png *.jpg *.jpeg *.bmp *.gif);;All Files (*)")
        if not path:
            return
        cursor = self.editor.textCursor()
        image = QImage(path)
        if image.isNull():
            QMessageBox.warning(self, self.trans["insert_image"], f"{self.trans.get('insert_image_error', 'Cannot load image:')} {path}")
            return
        img_fmt = QTextImageFormat()
        file_url = 'file:///' + path.replace('\\', '/')
        img_fmt.setName(file_url)
        w = image.width()
        h = image.height()
        max_w = 600
        if w > max_w:
            ratio = max_w / w
            img_fmt.setWidth(w * ratio)
            img_fmt.setHeight(h * ratio)
        cursor.insertImage(img_fmt)

    def insert_table(self):
        # 获取行列，使用翻译文本
        rows, ok1 = QInputDialog.getInt(self, self.trans.get("table_title", "Insert Table"), self.trans.get("rows_label", "Rows:"), 2, 1, 50)
        if not ok1:
            return
        cols, ok2 = QInputDialog.getInt(self, self.trans.get("table_title", "Insert Table"), self.trans.get("cols_label", "Columns:"), 2, 1, 20)
        if not ok2:
            return
        cursor = self.editor.textCursor()
        # 使用 QTextTableFormat 避免游标问题
        try:
            table_fmt = QTextTableFormat()
            table_fmt.setBorder(1)
            table_fmt.setCellPadding(4)
            table_fmt.setCellSpacing(2)
            table = cursor.createTable(rows, cols, table_fmt)
            # 将光标定位到第一个单元格的起始位置（更安全）
            try:
                cell = table.cellAt(0, 0)
                new_cursor = cell.firstCursorPosition()
                self.editor.setTextCursor(new_cursor)
            except Exception:
                pass
        except Exception as e:
            QMessageBox.warning(self, self.trans.get("table_title", "Insert Table"), f"Failed to insert table:\n{e}")

    # ---------------------------
    # File operations
    # ---------------------------
    def action_new(self):
        reply = QMessageBox.question(self, self.trans["new"], self.trans["confirm_clear"],
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.current_filepath = None
            self.editor.clear()

    def action_open(self):
        desktop = get_desktop_path()
        fname, _ = QFileDialog.getOpenFileName(self, self.trans["open"], desktop, "VOC files (*.voc);;All Files (*)")
        if fname:
            try:
                with open(fname, "r", encoding="utf-8") as f:
                    data = json.load(f)
                html = data.get("content", "")
                self.editor.setHtml(html)
                self.current_filepath = fname
            except Exception as e:
                QMessageBox.warning(self, self.trans["open"], f"{self.trans['open_error']}\n{e}")

    def action_save(self):
        desktop = get_desktop_path()
        if platform.system() != "Windows":
            QMessageBox.information(self, "Info", self.trans["not_windows"])
        if self.current_filepath:
            try:
                if os.path.commonpath([os.path.abspath(self.current_filepath), os.path.abspath(desktop)]) == os.path.abspath(desktop):
                    target = self.current_filepath
                else:
                    name, ok = QInputDialog.getText(self, self.trans["save"], self.trans["filename_prompt"])
                    if not ok or not name.strip():
                        return
                    filename = ensure_extension(name.strip())
                    target = os.path.join(desktop, filename)
            except Exception:
                name, ok = QInputDialog.getText(self, self.trans["save"], self.trans["filename_prompt"])
                if not ok or not name.strip():
                    return
                filename = ensure_extension(name.strip())
                target = os.path.join(desktop, filename)
        else:
            name, ok = QInputDialog.getText(self, self.trans["save"], self.trans["filename_prompt"])
            if not ok or not name.strip():
                return
            filename = ensure_extension(name.strip())
            target = os.path.join(desktop, filename)

        self._write_voc_file(target)
        self.current_filepath = target
        QMessageBox.information(self, self.trans["save"], self.trans["saved"])

    def action_save_as(self):
        desktop = get_desktop_path()
        if platform.system() != "Windows":
            QMessageBox.information(self, "Info", self.trans["not_windows"])
        name, ok = QInputDialog.getText(self, self.trans["save_as"], self.trans["filename_prompt"])
        if not ok or not name.strip():
            return
        filename = ensure_extension(name.strip())
        target = os.path.join(desktop, filename)
        self._write_voc_file(target)
        self.current_filepath = target
        QMessageBox.information(self, self.trans["save_as"], self.trans["saved"])

    def _write_voc_file(self, path):
        try:
            html = self.editor.toHtml()
            html = _replace_file_src_with_data_uris(html)
            data = {
                "content": html,
                "meta": {
                    "saved_by": "Voc Editor (Python/PyQt5)",
                    "platform": platform.platform(),
                }
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.warning(self, self.trans["save"], f"Save failed: {e}")

    # ---------------------------
    # Search & quick open
    # ---------------------------
    def search_voc_files(self):
        self.files_list.clear()
        search_term = self.search_input.text().strip().lower()
        dirpath = self.search_dir_combo.currentData()
        if not dirpath:
            dirpath = get_desktop_path()
        results = []
        try:
            for entry in os.listdir(dirpath):
                full = os.path.join(dirpath, entry)
                if os.path.isfile(full) and entry.lower().endswith('.voc'):
                    if not search_term or search_term in entry.lower():
                        results.append(full)
        except Exception:
            pass

        for p in results:
            item = QListWidgetItem(p)
            self.files_list.addItem(item)

    def open_voc_from_list(self, item: QListWidgetItem):
        path = item.text()
        if not path:
            QMessageBox.warning(self, self.trans["open"], self.trans["no_file_selected"])
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            html = data.get("content", "")
            self.editor.setHtml(html)
            self.current_filepath = path
        except Exception as e:
            QMessageBox.warning(self, self.trans["open"], f"{self.trans['open_error']}\n{e}")

    # ---------------------------
    # Language
    # ---------------------------
    def lang_changed(self, index):
        code = self.lang_combo.itemData(index)
        if not code:
            return
        self.lang = code
        self.trans = TRANSLATIONS.get(self.lang, TRANSLATIONS["en-US"])
        self.retranslate_ui()

    def retranslate_ui(self):
        # Window title and menus/actions
        self.setWindowTitle(self.trans["app_title"])
        self.act_new.setText(self.trans["new"])
        self.act_open.setText(self.trans["open"])
        self.act_save.setText(self.trans["save"])
        self.act_save_as.setText(self.trans["save_as"])
        self.act_exit.setText(self.trans["exit"])

        # Update menu titles
        self.menu_file.setTitle(self.trans.get("menu_file", "File"))
        self.menu_edit.setTitle(self.trans.get("menu_edit", "Edit"))
        self.menu_insert.setTitle(self.trans.get("menu_insert", "Insert"))

        # Tooltips and buttons
        self.act_bold.setToolTip(self.trans["bold"])
        self.act_italic.setToolTip(self.trans["italic"])
        self.act_underline.setToolTip(self.trans["underline"])
        self.act_align_left.setToolTip(self.trans["align_left"])
        self.act_align_center.setToolTip(self.trans["align_center"])
        self.act_align_right.setToolTip(self.trans["align_right"])

        self.search_input.setPlaceholderText(self.trans["search"])
        self.search_btn.setText(self.trans["search"])
        self.search_dir_combo.setItemText(0, self.trans["desktop_search"])

        self.act_insert_image.setText(self.trans["insert_image"])
        self.act_insert_table.setText(self.trans["insert_table"])

        self.act_undo.setText(self.trans["undo"])
        self.act_redo.setText(self.trans["redo"])

        # Style label and combo items
        self.style_label.setText(self.trans.get("style_label", "Style:"))
        current_index = self.style_combo.currentIndex()
        self.style_combo.blockSignals(True)
        self.style_combo.clear()
        self.style_combo.addItem(self.trans.get("style_normal", "Normal"))
        self.style_combo.addItem(self.trans.get("style_title", "Title"))
        self.style_combo.addItem(self.trans.get("style_subtitle", "Subtitle"))
        self.style_combo.addItem(self.trans.get("style_subtitle2", "Subtitle 2"))
        self.style_combo.addItem(self.trans.get("style_h1", "Heading 1"))
        self.style_combo.addItem(self.trans.get("style_h2", "Heading 2"))
        self.style_combo.addItem(self.trans.get("style_quote", "Quote"))
        self.style_combo.addItem(self.trans.get("style_code", "Code"))
        self.style_combo.addItem(self.trans.get("style_mono", "Monospace"))
        self.style_combo.addItem(self.trans.get("style_small", "Small"))
        self.style_combo.addItem(self.trans.get("style_large", "Large"))
        self.style_combo.addItem(self.trans.get("style_emphasis", "Emphasis"))
        if 0 <= current_index < self.style_combo.count():
            self.style_combo.setCurrentIndex(current_index)
        self.style_combo.blockSignals(False)

        # Update color button tooltips
        self.text_color_btn.setToolTip(self.trans.get("text_color", "Text Color"))
        self.bg_color_btn.setToolTip(self.trans.get("bg_color", "Highlight Color"))

        # Reapply popup styles
        try:
            self.style_combo.view().setStyleSheet(COMBO_POPUP_QSS)
        except Exception:
            pass
        try:
            self.font_combo.view().setStyleSheet(COMBO_POPUP_QSS)
        except Exception:
            pass
        try:
            self.lang_combo.view().setStyleSheet(COMBO_POPUP_QSS)
        except Exception:
            pass
        try:
            self.search_dir_combo.view().setStyleSheet(COMBO_POPUP_QSS)
        except Exception:
            pass


# ---------------------------
# Main
# ---------------------------
def main():
    app = QApplication(sys.argv)
    # App-wide style (light blue + white) + menus/combo style
    style = """
    QMainWindow { background: #E7F0FA; }
    QMenuBar { background: #FFFFFF; }
    QMenu { background: #FFFFFF; }
    QToolBar { background: #FFFFFF; }
    QPushButton { background: #FFFFFF; border: 1px solid #cfe8fb; padding: 4px; }
    QListWidget { background: #FFFFFF; }
    QLineEdit { background: #FFFFFF; padding: 4px; border: 1px solid #cfe8fb; }
    """ + GLOBAL_QSS_MENU
    app.setStyleSheet(style)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
