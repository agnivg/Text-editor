
import sys
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *
import docx2txt


class DocApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # window title
        self.title = "GoDocs"
        self.setWindowTitle(self.title)
        
        # editor section
        self.editor = QTextEdit(self) 
        self.setCentralWidget(self.editor)

        # create menuBar and toolbar first
        self.create_menu_bar()
        self.create_toolbar()

        # after creating toolbar we select font and font size
        font = QFont('Times', 12)
        self.editor.setFont(font)
        self.editor.setFontPointSize(12)

        # stores path
        self.path = ''

    def create_menu_bar(self):
        menuBar = QMenuBar(self)

        # file menu
        file_menu = QMenu("File", self)
        menuBar.addMenu(file_menu)

        save = QAction('Save', self)
        save.triggered.connect(self.file_save)
        file_menu.addAction(save)

        open = QAction('Open', self)
        open.triggered.connect(self.file_open)
        file_menu.addAction(open)

        saveas = QAction('Save As', self)
        saveas.triggered.connect(self.file_saveas)
        file_menu.addAction(saveas)

        pdf = QAction("Save as PDF", self)
        pdf.triggered.connect(self.save_pdf)
        file_menu.addAction(pdf)

        # edit menu
        edit_menu = QMenu("Edit", self)
        menuBar.addMenu(edit_menu)

        # cut
        cut = QAction('Cut', self)
        cut.triggered.connect(self.editor.cut)
        edit_menu.addAction(cut)

        # copy
        copy = QAction('Copy', self)
        copy.triggered.connect(self.editor.copy)
        edit_menu.addAction(copy)

        # paste
        paste = QAction('Paste', self)
        paste.triggered.connect(self.editor.paste)
        edit_menu.addAction(paste)

        # Clear Screen
        clearsc = QAction('ClearScreen', self)
        clearsc.triggered.connect(self.editor.clear)
        edit_menu.addAction(clearsc)

        # Select All
        selectall = QAction('Select All', self)
        selectall.triggered.connect(self.editor.selectAll)
        edit_menu.addAction(selectall)

        # view menu
        view_menu = QMenu("View", self)
        menuBar.addMenu(view_menu)

        # full screen
        fullsc = QAction('Full Screen View', self)
        fullsc.triggered.connect(lambda: self.showFullScreen())
        view_menu.addAction(fullsc)

        # normal screen
        normsc = QAction('Normal View', self)
        normsc.triggered.connect(lambda: self.showNormal())
        view_menu.addAction(normsc)

        # minimize
        minsc = QAction('Minimize', self)
        minsc.triggered.connect(lambda: self.showMinimized())
        view_menu.addAction(minsc)

        self.setMenuBar(menuBar)

    def create_toolbar(self):
        # Using a title
        toolBar = QToolBar("Tools", self)

        # save
        save = QAction(QIcon("save.png"), 'Save', self)
        save.triggered.connect(self.file_save)
        toolBar.addAction(save)

        # undo
        undo = QAction(QIcon("undo.png"), 'Undo', self)
        undo.triggered.connect(self.editor.undo)
        toolBar.addAction(undo)

        # redo
        redo = QAction(QIcon("redo.png"), 'Redo', self)
        redo.triggered.connect(self.editor.redo)
        toolBar.addAction(redo)

        # adding separator
        toolBar.addSeparator()

        # cut
        cut = QAction(QIcon("cut.png"), 'Cut', self)
        cut.triggered.connect(self.editor.cut)
        toolBar.addAction(cut)

        # copy
        copy = QAction(QIcon("copy.png"), 'Copy', self)
        copy.triggered.connect(self.editor.copy)
        toolBar.addAction(copy)

        # paste
        paste = QAction(QIcon("paste.png"), 'Paste', self)
        paste.triggered.connect(self.editor.paste)
        toolBar.addAction(paste)

        # adding separator
        toolBar.addSeparator()
        toolBar.addSeparator()

        # fonts
        self.font_name = QComboBox(self)
        self.font_name.addItems(["Courier", "Serif", "Courier", "Typewriter", "OldEnglish", "Decorative", "Helvetica", "Fantasy", "SansSerif", "Times", "Monospace", "Cursive"])
        self.font_name.activated.connect(self.set_font_name)
        toolBar.addWidget(self.font_name)
        toolBar.addSeparator()

        # font size
        self.font_size = QSpinBox(self)   
        self.font_size.setValue(12)  
        self.font_size.valueChanged.connect(self.set_font_size)
        toolBar.addWidget(self.font_size)

        # separator
        toolBar.addSeparator()
        toolBar.addSeparator()

        # bold
        bold = QAction(QIcon("bold.png"), 'Bold', self)
        bold.triggered.connect(self.bold_text)
        toolBar.addAction(bold)

        # underline
        underline = QAction(QIcon("underline.png"), 'Underline', self)
        underline.triggered.connect(self.underline_text)
        toolBar.addAction(underline)

        # italic
        italic = QAction(QIcon("italic.png"), 'Italic', self)
        italic.triggered.connect(self.italic_text)
        toolBar.addAction(italic)

        # separator
        toolBar.addSeparator()
        toolBar.addSeparator()

        # text alignment
        left_alignment = QAction(QIcon("left-align.png"), 'Left', self)
        left_alignment.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignLeft))
        toolBar.addAction(left_alignment)
        toolBar.addSeparator()

        center_alignment = QAction(QIcon("center-align.png"), 'Center', self)
        center_alignment.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignCenter))
        toolBar.addAction(center_alignment)
        toolBar.addSeparator()

        right_alignment = QAction(QIcon("right-align.png"), 'Right', self)
        right_alignment.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignRight))
        toolBar.addAction(right_alignment)

        # separator
        toolBar.addSeparator()
        toolBar.addSeparator()

        # zoom in
        zoomin = QAction(QIcon("zoom-in.png"), 'Zoom in', self)
        zoomin.triggered.connect(self.editor.zoomIn)
        toolBar.addAction(zoomin)

        # zoom out
        zoomout = QAction(QIcon("zoom-out.png"), 'Zoom out', self)
        zoomout.triggered.connect(self.editor.zoomOut)
        toolBar.addAction(zoomout)

        # separator
        toolBar.addSeparator()
        toolBar.addSeparator()
        
        self.addToolBar(toolBar)

    def italic_text(self):
        # make normal if already italic, else italic
        state = self.editor.fontItalic()
        self.editor.setFontItalic(not state)

    def underline_text(self):
        # make normal if already underlined, else underlined
        state = self.editor.fontUnderline()
        self.editor.setFontUnderline(not state)

    def bold_text(self):
        # make normal if already bold, else bold
        if self.editor.fontWeight() != QFont.Bold:
            self.editor.setFontWeight(QFont.Bold)
        else:
            self.editor.setFontWeight(QFont.Normal)

    def set_font_name(self):
        font = self.font_name.currentText()
        self.editor.setCurrentFont(QFont(font))

    def set_font_size(self):
        value = self.font_size.value()
        self.editor.setFontPointSize(value)  

    def file_open(self):
        self.path, _ = QFileDialog.getOpenFileName(self, "Open file", "", "Text documents (*.text);All files (*.*)")

        try:
            text = docx2txt.process(self.path)
        except Exception as e:
            print(e)
        else:
            self.editor.setText(text)
            self.update_title()

    def file_save(self):
        if self.path == '':
            # If we do not have a path, we need to use Save As.
            self.file_saveas()

        text = self.editor.toPlainText()

        try:
            with open(self.path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def file_saveas(self):
        self.path, _ = QFileDialog.getSaveFileName(self, "Save file", "", "Text documents (*.text);All files (*.*)")

        if self.path == '':
            return   # If Save As is cancelled

        text = self.editor.toPlainText()

        try:
            with open(path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def update_title(self):
        self.setWindowTitle(self.title + ' ' + self.path)

    def save_pdf(self):
        f_name, _ = QFileDialog.getSaveFileName(self, "Export PDF", None, "PDF files (.pdf);;All files()")
        if f_name != '':  # if name not empty
           printer = QPrinter(QPrinter.HighResolution)
           printer.setOutputFormat(QPrinter.PdfFormat)
           printer.setOutputFileName(f_name)
           self.editor.document().print_(printer)
    

app = QApplication(sys.argv)
window = DocApp()
window.show()
sys.exit(app.exec_())