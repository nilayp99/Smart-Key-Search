import sys

import fitz

from PyQt5.QtWidgets import (

    QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, QFileDialog, QGraphicsTransform, QSizePolicy,

    QPushButton, QTextEdit, QScrollArea, QListWidget, QListWidgetItem, QHBoxLayout, QSplitter, QItemDelegate, QGroupBox

)

from PyQt5.QtGui import QPixmap, QImage, QColor, QFont, QTransform, QPainter, QIcon

from PyQt5.QtCore import Qt, QSize

import os

import io

from reportlab.lib.pagesizes import letter

from reportlab.pdfgen import canvas

from reportlab.lib import colors

from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

from reportlab.lib.colors import yellow, cyan, magenta, red, blue, pink, orange, green

from docx import Document

import concurrent.futures

 

class VerticalLabel(QLabel):

    def __init__(self, text, parent=None):

        super().__init__(text, parent)

 

    def paintEvent(self, event):

        painter = QPainter(self)

        painter.setPen(QColor(85, 85, 85))

        painter.setFont(QFont("Arial", 6))

 

        painter.translate(self.width()-3, self.height()-14)

        painter.rotate(-90)

       

        painter.drawText(0, 0, self.text())

        painter.end()

 

    def sizeHint(self):

        return QSize(8, 1100)

 

class PDFViewer(QWidget):

    def __init__(self, parent=None):

        super().__init__(parent)

        self.zoom_level = 1.0

        self.current_page = 0

        self.doc = None

        self.initUI()

 

    def initUI(self):

        self.layout = QVBoxLayout()

 

        #--Scroll Area----------------------------------------------------------------------------------#

        self.scrollArea = QScrollArea()

        self.scrollWidget = QWidget()

        self.scrollLayout = QVBoxLayout()

        self.scrollLayout.setAlignment(Qt.AlignCenter)

        self.scrollWidget.setLayout(self.scrollLayout)

        self.scrollArea.setWidget(self.scrollWidget)

        self.scrollArea.setWidgetResizable(True)

        self.scrollArea.verticalScrollBar().valueChanged.connect(self.on_scroll)

        self.layout.addWidget(self.scrollArea)

        #----------------------------------------------------------------------------------------------#

 

        #--Bottom Functions----------------------------------------------------------------------------#

        self.button_layout = QHBoxLayout()

        # self.button_layout.addStretch()

 

        self.file_name_label = QLabel(self)

        self.file_name_label.setFixedSize(QSize(315, 18))

        self.button_layout.addWidget(self.file_name_label, alignment=Qt.AlignCenter | Qt.AlignLeft)

 

        self.page_label = QLabel(self)

        self.page_label.setFixedSize(QSize(150, 30))

        self.page_label.setAlignment(Qt.AlignRight | Qt.AlignCenter)

        self.button_layout.addWidget(self.page_label)

 

        self.zoom_out_button = QPushButton('-', self)

        self.zoom_out_button.setFixedSize(QSize(30, 30))

        self.zoom_out_button.clicked.connect(self.zoom_out)

        self.button_layout.addWidget(self.zoom_out_button)

 

        self.zoom_in_button = QPushButton('+', self)

        self.zoom_in_button.setFixedSize(QSize(30, 30))

        self.zoom_in_button.clicked.connect(self.zoom_in)

        self.button_layout.addWidget(self.zoom_in_button)

 

        self.layout.addLayout(self.button_layout)

        #----------------------------------------------------------------------------------------------#

 

        self.setLayout(self.layout)

        self.scrollArea.viewport().installEventFilter(self)

 

    def eventFilter(self, source, event):

        return super(PDFViewer, self).eventFilter(source, event)

 

    def clear_layout(self, layout):

        if layout is not None:

            while layout.count():

                child = layout.takeAt(0)

                if child.widget():

                    child.widget().deleteLater()

 

    def display_pdf(self, pdf_path, page_num=0):

        self.pdf_path = pdf_path

        self.file_name_label.setText(os.path.basename(pdf_path).split('_highlighted')[0]+'.pdf')

 

        if hasattr(self, 'pdf_path'):

            self.doc = fitz.open(self.pdf_path)

            self.clear_layout(self.scrollLayout)

 

            self.page_labels = []

            for p_num in range(len(self.doc)):

                self.add_page_to_layout(p_num)

            self.current_page = page_num

            self.scroll_to_page(page_num)

            self.update_page_label()

 

    def add_page_to_layout(self, page_num):

        page = self.doc.load_page(page_num)

        zoom_matrix = fitz.Matrix(self.zoom_level, self.zoom_level)

        pix = page.get_pixmap(matrix=zoom_matrix)

        image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)

        pixmap = QPixmap.fromImage(image)

        label = QLabel()

        label.setPixmap(pixmap)

        if page_num < len(self.page_labels):

            self.scrollLayout.replaceWidget(self.page_labels[page_num], label)

            self.page_labels[page_num].deleteLater()

            self.page_labels[page_num] = label

        else:

            self.scrollLayout.addWidget(label)

            self.page_labels.append(label)

 

    def scroll_to_page(self, page_num):

        page_height = self.scrollArea.widget().height() / len(self.doc)

        self.scrollArea.verticalScrollBar().setValue(int(page_height * page_num))

        self.current_page = page_num

        self.update_page_label()

 

    def zoom_in(self):

        try:

            self.zoom_level += 0.1

            self.update_zoom()

        except:

            pass

 

    def zoom_out(self):

        try:

            if self.zoom_level > 0.1:

                self.zoom_level -= 0.1

            self.update_zoom()

        except:

            pass

 

    def update_zoom(self):

        for page_num in range(len(self.page_labels)):

            self.add_page_to_layout(page_num)

        self.update_page_label()

 

    def update_page_label(self):

        self.page_label.setText(f"Page {self.current_page+1} of {len(self.doc)}")

 

    def on_scroll(self):

        scroll_value = self.scrollArea.verticalScrollBar().value()

        try:

            page_height = self.scrollArea.widget().height() / len(self.doc)

            self.current_page = int(scroll_value / page_height)

            self.update_page_label()

        except:

            pass

 

    def close_pdf(self):

        if self.doc:

            self.doc.close()

            self.doc = None

            self.clear_layout(self.scrollLayout)

 

class PDFHighlighter(QMainWindow):

    def __init__(self):

        super().__init__()

        self.initUI()

 

    def initUI(self):

 

        self.setWindowTitle('Smart Key Search')

        self.setGeometry(100, 100, 1200, 900)

 

        self.setWindowIcon(QApplication.style().standardIcon(QApplication.style().SP_FileDialogContentsView))

 

        self.layout = QHBoxLayout()

        self.splitter = QSplitter(Qt.Horizontal)

 

        #--Left Section-------------------------------------------------------------------------------------#

        self.left_panel = QWidget()

        self.left_layout = QVBoxLayout()

       

        #--Search Section-----------------------------------------------------------------------------------#

        self.searchGroupBox = QGroupBox(' Search Directory ')

        self.searchGroupBoxLayout = QVBoxLayout()

 

        self.openFolderButton = QPushButton('Open Folder', self)

        self.openFolderButton.clicked.connect(self.open_folder)

        self.searchGroupBoxLayout.addWidget(self.openFolderButton)

 

        self.keywordInput = QTextEdit(self)

        self.keywordInput.setPlaceholderText('Enter keywords separated by commas')

        self.searchGroupBoxLayout.addWidget(self.keywordInput)

 

        self.searchButton = QPushButton('Search Keywords', self)

        self.searchButton.clicked.connect(self.search_keywords)

        self.searchGroupBoxLayout.addWidget(self.searchButton)

 

        self.searchGroupBox.setLayout(self.searchGroupBoxLayout)

        self.left_layout.addWidget(self.searchGroupBox)

 

        #--File list Section--------------------------------------------------------------------------------#

        self.fileListGroupBox = QGroupBox(' Result File List ')

        self.fileListGroupBoxLayout = QVBoxLayout()

 

        self.fileList = QListWidget(self)

        self.fileList.itemClicked.connect(self.open_selected_pdf)

        self.fileListGroupBoxLayout.addWidget(self.fileList)

       

        self.fileListGroupBox.setLayout(self.fileListGroupBoxLayout)

        self.left_layout.addWidget(self.fileListGroupBox)

        #---------------------------------------------------------------------------------------------------

 

        self.left_panel.setLayout(self.left_layout)

        #---------------------------------------------------------------------------------------------------

 

        #--Right Section------------------------------------------------------------------------------------#

        self.right_panel = QWidget()

        self.right_layout = QVBoxLayout()

       

        #--Keyword List-------------------------------------------------------------------------------------#

        self.keyWordListGroupBox = QGroupBox(' Keyword List ')

        self.keyWordListGroupBoxLayout = QVBoxLayout()

 

        self.keywordList = QListWidget(self)

        self.keyWordListGroupBoxLayout.addWidget(self.keywordList)

        self.keywordList.itemClicked.connect(self.on_keyword_selected)

 

        self.up_down_button_layout = QHBoxLayout()

 

        self.upButton = QPushButton('Up', self)

        self.upButton.clicked.connect(self.scroll_keyword_up)

        self.up_down_button_layout.addWidget(self.upButton)

 

        self.downButton = QPushButton('Down', self)

        self.downButton.clicked.connect(self.scroll_keyword_down)

        self.up_down_button_layout.addWidget(self.downButton)

 

        self.keyWordListGroupBoxLayout.addLayout(self.up_down_button_layout)

 

        self.keyWordListGroupBox.setLayout(self.keyWordListGroupBoxLayout)

        self.right_layout.addWidget(self.keyWordListGroupBox)

        #---------------------------------------------------------------------------------------------------

 

        #--Page Line List----------------------------------------------------------------------------------#

        self.pageLineListGroupBox = QGroupBox(' Position Details ')

        self.pageLineListGroupBoxLayout = QVBoxLayout()

 

        self.page_line_list = QListWidget(self)

        self.page_line_list.itemClicked.connect(self.navigate_to_page_line)

        self.pageLineListGroupBoxLayout.addWidget(self.page_line_list)

 

        self.pageLineListGroupBox.setLayout(self.pageLineListGroupBoxLayout)

        self.right_layout.addWidget(self.pageLineListGroupBox)

 

        #---------------------------------------------------------------------------------------------------

        self.right_panel.setLayout(self.right_layout)

        #---------------------------------------------------------------------------------------------------

 

        self.pdf_viewer = PDFViewer()

 

        self.splitter.addWidget(self.left_panel)

        self.splitter.addWidget(self.pdf_viewer)

        self.splitter.addWidget(self.right_panel)

 

        # self.splitter.setStretchFactor(0, 1)

        # self.splitter.setStretchFactor(1, 3)

        # self.splitter.setStretchFactor(2, 1)

 

        self.splitter.setSizes([300, 800, 300])

 

        self.layout.addWidget(self.splitter)

 

        container = QWidget()

        container.setLayout(self.layout)

        self.setCentralWidget(container)

 

        self.temp_files = []

        self.page_positions = {}

        self.keyword_positions = {}

 

        self.apply_stylesheet()

 

    def apply_stylesheet(self):

        self.setStyleSheet("""

            QWidget {

                background-color: #F0F0F0;

                color: #333;

                font-family: Arial, sans-serif;

            }

            QPushButton {

                background-color: #E0E0E0;

                border: 1px solid #999;

                padding: 5px;

                border-radius: 5px;

            }

            QPushButton#searchButton {

                background-color: #D32F2F;

                color: white;

            }

            QListWidget::item:selected {

                background-color: #D32F2F;

                color: white;

                border: 1px solid #D32F2F;

            }

            QPushButton:hover {

                background-color: #C0C0C0;

            }

            QTextEdit, QListWidget {

                background-color: #FFFFFF;

                border: 1px solid #999;

                padding: 5px;

                border-radius: 5px;

            }

            QGroupBox {

                border: 1px solid #999;

                padding-top: 10px;

                border-radius: 5px;

            }

            QSplitter::handle {

                background-color: #858484;

                width: 1px;

                height: 1px;

                border-radius: 1px;

                ;

            }

        """)

 

    def open_folder(self):

        options = QFileDialog.Options()

        self.folder_path = QFileDialog.getExistingDirectory(self, 'Open Folder', '', options=options)

        self.keywordInput.setPlaceholderText('Enter keywords separated by commas')

        if self.folder_path:

            self.keywordInput.setPlaceholderText(self.folder_path)

            self.keywordList.clear()

            self.page_line_list.clear()

            self.fileList.clear()

            self.pdf_viewer.close_pdf()

            for temp_file in self.temp_files:

                try:

                    os.remove(temp_file)

                    self.temp_files = []

                except:

                    pass

 

    def convert_txt_to_pdf(self, txt_path):

        pdf_path = txt_path.rsplit('.', 1)[0] + '.pdf'

       

        with open(txt_path, 'r', encoding='utf-8') as file:

            content = file.read()

 

        pdf = SimpleDocTemplate(pdf_path, pagesize=letter)

        styles = getSampleStyleSheet()

        story = [Paragraph(content.replace('\n', '<br/>'), styles['Normal'])]

        pdf.build(story)

 

        self.temp_files.append(pdf_path)

        return pdf_path

 

    def search_keywords(self):

        self.keywordList.clear()

        self.page_line_list.clear()

        self.pdf_viewer.close_pdf()

        for temp_file in self.temp_files:

            try:

                os.remove(temp_file)

                self.temp_files = []

            except:

                pass

 

        if hasattr(self, 'folder_path'):

            keywords = self.keywordInput.toPlainText().split(',')

            keywords = [k.strip() for k in keywords]

            self.fileList.clear()

 

            files_to_search = []

            for root, _, files in os.walk(self.folder_path):

                for file in files:

                    if file.lower().endswith(('.pdf','.docx','.txt')):

                        # pdf_path = os.path.join(root, file)

                        # try:

                        #     if self.pdf_contains_keywords(pdf_path, keywords):

                        #         self.fileList.addItem(os.path.basename(pdf_path))

                        # except:

                        #     pass

                        files_to_search.append(os.path.join(root, file))

 

                with concurrent.futures.ThreadPoolExecutor() as executor:

                    future_to_file = {executor.submit(self.pdf_contains_keywords, file, keywords): file for file in files_to_search}

 

                    for future in concurrent.futures.as_completed(future_to_file):

                        file = future_to_file[future]

                        try:

                            if future.result():

                                self.fileList.addItem(os.path.basename(file))

                        except:

                            pass

 

    def pdf_contains_keywords(self, pdf_path, keywords):

        file_extension = os.path.splitext(pdf_path)[1].lower()

 

        if file_extension == '.txt':

            pdf_path = self.convert_txt_to_pdf(pdf_path)

 

        document = fitz.open(pdf_path)

        for keyword in keywords:

            keyword_found = False

            for page_num in range(len(document)):

                page = document.load_page(page_num)

                if page.search_for(keyword):

                    keyword_found = True

                    break

            if not keyword_found:

                document.close()

                return False

        document.close()

        return True

 

    def open_selected_pdf(self, item):

        self.page_line_list.clear()

 

        keywords = self.keywordInput.toPlainText().split(',')

        keywords = [k.strip() for k in keywords]

 

        pdf_path = self.folder_path + '/' + item.text()

        pdf_or_doc = item.text().split('.')[len(item.text().split('.'))-1]

 

        if pdf_or_doc == 'txt':

            pdf_path = self.convert_txt_to_pdf(pdf_path)

 

        if pdf_or_doc in ['pdf','txt']:

            highlighted_pdf_path, highlights = self.highlight_keywords_in_pdf(pdf_path, keywords)

        elif pdf_or_doc == 'docx':

            highlighted_pdf_path = self.highlight_keywords_in_docx(pdf_path, keywords)

 

        if highlighted_pdf_path:

            current_page = self.page_positions.get(pdf_path, 0)

            self.pdf_viewer.display_pdf(highlighted_pdf_path, current_page)

            self.keyword_positions = self.get_keyword_positions(highlighted_pdf_path, keywords)

            self.populate_keyword_list(keywords)

 

    def highlight_keywords_in_pdf(self, pdf_path, keywords):

        base, ext = os.path.splitext(pdf_path)

        existing_highlighted_pdf_path = f"{base}_highlighted_{'_'.join(keywords)}{ext}"

        if os.path.exists(existing_highlighted_pdf_path):

            return existing_highlighted_pdf_path, []

 

        document = fitz.open(pdf_path)

        highlight_colors = [fitz.utils.getColor("yellow"),

                            fitz.utils.getColor("cyan"),

                            fitz.utils.getColor("magenta"),

                            fitz.utils.getColor("red"),

                            fitz.utils.getColor("blue"),

                            fitz.utils.getColor("pink"),

                            fitz.utils.getColor("orange"),

                            fitz.utils.getColor("green")]

 

        highlights = []

        for page_num in range(len(document)):

            page = document.load_page(page_num)

            for i, keyword in enumerate(keywords):

                text_instances = page.search_for(keyword, quads = True)

                for inst in text_instances:

                    highlight = page.add_highlight_annot(inst)

                    highlight.set_colors({"stroke": highlight_colors[i % len(highlight_colors)]})

                    highlight.update()

                    highlights.append((page_num, inst))

 

        document.save(existing_highlighted_pdf_path)

        document.close()

        self.temp_files.append(existing_highlighted_pdf_path)

        return existing_highlighted_pdf_path, highlights

 

    def highlight_keywords_in_docx(self, docx_path, keywords):

        doc = Document(docx_path)

       

        # Create an in-memory PDF

        buffer = io.BytesIO()

        pdf = SimpleDocTemplate(buffer, pagesize=letter)

        styles = getSampleStyleSheet()

        story = []

        colors = ['yellow', 'cyan', 'magenta', 'red', 'blue', 'pink', 'orange', 'green']

 

        for paragraph in doc.paragraphs:

            words = paragraph.text.split()

            highlighted_words = []

 

            for word in words:

                found_keyword = False

                for i, keyword in enumerate(keywords):

                    if keyword.lower() in word.lower():

                        highlighted_word = f'<span bgcolor="{colors[i % len(colors)]}">{word}</span>'

                        highlighted_words.append(highlighted_word)

                        found_keyword = True

                        break

                if not found_keyword:

                    highlighted_words.append(word)

 

            highlighted_text = Paragraph(' '.join(highlighted_words), styles['Normal'])

            story.append(highlighted_text)

            story.append(Spacer(1, 12))

 

        pdf.build(story)

 

        # Save the highlighted PDF

        base, _ = os.path.splitext(docx_path)

        highlighted_pdf_path = f"{base}_highlighted_{'_'.join(keywords)}.pdf"

        with open(highlighted_pdf_path, 'wb') as f:

            f.write(buffer.getvalue())

 

        self.temp_files.append(highlighted_pdf_path)

        return highlighted_pdf_path

 

    def get_keyword_positions(self, pdf_path, keywords):

        document = fitz.open(pdf_path)

        keyword_positions = {keyword: [] for keyword in keywords}

 

        for page_num in range(len(document)):

            page = document.load_page(page_num)

            blocks = page.get_text("dict")["blocks"]

 

            lines = []

            line_number = 1

 

            for block in blocks:

                for line in block.get("lines", []):

                    line_bbox = line["bbox"]

                    line_y0 = line_bbox[1]

                    line_y1 = line_bbox[3]

                    line_text = " ".join([span["text"] for span in line["spans"] if span["text"].strip() != ""])

 

                    if line_text.strip():

                        lines.append((line_number, line_y0, line_y1, line_text))

                        line_number += 1

 

            for keyword in keywords:

                text_instances = page.search_for(keyword)

                for inst in text_instances:

                    y0 = inst.y0

                    y1 = inst.y1

                    line_text = None

 

                    for line_num, line_y0, line_y1, text in lines:

                        if abs(line_y0 - y0) < 2 and abs(line_y1 - y1) < 2:

                            line_text = text

                            break

 

                    if line_text:

                        snippet = line_text[:20] + "..."

                        keyword_positions[keyword].append((page_num, line_num, snippet))

 

        document.close()

        return keyword_positions

 

    def populate_keyword_list(self, keywords):

        self.keywordList.clear()

        colors = [QColor("yellow"), QColor("cyan"), QColor("magenta"), QColor("red"), QColor("blue"), QColor("pink"), QColor("orange"), QColor("green")]

 

        for i, keyword in enumerate(keywords):

            item = QListWidgetItem(keyword)

            item.setBackground(colors[i % len(colors)])

            self.keywordList.addItem(item)

 

    def on_keyword_selected(self, item):

        selected_keyword = item.text()

        self.populate_page_line_list(selected_keyword)

 

    def scroll_keyword_up(self):

        selected_items = self.keywordList.selectedItems()

        if not selected_items:

            return

        selected_keyword = selected_items[0].text()

        if selected_keyword in self.keyword_positions:

            current_page = self.pdf_viewer.current_page

            pages = [pos[0] for pos in self.keyword_positions[selected_keyword]]

            previous_pages = [p for p in pages if p < current_page]

            if previous_pages:

                self.pdf_viewer.scroll_to_page(previous_pages[-1])

            else:

                self.pdf_viewer.scroll_to_page(pages[-1])

            self.populate_page_line_list(selected_keyword)

 

    def scroll_keyword_down(self):

        selected_items = self.keywordList.selectedItems()

        if not selected_items:

            return

        selected_keyword = selected_items[0].text()

        if selected_keyword in self.keyword_positions:

            current_page = self.pdf_viewer.current_page

            pages = [pos[0] for pos in self.keyword_positions[selected_keyword]]

            next_pages = [p for p in pages if p > current_page]

            if next_pages:

                self.pdf_viewer.scroll_to_page(next_pages[0])

            else:

                self.pdf_viewer.scroll_to_page(pages[0])

            self.populate_page_line_list(selected_keyword)

 

    def populate_page_line_list(self, keyword):

        self.page_line_list.clear()

        if keyword in self.keyword_positions:

            for page_num, y_pos, snippet in self.keyword_positions[keyword]:

                item = QListWidgetItem(f"Page {page_num + 1}: Line {int(y_pos)} | {snippet}")

                item.setData(Qt.UserRole, (page_num, y_pos))

                self.page_line_list.addItem(item)

 

    def navigate_to_page_line(self, item):

        page_num, y_pos = item.data(Qt.UserRole)

        self.pdf_viewer.scroll_to_page(page_num)

        scroll_value = self.pdf_viewer.scrollArea.verticalScrollBar().value() + y_pos + 2

        self.pdf_viewer.scrollArea.verticalScrollBar().setValue(scroll_value)

 

    def closeEvent(self, event):

        self.pdf_viewer.close_pdf()

        for temp_file in self.temp_files:

            try:

                os.remove(temp_file)

            except:

                pass

        event.accept()

 

if __name__ == '__main__':

    app = QApplication(sys.argv)

    ex = PDFHighlighter()

    ex.show()

    sys.exit(app.exec_())

 

 


