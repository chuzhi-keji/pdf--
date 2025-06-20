#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# 作者: AI Assistant
# 日期: 2025-06-20
# 版本: 1.0
# 描述: 一个使用PyQt6和PyMuPDF实现的PDF转图片工具

import sys
import os
import fitz  # PyMuPDF

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QListWidget, QComboBox,
    QRadioButton, QGroupBox, QLineEdit, QMessageBox, QProgressBar
)
from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtGui import QDragEnterEvent, QDropEvent  # QIcon (if you want an icon)

# --- 全局常量与配置 ---
APP_TITLE = "PDF 转图片工具"
DEFAULT_RESOLUTION = 200  # DPI
SUPPORTED_FORMATS = ["png", "jpg"]
SOFTWARE_INFO = "Powered by PyQt6 & PyMuPDF | Author: AI Assistant | Version 1.0"

# 保存选项常量
SAVE_OPTION_SOURCE = "源文件位置"
SAVE_OPTION_SPECIFIED = "指定文件夹"
SAVE_OPTION_SUBFOLDER = "源文件位置创建子文件夹"


class PdfConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf_files = []  # 存储待转换的PDF文件路径
        self.output_directory = ""  # 存储用户指定的输出目录
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(APP_TITLE)
        self.setGeometry(100, 100, 650, 550)  # x, y, width, height - slightly wider
        self.setAcceptDrops(True)  # 允许窗口接收拖拽事件

        # --- 主布局 ---
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)  # Add some spacing between widgets
        main_layout.setContentsMargins(15, 15, 15, 15)  # Add padding around the main layout

        # --- 1. 文件列表区域 ---
        self.file_list_widget = QListWidget()
        self.file_list_widget.setToolTip("将PDF文件拖拽到此处，或点击下方按钮添加")
        self.file_list_widget.setStyleSheet("QListWidget { border: 1px solid #ccc; border-radius: 4px; }")
        main_layout.addWidget(QLabel("待转换 PDF 文件列表:"))
        main_layout.addWidget(self.file_list_widget, 1)  # 占据更多垂直空间

        # --- 文件操作按钮 ---
        file_button_layout = QHBoxLayout()
        add_files_button = QPushButton("添加文件...")
        add_files_button.clicked.connect(self.add_files_dialog)
        clear_files_button = QPushButton("清空列表")
        clear_files_button.clicked.connect(self.clear_files)
        file_button_layout.addWidget(add_files_button)
        file_button_layout.addWidget(clear_files_button)
        file_button_layout.addStretch()
        main_layout.addLayout(file_button_layout)

        # --- 2. 转换参数配置区域 ---
        settings_group = QGroupBox("转换设置")
        settings_layout = QVBoxLayout()
        settings_layout.setSpacing(8)

        # 分辨率选择
        res_layout = QHBoxLayout()
        res_layout.addWidget(QLabel("图片分辨率 (DPI):"))
        self.resolution_combo = QComboBox()
        self.resolution_combo.addItems(["72", "96", "150", "200", "300", "400", "600"])
        self.resolution_combo.setCurrentText(str(DEFAULT_RESOLUTION))
        res_layout.addWidget(self.resolution_combo)
        res_layout.addStretch()
        settings_layout.addLayout(res_layout)

        # 图片格式选择
        format_layout = QHBoxLayout()
        format_layout.addWidget(QLabel("图片格式:"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(SUPPORTED_FORMATS)
        format_layout.addWidget(self.format_combo)
        format_layout.addStretch()
        settings_layout.addLayout(format_layout)

        settings_group.setLayout(settings_layout)
        main_layout.addWidget(settings_group)

        # --- 3. 保存位置选项区域 ---
        save_options_group = QGroupBox("保存选项")
        save_options_layout = QVBoxLayout()
        save_options_layout.setSpacing(8)

        self.radio_save_source = QRadioButton(SAVE_OPTION_SOURCE)
        self.radio_save_source.setChecked(True)
        save_options_layout.addWidget(self.radio_save_source)

        self.radio_save_subfolder = QRadioButton(SAVE_OPTION_SUBFOLDER)
        save_options_layout.addWidget(self.radio_save_subfolder)

        specified_dir_layout = QHBoxLayout()
        self.radio_save_specified = QRadioButton(SAVE_OPTION_SPECIFIED + ":")  # Add colon for clarity
        specified_dir_layout.addWidget(self.radio_save_specified)

        self.specified_dir_edit = QLineEdit()
        self.specified_dir_edit.setPlaceholderText("选择或输入文件夹路径")
        self.specified_dir_edit.setEnabled(False)
        self.specified_dir_edit.setStyleSheet("QLineEdit[enabled=\"false\"] { background-color: #f0f0f0; }")

        browse_button = QPushButton("浏览...")
        browse_button.clicked.connect(self.browse_output_directory)
        browse_button.setEnabled(False)

        specified_dir_layout.addWidget(self.specified_dir_edit, 1)
        specified_dir_layout.addWidget(browse_button)
        save_options_layout.addLayout(specified_dir_layout)

        self.radio_save_specified.toggled.connect(self.specified_dir_edit.setEnabled)
        self.radio_save_specified.toggled.connect(browse_button.setEnabled)
        # Ensure output_directory is cleared if this option is deselected
        self.radio_save_specified.toggled.connect(lambda checked: self.output_directory_toggled(checked))

        save_options_group.setLayout(save_options_layout)
        main_layout.addWidget(save_options_group)

        # --- 4. 开始转换按钮 ---
        self.convert_button = QPushButton("开始转换")
        self.convert_button.setStyleSheet(
            "QPushButton { background-color: #28a745; color: white; padding: 10px 15px; font-size: 16px; border-radius: 5px; }"
            "QPushButton:hover { background-color: #218838; }"
            "QPushButton:pressed { background-color: #1e7e34; }"
        )
        self.convert_button.clicked.connect(self.start_conversion)
        main_layout.addWidget(self.convert_button)

        # --- 5. 进度条 ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%v / %m 文件")  # Show current/total
        main_layout.addWidget(self.progress_bar)

        # --- 6. 软件说明行 ---
        self.info_label = QLabel(SOFTWARE_INFO)
        self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.info_label.setStyleSheet("QLabel { margin-top: 10px; font-style: italic; color: #555; font-size: 0.9em; }")
        main_layout.addWidget(self.info_label)

    def output_directory_toggled(self, checked):
        if not checked:
            self.output_directory = ""  # Clear if not selected
            # self.specified_dir_edit.clear() # Optionally clear text field too

    def add_files_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择 PDF 文件", self.output_directory, "PDF Files (*.pdf)")
        if files:
            for file_path in files:
                if file_path not in self.pdf_files:
                    self.pdf_files.append(file_path)
                    self.file_list_widget.addItem(os.path.basename(file_path))
            self.update_progress_bar_max()

    def clear_files(self):
        self.pdf_files.clear()
        self.file_list_widget.clear()
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(0)  # Reset max
        self.progress_bar.setFormat("%v / %m 文件")

    def browse_output_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "选择保存文件夹",
                                                     self.output_directory or os.path.expanduser("~"))
        if directory:
            self.output_directory = directory
            self.specified_dir_edit.setText(directory)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            all_pdfs = True
            for url in event.mimeData().urls():
                if not (url.isLocalFile() and url.toLocalFile().lower().endswith('.pdf')):
                    all_pdfs = False
                    break
            if all_pdfs:
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        added_count = 0
        for url in urls:
            if url.isLocalFile():
                file_path = url.toLocalFile()
                if file_path.lower().endswith('.pdf') and file_path not in self.pdf_files:
                    self.pdf_files.append(file_path)
                    self.file_list_widget.addItem(os.path.basename(file_path))
                    added_count += 1
        if added_count > 0:
            self.update_progress_bar_max()
        event.acceptProposedAction()

    def update_progress_bar_max(self):
        self.progress_bar.setMaximum(len(self.pdf_files))
        self.progress_bar.setValue(0)  # Reset progress value
        self.progress_bar.setFormat("%v / %m 文件")

    def start_conversion(self):
        if not self.pdf_files:
            QMessageBox.warning(self, "提示", "请先添加至少一个 PDF 文件！")
            return

        try:
            resolution = int(self.resolution_combo.currentText())
        except ValueError:
            QMessageBox.critical(self, "错误", "无效的分辨率设置！")
            return

        image_format = self.format_combo.currentText()

        save_option = ""
        if self.radio_save_source.isChecked():
            save_option = SAVE_OPTION_SOURCE
        elif self.radio_save_subfolder.isChecked():
            save_option = SAVE_OPTION_SUBFOLDER
        elif self.radio_save_specified.isChecked():
            save_option = SAVE_OPTION_SPECIFIED
            if not self.output_directory or not os.path.isdir(self.output_directory):
                QMessageBox.warning(self, "提示", "请选择一个有效的指定保存文件夹！")
                self.browse_output_directory()  # Prompt user to select again
                if not self.output_directory or not os.path.isdir(self.output_directory):  # Check again
                    return

        self.convert_button.setEnabled(False)
        self.file_list_widget.setEnabled(False)  # Disable list during conversion
        self.progress_bar.setValue(0)
        # Max is already set by add_files or dropEvent via update_progress_bar_max

        total_files = len(self.pdf_files)
        success_count = 0
        errors = []

        for i, pdf_path in enumerate(self.pdf_files):
            try:
                self.convert_single_pdf(pdf_path, resolution, image_format, save_option)
                success_count += 1
            except Exception as e:
                error_msg = f"处理文件 '{os.path.basename(pdf_path)}' 失败: {str(e)}"
                errors.append(error_msg)
                print(f"Error: {error_msg}")  # Log to console

            self.progress_bar.setValue(i + 1)
            QApplication.processEvents()  # Process UI events to prevent freezing

        self.convert_button.setEnabled(True)
        self.file_list_widget.setEnabled(True)

        if not errors:
            QMessageBox.information(self, "转换完成", f"所有 {total_files} 个 PDF 文件已成功转换为图片！")
        else:
            error_summary = f"{success_count} 个文件成功转换。\n{len(errors)} 个文件转换失败:\n\n" + "\n".join(errors)
            QMessageBox.warning(self, "转换部分完成", error_summary)

        # self.progress_bar.setValue(total_files) # Ensure progress bar is full if all successful
        # Or reset if preferred:
        # self.clear_files() # Optionally clear list after conversion

    def convert_single_pdf(self, pdf_path, resolution, image_format, save_option):
        doc = None  # Initialize to ensure it's defined for finally block
        try:
            doc = fitz.open(pdf_path)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]

            output_dir_final = ""
            source_dir = os.path.dirname(pdf_path)

            if save_option == SAVE_OPTION_SOURCE:
                output_dir_final = source_dir
            elif save_option == SAVE_OPTION_SPECIFIED:
                output_dir_final = self.output_directory
            elif save_option == SAVE_OPTION_SUBFOLDER:
                output_dir_final = os.path.join(source_dir, base_name + "_images")

            if not os.path.exists(output_dir_final):
                os.makedirs(output_dir_final, exist_ok=True)

            num_pages = len(doc)
            for page_num in range(num_pages):
                page = doc.load_page(page_num)
                zoom = resolution / 72.0  # PyMuPDF default DPI is 72
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat,
                                      alpha=False if image_format.lower() == 'jpg' else True)  # JPG doesn't support alpha

                # Pad page number for better sorting if many pages
                page_indicator = f"page_{str(page_num + 1).zfill(len(str(num_pages)))}"
                image_filename = f"{base_name}_{page_indicator}.{image_format}"
                output_image_path = os.path.join(output_dir_final, image_filename)

                if image_format.lower() == "png":
                    pix.save(output_image_path)
                elif image_format.lower() == "jpg":
                    pix.save(output_image_path, jpg_quality=95)  # PyMuPDF default is 92, 95 is common high quality

            print(f"文件 '{os.path.basename(pdf_path)}' 已成功转换为图片并保存至 '{output_dir_final}'")
        finally:
            if doc:
                doc.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # Example for setting an application icon (optional)
    # script_dir = os.path.dirname(os.path.realpath(__file__))
    # icon_path = os.path.join(script_dir, "your_icon.png") # Place your_icon.png in the same directory
    # if os.path.exists(icon_path):
    #    app.setWindowIcon(QIcon(icon_path))

    main_window = PdfConverterApp()
    main_window.show()
    sys.exit(app.exec())