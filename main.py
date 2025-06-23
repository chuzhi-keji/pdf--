#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import pathlib
import tempfile
import shutil
import gc
import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QListWidget, QPushButton, QLabel, QLineEdit, QRadioButton, QButtonGroup,
    QFileDialog, QMessageBox, QStatusBar, QTabWidget, QGroupBox, QComboBox,
    QProgressBar, QMenu, QListWidgetItem, QTextBrowser, QDialog
)
from PyQt6.QtCore import Qt, QUrl, QThread, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QAction
import time


class PdfWorker(QThread):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)  # 新增：进度信号
    paused = pyqtSignal()  # 新增：暂停信号
    resumed = pyqtSignal()  # 新增：恢复信号

    def __init__(self, task_type, *args):
        super().__init__()
        self.task_type = task_type
        self.args = args
        self._is_paused = False  # 新增：暂停状态
        self._should_stop = False  # 新增：停止标志

    def run(self):
        try:
            if self.task_type == "merge":
                result = merge_pdf_documents(*self.args)
                self.finished.emit(result)
            elif self.task_type == "split":
                results = split_pdf_document(*self.args)
                self.finished.emit(results)
            elif self.task_type == "convert":
                # 新增：支持暂停功能
                results = convert_pdf_to_images(self, *self.args)
                self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))

    # 新增：暂停方法
    def pause(self):
        self._is_paused = True
        self.paused.emit()

    # 新增：恢复方法
    def resume(self):
        self._is_paused = False
        self.resumed.emit()

    # 新增：停止方法
    def stop(self):
        self._should_stop = True
        if self._is_paused:
            self.resume()  # 确保线程能退出

    # 新增：检查暂停状态
    def check_paused(self):
        while self._is_paused:
            time.sleep(0.1)  # 短暂睡眠减少CPU占用
            if self._should_stop:
                break
        return self._should_stop


def create_output_path(source_file_for_ref_dir: str, output_config: dict, target_filename: str) -> str:
    """
    根据输出配置计算并创建目标文件的完整路径。
    :param source_file_for_ref_dir: 用于确定源目录的参考文件路径。
    :param output_config: {"mode": "source_dir" | "new_subdir_in_source" | "custom_dir",
                           "path": "自定义路径 (仅custom_dir模式)",
                           "subfolder_name": "子文件夹名 (仅new_subdir_in_source模式)"}
    :param target_filename: 目标文件名 (带扩展名)。
    :return: 完整的目标文件路径。
    :raises: ValueError 如果配置无效或路径创建失败。
    """
    mode = output_config.get("mode")

    if mode == "source_dir":
        base_dir = os.path.dirname(source_file_for_ref_dir)
        final_path = pathlib.Path(base_dir) / target_filename
    elif mode == "new_subdir_in_source":
        base_dir = os.path.dirname(source_file_for_ref_dir)
        subfolder = output_config.get("subfolder_name", "output_files").strip()
        if not subfolder:  # 防止用户输入空或纯空格的子文件夹名
            subfolder = "output_files"
        output_subdir = pathlib.Path(base_dir) / subfolder
        try:
            output_subdir.mkdir(parents=True, exist_ok=True)  # 创建子文件夹
        except OSError as e:
            raise ValueError(f"创建子文件夹 '{output_subdir}' 失败: {e}")
        final_path = output_subdir / target_filename
    elif mode == "custom_dir":
        custom_path_str = output_config.get("path")
        if not custom_path_str or not os.path.isdir(custom_path_str):
            raise ValueError(f"自定义路径 '{custom_path_str}' 无效或不存在。")
        final_path = pathlib.Path(custom_path_str) / target_filename
    else:
        raise ValueError(f"未知的输出模式: {mode}")

    return str(final_path.resolve())  # 返回绝对路径


def is_valid_pdf(file_path: str) -> bool:
    """检查文件是否为有效的PDF"""
    try:
        with open(file_path, 'rb') as f:
            PdfReader(f)
        return True
    except Exception as e:
        print(f"无效的PDF文件: {file_path}, 错误: {str(e)}")
        return False


def _parse_page_ranges(ranges_str: str, total_pages: int) -> list:
    """
    辅助函数：解析页码范围字符串。
    例如 '1-3,5; 7-end' 会被解析成 [[1,2,3,5], [7,8,...,total_pages]]
    其中分号分隔不同的输出文件，逗号分隔同一文件内的不同范围。
    页码是1-based。
    """
    if not ranges_str:
        return []

    output_file_segments = []
    segments_str_list = ranges_str.split(';')  # 按分号分割成不同的输出文件段

    for segment_str in segments_str_list:
        current_segment_pages = set()  # 使用set避免重复页码
        parts = segment_str.strip().split(',')
        for part in parts:
            part = part.strip().lower()
            if not part:
                continue

            part_resolved = part.replace("end", str(total_pages))

            try:
                if '-' in part_resolved:
                    start, end = map(int, part_resolved.split('-'))
                    if start <= 0 or end < start or end > total_pages:
                        # 可以选择抛出异常或记录警告并跳过此范围
                        print(f"警告: 无效范围 '{part}' (总页数: {total_pages})，已跳过。")
                        continue
                    current_segment_pages.update(range(start, end + 1))
                else:
                    page_num = int(part_resolved)
                    if page_num <= 0 or page_num > total_pages:
                        print(f"警告: 无效页码 '{part}' (总页数: {total_pages})，已跳过。")
                        continue
                    current_segment_pages.add(page_num)
            except ValueError:
                print(f"警告: 无法解析页码或范围 '{part}'，已跳过。")
                continue  # 跳过无法解析的部分

        if current_segment_pages:
            output_file_segments.append(sorted(list(current_segment_pages)))

    return output_file_segments


def _write_pdf_with_tempfile(writer, output_path: str) -> bool:
    """
    使用临时文件安全写入PDF
    """
    temp_file = None
    temp_path = None
    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            temp_path = temp_file.name
            writer.write(temp_file)

        # 移动临时文件到目标位置
        shutil.move(temp_path, output_path)
        return True
    except Exception as e:
        # 清理临时文件
        if temp_path and os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except:
                pass
        raise e


def split_pdf_document(source_pdf_path: str, output_dir_config: dict, split_options: dict) -> list:
    """
    拆分指定的PDF文件。
    :param source_pdf_path: 原始PDF文件路径。
    :param output_dir_config: 输出配置字典。
    :param split_options: 拆分参数字典 {"type": "range" | "all_pages_separately",
                                       "ranges_str": "页码范围字符串 (e.g., '1-3, 5; 7-end')"}
    :return: 包含拆分结果信息的列表。
    """
    results = []

    # 验证PDF文件
    if not is_valid_pdf(source_pdf_path):
        results.append({
            "status": "failure",
            "file_path": None,
            "error_message": f"无效的PDF文件: {os.path.basename(source_pdf_path)}"
        })
        return results

    try:
        with open(source_pdf_path, 'rb') as src_file:
            reader = PdfReader(src_file)
            total_pages_in_doc = len(reader.pages)

            if split_options["type"] == "all_pages_separately":
                for i in range(total_pages_in_doc):
                    writer = PdfWriter()
                    writer.add_page(reader.pages[i])
                    page_num_one_based = i + 1
                    base_name, ext = os.path.splitext(os.path.basename(source_pdf_path))
                    output_filename = f"{base_name}_page_{page_num_one_based}{ext}"

                    try:
                        full_output_path = create_output_path(source_pdf_path, output_dir_config, output_filename)
                        _write_pdf_with_tempfile(writer, full_output_path)
                        results.append({"status": "success", "file_path": full_output_path, "error_message": None})
                    except Exception as e:
                        results.append({"status": "failure", "file_path": output_filename, "error_message": str(e)})

                    # 定期清理内存
                    if (i + 1) % 10 == 0:
                        writer = None
                        import gc
                        gc.collect()

            elif split_options["type"] == "range":
                ranges_str = split_options.get("ranges_str", "")
                parsed_page_segments = _parse_page_ranges(ranges_str, total_pages_in_doc)

                if not parsed_page_segments and ranges_str:  # 有输入但解析为空，说明输入格式可能有问题
                    results.append({"status": "failure", "file_path": None,
                                    "error_message": f"无法解析页码范围: '{ranges_str}'。请检查格式。"})
                    return results
                if not parsed_page_segments:  # 没有有效范围
                    results.append({"status": "failure", "file_path": None, "error_message": "未提供有效的拆分范围。"})
                    return results

                for idx, page_selection in enumerate(parsed_page_segments):
                    if not page_selection:  # 当前段没有有效页面
                        continue
                    writer = PdfWriter()
                    for page_num_one_based in page_selection:
                        # _parse_page_ranges 已经处理了页码有效性，这里直接使用
                        writer.add_page(reader.pages[page_num_one_based - 1])

                    if len(writer.pages) > 0:  # 确保有页面被添加
                        base_name, ext = os.path.splitext(os.path.basename(source_pdf_path))
                        # 文件名可以根据范围或序号生成
                        range_desc = f"{page_selection[0]}-{page_selection[-1]}" if len(page_selection) > 1 else str(
                            page_selection[0])
                        output_filename_suffix = f"_split_{idx + 1}_pages_{range_desc}" if len(
                            parsed_page_segments) > 1 else f"_pages_{range_desc}"
                        output_filename = f"{base_name}{output_filename_suffix}{ext}"
                        try:
                            full_output_path = create_output_path(source_pdf_path, output_dir_config, output_filename)
                            _write_pdf_with_tempfile(writer, full_output_path)
                            results.append({"status": "success", "file_path": full_output_path, "error_message": None})
                        except Exception as e:
                            results.append({"status": "failure", "file_path": output_filename, "error_message": str(e)})

                    # 定期清理内存
                    if (idx + 1) % 10 == 0:
                        writer = None
                        import gc
                        gc.collect()
            else:
                results.append({"status": "failure", "file_path": None, "error_message": "未知的拆分类型"})

    except Exception as e:
        # 捕获所有异常并记录
        results.append({"status": "failure", "file_path": None,
                        "error_message": f"处理 '{os.path.basename(source_pdf_path)}' 时发生错误: {str(e)}"})
    return results


def merge_pdf_documents(input_pdf_paths: list[str], output_dir_config: dict, merged_filename: str) -> dict:
    """
    合并多个PDF文件到一个文件。
    :param input_pdf_paths: 待合并PDF文件路径列表。
    :param output_dir_config: 输出配置字典。
    :param merged_filename: 合并后文件名 (不含路径), e.g., "merged_output.pdf"。
    :return: 包含合并结果信息的字典。
    """
    writer = PdfWriter()  # 使用 PdfWriter 替代 PdfMerger
    temp_file = None
    temp_path = None

    try:
        # 验证输入文件
        invalid_files = [path for path in input_pdf_paths if not is_valid_pdf(path)]
        if invalid_files:
            invalid_names = ", ".join(os.path.basename(path) for path in invalid_files)
            return {
                "status": "failure",
                "file_path": None,
                "error_message": f"无效的PDF文件: {invalid_names}"
            }

        if not input_pdf_paths:
            return {"status": "failure", "file_path": None, "error_message": "没有可合并的文件。"}

        # 添加文件
        for pdf_path in input_pdf_paths:
            try:
                with open(pdf_path, 'rb') as f:
                    reader = PdfReader(f)
                    # 将整个文档的所有页面添加到 writer
                    for page in reader.pages:
                        writer.add_page(page)
            except Exception as e:
                return {
                    "status": "failure",
                    "file_path": None,
                    "error_message": f"添加文件 '{os.path.basename(pdf_path)}' 失败: {str(e)}"
                }

        reference_source_path = input_pdf_paths[0]
        full_output_path = create_output_path(reference_source_path, output_dir_config, merged_filename)

        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            temp_path = temp_file.name
            writer.write(temp_file)

        # 移动临时文件到目标位置
        shutil.move(temp_path, full_output_path)

        return {
            "status": "success",
            "file_path": full_output_path,
            "error_message": None
        }

    except Exception as e:
        return {
            "status": "failure",
            "file_path": None,
            "error_message": f"合并PDF时发生错误: {str(e)}"
        }
    finally:
        try:
            # 关闭 writer 释放资源
            if hasattr(writer, 'close'):
                writer.close()
            # 清理临时文件（如果移动失败）
            if temp_path and os.path.exists(temp_path):
                os.unlink(temp_path)
        except:
            pass


# 修改：添加worker参数以支持暂停功能
def convert_pdf_to_images(worker: PdfWorker, pdf_files: list, resolution: int, image_format: str, save_option: str,
                          output_dir: str = "", page_range_str: str = "") -> list:
    """
    将PDF转换为图片
    :param worker: PDF工作线程对象（用于暂停控制）
    :param pdf_files: PDF文件路径列表
    :param resolution: 分辨率(DPI)
    :param image_format: 图片格式(png/jpg)
    :param save_option: 保存选项(source/specified/subfolder)
    :param output_dir: 指定目录(仅当save_option为specified时有效)
    :param page_range_str: 页码范围字符串(可选)
    :return: 包含转换结果的列表
    """
    results = []
    total_files = len(pdf_files)

    for file_idx, pdf_path in enumerate(pdf_files):
        # 检查是否应该停止
        if worker and worker._should_stop:
            results.append({
                "status": "cancelled",
                "file_path": pdf_path,
                "error_message": "用户取消操作"
            })
            break

        try:
            doc = fitz.open(pdf_path)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            source_dir = os.path.dirname(pdf_path)
            total_pages = len(doc)

            # 解析页码范围
            page_selection = []
            if page_range_str:
                parsed_segments = _parse_page_ranges(page_range_str, total_pages)
                # 合并所有段
                for segment in parsed_segments:
                    page_selection.extend(segment)
            else:
                # 如果没有指定范围，转换所有页面
                page_selection = list(range(1, total_pages + 1))

            # 去重并排序
            page_selection = sorted(set(page_selection))

            # 检查是否有有效页面
            if not page_selection:
                results.append({
                    "status": "failure",
                    "file_path": pdf_path,
                    "error_message": "没有有效的页面可供转换"
                })
                continue

            # 确定输出目录
            if save_option == "source_dir":
                output_dir_final = source_dir
            elif save_option == "custom_dir":
                output_dir_final = output_dir
            elif save_option == "new_subdir_in_source":
                output_dir_final = os.path.join(source_dir, base_name + "_images")

            # 创建输出目录
            if not os.path.exists(output_dir_final):
                os.makedirs(output_dir_final, exist_ok=True)

            num_pages_to_convert = len(page_selection)
            for idx, page_num in enumerate(page_selection):
                # 检查是否应该暂停
                if worker and worker.check_paused():
                    results.append({
                        "status": "cancelled",
                        "file_path": pdf_path,
                        "error_message": "用户取消操作"
                    })
                    doc.close()
                    return results

                # 确保页码在有效范围内
                if page_num < 1 or page_num > total_pages:
                    continue

                # 获取页面索引 (0-based)
                page_index = page_num - 1
                page = doc.load_page(page_index)
                zoom = resolution / 72.0  # PyMuPDF默认DPI是72
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False if image_format.lower() == 'jpg' else True)

                # 格式化页码
                page_indicator = f"page_{str(page_num).zfill(len(str(total_pages)))}"
                image_filename = f"{base_name}_{page_indicator}.{image_format}"
                output_image_path = os.path.join(output_dir_final, image_filename)

                # 保存图片
                if image_format.lower() == "png":
                    pix.save(output_image_path)
                elif image_format.lower() == "jpg":
                    pix.save(output_image_path, jpg_quality=95)

                # 更新进度 (每页)
                if worker:
                    # 计算总体进度：当前文件进度 + 已完成的文件
                    progress = int(((file_idx + (idx + 1) / num_pages_to_convert) / total_files) * 100)
                    worker.progress.emit(progress)

            results.append({
                "status": "success",
                "file_path": pdf_path,
                "output_dir": output_dir_final,
                "pages": num_pages_to_convert,
                "error_message": None
            })
        except Exception as e:
            results.append({
                "status": "failure",
                "file_path": pdf_path,
                "error_message": str(e)
            })
        finally:
            if 'doc' in locals() and doc:
                doc.close()

        # 更新文件级别进度
        if worker:
            progress = int(((file_idx + 1) / total_files) * 100)
            worker.progress.emit(progress)

    return results


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于 PDF 工具箱")
        self.setFixedSize(600, 500)

        layout = QVBoxLayout(self)

        title_label = QLabel("PDF 工具箱 v1.0")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        text_browser = QTextBrowser()
        text_browser.setOpenExternalLinks(True)
        text_browser.setHtml(self.get_about_html())
        text_browser.setStyleSheet("font-size: 12px;")

        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.close)
        close_btn.setFixedWidth(100)

        layout.addWidget(title_label)
        layout.addWidget(text_browser, 1)
        layout.addWidget(close_btn, 0, Qt.AlignmentFlag.AlignCenter)

    def get_about_html(self):
        return """
        <h2>功能说明</h2>
        <p><b>1. 合并PDF</b>：将多个PDF文件合并成一个PDF文档</p>
        <p><b>2. 拆分PDF</b>：</p>
        <ul>
            <li>按指定范围拆分（例如：1-3,5,7-end）</li>
            <li>每页拆分为单独文件</li>
        </ul>
        <p><b>3. PDF转图片</b>：</p>
        <ul>
            <li>支持PNG和JPG格式</li>
            <li>可调整分辨率（72-600 DPI）</li>
            <li>支持单页和多页转换</li>
            <li>新增暂停/继续功能</li>
            <li>新增页码范围选择功能</li>
        </ul>

        <h2>使用说明</h2>
        <ol>
            <li>拖放PDF文件到左侧文件列表区域</li>
            <li>选择功能选项卡（合并/拆分/转换）</li>
            <li>配置功能选项</li>
            <li>设置输出位置</li>
            <li>点击"开始执行"按钮</li>
            <li>转换过程中可使用"暂停/继续"按钮</li>
        </ol>

        <h2>输出选项</h2>
        <ul>
            <li><b>保存到源文件目录</b>：在原始PDF所在目录输出</li>
            <li><b>创建新子文件夹</b>：在原始目录下创建指定名称的子文件夹</li>
            <li><b>保存到指定目录</b>：选择自定义输出目录</li>
        </ul>

        <h2>注意事项</h2>
        <ul>
            <li>拆分功能仅支持单个PDF文件操作</li>
            <li>合并功能需要至少两个PDF文件</li>
            <li>高分辨率转换可能需要更多内存和处理时间</li>
            <li>使用"end"表示文档的最后一页</li>
            <li>暂停功能仅适用于PDF转图片操作</li>
            <li>页码范围选择支持多段范围（如：1-3,5,7-end）</li>
        </ul>

        <p style="text-align: center; margin-top: 20px;">
            <b>版本信息：</b> v1.2 (2025-06-26) - 新增页码范围选择功能
        </p>
        """


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('PDF 工具箱')
        self.setGeometry(100, 100, 850, 650)
        self.dropped_files = []  # 存储所有拖拽的文件路径
        self.worker = None  # PDF处理工作线程
        self.convert_files = []  # 存储转换文件的路径
        self.is_paused = False  # 新增：暂停状态

        self._setup_ui()
        self._setup_menu()  # 添加菜单栏
        self._connect_signals()
        self._apply_styles()
        self.status_bar.showMessage("准备就绪。请拖拽PDF文件到左侧列表。")

    def _setup_menu(self):
        """添加菜单栏和关于菜单"""
        menubar = self.menuBar()

        # 文件菜单
        file_menu = menubar.addMenu("文件")

        exit_action = QAction("退出", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 帮助菜单
        help_menu = menubar.addMenu("帮助")

        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

    def show_about_dialog(self):
        """显示关于对话框"""
        dialog = AboutDialog(self)
        dialog.exec()

    def _setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # --- 左侧面板: 文件列表 ---
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)

        left_layout.addWidget(QLabel("待处理PDF文件列表:"))
        self.file_list_widget = QListWidget()
        self.file_list_widget.setAcceptDrops(True)
        self.file_list_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.file_list_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        left_layout.addWidget(self.file_list_widget, 1)

        # 文件操作按钮
        file_btn_layout = QHBoxLayout()
        self.delete_selected_btn = QPushButton("删除选中")
        self.clear_list_btn = QPushButton("清空列表")
        file_btn_layout.addWidget(self.delete_selected_btn)
        file_btn_layout.addWidget(self.clear_list_btn)
        left_layout.addLayout(file_btn_layout)

        main_layout.addWidget(left_panel, 1)

        # --- 右侧面板: 功能选项卡 ---
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        self.tabs = QTabWidget()
        self.merge_tab = self._create_merge_tab()
        self.split_tab = self._create_split_tab()
        self.convert_tab = self._create_convert_tab()
        self.tabs.addTab(self.merge_tab, "合并PDF")
        self.tabs.addTab(self.split_tab, "拆分PDF")
        self.tabs.addTab(self.convert_tab, "PDF转图片")
        right_layout.addWidget(self.tabs, 1)

        # --- 输出位置配置 ---
        output_group = QGroupBox("输出位置选项")
        output_layout = QVBoxLayout(output_group)

        self.save_source_radio = QRadioButton("保存到源文件目录")
        self.save_subdir_radio = QRadioButton("在源文件目录创建新子文件夹")
        self.subdir_name_edit = QLineEdit("output_files")
        self.subdir_name_edit.setEnabled(False)

        custom_path_widget = QWidget()
        custom_layout = QHBoxLayout(custom_path_widget)
        custom_layout.setContentsMargins(0, 0, 0, 0)
        self.save_custom_radio = QRadioButton("保存到指定目录:")
        self.custom_dir_edit = QLineEdit()
        self.custom_dir_edit.setReadOnly(True)
        self.browse_dir_btn = QPushButton("浏览...")
        self.browse_dir_btn.setEnabled(False)

        custom_layout.addWidget(self.custom_dir_edit, 1)
        custom_layout.addWidget(self.browse_dir_btn)

        output_layout.addWidget(self.save_source_radio)
        output_layout.addWidget(self.save_subdir_radio)
        output_layout.addWidget(self.subdir_name_edit)
        output_layout.addWidget(self.save_custom_radio)
        output_layout.addWidget(custom_path_widget)

        right_layout.addWidget(output_group)

        # --- 执行按钮和状态栏 ---
        # 新增：暂停/继续按钮
        self.action_layout = QHBoxLayout()

        self.execute_btn = QPushButton("开始执行")
        self.execute_btn.setFixedHeight(40)
        self.action_layout.addWidget(self.execute_btn)

        # 新增：暂停/继续按钮
        self.pause_btn = QPushButton("暂停")
        self.pause_btn.setFixedHeight(40)
        self.pause_btn.setEnabled(False)  # 初始不可用
        self.action_layout.addWidget(self.pause_btn)

        right_layout.addLayout(self.action_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        right_layout.addWidget(self.progress_bar)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        main_layout.addWidget(right_panel, 2)

    def _create_merge_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        group = QGroupBox("合并选项")
        grid = QGridLayout(group)

        grid.addWidget(QLabel("合并后文件名:"), 0, 0)
        self.merge_filename_edit = QLineEdit("merged_document.pdf")
        grid.addWidget(self.merge_filename_edit, 0, 1)

        layout.addWidget(group)
        layout.addStretch()
        return tab

    def _create_split_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        group = QGroupBox("拆分选项")
        vbox = QVBoxLayout(group)

        self.split_range_radio = QRadioButton("按范围拆分 (如: 1-3; 5; 7-end)")
        self.split_range_edit = QLineEdit("1-end")
        self.split_all_radio = QRadioButton("每页单独拆分")
        self.split_range_radio.setChecked(True)

        vbox.addWidget(self.split_range_radio)
        vbox.addWidget(self.split_range_edit)
        vbox.addWidget(self.split_all_radio)

        layout.addWidget(group)
        layout.addStretch()
        return tab

    def _create_convert_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # 转换设置
        settings_group = QGroupBox("转换设置")
        settings_layout = QGridLayout(settings_group)

        settings_layout.addWidget(QLabel("分辨率 (DPI):"), 0, 0)
        self.resolution_combo = QComboBox()
        self.resolution_combo.addItems(["72", "96", "150", "200", "300", "400", "600"])
        self.resolution_combo.setCurrentText("200")
        settings_layout.addWidget(self.resolution_combo, 0, 1)

        settings_layout.addWidget(QLabel("图片格式:"), 1, 0)
        self.format_combo = QComboBox()
        self.format_combo.addItems(["png", "jpg"])
        settings_layout.addWidget(self.format_combo, 1, 1)

        # 新增：页面范围选择
        settings_layout.addWidget(QLabel("页面范围:"), 2, 0)
        page_range_layout = QHBoxLayout()

        self.convert_all_radio = QRadioButton("全部页面")
        self.convert_range_radio = QRadioButton("指定范围:")
        self.convert_range_edit = QLineEdit()
        self.convert_range_edit.setPlaceholderText("例如: 1-3,5,7-end")
        self.convert_range_edit.setEnabled(False)

        page_range_layout.addWidget(self.convert_all_radio)
        page_range_layout.addWidget(self.convert_range_radio)
        page_range_layout.addWidget(self.convert_range_edit, 1)  # 输入框占据剩余空间

        settings_layout.addLayout(page_range_layout, 2, 1)

        # 设置默认值
        self.convert_all_radio.setChecked(True)
        self.convert_range_radio.toggled.connect(self.convert_range_edit.setEnabled)

        layout.addWidget(settings_group)
        layout.addStretch()
        return tab

    def _connect_signals(self):
        # 文件列表操作
        self.file_list_widget.customContextMenuRequested.connect(self.show_context_menu)
        self.delete_selected_btn.clicked.connect(self.delete_selected_files)
        self.clear_list_btn.clicked.connect(self.clear_file_list)
        self.browse_dir_btn.clicked.connect(self.browse_custom_dir)

        # 单选按钮切换
        self.save_subdir_radio.toggled.connect(
            lambda checked: self.subdir_name_edit.setEnabled(checked)
        )

        self.save_custom_radio.toggled.connect(
            lambda checked: self._set_custom_dir_enabled(checked)
        )

        # 执行按钮
        self.execute_btn.clicked.connect(self.execute_action)

        # 新增：暂停按钮信号
        self.pause_btn.clicked.connect(self.toggle_pause)

    def _set_custom_dir_enabled(self, checked):
        """设置自定义目录相关控件的启用状态"""
        self.browse_dir_btn.setEnabled(checked)
        self.custom_dir_edit.setEnabled(checked)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        new_files = []
        for url in event.mimeData().urls():
            if url.isLocalFile() and url.toLocalFile().lower().endswith('.pdf'):
                file_path = url.toLocalFile()
                if file_path not in self.dropped_files:
                    self.dropped_files.append(file_path)
                    item = QListWidgetItem(os.path.basename(file_path))
                    item.setData(Qt.ItemDataRole.UserRole, file_path)
                    item.setToolTip(file_path)
                    self.file_list_widget.addItem(item)
                    new_files.append(file_path)

        if new_files:
            self.status_bar.showMessage(f"已添加 {len(new_files)} 个PDF文件")
        else:
            self.status_bar.showMessage("未添加新文件 (可能已存在或非PDF)")

    def show_context_menu(self, pos):
        menu = QMenu()
        delete_action = menu.addAction("删除选中项")
        action = menu.exec(self.file_list_widget.mapToGlobal(pos))
        if action == delete_action:
            self.delete_selected_files()

    def delete_selected_files(self):
        selected_items = self.file_list_widget.selectedItems()
        if not selected_items:
            return

        for item in selected_items:
            row = self.file_list_widget.row(item)
            file_path = item.data(Qt.ItemDataRole.UserRole)
            if file_path in self.dropped_files:
                self.dropped_files.remove(file_path)
            self.file_list_widget.takeItem(row)

        self.status_bar.showMessage(f"已删除 {len(selected_items)} 个文件")

    def clear_file_list(self):
        self.file_list_widget.clear()
        self.dropped_files.clear()
        self.status_bar.showMessage("文件列表已清空")

    def browse_custom_dir(self):
        start_dir = self.custom_dir_edit.text() or os.path.expanduser("~")
        dir_path = QFileDialog.getExistingDirectory(self, "选择保存目录", start_dir)
        if dir_path:
            self.custom_dir_edit.setText(dir_path)

    # 新增：暂停/继续切换
    def toggle_pause(self):
        if not self.worker:
            return

        if self.is_paused:
            self.worker.resume()
            self.pause_btn.setText("暂停")
            self.status_bar.showMessage("操作已继续...")
            self.is_paused = False
        else:
            self.worker.pause()
            self.pause_btn.setText("继续")
            self.status_bar.showMessage("操作已暂停")
            self.is_paused = True

    def _get_output_config(self):
        config = {}
        if self.save_source_radio.isChecked():
            config["mode"] = "source_dir"
        elif self.save_subdir_radio.isChecked():
            config["mode"] = "new_subdir_in_source"
            config["subfolder_name"] = self.subdir_name_edit.text().strip() or "output_files"
        elif self.save_custom_radio.isChecked():
            config["mode"] = "custom_dir"
            config["path"] = self.custom_dir_edit.text()
            if not config["path"] or not os.path.isdir(config["path"]):
                QMessageBox.warning(self, "错误", "请选择有效的保存目录")
                return None
        return config

    def execute_action(self):
        if not self.dropped_files:
            QMessageBox.warning(self, "提示", "请先添加PDF文件")
            return

        output_config = self._get_output_config()
        if not output_config:
            return

        self.execute_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_bar.showMessage("正在处理...")
        self.is_paused = False  # 重置暂停状态
        QApplication.processEvents()

        tab_index = self.tabs.currentIndex()

        if tab_index == 0:  # 合并
            if len(self.dropped_files) < 2:
                QMessageBox.warning(self, "提示", "合并需要至少两个PDF文件")
                self.execute_btn.setEnabled(True)
                self.progress_bar.setVisible(False)
                return

            filename = self.merge_filename_edit.text().strip()
            if not filename:
                filename = "merged_document.pdf"
            if not filename.lower().endswith(".pdf"):
                filename += ".pdf"

            self.worker = PdfWorker("merge", self.dropped_files, output_config, filename)
            self.worker.finished.connect(self.handle_merge_result)
            self.worker.error.connect(self.handle_worker_error)
            self.worker.start()
            self.pause_btn.setEnabled(False)  # 合并操作不支持暂停

        elif tab_index == 1:  # 拆分
            if not self.dropped_files:
                return

            split_options = {}
            if self.split_all_radio.isChecked():
                split_options["type"] = "all_pages_separately"
            else:
                split_options["type"] = "range"
                split_options["ranges_str"] = self.split_range_edit.text().strip()
                if not split_options["ranges_str"]:
                    QMessageBox.warning(self, "提示", "请输入拆分范围")
                    self.execute_btn.setEnabled(True)
                    self.progress_bar.setVisible(False)
                    return

            self.worker = PdfWorker("split", self.dropped_files[0], output_config, split_options)
            self.worker.finished.connect(self.handle_split_result)
            self.worker.error.connect(self.handle_worker_error)
            self.worker.start()
            self.pause_btn.setEnabled(False)  # 拆分操作不支持暂停

        elif tab_index == 2:  # 转换
            self.progress_bar.setMaximum(100)
            self.progress_bar.setValue(0)

            try:
                resolution = int(self.resolution_combo.currentText())
            except:
                resolution = 200

            image_format = self.format_combo.currentText()
            save_option = "source_dir"  # 默认为源目录

            if self.save_subdir_radio.isChecked():
                save_option = "new_subdir_in_source"
            elif self.save_custom_radio.isChecked():
                save_option = "custom_dir"

            output_dir = self.custom_dir_edit.text() if save_option == "custom_dir" else ""

            # 获取页码范围
            if self.convert_all_radio.isChecked():
                page_range_str = ""
            else:
                page_range_str = self.convert_range_edit.text().strip()

            self.worker = PdfWorker("convert", self.dropped_files, resolution,
                                    image_format, save_option, output_dir, page_range_str)
            self.worker.finished.connect(self.handle_convert_result)
            self.worker.error.connect(self.handle_worker_error)
            self.worker.progress.connect(self.update_progress)  # 新增：连接进度信号
            self.worker.start()

            # 启用暂停按钮
            self.pause_btn.setEnabled(True)
            self.pause_btn.setText("暂停")

    # 新增：更新进度条
    def update_progress(self, value):
        self.progress_bar.setValue(value)
        self.status_bar.showMessage(f"处理中... {value}% 完成")

    def handle_merge_result(self, result):
        self.progress_bar.setVisible(False)
        self.execute_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)  # 禁用暂停按钮

        if result["status"] == "success":
            QMessageBox.information(self, "成功", f"PDF合并完成!\n保存至: {result['file_path']}")
            self.status_bar.showMessage("合并完成")
        else:
            QMessageBox.critical(self, "失败", f"错误: {result['error_message']}")
            self.status_bar.showMessage("合并失败")

    def handle_split_result(self, results):
        self.progress_bar.setVisible(False)
        self.execute_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)  # 禁用暂停按钮

        success_files = [r for r in results if r["status"] == "success"]
        errors = [r["error_message"] for r in results if r["status"] == "failure"]

        if success_files:
            QMessageBox.information(self, "成功", f"拆分完成! 生成 {len(success_files)} 个文件")
            self.status_bar.showMessage(f"拆分完成，生成 {len(success_files)} 个文件")
        elif errors:
            QMessageBox.critical(self, "失败", "拆分失败:\n" + "\n".join(errors))
            self.status_bar.showMessage("拆分失败")
        else:
            QMessageBox.warning(self, "提示", "未生成任何文件，请检查设置")
            self.status_bar.showMessage("未生成文件")

    def handle_convert_result(self, results):
        self.progress_bar.setVisible(False)
        self.execute_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)  # 禁用暂停按钮

        success_count = sum(1 for r in results if r["status"] == "success")
        cancelled_count = sum(1 for r in results if r["status"] == "cancelled")
        errors = [r["error_message"] for r in results if r["status"] == "failure"]

        if cancelled_count > 0:
            QMessageBox.warning(self, "操作取消", f"用户取消了操作，已完成 {success_count} 个文件")
            self.status_bar.showMessage(f"操作已取消，完成 {success_count} 个文件")
        elif success_count == len(results):
            QMessageBox.information(self, "成功", f"所有 {len(results)} 个PDF已成功转换为图片")
            self.status_bar.showMessage("转换完成")
        elif success_count > 0:
            QMessageBox.warning(self, "部分成功",
                                f"{success_count} 个文件成功转换\n{len(errors)} 个文件失败:\n" + "\n".join(errors))
            self.status_bar.showMessage(f"部分成功 ({success_count}/{len(results)})")
        else:
            QMessageBox.critical(self, "失败", "所有转换均失败:\n" + "\n".join(errors))
            self.status_bar.showMessage("转换失败")

    def handle_worker_error(self, error_msg):
        self.progress_bar.setVisible(False)
        self.execute_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)  # 禁用暂停按钮
        QMessageBox.critical(self, "错误", f"处理过程中出错:\n{error_msg}")
        self.status_bar.showMessage("处理错误")

    def closeEvent(self, event):
        """关闭窗口时确保工作线程停止"""
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait(2000)  # 等待2秒线程结束
        event.accept()

    def _apply_styles(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d0d0d0;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 20px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 8px;
                left: 12px;
            }
            QPushButton {
                padding: 6px 12px;
                background-color: #e0e0e0;
                border: 1px solid #c0c0c0;
                border-radius: 4px;
                min-height: 28px;
            }
            QPushButton:hover {
                background-color: #d0d0d0;
            }
            QPushButton:pressed {
                background-color: #c0c0c0;
            }
            QPushButton#execute_btn {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton#execute_btn:hover {
                background-color: #45a049;
            }
            QPushButton#execute_btn:pressed {
                background-color: #3d8b40;
            }
            QPushButton#pause_btn {
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton#pause_btn:hover {
                background-color: #0b7dda;
            }
            QPushButton#pause_btn:pressed {
                background-color: #0069c0;
            }
            QLineEdit, QComboBox {
                padding: 6px;
                border: 1px solid #c0c0c0;
                border-radius: 4px;
                background-color: white;
            }
            QListWidget {
                border: 1px solid #c0c0c0;
                background-color: white;
                border-radius: 4px;
            }
            QTabWidget::pane {
                border: 1px solid #d0d0d0;
                border-top: none;
                border-radius: 0 0 6px 6px;
                background-color: white;
            }
            QTabBar::tab {
                padding: 8px 16px;
                background: #e0e0e0;
                border: 1px solid #c0c0c0;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom-color: white;
            }
            QStatusBar {
                background-color: #e0e0e0;
                border-top: 1px solid #d0d0d0;
                padding: 2px;
            }
            QProgressBar {
                border: 1px solid #c0c0c0;
                border-radius: 4px;
                text-align: center;
                background-color: white;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                width: 10px;
            }
        """)
        self.execute_btn.setObjectName("execute_btn")
        self.pause_btn.setObjectName("pause_btn")  # 新增：暂停按钮样式


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())