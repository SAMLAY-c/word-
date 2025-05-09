import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser, simpledialog
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
from docx.oxml import parse_xml
import os
import threading
import json

class DocxFormatter:
    def __init__(self, root):
        self.root = root
        self.root.title("Word文档格式化工具")
        self.root.geometry("1000x700")
        
        # 设置样式变量
        self.title_font = tk.StringVar(value="黑体")
        self.title_size = tk.IntVar(value=22)
        self.title_color = "#000000"  # 默认黑色
        self.title_bold = tk.BooleanVar(value=True)
        
        self.h1_font = tk.StringVar(value="黑体")
        self.h1_size = tk.IntVar(value=18)
        self.h1_color = "#000000"
        self.h1_bold = tk.BooleanVar(value=True)
        
        self.h2_font = tk.StringVar(value="楷体")
        self.h2_size = tk.IntVar(value=16)
        self.h2_color = "#000000"
        self.h2_bold = tk.BooleanVar(value=True)
        
        self.h3_font = tk.StringVar(value="宋体")
        self.h3_size = tk.IntVar(value=14)
        self.h3_color = "#000000"
        self.h3_bold = tk.BooleanVar(value=True)
        
        self.normal_font = tk.StringVar(value="仿宋")
        self.normal_size = tk.IntVar(value=12)
        self.normal_color = "#000000"
        self.normal_bold = tk.BooleanVar(value=False)
        
        self.toc_title = tk.StringVar(value="目 录")
        self.filename = tk.StringVar()
        
        # 内容结构数据
        self.document_sections = []
        
        # 创建用于存储文档标题的变量
        self.document_title = tk.StringVar(value="公文标题示例")
        
        self.log_text = None
        self.progress_bar = None
        
        # 创建界面元素
        self.create_widgets()
        
    def create_widgets(self):
        # 创建选项卡 
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 基本设置选项卡
        basic_frame = ttk.Frame(notebook)
        notebook.add(basic_frame, text="基本设置")
        
        # 内容设置选项卡
        content_frame = ttk.Frame(notebook)
        notebook.add(content_frame, text="文档内容")
        
        # 标题设置选项卡
        title_frame = ttk.Frame(notebook)
        notebook.add(title_frame, text="标题设置")
        
        # 字体设置选项卡
        styles_frame = ttk.Frame(notebook)
        notebook.add(styles_frame, text="字体设置")
        
        # 日志选项卡
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="执行日志")
        
        # === 基本设置选项卡内容 ===
        ttk.Label(basic_frame, text="文件名称(不需要输入.docx):", font=("Arial", 10)).grid(row=0, column=0, sticky="w", padx=10, pady=10)
        ttk.Entry(basic_frame, textvariable=self.filename, width=30).grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(basic_frame, text="目录标题:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", padx=10, pady=10)
        ttk.Entry(basic_frame, textvariable=self.toc_title, width=30).grid(row=1, column=1, padx=10, pady=10)
        
        ttk.Label(basic_frame, text="文档标题:", font=("Arial", 10)).grid(row=2, column=0, sticky="w", padx=10, pady=10)
        ttk.Entry(basic_frame, textvariable=self.document_title, width=30).grid(row=2, column=1, padx=10, pady=10)
        
        # === 内容设置选项卡 ===
        self.setup_content_tab(content_frame)
        
        # === 标题设置选项卡内容 ===
        self.create_font_settings(title_frame, 0, "文档标题", self.title_font, self.title_size, self.title_bold, self.title_color)
        self.create_font_settings(title_frame, 1, "一级标题", self.h1_font, self.h1_size, self.h1_bold, self.h1_color)
        self.create_font_settings(title_frame, 2, "二级标题", self.h2_font, self.h2_size, self.h2_bold, self.h2_color)
        self.create_font_settings(title_frame, 3, "三级标题", self.h3_font, self.h3_size, self.h3_bold, self.h3_color)
        
        # === 字体设置选项卡内容 ===
        self.create_font_settings(styles_frame, 0, "正文", self.normal_font, self.normal_size, self.normal_bold, self.normal_color)
        
        # 添加首行缩进选项
        ttk.Label(styles_frame, text="首行缩进字符数:").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        self.indent_entry = ttk.Entry(styles_frame, width=5)
        self.indent_entry.insert(0, "2")
        self.indent_entry.grid(row=1, column=1, sticky="w", padx=10, pady=10)
        
        # === 日志选项卡内容 ===
        self.log_text = tk.Text(log_frame, height=20, width=80, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.log_text.config(state=tk.DISABLED)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 进度条
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=980, mode="determinate")
        self.progress_bar.pack(pady=10)
        
        # 底部按钮
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="生成文档", command=self.generate_document).pack(side=tk.RIGHT, padx=5)
        
        # 初始化一些默认章节
        self.init_default_sections()
    
    def init_default_sections(self):
        """初始化一些默认的章节结构作为示例"""
        # 添加一个默认的一级标题
        self.add_section(level=1, title="第一章 背景介绍", content="这里是背景介绍内容。")
    
    def setup_content_tab(self, parent):
        """设置文档内容选项卡"""
        # 创建左右分栏布局
        paned_window = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左侧为章节树形结构
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=1)
        
        # 右侧为章节内容编辑
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=2)
        
        # 设置左侧树形结构
        ttk.Label(left_frame, text="文档结构").pack(pady=5)
        
        # 创建树形视图
        self.tree = ttk.Treeview(left_frame, selectmode='browse')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 设置树形视图列
        self.tree["columns"] = ("title", "level")
        self.tree.column("#0", width=50, minwidth=50, stretch=tk.NO)
        self.tree.column("title", width=150, minwidth=100, stretch=tk.YES)
        self.tree.column("level", width=50, minwidth=50, stretch=tk.NO)
        
        # 配置列标题
        self.tree.heading("#0", text="序号")
        self.tree.heading("title", text="标题")
        self.tree.heading("level", text="级别")
        
        # 树形视图滚动条
        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        
        # 添加树形视图点击事件
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        # 在左侧添加按钮区域
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(button_frame, text="添加章节", command=self.add_section_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="编辑章节", command=self.edit_section_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="删除章节", command=self.delete_section).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="上移", command=lambda: self.move_section(-1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="下移", command=lambda: self.move_section(1)).pack(side=tk.LEFT, padx=2)
        
        # 设置右侧内容编辑区域
        ttk.Label(right_frame, text="章节标题:").pack(anchor=tk.W, padx=5, pady=5)
        self.section_title_var = tk.StringVar()
        self.section_title_entry = ttk.Entry(right_frame, textvariable=self.section_title_var, width=40)
        self.section_title_entry.pack(fill=tk.X, padx=5, pady=2)
        
        # 章节级别选择
        level_frame = ttk.Frame(right_frame)
        level_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(level_frame, text="章节级别:").pack(side=tk.LEFT)
        self.section_level_var = tk.IntVar(value=1)
        ttk.Radiobutton(level_frame, text="一级", variable=self.section_level_var, value=1).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(level_frame, text="二级", variable=self.section_level_var, value=2).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(level_frame, text="三级", variable=self.section_level_var, value=3).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(right_frame, text="章节内容:").pack(anchor=tk.W, padx=5, pady=5)
        self.section_content_text = tk.Text(right_frame, wrap=tk.WORD, height=15)
        self.section_content_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 内容编辑区域滚动条
        content_scroll = ttk.Scrollbar(right_frame, orient="vertical", command=self.section_content_text.yview)
        content_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.section_content_text.configure(yscrollcommand=content_scroll.set)
        
        # 右侧底部按钮
        bottom_frame = ttk.Frame(right_frame)
        bottom_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(bottom_frame, text="保存章节", command=self.save_current_section).pack(side=tk.RIGHT, padx=5)
        
        # 初始化当前编辑的章节ID
        self.current_section_id = None
    
    def add_section_dialog(self):
        """打开添加章节对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("添加章节")
        dialog.geometry("400x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="章节标题:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
        title_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=title_var, width=30).grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="章节级别:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        level_var = tk.IntVar(value=1)
        ttk.Radiobutton(dialog, text="一级", variable=level_var, value=1).grid(row=1, column=1, sticky=tk.W)
        ttk.Radiobutton(dialog, text="二级", variable=level_var, value=2).grid(row=1, column=1)
        ttk.Radiobutton(dialog, text="三级", variable=level_var, value=3).grid(row=1, column=1, sticky=tk.E)
        
        def on_confirm():
            title = title_var.get().strip()
            level = level_var.get()
            if title:
                self.add_section(level, title)
                dialog.destroy()
            else:
                messagebox.showwarning("警告", "章节标题不能为空")
        
        ttk.Button(dialog, text="确定", command=on_confirm).grid(row=2, column=1, sticky=tk.E, padx=10, pady=10)
        ttk.Button(dialog, text="取消", command=dialog.destroy).grid(row=2, column=0, sticky=tk.W, padx=10, pady=10)
        
        # 设置对话框为模态
        dialog.wait_window()
    
    def edit_section_dialog(self):
        """打开编辑章节对话框"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择要编辑的章节")
            return
        
        section_id = selected[0]
        section = self.find_section_by_id(section_id)
        if not section:
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑章节")
        dialog.geometry("400x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="章节标题:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)
        title_var = tk.StringVar(value=section['title'])
        ttk.Entry(dialog, textvariable=title_var, width=30).grid(row=0, column=1, padx=10, pady=10)
        
        ttk.Label(dialog, text="章节级别:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)
        level_var = tk.IntVar(value=section['level'])
        ttk.Radiobutton(dialog, text="一级", variable=level_var, value=1).grid(row=1, column=1, sticky=tk.W)
        ttk.Radiobutton(dialog, text="二级", variable=level_var, value=2).grid(row=1, column=1)
        ttk.Radiobutton(dialog, text="三级", variable=level_var, value=3).grid(row=1, column=1, sticky=tk.E)
        
        def on_confirm():
            title = title_var.get().strip()
            level = level_var.get()
            if title:
                section['title'] = title
                section['level'] = level
                self.update_tree()
                dialog.destroy()
            else:
                messagebox.showwarning("警告", "章节标题不能为空")
        
        ttk.Button(dialog, text="确定", command=on_confirm).grid(row=2, column=1, sticky=tk.E, padx=10, pady=10)
        ttk.Button(dialog, text="取消", command=dialog.destroy).grid(row=2, column=0, sticky=tk.W, padx=10, pady=10)
        
        # 设置对话框为模态
        dialog.wait_window()
    
    def add_section(self, level, title, content=""):
        """添加一个新章节"""
        section = {
            'id': f"section_{len(self.document_sections)}",
            'level': level,
            'title': title,
            'content': content
        }
        self.document_sections.append(section)
        self.update_tree()
        return section['id']
    
    def update_tree(self):
        """更新树形视图"""
        # 清空树形视图
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 填充树形视图
        for i, section in enumerate(self.document_sections):
            self.tree.insert("", tk.END, section['id'], text=str(i+1), 
                            values=(section['title'], f"级别{section['level']}"))
    
    def find_section_by_id(self, section_id):
        """根据ID查找章节"""
        for section in self.document_sections:
            if section['id'] == section_id:
                return section
        return None
    
    def on_tree_select(self, event):
        """处理树形视图的选择事件"""
        selected = self.tree.selection()
        if not selected:
            return
        
        section_id = selected[0]
        section = self.find_section_by_id(section_id)
        if section:
            # 保存当前编辑的章节
            if self.current_section_id:
                self.save_current_section()
            
            # 加载选中的章节到编辑区
            self.current_section_id = section_id
            self.section_title_var.set(section['title'])
            self.section_level_var.set(section['level'])
            
            # 清空并设置内容
            self.section_content_text.delete(1.0, tk.END)
            self.section_content_text.insert(tk.END, section['content'])
    
    def save_current_section(self):
        """保存当前正在编辑的章节"""
        if not self.current_section_id:
            return
        
        section = self.find_section_by_id(self.current_section_id)
        if section:
            section['title'] = self.section_title_var.get()
            section['level'] = self.section_level_var.get()
            section['content'] = self.section_content_text.get(1.0, tk.END).strip()
            self.update_tree()
    
    def delete_section(self):
        """删除选中的章节"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择要删除的章节")
            return
        
        section_id = selected[0]
        result = messagebox.askyesno("确认", "确定要删除选中的章节吗？")
        if result:
            # 删除章节
            self.document_sections = [s for s in self.document_sections if s['id'] != section_id]
            self.update_tree()
            
            # 清空编辑区
            if self.current_section_id == section_id:
                self.current_section_id = None
                self.section_title_var.set("")
                self.section_content_text.delete(1.0, tk.END)
    
    def move_section(self, direction):
        """上移或下移章节"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择要移动的章节")
            return
        
        section_id = selected[0]
        for i, section in enumerate(self.document_sections):
            if section['id'] == section_id:
                new_pos = i + direction
                if 0 <= new_pos < len(self.document_sections):
                    # 交换位置
                    self.document_sections[i], self.document_sections[new_pos] = \
                        self.document_sections[new_pos], self.document_sections[i]
                    self.update_tree()
                    self.tree.selection_set(section_id)
                break
    
    def create_font_settings(self, parent, row, label_text, font_var, size_var, bold_var, color_var_name):
        # 创建一个框架来容纳这一行的设置
        frame = ttk.LabelFrame(parent, text=label_text)
        frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=10, pady=10)
        
        # 字体选择
        ttk.Label(frame, text="字体:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        font_combo = ttk.Combobox(frame, textvariable=font_var, values=["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "Times New Roman", "Arial"])
        font_combo.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        # 字号选择
        ttk.Label(frame, text="字号:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        size_combo = ttk.Combobox(frame, textvariable=size_var, values=[8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72])
        size_combo.grid(row=0, column=3, sticky="w", padx=5, pady=5)
        
        # 加粗选项
        bold_check = ttk.Checkbutton(frame, text="加粗", variable=bold_var)
        bold_check.grid(row=0, column=4, sticky="w", padx=5, pady=5)
        
        # 颜色选择按钮
        color_button = ttk.Button(frame, text="颜色", command=lambda: self.choose_color(color_var_name))
        color_button.grid(row=0, column=5, sticky="w", padx=5, pady=5)
        
        # 颜色预览
        self.color_preview = tk.Canvas(frame, width=16, height=16, bg=color_var_name, highlightthickness=1, highlightbackground="black")
        self.color_preview.grid(row=0, column=6, sticky="w", padx=5, pady=5)
        
        # 存储引用以便更新
        setattr(self, f"{color_var_name.replace('#', '')}_preview", self.color_preview)
    
    def choose_color(self, color_var_name):
        color = colorchooser.askcolor(initial=color_var_name)[1]
        if color:
            # 更新颜色变量
            setattr(self, color_var_name.replace('#', ''), color)
            # 更新预览
            preview_canvas = getattr(self, f"{color_var_name.replace('#', '')}_preview")
            preview_canvas.config(bg=color)
    
    def log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
    
    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.root.update()
    
    def generate_document(self):
        # 先保存当前编辑的章节
        self.save_current_section()
        
        if not self.filename.get():
            messagebox.showerror("错误", "请输入文件名称")
            return
        
        # 检查是否有章节
        if not self.document_sections:
            messagebox.showerror("错误", "文档内容为空，请至少添加一个章节")
            return
        
        # 启动一个线程来执行文档生成，以保持UI响应
        threading.Thread(target=self.generate_document_thread, daemon=True).start()
    
    def generate_document_thread(self):
        try:
            # 清空日志
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete(1.0, tk.END)
            self.log_text.config(state=tk.DISABLED)
            
            self.log("开始文档生成过程...")
            self.update_progress(5)
            
            # 获取用户选择的目标文件夹
            save_dir = filedialog.askdirectory(title="选择保存文档的文件夹")
            if not save_dir:
                self.log("用户取消了文件夹选择，操作终止。")
                self.update_progress(0)
                return
            
            # 构建完整文件路径
            doc_path = os.path.join(save_dir, f"{self.filename.get()}.docx")
            self.log(f"文档将保存至: {doc_path}")
            self.update_progress(10)
            
            # 创建文档
            document = Document()
            self.log("已创建新文档...")
            self.update_progress(15)
            
            # 设置文档页面格式（A4纸，常见公文格式）
            self.log("设置文档页面格式...")
            sections = document.sections
            for section in sections:
                section.page_height = Inches(11.69)  # A4高度
                section.page_width = Inches(8.27)    # A4宽度
                section.left_margin = Inches(1)      # 左边距
                section.right_margin = Inches(1)     # 右边距
                section.top_margin = Inches(1)       # 上边距
                section.bottom_margin = Inches(0.8)  # 下边距
            self.update_progress(20)
            
            # 创建样式
            self.log("创建文档样式...")
            
            # 文档标题样式
            self.log("创建文档标题样式...")
            title_style = self.create_style(document, 'Title', '文档标题', 
                                           self.title_font.get(), self.title_size.get(), 
                                           self.title_bold.get(), self.title_color)
            self.update_progress(30)
            
            # 一级标题样式
            self.log("创建一级标题样式...")
            h1_style = self.create_style(document, 'Heading1', '一级标题', 
                                        self.h1_font.get(), self.h1_size.get(), 
                                        self.h1_bold.get(), self.h1_color, level=1)
            self.update_progress(40)
            
            # 二级标题样式
            self.log("创建二级标题样式...")
            h2_style = self.create_style(document, 'Heading2', '二级标题', 
                                        self.h2_font.get(), self.h2_size.get(), 
                                        self.h2_bold.get(), self.h2_color, level=2)
            self.update_progress(50)
            
            # 三级标题样式
            self.log("创建三级标题样式...")
            h3_style = self.create_style(document, 'Heading3', '三级标题', 
                                        self.h3_font.get(), self.h3_size.get(), 
                                        self.h3_bold.get(), self.h3_color, level=3)
            self.update_progress(60)
            
            # 正文样式
            self.log("创建正文样式...")
            normal_style = document.styles['Normal']
            normal_style.font.name = self.normal_font.get()
            normal_style.font.size = Pt(self.normal_size.get())
            normal_style.font.bold = self.normal_bold.get()
            normal_style.font.color.rgb = self.rgb_from_hex(self.normal_color)
            normal_style._element.rPr.rFonts.set(qn('w:eastAsia'), self.normal_font.get())
            
            # 设置首行缩进
            indent_chars = int(self.indent_entry.get())
            normal_style.paragraph_format.first_line_indent = Pt(indent_chars * self.normal_size.get())
            normal_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            self.update_progress(70)
            
            # 添加文档内容
            self.log("添加文档内容...")
            
            # 添加文档标题
            self.log("添加文档标题...")
            title_para = document.add_paragraph(self.document_title.get(), style='Title')
            title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self.update_progress(75)
            
            # 添加目录
            self.log("添加目录...")
            self.add_toc(document, self.toc_title.get())
            document.add_page_break()  # 目录后添加分页符
            self.update_progress(80)
            
            # 添加用户定义的内容
            self.log("添加文档主体内容...")
            self.add_user_document_content(document)
            self.update_progress(90)
            
            # 保存文档
            self.log("保存文档...")
            document.save(doc_path)
            self.update_progress(100)
            
            self.log(f"文档已成功保存至 {doc_path}")
            self.log("请在Word中打开文档，右键点击目录，选择'更新域'或按F9更新目录。")
            
            # 弹出成功消息
            messagebox.showinfo("成功", f"文档已成功生成并保存至:\n{doc_path}\n\n请在Word中打开文档后，右键点击目录，选择'更新域'更新目录。")
            
        except Exception as e:
            self.log(f"错误: {str(e)}")
            messagebox.showerror("错误", f"生成文档时发生错误:\n{str(e)}")
            self.update_progress(0)
    
    def add_user_document_content(self, document):
        """添加用户定义的文档内容"""
        self.log("开始添加用户定义的文档内容...")
        
        # 遍历所有章节并添加到文档
        for section in self.document_sections:
            self.log(f"添加章节: {section['title']} (级别 {section['level']})")
            
            # 根据级别选择样式
            if section['level'] == 1:
                style = 'Heading1'
            elif section['level'] == 2:
                style = 'Heading2'
            elif section['level'] == 3:
                style = 'Heading3'
            else:
                style = 'Normal'
            
            # 添加标题
            document.add_paragraph(section['title'], style=style)
            
            # 添加内容
            if section['content']:
                document.add_paragraph(section['content'], style='Normal')
        
        self.log("所有用户定义的内容已添加完成")
    
    def create_style(self, document, style_id, style_name, font_name, font_size, bold, color, level=None):
        """创建标题样式"""
        try:
            self.log(f"创建样式: {style_name}, 字体: {font_name}, 大小: {font_size}pt")
            
            # 检查样式是否已存在
            if style_id in document.styles:
                style = document.styles[style_id]
            else:
                style = document.styles.add_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
            
            # 设置样式基本属性
            style.name = style_name
            style.hidden = False
            style.quick_style = True
            
            # 如果是标题样式，设置大纲级别
            if level:
                style.base_style = document.styles['Heading ' + str(level)]
                style.next_paragraph_style = document.styles['Normal']
                # 为自动生成目录设置大纲级别
                p_pr = style._element.get_or_add_pPr()
                p_pr.get_or_add_numPr().get_or_add_ilvl().val = level - 1
                p_pr.get_or_add_numPr().get_or_add_numId().val = 0

            # 设置字体
            font = style.font
            font.name = font_name
            font.size = Pt(font_size)
            font.bold = bold
            font.color.rgb = self.rgb_from_hex(color)
            font._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            
            # 设置段落格式
            paragraph_format = style.paragraph_format
            paragraph_format.space_before = Pt(12)
            paragraph_format.space_after = Pt(4)
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            
            if style_id == 'Title':
                paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            return style
        except Exception as e:
            self.log(f"创建样式时出错: {str(e)}")
            return None
    
    def rgb_from_hex(self, hex_color):
        """将十六进制颜色值转换为RGB对象"""
        hex_color = hex_color.lstrip('#')
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return RGBColor(r, g, b)
    
    def add_toc(self, document, toc_title="目录"):
        """添加自动生成的目录"""
        try:
            self.log(f"添加目录标题: {toc_title}")
            # 添加目录标题
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(toc_title)
            run.font.name = self.h1_font.get()
            run.font.size = Pt(16)
            run.font.bold = True
            run._element.rPr.rFonts.set(qn('w:eastAsia'), self.h1_font.get())
            
            self.log("添加目录字段")
            # 添加目录本身
            p = document.add_paragraph()
            run = p.add_run()
            
            # 创建目录XML元素
            fldChar = OxmlElement('w:fldChar')
            fldChar.set(qn('w:fldCharType'), 'begin')
            
            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = r' TOC \o "1-3" \h \z \u ' # 这个字符串设置了目录参数

            fldChar2 = OxmlElement('w:fldChar')
            fldChar2.set(qn('w:fldCharType'), 'separate')
            
            fldChar3 = OxmlElement('w:fldChar')
            fldChar3.set(qn('w:fldCharType'), 'end')
            
            # 将XML元素添加到段落中
            r_element = run._element
            r_element.append(fldChar)
            r_element.append(instrText)
            r_element.append(fldChar2)
            r_element.append(fldChar3)
            
            self.log("目录添加完成")
            return p
        except Exception as e:
            self.log(f"添加目录时出错: {str(e)}")
            return None

# 创建并运行应用
if __name__ == "__main__":
    root = tk.Tk()
    app = DocxFormatter(root)
    root.mainloop()