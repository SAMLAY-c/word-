import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser, simpledialog, scrolledtext
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
import os
import threading
import json
import requests
import re
import configparser
import time # Added for retry delay

class DocxFormatter:
    def __init__(self, root):
        self.root = root
        self.root.title("Word文档格式化工具") # Word Document Formatting Tool
        self.root.geometry("1000x700")
        
        # Style variables
        self.title_font = tk.StringVar(value="黑体") # HeiTi
        self.title_size = tk.IntVar(value=22)
        self.title_color_val = tk.StringVar(value="#000000") # Default black
        self.title_bold = tk.BooleanVar(value=True)
        
        self.h1_font = tk.StringVar(value="黑体") # HeiTi
        self.h1_size = tk.IntVar(value=18)
        self.h1_color_val = tk.StringVar(value="#000000")
        self.h1_bold = tk.BooleanVar(value=True)
        
        self.h2_font = tk.StringVar(value="楷体") # KaiTi
        self.h2_size = tk.IntVar(value=16)
        self.h2_color_val = tk.StringVar(value="#000000")
        self.h2_bold = tk.BooleanVar(value=True)
        
        self.h3_font = tk.StringVar(value="宋体") # SongTi
        self.h3_size = tk.IntVar(value=14)
        self.h3_color_val = tk.StringVar(value="#000000")
        self.h3_bold = tk.BooleanVar(value=True)
        
        self.normal_font = tk.StringVar(value="仿宋") # FangSong
        self.normal_size = tk.IntVar(value=12)
        self.normal_color_val = tk.StringVar(value="#000000")
        self.normal_bold = tk.BooleanVar(value=False)
        
        self.toc_title = tk.StringVar(value="目 录") # Table of Contents
        self.filename = tk.StringVar()
        
        self.document_sections = []
        self.document_title = tk.StringVar(value="公文标题示例") # Official Document Title Example
        
        self.deepseek_api_key = tk.StringVar()
        self.load_api_key()
        
        self.api_dependencies_status = tk.StringVar(value="未检查") # Not checked
        
        self.log_text = None 
        self.progress_bar = None
        self.color_previews = {}

        # Initialize a requests.Session object for API calls
        self.api_session = requests.Session()
        
        self.create_widgets()
        self.check_and_install_dependencies()
        
    def load_api_key(self):
        """从配置文件加载 API Key"""
        config = configparser.ConfigParser()
        config_file = "config.ini"
        if os.path.exists(config_file):
            config.read(config_file)
            if "DEEPSEEK" in config and "api_key" in config["DEEPSEEK"]:
                self.deepseek_api_key.set(config["DEEPSEEK"]["api_key"])
        else:
            config["DEEPSEEK"] = {"api_key": ""}
            with open(config_file, 'w') as f:
                config.write(f)
            
    def check_and_install_dependencies(self):
        """检查依赖项，并在缺失时提示用户手动安装"""
        dependencies = [
            ("docx", "python-docx"),
            ("requests", "requests")
        ]
        missing_dependencies = []
        for import_name, install_name in dependencies:
            try:
                __import__(import_name)
                self.log(f"{install_name} 已安装。")
            except ImportError:
                self.log(f"{install_name} 未安装。请在激活虚拟环境后运行: pip install {install_name}")
                missing_dependencies.append(install_name)
            except Exception as e:
                self.log(f"检查 {install_name} 时发生未知错误: {str(e)}")
                missing_dependencies.append(f"{install_name} (检查失败)")

        if not missing_dependencies:
            self.api_dependencies_status.set("所有依赖已就绪")
        else:
            self.api_dependencies_status.set(f"缺失依赖: {', '.join(missing_dependencies)}")
            messagebox.showwarning("依赖缺失", 
                                 f"以下依赖包未能加载或缺失，请在激活虚拟环境后，通过 pip 手动安装:\n\n{', '.join(missing_dependencies)}\n\n例如: pip install python-docx requests")

    def create_widgets(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.notebook = notebook
        
        basic_frame = ttk.Frame(notebook)
        notebook.add(basic_frame, text="基本设置")
        content_frame = ttk.Frame(notebook)
        notebook.add(content_frame, text="文档内容")
        ai_frame = ttk.Frame(notebook)
        notebook.add(ai_frame, text="AI 标题识别")
        title_settings_frame = ttk.Frame(notebook) 
        notebook.add(title_settings_frame, text="标题设置")
        styles_frame = ttk.Frame(notebook)
        notebook.add(styles_frame, text="字体设置")
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="执行日志")
        
        # === Basic Settings Tab Content ===
        ttk.Label(basic_frame, text="文件名称(不需要输入.docx):", font=("Arial", 10)).grid(row=0, column=0, sticky="w", padx=10, pady=10) 
        ttk.Entry(basic_frame, textvariable=self.filename, width=30).grid(row=0, column=1, padx=10, pady=10)
        ttk.Label(basic_frame, text="目录标题:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", padx=10, pady=10) 
        ttk.Entry(basic_frame, textvariable=self.toc_title, width=30).grid(row=1, column=1, padx=10, pady=10)
        ttk.Label(basic_frame, text="文档标题:", font=("Arial", 10)).grid(row=2, column=0, sticky="w", padx=10, pady=10) 
        ttk.Entry(basic_frame, textvariable=self.document_title, width=30).grid(row=2, column=1, padx=10, pady=10)
        
        api_frame = ttk.LabelFrame(basic_frame, text="DeepSeek API 设置")
        api_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=10)
        ttk.Label(api_frame, text="依赖状态:").grid(row=0, column=0, sticky="w", padx=5, pady=5) 
        ttk.Label(api_frame, textvariable=self.api_dependencies_status).grid(row=0, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(api_frame, text="API Key:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        api_entry = ttk.Entry(api_frame, textvariable=self.deepseek_api_key, width=40, show="*")
        api_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        def toggle_api_visibility():
            api_entry['show'] = '' if api_entry['show'] == '*' else '*'
            toggle_btn['text'] = '隐藏' if api_entry['show'] == '' else '显示'
                
        toggle_btn = ttk.Button(api_frame, text="显示", width=5, command=toggle_api_visibility)
        toggle_btn.grid(row=1, column=2, padx=5, pady=5)
        
        def save_api_key_action(): 
            config = configparser.ConfigParser()
            config["DEEPSEEK"] = {"api_key": self.deepseek_api_key.get()} 
            with open("config.ini", 'w') as f:
                config.write(f)
            messagebox.showinfo("成功", "API Key 已保存")
            
        ttk.Button(api_frame, text="保存", command=save_api_key_action).grid(row=1, column=3, padx=5, pady=5)
        
        self.setup_ai_tab(ai_frame)
        self.setup_content_tab(content_frame)
        
        self.create_font_settings(title_settings_frame, 0, "文档标题", self.title_font, self.title_size, self.title_bold, self.title_color_val, "title_color")
        self.create_font_settings(title_settings_frame, 1, "一级标题", self.h1_font, self.h1_size, self.h1_bold, self.h1_color_val, "h1_color")
        self.create_font_settings(title_settings_frame, 2, "二级标题", self.h2_font, self.h2_size, self.h2_bold, self.h2_color_val, "h2_color")
        self.create_font_settings(title_settings_frame, 3, "三级标题", self.h3_font, self.h3_size, self.h3_bold, self.h3_color_val, "h3_color")
        
        self.create_font_settings(styles_frame, 0, "正文", self.normal_font, self.normal_size, self.normal_bold, self.normal_color_val, "normal_color")
        ttk.Label(styles_frame, text="首行缩进字符数:").grid(row=1, column=0, sticky="w", padx=10, pady=10) 
        self.indent_entry = ttk.Entry(styles_frame, width=5)
        self.indent_entry.insert(0, "2")
        self.indent_entry.grid(row=1, column=1, sticky="w", padx=10, pady=10)
        
        self.log_text_widget = tk.Text(log_frame, height=20, width=80, wrap=tk.WORD) 
        self.log_text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.log_text_widget.config(state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text_widget.config(yscrollcommand=scrollbar.set)
        
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=980, mode="determinate")
        self.progress_bar.pack(pady=10)
        
        button_frame_main = ttk.Frame(self.root) 
        button_frame_main.pack(fill=tk.X, padx=10, pady=10)
        ttk.Button(button_frame_main, text="生成文档", command=self.generate_document).pack(side=tk.RIGHT, padx=5)
        
        self.init_default_sections()
    
    def init_default_sections(self):
        self.add_section(level=1, title="第一章 背景介绍", content="这里是背景介绍内容。") 
    
    def setup_content_tab(self, parent):
        paned_window = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=1)
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=2)
        
        ttk.Label(left_frame, text="文档结构").pack(pady=5)
        self.tree = ttk.Treeview(left_frame, selectmode='browse')
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.tree["columns"] = ("title", "level")
        self.tree.column("#0", width=50, minwidth=50, stretch=tk.NO)
        self.tree.column("title", width=150, minwidth=100, stretch=tk.YES)
        self.tree.column("level", width=50, minwidth=50, stretch=tk.NO)
        self.tree.heading("#0", text="序号") 
        self.tree.heading("title", text="标题") 
        self.tree.heading("level", text="级别") 
        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        button_frame_tree = ttk.Frame(left_frame) 
        button_frame_tree.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(button_frame_tree, text="添加章节", command=self.add_section_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame_tree, text="编辑章节", command=self.edit_section_dialog).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame_tree, text="删除章节", command=self.delete_section).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame_tree, text="上移", command=lambda: self.move_section(-1)).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame_tree, text="下移", command=lambda: self.move_section(1)).pack(side=tk.LEFT, padx=2)
        
        ttk.Label(right_frame, text="章节标题:").pack(anchor=tk.W, padx=5, pady=5)
        self.section_title_var = tk.StringVar()
        self.section_title_entry = ttk.Entry(right_frame, textvariable=self.section_title_var, width=40)
        self.section_title_entry.pack(fill=tk.X, padx=5, pady=2)
        
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
        content_scroll = ttk.Scrollbar(right_frame, orient="vertical", command=self.section_content_text.yview)
        content_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.section_content_text.configure(yscrollcommand=content_scroll.set)
        
        bottom_frame_content = ttk.Frame(right_frame) 
        bottom_frame_content.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(bottom_frame_content, text="保存章节", command=self.save_current_section).pack(side=tk.RIGHT, padx=5)
        self.current_section_id = None
    
    def add_section_dialog(self):
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
        dialog.wait_window()
    
    def edit_section_dialog(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择要编辑的章节")
            return
        section_id = selected[0]
        section = self.find_section_by_id(section_id)
        if not section: return
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
        dialog.wait_window()
    
    def add_section(self, level, title, content=""):
        section = {'id': f"section_{len(self.document_sections)}_{title.replace(' ','_')}", 'level': level, 'title': title, 'content': content}
        self.document_sections.append(section)
        self.update_tree()
        return section['id']
    
    def update_tree(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for i, section in enumerate(self.document_sections):
            self.tree.insert("", tk.END, section['id'], text=str(i+1), values=(section['title'], f"级别{section['level']}"))
    
    def find_section_by_id(self, section_id):
        for section in self.document_sections:
            if section['id'] == section_id: return section
        return None
    
    def setup_ai_tab(self, parent):
        ttk.Label(parent, text="粘贴您的文本内容，AI 将自动识别标题结构：").pack(anchor=tk.W, padx=10, pady=5)
        self.ai_input_text = scrolledtext.ScrolledText(parent, wrap=tk.WORD, height=15)
        self.ai_input_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        button_frame_ai = ttk.Frame(parent)
        button_frame_ai.pack(fill=tk.X, padx=10, pady=5)
        self.ai_status_var = tk.StringVar(value="准备就绪")
        status_label = ttk.Label(button_frame_ai, textvariable=self.ai_status_var)
        status_label.pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame_ai, text="识别标题并导入", command=self.analyze_with_deepseek).pack(side=tk.RIGHT, padx=5)
        tip_frame = ttk.LabelFrame(parent, text="使用说明")
        tip_frame.pack(fill=tk.X, padx=10, pady=10)
        tips = ("1. 粘贴您的文本内容到上方文本框中\n2. 点击「识别标题并导入」按钮\n3. AI 将分析文本，识别各级标题\n"
                "4. 识别出的标题结构将自动导入到「文档内容」选项卡中\n5. 您可以在「文档内容」选项卡进一步编辑调整\n\n"
                "提示：标题识别最适合结构化文档，如学术论文、报告等带有明确章节标题的文本\n\n"
                "注意：首次使用时需要联网安装必要的依赖包。\n请确保您已经安装了 Python 并设置了正确的 DeepSeek API Key。")
        ttk.Label(tip_frame, text=tips, justify=tk.LEFT).pack(padx=10, pady=10)
    
    def analyze_with_deepseek(self):
        api_key = self.deepseek_api_key.get()
        if not api_key:
            messagebox.showerror("错误", "请先设置 DeepSeek API Key")
            return
        text = self.ai_input_text.get(1.0, tk.END).strip()
        if not text:
            messagebox.showerror("错误", "请输入要分析的文本")
            return
        if self.api_dependencies_status.get().startswith("安装失败"):
            if messagebox.askyesno("依赖安装", "AI 识别所需的依赖包安装失败，是否尝试重新安装？"):
                self.check_and_install_dependencies()
            return 
        elif self.api_dependencies_status.get() == "正在安装...":
            messagebox.showinfo("提示", "依赖包仍在安装中，请稍后再试。")
            return
        self.ai_status_var.set("正在分析中...") 
        self.root.update_idletasks() 
        threading.Thread(target=self.run_deepseek_analysis, args=(text,), daemon=True).start()
    
    def run_deepseek_analysis(self, text):
        """在线程中运行 DeepSeek API 分析, 包含超时、会话和重试机制"""
        api_key = self.deepseek_api_key.get()
        api_url = "https://api.deepseek.com/chat/completions"
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        system_message_content = "你是一个专业的文本分析助手，负责识别文本中的标题结构，并严格按照用户指定的JSON格式返回结果。"
        user_prompt_content = f"""
请分析以下文本，识别其中的标题结构。将结果按如下JSON格式返回：
```json
[
    {{"level": 1, "title": "一级标题1", "content": "一级标题1下的正文内容"}},
    {{"level": 2, "title": "二级标题1", "content": "二级标题1下的正文内容"}},
    ...
]
```
规则：
1. level 表示标题级别，1为一级标题，2为二级标题，3为三级标题。
2. 识别标题时考虑格式特征，如数字编号、字体大小、缩进等。
3. content 是标题下的正文内容，不包括下一个标题。
4. 同一级别的标题，内容结构应保持一致。
5. 仅返回JSON格式，不要有其他解释文字。
以下是要分析的文本：
{text}
"""
        data = {
            "model": "deepseek-reasoner",
            "messages": [
                {"role": "system", "content": system_message_content},
                {"role": "user", "content": user_prompt_content}
            ],
            "stream": False
        }
        
        # --- 优化：增加超时时间、重试次数和退避因子 ---
        max_retries = 3
        backoff_factor = 0.5 # 每次重试后等待的时间增加因子
        request_timeout = 90  # 增加超时时间到90秒

        for attempt in range(max_retries):
            try:
                self.log(f"向 DeepSeek API 发送请求 (尝试 {attempt + 1}/{max_retries}): {api_url}")
                # 使用 self.api_session 发送请求
                response = self.api_session.post(api_url, headers=headers, json=data, timeout=request_timeout)
                self.log(f"DeepSeek API 响应状态码: {response.status_code}")

                if response.status_code == 200:
                    result = response.json()
                    if not result.get('choices') or not result['choices'][0].get('message') or not result['choices'][0]['message'].get('content'):
                        self.log(f"DeepSeek API 响应结构不符合预期: {result}")
                        # 不立即返回，允许重试其他错误
                        if attempt == max_retries - 1: # 如果是最后一次尝试
                             self.root.after(0, lambda: self.handle_ai_error("AI未能生成有效响应内容或响应结构错误。"))
                        raise requests.exceptions.RequestException("API响应结构错误") # 触发重试

                    text_response = result['choices'][0]['message']['content']
                    if not text_response.strip():
                        self.log("DeepSeek API 返回了空的 content。")
                        if attempt == max_retries - 1:
                            self.root.after(0, lambda: self.handle_ai_error("AI未能生成响应或响应为空。"))
                        raise requests.exceptions.RequestException("API返回空内容") # 触发重试

                    self.log(f"从 DeepSeek API 收到的原始响应内容 (前500字符): {text_response[:500]}")
                    json_match = re.search(r'```json\s*([\s\S]*?)\s*```', text_response, re.DOTALL)
                    json_str = json_match.group(1).strip() if json_match else text_response.strip()
                    
                    try:
                        sections = json.loads(json_str)
                        self.log(f"JSON解析成功，识别到 {len(sections)} 个章节。")
                        self.root.after(0, lambda: self.import_ai_sections(sections))
                        return # 成功，退出重试循环
                    except json.JSONDecodeError as json_e:
                        self.log(f"JSON解析错误: {json_e}")
                        self.log(f"尝试解析的JSON字符串 (前500字符): {json_str[:500]}...")
                        if attempt == max_retries - 1:
                            self.root.after(0, lambda: self.handle_ai_error(f"AI返回的JSON格式无效: {json_e}\n请检查日志获取更多信息。"))
                        # 不立即返回，允许重试
                        raise requests.exceptions.RequestException("JSON解析错误") # 触发重试
                
                elif response.status_code >= 500: # 服务器端错误，可以重试
                    self.log(f"DeepSeek API 服务器错误。状态码: {response.status_code}, 详情: {response.text[:200]}")
                    if attempt == max_retries - 1:
                         self.root.after(0, lambda r=response: self.handle_ai_error(f"AI请求失败(服务器错误)，错误码：{r.status_code}\n详情: {r.text[:200]}..."))
                    # 继续重试
                
                else: # 客户端错误 (4xx)，通常不可重试
                    error_detail = response.text
                    self.log(f"DeepSeek API 请求失败(客户端错误)。状态码: {response.status_code}, 详情: {error_detail}")
                    self.root.after(0, lambda r=response: self.handle_ai_error(f"AI请求失败(客户端错误)，错误码：{r.status_code}\n详情: {r.text[:200]}..."))
                    return # 不可重试的错误，直接返回

            except requests.exceptions.Timeout as timeout_e:
                self.log(f"DeepSeek API 请求超时 (尝试 {attempt + 1}/{max_retries}): {str(timeout_e)}")
                if attempt == max_retries - 1: # 如果是最后一次尝试
                    self.root.after(0, lambda: self.handle_ai_error(f"网络请求超时: {str(timeout_e)}"))
                # 继续重试
            except requests.exceptions.RequestException as req_e: # 其他网络相关错误
                self.log(f"DeepSeek API 请求时发生网络错误 (尝试 {attempt + 1}/{max_retries}): {str(req_e)}")
                if attempt == max_retries - 1:
                    self.root.after(0, lambda: self.handle_ai_error(f"网络请求错误: {str(req_e)}"))
                # 继续重试
            
            # 如果不是最后一次尝试，则等待后重试
            if attempt < max_retries - 1:
                wait_time = backoff_factor * (2 ** attempt) # 指数退避
                self.log(f"等待 {wait_time:.2f} 秒后重试...")
                time.sleep(wait_time)
            else: # 所有重试均失败
                self.log("所有重试均失败。")
                # 确保在所有重试失败后，如果之前没有调用handle_ai_error，这里会调用
                # (通常上面的逻辑会在最后一次尝试时调用，但作为保险)
                # self.root.after(0, lambda: self.handle_ai_error("AI分析失败，已达到最大重试次数。"))
                return # 结束函数执行

        # 如果循环结束仍未成功 (理论上应该在循环内通过 return 退出或在最后一次尝试的 except 块中处理)
        # 为确保状态被更新，如果意外到达这里
        self.root.after(0, lambda: self.handle_ai_error("AI分析失败，请检查网络连接和API Key。"))


    def import_ai_sections(self, sections_data):
        try:
            if not isinstance(sections_data, list): 
                self.handle_ai_error(f"AI返回的数据格式不正确，期望列表但得到 {type(sections_data)}。")
                return
            if self.document_sections:
                if messagebox.askyesno("确认", "是否清空现有文档内容，重新导入？"):
                    self.document_sections = []
                    self.current_section_id = None
                    self.section_title_var.set("")
                    self.section_content_text.delete(1.0, tk.END)
            imported_count = 0
            for section_item in sections_data:
                if not isinstance(section_item, dict):
                    self.log(f"跳过无效的章节项目: {section_item}")
                    continue
                self.add_section(level=section_item.get('level', 1), title=section_item.get('title', '未命名标题'), content=section_item.get('content', ''))
                imported_count +=1
            self.update_tree()
            target_tab_text = "文档内容"
            target_tab_index = next((i for i, tab_id in enumerate(self.notebook.tabs()) if self.notebook.tab(tab_id, "text") == target_tab_text), -1)
            if target_tab_index != -1: self.notebook.select(target_tab_index)
            self.ai_status_var.set(f"成功导入 {imported_count} 个章节")
            messagebox.showinfo("成功", f"已成功识别并导入 {imported_count} 个章节")
        except Exception as e:
            self.handle_ai_error(f"导入章节时出错: {str(e)}")
    
    def handle_ai_error(self, error_msg):
        self.ai_status_var.set("发生错误")
        self.log(f"AI错误: {error_msg}")
        messagebox.showerror("错误", error_msg)
        
    def on_tree_select(self, event):
        selected_items = self.tree.selection()
        if not selected_items: return
        section_id = selected_items[0]
        section = self.find_section_by_id(section_id)
        if section:
            if self.current_section_id and self.current_section_id != section_id: self.save_current_section()
            self.current_section_id = section_id
            self.section_title_var.set(section['title'])
            self.section_level_var.set(section['level'])
            self.section_content_text.delete(1.0, tk.END)
            self.section_content_text.insert(tk.END, section['content'])
    
    def save_current_section(self):
        if not self.current_section_id: return
        section = self.find_section_by_id(self.current_section_id)
        if section:
            new_title, new_level, new_content = self.section_title_var.get(), self.section_level_var.get(), self.section_content_text.get(1.0, tk.END).strip()
            if (section['title'] != new_title or section['level'] != new_level or section['content'] != new_content):
                section.update({'title': new_title, 'level': new_level, 'content': new_content})
                self.update_tree()
                self.log(f"章节 '{new_title}' 已保存。")
    
    def delete_section(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择要删除的章节")
            return
        section_id = selected[0]
        if messagebox.askyesno("确认", "确定要删除选中的章节吗？"):
            self.document_sections = [s for s in self.document_sections if s['id'] != section_id]
            self.update_tree()
            if self.current_section_id == section_id:
                self.current_section_id = None
                self.section_title_var.set("")
                self.section_level_var.set(1)
                self.section_content_text.delete(1.0, tk.END)
    
    def move_section(self, direction):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择要移动的章节")
            return
        section_id = selected[0]
        for i, section_item in enumerate(self.document_sections):
            if section_item['id'] == section_id:
                new_pos = i + direction
                if 0 <= new_pos < len(self.document_sections):
                    self.document_sections[i], self.document_sections[new_pos] = self.document_sections[new_pos], self.document_sections[i]
                    self.update_tree()
                    self.tree.selection_set(section_id)
                    self.tree.focus(section_id)
                break
    
    def create_font_settings(self, parent, r_idx, lbl_txt, fnt_var, sz_var, bld_var, clr_tk_var, clr_key): # Shortened var names
        frame = ttk.LabelFrame(parent, text=lbl_txt)
        frame.grid(row=r_idx, column=0, columnspan=3, sticky="ew", padx=10, pady=10)
        ttk.Label(frame, text="字体:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Combobox(frame, textvariable=fnt_var, values=["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "Times New Roman", "Arial"]).grid(row=0, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(frame, text="字号:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        ttk.Combobox(frame, textvariable=sz_var, values=[str(s) for s in [8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72]], width=5).grid(row=0, column=3, sticky="w", padx=5, pady=5)
        ttk.Checkbutton(frame, text="加粗", variable=bld_var).grid(row=0, column=4, sticky="w", padx=5, pady=5)
        ttk.Button(frame, text="颜色", command=lambda k=clr_key: self.choose_color(k)).grid(row=0, column=5, sticky="w", padx=5, pady=5)
        preview = tk.Canvas(frame, width=20, height=20, bg=clr_tk_var.get(), highlightthickness=1, highlightbackground="black")
        preview.grid(row=0, column=6, sticky="w", padx=5, pady=5)
        self.color_previews[clr_key] = preview
    
    def choose_color(self, color_key_name):
        color_tk_var = getattr(self, f"{color_key_name}_val", None)
        if not color_tk_var:
            self.log(f"错误：未找到颜色变量 {color_key_name}_val")
            return
        chosen_color = colorchooser.askcolor(initialcolor=color_tk_var.get(), title="选择颜色")
        if chosen_color and chosen_color[1]:
            color_tk_var.set(chosen_color[1])
            if color_key_name in self.color_previews: self.color_previews[color_key_name].config(bg=chosen_color[1])
            else: self.log(f"警告：未找到颜色预览 {color_key_name}")
    
    def log(self, message):
        if hasattr(self, 'log_text_widget') and self.log_text_widget:
            self.log_text_widget.config(state=tk.NORMAL)
            self.log_text_widget.insert(tk.END, message + "\n")
            self.log_text_widget.see(tk.END)
            self.log_text_widget.config(state=tk.DISABLED)
            self.root.update_idletasks()
        else: print(f"LOG (widget not ready): {message}")
    
    def update_progress(self, value):
        if self.progress_bar:
            self.progress_bar["value"] = value
            self.root.update_idletasks()
    
    def generate_document(self):
        self.save_current_section()
        if not self.filename.get(): messagebox.showerror("错误", "请输入文件名称"); return
        if not self.document_sections: messagebox.showerror("错误", "文档内容为空，请至少添加一个章节"); return
        threading.Thread(target=self.generate_document_thread, daemon=True).start()
    
    def generate_document_thread(self):
        try:
            if self.log_text_widget:
                self.log_text_widget.config(state=tk.NORMAL); self.log_text_widget.delete(1.0, tk.END); self.log_text_widget.config(state=tk.DISABLED)
            self.log("开始文档生成过程..."); self.update_progress(5)
            save_dir = filedialog.askdirectory(title="选择保存文档的文件夹")
            if not save_dir: self.log("用户取消了文件夹选择，操作终止。"); self.update_progress(0); return
            doc_path = os.path.join(save_dir, f"{self.filename.get().strip()}.docx")
            self.log(f"文档将保存至: {doc_path}"); self.update_progress(10)
            doc = Document(); self.log("已创建新文档..."); self.update_progress(15) # Renamed 'document' to 'doc'
            
            self.log("设置文档页面格式...")
            for sec in doc.sections:
                sec.page_height, sec.page_width = Inches(11.69), Inches(8.27)
                sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Inches(1), Inches(1), Inches(1), Inches(0.8)
            self.update_progress(20)
            
            self.log("创建文档样式...")
            self.create_style(doc, 'DocTitleStyle', '文档标题', self.title_font.get(), self.title_size.get(), self.title_bold.get(), self.title_color_val.get())
            self.update_progress(30)
            self.create_style(doc, 'Heading1Style', '一级标题', self.h1_font.get(), self.h1_size.get(), self.h1_bold.get(), self.h1_color_val.get(), level=1)
            self.update_progress(40)
            self.create_style(doc, 'Heading2Style', '二级标题', self.h2_font.get(), self.h2_size.get(), self.h2_bold.get(), self.h2_color_val.get(), level=2)
            self.update_progress(50)
            self.create_style(doc, 'Heading3Style', '三级标题', self.h3_font.get(), self.h3_size.get(), self.h3_bold.get(), self.h3_color_val.get(), level=3)
            self.update_progress(60)
            
            self.log("创建正文样式...")
            n_style = doc.styles['Normal'] # Renamed 'normal_style'
            n_style.font.name, n_style.font.size, n_style.font.bold = self.normal_font.get(), Pt(self.normal_size.get()), self.normal_bold.get()
            n_style.font.color.rgb = self.rgb_from_hex(self.normal_color_val.get())
            n_style.element.rPr.rFonts.set(qn('w:eastAsia'), self.normal_font.get())
            try: indent_chars = int(self.indent_entry.get())
            except ValueError: self.log("警告：首行缩进字符数无效，默认为0。"); indent_chars = 0
            n_style.paragraph_format.first_line_indent = Pt(indent_chars * self.normal_size.get() * 0.75)
            n_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            self.update_progress(70)
            
            self.log("添加文档内容..."); self.log("添加文档标题...")
            title_p = doc.add_paragraph(self.document_title.get(), style='DocTitleStyle') # Renamed 'title_para'
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER; self.update_progress(75)
            
            self.log("添加目录..."); self.add_toc(doc, self.toc_title.get()); doc.add_page_break(); self.update_progress(80)
            self.log("添加文档主体内容..."); self.add_user_document_content(doc); self.update_progress(90)
            
            self.log("保存文档..."); doc.save(doc_path); self.update_progress(100)
            self.log(f"文档已成功保存至 {doc_path}")
            self.log("请在Word中打开文档，右键点击目录，选择'更新域'或按F9更新目录。")
            self.root.after(0, lambda: messagebox.showinfo("成功", f"文档已成功生成并保存至:\n{doc_path}\n\n请在Word中打开文档后，右键点击目录，选择'更新域'更新目录。"))
        except Exception as e:
            err_msg = f"生成文档时发生错误:\n{str(e)}" # Renamed 'error_message'
            self.log(f"错误: {str(e)}")
            self.root.after(0, lambda em=err_msg: messagebox.showerror("错误", em)) # Renamed 'emsg'
            self.update_progress(0)
    
    def add_user_document_content(self, document): # Parameter name kept as 'document' for clarity with caller
        self.log("开始添加用户定义的文档内容...")
        for sec_item in self.document_sections: # Renamed 'section_item'
            self.log(f"添加章节: {sec_item['title']} (级别 {sec_item['level']})")
            style_name = 'Normal'
            if sec_item['level'] == 1: style_name = 'Heading1Style'
            elif sec_item['level'] == 2: style_name = 'Heading2Style'
            elif sec_item['level'] == 3: style_name = 'Heading3Style'
            try: document.add_paragraph(sec_item['title'], style=style_name)
            except KeyError: self.log(f"警告：样式 '{style_name}' 未找到，使用 Normal 样式替代。"); document.add_paragraph(sec_item['title'], style='Normal')
            if sec_item['content']:
                for line in sec_item['content'].splitlines():
                    if line.strip(): document.add_paragraph(line, style='Normal')
        self.log("所有用户定义的内容已添加完成")
    
    def create_style(self, document, style_id, style_name_ui, font_name, font_size_pt, is_bold, color_hex, level=None): # Parameter name kept
        try:
            self.log(f"创建样式: {style_name_ui}, 字体: {font_name}, 大小: {font_size_pt}pt")
            try: style = document.styles[style_id]
            except KeyError: style = document.styles.add_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
            style.name, style.hidden, style.quick_style = style_name_ui, False, True
            if level:
                try:
                    base_style_name = f'Heading {level}'
                    style.base_style = document.styles[base_style_name] if base_style_name in document.styles else document.styles['Normal']
                except KeyError: self.log(f"警告: 内置标题样式 'Heading {level}' 未找到，基于 Normal 创建。"); style.base_style = document.styles['Normal']
                style.next_paragraph_style = document.styles['Normal']
                pPr = style.element.get_or_add_pPr(); numPr = pPr.get_or_add_numPr(); ilvl = numPr.get_or_add_ilvl(); ilvl.val = level - 1
            font = style.font
            font.name, font.size, font.bold, font.color.rgb = font_name, Pt(font_size_pt), is_bold, self.rgb_from_hex(color_hex)
            rpr_fonts = font.element.rPr.rFonts
            rpr_fonts.set(qn('w:eastAsia'), font_name); rpr_fonts.set(qn('w:ascii'), font_name); rpr_fonts.set(qn('w:hAnsi'), font_name)
            p_fmt = style.paragraph_format # Renamed 'paragraph_format'
            if level: p_fmt.space_before, p_fmt.space_after = Pt(12 if level == 1 else 8), Pt(6 if level == 1 else 4)
            else: p_fmt.space_before, p_fmt.space_after = Pt(12), Pt(12)
            p_fmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
            if style_id == 'DocTitleStyle': p_fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif level: p_fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT
            return style
        except Exception as e: self.log(f"创建样式 '{style_name_ui}' 时出错: {str(e)}"); return None
    
    def rgb_from_hex(self, hex_color_str):
        hex_color_str = hex_color_str.lstrip('#')
        if len(hex_color_str) == 6: return RGBColor(int(hex_color_str[0:2],16), int(hex_color_str[2:4],16), int(hex_color_str[4:6],16))
        else: self.log(f"警告：无效的十六进制颜色 '{hex_color_str}'，使用黑色。"); return RGBColor(0,0,0)
    
    def add_toc(self, document, toc_main_title="目录"): # Parameter name kept
        try:
            self.log(f"添加目录标题: {toc_main_title}")
            p_title = document.add_paragraph(); p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_title = p_title.add_run(toc_main_title)
            run_title.font.name = self.h1_font.get()
            run_title.font.size = Pt(self.h1_size.get() if self.h1_size.get() > 16 else 16)
            run_title.font.bold = True
            run_title.font.element.rPr.rFonts.set(qn('w:eastAsia'), self.h1_font.get())
            p_title.paragraph_format.space_after = Pt(12)
            self.log("添加目录字段")
            p_toc_field = document.add_paragraph(); run_toc_field = p_toc_field.add_run()
            fldChar_begin = OxmlElement('w:fldChar'); fldChar_begin.set(qn('w:fldCharType'),'begin')
            instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'),'preserve'); instrText.text = r' TOC \o "1-3" \h \z \u '
            fldChar_separate = OxmlElement('w:fldChar'); fldChar_separate.set(qn('w:fldCharType'),'separate')
            fldChar_end = OxmlElement('w:fldChar'); fldChar_end.set(qn('w:fldCharType'),'end')
            r_elm = run_toc_field._r # Renamed 'r_element'
            r_elm.extend([fldChar_begin, instrText, fldChar_separate, fldChar_end]) # Use extend for list
            self.log("目录添加完成")
            return p_toc_field
        except Exception as e: self.log(f"添加目录时出错: {str(e)}"); return None

if __name__ == "__main__":
    root = tk.Tk()
    app = DocxFormatter(root)
    root.mainloop()
