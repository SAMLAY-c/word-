import tkinter as tk
from tkinter import ttk, colorchooser

# --- General UI Creation Functions ---
def create_font_settings_panel(parent_frame, label_text, 
                               style_font_var, style_size_var, style_bold_var, 
                               current_color_hex, color_hex_attr_name_in_config, # For identifying which color to update in config
                               choose_color_callback_with_args):
    """创建字体、字号、加粗、颜色的通用设置面板。"""
    frame = ttk.LabelFrame(parent_frame, text=label_text, padding="5 5 5 5")
    # frame.grid(row=row_index, column=0, columnspan=7, sticky="ew", padx=10, pady=5) # Grid in calling function
    
    ttk.Label(frame, text="字体:").grid(row=0, column=0, sticky=tk.W, padx=(0,5), pady=5)
    font_combo = ttk.Combobox(frame, textvariable=style_font_var, 
                              values=["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "Times New Roman", "Arial"], width=12)
    font_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
    
    ttk.Label(frame, text="字号:").grid(row=0, column=2, sticky=tk.W, padx=(10,5), pady=5)
    size_combo = ttk.Combobox(frame, textvariable=style_size_var, 
                              values=[8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72], width=5)
    size_combo.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
    
    bold_check = ttk.Checkbutton(frame, text="加粗", variable=style_bold_var)
    bold_check.grid(row=0, column=4, sticky=tk.W, padx=(10,5), pady=5)
    
    color_preview_canvas = tk.Canvas(frame, width=20, height=20, bg=current_color_hex, relief=tk.SUNKEN, borderwidth=1)
    color_preview_canvas.grid(row=0, column=5, sticky=tk.W, padx=(5,0), pady=5)

    # Pass necessary info to the callback
    color_button = ttk.Button(frame, text="颜色", 
                              command=lambda: choose_color_callback_with_args(color_hex_attr_name_in_config, color_preview_canvas, current_color_hex))
    color_button.grid(row=0, column=6, sticky=tk.W, padx=5, pady=5)
    
    return frame, color_preview_canvas # Return canvas for potential updates if needed

# --- Tab Creation Functions ---

def create_basic_settings_tab(parent_tab, style_config):
    frame = ttk.Frame(parent_tab, padding="10 10 10 10")
    frame.pack(expand=True, fill=tk.BOTH)

    ttk.Label(frame, text="文件名称(不需要.docx):", font=("Arial", 10)).grid(row=0, column=0, sticky="w", padx=10, pady=10)
    ttk.Entry(frame, textvariable=style_config.filename_var, width=40).grid(row=0, column=1, padx=10, pady=10, sticky=tk.EW)
    
    ttk.Label(frame, text="目录标题:", font=("Arial", 10)).grid(row=1, column=0, sticky="w", padx=10, pady=10)
    ttk.Entry(frame, textvariable=style_config.toc_title_var, width=40).grid(row=1, column=1, padx=10, pady=10, sticky=tk.EW)
    
    ttk.Label(frame, text="文档标题:", font=("Arial", 10)).grid(row=2, column=0, sticky="w", padx=10, pady=10)
    ttk.Entry(frame, textvariable=style_config.document_title_var, width=40).grid(row=2, column=1, padx=10, pady=10, sticky=tk.EW)
    
    frame.columnconfigure(1, weight=1)
    return frame

def create_content_tab(parent_tab, section_editor_vars, callbacks):
    paned_window = ttk.PanedWindow(parent_tab, orient=tk.HORIZONTAL)
    paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # Left side: Tree and buttons
    left_frame = ttk.Frame(paned_window, width=300) # Give it an initial width
    paned_window.add(left_frame, weight=1)
    left_frame.pack_propagate(False) # Prevent resizing based on children

    # Right side: Section editor
    right_frame = ttk.Frame(paned_window, width=600) # Give it an initial width
    paned_window.add(right_frame, weight=2)
    right_frame.pack_propagate(False)

    # --- Left Frame Content ---
    ttk.Label(left_frame, text="文档结构").pack(pady=(0,5), anchor=tk.W, padx=5)
    
    tree_frame = ttk.Frame(left_frame)
    tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0,5))

    tree = ttk.Treeview(tree_frame, selectmode='browse')
    tree["columns"] = ("title", "level")
    tree.column("#0", width=50, minwidth=40, stretch=tk.NO, anchor=tk.W)
    tree.column("title", width=150, minwidth=100, stretch=tk.YES, anchor=tk.W)
    tree.column("level", width=60, minwidth=50, stretch=tk.NO, anchor=tk.W)
    tree.heading("#0", text="序号", anchor=tk.W)
    tree.heading("title", text="标题", anchor=tk.W)
    tree.heading("level", text="级别", anchor=tk.W)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscrollcommand=tree_scroll.set)
    
    tree.bind('<<TreeviewSelect>>', callbacks['on_tree_select'])

    # Buttons below tree
    button_area = ttk.Frame(left_frame)
    button_area.pack(fill=tk.X, padx=5, pady=5)
    ttk.Button(button_area, text="添加", command=callbacks['on_add_section'], width=6).pack(side=tk.LEFT, padx=2)
    ttk.Button(button_area, text="编辑", command=callbacks['on_edit_section'], width=6).pack(side=tk.LEFT, padx=2)
    ttk.Button(button_area, text="删除", command=callbacks['on_delete_section'], width=6).pack(side=tk.LEFT, padx=2)
    ttk.Button(button_area, text="上移", command=callbacks['on_move_up'], width=6).pack(side=tk.LEFT, padx=2)
    ttk.Button(button_area, text="下移", command=callbacks['on_move_down'], width=6).pack(side=tk.LEFT, padx=2)

    # --- Right Frame Content (Section Editor) ---
    editor_frame = ttk.Frame(right_frame, padding="5 5 5 5")
    editor_frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(editor_frame, text="章节标题:").grid(row=0, column=0, sticky=tk.W, pady=(0,5))
    section_title_entry = ttk.Entry(editor_frame, textvariable=section_editor_vars['title_var'], width=50)
    section_title_entry.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(0,10))
    
    level_ui_frame = ttk.Frame(editor_frame)
    level_ui_frame.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(0,10))
    ttk.Label(level_ui_frame, text="章节级别:").pack(side=tk.LEFT, anchor=tk.W)
    ttk.Radiobutton(level_ui_frame, text="一级", variable=section_editor_vars['level_var'], value=1).pack(side=tk.LEFT, padx=(10,5))
    ttk.Radiobutton(level_ui_frame, text="二级", variable=section_editor_vars['level_var'], value=2).pack(side=tk.LEFT, padx=5)
    ttk.Radiobutton(level_ui_frame, text="三级", variable=section_editor_vars['level_var'], value=3).pack(side=tk.LEFT, padx=5)
    
    ttk.Label(editor_frame, text="章节内容:").grid(row=3, column=0, sticky=tk.W, pady=(0,5))
    text_area_frame = ttk.Frame(editor_frame)
    text_area_frame.grid(row=4, column=0, columnspan=2, sticky=tk.NSEW, pady=(0,10))
    text_area_frame.rowconfigure(0, weight=1)
    text_area_frame.columnconfigure(0, weight=1)

    section_content_text = tk.Text(text_area_frame, wrap=tk.WORD, height=10) # Initial height, will expand
    section_content_text.grid(row=0, column=0, sticky=tk.NSEW)
    content_scroll = ttk.Scrollbar(text_area_frame, orient="vertical", command=section_content_text.yview)
    content_scroll.grid(row=0, column=1, sticky=tk.NS)
    section_content_text.configure(yscrollcommand=content_scroll.set)
    
    save_button = ttk.Button(editor_frame, text="保存当前章节修改", command=callbacks['on_save_section'])
    save_button.grid(row=5, column=0, columnspan=2, sticky=tk.E, pady=5)

    editor_frame.columnconfigure(0, weight=1) # Allow title entry to expand
    editor_frame.rowconfigure(4, weight=1) # Allow text area to expand

    # Return references to widgets that main_app might need to interact with directly
    return paned_window, tree, section_title_entry, section_content_text


def create_title_style_settings_tab(parent_tab, style_config, choose_color_callback):
    frame = ttk.Frame(parent_tab, padding="10 10 10 10")
    frame.pack(expand=True, fill=tk.BOTH)

    # Doc Title
    doc_title_attrs = style_config.get_style_attributes("title")
    panel_doc_title, cv_doc_title = create_font_settings_panel(frame, "文档主标题样式", 
                                            doc_title_attrs["font_var"], doc_title_attrs["size_var"], 
                                            doc_title_attrs["bold_var"], doc_title_attrs["current_color_hex"], 
                                            doc_title_attrs["color_hex_attr"], choose_color_callback)
    panel_doc_title.grid(row=0, column=0, sticky=tk.EW, pady=(0,10))
    setattr(frame, 'cv_doc_title', cv_doc_title) # Store canvas reference

    # H1 Title
    h1_attrs = style_config.get_style_attributes("h1")
    panel_h1, cv_h1 = create_font_settings_panel(frame, "一级标题样式 (H1)", 
                                     h1_attrs["font_var"], h1_attrs["size_var"], 
                                     h1_attrs["bold_var"], h1_attrs["current_color_hex"], 
                                     h1_attrs["color_hex_attr"], choose_color_callback)
    panel_h1.grid(row=1, column=0, sticky=tk.EW, pady=(0,10))
    setattr(frame, 'cv_h1', cv_h1)

    # H2 Title
    h2_attrs = style_config.get_style_attributes("h2")
    panel_h2, cv_h2 = create_font_settings_panel(frame, "二级标题样式 (H2)", 
                                     h2_attrs["font_var"], h2_attrs["size_var"], 
                                     h2_attrs["bold_var"], h2_attrs["current_color_hex"], 
                                     h2_attrs["color_hex_attr"], choose_color_callback)
    panel_h2.grid(row=2, column=0, sticky=tk.EW, pady=(0,10))
    setattr(frame, 'cv_h2', cv_h2)

    # H3 Title
    h3_attrs = style_config.get_style_attributes("h3")
    panel_h3, cv_h3 = create_font_settings_panel(frame, "三级标题样式 (H3)", 
                                     h3_attrs["font_var"], h3_attrs["size_var"], 
                                     h3_attrs["bold_var"], h3_attrs["current_color_hex"], 
                                     h3_attrs["color_hex_attr"], choose_color_callback)
    panel_h3.grid(row=3, column=0, sticky=tk.EW, pady=(0,10))
    setattr(frame, 'cv_h3', cv_h3)
    
    frame.columnconfigure(0, weight=1)
    return frame

def create_normal_font_settings_tab(parent_tab, style_config, choose_color_callback):
    frame = ttk.Frame(parent_tab, padding="10 10 10 10")
    frame.pack(expand=True, fill=tk.BOTH)

    # Normal Text
    normal_attrs = style_config.get_style_attributes("normal")
    panel_normal, cv_normal = create_font_settings_panel(frame, "正文文本样式", 
                                           normal_attrs["font_var"], normal_attrs["size_var"], 
                                           normal_attrs["bold_var"], normal_attrs["current_color_hex"], 
                                           normal_attrs["color_hex_attr"], choose_color_callback)
    panel_normal.grid(row=0, column=0, sticky=tk.EW, pady=(0,10))
    setattr(frame, 'cv_normal', cv_normal)

    # Indent
    indent_frame = ttk.LabelFrame(frame, text="首行缩进", padding="5 5 5 5")
    indent_frame.grid(row=1, column=0, sticky=tk.EW, pady=(5,10))
    ttk.Label(indent_frame, text="缩进字符数 (基于正文字号):", width=25).grid(row=0, column=0, sticky=tk.W, padx=(0,5), pady=5)
    indent_entry = ttk.Entry(indent_frame, textvariable=style_config.indent_chars_var, width=8)
    indent_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

    frame.columnconfigure(0, weight=1)
    return frame, indent_entry # Return indent_entry if needed elsewhere, though its var is in style_config

def create_log_tab(parent_tab):
    frame = ttk.Frame(parent_tab, padding="10 10 10 10")
    frame.pack(expand=True, fill=tk.BOTH)
    
    log_text = tk.Text(frame, height=15, width=80, wrap=tk.WORD, state=tk.DISABLED)
    log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    scrollbar = ttk.Scrollbar(frame, command=log_text.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    log_text.config(yscrollcommand=scrollbar.set)
    
    return frame, log_text

def create_bottom_bar(parent_root, generate_doc_callback):
    bottom_frame = ttk.Frame(parent_root, padding="10 0 10 10") # Pad only bottom and sides
    bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)

    progress_bar = ttk.Progressbar(bottom_frame, orient="horizontal", length=200, mode="determinate") # Length is flexible
    progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,10), pady=(0,5))
    
    generate_button = ttk.Button(bottom_frame, text="生成文档", command=generate_doc_callback)
    generate_button.pack(side=tk.RIGHT, pady=(0,5))
    
    return bottom_frame, progress_bar, generate_button 