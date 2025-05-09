import tkinter as tk

class StyleConfig:
    def __init__(self):
        # 文档标题样式变量
        self.title_font_var = tk.StringVar(value="黑体")
        self.title_size_var = tk.IntVar(value=22)
        self.title_color_hex = "#000000"  # 默认黑色
        self.title_bold_var = tk.BooleanVar(value=True)
        
        # 一级标题样式变量
        self.h1_font_var = tk.StringVar(value="黑体")
        self.h1_size_var = tk.IntVar(value=18)
        self.h1_color_hex = "#000000"
        self.h1_bold_var = tk.BooleanVar(value=True)
        
        # 二级标题样式变量
        self.h2_font_var = tk.StringVar(value="楷体")
        self.h2_size_var = tk.IntVar(value=16)
        self.h2_color_hex = "#000000"
        self.h2_bold_var = tk.BooleanVar(value=True)
        
        # 三级标题样式变量
        self.h3_font_var = tk.StringVar(value="宋体")
        self.h3_size_var = tk.IntVar(value=14)
        self.h3_color_hex = "#000000"
        self.h3_bold_var = tk.BooleanVar(value=True)
        
        # 正文样式变量
        self.normal_font_var = tk.StringVar(value="仿宋")
        self.normal_size_var = tk.IntVar(value=12)
        self.normal_color_hex = "#000000"
        self.normal_bold_var = tk.BooleanVar(value=False)
        self.indent_chars_var = tk.StringVar(value="2") # 首行缩进字符数

        # 基本设置变量
        self.filename_var = tk.StringVar()
        self.toc_title_var = tk.StringVar(value="目 录")
        self.document_title_var = tk.StringVar(value="公文标题示例")

    def get_style_attributes(self, style_type):
        """根据类型获取样式属性字典"""
        if style_type == "title":
            return {
                "font_var": self.title_font_var, "size_var": self.title_size_var,
                "bold_var": self.title_bold_var, "color_hex_attr": "title_color_hex",
                "current_color_hex": self.title_color_hex
            }
        elif style_type == "h1":
            return {
                "font_var": self.h1_font_var, "size_var": self.h1_size_var,
                "bold_var": self.h1_bold_var, "color_hex_attr": "h1_color_hex",
                "current_color_hex": self.h1_color_hex
            }
        elif style_type == "h2":
            return {
                "font_var": self.h2_font_var, "size_var": self.h2_size_var,
                "bold_var": self.h2_bold_var, "color_hex_attr": "h2_color_hex",
                "current_color_hex": self.h2_color_hex
            }
        elif style_type == "h3":
            return {
                "font_var": self.h3_font_var, "size_var": self.h3_size_var,
                "bold_var": self.h3_bold_var, "color_hex_attr": "h3_color_hex",
                "current_color_hex": self.h3_color_hex
            }
        elif style_type == "normal":
            return {
                "font_var": self.normal_font_var, "size_var": self.normal_size_var,
                "bold_var": self.normal_bold_var, "color_hex_attr": "normal_color_hex",
                "current_color_hex": self.normal_color_hex
            }
        return {}

    def set_color_hex(self, color_hex_attr_name, value):
        if hasattr(self, color_hex_attr_name):
            setattr(self, color_hex_attr_name, value)
        else:
            print(f"Warning: Attribute {color_hex_attr_name} not found in StyleConfig") 