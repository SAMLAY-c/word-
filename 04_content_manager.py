import tkinter as tk
from tkinter import ttk, messagebox

class ContentManager:
    def __init__(self):
        self.document_sections = []
        self.tree_widget = None
        self.next_section_id_counter = 0

    def set_tree_widget(self, tree_widget):
        self.tree_widget = tree_widget

    def init_default_sections(self):
        """初始化一些默认的章节结构作为示例"""
        self.add_section(level=1, title="第一章 背景介绍", content="这里是背景介绍内容。")

    def _generate_section_id(self):
        new_id = f"section_{self.next_section_id_counter}"
        self.next_section_id_counter += 1
        return new_id

    def add_section(self, level, title, content=""):
        """添加一个新章节并更新UI"""
        section_id = self._generate_section_id()
        section = {
            'id': section_id,
            'level': level,
            'title': title,
            'content': content
        }
        self.document_sections.append(section)
        self.update_tree_ui()
        return section_id

    def update_tree_ui(self):
        """更新树形视图"""
        if not self.tree_widget:
            return
        # 清空树形视图
        for item in self.tree_widget.get_children():
            self.tree_widget.delete(item)
        
        # 填充树形视图
        for i, section in enumerate(self.document_sections):
            self.tree_widget.insert("", tk.END, section['id'], text=str(i+1), 
                                    values=(section['title'], f"级别{section['level']}"))

    def get_section_by_id(self, section_id):
        """根据ID查找章节"""
        for section in self.document_sections:
            if section['id'] == section_id:
                return section
        return None

    def edit_section_attributes(self, section_id, new_title, new_level):
        """编辑章节的标题和级别，并更新UI"""
        section = self.get_section_by_id(section_id)
        if section:
            section['title'] = new_title
            section['level'] = new_level
            self.update_tree_ui()
            return True
        return False

    def save_section_content(self, section_id, new_content):
        """保存指定章节的内容"""
        section = self.get_section_by_id(section_id)
        if section:
            section['content'] = new_content
            # 内容的改变不直接反映在树上，所以不需要调用 update_tree_ui()
            # 但如果需要在树上显示内容摘要等，则需要更新
            return True
        return False

    def delete_section(self, section_id):
        """删除选中的章节并更新UI"""
        section_to_delete = self.get_section_by_id(section_id)
        if section_to_delete:
            self.document_sections = [s for s in self.document_sections if s['id'] != section_id]
            self.update_tree_ui()
            return True
        return False

    def move_section(self, section_id, direction):
        """上移或下移章节并更新UI"""
        for i, section in enumerate(self.document_sections):
            if section['id'] == section_id:
                new_pos = i + direction
                if 0 <= new_pos < len(self.document_sections):
                    # 交换位置
                    self.document_sections[i], self.document_sections[new_pos] = \
                        self.document_sections[new_pos], self.document_sections[i]
                    self.update_tree_ui()
                    # 重新选中移动的项
                    if self.tree_widget:
                        self.tree_widget.selection_set(section_id)
                        self.tree_widget.focus(section_id) # 确保它可见
                    return True
                break
        return False

    def get_all_sections(self):
        """获取所有章节数据的列表副本"""
        return list(self.document_sections) # 返回副本以防外部修改 