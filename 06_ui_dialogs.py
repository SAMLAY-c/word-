import tkinter as tk
from tkinter import ttk, messagebox

def show_add_section_dialog(parent_root, confirm_callback):
    """打开添加章节对话框"""
    dialog = tk.Toplevel(parent_root)
    dialog.title("添加章节")
    dialog.geometry("400x180") # Adjusted height for better spacing
    dialog.transient(parent_root)
    dialog.grab_set()
    dialog.resizable(False, False)

    main_frame = ttk.Frame(dialog, padding="10 10 10 10")
    main_frame.pack(expand=True, fill=tk.BOTH)

    ttk.Label(main_frame, text="章节标题:").grid(row=0, column=0, sticky=tk.W, pady=5)
    title_var = tk.StringVar()
    title_entry = ttk.Entry(main_frame, textvariable=title_var, width=40)
    title_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
    title_entry.focus_set()

    ttk.Label(main_frame, text="章节级别:").grid(row=1, column=0, sticky=tk.W, pady=5)
    level_var = tk.IntVar(value=1)
    level_frame = ttk.Frame(main_frame)
    level_frame.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
    ttk.Radiobutton(level_frame, text="一级", variable=level_var, value=1).pack(side=tk.LEFT, padx=(0, 10))
    ttk.Radiobutton(level_frame, text="二级", variable=level_var, value=2).pack(side=tk.LEFT, padx=(0, 10))
    ttk.Radiobutton(level_frame, text="三级", variable=level_var, value=3).pack(side=tk.LEFT)

    button_frame = ttk.Frame(main_frame)
    button_frame.grid(row=2, column=0, columnspan=2, pady=(20, 0), sticky=tk.E)

    def on_confirm_internal():
        title = title_var.get().strip()
        level = level_var.get()
        if title:
            if confirm_callback(title, level):
                dialog.destroy()
        else:
            messagebox.showwarning("警告", "章节标题不能为空。", parent=dialog)
            title_entry.focus_set()

    confirm_button = ttk.Button(button_frame, text="确定", command=on_confirm_internal)
    confirm_button.pack(side=tk.RIGHT, padx=(10,0))
    cancel_button = ttk.Button(button_frame, text="取消", command=dialog.destroy)
    cancel_button.pack(side=tk.RIGHT)
    
    dialog.bind("<Return>", lambda event: on_confirm_internal())
    dialog.bind("<Escape>", lambda event: dialog.destroy())

    dialog.wait_window()


def show_edit_section_dialog(parent_root, current_title, current_level, confirm_callback):
    """打开编辑章节对话框"""
    dialog = tk.Toplevel(parent_root)
    dialog.title("编辑章节")
    dialog.geometry("400x180") # Adjusted height
    dialog.transient(parent_root)
    dialog.grab_set()
    dialog.resizable(False, False)

    main_frame = ttk.Frame(dialog, padding="10 10 10 10")
    main_frame.pack(expand=True, fill=tk.BOTH)

    ttk.Label(main_frame, text="章节标题:").grid(row=0, column=0, sticky=tk.W, pady=5)
    title_var = tk.StringVar(value=current_title)
    title_entry = ttk.Entry(main_frame, textvariable=title_var, width=40)
    title_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
    title_entry.focus_set()
    title_entry.selection_range(0, tk.END)

    ttk.Label(main_frame, text="章节级别:").grid(row=1, column=0, sticky=tk.W, pady=5)
    level_var = tk.IntVar(value=current_level)
    level_frame = ttk.Frame(main_frame)
    level_frame.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
    ttk.Radiobutton(level_frame, text="一级", variable=level_var, value=1).pack(side=tk.LEFT, padx=(0, 10))
    ttk.Radiobutton(level_frame, text="二级", variable=level_var, value=2).pack(side=tk.LEFT, padx=(0, 10))
    ttk.Radiobutton(level_frame, text="三级", variable=level_var, value=3).pack(side=tk.LEFT)

    button_frame = ttk.Frame(main_frame)
    button_frame.grid(row=2, column=0, columnspan=2, pady=(20, 0), sticky=tk.E)

    def on_confirm_internal():
        title = title_var.get().strip()
        level = level_var.get()
        if title:
            if confirm_callback(title, level):
                dialog.destroy()
        else:
            messagebox.showwarning("警告", "章节标题不能为空。", parent=dialog)
            title_entry.focus_set()

    confirm_button = ttk.Button(button_frame, text="确定", command=on_confirm_internal)
    confirm_button.pack(side=tk.RIGHT, padx=(10,0))
    cancel_button = ttk.Button(button_frame, text="取消", command=dialog.destroy)
    cancel_button.pack(side=tk.RIGHT)

    dialog.bind("<Return>", lambda event: on_confirm_internal())
    dialog.bind("<Escape>", lambda event: dialog.destroy())

    dialog.wait_window() 