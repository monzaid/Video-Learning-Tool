import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import re
import subprocess
import platform
from pathlib import Path
import urllib.parse

# 尝试导入tkinterdnd2用于文件拖拽功能
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

class SRTToTXTConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("SRT字幕转TXT工具")
        self.root.geometry("800x700")  # 增加高度以容纳搜索框和新选项
        
        # 存储文件信息：{文件路径: {'var': BooleanVar, 'frame': Frame}}
        self.file_items = {}
        
        # 文件覆盖选择状态：None=未选择, True=全部覆盖, False=全部不覆盖
        self.overwrite_all = None
        
        # 当前选中的文件（用于右键菜单高亮显示）
        self.selected_file = None
        
        # 输出文件夹路径
        self.output_folder = None
        
        # 文件添加顺序计数器（用于原序列排序）
        self.file_order_counter = 0
        
        # 拖拽框选相关变量
        self.drag_start_x = None
        self.drag_start_y = None
        self.drag_rect = None
        self.is_dragging = False
        self.drag_highlighted_items = set()  # 存储拖拽过程中高亮的文件项
        
        # 文件拖拽导入相关变量
        self.is_drag_over = False  # 是否有文件拖拽到区域上方
        
        # 功能选择相关变量
        self.function_mode = tk.StringVar(value="srt转txt")
        self.function_descriptions = {
            "srt转txt": "将SRT字幕文件转换为纯文本TXT文件，去除时间戳和序号，只保留字幕内容",
            "mp4转srt": "从MP4视频文件中提取音频并生成SRT字幕文件（需要语音识别功能）",
            "mp4语音翻译": "提取MP4视频中的语音内容并翻译为指定语言的文本",
            "srt文本翻译": "将SRT字幕文件中的文本内容翻译为其他语言",
            "txt文本翻译": "将TXT文本文件内容翻译为其他语言",
            "txt文本总结笔记": "对TXT文本文件内容进行智能总结，生成要点笔记"
        }
        
        # 创建GUI界面
        self.create_widgets()
    
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 第一行：文件选择按钮
        ttk.Button(file_frame, text="选择SRT文件",
                  command=self.select_files).grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(file_frame, text="选择文件夹",
                  command=self.select_folder).grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(file_frame, text="清空所有文件",
                  command=self.clear_all_files).grid(row=0, column=2, padx=(0, 10))
        
        ttk.Button(file_frame, text="删除选中文件",
                  command=self.remove_selected_files).grid(row=0, column=3)
        
        # 创建右对齐的功能选择框架
        function_frame = ttk.Frame(file_frame)
        function_frame.grid(row=0, column=4, sticky=tk.E, padx=(20, 0))
        
        # 功能选择下拉框
        ttk.Label(function_frame, text="功能：").pack(side=tk.LEFT, padx=(0, 5))
        
        self.function_combobox = ttk.Combobox(
            function_frame,
            textvariable=self.function_mode,
            values=list(self.function_descriptions.keys()),
            state="readonly",
            width=12
        )
        self.function_combobox.pack(side=tk.LEFT, padx=(0, 5))
        self.function_combobox.bind('<<ComboboxSelected>>', self.on_function_changed)
        
        # 帮助按钮
        self.help_button = ttk.Button(function_frame, text="?", width=3)
        self.help_button.pack(side=tk.LEFT)
        
        # 为帮助按钮创建悬浮提示
        self.create_function_help_tooltip(self.help_button)
        
        # 配置file_frame的列权重，让第4列可以扩展
        file_frame.columnconfigure(4, weight=1)
        
        # 第二行：递归选项
        self.recursive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(file_frame, text="递归搜索子文件夹中的SRT文件",
                       variable=self.recursive_var,
                       command=self.on_recursive_changed).grid(row=1, column=0, columnspan=5,
                                                              sticky=tk.W, pady=(10, 0))
       
       
        # 搜索区域（在文件列表上方）
        search_frame = ttk.Frame(main_frame)
        search_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        
        ttk.Label(search_frame, text="搜索文件：").pack(side=tk.LEFT)
        
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=(5, 5))
        self.search_var.trace('w', self.on_search_changed)
        
        ttk.Button(search_frame, text="清除", command=self.clear_search).pack(side=tk.LEFT, padx=(0, 10))
        
        # 搜索状态标签
        self.search_status_label = ttk.Label(search_frame, text="", foreground="gray")
        self.search_status_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # 正则表达式选项
        self.regex_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(search_frame, text="正则表达式", variable=self.regex_var,
                       command=self.on_search_changed).pack(side=tk.RIGHT)
        
        # 只处理搜索结果选项
        self.process_search_only_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(search_frame, text="只处理搜索结果", variable=self.process_search_only_var).pack(side=tk.RIGHT, padx=(0, 10))
        
        # 文件列表显示区域
        # 根据是否支持拖拽功能显示不同的标题
        if HAS_DND:
            list_title = "文件列表（勾选要处理的文件，支持拖拽SRT文件到此区域，支持Ctrl+V粘贴文件路径）"
        else:
            list_title = "文件列表（勾选要处理的文件，支持Ctrl+V粘贴文件路径）"
        list_frame = ttk.LabelFrame(main_frame, text=list_title, padding="10")
        list_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 创建滚动区域
        canvas = tk.Canvas(list_frame, height=200)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 存储canvas引用以便后续使用
        self.canvas = canvas
        
        # 绑定拖拽框选事件
        canvas.bind("<Button-1>", self.on_drag_start)
        canvas.bind("<B1-Motion>", self.on_drag_motion)
        canvas.bind("<ButtonRelease-1>", self.on_drag_end)
        
        # 绑定文件拖拽导入事件（如果支持的话）
        if HAS_DND:
            try:
                canvas.drop_target_register(DND_FILES)
                canvas.dnd_bind('<<DropEnter>>', self.on_file_drag_enter)
                canvas.dnd_bind('<<DropPosition>>', self.on_file_drag_motion)
                canvas.dnd_bind('<<DropLeave>>', self.on_file_drag_leave)
                canvas.dnd_bind('<<Drop>>', self.on_file_drop)
            except Exception as e:
                print(f"拖拽功能初始化失败: {e}")
        
        # 绑定点击事件来清除选中状态（在scrollable_frame上）
        def clear_selection_on_click(event):
            self.clear_selected_file()
        
        self.scrollable_frame.bind("<Button-1>", clear_selection_on_click)
        
        # 绑定鼠标滚轮事件
        self.bind_mousewheel(canvas)
        
        # 绑定键盘事件（Ctrl+V粘贴）
        # 需要让canvas获得焦点才能接收键盘事件
        canvas.focus_set()
        canvas.bind("<Control-v>", self.on_paste_files)
        canvas.bind("<Control-V>", self.on_paste_files)  # 大写V也支持
        
        # 让整个窗口也支持粘贴
        self.root.bind("<Control-v>", self.on_paste_files)
        self.root.bind("<Control-V>", self.on_paste_files)
        
        # 绑定Canvas右键菜单事件
        canvas.bind("<Button-3>", self.show_canvas_context_menu)
        # 也给scrollable_frame绑定右键事件，确保空白区域都能响应
        self.scrollable_frame.bind("<Button-3>", self.show_canvas_context_menu)
        
        # 全选/取消全选按钮和显示选项
        select_frame = ttk.Frame(list_frame)
        select_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        
        ttk.Button(select_frame, text="全选",
                  command=self.select_all_files).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(select_frame, text="取消全选",
                  command=self.deselect_all_files).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(select_frame, text="反向选择",
                  command=self.invert_selection).pack(side=tk.LEFT, padx=(0, 20))
        
        # 显示文件夹路径选项
        self.show_folder_path_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(select_frame, text="显示文件夹路径",
                       variable=self.show_folder_path_var,
                       command=self.on_show_path_changed).pack(side=tk.LEFT)
        
        # 排序选项框架（新的一行）
        sort_frame = ttk.Frame(list_frame)
        sort_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # 排序选项
        ttk.Label(sort_frame, text="排序：").pack(side=tk.LEFT)
        
        # 原序列选项
        self.sort_original_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(sort_frame, text="原序列", variable=self.sort_original_var,
                       command=self.on_sort_option_changed).pack(side=tk.LEFT, padx=(5, 5))
        
        # 文件名升序选项
        self.sort_name_asc_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(sort_frame, text="文件名升序", variable=self.sort_name_asc_var,
                       command=self.on_sort_option_changed).pack(side=tk.LEFT, padx=(0, 5))
        
        # 文件名降序选项
        self.sort_name_desc_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(sort_frame, text="文件名降序", variable=self.sort_name_desc_var,
                       command=self.on_sort_option_changed).pack(side=tk.LEFT, padx=(0, 5))
        
        # 已勾选优先选项
        self.sort_checked_first_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(sort_frame, text="已勾选优先", variable=self.sort_checked_first_var,
                       command=self.on_sort_option_changed).pack(side=tk.LEFT, padx=(0, 5))
        
        # 未勾选优先选项
        self.sort_unchecked_first_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(sort_frame, text="未勾选优先", variable=self.sort_unchecked_first_var,
                       command=self.on_sort_option_changed).pack(side=tk.LEFT, padx=(0, 5))
        
        # 输出选项
        option_frame = ttk.LabelFrame(main_frame, text="输出选项", padding="10")
        option_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 输出模式选择
        self.output_mode = tk.StringVar(value="separate")
        ttk.Radiobutton(option_frame, text="输出对应文件（每个SRT对应一个TXT）",
                       variable=self.output_mode, value="separate",
                       command=self.on_output_mode_changed).grid(row=0, column=0, sticky=tk.W)
        
        # 合成输出选项行
        merge_frame = ttk.Frame(option_frame)
        merge_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        ttk.Radiobutton(merge_frame, text="合成输出一个文件",
                       variable=self.output_mode, value="merge",
                       command=self.on_output_mode_changed).pack(side=tk.LEFT)
        
        # 合成输出文件名显示选项（放在合成输出右边，初始隐藏）
        self.show_merge_path_var = tk.BooleanVar(value=False)
        self.show_merge_path_checkbox = ttk.Checkbutton(
            merge_frame,
            text="合成输出时显示被合成文件的绝对路径",
            variable=self.show_merge_path_var
        )
        # 初始状态下不显示
        self.show_merge_path_checkbox.pack_forget()
        
        # 按文件夹合并选项（放在显示路径选项右边，初始隐藏）
        self.merge_by_folder_var = tk.BooleanVar(value=False)
        self.merge_by_folder_checkbox = ttk.Checkbutton(
            merge_frame,
            text="按文件夹合并（每个文件夹生成一个summary.txt）",
            variable=self.merge_by_folder_var
        )
        # 初始状态下不显示
        self.merge_by_folder_checkbox.pack_forget()
        
        # 输出文件夹选项
        output_folder_frame = ttk.Frame(option_frame)
        output_folder_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 输出到同一个文件夹里选项
        self.output_to_same_folder_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(output_folder_frame, text="输出到同一个文件夹里",
                       variable=self.output_to_same_folder_var,
                       command=self.on_output_folder_changed).pack(side=tk.LEFT)
        
        # 选择输出文件夹按钮
        self.select_output_folder_btn = ttk.Button(output_folder_frame, text="选择输出文件夹",
                                                  command=self.select_output_folder)
        self.select_output_folder_btn.pack(side=tk.LEFT, padx=(10, 0))
        self.select_output_folder_btn.config(state=tk.DISABLED)  # 初始禁用
        
        # 输出文件夹路径显示
        output_path_frame = ttk.Frame(option_frame)
        output_path_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        ttk.Label(output_path_frame, text="输出文件夹：").pack(side=tk.LEFT)
        self.output_folder_label = ttk.Label(output_path_frame, text="未选择", foreground="gray")
        self.output_folder_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # 转换按钮
        convert_frame = ttk.Frame(main_frame)
        convert_frame.grid(row=4, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(convert_frame, text="转换选中文件",
                  command=self.convert_selected_files, style="Accent.TButton").pack()
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
    
    def on_output_mode_changed(self):
        """输出模式变化时的回调"""
        if self.output_mode.get() == "merge":
            # 显示合成输出文件名显示选项（在合成输出右边）
            self.show_merge_path_checkbox.pack(side=tk.LEFT, padx=(20, 0))
        else:
            # 隐藏合成输出文件名显示选项
            self.show_merge_path_checkbox.pack_forget()
    
    def on_output_folder_changed(self):
        """输出文件夹选项变化时的回调"""
        if self.output_to_same_folder_var.get():
            # 启用选择输出文件夹按钮
            self.select_output_folder_btn.config(state=tk.NORMAL)
        else:
            # 禁用选择输出文件夹按钮，重置输出文件夹
            self.select_output_folder_btn.config(state=tk.DISABLED)
            self.output_folder = None
            self.output_folder_label.config(text="未选择", foreground="gray")
    
    def select_output_folder(self):
        """选择输出文件夹"""
        folder = filedialog.askdirectory(title="选择输出文件夹")
        
        if folder:
            self.output_folder = folder
            # 显示选择的文件夹路径（使用正确的路径分隔符）
            display_path = os.path.normpath(folder)
            self.output_folder_label.config(text=display_path, foreground="black")
        else:
            # 用户取消选择，保持当前状态
            pass
    
    def select_files(self):
        """选择SRT文件"""
        files = filedialog.askopenfilenames(
            title="选择SRT字幕文件",
            filetypes=[("SRT文件", "*.srt"), ("所有文件", "*.*")]
        )
        
        for file in files:
            if file not in self.file_items:
                # 对于单独选择的文件，文件夹路径就是文件所在目录
                folder_path = os.path.dirname(file)
                self.add_file_item(file, folder_path)
    
    def on_recursive_changed(self):
        """递归选项变化时的回调"""
        if self.recursive_var.get():
            # 显示按文件夹合并选项（在显示路径选项右边）
            self.merge_by_folder_checkbox.pack(side=tk.LEFT, padx=(20, 0))
        else:
            # 隐藏按文件夹合并选项
            self.merge_by_folder_checkbox.pack_forget()
            self.merge_by_folder_var.set(False)
    
    def select_folder(self):
        """选择文件夹并获取其中所有SRT文件"""
        folder = filedialog.askdirectory(title="选择包含SRT文件的文件夹")
        
        if folder:
            srt_files = []
            
            if self.recursive_var.get():
                # 递归搜索子文件夹
                for root, dirs, files in os.walk(folder):
                    for file in files:
                        if file.lower().endswith('.srt'):
                            full_path = os.path.join(root, file)
                            if full_path not in self.file_items:
                                srt_files.append(full_path)
                                self.add_file_item(full_path, root)  # 传递文件夹路径
            else:
                # 只搜索当前文件夹
                for file in os.listdir(folder):
                    if file.lower().endswith('.srt'):
                        full_path = os.path.join(folder, file)
                        if full_path not in self.file_items:
                            srt_files.append(full_path)
                            self.add_file_item(full_path, folder)  # 传递文件夹路径
            
            if srt_files:
                search_type = "递归搜索" if self.recursive_var.get() else "当前文件夹"
                messagebox.showinfo("成功", f"通过{search_type}找到并添加了 {len(srt_files)} 个SRT文件")
            else:
                search_type = "递归搜索" if self.recursive_var.get() else "当前文件夹"
                messagebox.showwarning("警告", f"通过{search_type}没有找到新的SRT文件")
    
    def add_file_item(self, file_path, folder_path=None):
        """添加文件项到列表"""
        # 创建复选框变量
        var = tk.BooleanVar(value=True)  # 默认选中
        
        # 创建文件项框架
        item_frame = ttk.Frame(self.scrollable_frame)
        item_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # 配置框架的列权重，让文本可以扩展
        item_frame.columnconfigure(1, weight=1)
        
        # 创建复选框（不包含文本）
        checkbox = ttk.Checkbutton(item_frame, variable=var)
        checkbox.grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        # 创建文件名标签，支持自动换行和工具提示
        # 根据显示选项决定显示内容
        if self.show_folder_path_var.get():
            # 显示完整路径，确保使用正确的路径分隔符
            display_text = os.path.normpath(file_path)
        else:
            display_text = os.path.basename(file_path)  # 只显示文件名
        
        file_label = tk.Label(
            item_frame,
            text=display_text,
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=700,  # 增加换行宽度以适应路径
            bg=self.root.cget('bg')  # 使用窗口背景色
        )
        file_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        # 添加工具提示显示完整路径，确保使用正确的路径分隔符
        self.create_tooltip(file_label, os.path.normpath(file_path))
        
        # 绑定标签点击事件来切换复选框状态
        def toggle_checkbox(event):
            var.set(not var.get())
            # 左键点击时清除选中状态
            self.clear_selected_file()
        
        file_label.bind("<Button-1>", toggle_checkbox)
        
        # 绑定右键菜单
        def show_context_menu(event):
            self.show_file_context_menu(event, file_path)
        
        file_label.bind("<Button-3>", show_context_menu)  # 右键点击
        
        # 存储文件信息，包括文件夹路径和原始顺序
        self.file_items[file_path] = {
            'var': var,
            'frame': item_frame,
            'checkbox': checkbox,
            'label': file_label,
            'folder': folder_path or os.path.dirname(file_path),  # 存储文件夹路径
            'order': self.file_order_counter  # 存储原始添加顺序
        }
        
        # 增加顺序计数器
        self.file_order_counter += 1
        
        # 更新滚动区域
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def create_tooltip(self, widget, text):
        """为控件创建工具提示"""
        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            label = tk.Label(
                tooltip,
                text=text,
                background="lightyellow",
                relief="solid",
                borderwidth=1,
                wraplength=400
            )
            label.pack()
            
            widget.tooltip = tooltip
        
        def on_leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
        
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
    
    def bind_mousewheel(self, canvas):
        """绑定鼠标滚轮事件到Canvas"""
        def on_mousewheel(event):
            # Windows和Linux系统的滚轮事件处理
            if event.delta:
                # Windows系统
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                # Linux系统
                if event.num == 4:
                    canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    canvas.yview_scroll(1, "units")
        
        def bind_to_mousewheel(event):
            # 当鼠标进入Canvas区域时绑定滚轮事件
            canvas.bind_all("<MouseWheel>", on_mousewheel)  # Windows
            canvas.bind_all("<Button-4>", on_mousewheel)    # Linux向上滚动
            canvas.bind_all("<Button-5>", on_mousewheel)    # Linux向下滚动
        
        def unbind_from_mousewheel(event):
            # 当鼠标离开Canvas区域时解绑滚轮事件
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")
        
        # 绑定鼠标进入和离开事件
        canvas.bind('<Enter>', bind_to_mousewheel)
        canvas.bind('<Leave>', unbind_from_mousewheel)
        
        # 同时为scrollable_frame绑定事件，确保在内容区域也能滚动
        def bind_frame_mousewheel(event):
            canvas.bind_all("<MouseWheel>", on_mousewheel)
            canvas.bind_all("<Button-4>", on_mousewheel)
            canvas.bind_all("<Button-5>", on_mousewheel)
        
        def unbind_frame_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")
        
        self.scrollable_frame.bind('<Enter>', bind_frame_mousewheel)
        self.scrollable_frame.bind('<Leave>', unbind_frame_mousewheel)
    
    def on_show_path_changed(self):
        """显示路径选项变化时的回调"""
        # 更新所有文件项的显示文本
        for file_path, item_info in self.file_items.items():
            self.update_file_display_text(file_path, item_info)
        # 重新执行排序（因为排序依据可能发生变化）
        self.sort_file_list()
        # 重新应用搜索过滤
        self.filter_file_list()
    
    def on_sort_option_changed(self):
        """排序选项变化时的回调，处理互斥逻辑"""
        # 防止递归调用
        if hasattr(self, '_updating_sort_options') and self._updating_sort_options:
            return
        
        self._updating_sort_options = True
        
        try:
            # 获取当前所有选项的状态
            original_checked = self.sort_original_var.get()
            name_asc_checked = self.sort_name_asc_var.get()
            name_desc_checked = self.sort_name_desc_var.get()
            checked_first_checked = self.sort_checked_first_var.get()
            unchecked_first_checked = self.sort_unchecked_first_var.get()
            
            # 初始化上次状态（如果不存在）
            if not hasattr(self, '_last_sort_states'):
                self._last_sort_states = {
                    'original': True,
                    'name_asc': False,
                    'name_desc': False,
                    'checked_first': False,
                    'unchecked_first': False
                }
            
            current_states = {
                'original': original_checked,
                'name_asc': name_asc_checked,
                'name_desc': name_desc_checked,
                'checked_first': checked_first_checked,
                'unchecked_first': unchecked_first_checked
            }
            
            # 找出状态发生变化的选项
            changed_options = []
            for option, current_state in current_states.items():
                if current_state != self._last_sort_states[option]:
                    changed_options.append((option, current_state))
            
            # 处理互斥逻辑
            for option, is_selected in changed_options:
                if is_selected:  # 如果选项被选中
                    if option == 'original':
                        # 原序列与其他所有选项互斥
                        self.sort_name_asc_var.set(False)
                        self.sort_name_desc_var.set(False)
                        self.sort_checked_first_var.set(False)
                        self.sort_unchecked_first_var.set(False)
                    elif option == 'name_asc':
                        # 文件名升序与原序列和文件名降序互斥
                        self.sort_original_var.set(False)
                        self.sort_name_desc_var.set(False)
                    elif option == 'name_desc':
                        # 文件名降序与原序列和文件名升序互斥
                        self.sort_original_var.set(False)
                        self.sort_name_asc_var.set(False)
                    elif option == 'checked_first':
                        # 已勾选优先与原序列和未勾选优先互斥
                        self.sort_original_var.set(False)
                        self.sort_unchecked_first_var.set(False)
                    elif option == 'unchecked_first':
                        # 未勾选优先与原序列和已勾选优先互斥
                        self.sort_original_var.set(False)
                        self.sort_checked_first_var.set(False)
            
            # 检查是否没有任何选项被选中，如果是则默认选择原序列
            if not any([self.sort_original_var.get(), self.sort_name_asc_var.get(),
                       self.sort_name_desc_var.get(), self.sort_checked_first_var.get(),
                       self.sort_unchecked_first_var.get()]):
                self.sort_original_var.set(True)
            
            # 更新上次状态
            self._last_sort_states = {
                'original': self.sort_original_var.get(),
                'name_asc': self.sort_name_asc_var.get(),
                'name_desc': self.sort_name_desc_var.get(),
                'checked_first': self.sort_checked_first_var.get(),
                'unchecked_first': self.sort_unchecked_first_var.get()
            }
            
            # 执行排序
            self.sort_file_list()
            # 重新应用搜索过滤以保持搜索结果
            self.filter_file_list()
            
        finally:
            self._updating_sort_options = False
    
    def sort_file_list(self):
        """根据选择的排序方式对文件列表进行排序，实现分层排序逻辑"""
        
        # 获取所有文件项
        all_items = list(self.file_items.items())
        
        # 检查是否选择了勾选状态排序
        has_checked_sort = self.sort_checked_first_var.get() or self.sort_unchecked_first_var.get()
        # 检查是否选择了文件名排序
        has_name_sort = self.sort_name_asc_var.get() or self.sort_name_desc_var.get()
        
        if has_checked_sort and has_name_sort:
            # 分层排序：先按勾选状态分组，然后在每个组内按文件名排序
            
            # 分离已勾选和未勾选的文件
            checked_items = []
            unchecked_items = []
            
            for file_path, item_info in all_items:
                if item_info['var'].get():
                    checked_items.append((file_path, item_info))
                else:
                    unchecked_items.append((file_path, item_info))
            
            # 定义文件名排序函数
            def sort_by_filename(items):
                if self.sort_name_asc_var.get():
                    # 文件名升序
                    if self.show_folder_path_var.get():
                        # 按绝对路径排序
                        return sorted(items, key=lambda x: (x[0].lower(), x[1]['order']))
                    else:
                        # 按文件名排序
                        return sorted(items, key=lambda x: (os.path.basename(x[0]).lower(), x[1]['order']))
                elif self.sort_name_desc_var.get():
                    # 文件名降序
                    if self.show_folder_path_var.get():
                        # 按绝对路径排序
                        return sorted(items, key=lambda x: (x[0].lower(), x[1]['order']), reverse=True)
                    else:
                        # 按文件名排序
                        return sorted(items, key=lambda x: (os.path.basename(x[0]).lower(), x[1]['order']), reverse=True)
                else:
                    # 如果没有文件名排序，按原序列排序
                    return sorted(items, key=lambda x: x[1]['order'])
            
            # 对每个组内的文件按文件名排序
            sorted_checked = sort_by_filename(checked_items)
            sorted_unchecked = sort_by_filename(unchecked_items)
            
            # 根据勾选状态排序选项决定最终顺序
            if self.sort_checked_first_var.get():
                # 已勾选优先：已勾选文件在前，未勾选文件在后
                sorted_items = sorted_checked + sorted_unchecked
            else:  # self.sort_unchecked_first_var.get()
                # 未勾选优先：未勾选文件在前，已勾选文件在后
                sorted_items = sorted_unchecked + sorted_checked
                
        elif has_checked_sort:
            # 仅按勾选状态排序
            if self.sort_checked_first_var.get():
                # 已勾选优先
                sorted_items = sorted(all_items, key=lambda x: (0 if x[1]['var'].get() else 1, x[1]['order']))
            else:  # self.sort_unchecked_first_var.get()
                # 未勾选优先
                sorted_items = sorted(all_items, key=lambda x: (1 if x[1]['var'].get() else 0, x[1]['order']))
                
        elif has_name_sort:
            # 仅按文件名排序
            if self.sort_name_asc_var.get():
                # 文件名升序
                if self.show_folder_path_var.get():
                    # 按绝对路径排序
                    sorted_items = sorted(all_items, key=lambda x: (x[0].lower(), x[1]['order']))
                else:
                    # 按文件名排序
                    sorted_items = sorted(all_items, key=lambda x: (os.path.basename(x[0]).lower(), x[1]['order']))
            else:  # self.sort_name_desc_var.get()
                # 文件名降序
                if self.show_folder_path_var.get():
                    # 按绝对路径排序
                    sorted_items = sorted(all_items, key=lambda x: (x[0].lower(), x[1]['order']), reverse=True)
                else:
                    # 按文件名排序
                    sorted_items = sorted(all_items, key=lambda x: (os.path.basename(x[0]).lower(), x[1]['order']), reverse=True)
                
        else:
            # 原序列排序
            sorted_items = sorted(all_items, key=lambda x: x[1]['order'])
        
        # 重新排列界面中的文件项
        for file_path, item_info in sorted_items:
            # 将frame移到最后（这会重新排列顺序）
            item_info['frame'].pack_forget()
            item_info['frame'].pack(fill=tk.X, padx=5, pady=2)
        
        # 更新滚动区域
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def on_search_changed(self, *args):
        """搜索框内容变化时的回调"""
        self.filter_file_list()
    
    def clear_search(self):
        """清除搜索框内容"""
        self.search_var.set("")
        self.filter_file_list()
    
    def filter_file_list(self):
        """根据搜索条件过滤文件列表"""
        search_text = self.search_var.get().strip()
        
        if not search_text:
            # 如果搜索框为空，显示所有文件
            for file_path, item_info in self.file_items.items():
                item_info['frame'].pack(fill=tk.X, padx=5, pady=2)
            self.search_status_label.config(text="")
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            return
        
        matched_count = 0
        total_count = len(self.file_items)
        
        try:
            for file_path, item_info in self.file_items.items():
                # 根据显示路径选项决定搜索的文本
                if self.show_folder_path_var.get():
                    search_target = os.path.normpath(file_path)
                else:
                    search_target = os.path.basename(file_path)
                
                # 根据正则表达式选项进行匹配
                if self.regex_var.get():
                    # 使用正则表达式搜索
                    import re
                    if re.search(search_text, search_target, re.IGNORECASE):
                        item_info['frame'].pack(fill=tk.X, padx=5, pady=2)
                        matched_count += 1
                    else:
                        item_info['frame'].pack_forget()
                else:
                    # 使用普通文本搜索（不区分大小写）
                    if search_text.lower() in search_target.lower():
                        item_info['frame'].pack(fill=tk.X, padx=5, pady=2)
                        matched_count += 1
                    else:
                        item_info['frame'].pack_forget()
            
            # 更新搜索状态
            self.search_status_label.config(
                text=f"找到 {matched_count}/{total_count} 个文件",
                foreground="blue" if matched_count > 0 else "red"
            )
            
        except re.error as e:
            # 正则表达式错误
            self.search_status_label.config(
                text=f"正则表达式错误: {str(e)}",
                foreground="red"
            )
            # 显示所有文件
            for file_path, item_info in self.file_items.items():
                item_info['frame'].pack(fill=tk.X, padx=5, pady=2)
        
        except Exception as e:
            # 其他错误
            self.search_status_label.config(
                text=f"搜索错误: {str(e)}",
                foreground="red"
            )
        
        # 更新滚动区域
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def update_file_display_text(self, file_path, item_info):
        """更新文件项的显示文本"""
        if self.show_folder_path_var.get():
            # 显示绝对路径，确保使用正确的路径分隔符
            display_text = os.path.normpath(file_path)
        else:
            # 只显示文件名
            display_text = os.path.basename(file_path)
        
        # 更新标签文本
        item_info['label'].config(text=display_text)
    
    def clear_all_files(self):
        """清空所有文件"""
        if self.file_items:
            result = messagebox.askyesno("确认清空", "确定要清空文件列表吗？")
            if result:
                for file_path in list(self.file_items.keys()):
                    self.file_items[file_path]['frame'].destroy()
                self.file_items.clear()
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def remove_selected_files(self):
        """删除选中的文件"""
        if self.file_items:
            result = messagebox.askyesno("确认删除", "确定要删除选择的文件吗？")
            if result:
                to_remove = []
                for file_path, item_info in self.file_items.items():
                    if item_info['var'].get():
                        to_remove.append(file_path)
                
                if not to_remove:
                    messagebox.showwarning("警告", "请先勾选要删除的文件")
                    return
                
                for file_path in to_remove:
                    self.file_items[file_path]['frame'].destroy()
                    del self.file_items[file_path]
                
                self.canvas.configure(scrollregion=self.canvas.bbox("all"))
                messagebox.showinfo("成功", f"已删除 {len(to_remove)} 个文件")
    
    def select_all_files(self):
        """全选所有显示的文件"""
        for file_path, item_info in self.file_items.items():
            # 只选择当前显示的文件（未被搜索过滤掉的）
            if item_info['frame'].winfo_viewable():
                item_info['var'].set(True)
    
    def deselect_all_files(self):
        """取消全选所有显示的文件"""
        for file_path, item_info in self.file_items.items():
            # 只取消选择当前显示的文件（未被搜索过滤掉的）
            if item_info['frame'].winfo_viewable():
                item_info['var'].set(False)
    
    def invert_selection(self):
        """反向选择所有显示的文件"""
        for file_path, item_info in self.file_items.items():
            # 只反向选择当前显示的文件（未被搜索过滤掉的）
            if item_info['frame'].winfo_viewable():
                current_value = item_info['var'].get()
                item_info['var'].set(not current_value)
    
    def on_drag_start(self, event):
        """开始拖拽选择"""
        # 清除之前的选中状态
        self.clear_selected_file()
        
        # 记录拖拽起始位置
        self.drag_start_x = self.canvas.canvasx(event.x)
        self.drag_start_y = self.canvas.canvasy(event.y)
        self.is_dragging = True
        
        # 清除之前的高亮
        self.clear_drag_highlights()
    
    def on_drag_motion(self, event):
        """拖拽过程中的处理"""
        if not self.is_dragging:
            return
        
        # 获取当前鼠标位置
        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        
        # 删除之前的选择框
        if self.drag_rect:
            self.canvas.delete(self.drag_rect)
        
        # 绘制新的选择框
        self.drag_rect = self.canvas.create_rectangle(
            self.drag_start_x, self.drag_start_y, current_x, current_y,
            outline="blue", width=2, fill="lightblue", stipple="gray25"
        )
        
        # 更新文件项的高亮状态
        self.update_drag_highlights(self.drag_start_x, self.drag_start_y, current_x, current_y)
    
    def on_drag_end(self, event):
        """结束拖拽选择"""
        if not self.is_dragging:
            return
        
        # 获取最终位置
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        
        # 执行反选操作
        self.apply_drag_selection(self.drag_start_x, self.drag_start_y, end_x, end_y)
        
        # 清理拖拽状态
        if self.drag_rect:
            self.canvas.delete(self.drag_rect)
            self.drag_rect = None
        
        self.clear_drag_highlights()
        self.is_dragging = False
        self.drag_start_x = None
        self.drag_start_y = None
    
    def update_drag_highlights(self, x1, y1, x2, y2):
        """更新拖拽过程中的文件项高亮"""
        # 确保坐标顺序正确
        min_x, max_x = min(x1, x2), max(x1, x2)
        min_y, max_y = min(y1, y2), max(y1, y2)
        
        # 清除之前的高亮
        self.clear_drag_highlights()
        
        # 检查每个文件项是否在选择区域内
        for file_path, item_info in self.file_items.items():
            frame = item_info['frame']
            if not frame.winfo_viewable():
                continue
            
            try:
                # 获取文件项相对于scrollable_frame的位置
                frame_x = frame.winfo_x()
                frame_y = frame.winfo_y()
                frame_width = frame.winfo_width()
                frame_height = frame.winfo_height()
                
                # 转换为Canvas坐标系
                # 获取scrollable_frame在Canvas中的位置
                canvas_window = self.canvas.find_all()[0] if self.canvas.find_all() else None
                if canvas_window:
                    scroll_x, scroll_y = self.canvas.coords(canvas_window)
                    canvas_frame_x = frame_x + scroll_x
                    canvas_frame_y = frame_y + scroll_y
                    
                    # 检查是否与选择框相交
                    if (canvas_frame_x < max_x and canvas_frame_x + frame_width > min_x and
                        canvas_frame_y < max_y and canvas_frame_y + frame_height > min_y):
                        # 添加高亮效果 - 对label设置背景色，因为ttk.Frame不支持background属性
                        frame.configure(relief="solid", borderwidth=2)
                        label = self.file_items[file_path]['label']
                        label.configure(background="lightblue")
                        self.drag_highlighted_items.add(file_path)
            except tk.TclError:
                # 如果widget已被销毁，跳过
                continue
    
    def clear_drag_highlights(self):
        """清除拖拽高亮效果"""
        for file_path in self.drag_highlighted_items:
            if file_path in self.file_items:
                frame = self.file_items[file_path]['frame']
                label = self.file_items[file_path]['label']
                try:
                    # 重置边框样式
                    frame.configure(relief="flat", borderwidth=0)
                    # 恢复label的原始背景色
                    label.configure(background=self.root.cget('bg'))
                except tk.TclError:
                    # 如果配置失败，忽略错误
                    pass
        self.drag_highlighted_items.clear()
    
    def apply_drag_selection(self, x1, y1, x2, y2):
        """应用拖拽选择的反选操作"""
        # 确保坐标顺序正确
        min_x, max_x = min(x1, x2), max(x1, x2)
        min_y, max_y = min(y1, y2), max(y1, y2)
        
        # 对选择区域内的文件执行反选
        for file_path, item_info in self.file_items.items():
            frame = item_info['frame']
            if not frame.winfo_viewable():
                continue
            
            try:
                # 获取文件项相对于scrollable_frame的位置
                frame_x = frame.winfo_x()
                frame_y = frame.winfo_y()
                frame_width = frame.winfo_width()
                frame_height = frame.winfo_height()
                
                # 转换为Canvas坐标系
                # 获取scrollable_frame在Canvas中的位置
                canvas_window = self.canvas.find_all()[0] if self.canvas.find_all() else None
                if canvas_window:
                    scroll_x, scroll_y = self.canvas.coords(canvas_window)
                    canvas_frame_x = frame_x + scroll_x
                    canvas_frame_y = frame_y + scroll_y
                    
                    # 检查是否与选择框相交
                    if (canvas_frame_x < max_x and canvas_frame_x + frame_width > min_x and
                        canvas_frame_y < max_y and canvas_frame_y + frame_height > min_y):
                        # 执行反选
                        current_value = item_info['var'].get()
                        item_info['var'].set(not current_value)
            except tk.TclError:
                # 如果widget已被销毁，跳过
                continue
    
    def get_selected_files(self):
        """获取选中的文件列表"""
        selected = []
        for file_path, item_info in self.file_items.items():
            if item_info['var'].get():
                # 如果选择了"只处理搜索结果"，则只包含当前显示的文件
                if self.process_search_only_var.get():
                    # 检查文件是否在当前搜索结果中（即是否可见）
                    if self.is_file_visible_in_search(file_path):
                        selected.append(file_path)
                else:
                    # 否则包含所有选中的文件
                    selected.append(file_path)
        return selected
    
    def is_file_visible_in_search(self, file_path):
        """检查文件是否在当前搜索结果中可见"""
        search_text = self.search_var.get().strip()
        
        # 如果没有搜索条件，所有文件都可见
        if not search_text:
            return True
        
        # 根据显示路径选项决定搜索的文本
        if self.show_folder_path_var.get():
            search_target = os.path.normpath(file_path)
        else:
            search_target = os.path.basename(file_path)
        
        try:
            # 根据正则表达式选项进行匹配
            if self.regex_var.get():
                # 使用正则表达式搜索
                import re
                return bool(re.search(search_text, search_target, re.IGNORECASE))
            else:
                # 使用普通文本搜索（不区分大小写）
                return search_text.lower() in search_target.lower()
        except re.error:
            # 正则表达式错误时，返回False
            return False
        except Exception:
            # 其他错误时，返回False
            return False
    
    def parse_srt_file(self, file_path):
        """解析SRT文件，提取字幕文本"""
        subtitles = []
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            # 尝试其他编码
            try:
                with open(file_path, 'r', encoding='gbk') as f:
                    content = f.read()
            except UnicodeDecodeError:
                with open(file_path, 'r', encoding='latin-1') as f:
                    content = f.read()
        
        # 分割字幕块
        subtitle_blocks = re.split(r'\n\s*\n', content.strip())
        
        for block in subtitle_blocks:
            lines = block.strip().split('\n')
            if len(lines) >= 3:
                # 跳过序号和时间戳，提取字幕文本
                subtitle_text = '\n'.join(lines[2:]).strip()
                if subtitle_text:
                    subtitles.append(subtitle_text)
        
        return subtitles
    
    def sanitize_filename(self, filename):
        """清理文件名中的无效字符"""
        # Windows系统中文件名不能包含的字符
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename
    
    def convert_selected_files(self):
        """转换选中的文件"""
        selected_files = self.get_selected_files()
        
        if not selected_files:
            messagebox.showwarning("警告", "请先勾选要转换的SRT文件")
            return
        
        try:
            if self.output_mode.get() == "separate":
                self.convert_separate(selected_files)
            else:
                self.convert_merge(selected_files)
            
            messagebox.showinfo("成功", "转换完成！")
            
        except Exception as e:
            messagebox.showerror("错误", f"转换过程中发生错误：{str(e)}")
    
    def convert_separate(self, files_to_convert):
        """分别转换每个文件"""
        converted_count = 0
        failed_files = []
        
        # 检查是否需要输出到同一个文件夹
        if self.output_to_same_folder_var.get():
            if not self.output_folder:
                messagebox.showwarning("警告", "请先选择输出文件夹")
                return
        
        for srt_file in files_to_convert:
            try:
                # 解析SRT文件
                subtitles = self.parse_srt_file(srt_file)
                
                if subtitles:
                    # 生成输出文件名
                    if self.output_to_same_folder_var.get() and self.output_folder:
                        # 输出到指定文件夹，文件名格式：原文件名(绝对父目录路径).txt
                        base_name = os.path.splitext(os.path.basename(srt_file))[0]
                        parent_dir = os.path.normpath(os.path.dirname(srt_file))
                        # 清理文件名中的无效字符
                        safe_filename = self.sanitize_filename(f"{base_name}({parent_dir})")
                        new_filename = f"{safe_filename}.txt"
                        output_file = os.path.join(self.output_folder, new_filename)
                    else:
                        # 输出到原文件所在目录
                        output_file = os.path.splitext(srt_file)[0] + '.txt'
                    
                    # 检查文件覆盖
                    new_content = '，'.join(subtitles) + '，'
                    if not self.check_file_overwrite(output_file, new_content):
                        failed_files.append(f"{os.path.basename(srt_file)} (用户选择不覆盖)")
                        continue
                    
                    # 写入TXT文件
                    try:
                        with open(output_file, 'w', encoding='utf-8') as f:
                            f.write('，'.join(subtitles) + '，')
                        converted_count += 1
                    except (IOError, OSError, PermissionError) as write_error:
                        failed_files.append(f"{os.path.basename(srt_file)} (写入失败: {str(write_error)})")
                else:
                    failed_files.append(f"{os.path.basename(srt_file)} (无字幕内容)")
                
            except Exception as e:
                failed_files.append(f"{os.path.basename(srt_file)} ({str(e)})")
        
        # 显示转换结果
        result_msg = f"成功转换了 {converted_count} 个文件"
        if self.output_to_same_folder_var.get() and self.output_folder:
            result_msg += f"\n输出位置：{os.path.normpath(self.output_folder)}"
        if failed_files:
            result_msg += f"\n失败的文件：\n" + "\n".join(failed_files)
        
        messagebox.showinfo("转换完成", result_msg)
    
    def convert_merge(self, files_to_convert):
        """合并所有文件到一个TXT"""
        # 检查是否启用了递归搜索和按文件夹合并
        if self.recursive_var.get() and self.merge_by_folder_var.get():
            self.convert_merge_by_folder(files_to_convert)
        else:
            self.convert_merge_all(files_to_convert)
    
    def convert_merge_by_folder(self, files_to_convert):
        """按文件夹合并文件"""
        # 重置覆盖选择状态
        self.overwrite_all = None
        
        # 检查是否需要输出到同一个文件夹
        if self.output_to_same_folder_var.get():
            if not self.output_folder:
                messagebox.showwarning("警告", "请先选择输出文件夹")
                return
        
        # 按文件夹分组文件
        folder_groups = {}
        for file_path in files_to_convert:
            folder_path = self.file_items[file_path]['folder']
            if folder_path not in folder_groups:
                folder_groups[folder_path] = []
            folder_groups[folder_path].append(file_path)
        
        total_folders = len(folder_groups)
        successful_folders = 0
        failed_files = []
        
        for folder_path, files in folder_groups.items():
            merged_content = []
            folder_successful_count = 0
            
            for srt_file in files:
                try:
                    subtitles = self.parse_srt_file(srt_file)
                    if subtitles:
                        # 根据选项决定文件名格式
                        if self.show_merge_path_var.get():
                            # 显示绝对路径（不含扩展名）
                            filename = os.path.splitext(os.path.normpath(srt_file))[0]
                        else:
                            # 只显示文件名（不含扩展名）
                            filename = os.path.splitext(os.path.basename(srt_file))[0]
                        
                        # 将字幕内容用逗号连接
                        content = '，'.join(subtitles) + '，'
                        # 格式：文件名 + 换行 + 内容
                        file_section = f"{filename}\n{content}"
                        merged_content.append(file_section)
                        folder_successful_count += 1
                    else:
                        failed_files.append(f"{os.path.basename(srt_file)} (无字幕内容)")
                except Exception as e:
                    failed_files.append(f"{os.path.basename(srt_file)} ({str(e)})")
            
            if merged_content:
                # 决定输出文件位置
                if self.output_to_same_folder_var.get() and self.output_folder:
                    # 输出到指定文件夹，使用文件夹名作为文件名前缀
                    folder_name = os.path.normpath(os.path.dirname(srt_file))
                    # 清理文件名中的无效字符
                    safe_filename = self.sanitize_filename(f"summary({folder_name})")
                    output_file = os.path.join(self.output_folder, f"{safe_filename}.txt")
                else:
                    # 在每个文件夹下生成summary.txt
                    output_file = os.path.join(folder_path, "summary.txt")
                
                final_content = '\n\n'.join(merged_content)
                
                # 检查文件覆盖
                if not self.check_file_overwrite(output_file, final_content):
                    failed_files.append(f"文件夹 {os.path.basename(folder_path)} (用户选择不覆盖summary.txt)")
                    continue
                
                try:
                    with open(output_file, 'w', encoding='utf-8') as f:
                        f.write(final_content)
                    successful_folders += 1
                except (IOError, OSError, PermissionError) as write_error:
                    failed_files.append(f"文件夹 {os.path.basename(folder_path)} (写入summary.txt失败: {str(write_error)})")
        
        # 显示结果
        result_msg = f"成功在 {successful_folders} 个文件夹中生成了summary.txt文件"
        if self.output_to_same_folder_var.get() and self.output_folder:
            result_msg += f"\n输出位置：{os.path.normpath(self.output_folder)}"
        if failed_files:
            result_msg += f"\n处理失败的文件：\n" + "\n".join(failed_files)
        
        messagebox.showinfo("按文件夹合并完成", result_msg)
    
    def convert_merge_all(self, files_to_convert):
        """合并所有文件到一个TXT"""
        # 重置覆盖选择状态
        self.overwrite_all = None
        
        merged_content = []
        failed_files = []
        successful_count = 0
        
        for srt_file in files_to_convert:
            try:
                subtitles = self.parse_srt_file(srt_file)
                if subtitles:
                    # 根据选项决定文件名格式
                    if self.show_merge_path_var.get():
                        # 显示绝对路径（不含扩展名）
                        filename = os.path.splitext(os.path.normpath(srt_file))[0]
                    else:
                        # 只显示文件名（不含扩展名）
                        filename = os.path.splitext(os.path.basename(srt_file))[0]
                    
                    # 将字幕内容用逗号连接
                    content = '，'.join(subtitles) + '，'
                    # 格式：文件名 + 换行 + 内容
                    file_section = f"{filename}\n{content}"
                    merged_content.append(file_section)
                    successful_count += 1
                else:
                    failed_files.append(f"{os.path.basename(srt_file)} (无字幕内容)")
            except Exception as e:
                failed_files.append(f"{os.path.basename(srt_file)} ({str(e)})")
        
        if merged_content:
            # 弹窗让用户输入文件名
            output_file = filedialog.asksaveasfilename(
                title="保存合并的TXT文件",
                defaultextension=".txt",
                filetypes=[("TXT文件", "*.txt"), ("所有文件", "*.*")]
            )
            
            if not output_file:  # 用户取消了保存
                return
            
            # 用换行符连接每个文件的处理结果
            final_content = '\n\n'.join(merged_content)
            
            # 检查文件覆盖
            if not self.check_file_overwrite(output_file, final_content):
                error_msg = "用户选择不覆盖文件"
                if failed_files:
                    error_msg += f"\n处理失败的文件：\n" + "\n".join(failed_files)
                messagebox.showinfo("操作取消", error_msg)
                return
            
            try:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(final_content)
                
                # 显示结果
                result_msg = f"成功合并了 {successful_count} 个文件的内容到 {os.path.basename(output_file)}"
                if self.output_to_same_folder_var.get() and self.output_folder:
                    result_msg += f"\n输出位置：{os.path.normpath(self.output_folder)}"
                if failed_files:
                    result_msg += f"\n处理失败的文件：\n" + "\n".join(failed_files)
                
                messagebox.showinfo("合并完成", result_msg)
            except (IOError, OSError, PermissionError) as write_error:
                error_msg = f"写入合并文件失败: {str(write_error)}"
                if failed_files:
                    error_msg += f"\n处理失败的文件：\n" + "\n".join(failed_files)
                messagebox.showerror("写入失败", error_msg)
        else:
            messagebox.showwarning("警告", "没有提取到任何字幕内容")

    def check_file_overwrite(self, output_file, new_content=None):
        """检查文件是否存在，如果存在则询问用户是否覆盖
        参数：
        - output_file: 输出文件路径
        - new_content: 新文件内容（用于对比显示）
        返回值：True=允许写入, False=跳过写入
        """
        if not os.path.exists(output_file):
            return True  # 文件不存在，可以直接写入
        
        # 如果已经选择了全部覆盖或全部不覆盖
        if self.overwrite_all is True:
            return True
        elif self.overwrite_all is False:
            return False
        
        # 创建自定义对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("文件已存在")
        dialog.geometry("600x400")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        result = {'choice': None}
        
        # 获取绝对路径
        abs_path = os.path.abspath(output_file)
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(main_frame, text="文件已存在", font=("", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # 文件路径框架
        path_frame = ttk.LabelFrame(main_frame, text="文件路径", padding="10")
        path_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 可选择的文本框显示完整路径
        path_text = tk.Text(path_frame, height=2, wrap=tk.WORD, font=("Consolas", 9))
        path_text.insert(tk.END, abs_path)
        path_text.config(state=tk.DISABLED)
        path_text.pack(fill=tk.X, pady=(0, 10))
        
        # 路径操作按钮框架
        path_btn_frame = ttk.Frame(path_frame)
        path_btn_frame.pack(fill=tk.X)
        
        def copy_path():
            self.copy_file_path(abs_path, parent=dialog)
        
        def open_file():
            try:
                if platform.system() == "Windows":
                    os.startfile(abs_path)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", abs_path])
                else:  # Linux
                    subprocess.run(["xdg-open", abs_path])
            except Exception as e:
                messagebox.showerror("打开失败", f"无法打开文件：{str(e)}", parent=dialog)
        
        def compare_files():
            self.show_file_comparison(abs_path, dialog, new_content)
        
        ttk.Button(path_btn_frame, text="复制路径", command=copy_path).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(path_btn_frame, text="打开文件", command=open_file).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(path_btn_frame, text="对比文件", command=compare_files).pack(side=tk.LEFT)
        
        # 操作选择框架
        action_frame = ttk.LabelFrame(main_frame, text="请选择操作", padding="10")
        action_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(action_frame, text="该文件已存在，请选择如何处理：", font=("", 10)).pack(anchor=tk.W, pady=(0, 10))
        
        # 按钮框架
        btn_frame = ttk.Frame(action_frame)
        btn_frame.pack(fill=tk.X)
        
        def on_choice(choice):
            result['choice'] = choice
            if choice in ['overwrite_all', 'skip_all']:
                self.overwrite_all = (choice == 'overwrite_all')
            dialog.destroy()
        
        # 按钮
        ttk.Button(btn_frame, text="覆盖", command=lambda: on_choice('overwrite')).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="不覆盖", command=lambda: on_choice('skip')).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="全部覆盖", command=lambda: on_choice('overwrite_all')).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="全部不覆盖", command=lambda: on_choice('skip_all')).pack(side=tk.LEFT)
        
        # 等待用户选择
        dialog.wait_window()
        
        choice = result['choice']
        if choice in ['overwrite', 'overwrite_all']:
            return True
        elif choice in ['skip', 'skip_all']:
            return False
        else:
            return False  # 默认不覆盖

    def set_selected_file(self, file_path):
        """设置选中的文件并更新视觉反馈"""
        # 清除之前的选中状态
        if self.selected_file and self.selected_file in self.file_items:
            old_item = self.file_items[self.selected_file]
            old_item['frame'].configure(relief=tk.FLAT, borderwidth=0)
            old_item['label'].configure(bg=self.root.cget('bg'))
        
        # 设置新的选中状态
        self.selected_file = file_path
        if file_path and file_path in self.file_items:
            new_item = self.file_items[file_path]
            new_item['frame'].configure(relief=tk.SOLID, borderwidth=2)
            new_item['label'].configure(bg='lightblue')
    
    def clear_selected_file(self):
        """清除选中状态"""
        self.set_selected_file(None)
    
    def show_file_context_menu(self, event, file_path):
        """显示文件右键菜单"""
        # 设置选中状态和高亮显示
        self.set_selected_file(file_path)
        
        # 创建右键菜单
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="预览转换结果", command=lambda: self.preview_conversion_result(file_path))
        context_menu.add_command(label="转换当前文件", command=lambda: self.convert_single_file(file_path))
        context_menu.add_separator()
        context_menu.add_command(label="复制路径", command=lambda: self.copy_file_path(file_path))
        context_menu.add_command(label="在文件资源管理器中显示", command=lambda: self.open_file_location(file_path))
        context_menu.add_command(label="打开文件", command=lambda: self.open_file_with_editor(file_path))
        
        # 在鼠标位置显示菜单
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
            # 不再自动清除选中状态，保持高亮显示
    
    def preview_conversion_result(self, file_path):
        """预览转换结果"""
        try:
            # 解析SRT文件
            subtitles = self.parse_srt_file(file_path)
            if not subtitles:
                messagebox.showwarning("预览失败", f"无法解析SRT文件：{os.path.basename(file_path)}")
                return
            
            # 转换为TXT内容
            txt_content = "，".join(subtitles)
            
            # 创建预览窗口
            preview_dialog = tk.Toplevel(self.root)
            preview_dialog.title(f"转换结果预览 - {os.path.basename(file_path)}")
            preview_dialog.geometry("900x650")
            preview_dialog.transient(self.root)
            
            # 主框架
            main_frame = ttk.Frame(preview_dialog, padding="10")
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # 标题
            title_frame = ttk.Frame(main_frame)
            title_frame.pack(fill=tk.X, pady=(0, 10))
            
            ttk.Label(title_frame, text="转换结果预览", font=("", 12, "bold")).pack()
            # 修复路径显示问题，使用正确的路径分隔符
            normalized_path = os.path.normpath(file_path)
            ttk.Label(title_frame, text=f"源文件：{normalized_path}", font=("", 9)).pack(pady=(5, 0))
            
            # 文本框和滚动条
            text_frame = ttk.Frame(main_frame)
            text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
            
            text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 10))
            scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # 显示转换结果，设置为可编辑
            text_widget.insert(tk.END, txt_content)
            # 不设置为DISABLED，保持可编辑状态
            
            # 按钮框架
            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(fill=tk.X)
            
            def copy_content():
                # 获取当前文本框内容（可能已被用户编辑）
                current_content = text_widget.get("1.0", tk.END).strip()
                preview_dialog.clipboard_clear()
                preview_dialog.clipboard_append(current_content)
                messagebox.showinfo("复制成功", "转换结果已复制到剪贴板", parent=preview_dialog)
            
            def open_source_file():
                """用系统记事本打开源文件"""
                self.open_file_with_editor(file_path, parent=preview_dialog)
            
            def convert_and_save():
                """转换并输出文件"""
                try:
                    # 获取当前文本框内容（可能已被用户编辑）
                    current_content = text_widget.get("1.0", tk.END).strip()
                    
                    # 生成默认文件名
                    base_name = os.path.splitext(os.path.basename(file_path))[0]
                    default_filename = f"{base_name}.txt"
                    
                    # 让用户选择保存位置和文件名
                    output_file = filedialog.asksaveasfilename(
                        title="保存转换结果",
                        defaultextension=".txt",
                        filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
                        initialfile=default_filename,
                        parent=preview_dialog
                    )
                    
                    if not output_file:
                        return
                    
                    # 写入文件
                    with open(output_file, 'w', encoding='utf-8') as f:
                        f.write(current_content)
                    
                    messagebox.showinfo(
                        "转换成功",
                        f"文件已保存到：\n{output_file}",
                        parent=preview_dialog
                    )
                    
                except Exception as e:
                    messagebox.showerror("转换失败", f"转换并保存文件时发生错误：{str(e)}", parent=preview_dialog)
            
            # 左侧按钮
            left_btn_frame = ttk.Frame(btn_frame)
            left_btn_frame.pack(side=tk.LEFT)
            
            ttk.Button(left_btn_frame, text="复制内容", command=copy_content).pack(side=tk.LEFT, padx=(0, 5))
            ttk.Button(left_btn_frame, text="打开源文件", command=open_source_file).pack(side=tk.LEFT, padx=(0, 5))
            ttk.Button(left_btn_frame, text="转换并输出文件", command=convert_and_save).pack(side=tk.LEFT, padx=(0, 5))
            
            # 右侧按钮
            ttk.Button(btn_frame, text="关闭", command=preview_dialog.destroy).pack(side=tk.RIGHT)
            
        except Exception as e:
            messagebox.showerror("预览失败", f"预览转换结果时发生错误：{str(e)}")
    
    def convert_single_file(self, file_path):
        """转换单个文件（右键菜单调用，弹窗选择保存位置和文件名）"""
        try:
            # 解析SRT文件
            subtitles = self.parse_srt_file(file_path)
            if not subtitles:
                messagebox.showwarning("转换失败", f"无法解析SRT文件：{os.path.basename(file_path)}")
                return
            
            # 转换为TXT内容
            txt_content = "，".join(subtitles)
            
            # 生成默认文件名
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            default_filename = f"{base_name}.txt"
            
            # 让用户选择保存位置和文件名
            output_file = filedialog.asksaveasfilename(
                title="保存转换结果",
                defaultextension=".txt",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
                initialfile=default_filename,
                parent=self.root
            )
            
            if not output_file:
                return
            
            # 写入文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(txt_content)
            
            messagebox.showinfo(
                "转换成功",
                f"文件已保存到：\n{output_file}",
                parent=self.root
            )
                
        except Exception as e:
            messagebox.showerror("转换失败", f"转换文件时发生错误：{str(e)}")
    def copy_file_path(self, file_path, parent=None):
        """复制文件绝对路径到剪贴板"""
        try:
            # 获取绝对路径
            abs_path = os.path.abspath(file_path)
            
            # 复制到剪贴板
            self.root.clipboard_clear()
            self.root.clipboard_append(abs_path)
            self.root.update()  # 确保剪贴板更新
            
        except Exception as e:
            messagebox.showerror("复制失败", f"复制文件路径时发生错误：{str(e)}")
    
    def open_file_location(self, file_path):
        """在文件浏览器中打开文件位置"""
        try:
            # 获取绝对路径
            abs_path = os.path.abspath(file_path)
            
            # 检查文件是否存在
            if not os.path.exists(abs_path):
                messagebox.showwarning("文件不存在", f"文件不存在：\n{abs_path}")
                return
            
            # 根据操作系统选择合适的命令
            import platform
            system = platform.system()
            
            if system == "Windows":
                # Windows: 使用explorer /select命令
                # 将路径转换为Windows格式（使用反斜杠）
                windows_path = abs_path.replace('/', '\\')
                subprocess.run(['explorer', '/select,', windows_path])
            elif system == "Darwin":  # macOS
                # macOS: 使用open -R命令
                subprocess.run(['open', '-R', abs_path])
            elif system == "Linux":
                # Linux: 尝试多种文件管理器
                file_managers = ['nautilus', 'dolphin', 'thunar', 'pcmanfm', 'caja']
                folder_path = os.path.dirname(abs_path)
                
                success = False
                for fm in file_managers:
                    try:
                        subprocess.run([fm, folder_path])
                        success = True
                        break
                    except (subprocess.CalledProcessError, FileNotFoundError):
                        continue
                
                if not success:
                    # 如果所有文件管理器都失败，尝试使用xdg-open打开文件夹
                    try:
                        subprocess.run(['xdg-open', folder_path], check=True)
                    except (subprocess.CalledProcessError, FileNotFoundError):
                        messagebox.showwarning(
                            "无法打开文件位置",
                            f"无法找到合适的文件管理器。\n文件位置：{folder_path}"
                        )
            else:
                messagebox.showwarning("不支持的系统", f"当前系统 {system} 不支持此功能")
                
        except Exception as e:
            messagebox.showerror("打开失败", f"打开文件位置时发生错误：{str(e)}")
    
    def open_file_with_editor(self, file_path, parent=None):
        """用系统文本编辑器打开文件"""
        try:
            # 获取绝对路径
            abs_path = os.path.abspath(file_path)
            
            # 检查文件是否存在
            if not os.path.exists(abs_path):
                messagebox.showwarning("文件不存在", f"文件不存在：\n{abs_path}", parent=parent)
                return
            
            # 根据操作系统选择合适的编辑器
            import platform
            system = platform.system()
            
            if system == "Windows":
                # Windows: 使用记事本
                subprocess.run(['notepad.exe', abs_path], check=True)
            elif system == "Darwin":  # macOS
                # macOS: 使用TextEdit
                subprocess.run(['open', '-a', 'TextEdit', abs_path], check=True)
            elif system == "Linux":
                # Linux: 尝试多种文本编辑器
                editors = ['gedit', 'kate', 'mousepad', 'leafpad', 'pluma', 'nano', 'vim']
                
                success = False
                for editor in editors:
                    try:
                        subprocess.run([editor, abs_path], check=True)
                        success = True
                        break
                    except (subprocess.CalledProcessError, FileNotFoundError):
                        continue
                
                if not success:
                    # 如果所有编辑器都失败，尝试使用xdg-open
                    try:
                        subprocess.run(['xdg-open', abs_path], check=True)
                    except (subprocess.CalledProcessError, FileNotFoundError):
                        messagebox.showwarning(
                            "无法打开文件",
                            f"无法找到合适的文本编辑器。\n文件路径：{abs_path}",
                            parent=parent
                        )
            else:
                messagebox.showwarning("不支持的系统", f"当前系统 {system} 不支持此功能", parent=parent)
                
        except Exception as e:
            messagebox.showerror("打开失败", f"打开文件时发生错误：{str(e)}", parent=parent)


    def show_file_comparison(self, file_path, parent_dialog, new_content=None):
        """显示文件对比窗口"""
        # 创建对比窗口
        compare_dialog = tk.Toplevel(parent_dialog)
        compare_dialog.title(f"文件内容对比 - {os.path.basename(file_path)}")
        compare_dialog.geometry("1200x700")
        compare_dialog.transient(parent_dialog)
        
        # 主框架
        main_frame = ttk.Frame(compare_dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(title_frame, text="文件内容对比", font=("", 12, "bold")).pack()
        ttk.Label(title_frame, text=f"文件路径: {file_path}", font=("", 9)).pack(pady=(5, 0))
        
        # 对比区域框架
        compare_frame = ttk.Frame(main_frame)
        compare_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 左侧 - 现有文件内容
        left_frame = ttk.LabelFrame(compare_frame, text="现有文件内容", padding="5")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 左侧文本框和滚动条
        left_text_frame = ttk.Frame(left_frame)
        left_text_frame.pack(fill=tk.BOTH, expand=True)
        
        left_text = tk.Text(left_text_frame, wrap=tk.WORD, font=("Consolas", 9))
        left_scrollbar = ttk.Scrollbar(left_text_frame, orient=tk.VERTICAL, command=left_text.yview)
        left_text.configure(yscrollcommand=left_scrollbar.set)
        
        left_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 读取现有文件内容
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    existing_content = f.read()
                    left_text.insert(tk.END, existing_content)
            except Exception as e:
                left_text.insert(tk.END, f"无法读取文件内容：{str(e)}")
        else:
            left_text.insert(tk.END, "文件不存在")
        
        left_text.config(state=tk.DISABLED)
        
        # 右侧 - 新文件内容
        right_frame = ttk.LabelFrame(compare_frame, text="新文件内容", padding="5")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 右侧文本框和滚动条
        right_text_frame = ttk.Frame(right_frame)
        right_text_frame.pack(fill=tk.BOTH, expand=True)
        
        right_text = tk.Text(right_text_frame, wrap=tk.WORD, font=("Consolas", 9))
        right_scrollbar = ttk.Scrollbar(right_text_frame, orient=tk.VERTICAL, command=right_text.yview)
        right_text.configure(yscrollcommand=right_scrollbar.set)
        
        right_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        right_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 显示新文件内容
        if new_content:
            right_text.insert(tk.END, new_content)
        else:
            right_text.insert(tk.END, "无新内容")
        
        right_text.config(state=tk.DISABLED)
        
        # 按钮框架
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)
        
        def copy_existing():
            if os.path.exists(file_path):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    compare_dialog.clipboard_clear()
                    compare_dialog.clipboard_append(content)
                    messagebox.showinfo("复制成功", "现有文件内容已复制到剪贴板", parent=compare_dialog)
                except Exception as e:
                    messagebox.showerror("复制失败", f"无法复制现有文件内容：{str(e)}", parent=compare_dialog)
        
        def copy_new():
            if new_content:
                compare_dialog.clipboard_clear()
                compare_dialog.clipboard_append(new_content)
                messagebox.showinfo("复制成功", "新文件内容已复制到剪贴板", parent=compare_dialog)
        
        def open_in_editor():
            if os.path.exists(file_path):
                try:
                    if platform.system() == "Windows":
                        subprocess.run(["notepad", file_path])
                    elif platform.system() == "Darwin":  # macOS
                        subprocess.run(["open", "-t", file_path])
                    else:  # Linux
                        subprocess.run(["xdg-open", file_path])
                except Exception as e:
                    messagebox.showerror("打开失败", f"无法在编辑器中打开文件：{str(e)}", parent=compare_dialog)
        
        # 同步滚动
        def sync_scroll(*args):
            left_text.yview_moveto(args[0])
            right_text.yview_moveto(args[0])
        
        left_text.configure(yscrollcommand=lambda *args: (left_scrollbar.set(*args), sync_scroll(*args)))
        right_text.configure(yscrollcommand=lambda *args: (right_scrollbar.set(*args), sync_scroll(*args)))
        
        ttk.Button(btn_frame, text="复制现有内容", command=copy_existing).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="复制新内容", command=copy_new).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="在编辑器中打开源文件", command=open_in_editor).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="关闭", command=compare_dialog.destroy).pack(side=tk.RIGHT)

    # 文件拖拽导入事件处理方法
    def on_file_drag_enter(self, event):
        """文件拖拽进入事件"""
        if HAS_DND:
            self.is_drag_over = True
            # 添加视觉反馈：改变Canvas背景色
            self.canvas.configure(bg="lightblue")
    
    def on_file_drag_motion(self, event):
        """文件拖拽移动事件"""
        if HAS_DND:
            # 可以在这里添加更多的视觉反馈
            pass
    
    def on_file_drag_leave(self, event):
        """文件拖拽离开事件"""
        if HAS_DND:
            self.is_drag_over = False
            # 恢复Canvas原始背景色
            self.canvas.configure(bg="white")
    
    def on_paste_files(self, event):
        """处理Ctrl+V粘贴文件路径事件"""
        try:
            # 获取剪贴板内容
            clipboard_content = self.root.clipboard_get()
            
            if not clipboard_content.strip():
                messagebox.showwarning("粘贴导入", "剪贴板为空")
                return
            
            # 解析剪贴板内容中的文件路径
            srt_files = []
            lines = clipboard_content.strip().split('\n')
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # 处理可能的文件路径格式
                file_paths = []
                
                # 处理用引号包围的路径
                if line.startswith('"') and line.endswith('"'):
                    file_paths.append(line[1:-1])
                elif line.startswith("'") and line.endswith("'"):
                    file_paths.append(line[1:-1])
                # 处理file://协议的URL
                elif line.startswith('file://'):
                    try:
                        # 解码URL
                        decoded_path = urllib.parse.unquote(line[7:])
                        # Windows路径处理
                        if platform.system() == "Windows" and decoded_path.startswith('/'):
                            decoded_path = decoded_path[1:]
                        file_paths.append(decoded_path)
                    except Exception:
                        file_paths.append(line)
                # 处理多个路径用空格分隔的情况
                elif ' ' in line and not os.path.exists(line):
                    # 尝试按空格分割，但要考虑路径中可能包含空格的情况
                    parts = line.split()
                    current_path = ""
                    for part in parts:
                        if current_path:
                            current_path += " " + part
                        else:
                            current_path = part
                        
                        # 检查当前路径是否存在
                        if os.path.exists(current_path):
                            file_paths.append(current_path)
                            current_path = ""
                    
                    # 如果最后还有未处理的路径
                    if current_path and os.path.exists(current_path):
                        file_paths.append(current_path)
                else:
                    file_paths.append(line)
                
                # 检查每个路径并添加SRT文件
                for file_path in file_paths:
                    if os.path.isfile(file_path) and file_path.lower().endswith('.srt'):
                        srt_files.append(file_path)
                    elif os.path.isdir(file_path):
                        # 如果是文件夹，递归查找.srt文件
                        for root, dirs, files in os.walk(file_path):
                            for file in files:
                                if file.lower().endswith('.srt'):
                                    srt_files.append(os.path.join(root, file))
            
            # 添加文件到列表
            if srt_files:
                added_count = 0
                for srt_file in srt_files:
                    if srt_file not in self.file_items:
                        self.add_file_item(srt_file)
                        added_count += 1
                
                # 显示导入结果
                if added_count > 0:
                    messagebox.showinfo("粘贴导入成功", f"成功导入 {added_count} 个SRT文件")
                    # 重新应用排序和过滤
                    self.sort_file_list()
                    self.filter_file_list()
                else:
                    messagebox.showinfo("粘贴导入", "所有文件都已存在于列表中")
            else:
                messagebox.showwarning("粘贴导入失败", "剪贴板中未找到有效的SRT文件路径")
                
        except tk.TclError:
            messagebox.showwarning("粘贴导入", "无法访问剪贴板")
        except Exception as e:
            messagebox.showerror("粘贴导入错误", f"处理粘贴内容时发生错误：{str(e)}")
    
    def on_file_drop(self, event):
        """文件拖拽放下事件"""
        if not HAS_DND:
            return
        
        try:
            # 恢复Canvas原始背景色
            self.canvas.configure(bg="white")
            self.is_drag_over = False
            
            # 获取拖拽的文件路径
            files = event.data.split()
            srt_files = []
            
            # 过滤出.srt文件
            for file_path in files:
                # 移除可能的引号
                file_path = file_path.strip('{}').strip('"').strip("'")
                if os.path.isfile(file_path) and file_path.lower().endswith('.srt'):
                    srt_files.append(file_path)
                elif os.path.isdir(file_path):
                    # 如果是文件夹，递归查找.srt文件
                    for root, dirs, files in os.walk(file_path):
                        for file in files:
                            if file.lower().endswith('.srt'):
                                srt_files.append(os.path.join(root, file))
            
            # 添加文件到列表
            if srt_files:
                added_count = 0
                for srt_file in srt_files:
                    if srt_file not in self.file_items:
                        self.add_file_item(srt_file)
                        added_count += 1
                
                # 显示导入结果
                if added_count > 0:
                    messagebox.showinfo("拖拽导入成功", f"成功导入 {added_count} 个SRT文件")
                    # 重新应用排序和过滤
                    self.sort_file_list()
                    self.filter_file_list()
                else:
                    messagebox.showinfo("拖拽导入", "所有文件都已存在于列表中")
            else:
                messagebox.showwarning("拖拽导入失败", "未找到有效的SRT文件")
                
        except Exception as e:
            messagebox.showerror("拖拽导入错误", f"处理拖拽文件时发生错误：{str(e)}")
    
    def show_canvas_context_menu(self, event):
        """显示Canvas空白区域右键菜单"""
        # 创建右键菜单
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="粘贴", command=lambda: self.on_paste_files(event))
        context_menu.add_separator()
        context_menu.add_command(label="全选", command=self.select_all_files)
        context_menu.add_command(label="取消全选", command=self.deselect_all_files)
        context_menu.add_command(label="反向选择", command=self.invert_selection)
        context_menu.add_separator()
        context_menu.add_command(label="清空文件列表", command=self.clear_all_files)
        context_menu.add_command(label="删除选中文件", command=self.remove_selected_files)
        
        # 在鼠标位置显示菜单
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    def on_function_changed(self, event=None):
        """功能选择下拉框变化时的回调"""
        selected_function = self.function_mode.get()
        print(f"选择的功能: {selected_function}")
        
        # 根据选择的功能更新界面状态
        if selected_function == "srt转txt":
            # 当前功能，保持现有状态
            pass
        else:
            # 其他功能暂未实现，显示提示
            messagebox.showinfo("功能提示", f"'{selected_function}' 功能正在开发中，敬请期待！")
            # 重置为默认功能
            self.function_mode.set("srt转txt")
    
    def create_function_help_tooltip(self, widget):
        """为帮助按钮创建动态悬浮提示"""
        def on_enter(event):
            # 获取当前选择的功能
            current_function = self.function_mode.get()
            description = self.function_descriptions.get(current_function, "暂无描述")
            
            # 创建提示窗口
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root-30}")
            
            # 创建提示标签
            label = tk.Label(
                tooltip,
                text=f"{current_function}:\n{description}",
                background="lightyellow",
                relief="solid",
                borderwidth=1,
                wraplength=300,
                justify=tk.LEFT,
                font=("", 9),
                padx=8,
                pady=6
            )
            label.pack()
            
            # 存储提示窗口引用
            widget.tooltip = tooltip
        
        def on_leave(event):
            # 销毁提示窗口
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
        
        # 绑定鼠标事件
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)


def main():
    # 根据是否支持拖拽功能选择不同的根窗口类型
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    
    app = SRTToTXTConverter(root)
    
    # 如果不支持拖拽功能，显示提示
    if not HAS_DND:
        print("提示：未安装tkinterdnd2库，文件拖拽功能不可用")
        print("可以通过 'pip install tkinterdnd2' 安装以启用拖拽功能")
    
    root.mainloop()

if __name__ == "__main__":
    main()