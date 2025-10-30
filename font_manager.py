"""
Windows字体管理工具
功能：安装、卸载、预览、搜索系统字体
"""

import os
import sys
import shutil
import winreg
import ctypes
from ctypes import wintypes
from tkinter import *
from tkinter import ttk, filedialog, messagebox, font
from pathlib import Path
import threading
from PIL import Image, ImageTk

class FontManager:
    """字体管理核心类"""
    
    # Windows字体注册表路径
    SYSTEM_FONTS_REG_PATH = r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
    USER_FONTS_REG_PATH = r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
    
    def __init__(self):
        # 系统字体目录
        self.system_fonts_dir = os.path.join(os.environ['WINDIR'], 'Fonts')
        # 用户字体目录
        self.user_fonts_dir = os.path.join(
            os.environ['LOCALAPPDATA'], 
            'Microsoft', 'Windows', 'Fonts'
        )
        # 确保用户字体目录存在
        os.makedirs(self.user_fonts_dir, exist_ok=True)
        
    def get_installed_fonts(self):
        """获取已安装的字体列表（包括系统和用户字体）"""
        fonts = []
        
        # 用于去重的集合（避免同一字体在两个注册表中都有记录）
        added_fonts = set()
        
        # 读取系统字体 (HKEY_LOCAL_MACHINE)
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, self.SYSTEM_FONTS_REG_PATH, 0, 
                              winreg.KEY_READ) as key:
                i = 0
                while True:
                    try:
                        name, value, _ = winreg.EnumValue(key, i)
                        font_path = os.path.join(self.system_fonts_dir, value) if not os.path.isabs(value) else value
                        
                        # 只添加文件真实存在的字体
                        if os.path.exists(font_path):
                            # 根据实际路径判断字体类型
                            is_system = font_path.lower().startswith(self.system_fonts_dir.lower())
                            font_key = (name, font_path.lower())
                            
                            if font_key not in added_fonts:
                                fonts.append({
                                    'name': name,
                                    'file': value,
                                    'path': font_path,
                                    'type': '系统字体' if is_system else '用户字体'
                                })
                                added_fonts.add(font_key)
                        
                        i += 1
                    except WindowsError:
                        break
        except Exception as e:
            print(f"读取系统字体注册表错误: {e}")
        
        # 读取用户字体 (HKEY_CURRENT_USER)
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.USER_FONTS_REG_PATH, 0, 
                              winreg.KEY_READ) as key:
                i = 0
                while True:
                    try:
                        name, value, _ = winreg.EnumValue(key, i)
                        font_path = os.path.join(self.user_fonts_dir, value) if not os.path.isabs(value) else value
                        
                        # 只添加文件真实存在的字体
                        if os.path.exists(font_path):
                            # 根据实际路径判断字体类型（避免将C:\Windows\Fonts\下的字体误判为用户字体）
                            is_system = font_path.lower().startswith(self.system_fonts_dir.lower())
                            font_key = (name, font_path.lower())
                            
                            # 避免重复添加（系统注册表中已有的字体）
                            if font_key not in added_fonts:
                                fonts.append({
                                    'name': name,
                                    'file': value,
                                    'path': font_path,
                                    'type': '系统字体' if is_system else '用户字体'
                                })
                                added_fonts.add(font_key)
                        
                        i += 1
                    except WindowsError:
                        break
        except Exception as e:
            print(f"读取用户字体注册表错误: {e}")
        
        return sorted(fonts, key=lambda x: x['name'])
    
    def install_font(self, font_path, install_type='user'):
        """安装字体文件
        Args:
            font_path: 字体文件路径
            install_type: 'user' 用户字体（无需管理员权限）或 'system' 系统字体（需要管理员权限）
        """
        if not os.path.exists(font_path):
            return False, "字体文件不存在"
        
        # 检查文件扩展名
        ext = os.path.splitext(font_path)[1].lower()
        if ext not in ['.ttf', '.otf', '.ttc']:
            return False, "不支持的字体格式，仅支持 .ttf, .otf, .ttc"
        
        font_filename = os.path.basename(font_path)
        
        # 根据安装类型选择目标目录和注册表
        if install_type == 'system':
            target_dir = self.system_fonts_dir
            reg_hkey = winreg.HKEY_LOCAL_MACHINE
            reg_path = self.SYSTEM_FONTS_REG_PATH
            type_name = "系统字体"
        else:  # user
            target_dir = self.user_fonts_dir
            reg_hkey = winreg.HKEY_CURRENT_USER
            reg_path = self.USER_FONTS_REG_PATH
            type_name = "用户字体"
        
        target_path = os.path.join(target_dir, font_filename)
        
        # 检查字体是否已存在
        if os.path.exists(target_path):
            return False, f"字体已存在于{type_name}中"
        
        try:
            # 复制字体文件到目标目录
            shutil.copy2(font_path, target_path)
            
            # 获取字体名称（简化处理，使用文件名）
            font_name = os.path.splitext(font_filename)[0]
            if ext == '.ttf':
                reg_name = f"{font_name} (TrueType)"
            elif ext == '.otf':
                reg_name = f"{font_name} (OpenType)"
            else:
                reg_name = f"{font_name}"
            
            # 写入注册表
            with winreg.OpenKey(reg_hkey, reg_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, reg_name, 0, winreg.REG_SZ, font_filename)
            
            # 通知系统字体已更改
            self._notify_font_change()
            
            return True, f"字体已成功安装为{type_name}"
            
        except PermissionError:
            if install_type == 'system':
                return False, "安装系统字体需要管理员权限，请以管理员身份运行程序"
            else:
                return False, "权限不足"
        except Exception as e:
            return False, f"安装失败: {str(e)}"
    
    def uninstall_font(self, font_name, font_file, font_type):
        """卸载字体
        Args:
            font_name: 字体名称
            font_file: 字体文件名
            font_type: '系统字体' 或 '用户字体'
        """
        try:
            # 根据字体类型选择注册表和目录
            if font_type == '系统字体':
                reg_hkey = winreg.HKEY_LOCAL_MACHINE
                reg_path = self.SYSTEM_FONTS_REG_PATH
                font_dir = self.system_fonts_dir
            else:  # 用户字体
                reg_hkey = winreg.HKEY_CURRENT_USER
                reg_path = self.USER_FONTS_REG_PATH
                font_dir = self.user_fonts_dir
            
            # 从注册表删除
            with winreg.OpenKey(reg_hkey, reg_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, font_name)
            
            # 删除字体文件
            font_path = os.path.join(font_dir, font_file)
            if os.path.exists(font_path):
                os.remove(font_path)
            
            # 通知系统字体已更改
            self._notify_font_change()
            
            return True, f"{font_type}卸载成功"
            
        except PermissionError:
            if font_type == '系统字体':
                return False, "卸载系统字体需要管理员权限，请以管理员身份运行程序"
            else:
                return False, "权限不足"
        except FileNotFoundError:
            return False, "注册表项不存在"
        except Exception as e:
            return False, f"卸载失败: {str(e)}"
    
    def _notify_font_change(self):
        """通知Windows系统字体已更改"""
        HWND_BROADCAST = 0xFFFF
        WM_FONTCHANGE = 0x001D
        SMTO_ABORTIFHUNG = 0x0002
        
        try:
            result = ctypes.c_long()
            ctypes.windll.user32.SendMessageTimeoutW(
                HWND_BROADCAST, WM_FONTCHANGE, 0, 0,
                SMTO_ABORTIFHUNG, 5000, ctypes.byref(result)
            )
        except Exception as e:
            print(f"通知系统失败: {e}")


class FontManagerGUI:
    """字体管理器GUI界面"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("字体管理器")
        self.root.geometry("1400x800")
        
        # 获取资源文件路径（支持打包后的exe）
        if getattr(sys, 'frozen', False):
            # 打包后的exe路径
            self.base_path = sys._MEIPASS
        else:
            # 开发环境路径
            self.base_path = os.path.dirname(os.path.abspath(__file__))
        
        # 设置窗口图标 - 使用多种方式确保兼容性
        icon_set = False
        
        # 方式1: 尝试从打包路径加载.ico文件
        try:
            icon_path = os.path.join(self.base_path, 'icon.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                icon_set = True
        except Exception as e:
            pass
        
        # 方式2: 尝试从当前目录加载.ico文件
        if not icon_set:
            try:
                if os.path.exists('icon.ico'):
                    self.root.iconbitmap('icon.ico')
                    icon_set = True
            except Exception as e:
                pass
        
        # 方式3: 使用LOGO.jpg通过PhotoImage设置（备用方案）
        try:
            logo_path = os.path.join(self.base_path, 'LOGO.jpg')
            if not os.path.exists(logo_path) and os.path.exists('LOGO.jpg'):
                logo_path = 'LOGO.jpg'
            
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((64, 64), Image.LANCZOS)
                self.icon_photo = ImageTk.PhotoImage(logo_img)
                self.root.iconphoto(True, self.icon_photo)
        except Exception as e:
            pass
        
        # 现代化配色方案
        self.colors = {
            'bg': '#F5F7FA',           # 浅灰背景
            'sidebar': '#FFFFFF',       # 侧边栏白色
            'primary': '#1d7cff',      # 主色调蓝色
            'success': '#52C41A',      # 成功绿色
            'danger': '#FF4D4F',       # 危险红色
            'warning': '#FAAD14',      # 警告橙色
            'text': '#262626',         # 主文字颜色
            'text_secondary': '#8C8C8C', # 次要文字
            'border': '#E8E8E8',       # 边框颜色
            'hover': '#E6F7FF',        # 悬停色
            'user_tag': '#E6F7FF',     # 用户字体标签
            'system_tag': '#FFF1F0'    # 系统字体标签
        }
        
        # 配置root背景色
        self.root.configure(bg=self.colors['bg'])
        
        self.font_manager = FontManager()
        self.all_fonts = []  # 所有字体
        self.current_fonts = []  # 根据tab过滤后的字体
        self.filtered_fonts = []  # 再根据搜索过滤后的字体
        self.current_tab = "user"  # 当前选中的tab，默认显示用户字体
        self.current_preview_font = None  # 当前预览的字体
        
        self._check_admin_rights()
        self._create_widgets()
        self._load_fonts()
    
    def _check_admin_rights(self):
        """检查是否有管理员权限"""
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            if not is_admin:
                messagebox.showwarning(
                    "权限提示",
                    "程序未以管理员身份运行，安装和卸载字体功能可能受限。\n"
                    "建议右键点击程序，选择'以管理员身份运行'。"
                )
        except:
            pass
    
    def _create_widgets(self):
        """创建界面组件"""
        # ===== 顶部标题栏 =====
        header = Frame(self.root, bg=self.colors['sidebar'], height=70)
        header.pack(side=TOP, fill=X)
        header.pack_propagate(False)
        
        # LOGO + 标题
        title_frame = Frame(header, bg=self.colors['sidebar'])
        title_frame.pack(side=LEFT, padx=30, pady=15)
        
        # 加载并显示LOGO
        try:
            logo_path = os.path.join(self.base_path, 'LOGO.jpg')
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((40, 40), Image.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                logo_label = Label(title_frame, image=self.logo_photo, 
                                 bg=self.colors['sidebar'])
                logo_label.pack(side=LEFT, padx=(0, 12))
        except Exception as e:
            pass  # 如果LOGO文件不存在，则跳过
        
        # 标题
        title_label = Label(title_frame, text="字体管理器", 
                           font=("微软雅黑", 20, "bold"),
                           bg=self.colors['sidebar'],
                           fg=self.colors['text'])
        title_label.pack(side=LEFT)
        
        # 搜索框 - 搜索图标在输入框内部
        search_container = Frame(header, bg='#FAFAFA', 
                                highlightthickness=1,
                                highlightbackground=self.colors['border'],
                                highlightcolor=self.colors['primary'])
        search_container.pack(side=RIGHT, padx=30, pady=20)
        
        # 搜索图标（在输入框内部）- 线性图标
        search_icon = Label(search_container, text="⌕", bg='#FAFAFA',
                           font=("微软雅黑", 16), fg=self.colors['text_secondary'])
        search_icon.pack(side=LEFT, padx=(12, 5))
        
        # 搜索输入框
        self.search_var = StringVar()
        self.search_var.trace('w', lambda *args: self._filter_fonts())
        search_entry = Entry(search_container, textvariable=self.search_var,
                            width=25, font=("微软雅黑", 11),
                            relief=FLAT, bg='#FAFAFA',
                            highlightthickness=0,
                            borderwidth=0)
        search_entry.pack(side=LEFT, ipady=6, padx=(0, 12))
        
        # 分隔线
        separator = Frame(self.root, height=1, bg=self.colors['border'])
        separator.pack(fill=X)
        
        # 中间主区域 - 使用PanedWindow实现左右分栏
        main_paned = ttk.PanedWindow(self.root, orient=HORIZONTAL)
        main_paned.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        # ===== 左侧：字体列表 =====
        left_frame = Frame(main_paned, bg=self.colors['sidebar'])
        main_paned.add(left_frame, weight=3)
        
        # 顶部操作栏
        action_bar = Frame(left_frame, bg=self.colors['sidebar'], height=60)
        action_bar.pack(side=TOP, fill=X)
        action_bar.pack_propagate(False)
        
        # Tab切换栏 - Arco Design风格，与列表左对齐
        tab_container = Frame(action_bar, bg=self.colors['sidebar'])
        tab_container.pack(side=LEFT, padx=(10, 0), pady=10)
        
        # Tab按钮容器
        tab_frame = Frame(tab_container, bg=self.colors['sidebar'])
        tab_frame.pack(side=TOP)
        
        # 下划线指示器容器
        indicator_frame = Frame(tab_container, bg=self.colors['sidebar'], height=2)
        indicator_frame.pack(side=TOP, fill=X, pady=(5, 0))
        
        # 创建Tab按钮
        self.tab_buttons = {}
        self.tab_indicators = {}
        tab_configs = [
            ("user", "用户字体", self.colors['primary']),
            ("system", "系统字体", self.colors['primary'])
        ]
        
        for idx, (tab_id, tab_text, color) in enumerate(tab_configs):
            # Tab按钮
            btn = Button(tab_frame, text=tab_text, 
                        command=lambda t=tab_id: self._switch_tab(t),
                        relief=FLAT,
                        bg=self.colors['sidebar'],
                        fg=self.colors['text_secondary'],
                        font=("微软雅黑", 11),
                        cursor="hand2",
                        activebackground=self.colors['sidebar'],
                        bd=0,
                        padx=16, pady=8)
            # 第一个Tab左边距为0，与列表对齐
            btn.pack(side=LEFT, padx=(0 if idx == 0 else 8, 8))
            
            # 绑定鼠标悬停事件
            btn.bind("<Enter>", lambda e, tid=tab_id: self._on_tab_hover(tid, True))
            btn.bind("<Leave>", lambda e, tid=tab_id: self._on_tab_hover(tid, False))
            
            # 为每个tab创建下划线指示器（初始时使用背景色隐藏）
            indicator = Frame(indicator_frame, bg=self.colors['sidebar'], height=2)
            # 第一个指示器左边距为0，与按钮对齐
            indicator.pack(side=LEFT, padx=(0 if idx == 0 else 8, 8))
            # 设置指示器宽度与按钮相同
            indicator.config(width=btn.winfo_reqwidth())
            
            self.tab_buttons[tab_id] = {"button": btn, "color": color}
            self.tab_indicators[tab_id] = indicator
        
        # 操作按钮
        btn_frame = Frame(action_bar, bg=self.colors['sidebar'])
        btn_frame.pack(side=RIGHT, padx=20, pady=15)
        
        # 全选按钮
        self.select_all_btn = Button(btn_frame, text="☑ 全选", command=self._select_all_fonts,
                             relief=FLAT, bg=self.colors['primary'], fg='white',
                             font=("微软雅黑", 10), cursor="hand2",
                             padx=15, pady=6)
        self.select_all_btn.pack(side=LEFT, padx=3)
        
        # 创建现代化按钮 - 线性图标风格（移除安装按钮）
        self.uninstall_btn = Button(btn_frame, text="− 批量卸载", command=self._uninstall_font,
                             relief=FLAT, bg=self.colors['danger'], fg='white',
                             font=("微软雅黑", 10), cursor="hand2",
                             padx=15, pady=6)
        self.uninstall_btn.pack(side=LEFT, padx=3)
        
        refresh_btn = Button(btn_frame, text="↻ 刷新", command=self._load_fonts,
                           relief=FLAT, bg=self.colors['primary'], fg='white',
                           font=("微软雅黑", 10), cursor="hand2",
                           padx=15, pady=6)
        refresh_btn.pack(side=LEFT, padx=3)
        
        # 设置默认选中
        self._update_tab_style()
        
        # 字体列表（带滚动条）
        list_frame = Frame(left_frame, bg='white', relief=FLAT, borderwidth=0)
        list_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))
        
        # 配置ttk样式 - DeepSeek Chat风格
        style = ttk.Style()
        style.theme_use('clam')
        
        # 表格样式 - 无边框
        style.configure("Treeview",
                       background="white",
                       foreground="#1D2129",
                       rowheight=38,
                       fieldbackground="white",
                       borderwidth=0,
                       highlightthickness=0,
                       relief=FLAT,
                       font=("微软雅黑", 10))
        style.configure("Treeview.Heading",
                       background="#F7F8FA",
                       foreground="#4E5969",
                       borderwidth=0,
                       relief=FLAT,
                       font=("微软雅黑", 10, "bold"))
        style.map('Treeview', 
                 background=[('selected', '#E8F3FF')],
                 foreground=[('selected', '#1D2129')])
        style.map('Treeview.Heading',
                 background=[('active', '#F7F8FA')])
        
        # 超轻量化滚动条样式 - 类似现代浏览器
        style.configure("Vertical.TScrollbar",
                       background="#C9CDD4",
                       troughcolor="white",
                       borderwidth=0,
                       arrowsize=0,  # 隐藏箭头
                       width=20)  # 20px宽度
        style.configure("Horizontal.TScrollbar",
                       background="#C9CDD4",
                       troughcolor="white",
                       borderwidth=0,
                       arrowsize=0,
                       width=20)
        style.map("Vertical.TScrollbar",
                 background=[('active', '#C9CDD4'), ('!active', '#C9CDD4')])
        style.map("Horizontal.TScrollbar",
                 background=[('active', '#C9CDD4'), ('!active', '#C9CDD4')])
        
        # 创建Treeview - 去掉类型列，支持多选
        columns = ('name', 'file', 'path')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', 
                                selectmode='extended', style="Treeview")
        
        self.tree.heading('name', text='字体名称')
        self.tree.heading('file', text='文件名')
        self.tree.heading('path', text='路径')
        
        # 不使用背景色标签，保持简洁
        # self.tree.tag_configure('system', ...)
        # self.tree.tag_configure('user', ...)
        
        # 轻量化滚动条
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview,
                           style="Vertical.TScrollbar")
        hsb = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tree.xview,
                           style="Horizontal.TScrollbar")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # 绑定选择事件 - 点击字体时预览
        self.tree.bind('<<TreeviewSelect>>', lambda e: self._on_font_select())
        
        # ===== 右侧：预览面板 =====
        right_frame = Frame(main_paned, bg=self.colors['sidebar'])
        main_paned.add(right_frame, weight=2)
        
        # 预览标题
        preview_title_frame = Frame(right_frame, bg=self.colors['sidebar'], height=50)
        preview_title_frame.pack(fill=X)
        preview_title_frame.pack_propagate(False)
        
        self.preview_title_label = Label(preview_title_frame, text="字体预览", 
                                         font=("微软雅黑", 14, "bold"),
                                         bg=self.colors['sidebar'],
                                         fg=self.colors['text'])
        self.preview_title_label.pack(side=LEFT, padx=20, pady=10)
        
        # 预览内容区域（带轻量化滚动条）
        preview_content_frame = Frame(right_frame, bg='white')
        preview_content_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))
        
        preview_scrollbar = ttk.Scrollbar(preview_content_frame, orient="vertical",
                                         style="Vertical.TScrollbar")
        preview_scrollbar.pack(side=RIGHT, fill=Y)
        
        self.preview_text = Text(preview_content_frame, wrap=WORD, 
                                yscrollcommand=preview_scrollbar.set,
                                font=("微软雅黑", 11), bg="white",
                                relief=FLAT,
                                borderwidth=0,
                                highlightthickness=0,
                                state=DISABLED, padx=20, pady=20)
        self.preview_text.pack(side=LEFT, fill=BOTH, expand=True)
        preview_scrollbar.config(command=self.preview_text.yview)
        
        # 显示默认提示
        self._show_preview_hint()
        
        # 底部状态栏
        status_frame = Frame(self.root, bg=self.colors['sidebar'], height=35)
        status_frame.pack(side=BOTTOM, fill=X)
        status_frame.pack_propagate(False)
        
        self.status_var = StringVar()
        self.status_var.set("就绪")
        status_label = Label(status_frame, textvariable=self.status_var,
                            bg=self.colors['sidebar'],
                            fg=self.colors['text_secondary'],
                            font=("微软雅黑", 9),
                            anchor=W)
        status_label.pack(side=LEFT, padx=20, fill=X, expand=True)
    
    def _switch_tab(self, tab_id):
        """切换Tab"""
        self.current_tab = tab_id
        self._update_tab_style()
        self._apply_filters()
    
    def _on_tab_hover(self, tab_id, is_enter):
        """Tab鼠标悬停效果 - Arco Design风格"""
        if tab_id != self.current_tab:  # 只对未选中的tab应用悬停效果
            btn = self.tab_buttons[tab_id]["button"]
            if is_enter:
                # 鼠标进入：文字颜色变深（接近黑色）
                btn.config(fg='#1D2129')
            else:
                # 鼠标离开：文字颜色恢复灰色
                btn.config(fg='#86909C')
    
    def _update_tab_style(self):
        """更新Tab按钮样式 - Arco Design风格（底部下划线）"""
        for tab_id, tab_info in self.tab_buttons.items():
            btn = tab_info["button"]
            indicator = self.tab_indicators[tab_id]
            color = tab_info["color"]
            
            if tab_id == self.current_tab:
                # 选中状态：黑色文字 + 加粗，显示蓝色下划线
                btn.config(fg='#1D2129',  # 深黑色
                          font=("微软雅黑", 12, "bold"),
                          activeforeground='#1D2129')
                # 显示明显的蓝色下划线指示器
                indicator.config(bg='#165DFF', height=3)  # Arco蓝色
            else:
                # 未选中状态：灰色文字，隐藏下划线
                btn.config(fg='#86909C',  # 中灰色
                          font=("微软雅黑", 12),
                          activeforeground='#4E5969')  # 鼠标按下时稍深
                # 隐藏下划线指示器（使用背景色）
                indicator.config(bg=self.colors['sidebar'], height=2)
        
        # 根据当前Tab更新按钮状态
        if self.current_tab == 'system':
            # 系统字体Tab：禁用全选和卸载按钮，灰色显示
            self.select_all_btn.config(
                state=DISABLED,
                bg='#C9CDD4',  # 灰色背景
                fg='#86909C',  # 灰色文字
                cursor='arrow'
            )
            self.uninstall_btn.config(
                state=DISABLED,
                bg='#C9CDD4',  # 灰色背景
                fg='#86909C',  # 灰色文字
                cursor='arrow'
            )
        else:
            # 用户字体Tab：启用全选和卸载按钮
            self.select_all_btn.config(
                state=NORMAL,
                bg=self.colors['primary'],
                fg='white',
                cursor='hand2'
            )
            self.uninstall_btn.config(
                state=NORMAL,
                bg=self.colors['danger'],
                fg='white',
                cursor='hand2'
            )
    
    def _load_fonts(self):
        """加载字体列表"""
        self.status_var.set("正在加载字体列表...")
        self.root.update()
        
        # 在后台线程加载
        def load():
            self.all_fonts = self.font_manager.get_installed_fonts()
            self.root.after(0, self._apply_filters)
        
        threading.Thread(target=load, daemon=True).start()
    
    def _update_tree(self):
        """更新树形视图"""
        # 清空现有项
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 统计计数
        system_count = 0
        user_count = 0
        
        # 添加字体 - 不显示类型列
        for font_info in self.filtered_fonts:
            font_type = font_info.get('type', '系统字体')
            
            if font_type == '系统字体':
                system_count += 1
            else:
                user_count += 1
            
            # 只插入name、file、path三列数据
            self.tree.insert('', 'end', values=(
                font_info['name'],
                font_info['file'],
                font_info['path']
            ))
        
        count = len(self.filtered_fonts)
        total = len(self.current_fonts)
        
        # 根据当前Tab显示对应的字体类型和数量
        if self.current_tab == 'system':
            font_type_text = "系统字体"
        else:
            font_type_text = "用户字体"
        
        if count == total:
            self.status_var.set(f"共 {total} 个{font_type_text}")
        else:
            self.status_var.set(f"显示 {count} 个{font_type_text}（共 {total} 个）")
    
    def _apply_filters(self):
        """应用tab和搜索过滤"""
        # 第一步：根据tab过滤
        if self.current_tab == "user":
            self.current_fonts = [f for f in self.all_fonts if f['type'] == '用户字体']
        else:  # system
            self.current_fonts = [f for f in self.all_fonts if f['type'] == '系统字体']
        
        # 第二步：根据搜索框过滤
        search_text = self.search_var.get().lower()
        if search_text:
            self.filtered_fonts = [
                f for f in self.current_fonts
                if search_text in f['name'].lower() or search_text in f['file'].lower()
            ]
        else:
            self.filtered_fonts = self.current_fonts
        
        self._update_tree()
    
    def _filter_fonts(self):
        """根据搜索框过滤字体"""
        self._apply_filters()
    
    def _show_preview_hint(self):
        """显示默认预览提示"""
        self.preview_text.config(state=NORMAL)
        self.preview_text.delete(1.0, END)
        
        hint_text = """



        👈 请在左侧列表中选择一个字体
        
        字体预览将在此处显示
        
        
        
        提示：
        • 点击字体即可预览
        • 系统字体不可卸载
        • 用户字体可自由管理
        
        """
        
        self.preview_text.insert(END, hint_text)
        self.preview_text.tag_add("center", "1.0", "end")
        self.preview_text.tag_config("center", justify='center', foreground='#888')
        self.preview_text.config(state=DISABLED)
    
    def _on_font_select(self):
        """当选择字体时触发"""
        selection = self.tree.selection()
        if not selection:
            self._show_preview_hint()
            return
        
        item = self.tree.item(selection[0])
        # 现在values顺序是: (name, file, path)
        font_name = item['values'][0]
        font_file = item['values'][1]
        
        # 根据当前Tab判断字体类型
        font_type = '系统字体' if self.current_tab == 'system' else '用户字体'
        
        self._preview_font_inline(font_type, font_name, font_file)
    
    def _preview_font_inline(self, font_type, font_name, font_file):
        """在右侧面板预览字体"""
        # 更新标题
        self.preview_title_label.config(text=f"字体预览 - {font_name}")
        
        # 清空并启用文本框
        self.preview_text.config(state=NORMAL)
        self.preview_text.delete(1.0, END)
        
        # 从font_name中提取实际的字体族名称
        # font_name格式如: "Arial (TrueType)", "微软雅黑 (TrueType)" 等
        # 需要去掉后面的类型标识
        actual_font_name = font_name
        for suffix in [' (TrueType)', ' (OpenType)', ' (TTC)', ' & ', '(TrueType)', '(OpenType)']:
            if suffix in actual_font_name:
                actual_font_name = actual_font_name.split(suffix)[0].strip()
                break
        
        # 插入字体信息
        self.preview_text.insert(END, f"字体名称: {font_name}\n", 'info')
        self.preview_text.insert(END, f"文件名: {font_file}\n", 'info')
        self.preview_text.insert(END, "\n" + "=" * 50 + "\n\n", 'separator')
        
        # 字母和数字
        self.preview_text.insert(END, "ABCDEFGHIJKLMNOPQRSTUVWXYZ\n", 'sample')
        self.preview_text.insert(END, "abcdefghijklmnopqrstuvwxyz\n", 'sample')
        self.preview_text.insert(END, "0123456789\n\n", 'sample')
        
        # 示例文字
        self.preview_text.insert(END, "中文字体预览示例\n", 'sample')
        self.preview_text.insert(END, "快速的棕色狐狸跳过懒狗。\n", 'sample')
        self.preview_text.insert(END, "The quick brown fox jumps over the lazy dog.\n\n", 'sample')
        
        # 不同字号
        self.preview_text.insert(END, "不同字号预览:\n\n", 'info')
        
        sizes = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36]
        
        for size in sizes:
            self.preview_text.insert(END, f"字号 {size}pt: ", 'info')
            self.preview_text.insert(END, f"样例文字 Sample Text Abc123\n", f'size{size}')
            try:
                # 使用提取出的实际字体名称
                self.preview_text.tag_config(f'size{size}', font=(actual_font_name, size))
            except Exception as e:
                # 如果字体名称不对，尝试使用默认字体
                self.preview_text.tag_config(f'size{size}', font=("Arial", size))
        
        # 配置样式标签
        self.preview_text.tag_config('info', foreground='#666', font=("Arial", 10))
        self.preview_text.tag_config('separator', foreground='#ccc')
        try:
            self.preview_text.tag_config('sample', font=(actual_font_name, 12))
        except:
            self.preview_text.tag_config('sample', font=("Arial", 12))
        
        self.preview_text.config(state=DISABLED)
    
    def _install_font(self):
        """安装字体"""
        # 创建选择对话框
        choice_dialog = Toplevel(self.root)
        choice_dialog.title("选择安装类型")
        choice_dialog.geometry("400x200")
        choice_dialog.transient(self.root)
        choice_dialog.grab_set()
        
        # 居中显示
        choice_dialog.geometry("+%d+%d" % (
            self.root.winfo_x() + self.root.winfo_width()//2 - 200,
            self.root.winfo_y() + self.root.winfo_height()//2 - 100
        ))
        
        selected_type = StringVar(value='user')
        
        frame = Frame(choice_dialog, padx=20, pady=20)
        frame.pack(fill=BOTH, expand=True)
        
        Label(frame, text="请选择字体安装类型：", font=("Arial", 11, "bold")).pack(pady=10)
        
        Radiobutton(frame, text="用户字体（推荐，无需管理员权限）", 
                   variable=selected_type, value='user', 
                   font=("Arial", 10)).pack(anchor=W, pady=5)
        
        Radiobutton(frame, text="系统字体（需要管理员权限）", 
                   variable=selected_type, value='system',
                   font=("Arial", 10)).pack(anchor=W, pady=5)
        
        btn_frame = Frame(frame)
        btn_frame.pack(pady=15)
        
        install_type = [None]  # 用列表来存储选择结果
        
        def on_ok():
            install_type[0] = selected_type.get()
            choice_dialog.destroy()
        
        def on_cancel():
            choice_dialog.destroy()
        
        Button(btn_frame, text="确定", command=on_ok, width=10, 
               bg="#4CAF50", fg="white").pack(side=LEFT, padx=5)
        Button(btn_frame, text="取消", command=on_cancel, width=10,
               bg="#f44336", fg="white").pack(side=LEFT, padx=5)
        
        self.root.wait_window(choice_dialog)
        
        if install_type[0] is None:
            return
        
        # 选择字体文件
        file_paths = filedialog.askopenfilenames(
            title="选择字体文件",
            filetypes=[
                ("字体文件", "*.ttf *.otf *.ttc"),
                ("TrueType字体", "*.ttf"),
                ("OpenType字体", "*.otf"),
                ("所有文件", "*.*")
            ]
        )
        
        if not file_paths:
            return
        
        success_count = 0
        fail_count = 0
        errors = []
        
        for font_path in file_paths:
            success, message = self.font_manager.install_font(font_path, install_type[0])
            if success:
                success_count += 1
            else:
                fail_count += 1
                errors.append(f"{os.path.basename(font_path)}: {message}")
        
        # 刷新列表
        self._load_fonts()
        
        # 显示结果
        result_msg = f"成功安装 {success_count} 个字体"
        if fail_count > 0:
            result_msg += f"\n失败 {fail_count} 个"
            if errors:
                result_msg += "\n\n错误详情:\n" + "\n".join(errors[:5])
                if len(errors) > 5:
                    result_msg += f"\n... 还有 {len(errors) - 5} 个错误"
        
        if success_count > 0:
            messagebox.showinfo("安装完成", result_msg)
        else:
            messagebox.showerror("安装失败", result_msg)
    
    def _select_all_fonts(self):
        """全选所有字体"""
        # 获取所有项目
        all_items = self.tree.get_children()
        if not all_items:
            messagebox.showinfo("提示", "当前列表中没有字体")
            return
        
        # 选中所有项目
        self.tree.selection_set(all_items)
        
        # 更新状态栏
        self.status_var.set(f"已选中 {len(all_items)} 个字体")
    
    def _uninstall_font(self):
        """批量卸载字体"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要卸载的字体")
            return
        
        # 根据当前Tab判断字体类型
        font_type = '系统字体' if self.current_tab == 'system' else '用户字体'
        
        # 检查是否为系统字体 - 系统字体不允许卸载（实际上按钮已禁用）
        if font_type == '系统字体':
            messagebox.showerror(
                "操作受限", 
                "无法卸载系统字体\n\n"
                "系统字体由Windows管理，不支持通过此工具卸载。"
            )
            return
        
        # 收集要卸载的字体信息
        fonts_to_uninstall = []
        for item_id in selection:
            item = self.tree.item(item_id)
            # 现在values顺序是 (name, file, path)
            font_name = item['values'][0]
            font_file = item['values'][1]
            font_path = item['values'][2]
            fonts_to_uninstall.append({
                'name': font_name,
                'file': font_file,
                'path': font_path
            })
        
        # 确认对话框
        count = len(fonts_to_uninstall)
        if count == 1:
            # 单个字体
            font = fonts_to_uninstall[0]
            confirm_msg = (
                f"确定要卸载字体 '{font['name']}' 吗？\n\n"
                f"文件: {font['file']}\n"
                f"路径: {font['path']}"
            )
        else:
            # 多个字体
            confirm_msg = (
                f"确定要批量卸载选中的 {count} 个字体吗？\n\n"
                f"即将卸载的字体包括:\n"
            )
            # 显示前5个字体名称
            for i, font in enumerate(fonts_to_uninstall[:5]):
                confirm_msg += f"• {font['name']}\n"
            if count > 5:
                confirm_msg += f"... 还有 {count - 5} 个字体\n"
        
        if not messagebox.askyesno("确认批量卸载", confirm_msg):
            return
        
        # 批量卸载
        success_count = 0
        fail_count = 0
        errors = []
        
        for font in fonts_to_uninstall:
            success, message = self.font_manager.uninstall_font(
                font['name'], font['file'], font_type
            )
            if success:
                success_count += 1
            else:
                fail_count += 1
                errors.append(f"{font['name']}: {message}")
        
        # 刷新列表
        self._load_fonts()
        # 清空预览
        self._show_preview_hint()
        
        # 显示结果
        result_msg = f"成功卸载 {success_count} 个字体"
        if fail_count > 0:
            result_msg += f"\n失败 {fail_count} 个"
            if errors:
                result_msg += "\n\n错误详情:\n" + "\n".join(errors[:5])
                if len(errors) > 5:
                    result_msg += f"\n... 还有 {len(errors) - 5} 个错误"
        
        if success_count > 0:
            messagebox.showinfo("卸载完成", result_msg)
        else:
            messagebox.showerror("卸载失败", result_msg)
    


def main():
    """主函数"""
    root = Tk()
    app = FontManagerGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()

