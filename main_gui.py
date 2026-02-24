import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
import tkinter as tk
import desktop_manager
import sys
import os
import json
import threading
import time
import datetime
import pystray
from PIL import Image, ImageDraw, ImageTk
import base64
import io
from ttkbootstrap.icons import Icon

import os

CONFIG_FILE = "desktop_layouts.json"

class LayoutManager:
    def __init__(self, filename=CONFIG_FILE):
        self.filename = filename
        self.layouts = []
        self.load()

    def load(self):
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.layouts = data.get("layouts", [])
            except Exception as e:
                print(f"Load failed: {e}")
                self.layouts = []
        else:
            if os.path.exists("desktop_layout.json"):
                try:
                    with open("desktop_layout.json", 'r', encoding='utf-8') as f:
                        old_data = json.load(f)
                        self.layouts.append({
                            "id": str(int(time.time())),
                            "name": "默认配置",
                            "saved": True,
                            "timestamp": time.time(),
                            "data": old_data
                        })
                except:
                    pass

    def save(self):
        data = {"layouts": self.layouts}
        with open(self.filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def add_layout(self):
        new_layout = {
            "id": str(int(time.time()*1000)),
            "name": "",
            "saved": False,
            "timestamp": None,
            "data": None
        }
        self.layouts.append(new_layout)
        return new_layout

    def delete_layout(self, index):
        if 0 <= index < len(self.layouts):
            del self.layouts[index]
            self.save()

    def update_layout(self, index, name=None, data=None):
        if 0 <= index < len(self.layouts):
            if name is not None:
                self.layouts[index]["name"] = name
            if data is not None:
                self.layouts[index]["data"] = data
                self.layouts[index]["saved"] = True
                self.layouts[index]["timestamp"] = time.time()
            self.save()
            
    def move_layout(self, from_index, to_index):
        if 0 <= from_index < len(self.layouts) and 0 <= to_index < len(self.layouts):
            item = self.layouts.pop(from_index)
            self.layouts.insert(to_index, item)
            self.save()

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        self.canvas.bind('<Configure>', self._on_canvas_configure)

        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self._update_colors()
    
    def _update_colors(self):
        try:
             bg = self.master.cget('background')
             self.canvas.configure(bg=bg)
        except:
             pass

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        # Only scroll if content is larger than canvas
        if self.scrollable_frame.winfo_height() > self.canvas.winfo_height():
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

class LayoutRow(ttk.Frame):
    def __init__(self, parent, app, index, layout, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.app = app
        self.index = index
        self.layout = layout
        self.parent = parent
        
        # Style - Reduced padding for sleeker look
        self.configure(bootstyle="light", padding="2")
        
        # Inner Frame - Reduced vertical padding
        self.inner = ttk.Frame(self, bootstyle="light", padding="10 5")
        self.inner.pack(fill="x")
        self.inner.columnconfigure(2, weight=1)

        # 0. Indicator Icon (New)
        self.indicator_lbl = ttk.Label(self.inner, text="", font=("Segoe UI Emoji", 12), width=3, anchor="center", bootstyle="warning")
        self.indicator_lbl.grid(row=0, column=0, padx=(0, 5))

        # 1. Name Entry
        self.name_var = tk.StringVar(value=layout["name"])
        if layout["saved"]:
            name_entry = ttk.Entry(self.inner, textvariable=self.name_var, state="readonly", font=("Microsoft YaHei UI", 10), width=15, bootstyle="secondary")
        else:
            name_entry = ttk.Entry(self.inner, textvariable=self.name_var, font=("Microsoft YaHei UI", 10), width=15, bootstyle="primary")
            def on_name_change(*args):
                self.app.manager.update_layout(self.index, name=self.name_var.get())
            name_entry.bind("<FocusOut>", on_name_change)
            name_entry.bind("<Return>", on_name_change)
        
        name_entry.grid(row=0, column=1, padx=(0, 10), sticky="w")
        
        # 2. Status Info
        timestamp = layout.get("timestamp")
        if timestamp:
            dt = datetime.datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M')
            icon_count = len(layout["data"]["icons"]) if layout["data"] else 0
            info_text = f"📅 {dt}   📁 {icon_count}个图标"
            bootstyle = "secondary"
        else:
            info_text = "⚠️ 未保存"
            bootstyle = "warning"
            
        info_lbl = ttk.Label(self.inner, text=info_text, font=("Microsoft YaHei UI", 10), bootstyle=bootstyle)
        info_lbl.grid(row=0, column=2, sticky="w", padx=5)

        # 3. Buttons
        btn_frame = ttk.Frame(self.inner)
        btn_frame.grid(row=0, column=3, sticky="e")

        # Compact buttons with outline style
        save_btn = ttk.Button(btn_frame, text="保存", width=6, bootstyle="outline-primary",
                            command=lambda: self.app.save_action(self.index, self.name_var))
        save_btn.pack(side=LEFT, padx=2)
        
        restore_btn = ttk.Button(btn_frame, text="恢复", width=6, bootstyle="outline-success",
                               state="normal" if layout["saved"] else "disabled",
                               command=lambda: self.app.restore_action(self.index))
        restore_btn.pack(side=LEFT, padx=2)
        
        # Monitor Layout Button
        has_monitors = False
        if layout.get("saved") and layout.get("data") and layout["data"].get("monitors"):
            has_monitors = True
            
        layout_btn = ttk.Button(btn_frame, text="布局", width=6, bootstyle="outline-info",
                              state="normal" if has_monitors else "disabled",
                              command=lambda: self.app.show_saved_monitor_layout(self.index))
        layout_btn.pack(side=LEFT, padx=2)
        
        del_btn = ttk.Button(btn_frame, text="删除", width=6, bootstyle="outline-danger",
                           command=lambda: self.app.delete_action(self.index))
        del_btn.pack(side=LEFT, padx=2)

    def set_active(self, active):
        if active:
            self.indicator_lbl.configure(text="⭐") # Star
        else:
            self.indicator_lbl.configure(text="")


class DesktopLayoutApp:
    def __init__(self, root):
        self.root = root
        # self.root.title("桌面图标管理") # Already set in main
        # self.root.geometry("850x650") # Already set in main
        
        self.manager = LayoutManager()
        
        self._init_ui()
        self.refresh_list()

    def _init_ui(self):
        # Header Area
        header_frame = ttk.Frame(self.root, padding="30 20")
        header_frame.pack(fill="x")
        
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=LEFT, fill="x", expand=True)
        
        title = ttk.Label(title_frame, text="桌面图标布局管理", font=("Microsoft YaHei UI", 20, "bold"), bootstyle="primary")
        title.pack(side=LEFT)
        
        add_btn = ttk.Button(header_frame, text="➕ 新增配置", command=self.add_row, bootstyle="success", width=15)
        add_btn.pack(side=RIGHT)
        
        monitor_btn = ttk.Button(header_frame, text="🖥️ 显示器布局", command=self.show_monitor_layout, bootstyle="info", width=15)
        monitor_btn.pack(side=RIGHT, padx=10)

        # List Container
        self.list_container = ScrollableFrame(self.root, padding="20 10")
        self.list_container.pack(fill="both", expand=True)
        
        # Footer
        footer_frame = ttk.Frame(self.root, padding="20")
        footer_frame.pack(fill="x", side="bottom")
        
        self.status_var = tk.StringVar(value="准备就绪")
        self.progress_var = tk.StringVar(value="")
        
        status_lbl = ttk.Label(footer_frame, textvariable=self.status_var, bootstyle="secondary", font=("Microsoft YaHei UI", 10))
        status_lbl.pack(side=LEFT)
        
        progress_lbl = ttk.Label(footer_frame, textvariable=self.progress_var, bootstyle="info", font=("Microsoft YaHei UI", 10))
        progress_lbl.pack(side=RIGHT)

    def refresh_list(self):
        for widget in self.list_container.scrollable_frame.winfo_children():
            widget.destroy()
            
        for index, layout in enumerate(self.manager.layouts):
            row = LayoutRow(self.list_container.scrollable_frame, self, index, layout)
            row.pack(fill="x", pady=5)
            
        # Check for matching layout
        self.check_layout_match()

    def check_layout_match(self):
        try:
            current = desktop_manager.get_monitors_info()
            # Sort by x, then y
            current.sort(key=lambda x: (x['rect'][0], x['rect'][1]))
            
            for child in self.list_container.scrollable_frame.winfo_children():
                if isinstance(child, LayoutRow):
                    layout = child.layout
                    match = False
                    
                    if layout.get('saved') and layout.get('data'):
                        saved = layout['data'].get('monitors')
                        if saved and isinstance(saved, list) and len(saved) == len(current):
                            try:
                                # Convert saved rects to tuples if they are lists (JSON)
                                saved_sorted = sorted(saved, key=lambda x: (x['rect'][0], x['rect'][1]))
                                
                                is_same = True
                                for c, s in zip(current, saved_sorted):
                                    # Compare resolution and position (tuple comparison)
                                    c_res = tuple(c['resolution'])
                                    s_res = tuple(s['resolution'])
                                    c_rect = tuple(c['rect'])
                                    s_rect = tuple(s['rect'])
                                    
                                    if c_res != s_res or c_rect != s_rect:
                                        is_same = False
                                        break
                                
                                if is_same:
                                    match = True
                            except Exception as e:
                                print(f"Compare error: {e}")
                    
                    child.set_active(match)
                    
        except Exception as e:
            print(f"Layout match check failed: {e}")

    def add_row(self):
        self.manager.add_layout()
        self.refresh_list()
        
    def delete_action(self, index):
        if Messagebox.show_question("确定要删除这个配置吗？", "确认删除"):
            self.manager.delete_layout(index)
            self.refresh_list()

    def save_action(self, index, name_var):
        name = name_var.get().strip()
        if not name:
            Messagebox.show_warning("请输入配置名称！", "提示")
            return
            
        try:
            self.manager.update_layout(index, name=name)
            data = desktop_manager.get_current_layout_data()
            count = len(data['icons'])
            self.manager.update_layout(index, data=data)
            
            self.status_var.set(f"已保存: {name}")
            self.refresh_list()
            
        except Exception as e:
            Messagebox.show_error(f"保存失败: {str(e)}", "错误")

    def restore_action(self, index):
        layout = self.manager.layouts[index]
        if not layout["saved"] or not layout["data"]:
            return
        
        def run_restore():
            try:
                self.status_var.set(f"正在恢复: {layout['name']}...")
                
                def progress(current, total):
                    self.progress_var.set(f"进度: {current}/{total}")
                
                count = desktop_manager.restore_from_data(layout["data"], progress_callback=progress)
                
                self.status_var.set(f"恢复完成: {layout['name']}")
                self.progress_var.set(f"成功恢复 {count} 个图标")
                
            except Exception as e:
                self.status_var.set("恢复失败")
                self.progress_var.set(f"错误: {str(e)}")
        
        threading.Thread(target=run_restore, daemon=True).start()

    def show_saved_monitor_layout(self, index):
        layout = self.manager.layouts[index]
        if not layout.get("saved") or not layout.get("data"):
            Messagebox.show_warning("该配置未保存或数据损坏", "提示")
            return
            
        monitors = layout["data"].get("monitors")
        if not monitors:
            Messagebox.show_warning("该配置不包含显示器布局信息", "提示")
            return
            
        self.show_monitor_visualization(monitors, title=f"布局: {layout['name']}")

    def show_monitor_layout(self):
        monitors = desktop_manager.get_monitors_info()
        self.show_monitor_visualization(monitors, title="当前显示器布局")

    def show_monitor_visualization(self, monitors, title="显示器布局"):
        # Create a Toplevel window
        top = ttk.Toplevel(self.root)
        top.title(title)
        top.geometry("900x650")
        top.place_window_center()
        
        if not monitors:
            ttk.Label(top, text="无法获取显示器信息", bootstyle="danger").pack(pady=20)
            return

        # Canvas for drawing - Dark Theme for contrast
        canvas_frame = ttk.Frame(top, padding=0)
        canvas_frame.pack(fill="both", expand=True)
        
        # Modern Dark Background
        canvas = tk.Canvas(canvas_frame, bg="#2b2b2b", highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        
        # Calculate bounding box of all monitors
        min_x = min(m['rect'][0] for m in monitors)
        min_y = min(m['rect'][1] for m in monitors)
        max_x = max(m['rect'][2] for m in monitors)
        max_y = max(m['rect'][3] for m in monitors)
        
        virtual_width = max_x - min_x
        virtual_height = max_y - min_y
        
        # Draw logic (defer until canvas is resized to fit)
        def draw_layout(event=None):
            canvas.delete("all")
            
            w = canvas.winfo_width()
            h = canvas.winfo_height()
            if w <= 1 or h <= 1: return # Not ready
            
            # Calculate scale to fit with padding
            padding = 60 # More padding for "desk" feel
            avail_w = w - 2 * padding
            avail_h = h - 2 * padding
            
            if virtual_width == 0 or virtual_height == 0:
                return

            scale_x = avail_w / virtual_width
            scale_y = avail_h / virtual_height
            scale = min(scale_x, scale_y)
            
            # Center offset
            draw_w = virtual_width * scale
            draw_h = virtual_height * scale
            offset_x = (w - draw_w) / 2
            offset_y = (h - draw_h) / 2
            
            # Draw each monitor
            for i, m in enumerate(monitors):
                rect = m['rect']
                # Normalized coordinates
                rel_x = (rect[0] - min_x) * scale
                rel_y = (rect[1] - min_y) * scale
                rel_w = (rect[2] - rect[0]) * scale
                rel_h = (rect[3] - rect[1]) * scale
                
                x1 = offset_x + rel_x
                y1 = offset_y + rel_y
                x2 = x1 + rel_w
                y2 = y1 + rel_h
                
                # Bezel (Outer Frame)
                bezel = max(2, int(4 * scale)) # Dynamic bezel
                canvas.create_rectangle(x1, y1, x2, y2, fill="#1a1a1a", outline="#555555", width=1)
                
                # Screen (Inner Area)
                sx1 = x1 + bezel
                sy1 = y1 + bezel
                sx2 = x2 - bezel
                sy2 = y2 - bezel
                
                is_primary = m.get('is_primary')
                # Screen Color: Primary gets a nice blue, others grey
                screen_color = "#2980b9" if is_primary else "#7f8c8d" 
                
                canvas.create_rectangle(sx1, sy1, sx2, sy2, fill=screen_color, outline="")
                
                # Screen Glare/Reflection (Subtle top sheen)
                # canvas.create_rectangle(sx1, sy1, sx2, sy1 + (sy2-sy1)*0.4, fill="#ffffff", stipple="gray12", outline="")
                # Note: stipple is not supported well on Windows with ttkbootstrap sometimes, skipping for safety or use simpler method
                
                # Text Info
                res_text = f"{m['resolution'][0]} x {m['resolution'][1]}"
                rate_text = f"{m['refresh_rate']} Hz"
                pos_text = f"Pos: ({m['position'][0]}, {m['position'][1]})"
                
                # Center text in rect
                cx = (x1 + x2) / 2
                cy = (y1 + y2) / 2
                
                # Draw Primary Label
                if is_primary:
                    canvas.create_text(cx, cy - 30, text="★ 主显示器", font=("Microsoft YaHei UI", 11, "bold"), fill="#f1c40f")
                
                # Draw other info
                canvas.create_text(cx, cy, text=res_text, font=("Segoe UI", 10, "bold"), fill="white")
                canvas.create_text(cx, cy + 20, text=rate_text, font=("Segoe UI", 9), fill="#ecf0f1")
                canvas.create_text(cx, cy + 40, text=pos_text, font=("Segoe UI", 8), fill="#bdc3c7")
                
                # Monitor ID Tag (Corner)
                canvas.create_text(x1 + 10, y1 + 10, text=f"#{i+1}", font=("Arial", 12, "bold"), fill="#ecf0f1", anchor="nw")

        canvas.bind("<Configure>", draw_layout)

    def create_tray_icon(self):
        def show_window(icon, item):
            icon.stop()
            self.root.after(0, self.root.deiconify)

        def exit_app(icon, item):
            icon.stop()
            self.root.quit()

        # 动态构建菜单
        menu_items = []
        
        # 1. 添加已保存的配置
        for i, layout in enumerate(self.manager.layouts):
            if layout.get("saved"):
                # 使用闭包捕获 index
                def make_restore_callback(index):
                    # 使用 root.after 确保在主线程触发 restore_action (虽然 restore_action 内部又开了线程，但为了安全)
                    return lambda icon, item: self.root.after(0, lambda: self.restore_action(index))
                
                menu_items.append(pystray.MenuItem(
                    f"恢复: {layout['name']}", 
                    make_restore_callback(i)
                ))
        
        # 2. 添加分隔符 (如果有配置项)
        if menu_items:
             menu_items.append(pystray.Menu.SEPARATOR)

        # 3. 添加标准选项
        menu_items.append(pystray.MenuItem('显示主界面', show_window))
        menu_items.append(pystray.MenuItem('退出', exit_app))

        self.icon = pystray.Icon("name", self.icon_image, "桌面图标管理", pystray.Menu(*menu_items))
        self.icon.run()

    def minimize_to_tray(self):
        self.root.withdraw()
        threading.Thread(target=self.create_tray_icon, daemon=True).start()

    def on_unmap(self, event):
        if self.root.state() == 'iconic':
            self.minimize_to_tray()

def main():
    # Use 'litera' for a clean, light, slightly rounded look
    app = ttk.Window(title="桌面图标管理", themename="litera", size=(950, 450))
    app.withdraw() # 先隐藏，避免启动时出现在左上角
    app.place_window_center() # 居中
    app.deiconify() # 再显示
    app.resizable(False, False)
    
    # Try to set window icon from file if available (best for Windows Taskbar)
    # When running from PyInstaller --onefile, we might not have app.ico extracted unless we handle it.
    # But since we use --icon in PyInstaller, the EXE itself has the icon.
    # For the window title bar/taskbar at runtime, ttkbootstrap sets it via iconphoto.
    # We can reinforce it if app.ico exists.
    if os.path.exists("app.ico"):
        try:
            app.iconbitmap("app.ico")
        except Exception:
            pass

    # Use ttkbootstrap default icon for tray to match window title
    # Decode base64 icon from ttkbootstrap
    icon_data = base64.b64decode(Icon.icon)
    image = Image.open(io.BytesIO(icon_data))
    
    gui = DesktopLayoutApp(app)
    gui.icon_image = image # Store for pystray
    
    # Bind Unmap event to intercept minimization
    app.bind('<Unmap>', gui.on_unmap)
    
    app.mainloop()

if __name__ == "__main__":
    main()
