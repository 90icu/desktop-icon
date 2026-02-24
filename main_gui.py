import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

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
                    self.layouts = json.load(f).get("layouts", [])
            except Exception as e:
                print(f"Load failed: {e}")
                self.layouts = []
        elif os.path.exists("desktop_layout.json"):
            try:
                with open("desktop_layout.json", 'r', encoding='utf-8') as f:
                    old_data = json.load(f)
                    self.layouts.append({
                        "id": str(int(time.time())),
                        "name": "默认配置",
                        "saved": True,
                        "timestamp": time.time(),
                        "data": old_data,
                    })
            except Exception:
                pass

    def save(self):
        with open(self.filename, 'w', encoding='utf-8') as f:
            json.dump({"layouts": self.layouts}, f, indent=2, ensure_ascii=False)

    def add_layout(self):
        new_layout = {
            "id": str(int(time.time() * 1000)),
            "name": "",
            "saved": False,
            "timestamp": None,
            "data": None,
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
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self._update_colors()

    def _update_colors(self):
        try:
            self.canvas.configure(bg=self.master.cget('background'))
        except Exception:
            pass

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        if self.scrollable_frame.winfo_height() > self.canvas.winfo_height():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class LayoutRow(ttk.Frame):
    def __init__(self, parent, app, index, layout, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.app = app
        self.index = index
        self.layout = layout

        self.configure(bootstyle="light", padding="2")
        inner = ttk.Frame(self, bootstyle="light", padding="10 5")
        inner.pack(fill="x")
        inner.columnconfigure(2, weight=1)

        self.indicator_lbl = ttk.Label(inner, text="", font=("Segoe UI Emoji", 12),
                                       width=3, anchor="center", bootstyle="warning")
        self.indicator_lbl.grid(row=0, column=0, padx=(0, 5))

        self.name_var = tk.StringVar(value=layout["name"])
        if layout["saved"]:
            name_entry = ttk.Entry(inner, textvariable=self.name_var, state="readonly",
                                   font=("Microsoft YaHei UI", 10), width=15, bootstyle="secondary")
        else:
            name_entry = ttk.Entry(inner, textvariable=self.name_var,
                                   font=("Microsoft YaHei UI", 10), width=15, bootstyle="primary")
            def on_name_change(*args):
                self.app.manager.update_layout(self.index, name=self.name_var.get())
            name_entry.bind("<FocusOut>", on_name_change)
            name_entry.bind("<Return>", on_name_change)
        name_entry.grid(row=0, column=1, padx=(0, 10), sticky="w")

        timestamp = layout.get("timestamp")
        if timestamp:
            dt = datetime.datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M')
            icon_count = len(layout["data"]["icons"]) if layout["data"] else 0
            info_text = f"📅 {dt}   📁 {icon_count}个图标"
            bootstyle = "secondary"
        else:
            info_text = "⚠️ 未保存"
            bootstyle = "warning"
        ttk.Label(inner, text=info_text, font=("Microsoft YaHei UI", 10),
                  bootstyle=bootstyle).grid(row=0, column=2, sticky="w", padx=5)

        btn_frame = ttk.Frame(inner)
        btn_frame.grid(row=0, column=3, sticky="e")

        ttk.Button(btn_frame, text="保存", width=6, bootstyle="outline-primary",
                   command=lambda: self.app.save_action(self.index, self.name_var)
                   ).pack(side=LEFT, padx=2)

        ttk.Button(btn_frame, text="恢复", width=6, bootstyle="outline-success",
                   state="normal" if layout["saved"] else "disabled",
                   command=lambda: self.app.restore_action(self.index)
                   ).pack(side=LEFT, padx=2)

        has_monitors = bool(layout.get("saved") and layout.get("data") and
                            layout["data"].get("monitors"))
        ttk.Button(btn_frame, text="布局", width=6, bootstyle="outline-info",
                   state="normal" if has_monitors else "disabled",
                   command=lambda: self.app.show_saved_monitor_layout(self.index)
                   ).pack(side=LEFT, padx=2)

        ttk.Button(btn_frame, text="删除", width=6, bootstyle="outline-danger",
                   command=lambda: self.app.delete_action(self.index)
                   ).pack(side=LEFT, padx=2)

    def set_active(self, active):
        self.indicator_lbl.configure(text="⭐" if active else "")


class DesktopLayoutApp:
    def __init__(self, root):
        self.root = root
        self.manager = LayoutManager()
        self._viz_win = None
        self._viz_canvas = None
        self._init_ui()
        self.refresh_list()

    def _init_ui(self):
        header_frame = ttk.Frame(self.root, padding="30 20")
        header_frame.pack(fill="x")

        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=LEFT, fill="x", expand=True)
        ttk.Label(title_frame, text="桌面图标布局管理",
                  font=("Microsoft YaHei UI", 20, "bold"),
                  bootstyle="primary").pack(side=LEFT)

        ttk.Button(header_frame, text="🖥️ 显示器布局",
                   command=self.show_monitor_layout,
                   bootstyle="info", width=15).pack(side=RIGHT, padx=10)
        ttk.Button(header_frame, text="➕ 新增配置",
                   command=self.add_row,
                   bootstyle="success", width=15).pack(side=RIGHT)

        self.list_container = ScrollableFrame(self.root, padding="20 10")
        self.list_container.pack(fill="both", expand=True)

        footer_frame = ttk.Frame(self.root, padding="20")
        footer_frame.pack(fill="x", side="bottom")

        self.status_var = tk.StringVar(value="准备就绪")
        self.progress_var = tk.StringVar(value="")

        ttk.Label(footer_frame, textvariable=self.status_var,
                  bootstyle="secondary", font=("Microsoft YaHei UI", 10)).pack(side=LEFT)
        ttk.Label(footer_frame, textvariable=self.progress_var,
                  bootstyle="info", font=("Microsoft YaHei UI", 10)).pack(side=RIGHT)

    def refresh_list(self):
        for widget in self.list_container.scrollable_frame.winfo_children():
            widget.destroy()
        for index, layout in enumerate(self.manager.layouts):
            LayoutRow(self.list_container.scrollable_frame, self, index, layout).pack(fill="x", pady=5)
        self.check_layout_match()

    def check_layout_match(self):
        try:
            current = sorted(desktop_manager.get_monitors_info(),
                             key=lambda x: (x['rect'][0], x['rect'][1]))
            for child in self.list_container.scrollable_frame.winfo_children():
                if not isinstance(child, LayoutRow):
                    continue
                match = False
                layout = child.layout
                if layout.get('saved') and layout.get('data'):
                    saved = layout['data'].get('monitors')
                    if saved and len(saved) == len(current):
                        saved_sorted = sorted(saved, key=lambda x: (x['rect'][0], x['rect'][1]))
                        match = all(
                            tuple(c['resolution']) == tuple(s['resolution']) and
                            tuple(c['rect']) == tuple(s['rect'])
                            for c, s in zip(current, saved_sorted)
                        )
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
        self.show_monitor_visualization(
            monitors,
            icons=layout["data"].get("icons", []),
            title=f"布局: {layout['name']}")

    def show_monitor_layout(self):
        try:
            data = desktop_manager.get_current_layout_data()
            self.show_monitor_visualization(
                data.get("monitors", []),
                icons=data.get("icons", []),
                title="当前显示器布局")
        except Exception as e:
            monitors = desktop_manager.get_monitors_info()
            self.show_monitor_visualization(monitors, title="当前显示器布局")

    def show_monitor_visualization(self, monitors, icons=None, title="显示器布局"):
        # 复用已有窗口：存在且未被关闭则直接更新，否则新建
        if self._viz_win is not None:
            try:
                self._viz_win.winfo_exists()
                self._viz_win.title(title)
                self._viz_win.lift()
                self._viz_win.focus_force()
                canvas = self._viz_canvas
                canvas.delete("all")
            except Exception:
                self._viz_win = None
                self._viz_canvas = None

        if self._viz_win is None:
            top = ttk.Toplevel(self.root)
            top.title(title)
            top.geometry("900x650")
            top.place_window_center()
            top.protocol("WM_DELETE_WINDOW", lambda: self._on_viz_close(top))
            self._viz_win = top

            if not monitors:
                ttk.Label(top, text="无法获取显示器信息", bootstyle="danger").pack(pady=20)
                return

            canvas = tk.Canvas(
                ttk.Frame(top, padding=0),
                bg="#2b2b2b", highlightthickness=0)
            canvas.master.pack(fill="both", expand=True)
            canvas.pack(fill="both", expand=True)
            self._viz_canvas = canvas
        else:
            if not monitors:
                return

        min_x = min(m['rect'][0] for m in monitors)
        min_y = min(m['rect'][1] for m in monitors)
        max_x = max(m['rect'][2] for m in monitors)
        max_y = max(m['rect'][3] for m in monitors)
        virtual_width  = max_x - min_x
        virtual_height = max_y - min_y

        idx_to_pos = {m.get('index', j): j for j, m in enumerate(monitors)}

        icons_by_monitor = {}
        if icons:
            for icon in icons:
                j = idx_to_pos.get(icon.get("monitor", 0), 0)
                icons_by_monitor.setdefault(j, []).append(icon)

        def draw_layout(event=None):
            canvas.delete("all")
            w = canvas.winfo_width()
            h = canvas.winfo_height()
            if w <= 1 or h <= 1 or virtual_width == 0 or virtual_height == 0:
                return

            padding = 60
            avail_w = w - 2 * padding
            avail_h = h - 2 * padding
            scale = min(avail_w / virtual_width, avail_h / virtual_height)

            draw_w = virtual_width * scale
            draw_h = virtual_height * scale
            offset_x = (w - draw_w) / 2
            offset_y = (h - draw_h) / 2

            for i, m in enumerate(monitors):
                rect = m['rect']
                mon_w = rect[2] - rect[0]
                mon_h = rect[3] - rect[1]

                x1 = offset_x + (rect[0] - min_x) * scale
                y1 = offset_y + (rect[1] - min_y) * scale
                x2 = x1 + mon_w * scale
                y2 = y1 + mon_h * scale

                bezel = max(2, int(4 * scale))
                canvas.create_rectangle(x1, y1, x2, y2, fill="#1a1a1a", outline="#555555", width=1)

                sx1, sy1 = x1 + bezel, y1 + bezel
                sx2, sy2 = x2 - bezel, y2 - bezel
                screen_w = sx2 - sx1

                canvas.create_rectangle(sx1, sy1, sx2, sy2,
                                        fill="#1e3a5f" if m.get('is_primary') else "#2d3436",
                                        outline="")

                cx = (x1 + x2) / 2
                base_size = max(7, min(11, int(min(x2 - x1, y2 - y1) / 28)))

                res = m.get('resolution', (mon_w, mon_h))
                pos = m.get('position', (rect[0], rect[1]))
                primary_tag = " ★" if m.get('is_primary') else ""
                info_text = (f"#{i + 1}{primary_tag}  "
                             f"{res[0]}×{res[1]}  "
                             f"{m.get('refresh_rate', '?')}Hz  "
                             f"({pos[0]}, {pos[1]})")

                info_bar_h = base_size + 12
                canvas.create_rectangle(sx1, sy1, sx2, sy1 + info_bar_h, fill="#111111", outline="")
                canvas.create_text(cx, sy1 + info_bar_h / 2, text=info_text,
                                   font=("Microsoft YaHei UI", base_size, "bold"), fill="white")

                m_icons = icons_by_monitor.get(i, [])
                if not m_icons:
                    if not icons:
                        canvas.create_text(cx, (sy1 + sy2) / 2, text="(无图标数据)",
                                           font=("Microsoft YaHei UI", 9), fill="#666666")
                    continue

                icon_area_x = sx1 + 6
                icon_area_y = sy1 + info_bar_h + 4
                icon_area_w = screen_w - 12
                icon_area_h = sy2 - icon_area_y - 4

                if icon_area_w <= 0 or icon_area_h <= 0:
                    continue

                if all('row' in ic and 'col' in ic for ic in m_icons):
                    min_col = min(ic['col'] for ic in m_icons)
                    min_row = min(ic['row'] for ic in m_icons)
                    cols = max(ic['col'] for ic in m_icons) - min_col + 1
                    rows = max(ic['row'] for ic in m_icons) - min_row + 1

                    cell = min(icon_area_w / max(cols, 1), icon_area_h / max(rows, 1))
                    dot_r = max(2, min(5, cell / 3))

                    for ic in m_icons:
                        dx = icon_area_x + (ic['col'] - min_col) * cell + cell / 2
                        dy = icon_area_y + (ic['row'] - min_row) * cell + cell / 2
                        canvas.create_oval(dx - dot_r, dy - dot_r, dx + dot_r, dy + dot_r,
                                           fill="#f39c12", outline="#e67e22", width=1)
                else:
                    for ic in m_icons:
                        rx = max(0.0, min(1.0, (ic.get("x", 0) - rect[0]) / mon_w)) if mon_w else 0
                        ry = max(0.0, min(1.0, (ic.get("y", 0) - rect[1]) / mon_h)) if mon_h else 0
                        dx = icon_area_x + rx * icon_area_w
                        dy = icon_area_y + ry * icon_area_h
                        canvas.create_oval(dx - 3, dy - 3, dx + 3, dy + 3,
                                           fill="#f39c12", outline="#e67e22", width=1)

        canvas.bind("<Configure>", draw_layout)
        # 窗口已存在时尺寸不变，手动触发一次重绘
        self._viz_win.after(50, draw_layout)

    def _on_viz_close(self, win):
        self._viz_win = None
        self._viz_canvas = None
        win.destroy()

    def create_tray_icon(self):
        def show_window(icon, item):
            icon.stop()
            self.root.after(0, self.root.deiconify)

        def exit_app(icon, item):
            icon.stop()
            self.root.quit()

        menu_items = []
        for i, layout in enumerate(self.manager.layouts):
            if layout.get("saved"):
                def make_restore_callback(index):
                    return lambda icon, item: self.root.after(0, lambda: self.restore_action(index))
                menu_items.append(pystray.MenuItem(f"恢复: {layout['name']}", make_restore_callback(i)))

        if menu_items:
            menu_items.append(pystray.Menu.SEPARATOR)
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
    app = ttk.Window(title="桌面图标管理", themename="litera", size=(950, 450))
    app.withdraw()
    app.place_window_center()
    app.deiconify()
    app.resizable(False, False)

    if os.path.exists("app.ico"):
        try:
            app.iconbitmap("app.ico")
        except Exception:
            pass

    icon_data = base64.b64decode(Icon.icon)
    image = Image.open(io.BytesIO(icon_data))

    gui = DesktopLayoutApp(app)
    gui.icon_image = image
    app.bind('<Unmap>', gui.on_unmap)
    app.mainloop()


if __name__ == "__main__":
    main()
