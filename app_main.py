"""Tabbed launcher for ITR tools."""

import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from itr_modules import CheckItemsTestTab, ITRAutofillTab
from itr_modules.itr_autofill_tab import APP_NAME, APP_VERSION

from itr_modules.shared.paths import OUTPUT_ROOT, REPORT_ROOT, open_in_file_explorer
from itr_modules.shared.ui_utils import apply_global_font


class TabManager:
    def __init__(self, header: tk.Frame, notebook: ttk.Notebook):
        self.header = header
        self.notebook = notebook
        self.tabs: dict[str, ttk.Frame] = {}
        self.headers: dict[str, dict[str, tk.Widget | None]] = {}
        self.factories: dict[str, tuple[str, callable]] = {}
        self.home_key = "home"
        self.active_key: str | None = None
        self.close_mark = "✕"
        self.colors = {
            "normal": "#f5f5f5",
            "hover": "#e8e8e8",
            "active": "#dcdcdc",
        }
        self.padding = {"x": 10, "y": 6, "close_pad": 6}
        self.title_font = ("Microsoft YaHei UI", 11)

    def register_tab(self, key: str, title: str, factory) -> None:
        self.factories[key] = (title, factory)

    def register_home(self, key: str, title: str, frame: ttk.Frame) -> None:
        self.home_key = key
        self.tabs[key] = frame
        self.notebook.add(frame)
        self._create_tab_header(key, title, closable=False)
        self.focus_tab(key)

    def open_tab(self, key: str) -> None:
        if key in self.tabs:
            self.focus_tab(key)
            return
        if key not in self.factories:
            return
        title, factory = self.factories[key]
        frame = ttk.Frame(self.notebook)
        content = factory(frame)
        content.pack(fill="both", expand=True)
        self.tabs[key] = frame
        self.notebook.add(frame)
        self._create_tab_header(key, title, closable=True)
        self.focus_tab(key)

    def close_tab(self, key: str) -> None:
        if key == self.home_key:
            return
        frame = self.tabs.pop(key, None)
        if frame is None:
            return
        header = self.headers.pop(key, None)
        if header and header["frame"]:
            header["frame"].destroy()
        self.notebook.forget(frame)
        self.focus_tab(self.home_key)

    def focus_tab(self, key: str) -> None:
        frame = self.tabs.get(key)
        if frame is None:
            return
        self.notebook.select(frame)
        self.active_key = key
        self._refresh_header_styles()

    def _create_tab_header(self, key: str, title: str, closable: bool) -> None:
        tab = tk.Frame(self.header, bg=self.colors["normal"], bd=1, relief="solid")
        tab.pack(side="left", padx=(0, 6), pady=2)

        tab.columnconfigure(0, weight=0)
        tab.columnconfigure(1, weight=1)
        tab.columnconfigure(2, weight=0)

        close_btn = None
        if closable:
            icon_holder = tk.Label(tab, text="", bg=self.colors["normal"])
            icon_holder.grid(row=0, column=0, padx=(self.padding["x"], 4), pady=self.padding["y"])

            label = tk.Label(tab, text=title, bg=self.colors["normal"], font=self.title_font, anchor="center")
            label.grid(row=0, column=1, sticky="ew", pady=self.padding["y"])
        else:
            icon_holder = tk.Label(tab, text="", bg=self.colors["normal"], width=4)
            icon_holder.grid(row=0, column=0, padx=(self.padding["x"], 4), pady=self.padding["y"])

            label = tk.Label(tab, text=title, bg=self.colors["normal"], font=self.title_font, anchor="center")
            label.grid(row=0, column=1, sticky="ew", pady=self.padding["y"])
            close_btn = tk.Label(tab, text="", bg=self.colors["normal"], width=4)
            close_btn.grid(row=0, column=2, padx=(4, self.padding["x"]), pady=self.padding["y"])
        if closable:
            close_btn = tk.Label(
                tab,
                text=self.close_mark,
                bg=self.colors["normal"],
                bd=1,
                relief="solid",
                padx=self.padding["close_pad"],
                pady=0,
                font=self.title_font,
                cursor="hand2",
            )
            close_btn.grid(row=0, column=2, padx=(6, self.padding["x"]), pady=2)
            close_btn.bind("<Button-1>", lambda _e, k=key: self.close_tab(k))

        def on_enter(_event=None):
            if self.active_key == key:
                return
            self._set_tab_style(tab, label, icon_holder, close_btn, "hover")

        def on_leave(_event=None):
            if self.active_key == key:
                return
            self._set_tab_style(tab, label, icon_holder, close_btn, "normal")

        def on_click(_event=None):
            self.focus_tab(key)

        for widget in (tab, label, icon_holder):
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
            widget.bind("<Button-1>", on_click)
            widget.bind("<Button-3>", lambda _e: None)

        self.headers[key] = {
            "frame": tab,
            "label": label,
            "icon": icon_holder,
            "close": close_btn,
        }
        self._set_tab_style(tab, label, icon_holder, close_btn, "normal")

    def _set_tab_style(self, tab, label, icon_holder, close_btn, state: str) -> None:
        bg = self.colors[state]
        tab.configure(bg=bg)
        label.configure(bg=bg)
        icon_holder.configure(bg=bg)
        if close_btn:
            close_btn.configure(bg=bg, activebackground=self.colors["hover"])

    def _refresh_header_styles(self) -> None:
        for key, widgets in self.headers.items():
            state = "active" if key == self.active_key else "normal"
            tab = widgets["frame"]
            label = widgets["label"]
            icon_holder = widgets["icon"]
            close_btn = widgets["close"]
            self._set_tab_style(tab, label, icon_holder, close_btn, state)


def open_folder(path: str) -> None:
    open_in_file_explorer(Path(path))


def open_help(parent: tk.Misc) -> None:
    win = tk.Toplevel(parent)
    win.title("使用说明")
    win.geometry("1120x760")

    container = ttk.Frame(win, padding=10)
    container.pack(fill="both", expand=True)
    container.columnconfigure(1, weight=1)
    container.rowconfigure(0, weight=1)

    sections = {
        "软件介绍": (
            "适用场景：\n"
            "- 把 ITR PDF 表格字段自动填充为 Excel 台账数据。\n"
            "- 对无法自动识别的字段，允许人工校对后再导出。\n\n"
            "核心产出：\n"
            "- output/ 下导出的 PDF（填好字段）\n"
            "- report/ 下的报告（定位问题、统计空字段）"
        ),
        "表头预填（ITR Autofill）": (
            "准备工作：\n"
            "1）准备 ITR PDF（可多套、多文件）\n"
            "2）准备 Excel 台账（.xlsx）\n\n"
            "操作步骤：\n"
            "1）进入“预设管理”选择或新建预设\n"
            "2）设置 Excel 表头行、匹配键、字段映射\n"
            "3）先做“PDF 定位测试（画框）”确认定位\n"
            "4）回到主界面选择 Excel 和 PDF\n"
            "5）点击“解析&预填”，检查左侧 ITR 列表\n"
            "6）在右侧列表人工修改字段（如 Serial Number）\n"
            "7）点击“导出填好的PDF + report.xlsx”\n\n"
            "输出在哪里：\n"
            "- output/itr_autofill/filled/<batch>/\n"
            "- report/itr_autofill/<batch>/report.xlsx\n\n"
            "常见错误处理：\n"
            "- 解析无结果：检查匹配键正则是否能在 PDF 中抓到 Tag。\n"
            "- 字段为空：检查 excel_col_norm 是否与表头归一化一致。\n"
            "- 定位偏移：先使用“测试PDF定位(画框)”确认位置。"
        ),
        "NA 自动勾选（NA Check）": (
            "准备工作：\n"
            "1）准备 ITR PDF（可多选）\n\n"
            "操作步骤：\n"
            "1）点击“批量导入 PDF”\n"
            "2）点击“解析（抓锚点）”完成结构识别\n"
            "3）需要验证时点击“测试（生成框图 PDF）”\n"
            "4）确认无误后点击“打勾（NA）”\n\n"
            "输出在哪里：\n"
            "- output/na_check/test/<batch>/（测试框图）\n"
            "- output/na_check/filled/<batch>/（打勾结果）\n"
            "- report/na_check/<batch>/（如有）\n\n"
            "常见错误处理：\n"
            "- 解析失败：检查 PDF 是否扫描件/不可选中文本。\n"
            "- 打勾位置偏移：先生成测试框图确认表格边界。\n"
            "- 无输出：确认是否先完成“解析”步骤。"
        ),
        "常见问题": (
            "Q：report.xlsx 有什么用？\n"
            "A：记录每套 ITR 的匹配情况、空字段清单，便于补录与核对。\n\n"
            "Q：output 和 report 目录在哪里？\n"
            "A：程序根目录下的 output/ 与 report/，主页右下角按钮可直接打开。\n\n"
            "Q：如何确认匹配键正确？\n"
            "A：在预设里设置 PDF 提取正则，并通过测试 PDF 验证 Tag 是否被识别。"
        ),
    }

    nav = tk.Listbox(container, height=12)
    for title in sections:
        nav.insert(tk.END, title)
    nav.grid(row=0, column=0, sticky="ns", padx=(0, 10))

    text = ScrolledText(container, wrap="word")
    text.grid(row=0, column=1, sticky="nsew")

    def show_section(_event=None):
        sel = nav.curselection()
        if not sel:
            return
        title = nav.get(sel[0])
        text.config(state="normal")
        text.delete("1.0", tk.END)
        text.insert("1.0", sections.get(title, ""))
        text.config(state="disabled")

    nav.bind("<<ListboxSelect>>", show_section)
    nav.selection_set(0)
    show_section()


def main() -> None:
    root = tk.Tk()
    root.title(f"{APP_NAME} {APP_VERSION} - 工具合集")
    root.geometry("1500x900")
    apply_global_font(root)

    tab_header = tk.Frame(root, bg="#f5f5f5")
    tab_header.pack(fill="x", padx=12, pady=(6, 2))

    style = ttk.Style(root)
    style.layout("Hidden.TNotebook.Tab", [])
    style.layout("Hidden.TNotebook", [("Notebook.client", {"sticky": "nswe"})])
    notebook = ttk.Notebook(root, style="Hidden.TNotebook")
    notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    manager = TabManager(tab_header, notebook)
    manager.register_tab("itr_autofill", "表头预填（ITR Autofill）", ITRAutofillTab)
    manager.register_tab("check_items_test", "非防爆 CheckItems（测试）", CheckItemsTestTab)

    home = ttk.Frame(notebook)
    manager.register_home("home", "主页", home)

    home.columnconfigure(0, weight=1)
    home.rowconfigure(2, weight=1)

    title_frame = ttk.Frame(home)
    title_frame.grid(row=0, column=0, sticky="ew", padx=24, pady=(16, 8))
    ttk.Label(title_frame, text="ITR辅助填写软件", font=("Microsoft YaHei UI", 18, "bold")).pack(anchor="w")
    ttk.Label(title_frame, text="作者：马瑞泽", font=("Microsoft YaHei UI", 10)).pack(anchor="w", pady=(4, 0))

    ttk.Separator(home, orient="horizontal").grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 10))

    content = ttk.Frame(home)
    content.grid(row=2, column=0, sticky="nsew", padx=24, pady=20)
    ttk.Label(content, text="功能入口：", font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w")
    btns = ttk.Frame(content)
    btns.pack(anchor="w", pady=(12, 0))
    ttk.Button(
        btns,
        text="表头预填（ITR Autofill）",
        command=lambda: manager.open_tab("itr_autofill"),
        width=28,
    ).pack(side="left", padx=(0, 12))
    ttk.Button(
        btns,
        text="非防爆 CheckItems（测试）",
        command=lambda: manager.open_tab("check_items_test"),
        width=30,
    ).pack(side="left")

    bottom = ttk.Frame(home)
    bottom.grid(row=3, column=0, sticky="sew", padx=24, pady=20)
    bottom.columnconfigure(0, weight=1)
    action_group = ttk.Frame(bottom)
    action_group.grid(row=0, column=0, sticky="e")
    ttk.Button(action_group, text="使用说明", command=lambda: open_help(root)).pack(side="left", padx=(0, 8))
    ttk.Button(action_group, text="打开 output", command=lambda: open_folder(OUTPUT_ROOT)).pack(side="left", padx=(0, 8))
    ttk.Button(action_group, text="打开 report", command=lambda: open_folder(REPORT_ROOT)).pack(side="left")

    root.mainloop()


if __name__ == "__main__":
    main()
