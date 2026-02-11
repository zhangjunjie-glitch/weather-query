# -*- coding: utf-8 -*-
"""
天气数据可视化应用 - 日历视图 + 点击日期查当日天气
选择路径指向项目根目录（如 H:\\zhangjunjie_stage_1），
程序自动定位到 .../RawAssets/DesignerAssets/NewDatabase/logic/weather.xlsx，并记住路径。
"""
import calendar
import json
import os
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import urllib.request
import webbrowser
from tkinter import ttk, messagebox, scrolledtext, filedialog
from datetime import date, datetime

from weather import Weather

# 相对路径：在所选根目录下拼接此路径得到 weather.xlsx
EXCEL_REL_PATH = os.path.join("RawAssets", "DesignerAssets", "NewDatabase", "logic", "weather.xlsx")

# 应用版本号，用于「检查更新」；发布前请修改并打 tag（如 v1.0.1）
__version__ = "1.0.1"
# GitHub 仓库：用于检查更新与打开发布页（用户名/仓库名）
GITHUB_REPO = "zhangjunjie-glitch/weather-query"


def _app_dir():
    """应用所在目录：打包为 exe 时用 exe 所在目录，否则用脚本所在目录"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _config_path():
    """配置文件的路径（与 weather_app / exe 同目录）"""
    return os.path.join(_app_dir(), "weather_app_config.json")


def _load_config():
    """读取配置（last_folder, save_folder），返回 (data_dict, last_folder, save_folder)"""
    p = _config_path()
    if not os.path.isfile(p):
        return {}, None, None
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        data = data if isinstance(data, dict) else {}
        return data, data.get("last_folder") or None, data.get("save_folder") or None
    except Exception:
        return {}, None, None


def _load_saved_folder():
    _, folder, _ = _load_config()
    return folder


def _save_folder(folder):
    """保存项目根目录路径（保留 save_folder）"""
    data, _, save_folder = _load_config()
    data = data if isinstance(data, dict) else {}
    data["last_folder"] = folder
    if save_folder is not None:
        data["save_folder"] = save_folder
    p = _config_path()
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _load_save_folder():
    """读取上次选择的保存文件目录，首次返回 None"""
    _, _, folder = _load_config()
    return folder


def _save_save_folder(folder):
    """保存「保存路径」到配置（保留 last_folder）"""
    data, last_folder, _ = _load_config()
    data = data if isinstance(data, dict) else {}
    data["save_folder"] = folder
    if last_folder is not None:
        data["last_folder"] = last_folder
    p = _config_path()
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# 星期表头（周日为第一天，与 calendar.monthcalendar 一致）
WEEK_HEADERS = ("日", "一", "二", "三", "四", "五", "六")

# 特殊天气：属性名 -> 包含的天气 ID 列表（用于多选查询）
SPECIAL_WEATHER_ATTRS = {
    "晴天": [101, 102, 105],
    "雨天": [201, 202, 203, 204],
    "酷暑": [104],
    "花瓣雨": [106],
    "流星雨": list(range(107, 119)),  # 107-118
    "极光": [119, 120, 121],
    "雪天": [202, 203, 204],  # 小雪、中雪、大雪
    "彩虹": [301, 302, 303, 304, 305],
}


class WeatherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("天气数据查询 · 日历  v%s" % __version__)
        self.root.geometry("1600x960")
        self.root.minsize(1600, 960)
        self.root.configure(bg="#f0f0f0")

        self.weather = None
        self._data_loaded = False
        self._last_text = ""
        self._last_file_path = None
        self._current_folder = _load_saved_folder()
        self._save_folder = _load_save_folder()
        self._compare_path_b = None  # 分支对比用路径 B，默认不展示

        self.font = ("Microsoft YaHei UI", 11)
        self.font_bold = ("Microsoft YaHei UI", 11, "bold")
        self.font_small = ("Microsoft YaHei UI", 10)
        self._cal_year = date.today().year
        self._cal_month = date.today().month
        self._day_buttons = []

        self._setup_styles()
        self._build_ui()
        # 若有已保存路径且文件存在，进入应用后自动加载
        if self._current_folder and os.path.isdir(self._current_folder):
            excel_path = os.path.join(self._current_folder, EXCEL_REL_PATH)
            if os.path.isfile(excel_path):
                self.root.after(80, self._auto_load)

    def _setup_styles(self):
        """统一放大并美化 ttk 控件样式"""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        # 通用字体与内边距
        style.configure(".", font=self.font_small, padding=4)
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", font=self.font_small, background="#f0f0f0", padding=(4, 6))
        style.configure("TLabelFrame", font=self.font_bold, padding=(12, 8))
        style.configure("TLabelFrame.Label", font=self.font_bold, padding=(0, 4))
        style.configure("TButton", font=self.font_small, padding=(10, 6))
        style.configure("TNotebook", padding=2)
        style.configure("TNotebook.Tab", font=self.font_small, padding=(12, 6))

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=16)
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=0)
        main.columnconfigure(1, weight=0)
        main.columnconfigure(2, weight=1)
        main.rowconfigure(1, weight=1)

        # ----- 顶部 -----
        top = ttk.Frame(main)
        top.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        top.columnconfigure(1, weight=1)

        ttk.Button(top, text="选择路径", command=self._on_choose_path).grid(row=0, column=0, padx=(0, 8))
        self.path_var = tk.StringVar(value=self._current_folder or "未选择（点击「选择路径」选到项目根目录，如 H:\\zhangjunjie_stage_1）")
        path_lbl = ttk.Label(top, textvariable=self.path_var, font=self.font_small)
        path_lbl.grid(row=0, column=1, padx=8, sticky="w")

        self.status_var = tk.StringVar(value="请点击「选择路径」选择项目根目录" if not self._current_folder else "正在加载…" if self._data_loaded else "已记住路径")
        ttk.Label(top, textvariable=self.status_var, font=self.font_small).grid(row=0, column=2, padx=12)
        ttk.Button(top, text="检查更新", command=self._check_update).grid(row=0, column=3, padx=4)

        # ----- 左：日历 + 保存/打开栏（上下排列） -----
        left_wrapper = ttk.Frame(main)
        left_wrapper.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        left_wrapper.columnconfigure(0, weight=1)
        left_wrapper.rowconfigure(0, weight=1)
        left_wrapper.rowconfigure(1, weight=0)

        cal_frame = ttk.LabelFrame(left_wrapper, text="  日  历  ", padding=8)
        cal_frame.grid(row=0, column=0, sticky="nsew")
        for c in range(7):
            cal_frame.columnconfigure(c, minsize=36, weight=1)

        nav = ttk.Frame(cal_frame)
        nav.grid(row=0, column=0, columnspan=7, sticky="ew", pady=(0, 4))
        nav.columnconfigure(1, weight=1)

        self.btn_prev = ttk.Button(nav, text="◀ 上月", width=9, command=self._cal_prev_month)
        self.btn_prev.grid(row=0, column=0, padx=4)
        self.cal_title_var = tk.StringVar()
        self.cal_title_lbl = ttk.Label(nav, textvariable=self.cal_title_var, font=("Microsoft YaHei UI", 12, "bold"))
        self.cal_title_lbl.grid(row=0, column=1, padx=8)
        self.btn_next = ttk.Button(nav, text="下月 ▶", width=9, command=self._cal_next_month)
        self.btn_next.grid(row=0, column=2, padx=4)

        for c, h in enumerate(WEEK_HEADERS):
            ttk.Label(cal_frame, text=h, font=self.font_small, width=2, anchor="center").grid(
                row=1, column=c, padx=2, pady=2, sticky="ew"
            )

        self.cal_grid_frame = ttk.Frame(cal_frame)
        self.cal_grid_frame.grid(row=2, column=0, columnspan=7, sticky="nsew", pady=2)
        for c in range(7):
            self.cal_grid_frame.columnconfigure(c, minsize=36, weight=1)

        self._refresh_calendar()

        save_frame = ttk.LabelFrame(left_wrapper, text="  保存 / 打开  ", padding=8)
        save_frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        save_frame.columnconfigure(0, weight=1)
        self.save_path_var = tk.StringVar(value=self._save_folder or "未选择（点击后保存文件将写入该目录）")
        ttk.Button(save_frame, text="选择保存路径", command=self._on_choose_save_path).grid(row=0, column=0, sticky="w", pady=2)
        ttk.Label(save_frame, textvariable=self.save_path_var, font=self.font_small).grid(row=1, column=0, sticky="w", padx=0, pady=(0, 4))
        ttk.Button(save_frame, text="保存当前结果到文件", command=self._save_current_result).grid(row=2, column=0, sticky="w", pady=2)
        ttk.Button(save_frame, text="打开刚刚保存的文件", command=self._open_last_saved_file).grid(row=3, column=0, sticky="w", pady=2)

        # ----- 中：天气 ID 含义（与日历、查询结果同高） -----
        mid_wrapper = ttk.Frame(main)
        mid_wrapper.grid(row=1, column=1, sticky="nsew", padx=4)
        mid_wrapper.columnconfigure(0, weight=1)
        mid_wrapper.rowconfigure(0, weight=1)
        id_frame = ttk.LabelFrame(mid_wrapper, text="  天气 ID 含义  ", padding=10)
        id_frame.grid(row=0, column=0, sticky="nsew")
        id_frame.columnconfigure(0, weight=1)
        id_frame.rowconfigure(0, weight=1)
        self.id_meanings_text = scrolledtext.ScrolledText(
            id_frame, wrap=tk.WORD, font=("Consolas", 10), width=30, state="disabled"
        )
        self.id_meanings_text.grid(row=0, column=0, sticky="nsew")
        self._set_id_meanings_placeholder()

        # ----- 右：查询结果 -----
        right = ttk.Frame(main)
        right.grid(row=1, column=2, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        self.result_title_var = tk.StringVar(value="点击左侧日历中的日期，查看该日 24 小时天气")
        result_frame = ttk.LabelFrame(right, text="  查询结果  ", padding=10)
        result_frame.grid(row=0, column=0, sticky="nsew")
        result_frame.columnconfigure(0, weight=1, minsize=380)
        result_frame.columnconfigure(1, weight=0)
        result_frame.rowconfigure(1, weight=1)
        ttk.Label(result_frame, textvariable=self.result_title_var, font=self.font_bold).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )

        result_left = ttk.Frame(result_frame)
        result_left.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        result_left.columnconfigure(0, weight=1)
        result_left.rowconfigure(0, weight=1)

        self.result_text = scrolledtext.ScrolledText(
            result_left, wrap=tk.WORD, font=("Consolas", 11), state="normal"
        )
        self.result_text.grid(row=0, column=0, sticky="nsew")

        tree_container = ttk.Frame(result_left)
        tree_container.grid(row=0, column=0, sticky="nsew")
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)
        self.result_tree = ttk.Treeview(tree_container, show="headings", height=24)
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        tree_sb = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.result_tree.yview)
        tree_sb.grid(row=0, column=1, sticky="ns")
        self.result_tree.configure(yscrollcommand=tree_sb.set)
        tree_container.grid_remove()
        tree_container.bind("<Configure>", self._on_result_tree_configure)

        self._result_tree_container = tree_container

        special_frame = ttk.LabelFrame(result_frame, text="  本日特殊天气  ", padding=8)
        special_frame.grid(row=1, column=1, sticky="nsew")
        special_frame.columnconfigure(0, weight=1)
        special_frame.rowconfigure(0, weight=1)
        self.special_weather_text = scrolledtext.ScrolledText(
            special_frame, wrap=tk.WORD, font=("Consolas", 10), width=40, state="disabled"
        )
        self.special_weather_text.grid(row=0, column=0, sticky="nsew")
        self._set_special_weather_placeholder()

        # ----- 下方：更多功能 -----
        self.notebook = ttk.Notebook(main)
        self.notebook.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        main.rowconfigure(2, weight=0)

        self._add_tab_range()
        self._add_tab_all()
        self._add_tab_find_weather_id()
        self._add_tab_special()
        self._add_tab_compare()

    def _set_id_meanings_placeholder(self):
        """未加载时显示的占位文字"""
        self.id_meanings_text.config(state="normal")
        self.id_meanings_text.delete("1.0", tk.END)
        self.id_meanings_text.insert(tk.END, "请先加载数据，此处将显示所有天气 ID 与含义。")
        self.id_meanings_text.config(state="disabled")

    def _refresh_weather_id_meanings(self):
        """根据已加载的 weather 数据填充「天气 ID 含义」"""
        self.id_meanings_text.config(state="normal")
        self.id_meanings_text.delete("1.0", tk.END)
        if not self.weather or not self._data_loaded:
            self.id_meanings_text.insert(tk.END, "请先加载数据，此处将显示所有天气 ID 与含义。")
            self.id_meanings_text.config(state="disabled")
            return
        try:
            ids = sorted(self.weather.df_weather_type["id"].dropna().unique().tolist())
            # 转为 int 再排序，避免 119.0 与 119 重复
            seen = set()
            id_list = []
            for x in ids:
                k = int(x) if isinstance(x, float) else x
                if k not in seen:
                    seen.add(k)
                    id_list.append(k)
            id_list.sort()
            # 不展示未使用的 ID：1、2、3、4、399
            exclude_ids = {1, 2, 3, 4, 399}
            id_list = [x for x in id_list if x not in exclude_ids]
            lines = []
            for wid in id_list:
                name = self.weather.get_weather_type(wid)
                lines.append(f"  {wid:>4}  →  {name}")
            self.id_meanings_text.insert(tk.END, "\n".join(lines) if lines else "无数据")
        except Exception:
            self.id_meanings_text.insert(tk.END, "加载 ID 含义时出错")
        self.id_meanings_text.config(state="disabled")

    def _refresh_calendar(self):
        """根据当前年月重绘日历格子"""
        self.cal_title_var.set(f"{self._cal_year}年 {self._cal_month}月")
        # 清空旧按钮
        for w in self.cal_grid_frame.winfo_children():
            w.destroy()
        self._day_buttons.clear()

        weeks = calendar.monthcalendar(self._cal_year, self._cal_month)
        for r, week in enumerate(weeks):
            for c, day in enumerate(week):
                if day == 0:
                    lbl = ttk.Label(self.cal_grid_frame, text="", width=2)
                    lbl.grid(row=r, column=c, padx=2, pady=2)
                else:
                    btn = ttk.Button(
                        self.cal_grid_frame, text=str(day), width=3,
                        command=lambda d=day: self._on_cal_day_click(d)
                    )
                    btn.grid(row=r, column=c, padx=2, pady=2)
                    self._day_buttons.append((self._cal_month, day, btn))

    def _cal_prev_month(self):
        if self._cal_month <= 1:
            self._cal_month = 12
            self._cal_year -= 1
        else:
            self._cal_month -= 1
        self._refresh_calendar()

    def _cal_next_month(self):
        if self._cal_month >= 12:
            self._cal_month = 1
            self._cal_year += 1
        else:
            self._cal_month += 1
        self._refresh_calendar()

    def _on_cal_day_click(self, day):
        """点击日历某日：查询当日 24 小时天气并显示在右侧；右侧同时展示该日特殊天气"""
        month = self._cal_month
        if not self._data_loaded or self.weather is None:
            self.result_title_var.set(f"{month}月{day}日 — 请先加载数据")
            self._set_result("请先选择分支并点击「加载数据」，再点击日期查询。")
            return
        try:
            _, text, _, cols, rows = self.weather.get_weather_list_by_day(month=month, day=day)
            self.result_title_var.set(f"{month}月{day}日 全天天气")
            self._last_file_path = None
            if cols and rows:
                self._set_result_table(cols, rows, text_for_save=text or "")
            else:
                self._set_result(text or "未找到该日期数据。")
            special_text = self.weather.get_special_weather_for_day(month, day)
            if special_text:
                special_text = f"  ┌─ {month}月{day}日\n{special_text}\n  └" + "─" * 10
            self._set_special_weather_content(special_text)
        except Exception as e:
            self.result_title_var.set(f"{month}月{day}日 — 查询出错")
            self._set_result(f"查询失败：{e}")

    def _add_tab_range(self):
        f = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(f, text="日期范围")
        ttk.Label(f, text="开始:").grid(row=0, column=0, padx=4, pady=6)
        self.range_sm = ttk.Combobox(f, values=list(range(1, 13)), width=5)
        self.range_sm.grid(row=0, column=1, padx=4, pady=6)
        ttk.Label(f, text="月").grid(row=0, column=2, padx=2)
        self.range_sd = ttk.Combobox(f, values=list(range(1, 32)), width=5)
        self.range_sd.grid(row=0, column=3, padx=4, pady=6)
        ttk.Label(f, text="日").grid(row=0, column=4, padx=2)
        ttk.Label(f, text="结束:").grid(row=0, column=5, padx=(14, 4), pady=6)
        self.range_em = ttk.Combobox(f, values=list(range(1, 13)), width=5)
        self.range_em.grid(row=0, column=6, padx=4, pady=6)
        ttk.Label(f, text="月").grid(row=0, column=7, padx=2)
        self.range_ed = ttk.Combobox(f, values=list(range(1, 32)), width=5)
        self.range_ed.grid(row=0, column=8, padx=4, pady=6)
        ttk.Label(f, text="日").grid(row=0, column=9, padx=2)
        ttk.Button(f, text="查询", command=self._query_range).grid(row=0, column=10, padx=8)
        ttk.Button(f, text="查询并保存", command=self._query_range_save).grid(row=0, column=11, padx=4)
        ttk.Label(f, text="使用方法：选择开始、结束的月与日，点击「查询」查看该范围内每日逐小时天气；「查询并保存」将结果保存到已选择的保存路径。", font=self.font_small, wraplength=900).grid(row=1, column=0, columnspan=12, sticky="w", padx=6, pady=(8, 0))

    def _add_tab_all(self):
        f = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(f, text="全部日期")
        ttk.Button(f, text="查询全部日期天气", command=self._query_all).grid(row=0, column=0, padx=8, pady=6)
        ttk.Button(f, text="查询并保存到文件", command=self._query_all_save).grid(row=0, column=1, padx=8, pady=6)
        ttk.Label(f, text="使用方法：点击「查询全部日期天气」可查看全年每一天的逐小时天气；「查询并保存到文件」将结果写入已选择的保存路径。", font=self.font_small, wraplength=900).grid(row=1, column=0, columnspan=12, sticky="w", padx=6, pady=(8, 0))

    def _add_tab_find_weather_id(self):
        """查找天气 ID / 时间段：支持单个或多个 ID，统一输出为「日期 + 时间段 + 天气名」"""
        f = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(f, text="查找天气ID")
        ttk.Label(f, text="天气 ID（可填一个或多个，逗号分隔，如 119 或 119,120,121）:").grid(row=0, column=0, padx=6, pady=6)
        self.find_weather_ids_var = tk.StringVar(value="119,120,121")
        ttk.Entry(f, textvariable=self.find_weather_ids_var, width=28).grid(row=0, column=1, padx=6, pady=6)
        ttk.Button(f, text="查询", command=self._query_find_weather_id).grid(row=0, column=2, padx=8)
        ttk.Button(f, text="查询并保存", command=self._query_find_weather_id_save).grid(row=0, column=3, padx=4)
        ttk.Label(f, text="使用方法：输入一个或多个天气 ID（逗号分隔，如 119 或 301,302,303），点击「查询」可查看这些天气在全年中的连续时间段（按 ID 分组、按时间排序）；「查询并保存」将结果保存到已选路径。", font=self.font_small, wraplength=900).grid(row=1, column=0, columnspan=12, sticky="w", padx=6, pady=(8, 0))

    def _add_tab_special(self):
        f = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(f, text="特殊天气")
        ttk.Label(f, text="选择天气属性（可多选）:").grid(row=0, column=0, columnspan=2, sticky="w", padx=4, pady=(0, 6))
        self.spec_attr_vars = {}
        for col, name in enumerate(SPECIAL_WEATHER_ATTRS):
            self.spec_attr_vars[name] = tk.IntVar(value=0)
            ttk.Checkbutton(f, text=name, variable=self.spec_attr_vars[name]).grid(
                row=1 + col // 4, column=col % 4, sticky="w", padx=8, pady=2
            )
        btn_row = 1 + (len(SPECIAL_WEATHER_ATTRS) + 3) // 4
        ttk.Button(f, text="查询", command=self._query_special).grid(row=btn_row, column=0, padx=8, pady=8)
        ttk.Button(f, text="查询并保存", command=self._query_special_save).grid(row=btn_row, column=1, padx=4, pady=8)
        ttk.Label(f, text="使用方法：勾选一个或多个天气属性，点击「查询」展示对应天气 ID 在全年中的时间段（按 ID 分组、按时间排序）；「查询并保存」将结果保存到已选路径。", font=self.font_small, wraplength=900).grid(row=btn_row + 1, column=0, columnspan=4, sticky="w", padx=6, pady=(4, 0))

    def _add_tab_compare(self):
        """分支对比：路径 A = 当前加载路径，路径 B = 选择路径（同加载逻辑），默认不展示"""
        f = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(f, text="分支对比")
        ttk.Label(f, text="路径 A（当前）:").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        self.compare_path_a_var = tk.StringVar(value="未加载（请先选择路径并加载）")
        ttk.Label(f, textvariable=self.compare_path_a_var, font=self.font_small).grid(row=0, column=1, columnspan=3, padx=6, pady=6, sticky="w")
        ttk.Label(f, text="路径 B:").grid(row=1, column=0, padx=6, pady=6, sticky="w")
        ttk.Button(f, text="选择路径 B", command=self._on_choose_compare_path_b).grid(row=1, column=1, padx=4, pady=6)
        self.compare_path_b_var = tk.StringVar(value="未选择")
        ttk.Label(f, textvariable=self.compare_path_b_var, font=self.font_small).grid(row=1, column=2, columnspan=2, padx=6, pady=6, sticky="w")
        ttk.Button(f, text="对比", command=self._query_compare).grid(row=2, column=0, columnspan=2, padx=8, pady=8)
        ttk.Button(f, text="对比并保存", command=self._query_compare_save).grid(row=2, column=2, columnspan=2, padx=4, pady=8)
        ttk.Label(f, text="使用方法：路径 A 为当前已加载的项目根目录，路径 B 需点击「选择路径 B」选择另一项目根目录（与加载时选择方式相同）。点击「对比」可比较两路径下 weather.xlsx 的 weatherType / weatherList 差异；「对比并保存」将报告保存到已选路径。", font=self.font_small, wraplength=900).grid(row=3, column=0, columnspan=4, sticky="w", padx=6, pady=(8, 4))

    def _ensure_loaded(self):
        if not self._data_loaded or self.weather is None:
            messagebox.showwarning("未加载数据", "请先点击「选择路径」选择项目根目录。")
            return False
        return True

    def _load_from_path(self, excel_path):
        """在后台线程中从 excel_path 加载数据，成功后更新界面"""
        self.status_var.set("正在加载…")
        self.root.update_idletasks()

        def do_load():
            try:
                w = Weather(custom_excel_path=excel_path)
                w.read_file()
                self.weather = w
                self._data_loaded = True
                def _after_load():
                    self.status_var.set("已加载，可点击日历日期或使用下方功能")
                    self._refresh_weather_id_meanings()
                    if hasattr(self, 'compare_path_a_var'):
                        self.compare_path_a_var.set(self._current_folder or "未加载")
                self.root.after(0, _after_load)
            except Exception as e:
                self.root.after(0, lambda: self._load_error(str(e)))

        threading.Thread(target=do_load, daemon=True).start()

    def _auto_load(self):
        """启动时若有已保存路径且文件存在，自动加载"""
        if not self._current_folder or not os.path.isdir(self._current_folder):
            return
        excel_path = os.path.join(self._current_folder, EXCEL_REL_PATH)
        if os.path.isfile(excel_path):
            self._load_from_path(excel_path)

    def _check_update(self):
        """检查 GitHub 最新发布版本，支持自动更新或打开发布页。"""
        if "YOUR_USERNAME" in GITHUB_REPO or "/" not in GITHUB_REPO:
            messagebox.showinfo("检查更新", "请先在代码中配置 GITHUB_REPO 为你的 用户名/仓库名。")
            return
        self.status_var.set("正在检查更新…")
        def do_check():
            try:
                url = "https://api.github.com/repos/%s/releases/latest" % GITHUB_REPO.strip()
                req = urllib.request.Request(url, headers={"Accept": "application/vnd.github.v3+json"})
                with urllib.request.urlopen(req, timeout=10) as resp:
                    data = json.loads(resp.read().decode())
                tag = (data.get("tag_name") or "").strip().lstrip("v")
                html_url = data.get("html_url") or ("https://github.com/%s/releases" % GITHUB_REPO)
                assets = data.get("assets") or []
                download_url = None
                for a in assets:
                    u = (a.get("browser_download_url") or "").strip()
                    if u and u.endswith(".zip"):
                        download_url = u
                        break
                def on_done():
                    self.status_var.set("已加载，可点击日历日期或使用下方功能" if self._data_loaded else "已记住路径")
                    if not tag:
                        messagebox.showinfo("检查更新", "无法获取最新版本号。")
                        return
                    if self._version_less(__version__, tag):
                        msg = "当前版本：%s\n最新版本：%s\n\n是否自动更新？程序将下载并替换后重启。\n选「否」则仅打开发布页。" % (__version__, tag)
                        if messagebox.askyesno("发现新版本", msg):
                            if download_url:
                                self._do_auto_update(download_url, tag)
                            else:
                                messagebox.showwarning("自动更新", "未找到可下载的 zip，请从发布页手动下载。")
                                webbrowser.open(html_url)
                        else:
                            webbrowser.open(html_url)
                    else:
                        messagebox.showinfo("检查更新", "当前已是最新版本（v%s）。" % __version__)
                self.root.after(0, on_done)
            except Exception as e:
                def on_fail():
                    self.status_var.set("已加载，可点击日历日期或使用下方功能" if self._data_loaded else "已记住路径")
                    messagebox.showerror("检查更新失败", str(e))
                self.root.after(0, on_fail)
        threading.Thread(target=do_check, daemon=True).start()

    def _do_auto_update(self, download_url, tag):
        """后台下载 zip，写 updater 脚本，运行后退出；脚本会解压覆盖并重启。"""
        def work():
            try:
                self.root.after(0, lambda: self.status_var.set("正在下载新版本…"))
                zip_path = os.path.join(tempfile.gettempdir(), "WeatherQuery-%s.zip" % tag)
                req = urllib.request.Request(download_url, headers={"Accept": "application/octet-stream"})
                with urllib.request.urlopen(req, timeout=60) as resp:
                    with open(zip_path, "wb") as f:
                        while True:
                            chunk = resp.read(65536)
                            if not chunk:
                                break
                            f.write(chunk)
                install_dir = _app_dir()
                bat = self._write_updater_bat(zip_path, install_dir)
                if not bat:
                    self.root.after(0, lambda: messagebox.showerror("自动更新", "无法创建更新脚本。"))
                    return
                CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
                subprocess.Popen(
                    ["cmd", "/c", bat, zip_path, install_dir],
                    creationflags=CREATE_NO_WINDOW,
                    shell=False,
                )
                self.root.after(0, lambda: (messagebox.showinfo("自动更新", "程序即将退出并完成更新，请稍候重新打开。"), sys.exit(0)))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("自动更新失败", str(e)))

    @staticmethod
    def _write_updater_bat(zip_path, install_dir):
        """写入用于解压覆盖并重启的 bat 到临时目录，返回 bat 路径。"""
        content = r"""@echo off
set "ZIP_PATH=%~1"
set "INSTALL_DIR=%~2"
if not defined ZIP_PATH exit /b 1
if not defined INSTALL_DIR exit /b 1
timeout /t 2 /nobreak >nul
set "TEMP_EXTRACT=%TEMP%\WeatherQuery_update%RANDOM%"
mkdir "%TEMP_EXTRACT%" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -LiteralPath '%ZIP_PATH%' -DestinationPath '%TEMP_EXTRACT%' -Force"
xcopy /E /Y /H "%TEMP_EXTRACT%\WeatherQuery\*" "%INSTALL_DIR%\" >nul 2>&1
start "" "%INSTALL_DIR%\WeatherQuery.exe"
rd /s /q "%TEMP_EXTRACT%" 2>nul
del "%ZIP_PATH%" 2>nul
"""
        try:
            fd, path = tempfile.mkstemp(suffix=".bat", prefix="WeatherQuery_updater_", text=True)
            os.write(fd, content.encode("utf-8"))
            os.close(fd)
            return path
        except Exception:
            return None

    @staticmethod
    def _version_less(current, latest):
        """比较版本号，若 current < latest 返回 True。"""
        def parse(v):
            return [int(x) if x.isdigit() else 0 for x in (v or "0").replace("-", ".").split(".")[:4]]
        return parse(current) < parse(latest)

    def _on_choose_path(self):
        """打开选择文件夹对话框，选到项目根目录；选择后自动加载"""
        initial = self._current_folder if self._current_folder and os.path.isdir(self._current_folder) else None
        folder = filedialog.askdirectory(title="选择项目根目录（如 zhangjunjie_stage_1 所在目录）", initialdir=initial)
        if not folder:
            return
        excel_path = os.path.join(folder, EXCEL_REL_PATH)
        if not os.path.isfile(excel_path):
            messagebox.showwarning(
                "路径无效",
                f"该目录下未找到 weather.xlsx：\n{excel_path}\n\n请选择包含 RawAssets\\DesignerAssets\\NewDatabase\\logic 的项目根目录。"
            )
            return
        self._current_folder = folder
        _save_folder(folder)
        self.path_var.set(folder)
        if hasattr(self, 'compare_path_a_var'):
            self.compare_path_a_var.set(folder)
        self._load_from_path(excel_path)

    def _load_error(self, msg):
        self._data_loaded = False
        self.status_var.set("加载失败")
        messagebox.showerror("加载失败", f"无法读取 weather.xlsx，请检查路径与文件是否存在。\n\n{msg}")

    def _set_special_weather_placeholder(self):
        """未选日期时右侧「本日特殊天气」占位"""
        self.special_weather_text.config(state="normal")
        self.special_weather_text.delete("1.0", tk.END)
        self.special_weather_text.insert(
            tk.END,
            "选择日历日期后，此处显示该日的特殊天气（ID：104、106、107-121、211-213、301-305）。",
        )
        self.special_weather_text.config(state="disabled")

    def _set_special_weather_content(self, text):
        """设置右侧「本日特殊天气」内容；无特殊天气时不显示（留空）。"""
        self.special_weather_text.config(state="normal")
        self.special_weather_text.delete("1.0", tk.END)
        self.special_weather_text.insert(tk.END, text or "")
        self.special_weather_text.config(state="disabled")

    def _set_result(self, text, file_path=None):
        """纯文本结果：显示在 ScrolledText，隐藏表格。"""
        self._last_text = text or ""
        self._last_file_path = file_path
        self._result_tree_container.grid_remove()
        self.result_text.grid()
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, self._last_text)
        self._set_special_weather_placeholder()

    def _result_tree_column_layout(self, col_name):
        """每列的最小宽度与权重；天气列约为原先的 1/4 占比。"""
        layout = {
            "日期": (72, 1),
            "时间段": (80, 1),
            "天气": (70, 1),
            "ID": (48, 1),
        }
        return layout.get(col_name, (80, 1))

    def _on_result_tree_configure(self, event):
        """使表格列宽填满可用宽度，天气列加宽以免长名显示不全。"""
        w = self.result_tree.winfo_width()
        cols = self.result_tree["columns"]
        if not cols or w <= 0:
            return
        total = max(w - 20, 200)
        mins = []
        weights = []
        for c in cols:
            m, g = self._result_tree_column_layout(c)
            mins.append(m)
            weights.append(g)
        sum_min = sum(mins)
        sum_weight = sum(weights)
        extra = max(0, total - sum_min)
        for i, c in enumerate(cols):
            col_width = mins[i] + (int(extra * weights[i] / sum_weight) if sum_weight else 0)
            self.result_tree.column(c, width=max(mins[i], col_width))

    def _set_result_table(self, columns, rows, text_for_save=None, file_path=None):
        """表格结果：显示在 Treeview，隐藏文本；保存时使用 text_for_save。"""
        self._last_text = text_for_save if text_for_save is not None else ""
        self._last_file_path = file_path
        self.result_text.grid_remove()
        self._result_tree_container.grid()
        for iid in self.result_tree.get_children(""):
            self.result_tree.delete(iid)
        self.result_tree["columns"] = columns
        for c in columns:
            self.result_tree.heading(c, text=c)
            minw, _ = self._result_tree_column_layout(c)
            self.result_tree.column(c, width=minw, minwidth=minw)
        for row in rows:
            self.result_tree.insert("", tk.END, values=tuple(row))
        self.root.update_idletasks()
        self._on_result_tree_configure(None)
        self._set_special_weather_placeholder()

    def _on_choose_save_path(self):
        """选择保存文件时的目标目录，并记住"""
        initial = self._save_folder if self._save_folder and os.path.isdir(self._save_folder) else None
        folder = filedialog.askdirectory(title="选择保存文件的目录", initialdir=initial)
        if not folder:
            return
        self._save_folder = folder
        _save_save_folder(folder)
        self.save_path_var.set(folder)

    def _write_to_save_folder(self, content, filename_prefix="weather"):
        """将 content 写入已选择的保存目录，文件名 prefix_时间戳.txt。未选择保存路径时提示并返回 None。"""
        if not self._save_folder or not os.path.isdir(self._save_folder):
            messagebox.showwarning("请选择保存路径", "请先点击「选择保存路径」选择保存文件的目录。")
            return None
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{filename_prefix}_{timestamp}.txt"
        path = os.path.join(self._save_folder, filename)
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            return path
        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            return None

    def _save_current_result(self):
        if not self._last_text.strip():
            messagebox.showinfo("提示", "当前没有可保存的结果。")
            return
        path = self._write_to_save_folder(self._last_text, "weather_result")
        if path:
            self._last_file_path = path
            messagebox.showinfo("保存成功", f"已写入:\n{path}")

    def _open_last_saved_file(self):
        """用系统默认程序打开刚刚保存的文件"""
        if not self._last_file_path or not os.path.isfile(self._last_file_path):
            messagebox.showinfo("提示", "尚未保存过文件，或文件已不存在。请先使用「保存当前结果到文件」或各功能的「并保存」后再打开。")
            return
        try:
            if sys.platform == "win32":
                os.startfile(self._last_file_path)
            elif sys.platform == "darwin":
                subprocess.run(["open", self._last_file_path], check=False)
            else:
                subprocess.run(["xdg-open", self._last_file_path], check=False)
        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开文件：{e}")

    def _query_range(self):
        self._query_range_impl(save_to_file=False)

    def _query_range_save(self):
        self._query_range_impl(save_to_file=True)

    def _query_range_impl(self, save_to_file=False):
        if not self._ensure_loaded():
            return
        try:
            sm = int(self.range_sm.get())
            sd = int(self.range_sd.get())
            em = int(self.range_em.get())
            ed = int(self.range_ed.get())
        except (ValueError, TypeError):
            messagebox.showwarning("输入错误", "请填写有效的开始/结束月、日。")
            return
        try:
            _, text, _, cols, rows = self.weather.get_weather_list_by_day(
                start_month=sm, start_day=sd, end_month=em, end_day=ed, save_to_file=False
            )
            self.result_title_var.set(f"日期范围 {sm}月{sd}日～{em}月{ed}日")
            if cols and rows:
                self._set_result_table(cols, rows, text_for_save=text or "")
            else:
                self._set_result(text or "范围内无数据。")
            special_text = self.weather.get_special_weather_for_range(sm, sd, em, ed)
            self._set_special_weather_content(special_text)
            if save_to_file and text:
                path = self._write_to_save_folder(text, "weather_range")
                if path:
                    self._last_file_path = path
                    messagebox.showinfo("已保存", f"已保存到:\n{path}")
        except Exception as e:
            messagebox.showerror("查询失败", str(e))

    def _query_all(self):
        if not self._ensure_loaded():
            return
        try:
            _, text, _, cols, rows = self.weather.get_weather_list_by_day(show_all=True)
            self.result_title_var.set("全部日期天气")
            if cols and rows:
                self._set_result_table(cols, rows, text_for_save=text or "")
            else:
                self._set_result(text or "无数据。")
        except Exception as e:
            messagebox.showerror("查询失败", str(e))

    def _query_all_save(self):
        if not self._ensure_loaded():
            return
        try:
            _, text, _, cols, rows = self.weather.get_weather_list_by_day(show_all=True, save_to_file=False)
            self.result_title_var.set("全部日期天气")
            if cols and rows:
                self._set_result_table(cols, rows, text_for_save=text or "")
            else:
                self._set_result(text or "无数据。")
            if text:
                path = self._write_to_save_folder(text, "weather_all")
                if path:
                    self._last_file_path = path
                    messagebox.showinfo("已保存", f"已保存到:\n{path}")
        except Exception as e:
            messagebox.showerror("查询失败", str(e))

    def _query_find_weather_id(self):
        self._query_find_weather_id_impl(save_to_file=False)

    def _query_find_weather_id_save(self):
        self._query_find_weather_id_impl(save_to_file=True)

    def _query_find_weather_id_impl(self, save_to_file=False):
        """查找天气 ID 时间段：支持单个或多个 ID，统一输出为「日期 + 时间段 + 天气名」"""
        if not self._ensure_loaded():
            return
        raw = self.find_weather_ids_var.get().strip()
        if not raw:
            messagebox.showwarning("输入错误", "请填写天气 ID，可填一个或多个（逗号分隔），如 119 或 119,120,121")
            return
        try:
            ids = [int(x.strip()) for x in raw.split(",") if x.strip()]
        except ValueError:
            messagebox.showwarning("输入错误", "ID 应为数字，多个用逗号分隔。")
            return
        if not ids:
            messagebox.showwarning("输入错误", "至少填写一个天气 ID。")
            return
        try:
            _, text, _, cols, rows = self.weather.find_weather_ids_time_ranges(weather_ids=ids, save_to_file=False)
            title = f"天气 ID {ids} 时间段" if len(ids) > 1 else f"天气 ID {ids[0]} 时间段"
            self.result_title_var.set(title)
            if cols and rows:
                self._set_result_table(cols, rows, text_for_save=text)
            else:
                self._set_result(text)
            if save_to_file and text:
                path = self._write_to_save_folder(text, "weather_ids")
                if path:
                    self._last_file_path = path
                    messagebox.showinfo("已保存", f"已保存到:\n{path}")
        except Exception as e:
            messagebox.showerror("查询失败", str(e))

    def _query_special(self):
        self._query_special_impl(save_to_file=False)

    def _query_special_save(self):
        self._query_special_impl(save_to_file=True)

    def _query_special_impl(self, save_to_file=False):
        if not self._ensure_loaded():
            return
        # 按 SPECIAL_WEATHER_ATTRS 的键顺序收集选中属性对应的 ID，并统一为 Python int
        selected = [name for name in SPECIAL_WEATHER_ATTRS if self.spec_attr_vars[name].get()]
        ids = []
        seen = set()
        for name in selected:
            for wid in SPECIAL_WEATHER_ATTRS[name]:
                wid = int(wid)
                if wid not in seen:
                    seen.add(wid)
                    ids.append(wid)
        if not ids:
            messagebox.showwarning("请选择属性", "请至少勾选一个天气属性后再查询。")
            return
        try:
            _, text, _, cols, rows = self.weather.find_weather_ids_time_ranges(weather_ids=ids, save_to_file=False)
            self.result_title_var.set(f"特殊天气（{'、'.join(selected)}）时间段")
            if cols and rows:
                self._set_result_table(cols, rows, text_for_save=text)
            else:
                self._set_result(text)
            if save_to_file and text:
                path = self._write_to_save_folder(text, "weather_special")
                if path:
                    self._last_file_path = path
                    messagebox.showinfo("已保存", f"已保存到:\n{path}")
        except Exception as e:
            messagebox.showerror("查询失败", str(e))

    def _on_choose_compare_path_b(self):
        """选择路径 B（同加载时的选择方式：选项目根目录，程序校验 weather.xlsx 是否存在）"""
        initial = self._compare_path_b if self._compare_path_b and os.path.isdir(self._compare_path_b) else None
        folder = filedialog.askdirectory(title="选择路径 B 的项目根目录（如 zhangjunjie_obt_hotfix1_1 所在目录）", initialdir=initial)
        if not folder:
            return
        excel_path = os.path.join(folder, EXCEL_REL_PATH)
        if not os.path.isfile(excel_path):
            messagebox.showwarning(
                "路径无效",
                f"该目录下未找到 weather.xlsx：\n{excel_path}\n\n请选择包含 RawAssets\\DesignerAssets\\NewDatabase\\logic 的项目根目录。"
            )
            return
        self._compare_path_b = folder
        self.compare_path_b_var.set(folder)

    def _query_compare(self):
        """分支对比：仅对比并显示结果"""
        self._query_compare_impl(save_to_file=False)

    def _query_compare_save(self):
        """分支对比：对比并保存到已选择的保存路径"""
        self._query_compare_impl(save_to_file=True)

    def _query_compare_impl(self, save_to_file=False):
        if not self._current_folder or not os.path.isdir(self._current_folder):
            messagebox.showwarning("分支对比", "路径 A 未设置。请先在顶部「选择路径」并加载数据。")
            return
        if not self._compare_path_b or not os.path.isdir(self._compare_path_b):
            messagebox.showwarning("分支对比", "请点击「选择路径 B」选择要对比的另一个项目根目录。")
            return
        path_a = self._current_folder
        path_b = self._compare_path_b
        if path_a == path_b:
            messagebox.showwarning("分支对比", "路径 A 与路径 B 不能相同。")
            return

        self.result_title_var.set("分支对比（对比中…）")
        self._set_result("正在读取两路径的 weather.xlsx，请稍候…")

        def do_compare():
            try:
                diff_dict, report, _ = Weather.compare_two_paths(
                    path_a, path_b,
                    label_a=path_a,
                    label_b=path_b,
                    save_to_file=False
                )
                def _after():
                    self.result_title_var.set("分支对比 路径 A vs 路径 B")
                    self._set_result(report)
                    if save_to_file and report:
                        path = self._write_to_save_folder(report, "weather_compare")
                        if path:
                            self._last_file_path = path
                            messagebox.showinfo("已保存", f"对比结果已保存到:\n{path}")
                self.root.after(0, _after)
            except FileNotFoundError as e:
                self.root.after(0, lambda: messagebox.showerror("分支对比", str(e)))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("分支对比", str(e)))

        threading.Thread(target=do_compare, daemon=True).start()


def main():
    root = tk.Tk()
    app = WeatherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
