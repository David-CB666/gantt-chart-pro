#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI 橫道圖生成器 v1.0
=====================
雙擊運行，可視化配置後一鍵生成橫道圖。

無需安裝額外依賴 — 基於 Python 內建 tkinter。

也可從 gen_gantt.py 雙擊啟動（無 CLI 參數時自動進入 GUI 模式）。
"""

import json
import os
import sys
import threading
from datetime import date, timedelta
from tkinter import (
    Tk, Toplevel, Frame, Label, Entry, Button, Checkbutton, Listbox,
    Radiobutton, StringVar, IntVar, BooleanVar, Text, Scrollbar,
    filedialog, messagebox, ttk, END, DISABLED, NORMAL, HORIZONTAL, WORD,
)
from tkinter.font import Font

# ---- Ensure script dir on path for local imports ----
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

# 加載配置
def _load_json(name):
    p = os.path.join(SCRIPT_DIR, name)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

CFG = _load_json("config_v2.json")
STY = _load_json("gantt_styles.json")

SCHEMES = CFG.get("schemes", {})
I18N = CFG.get("i18n", {})
LOCALES = I18N.get("locales", ["zh_TW"])
TRANSLATIONS = I18N.get("translations", {})
HOLIDAYS_DEFAULT = CFG.get("calendar", {}).get("holidays", [])

# 配色預覽色塊
SCHEME_PREVIEW_KEYS = ["HEADER_BG", "TASK_BG_END", "MILESTONE_BG"]
STAGE_PREVIEW_KEYS = ["A", "B", "C", "D", "E"]


# ================================================================
#  主窗口
# ================================================================

class GanttGUI:
    def __init__(self, root: Tk, default_edf: str = ""):
        self.root = root
        self.root.title("橫道圖生成器 v3.1 — 可視化配置面板")
        self.root.geometry("680x760")
        self.root.minsize(620, 680)
        self.root.configure(bg="#F5F6FA")
        self.root.resizable(True, True)

        # ---- Style ----
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self._configure_styles()

        # ---- Vars ----
        self.edf_path = StringVar(value=default_edf)
        self.project_name = StringVar(value="")
        self.subtitle = StringVar(value="施工進度橫道圖")
        self.contractor = StringVar(value="")
        self.cal_type = StringVar(value="calendar_days")
        self.scheme_name = StringVar(value=CFG.get("active_scheme", "blue_pro"))
        self.locale = StringVar(value=I18N.get("current_locale", "zh_TW"))
        self.gran_mode = StringVar(value=CFG.get("granularity", {}).get("manual", "auto"))
        self.days_per_col = StringVar(value=str(CFG.get("granularity", {}).get("days_per_col", 7)))
        self.output_path = StringVar(value="")
        self.status_text = StringVar(value="就緒 — 請選擇 EDF.xlsx 數據源")
        self.show_weekends = BooleanVar(value=True)

        # ---- Build UI ----
        self._build_source_section()
        self._build_project_section()
        self._build_calendar_section()
        self._build_display_section()
        self._build_output_section()
        self._build_action_bar()
        self._build_status_bar()

        # ---- Key bindings ----
        self.root.bind("<Return>", lambda e: self._generate())
        self.root.bind("<Escape>", lambda e: self.root.destroy())

        # ---- Focus ----
        self.edf_entry.focus_set()

    # ============================================================
    #  Styles
    # ============================================================

    def _configure_styles(self):
        self.style.configure("Title.TLabel", font=("Microsoft YaHei UI", 13, "bold"),
                             foreground="#1E3B4A", background="#F5F6FA")
        self.style.configure("Section.TLabel", font=("Microsoft YaHei UI", 10, "bold"),
                             foreground="#2C3E50", background="#F5F6FA")
        self.style.configure("Hint.TLabel", font=("Microsoft YaHei UI", 8),
                             foreground="#8899A6", background="#F5F6FA")
        self.style.configure("Card.TFrame", background="#FFFFFF", relief="solid", borderwidth=1)
        self.style.configure("Generate.TButton",
                             font=("Microsoft YaHei UI", 11, "bold"),
                             padding=(30, 8))
        self.style.configure("Status.TLabel",
                             font=("Microsoft YaHei UI", 9),
                             foreground="#546E7A", background="#E8ECF1")

    # ============================================================
    #  Section 1: Data Source
    # ============================================================

    def _build_source_section(self):
        card = ttk.Frame(self.root, style="Card.TFrame")
        card.pack(fill="x", padx=14, pady=(12, 4), ipady=6)

        ttk.Label(card, text="📂 數據來源", style="Section.TLabel").pack(anchor="w", padx=12, pady=(8, 2))
        ttk.Label(card, text="選擇 EDF.xlsx 工程進度表，或現有 gantt_config.json",
                  style="Hint.TLabel").pack(anchor="w", padx=12)

        row = Frame(card, bg="#FFFFFF")
        row.pack(fill="x", padx=12, pady=(4, 10))
        self.edf_entry = Entry(row, textvariable=self.edf_path, font=("Consolas", 10),
                               relief="solid", borderwidth=1)
        self.edf_entry.pack(side="left", fill="x", expand=True, ipady=3)
        self.edf_entry.bind("<Button-3>", self._paste_from_clipboard)

        Button(row, text="瀏覽…", command=self._browse_edf, bg="#2E86AB", fg="white",
               font=("Microsoft YaHei UI", 9), padx=12, relief="flat", cursor="hand2",
               activebackground="#2471A3", activeforeground="white").pack(side="left", padx=(8, 0))

    def _browse_edf(self):
        path = filedialog.askopenfilename(
            title="選擇數據源",
            filetypes=[
                ("Excel / JSON", "*.xlsx;*.xls;*.xlsm;*.json"),
                ("Excel 文件", "*.xlsx;*.xls;*.xlsm"),
                ("JSON 文件", "*.json"),
                ("所有文件", "*.*"),
            ],
            initialdir=SCRIPT_DIR,
        )
        if path:
            self.edf_path.set(path)
            # 自動推導項目名
            base = os.path.splitext(os.path.basename(path))[0]
            if not self.project_name.get():
                self.project_name.set(base)
            if not self.output_path.get():
                self.output_path.set(os.path.join(SCRIPT_DIR, f"{base}-施工進度橫道圖.xlsx"))

    def _paste_from_clipboard(self, event):
        try:
            clip = self.root.clipboard_get()
            if os.path.exists(clip.strip()):
                self.edf_path.set(clip.strip())
        except Exception:
            pass

    # ============================================================
    #  Section 2: Project Info
    # ============================================================

    def _build_project_section(self):
        card = ttk.Frame(self.root, style="Card.TFrame")
        card.pack(fill="x", padx=14, pady=4, ipady=6)

        ttk.Label(card, text="📋 項目信息", style="Section.TLabel").pack(anchor="w", padx=12, pady=(8, 4))

        fields = [
            ("項目名稱", self.project_name, "自動從文件名推導"),
            ("副標題", self.subtitle, "顯示在橫道圖標題行"),
            ("承建商", self.contractor, "可選"),
        ]
        for label, var, hint in fields:
            self._labeled_input(card, label, var, hint)

    def _labeled_input(self, parent, label, var, hint=""):
        row = Frame(parent, bg="#FFFFFF")
        row.pack(fill="x", padx=12, pady=2)
        ttk.Label(row, text=label, width=10, anchor="e").pack(side="left", padx=(0, 6))
        e = Entry(row, textvariable=var, font=("Microsoft YaHei UI", 10),
                  relief="solid", borderwidth=1)
        e.pack(side="left", fill="x", expand=True, ipady=2)
        if hint:
            ttk.Label(row, text=hint, style="Hint.TLabel").pack(side="left", padx=(6, 0))

    # ============================================================
    #  Section 3: Calendar
    # ============================================================

    def _build_calendar_section(self):
        card = ttk.Frame(self.root, style="Card.TFrame")
        card.pack(fill="x", padx=14, pady=4, ipady=6)

        ttk.Label(card, text="📅 工期計算方式", style="Section.TLabel").pack(anchor="w", padx=12, pady=(8, 4))

        radio_row = Frame(card, bg="#FFFFFF")
        radio_row.pack(fill="x", padx=12, pady=2)
        Radiobutton(radio_row, text="日曆天（含週末與假日）", variable=self.cal_type,
                    value="calendar_days", bg="#FFFFFF", font=("Microsoft YaHei UI", 10),
                    activebackground="#E8F0FE", command=self._on_cal_change).pack(anchor="w")
        Radiobutton(radio_row, text="工作天（週一至週五，排除假日）", variable=self.cal_type,
                    value="workdays", bg="#FFFFFF", font=("Microsoft YaHei UI", 10),
                    activebackground="#E8F0FE", command=self._on_cal_change).pack(anchor="w", pady=(4, 0))

    def _on_cal_change(self):
        pass  # 可擴充

    # ============================================================
    #  Section 4: Display / Scheme
    # ============================================================

    def _build_display_section(self):
        card = ttk.Frame(self.root, style="Card.TFrame")
        card.pack(fill="x", padx=14, pady=4, ipady=6)

        ttk.Label(card, text="🎨 顯示設置", style="Section.TLabel").pack(anchor="w", padx=12, pady=(8, 4))

        # ------- Color scheme selector with preview --------
        scheme_frame = Frame(card, bg="#FFFFFF")
        scheme_frame.pack(fill="x", padx=12, pady=2)
        ttk.Label(scheme_frame, text="配色方案", width=10, anchor="e").pack(side="left", padx=(0, 6))

        self.scheme_cb = ttk.Combobox(scheme_frame, textvariable=self.scheme_name,
                                       state="readonly", font=("Microsoft YaHei UI", 10))
        self.scheme_cb["values"] = list(SCHEMES.keys())
        self.scheme_cb.pack(side="left", ipadx=2)
        self.scheme_cb.bind("<<ComboboxSelected>>", self._on_scheme_change)

        self.scheme_preview = Frame(scheme_frame, bg="#FFFFFF", width=180, height=24)
        self.scheme_preview.pack(side="left", padx=(12, 0))
        self._update_scheme_preview()

        # ------- Locale --------
        loc_frame = Frame(card, bg="#FFFFFF")
        loc_frame.pack(fill="x", padx=12, pady=4)
        ttk.Label(loc_frame, text="界面語言", width=10, anchor="e").pack(side="left", padx=(0, 6))

        locale_map = {"zh_TW": "繁體中文", "en_US": "English", "pt_BR": "Português"}
        self.locale_cb = ttk.Combobox(loc_frame, textvariable=self.locale,
                                       state="readonly", font=("Microsoft YaHei UI", 10))
        self.locale_cb["values"] = list(locale_map.keys())
        # 顯示 label 而非 key
        self.locale_cb.pack(side="left", ipadx=2)

        # ------- 時間顆粒度模式 --------
        gran_frame = Frame(card, bg="#FFFFFF")
        gran_frame.pack(fill="x", padx=12, pady=4)
        ttk.Label(gran_frame, text="時間顆粒度", width=10, anchor="e").pack(side="left", padx=(0, 6))
        self.gran_cb = ttk.Combobox(gran_frame, textvariable=self.gran_mode,
                                    state="readonly", font=("Microsoft YaHei UI", 10), width=14)
        self.gran_cb["values"] = ["auto", "day", "week_grouped", "week", "compressed", "month"]
        self.gran_cb.pack(side="left", ipadx=2)
        ttk.Label(gran_frame, text="壓縮日/列", width=8, anchor="e").pack(side="left", padx=(10, 2))
        self.dpc_entry = Entry(gran_frame, textvariable=self.days_per_col,
                               font=("Microsoft YaHei UI", 10), width=5)
        self.dpc_entry.pack(side="left")

        # ------- 週末標註 --------
        chk_frame = Frame(card, bg="#FFFFFF")
        chk_frame.pack(fill="x", padx=12, pady=4)
        ttk.Label(chk_frame, text="週末標註", width=10, anchor="e").pack(side="left", padx=(0, 6))
        Checkbutton(chk_frame, text="顯示週末 / 假日（灰色標註）", variable=self.show_weekends,
                    bg="#FFFFFF", font=("Microsoft YaHei UI", 10),
                    activebackground="#E8F0FE").pack(anchor="w")

    def _on_scheme_change(self, event=None):
        self._update_scheme_preview()

    def _update_scheme_preview(self):
        for w in self.scheme_preview.winfo_children():
            w.destroy()
        sname = self.scheme_name.get()
        scheme = SCHEMES.get(sname, {})
        colors = scheme.get("colors", {})
        stage_colors = scheme.get("stage_colors", {})

        # Show header + task + milestone swatches
        preview_keys = [
            ("H", colors.get("HEADER_BG", "CCC")),
            ("T", colors.get("TASK_BG_END", "CCC")),
            ("M", colors.get("MILESTONE_BG", "CCC")),
        ]
        for label, color in preview_keys:
            b = Frame(self.scheme_preview, bg=f"#{color}", width=18, height=18, relief="solid",
                      borderwidth=1)
            b.pack(side="left", padx=1)
            Label(b, text=label, bg=f"#{color}", fg="white",
                  font=("Arial", 7, "bold")).place(relx=0.5, rely=0.5, anchor="center")

        # Stage colors
        for cat in ["A", "B", "C", "D", "E"]:
            sc = stage_colors.get(cat, "CCC")
            b = Frame(self.scheme_preview, bg=f"#{sc}", width=12, height=18, relief="solid",
                      borderwidth=1)
            b.pack(side="left", padx=1)

    # ============================================================
    #  Section 5: Output
    # ============================================================

    def _build_output_section(self):
        card = ttk.Frame(self.root, style="Card.TFrame")
        card.pack(fill="x", padx=14, pady=4, ipady=6)

        ttk.Label(card, text="💾 輸出設置", style="Section.TLabel").pack(anchor="w", padx=12, pady=(8, 4))

        row = Frame(card, bg="#FFFFFF")
        row.pack(fill="x", padx=12, pady=(2, 10))
        ttk.Label(row, text="輸出路徑", width=10, anchor="e").pack(side="left", padx=(0, 6))
        Entry(row, textvariable=self.output_path, font=("Consolas", 10),
              relief="solid", borderwidth=1).pack(side="left", fill="x", expand=True, ipady=3)
        Button(row, text="另存為…", command=self._browse_output, bg="#7D8B9A", fg="white",
               font=("Microsoft YaHei UI", 9), padx=10, relief="flat", cursor="hand2",
               activebackground="#5D6D7E", activeforeground="white").pack(side="left", padx=(8, 0))

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="保存橫道圖",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            initialdir=SCRIPT_DIR,
            initialfile=self.output_path.get() or "施工進度橫道圖.xlsx",
        )
        if path:
            self.output_path.set(path)

    # ============================================================
    #  Action Bar
    # ============================================================

    def _build_action_bar(self):
        bar = Frame(self.root, bg="#F5F6FA")
        bar.pack(fill="x", padx=14, pady=(8, 4))

        self.gen_btn = Button(
            bar, text="🚀 開始生成橫道圖", command=self._generate,
            bg="#1E3B4A", fg="white", font=("Microsoft YaHei UI", 12, "bold"),
            padx=40, pady=8, relief="flat", cursor="hand2",
            activebackground="#2E86AB", activeforeground="white",
        )
        self.gen_btn.pack(side="left")

    # ============================================================
    #  Status Bar + Log
    # ============================================================

    def _build_status_bar(self):
        bar = Frame(self.root, bg="#E8ECF1", height=28)
        bar.pack(fill="x", side="bottom", padx=0, pady=0)
        ttk.Label(bar, textvariable=self.status_text, style="Status.TLabel",
                  background="#E8ECF1").pack(side="left", padx=12, pady=4)

        # Log area
        log_frame = Frame(self.root, bg="#1E2A33")
        log_frame.pack(fill="both", expand=True, padx=14, pady=(0, 10))
        self.log = Text(log_frame, bg="#1E2A33", fg="#C8D6E5", font=("Consolas", 9),
                        wrap=WORD, relief="flat", borderwidth=0, padx=10, pady=8,
                        state=DISABLED, height=8)
        self.log.pack(fill="both", expand=True)

        scroll = Scrollbar(self.log, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)

    # ============================================================
    #  Logging
    # ============================================================

    def _log(self, msg: str, tag: str = "info"):
        colors = {"info": "#C8D6E5", "warn": "#F5D76E", "error": "#E57373",
                  "success": "#81C784", "header": "#64B5F6"}
        self.log.configure(state=NORMAL)
        self.log.insert(END, msg + "\n", tag)
        self.log.tag_config(tag, foreground=colors.get(tag, "#C8D6E5"))
        self.log.see(END)
        self.log.configure(state=DISABLED)
        self.root.update_idletasks()

    # ============================================================
    #  Validation
    # ============================================================

    def _validate(self) -> list:
        errors = []
        src = self.edf_path.get().strip()
        if not src:
            errors.append("請選擇數據源文件（EDF.xlsx 或 gantt_config.json）")
        elif not os.path.exists(src):
            errors.append(f"數據源文件不存在：{src}")
        elif not src.lower().endswith((".xlsx", ".xls", ".xlsm", ".json")):
            errors.append("數據源僅支援 .xlsx / .xls / .json 格式")
        return errors

    # ============================================================
    #  Core: Generate
    # ============================================================

    def _generate(self):
        errors = self._validate()
        if errors:
            for e in errors:
                self._log(f"❌ {e}", "error")
                messagebox.showerror("輸入錯誤", e)
            return

        # ---- 禁用按鈕 ----
        self.gen_btn.configure(state=DISABLED, text="⏳ 生成中…")
        self.status_text.set("正在處理…")
        self._log("─" * 50, "info")
        self._log("🚀 開始生成橫道圖", "header")

        # ---- 後台線程執行 ----
        t = threading.Thread(target=self._generate_thread, daemon=True)
        t.start()

    def _generate_thread(self):
        try:
            self._run_generate()
            self.root.after(0, self._on_done)
        except Exception as e:
            import traceback
            self._log(f"❌ 生成失敗: {e}", "error")
            self._log(traceback.format_exc(), "error")
            self.root.after(0, lambda: self.gen_btn.configure(
                state=NORMAL, text="🚀 開始生成橫道圖"))
            self.root.after(0, lambda: self.status_text.set("生成失敗 — 請檢查數據格式"))
            self.root.after(0, lambda: messagebox.showerror("生成失敗",
                               f"{e}\n\n請確認 EDF 列頭包含：施工內容 / 開始日期 / 完成日期"))

    def _run_generate(self):
        from gen_gantt import GanttChart

        src = self.edf_path.get().strip()
        src_ext = os.path.splitext(src)[1].lower()
        script_dir = SCRIPT_DIR

        # ---- Step 1: Load / Import task data ----
        if src_ext in (".xlsx", ".xlsm", ".xls"):
            self._log(f"📂 導入 EDF 文件: {os.path.basename(src)}", "info")
            from edf_importer import import_edf
            task_config = import_edf(
                src,
                project_name=self.project_name.get().strip() or None,
                subtitle=self.subtitle.get().strip() or None,
            )
        else:
            self._log(f"📂 讀取 JSON 配置: {os.path.basename(src)}", "info")
            with open(src, "r", encoding="utf-8") as f:
                task_config = json.load(f)

        self._log(f"   任務: {sum(len(s['tasks']) for s in task_config['sections'])} 個", "info")

        # ---- Step 2: Load & patch system config ----
        sys_config = {}
        sys_path = os.path.join(script_dir, "config_v2.json")
        if os.path.exists(sys_path):
            with open(sys_path, "r", encoding="utf-8") as f:
                sys_config = json.load(f)

        # Apply GUI selections to sys_config
        sys_config["active_scheme"] = self.scheme_name.get()
        sys_config["i18n"]["current_locale"] = self.locale.get()

        # 時間顆粒度（含 week_grouped / compressed 新模式）
        g_cfg = sys_config.setdefault("granularity", {})
        g_cfg["manual"] = self.gran_mode.get()
        g_cfg["auto"] = (self.gran_mode.get() == "auto")
        dpc = self.days_per_col.get().strip()
        if dpc:
            try:
                g_cfg["days_per_col"] = int(dpc)
                g_cfg["days_per_col_auto"] = False
            except ValueError:
                pass
        self._log(f"   顆粒度: {self.gran_mode.get()}", "info")

        cal_type = self.cal_type.get()
        sys_config["calendar"]["type"] = cal_type
        cal_label = "日曆天" if cal_type == "calendar_days" else "工作天"
        self._log(f"   工期: {cal_label}", "info")
        self._log(f"   配色: {SCHEMES.get(self.scheme_name.get(), {}).get('name', self.scheme_name.get())}", "info")

        # Apply contractor override
        contractor = self.contractor.get().strip()
        if contractor:
            task_config.setdefault("project", {})["contractor"] = contractor

        # ---- Step 3: Load styles ----
        styles_config = {}
        sty_path = os.path.join(script_dir, "gantt_styles.json")
        if os.path.exists(sty_path):
            with open(sty_path, "r", encoding="utf-8") as f:
                styles_config = json.load(f)

        # ---- Step 4: Output path ----
        out = self.output_path.get().strip()
        if not out:
            pname = task_config.get("project", {}).get("name", "工程項目")
            out = os.path.join(script_dir, f"{pname}-施工進度橫道圖.xlsx")
        self._log(f"   輸出: {os.path.basename(out)}", "info")

        # ---- Step 5: Render ----
        self._log("🎨 渲染橫道圖…", "info")
        chart = GanttChart(task_config, out, sys_config, styles_config)

        if chart.warnings:
            for w in chart.warnings:
                self._log(f"   ⚠ {w}", "warn")

        if chart.recommendation:
            r = chart.recommendation["recommended"]
            self._log(f"   💡 建議: {r['paper']} 橫向 {r['days_per_col']} 日/列（共 {r['cols']} 列）",
                     "header")

        self._output_path = out
        self._log(f"✅ 完成！", "success")
        self._log(f"📁 {out}", "success")

    def _on_done(self):
        self.gen_btn.configure(state=NORMAL, text="🚀 開始生成橫道圖")
        self.status_text.set(f"生成完成 — {os.path.basename(self._output_path)}")

        if messagebox.askyesno("生成完成",
                               f"橫道圖已保存至：\n{self._output_path}\n\n是否立即打開？"):
            os.startfile(self._output_path)


# ================================================================
#  Entry
# ================================================================

def launch_gui(default_edf: str = ""):
    root = Tk()
    app = GanttGUI(root, default_edf)
    root.mainloop()


def main():
    # 支持命令行傳入 EDF 路徑作為默認值
    default_edf = ""
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if os.path.exists(arg):
            default_edf = arg
    launch_gui(default_edf)


if __name__ == "__main__":
    main()
