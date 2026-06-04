"""
Neware Batch Export Tool
Kami-style UI: warm parchment + ink-blue, serif hierarchy, EN/ZH toggle
"""

# ── Python 3.14 tkinter __repr__ recursion fix ───────────────────────
import tkinter as _tk
if hasattr(_tk.Misc, "__repr__") and _tk.Misc.__repr__ is not object.__repr__:
    _tk.Misc.__repr__ = lambda self: f"<{type(self).__name__}>"
del _tk

import customtkinter as ctk
from tkinter import filedialog, messagebox
import re, os, subprocess, threading, json, shutil, time, csv, zipfile
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── Appearance ───────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Kami palette: warm parchment + ink-blue
C_BG      = "#F5F0E8"   # warm parchment
C_CARD    = "#FDFAF4"   # lighter parchment
C_BORDER  = "#D8D0BC"   # warm border
C_ACCENT  = "#4E80C0"   # light blue
C_ACCENT2 = "#EAF1FB"   # light blue tint
C_TEXT    = "#1A1A1A"
C_MUTED   = "#7A7060"   # warm gray
C_OK      = "#2D6A4F"
C_FAIL    = "#A53030"
C_SKIP    = "#8B6914"
C_INFO    = "#4E80C0"
C_LOG_BG  = "#EDE8DC"

# Fonts: serif-led (kami style), CN fallback
FONT_SERIF = ["Georgia", "Times New Roman", "serif"]
FONT_CN    = ["TsangerJinKai02", "STSong", "SimSun", "Microsoft YaHei"]
FONT_MONO  = ["Consolas", "Courier New", "monospace"]

def serif(size, weight="normal"):
    return ctk.CTkFont(family="Georgia", size=size,
                       weight="bold" if weight == "bold" else "normal")

def mono(size):
    return ctk.CTkFont(family="Consolas", size=size)

# ── i18n ─────────────────────────────────────────────────────────────
LANG = "en"

STRINGS = {
    "en": {
        "title":        "Neware Export",
        "cycle_mode":   "CYCLE MODE",
        "seg_labels":   ["Step Default", "Chg→Dchg", "Dchg→Chg", "Custom Step"],
        "input_paths":  "INPUT PATHS",
        "scan_folder":  "Scan Folder",
        "clear":        "Clear",
        "files_none":   "0 files",
        "files_n":      "{n} file",
        "files_n_pl":   "{n} files",
        "output_dir":   "OUTPUT",
        "same_folder":  "Same folder as source",
        "custom_dir":   "Custom:",
        "options":      "OPTIONS",
        "skip_existing":"Skip existing .xlsx",
        "workers":      "Workers:",
        "email_report": "Email report on finish",
        "export_btn":   "Export",
        "exporting":    "Exporting…",
        "clear_log":    "Clear",
        "log_title":    "LOG",
        "iconf_label":  "iconf: {name} ({val})",
        "sel_count":    "{n} selected",
        "add_selected": "Add selected",
        "cancel":       "Cancel",
        "path_hint":    'Tip: select files in Explorer → right-click → "Copy as path" → paste here\n"C:\\data\\experiment_01.ndax"\n"C:\\data\\experiment_02.ndax"',
        "no_files_warn":"No .ndax paths detected.",
        "missing_title":"Missing files",
        "missing_msg":  "{n} file(s) not found. Continue with the rest?",
        "btsda_title":  "BTSDA Running",
        "btsda_msg":    "BTSDA.exe is open — iconf may be overwritten mid-export. Continue?",
        "no_outdir":    "Enter a custom output directory.",
        "log_start":    "▸  {n} files   mode: {mode}   workers: {w}",
        "log_iconf":    "   iconf CycleMode → {val}  (backup: {bak})",
        "log_iconf_err":"   ERROR — iconf write failed: {e}",
        "log_ok":       "   ✓  {name}   {kb} KB  {cyc} cycles  ({s}s)",
        "log_ok_nokb":  "   ✓  {name}   {kb} KB  ({s}s)",
        "log_skip":     "   –  {name}  skipped",
        "log_fail":     "   ✗  {name}  FAILED",
        "log_done":     "\n   Done  {ok} exported  {fail} failed  {skip} skipped  {t}s",
        "log_csv":      "   Summary → {path}",
        "log_csv_err":  "   CSV failed: {e}",
        "log_email":    "   Sending email…",
        "log_email_ok": "   ✓ Email → {addr}",
        "log_email_err":"   ✗ Email failed: {e}",
        "progress":     "{done}/{total}   {t}s",
    },
    "zh": {
        "title":        "新威批量导出",
        "cycle_mode":   "循环统计方式",
        "seg_labels":   ["工步默认", "先充后放", "先放后充", "起始工步"],
        "input_paths":  "输入路径",
        "scan_folder":  "扫描文件夹",
        "clear":        "清空",
        "files_none":   "0 个文件",
        "files_n":      "{n} 个文件",
        "files_n_pl":   "{n} 个文件",
        "output_dir":   "输出目录",
        "same_folder":  "原文件同目录",
        "custom_dir":   "指定目录：",
        "options":      "选项",
        "skip_existing":"跳过已存在的 .xlsx",
        "workers":      "并行数：",
        "email_report": "完成后发邮件简报",
        "export_btn":   "开始导出",
        "exporting":    "导出中…",
        "clear_log":    "清空",
        "log_title":    "日志",
        "iconf_label":  "iconf: {name} ({val})",
        "sel_count":    "已选 {n} 个",
        "add_selected": "添加选中",
        "cancel":       "取消",
        "path_hint":    '提示：文件管理器多选文件 → 右击 → "复制为路径" → 粘贴到此处\n"C:\\data\\experiment_01.ndax"\n"C:\\data\\experiment_02.ndax"',
        "no_files_warn":"未识别到任何 .ndax 路径。",
        "missing_title":"部分文件缺失",
        "missing_msg":  "{n} 个文件不存在，继续导出其余文件？",
        "btsda_title":  "BTSDA 正在运行",
        "btsda_msg":    "检测到 BTSDA.exe 运行中，iconf 可能被覆盖。是否继续？",
        "no_outdir":    "请输入自定义输出目录。",
        "log_start":    "▸  {n} 个文件   方式：{mode}   并行：{w}",
        "log_iconf":    "   iconf CycleMode → {val}  (备份：{bak})",
        "log_iconf_err":"   错误 — iconf 写入失败：{e}",
        "log_ok":       "   ✓  {name}   {kb} KB  {cyc} 个循环  ({s}s)",
        "log_ok_nokb":  "   ✓  {name}   {kb} KB  ({s}s)",
        "log_skip":     "   –  {name}  已跳过",
        "log_fail":     "   ✗  {name}  失败",
        "log_done":     "\n   完成  {ok} 成功  {fail} 失败  {skip} 跳过  {t}s",
        "log_csv":      "   汇总 CSV → {path}",
        "log_csv_err":  "   CSV 写入失败：{e}",
        "log_email":    "   发送邮件中…",
        "log_email_ok": "   ✓ 邮件已发送 → {addr}",
        "log_email_err":"   ✗ 邮件发送失败：{e}",
        "progress":     "{done}/{total}   {t}s",
    },
}

CYCLE_MODE_VALUES = [0, 1, 2, 3]

def T(key, **kw):
    s = STRINGS[LANG].get(key, key)
    return s.format(**kw) if kw else s

# ── Constants ────────────────────────────────────────────────────────
BTSDA_EXE    = r"E:\software\BTSClient80\BTSDAExReport.exe"
ICONF_PATH   = r"C:\Users\omwtbarca\Documents\NEWARE\BTSClient\BTSDAConfig.iconf"
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(SCRIPT_DIR, "neware_export_history.json")

SMTP_SERVER     = "smtp.your-provider.com"  # e.g. smtp.qq.com / smtp.gmail.com
SMTP_PORT       = 465                    # 465 for SSL, 587 for STARTTLS
SENDER_EMAIL    = "your_sender@email.com"
SENDER_PASSWORD = "your_smtp_auth_code"  # App password / SMTP auth code, NOT login password
RECEIVER_EMAIL  = "your_receiver@email.com"

# ── iconf ────────────────────────────────────────────────────────────
def _read_iconf():
    with open(ICONF_PATH, "rb") as f: raw = f.read()
    bom = raw[:2] if raw[:2] == b"\xff\xfe" else b""
    return bom, raw[len(bom):].decode("utf-16-le")

def _write_iconf(bom, text):
    with open(ICONF_PATH, "wb") as f:
        f.write(bom); f.write(text.encode("utf-16-le"))

def get_iconf_mode():
    _, t = _read_iconf()
    m = re.search(r"CycleMode\s*=\s*(\d+)", t)
    return int(m.group(1)) if m else 0

def set_iconf_mode(val):
    bom, t = _read_iconf()
    _write_iconf(bom, re.sub(r"(CycleMode\s*=\s*)\d+",
                             lambda m: m.group(1)+str(val), t, count=1))

# ── Helpers ──────────────────────────────────────────────────────────
def parse_paths(text):
    paths = []
    for line in text.splitlines():
        s = line.strip()
        if not s or s.startswith("#"): continue
        for p in re.findall(r'r?"([^"]+\.ndax)"', s, re.IGNORECASE):
            if p and p not in paths: paths.append(p)
    return paths

def is_btsda_running():
    r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq BTSDA.exe"],
                       capture_output=True, text=True, creationflags=0x08000000)
    return "BTSDA.exe" in r.stdout

def get_cycle_count(xlsx):
    try:
        with zipfile.ZipFile(xlsx) as z:
            wb   = z.read("xl/workbook.xml").decode("utf-8", errors="replace")
            shs  = re.findall(r'<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"', wb)
            rid  = next((r for n, r in shs if n == "cycle"), None)
            if not rid: return None
            rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="replace")
            tgt  = re.search(rf'Id="{re.escape(rid)}"[^>]+Target="([^"]+)"', rels)
            if not tgt: return None
            xml  = z.read("xl/"+tgt.group(1)).decode("utf-8", errors="replace")
            return max(0, len(re.findall(r"<row ", xml))-1)
    except Exception: return None

def export_one(ndax, output_dir, skip_existing):
    base = os.path.splitext(os.path.basename(ndax))[0]
    xlsx = os.path.join(output_dir, base+".xlsx") if output_dir \
           else os.path.splitext(ndax)[0]+".xlsx"
    if skip_existing and os.path.isfile(xlsx):
        return {"ndax": ndax, "xlsx": xlsx, "status": "skipped",
                "cycles": None, "size_kb": None, "elapsed": 0,
                "ts": time.strftime("%H:%M:%S")}
    t0   = time.time()
    proc = subprocess.run([BTSDA_EXE, "export", "custom", ndax, xlsx],
                          capture_output=True, creationflags=0x08000000)
    elapsed = round(time.time()-t0, 1)
    if proc.returncode == 0 and os.path.isfile(xlsx):
        return {"ndax": ndax, "xlsx": xlsx, "status": "ok",
                "cycles": get_cycle_count(xlsx),
                "size_kb": round(os.path.getsize(xlsx)/1024, 1),
                "elapsed": elapsed, "ts": time.strftime("%H:%M:%S")}
    return {"ndax": ndax, "xlsx": xlsx, "status": "fail",
            "cycles": None, "size_kb": None,
            "elapsed": elapsed, "ts": time.strftime("%H:%M:%S")}

def write_csv(results, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Filename","Input","Output","Cycles","Size_KB","Elapsed_s","Time","Status"])
        for r in results:
            w.writerow([os.path.basename(r["ndax"]), r["ndax"], r["xlsx"],
                        r.get("cycles",""), r.get("size_kb",""),
                        r.get("elapsed",""), r.get("ts",""), r["status"]])

def send_email(results, mode_name, csv_path):
    ok   = sum(1 for r in results if r["status"]=="ok")
    fail = sum(1 for r in results if r["status"]=="fail")
    skip = sum(1 for r in results if r["status"]=="skipped")
    rows = "".join(
        f'<tr style="background:{"#d4edda" if r["status"]=="ok" else "#f8d7da" if r["status"]=="fail" else "#fff3cd"}">'
        f'<td>{os.path.basename(r["ndax"])}</td><td>{r["status"]}</td>'
        f'<td>{r.get("cycles","")}</td><td>{r.get("size_kb","")}</td>'
        f'<td>{r.get("elapsed","")}</td></tr>' for r in results)
    html = (f"<h3>Neware Batch Export Complete</h3>"
            f"<p>Mode: <b>{mode_name}</b> | "
            f"OK <b style='color:green'>{ok}</b>  "
            f"Fail <b style='color:red'>{fail}</b>  Skip {skip}</p>"
            f"<table border='1' cellspacing='0' cellpadding='4' "
            f"style='border-collapse:collapse;font-size:12px'>"
            f"<tr style='background:#eee'><th>File</th><th>Status</th>"
            f"<th>Cycles</th><th>KB</th><th>s</th></tr>{rows}</table>")
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Neware Export — {ok}✓/{fail}✗ [{time.strftime('%m-%d %H:%M')}]"
    msg["From"] = SENDER_EMAIL; msg["To"] = RECEIVER_EMAIL
    msg.attach(MIMEText(html, "html", "utf-8"))
    if csv_path and os.path.isfile(csv_path):
        with open(csv_path, "rb") as f:
            att = MIMEApplication(f.read(), Name=os.path.basename(csv_path))
        att["Content-Disposition"] = f'attachment; filename="{os.path.basename(csv_path)}"'
        msg.attach(att)
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=ctx) as s:
        s.login(SENDER_EMAIL, SENDER_PASSWORD)
        s.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())

# ── Folder Scan Dialog ───────────────────────────────────────────────
class FolderScanDialog(ctk.CTkToplevel):
    def __init__(self, parent, on_add):
        super().__init__(parent)
        self.title(T("scan_folder"))
        self.geometry("700x480")
        self.resizable(True, True)
        self.grab_set()
        self.configure(fg_color=C_BG)
        self.on_add  = on_add
        self._checks = []
        self._build()

    def _build(self):
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=16, pady=(14,6))
        self._btn_browse = ctk.CTkButton(
            top, text=T("scan_folder"), width=110, height=30,
            font=serif(12), fg_color=C_ACCENT, hover_color="#3A6BAF",
            text_color="white", corner_radius=4, command=self._browse)
        self._btn_browse.pack(side="left")
        ctk.CTkButton(top, text="✓ All", width=56, height=30,
                      font=serif(11), corner_radius=4,
                      fg_color=C_ACCENT2, hover_color=C_BORDER,
                      text_color=C_ACCENT, command=self._all).pack(side="left", padx=6)
        ctk.CTkButton(top, text="✗ None", width=64, height=30,
                      font=serif(11), corner_radius=4,
                      fg_color=C_ACCENT2, hover_color=C_BORDER,
                      text_color=C_MUTED, command=self._none).pack(side="left")
        self.lbl_dir = ctk.CTkLabel(top, text="", text_color=C_MUTED,
                                    font=mono(10), anchor="w")
        self.lbl_dir.pack(side="left", padx=10, fill="x", expand=True)

        self.scroll = ctk.CTkScrollableFrame(self, fg_color=C_CARD,
                                              border_width=1,
                                              border_color=C_BORDER,
                                              corner_radius=6)
        self.scroll.pack(fill="both", expand=True, padx=16, pady=4)

        bot = ctk.CTkFrame(self, fg_color="transparent")
        bot.pack(fill="x", padx=16, pady=(6,14))
        self.lbl_sel = ctk.CTkLabel(bot, text=T("sel_count", n=0),
                                    text_color=C_MUTED, font=serif(11))
        self.lbl_sel.pack(side="left")
        ctk.CTkButton(bot, text=T("cancel"), width=72, height=30,
                      font=serif(12), corner_radius=4,
                      fg_color=C_ACCENT2, hover_color=C_BORDER,
                      text_color=C_MUTED, command=self.destroy).pack(side="right", padx=6)
        ctk.CTkButton(bot, text=T("add_selected"), width=110, height=30,
                      font=serif(12), corner_radius=4,
                      fg_color=C_ACCENT, hover_color="#3A6BAF",
                      text_color="white", command=self._confirm).pack(side="right")

    def _browse(self):
        folder = filedialog.askdirectory(parent=self)
        if not folder: return
        self.lbl_dir.configure(text=folder)
        files = sorted(
            os.path.join(root, f)
            for root, _, fnames in os.walk(folder)
            for f in fnames if f.lower().endswith(".ndax"))
        for w in self.scroll.winfo_children(): w.destroy()
        self._checks.clear()
        for p in files:
            var = ctk.BooleanVar(value=True)
            cb  = ctk.CTkCheckBox(self.scroll, text=p, variable=var,
                                  font=mono(10), text_color=C_TEXT,
                                  fg_color=C_ACCENT,
                                  command=self._update_count)
            cb.pack(anchor="w", padx=4, pady=1)
            self._checks.append((var, p))
        self._update_count()

    def _all(self):  [v.set(True)  for v,_ in self._checks]; self._update_count()
    def _none(self): [v.set(False) for v,_ in self._checks]; self._update_count()
    def _update_count(self):
        n = sum(1 for v,_ in self._checks if v.get())
        self.lbl_sel.configure(text=T("sel_count", n=n))
    def _confirm(self):
        sel = [p for v,p in self._checks if v.get()]
        if sel: self.on_add(sel)
        self.destroy()

# ── Main App ─────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("840x760")
        self.minsize(720, 640)
        self.configure(fg_color=C_BG)
        self._results   = []
        self._t_start   = None
        self._exporting = False
        self._lang      = "en"
        self._widgets   = {}   # key → widget (for lang refresh)
        self._build()
        self._load_history()

    # ── Build ────────────────────────────────────────────────────────
    def _build(self):
        global LANG
        LANG = self._lang
        self.title(T("title"))
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(12, weight=1)

        # Header
        hdr = ctk.CTkFrame(self, fg_color=C_CARD, corner_radius=0, height=58)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_propagate(False)
        hdr.grid_columnconfigure(1, weight=1)

        self._reg("lbl_title", ctk.CTkLabel(
            hdr, text=T("title"),
            font=serif(18, "bold"), text_color=C_TEXT, anchor="w"))
        self._reg("lbl_title").grid(row=0, column=0, padx=22, pady=14, sticky="w")

        # Language toggle
        self.lang_btn = ctk.CTkButton(
            hdr, text="中文", width=52, height=28,
            font=serif(11), corner_radius=4,
            fg_color=C_ACCENT2, hover_color=C_BORDER,
            text_color=C_ACCENT, command=self._toggle_lang)
        self.lang_btn.grid(row=0, column=1, padx=(0,10), sticky="e")

        self._reg("lbl_iconf", ctk.CTkLabel(
            hdr, text="", font=serif(10), text_color=C_MUTED))
        self._reg("lbl_iconf").grid(row=0, column=2, padx=(0,22), sticky="e")
        self._refresh_iconf_label()

        # Divider
        ctk.CTkFrame(self, height=1, fg_color=C_BORDER,
                     corner_radius=0).grid(row=1, column=0, sticky="ew")

        # Cycle mode
        c1 = self._card(2, "cycle_mode")
        self.seg_mode = ctk.CTkSegmentedButton(
            c1, values=T("seg_labels"),
            font=serif(12),
            fg_color=C_BORDER,
            selected_color=C_ACCENT,
            selected_hover_color="#3A6BAF",
            unselected_color=C_CARD,
            unselected_hover_color=C_ACCENT2,
            text_color=C_TEXT,
            width=520)
        self.seg_mode.pack(padx=16, pady=12, anchor="w")
        self.seg_mode.set(T("seg_labels")[get_iconf_mode()])

        # Input paths
        c2 = self._card(4, "input_paths")
        tb_top = ctk.CTkFrame(c2, fg_color="transparent")
        tb_top.pack(fill="x", padx=14, pady=(8,4))
        self._reg("btn_scan", ctk.CTkButton(
            tb_top, text=T("scan_folder"), width=120, height=28,
            font=serif(11), corner_radius=4,
            fg_color=C_ACCENT2, hover_color=C_BORDER,
            text_color=C_ACCENT, command=self._open_scan))
        self._reg("btn_scan").pack(side="left")
        self._reg("btn_clear_paths", ctk.CTkButton(
            tb_top, text=T("clear"), width=56, height=28,
            font=serif(11), corner_radius=4,
            fg_color=C_ACCENT2, hover_color=C_BORDER,
            text_color=C_MUTED, command=self._clear_paths))
        self._reg("btn_clear_paths").pack(side="left", padx=6)
        self._reg("lbl_count", ctk.CTkLabel(
            tb_top, text=T("files_none"),
            font=serif(11, "bold"), text_color=C_MUTED))
        self._reg("lbl_count").pack(side="right")

        self.txt_paths = ctk.CTkTextbox(
            c2, height=130, font=mono(11),
            fg_color=C_LOG_BG, border_color=C_BORDER, border_width=1,
            corner_radius=4, text_color=C_TEXT, wrap="none")
        self.txt_paths.pack(fill="x", padx=14, pady=(0,4))
        self.txt_paths.bind("<KeyRelease>", self._on_paths_change)

        self._reg("lbl_path_hint", ctk.CTkLabel(
            c2, text=T("path_hint"),
            font=mono(10), text_color=C_MUTED,
            justify="left", anchor="w"))
        self._reg("lbl_path_hint").pack(fill="x", padx=16, pady=(0,10))

        # Output dir
        c3 = self._card(6, "output_dir")
        out_row = ctk.CTkFrame(c3, fg_color="transparent")
        out_row.pack(fill="x", padx=14, pady=10)
        self.out_mode = ctk.StringVar(value="same")
        self._reg("rb_same", ctk.CTkRadioButton(
            out_row, text=T("same_folder"),
            variable=self.out_mode, value="same",
            font=serif(12), fg_color=C_ACCENT, text_color=C_TEXT,
            command=self._toggle_out))
        self._reg("rb_same").pack(side="left")
        self._reg("rb_custom", ctk.CTkRadioButton(
            out_row, text=T("custom_dir"),
            variable=self.out_mode, value="custom",
            font=serif(12), fg_color=C_ACCENT, text_color=C_TEXT,
            command=self._toggle_out))
        self._reg("rb_custom").pack(side="left", padx=20)
        self.ent_outdir = ctk.CTkEntry(
            out_row, width=240, height=28,
            font=mono(11), fg_color=C_LOG_BG,
            border_color=C_BORDER, text_color=C_TEXT)
        self.ent_outdir.pack(side="left")
        self.btn_browse = ctk.CTkButton(
            out_row, text="…", width=30, height=28,
            font=serif(12), corner_radius=4,
            fg_color=C_ACCENT2, hover_color=C_BORDER,
            text_color=C_ACCENT, command=self._browse_out)
        self.btn_browse.pack(side="left", padx=4)
        self._toggle_out()

        # Options
        c4 = self._card(8, "options")
        opt = ctk.CTkFrame(c4, fg_color="transparent")
        opt.pack(fill="x", padx=14, pady=10)
        self.chk_skip = ctk.CTkCheckBox(
            opt, text=T("skip_existing"),
            font=serif(12), fg_color=C_ACCENT,
            text_color=C_TEXT, hover_color=C_ACCENT2)
        self.chk_skip.select()
        self.chk_skip.pack(side="left")
        self._reg("lbl_workers", ctk.CTkLabel(
            opt, text=T("workers"), font=serif(12), text_color=C_TEXT))
        self._reg("lbl_workers").pack(side="left", padx=(22,4))
        self.opt_workers = ctk.CTkOptionMenu(
            opt, values=["1","2","3","4","5","6","8"],
            width=68, height=28, font=serif(12),
            fg_color=C_CARD, button_color=C_ACCENT,
            button_hover_color="#3A6BAF", text_color=C_TEXT,
            dropdown_fg_color=C_CARD, dropdown_text_color=C_TEXT)
        self.opt_workers.set("3")
        self.opt_workers.pack(side="left")
        self.chk_email = ctk.CTkCheckBox(
            opt, text=T("email_report"),
            font=serif(12), fg_color=C_ACCENT,
            text_color=C_TEXT, hover_color=C_ACCENT2)
        self.chk_email.pack(side="left", padx=22)

        # Action row
        act = ctk.CTkFrame(self, fg_color="transparent")
        act.grid(row=10, column=0, sticky="ew", padx=16, pady=(6,4))
        act.grid_columnconfigure(1, weight=1)
        self._reg("btn_export", ctk.CTkButton(
            act, text=T("export_btn"), width=120, height=38,
            font=serif(14, "bold"), corner_radius=6,
            fg_color=C_ACCENT, hover_color="#3A6BAF",
            text_color="white", command=self._start_export))
        self._reg("btn_export").grid(row=0, column=0, padx=(0,12))
        self.progress = ctk.CTkProgressBar(
            act, height=8, mode="determinate",
            fg_color=C_BORDER, progress_color=C_ACCENT, corner_radius=4)
        self.progress.set(0)
        self.progress.grid(row=0, column=1, sticky="ew", padx=(0,12))
        self._reg("lbl_status", ctk.CTkLabel(
            act, text="", width=130, font=serif(11), text_color=C_MUTED))
        self._reg("lbl_status").grid(row=0, column=2)

        # Log
        log_bar = ctk.CTkFrame(self, fg_color="transparent")
        log_bar.grid(row=11, column=0, sticky="ew", padx=20, pady=(4,0))
        self._reg("lbl_log_title", ctk.CTkLabel(
            log_bar, text=T("log_title"),
            font=serif(11, "bold"), text_color=C_MUTED))
        self._reg("lbl_log_title").pack(side="left")
        self._reg("btn_clear_log", ctk.CTkButton(
            log_bar, text=T("clear_log"), width=52, height=22,
            font=serif(10), corner_radius=4,
            fg_color=C_ACCENT2, hover_color=C_BORDER,
            text_color=C_MUTED, command=self._clear_log))
        self._reg("btn_clear_log").pack(side="right")

        self.txt_log = ctk.CTkTextbox(
            self, height=160, font=mono(11),
            fg_color=C_LOG_BG, border_color=C_BORDER, border_width=1,
            corner_radius=6, state="disabled", text_color=C_TEXT, wrap="none")
        self.txt_log.grid(row=12, column=0, sticky="nsew", padx=16, pady=(4,16))

    def _card(self, row, label_key):
        lbl = ctk.CTkLabel(self, text=T(label_key),
                           font=serif(10, "bold"),
                           text_color=C_MUTED, anchor="w")
        lbl.grid(row=row-1, column=0, sticky="w", padx=22, pady=(8,1))
        self._reg(f"sec_{label_key}", lbl)
        card = ctk.CTkFrame(self, fg_color=C_CARD,
                            border_color=C_BORDER, border_width=1,
                            corner_radius=8)
        card.grid(row=row, column=0, sticky="ew", padx=16, pady=(0,2))
        return card

    def _reg(self, key, widget=None):
        if widget is not None:
            self._widgets[key] = widget
        return self._widgets.get(key)

    # ── Language toggle ──────────────────────────────────────────────
    def _toggle_lang(self):
        global LANG
        self._lang = "zh" if self._lang == "en" else "en"
        LANG = self._lang
        self.lang_btn.configure(text="English" if LANG == "zh" else "中文")
        self._refresh_all_strings()

    def _refresh_all_strings(self):
        self.title(T("title"))
        for key, w in self._widgets.items():
            if key == "lbl_title":       w.configure(text=T("title"))
            elif key == "lbl_log_title": w.configure(text=T("log_title"))
            elif key == "btn_scan":      w.configure(text=T("scan_folder"))
            elif key == "btn_clear_paths": w.configure(text=T("clear"))
            elif key == "btn_clear_log": w.configure(text=T("clear_log"))
            elif key == "rb_same":       w.configure(text=T("same_folder"))
            elif key == "rb_custom":     w.configure(text=T("custom_dir"))
            elif key == "lbl_workers":   w.configure(text=T("workers"))
            elif key.startswith("sec_"): w.configure(text=T(key[4:]))
            elif key == "lbl_path_hint": w.configure(text=T("path_hint"))
            elif key == "btn_export" and not self._exporting:
                w.configure(text=T("export_btn"))
        self.chk_skip.configure(text=T("skip_existing"))
        self.chk_email.configure(text=T("email_report"))
        # Update seg mode
        cur_idx = CYCLE_MODE_VALUES.index(
            STRINGS["en"]["seg_labels"].index(self.seg_mode.get())
            if self._lang == "zh"
            else STRINGS["zh"]["seg_labels"].index(self.seg_mode.get())
        ) if False else None
        # Simpler: find current value index by position
        old_labels = STRINGS["zh" if self._lang == "en" else "en"]["seg_labels"]
        new_labels = T("seg_labels")
        try:
            idx = old_labels.index(self.seg_mode.get())
        except ValueError:
            idx = 0
        self.seg_mode.configure(values=new_labels)
        self.seg_mode.set(new_labels[idx])
        self._refresh_iconf_label()
        self._on_paths_change()

    # ── UI helpers ───────────────────────────────────────────────────
    def _toggle_out(self):
        custom = self.out_mode.get() == "custom"
        self.ent_outdir.configure(state="normal" if custom else "disabled")
        self.btn_browse.configure(state="normal" if custom else "disabled")

    def _browse_out(self):
        d = filedialog.askdirectory(parent=self)
        if d:
            self.ent_outdir.delete(0, "end")
            self.ent_outdir.insert(0, d)

    def _clear_paths(self):
        self.txt_paths.delete("1.0", "end")
        self._on_paths_change()

    def _on_paths_change(self, _=None):
        n = len(parse_paths(self.txt_paths.get("1.0","end")))
        key = "files_none" if n == 0 else ("files_n" if n == 1 else "files_n_pl")
        self._reg("lbl_count").configure(
            text=T(key, n=n),
            text_color=C_ACCENT if n else C_MUTED)

    def _refresh_iconf_label(self):
        cur  = get_iconf_mode()
        names_en = STRINGS["en"]["seg_labels"]
        name = names_en[cur] if cur < len(names_en) else str(cur)
        self._reg("lbl_iconf").configure(
            text=T("iconf_label", name=name, val=cur))

    def _open_scan(self):
        FolderScanDialog(self, self._append_scanned)

    def _append_scanned(self, paths):
        existing = self.txt_paths.get("1.0","end").rstrip()
        sep  = "\n" if existing else ""
        text = "\n".join(f'r"{p}",' for p in paths)
        self.txt_paths.insert("end", sep+text)
        self._on_paths_change()

    # ── Export ───────────────────────────────────────────────────────
    def _start_export(self):
        if self._exporting: return
        paths = parse_paths(self.txt_paths.get("1.0","end"))
        if not paths:
            messagebox.showwarning("", T("no_files_warn")); return

        missing = [p for p in paths if not os.path.isfile(p)]
        if missing:
            if not messagebox.askyesno(
                T("missing_title"),
                T("missing_msg", n=len(missing))): return
            paths = [p for p in paths if os.path.isfile(p)]

        if is_btsda_running():
            if not messagebox.askyesno(T("btsda_title"), T("btsda_msg")): return

        out_dir = ""
        if self.out_mode.get() == "custom":
            out_dir = self.ent_outdir.get().strip()
            if not out_dir:
                messagebox.showwarning("", T("no_outdir")); return
            os.makedirs(out_dir, exist_ok=True)

        self._save_history()
        self._results   = []
        self._exporting = True
        self._t_start   = time.time()
        self.progress.set(0)
        self._reg("btn_export").configure(state="disabled",
                                        text=T("exporting"))

        seg_labels = T("seg_labels")
        mode_label = self.seg_mode.get()
        try:
            mode_val = CYCLE_MODE_VALUES[seg_labels.index(mode_label)]
        except ValueError:
            mode_val = 0
        # Keep English name for iconf logging
        en_labels = STRINGS["en"]["seg_labels"]
        en_name   = en_labels[mode_val] if mode_val < len(en_labels) else str(mode_val)

        threading.Thread(
            target=self._worker,
            args=(paths, mode_val, mode_label, en_name, out_dir,
                  bool(self.chk_skip.get()), int(self.opt_workers.get()),
                  bool(self.chk_email.get())),
            daemon=True).start()

    def _worker(self, paths, mode_val, mode_label, en_name,
                out_dir, skip, n_workers, do_email):
        total = len(paths)
        done  = [0]

        self._log(T("log_start", n=total, mode=mode_label, w=n_workers), C_INFO)

        backup = ICONF_PATH + ".bak"
        try:
            shutil.copy2(ICONF_PATH, backup)
            set_iconf_mode(mode_val)
            self._log(T("log_iconf", val=mode_val,
                        bak=os.path.basename(backup)), C_MUTED)
        except Exception as e:
            self._log(T("log_iconf_err", e=e), C_FAIL)
            self.after(0, self._finish_ui); return

        def _handle(r):
            done[0] += 1
            elapsed = round(time.time()-self._t_start, 1)
            name    = os.path.basename(r["ndax"])
            if r["status"] == "ok":
                if r["cycles"] is not None:
                    self._log(T("log_ok", name=name, kb=r["size_kb"],
                                cyc=r["cycles"], s=r["elapsed"]), C_OK)
                else:
                    self._log(T("log_ok_nokb", name=name,
                                kb=r["size_kb"], s=r["elapsed"]), C_OK)
            elif r["status"] == "skipped":
                self._log(T("log_skip", name=name), C_SKIP)
            else:
                self._log(T("log_fail", name=name), C_FAIL)
            self.after(0, lambda p=done[0]/total, d=done[0], e=elapsed: (
                self.progress.set(p),
                self._reg("lbl_status").configure(
                    text=T("progress", done=d, total=total, t=e))))

        with ThreadPoolExecutor(max_workers=n_workers) as ex:
            futures = {ex.submit(export_one, p, out_dir, skip): p for p in paths}
            for fut in as_completed(futures):
                r = fut.result()
                self._results.append(r)
                _handle(r)

        ok     = sum(1 for r in self._results if r["status"]=="ok")
        fail   = sum(1 for r in self._results if r["status"]=="fail")
        skip_n = sum(1 for r in self._results if r["status"]=="skipped")
        total_t = round(time.time()-self._t_start, 1)
        self._log(T("log_done", ok=ok, fail=fail, skip=skip_n, t=total_t),
                  C_OK if fail==0 else C_FAIL)

        csv_path = os.path.join(
            SCRIPT_DIR, f"export_summary_{time.strftime('%Y%m%d_%H%M%S')}.csv")
        try:
            write_csv(self._results, csv_path)
            self._log(T("log_csv", path=csv_path), C_MUTED)
        except Exception as e:
            self._log(T("log_csv_err", e=e), C_FAIL); csv_path = None

        if do_email:
            self._log(T("log_email"), C_MUTED)
            try:
                send_email(self._results, en_name, csv_path)
                self._log(T("log_email_ok", addr=RECEIVER_EMAIL), C_OK)
            except Exception as e:
                self._log(T("log_email_err", e=e), C_FAIL)

        self.after(0, self._refresh_iconf_label)
        self.after(0, self._finish_ui)

    def _finish_ui(self):
        self._exporting = False
        self._reg("btn_export").configure(state="normal",
                                        text=T("export_btn"))

    # ── Log ──────────────────────────────────────────────────────────
    def _log(self, msg, color=None):
        def _do():
            self.txt_log.configure(state="normal")
            if color:
                tag = f"t{color.replace('#','')}"
                self.txt_log.tag_config(tag, foreground=color)
                self.txt_log.insert("end", msg+"\n", tag)
            else:
                self.txt_log.insert("end", msg+"\n")
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
        self.after(0, _do)

    def _clear_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0","end")
        self.txt_log.configure(state="disabled")

    # ── History ──────────────────────────────────────────────────────
    def _save_history(self):
        try:
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "paths":    self.txt_paths.get("1.0","end").rstrip(),
                    "mode_idx": CYCLE_MODE_VALUES[
                        T("seg_labels").index(self.seg_mode.get())
                        if self.seg_mode.get() in T("seg_labels") else 0],
                    "out_mode": self.out_mode.get(),
                    "out_dir":  self.ent_outdir.get(),
                    "skip":     bool(self.chk_skip.get()),
                    "workers":  self.opt_workers.get(),
                    "email":    bool(self.chk_email.get()),
                    "lang":     self._lang,
                }, f, ensure_ascii=False, indent=2)
        except Exception: pass

    def _load_history(self):
        if not os.path.isfile(HISTORY_FILE): return
        try:
            with open(HISTORY_FILE, encoding="utf-8") as f:
                h = json.load(f)
            if h.get("lang") and h["lang"] in ("en","zh"):
                global LANG
                self._lang = h["lang"]
                LANG = self._lang
                self.lang_btn.configure(
                    text="English" if LANG=="zh" else "中文")
                self._refresh_all_strings()
            if h.get("paths"):
                self.txt_paths.insert("1.0", h["paths"])
                self._on_paths_change()
            if "mode_idx" in h:
                labels = T("seg_labels")
                idx = h["mode_idx"]
                if 0 <= idx < len(labels):
                    self.seg_mode.set(labels[idx])
            if h.get("out_mode"):
                self.out_mode.set(h["out_mode"]); self._toggle_out()
            if h.get("out_dir"):
                self.ent_outdir.delete(0,"end")
                self.ent_outdir.insert(0, h["out_dir"])
            if not h.get("skip", True): self.chk_skip.deselect()
            if h.get("workers") in ["1","2","3","4","5","6","8"]:
                self.opt_workers.set(h["workers"])
            if h.get("email"): self.chk_email.select()
        except Exception: pass


if __name__ == "__main__":
    App().mainloop()
