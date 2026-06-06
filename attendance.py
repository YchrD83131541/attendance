import datetime
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import hashlib
import hmac

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import labor

# make both names available: `labor` module and `attendance` alias used by
# existing GUI code. Avoid `import attendance` to prevent importing this file.
attendance = labor

# backward-compatibility: expose backend functions at module level so
# other modules importing `attendance` can call `attendance.read_rows()` etc.
try:
    read_rows = attendance.read_rows
    calculate_actual_hours = attendance.calculate_actual_hours
    find_row = attendance.find_row
    checkin = attendance.checkin
    checkout = attendance.checkout
    add_lessons = attendance.add_lessons
    update_record = attendance.update_record
    delete_record = attendance.delete_record
except Exception:

    pass


# ── カラーパレット ────────────────────────────────
C_BG        = "#F0F4FF"   # 全体背景
C_HEADER    = "#4A7FD4"   # ヘッダー帯
C_HEADER_FG = "#FFFFFF"
C_BTN_CHECK = "#5CB85C"   # 出勤ボタン
C_BTN_OUT   = "#E87C2B"   # 退勤ボタン
C_BTN_LESS  = "#5BC0DE"   # レッスン追加
C_BTN_SAVE  = "#9B59B6"   # 訂正保存
C_BTN_DEL   = "#E74C3C"   # 削除
C_BTN_REFR  = "#7F8C8D"   # 更新
C_BTN_REP   = "#2E86AB"   # 集計
C_BTN_XLS   = "#27AE60"   # Excel出力
C_ROW_ODD   = "#FFFFFF"
C_ROW_EVEN  = "#EAF0FB"
C_SELECT    = "#AED6F1"
C_STATUS_OK = "#1A7A3C"
C_STATUS_NG = "#C0392B"
FONT_MAIN   = ("Meiryo UI", 10)
FONT_BOLD   = ("Meiryo UI", 10, "bold")
FONT_TITLE  = ("Meiryo UI", 13, "bold")

STAFF_FILE      = os.path.join(os.path.expanduser("~"), "Documents", "AttendanceApp", "staff_list.json")
USER_FILE       = os.path.join(os.path.expanduser("~"), "Documents", "AttendanceApp", "users.json")
PAID_LEAVE_FILE = os.path.join(os.path.expanduser("~"), "Documents", "AttendanceApp", "paid_leave.json")

# ── ユーザー認証関連 ────────────────────────────
def _hash_password(password):
    """パスワードをハッシュ化"""
    return hashlib.sha256(password.encode()).hexdigest()

def _load_users():
    """ユーザー情報を読み込む"""
    if os.path.exists(USER_FILE):
        try:
            with open(USER_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def _save_users(users):
    """ユーザー情報を保存"""
    os.makedirs(os.path.dirname(USER_FILE), exist_ok=True)
    with open(USER_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def _verify_user(username, password):
    """ユーザー認証を確認"""
    users = _load_users()
    if username not in users:
        return False
    stored_hash = users[username].get("password_hash", "")
    return hmac.compare_digest(stored_hash, _hash_password(password))

def _create_default_user():
    """デフォルトユーザーを作成"""
    users = _load_users()
    if "admin" not in users:
        users["admin"] = {
            "password_hash": _hash_password("admin"),
            "role": "管理者"
        }
        _save_users(users)
def _load_staff_list():
    if os.path.exists(STAFF_FILE):
        try:
            with open(STAFF_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def _save_staff_list(names):
    os.makedirs(os.path.dirname(STAFF_FILE), exist_ok=True)
    with open(STAFF_FILE, "w", encoding="utf-8") as f:
        json.dump(names, f, ensure_ascii=False, indent=2)


def _load_paid_leave():
    if os.path.exists(PAID_LEAVE_FILE):
        try:
            with open(PAID_LEAVE_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"grants": [], "usages": []}
    return {"grants": [], "usages": []}


def _save_paid_leave(data):
    os.makedirs(os.path.dirname(PAID_LEAVE_FILE), exist_ok=True)
    with open(PAID_LEAVE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _next_pl_id(records):
    if not records:
        return 1
    return max(int(r.get("id", 0)) for r in records) + 1


def _calc_paid_leave_balance(name, data, as_of=None):
    """(残日数, 累計付与, 累計取得) を返す。期限切れ付与は除外。"""
    if as_of is None:
        as_of = datetime.date.today().isoformat()
    total_granted = sum(
        float(g.get("days", 0)) for g in data["grants"]
        if g["name"] == name and g.get("expiry_date", "9999-12-31") >= as_of
    )
    total_used = sum(float(u.get("days", 0)) for u in data["usages"] if u["name"] == name)
    remaining = round(total_granted - total_used, 1)
    return remaining, round(total_granted, 1), round(total_used, 1)


def _color_btn(parent, text, command, bg, fg="#FFFFFF", width=10):
    """角丸風の色付きボタン（tk.Button で実装）"""
    return tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=fg, activebackground=bg, activeforeground=fg,
        relief=tk.FLAT, bd=0, padx=10, pady=6,
        font=FONT_BOLD, cursor="hand2", width=width,
    )

# ── ログイン画面 ──────────────────────────────
class LoginWindow(tk.Tk):
    """ログイン画面"""
    def __init__(self, callback=None):
        super().__init__()
        self.title("勤怠管理システム - ログイン")
        self.geometry("350x250")
        self.resizable(False, False)
        self.configure(bg=C_BG)
        self.callback = callback
        self.logged_in_user = None
        
        self._create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        # タイトル
        title_frame = tk.Frame(self, bg=C_HEADER, height=50)
        title_frame.pack(fill=tk.X)
        tk.Label(
            title_frame, text="🔐 ログイン",
            bg=C_HEADER, fg=C_HEADER_FG, font=FONT_TITLE
        ).pack(pady=12)
        
        # メインフレーム
        main_frame = tk.Frame(self, bg=C_BG)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # ユーザー名
        tk.Label(main_frame, text="ユーザー名:", bg=C_BG, font=FONT_MAIN).pack(anchor=tk.W, pady=(0, 4))
        self.username_var = tk.StringVar()
        username_entry = ttk.Entry(main_frame, textvariable=self.username_var, width=25)
        username_entry.pack(fill=tk.X, pady=(0, 12))
        username_entry.focus()
        
        # パスワード
        tk.Label(main_frame, text="パスワード:", bg=C_BG, font=FONT_MAIN).pack(anchor=tk.W, pady=(0, 4))
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(main_frame, textvariable=self.password_var, width=25, show="*")
        password_entry.pack(fill=tk.X, pady=(0, 20))
        password_entry.bind("<Return>", lambda e: self.on_login())
        
        # ボタンフレーム
        btn_frame = tk.Frame(main_frame, bg=C_BG)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        _color_btn(btn_frame, "ログイン", self.on_login, C_BTN_CHECK, width=12).pack(side=tk.LEFT, padx=(0, 5))
        _color_btn(btn_frame, "キャンセル", self.on_close, C_BTN_DEL, width=12).pack(side=tk.LEFT)
        
        # メッセージラベル
        self.message_var = tk.StringVar()
        tk.Label(main_frame, textvariable=self.message_var, bg=C_BG, fg=C_STATUS_NG, font=("Meiryo UI", 9)).pack(pady=(15, 0))
    
    def on_login(self):
        """ログイン処理"""
        username = self.username_var.get().strip()
        password = self.password_var.get()
        
        if not username or not password:
            self.message_var.set("ユーザー名とパスワードを入力してください")
            return
        
        if _verify_user(username, password):
            self.logged_in_user = username
            self.destroy()
        else:
            self.message_var.set("ユーザー名またはパスワードが正しくありません")
            self.password_var.set("")
    
    def on_close(self):
        """ウィンドウを閉じる"""
        self.destroy()

#勤怠管理アプリの画面を作るクラス
class AttendanceGUI(tk.Tk):
    def __init__(self, username=None):
        super().__init__() #tkinterのウインドウ機能を有効化
        self.logged_in_user = username
        self.title("スポーツクラブ 勤怠管理")
        self.geometry("1060x580")
        self.configure(bg=C_BG)
        self.sort_by_name = False
        self.staff_list = _load_staff_list()
        self._edit_original_date = None
        self._apply_style()
        self.create_widgets()
        self.refresh_records()
        # 起動時に36協定の警告をチェック
        try:
            self.check_overtime_alerts()
        except Exception:
            pass

    # ── スタイル設定 ──────────────────────────────
    def _apply_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".", background=C_BG, font=FONT_MAIN)
        style.configure("TFrame", background=C_BG)
        style.configure("TLabel", background=C_BG, font=FONT_MAIN)
        style.configure("TEntry", font=FONT_MAIN, fieldbackground="#FDFEFF")
        style.configure("TCombobox", font=FONT_MAIN)
        style.configure("TRadiobutton", background=C_BG, font=FONT_MAIN)

        # Treeview
        style.configure(
            "Treeview",
            background=C_ROW_ODD,
            fieldbackground=C_ROW_ODD,
            rowheight=26,
            font=FONT_MAIN,
        )
        style.configure(
            "Treeview.Heading",
            background=C_HEADER,
            foreground=C_HEADER_FG,
            font=FONT_BOLD,
            relief=tk.FLAT,
        )
        style.map("Treeview", background=[("selected", C_SELECT)])
        style.map("Treeview.Heading", background=[("active", "#3A6FC4")])

    # ── ウィジェット構築 ──────────────────────────
    def create_widgets(self):
        # ─ タイトルヘッダー帯
        hdr = tk.Frame(self, bg=C_HEADER, height=44)
        hdr.pack(fill=tk.X)
        tk.Label(
            hdr, text="⚽ スポーツクラブ 勤怠管理システム",
            bg=C_HEADER, fg=C_HEADER_FG, font=FONT_TITLE,
        ).pack(side=tk.LEFT, padx=16, pady=8)
        
        # ログイン情報とログアウトボタン
        if self.logged_in_user:
            user_frame = tk.Frame(hdr, bg=C_HEADER)
            user_frame.pack(side=tk.RIGHT, padx=16, pady=8)
            tk.Label(
                user_frame, text=f"👤 {self.logged_in_user}",
                bg=C_HEADER, fg=C_HEADER_FG, font=FONT_MAIN
            ).pack(side=tk.LEFT, padx=(0, 10))
            _color_btn(
                user_frame, "🚪 ログアウト", self.on_logout, C_BTN_OUT, width=8
            ).pack(side=tk.LEFT)

        # ─ メインフレーム
        frame = tk.Frame(self, bg=C_BG, padx=14, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # ─ ソートボタン行
        sort_row = tk.Frame(frame, bg=C_BG)
        sort_row.pack(fill=tk.X, pady=(0, 4))
        self.sort_button = _color_btn(
            sort_row, "名前順ソート", self.on_toggle_sort, C_BTN_REFR, width=12
        )
        self.sort_button.pack(side=tk.LEFT)
        _color_btn(
            sort_row, "スタッフ管理", self.open_staff_manager, C_BTN_SAVE, width=12
        ).pack(side=tk.LEFT, padx=(6, 0))

        # ─ 入力行
        top = tk.Frame(frame, bg=C_BG)
        top.pack(fill=tk.X, pady=4)

        def lbl(parent, text, row, col, padx=0):
            tk.Label(parent, text=text, bg=C_BG, font=FONT_MAIN).grid(
                row=row, column=col, sticky=tk.W, padx=(padx, 2)
            )

        lbl(top, "スタッフ名:", 0, 0)
        self.name_var = tk.StringVar()
        self.name_combo = ttk.Combobox(top, textvariable=self.name_var, width=20, state="readonly")
        self.name_combo.grid(row=0, column=1, sticky=tk.W)
        self._update_staff_combobox()

        # ロール（メイン/サブ）
        role_frame = tk.Frame(top, bg=C_BG)
        role_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(2, 0))
        self.role_radio_var = tk.StringVar(value="メイン")
        ttk.Radiobutton(role_frame, text="メイン", variable=self.role_radio_var, value="メイン",
                        command=lambda: self.role_var.set("メイン")).pack(side=tk.LEFT)
        ttk.Radiobutton(role_frame, text="サブ",   variable=self.role_radio_var, value="サブ",
                        command=lambda: self.role_var.set("サブ")).pack(side=tk.LEFT)

        # 勤務区分
        lbl(top, "勤務区分:", 1, 2, padx=12)
        self.work_type_var = tk.StringVar(value="通常出勤")
        ttk.Combobox(
            top, textvariable=self.work_type_var,
            values=("通常出勤", "所定休日出勤", "法定休日出勤"),
            width=13, state="readonly",
        ).grid(row=1, column=3, sticky=tk.W)

        lbl(top, "日付:", 0, 2, padx=12)
        self.date_var = tk.StringVar(value=datetime.date.today().isoformat())
        ttk.Entry(top, textvariable=self.date_var, width=14).grid(row=0, column=3, sticky=tk.W)

        lbl(top, "レッスン数:", 0, 4, padx=12)
        self.lessons_var = tk.StringVar(value="1")
        ttk.Entry(top, textvariable=self.lessons_var, width=5).grid(row=0, column=5, sticky=tk.W)

        lbl(top, "休憩(分):", 0, 6, padx=12)
        self.break_var = tk.StringVar(value="0")
        ttk.Entry(top, textvariable=self.break_var, width=5).grid(row=0, column=7, sticky=tk.W)

        lbl(top, "事務作業:", 0, 8, padx=12)
        self.admin_hours_var   = tk.StringVar(value="0")
        self.admin_minutes_var = tk.StringVar(value="0")
        ttk.Entry(top, textvariable=self.admin_hours_var,   width=4).grid(row=0, column=9,  sticky=tk.W)
        tk.Label(top, text="h", bg=C_BG).grid(row=0, column=10, sticky=tk.W)
        ttk.Entry(top, textvariable=self.admin_minutes_var, width=4).grid(row=0, column=11, sticky=tk.W, padx=(4, 0))
        tk.Label(top, text="m", bg=C_BG).grid(row=0, column=12, sticky=tk.W)

        lbl(top, "担当:", 0, 13, padx=12)
        self.role_var = tk.StringVar(value="メイン")
        ttk.Combobox(top, textvariable=self.role_var, values=("メイン", "サブ"),
                     width=7, state="readonly").grid(row=0, column=14, sticky=tk.W)

        # ─ ボタン行
        btn_row = tk.Frame(frame, bg=C_BG)
        btn_row.pack(fill=tk.X, pady=(10, 4))

        buttons = [
            ("✅ 登録",       self.on_register,         C_BTN_CHECK),
            ("📚 レッスン追加", self.on_add_lessons,    C_BTN_LESS),
            ("💾 訂正保存",   self.on_save_corrections,  C_BTN_SAVE),
            ("🗑 削除",       self.on_delete_selected,   C_BTN_DEL),
            ("🔄 更新",       self.refresh_records,      C_BTN_REFR),
            ("📊 集計",       self.show_report,          C_BTN_REP),
            ("📝 36協定管理",  self.open_labor_manager,   "#D35400"),
            ("👤 ユーザー管理", self.open_user_manager,   "#8E44AD"),
            ("📥 Excel出力",  self.export_excel,          C_BTN_XLS),
            ("🌴 有給管理",  self.open_paid_leave_manager, "#16A085"),
        ]
        for text, cmd, color in buttons:
            _color_btn(btn_row, text, cmd, color, width=11).pack(side=tk.LEFT, padx=3)

        # ─ ステータスバー
        self.status_var = tk.StringVar(value="準備完了")
        self.status_label = tk.Label(
            frame, textvariable=self.status_var,
            bg=C_BG, fg=C_STATUS_OK, font=("Meiryo UI", 9, "italic"), anchor=tk.W
        )
        self.status_label.pack(fill=tk.X, pady=(2, 6))

        # ─ 時刻表示行
        time_frame = tk.Frame(frame, bg=C_BG)
        time_frame.pack(fill=tk.X, pady=(0, 6))

        def tlbl(text, col, padx=0):
            tk.Label(time_frame, text=text, bg=C_BG, font=FONT_MAIN).grid(
                row=0, column=col, sticky=tk.W, padx=(padx, 2)
            )

        tlbl("出勤時間:", 0)
        self.checkin_time_var = tk.StringVar(value="")
        ttk.Entry(time_frame, textvariable=self.checkin_time_var, width=22).grid(row=0, column=1, sticky=tk.W)

        tlbl("退勤時間:", 2, padx=10)
        self.checkout_time_var = tk.StringVar(value="")
        ttk.Entry(time_frame, textvariable=self.checkout_time_var, width=22).grid(row=0, column=3, sticky=tk.W)

        tlbl("休憩:", 4, padx=8)
        self.break_display_var = tk.StringVar(value="0分")
        tk.Label(time_frame, textvariable=self.break_display_var, bg=C_BG, width=8, anchor=tk.W).grid(row=0, column=5, sticky=tk.W)

        tlbl("総勤務:", 6, padx=8)
        self.total_work_var = tk.StringVar(value="-")
        tk.Label(time_frame, textvariable=self.total_work_var, bg=C_BG, width=10, anchor=tk.W).grid(row=0, column=7, sticky=tk.W)

        tk.Label(time_frame, text="実働時間:", bg=C_BG, font=FONT_MAIN).grid(row=1, column=0, sticky=tk.W)
        self.actual_hours_var = tk.StringVar(value="-")
        tk.Label(time_frame, textvariable=self.actual_hours_var, bg=C_BG, width=22, anchor=tk.W).grid(row=1, column=1, sticky=tk.W)

        # ─ Treeview
        tree_frame = tk.Frame(frame, bg=C_BG)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("date", "name", "work_type", "checkin", "checkout", "break", "total", "lessons", "admin")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15)

        headers = {
            "date":      ("日付",          90,  tk.CENTER),
            "name":      ("スタッフ名",    120, tk.W),
            "work_type": ("勤務区分",      110, tk.CENTER),
            "checkin":   ("出勤",          125, tk.CENTER),
            "checkout":  ("退勤",          125, tk.CENTER),
            "break":     ("休憩",           65, tk.CENTER),
            "total":     ("総勤務",         65, tk.CENTER),
            "lessons":   ("レッスン(M/S)", 120, tk.CENTER),
            "admin":     ("事務作業",      100, tk.CENTER),
        }
        for cid, (htext, w, anchor) in headers.items():
            self.tree.heading(cid, text=htext)
            self.tree.column(cid, width=w, anchor=anchor)

        self.tree.tag_configure("odd",  background=C_ROW_ODD)
        self.tree.tag_configure("even", background=C_ROW_EVEN)

        sb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

    # ── スタッフ管理 ──────────────────────────────
    def _update_staff_combobox(self):
        self.name_combo["values"] = self.staff_list

    def open_staff_manager(self):
        dlg = tk.Toplevel(self)
        dlg.title("スタッフ管理")
        dlg.geometry("280x380")
        dlg.resizable(False, False)
        dlg.configure(bg=C_BG)
        dlg.grab_set()

        tk.Label(dlg, text="スタッフ一覧", bg=C_BG, font=FONT_BOLD).pack(pady=(10, 4))

        list_frame = tk.Frame(dlg, bg=C_BG)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10)

        lb = tk.Listbox(list_frame, font=FONT_MAIN, selectbackground=C_SELECT, height=12, relief=tk.FLAT, bd=1)
        sb_lb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=lb.yview)
        lb.configure(yscrollcommand=sb_lb.set)
        for name in self.staff_list:
            lb.insert(tk.END, name)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_lb.pack(side=tk.RIGHT, fill=tk.Y)

        input_frame = tk.Frame(dlg, bg=C_BG)
        input_frame.pack(fill=tk.X, padx=10, pady=(8, 2))
        tk.Label(input_frame, text="名前:", bg=C_BG, font=FONT_MAIN).pack(side=tk.LEFT)
        new_name_var = tk.StringVar()
        entry = ttk.Entry(input_frame, textvariable=new_name_var, width=16)
        entry.pack(side=tk.LEFT, padx=(4, 0))

        def refresh_lb():
            lb.delete(0, tk.END)
            for n in self.staff_list:
                lb.insert(tk.END, n)

        def add_staff():
            name = new_name_var.get().strip()
            if not name:
                return
            if name in self.staff_list:
                messagebox.showwarning("重複", f"「{name}」は既に登録されています。", parent=dlg)
                return
            self.staff_list.append(name)
            self.staff_list.sort()
            _save_staff_list(self.staff_list)
            self._update_staff_combobox()
            refresh_lb()
            new_name_var.set("")
            entry.focus()

        entry.bind("<Return>", lambda _: add_staff())

        def delete_staff():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("削除", "削除するスタッフを選択してください。", parent=dlg)
                return
            name = lb.get(sel[0])
            if messagebox.askyesno("確認", f"「{name}」を削除しますか？", parent=dlg):
                self.staff_list.remove(name)
                _save_staff_list(self.staff_list)
                self._update_staff_combobox()
                refresh_lb()

        btn_frame = tk.Frame(dlg, bg=C_BG)
        btn_frame.pack(pady=(6, 10))
        _color_btn(btn_frame, "追加", add_staff,    C_BTN_CHECK, width=8).pack(side=tk.LEFT, padx=4)
        _color_btn(btn_frame, "削除", delete_staff, C_BTN_DEL,   width=8).pack(side=tk.LEFT, padx=4)

    # ── ユーザー管理 ──────────────────────────────
    def open_user_manager(self):
        """ユーザー管理画面を開く"""
        dlg = tk.Toplevel(self)
        dlg.title("ユーザー管理")
        dlg.geometry("350x450")
        dlg.resizable(False, False)
        dlg.configure(bg=C_BG)
        dlg.grab_set()

        tk.Label(dlg, text="ユーザー一覧", bg=C_BG, font=FONT_BOLD).pack(pady=(10, 4))

        # ユーザーリスト表示フレーム
        list_frame = tk.Frame(dlg, bg=C_BG)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10)

        lb = tk.Listbox(list_frame, font=FONT_MAIN, selectbackground=C_SELECT, height=12, relief=tk.FLAT, bd=1)
        sb_lb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=lb.yview)
        lb.configure(yscrollcommand=sb_lb.set)

        def refresh_user_lb():
            lb.delete(0, tk.END)
            users = _load_users()
            for username in sorted(users.keys()):
                role = users[username].get("role", "ユーザー")
                lb.insert(tk.END, f"{username} ({role})")

        refresh_user_lb()
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_lb.pack(side=tk.RIGHT, fill=tk.Y)

        # ユーザー名入力フレーム
        input_frame = tk.Frame(dlg, bg=C_BG)
        input_frame.pack(fill=tk.X, padx=10, pady=(8, 4))
        tk.Label(input_frame, text="ユーザー名:", bg=C_BG, font=FONT_MAIN).pack(side=tk.LEFT)
        username_var = tk.StringVar()
        username_entry = ttk.Entry(input_frame, textvariable=username_var, width=18)
        username_entry.pack(side=tk.LEFT, padx=(4, 0))

        # パスワード入力フレーム
        pwd_frame = tk.Frame(dlg, bg=C_BG)
        pwd_frame.pack(fill=tk.X, padx=10, pady=(0, 4))
        tk.Label(pwd_frame, text="パスワード:", bg=C_BG, font=FONT_MAIN).pack(side=tk.LEFT)
        password_var = tk.StringVar()
        password_entry = ttk.Entry(pwd_frame, textvariable=password_var, width=18, show="*")
        password_entry.pack(side=tk.LEFT, padx=(4, 0))

        # ロール選択フレーム
        role_frame = tk.Frame(dlg, bg=C_BG)
        role_frame.pack(fill=tk.X, padx=10, pady=(0, 8))
        tk.Label(role_frame, text="ロール:", bg=C_BG, font=FONT_MAIN).pack(side=tk.LEFT)
        role_var = tk.StringVar(value="ユーザー")
        ttk.Combobox(role_frame, textvariable=role_var, values=("ユーザー", "管理者"),
                     width=15, state="readonly").pack(side=tk.LEFT, padx=(4, 0))

        def add_user():
            """新しいユーザーを追加"""
            username = username_var.get().strip()
            password = password_var.get()
            
            if not username or not password:
                messagebox.showwarning("入力エラー", "ユーザー名とパスワードを入力してください。", parent=dlg)
                return
            
            users = _load_users()
            if username in users:
                messagebox.showwarning("重複", f"「{username}」は既に存在します。", parent=dlg)
                return
            
            users[username] = {
                "password_hash": _hash_password(password),
                "role": role_var.get()
            }
            _save_users(users)
            refresh_user_lb()
            username_var.set("")
            password_var.set("")
            username_entry.focus()
            messagebox.showinfo("成功", f"ユーザー「{username}」を作成しました。", parent=dlg)

        def delete_user():
            """選択されたユーザーを削除"""
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("削除", "削除するユーザーを選択してください。", parent=dlg)
                return
            
            user_text = lb.get(sel[0])
            username = user_text.split(" (")[0]
            
            if username == "admin":
                messagebox.showerror("エラー", "管理者ユーザーは削除できません。", parent=dlg)
                return
            
            if messagebox.askyesno("確認", f"ユーザー「{username}」を削除しますか？", parent=dlg):
                users = _load_users()
                if username in users:
                    del users[username]
                    _save_users(users)
                    refresh_user_lb()
                    messagebox.showinfo("成功", f"ユーザー「{username}」を削除しました。", parent=dlg)

        def change_password():
            """選択されたユーザーのパスワードを変更"""
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("変更", "変更するユーザーを選択してください。", parent=dlg)
                return
            
            new_pwd = password_var.get()
            if not new_pwd:
                messagebox.showwarning("入力エラー", "新しいパスワードを入力してください。", parent=dlg)
                return
            
            user_text = lb.get(sel[0])
            username = user_text.split(" (")[0]
            
            if messagebox.askyesno("確認", f"ユーザー「{username}」のパスワードを変更しますか？", parent=dlg):
                users = _load_users()
                if username in users:
                    users[username]["password_hash"] = _hash_password(new_pwd)
                    _save_users(users)
                    password_var.set("")
                    messagebox.showinfo("成功", f"パスワードを変更しました。", parent=dlg)

        # ボタンフレーム
        btn_frame = tk.Frame(dlg, bg=C_BG)
        btn_frame.pack(pady=(6, 10))
        _color_btn(btn_frame, "作成", add_user, C_BTN_CHECK, width=8).pack(side=tk.LEFT, padx=2)
        _color_btn(btn_frame, "パスワード変更", change_password, C_BTN_SAVE, width=14).pack(side=tk.LEFT, padx=2)
        _color_btn(btn_frame, "削除", delete_user, C_BTN_DEL, width=8).pack(side=tk.LEFT, padx=2)

    # ── ヘルパー ─────────────────────────────────
    def parse_date(self):
        text = self.date_var.get().strip()
        if not text:
            return None
        try:
            return datetime.datetime.fromisoformat(text)
        except ValueError:
            try:
                d = datetime.date.fromisoformat(text)
                return datetime.datetime.combine(d, datetime.datetime.now().time())
            except ValueError:
                raise ValueError("日付フォーマットが不正です（YYYY-MM-DD）")

    def parse_names(self):
        text = self.name_combo.get().strip()
        if not text:
            raise ValueError("スタッフ名を選択してください。")
        return [text]

    def format_timestamp(self, ts):
        if not ts:
            return "-"
        try:
            return datetime.datetime.fromisoformat(ts).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return ts

    def format_lessons(self, row):
        main = int(row.get("lessons_main") or 0)
        sub  = int(row.get("lessons_sub")  or 0)
        if main == 0 and sub == 0:
            return "0"
        return f"メイン{main} / サブ{sub}"

    def format_admin_time(self, row):
        minutes = int(row.get("admin_minutes") or 0)
        if minutes == 0:
            return "0分"
        return f"{minutes // 60}時間{minutes % 60}分"

    def format_break_time(self, row):
        m = int(row.get("break_minutes") or 0)
        return f"{m // 60}:{m % 60:02d}"

    def format_total_work(self, row):
        v = attendance.calculate_actual_hours(row)
        return f"{v:.2f}".rstrip("0").rstrip(".")

    def parse_timestamp(self, text):
        text = text.strip()
        if not text or text == "-":
            return None
        try:
            return datetime.datetime.fromisoformat(text)
        except ValueError:
            try:
                t = datetime.time.fromisoformat(text)
                d = self.parse_date() or datetime.datetime.now()
                return datetime.datetime.combine(d.date(), t)
            except ValueError:
                raise ValueError("日時フォーマットが不正です（YYYY-MM-DD HH:MM）")

    def on_tree_select(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        vals = self.tree.item(selected[0], "values")
        if len(vals) < 5:
            return
        # cols: date, name, work_type, checkin, checkout, ...
        date, name, work_type, checkin, checkout = vals[0], vals[1], vals[2], vals[3], vals[4]

        self.name_combo.set(name)
        self.name_var.set(name)
        self._edit_original_date = date
        self.date_var.set(date)
        self.work_type_var.set(work_type or "通常出勤")

        self.checkin_time_var.set("" if checkin == "-" else checkin)
        self.checkout_time_var.set("" if checkout == "-" else checkout)

        rows = attendance.read_rows()
        row = attendance.find_row(rows, name, date)
        if row:
            bk = int(row.get("break_minutes") or 0)
            self.break_var.set(str(bk))
            self.break_display_var.set(f"{bk // 60}:{bk % 60:02d}")
            am = int(row.get("admin_minutes") or 0)
            self.admin_hours_var.set(str(am // 60))
            self.admin_minutes_var.set(str(am % 60))
            self.actual_hours_var.set(f"{attendance.calculate_actual_hours(row):.2f}h")
            self.total_work_var.set(self.format_total_work(row))

    def update_time_labels(self):
        self.checkin_time_var.set("-")
        self.checkout_time_var.set("-")
        self.actual_hours_var.set("-")
        self.total_work_var.set("-")
        try:
            names = self.parse_names()
            if len(names) != 1:
                for v in (self.checkin_time_var, self.checkout_time_var,
                           self.actual_hours_var, self.total_work_var):
                    v.set("複数選択中")
                return
            when = self.parse_date()
            if not when:
                return
            row = attendance.find_row(attendance.read_rows(), names[0], when.date().isoformat())
            if not row:
                return
            self.checkin_time_var.set(self.format_timestamp(row.get("checkin")))
            self.checkout_time_var.set(self.format_timestamp(row.get("checkout")))
            self.break_var.set(str(int(row.get("break_minutes") or 0)))
            m = int(row.get("break_minutes") or 0)
            self.break_display_var.set(f"{m // 60}:{m % 60:02d}")
            am = int(row.get("admin_minutes") or 0)
            self.admin_hours_var.set(str(am // 60))
            self.admin_minutes_var.set(str(am % 60))
            self.actual_hours_var.set(f"{attendance.calculate_actual_hours(row):.2f}h")
            self.total_work_var.set(self.format_total_work(row))
        except Exception:
            pass

    def set_status(self, message, error=False):
        self.status_var.set(message)
        self.status_label.configure(fg=C_STATUS_NG if error else C_STATUS_OK)

    # ── イベントハンドラ ──────────────────────────
    def on_register(self):
        """全項目を入力してから一括登録する。"""
        try:
            names = self.parse_names()
            if len(names) != 1:
                raise ValueError("スタッフを1名選択してください。")
            name = names[0]
            dt = self.parse_date()
            if not dt:
                raise ValueError("日付を入力してください。")
            date_str = dt.date().isoformat()
            work_type = self.work_type_var.get()

            checkin_dt  = self.parse_timestamp(self.checkin_time_var.get())
            checkout_dt = self.parse_timestamp(self.checkout_time_var.get())

            bk        = int(self.break_var.get() or 0)
            ah        = int(self.admin_hours_var.get() or 0)
            am_val    = int(self.admin_minutes_var.get() or 0)
            admin_min = ah * 60 + am_val
            lessons   = int(self.lessons_var.get() or 0)
            role      = self.role_var.get()

            am_obj  = labor.AttendanceManager()
            rows    = am_obj.read_rows()
            existing = labor.find_row(rows, name, date_str)

            if existing:
                if not messagebox.askyesno("上書き確認",
                        f"「{name}」{date_str} の記録が既に存在します。\n上書きしますか？"):
                    return
                existing["work_type"]     = work_type
                existing["checkin"]       = checkin_dt.isoformat()  if checkin_dt  else existing.get("checkin", "")
                existing["checkout"]      = checkout_dt.isoformat() if checkout_dt else ""
                existing["break_minutes"] = str(bk)
                existing["admin_minutes"] = str(admin_min)
                if role == "メイン":
                    existing["lessons_main"] = str(lessons)
                else:
                    existing["lessons_sub"] = str(lessons)
            else:
                new_row = {
                    "id":            labor._next_id(rows),
                    "name":          name,
                    "date":          date_str,
                    "work_type":     work_type,
                    "checkin":       checkin_dt.isoformat()  if checkin_dt  else "",
                    "checkout":      checkout_dt.isoformat() if checkout_dt else "",
                    "break_minutes": str(bk),
                    "admin_minutes": str(admin_min),
                    "lessons_main":  str(lessons) if role == "メイン" else "0",
                    "lessons_sub":   str(lessons) if role == "サブ"   else "0",
                }
                rows.append(new_row)

            am_obj.save_rows(rows)
            self._edit_original_date = date_str
            self.set_status(f"「{name}」{date_str} を登録しました。")
            self.refresh_records()
            try:
                self.check_overtime_alerts()
            except Exception:
                pass
        except ValueError as exc:
            messagebox.showerror("入力エラー", str(exc))
            self.set_status(str(exc), error=True)
        except Exception as exc:
            messagebox.showerror("エラー", str(exc))
            self.set_status(str(exc), error=True)

    def on_toggle_sort(self):
        self.sort_by_name = not self.sort_by_name
        self.sort_button.configure(text="日付順に戻す" if self.sort_by_name else "名前順ソート")
        self.refresh_records()

    def on_add_lessons(self):
        try:
            names = self.parse_names()
            dt = self.parse_date()
            count = int(self.lessons_var.get())
            role = self.role_var.get()
            msgs = []
            for name in names:
                when = datetime.datetime.combine(dt.date(), datetime.datetime.now().time()) if dt else None
                msgs.append(attendance.add_lessons(name, count, when=when, role=role))
            self.set_status("; ".join(msgs))
            self.refresh_records()
        except ValueError:
            err = "レッスン数は整数を入力してください。"
            messagebox.showerror("入力エラー", err)
            self.set_status(err, error=True)
        except Exception as exc:
            messagebox.showerror("エラー", str(exc))
            self.set_status(str(exc), error=True)

    def refresh_records(self):
        rows = attendance.read_rows()
        if self.sort_by_name:
            rows = sorted(rows, key=lambda r: (r["name"].lower(), r["date"], int(r["id"])))
        else:
            rows = sorted(rows, key=lambda r: (r["date"], int(r["id"])), reverse=True)

        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, r in enumerate(rows[:100]):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", tk.END, tags=(tag,), values=(
                r["date"],
                r["name"],
                r.get("work_type") or "通常出勤",
                self.format_timestamp(r.get("checkin")),
                self.format_timestamp(r.get("checkout")),
                self.format_break_time(r),
                self.format_total_work(r),
                self.format_lessons(r),
                self.format_admin_time(r),
            ))
        self.set_status(f"レコードを {len(rows[:100])} 件表示しました。")

    def on_save_corrections(self):
        try:
            names = self.parse_names()
            if len(names) != 1:
                raise ValueError("修正は1名のみ指定してください。")
            dt = self.parse_date()
            when_date = dt.date().isoformat() if dt else None
            if not when_date:
                raise ValueError("日付を入力してください。")

            lookup_date = self._edit_original_date if self._edit_original_date else when_date
            new_date = when_date if when_date != lookup_date else None

            checkin  = self.parse_timestamp(self.checkin_time_var.get())
            checkout = self.parse_timestamp(self.checkout_time_var.get())

            # 日付が変わった場合、出勤・退勤の日付部分も新しい日付に合わせる
            if new_date:
                if checkin:
                    checkin = datetime.datetime.combine(dt.date(), checkin.time())
                if checkout:
                    checkout = datetime.datetime.combine(dt.date(), checkout.time())

            bk = int(self.break_var.get())
            ah = int(self.admin_hours_var.get())
            am = int(self.admin_minutes_var.get())
            msg = attendance.update_record(
                names[0], lookup_date,
                checkin=checkin, checkout=checkout,
                break_minutes=bk, admin_minutes=ah * 60 + am,
                role=self.role_var.get(),
                work_type=self.work_type_var.get(),
                new_date=new_date,
            )
            self._edit_original_date = when_date
            self.set_status(msg)
            self.refresh_records()
        except (ValueError, Exception) as exc:
            messagebox.showerror("エラー", str(exc))
            self.set_status(str(exc), error=True)

    def on_delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("削除", "削除するレコードを選択してください。")
            return
        if not messagebox.askyesno("確認", "選択したレコードを削除しますか？"):
            return
        msgs = []
        for item in selected:
            vals = self.tree.item(item, "values")
            if len(vals) >= 2:
                msgs.append(attendance.delete_record(vals[1], vals[0]))
        self.set_status("; ".join(msgs))
        self.refresh_records()

    def show_report(self):
        rows = attendance.read_rows()
        by_name = {}
        for r in rows:
            name = r["name"]
            if name not in by_name:
                by_name[name] = {"hours": 0.0, "lessons_main": 0, "lessons_sub": 0}
            by_name[name]["hours"]        += attendance.calculate_actual_hours(r)
            by_name[name]["lessons_main"] += int(r.get("lessons_main") or 0)
            by_name[name]["lessons_sub"]  += int(r.get("lessons_sub")  or 0)

        lines = [
            f"{n}: 勤務 {v['hours']:.2f}h / レッスン メイン{v['lessons_main']} サブ{v['lessons_sub']}"
            for n, v in sorted(by_name.items())
        ]
        messagebox.showinfo("集計結果", "\n".join(lines) if lines else "記録がありません。")

    # ── 36協定管理ウィンドウ ─────────────────────
    def open_labor_manager(self):
        dlg = tk.Toplevel(self)
        dlg.title("36協定管理")
        dlg.geometry("760x420")
        dlg.configure(bg=C_BG)
        dlg.grab_set()

        top = tk.Frame(dlg, bg=C_BG)
        top.pack(fill=tk.X, padx=8, pady=6)

        tk.Label(top, text="対象年月(YYYY-MM):", bg=C_BG, font=FONT_MAIN).pack(side=tk.LEFT)
        ym_var = tk.StringVar(value=datetime.date.today().strftime("%Y-%m"))
        ttk.Entry(top, textvariable=ym_var, width=12).pack(side=tk.LEFT, padx=(6, 8))
        def do_refresh():
            v = ym_var.get().strip()
            try:
                y, m = v.split("-")
                y = int(y); m = int(m)
            except Exception:
                messagebox.showerror("入力エラー", "年月は YYYY-MM の形式で入力してください。", parent=dlg)
                return
            lm = labor.LaborAgreementManager(labor.AttendanceManager())
            data = lm.compute_monthly_summary(y, m)
            for it in tree.get_children():
                tree.delete(it)
            for i, r in enumerate(data):
                tag = "even" if i % 2 == 0 else "odd"
                tree.insert("", tk.END, tags=(tag,), values=(r["name"], r["total_hours"], r["overtime_hours"], r["judgement"], r["warning_level"]))

        ttk.Button(top, text="表示", command=do_refresh).pack(side=tk.LEFT)
        def export_csv():
            v = ym_var.get().strip()
            try:
                y, m = v.split("-")
                y = int(y); m = int(m)
            except Exception:
                messagebox.showerror("入力エラー", "年月は YYYY-MM の形式で入力してください。", parent=dlg)
                return
            path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], parent=dlg)
            if not path:
                return
            lm = labor.LaborAgreementManager(labor.AttendanceManager())
            lm.export_csv(path, y, m)
            messagebox.showinfo("出力完了", f"CSVを保存しました。\n{path}", parent=dlg)
        ttk.Button(top, text="CSV出力", command=export_csv).pack(side=tk.LEFT, padx=6)
        def export_annual_csv():
            y = datetime.date.today().year
            path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], parent=dlg)
            if not path:
                return
            lm = labor.LaborAgreementManager(labor.AttendanceManager())
            lm.export_annual_csv(path, y)
            messagebox.showinfo("出力完了", f"CSVを保存しました。\n{path}", parent=dlg)
        ttk.Button(top, text="年間CSV出力", command=export_annual_csv).pack(side=tk.LEFT, padx=6)
        def export_multi_csv():
            ym = avg_entry.get().strip()
            try:
                y, m = ym.split("-")
                y = int(y); m = int(m)
            except Exception:
                messagebox.showerror("入力エラー", "年月は YYYY-MM 形式で入力してください。", parent=dlg)
                return
            w = int(months_var.get() or 3)
            path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], parent=dlg)
            if not path:
                return
            lm = labor.LaborAgreementManager(labor.AttendanceManager())
            lm.export_multi_month_csv(path, y, m, window=w)
            messagebox.showinfo("出力完了", f"CSVを保存しました。\n{path}", parent=dlg)
        ttk.Button(top, text="複数月CSV出力", command=export_multi_csv).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="年間720チェック", command=lambda: self._do_annual_check(dlg)).pack(side=tk.LEFT, padx=6)
        ttk.Label(top, text=" / 複数月平均(末月 YYYY-MM):", bg=C_BG, font=FONT_MAIN).pack(side=tk.LEFT, padx=(12,4))
        avg_entry = ttk.Entry(top, width=8)
        avg_entry.pack(side=tk.LEFT)
        avg_entry.insert(0, datetime.date.today().strftime("%Y-%m"))
        months_var = tk.StringVar(value="3")
        ttk.Entry(top, textvariable=months_var, width=3).pack(side=tk.LEFT, padx=(6,0))
        ttk.Button(top, text="複数月平均チェック", command=lambda: self._do_multi_month_check(dlg, avg_entry.get(), int(months_var.get() or 3))).pack(side=tk.LEFT, padx=6)

        cols = ("name", "total", "overtime", "judgement", "level")
        tree = ttk.Treeview(dlg, columns=cols, show="headings", height=16)
        headers = {
            "name": ("スタッフ名", 220),
            "total": ("月間勤務時間", 120),
            "overtime": ("時間外労働時間", 120),
            "judgement": ("36協定判定", 120),
            "level": ("警告レベル", 120),
        }
        for cid, (htext, w) in headers.items():
            tree.heading(cid, text=htext)
            tree.column(cid, width=w)
        tree.tag_configure("odd",  background=C_ROW_ODD)
        tree.tag_configure("even", background=C_ROW_EVEN)

        sb = ttk.Scrollbar(dlg, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=sb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8,0), pady=8)
        sb.pack(side=tk.RIGHT, fill=tk.Y, pady=8)

        # keep reference so other methods can update this Treeview
        self.labor_tree = tree

        do_refresh()

    def _do_annual_check(self, parent):
        y = datetime.date.today().year
        lm = labor.LaborAgreementManager(labor.AttendanceManager())
        data = lm.compute_annual_summary(y)
        offenders = [f"{r['name']}: {r['total_hours']}h" for r in data if r['over_720']]
        if offenders:
            messagebox.showwarning("年間720時間超過", "\n".join(offenders), parent=parent)
        else:
            messagebox.showinfo("年間720時間チェック", "超過者はいません。", parent=parent)

    def _do_multi_month_check(self, parent, ym_text, window):
        try:
            y, m = ym_text.split("-")
            y = int(y); m = int(m)
        except Exception:
            messagebox.showerror("入力エラー", "年月は YYYY-MM 形式で入力してください。", parent=parent)
            return
        lm = labor.LaborAgreementManager(labor.AttendanceManager())
        data = lm.compute_multi_month_average(y, m, window=window)
        tree = getattr(self, 'labor_tree', None)
        if tree is None:
            messagebox.showerror("内部エラー", "Treeview が見つかりません。36協定管理画面を再表示してください。", parent=parent)
            return
        # reconfigure tree to show multi-month average
        tree_cols = ("name", "avg", "window", "flag")
        tree.config(columns=tree_cols)
        tree.delete(*tree.get_children())
        tree.heading("name", text="スタッフ名")
        tree.column("name", width=260)
        tree.heading("avg", text=f"平均時間外(h)")
        tree.column("avg", width=140)
        tree.heading("window", text="ウィンドウ(月)")
        tree.column("window", width=100)
        tree.heading("flag", text="超過フラグ")
        tree.column("flag", width=120)
        for i, r in enumerate(data):
            tag = "even" if i % 2 == 0 else "odd"
            flag = "超過" if r.get("over_80_avg") else "正常"
            tree.insert("", tk.END, tags=(tag,), values=(r["name"], r["average_overtime"], r.get("window", window), flag))

    def check_overtime_alerts(self):
        today = datetime.date.today()
        lm = labor.LaborAgreementManager(labor.AttendanceManager())
        data = lm.compute_monthly_summary(today.year, today.month)
        alerts = []
        for r in data:
            if r["overtime_hours"] >= 45:
                alerts.append(f"{r['name']}さんの今月の時間外労働は{r['overtime_hours']}時間です。")
        if alerts:
            messagebox.showwarning("36協定警告", "\n".join(alerts))

    # ── Excel 出力 ────────────────────────────────
    def _ask_export_month(self, all_rows):
        """月選択ダイアログを表示し、選択された YYYY-MM 文字列（または None でキャンセル）を返す。"""
        months = sorted(
            {r["date"][:7] for r in all_rows if r.get("date", "") >= "2000"},
            reverse=True,
        )
        if not months:
            return None

        result = {"value": None}
        dlg = tk.Toplevel(self)
        dlg.title("出力月を選択")
        dlg.geometry("280x140")
        dlg.resizable(False, False)
        dlg.configure(bg=C_BG)
        dlg.grab_set()

        tk.Label(dlg, text="出力対象の年月を選択してください", bg=C_BG, font=FONT_MAIN).pack(pady=(14, 6))

        var = tk.StringVar(value=months[0])
        combo = ttk.Combobox(dlg, textvariable=var, values=months, state="readonly", width=14)
        combo.pack()

        def on_ok():
            result["value"] = var.get()
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        btn_frame = tk.Frame(dlg, bg=C_BG)
        btn_frame.pack(pady=14)
        _color_btn(btn_frame, "OK",       on_ok,     C_BTN_CHECK, width=8).pack(side=tk.LEFT, padx=6)
        _color_btn(btn_frame, "キャンセル", on_cancel, C_BTN_DEL,   width=10).pack(side=tk.LEFT, padx=6)

        dlg.bind("<Return>", lambda _: on_ok())
        dlg.bind("<Escape>", lambda _: on_cancel())
        self.wait_window(dlg)
        return result["value"]

    def export_excel(self):
        all_rows = attendance.read_rows()
        if not all_rows:
            messagebox.showwarning("Excel出力", "出力するデータがありません。")
            return

        month_str = self._ask_export_month(all_rows)
        if month_str is None:
            return

        rows = [r for r in all_rows if r.get("date", "").startswith(month_str)]
        if not rows:
            messagebox.showwarning("Excel出力", f"{month_str} のデータがありません。")
            return

        default_name = f"勤怠集計_{month_str}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Excel ファイルの保存先",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel ファイル", "*.xlsx")],
        )
        if not path:
            return

        wb = openpyxl.Workbook()

        # ── シート1: 詳細レコード ─────────────────
        ws1 = wb.active
        ws1.title = "勤怠詳細"

        hdr_fill  = PatternFill("solid", fgColor="4A7FD4")
        hdr_font  = Font(bold=True, color="FFFFFF", name="Meiryo UI", size=10)
        body_font = Font(name="Meiryo UI", size=10)
        even_fill = PatternFill("solid", fgColor="EAF0FB")
        center    = Alignment(horizontal="center", vertical="center")
        left      = Alignment(horizontal="left",   vertical="center")
        thin      = Side(style="thin", color="CCCCCC")
        border    = Border(left=thin, right=thin, top=thin, bottom=thin)

        detail_headers = ["日付", "スタッフ名", "勤務区分", "出勤時間", "退勤時間",
                          "休憩(分)", "総勤務(h)", "メインレッスン", "サブレッスン", "事務作業(分)"]
        col_widths     = [14, 16, 14, 22, 22, 10, 12, 14, 14, 14]

        for ci, (h, w) in enumerate(zip(detail_headers, col_widths), 1):
            cell = ws1.cell(row=1, column=ci, value=h)
            cell.fill      = hdr_fill
            cell.font      = hdr_font
            cell.alignment = center
            cell.border    = border
            ws1.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

        sorted_rows = sorted(rows, key=lambda r: (r["date"], int(r["id"])), reverse=True)
        for ri, r in enumerate(sorted_rows, 2):
            fill = even_fill if ri % 2 == 0 else PatternFill()
            vals = [
                r.get("date", ""),
                r.get("name", ""),
                r.get("work_type") or "通常出勤",
                self.format_timestamp(r.get("checkin")),
                self.format_timestamp(r.get("checkout")),
                int(r.get("break_minutes") or 0),
                round(attendance.calculate_actual_hours(r), 2),
                int(r.get("lessons_main") or 0),
                int(r.get("lessons_sub")  or 0),
                int(r.get("admin_minutes") or 0),
            ]
            for ci, v in enumerate(vals, 1):
                cell = ws1.cell(row=ri, column=ci, value=v)
                cell.font      = body_font
                cell.fill      = fill
                cell.border    = border
                cell.alignment = center if ci != 2 else left
        ws1.freeze_panes = "A2"

        # ── シート2: スタッフ別集計（勤務区分別） ────────────────
        ws2 = wb.create_sheet("スタッフ別集計")

        fill_normal    = PatternFill("solid", fgColor="4A7FD4")
        fill_specified = PatternFill("solid", fgColor="C87D00")
        fill_legal     = PatternFill("solid", fgColor="A93226")
        total_fill     = PatternFill("solid", fgColor="D5E8D4")
        total_font     = Font(bold=True, name="Meiryo UI", size=10)

        # ─ 2行ヘッダー
        # 列レイアウト:
        # 1=スタッフ名, 2-3=通常出勤, 4-5=所定休日出勤, 6-7=法定休日出勤,
        # 8=総勤務時間, 9=メインL, 10=サブL, 11=総L, 12=事務作業

        def _ws2_hdr(col, text, row1_fill, merge_end_col=None, row2_sub=None):
            if merge_end_col:
                ws2.merge_cells(start_row=1, start_column=col, end_row=1, end_column=merge_end_col)
            else:
                ws2.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
            c = ws2.cell(row=1, column=col, value=text)
            c.fill = row1_fill; c.font = hdr_font; c.alignment = center; c.border = border
            if row2_sub:
                for offset, sub in enumerate(row2_sub):
                    c2 = ws2.cell(row=2, column=col + offset, value=sub)
                    c2.fill = row1_fill; c2.font = hdr_font; c2.alignment = center; c2.border = border

        _ws2_hdr(1,  "スタッフ名",   hdr_fill)
        _ws2_hdr(2,  "通常出勤",     fill_normal,    merge_end_col=3,  row2_sub=["日数", "時間(h)"])
        _ws2_hdr(4,  "所定休日出勤", fill_specified, merge_end_col=5,  row2_sub=["日数", "時間(h)"])
        _ws2_hdr(6,  "法定休日出勤", fill_legal,     merge_end_col=7,  row2_sub=["日数", "時間(h)"])
        _ws2_hdr(8,  "総勤務時間(h)",  hdr_fill)
        _ws2_hdr(9,  "メインレッスン", hdr_fill)
        _ws2_hdr(10, "サブレッスン",   hdr_fill)
        _ws2_hdr(11, "総レッスン",     hdr_fill)
        _ws2_hdr(12, "事務作業(h)",    hdr_fill)

        col_widths_s2 = [18, 8, 10, 10, 12, 10, 12, 14, 14, 12, 10, 12]
        for ci, w in enumerate(col_widths_s2, 1):
            ws2.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

        # ─ データ集計
        by_name = {}
        for r in rows:
            n = r["name"]
            if n not in by_name:
                by_name[n] = {
                    "days_n": 0, "hours_n": 0.0,
                    "days_s": 0, "hours_s": 0.0,
                    "days_l": 0, "hours_l": 0.0,
                    "hours_total": 0.0,
                    "main": 0, "sub": 0, "admin": 0,
                }
            h = attendance.calculate_actual_hours(r)
            by_name[n]["hours_total"] += h
            by_name[n]["main"]  += int(r.get("lessons_main")  or 0)
            by_name[n]["sub"]   += int(r.get("lessons_sub")   or 0)
            by_name[n]["admin"] += int(r.get("admin_minutes") or 0)
            wt = r.get("work_type") or "通常出勤"
            if wt == "所定休日出勤":
                by_name[n]["days_s"] += 1; by_name[n]["hours_s"] += h
            elif wt == "法定休日出勤":
                by_name[n]["days_l"] += 1; by_name[n]["hours_l"] += h
            else:
                by_name[n]["days_n"] += 1; by_name[n]["hours_n"] += h

        # ─ データ行（3行目〜）
        for ri, (name, v) in enumerate(sorted(by_name.items()), 3):
            fill = even_fill if ri % 2 == 0 else PatternFill()
            row_vals = [
                name,
                v["days_n"],  round(v["hours_n"], 2),
                v["days_s"],  round(v["hours_s"], 2),
                v["days_l"],  round(v["hours_l"], 2),
                round(v["hours_total"], 2),
                v["main"], v["sub"], v["main"] + v["sub"],
                round(v["admin"] / 60, 2),
            ]
            for ci, val in enumerate(row_vals, 1):
                c = ws2.cell(row=ri, column=ci, value=val)
                c.font = body_font; c.fill = fill; c.border = border
                c.alignment = left if ci == 1 else center
        ws2.freeze_panes = "A3"

        # ─ 合計行
        total_row = len(by_name) + 3
        total_vals = [
            "【合計】",
            sum(v["days_n"]  for v in by_name.values()),
            round(sum(v["hours_n"]  for v in by_name.values()), 2),
            sum(v["days_s"]  for v in by_name.values()),
            round(sum(v["hours_s"]  for v in by_name.values()), 2),
            sum(v["days_l"]  for v in by_name.values()),
            round(sum(v["hours_l"]  for v in by_name.values()), 2),
            round(sum(v["hours_total"] for v in by_name.values()), 2),
            sum(v["main"] for v in by_name.values()),
            sum(v["sub"]  for v in by_name.values()),
            sum(v["main"] + v["sub"] for v in by_name.values()),
            round(sum(v["admin"] for v in by_name.values()) / 60, 2),
        ]
        for ci, val in enumerate(total_vals, 1):
            c = ws2.cell(row=total_row, column=ci, value=val)
            c.font = total_font; c.fill = total_fill; c.border = border
            c.alignment = left if ci == 1 else center

        try:
            wb.save(path)
            self.set_status(f"Excel出力完了: {os.path.basename(path)}")
            if messagebox.askyesno("完了", f"保存しました。\n{path}\n\nファイルを開きますか？"):
                os.startfile(path)
        except PermissionError:
            messagebox.showerror("書き込みエラー", "ファイルが他のアプリで開かれています。閉じてから再試行してください。")
        except Exception as exc:
            messagebox.showerror("エラー", str(exc))

    # ── 有給休暇管理 ──────────────────────────────
    def open_paid_leave_manager(self):
        dlg = tk.Toplevel(self)
        dlg.title("有給休暇管理")
        dlg.geometry("800x540")
        dlg.configure(bg=C_BG)
        dlg.grab_set()

        notebook = ttk.Notebook(dlg)
        notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        tab_balance = tk.Frame(notebook, bg=C_BG)
        tab_grant   = tk.Frame(notebook, bg=C_BG)
        tab_usage   = tk.Frame(notebook, bg=C_BG)
        tab_history = tk.Frame(notebook, bg=C_BG)
        notebook.add(tab_balance, text="残高一覧")
        notebook.add(tab_grant,   text="付与登録")
        notebook.add(tab_usage,   text="取得登録")
        notebook.add(tab_history, text="履歴一覧")

        C_GRANT = "#16A085"

        # ── タブ1: 残高一覧 ──────────────────────
        ctrl_bal = tk.Frame(tab_balance, bg=C_BG)
        ctrl_bal.pack(fill=tk.X, padx=8, pady=6)
        tk.Label(ctrl_bal, text="基準日:", bg=C_BG, font=FONT_MAIN).pack(side=tk.LEFT)
        bal_date_var = tk.StringVar(value=datetime.date.today().isoformat())
        ttk.Entry(ctrl_bal, textvariable=bal_date_var, width=14).pack(side=tk.LEFT, padx=(4, 8))

        bal_cols = ("name", "granted", "used", "remaining", "alert")
        bal_tree = ttk.Treeview(tab_balance, columns=bal_cols, show="headings", height=14)
        for cid, ht, w, anch in [
            ("name",      "スタッフ名",      160, tk.W),
            ("granted",   "累計付与日数",    110, tk.CENTER),
            ("used",      "取得済日数",      110, tk.CENTER),
            ("remaining", "残日数",           90, tk.CENTER),
            ("alert",     "年5日義務",        140, tk.CENTER),
        ]:
            bal_tree.heading(cid, text=ht)
            bal_tree.column(cid, width=w, anchor=anch)
        bal_tree.tag_configure("odd",  background=C_ROW_ODD)
        bal_tree.tag_configure("even", background=C_ROW_EVEN)
        bal_tree.tag_configure("warn", background="#FFF3CD")

        sb_bal = ttk.Scrollbar(tab_balance, orient=tk.VERTICAL, command=bal_tree.yview)
        bal_tree.configure(yscroll=sb_bal.set)
        bal_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=(0, 8))
        sb_bal.pack(side=tk.RIGHT, fill=tk.Y, pady=(0, 8))

        def refresh_balance():
            as_of = bal_date_var.get().strip() or datetime.date.today().isoformat()
            year  = as_of[:4]
            data  = _load_paid_leave()
            for it in bal_tree.get_children():
                bal_tree.delete(it)
            for i, name in enumerate(sorted(self.staff_list)):
                remaining, granted, used = _calc_paid_leave_balance(name, data, as_of)
                used_yr = sum(
                    float(u.get("days", 0)) for u in data["usages"]
                    if u["name"] == name and u.get("date", "").startswith(year)
                )
                need = round(5 - used_yr, 1)
                alert = "✓ 達成" if used_yr >= 5 else f"要取得(あと{need}日)"
                tag   = "warn" if used_yr < 5 else ("even" if i % 2 == 0 else "odd")
                bal_tree.insert("", tk.END, tags=(tag,), values=(name, granted, used, remaining, alert))

        _color_btn(ctrl_bal, "更新", refresh_balance, C_BTN_REFR, width=8).pack(side=tk.LEFT)
        refresh_balance()

        # ── タブ2: 付与登録 ──────────────────────
        gf = tk.Frame(tab_grant, bg=C_BG)
        gf.pack(fill=tk.X, padx=12, pady=10)

        def glbl(text, row, col):
            tk.Label(gf, text=text, bg=C_BG, font=FONT_MAIN).grid(
                row=row, column=col, sticky=tk.W, padx=(0, 4), pady=4)

        glbl("スタッフ名:", 0, 0)
        g_name_var = tk.StringVar()
        ttk.Combobox(gf, textvariable=g_name_var, values=self.staff_list,
                     width=16, state="readonly").grid(row=0, column=1, sticky=tk.W)
        glbl("付与日:", 0, 2)
        g_date_var = tk.StringVar(value=datetime.date.today().isoformat())
        ttk.Entry(gf, textvariable=g_date_var, width=14).grid(row=0, column=3, sticky=tk.W)
        glbl("付与日数:", 1, 0)
        g_days_var = tk.StringVar(value="10")
        ttk.Entry(gf, textvariable=g_days_var, width=8).grid(row=1, column=1, sticky=tk.W)
        glbl("有効期限:", 1, 2)
        default_expiry = datetime.date.today().replace(
            year=datetime.date.today().year + 2).isoformat()
        g_expiry_var = tk.StringVar(value=default_expiry)
        ttk.Entry(gf, textvariable=g_expiry_var, width=14).grid(row=1, column=3, sticky=tk.W)
        glbl("備考:", 2, 0)
        g_note_var = tk.StringVar()
        ttk.Entry(gf, textvariable=g_note_var, width=36).grid(
            row=2, column=1, columnspan=3, sticky=tk.W + tk.E)

        grant_cols = ("id", "name", "grant_date", "days", "expiry", "note")
        grant_tree = ttk.Treeview(tab_grant, columns=grant_cols, show="headings", height=11)
        for cid, ht, w, anch in [
            ("id",         "ID",       50,  tk.CENTER),
            ("name",       "スタッフ名", 130, tk.W),
            ("grant_date", "付与日",   100, tk.CENTER),
            ("days",       "付与日数",  80, tk.CENTER),
            ("expiry",     "有効期限", 100, tk.CENTER),
            ("note",       "備考",     210, tk.W),
        ]:
            grant_tree.heading(cid, text=ht)
            grant_tree.column(cid, width=w, anchor=anch)
        grant_tree.tag_configure("odd",  background=C_ROW_ODD)
        grant_tree.tag_configure("even", background=C_ROW_EVEN)
        sb_grant = ttk.Scrollbar(tab_grant, orient=tk.VERTICAL, command=grant_tree.yview)
        grant_tree.configure(yscroll=sb_grant.set)

        def refresh_grant_tree():
            data = _load_paid_leave()
            for it in grant_tree.get_children():
                grant_tree.delete(it)
            for i, g in enumerate(sorted(data["grants"],
                                         key=lambda x: x.get("grant_date", ""), reverse=True)):
                tag = "even" if i % 2 == 0 else "odd"
                grant_tree.insert("", tk.END, tags=(tag,), values=(
                    g.get("id", ""), g.get("name", ""), g.get("grant_date", ""),
                    g.get("days", ""), g.get("expiry_date", ""), g.get("note", "")))

        def add_grant():
            name = g_name_var.get().strip()
            if not name:
                messagebox.showwarning("入力エラー", "スタッフ名を選択してください。", parent=dlg)
                return
            try:
                datetime.date.fromisoformat(g_date_var.get().strip())
                days = float(g_days_var.get().strip())
                if days <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("入力エラー", "付与日と付与日数（正の数）を正しく入力してください。", parent=dlg)
                return
            expiry = g_expiry_var.get().strip()
            if expiry:
                try:
                    datetime.date.fromisoformat(expiry)
                except ValueError:
                    messagebox.showwarning("入力エラー", "有効期限の形式が不正です（YYYY-MM-DD）。", parent=dlg)
                    return
            data = _load_paid_leave()
            data["grants"].append({
                "id":          _next_pl_id(data["grants"]),
                "name":        name,
                "grant_date":  g_date_var.get().strip(),
                "days":        days,
                "expiry_date": expiry or "9999-12-31",
                "note":        g_note_var.get().strip(),
            })
            _save_paid_leave(data)
            refresh_grant_tree()
            refresh_balance()
            self.set_status(f"「{name}」に {days}日 の有給を付与しました。")

        def delete_grant():
            sel = grant_tree.selection()
            if not sel:
                messagebox.showwarning("削除", "削除する付与レコードを選択してください。", parent=dlg)
                return
            vals = grant_tree.item(sel[0], "values")
            if not messagebox.askyesno("確認", f"ID:{vals[0]} の付与レコードを削除しますか？", parent=dlg):
                return
            data = _load_paid_leave()
            data["grants"] = [g for g in data["grants"] if str(g.get("id", "")) != str(vals[0])]
            _save_paid_leave(data)
            refresh_grant_tree()
            refresh_balance()

        btn_gf = tk.Frame(tab_grant, bg=C_BG)
        btn_gf.pack(pady=(2, 4))
        _color_btn(btn_gf, "付与登録", add_grant,    C_GRANT,    width=10).pack(side=tk.LEFT, padx=4)
        _color_btn(btn_gf, "削除",     delete_grant, C_BTN_DEL,  width=8).pack(side=tk.LEFT, padx=4)
        grant_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=(0, 8))
        sb_grant.pack(side=tk.RIGHT, fill=tk.Y, pady=(0, 8))
        refresh_grant_tree()

        # ── タブ3: 取得登録 ──────────────────────
        uf = tk.Frame(tab_usage, bg=C_BG)
        uf.pack(fill=tk.X, padx=12, pady=10)

        def ulbl(text, row, col):
            tk.Label(uf, text=text, bg=C_BG, font=FONT_MAIN).grid(
                row=row, column=col, sticky=tk.W, padx=(0, 4), pady=4)

        ulbl("スタッフ名:", 0, 0)
        u_name_var = tk.StringVar()
        ttk.Combobox(uf, textvariable=u_name_var, values=self.staff_list,
                     width=16, state="readonly").grid(row=0, column=1, sticky=tk.W)
        ulbl("取得日:", 0, 2)
        u_date_var = tk.StringVar(value=datetime.date.today().isoformat())
        ttk.Entry(uf, textvariable=u_date_var, width=14).grid(row=0, column=3, sticky=tk.W)
        ulbl("取得日数:", 1, 0)
        u_days_var = tk.StringVar(value="1")
        ttk.Entry(uf, textvariable=u_days_var, width=8).grid(row=1, column=1, sticky=tk.W)
        tk.Label(uf, text="(0.5=半日)", bg=C_BG, font=("Meiryo UI", 9)).grid(
            row=1, column=2, sticky=tk.W)
        ulbl("備考:", 2, 0)
        u_note_var = tk.StringVar()
        ttk.Entry(uf, textvariable=u_note_var, width=36).grid(
            row=2, column=1, columnspan=3, sticky=tk.W + tk.E)
        u_remain_var = tk.StringVar(value="残日数: - 日")
        tk.Label(uf, textvariable=u_remain_var, bg=C_BG, font=FONT_BOLD,
                 fg=C_GRANT).grid(row=0, column=4, padx=(16, 0), sticky=tk.W)

        def update_remain_lbl(*_):
            name = u_name_var.get().strip()
            if name:
                rem, _, _ = _calc_paid_leave_balance(name, _load_paid_leave())
                u_remain_var.set(f"残日数: {rem} 日")
            else:
                u_remain_var.set("残日数: - 日")
        u_name_var.trace_add("write", update_remain_lbl)

        usage_cols = ("id", "name", "date", "days", "note")
        usage_tree = ttk.Treeview(tab_usage, columns=usage_cols, show="headings", height=11)
        for cid, ht, w, anch in [
            ("id",   "ID",        50,  tk.CENTER),
            ("name", "スタッフ名", 130, tk.W),
            ("date", "取得日",    100, tk.CENTER),
            ("days", "取得日数",   80, tk.CENTER),
            ("note", "備考",      310, tk.W),
        ]:
            usage_tree.heading(cid, text=ht)
            usage_tree.column(cid, width=w, anchor=anch)
        usage_tree.tag_configure("odd",  background=C_ROW_ODD)
        usage_tree.tag_configure("even", background=C_ROW_EVEN)
        sb_usage = ttk.Scrollbar(tab_usage, orient=tk.VERTICAL, command=usage_tree.yview)
        usage_tree.configure(yscroll=sb_usage.set)

        def refresh_usage_tree():
            data = _load_paid_leave()
            for it in usage_tree.get_children():
                usage_tree.delete(it)
            for i, u in enumerate(sorted(data["usages"],
                                         key=lambda x: x.get("date", ""), reverse=True)):
                tag = "even" if i % 2 == 0 else "odd"
                usage_tree.insert("", tk.END, tags=(tag,), values=(
                    u.get("id", ""), u.get("name", ""), u.get("date", ""),
                    u.get("days", ""), u.get("note", "")))

        def add_usage():
            name = u_name_var.get().strip()
            if not name:
                messagebox.showwarning("入力エラー", "スタッフ名を選択してください。", parent=dlg)
                return
            try:
                datetime.date.fromisoformat(u_date_var.get().strip())
                days = float(u_days_var.get().strip())
                if days <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("入力エラー", "取得日と取得日数（正の数）を正しく入力してください。", parent=dlg)
                return
            data = _load_paid_leave()
            rem, _, _ = _calc_paid_leave_balance(name, data)
            if days > rem:
                if not messagebox.askyesno(
                        "残日数不足",
                        f"残日数({rem}日)を超えています。\n続けますか？", parent=dlg):
                    return
            data["usages"].append({
                "id":   _next_pl_id(data["usages"]),
                "name": name,
                "date": u_date_var.get().strip(),
                "days": days,
                "note": u_note_var.get().strip(),
            })
            _save_paid_leave(data)
            refresh_usage_tree()
            refresh_balance()
            update_remain_lbl()
            self.set_status(f"「{name}」{u_date_var.get().strip()} 有給{days}日を登録しました。")

        def delete_usage():
            sel = usage_tree.selection()
            if not sel:
                messagebox.showwarning("削除", "削除する取得レコードを選択してください。", parent=dlg)
                return
            vals = usage_tree.item(sel[0], "values")
            if not messagebox.askyesno("確認", f"ID:{vals[0]} の取得レコードを削除しますか？", parent=dlg):
                return
            data = _load_paid_leave()
            data["usages"] = [u for u in data["usages"] if str(u.get("id", "")) != str(vals[0])]
            _save_paid_leave(data)
            refresh_usage_tree()
            refresh_balance()
            update_remain_lbl()

        btn_uf = tk.Frame(tab_usage, bg=C_BG)
        btn_uf.pack(pady=(2, 4))
        _color_btn(btn_uf, "取得登録", add_usage,    C_GRANT,   width=10).pack(side=tk.LEFT, padx=4)
        _color_btn(btn_uf, "削除",     delete_usage, C_BTN_DEL, width=8).pack(side=tk.LEFT, padx=4)
        usage_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=(0, 8))
        sb_usage.pack(side=tk.RIGHT, fill=tk.Y, pady=(0, 8))
        refresh_usage_tree()

        # ── タブ4: 履歴一覧 ──────────────────────
        hist_ctrl = tk.Frame(tab_history, bg=C_BG)
        hist_ctrl.pack(fill=tk.X, padx=8, pady=6)
        tk.Label(hist_ctrl, text="スタッフ:", bg=C_BG, font=FONT_MAIN).pack(side=tk.LEFT)
        hist_name_var = tk.StringVar(value="(全員)")
        ttk.Combobox(hist_ctrl, textvariable=hist_name_var,
                     values=["(全員)"] + self.staff_list,
                     width=14, state="readonly").pack(side=tk.LEFT, padx=(4, 10))
        tk.Label(hist_ctrl, text="年:", bg=C_BG, font=FONT_MAIN).pack(side=tk.LEFT)
        hist_year_var = tk.StringVar(value=str(datetime.date.today().year))
        ttk.Entry(hist_ctrl, textvariable=hist_year_var, width=7).pack(side=tk.LEFT, padx=(4, 8))

        hist_cols = ("kind", "name", "date", "days", "note")
        hist_tree = ttk.Treeview(tab_history, columns=hist_cols, show="headings", height=14)
        for cid, ht, w, anch in [
            ("kind", "種別",       70,  tk.CENTER),
            ("name", "スタッフ名", 130, tk.W),
            ("date", "日付",      100, tk.CENTER),
            ("days", "日数",       80, tk.CENTER),
            ("note", "備考",      290, tk.W),
        ]:
            hist_tree.heading(cid, text=ht)
            hist_tree.column(cid, width=w, anchor=anch)
        hist_tree.tag_configure("grant", background="#D5F5E3")
        hist_tree.tag_configure("usage", background="#FDEBD0")
        sb_hist = ttk.Scrollbar(tab_history, orient=tk.VERTICAL, command=hist_tree.yview)
        hist_tree.configure(yscroll=sb_hist.set)

        def refresh_history():
            nf   = hist_name_var.get().strip()
            yf   = hist_year_var.get().strip()
            data = _load_paid_leave()
            for it in hist_tree.get_children():
                hist_tree.delete(it)
            entries = []
            for g in data["grants"]:
                if (nf == "(全員)" or g.get("name") == nf) and \
                        (not yf or g.get("grant_date", "").startswith(yf)):
                    entries.append(("付与", g["name"], g.get("grant_date", ""),
                                    g.get("days", ""), g.get("note", "")))
            for u in data["usages"]:
                if (nf == "(全員)" or u.get("name") == nf) and \
                        (not yf or u.get("date", "").startswith(yf)):
                    entries.append(("取得", u["name"], u.get("date", ""),
                                    u.get("days", ""), u.get("note", "")))
            entries.sort(key=lambda x: x[2], reverse=True)
            for rec in entries:
                hist_tree.insert("", tk.END, tags=("grant" if rec[0] == "付与" else "usage",),
                                 values=rec)

        _color_btn(hist_ctrl, "表示", refresh_history, C_BTN_REFR, width=8).pack(side=tk.LEFT)
        hist_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=(0, 8))
        sb_hist.pack(side=tk.RIGHT, fill=tk.Y, pady=(0, 8))
        refresh_history()

    # ── ログアウト ────────────────────────────────
    def on_logout(self):
        """ログアウト処理"""
        if messagebox.askyesno("ログアウト", "ログアウトしますか？"):
            self.destroy()


if __name__ == "__main__":
    _create_default_user()

    login = LoginWindow()
    login.mainloop()
    username = login.logged_in_user
    if username:
        app = AttendanceGUI(username=username)
        app.mainloop()
