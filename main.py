import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import json, os, threading, time
from pynput import keyboard, mouse
import configparser

# グローバル変数：各セクションの記録データ
click_data = {"pre": [], "clicks": [], "post": []}
auto_running = False
CONFIG_FILENAME = "config.ini"

# -------------------------- GUIクラス --------------------------
class ClickToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("いいかんじクリッカー")
        self.geometry("400x700")
        self.default_interval = tk.IntVar(value=500)
        self.loop_var = tk.BooleanVar(value=False)
        self.loop_count = tk.IntVar(value=1)  # ループ回数（正の数なら指定回数、0以下なら無限）
        self.last_used_file = tk.StringVar(value="")

        # 記録先指定はラジオボタンで管理
        self.record_phase_var = tk.StringVar(value="clicks")
        self.phase_buttons = {}

        self.create_widgets()
        self.load_file_list()
        self.start_listeners()
        self.load_settings()
        if self.last_used_file.get():
            self.load_file(self.last_used_file.get())
        self.record_phase_var.trace_add("write", self.update_phase_appearances)
        # ループON/OFFに応じてループ回数入力を無効化する
        self.loop_var.trace_add("write", self.update_loop_count_state)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        # 上部設定領域：デフォルト間隔、ループ設定、記録先指定
        top_frame = ttk.Frame(self)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(top_frame, text="デフォルト間隔(ms):").pack(side=tk.LEFT)
        ttk.Entry(top_frame, textvariable=self.default_interval, width=6).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(top_frame, text="ループ", variable=self.loop_var).pack(side=tk.LEFT, padx=10)
        ttk.Label(top_frame, text="回数:").pack(side=tk.LEFT, padx=5)
        self.loop_count_entry = ttk.Entry(top_frame, textvariable=self.loop_count, width=4)
        self.loop_count_entry.pack(side=tk.LEFT)

        # 記録先ラジオボタン群
        phase_frame = ttk.Frame(self)
        phase_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(phase_frame, text="記録先: (F12)").pack(side=tk.LEFT)
        for phase in ["pre", "clicks", "post"]:
            rb = tk.Radiobutton(phase_frame,
                                text=phase.upper(),
                                value=phase,
                                variable=self.record_phase_var,
                                indicatoron=True,
                                width=8)
            rb.pack(side=tk.LEFT, padx=3)
            self.phase_buttons[phase] = rb
        self.update_phase_appearances()

        # 各セクション（PRE, CLICKS, POST）
        self.trees = {}
        self.section_frames = {}
        for section in ["pre", "clicks", "post"]:
            cont = ttk.Frame(self)
            cont.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            self.section_frames[section] = cont

            # ヘッダーエリア：セクション名＋操作ボタン
            head = ttk.Frame(cont)
            head.pack(fill=tk.X)
            ttk.Label(head, text=section.upper()).pack(side=tk.LEFT)

            # ボタンは右詰めで配置（選択中レコードの操作対象となる）
            btn_delete = tk.Button(head, text="X", command=lambda s=section: self.delete_item_in_section(s), width=2)
            btn_delete.pack(side=tk.RIGHT, padx=1, pady=1)
            btn_down = tk.Button(head, text="▼", command=lambda s=section: self.move_item_down_in_section(s), width=2)
            btn_down.pack(side=tk.RIGHT, padx=1, pady=1)
            btn_up = tk.Button(head, text="▲", command=lambda s=section: self.move_item_up_in_section(s), width=2)
            btn_up.pack(side=tk.RIGHT, padx=1, pady=1)
            # 追加：選択中の座標に移動するボタン
            btn_move = tk.Button(head, text="移動", command=lambda s=section: self.move_cursor_to_selected(s), width=3)
            btn_move.pack(side=tk.RIGHT, padx=1, pady=1)

            # ツリービュー
            tree = ttk.Treeview(cont, columns=("x", "y", "interval"), show="headings", height=5)
            tree.heading("x", text="X")
            tree.heading("y", text="Y")
            tree.heading("interval", text="間隔(ms)")
            tree.column("x", width=50, anchor="center")
            tree.column("y", width=50, anchor="center")
            tree.column("interval", width=70, anchor="center")
            tree.pack(fill=tk.BOTH, expand=True)
            tree.bind("<Double-1>", lambda e, s=section: self.edit_cell(e, s))
            self.trees[section] = tree

        # 下部操作領域
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="新規", command=self.clear_records).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="設定保存", command=self.save_current).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="開始", command=self.start_auto_click).pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_frame, text="※実行は右クリックで停止").pack(side=tk.LEFT, padx=5)

        file_frame = ttk.Frame(self)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        self.file_list = tk.Listbox(file_frame, height=4)
        self.file_list.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="読み込み", command=self.load_selected).pack(side=tk.LEFT, padx=5)

    def update_phase_appearances(self, *args):
        selected = self.record_phase_var.get()
        for phase, rb in self.phase_buttons.items():
            rb.configure(fg="black" if phase == selected else "gray")

    def update_loop_count_state(self, *args):
        if self.loop_var.get():
            self.loop_count_entry.config(state="disabled")
        else:
            self.loop_count_entry.config(state="normal")

    def delete_item_in_section(self, section):
        tree = self.trees[section]
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("削除", f"{section.upper()} で削除する行を選択してください。")
            return
        for item in selected:
            tree.delete(item)
        self.update_record_from_tree(section)

    def move_item_up_in_section(self, section):
        tree = self.trees[section]
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("移動", f"{section.upper()} で移動する行を選択してください。")
            return
        for item in selected:
            index = tree.index(item)
            if index > 0:
                tree.move(item, "", index - 1)
        self.update_record_from_tree(section)

    def move_item_down_in_section(self, section):
        tree = self.trees[section]
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("移動", f"{section.upper()} で移動する行を選択してください。")
            return
        for item in reversed(selected):
            index = tree.index(item)
            if index < len(tree.get_children()) - 1:
                tree.move(item, "", index + 1)
        self.update_record_from_tree(section)

    def edit_cell(self, event, section):
        item_id = self.trees[section].focus()
        col = self.trees[section].identify_column(event.x)
        if col not in ("#1", "#2", "#3"):
            return
        x, y, width, height = self.trees[section].bbox(item_id, col)
        value = self.trees[section].set(item_id, col)
        entry = tk.Entry(self.trees[section])
        entry.insert(0, value)
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus()

        def on_enter(event):
            self.trees[section].set(item_id, col, entry.get())
            entry.destroy()
            self.update_record_from_tree(section)

        entry.bind("<Return>", on_enter)
        entry.bind("<FocusOut>", lambda e: entry.destroy())

    def move_cursor_to_selected(self, section):
        tree = self.trees[section]
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("移動", f"{section.upper()} で座標を選択してください。")
            return
        # 先頭の選択アイテムの座標（X, Y）を取得
        vals = tree.item(selected[0], "values")
        try:
            x = int(vals[0])
            y = int(vals[1])
            mouse.Controller().position = (x, y)
        except Exception as e:
            messagebox.showerror("移動エラー", str(e))

    def add_recorded_click(self, rec, section):
        self.trees[section].insert("", tk.END, values=(rec["x"], rec["y"], rec["interval"]))

    def clear_records(self):
        for section in ("pre", "clicks", "post"):
            click_data[section] = []
            for item in self.trees[section].get_children():
                self.trees[section].delete(item)

    def update_record_from_tree(self, section=None):
        targets = [section] if section else ["pre", "clicks", "post"]
        for sec in targets:
            click_data[sec] = []
            for item in self.trees[sec].get_children():
                vals = self.trees[sec].item(item, "values")
                click_data[sec].append({"x": int(vals[0]), "y": int(vals[1]), "interval": int(vals[2])})

    def save_current(self):
        self.save_settings()
        self.update_record_from_tree()
        filename = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if filename:
            with open(filename, "w") as f:
                json.dump(click_data, f, indent=2)
            self.last_used_file.set(filename)
            self.save_settings()
            self.load_file_list()

    def load_selected(self):
        sel = self.file_list.curselection()
        if not sel:
            return
        fname = self.file_list.get(sel[0])
        self.load_file(fname)
        self.last_used_file.set(fname)
        self.save_settings()

    def load_file(self, fname):
        global click_data
        try:
            with open(fname, "r") as f:
                click_data = json.load(f)
            self.show_all_sections()
        except Exception as e:
            messagebox.showerror("読み込みエラー", str(e))

    def load_file_list(self):
        self.file_list.delete(0, tk.END)
        for f in os.listdir('.'):
            if f.endswith(".json"):
                self.file_list.insert(tk.END, f)

    def load_last_used(self):
        if self.last_used_file.get() and os.path.exists(self.last_used_file.get()):
            self.load_file(self.last_used_file.get())

    def show_all_sections(self):
        for section in ("pre", "clicks", "post"):
            for item in self.trees[section].get_children():
                self.trees[section].delete(item)
            for rec in click_data.get(section, []):
                self.add_recorded_click(rec, section)

    def start_auto_click(self):
        global auto_running
        self.update_record_from_tree()
        auto_running = True
        threading.Thread(target=auto_clicker, daemon=True).start()

    def start_listeners(self):
        keyboard.Listener(on_press=on_key_press, daemon=True).start()

    def load_settings(self):
        config = configparser.ConfigParser()
        if os.path.exists(CONFIG_FILENAME):
            try:
                config.read(CONFIG_FILENAME)
                if 'Settings' in config:
                    self.default_interval.set(config['Settings'].getint("default_interval", 300))
                    self.loop_var.set(config['Settings'].getboolean("loop_var", False))
                    self.loop_count.set(config['Settings'].getint("loop_count", 1))
                    self.last_used_file.set(config['Settings'].get("last_used_file", ""))
            except configparser.Error as e:
                print(f"設定ファイルの読み込みエラー: {e}")

    def save_settings(self):
        config = configparser.ConfigParser()
        config['Settings'] = {
            'default_interval': self.default_interval.get(),
            'loop_var': self.loop_var.get(),
            'loop_count': self.loop_count.get(),
            'last_used_file': self.last_used_file.get()
        }
        try:
            with open(CONFIG_FILENAME, "w") as f:
                config.write(f)
        except IOError as e:
            print(f"設定ファイルの保存エラー: {e}")

    def on_closing(self):
        self.save_settings()
        self.destroy()

# -------------------------- F12記録 --------------------------
def on_key_press(key):
    try:
        if key == keyboard.Key.f12:
            pos = mouse.Controller().position
            rec = {"x": pos[0], "y": pos[1], "interval": app.default_interval.get()}
            phase = app.record_phase_var.get()
            click_data[phase].append(rec)
            app.add_recorded_click(rec, phase)
    except Exception as e:
        print("F12記録エラー:", e)

# -------------------------- 自動クリック実行 --------------------------
def auto_clicker():
    global auto_running
    mctrl = mouse.Controller()
    stop_requested = False

    def do_clicks(seq, ignore_stop=False):
        for rec in seq:
            if not ignore_stop and stop_requested:
                break
            mctrl.position = (rec["x"], rec["y"])
            mctrl.click(mouse.Button.left)
            time.sleep(rec["interval"] / 1000.0)

    def on_right_click(x, y, button, pressed):
        nonlocal stop_requested
        if button == mouse.Button.right and pressed:
            stop_requested = True
            return False

    mouse.Listener(on_click=on_right_click, daemon=True).start()

    # PRE工程
    do_clicks(click_data.get("pre", []))
    # CLICKS工程
    if app.loop_var.get():
        while not stop_requested:
            do_clicks(click_data.get("clicks", []))
    else:
        do_clicks(click_data.get("clicks", []))
    # POST工程（停止要求を無視）
    do_clicks(click_data.get("post", []), ignore_stop=True)
    auto_running = False

# -------------------------- 起動 --------------------------
if __name__ == "__main__":
    app = ClickToolApp()
    app.mainloop()
