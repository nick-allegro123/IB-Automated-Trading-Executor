import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import threading, time, os, json
import requests
from datetime import datetime
import winsound
import queue
from concurrent.futures import ThreadPoolExecutor
import threading

# OANDA API 全域變數
BASE_URL = None
account_type = None
API_TOKEN = None
ACCOUNT_ID = None
DEMO_URL = "https://api-fxpractice.oanda.com/v3"
LIVE_URL = "https://api-fxtrade.oanda.com/v3"

# Telegram 全域變數
TELEGRAM_BOT_TOKEN = None
TELEGRAM_CHAT_ID = None
TELEGRAM_ENABLED = False
TELEGRAM_MESSAGE_QUEUE = queue.Queue()

# 修改：設定檔使用程式所在目錄的絕對路徑
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "oanda_order_settings.json")

#[Telegram]
#--------------------------------------------------------------------

def send_telegram_message():
    """發送 Telegram 通知"""
    global TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, TELEGRAM_ENABLED
    markdown_escape_chars = r'_*[]()~`>#+-=|{}.!'
    
    while True:
        try:
            message = TELEGRAM_MESSAGE_QUEUE.get()
            if not TELEGRAM_ENABLED or not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
                TELEGRAM_MESSAGE_QUEUE.task_done()
                continue
            escaped_message = ''.join('\\' + c if c in markdown_escape_chars else c for c in message)
            url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
            payload = {
                "chat_id": TELEGRAM_CHAT_ID,
                "text": escaped_message,
                "parse_mode": "MarkdownV2"
            }
            response = requests.post(url, json=payload)
            if response.status_code != 200:
                print(f"Telegram 訊息發送失敗: {response.text}")
            TELEGRAM_MESSAGE_QUEUE.task_done()
        except Exception as e:
            print(f"Telegram 訊息發送錯誤: {e}")
            TELEGRAM_MESSAGE_QUEUE.task_done()

#--------------------------------------------------------------------

def check_auth(app_instance=None):
    """檢查 OANDA API 連線是否正常"""
    url = f"{BASE_URL}/accounts/{ACCOUNT_ID}"
    headers = {
        "Authorization": f"Bearer {API_TOKEN}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return True
    else:
        print(f"OANDA API 連線失敗: {response.text}")
        if app_instance:
            app_instance.log(f"OANDA API 連線失敗: {response.text}")
        return False

def place_order(product_code, direction, size, app_instance=None, is_manual=False):
    """下單，使用 OANDA v20 API"""
    if not is_manual and app_instance and not app_instance.is_ordering_enabled:
        return {"status": "error", "message": "下單功能已停止，僅允許手動下單"}
    
    url = f"{BASE_URL}/accounts/{ACCOUNT_ID}/orders"
    headers = {
        "Authorization": f"Bearer {API_TOKEN}",
        "Content-Type": "application/json"
    }
    units = float(size) if direction == "BUY" else -float(size)
    payload = {
        "order": {
            "type": "MARKET",
            "instrument": product_code,
            "units": str(units),
            "timeInForce": "FOK",
            "positionFill": "DEFAULT"
        }
    }
    
    response = requests.post(url, json=payload, headers=headers)
    if response.status_code == 201:
        result = response.json()
        print(f"[{product_code}] 下單: {direction} {size} 單 - 成功: {result}")
        if app_instance:
            app_instance.log(f"[{product_code}] 下單成功: {direction} {size} 單")
        winsound.Beep(1000, 200)
        winsound.Beep(1200, 200)
        return {"status": "success", "orderId": result.get("orderFillTransaction", {}).get("id")}
    else:
        print(f"[{product_code}] 下單失敗: {response.text}")
        if app_instance:
            app_instance.log(f"[{product_code}] 下單失敗: {response.text}")
        return {"status": "error", "message": response.text}

def get_account_balance(app_instance=None):
    """查詢 OANDA 帳戶餘額"""
    url = f"{BASE_URL}/accounts/{ACCOUNT_ID}/summary"
    headers = {
        "Authorization": f"Bearer {API_TOKEN}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        result = response.json()
        balance = result["account"]["balance"]
        currency = result["account"]["currency"]
        if app_instance:
            app_instance.log(f"帳戶餘額: {balance} {currency}")
        return balance
    else:
        if app_instance:
            app_instance.log(f"查詢餘額失敗: {response.text}")
        return None

def get_account_nav(app_instance=None):
    """查詢 OANDA 帳戶淨值"""
    url = f"{BASE_URL}/accounts/{ACCOUNT_ID}/summary"
    headers = {
        "Authorization": f"Bearer {API_TOKEN}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        result = response.json()
        balance = float(result["account"]["balance"])
        unrealized_pl = float(result["account"]["unrealizedPL"])
        nav = balance + unrealized_pl
        currency = result["account"]["currency"]
        if app_instance:
            app_instance.log(f"帳戶淨值: {nav:.4f} {currency} (餘額: {balance}, 未實現損益: {unrealized_pl})")
        return nav
    else:
        if app_instance:
            app_instance.log(f"查詢淨值失敗: {response.text}")
        return None

def check_positions(app_instance=None):
    """查詢 OANDA 帳戶當前持有部位並與策略比較"""
    url = f"{BASE_URL}/accounts/{ACCOUNT_ID}/positions"
    headers = {
        "Authorization": f"Bearer {API_TOKEN}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        if app_instance:
            app_instance.log(f"查詢部位失敗: {response.text}")
        return
    
    result = response.json()
    api_positions = {}
    for pos in result.get("positions", []):
        instrument = pos["instrument"]
        long_units = float(pos["long"]["units"])
        short_units = float(pos["short"]["units"])
        net_units = long_units + short_units
        api_positions[instrument] = net_units

    strategy_positions = {}
    for strat in app_instance.strategies:
        product = strat.product_code
        total = strat.actual_position * strat.unit
        if product in strategy_positions:
            strategy_positions[product] += total
        else:
            strategy_positions[product] = total

    inconsistent = {}
    all_products = set(api_positions.keys()).union(strategy_positions.keys())
    for product in all_products:
        api_pos = api_positions.get(product, 0)
        strat_pos = strategy_positions.get(product, 0)
        if api_pos != strat_pos:
            inconsistent[product] = api_pos

    if not inconsistent:
        app_instance.log("帳戶倉位與部位總和一致")
    else:
        inconsistent_str = ", ".join(f"{prod}: {pos}" for prod, pos in inconsistent.items())
        app_instance.log(f"{inconsistent_str} 與部位總和不一致，請確認是否故意不一致或有其它異常狀況")

class Strategy:
    def __init__(self, file_path, product_code, unit, init_position):
        self.file_path = file_path
        self.product_code = product_code
        self.unit = float(unit)
        self.init_position = int(init_position)
        self.actual_position = int(init_position)
        self.last_strategy_value = None
        self.current_signal = None
        self.change_timestamps = []
        self.monitoring = False

    def to_dict(self):
        return {
            "file_path": self.file_path,
            "product_code": self.product_code,
            "unit": self.unit,
            "init_position": self.init_position,
            "actual_position": self.actual_position,
        }

    @classmethod
    def from_dict(cls, d):
        strat = cls(d["file_path"], d["product_code"], d["unit"], d["init_position"])
        strat.actual_position = d.get("actual_position", strat.init_position)
        return strat

class OandaOrderApp:
    def __init__(self, master):
        self.master = master
        self.master.title("進階 OANDA API 下單程式")
        self.strategies = []
        self.strategy_controls = {}
        self.global_freq = 1.0
        self.sort_reverse = {}
        self.last_click_time = {}
        self.risk_window = 50.0
        self.risk_threshold = 5
        self.is_ordering_enabled = False
        self.last_order_toggle_click = 0
        self.order_status_canvas = None
        self.telegram_enabled_var = None
        self.create_widgets()
        self.load_settings()
        self.executor = ThreadPoolExecutor(max_workers=99)
        telegram_thread = threading.Thread(target=send_telegram_message, daemon=True)
        telegram_thread.start()
        if check_auth(self):
            self.log("OANDA API 連線成功!!")
            self.log("邏輯判定上只要成功送出下單指令就是下單成功，除非要執行系統測試否則不要在非交易時段下單，不小心在非交易時段下單也不用擔心只需調整實際部位到正確數字即可，非交易時段訂單會自動被OANDA後台ORDER_CANCEL，不會轉為日內有效單。")
            print("OANDA API 連線成功!!")
        else:
            self.log("無法連接到 OANDA API，請檢查憑證")
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.update_time()
        self.center_window()

    def center_window(self):
        self.master.update_idletasks()
        width = self.master.winfo_width()
        height = self.master.winfo_height()
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.master.geometry(f"{width}x{height}+{x}+{y}")

    def on_closing(self):
        self.executor.shutdown(wait=True)
        self.save_settings()
        self.master.destroy()

    def create_widgets(self):
        main_frame = tk.Frame(self.master)
        main_frame.pack(fill=tk.BOTH, expand=True)

        top_frame = tk.Frame(main_frame)
        top_frame.pack(padx=5, pady=5, fill=tk.X)
        tk.Label(top_frame, text="共用監控頻率 (秒/次):").pack(side=tk.LEFT)
        self.freq_var = tk.StringVar(value=str(self.global_freq))
        self.freq_entry = tk.Entry(top_frame, width=5, textvariable=self.freq_var)
        self.freq_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(top_frame, text="更新頻率", command=self.update_global_freq).pack(side=tk.LEFT, padx=5)
        tk.Label(top_frame, text="當前時間:").pack(side=tk.LEFT, padx=10)
        self.time_label = tk.Label(top_frame, text="")
        self.time_label.pack(side=tk.LEFT)
        tk.Button(top_frame, text="設定 Telegram", command=self.set_telegram_settings).pack(side=tk.RIGHT, padx=6)
        self.telegram_enabled_var = tk.BooleanVar(value=TELEGRAM_ENABLED)
        tk.Checkbutton(top_frame, text="通知(傳送 LOG)", variable=self.telegram_enabled_var, command=self.toggle_telegram).pack(side=tk.RIGHT, padx=10)

        middle_frame = tk.Frame(main_frame)
        middle_frame.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(middle_frame, columns=("file", "product", "unit", "position", "signal"), show="tree headings", height=10)
        self.tree.heading("#0", text="狀態")
        self.tree.heading("file", text="策略檔案")
        self.tree.heading("product", text="商品代號")
        self.tree.heading("unit", text="單位")
        self.tree.heading("position", text="實際部位")
        self.tree.heading("signal", text="策略訊號")
        self.tree.column("#0", width=80, anchor="center")
        self.tree.column("file", width=200)
        self.tree.column("product", width=120)
        self.tree.column("unit", width=60)
        self.tree.column("position", width=80)
        self.tree.column("signal", width=80)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree.bind("<Double-1>", self.on_tree_double_click)
        for col in ("#0", "file", "product", "unit", "position", "signal"):
            self.tree.heading(col, command=lambda c=col: self.check_double_click(c))

        self.position_sum_frame = tk.LabelFrame(middle_frame, text="實際部位總和", padx=5, pady=5)
        self.position_sum_frame.pack(side=tk.RIGHT, padx=5, fill=tk.Y)
        self.position_sum_tree = ttk.Treeview(self.position_sum_frame, columns=("product", "quantity"), show="headings", height=10)
        self.position_sum_tree.heading("product", text="商品")
        self.position_sum_tree.heading("quantity", text="數量")
        self.position_sum_tree.column("product", width=120, anchor="center")
        self.position_sum_tree.column("quantity", width=80, anchor="center")
        self.position_sum_tree.pack(fill=tk.Y)
        self.update_position_sum()

        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(padx=5, pady=5, fill=tk.X)
        tk.Button(btn_frame, text="新增策略", command=self.add_strategy).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="刪除策略", command=self.remove_strategy).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="更新實際部位", command=self.manual_update_position).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="手動下單", command=self.manual_order).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="查詢餘額", command=self.check_balance).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="查詢淨值", command=self.check_nav).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="檢查倉位", command=self.check_positions).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="設定風控 (共用)", command=self.set_risk_settings).pack(side=tk.LEFT, padx=5)
        self.order_toggle_button = tk.Button(btn_frame, text="開始/停止下單 (雙擊)")
        self.order_toggle_button.pack(side=tk.RIGHT, padx=5)
        self.order_toggle_button.bind("<Button-1>", self.check_order_toggle_click)
        self.order_status_canvas = tk.Canvas(btn_frame, width=15, height=15)
        self.order_status_canvas.pack(side=tk.RIGHT, padx=5)
        self.update_order_status_light()

        self.txt_status = tk.Text(main_frame, height=15, width=100)
        self.txt_status.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)

    def log(self, msg):
        timestamp = time.strftime("[%Y-%m-%d %H:%M:%S]")
        def update_log():
            self.txt_status.insert(tk.END, f"{timestamp} {msg}\n")
            self.txt_status.see(tk.END)
            lines = self.txt_status.get("1.0", tk.END).splitlines()
            if len(lines) > 600:
                self.txt_status.delete("1.0", f"{len(lines) - 200}.0")
        self.master.after(0, update_log)
        if TELEGRAM_ENABLED:
            TELEGRAM_MESSAGE_QUEUE.put(f"{timestamp} {msg}")

    def update_time(self):
        current_time = time.strftime("%Y-%m-%d %H:%M:%S")
        self.time_label.config(text=current_time)
        self.master.after(1000, self.update_time)

    def update_global_freq(self):
        try:
            self.global_freq = float(self.freq_var.get())
            self.log(f"更新共用監控頻率為 {self.global_freq} 秒/次")
        except Exception as e:
            messagebox.showwarning("錯誤", f"監控頻率必須為數字：{e}", parent=self.master)

    def set_risk_settings(self):
        class RiskSettingsDialog(simpledialog.Dialog):
            def __init__(self, parent, title, app_instance):
                self.app_instance = app_instance
                super().__init__(parent, title)

            def body(self, master):
                tk.Label(master, text="[非高頻策略通常10秒內超過3次即為異常]").pack()
                tk.Label(master, text=" ").pack()
                tk.Label(master, text="多少時間內(秒):").pack()
                self.window_var = tk.StringVar(value=str(self.app_instance.risk_window))
                tk.Entry(master, textvariable=self.window_var).pack(padx=5, pady=5)
                tk.Label(master, text="N次後風控觸發:").pack()
                self.threshold_var = tk.StringVar(value=str(self.app_instance.risk_threshold))
                tk.Entry(master, textvariable=self.threshold_var).pack(padx=5, pady=5)
                return None

            def apply(self):
                self.result = (self.window_var.get(), self.threshold_var.get())

        dialog = RiskSettingsDialog(self.master, "設定風控參數", self)
        if dialog.result:
            try:
                new_window = float(dialog.result[0])
                new_threshold = int(dialog.result[1])
                if new_window <= 0 or new_threshold <= 0:
                    raise ValueError("時間窗口和次數閾值必須為正數")
                self.risk_window = new_window
                self.risk_threshold = new_threshold
                self.log(f"更新風控設定為 {self.risk_window} 秒內 {self.risk_threshold} 次")
            except Exception as e:
                messagebox.showwarning("錯誤", f"風控設定無效：{e}", parent=self.master)

    def set_telegram_settings(self):
        class TelegramSettingsDialog(simpledialog.Dialog):
            def __init__(self, parent, title, app_instance):
                self.app_instance = app_instance
                super().__init__(parent, title)

            def body(self, master):
                tk.Label(master, text="Telegram Bot Token:").pack()
                self.token_var = tk.StringVar(value=TELEGRAM_BOT_TOKEN or "")
                tk.Entry(master, textvariable=self.token_var, width=50).pack(padx=5, pady=5)
                tk.Label(master, text="Telegram Chat ID:").pack()
                self.chat_id_var = tk.StringVar(value=TELEGRAM_CHAT_ID or "")
                tk.Entry(master, textvariable=self.chat_id_var, width=50).pack(padx=5, pady=5)
                return None

            def apply(self):
                self.result = (self.token_var.get(), self.chat_id_var.get())

        dialog = TelegramSettingsDialog(self.master, "設定 Telegram 通知", self)
        if dialog.result:
            try:
                global TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
                TELEGRAM_BOT_TOKEN = dialog.result[0]
                TELEGRAM_CHAT_ID = dialog.result[1]
                if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
                    raise ValueError("Token 和 Chat ID 不能為空")
                self.log("Telegram 設定更新成功")
                self.save_settings()
            except Exception as e:
                messagebox.showwarning("錯誤", f"Telegram 設定無效：{e}", parent=self.master)

    def toggle_telegram(self):
        global TELEGRAM_ENABLED
        TELEGRAM_ENABLED = self.telegram_enabled_var.get()
        status = "啟用" if TELEGRAM_ENABLED else "停用"
        self.log(f"Telegram 通知已{status}")
        self.save_settings()

    def update_position_sum(self):
        for item in self.position_sum_tree.get_children():
            self.position_sum_tree.delete(item)

        position_totals = {}
        for strat in self.strategies:
            product = strat.product_code
            total = strat.actual_position * strat.unit
            if product in position_totals:
                position_totals[product] += total
            else:
                position_totals[product] = total
        
        for product, total in position_totals.items():
            self.position_sum_tree.insert("", tk.END, values=(product, f"{total:.2f}"))

    def queue_gui_update(self, func, *args):
        self.master.after(0, func, *args)

    def update_tree_signal(self, fid, value):
        self.tree.set(fid, column="signal", value=str(value))

    def update_tree_position(self, fid, value):
        self.tree.set(fid, column="position", value=str(value))

    def update_tree_status(self, fid, text):
        self.tree.item(fid, text=text)

    def check_order_toggle_click(self, event):
        current_time = time.time()
        if current_time - self.last_order_toggle_click < 0.5:
            self.toggle_ordering()
        self.last_order_toggle_click = current_time
        return "break"

    def toggle_ordering(self):
        self.is_ordering_enabled = not self.is_ordering_enabled
        self.update_order_status_light()
        status = "開始下單" if self.is_ordering_enabled else "停止下單"
        self.log(f"下單狀態切換為: {status}")

    def update_order_status_light(self):
        self.order_status_canvas.delete("all")
        color = "#00FF00" if self.is_ordering_enabled else "#FF0000"
        self.order_status_canvas.create_oval(5, 5, 15, 15, fill=color, outline="black")

    def execute_order_in_main_thread(self, product_code, direction, size):
        result = {"status": None, "message": None}
        event = threading.Event()

        def run_order():
            nonlocal result
            try:
                result = place_order(product_code, direction, size, self, is_manual=False)
            except Exception as e:
                result = {"status": "error", "message": str(e)}
            finally:
                event.set()

        self.executor.submit(run_order)
        event.wait()
        return result

    def add_strategy(self):
        while True:
            file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")], parent=self.master)
            if not file_path:
                return
            if any(strat.file_path == file_path for strat in self.strategies):
                messagebox.showwarning("警告", "此策略檔案已存在，請選擇其他檔案", parent=self.master)
                continue
            break
        
        product_options = [
            "======(Forex)","EUR_USD", "USD_JPY", "EUR_JPY", "GBP_JPY","GBP_USD", "EUR_GBP", "USD_CNH",
            "======(Crypto)","BTC_USD",
            "======(Gold)","XAU_USD",
            "======(Oil)","WTICO_USD",
            "======(Index)","SPX500_USD", "NAS100_USD", "US2000_USD", "JP225_USD",
            "======(Metal)","XCU_USD",
            "======(Bond)","USB10Y_USD", "USB05Y_USD", "USB02Y_USD", "USB30Y_USD"
        ]

        class ProductDialog(simpledialog.Dialog):
            def body(self, master):
                tk.Label(master, text="請選擇或輸入商品代號：").pack()
                self.product_var = tk.StringVar(value="EUR_USD")
                self.product_menu = ttk.Combobox(master, textvariable=self.product_var, values=product_options)
                self.product_menu.pack(padx=5, pady=5)
                return self.product_menu

            def apply(self):
                self.result = self.product_var.get()

        dialog = ProductDialog(self.master, title="選擇商品代號")
        product_code = dialog.result
        if not product_code:
            return

        unit = simpledialog.askstring("輸入", "請輸入下單單位 (例如 1000)：", parent=self.master)
        if not unit:
            return
        init_position = simpledialog.askstring("輸入", "請輸入目前實際持有部位 (例如 0)：", parent=self.master)
        if not init_position:
            return
        
        try:
            float(unit)
            int(init_position)
        except ValueError:
            messagebox.showwarning("錯誤", "單位必須為數字，實際持有部位必須為整數！", parent=self.master)
            return
        
        strat = Strategy(file_path, product_code, unit, init_position)
        self.strategies.append(strat)
        self.tree.insert("", tk.END, iid=strat.file_path, text="OFF", values=(file_path, product_code, unit, strat.actual_position, ""))
        self.strategy_controls[strat.file_path] = {"running": False, "thread": None}
        self.log(f"新增策略: {file_path} | {product_code}")
        self.update_position_sum()

    def remove_strategy(self):
        selected = self.tree.selection()
        if not selected:
            return
        fid = selected[0]
        if fid in self.strategy_controls and self.strategy_controls[fid]["running"]:
            self.strategy_controls[fid]["running"] = False
        self.tree.delete(fid)
        self.strategies = [s for s in self.strategies if s.file_path != fid]
        if fid in self.strategy_controls:
            del self.strategy_controls[fid]
        self.log(f"已刪除策略: {fid}")
        self.update_position_sum()

    def manual_update_position(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "請選擇要更新的策略", parent=self.master)
            return
        fid = selected[0]
        strat = next((s for s in self.strategies if s.file_path == fid), None)
        if not strat:
            return
        new_position = simpledialog.askstring("更新實際部位", f"請輸入新的實際部位 (目前為 {strat.actual_position})：", parent=self.master)
        try:
            strat.actual_position = int(new_position)
            self.log(f"[{strat.product_code}] 實際部位更新為 {strat.actual_position}")
            self.tree.set(fid, column="position", value=str(strat.actual_position))
            self.update_position_sum()
        except Exception as e:
            messagebox.showwarning("錯誤", f"更新失敗: {e}", parent=self.master)

    def manual_order(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "請選擇要手動下單的策略", parent=self.master)
            return
        fid = selected[0]
        strat = next((s for s in self.strategies if s.file_path == fid), None)
        if not strat:
            return

        class ManualOrderDialog(simpledialog.Dialog):
            def body(self, master):
                tk.Label(master, text=f"商品: {strat.product_code}").pack()
                tk.Label(master, text="請選擇方向:").pack()
                self.direction_var = tk.StringVar(value="BUY")
                tk.Radiobutton(master, text="BUY", variable=self.direction_var, value="BUY").pack(anchor=tk.W)
                tk.Radiobutton(master, text="SELL", variable=self.direction_var, value="SELL").pack(anchor=tk.W)
                tk.Label(master, text="請輸入下單單位:").pack()
                self.size_var = tk.StringVar(value=str(strat.unit))
                tk.Entry(master, textvariable=self.size_var).pack()
                return None

            def apply(self):
                self.result = (self.direction_var.get(), self.size_var.get())

        dialog = ManualOrderDialog(self.master, title="手動下單")
        if not dialog.result:
            return
        direction, size = dialog.result

        try:
            size = float(size)
            if size <= 0:
                raise ValueError("下單單位必須為正數")
        except ValueError as e:
            messagebox.showwarning("錯誤", f"下單單位無效: {e}", parent=self.master)
            return

        result = place_order(strat.product_code, direction, size, self, is_manual=True)
        if result["status"] == "success":
            if direction == "BUY":
                strat.actual_position += int(size / strat.unit)
            else:
                strat.actual_position -= int(size / strat.unit)
            self.tree.set(fid, column="position", value=str(strat.actual_position))
            self.log(f"[{strat.product_code}] 手動下單成功: {direction} {size} 單")
            self.update_position_sum()
        else:
            self.log(f"[{strat.product_code}] 手動下單失敗: {result['message']}")

    def check_balance(self):
        get_account_balance(self)

    def check_nav(self):
        get_account_nav(self)

    def check_positions(self):
        check_positions(self)

    def on_tree_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.toggle_monitor(item)

    def toggle_monitor(self, fid):
        strat = next((s for s in self.strategies if s.file_path == fid), None)
        if not strat:
            return
        ctrl = self.strategy_controls.get(fid)
        if not ctrl:
            return

        if not ctrl["running"]:
            ctrl["running"] = True
            strat.monitoring = True
            self.tree.item(fid, text="ON")
            ctrl["thread"] = threading.Thread(target=self.monitor_strategy, args=(strat, ctrl), daemon=True)
            ctrl["thread"].start()
            self.log(f"[{strat.file_path}] 開始監控。")
        else:
            ctrl["running"] = False
            strat.monitoring = False
            strat.last_strategy_value = None
            strat.change_timestamps = []
            self.tree.item(fid, text="OFF")
            self.log(f"[{strat.file_path}] 停止監控。")

    def monitor_strategy(self, strat: Strategy, control: dict):
        while control["running"]:
            try:
                if not os.path.isfile(strat.file_path):
                    self.log(f"檔案不存在: {strat.file_path}")
                    time.sleep(self.global_freq)
                    continue

                with open(strat.file_path, "r") as f:
                    lines = f.readlines()
                if not lines:
                    time.sleep(self.global_freq)
                    continue

                last_line = lines[-1].strip()
                parts = last_line.split()
                if len(parts) < 2:
                    self.log(f"檔案格式錯誤 (不足兩段): {strat.file_path}")
                    time.sleep(self.global_freq)
                    continue
                strat_part = parts[1]
                strat_values = strat_part.split(",")
                if len(strat_values) < 2:
                    self.log(f"檔案格式錯誤 (策略值不足): {strat.file_path}")
                    time.sleep(self.global_freq)
                    continue
                try:
                    current_strategy_value = int(strat_values[1])
                except ValueError:
                    self.log(f"策略值格式錯誤: {strat.file_path}")
                    time.sleep(self.global_freq)
                    continue

                strat.current_signal = current_strategy_value
                self.queue_gui_update(self.update_tree_signal, strat.file_path, current_strategy_value)

                now = time.time()
                if strat.last_strategy_value is None or current_strategy_value != strat.last_strategy_value:
                    strat.change_timestamps.append(now)
                    strat.change_timestamps = [t for t in strat.change_timestamps if now - t <= self.risk_window]
                    if len(strat.change_timestamps) >= (self.risk_threshold + 1):
                        self.log(f"[{strat.file_path}] 風控觸發: {self.risk_window}秒內變動超過{self.risk_threshold}次，關閉策略")
                        control["running"] = False
                        strat.monitoring = False
                        self.queue_gui_update(self.update_tree_status, strat.file_path, "OFF")
                        if strat.actual_position != 0:
                            direction = "SELL" if strat.actual_position > 0 else "BUY"
                            order_size = abs(strat.actual_position) * strat.unit
                            self.log(f"[{strat.product_code}] 平倉: {direction} {order_size} 單")
                            result = self.execute_order_in_main_thread(strat.product_code, direction, order_size)
                            self.log(f"[{strat.file_path}] 平倉結果: {result}")
                            if result["status"] == "success":
                                strat.actual_position = 0
                                self.queue_gui_update(self.update_tree_position, strat.file_path, 0)
                                self.queue_gui_update(self.update_position_sum)
                        strat.last_strategy_value = None
                        strat.change_timestamps = []
                        self.log(f"[{strat.file_path}] 策略訊號初始值已重置為 None")
                        break

                if strat.last_strategy_value is None:
                    strat.last_strategy_value = current_strategy_value
                    self.log(f"[{strat.file_path}] 初始策略值: {current_strategy_value}, 實際部位: {strat.actual_position}")
                elif current_strategy_value != strat.last_strategy_value:
                    diff = current_strategy_value - strat.actual_position
                    if diff != 0:
                        order_size = abs(diff) * strat.unit
                        direction = "BUY" if diff > 0 else "SELL"
                        self.log(f"[{strat.file_path}] 策略變化: {strat.last_strategy_value} -> {current_strategy_value}，實際部位: {strat.actual_position}，預計下單: {direction} {order_size} 單")
                        result = self.execute_order_in_main_thread(strat.product_code, direction, order_size)
                        self.log(f"[{strat.product_code}] 下單結果: {result}")
                        strat.last_strategy_value = current_strategy_value
                        if result["status"] == "success":
                            strat.actual_position = current_strategy_value
                            self.queue_gui_update(self.update_tree_position, strat.file_path, strat.actual_position)
                            self.queue_gui_update(self.update_position_sum)
                        else:
                            self.log(f"[{strat.product_code}] 注意：下單未完成，實際部位未更新，請檢查 OANDA 或市場狀態")
                    else:
                        self.log(f"[{strat.product_code}] 策略值變動但下單數量為零，跳過下單")
                        strat.last_strategy_value = current_strategy_value
                time.sleep(self.global_freq)
            except Exception as e:
                self.log(f"[{strat.product_code}] 監控錯誤: {e}")
                time.sleep(self.global_freq)

    def save_settings(self):
        data = {
            "api_token": API_TOKEN,
            "account_id": ACCOUNT_ID,
            "account_type": "Live" if BASE_URL == LIVE_URL else "Demo",
            "global_freq": self.global_freq,
            "risk_window": self.risk_window,
            "risk_threshold": self.risk_threshold,
            "is_ordering_enabled": self.is_ordering_enabled,
            "telegram_bot_token": TELEGRAM_BOT_TOKEN,
            "telegram_chat_id": TELEGRAM_CHAT_ID,
            "telegram_enabled": TELEGRAM_ENABLED,
            "strategies": [s.to_dict() for s in self.strategies],
            "strategy_controls": {fid: {"running": ctrl["running"]} for fid, ctrl in self.strategy_controls.items()}
        }
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(data, f, indent=4)
            self.log("設定儲存成功。")
        except Exception as e:
            self.log(f"設定儲存錯誤: {e}")

    def load_settings(self):
        global TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, TELEGRAM_ENABLED, BASE_URL
        if not os.path.isfile(SETTINGS_FILE):
            return
        try:
            with open(SETTINGS_FILE, "r") as f:
                data = json.load(f)
            global API_TOKEN, ACCOUNT_ID
            self.global_freq = data.get("global_freq", 1.0)
            self.risk_window = data.get("risk_window", 10.0)
            self.risk_threshold = data.get("risk_threshold", 5)
            self.is_ordering_enabled = data.get("is_ordering_enabled", False)
            TELEGRAM_BOT_TOKEN = data.get("telegram_bot_token", None)
            TELEGRAM_CHAT_ID = data.get("telegram_chat_id", None)
            TELEGRAM_ENABLED = data.get("telegram_enabled", False)
            self.freq_var.set(str(self.global_freq))
            self.telegram_enabled_var.set(TELEGRAM_ENABLED)
            self.strategies = [Strategy.from_dict(d) for d in data.get("strategies", [])]
            for strat in self.strategies:
                self.tree.insert("", tk.END, iid=strat.file_path, text="OFF", 
                                values=(strat.file_path, strat.product_code, strat.unit, strat.actual_position, ""))
                self.strategy_controls[strat.file_path] = {"running": False, "thread": None}
            saved_controls = data.get("strategy_controls", {})
            for fid, ctrl_data in saved_controls.items():
                if fid in self.strategy_controls and ctrl_data["running"]:
                    strat = next((s for s in self.strategies if s.file_path == fid), None)
                    if strat:
                        self.toggle_monitor(fid)
            self.log("設定載入成功。")
            self.update_position_sum()
            self.update_order_status_light()
        except Exception as e:
            self.log(f"設定載入錯誤: {e}")

    def check_double_click(self, col):
        current_time = time.time()
        last_time = self.last_click_time.get(col, 0)
        self.last_click_time[col] = current_time
        if current_time - last_time < 0.5:
            self.sort_column(col)

    def sort_column(self, col):
        if col not in self.sort_reverse:
            self.sort_reverse[col] = False
        else:
            self.sort_reverse[col] = not self.sort_reverse[col]
        
        reverse = self.sort_reverse[col]
        items = [(self.tree.set(k, col) if col != "#0" else self.tree.item(k, "text"), k) for k in self.tree.get_children('')]
        
        if col in ("unit", "position", "signal"):
            items.sort(key=lambda x: float(x[0]) if x[0] else 0, reverse=reverse)
        elif col == "#0":
            items.sort(key=lambda x: x[0], reverse=reverse)
        else:
            items.sort(key=lambda x: x[0], reverse=reverse)
        
        for index, (val, k) in enumerate(items):
            self.tree.move(k, '', index)

        base_text = {
            "#0": "狀態",
            "file": "策略檔案",
            "product": "商品代號",
            "unit": "單位",
            "position": "實際部位",
            "signal": "策略訊號"
        }

        for c in base_text:
            self.tree.heading(c, text=base_text[c])

        arrow = "↓" if reverse else "↑"
        self.tree.heading(col, text=f"{base_text[col]} {arrow}")

def load_credentials():
    global API_TOKEN, ACCOUNT_ID, BASE_URL
    account_type = "Demo"
    
    try:
        if os.path.isfile(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r") as f:
                data = json.load(f)
                API_TOKEN = data.get("api_token", "")
                ACCOUNT_ID = data.get("account_id", "")
                account_type = data.get("account_type", "Demo")
                BASE_URL = LIVE_URL if account_type == "Live" else DEMO_URL
    except Exception:
        pass

    root = tk.Tk()
    root.withdraw()

    class CredentialsDialog(simpledialog.Dialog):
        def __init__(self, parent, title, token, account, account_type):
            self.token = token
            self.account = account
            self.account_type = account_type
            super().__init__(parent, title)

        def body(self, master):
            tk.Label(master, text="OANDA API Token #(提示:很長、英文數字混合):").pack()
            self.token_var = tk.StringVar(value=self.token)
            tk.Entry(master, textvariable=self.token_var, width=50).pack(padx=5, pady=5)
            tk.Label(master, text="OANDA ID #(提示:數字15碼):").pack()
            self.account_var = tk.StringVar(value=self.account)
            tk.Entry(master, textvariable=self.account_var, width=50).pack(padx=5, pady=5)
            tk.Label(master, text="帳戶類型:").pack()
            self.account_type_var = tk.StringVar(value=self.account_type)
            ttk.Combobox(master, textvariable=self.account_type_var, values=["Live", "Demo"]).pack(padx=5, pady=5)
            return None

        def apply(self):
            self.result = (self.token_var.get(), self.account_var.get(), self.account_type_var.get())

    while True:
        dialog = CredentialsDialog(root, "輸入 OANDA 憑證", API_TOKEN or "", ACCOUNT_ID or "", account_type)
        if not dialog.result:
            root.destroy()
            return False
        
        API_TOKEN, ACCOUNT_ID, account_type = dialog.result
        BASE_URL = LIVE_URL if account_type == "Live" else DEMO_URL
        if check_auth():
            root.destroy()
            return True
        else:
            messagebox.showerror("錯誤", "OANDA API 連線失敗，請檢查憑證並重新輸入", parent=root)

if __name__ == "__main__":
    if load_credentials():
        root = tk.Tk()
        app = OandaOrderApp(root)
        root.mainloop()
