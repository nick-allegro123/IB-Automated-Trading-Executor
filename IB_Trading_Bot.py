import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import threading, time, os, json
from ib_insync import IB, Stock, MarketOrder, Future
import winsound
import requests
import queue
import asyncio
from concurrent.futures import ThreadPoolExecutor
import threading

# IB 連線參數
IB_HOST = None
IB_PORT = None
IB_CLIENT_ID = 1
SETTINGS_FILE = "ib_order_settings.json"

ib = None

# IB 全域變數
ORDER_EXECUTOR = ThreadPoolExecutor(max_workers=32)#IB最多32支子帳戶同時連線
IB_CONNECTION_LOCK = threading.Lock()

# Telegram 全域變數
TELEGRAM_BOT_TOKEN = None
TELEGRAM_CHAT_ID = None
TELEGRAM_ENABLED = False
TELEGRAM_MESSAGE_QUEUE = queue.Queue()

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ib_order_settings.json")

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

def connect_ib(app_instance=None):
    global ib
    if ib is None or not ib.isConnected():
        ib = IB()
        try:
            ib.connect(IB_HOST, IB_PORT, clientId=IB_CLIENT_ID)
            ib.disconnectedEvent += lambda: on_disconnect(app_instance)
            print("IB API 連線成功!!")
            if app_instance:
                app_instance.log("IB API 連線成功")
            return True
        except Exception as e:
            print(f"IB API 連線失敗: {e}")
            if app_instance:
                app_instance.log(f"IB API 連線失敗: {e}")
            return False
    return True

def on_disconnect(app_instance=None):
    if app_instance:
        app_instance.log("IB API 連線斷開，請檢查 TWS/Gateway 狀態")
    global ib
    ib = None

def check_auth(app_instance=None):
    return connect_ib(app_instance)

def place_order(product_code, direction, size, app_instance=None, is_futures=False, expiry=None, exchange="GLOBEX", is_manual=False):
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    
    local_ib = IB()
    try:
        with IB_CONNECTION_LOCK:
            local_ib.connect(IB_HOST, IB_PORT, clientId = threading.get_ident())
        if not is_manual and app_instance and not app_instance.is_ordering_enabled:
            local_ib.disconnect()
            loop.close()
            return {"status": "error", "message": "下單功能已停止，僅允許手動下單"}
        
        try:
            if is_futures:
                currency = "USD"
                if product_code == "MGC":
                    exchange = "COMEX"
                elif product_code == "ZN":
                    exchange = "CBOT"
                
                contract = Future(
                    symbol=product_code,
                    lastTradeDateOrContractMonth=expiry,
                    exchange=exchange,
                    currency=currency
                )
            else:
                contract = Stock(product_code, "SMART", "USD")
            
            units = float(size) if direction == "BUY" else -float(size)
            order = MarketOrder("BUY" if units > 0 else "SELL", abs(units))
            trade = local_ib.placeOrder(contract, order)
            local_ib.sleep(0.1)
            
            status = trade.orderStatus.status
            if status in ("Filled", "PendingSubmit", "PreSubmitted", "Submitted"):
                print(f"[{product_code}] 下單: {direction} {size} 單 - 成功")
                if app_instance:
                    app_instance.log(f"[{product_code}] 下單成功: {direction} {size} 單")
                    if status in ("PendingSubmit", "PreSubmitted", "Submitted"):
                        app_instance.log(f"[{product_code}] 注意：訂單狀態為 {status}，請確認目前是否為交易時段以及保證金是否充足，非交易時段訂單將轉為日內有效單，保證金不足會自動取消訂單，若無上述情形請忽略即可")
                winsound.Beep(1000, 200)
                winsound.Beep(1200, 200)
                local_ib.disconnect()
                loop.close()
                return {"status": "success", "orderId": trade.order.orderId}
            else:
                print(f"[{product_code}] 下單失敗: {status}")
                if app_instance:
                    app_instance.log(f"[{product_code}] 下單失敗: {status}")
                local_ib.disconnect()
                loop.close()
                return {"status": "error", "message": status}
        except Exception as e:
            print(f"[{product_code}] 下單失敗: {e}")
            if app_instance:
                app_instance.log(f"[{product_code}] 下單失敗: {e}")
            local_ib.disconnect()
            loop.close()
            return {"status": "error", "message": str(e)}
    except Exception as e:
        print(f"[{product_code}] IB API 連線失敗: {e}")
        if app_instance:
            app_instance.log(f"[{product_code}] IB API 連線失敗: {e}")
        local_ib.disconnect()
        loop.close()
        return {"status": "error", "message": "未連線到 IB API"}

class Strategy:
    def __init__(self, file_path, product_code, unit, init_position, is_futures=False, expiry=None, exchange="GLOBEX"):
        self.file_path = file_path
        self.product_code = product_code
        self.unit = float(unit)
        self.init_position = int(init_position)
        self.actual_position = int(init_position)
        self.last_strategy_value = None
        self.current_signal = None
        self.change_timestamps = []
        self.monitoring = False
        self.is_futures = is_futures
        self.expiry = expiry
        self.exchange = exchange
        self.unique_id = f"{product_code}_{expiry}" if is_futures else product_code

    def to_dict(self):
        return {
            "file_path": self.file_path,
            "product_code": self.product_code,
            "unit": self.unit,
            "init_position": self.init_position,
            "actual_position": self.actual_position,
            "is_futures": self.is_futures,
            "expiry": self.expiry,
            "exchange": self.exchange
        }

    @classmethod
    def from_dict(cls, d):
        strat = cls(
            d["file_path"], d["product_code"], d["unit"], d["init_position"],
            d.get("is_futures", False), d.get("expiry"), d.get("exchange", "GLOBEX")
        )
        strat.actual_position = d.get("actual_position", strat.init_position)
        return strat

class IBOrderApp:
    def __init__(self, master):
        self.master = master
        self.master.title("IB下單機")
        self.strategies = []
        self.strategy_controls = {}
        self.global_freq = None
        self.sort_reverse = {}
        self.last_click_time = {}
        self.risk_window = None
        self.risk_threshold = None
        self.is_ordering_enabled = False
        self.last_order_toggle_click = 0
        self.order_status_canvas = None
        self.telegram_enabled_var = None
        self.create_widgets()
        self.load_settings()
        telegram_thread = threading.Thread(target=send_telegram_message, daemon=True)
        telegram_thread.start()
        if check_auth(self):
            self.log("成功連接到 IB API")
        else:
            self.log("無法連接到 IB API，請檢查 TWS 是否運行")
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.update_time()
        self.center_window()
        self.start_server_time_heartbeat()

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
        self.save_settings()
        if ib and ib.isConnected():
            ib.disconnect()
        ORDER_EXECUTOR.shutdown(wait=True)
        self.master.destroy()

    def create_widgets(self):
        main_frame = tk.Frame(self.master)
        main_frame.pack(fill=tk.BOTH, expand=True)

        top_frame = tk.Frame(main_frame)
        top_frame.pack(padx=5, pady=5, fill=tk.X)
        tk.Label(top_frame, text="共用監控頻率 (秒/次):").pack(side=tk.LEFT)
        self.freq_var = tk.StringVar(value="1.0")
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

        self.tree = ttk.Treeview(middle_frame, columns=("file", "product", "unit", "position", "signal", "expiry"), show="tree headings", height=10)
        self.tree.heading("#0", text="狀態")
        self.tree.heading("file", text="策略檔案")
        self.tree.heading("product", text="商品代號")
        self.tree.heading("unit", text="單位")
        self.tree.heading("position", text="實際部位")
        self.tree.heading("signal", text="策略訊號")
        self.tree.heading("expiry", text="合約月份")
        self.tree.column("#0", width=80, anchor="center")
        self.tree.column("file", width=200)
        self.tree.column("product", width=120)
        self.tree.column("unit", width=60)
        self.tree.column("position", width=80)
        self.tree.column("signal", width=80)
        self.tree.column("expiry", width=100)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree.bind("<Double-1>", self.on_tree_double_click)
        for col in ("#0", "file", "product", "unit", "position", "signal", "expiry"):
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
        tk.Button(btn_frame, text="設定風控 (所有策略共用)", command=self.set_risk_settings).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="更換月份", command=self.change_expiry).pack(side=tk.LEFT, padx=5)
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
                tk.Label(master, text="[非高頻策略通常50秒內超過5次即為異常，建議設定值為50、5]").pack()
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
        """設置 Telegram Bot Token 和 Chat ID"""
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
                self.save_settings()  # 儲存新的 Telegram 設定
            except Exception as e:
                messagebox.showwarning("錯誤", f"Telegram 設定無效：{e}", parent=self.master)

    def toggle_telegram(self):
        """切換 Telegram 通知啟用狀態"""
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
            key = strat.unique_id
            total = strat.actual_position * strat.unit
            position_totals[key] = position_totals.get(key, 0) + total
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

    def update_tree_expiry(self, fid, value):
        self.tree.set(fid, column="expiry", value=value)

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

    def add_strategy(self):
        while True:
            file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")], parent=self.master)
            if not file_path:
                return
            if any(strat.file_path == file_path for strat in self.strategies):
                messagebox.showwarning("警告", "此策略檔案已存在，請選擇其他檔案", parent=self.master)
                continue
            break
        
        class ProductDialog(simpledialog.Dialog):
            def body(self, master):
                tk.Label(master, text="請輸入商品代號（股票如 AAPL，期貨如 MNQ）：").pack()
                self.product_var = tk.StringVar(value="AAPL")
                self.product_menu = ttk.Combobox(master, textvariable=self.product_var, values=[
                    "AAPL", "MSFT", "NVDA", "QQQ", "SPY", "------",
                    "期貨策略請勾選[期貨]",
                    "文字部分是介紹請勿勾選","------",
                    "NQ","(納斯達克100期貨, CME)",
                    "MNQ","(微型納指, CME)",
                    "ES","(標普500期貨, CME)",
                    "MES","(微型標普, CME)",
                    "MGC","(微型黃金, COMEX)",
                    "ZN","(10年公債期貨, CBOT)",
                    "其餘請上IB官網查詢"])
                
                self.product_menu.pack(padx=5, pady=5)
                self.product_entry = tk.Entry(master, textvariable=self.product_var)

                self.is_futures_var = tk.BooleanVar(value=False)
                self.futures_check = tk.Checkbutton(master, text="期貨策略", variable=self.is_futures_var, 
                                                  command=self.toggle_futures_fields)
                self.futures_check.pack(pady=5)

                self.expiry_frame = tk.Frame(master)
                tk.Label(self.expiry_frame, text="到期日 (YYYYMM，例如 203012):").pack(side=tk.LEFT)
                self.expiry_var = tk.StringVar()
                self.expiry_entry = tk.Entry(self.expiry_frame, textvariable=self.expiry_var, width=10)
                self.expiry_entry.pack(side=tk.LEFT, padx=5)

                self.exchange_frame = tk.Frame(master)
                tk.Label(self.exchange_frame, text="交易所 (例如 CME):").pack(side=tk.LEFT)
                self.exchange_var = tk.StringVar(value="CME")
                self.exchange_entry = tk.Entry(self.exchange_frame, textvariable=self.exchange_var, width=10)
                self.exchange_entry.pack(side=tk.LEFT, padx=5)

                tk.Label(master, text="期貨示例：MNQ (微型納指), 到期日 YYYYMM, 交易所 CME").pack(padx=5, pady=5)
                tk.Label(master, text="#若在非交易時段下單會顯示[交易失敗]").pack()
                tk.Label(master, text="訂單會轉為日內有效單，請特別留意!!").pack()
                tk.Label(master, text="若需要取消請到TWS處理").pack()
                
                self.toggle_futures_fields()
                return self.product_entry

            def toggle_futures_fields(self):
                if self.is_futures_var.get():
                    self.expiry_frame.pack(pady=5)
                    self.exchange_frame.pack(pady=5)
                else:
                    self.expiry_frame.pack_forget()
                    self.exchange_frame.pack_forget()

            def validate(self):
                if self.is_futures_var.get():
                    if not self.expiry_var.get():
                        messagebox.showwarning("錯誤", "期貨策略需輸入到期日 (例如 202506)", parent=self)
                        return False
                    if not self.exchange_var.get():
                        messagebox.showwarning("錯誤", "期貨策略需輸入交易所 (例如 CME)", parent=self)
                        return False
                    expiry = self.expiry_var.get()
                    if not (len(expiry) == 6 and expiry.isdigit()):
                        messagebox.showwarning("錯誤", "到期日格式必須為 YYYYMM (例如 202506)", parent=self)
                        return False
                return True

            def apply(self):
                self.result = (self.product_var.get(), self.is_futures_var.get(), 
                             self.expiry_var.get() if self.is_futures_var.get() else None, 
                             self.exchange_var.get() if self.is_futures_var.get() else "GLOBEX")

        dialog = ProductDialog(self.master, title="選擇商品代號")
        if not dialog.result:
            return
        product_code, is_futures, expiry, exchange = dialog.result

        unit = simpledialog.askstring("輸入", "請輸入下單單位 (股票例如 100，期貨例如 1)：", parent=self.master)
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
        
        strat = Strategy(file_path, product_code, unit, init_position, is_futures, expiry, exchange)
        self.strategies.append(strat)
        self.tree.insert("", tk.END, iid=strat.file_path, text="OFF", 
                        values=(file_path, product_code, unit, strat.actual_position, "", strat.expiry if is_futures else ""))
        self.strategy_controls[strat.file_path] = {"running": False, "thread": None}
        self.log(f"新增策略: {file_path} | {strat.unique_id} {'(期貨)' if is_futures else ''}")
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
            self.log(f"[{strat.unique_id}] 實際部位更新為 {strat.actual_position}")
            self.tree.set(fid, column="position", value=str(strat.actual_position))
            self.update_position_sum()
        except Exception as e:
            messagebox.showwarning("錯誤", f"更新失敗:請輸入正確數值", parent=self.master)

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
                tk.Label(master, text=f"商品: {strat.unique_id}").pack()
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

        result = place_order(strat.product_code, direction, size, self, strat.is_futures, strat.expiry, strat.exchange, is_manual=True)
        if result["status"] == "success":
            if direction == "BUY":
                strat.actual_position += int(size / strat.unit)
            else:
                strat.actual_position -= int(size / strat.unit)
            self.tree.set(fid, column="position", value=str(strat.actual_position))
            self.log(f"[{strat.unique_id}] 手動下單成功: {direction} {size} 單")
            self.update_position_sum()
        else:
            self.log(f"[{strat.unique_id}] 手動下單失敗: {result['message']}")

    def change_expiry(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "請選擇要更換月份的策略", parent=self.master)
            return
        fid = selected[0]
        strat = next((s for s in self.strategies if s.file_path == fid), None)
        if not strat:
            return
        if not strat.is_futures:
            messagebox.showinfo("提示", "此策略不是期貨策略，無法更換月份", parent=self.master)
            return

        new_expiry = simpledialog.askstring("更換合約月份", 
                                           f"目前月份: {strat.expiry}\n請輸入新的到期日 (YYYYMM，例如 202506):", 
                                           parent=self.master)
        if not new_expiry:
            return
        
        if not (len(new_expiry) == 6 and new_expiry.isdigit()):
            messagebox.showwarning("錯誤", "到期日格式必須為 YYYYMM (例如 202506)", parent=self.master)
            return
        
        old_expiry = strat.expiry
        strat.expiry = new_expiry
        strat.unique_id = f"{strat.product_code}_{new_expiry}"
        self.update_tree_expiry(fid, new_expiry)
        self.log(f"[{strat.product_code}] 合約月份從 {old_expiry} 更換為 {new_expiry}")
        self.update_position_sum()

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

    def execute_order_in_thread(self, product_code, direction, size, is_futures, expiry, exchange):
        def run_order():
            result = place_order(product_code, direction, size, self, is_futures, expiry, exchange, is_manual=False)
            return result

        future = ORDER_EXECUTOR.submit(run_order)
        return future.result()

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
                            self.log(f"[{strat.unique_id}] 平倉: {direction} {order_size} 單")
                            result = self.execute_order_in_thread(
                                strat.product_code, direction, order_size,
                                strat.is_futures, strat.expiry, strat.exchange
                            )
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
                        result = self.execute_order_in_thread(
                            strat.product_code, direction, order_size,
                            strat.is_futures, strat.expiry, strat.exchange
                        )
                        self.log(f"[{strat.unique_id}] 下單結果: {result}")
                        strat.last_strategy_value = current_strategy_value
                        if result["status"] == "success":
                            strat.actual_position = current_strategy_value
                            self.queue_gui_update(self.update_tree_position, strat.file_path, strat.actual_position)
                            self.queue_gui_update(self.update_position_sum)
                        else:
                            self.log(f"[{strat.unique_id}] 注意：下單未完成，實際部位未更新，請檢查 TWS 或市場狀態")
                    else:
                        self.log(f"[{strat.unique_id}] 策略值變動但下單數量為零，跳過下單")
                        strat.last_strategy_value = current_strategy_value
                time.sleep(self.global_freq)
            except Exception as e:
                self.log(f"[{strat.unique_id}] 監控錯誤: {e}")
                time.sleep(self.global_freq)

    def save_settings(self):
        data = {
            "host": IB_HOST,
            "port": IB_PORT,
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
        global TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, TELEGRAM_ENABLED
        if not os.path.exists(SETTINGS_FILE):
            self.global_freq = 1.0
            self.risk_window = 50.0
            self.risk_threshold = 5
            self.is_ordering_enabled = False
            self.freq_var.set(str(self.global_freq))
            self.update_order_status_light()
            return
        try:
            with open(SETTINGS_FILE, "r") as f:
                data = json.load(f)
            global IB_HOST, IB_PORT
            self.global_freq = data.get("global_freq", 1.0)
            self.risk_window = data.get("risk_window", 50.0)
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
                                values=(strat.file_path, strat.product_code, strat.unit, strat.actual_position, "", 
                                        strat.expiry if strat.is_futures else ""))
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
            self.global_freq = 1.0
            self.risk_window = 50.0
            self.risk_threshold = 5
            self.is_ordering_enabled = False
            self.freq_var.set(str(self.global_freq))
            self.update_order_status_light()

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
        else:
            items.sort(key=lambda x: x[0], reverse=reverse)
        
        for index, (val, k) in enumerate(items):
            self.tree.move(k, '', index)

        base_text = {"#0": "狀態", "file": "策略檔案", "product": "商品代號", "unit": "單位", "position": "實際部位", "signal": "策略訊號", "expiry": "合約月份"}
        for c in base_text:
            self.tree.heading(c, text=base_text[c])
        arrow = "↓" if reverse else "↑"
        self.tree.heading(col, text=f"{base_text[col]} {arrow}")

    def start_server_time_heartbeat(self):
        def heartbeat():
            try:
                connect_ib(self)
                ib.reqCurrentTime()
            except Exception as e:
                self.log(f"[Heartbeat] 伺服器連線錯誤: {e}")
            self.master.after(3000, heartbeat)
        self.master.after(3000, heartbeat)

def load_credentials():
    global IB_HOST, IB_PORT, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, TELEGRAM_ENABLED
    
    try:
        if os.path.isfile(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r") as f:
                data = json.load(f)
                IB_HOST = data.get("host", "127.0.0.1")
                IB_PORT = data.get("port", 7496)
                TELEGRAM_BOT_TOKEN = data.get("telegram_bot_token", None)
                TELEGRAM_CHAT_ID = data.get("telegram_chat_id", None)
                TELEGRAM_ENABLED = data.get("telegram_enabled", False)
    except Exception:
        pass

    root = tk.Tk()
    root.withdraw()

    class CredentialsDialog(simpledialog.Dialog):
        def __init__(self, parent, title, token, account):
            self.token = token
            self.account = account
            super().__init__(parent, title)

        def body(self, master):
            tk.Label(master, text="請輸入 IB TWS 主機地址 (預設 127.0.0.1):").pack()
            self.host_var = tk.StringVar(value=IB_HOST or "127.0.0.1")
            tk.Entry(master, textvariable=self.host_var, width=50).pack(padx=5, pady=5)
            tk.Label(master, text="請輸入 IB TWS 埠號 (預設 7497 為模擬帳戶， 7496 為真實帳戶):").pack()
            self.port_var = tk.StringVar(value=IB_PORT or "7496")
            tk.Entry(master, textvariable=self.port_var, width=50).pack(padx=5, pady=5)
            return None

        def apply(self):
            self.result = (self.host_var.get(), self.port_var.get())

    while True:
        dialog = CredentialsDialog(root, "輸入 IB TWS 埠號與IP", IB_HOST or "", IB_PORT or "")
        if not dialog.result:
            root.destroy()
            return False
        
        IB_HOST, IB_PORT = dialog.result
        if check_auth():
            root.destroy()
            return True
        else:
            messagebox.showerror("錯誤", "IB TWS 連線失敗，請檢查憑證並重新輸入", parent=root)

if __name__ == "__main__":
    if load_credentials():
        root = tk.Tk()
        app = IBOrderApp(root)
        root.mainloop()
