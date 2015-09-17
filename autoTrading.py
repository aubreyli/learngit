import time
import win32gui
import win32api
import win32con
import tushare as ts
import threading
import tkinter as tk
import tkinter.messagebox

TIME = 0.2

is_start = False
is_stop = False
items_list = []
stock_name = ''
stock_price = ''


def hold_mouse():
    """
    固定鼠标位置
    """
    # 获取屏幕的像素大小
    screen_width_pixel = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    # 固定鼠标位置
    mouse_position = [
        screen_width_pixel - 200, 200, screen_width_pixel - 200, 200]
    win32api.ClipCursor(mouse_position)


def release_mouse():
    mouse_position = [0, 0, 0, 0]
    win32api.ClipCursor(mouse_position)


def TAB_KEY():
    win32api.keybd_event(9, 0, 0, 0)  # TAB
    win32api.keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(TIME)


def F1_KEY():
    win32api.keybd_event(112, 0, 0, 0)
    win32api.keybd_event(112, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(TIME)


def F2_KEY():
    win32api.keybd_event(113, 0, 0, 0)
    win32api.keybd_event(113, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(TIME)


def F6_KEY():
    win32api.keybd_event(117, 0, 0, 0)
    win32api.keybd_event(117, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(TIME)


def ESC_KEY():
    win32api.keybd_event(27, 0, 0, 0)
    win32api.keybd_event(27, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(TIME)


def translate_str_to_keys(str):
    """把字符串传转化成按键码
    """

    clean_str_lst = str.split()

    for char in clean_str_lst[0]:

        if char == '.':
            # 数字符号 .
            win32api.keybd_event(110, 0, 0, 0)  # .
            win32api.keybd_event(110, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(TIME)
            continue

        win32api.keybd_event(ord(char), 0, 0, 0)
        win32api.keybd_event(ord(char), 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(TIME)

    TAB_KEY()


def ready_trade():
    # 获取窗口焦点
    win_name = "网上股票交易系统5.0"
    w1hd = win32gui.FindWindow(0, win_name)
    win32gui.ShowWindow(w1hd, win32con.SW_SHOWMAXIMIZED)
    win32gui.SetForegroundWindow(w1hd)
    hold_mouse()
    ESC_KEY()
    F6_KEY()


def buy(stock_code, stock_number):
    """
    股票买入函数
    :param stock_code:    股票名称
    :param stock_number:  股票数量

    """
    ready_trade()
    F1_KEY()
    translate_str_to_keys(stock_code)
    TAB_KEY()
    TAB_KEY()
    translate_str_to_keys(stock_number)
    translate_str_to_keys("B")
    ESC_KEY()
    release_mouse()


def sell(stock_code, stock_number):
    """
    股票卖出函数
    :param stock_name:    股票名称
    :param stock_number:  股票数量
    """
    ready_trade()
    F2_KEY()
    translate_str_to_keys(stock_code)
    TAB_KEY()
    TAB_KEY()
    translate_str_to_keys(stock_number)
    translate_str_to_keys("S")
    ESC_KEY()
    release_mouse()


def get_stock_data(stock_code):
    """
    :param stock_code: 股票代码
    :return: 字符串
    """
    stock_data = []
    df = ts.get_realtime_quotes(stock_code)
    stock_data.append(df['name'][0])
    stock_data.append(float(df['price'][0]))
    return stock_data


def is_digit(str1):
    if str1 == '':
        return False
    for ch in str1:
        if not ch.isdigit():
            if ch != '.':
                return False
    return True


def monitor():
    global is_start, is_stop, stock_name, stock_price, items_list
    buy_times = 1
    sell_times = 1
    while not is_stop:
        while is_start:
            stock_name, stock_price = get_stock_data(items_list[0])
            if len(items_list) and items_list[0] != '':
                if items_list[3] != '':
                    if buy_times and (stock_price < items_list[1] or stock_price > items_list[2]):
                        sell(items_list[0], items_list[3])
                        buy_times = buy_times - 1
                time.sleep(3)
                if items_list[5] != '':
                    if sell_times and (stock_price > items_list[4]):
                        buy(items_list[0], items_list[5])
                        sell_times = sell_times - 1
            time.sleep(3)
        time.sleep(3)


class StockGui:

    def __init__(self):
        window = tk.Tk()
        window.title("股票交易伴侣")

        label_frame = tk.LabelFrame(window, text="股票")
        label_frame.pack()

        tk.Label(label_frame, text="股票代码", width=10).grid(
            row=1, column=1, sticky=tk.W)
        tk.Label(label_frame, text="股票名称", width=10).grid(
            row=2, column=1, sticky=tk.W)
        tk.Label(label_frame, text="当前价格", width=10).grid(
            row=3, column=1, sticky=tk.W)
        self.stock_code = tk.StringVar()
        self.stock_code_entry = tk.Entry(label_frame, textvariable=self.stock_code, width=10,
                                         justify=tk.RIGHT)
        self.stock_code_entry.grid(row=1, column=2)
        self.stock_name_label = tk.Label(
            label_frame, width=10, bg="yellow", justify=tk.RIGHT)
        self.stock_name_label.grid(row=2, column=2)
        self.stock_price_label = tk.Label(
            label_frame, width=10, bg="yellow", justify=tk.RIGHT)
        self.stock_price_label.grid(row=3, column=2)

        # 卖出
        label_frame1 = tk.LabelFrame(window, text="卖出")
        label_frame1.pack()
        tk.Label(label_frame1, text="止损价格", width=10, fg="blue").grid(
            row=1, column=1, sticky=tk.W)
        tk.Label(label_frame1, text="止盈价格", width=10, fg="blue").grid(
            row=2, column=1, sticky=tk.W)
        tk.Label(label_frame1, text="卖出数量", width=10, fg="blue").grid(
            row=3, column=1, sticky=tk.W)
        self.stop_loss_price = tk.StringVar()
        self.stop_loss_price_entry = tk.Entry(label_frame1, textvariable=self.stop_loss_price, width=10,
                                              justify=tk.RIGHT)
        self.stop_loss_price_entry.grid(row=1, column=2)

        self.stop_profit_price = tk.StringVar()
        self.stop_profit_price_entry = tk.Entry(label_frame1, textvariable=self.stop_profit_price, width=10,
                                                justify=tk.RIGHT)
        self.stop_profit_price_entry.grid(row=2, column=2)

        self.sell_stock_number = tk.StringVar()
        self.sell_stock_number_entry = tk.Entry(label_frame1, textvariable=self.sell_stock_number, width=10,
                                                justify=tk.RIGHT)
        self.sell_stock_number_entry.grid(row=3, column=2)

        # 买入
        ##########################################################
        label_frame2 = tk.LabelFrame(window, text="买入")
        label_frame2.pack()
        tk.Label(label_frame2, text="买入价格", width=10, fg="red").grid(
            row=1, column=1, sticky=tk.W)
        tk.Label(label_frame2, text="买入数量", width=10, fg="red").grid(
            row=2, column=1, sticky=tk.W)

        self.buy_stock_price = tk.StringVar()
        self.buy_stock_price_entry = tk.Entry(label_frame2, textvariable=self.buy_stock_price, width=10,
                                              justify=tk.RIGHT)
        self.buy_stock_price_entry.grid(row=1, column=2)

        self.buy_stock_number = tk.StringVar()
        self.buy_stock_number_entry = tk.Entry(label_frame2, textvariable=self.buy_stock_number, width=10,
                                               justify=tk.RIGHT)
        self.buy_stock_number_entry.grid(row=2, column=2)

        self.start_bt = tk.Button(window, text="启动", command=self.start_stop)
        self.start_bt.pack()

        self.quit_bt = tk.Button(
            window, text="退出", command=lambda: self.close(window))
        self.quit_bt.pack()

        window.protocol(name="WM_DELETE_WINDOW", func=self.callback)

        window.mainloop()

    def start_stop(self):
        global is_start, items_list, stock_name, stock_price

        if is_start is False:
            is_start = True
        else:
            is_start = False

        if is_start:

            items_list = []
            self.get_items()
            items_list = self.items_list

            self.start_bt['text'] = '停止'
            self.disable_widget()

            # while is_start:
            #     self.stock_name_label['text'] = stock_name
            #     self.stock_price_label['text'] = str(stock_price)
            # self.stock_name_label.after(3000)

            # time.sleep(3)
            # self.set_items(stock_name, stock_price)

        else:
            self.enable_widget()
            self.start_bt['text'] = '启动'

    def close(self, window):
        global is_start, is_stop
        is_stop = True
        is_start = False
        window.quit()

    def callback(self):
        pass

    def enable_widget(self):
        self.stock_code_entry['state'] = tk.NORMAL
        self.stop_loss_price_entry['state'] = tk.NORMAL
        self.stop_profit_price_entry['state'] = tk.NORMAL
        self.sell_stock_number_entry['state'] = tk.NORMAL
        self.buy_stock_price_entry['state'] = tk.NORMAL
        self.buy_stock_number_entry['state'] = tk.NORMAL

    def disable_widget(self):
        self.stock_code_entry['state'] = tk.DISABLED
        self.stop_loss_price_entry['state'] = tk.DISABLED
        self.stop_profit_price_entry['state'] = tk.DISABLED
        self.sell_stock_number_entry['state'] = tk.DISABLED
        self.buy_stock_price_entry['state'] = tk.DISABLED
        self.buy_stock_number_entry['state'] = tk.DISABLED

    def get_items(self):

        self.items_list = []

        stock_code = self.stock_code.get().strip()
        if stock_code.isdigit() and len(stock_code) == 6:
            self.items_list.append(stock_code)
        else:
            self.items_list.append('')

        stock_loss_price = self.stop_loss_price.get().strip()
        if is_digit(stock_loss_price):
            self.items_list.append(float(stock_loss_price))
        else:
            self.items_list.append(0)

        stop_profit_price = self.stop_profit_price.get().strip()
        if is_digit(stop_profit_price):
            self.items_list.append(float(stop_profit_price))
        else:
            self.items_list.append(10000)

        sell_stock_number = self.sell_stock_number.get().strip()
        if sell_stock_number.isdigit():
            self.items_list.append(sell_stock_number)
        else:
            self.items_list.append('')

        buy_stock_price = self.buy_stock_price.get().strip()
        if is_digit(buy_stock_price):
            self.items_list.append(float(buy_stock_price))
        else:
            self.items_list.append(10000)

        buy_stock_price = self.buy_stock_number.get().strip()
        if buy_stock_price.isdigit():
            self.items_list.append(buy_stock_price)
        else:
            self.items_list.append('')

    def set_items(self, stock_name, stock_price):
        self.stock_name_label['text'] = stock_name
        self.stock_price_label['text'] = str(stock_price)


t1 = threading.Thread(target=StockGui)
t2 = threading.Thread(target=monitor)

t1.start()
t2.start()
t1.join()
t2.join()
