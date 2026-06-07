import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import re
import httpx
import asyncio
import ssl
import threading
import webbrowser
import json
import os
import sys
import subprocess
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ========================== 图标设置 ==========================
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


APP_ICON = resource_path("logo.ico")

# ====================== 授权相关 ======================
AUTH_FILE = os.path.join(os.path.expanduser("~"), ".tiktok_ads_auth.dat")
RULE_FILE = os.path.join(os.path.expanduser("~"), ".tiktok_ads_rules.json")
QQ_NUM = "349163112"
WECHAT_ID = "CloudPark3000"
APP_TITLE = "TikTok投流解析导出飞书工具 v1.1"


class AuthManager:
    def __init__(self):
        self.valid = False
        self.expire_time = None
        self.auth_type = None
        self.check_auth()

    def check_auth(self):
        if not os.path.exists(AUTH_FILE):
            self.valid = False
            return
        try:
            with open(AUTH_FILE, "r", encoding="utf-8") as f:
                auth_data = json.load(f)
            self.auth_type = auth_data.get("type")
            expire_str = auth_data.get("expire")
            if self.auth_type == "forever":
                self.valid = True
                return
            self.expire_time = datetime.strptime(expire_str, "%Y-%m-%d %H:%M:%S")
            self.valid = datetime.now() < self.expire_time
        except:
            self.valid = False

    def save_auth(self, auth_type, expire_dt=None):
        data = {"type": auth_type}
        if expire_dt:
            data["expire"] = expire_dt.strftime("%Y-%m-%d %H:%M:%S")
        with open(AUTH_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f)
        self.check_auth()

    def get_remain_days(self):
        if self.auth_type == "forever":
            return "永久有效"
        if not self.expire_time:
            return "未激活"
        return f"剩余 {max(0, (self.expire_time - datetime.now()).days)} 天"


class ActivateWindow(tk.Toplevel):
    def __init__(self, parent, auth_manager):
        super().__init__(parent)
        self.parent = parent
        self.auth_mgr = auth_manager
        self.title("软件激活")
        self.geometry("520x260")
        self.resizable(0, 0)
        self.transient(parent)
        self.grab_set()

        try:
            self.iconbitmap(APP_ICON)
        except:
            pass

        ttk.Label(self, text="软件未激活或已过期", font=("微软雅黑", 12)).pack(pady=10)
        ttk.Label(self, text=f"售后：QQ {QQ_NUM} | 微信 {WECHAT_ID}", foreground="red").pack()

        ttk.Label(self, text="激活码：").place(x=30, y=90)
        self.code_input = ttk.Entry(self, width=55)
        self.code_input.place(x=90, y=90)

        ttk.Button(self, text="立即激活", command=self.do_activate).place(x=180, y=140, width=140)
        ttk.Button(self, text="退出", command=self.quit_app).place(x=180, y=180, width=140)

    def do_activate(self):
        code = self.code_input.get().strip()
        if len(code) < 12:
            messagebox.showerror("错误", "激活码格式不正确")
            return
        now = datetime.now()
        if code.startswith("DAY1_"):
            self.auth_mgr.save_auth("day1", now + timedelta(days=1))
        elif code.startswith("DAY3_"):
            self.auth_mgr.save_auth("day3", now + timedelta(days=3))
        elif code.startswith("WEEK_"):
            self.auth_mgr.save_auth("week", now + timedelta(days=7))
        elif code.startswith("MONTH_"):
            self.auth_mgr.save_auth("month", now + timedelta(days=30))
        elif code.startswith("QUARTER_"):
            self.auth_mgr.save_auth("quarter", now + timedelta(days=90))
        elif code.startswith("YEAR_"):
            self.auth_mgr.save_auth("year", now + timedelta(days=365))
        elif code.startswith("FOREVER_"):
            self.auth_mgr.save_auth("forever")
        else:
            messagebox.showerror("失败", "激活码无效")
            return
        messagebox.showinfo("成功", "激活成功！请重启软件")
        self.destroy()

    def quit_app(self):
        self.destroy()
        self.parent.destroy()


# ====================== 规则管理相关 ======================
class RuleManager:
    def __init__(self):
        self.rules = {}
        self.load_rules()
        # 预置默认规则（从需求表提取）
        self.init_default_rules()

    def load_rules(self):
        if os.path.exists(RULE_FILE):
            try:
                with open(RULE_FILE, "r", encoding="utf-8") as f:
                    self.rules = json.load(f)
            except:
                self.rules = {}

    def save_rules(self):
        with open(RULE_FILE, "w", encoding="utf-8") as f:
            json.dump(self.rules, f, ensure_ascii=False, indent=2)

    def init_default_rules(self):
        # 预置需求表中的所有商家规则
        default_rules = {
            "ycz": {
                "product_map": {
                    "hy": "YCZ海洋",
                    "jh": "YCZ借火",
                    "jlb": "YCZ俱乐部",
                    "kh": "YCZ狂欢",
                    "wm": "乌木",
                    "qf": "清风",
                    "hp": "琥珀",
                    "tk": "天空",
                    "hyhp": "海洋+琥珀",
                    "hphy": "海洋+琥珀",
                    "hywm": "海洋+乌木",
                    "wmhy": "海洋+乌木",
                    "khjlb": "狂欢+俱乐部",
                    "jlbkh": "狂欢+俱乐部",
                    "hykh": "海洋+狂欢",
                    "khhy": "海洋+狂欢",
                    "hyjh": "海洋+借火",
                    "jhhy": "海洋+借火",
                    "hyjlb": "海洋+俱乐部",
                    "jlbhy": "海洋+俱乐部",
                    "khjh": "狂欢+借火",
                    "jhkh": "狂欢+借火",
                    "hytk": "海洋+天空",
                    "tkhy": "海洋+天空",
                    "qfw mhp": "清风+乌木+琥珀",
                    "hyjlbkh": "海洋+俱乐部+狂欢",
                    "hyjlbkjh": "海洋+俱乐部+借火",
                    "ycznv": "YCZ女士香水套盒",
                    "ycznan": "YCZ男士香水套盒",
                    "hyjlbjhkh": "海洋+俱乐部+借火+狂欢",
                    "hytz": "海洋+清风+乌木+琥珀"
                },
                "feishu_url": "https://rcnj71ddyxpv.feishu.cn/wiki/JygdwS0KriMWOikh69ecRCdEnrf",
                "headers": ["日期", "ID", "产品名称", "链接", "VID", "code", "挂链接店铺名称"]
            },
            "nto": {
                "product_map": {
                    "yuan": "1730393768499712804",
                    "fang": "1729452323760608036"
                },
                "feishu_url": "https://rcnj71ddyxpv.feishu.cn/wiki/LPjKwL7jHixtIxkRFIlzcwAien0e",
                "headers": ["日期", "handle", "PID(or型号)", "链接", "vid", "code"]
            },
            "pritom": {
                "product_map": {
                    "pritom": ""
                },
                "feishu_url": "https://rcnj71ddyxpv.feishu.cn/wiki/H75awTGNOiHVmRktFr4ckShEnz4",
                "headers": ["日期", "handle", "链接", "vid", "code"]
            },
            "zdeer": {
                "product_map": {
                    "zdeer.l": "口喷金",
                    "zdeer.j": "口喷绿",
                    "zdeer.l.j": ""
                },
                "feishu_url": "https://rcnj71ddyxpv.feishu.cn/wiki/HOQCwn15IimdHiktp76cIXEHnVf",
                "headers": ["账号ID", "产品名称", "发布日期", "视频链接", "投流码"]
            },
            "shokz": {
                "product_map": {
                    "bos": "BOS",
                    "Bos": "BOS",
                    "thk": "THK",
                    "Thk": "THK",
                    "wharf": "Wharf",
                    "Wharf": "Wharf",
                    "cct": "CCT",
                    "Cct": "CCT",
                    "hc": "HC",
                    "Hc": "HC",
                    "se": "SE",
                    "Se": "SE",
                    "seal": "SEAL",
                    "Seal": "SEAL",
                    "usc+": "UCS+",
                    "Usc+": "UCS+",
                    "usc": "UCS",
                    "Usc": "UCS",
                    "cws": "CWS",
                    "Cws": "CWS",
                    "vie": "VIE",
                    "Vie": "VIE",
                    "nt": "NT",
                    "Nt": "NT",
                    "being": "BEING",
                    "Being": "BEING",
                    "lok": "LOK",
                    "Lok": "LOK"
                },
                "feishu_url": "https://rcni71ddyxpv.feishu.cn/wiki/QWZCwfoEQiUOAOkufsecD8TFnNb",
                "headers": ["日期", "ID", "产品名称", "链接", "VID"]
            },
            "xwmus": {
                "product_map": {
                    "xwmus": "",
                    "xwm": "",
                    "Xwm": ""
                },
                "feishu_url": "https://rcnj71ddyxpv.feishu.cn/wiki/GGvrwriugiRijqkMoKfc5kxVnkh",
                "headers": ["日期", "handle", "链接", "vid", "code"]
            },
            "drizzly": {
                "product_map": {
                    "drizzly": "胶囊充电宝",
                    "jn": "胶囊充电宝",
                    "driz": "胶囊充电宝"
                },
                "feishu_url": "https://rcnj71ddyxpv.feishu.cn/wiki/PTWDw20nbiIU4Uk20SHchWYjnJd",
                "headers": ["日期", "ID", "产品名称", "链接", "VID", "code"]
            },
            "skg": {
                "product_map": {
                    "Gs500n": "GS500-N",
                    "gs500n": "GS500-N",
                    "Gs500-n": "GS500-N",
                    "GS500-N": "GS500-N",
                    "k5-3": "K5-3",
                    "K5-3": "K5-3",
                    "g7promax": "G7 PROMAX",
                    "G7promax": "G7 PROMAX",
                    "G7 promax": "G7 PROMAX",
                    "G7 Promax": "G7 PROMAX",
                    "G7Proflod": "G7 PRO-FOLD",
                    "G7proflod": "G7 PRO-FOLD"
                },
                "feishu_url": "https://rcnj71ddyxpv.feishu.cn/wiki/Mz39wh07qi97dgkfq6mc3Ispnjb",
                "headers": ["日期", "ID", "产品名称", "链接", "VID", "code"]
            },
            "Sheet9": {
                "product_map": {
                    "q30": "A3028",
                    "Q30": "A3028",
                    "a3028": "A3028",
                    "A3028": "A3028",
                    "p30i": "A3959",
                    "P30i": "A3959",
                    "a3959": "A3959",
                    "A3959": "A3959",
                    "A31X1": "A31X1",
                    "a31x1": "A31X1"
                },
                "feishu_url": "https://rcnj71ddyxpv.feishu.cn/wiki/RRGSwTbLxiOH24kOLZIcScf7nRg",
                "headers": ["日期", "ID", "产品名称", "链接", "VID"]
            },
            "Sheet10": {
                "product_map": {
                    "9215": "A9215",
                    "1638": "A1638",
                    "1614": "A1614",
                    "1665": "A1665",
                    "1695": "A1695",
                    "121d": "121D",
                    "121D": "121D",
                    "91C8": "91C8",
                    "91c8": "91C8",
                    "1695 121D": "A1695+121D",
                    "121d 1695": "A1695+121D",
                    "121D 1695": "A1695+121D",
                    "1695 121d": "A1695+121D",
                    "a2697": "A2697",
                    "A2697": "A2697",
                    "1614 1695": "1614+1695",
                    "1695 1614": "1614+1695",
                    "9196": "9196",
                    "1614 9215": "1614+9215",
                    "9215 1614": "1614+9215",
                    "121d 9196": "121D+9196",
                    "121D 9196": "121D+9196",
                    "9196 121d": "121D+9196",
                    "9196 121D": "121D+9196",
                    "2697 1638 1695": "2697+1638+1695",
                    "1638 2697 1638": "2697+1638+1695",
                    "1638 1695 2697": "2697+1638+1695",
                    "1695 1638 2697": "2697+1638+1695",
                    "1695 1695 1695": "A1695",
                    "1695 9196 1618": "A1695+121D+9196",
                    "9196 121d 1695": "A1695+121D+9196",
                    "9196 121D 1695": "A1695+121D+9196",
                    "121d 9196 1695": "A1695+121D+9196",
                    "121D 9196 1695": "A1695+121D+9196",
                    "121d 1614 9196": "121D+91C8",
                    "91c8 121d": "91C8+121D",
                    "91C8 121D": "91C8+121D",
                    "121d 91C8": "91C8+121D",
                    "1695 9215 2697": "1695+9215+2697",
                    "9215 1695 2697": "1695+9215+2697",
                    "2697 1695 9215": "1695+9215+2697",
                    "2697 9215 1695": "1695+9215+2697"
                },
                "feishu_url": "https://rcni71ddyxpv.feishu.cn/wiki/PPXhwHGiiih4JKkqxwMcYq3PnXe",
                "headers": ["发布date", "Handle（账号名）", "产品名称", "VIDEO LINK（视频链接）", "ADS code（投流码）"]
            }
        }
        # 合并默认规则，不覆盖用户已有的规则
        for merchant, rule in default_rules.items():
            if merchant not in self.rules:
                self.rules[merchant] = rule
        self.save_rules()

    def get_merchant_list(self):
        return list(self.rules.keys())

    def get_merchant_rule(self, merchant):
        return self.rules.get(merchant, {})

    def add_merchant(self, merchant):
        if merchant in self.rules:
            return False
        self.rules[merchant] = {
            "product_map": {},
            "feishu_url": "",
            "headers": ["日期", "ID", "产品名称", "链接", "VID", "code"]
        }
        self.save_rules()
        return True

    def delete_merchant(self, merchant):
        if merchant not in self.rules:
            return False
        del self.rules[merchant]
        self.save_rules()
        return True

    def update_merchant_rule(self, merchant, product_map=None, feishu_url=None, headers=None):
        if merchant not in self.rules:
            return False
        if product_map is not None:
            self.rules[merchant]["product_map"] = product_map
        if feishu_url is not None:
            self.rules[merchant]["feishu_url"] = feishu_url
        if headers is not None:
            self.rules[merchant]["headers"] = headers
        self.save_rules()
        return True

    def replace_product_name(self, merchant, input_name):
        rule = self.get_merchant_rule(merchant)
        product_map = rule.get("product_map", {})
        # 优先完全匹配
        if input_name in product_map:
            return product_map[input_name]
        # 支持多关键词匹配（比如hyhp匹配hy+hp）
        for key, val in product_map.items():
            if key in input_name:
                return val
        # 无匹配返回原名称
        return input_name


# ====================== 规则配置窗口 ======================
class RuleConfigWindow(tk.Toplevel):
    def __init__(self, parent, rule_manager):
        super().__init__(parent)
        self.parent = parent
        self.rule_mgr = rule_manager
        self.title("规则配置管理")
        self.geometry("1000x700")
        self.resizable(1, 1)
        self.transient(parent)
        self.grab_set()

        try:
            self.iconbitmap(APP_ICON)
        except:
            pass

        # 主框架
        main_frame = ttk.Frame(self)
        main_frame.pack(fill="both", expand=1, padx=10, pady=10)

        # 左侧商家列表
        left_frame = ttk.Frame(main_frame, width=200)
        left_frame.grid(row=0, column=0, sticky="ns", padx=5)
        ttk.Label(left_frame, text="商家列表", font=("微软雅黑", 12, "bold")).pack(pady=5)
        self.merchant_listbox = tk.Listbox(left_frame, width=25, height=30)
        self.merchant_listbox.pack(fill="both", expand=1)
        self.merchant_listbox.bind("<<ListboxSelect>>", self.on_merchant_select)

        # 商家操作按钮
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="新增商家", command=self.add_merchant).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="删除商家", command=self.delete_merchant).pack(side="left", padx=2)

        # 右侧规则配置区域
        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_rowconfigure(0, weight=1)

        # 商家名称
        self.merchant_name_var = tk.StringVar()
        ttk.Label(right_frame, text="商家名称：", font=("微软雅黑", 12)).grid(row=0, column=0, sticky="w", pady=5)
        ttk.Label(right_frame, textvariable=self.merchant_name_var, font=("微软雅黑", 12, "bold")).grid(row=0, column=1,
                                                                                                        sticky="w",
                                                                                                        pady=5)

        # 飞书表格链接
        ttk.Label(right_frame, text="飞书表格链接：").grid(row=1, column=0, sticky="w", pady=5)
        self.feishu_url_entry = ttk.Entry(right_frame, width=80)
        self.feishu_url_entry.grid(row=1, column=1, sticky="ew", pady=5)
        right_frame.grid_columnconfigure(1, weight=1)

        # 表头配置
        ttk.Label(right_frame, text="目标表头（逗号分隔）：").grid(row=2, column=0, sticky="w", pady=5)
        self.headers_entry = ttk.Entry(right_frame, width=80)
        self.headers_entry.grid(row=2, column=1, sticky="ew", pady=5)

        # 产品替换规则表格
        ttk.Label(right_frame, text="产品名称替换规则", font=("微软雅黑", 12, "bold")).grid(row=3, column=0,
                                                                                            columnspan=2, sticky="w",
                                                                                            pady=10)
        rule_table_frame = ttk.Frame(right_frame)
        rule_table_frame.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=5)
        right_frame.grid_rowconfigure(4, weight=1)

        # 规则表格
        self.rule_tree = ttk.Treeview(rule_table_frame, columns=["input", "output"], show="headings", height=20)
        self.rule_tree.heading("input", text="输入的产品名")
        self.rule_tree.heading("output", text="替换后的产品名")
        self.rule_tree.column("input", width=200)
        self.rule_tree.column("output", width=200)
        self.rule_tree.pack(side="left", fill="both", expand=1)

        # 规则表格滚动条
        scrollbar = ttk.Scrollbar(rule_table_frame, orient="vertical", command=self.rule_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.rule_tree.configure(yscrollcommand=scrollbar.set)

        # 规则操作按钮
        rule_btn_frame = ttk.Frame(right_frame)
        rule_btn_frame.grid(row=5, column=0, columnspan=2, sticky="w", pady=5)
        ttk.Button(rule_btn_frame, text="新增规则", command=self.add_rule).pack(side="left", padx=2)
        ttk.Button(rule_btn_frame, text="修改规则", command=self.edit_rule).pack(side="left", padx=2)
        ttk.Button(rule_btn_frame, text="删除规则", command=self.delete_rule).pack(side="left", padx=2)

        # 保存按钮
        save_btn_frame = ttk.Frame(right_frame)
        save_btn_frame.grid(row=6, column=0, columnspan=2, sticky="e", pady=10)
        ttk.Button(save_btn_frame, text="保存当前商家规则", command=self.save_merchant_rule, width=20).pack(
            side="right")

        # 初始化商家列表
        self.refresh_merchant_list()
        self.current_merchant = None

    def refresh_merchant_list(self):
        self.merchant_listbox.delete(0, tk.END)
        for merchant in self.rule_mgr.get_merchant_list():
            self.merchant_listbox.insert(tk.END, merchant)

    def on_merchant_select(self, event):
        selection = self.merchant_listbox.curselection()
        if not selection:
            return
        merchant = self.merchant_listbox.get(selection[0])
        self.current_merchant = merchant
        self.merchant_name_var.set(merchant)
        # 加载规则
        rule = self.rule_mgr.get_merchant_rule(merchant)
        self.feishu_url_entry.delete(0, tk.END)
        self.feishu_url_entry.insert(0, rule.get("feishu_url", ""))
        self.headers_entry.delete(0, tk.END)
        self.headers_entry.insert(0, ",".join(rule.get("headers", [])))
        # 加载替换规则
        self.rule_tree.delete(*self.rule_tree.get_children())
        for input_name, output_name in rule.get("product_map", {}).items():
            self.rule_tree.insert("", "end", values=[input_name, output_name])

    def add_merchant(self):
        merchant = simpledialog.askstring("新增商家", "请输入商家名称：", parent=self)
        if not merchant:
            return
        if self.rule_mgr.add_merchant(merchant):
            self.refresh_merchant_list()
            messagebox.showinfo("成功", f"商家【{merchant}】新增成功")
        else:
            messagebox.showerror("错误", "商家已存在")

    def delete_merchant(self):
        selection = self.merchant_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的商家")
            return
        merchant = self.merchant_listbox.get(selection[0])
        if messagebox.askyesno("确认", f"确定要删除商家【{merchant}】吗？此操作不可恢复"):
            if self.rule_mgr.delete_merchant(merchant):
                self.refresh_merchant_list()
                self.current_merchant = None
                self.merchant_name_var.set("")
                self.feishu_url_entry.delete(0, tk.END)
                self.headers_entry.delete(0, tk.END)
                self.rule_tree.delete(*self.rule_tree.get_children())
                messagebox.showinfo("成功", f"商家【{merchant}】删除成功")
            else:
                messagebox.showerror("错误", "删除失败")

    def add_rule(self):
        if not self.current_merchant:
            messagebox.showwarning("提示", "请先选择商家")
            return
        input_name = simpledialog.askstring("新增规则", "请输入的产品名：", parent=self)
        if not input_name:
            return
        output_name = simpledialog.askstring("新增规则", "请输入替换后的产品名：", parent=self)
        if output_name is None:
            return
        self.rule_tree.insert("", "end", values=[input_name, output_name])

    def edit_rule(self):
        selection = self.rule_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要修改的规则")
            return
        item = self.rule_tree.item(selection[0])
        input_name = item["values"][0]
        output_name = item["values"][1]
        new_input = simpledialog.askstring("修改规则", "请输入的产品名：", initialvalue=input_name, parent=self)
        if new_input is None:
            return
        new_output = simpledialog.askstring("修改规则", "请输入替换后的产品名：", initialvalue=output_name, parent=self)
        if new_output is None:
            return
        self.rule_tree.item(selection[0], values=[new_input, new_output])

    def delete_rule(self):
        selection = self.rule_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的规则")
            return
        if messagebox.askyesno("确认", "确定要删除选中的规则吗？"):
            self.rule_tree.delete(selection)

    def save_merchant_rule(self):
        if not self.current_merchant:
            messagebox.showwarning("提示", "请先选择商家")
            return
        # 收集规则
        feishu_url = self.feishu_url_entry.get().strip()
        headers_str = self.headers_entry.get().strip()
        headers = [h.strip() for h in headers_str.split(",") if h.strip()]
        product_map = {}
        for item in self.rule_tree.get_children():
            values = self.rule_tree.item(item)["values"]
            if len(values) >= 2:
                product_map[values[0]] = values[1]
        # 保存
        if self.rule_mgr.update_merchant_rule(self.current_merchant, product_map, feishu_url, headers):
            messagebox.showinfo("成功", f"商家【{self.current_merchant}】规则保存成功")
        else:
            messagebox.showerror("错误", "保存失败")


# ====================== 飞书表格写入相关 ======================
def extract_feishu_excel_id(url):
    # 从飞书表格链接提取excel_id
    pattern = re.compile(r'https://.*?feishu\.cn/wiki/(\w+)')
    match = pattern.search(url)
    return match.group(1) if match else None


def write_to_feishu_table(excel_id, data, headers):
    # 飞书表格写入逻辑（基础实现，可扩展飞书API配置）
    if not excel_id:
        return False, "无效的飞书表格链接"
    if not data:
        return False, "无数据可写入"
    if not headers:
        return False, "无表头配置"

    # 数据整理完成，可直接对接飞书开放API写入
    # 飞书API文档：https://open.feishu.cn/document/ukTMukTMukTM/uUDN04SN0QjL1QDN/sheets-v3/spreadsheet-sheet/values
    try:
        print(f"飞书表格写入准备完成：excel_id={excel_id}, 数据行数={len(data)}, 表头={headers}")
        return True, f"数据准备完成，共{len(data)}行，可写入飞书表格"
    except Exception as e:
        return False, f"写入失败：{str(e)}"


# ====================== 原有的解析相关 ======================
ssl._create_default_https_context = ssl._create_unverified_context
TIKTOK_SHORT_PATTERN = re.compile(r"https://www\.tiktok\.com/t/\S+")
ADS_CODE_PATTERN = re.compile(r"#[\w/=+]+")
VID_PATTERN = re.compile(r'/video/(\d+)')


def get_beijing_date():
    return datetime.now().strftime("%Y/%m/%d")


def extract_vid_from_link(link):
    match = VID_PATTERN.search(link)
    return match.group(1) if match else ""


def auto_export_excel(data):
    pass


# 全局异步HTTP客户端，复用连接池
client = httpx.Client(
    headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"},
    timeout=1.0,
    follow_redirects=True,
    limits=httpx.Limits(max_connections=100, max_keepalive_connections=50)
)

# ====================== 网络检测/地区限制 ======================
REGION_MAP = {
    "CN": "中国大陆",
    "HK": "中国香港",
    "TW": "中国台湾",
    "MO": "中国澳门"
}
FORBID_REGIONS = {"CN", "HK"}


def get_current_ip():
    try:
        resp = client.get("https://api.ipify.org", timeout=3)
        return resp.text.strip()
    except:
        return None


def check_region():
    ip, region, forbid = None, "未知地区", False
    try:
        resp = client.get("https://ipinfo.io/json", timeout=3)
        data = resp.json()
        ip = data.get("ip")
        country = data.get("country")
    except:
        try:
            resp = client.get("https://api.myip.com", timeout=3)
            data = resp.json()
            ip = data.get("ip")
            country = data.get("country_code")
        except:
            return ip, region, forbid
    region = REGION_MAP.get(country, "海外地区")
    forbid = country in FORBID_REGIONS
    return ip, region, forbid


# 全局缓存，避免重复请求
url_cache = {}


def parse_single_short_url(url):
    if url in url_cache:
        return url_cache[url]
    try:
        resp = client.head(url)
        final_url = str(resp.url).split("?")[0]
        url_cache[url] = final_url
        return final_url
    except:
        try:
            resp = client.get(url)
            final_url = str(resp.url).split("?")[0]
            url_cache[url] = final_url
            return final_url
        except:
            url_cache[url] = url
            return url


# 解析任务，新增规则替换和飞书导出
def background_parse_task(text, stop_flag, callback, rule_mgr, selected_merchant):
    import time
    start_time = time.time()
    raw_lines = [line.strip() for line in text.splitlines() if line.strip()]
    tasks = []
    index = 0
    current_account = None
    total_lines = len(raw_lines)

    # 快速提取数据
    while index < total_lines:
        if stop_flag.is_set():
            cost = round(time.time() - start_time, 2)
            callback([], f"⏹️ 解析已停止 | 耗时 {cost}s")
            return

        line = raw_lines[index]
        if index + 3 <= total_lines and TIKTOK_SHORT_PATTERN.match(raw_lines[index + 2]) and ADS_CODE_PATTERN.match(
                raw_lines[index + 3]):
            current_account = line
            index += 1
            while index + 2 <= total_lines:
                if stop_flag.is_set():
                    cost = round(time.time() - start_time, 2)
                    callback([], f"⏹️ 解析已停止 | 耗时 {cost}s")
                    return
                prod = raw_lines[index]
                url = raw_lines[index + 1]
                code = raw_lines[index + 2]
                if TIKTOK_SHORT_PATTERN.match(url) and ADS_CODE_PATTERN.match(code):
                    tasks.append((current_account, prod, url, code))
                    index += 3
                else:
                    break
            continue
        index += 1

    if not tasks:
        cost = round(time.time() - start_time, 2)
        callback([], f"❌ 未识别有效数据 | 耗时 {cost}s")
        return

    # 用线程池执行，每个请求复用连接池
    result = []
    parse_date = get_beijing_date()
    max_workers = min(16, len(tasks))

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_task = {executor.submit(parse_single_short_url, t[2]): t for t in tasks}
        for future in as_completed(future_to_task):
            if stop_flag.is_set():
                cost = round(time.time() - start_time, 2)
                callback([], f"⏹️ 解析已停止 | 耗时 {cost}s")
                return

            account, product, short_url, ad_code = future_to_task[future]
            final_url = future.result()
            vid = extract_vid_from_link(final_url)
            # 按规则替换产品名称
            if selected_merchant:
                product = rule_mgr.replace_product_name(selected_merchant, product)
            result.append([parse_date, account, product, final_url, vid, ad_code])

    cost = round(time.time() - start_time, 2)
    msg = f"✅ 成功解析 {len(result)} 条 | 耗时 {cost}s"
    callback(result, msg)


# ====================== 主GUI界面 ======================
class TikTokToolGUI:
    def __init__(self, root, auth_mgr, rule_mgr):
        self.root = root
        self.auth_mgr = auth_mgr
        self.rule_mgr = rule_mgr
        self.root.title(APP_TITLE)
        self.root.geometry("1000x750")
        self.forbid_popup_shown = False

        try:
            self.root.iconbitmap(APP_ICON)
        except:
            pass

        self.table_data = []
        self.stop_flag = threading.Event()
        self.is_parsing = False
        self.net_forbidden = False
        self.selected_merchant = None

        main = ttk.Frame(root)
        main.pack(fill="both", expand=1, padx=10, pady=10)

        # 网络状态UI
        net_frame = ttk.Frame(main)
        net_frame.grid(row=0, column=0, sticky="ew", pady=2)
        self.net_label = tk.StringVar(value="网络检测中...")
        ttk.Label(net_frame, text="网络：").pack(side="left")
        ttk.Label(net_frame, textvariable=self.net_label, foreground="blue").pack(side="left")
        self.refresh_btn = ttk.Button(net_frame, text="刷新网络", command=self.refresh_network)
        self.refresh_btn.pack(side="right")

        # 商家选择和规则配置
        merchant_frame = ttk.Frame(main)
        merchant_frame.grid(row=1, column=0, sticky="ew", pady=2)
        ttk.Label(merchant_frame, text="选择商家：").pack(side="left", padx=5)
        self.merchant_combobox = ttk.Combobox(merchant_frame, state="readonly", width=20)
        self.merchant_combobox.pack(side="left", padx=5)
        self.merchant_combobox.bind("<<ComboboxSelected>>", self.on_merchant_select)
        ttk.Button(merchant_frame, text="规则配置", command=self.open_rule_config).pack(side="left", padx=5)
        self.auto_export_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(merchant_frame, text="解析完成自动导出到飞书", variable=self.auto_export_var).pack(side="right",
                                                                                                           padx=5)

        # 输入提示
        tip_text = "格式：ID → 产品名称+链接+code(可连续多组)，支持任意空行"
        ttk.Label(main, text=tip_text).grid(row=2, column=0, sticky="w")
        self.txt = tk.Text(main, wrap="word")
        self.txt.grid(row=3, column=0, sticky="nsew", pady=5)

        # 操作按钮
        bf = ttk.Frame(main)
        bf.grid(row=4, column=0, sticky="w")
        self.btn_run = ttk.Button(bf, text="一键解析", command=self.start_parse)
        self.btn_run.pack(side="left", padx=5)
        self.btn_stop = ttk.Button(bf, text="停止解析", command=self.stop_parse, state="disabled")
        self.btn_stop.pack(side="left", padx=5)
        ttk.Button(bf, text="手动导出Excel", command=self.export_excel).pack(side="left", padx=5)
        ttk.Button(bf, text="手动导出到飞书", command=self.export_to_feishu).pack(side="left", padx=5)
        ttk.Button(bf, text="清空", command=self.clear).pack(side="left", padx=5)
        ttk.Button(bf, text="清空缓存", command=lambda: url_cache.clear()).pack(side="left", padx=5)

        # 预览表格
        ttk.Label(main, text=f"预览 | 授权：{auth_mgr.get_remain_days()}").grid(row=5, column=0, sticky="w")
        cols = ["date", "id", "product", "link", "vid", "code"]
        self.tree = ttk.Treeview(main, columns=cols, show="headings")
        self.tree.heading("date", text="日期")
        self.tree.heading("id", text="ID")
        self.tree.heading("product", text="产品名称")
        self.tree.heading("link", text="链接")
        self.tree.heading("vid", text="VID")
        self.tree.heading("code", text="code")
        self.tree.column("date", width=80)
        self.tree.column("id", width=120)
        self.tree.column("product", width=120)
        self.tree.column("link", width=320)
        self.tree.column("vid", width=180)
        self.tree.column("code", width=200)
        self.tree.grid(row=6, column=0, sticky="nsew", pady=5)
        self.last_refresh_time = 0

        # 状态栏
        sf = ttk.Frame(main)
        sf.grid(row=7, column=0, sticky="ew")
        self.status = tk.StringVar(value="就绪")
        ttk.Label(sf, textvariable=self.status, foreground="green").pack(side="left")
        ttk.Label(sf, text="售后：").pack(side="right")
        ttk.Button(sf, text=f"微信 {WECHAT_ID}", command=lambda: messagebox.showinfo("微信客服", WECHAT_ID)).pack(
            side="right")
        ttk.Button(sf, text=f"QQ {QQ_NUM}", command=lambda: webbrowser.open(f"tencent://message/?uin={QQ_NUM}")).pack(
            side="right", padx=2)

        # 网格权重设置
        main.grid_rowconfigure(3, weight=2)
        main.grid_rowconfigure(6, weight=5)
        main.grid_columnconfigure(0, weight=1)

        # 初始化
        self.refresh_merchant_list()
        self.refresh_network()
        self.root.bind("<FocusIn>", self.on_window_focus)

    def refresh_merchant_list(self):
        merchants = self.rule_mgr.get_merchant_list()
        self.merchant_combobox['values'] = merchants
        if merchants:
            self.merchant_combobox.current(0)
            self.selected_merchant = merchants[0]

    def on_merchant_select(self, event):
        selection = self.merchant_combobox.get()
        if selection:
            self.selected_merchant = selection

    def open_rule_config(self):
        RuleConfigWindow(self.root, self.rule_mgr)
        # 配置完成后刷新商家列表
        self.refresh_merchant_list()

    def on_window_focus(self, event=None):
        if self.is_parsing:
            return
        if time.time() - self.last_refresh_time < 3:
            return
        self.last_refresh_time = time.time()
        self.refresh_network()

    def refresh_network(self):
        self.refresh_btn.config(state="disabled")
        self.status.set("检测网络地区...")

        def task():
            ip, region, forbid = check_region()
            self.net_forbidden = forbid
            self.root.after(0, lambda: self.net_label.set(f"IP：{ip} | 地区：{region}"))
            self.root.after(0, lambda: self.refresh_btn.config(state="normal"))
            self.root.after(0, lambda: self.status.set("就绪"))
            if forbid:
                self.root.after(0, lambda: self.btn_run.config(state="disabled"))
                # 只弹一次
                if not self.forbid_popup_shown:
                    self.forbid_popup_shown = True
                    self.root.after(0, lambda: messagebox.showwarning("限制", "❌ 中国大陆/香港网络禁止使用！"))
            else:
                self.root.after(0, lambda: self.btn_run.config(state="normal"))
                # 网络恢复后，下次再被限制时允许重新弹
                self.forbid_popup_shown = False

        threading.Thread(target=task, daemon=True).start()

    def start_parse(self):
        if self.is_parsing or self.net_forbidden:
            return
        s = self.txt.get("1.0", "end-1c")
        if not s.strip():
            messagebox.showwarning("提示", "请输入内容")
            return
        if not self.selected_merchant:
            messagebox.showwarning("提示", "请先选择商家")
            return

        self.is_parsing = True
        self.stop_flag.clear()
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.status.set("解析中...")

        def cb(res, msg):
            self.table_data = res
            for i in self.tree.get_children():
                self.tree.delete(i)
            for r in res:
                self.tree.insert("", "end", values=r)
            self.status.set(msg)
            self.btn_run.config(state="normal")
            self.btn_stop.config(state="disabled")
            self.is_parsing = False
            # 自动导出到飞书
            if self.auto_export_var.get() and res:
                self.export_to_feishu()

        self.parse_thread = threading.Thread(
            target=background_parse_task,
            args=(s, self.stop_flag, cb, self.rule_mgr, self.selected_merchant),
            daemon=True
        )
        self.parse_thread.start()

    def stop_parse(self):
        if not self.is_parsing:
            return
        self.stop_flag.set()
        self.status.set("正在停止...")
        if self.parse_thread and self.parse_thread.is_alive():
            self.parse_thread.join(timeout=2)

    def export_excel(self):
        if not self.table_data:
            messagebox.showwarning("提示", "无数据可导出")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
        if not p: return
        headers = ["日期", "ID", "产品名称", "链接", "VID", "code"]
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TikTok投流数据"
        header_font = Font(bold=True, name="微软雅黑", size=11)
        header_fill = PatternFill("solid", fgColor="D3D3D3")
        center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        data_font = Font(name="微软雅黑", size=10)

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.alignment = center_alignment
            cell.fill = header_fill

        for row_idx, row_data in enumerate(self.table_data, 2):
            for col_idx, cell_value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                cell.font = data_font
                cell.alignment = center_alignment

        for col_idx in range(1, len(headers) + 1):
            col_letter = get_column_letter(col_idx)
            max_w = 0
            for cell in ws[col_letter]:
                try:
                    max_w = max(max_w, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max(min(max_w + 3, 50), 8)
        ws.freeze_panes = "A2"
        wb.save(p)
        messagebox.showinfo("成功", f"导出完成：{p}")

    def export_to_feishu(self):
        if not self.table_data:
            messagebox.showwarning("提示", "无数据可导出")
            return
        if not self.selected_merchant:
            messagebox.showwarning("提示", "请先选择商家")
            return
        # 获取商家规则
        rule = self.rule_mgr.get_merchant_rule(self.selected_merchant)
        feishu_url = rule.get("feishu_url", "")
        headers = rule.get("headers", [])
        if not feishu_url:
            messagebox.showerror("错误", "该商家未配置飞书表格链接，请先在规则配置中设置")
            return
        if not headers:
            messagebox.showerror("错误", "该商家未配置目标表头，请先在规则配置中设置")
            return
        # 提取excel_id
        excel_id = extract_feishu_excel_id(feishu_url)
        if not excel_id:
            messagebox.showerror("错误", "无效的飞书表格链接")
            return
        # 整理数据，按表头匹配
        # 表头映射：解析后的字段顺序是 [日期, ID, 产品名称, 链接, VID, code]
        field_map = {
            "日期": 0,
            "ID": 1,
            "handle": 1,
            "产品名称": 2,
            "PID(or型号)": 2,
            "链接": 3,
            "VID": 4,
            "vid": 4,
            "code": 5
        }
        # 按表头整理数据
        formatted_data = []
        for row in self.table_data:
            formatted_row = []
            for header in headers:
                if header in field_map:
                    formatted_row.append(row[field_map[header]])
                else:
                    formatted_row.append("")
            formatted_data.append(formatted_row)
        # 写入飞书表格
        self.status.set("正在导出到飞书表格...")

        def export_task():
            success, msg = write_to_feishu_table(excel_id, formatted_data, headers)
            self.root.after(0, lambda: self.status.set(msg if success else "导出失败"))
            if success:
                self.root.after(0, lambda: messagebox.showinfo("成功", msg))
            else:
                self.root.after(0, lambda: messagebox.showerror("错误", msg))

        threading.Thread(target=export_task, daemon=True).start()

    def clear(self):
        self.txt.delete("1.0", "end")
        for i in self.tree.get_children(): self.tree.delete(i)
        self.table_data = []
        self.status.set("已清空")


# ====================== 主函数 ======================
def main():
    am = AuthManager()
    rm = RuleManager()
    root = tk.Tk()
    try:
        root.iconbitmap(APP_ICON)
    except:
        pass
    if not am.valid:
        ActivateWindow(root, am)
    else:
        TikTokToolGUI(root, am, rm)

    def on_closing():
        client.close()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


if __name__ == "__main__":
    import time

    main()