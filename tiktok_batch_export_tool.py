import tkinter as tk
from tkinter import ttk, messagebox, filedialog
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
QQ_NUM = "349163112"
WECHAT_ID = "CloudPark3000"
APP_TITLE = "TikTok投流解析导出飞书工具 v1.0"  # 1. 加版本号


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


ssl._create_default_https_context = ssl._create_unverified_context
TIKTOK_SHORT_PATTERN = re.compile(r"https://www\.tiktok\.com/t/\S+")
ADS_CODE_PATTERN = re.compile(r"#[\w/=+]+")
VID_PATTERN = re.compile(r'/video/(\d+)')


# 2. 日期改为北京时间
def get_beijing_date():
    return datetime.now().strftime("%Y/%m/%d")

def extract_vid_from_link(link):
    match = VID_PATTERN.search(link)
    return match.group(1) if match else ""


# 移除自动导出Excel，只保留手动导出
def auto_export_excel(data):
    pass


# 全局异步HTTP客户端，复用连接池
client = httpx.Client(
    headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"},
    timeout=1.0,  # 1秒超时，快速失败
    follow_redirects=True,
    limits=httpx.Limits(max_connections=100, max_keepalive_connections=50)
)

# ====================== 网络检测/地区限制（修复版） ======================
REGION_MAP = {
    "CN": "中国大陆",
    "HK": "中国香港",
    "TW": "中国台湾",
    "MO": "中国澳门"
}
FORBID_REGIONS = {"CN", "HK"}

def get_current_ip():
    try:
        # 改用httpx请求，和全局客户端复用配置
        resp = client.get("https://api.ipify.org", timeout=3)
        return resp.text.strip()
    except:
        return None

def check_region():
    ip, region, forbid = None, "未知地区", False
    try:
        # 主接口
        resp = client.get("https://ipinfo.io/json", timeout=3)
        data = resp.json()
        ip = data.get("ip")
        country = data.get("country")
    except:
        try:
            # 备用接口
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
        resp = client.head(url)  # 用HEAD请求，比GET快一倍
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

# 替换background_parse_task为异步并发版本
def background_parse_task(text, stop_flag, callback):
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
        if index + 3 <= total_lines and TIKTOK_SHORT_PATTERN.match(raw_lines[index+2]) and ADS_CODE_PATTERN.match(raw_lines[index+3]):
            current_account = line
            index += 1
            while index + 2 <= total_lines:
                if stop_flag.is_set():
                    cost = round(time.time() - start_time, 2)
                    callback([], f"⏹️ 解析已停止 | 耗时 {cost}s")
                    return
                prod = raw_lines[index]
                url = raw_lines[index+1]
                code = raw_lines[index+2]
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
    max_workers = min(16, len(tasks))  # 16线程足够，再多会被限流

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
            result.append([parse_date, account, product, final_url, vid, ad_code])

    cost = round(time.time() - start_time, 2)
    msg = f"✅ 成功解析 {len(result)} 条 | 耗时 {cost}s"
    callback(result, msg)


class TikTokToolGUI:
    def __init__(self, root, auth_mgr):
        self.root = root
        self.auth_mgr = auth_mgr
        self.root.title(APP_TITLE)
        self.root.geometry("900x700")

        try:
            self.root.iconbitmap(APP_ICON)
        except:
            pass

        self.table_data = []
        self.stop_flag = threading.Event()
        self.is_parsing = False
        self.net_forbidden = False

        main = ttk.Frame(root)
        main.pack(fill="both", expand=1, padx=10, pady=10)

        # 3. 新增：网络状态UI
        net_frame = ttk.Frame(main)
        net_frame.grid(row=0, column=0, sticky="ew", pady=2)
        self.net_label = tk.StringVar(value="网络检测中...")
        ttk.Label(net_frame, text="网络：").pack(side="left")
        ttk.Label(net_frame, textvariable=self.net_label, foreground="blue").pack(side="left")
        self.refresh_btn = ttk.Button(net_frame, text="刷新网络", command=self.refresh_network)
        self.refresh_btn.pack(side="right")

        tip_text = "格式：ID → 产品名称+链接+code(可连续多组)，支持任意空行"
        ttk.Label(main, text=tip_text).grid(row=1, column=0, sticky="w")
        self.txt = tk.Text(main, wrap="word")
        self.txt.grid(row=2, column=0, sticky="nsew", pady=5)

        bf = ttk.Frame(main)
        bf.grid(row=3, column=0, sticky="w")
        self.btn_run = ttk.Button(bf, text="一键解析", command=self.start_parse)
        self.btn_run.pack(side="left", padx=5)

        # 4. 新增：停止解析按钮
        self.btn_stop = ttk.Button(bf, text="停止解析", command=self.stop_parse, state="disabled")
        self.btn_stop.pack(side="left", padx=5)

        ttk.Button(bf, text="手动导出Excel", command=self.export_excel).pack(side="left", padx=5)
        ttk.Button(bf, text="清空", command=self.clear).pack(side="left", padx=5)
        ttk.Button(bf, text="清空缓存", command=lambda: url_cache.clear()).pack(side="left", padx=5)

        ttk.Label(main, text=f"预览 | 授权：{auth_mgr.get_remain_days()}").grid(row=4, column=0, sticky="w")
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
        self.tree.grid(row=5, column=0, sticky="nsew", pady=5)

        sf = ttk.Frame(main)
        sf.grid(row=6, column=0, sticky="ew")
        self.status = tk.StringVar(value="就绪")
        ttk.Label(sf, textvariable=self.status, foreground="green").pack(side="left")
        ttk.Label(sf, text="售后：").pack(side="right")
        ttk.Button(sf, text=f"微信 {WECHAT_ID}", command=lambda: messagebox.showinfo("微信客服", WECHAT_ID)).pack(
            side="right")
        ttk.Button(sf, text=f"QQ {QQ_NUM}", command=lambda: webbrowser.open(f"tencent://message/?uin={QQ_NUM}")).pack(
            side="right", padx=2)

        main.grid_rowconfigure(2, weight=2)
        main.grid_rowconfigure(5, weight=5)
        main.grid_columnconfigure(0, weight=1)

        # 初始化网络检测 + 自动刷新
        self.refresh_network()
        #self.start_auto_net_refresh()

    # 3. 网络刷新 + 5. 地区限制
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
                self.root.after(0, lambda: messagebox.showwarning("限制", "❌ 中国大陆/香港网络禁止使用！"))
            else:
                self.root.after(0, lambda: self.btn_run.config(state="normal"))

        threading.Thread(target=task, daemon=True).start()

    def start_auto_net_refresh(self):
        def loop():
            while True:
                time.sleep(5)
                if not self.is_parsing:
                    self.root.after(0, self.refresh_network)

        threading.Thread(target=loop, daemon=True).start()

    def start_parse(self):
        if self.is_parsing or self.net_forbidden:
            return
        s = self.txt.get("1.0", "end-1c")
        if not s.strip():
            messagebox.showwarning("提示", "请输入内容")
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
            self.status.set(msg)  # 这里会自动显示 耗时xx秒
            self.btn_run.config(state="normal")
            self.btn_stop.config(state="disabled")
            self.is_parsing = False

        self.parse_thread = threading.Thread(
            target=background_parse_task,
            args=(s, self.stop_flag, cb),
            daemon=True
        )
        self.parse_thread.start()

    # 4. 停止解析功能
    def stop_parse(self):
        if not self.is_parsing:
            return
        self.stop_flag.set()
        self.status.set("正在停止...")
        # 等待线程结束（可选，增强稳定性）
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

    def clear(self):
        self.txt.delete("1.0", "end")
        for i in self.tree.get_children(): self.tree.delete(i)
        self.table_data = []
        self.status.set("已清空")


def main():
    am = AuthManager()
    root = tk.Tk()
    try:
        root.iconbitmap(APP_ICON)
    except:
        pass
    if not am.valid:
        ActivateWindow(root, am)
    else:
        TikTokToolGUI(root, am)
    def on_closing():
        client.close()  # 关闭httpx连接池
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()


if __name__ == "__main__":
    import time

    main()