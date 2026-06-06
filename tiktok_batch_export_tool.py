import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re
import urllib.request
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
APP_TITLE = "TikTok投流码专属自动化客户端"


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
HEADERS = {"User-Agent": "Mozilla/5.0"}
TIKTOK_SHORT_PATTERN = re.compile(r"https://www\.tiktok\.com/t/\S+")
ADS_CODE_PATTERN = re.compile(r"#[\w/=+]+")
# 提取VID的正则，匹配/video/后的数字串
VID_PATTERN = re.compile(r'/video/(\d+)')


def get_us_date():
    return (datetime.now() - timedelta(1)).strftime("%Y/%m/%d")


def parse_single_short_url(url):
    try:
        resp = urllib.request.urlopen(urllib.request.Request(url, headers=HEADERS), timeout=5)
        return resp.geturl().split("?")[0]
    except:
        return url


# 从链接中提取VID
def extract_vid_from_link(link):
    match = VID_PATTERN.search(link)
    return match.group(1) if match else ""


# ====================== 【核心优化】导出带居中对齐的XLSX Excel文件 ======================
def auto_export_excel(data):
    # 生成带时间戳的文件名，避免覆盖
    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"TikTok投流数据_{now_str}.xlsx"

    # 优先保存到桌面，无权限则保存到我的文档
    try:
        save_path = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.access(save_path, os.W_OK):
            save_path = os.path.join(os.path.expanduser("~"), "Documents")
        full_file_path = os.path.join(save_path, file_name)
    except:
        full_file_path = os.path.join(os.path.abspath("."), file_name)

    # 定义表头
    headers = ["日期", "ID", "产品名称", "链接", "VID", "code"]

    # 创建Excel工作簿和工作表
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TikTok投流数据"

    # 定义样式
    # 表头样式：加粗、水平垂直居中、浅灰色背景
    header_font = Font(bold=True, name="微软雅黑", size=11)
    header_fill = PatternFill("solid", fgColor="D3D3D3")
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 数据行样式：水平垂直居中、自动换行
    data_font = Font(name="微软雅黑", size=10)

    # 写入表头
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.alignment = center_alignment
        cell.fill = header_fill

    # 写入数据行，全部居中对齐
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, cell_value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
            cell.font = data_font
            cell.alignment = center_alignment

    # 自动适配列宽
    for col_idx in range(1, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        max_content_width = 0
        for cell in ws[col_letter]:
            try:
                cell_content = str(cell.value)
                if len(cell_content) > max_content_width:
                    max_content_width = len(cell_content)
            except:
                continue
        # 列宽最小8，最大50，避免过宽
        final_width = max(min(max_content_width + 3, 50), 8)
        ws.column_dimensions[col_letter].width = final_width

    # 冻结表头，滚动时固定
    ws.freeze_panes = "A2"

    # 保存Excel文件
    wb.save(full_file_path)

    # 自动用系统默认Excel/WPS程序打开文件
    try:
        if os.name == "nt":  # Windows系统
            os.startfile(full_file_path)
        else:  # Mac/Linux系统
            open_cmd = "open" if os.name == "posix" else "xdg-open"
            subprocess.run([open_cmd, full_file_path], check=True)
    except:
        pass  # 打开失败不中断流程

    return full_file_path


# ====================== 核心解析：账号 + 多条3行产品 ======================
def background_parse_task(text, callback):
    raw_lines = [x.strip() for x in text.splitlines() if x.strip()]
    tasks = []
    idx = 0
    cur_account = None

    while idx < len(raw_lines):
        line = raw_lines[idx]
        # 新账号判定：后面至少3行：产品、链接、投流码
        if idx + 2 < len(raw_lines) and TIKTOK_SHORT_PATTERN.match(raw_lines[idx + 2]) and ADS_CODE_PATTERN.match(
                raw_lines[idx + 3]):
            cur_account = line
            idx += 1
            # 循环读取该账号下所有产品(每组3行)
            while idx + 2 < len(raw_lines):
                p = raw_lines[idx]
                u = raw_lines[idx + 1]
                a = raw_lines[idx + 2]
                if TIKTOK_SHORT_PATTERN.match(u) and ADS_CODE_PATTERN.match(a):
                    tasks.append((cur_account, p, u, a))
                    idx += 3
                else:
                    break
        else:
            idx += 1

    if not tasks:
        callback([], "❌ 未识别有效数据")
        return

    res = []
    day = get_us_date()
    with ThreadPoolExecutor(5) as pool:
        future_map = {pool.submit(parse_single_short_url, t[2]): t for t in tasks}
        for future in as_completed(future_map):
            acc, mod, link, adcode = future_map[future]
            final_link = future.result()
            # 提取VID
            vid = extract_vid_from_link(final_link)
            # 行数据：日期、ID、产品名称、链接、VID、code
            row = [day, acc, mod, final_link, vid, adcode]
            res.append(row)

    # 自动导出+打开
    excel_file_path = auto_export_excel(res)
    msg = f"✅ 成功解析 {len(res)} 条\nExcel文件已自动保存并打开：\n{excel_file_path}"
    callback(res, msg)


class TikTokToolGUI:
    def __init__(self, root, auth_mgr):
        self.root = root
        self.auth_mgr = auth_mgr
        self.root.title(APP_TITLE)
        self.root.geometry("900x650")

        try:
            self.root.iconbitmap(APP_ICON)
        except:
            pass

        self.table_data = []
        main = ttk.Frame(root)
        main.pack(fill="both", expand=1, padx=10, pady=10)

        tip_text = "格式：ID → 产品名称+链接+code(可连续多组)，支持任意空行"
        ttk.Label(main, text=tip_text).grid(row=1, column=0, sticky="w")
        self.txt = tk.Text(main, wrap="word")
        self.txt.grid(row=2, column=0, sticky="nsew", pady=5)

        bf = ttk.Frame(main)
        bf.grid(row=3, column=0, sticky="w")
        self.btn_run = ttk.Button(bf, text="一键解析", command=self.start)
        self.btn_run.pack(side="left", padx=5)
        ttk.Button(bf, text="手动导出Excel", command=self.export_excel).pack(side="left", padx=5)
        ttk.Button(bf, text="清空", command=self.clear).pack(side="left", padx=5)

        ttk.Label(main, text=f"预览 | 授权：{auth_mgr.get_remain_days()}").grid(row=4, column=0, sticky="w")
        # 列名对应新表头
        cols = ["date", "id", "product", "link", "vid", "code"]
        self.tree = ttk.Treeview(main, columns=cols, show="headings")
        # 表头文字
        self.tree.heading("date", text="日期")
        self.tree.heading("id", text="ID")
        self.tree.heading("product", text="产品名称")
        self.tree.heading("link", text="链接")
        self.tree.heading("vid", text="VID")
        self.tree.heading("code", text="code")
        # 列宽适配
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

    def start(self):
        s = self.txt.get("1.0", "end-1c")
        if not s.strip():
            messagebox.showwarning("提示", "请输入内容")
            return
        self.btn_run.config(state="disabled")
        self.status.set("解析中...")

        def cb(res, msg):
            self.table_data = res
            for i in self.tree.get_children(): self.tree.delete(i)
            for r in res: self.tree.insert("", "end", values=r)
            self.status.set(msg.split("\n")[0])
            self.btn_run.config(state="normal")
            messagebox.showinfo("完成", msg)

        threading.Thread(target=background_parse_task, args=(s, cb), daemon=1).start()

    def export_excel(self):
        if not self.table_data:
            messagebox.showwarning("提示", "无数据可导出")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
        if not p: return
        # 手动导出也使用带格式的Excel
        headers = ["日期", "ID", "产品名称", "链接", "VID", "code"]
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TikTok投流数据"

        # 样式
        header_font = Font(bold=True, name="微软雅黑", size=11)
        header_fill = PatternFill("solid", fgColor="D3D3D3")
        center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        data_font = Font(name="微软雅黑", size=10)

        # 写入表头
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = header_font
            cell.alignment = center_alignment
            cell.fill = header_fill

        # 写入数据
        for row_idx, row_data in enumerate(self.table_data, 2):
            for col_idx, cell_value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                cell.font = data_font
                cell.alignment = center_alignment

        # 适配列宽
        for col_idx in range(1, len(headers) + 1):
            col_letter = get_column_letter(col_idx)
            max_content_width = 0
            for cell in ws[col_letter]:
                try:
                    cell_content = str(cell.value)
                    if len(cell_content) > max_content_width:
                        max_content_width = len(cell_content)
                except:
                    continue
            final_width = max(min(max_content_width + 3, 50), 8)
            ws.column_dimensions[col_letter].width = final_width

        ws.freeze_panes = "A2"
        wb.save(p)
        messagebox.showinfo("成功", f"Excel已导出到：{p}")

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
    root.mainloop()


if __name__ == "__main__":
    main()