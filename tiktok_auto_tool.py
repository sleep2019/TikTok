import tkinter as tk
from tkinter import ttk, messagebox
import requests
import re
import pyperclip

# 禁用警告
import warnings
warnings.filterwarnings("ignore")

def get_real_url(short_url):
    """
    后台静默获取TikTok短链跳转后的真实长链接
    不打开浏览器！
    """
    try:
        session = requests.Session()
        # 模拟手机请求头，防止被拦截
        headers = {
            "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit"
        }
        # 只获取跳转地址，不下载内容
        response = session.head(
            short_url,
            headers=headers,
            allow_redirects=True,
            timeout=10,
            verify=False
        )
        return response.url
    except:
        return None

def process_auto():
    short_url = input_entry.get().strip()
    if not short_url:
        messagebox.showwarning("提示", "请输入TikTok短链")
        return

    status_label.config(text="🔄 正在解析链接...")
    root.update()

    # 自动获取长链
    real_url = get_real_url(short_url)

    if not real_url:
        messagebox.showerror("错误", "解析失败，请检查链接是否有效")
        status_label.config(text="❌ 解析失败")
        return

    # 裁剪 ? 及后面所有内容
    clean_url = re.split(r"\?", real_url)[0]

    # 显示结果
    result_entry.delete(0, tk.END)
    result_entry.insert(0, clean_url)
    status_label.config(text="✅ 处理完成！已自动去参数")

def copy_result():
    url = result_entry.get().strip()
    if url:
        pyperclip.copy(url)
        status_label.config(text="✅ 已复制到剪贴板")

# ==================== GUI界面 ====================
root = tk.Tk()
root.title("TikTok短链全自动解析工具")
root.geometry("650x220")
root.resizable(False, False)

# 输入
ttk.Label(root, text="输入TikTok短链：").place(x=20, y=30)
input_entry = ttk.Entry(root, width=60, font=("微软雅黑", 10))
input_entry.place(x=150, y=30)

# 开始按钮
start_btn = ttk.Button(root, text="开始解析", command=process_auto)
start_btn.place(x=250, y=80, width=120)

# 结果
ttk.Label(root, text="最终纯净链接：").place(x=20, y=130)
result_entry = ttk.Entry(root, width=60, font=("微软雅黑", 10))
result_entry.place(x=150, y=130)

# 复制按钮
copy_btn = ttk.Button(root, text="一键复制", command=copy_result)
copy_btn.place(x=250, y=170, width=120)

# 状态
status_label = ttk.Label(root, text="就绪", foreground="green")
status_label.place(x=20, y=170)

root.mainloop()