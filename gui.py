from tkinter import messagebox, ttk, filedialog, PhotoImage, Menu, Label, StringVar
from version import VERSION
from datetime import datetime
from PIL import Image, ImageTk
import os
import tkinter as tk
import pandas as pd
import sys
import webbrowser
from scraper import get_stock_names, fetch_mops_today_page, parse_mops_data, fetch_mops_yesterday_page,parse_mops_yesterday_data


def create_gui():
    # 初始化
    root = tk.Tk()
    root.title(f"Ethen's Dividend NewsPal {VERSION}")
    root.geometry("1300x600")

    fetch_icon, yesterday_icon, csv_icon, open_website_icon = import_button_icon()
    fetch_button, yesterday_button, save_button, open_website_button = setup_buttons(
        root, fetch_icon, yesterday_icon, csv_icon, open_website_icon,  # 添加 open_website_icon
        lambda: fetch_and_display(tree, status_var),
        lambda: fetch_yesterday_and_display(tree, status_var),
        lambda: save_data_as_csv(tree),
        lambda: open_website("https://mops.twse.com.tw/mops/web/t05sr01_1")  # 導向指定網站
    )
    fetch_button.image = fetch_icon
    yesterday_button.image = yesterday_icon
    save_button.image = csv_icon
    open_website_button.image = open_website_icon  # 新增的按鈕圖標設置

    set_gui_icon(root)
    create_menu(root)
    status_var, status_bar = create_status_var(root)
    create_time_label(status_bar)

    tree = create_treeview(root)

    fetch_button.bind("<Enter>", lambda event: fetch_on_enter(status_var, event))
    fetch_button.bind("<Leave>", lambda event: on_leave(status_var, event))
    yesterday_button.bind("<Enter>", lambda event: fetch_yesterday_on_enter(status_var, event))
    yesterday_button.bind("<Leave>", lambda event: on_leave(status_var, event))
    save_button.bind("<Enter>", lambda event: save_on_enter(status_var, event))
    save_button.bind("<Leave>", lambda event: on_leave(status_var, event))
    open_website_button.bind("<Enter>", lambda event: open_website_on_enter(status_var, event))
    open_website_button.bind("<Leave>", lambda event: on_leave(status_var, event))

    root.mainloop()


def open_website(url):
    webbrowser.open(url)


def fetch_and_display(tree, status_var, attempt=0):
    headings = ("公司代號", "公司簡稱", "發言日期", "發言時間", "主旨")
    configure_treeview(tree, headings)
    stock_list = get_stock_names()
    url = "https://mops.twse.com.tw/mops/web/t05sr01_1"
    try:
        soup = fetch_mops_today_page(url)
        stock_info = parse_mops_data(soup, stock_list)

        if stock_info:
            for i in tree.get_children():
                tree.delete(i)
            for info in stock_info:
                tree.insert("", "end", values=info)
            messagebox.showinfo("成功", "已蒐集成功本日之股票股利標的之公開資訊")
        else:
            if attempt < 3:
                if messagebox.askyesno("失敗", "蒐集到0筆資料，請問需要再蒐集一次嗎？"):
                    attempt += 1
                    return fetch_and_display(tree, status_var, attempt)
            else:
                messagebox.showerror("提示", "很抱歉，目前可能沒有可蒐集的資料。請稍後再試。")
    except Exception as e:
        messagebox.showerror("錯誤", f"發生錯誤: {e}")


def fetch_yesterday_and_display(tree, status_var, attempt=0):
    headings = ("公司代號", "公司簡稱", "發言日期", "發言時間", "主旨")
    configure_treeview(tree, headings)
    stock_list = get_stock_names()
    url = "https://mops.twse.com.tw/mops/web/t05sr01_1"
    try:
        soup = fetch_mops_yesterday_page(url)
        stock_info = parse_mops_yesterday_data(soup, stock_list)

        if stock_info:
            for i in tree.get_children():
                tree.delete(i)
            for info in stock_info:
                tree.insert("", "end", values=info)
            messagebox.showinfo("成功", "已蒐集成功昨日之股票股利標的之公開資訊")
        else:
            if attempt < 3:
                if messagebox.askyesno("失敗", "蒐集到0筆資料，請問需要再蒐集一次嗎？"):
                    attempt += 1
                    return fetch_yesterday_and_display(tree, status_var, attempt)
            else:
                messagebox.showerror("提示", "很抱歉，目前可能沒有可蒐集的資料。請稍後再試。")
    except Exception as e:
        messagebox.showerror("錯誤", f"發生錯誤: {e}")


def set_gui_icon(root):
    # 設置EXE ICON
    current_dir = os.path.dirname(__file__)
    icon_path = os.path.join(current_dir, "icon.ico")

    # 設置窗口圖標
    try:
        root.iconbitmap(icon_path)
    except FileNotFoundError as e:
        messagebox.showerror("錯誤", "載入圖標時發現錯誤")


def create_time_label(status_bar):
    time_label = tk.Label(status_bar, font=('Helvetica', 12))
    time_label.pack(side='right', padx=5)
    update_time(time_label)


def update_time(time_label):
    now = datetime.now()
    current_time = now.strftime("%Y-%m-%d %H:%M:%S")
    time_label.config(text=current_time)
    time_label.after(1000, lambda: update_time(time_label))


def save_data_as_csv(tree):
    # 檢查資料是否存在
    if not tree.get_children():
        messagebox.showerror("錯誤", "沒有資料可以匯出")
        return
    # 準備數據
    data = []
    columns = [tree.heading(col)['text'] for col in tree['columns']]
    for child in tree.get_children():
        row_data = [tree.set(child, col) for col in tree['columns']]
        data.append(row_data)
    # 檢查是否真的獲取到了數據
    if not data:
        messagebox.showerror("錯誤", "沒有資料可以匯出")
        return
    # 轉換成DF
    df = pd.DataFrame(data, columns=columns)
    # 保存為CSV
    today_date = datetime.today().strftime('%Y-%m-%d')
    default_filename = f"{today_date}_export.csv"
    file_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[("CSV files", "*.csv")],
                                             initialfile=default_filename)
    if file_path:
        df.to_csv(file_path, index=False, encoding='utf-8-sig')
        messagebox.showinfo("成功", "儲存成功！")


def create_status_var(root):
    status_var = StringVar()
    status_var.set("Ready")
    statusbar = Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor="w", font=('Helvetica', 11))
    statusbar.pack(side=tk.BOTTOM, fill=tk.X)
    return status_var, statusbar


def create_menu(root):
    menubar = Menu(root)
    root.config(menu=menubar)
    help_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", """
    如果您在使用本程式時遇到任何問題、有任何建議或者想要提供反饋意見，請隨時通過郵件與我們聯絡。\n
    您可以寄送郵件至 hamiltondevjourney@gmail.com，我們將會盡快回覆您的郵件。\n
    感謝您的支持與理解！
    """))


def configure_treeview(tree, columns):
    for _ in tree.get_children():
        tree.delete(_)
    tree["columns"] = columns
    for col in tree["columns"]:
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor=tk.CENTER, stretch=tk.YES)


def clear_and_display_data(tree, data):
    tree.delete(*tree.get_children())
    for info in data:
        tree.insert("", "end", values=info)


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def import_button_icon():
    fetch = resource_path('fetch.png')
    yesterday = resource_path('yesterday.png')
    csv_icon_path = resource_path('csv.png')
    open_website = resource_path('open_website.png')
    analysis_icon = load_and_resize(fetch, 50, 55)
    yesterday_icon = load_and_resize(yesterday, 50, 55)
    csv_icon = load_and_resize(csv_icon_path, 50, 55)
    open_website_icon = load_and_resize(open_website, 50, 55)
    return analysis_icon, yesterday_icon, csv_icon, open_website_icon


def load_and_resize(icon_path, width, height):
    original_icon = Image.open(icon_path)
    resized_icon = original_icon.resize((width, height))
    return ImageTk.PhotoImage(resized_icon)


def create_button(frame, text, icon, command, width=20):
    button = ttk.Button(frame, text=text, image=icon, command=command, width=width)
    button.pack(side='left', padx=5)
    return button


def setup_buttons(root, fetch_icon, yesterday_icon, csv_icon, open_website_icon,
                  fetch_command, yesterday_command, save_command, open_website_command):
    button_frame = tk.Frame(root)
    button_frame.pack(side='top', fill='x', pady=5)

    fetch_button = create_button(button_frame, "Fetch Stock Data", fetch_icon, fetch_command)
    yesterday_button = create_button(button_frame, "Fetch Yesterday Stock Data", yesterday_icon, yesterday_command)
    save_button = create_button(button_frame, "Save As Csv", csv_icon, save_command)
    open_website_button = create_button(button_frame, "Open MOPS Website", open_website_icon, open_website_command)  # 新增的按鈕

    return fetch_button, yesterday_button, save_button, open_website_button


def create_treeview(root):
    tree = ttk.Treeview(root, columns=("Col1", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7", "Col8"), show='headings')
    tree.pack(expand=True, fill='both')
    return tree


def fetch_on_enter(status_var, event):
    status_var.set("至目標網站蒐集股票股利標的本日所公開之資訊")


def save_on_enter(status_var, event):
    status_var.set("將蒐集之內容儲存至CSV檔案")


def open_website_on_enter(status_var, event):
    status_var.set("前往公開資訊觀測站")


def fetch_yesterday_on_enter(status_var, event):
    status_var.set("至目標網站蒐集股票股利標的昨日所公開之資訊，(本日重大訊息包含前一日17:30以後之訊息)")


def on_leave(status_var, event):
    status_var.set("Ready")


if __name__ == "__main__":
    create_gui()
