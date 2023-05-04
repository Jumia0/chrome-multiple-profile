import os
import shutil
import subprocess
from tkinter import *
from tkinter import ttk, simpledialog, messagebox

import win32com.client

window = Tk()
window.title("Chrome多开")
window.geometry('248x240')
window.resizable(False, False)

chrome_data_path = os.path.join(os.environ['LOCALAPPDATA'], 'Google', 'Chrome', 'User Data')
listbox = Listbox(window, height=12, width=20)
listbox.grid(row=0, column=0, rowspan=5, padx=10, pady=10)
profile_prefix = 'Profile_'


def refresh():
    profiles = []
    for item in os.listdir(chrome_data_path):
        if item.startswith(profile_prefix):
            item_path = os.path.join(chrome_data_path, item)
            if os.path.isdir(item_path):
                profiles.append(item[len(profile_prefix):])

    print(profiles)

    selection = listbox.curselection()
    listbox.delete(0, END)
    for index, item in enumerate(profiles):
        listbox.insert(index + 1, item)
    if selection:
        try:
            listbox.select_set(selection)
        except Exception:
            pass
    return profiles


def open(profile=None):
    if profile is None:
        selection = listbox.curselection()
        if selection:
            profile = listbox.get(selection[0])
        else:
            return False
    chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
    cmd = [chrome_path, "-profile-directory=%s%s" % (profile_prefix, profile)]
    print(cmd)
    subprocess.run(cmd)


def create():
    result = simpledialog.askstring("请输入", "配置名称")
    if result:
        open(result)
        messagebox.showinfo("提示", "创建配置成功")
    refresh()


def shortcut(profile=None):
    if profile is None:
        selection = listbox.curselection()
        if selection:
            profile = listbox.get(selection[0])
        else:
            return False

    try:
        # 创建 Shell 对象
        shell = win32com.client.Dispatch("WScript.Shell")

        # 要创建快捷方式的 .exe 文件路径
        chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

        # 快捷方式名称和所在目录路径
        shortcut_name = profile
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

        # 命令行参数（以空格分隔）
        args = "-profile-directory=%s%s" % (profile_prefix, profile)

        # 创建快捷方式
        sc = shell.CreateShortcut(f"{desktop_path}\\{shortcut_name}.lnk")  # 注意要加上 .lnk 扩展名
        sc.TargetPath = chrome_path
        sc.Arguments = args
        sc.Save()

        messagebox.showinfo("提示", "创建快捷方式成功")

    except Exception as e:
        messagebox.showerror("错误", repr(e))


def edit(profile=None):
    if profile is None:
        selection = listbox.curselection()
        if selection:
            profile = listbox.get(selection[0])
        else:
            return False

    result = simpledialog.askstring("请输入", "新的名称")
    if result:
        old_dir = os.path.join(chrome_data_path, '%s%s' % (profile_prefix, profile))
        new_dir = os.path.join(chrome_data_path, '%s%s' % (profile_prefix, result))
        try:
            os.rename(old_dir, new_dir)
            messagebox.showinfo("提示", "重命名成功")
        except Exception as e:
            messagebox.showerror("错误", repr(e))

    refresh()


def remove(profile=None):
    if profile is None:
        selection = listbox.curselection()
        if selection:
            profile = listbox.get(selection[0])
        else:
            return False

    result = messagebox.askquestion("确认删除", "确认删除 %s ？" % profile)
    if result == "yes":
        profile_dir = os.path.join(chrome_data_path, '%s%s' % (profile_prefix, profile))
        print(profile_dir)
        try:
            shutil.rmtree(profile_dir)
            messagebox.showinfo("提示", "删除成功")
        except Exception as e:
            messagebox.showerror("错误", repr(e))

    refresh()


ttk.Button(window, text="打开", command=open, width=8).grid(row=0, column=1, padx=5, pady=5)
ttk.Button(window, text="创建配置", command=create, width=8).grid(row=1, column=1, padx=5, pady=5)
ttk.Button(window, text="快捷方式", command=shortcut, width=8).grid(row=2, column=1, padx=5, pady=5)
ttk.Button(window, text="重命名", command=edit, width=8).grid(row=3, column=1, padx=5, pady=5)
ttk.Button(window, text="删除配置", command=remove, width=8).grid(row=4, column=1, padx=5, pady=5)
refresh()

window.mainloop()
