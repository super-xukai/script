import os
import datetime
import subprocess
import sys
import platform
import tkinter as tk
from tkinter import ttk, messagebox
import json

def load_archive_config():
    """加载归档目录配置"""
    config_file = 'archive_config.json'
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        # 如果配置文件不存在，创建一个示例配置文件
        default_config = [
            {
                "archive_dir": "E:\\archive\\民宿项目记录",
                "shortcut_prefix": "民宿项目",
                "display_name": "民宿项目记录"
            },
            {
                "archive_dir": "E:\\archive\\民宿财务记录", 
                "shortcut_prefix": "民宿财务",
                "display_name": "民宿财务记录"
            }
        ]
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, ensure_ascii=False, indent=4)
        print(f"配置文件 {config_file} 不存在，已创建示例配置文件，请修改后重新运行程序。")
        return None
    except Exception as e:
        print(f"读取配置文件出错: {e}")
        return None

def get_desktop_path():
    """获取桌面路径，支持Windows和Linux系统"""
    if platform.system() == 'Windows':
        return os.path.join(os.path.expanduser('~'), 'Desktop')
    else:  # 假设是Linux或macOS
        return os.path.join(os.path.expanduser('~'), 'Desktop')

def create_folder(folder_path):
    """创建文件夹，如果不存在的话"""
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"文件夹已创建: {folder_path}")
        return True
    print(f"文件夹已存在: {folder_path}")
    return False

def create_shortcut(target_path, shortcut_path, shortcut_prefix, today_str):
    """创建快捷方式，支持Windows和Linux系统"""
    if platform.system() == 'Windows':
        try:
            # 尝试使用pywin32
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = target_path
            shortcut.Save()
        except Exception as e:
            # 备选方案：使用PowerShell创建快捷方式
            ps_command = f'''
            $WshShell = New-Object -comObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
            $Shortcut.TargetPath = "{target_path}"
            $Shortcut.Save()
            '''
            subprocess.run(["powershell", "-Command", ps_command], check=True)
    else:  # 为Linux或macOS创建.desktop文件
        desktop_entry = f"""[Desktop Entry]
Name={shortcut_prefix} {today_str}
Type=Directory
Path={target_path}
Icon=folder
"""
        with open(shortcut_path, 'w') as f:
            f.write(desktop_entry)
        os.chmod(shortcut_path, 0o755)  # 确保可执行

def delete_old_shortcuts(desktop_path, today_str, shortcut_prefix):
    """删除桌面上除了今天之外的所有archive快捷方式"""
    for item in os.listdir(desktop_path):
        item_path = os.path.join(desktop_path, item)
        # Windows快捷方式检查
        if platform.system() == 'Windows':
            if item.endswith('.lnk') and item.startswith(f'{shortcut_prefix} '):
                date_str = item[len(shortcut_prefix)+1:-4]  # 提取日期部分
                if date_str != today_str:
                    os.remove(item_path)
                    print(f"已删除旧快捷方式: {item_path}")
        # Linux/macOS快捷方式检查
        else:
            if item.endswith('.desktop') and item.startswith(f'{shortcut_prefix} '):
                date_str = item[len(shortcut_prefix)+1:-8]  # 提取日期部分
                if date_str != today_str:
                    os.remove(item_path)
                    print(f"已删除旧快捷方式: {item_path}")

def open_folder(folder_path):
    """打开指定文件夹"""
    if platform.system() == 'Windows':
        os.startfile(folder_path)
    elif platform.system() == 'Darwin':  # macOS
        subprocess.run(['open', folder_path])
    else:  # Linux
        subprocess.run(['xdg-open', folder_path])

def select_archive_dir(archive_configs):
    """显示选择窗口并返回选中的归档目录配置"""
    root = tk.Tk()
    root.title("选择归档目录")
    root.geometry("400x300")
    root.resizable(False, False)
    
    selected_config = {}
    
    # 创建标签
    label = ttk.Label(root, text="请选择归档目录:", font=("Arial", 12))
    label.pack(pady=10)
    
    # 创建列表框
    listbox = tk.Listbox(root, font=("Arial", 10))
    listbox.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
    
    # 添加目录到列表框
    for i, config in enumerate(archive_configs):
        display_name = config["display_name"]
        listbox.insert(tk.END, display_name)
    
    # 设置默认选中第一个
    if archive_configs:
        listbox.selection_set(0)
        selected_config.update(archive_configs[0])
    
    def on_select(event=None):
        selection = listbox.curselection()
        if selection:
            index = selection[0]
            selected_config.clear()
            selected_config.update(archive_configs[index])
    
    def on_ok():
        if not selected_config:
            messagebox.showwarning("警告", "请选择一个归档目录")
            return
        root.quit()
    
    def on_cancel():
        selected_config.clear()
        root.quit()
    
    # 绑定选择事件
    listbox.bind('<<ListboxSelect>>', on_select)
    
    # 创建按钮框架
    button_frame = ttk.Frame(root)
    button_frame.pack(pady=10)
    
    # 创建确定按钮
    ok_button = ttk.Button(button_frame, text="确定", command=on_ok)
    ok_button.pack(side=tk.LEFT, padx=10)
    
    # 创建取消按钮
    cancel_button = ttk.Button(button_frame, text="取消", command=on_cancel)
    cancel_button.pack(side=tk.LEFT, padx=10)
    
    # 居中显示窗口
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()
    root.destroy()
    
    return selected_config if selected_config else None

if __name__ == "__main__":
    try:
        # 加载归档目录配置
        archive_configs = load_archive_config()
        
        # 检查是否有配置的归档目录
        if archive_configs is None:
            # 配置文件刚创建，退出程序
            sys.exit(0)
            
        if not archive_configs:
            print("错误: 未配置任何归档目录")
            sys.exit(1)
        
        # 显示选择窗口
        selected_config = select_archive_dir(archive_configs)
        
        # 如果用户取消选择，则退出
        if not selected_config:
            print("用户取消操作")
            sys.exit(0)
        
        # 获取选中的配置信息
        selected_archive_dir = selected_config["archive_dir"]
        shortcut_prefix = selected_config["shortcut_prefix"]
        display_name = selected_config["display_name"]
        
        # 获取今天的日期
        today = datetime.datetime.now()
        today_str = today.strftime('%Y-%m-%d')

        # 设置路径
        today_folder = os.path.join(selected_archive_dir, today_str)
        desktop_path = get_desktop_path()
        shortcut_name = f'{shortcut_prefix} {today_str}.lnk' if platform.system() == 'Windows' else f'{shortcut_prefix} {today_str}.desktop'
        shortcut_path = os.path.join(desktop_path, shortcut_name)

        # 创建文件夹
        folder_created = create_folder(today_folder)

        # 创建快捷方式
        if folder_created:
            create_shortcut(today_folder, shortcut_path, shortcut_prefix, today_str)
            print(f"快捷方式已创建: {shortcut_path}")

        # 删除旧快捷方式（只删除同名前缀的）
        delete_old_shortcuts(desktop_path, today_str, shortcut_prefix)

        # 打开文件夹
        open_folder(today_folder)

    except Exception as e:
        print(f"发生错误: {e}")
        sys.exit(1)