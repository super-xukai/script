import tkinter as tk
from tkinter import ttk, messagebox
import psutil

class PortCheckerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("端口检测工具")
        self.root.geometry("800x500")
        
        self.processes = {}  # 存储进程信息
        
        self.create_widgets()
        
    def create_widgets(self):        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="端口检测工具", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        # 端口输入区域
        input_frame = ttk.LabelFrame(main_frame, text="端口搜索", padding="10")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(3, weight=1)
        
        ttk.Label(input_frame, text="端口号:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        self.port_entry = ttk.Entry(input_frame, width=20)
        self.port_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        self.port_entry.bind("<Return>", lambda event: self.search_ports())
        
        # 状态筛选下拉框
        ttk.Label(input_frame, text="状态:").grid(row=0, column=2, sticky=tk.W, padx=(10, 5))
        
        # 状态选项：显示文本与实际值的映射
        self.status_options = {
            "全部状态": "ALL",
            "监听": "LISTEN",
            "已建立": "ESTABLISHED",
            "等待关闭": "CLOSE_WAIT",
            "时间等待": "TIME_WAIT"
        }
        
        self.status_var = tk.StringVar(value="监听")  # 默认选中"监听"
        self.status_combo = ttk.Combobox(input_frame, textvariable=self.status_var, 
                                        values=list(self.status_options.keys()), 
                                        state="readonly", width=12)
        self.status_combo.grid(row=0, column=3, sticky=(tk.W, tk.E), padx=(0, 5))
        
        self.search_btn = ttk.Button(input_frame, text="搜索端口", command=self.search_ports)
        self.search_btn.grid(row=0, column=4, padx=(0, 5))
        
        self.refresh_btn = ttk.Button(input_frame, text="刷新列表", command=self.refresh_list)
        self.refresh_btn.grid(row=0, column=5)
        
        # 端口列表
        list_frame = ttk.LabelFrame(main_frame, text="端口占用情况", padding="5")
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 创建Treeview来显示端口信息
        columns = ("port", "pid", "name", "status")
        self.port_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        self.port_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 设置列标题
        self.port_tree.heading("port", text="端口号")
        self.port_tree.heading("pid", text="进程ID")
        self.port_tree.heading("name", text="应用程序")
        self.port_tree.heading("status", text="状态")
        
        # 设置列宽
        self.port_tree.column("port", width=100, anchor=tk.CENTER)
        self.port_tree.column("pid", width=100, anchor=tk.CENTER)
        self.port_tree.column("name", width=300, anchor=tk.W)
        self.port_tree.column("status", width=100, anchor=tk.CENTER)
        
        # 添加滚动条
        tree_scroll_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.port_tree.yview)
        tree_scroll_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.port_tree.configure(yscrollcommand=tree_scroll_y.set)
        
        tree_scroll_x = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.port_tree.xview)
        tree_scroll_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.port_tree.configure(xscrollcommand=tree_scroll_x.set)
        
        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(10, 0))
        
        self.kill_btn = ttk.Button(button_frame, text="终止选中进程", command=self.kill_process, state=tk.DISABLED)
        self.kill_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.kill_all_btn = ttk.Button(button_frame, text="终止所有选中", command=self.kill_all_processes, state=tk.DISABLED)
        self.kill_all_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 绑定选择事件
        self.port_tree.bind("<<TreeviewSelect>>", self.on_select)
        
        # 初始化时搜索常用端口
        self.search_common_ports()
    
    def get_status_chinese(self, status):
        """将连接状态转换为中文"""
        status_map = {
            'ESTABLISHED': '已建立',
            'SYN_SENT': '连接请求已发送',
            'SYN_RECV': '连接请求已接收',
            'FIN_WAIT1': '等待终止请求1',
            'FIN_WAIT2': '等待终止请求2',
            'TIME_WAIT': '时间等待',
            'CLOSE': '关闭',
            'CLOSE_WAIT': '等待关闭',
            'LAST_ACK': '最后确认',
            'LISTEN': '监听',
            'CLOSING': '关闭中',
            'NONE': '无',
            'UNKNOWN': '未知'
        }
        return status_map.get(status.upper(), status)
    
    def get_status_english(self, status):
        """将中文状态转换为英文"""
        chinese_to_english = {
            '已建立': 'ESTABLISHED',
            '连接请求已发送': 'SYN_SENT',
            '连接请求已接收': 'SYN_RECV',
            '等待终止请求1': 'FIN_WAIT1',
            '等待终止请求2': 'FIN_WAIT2',
            '时间等待': 'TIME_WAIT',
            '关闭': 'CLOSE',
            '等待关闭': 'CLOSE_WAIT',
            '最后确认': 'LAST_ACK',
            '监听': 'LISTEN',
            '关闭中': 'CLOSING',
            '无': 'NONE',
            '未知': 'UNKNOWN'
        }
        return chinese_to_english.get(status, status)
    
    def search_common_ports(self):
        """初始化时搜索常用端口"""
        common_ports = [80, 443, 8080, 8000, 3306, 5432, 6379, 27017, 22, 21, 25, 110, 143, 993, 995]
        self.search_ports_by_list(common_ports)
    
    def search_ports(self):
        """根据输入搜索端口"""
        port_input = self.port_entry.get().strip()
        
        # 如果未输入端口号，则搜索所有端口
        if not port_input:
            self.search_all_ports()
            return
        
        # 解析输入（支持单个端口或端口范围）
        ports = []
        try:
            if '-' in port_input:
                # 端口范围
                start, end = map(int, port_input.split('-'))
                if start > end:
                    start, end = end, start
                ports = list(range(start, end + 1))
            elif ',' in port_input:
                # 多个端口
                ports = [int(p.strip()) for p in port_input.split(',')]
            else:
                # 单个端口
                ports = [int(port_input)]
            
            # 验证端口范围
            for port in ports:
                if not (1 <= port <= 65535):
                    messagebox.showwarning("输入错误", f"端口号 {port} 超出有效范围 (1-65535)")
                    return
            
            self.search_ports_by_list(ports)
        except ValueError:
            messagebox.showwarning("输入错误", "请输入有效的端口号（例如：8080 或 8000-8010 或 8080,8000,3306）")
    
    def search_all_ports(self):
        """搜索所有端口"""
        # 清空现有列表
        for item in self.port_tree.get_children():
            self.port_tree.delete(item)
        
        self.processes = {}
        
        # 获取所有网络连接
        try:
            connections = psutil.net_connections(kind='inet')
        except Exception as e:
            messagebox.showerror("错误", f"无法获取网络连接信息: {e}")
            return
        
        # 获取筛选状态
        selected_status_text = self.status_var.get()
        selected_status = self.status_options.get(selected_status_text, "ALL")
        
        # 查找所有端口的占用情况
        found_ports = {}
        for conn in connections:
            if conn.laddr:
                # 根据状态筛选
                if selected_status != "ALL" and conn.status.upper() != selected_status:
                    continue
                    
                port = conn.laddr.port
                if port not in found_ports:
                    found_ports[port] = []
                
                pid = conn.pid
                status = conn.status
                
                # 获取进程信息
                try:
                    if pid:
                        process = psutil.Process(pid)
                        name = process.name()
                    else:
                        name = "系统进程"
                        pid = "N/A"
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    name = "未知进程"
                    pid = "N/A"
                
                found_ports[port].append({
                    'pid': pid,
                    'name': name,
                    'status': status
                })
        
        # 显示结果
        for port in sorted(found_ports.keys()):
            port_info_list = found_ports[port]
            for i, info in enumerate(port_info_list):
                pid = info['pid']
                name = info['name']
                status = self.get_status_chinese(info['status'])
                
                # 为每个进程创建唯一ID
                item_id = f"{port}_{pid}_{i}"
                self.processes[item_id] = {
                    'port': port,
                    'pid': pid,
                    'name': name,
                    'status': info['status']  # 保存原始状态用于内部处理
                }
                
                # 添加到Treeview
                self.port_tree.insert("", "end", item_id, values=(port, pid, name, status))
        
        # 如果没有找到任何结果
        if not found_ports:
            self.port_tree.insert("", "end", "no_results", values=("-", "-", "未找到任何端口", "-"))
    
    def search_ports_by_list(self, ports):
        """根据端口列表搜索"""
        # 清空现有列表
        for item in self.port_tree.get_children():
            self.port_tree.delete(item)
        
        self.processes = {}
        
        # 获取所有网络连接
        try:
            connections = psutil.net_connections(kind='inet')
        except Exception as e:
            messagebox.showerror("错误", f"无法获取网络连接信息: {e}")
            return
        
        # 获取筛选状态
        selected_status_text = self.status_var.get()
        selected_status = self.status_options.get(selected_status_text, "ALL")
        
        # 查找指定端口的占用情况
        found_ports = {}
        for conn in connections:
            if conn.laddr and conn.laddr.port in ports:
                # 根据状态筛选
                if selected_status != "ALL" and conn.status.upper() != selected_status:
                    continue
                    
                port = conn.laddr.port
                if port not in found_ports:
                    found_ports[port] = []
                
                pid = conn.pid
                status = conn.status
                
                # 获取进程信息
                try:
                    if pid:
                        process = psutil.Process(pid)
                        name = process.name()
                    else:
                        name = "系统进程"
                        pid = "N/A"
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    name = "未知进程"
                    pid = "N/A"
                
                found_ports[port].append({
                    'pid': pid,
                    'name': name,
                    'status': status
                })
        
        # 显示结果
        for port in sorted(found_ports.keys()):
            port_info_list = found_ports[port]
            for i, info in enumerate(port_info_list):
                pid = info['pid']
                name = info['name']
                status = self.get_status_chinese(info['status'])
                
                # 为每个进程创建唯一ID
                item_id = f"{port}_{pid}_{i}"
                self.processes[item_id] = {
                    'port': port,
                    'pid': pid,
                    'name': name,
                    'status': info['status']  # 保存原始状态用于内部处理
                }
                
                # 添加到Treeview
                self.port_tree.insert("", "end", item_id, values=(port, pid, name, status))
        
        # 如果没有找到任何结果，显示未占用的端口
        if not found_ports:
            for port in ports:
                item_id = f"empty_{port}"
                self.port_tree.insert("", "end", item_id, values=(port, "-", "未占用", "-"))
    
    def refresh_list(self):
        """刷新端口列表"""
        # 获取当前搜索的端口
        current_ports = set()
        for item_id in self.port_tree.get_children():
            values = self.port_tree.item(item_id, "values")
            if values and values[0] not in ["-", "端口号"]:
                try:
                    port = int(values[0])
                    current_ports.add(port)
                except (ValueError, IndexError):
                    pass
        
        if current_ports:
            self.search_ports_by_list(list(current_ports))
        else:
            # 如果没有特定端口，则根据当前输入决定行为
            port_input = self.port_entry.get().strip()
            if port_input:
                self.search_ports()
            else:
                self.search_all_ports()
    
    def on_select(self, event):
        """选择项改变时的处理"""
        selection = self.port_tree.selection()
        if selection:
            self.kill_btn.config(state=tk.NORMAL)
            self.kill_all_btn.config(state=tk.NORMAL)
        else:
            self.kill_btn.config(state=tk.DISABLED)
            self.kill_all_btn.config(state=tk.DISABLED)
    
    def kill_process(self):
        """终止选中的进程"""
        selection = self.port_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要终止的进程")
            return
        
        item_id = selection[0]
        if item_id not in self.processes:
            messagebox.showwarning("警告", "无法找到选中进程的信息")
            return
        
        process_info = self.processes[item_id]
        pid = process_info['pid']
        port = process_info['port']
        name = process_info['name']
        
        if pid == "N/A":
            messagebox.showwarning("警告", "无法终止系统进程")
            return
        
        # 确认对话框
        result = messagebox.askyesno(
            "确认终止",
            f"确定要终止以下进程吗?\n\n应用程序: {name}\n进程ID: {pid}\n占用端口: {port}",
            icon='warning'
        )
        
        if result:
            try:
                process = psutil.Process(int(pid))
                process.terminate()
                # 等待进程结束
                process.wait(timeout=3)
                messagebox.showinfo("成功", f"进程 {name} (PID: {pid}) 已成功终止")
                
                # 刷新列表
                self.refresh_list()
            except psutil.NoSuchProcess:
                messagebox.showinfo("提示", "进程已经结束")
                self.refresh_list()
            except psutil.AccessDenied:
                messagebox.showerror("错误", "权限不足，无法终止该进程")
            except psutil.TimeoutExpired:
                # 如果进程没有及时终止，强制杀死
                try:
                    process.kill()
                    messagebox.showinfo("成功", f"进程 {name} (PID: {pid}) 已强制终止")
                    self.refresh_list()
                except Exception as e:
                    messagebox.showerror("错误", f"强制终止进程失败: {e}")
            except Exception as e:
                messagebox.showerror("错误", f"终止进程时发生错误: {e}")
    
    def kill_all_processes(self):
        """终止所有选中的进程"""
        selection = self.port_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要终止的进程")
            return
        
        # 收集要终止的进程信息
        processes_to_kill = []
        for item_id in selection:
            if item_id in self.processes:
                process_info = self.processes[item_id]
                if process_info['pid'] != "N/A":
                    processes_to_kill.append(process_info)
        
        if not processes_to_kill:
            messagebox.showwarning("警告", "没有可终止的进程")
            return
        
        # 确认对话框
        process_list = "\n".join([f"{p['name']} (PID: {p['pid']}, Port: {p['port']})" 
                                 for p in processes_to_kill])
        result = messagebox.askyesno(
            "确认终止",
            f"确定要终止以下 {len(processes_to_kill)} 个进程吗?\n\n{process_list}",
            icon='warning'
        )
        
        if result:
            success_count = 0
            fail_count = 0
            
            for process_info in processes_to_kill:
                try:
                    pid = int(process_info['pid'])
                    process = psutil.Process(pid)
                    process.terminate()
                    process.wait(timeout=3)
                    success_count += 1
                except psutil.NoSuchProcess:
                    success_count += 1  # 进程已经结束也算成功
                except psutil.AccessDenied:
                    fail_count += 1
                except psutil.TimeoutExpired:
                    try:
                        process.kill()
                        success_count += 1
                    except:
                        fail_count += 1
                except Exception:
                    fail_count += 1
            
            # 显示结果
            if fail_count == 0:
                messagebox.showinfo("成功", f"成功终止了 {success_count} 个进程")
            else:
                messagebox.showinfo("完成", f"成功终止了 {success_count} 个进程，{fail_count} 个进程终止失败")
            
            # 刷新列表
            self.refresh_list()

def main():
    """主函数"""
    root = tk.Tk()
    app = PortCheckerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()