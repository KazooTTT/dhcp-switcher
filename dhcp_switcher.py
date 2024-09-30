import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import psutil
import re
import locale
import json  
import ctypes
import sys

class DHCPSwitcher:
    def __init__(self, master):
        self.master = master
        master.title("DHCP 切换器")
        master.geometry("300x450")

        self.interface_label = ttk.Label(master, text="网络接口:")
        self.interface_label.pack(pady=5)

        self.interface_combobox = ttk.Combobox(master, state="readonly")
        self.interface_combobox.pack(pady=5)
        self.load_interfaces()
        self.interface_combobox.bind("<<ComboboxSelected>>", self.load_current_config)

        self.dhcp_var = tk.BooleanVar()
        self.dhcp_checkbox = ttk.Checkbutton(master, text="使用 DHCP", variable=self.dhcp_var, command=self.toggle_dhcp)
        self.dhcp_checkbox.pack(pady=10)

        self.static_frame = ttk.LabelFrame(master, text="静态 IP 设置")
        self.static_frame.pack(pady=10, padx=10, fill="x")

        self.ip_label = ttk.Label(self.static_frame, text="IP 地址:")
        self.ip_label.grid(row=0, column=0, sticky="w", pady=2)
        self.ip_entry = ttk.Entry(self.static_frame)
        self.ip_entry.grid(row=0, column=1, pady=2)

        self.subnet_label = ttk.Label(self.static_frame, text="子网掩码:")
        self.subnet_label.grid(row=1, column=0, sticky="w", pady=2)
        self.subnet_entry = ttk.Entry(self.static_frame)
        self.subnet_entry.grid(row=1, column=1, pady=2)

        self.gateway_label = ttk.Label(self.static_frame, text="默认网关:")
        self.gateway_label.grid(row=2, column=0, sticky="w", pady=2)
        self.gateway_entry = ttk.Entry(self.static_frame)
        self.gateway_entry.grid(row=2, column=1, pady=2)

        self.dns_label = ttk.Label(self.static_frame, text="DNS 服务器:")
        self.dns_label.grid(row=3, column=0, sticky="w", pady=2)
        self.dns_entry = ttk.Entry(self.static_frame)
        self.dns_entry.grid(row=3, column=1, pady=2)

        self.save_button = ttk.Button(master, text="保存设置", command=self.save_settings)
        self.save_button.pack(pady=10)

        self.apply_button = ttk.Button(master, text="应用设置", command=self.apply_settings)
        self.apply_button.pack(pady=10)

        self.load_current_config()

    def load_interfaces(self):
        active_interfaces = [iface for iface, addrs in psutil.net_if_addrs().items() 
                             if any(addr.family == 2 for addr in addrs)]
        
        # 将以太网接口移到列表的开头
        ethernet_interfaces = [iface for iface in active_interfaces if "以太网" in iface or "Ethernet" in iface]
        other_interfaces = [iface for iface in active_interfaces if iface not in ethernet_interfaces]
        sorted_interfaces = ethernet_interfaces + other_interfaces

        if not sorted_interfaces:
            messagebox.showwarning("警告", "未找到活动的网络接口")
            return
        
        self.interface_combobox['values'] = sorted_interfaces
        self.interface_combobox.set(sorted_interfaces[0])

    def load_current_config(self, event=None):
        interface = self.interface_combobox.get()
        config = self.get_network_config(interface)
        
        self.dhcp_var.set(config['is_dhcp'])
        self.ip_entry.delete(0, tk.END)
        self.ip_entry.insert(0, config['ip'])
        self.subnet_entry.delete(0, tk.END)
        self.subnet_entry.insert(0, config['subnet'])
        self.gateway_entry.delete(0, tk.END)
        self.gateway_entry.insert(0, config['gateway'])
        self.dns_entry.delete(0, tk.END)
        self.dns_entry.insert(0, config['dns'])
        
        # 如果不是 DHCP，确保静态 IP 设置可编辑
        if not config['is_dhcp']:
            for child in self.static_frame.winfo_children():
                if isinstance(child, ttk.Entry):
                    child.configure(state="normal")
        else:
            for child in self.static_frame.winfo_children():
                if isinstance(child, ttk.Entry):
                    child.configure(state="disabled")

    def get_network_config(self, interface):
        config = {'is_dhcp': True, 'ip': '', 'subnet': '', 'gateway': '', 'dns': ''}
        
        # 获取系统默认编码
        system_encoding = locale.getpreferredencoding()
        
        # 获取 IP 配置
        try:
            output = subprocess.check_output(f'netsh interface ip show config "{interface}"', shell=True).decode(system_encoding, errors='replace')
        except subprocess.CalledProcessError:
            messagebox.showerror("错误", f"无法获取接口 '{interface}' 的配置信息")
            return config

        # 更精确地判断是否使用 DHCP
        if "DHCP 已启用:                          是" not in output:
            config['is_dhcp'] = False
        
        ip_match = re.search(r'IP 地址:\s+(\d+\.\d+\.\d+\.\d+)', output)
        subnet_match = re.search(r'子网前缀:\s+(\d+\.\d+\.\d+\.\d+)/\d+\s+\(掩码\s+(\d+\.\d+\.\d+\.\d+)\)', output)
        gateway_match = re.search(r'默认网关:\s+(\d+\.\d+\.\d+\.\d+)', output)
        
        if ip_match:
            config['ip'] = ip_match.group(1)
        if subnet_match:
            config['subnet'] = subnet_match.group(2)  # 获取掩码部分
        if gateway_match:
            config['gateway'] = gateway_match.group(1)

        # 获取 DNS 配置
        try:
            dns_output = subprocess.check_output(f'netsh interface ip show dns "{interface}"', shell=True).decode(system_encoding, errors='replace')
            dns_matches = re.findall(r'静态配置的 DNS 服务器:\s+(\d+\.\d+\.\d+\.\d+)', dns_output)
            if dns_matches:
                config['dns'] = dns_matches[0]  # 使用第一个 DNS 服务器
        except subprocess.CalledProcessError:
            messagebox.showwarning("警告", f"无法获取接口 '{interface}' 的 DNS 信息")

        return config

    def toggle_dhcp(self):
        is_dhcp = self.dhcp_var.get()
        state = "disabled" if is_dhcp else "normal"
        for child in self.static_frame.winfo_children():
            if isinstance(child, ttk.Entry):
                child.configure(state=state)

    def save_settings(self):
        # 保存设置到一个 JSON 文件
        config = {
            'interface': self.interface_combobox.get(),
            'is_dhcp': self.dhcp_var.get(),
            'ip': self.ip_entry.get(),
            'subnet': self.subnet_entry.get(),
            'gateway': self.gateway_entry.get(),
            'dns': self.dns_entry.get()
        }
        with open('network_config.json', 'w') as f:
            json.dump(config, f)
        messagebox.showinfo("保存", "设置已保存")

    def load_settings(self):
        # 从 JSON 文件加载设置
        try:
            with open('network_config.json', 'r') as f:
                config = json.load(f)
                self.interface_combobox.set(config['interface'])
                self.dhcp_var.set(config['is_dhcp'])
                self.ip_entry.delete(0, tk.END)
                self.ip_entry.insert(0, config['ip'])
                self.subnet_entry.delete(0, tk.END)
                self.subnet_entry.insert(0, config['subnet'])
                self.gateway_entry.delete(0, tk.END)
                self.gateway_entry.insert(0, config['gateway'])
                self.dns_entry.delete(0, tk.END)
                self.dns_entry.insert(0, config['dns'])
                self.toggle_dhcp()
        except FileNotFoundError:
            messagebox.showwarning("警告", "未找到配置文件，使用默认设置。")

    def apply_settings(self):
        # 应用当前设置到网络接口
        interface = self.interface_combobox.get()
        is_dhcp = self.dhcp_var.get()

        if is_dhcp:
            self.set_dhcp(interface)
        else:
            ip = self.ip_entry.get()
            subnet = self.subnet_entry.get()
            gateway = self.gateway_entry.get()
            dns = self.dns_entry.get()
            self.set_static_ip(interface, ip, subnet, gateway, dns)

        # 在应用设置后自动保存配置
        self.save_settings()  # 添加这一行

    def set_dhcp(self, interface):
        commands = [
            f'netsh interface ip set address "{interface}" dhcp',
            f'netsh interface ip set dns "{interface}" dhcp'
        ]
        for command in commands:
            subprocess.run(command, shell=True)
        messagebox.showinfo("成功", f"{interface} 已设置为使用 DHCP")

    def set_static_ip(self, interface, ip, subnet, gateway, dns):
        commands = [
            f'netsh interface ip set address "{interface}" static {ip} {subnet} {gateway}',
            f'netsh interface ip set dns "{interface}" static {dns}'
        ]
        for command in commands:
            subprocess.run(command, shell=True)
        messagebox.showinfo("成功", f"{interface} 已设置为静态 IP")

if __name__ == "__main__":
    # 请求管理员权限
    if ctypes.windll.shell32.IsUserAnAdmin():
        root = tk.Tk()
        app = DHCPSwitcher(root)
        app.load_settings()  # 在启动时加载设置
        root.mainloop()
    else:
        # 以管理员身份重新启动程序
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)