import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import psutil
import socket
import os
import json
import ctypes
import sys

CONFIG_FILE = "ip_config.conf"  # ชื่อไฟล์ config ใหม่ที่ปรับให้กระชับขึ้น

# ฟังก์ชันแปลง CIDR เป็น Subnet Mask


def cidr_to_subnet_mask(cidr):
    try:
        cidr = int(cidr)
        mask = (0xffffffff >> (32 - cidr)) << (32 - cidr)
        return f"{(mask >> 24) & 0xff}.{(mask >> 16) & 0xff}.{(mask >> 8) & 0xff}.{mask & 0xff}"
    except ValueError:
        return None

# ฟังก์ชันตรวจสอบว่าโปรแกรมรันด้วยสิทธิ์ Admin หรือไม่


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


# ถ้าไม่ได้รันในโหมด Admin จะเปิดโปรแกรมใหม่พร้อมสิทธิ์ Admin
if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

# ฟังก์ชันตรวจสอบสถานะของ Network Adapter


def is_adapter_connected(adapter):
    try:
        result = subprocess.check_output(f'netsh interface show interface "{
                                         adapter}"', shell=True).decode()
        for line in result.splitlines():
            if "Connected" in line:
                return True
        return False
    except subprocess.CalledProcessError:
        return False

# ฟังก์ชันสำหรับดึงข้อมูลเครือข่าย (IP, Subnet, Gateway, DNS)


def get_network_info(adapter):
    try:
        result = subprocess.check_output(f'netsh interface ip show config name="{
                                         adapter}"', shell=True).decode()

        info = {
            "ip_address": None,
            "subnet_mask": None,
            "gateway": None,
            "dns1": None,
            "dns2": None
        }

        # แยกบรรทัดและดึงข้อมูล IP, Subnet, Gateway, DNS
        for line in result.splitlines():
            if "IP Address" in line and "Subnet Prefix" not in line:  # ตรวจ IP Address ที่ไม่ใช่ Subnet Prefix
                info["ip_address"] = line.split(":")[-1].strip()
            elif "Subnet Prefix" in line:  # ตรวจ Subnet Prefix เพื่อดึง Subnet Mask
                subnet_info = line.split()[2]
                cidr = subnet_info.split("/")[1]  # ดึงเฉพาะส่วน CIDR เช่น /25
                info["subnet_mask"] = cidr_to_subnet_mask(
                    cidr)  # แปลง CIDR เป็น Subnet Mask
            elif "Default Gateway" in line:
                info["gateway"] = line.split(":")[-1].strip()
            elif "DNS Servers" in line:
                info["dns1"] = line.split(":")[-1].strip()
            elif "Register with which" not in line and info["dns1"] and not info["dns2"]:
                info["dns2"] = line.strip()

        return info
    except subprocess.CalledProcessError:
        messagebox.showerror(
            "Error", f"Could not retrieve network info for {adapter}")
        return None

# ฟังก์ชันสำหรับบันทึกการตั้งค่าในไฟล์ config


def save_config():
    config = {}
    for adapter in adapters:  # ใช้ adapters ที่แสดงใน Dropdown เท่านั้น
        info = get_network_info(adapter)
        if info:
            config[adapter] = info

    with open(CONFIG_FILE, "w") as config_file:
        json.dump(config, config_file)

    messagebox.showinfo("Success", "Configuration saved successfully.")

# ฟังก์ชันสำหรับโหลดค่าจากไฟล์ config


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as config_file:
            config = json.load(config_file)
        return config
    return None

# ฟังก์ชันเมื่อเปลี่ยน Adapter


def on_adapter_change(event):
    selected_adapter = interface_var.get()
    if loaded_config and selected_adapter in loaded_config:
        info = loaded_config[selected_adapter]
    else:
        info = get_network_info(selected_adapter)

    if info:
        ip_entry.delete(0, tk.END)
        subnet_entry.delete(0, tk.END)
        gateway_entry.delete(0, tk.END)
        dns1_entry.delete(0, tk.END)
        dns2_entry.delete(0, tk.END)

        ip_entry.insert(0, info["ip_address"] if info["ip_address"] else "")
        subnet_entry.insert(0, info["subnet_mask"]
                            if info["subnet_mask"] else "")
        gateway_entry.insert(0, info["gateway"] if info["gateway"] else "")
        dns1_entry.insert(0, info["dns1"] if info["dns1"] else "")
        dns2_entry.insert(0, info["dns2"] if info["dns2"] else "")

# ฟังก์ชันสำหรับตั้งค่า Fixed IP (ภายใน บริษัท)


def set_fixed_ip():
    interface = interface_var.get()
    ip_address = ip_entry.get()
    subnet_mask = subnet_entry.get()
    gateway = gateway_entry.get()
    dns1 = dns1_entry.get()
    dns2 = dns2_entry.get()

    if not interface or not ip_address:
        messagebox.showerror("Error", "Please fill all fields")
        return

    try:
        subprocess.run(f'netsh interface ip set address "{interface}" static {ip_address} {
                       subnet_mask} {gateway}', shell=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        subprocess.run(f'netsh interface ip set dns "{interface}" static {
                       dns1}', shell=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        subprocess.run(f'netsh interface ip add dns "{interface}" {
                       dns2} index=2', shell=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        messagebox.showinfo("Success", f"Fixed IP set on {interface}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to set IP: {e}")

# ฟังก์ชันสำหรับตั้งค่า DHCP (ภายนอก บริษัท)


def set_dhcp():
    interface = interface_var.get()
    if not interface:
        messagebox.showerror("Error", "Please select an interface")
        return

    try:
        subprocess.run(f'netsh interface ip set address "{
                       interface}" dhcp', shell=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        subprocess.run(f'netsh interface ip set dns "{
                       interface}" dhcp', shell=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        messagebox.showinfo("Success", f"DHCP set on {interface}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to set DHCP: {e}")

# ฟังก์ชันสำหรับดึงรายชื่อ Network Adapters เฉพาะที่เชื่อมต่ออยู่เท่านั้น


def get_network_adapters():
    adapters = psutil.net_if_addrs()
    connected_adapters = []
    for adapter in adapters.keys():
        if is_adapter_connected(adapter):  # ตรวจเฉพาะ Adapter ที่เชื่อมต่อ
            connected_adapters.append(adapter)
    return connected_adapters


# GUI Layout
root = tk.Tk()
root.title("IP Configurator")

# ดึงรายชื่อ Network Adapters ที่เชื่อมต่อ และตั้งค่า Dropdown
tk.Label(root, text="Network Adapter:").grid(row=1, column=0, padx=10, pady=5)
adapters = get_network_adapters()
interface_var = tk.StringVar()
interface_dropdown = ttk.Combobox(
    root, textvariable=interface_var, values=adapters)
interface_dropdown.grid(row=1, column=1, padx=10, pady=5)
if adapters:
    interface_dropdown.current(0)  # ตั้งค่าให้เลือกค่าแรกโดยอัตโนมัติ

# ตั้งค่า IP Address ของเครื่องเองโดยอัตโนมัติเมื่อเปลี่ยน Adapter
tk.Label(root, text="IP Address:").grid(row=2, column=0, padx=10, pady=5)
ip_entry = tk.Entry(root)
ip_entry.grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Subnet Mask:").grid(row=3, column=0, padx=10, pady=5)
subnet_entry = tk.Entry(root)
subnet_entry.grid(row=3, column=1, padx=10, pady=5)

tk.Label(root, text="Default Gateway:").grid(row=4, column=0, padx=10, pady=5)
gateway_entry = tk.Entry(root)
gateway_entry.grid(row=4, column=1, padx=10, pady=5)

tk.Label(root, text="DNS Server 1:").grid(row=5, column=0, padx=10, pady=5)
dns1_entry = tk.Entry(root)
dns1_entry.grid(row=5, column=1, padx=10, pady=5)

tk.Label(root, text="DNS Server 2:").grid(row=6, column=0, padx=10, pady=5)
dns2_entry = tk.Entry(root)
dns2_entry.grid(row=6, column=1, padx=10, pady=5)

# ปุ่ม Save Config อยู่ด้านบนแบบเต็มความกว้าง พร้อมสีเขียว lightgreen
tk.Button(root, text="Save Config", command=save_config, bg="lightgreen",
          fg="black").grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

# ปุ่ม Fixed IP และ DHCP พร้อมคำอธิบาย แยกกันคนละบรรทัด
fixed_ip_btn = tk.Button(root, text="Set Fixed IP (ภายใน บริษัท)",
                         command=set_fixed_ip, bg="lightblue", fg="black")
dhcp_btn = tk.Button(root, text="Set DHCP (ภายนอก บริษัท)",
                     command=set_dhcp, bg="lightyellow", fg="black")

# วางปุ่ม Fixed IP และ DHCP แยกกันคนละบรรทัด
fixed_ip_btn.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
dhcp_btn.grid(row=8, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

# ใส่ข้อความ Copyright ด้านล่างสุดของหน้าต่าง
tk.Label(root, text="© 2024 : IP Configurator by IT Section/LSD", font=("Arial", 8),
         fg="gray").grid(row=9, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

# โหลด config ถ้ามีอยู่
loaded_config = load_config()

if loaded_config:
    fixed_ip_btn.grid()
    dhcp_btn.grid()
else:
    fixed_ip_btn.grid_remove()
    dhcp_btn.grid_remove()

# ผูก event เมื่อมีการเปลี่ยนแปลง Network Adapter
interface_dropdown.bind("<<ComboboxSelected>>", on_adapter_change)

# เรียก on_adapter_change เมื่อเปิดโปรแกรมครั้งแรก ถ้ามี adapter เชื่อมต่อ
if adapters:
    on_adapter_change(None)

root.mainloop()
