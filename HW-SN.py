import wmi
import os
import datetime
import subprocess
import sys  # Import sys for terminal detection

# ANSI color codes
COLOR_RESET = '\033[0m'
COLOR_RED = '\033[91m'
COLOR_GREEN = '\033[92m'
COLOR_YELLOW = '\033[93m'
COLOR_BLUE = '\033[94m'
COLOR_PURPLE = '\033[95m'
COLOR_CYAN = '\033[96m'
COLOR_GRAY = '\033[97m'

def colorize_text(text, color_code):
    """Conditionally colorizes text for terminal output only."""
    if sys.stdout.isatty():  # Check if writing to a terminal
        return color_code + text + COLOR_RESET
    else:
        return text  # Return plain text if not terminal

def get_gateway_mac(gateway_ip):
    """通过 arp -a 命令获取网关 MAC 地址，并格式化为冒号分隔"""
    try:
        arp_output = subprocess.check_output(['arp', '-a', gateway_ip], text=True, stderr=subprocess.DEVNULL)
        for line in arp_output.strip().split('\n'):
            if gateway_ip in line:
                parts = line.split()
                if len(parts) >= 2:
                    mac_address_hyphen = parts[1]
                    # 尝试将 MAC 地址从 hyphenated 格式转换为 colon 格式
                    mac_address_colon = mac_address_hyphen.replace('-', ':')
                    return mac_address_colon
    except Exception:
        return "N/A"
    return "N/A"

def get_gateway_ip_fallback():
    """使用 ipconfig /all 命令作为备选方案获取默认网关 IP"""
    try:
        ipconfig_output = subprocess.check_output(['ipconfig', '/all'], text=True, stderr=subprocess.DEVNULL, encoding='gbk') # Use gbk encoding for ipconfig output
        for line in ipconfig_output.splitlines():
            if "默认网关" in line:
                parts = line.split(":")
                if len(parts) > 1:
                    ip_address = parts[1].strip()
                    if ip_address and ip_address != "0.0.0.0": # Filter out 0.0.0.0 gateway
                        return ip_address
    except Exception:
        return "N/A"
    return "N/A"


def get_nvidia_gpu_info():
    """尝试使用 nvidia-smi -L 获取 NVIDIA 显卡信息 (包括 UUID)"""
    try:
        nvidia_smi_output = subprocess.check_output(['nvidia-smi', '-L'], text=True, stderr=subprocess.DEVNULL)
        gpu_info_list = []
        for line in nvidia_smi_output.strip().split('\n'):
            if line.startswith('GPU '):
                gpu_info_list.append(line.strip()) # Store the whole line as it is in BAT
        return gpu_info_list
    except FileNotFoundError:
        return ["nvidia-smi_not_found"]
    except subprocess.CalledProcessError:
        return ["nvidia_driver_error"]
    except Exception:
        return ["nvidia_gpu_error"]
    return ["N/A"]

def get_wmi_info(wmi_client, wmi_class, properties):
    """通用 WMI 信息获取函数，接受 wmi_client 参数, 抑制 Win32_BIOS 错误"""
    try:
        wmi_obj_list = wmi_client.query(f"SELECT {','.join(properties)} FROM {wmi_class}") # 使用 wmi_client
        if wmi_obj_list and wmi_obj_list[0]:
            return wmi_obj_list[0]
        else:
            return None
    except Exception as e:
        if wmi_class != "Win32_BIOS": # 抑制 Win32_BIOS 错误输出
            print(f"WMI Query Error for class '{wmi_class}' and properties '{properties}': {e}")
        return None


def create_hardware_report():
    # 创建保存目录
    os.makedirs('sn', exist_ok=True)

    # 获取当前时间戳
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join('sn', f'sn_{timestamp}.txt')

    # 初始化WMI对象
    c = wmi.WMI()

    report_data = {} # 使用字典存储数据

    report_data['cpu'] = {}
    # CPU 信息
    try:
        cpu_info = []
        for cpu in c.Win32_Processor():
            cpu_data = {}
            cpu_data['model'] = cpu.Name
            cpu_data['serial_number'] = cpu.ProcessorId or "N/A"
            cpu_info.append(cpu_data)
        report_data['cpu']['instances'] = cpu_info
    except:
        report_data['cpu']['instances'] = []

    report_data['motherboard'] = {}
    # 主板信息
    try:
        mb_data = {}
        baseboard_info = get_wmi_info(c, "Win32_BaseBoard", ["Product", "SerialNumber"])
        if baseboard_info:
            mb_data['model'] = baseboard_info.Product or "N/A"
            mb_data['serial_number'] = baseboard_info.SerialNumber or "N/A"
        else:
            mb_data['model'] = "Failed_to_get"
            mb_data['serial_number'] = "Failed_to_get"

        # BIOS Info - Removed BIOS SerialNumber and UUID retrieval

        csproduct_info = get_wmi_info(c, "Win32_ComputerSystemProduct", ["UUID"])
        if csproduct_info:
            mb_data['system_uuid'] = csproduct_info.UUID or "N/A"
        else:
            mb_data['system_uuid'] = "Failed_to_get"

        report_data['motherboard'] = mb_data
    except:
        report_data['motherboard'] = {}

    report_data['memory'] = {}
    # 内存信息
    try:
        mem_info = []
        for mem in c.Win32_PhysicalMemory():
            mem_data = {}
            mem_data['serial_number'] = mem.SerialNumber or "N/A"
            mem_info.append(mem_data)
        report_data['memory']['instances'] = mem_info
    except:
        report_data['memory']['instances'] = []

    report_data['disk'] = {}
    # 硬盘信息
    try:
        disk_info = []
        for disk in c.Win32_DiskDrive():
            disk_data = {}
            disk_data['model'] = disk.Model or "N/A"
            disk_data['serial_number'] = disk.SerialNumber.strip() if disk.SerialNumber else "N/A"
            disk_info.append(disk_data)
        report_data['disk']['instances'] = disk_info
    except:
        report_data['disk']['instances'] = []

    # Graphics Card Info - Removed

    report_data['network_adapter'] = {}
    # 网卡信息 (物理网卡，包含网关)
    try:
        nic_info = []
        for nic in c.Win32_NetworkAdapter():
            if nic.NetConnectionStatus == 2 and nic.AdapterTypeID == 0 and "PCI" in str(nic.PNPDeviceID): # 已连接的物理网卡
                nic_data = {}
                nic_data['name'] = nic.NetConnectionID or nic.Name or "N/A"
                nic_data['mac_address'] = nic.MACAddress or "N/A"
                nic_info.append(nic_data)
        report_data['network_adapter']['instances'] = nic_info

        # 获取默认网关信息 (IP 和 MAC)
        gateway_ip = "Failed_to_get"
        gateway_mac = "Failed_to_get"
        try:
            for gateway_config in c.Win32_NetworkAdapterConfiguration(DefaultIPGateway__like="%"):
                if gateway_config.DefaultIPGateway:
                    gateway_ip = gateway_config.DefaultIPGateway[0] # 假设只有一个网关
                    gateway_mac = get_gateway_mac(gateway_ip)
                    break # 找到网关配置就停止
        except:
            gateway_ip = "Failed_to_get_WMI" # WMI Failed
            gateway_mac = "Failed_to_get_WMI"

        if gateway_ip == "Failed_to_get_WMI" or gateway_ip == "Failed_to_get": # If WMI or initial method failed for IP, try fallback
            fallback_ip = get_gateway_ip_fallback()
            if fallback_ip and fallback_ip != "N/A":
                gateway_ip = fallback_ip
                gateway_mac = get_gateway_mac(gateway_ip) # Try to get MAC again with fallback IP
            else:
                gateway_ip = "Failed_to_get_Fallback" # Fallback also failed
                gateway_mac = "Failed_to_get_Fallback"


        report_data['network_adapter']['default_gateway_ip'] = gateway_ip
        report_data['network_adapter']['default_gateway_mac'] = gateway_mac


    except:
        report_data['network_adapter']['instances'] = []
        report_data['network_adapter']['default_gateway_ip'] = "Failed_to_get_NIC_info"
        report_data['network_adapter']['default_gateway_mac'] = "Failed_to_get_NIC_info"


    # Bluetooth Info - Removed

    nvidia_gpu_list = get_nvidia_gpu_info() # Get NVIDIA GPU info in BAT style

    # 写入文件 (代码友好和人类可读的格式) 并同时输出到屏幕
    output_lines = []
    output_lines.append(colorize_text("=" * 50, COLOR_GRAY))
    output_lines.append(colorize_text("     2049", COLOR_GRAY)) # Changed title
    output_lines.append(colorize_text("     硬件信息报告", COLOR_GRAY))
    output_lines.append(colorize_text("=" * 50, COLOR_GRAY))
    output_lines.append("")

    output_lines.append(colorize_text("[CPU信息]", COLOR_RED))
    if report_data['cpu']['instances']:
        for i, cpu_data in enumerate(report_data['cpu']['instances']):
            output_lines.append(f"  实例 {colorize_text(str(i+1), COLOR_GRAY)}:") # Color instance number
            output_lines.append(f"    型号: {colorize_text(cpu_data.get('model', 'N/A'), COLOR_CYAN)}") # Color model
            output_lines.append(f"    序列号: {colorize_text(cpu_data.get('serial_number', 'N/A'), COLOR_YELLOW)}") # Color serial
    else:
        output_lines.append("  没有可用的CPU信息.")
    output_lines.append("")

    output_lines.append(colorize_text("[主板信息]", COLOR_GREEN))
    if report_data['motherboard']:
        mb_data = report_data['motherboard']
        output_lines.append(f"  型号: {colorize_text(mb_data.get('model', 'N/A').replace('Failed_to_get', '获取失败'), COLOR_CYAN)}") # Color model
        output_lines.append(f"  序列号: {colorize_text(mb_data.get('serial_number', 'N/A').replace('Failed_to_get', '获取失败'), COLOR_YELLOW)}") # Color serial
        output_lines.append(f"  系统 UUID: {colorize_text(mb_data.get('system_uuid', 'N/A').replace('Failed_to_get', '获取失败'), COLOR_YELLOW)}") # System UUID
    else:
        output_lines.append("  没有可用的主板信息.")
    output_lines.append("")

    output_lines.append(colorize_text("[内存信息]", COLOR_BLUE))
    if report_data['memory']['instances']:
        for i, mem_data in enumerate(report_data['memory']['instances']):
            output_lines.append(f"  实例 {colorize_text(str(i+1), COLOR_GRAY)}:") # Color instance number
            output_lines.append(f"    序列号: {colorize_text(mem_data.get('serial_number', 'N/A'), COLOR_YELLOW)}") # Color serial
    else:
        output_lines.append("  没有可用的内存信息.")
    output_lines.append("")

    output_lines.append(colorize_text("[硬盘信息]", COLOR_PURPLE))
    if report_data['disk']['instances']:
        for i, disk_data in enumerate(report_data['disk']['instances']):
            output_lines.append(f"  实例 {colorize_text(str(i+1), COLOR_GRAY)}:") # Color instance number
            output_lines.append(f"    型号: {colorize_text(disk_data.get('model', 'N/A'), COLOR_CYAN)}") # Color model
            output_lines.append(f"    序列号: {colorize_text(disk_data.get('serial_number', 'N/A'), COLOR_YELLOW)}") # Color serial
    else:
        output_lines.append("  没有可用的硬盘信息.")
    output_lines.append("")

    output_lines.append(colorize_text("[网卡信息]", COLOR_CYAN))
    if report_data['network_adapter']['instances']:
        for i, nic_data in enumerate(report_data['network_adapter']['instances']):
            output_lines.append(f"  实例 {colorize_text(str(i+1), COLOR_GRAY)}:") # Color instance number
            output_lines.append(f"    名称: {colorize_text(nic_data.get('name', 'N/A'), COLOR_CYAN)}") # Color name
            output_lines.append(f"    MAC 地址: {colorize_text(nic_data.get('mac_address', 'N/A'), COLOR_YELLOW)}") # Color MAC
        gateway_ip = report_data['network_adapter'].get('default_gateway_ip')
        gateway_mac = report_data['network_adapter'].get('default_gateway_mac')
        if gateway_ip and gateway_ip != "Failed_to_get_NIC_info":
            output_lines.append(f"  默认网关 IP: {colorize_text(gateway_ip.replace('Failed_to_get_WMI', '获取失败(WMI)').replace('Failed_to_get_Fallback', '获取失败(备选)') if 'Failed_to_get' in gateway_ip else gateway_ip, COLOR_CYAN)}") # Color Gateway IP
            output_lines.append(f"  默认网关 MAC: {colorize_text(gateway_mac.replace('Failed_to_get_WMI', '获取失败(WMI)').replace('Failed_to_get_Fallback', '获取失败(备选)') if 'Failed_to_get' in gateway_mac else gateway_mac, COLOR_YELLOW)}") # Color Gateway MAC

    else:
        output_lines.append("  没有可用的网卡信息.")
    output_lines.append("")

    # Bluetooth Output Removed

    if nvidia_gpu_list and nvidia_gpu_list[0] not in ["N/A", "nvidia-smi_not_found", "nvidia_driver_error", "nvidia_gpu_error"]:
        output_lines.append(colorize_text("[显卡虚拟序列号 (NVIDIA)]", COLOR_PURPLE)) # Purple for GPU section like in BAT
        for gpu_info in nvidia_gpu_list:
            output_lines.append(f"  {colorize_text(gpu_info, COLOR_CYAN)}") # Color GPU info
        output_lines.append("")


    output_text = "\n".join(output_lines)

    with open(filename, 'w', encoding='utf-8') as f:
        f.write(output_text) # File output is plain text now

    print(output_text) # Terminal output is colored
    print(colorize_text(f"\n硬件信息已保存至：{filename}", COLOR_GRAY))
    input(colorize_text("按任意键结束...", COLOR_GRAY)) # Wait for user input to exit

if __name__ == "__main__":
    if os.name == 'nt': # Enable ANSI color support in Windows cmd if possible
        os.system('')
    create_hardware_report()
