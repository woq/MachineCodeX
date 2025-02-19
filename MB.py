import random
import string
import subprocess
import uuid
import os

# 调试信息
def debug_info(message):
    print(f"[DEBUG] {message}")

# 生成随机 UUID
def generate_uuid():
    # 生成一个随机 UUID
    random_uuid = str(uuid.uuid4())
    debug_info(f"生成的随机 UUID: {random_uuid}")
    return random_uuid

# 生成随机序列号（16位字母和数字）
def generate_serial_number():
    # 生成一个随机的 16 位序列号
    characters = string.ascii_uppercase + string.digits
    serial_number = ''.join(random.choice(characters) for _ in range(16))
    debug_info(f"生成的随机 SerialNumber: {serial_number}")
    return serial_number

# 执行 AMIDE 命令（模拟命令行）
def run_amide_command(uuid_value):
    amide_path = "AMIDEWINx64.EXE" if os.environ['PROCESSOR_ARCHITECTURE'] == 'AMD64' else "AMIDEWIN.EXE"
    debug_info(f"选择的 AMIDE 可执行文件: {amide_path}")

    commands = [
        f"{amide_path} /IVN {uuid_value}",
        f"{amide_path} /IV {uuid_value}",
        f"{amide_path} /ID {uuid_value}",
        f"{amide_path} /SS {uuid_value}",
        f"{amide_path} /SK {uuid_value}",
        f"{amide_path} /SF {uuid_value}",
        f"{amide_path} /SU AUTO",
        f"{amide_path} /BM {uuid_value}",
        f"{amide_path} /BV {uuid_value}",
        f"{amide_path} /BS {uuid_value}",
        f"{amide_path} /BT {uuid_value}",
        f"{amide_path} /BLC {uuid_value}",
        f"{amide_path} /CM {uuid_value}",
        f"{amide_path} /CV {uuid_value}",
        f"{amide_path} /CA {uuid_value}",
        f"{amide_path} /CS {uuid_value}",
        f"{amide_path} /CSK {uuid_value}",
        f"{amide_path} /PPN {uuid_value}",
        f"{amide_path} /PAT {uuid_value}",
        f"{amide_path} /PSN {uuid_value}"
    ]
    
    # 执行每个命令
    for command in commands:
        debug_info(f"执行命令: {command}")
        subprocess.run(command, shell=True)

# 终止 WmiPrvSE.exe 进程
def kill_wmiprse():
    for i in range(4):
        debug_info(f"终止 WmiPrvSE.exe 进程 (尝试 {i+1}/4)")
        subprocess.run("taskkill /F /IM WmiPrvSE.exe", shell=True)

# 主函数
def main():
    # 生成随机 UUID 和序列号
    uuid_value = generate_uuid()
    serial_number = generate_serial_number()
    input("Ready to Go")
    # 执行 AMIDE 命令
    run_amide_command(uuid_value)

    # 终止进程
    kill_wmiprse()

    debug_info("脚本执行完毕。")
    input("Done")

if __name__ == "__main__":
    main()
