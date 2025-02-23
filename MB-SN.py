import random
import string
import subprocess
import uuid
import os
import time
import logging

logging.basicConfig(
    level=logging.DEBUG,
    format="[%(levelname)s] %(asctime)s - %(message)s"
)

def debug_info(message):
    logging.debug(f"[+] {message}")

def generate_amide_commands(value):
    amide_path = "AMIDEWINx64.EXE" if os.environ.get('PROCESSOR_ARCHITECTURE') == 'AMD64' else "AMIDEWIN.EXE"
    debug_info(f"检测到系统架构: {amide_path}")

    command_params = [
        ('IVN', value), ('IV', value), ('ID', value),
        ('SS', value), ('SK', value), ('SF', value),
        ('SU', 'AUTO'), 
        ('BM', value), ('BV', value), ('BS', value),
        ('BT', value), ('BLC', value), ('CM', value),
        ('CV', value), ('CA', value), ('CS', value),
        ('CSK', value), ('PPN', value), ('PAT', value),
        ('PSN', value)
    ]

    commands = []
    for param, value in command_params:
        cmd = f"{amide_path} /{param} {value}"
        commands.append(cmd)
        debug_info(f"生成命令: {cmd}")
    
    return commands

def execute_commands(commands):
    for cmd in commands:
        debug_info(f"执行命令: {cmd}")
        try:
            result = subprocess.run(cmd, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=10)
            debug_info(f"命令输出: {result.stdout.decode()}")
        except subprocess.CalledProcessError as e:
            logging.error(f"命令执行失败: {cmd}\n错误码: {e.returncode}\n错误输出: {e.stderr.decode()}")
            raise

def kill_wmiprse():
    max_retries = 4
    delay = 2
    
    for attempt in range(max_retries):
        try:
            logging.info(f"尝试终止 WmiPrvSE.exe (第 {attempt+1}/{max_retries} 次)")
            subprocess.run("taskkill /F /IM WmiPrvSE.exe", shell=True, check=True, timeout=5)
            logging.success("进程终止成功")
            return
        except Exception as e:
            logging.warning(f"终止进程失败: {e}", exc_info=True)
            if attempt < max_retries - 1:
                time.sleep(delay)
    logging.error("所有重试均失败，进程可能仍在运行")

def main():
    #uuid_value = str(uuid.uuid4())
    #debug_info(f"生成UUID: {uuid_value}")
    
    serial_chars = string.ascii_uppercase + string.digits
    serial_number = ''.join(random.choices(serial_chars, k=16))
    debug_info(f"生成序列号: {serial_number}")
    
    input("Ready to proceed... (按回车继续)")

    commands = generate_amide_commands(serial_number)
    execute_commands(commands)

    kill_wmiprse()

    input("\n脚本执行完成，按回车退出")
    logging.info("脚本正常终止")

if __name__ == "__main__":
    main()
