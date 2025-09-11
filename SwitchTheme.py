import datetime
import subprocess
import os
import json
import ctypes
import time
import sys
import scheduler

# 说明
# 有两种执行方式：
# 1. 直接执行，则按照当前时区的日出日落时间切换主题
# 2. 带参数执行，则按照参数切换指定主题。参数格式：--mode light/dark

# 获取当前用户名
current_user = os.getlogin()

# 读取 datetime.json 文件
while True:
    try:
        with open(r'datetime.json', 'r') as f:
            data = json.load(f)
            break
    except:
        print("无法读取 datetime.json 文件")
        scheduler.outTimefile()

# 获取当前时间（带本地时区信息）
current_time = datetime.datetime.now(datetime.timezone.utc).astimezone()

print('当前时间：', current_time)
print('当前用户：',current_user)

# 根据当前用户名构建主题文件路径
# windows 的默认主题路径：C:\Windows\Resources\Themes
light_theme_path = r"C:\\Windows\\resources\\Themes\\aero.theme"
dark_theme_path  = r"C:\\Windows\\resources\\Themes\\dark.theme"

# 用户自定义的主题路径：%homepath%\AppData\Local\Microsoft\Windows\Themes
# light_theme_path = f'C:\\Users\\{current_user}\\AppData\\Local\\Microsoft\\Windows\\Themes\\自定义-浅色.theme'
# dark_theme_path  = f'C:\\Users\\{current_user}\\AppData\\Local\\Microsoft\\Windows\\Themes\\自定义-深色.theme'

# 获取日出和日落时间（转换为本地时区）
sunrise_time = datetime.datetime.fromisoformat(data['results']['sunrise']).astimezone()
sunset_time = datetime.datetime.fromisoformat(data['results']['sunset']).astimezone()

sunrise_time += datetime.timedelta(minutes=30)  # 日出时间延迟30分钟
sunset_time -= datetime.timedelta(minutes=30)   # 日落时间提前30分钟

# 提取时间部分（忽略日期）
sunrise_time = sunrise_time.time()
sunset_time = sunset_time.time()

print('日出时间：',sunrise_time)
print('日落时间：',sunset_time)
# 调用自动广播系统设置变更通知（防止出现主题切换可能不完整的情况）
def broadcast_setting_change():
    HWND_BROADCAST = 0xFFFF
    WM_SETTINGCHANGE = 0x001A
    SMTO_ABORTIFHUNG = 0x0002

    ctypes.windll.user32.SendMessageTimeoutW(
        HWND_BROADCAST,
        WM_SETTINGCHANGE,
        0,
        "Environment",
        SMTO_ABORTIFHUNG,
        5000,
        ctypes.byref(ctypes.c_ulong())
    )

# 当前主题和期望主题路径
def theme_contrast():
    # 获取当前时间（带时区），决定预期模式
    current_time = datetime.datetime.now(datetime.timezone.utc).astimezone()
    current_time_only = current_time.time()
    expected_mode = "light" if (sunrise_time <= current_time_only < sunset_time) else "dark"

    # 查询注册表当前模式
    result = subprocess.run(
        'reg query HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize /v AppsUseLightTheme',
        shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
    )

    current_mode = None
    if "0x1" in result.stdout:   # 1 = 浅色
        current_mode = "light"
    elif "0x0" in result.stdout: # 0 = 深色
        current_mode = "dark"

    if not current_mode:
        with open('err.log','a',encoding="utf-8") as f:
            f.write(f'{datetime.datetime.now()} - 无法解析当前主题模式，注册表返回：{result.stdout}')
        print("无法解析当前主题模式，注册表返回：")
        print(result.stdout)
        raise RuntimeError("未能获取 AppsUseLightTheme 的值")
        

    print("当前模式:", current_mode)
    print("期望模式:", expected_mode)

    # 返回模式对应的主题路径
    expected_theme = light_theme_path if expected_mode == "light" else dark_theme_path
    current_theme = light_theme_path if current_mode == "light" else dark_theme_path

    return current_theme, expected_theme


def main(themes_path):
    i = 0
    while True:
        if theme_contrast()[0] == themes_path:
            print("当前主题已符合要求,跳过更改")
            break
        else:
            print(f"正在切换主题为: {themes_path.split('\\')[-1]}")
            broadcast_setting_change()
            subprocess.run(f'start "" "{themes_path}"', shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            broadcast_setting_change()
            i += 1
            if i >= 10:
                error_msg = f"[ERROR] 超过 {i} 次仍未能切换到 {themes_path}\n"
                with open("err.log", "a", encoding="utf-8") as f:
                    f.write(f"{datetime.datetime.now()} - {error_msg}")
                print(error_msg.strip())
                break
    kill_settings_panel()

# 检测设置面板是否已经关闭
def kill_settings_panel():
    print("正在检测设置面板是否已经关闭...")
    for _ in range(20):
        result = subprocess.run('tasklist /fi "IMAGENAME eq SystemSettings.exe"', shell=True,stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if "systemsettings.exe" in result.stdout.lower():
            print("检测到设置面板，将其关闭…")
            subprocess.run('taskkill /f /im "SystemSettings.exe"', shell=True,stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        time.sleep(0.1)

if __name__=='__main__':

    if len(sys.argv) == 1:
        print("参数为空,按照时间条件执行")
        main(theme_contrast()[1])
    else:
        if "--mode" in sys.argv:
            try:
                mode_index = sys.argv.index("--mode") + 1
                mode_value = sys.argv[mode_index]
                if mode_value == "light":
                    print("使用参数指定更改到浅色主题")
                    main(light_theme_path)
                elif mode_value == "dark":
                    print("使用参数指定更改到深色主题")
                    main(dark_theme_path)
                else:
                    print('参数错误')
            except IndexError:
                print("未指定 mode 的值")
        else:
            print("未识别的参数")

    sys.exit(0)