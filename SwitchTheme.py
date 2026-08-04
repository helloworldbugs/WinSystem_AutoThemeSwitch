import datetime
import os
import yaml
import json
import ctypes
import sys
import configparser
import subprocess
import re
import time
import platform
import Scheduler

# 说明
# 1. 直接执行，则按照当前时区的日出日落时间切换主题
# 2. 带参数执行，则按照参数切换指定主题。参数格式：--mode light/dark


# 检测是否 Windows 10（Win10 没有 dark.theme 主题文件，深色模式纯靠注册表）
def is_win10():
    ver = platform.version()
    try:
        build = int(ver.split('.')[2])
        return build < 22000  # Win11 build 从 22000 开始
    except (ValueError, IndexError):
        return False


# 获取当前用户名
current_user = os.getlogin()
print('当前用户：',current_user)

# 读取配置文件获取主题路径和壁纸路径
with open('config.yaml', 'r', encoding='utf-8') as f:
    config = yaml.safe_load(f)
    light_theme_path = config['Theme_path']['light_theme_path']
    dark_theme_path = config['Theme_path']['dark_theme_path']
    light_wallpaper_path = config['Wallpaper_path']['light_wallpaper_path']
    dark_wallpaper_path = config['Wallpaper_path']['dark_wallpaper_path']

    # 记录用户是否配置了自定义主题（在展开和默认值之前）
    user_configured_theme = bool(light_theme_path or dark_theme_path)

    # 展开环境变量（支持 %homepath%、%userprofile% 等）
    light_theme_path = os.path.expandvars(light_theme_path) if light_theme_path else ''
    dark_theme_path = os.path.expandvars(dark_theme_path) if dark_theme_path else ''
    light_wallpaper_path = os.path.expandvars(light_wallpaper_path) if light_wallpaper_path else ''
    dark_wallpaper_path = os.path.expandvars(dark_wallpaper_path) if dark_wallpaper_path else ''

    time_offset = config.get('Time_offset', {})
    sunrise_offset_minutes = int(time_offset.get('sunrise_offset_minutes', 0) or 0)
    sunset_offset_minutes = int(time_offset.get('sunset_offset_minutes', 0) or 0)
    brightness_config = config.get('Brightness', {})
    light_brightness = brightness_config.get('light_brightness')
    dark_brightness = brightness_config.get('dark_brightness')

    # 如果主题路径为空，使用默认路径
    # Win10 没有 dark.theme，深色模式纯靠注册表，主题文件仍用 aero.theme
    if not light_theme_path:
        light_theme_path = r'C:\Windows\Resources\Themes\aero.theme'
    if not dark_theme_path:
        dark_theme_path = r'C:\Windows\Resources\Themes\aero.theme' if is_win10() else r'C:\Windows\Resources\Themes\dark.theme'

if "Users" in light_theme_path and "AppData" in light_theme_path:
    light_theme_path = re.sub(r'(?<=\\Users\\)[^\\]+(?=\\AppData)', current_user, light_theme_path)
if "Users" in dark_theme_path and "AppData" in dark_theme_path:
    dark_theme_path = re.sub(r'(?<=\\Users\\)[^\\]+(?=\\AppData)', current_user, dark_theme_path)

# 读取 datetime.json 文件
while True:
    try:
        with open(r'datetime.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
            break
    except:
        Scheduler.outTimefile()

# 获取日出和日落时间（转换为本地时区）
sunrise_time = datetime.datetime.fromisoformat(data['results']['sunrise']).astimezone()
sunset_time = datetime.datetime.fromisoformat(data['results']['sunset']).astimezone()

sunrise_time -= datetime.timedelta(minutes=sunrise_offset_minutes)
sunset_time -= datetime.timedelta(minutes=sunset_offset_minutes)

sunrise_time = sunrise_time.time()
sunset_time = sunset_time.time()

print("当前时间：", datetime.datetime.now(datetime.timezone.utc).astimezone())
print("日出时间：", sunrise_time)
print("日落时间：", sunset_time)


# 广播系统设置变更
def broadcast_setting_change():
    HWND_BROADCAST = 0xFFFF
    WM_SETTINGCHANGE = 0x001A
    SMTO_ABORTIFHUNG = 0x0002

    # 同时刷新主题和环境设置，保证资源管理器和应用同步
    for setting in ("ImmersiveColorSet", "Environment"):
        ctypes.windll.user32.SendMessageTimeoutW(
            HWND_BROADCAST,
            WM_SETTINGCHANGE,
            0,
            setting,
            SMTO_ABORTIFHUNG,
            5000,
            ctypes.byref(ctypes.c_ulong())
        )
    print("广播系统设置，刷新主题和环境设置")


# 获取当前主题模式
def get_current_mode():
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "light" if value == 1 else "dark"
    except Exception as e:
        print("获取当前主题失败:", e)
    return None


# 切换壁纸
def set_wallpaper_by_mode(mode):

    # Win10 且未配置壁纸路径 → 不碰壁纸
    if is_win10() and not light_wallpaper_path and not dark_wallpaper_path:
        print("Win10 未配置壁纸路径，跳过壁纸切换")
        return

    # 从 .theme 文件读取壁纸路径
    def get_wallpaper_from_theme(theme_file):
        if not os.path.exists(theme_file):
            print(f"主题文件不存在: {theme_file}")
            return None

        cfg = configparser.ConfigParser(interpolation=None)
        try:
            with open(theme_file, "r", encoding="mbcs", errors="ignore") as f:
                cfg.read_file(f)
            wallpaper = cfg["Control Panel\\Desktop"]["Wallpaper"]
            wallpaper = os.path.expandvars(wallpaper)
            print('壁纸路径：', wallpaper)
            return wallpaper if os.path.exists(wallpaper) else None
        except Exception as e:
            print(f"读取壁纸失败: {e}")
            return None

    #切换壁纸动作
    def set_wallpaper(path):
        SPI_SETDESKWALLPAPER = 20
        SPIF_UPDATEINIFILE = 1
        SPIF_SENDWININICHANGE = 2
        ctypes.windll.user32.SystemParametersInfoW(
            SPI_SETDESKWALLPAPER, 0, path,
            SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE
        )

    # 如果用户自定义了壁纸路径，直接使用
    if mode == "light" and light_wallpaper_path:
        wp = light_wallpaper_path
    elif mode == "dark" and dark_wallpaper_path:
        wp = dark_wallpaper_path
    else:
        theme_file = light_theme_path if mode == "light" else dark_theme_path
        wp = get_wallpaper_from_theme(theme_file)

    if wp:
        print(f"切换壁纸: {wp}")
        set_wallpaper(wp)
    else:
        print(f"未能获取壁纸")


# 设置屏幕亮度
def set_brightness(brightness):
    if brightness is None:
        return
    try:
        cmd = f'Get-CimInstance -Namespace root/WMI -ClassName WmiMonitorBrightnessMethods | Invoke-CimMethod -MethodName WmiSetBrightness -Arguments @{{Timeout = 0; Brightness = {brightness}}}'
        result = subprocess.run(
            ['powershell', '-Command', cmd],
            capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        if result.returncode == 0:
            print(f"屏幕亮度已设置为 {brightness}")
        else:
            err_msg = result.stderr.decode('utf-8', errors='replace').strip()
            print(f"设置屏幕亮度失败 (设备可能不支持屏幕亮度调节): {err_msg}")
    except FileNotFoundError:
        print("设置屏幕亮度失败: 未找到 PowerShell")
    except Exception as e:
        print(f"设置屏幕亮度失败: {e}")


# 设置主题模式（light=True 切浅色，False 切深色）
def set_theme(light=True):
    import winreg
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "AppsUseLightTheme", 0, winreg.REG_DWORD, 1 if light else 0)
            winreg.SetValueEx(key, "SystemUsesLightTheme", 0, winreg.REG_DWORD, 1 if light else 0)
        broadcast_setting_change()
        mode = "light" if light else "dark"
        print(f"主题已切换为 {mode}")
        set_wallpaper_by_mode(mode)   # 同时切换壁纸
    except Exception as e:
        print("切换主题失败:", e)
        return

    brightness = light_brightness if light else dark_brightness
    set_brightness(brightness)


# 根据时间决定期望模式
def expected_mode_by_time():
    now = datetime.datetime.now(datetime.timezone.utc).astimezone().time()
    return "light" if sunrise_time <= now < sunset_time else "dark"


# 当前主题文件路径不符合期望的时候切换
def theme_file_switch():

    # Win10 且未配置自定义主题 → 深色模式纯靠注册表，无需切换主题文件
    if is_win10() and not user_configured_theme:
        print("Win10 未配置自定义主题，跳过主题文件切换")
        return

    # 获取当前主题路径
    def get_theme_path():
        # 使用 winreg 查询当前主题路径，避免 subprocess 编码问题
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r"Software\Microsoft\Windows\CurrentVersion\Themes") as key:
                current_theme_path, _ = winreg.QueryValueEx(key, "CurrentTheme")
        except Exception:
            current_theme_path = ""
        expected_theme_path = light_theme_path if expected_mode_by_time() == "light" else dark_theme_path
        print('当前主题路径：',current_theme_path)
        print('期望主题路径：',expected_theme_path)
        return current_theme_path,expected_theme_path
    # 检测设置面板是否已经关闭
    def kill_settings_panel():
        print("正在检测设置面板是否已经关闭...")
        for _ in range(20):
            result = subprocess.run(
                'tasklist /fi "IMAGENAME eq SystemSettings.exe"',
                shell=True, capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            output = result.stdout.decode('utf-8', errors='replace')
            if "systemsettings.exe" in output.lower():
                print("检测到设置面板，将其关闭…")
                subprocess.run(
                    'taskkill /f /im "SystemSettings.exe"',
                    shell=True, capture_output=True,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            time.sleep(0.1)

    i = 0
    while True:
        # 检测主题文件是否一致
        print('检测主题路径是否一致')
        if os.path.normpath(get_theme_path()[0]).lower() == os.path.normpath(get_theme_path()[1]).lower():
            print('主题路径一致，无需切换')
            break
        else:
            print('主题文件路径不一致，切换主题文件')
            subprocess.run('start ms-settings:', shell=True,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           creationflags=subprocess.CREATE_NO_WINDOW)
            time.sleep(0.5)
            os.startfile(get_theme_path()[1])
            broadcast_setting_change()
            print('当前主题路径：',get_theme_path()[0])
            i += 1
            if i >= 15:
                error_msg = f"[ERROR] 超过 {i} 次仍未能切换到 {get_theme_path()[1]}\n"
                with open("err.log", "a", encoding="utf-8") as f:
                    f.write(f"{datetime.datetime.now()} - {error_msg}")
                print(error_msg.strip())
                break
    kill_settings_panel()

def main():
    # 按照需求修改注册表切换深浅色主题
    mode_to_set = None
    if len(sys.argv) == 1:
        mode_to_set = expected_mode_by_time()
        print(f"按时间切换，期望模式: {mode_to_set}")
    else:
        if "--mode" in sys.argv:
            try:
                mode_index = sys.argv.index("--mode") + 1
                mode_value = sys.argv[mode_index].lower()
                if mode_value in ("light", "dark"):
                    mode_to_set = mode_value
                    print(f"使用参数切换模式: {mode_to_set}")
                else:
                    print("参数错误，只能是 light 或 dark")
                    sys.exit(1)
            except IndexError:
                print("未指定 --mode 的值")
                sys.exit(1)
        else:
            print("未识别的参数")
            sys.exit(1)

    current_mode = get_current_mode()
    print(f"当前模式: {current_mode}")

    if mode_to_set == current_mode:
        print("当前模式已符合要求，无需切换")
    else:
        set_theme(light=(mode_to_set == "light"))


if __name__ == "__main__":
    main()
    sys.exit(0)
