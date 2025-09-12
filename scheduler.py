import requests
from datetime import datetime, timedelta
import subprocess
import tempfile
import os
import json
import sys

# 获取当前目录
pwd = os.getcwd()

# 获取用户名和 SID
result = subprocess.run(['whoami'], capture_output=True, text=True, check=True)
username = result.stdout.strip()

result_sid = subprocess.run(['whoami', '/user'], capture_output=True, text=True, check=True)
sid_line = result_sid.stdout.strip().split()[-1]

print(f"用户名: {username}")
print(f"SID: {sid_line}")

# 配置经纬度
LNG = 113.32446000000004  # 替换为你的经度
LAT = 23.106469999999987   # 替换为你的纬度

def update_task_scheduler(sunrise, sunset):
    """使用schtasks命令更新计划任务"""
    # 删除旧任务
    subprocess.run(
        'schtasks /delete /tn "AutoThemeSwitch\\Switch_light" /f',
        shell=True,
        stdout=subprocess.PIPE
    )
    subprocess.run(
        'schtasks /delete /tn "AutoThemeSwitch\\Switch_dark" /f',
        shell=True,
        stdout=subprocess.PIPE
    )
    # 创建新任务
    def create_task(name, time, mode):
        task_name = rf"\AutoThemeSwitch\{name}"
        start_time = time.isoformat()

        task_xml = rf"""
        <?xml version="1.0" encoding="UTF-16"?>
        <Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
        <RegistrationInfo>
            <Author>{username}</Author>
        </RegistrationInfo>
        <Triggers>
            <CalendarTrigger>
            <StartBoundary>{start_time}</StartBoundary>
            <Enabled>true</Enabled>
            <ScheduleByDay>
                <DaysInterval>1</DaysInterval>
            </ScheduleByDay>
            </CalendarTrigger>
        </Triggers>
        <Principals>
            <Principal id="Author">
            <UserId>{sid_line}</UserId>
            <LogonType>InteractiveToken</LogonType>
            <RunLevel>LeastPrivilege</RunLevel>
            </Principal>
        </Principals>
        <Settings>
            <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
            <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
            <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
            <AllowHardTerminate>true</AllowHardTerminate>
            <StartWhenAvailable>false</StartWhenAvailable>
            <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
            <IdleSettings>
            <Duration>PT10M</Duration>
            <WaitTimeout>PT1H</WaitTimeout>
            <StopOnIdleEnd>true</StopOnIdleEnd>
            <RestartOnIdle>false</RestartOnIdle>
            </IdleSettings>
            <AllowStartOnDemand>true</AllowStartOnDemand>
            <Enabled>true</Enabled>
            <Hidden>false</Hidden>
            <RunOnlyIfIdle>false</RunOnlyIfIdle>
            <WakeToRun>false</WakeToRun>
            <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
            <Priority>7</Priority>
        </Settings>
        <Actions Context="Author">
            <Exec>
            <Command>pythonw.exe</Command>
            <Arguments>"{pwd}\SwitchTheme.py" --mode {mode}</Arguments>
            </Exec>
        </Actions>
        </Task>
        """.lstrip()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="w", encoding="utf-16") as f:
            f.write(task_xml)
            temp_xml = f.name

        subprocess.run(["schtasks", "/Create", "/TN", task_name, "/XML", temp_xml, "/F"])

        os.remove(temp_xml)

        # print(f"计划任务 {task_name} 已创建完成。")


    # 更新任务
    create_task("Switch_light", sunrise, "light")
    create_task("Switch_dark", sunset, "dark")

def outTimefile():
    # 获取次日的日出日落时间
    target_date = (datetime.now() + timedelta(days=1)).date().isoformat()
    try:
        resp = requests.get(
            f"https://api.sunrise-sunset.org/json?lat={LAT}&lng={LNG}&formatted=0&date={target_date}",
            headers={'User-Agent': 'Mozilla/5.0'}
        )
        
        # data = resp.json()['results']
        if resp.status_code == 200:
            open(r'datetime.json','w+',encoding='utf-8').write(resp.text)
            print("日出日落时间更新成功")

        else:
            print('响应码：',resp.status_code)
            print('响应体：',resp.text)
            print("获取日出日落时间失败！")
            sys.exit(0)
    except:
        sys.exit(0)
    # 读取本地缓存的时间文件
    with open(r'datetime.json','r',encoding='utf-8') as f:
        data = json.load(f)
    return data

def main():
    data = outTimefile()

    sunrise = datetime.fromisoformat(data['results']['sunrise']).astimezone()
    sunset = datetime.fromisoformat(data['results']['sunset']).astimezone()

    sunrise += timedelta(minutes=30)    # 日出时间延迟30分钟
    sunset -= timedelta(minutes=30)     # 日落时间提前30分钟

    print(f"日出时间：{sunrise.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"日落时间：{sunset.strftime('%Y-%m-%d %H:%M:%S')}")
    
    update_task_scheduler(sunrise, sunset)
    print(f"已更新计划任务时间！")


if __name__ == "__main__":
    main()
    sys.exit(0)