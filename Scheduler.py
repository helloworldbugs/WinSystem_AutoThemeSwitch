import json
import os
import subprocess
import sys
import tempfile
from datetime import datetime, timedelta

import requests
import yaml


pwd = os.getcwd()

result = subprocess.run(['whoami'], shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
username = result.stdout.strip()

result_sid = subprocess.run(['whoami', '/user'], shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
sid_line = result_sid.stdout.strip().split()[-1]

with open(r'config.yaml', 'r', encoding='utf-8') as f:
    config = yaml.safe_load(f)
    LNG = config['Position']['LNG']
    LAT = config['Position']['LAT']
    time_offset = config.get('Time_offset', {})
    sunrise_offset_minutes = int(time_offset.get('sunrise_offset_minutes', 0) or 0)
    sunset_offset_minutes = int(time_offset.get('sunset_offset_minutes', 0) or 0)


def apply_time_offset(target_time, offset_minutes):
    return target_time - timedelta(minutes=offset_minutes)


def update_task_scheduler(sunrise, sunset):
    subprocess.run(
        'schtasks /delete /tn "AutoThemeSwitch\\mode_light" /f',
        shell=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    subprocess.run(
        'schtasks /delete /tn "AutoThemeSwitch\\mode_dark" /f',
        shell=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

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
            <WorkingDirectory>{pwd}</WorkingDirectory>
            </Exec>
        </Actions>
        </Task>
        """.lstrip()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="w", encoding="utf-16") as f:
            f.write(task_xml)
            temp_xml = f.name

        subprocess.run(
            ["schtasks", "/Create", "/TN", task_name, "/XML", temp_xml, "/F"],
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        os.remove(temp_xml)

    create_task("mode_light", sunrise, "light")
    create_task("mode_dark", sunset, "dark")


def outTimefile():
    target_date = (datetime.now() + timedelta(days=1)).date().isoformat()
    try:
        resp = requests.get(
            f"https://api.sunrise-sunset.org/json?lat={LAT}&lng={LNG}&formatted=0&date={target_date}",
            headers={'User-Agent': 'Mozilla/5.0'}
        )

        if resp.status_code == 200:
            open(r'datetime.json', 'w+', encoding='utf-8').write(resp.text)
            print("日出日落时间更新成功")
        else:
            print('响应码：', resp.status_code)
            print('响应体：', resp.text)
            print("获取日出日落时间失败")
            sys.exit(0)
    except Exception:
        sys.exit(0)

    with open(r'datetime.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data


def main():
    data = outTimefile()

    sunrise = apply_time_offset(
        datetime.fromisoformat(data['results']['sunrise']).astimezone(),
        sunrise_offset_minutes,
    )
    sunset = apply_time_offset(
        datetime.fromisoformat(data['results']['sunset']).astimezone(),
        sunset_offset_minutes,
    )

    print(f"日出时间：{sunrise.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"日落时间：{sunset.strftime('%Y-%m-%d %H:%M:%S')}")

    update_task_scheduler(sunrise, sunset)
    print("已更新计划任务时间！")


if __name__ == "__main__":
    main()
    sys.exit(0)
