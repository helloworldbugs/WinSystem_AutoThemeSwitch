import subprocess
import tempfile
import os
import scheduler

# 获取当前目录
pwd = os.getcwd()

# 获取用户名和 SID
result = subprocess.run(['whoami'], capture_output=True, text=True, check=True)
username = result.stdout.strip()

result_sid = subprocess.run(['whoami', '/user'], capture_output=True, text=True, check=True)
sid_line = result_sid.stdout.strip().split()[-1]

print(f"用户名: {username}")
print(f"SID: {sid_line}")

def create_AutoThemeSwitch():
  # 任务名称
  task_name = r"\AutoThemeSwitch\AutoThemeSwitch"

  # XML 模板
  task_xml = rf"""
  <?xml version="1.0" encoding="UTF-16"?>
  <Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
    <RegistrationInfo>
      <Author>{username}</Author>
    </RegistrationInfo>
    <Triggers>
      <LogonTrigger>
        <Enabled>true</Enabled>
      </LogonTrigger>
      <SessionStateChangeTrigger>
        <Enabled>true</Enabled>
        <StateChange>SessionUnlock</StateChange>
      </SessionStateChangeTrigger>
    </Triggers>
    <Principals>
      <Principal id="Author">
        <UserId>{sid_line}</UserId>
        <LogonType>InteractiveToken</LogonType>
        <RunLevel>LeastPrivilege</RunLevel>
      </Principal>
    </Principals>
    <Settings>
      <MultipleInstancesPolicy>Parallel</MultipleInstancesPolicy>
      <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
      <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
      <AllowHardTerminate>true</AllowHardTerminate>
      <StartWhenAvailable>false</StartWhenAvailable>
      <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
      <IdleSettings>
        <StopOnIdleEnd>true</StopOnIdleEnd>
        <RestartOnIdle>false</RestartOnIdle>
      </IdleSettings>
      <AllowStartOnDemand>true</AllowStartOnDemand>
      <Enabled>true</Enabled>
      <Hidden>false</Hidden>
      <RunOnlyIfIdle>false</RunOnlyIfIdle>
      <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
      <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
      <WakeToRun>false</WakeToRun>
      <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
      <Priority>7</Priority>
    </Settings>
    <Actions Context="Author">
      <Exec>
        <Command>pythonw.exe</Command>
        <Arguments>"{pwd}\SwitchTheme.py"</Arguments>
        <WorkingDirectory>{pwd}</WorkingDirectory>
      </Exec>
    </Actions>
  </Task>
  """.lstrip()

  # 写临时 xml 文件
  with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="w", encoding="utf-16") as f:
      f.write(task_xml)
      temp_xml = f.name

  # 注册任务
  subprocess.run(["schtasks", "/Create", "/TN", task_name, "/XML", temp_xml, "/F"], shell=True)

  # 删除临时文件
  os.remove(temp_xml)

  # print(f"计划任务 {task_name} 已创建完成。")

def create_AutoThemeScheduler():
  # 任务名称
  task_name = r"\AutoThemeSwitch\AutoThemeScheduler"

  # XML 模板
  task_xml = rf"""
  <?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2025-02-17T15:28:13.9902886</Date>
    <Author>{username}</Author>
  </RegistrationInfo>
  <Triggers>
    <BootTrigger>
      <Enabled>true</Enabled>
      <Delay>PT15M</Delay>
    </BootTrigger>
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
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>pythonw.exe</Command>
      <Arguments>"{pwd}\scheduler.py"</Arguments>
      <WorkingDirectory>{pwd}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
  """.lstrip()

  # 写临时 xml 文件
  with tempfile.NamedTemporaryFile(delete=False, suffix=".xml", mode="w", encoding="utf-16") as f:
      f.write(task_xml)
      temp_xml = f.name

  # 注册任务
  subprocess.run(["schtasks", "/Create", "/TN", task_name, "/XML", temp_xml, "/F"], shell=True)

  # 删除临时文件
  os.remove(temp_xml)

  # print(f"计划任务 {task_name} 已创建完成。")


if __name__ == "__main__":
  create_AutoThemeSwitch()
  create_AutoThemeScheduler()
  scheduler.main()
  input("\n完成!!!")
