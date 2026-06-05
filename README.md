# 说明

Windows 系统至今没有自动切换深浅色主题的功能，苦天下久矣，于是我自己写了一套 Python3 脚本搞定这事。

全程基于 Windows 计划任务实现，无后台常驻进程，不占资源也不存在进程保活问题。

## 工作流程总览

```
首次运行 Start_Run.py
  ├─ 创建计划任务：SwitchTheme（登录/解锁触发）
  ├─ 创建计划任务：Scheduler（开机15分钟触发）
  └─ 立即执行一次主题切换

Scheduler（每次开机 15 分钟后执行）
  ├─ 从 API 获取当天日出日落时间 → 缓存到 datetime.json
  ├─ 应用 config.yaml 中的时间偏移量
  └─ 更新 mode_light / mode_dark 计划任务的触发时间

SwitchTheme（每次登录 / 解锁时触发）
  ├─ 判断当前时间属于日间还是夜间
  ├─ 模式不匹配 → 修改注册表 → 广播刷新 → 切换壁纸 → 调节亮度
  └─ 模式已匹配 → 跳过

mode_light（日出时触发）  → 强制切换为浅色主题 + 壁纸 + 亮度
mode_dark （日落时触发）  → 强制切换为深色主题 + 壁纸 + 亮度
```

这样即便电脑在切换时间点处于休眠/锁屏/关机状态，下次登录或解锁时也会立刻补上切换。

## 快速开始

确保本机有 Python3，装好依赖，然后执行安装脚本即可：

```bash
pip install -r requirements.txt
py Start_Run.py
```

## 配置 `config.yaml`

```yaml
# ─── 经纬度（获取本地日出日落时间）───
Position:
  LNG: '113.264434'      # 经度
  LAT: '23.129162'       # 纬度

# ─── 时间偏移（分钟）───
# 正值 = 提前，负值 = 延后
Time_offset:
  sunrise_offset_minutes: 0    # 日出偏移
  sunset_offset_minutes: 0     # 日落偏移

# ─── 自定义主题路径 ───
# 留空使用系统默认（浅色 aero.theme / 深色 dark.theme）
Theme_path:
  light_theme_path: ''
  dark_theme_path: ''

# ─── 自定义壁纸路径 ───
# 留空则自动从主题文件中读取
Wallpaper_path:
  light_wallpaper_path: ''
  dark_wallpaper_path: ''

# ─── 屏幕亮度（0-100） ───
# 留空则切换主题时不改变亮度
# 注：仅支持笔记本电脑，台式机/外接显示器不支持亮度调节
Brightness:
  light_brightness:      # 浅色模式亮度
  dark_brightness:       # 深色模式亮度
```
>Windows默认主题路径：`C:\Windows\Resources\Themes`
>用户自定义的主题路径：`%homepath%\AppData\Local\Microsoft\Windows\Themes`

> 路径用单引号包裹，可直接粘贴 Windows 风格路径，无需反斜杠转义，支持 `%homepath%`、`%userprofile%`、`%appdata%` 等环境变量。

> 每次修改配置文件后，重新执行 `py Start_Run.py` 生效。

## 手动切换

```bash
py SwitchTheme.py                  # 根据当前时间自动判断
py SwitchTheme.py --mode light     # 强制浅色
py SwitchTheme.py --mode dark      # 强制深色
```

## 文件结构

| 文件 | 用途 |
|------|------|
| `Start_Run.py` | 安装入口，注册全部计划任务并首次执行 |
| `SwitchTheme.py` | 主题切换核心：读写注册表 + 切换壁纸 + 调节亮度 |
| `Scheduler.py` | 每日更新日出日落数据，刷新计划任务触发时间 |
| `config.yaml` | 用户配置 |
| `datetime.json` | 日出日落 API 数据本地缓存 |
| `requirements.txt` | Python 依赖（requests, pyyaml） |

## 已知缺陷

~~由于微软的屎山代码，切换主题时有概率自动弹出设置面板，无法避免。脚本内置了循环检测自动关闭设置面板的逻辑，影响大约 0.5 秒——你会看到设置面板一闪而过然后自动关掉。~~

v2.0 之后改为直接修改注册表 + 广播系统更新，不再弹窗。仅首次手动切换主题文件时可能会短暂弹出设置面板，后续自动任务完全无感。
