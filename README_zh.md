# neware-batch-export（新威批量导出工具）

**将新威 `.ndax` 文件批量导出为完整 8-sheet `.xlsx`——带 GUI 界面、循环统计方式选择、并行处理和邮件通知。**

> 📖 [English README](README.md)

> 本项目为 [新威开发者计划](https://github.com/NEWARE-Tech/neware-official/discussions/2) 贡献项目。

---

## 痛点

新威 BTS 8.0 软件的两种导出模式存在根本性的数据质量差距：

| 模式 | 导出内容 | 适用性 |
|------|---------|-------|
| **单文件手动导出** | 8 个完整 sheet——unit / test / cycle / **step** / **record** / log / idle / curve | ✅ 数据完整 |
| **内置批量导出** | 数据被阉割——**step 层和 record 层缺失** | ❌ 数据不完整 |

`step` 和 `record` 层是机器学习和数据驱动电池研究中最重要的数据——包含每个记录时刻的完整时间序列，没有这两层数据无法还原充放电曲线、提取特征或构建可靠的 ML 模型。

但手动单文件导出无法规模化：面对几十上百个 `.ndax` 文件逐一点击不现实。

**这一矛盾卡住了所有使用新威设备做 AI/ML 电池研究的课题组：要么批量但残缺，要么完整但手动，二者不可兼得。** 本项目解决这个问题。

---

## 解决方案

发现 `BTSDAExReport.exe`（BTS 8.0 自带）有一个未公开的命令行接口：

```
BTSDAExReport.exe export custom "<输入文件.ndax>" "<输出文件.xlsx>"
```

`custom` 模式会读取你在 BTSDA 软件中保存的 `BTSDAConfig.iconf` 配置，导出与手动单文件操作**完全一致的 8-sheet 完整数据**，包含完整的 step 层和 record 层。我们通过逐 sheet 对比行数的方式验证了这一点，结果完全一致。

本项目在此基础上封装了一个 GUI 工具，支持批量处理、并行导出和额外的工作流功能。

---

## 功能

- ✅ **完整 8-sheet 导出** — 与手动单文件导出完全一致（step/record 数据完整）
- ✅ **循环统计方式选择** — 工步默认 / 先充后放 / 先放后充 / 起始工步（导出前自动写入 iconf）
- ✅ **并行处理** — 可配置并行数（默认 3）
- ✅ **中英文 UI 一键切换**
- ✅ **灵活路径输入** — 支持单个或多个 `r"..."` Python 格式路径，`#` 注释行自动跳过
- ✅ **文件夹扫描** — 浏览目录，从勾选列表中选取 `.ndax` 文件
- ✅ **自定义输出目录** — 原文件同目录或指定目录
- ✅ **跳过已存在文件** — 默认勾选，重复运行不重复导出
- ✅ **汇总 CSV** — 每批次完成后自动生成（文件名/循环数/大小/用时）
- ✅ **邮件简报** — 可选：完成后发送含汇总表和 CSV 附件的通知邮件
- ✅ **历史记录** — 上次使用的路径和设置下次启动自动恢复
- ✅ **BTSDA 冲突检测** — 检测到 BTSDA.exe 运行时弹出警告
- ✅ **iconf 备份** — 修改前自动备份，出错自动还原

---

## 环境要求

- Windows（BTSDA 仅支持 Windows）
- 已安装新威 BTS 8.0 软件
- Python ≥ 3.10
- [`customtkinter`](https://github.com/TomSchimansky/CustomTkinter)

```bash
pip install customtkinter
```

---

## 安装

1. 克隆或下载本仓库
2. 安装依赖：`pip install customtkinter`
3. 配置路径（见[配置说明](#配置说明)）
4. 双击 **`neware-export.bat`** 启动

---

## 使用方式

### GUI 界面

```bash
python neware_export_gui.py
```

或直接双击 `neware-export.bat`。

**操作流程：**
1. 选择**循环统计方式**（大多数实验推荐使用工步默认）
2. 粘贴 `.ndax` 路径，支持 Python list 格式（含注释行）：
   ```python
   r"E:\data\experiment_01.ndax",
   r"E:\data\experiment_02.ndax",
   # r"E:\data\skip_this.ndax",    ← 注释行自动跳过
   r"E:\data\experiment_03.ndax",
   ```
3. 或点击**扫描文件夹**按钮，浏览目录并勾选文件
4. 设置输出目录
5. 点击**开始导出**

### PowerShell 脚本（无界面）

```powershell
.\neware_batch_export.ps1 -InputDir "E:\path\to\ndax"
.\neware_batch_export.ps1 -InputDir "E:\data" -OutputDir "E:\xlsx" -Recurse -Force
```

---

## 配置说明

打开 `neware_export_gui.py`，修改顶部常量：

```python
# ── 路径配置（根据你的安装路径修改）──────────────────────────────
BTSDA_EXE  = r"E:\software\BTSClient80\BTSDAExReport.exe"
ICONF_PATH = r"C:\Users\<你的用户名>\Documents\NEWARE\BTSClient\BTSDAConfig.iconf"
```

### 邮件通知

邮件功能**默认关闭**。勾选界面中的「完成后发邮件简报」后，配置 `neware_export_gui.py` 中的 SMTP 信息：

```python
SMTP_SERVER     = "smtp.qq.com"         # 根据你的邮件服务商修改
SMTP_PORT       = 465
SENDER_EMAIL    = "your@email.com"
SENDER_PASSWORD = "your_smtp_auth_code" # SMTP 授权码，非登录密码
RECEIVER_EMAIL  = "recipient@email.com"
```

常见邮件服务商配置：

| 服务商 | SMTP 地址 | 端口 |
|--------|-----------|------|
| QQ 邮箱 | `smtp.qq.com` | 465 |
| 163 邮箱 | `smtp.163.com` | 465 |
| Gmail | `smtp.gmail.com` | 465 |
| Outlook | `smtp.office365.com` | 587 |

> ⚠️ **请勿将真实凭据提交到 git。** 建议使用环境变量或本地配置文件。

---

## 循环统计方式说明

控制 BTS 如何统计每个循环的充放电容量，导出前写入 `BTSDAConfig.iconf`。

| 方式 | 值 | 说明 |
|------|---|------|
| 工步默认 | 0 | 按工步默认分配（推荐） |
| 先充后放 | 1 | 先充电后放电统计 |
| 先放后充 | 2 | 先放电后充电统计 |
| 起始工步 | 3 | 自定义起始工步 |

---

## 输出 Sheet 说明

| Sheet | 内容 |
|-------|------|
| `unit` | 单位设置 |
| `test` | 通道/测试元信息 |
| `cycle` | 逐循环统计（容量、库伦效率、能量、DCIR…） |
| `step` | 逐工步统计 |
| `record` | **完整时间序列**——每个记录点的数据 |
| `log` | 事件日志 |
| `idle` | 静置段数据 |
| `curve` | 计算曲线数据 |

`record` 层是 ML 特征提取的主要数据来源。

---

## 与 aurora-neware 的关系

[`aurora-neware`](https://github.com/EmpaEconversion/aurora-neware) 负责实验**运行期间**的设备控制（启停通道、监控状态），本项目处理**实验结束后**的数据导出，二者互补：

```
运行实验 → [aurora-neware：监控与控制]
                    ↓
            实验结束
                    ↓
    [neware-batch-export：完整 xlsx 数据集]
                    ↓
          ML / 数据分析
```

---

## 验证环境

- BTS 8.0（BTSDAExReport.exe 版本 2024.06.24）
- Windows 10 / 11
- CT 系列和 BTS 系列测试仪

---

## 许可证

MIT

---

