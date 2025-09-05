# nmap-xml2xlsx — 端口调研合并工具（Nmap XML / Excel / CSV -> 端口调研表）

> 将多个 Nmap XML 与本地 Excel/CSV 扫描结果合并、去重、标注“危险端口”，并导出美观的 Excel 报表。适合渗透测试、资产盘点与网管运维的端口调研流程。

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE) [![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org)

仓库：

* GitHub: `https://github.com/ouwenjin/nmap-xml2xlsx`
* Gitee: `https://gitee.com/zhkali/nmap-xml2xlsx`

---

## 主要功能

* 自动合并当前目录下所有 Nmap `.xml` 输出为一个临时 XML（可自定义名）。
* 解析 Nmap XML 中的 host/port 信息，支持多 address、IPv4/IPv6 优先判断。
* 解析本地 Excel/CSV（自动尝试 `utf-8`、回退 `gbk`；支持列名模糊匹配）。
* 合并所有来源的数据，字段归一化并严格去重（按 `IP, 端口/协议, 服务, 状态, 端口用途`）。
* 根据内置危险端口与服务字典自动标注危险端口（`是否必要开放` 列）。
* 导出为格式化的 Excel（冻结表头、自动筛选、表头样式、列宽、危险字体上色）。
* 原子写入输出文件（先写临时文件再替换，避免中间损坏）。
* 支持命令行参数：输入/输出文件、是否删除临时 XML、是否显示颜色/Unicode 框、开启详细日志等。

---

## 要求

* Python 3.8+
* 依赖包（示例）：

  ```
  pandas
  openpyxl
  tqdm
  ```

建议在虚拟环境中安装（venv / conda）。

示例 `requirements.txt`：

```
pandas>=1.3
openpyxl>=3.0
tqdm>=4.0
```

---

## 安装

```bash
# 克隆仓库
git clone https://github.com/ouwenjin/nmap-xml2xlsx.git
cd nmap-xml2xlsx

# 建议创建虚拟环境
python -m venv .venv
source .venv/bin/activate    # Linux / macOS
.venv\Scripts\activate     # Windows

# 安装依赖
pip install -r requirements.txt
```

或者把脚本直接拷到你的工具集中并保证依赖安装完毕即可运行。

---

## 使用说明（CLI）

脚本示例名：`nmap_merge.py`（或你实际保存的脚本名）。

基本用法：

```bash
python nmap_merge.py
```

可用参数（常用）：

```
--input, -i       输入 Excel/CSV 文件（默认 "开放端口.xlsx"）
--output, -o      输出文件名（默认 "端口调研表.xlsx"）
--temp-xml        临时合并 XML 名称（默认 "out.xml"）
--cleanup         处理完成后删除临时 out.xml
--no-unicode      使用 ASCII 边框（不使用 Unicode 盒绘字符）
--no-color        禁用颜色输出
--margin          横幅左侧外边距空格数
--pad             横幅内部左右边距
--verbose         开启 DEBUG 详细日志
```

示例：

1. 从默认 `开放端口.xlsx` + 目录中所有 `.xml` 生成表：

```bash
python nmap_merge.py
```

2. 指定输入与输出，并删除临时 `out.xml`：

```bash
python nmap_merge.py -i hosts.csv -o report.xlsx --cleanup
```

3. 在 CI / 无颜色终端运行、开启详细日志：

```bash
python nmap_merge.py --no-color --verbose
```

---

## 输出说明

输出 Excel 文件（默认 `端口调研表.xlsx`）包含如下列：

* `IP`：目标 IP（优先 IPv4）
* `端口/协议`：例如 `80/tcp`、`53/udp`
* `状态`：端口状态（open/closed 等）
* `服务`：服务名（来自 Nmap / 表格）
* `端口用途`：备注 / 用途（可由输入表填写）
* `是否必要开放`：当端口或服务被列入危险集合时，标注 `危险端口不允许对外开放`

Excel 特性：

* 首行冻结、自动筛选（方便按 IP/服务筛查）
* 危险行字体为红色
* 表头加粗并有浅色背景

示例数据：

| IP           | 端口/协议    | 状态   | 服务    | 端口用途 | 是否必要开放      |
| ------------ | -------- | ---- | ----- | ---- | ----------- |
| 192.168.1.10 | 22/tcp   | open | ssh   | 管理端口 |             |
| 10.0.0.5     | 3306/tcp | open | mysql | 数据库  | 危险端口不允许对外开放 |

---

## 实现流程（简述）

1. 在当前目录查找 `.xml` 文件，若存在则把它们合并为一个临时 XML（`out.xml`）。
2. 解析用户指定的 Excel/CSV（支持常见列名的模糊匹配）；读取 IP/端口/服务/备注。
3. 解析合并后的 Nmap XML，抽取 host -> port -> service。
4. 合并两部分数据，进行字段归一化（trim、小写、合并空格）。
5. 去重（`IP, 端口/协议, 服务, 状态, 端口用途`）。
6. 根据危险端口/服务字典标注 `是否必要开放`。
7. 以临时文件方式导出为 Excel 并格式化，最后替换目标文件。
8. 可选：删除临时 XML 文件。

---

## 可定制项

* 危险端口集合与危险服务集合在脚本中以 `dangerous_ports` 与 `dangerous_services` 定义，直接在脚本中编辑即可：

```python
dangerous_ports = {20,21,22,23,25,3306,3389,6379,11211,27017, ...}
dangerous_services = {'ftp','telnet','ssh','mysql','redis','mongodb', ...}
```

* 可扩展：把这两个集合移到外部配置文件（YAML/JSON），或通过 CLI 传入定制配置（可由后续 PR 实现）。

---

## 常见问题与排错

* **找不到 `openpyxl` / `pandas` 报错**
  → `pip install -r requirements.txt`。

* **Excel 打开后显示损坏或无法写入**
  → 请确保没有其他程序（如 Excel）正在占用目标文件；脚本使用临时文件再替换，若替换失败请检查写权限。

* **CSV 编码问题（读取失败/乱码）**
  → 脚本会自动尝试 `utf-8`，失败后尝试 `gbk`；如依然乱码，请手动转成 UTF-8。

* **没有检测到任何 .xml 文件**
  → 脚本会跳过合并步骤，此时仅使用 `--input` 指定的 Excel/CSV 数据生成报告。

* **去重结果不尽如人意**
  → 脚本在去重前会将关键字段转为小写并压缩空格，若你希望更宽松或更严格的去重规则，可修改 `auto_dedup` 函数。

---

## 开发 & 贡献指南

欢迎贡献（Issue / PR / Star）！

建议流程：

1. Fork 仓库 → 新建 feature 分支（`feature/xxxx`）。
2. 提交清晰的 commit（遵循语义化 commit），并在 PR 描述中写明更改内容与动机。
3. 保持代码风格一致（PEP8）；若增加依赖请更新 `requirements.txt`。
4. 增加或修改功能请提供使用示例与简单说明。

可考虑的改进（欢迎实现）：

* 将危险端口/服务拆到配置文件（YAML/JSON）。
* 生成 summary sheet（按 IP 统计、按服务统计）。
* 支持更多扫描器（Nessus/ AWVS / 绿盟 HTML）解析器插件化。
* 打包成 pip 包并发布到 PyPI。
* CI（GitHub Actions）跑 lint / unit tests。

---

## 作者 & 致谢

作者：`zhkali`
仓库：

* GitHub: `https://github.com/ouwenjin/nmap-xml2xlsx`
* Gitee: `https://gitee.com/zhkali/nmap-xml2xlsx`

感谢使用者的反馈与贡献！
