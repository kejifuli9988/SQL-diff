# SQL Excel 比较工具

这是一个可打包成 Windows `.exe` 的小工具，用来比较两个 Excel 里 `SQL语句` 这一列。

## 功能

- 支持选择两个 Excel 文件，表头可以不同
- 两个文件都必须包含 `SQL语句` 列
- 输出四个工作表：
  - `表1不在表2`
  - `表2不在表1`
  - `两个表都有_表1格式`
  - `两个表都有_表2格式`
- 默认忽略空白差异：
  - 首尾空格
  - 多个空格
  - 换行

## 运行方式

本地直接运行：

```bash
pip install -r requirements.txt
python sql_diff_gui.py
```

## 打包成 Windows 可执行文件

先在一台安装了 Python 的 Windows 电脑上执行：

```bash
pip install -r requirements.txt
build_windows_exe.bat
```

打包完成后会生成：

```text
dist\SQL语句比较工具.exe
```

把这个 `.exe` 发到没有 Python 的 Windows 电脑上也可以直接运行。

## 用 GitHub 自动打包 Windows EXE

如果你手头没有 Windows 电脑，也可以直接在 Mac 上开发，然后交给 GitHub Actions 自动打包。

项目里已经加好了工作流文件：

`.github/workflows/build-windows-exe.yml`

使用方法：

1. 把当前目录初始化成 Git 仓库并推到 GitHub
2. 把默认分支设为 `main` 或 `master`
3. 提交并推送代码
4. 打开 GitHub 仓库的 `Actions`
5. 运行 `Build Windows EXE`，或者直接在推送后自动触发
6. 运行完成后，在该次 Action 的 `Artifacts` 里下载 `SQL-Windows-EXE`

下载后你会拿到打包好的：

```text
SQL语句比较工具.exe
```

如果你愿意，也可以只用手动触发，不依赖每次推送自动构建；这个工作流已经支持 `workflow_dispatch`。

## 使用说明

1. 选择表1 Excel
2. 选择表2 Excel
3. 选择输出目录
4. 点击“开始比较”
5. 程序会生成一个新的结果文件，例如：

```text
SQL比较结果_20260703_153000.xlsx
```

## 规则说明

- 比较依据是 `SQL语句` 列
- 如果同一个文件中同一条 SQL 出现多次，只保留首次出现的那一行
- 空的 `SQL语句` 会被忽略

## 后续可扩展

如果你想继续完善，这个程序后面还可以再加：

- 拖拽上传文件
- 比较完成后自动打开结果文件
- 输出统计数量
- 支持多个工作表选择
- 保留你原始脚本里的更多业务规则
