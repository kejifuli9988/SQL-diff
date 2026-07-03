# SQL Excel 比较工具

这是一个可打包成 Windows `.exe` 的小工具，用来比较两个 Excel 里 `SQL语句` 这一列。

## 功能

- 支持选择两个 Excel 文件，表头可以不同
- 两个文件都必须包含 `SQL语句` 列
- 支持标准 `.xlsx`、`.xlsm`、`.xls`
- 对 `.xls` 文件会先按真实 Excel 读取，失败后再尝试按 HTML 表格读取，兼容部分“伪装成 xls”的导出文件
- 为支持 HTML 表格读取，依赖中包含 `lxml`
- 自动在前 30 行里识别真正表头，不要求第一行就是表头
- 输出 1 个结果 Excel，里面包含 4 个工作表：
  - `仅表1文件名`
  - `仅表2文件名`
  - `共有(表1文件名)`
  - `共有(表2文件名)`
- 输出文件名格式：`SQL比较结果_表1名_VS_表2名.xlsx`
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

1. 提交并推送代码到 GitHub 仓库
2. 打开 GitHub 仓库的 `Actions`
3. 运行 `Build Windows EXE`，或者直接在推送后自动触发
4. 运行完成后，在该次 Action 的 `Artifacts` 里下载 `SQL-Windows-EXE`

下载后你会拿到打包好的：

```text
SQL语句比较工具.exe
```

## 使用说明

1. 选择表1 Excel
2. 选择表2 Excel
3. 选择输出目录
4. 点击“开始比较”
5. 程序会生成类似下面的结果文件：

```text
SQL比较结果_表1名_VS_表2名.xlsx
```

## 规则说明

- 比较依据是 `SQL语句` 列
- 如果同一个文件中同一条 SQL 出现多次，只保留首次出现的那一行
- 空的 `SQL语句` 会被忽略
- 如果文件内容像 Excel 但实际不是标准 `.xlsx/.xlsm`，程序会给出更明确的错误提示

## 后续可扩展

如果你想继续完善，这个程序后面还可以再加：

- 拖拽上传文件
- 比较完成后自动打开结果文件
- 输出统计数量
- 支持多个工作表选择
- 进一步减少 exe 体积
